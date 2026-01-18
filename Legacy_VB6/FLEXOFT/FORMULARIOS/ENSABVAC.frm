VERSION 5.00
Begin VB.Form ENSABVAC 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ensayo de Bombas de Vacio"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
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
   ScaleHeight     =   8295
   ScaleWidth      =   11910
   Begin VB.CommandButton Command13 
      Caption         =   "List"
      Height          =   855
      Left            =   10920
      Picture         =   "ENSABVAC.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   165
      ToolTipText     =   "Emisión de Listado"
      Top             =   4000
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ver"
      Height          =   855
      Left            =   10920
      Picture         =   "ENSABVAC.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   159
      ToolTipText     =   "Ver y Re-Imprimir Comprobantes de Ensayo"
      Top             =   5670
      Width           =   855
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ensayo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   4830
      TabIndex        =   154
      Top             =   105
      Width           =   1590
      Begin VB.Label Label70 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "N°"
         Height          =   225
         Left            =   -585
         TabIndex        =   158
         Top             =   630
         Width           =   1080
      End
      Begin VB.Label Label71 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
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
         Height          =   285
         Left            =   525
         TabIndex        =   157
         Top             =   315
         Width           =   960
      End
      Begin VB.Label Label73 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   525
         TabIndex        =   156
         Top             =   630
         Width           =   960
      End
      Begin VB.Label Label72 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "F:"
         Height          =   225
         Left            =   -735
         TabIndex        =   155
         Top             =   315
         Width           =   1185
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Bomba"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   105
      TabIndex        =   115
      Top             =   105
      Width           =   4635
      Begin VB.TextBox Text22 
         Height          =   285
         Left            =   945
         TabIndex        =   3
         Text            =   " "
         Top             =   630
         Width           =   3585
      End
      Begin VB.TextBox Text23 
         Height          =   285
         Left            =   3150
         TabIndex        =   2
         Text            =   " "
         Top             =   315
         Width           =   1380
      End
      Begin VB.CommandButton BOTSEL 
         Caption         =   "..."
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
         Index           =   31
         Left            =   2310
         Picture         =   "ENSABVAC.frx":0614
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   315
         Width           =   175
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Modelo:"
         Height          =   225
         Left            =   -315
         TabIndex        =   118
         Top             =   315
         Width           =   1185
      End
      Begin VB.Label Label62 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Corresp:"
         Height          =   225
         Left            =   -840
         TabIndex        =   117
         Top             =   630
         Width           =   1710
      End
      Begin VB.Label Label63 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nro:"
         Height          =   225
         Left            =   1890
         TabIndex        =   116
         Top             =   315
         Width           =   1185
      End
      Begin VB.Label Label66 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Left            =   945
         TabIndex        =   0
         Top             =   315
         Width           =   1380
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Motor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   105
      TabIndex        =   67
      Top             =   1155
      Width           =   6315
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   5250
         TabIndex        =   10
         Text            =   " "
         Top             =   1090
         Width           =   960
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   2940
         TabIndex        =   9
         Text            =   " "
         Top             =   1090
         Width           =   960
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   945
         TabIndex        =   4
         Text            =   " "
         Top             =   315
         Width           =   2325
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   4200
         TabIndex        =   5
         Text            =   " "
         Top             =   315
         Width           =   2010
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   945
         TabIndex        =   6
         Text            =   " "
         Top             =   630
         Width           =   960
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   5250
         TabIndex        =   7
         Text            =   " "
         Top             =   630
         Width           =   960
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   945
         TabIndex        =   8
         Text            =   " "
         Top             =   1090
         Width           =   960
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Marca:"
         Height          =   225
         Left            =   -630
         TabIndex        =   74
         Top             =   315
         Width           =   1500
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nro:"
         Height          =   225
         Left            =   2625
         TabIndex        =   73
         Top             =   315
         Width           =   1500
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pot [CV]:"
         Height          =   225
         Left            =   -630
         TabIndex        =   72
         Top             =   630
         Width           =   1500
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "vel.[rpm]:"
         Height          =   225
         Left            =   3675
         TabIndex        =   71
         Top             =   630
         Width           =   1500
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tens. [V]:"
         Height          =   225
         Left            =   -630
         TabIndex        =   70
         Top             =   1090
         Width           =   1500
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Int.[A]:"
         Height          =   225
         Left            =   1365
         TabIndex        =   69
         Top             =   1095
         Width           =   1500
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "frec. [Hz]:"
         Height          =   225
         Left            =   3675
         TabIndex        =   68
         Top             =   1095
         Width           =   1500
      End
   End
   Begin VB.TextBox Text25 
      Height          =   540
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   41
      Top             =   7665
      Width           =   10620
   End
   Begin VB.TextBox Text24 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   210
      TabIndex        =   47
      Text            =   " "
      Top             =   2310
      Width           =   2010
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ensayo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4880
      Left            =   105
      TabIndex        =   75
      Top             =   2730
      Width           =   10620
      Begin VB.TextBox Text13 
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
         Height          =   285
         Index           =   4
         Left            =   9135
         TabIndex        =   37
         Text            =   " "
         Top             =   1365
         Width           =   1350
      End
      Begin VB.TextBox Text14 
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
         Height          =   285
         Index           =   4
         Left            =   9135
         TabIndex        =   38
         Text            =   " "
         Top             =   1995
         Width           =   1350
      End
      Begin VB.TextBox Text15 
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
         Height          =   285
         Index           =   4
         Left            =   9135
         TabIndex        =   39
         Text            =   " "
         Top             =   2625
         Width           =   1350
      End
      Begin VB.TextBox Text16 
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
         Height          =   285
         Index           =   4
         Left            =   9135
         TabIndex        =   36
         Top             =   1050
         Width           =   1350
      End
      Begin VB.TextBox Text18 
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
         Height          =   285
         Index           =   4
         Left            =   9135
         TabIndex        =   40
         Text            =   " "
         Top             =   2940
         Width           =   1350
      End
      Begin VB.TextBox Text13 
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
         Height          =   285
         Index           =   3
         Left            =   7665
         TabIndex        =   32
         Text            =   " "
         Top             =   1365
         Width           =   1350
      End
      Begin VB.TextBox Text14 
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
         Height          =   285
         Index           =   3
         Left            =   7665
         TabIndex        =   33
         Text            =   " "
         Top             =   1995
         Width           =   1350
      End
      Begin VB.TextBox Text15 
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
         Height          =   285
         Index           =   3
         Left            =   7665
         TabIndex        =   34
         Text            =   " "
         Top             =   2625
         Width           =   1350
      End
      Begin VB.TextBox Text16 
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
         Height          =   285
         Index           =   3
         Left            =   7665
         TabIndex        =   31
         Top             =   1050
         Width           =   1350
      End
      Begin VB.TextBox Text18 
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
         Height          =   285
         Index           =   3
         Left            =   7665
         TabIndex        =   35
         Text            =   " "
         Top             =   2940
         Width           =   1350
      End
      Begin VB.TextBox Text13 
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
         Height          =   285
         Index           =   2
         Left            =   6195
         TabIndex        =   27
         Text            =   " "
         Top             =   1365
         Width           =   1350
      End
      Begin VB.TextBox Text14 
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
         Height          =   285
         Index           =   2
         Left            =   6195
         TabIndex        =   28
         Text            =   " "
         Top             =   1995
         Width           =   1350
      End
      Begin VB.TextBox Text15 
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
         Height          =   285
         Index           =   2
         Left            =   6195
         TabIndex        =   29
         Text            =   " "
         Top             =   2625
         Width           =   1350
      End
      Begin VB.TextBox Text16 
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
         Height          =   285
         Index           =   2
         Left            =   6195
         TabIndex        =   26
         Top             =   1050
         Width           =   1350
      End
      Begin VB.TextBox Text18 
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
         Height          =   285
         Index           =   2
         Left            =   6195
         TabIndex        =   30
         Text            =   " "
         Top             =   2940
         Width           =   1350
      End
      Begin VB.TextBox Text13 
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
         Height          =   285
         Index           =   1
         Left            =   4725
         TabIndex        =   22
         Text            =   " "
         Top             =   1365
         Width           =   1350
      End
      Begin VB.TextBox Text14 
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
         Height          =   285
         Index           =   1
         Left            =   4725
         TabIndex        =   23
         Text            =   " "
         Top             =   1995
         Width           =   1350
      End
      Begin VB.TextBox Text15 
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
         Height          =   285
         Index           =   1
         Left            =   4725
         TabIndex        =   24
         Text            =   " "
         Top             =   2625
         Width           =   1350
      End
      Begin VB.TextBox Text16 
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
         Height          =   285
         Index           =   1
         Left            =   4725
         TabIndex        =   21
         Top             =   1050
         Width           =   1350
      End
      Begin VB.TextBox Text18 
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
         Height          =   285
         Index           =   1
         Left            =   4725
         TabIndex        =   25
         Text            =   " "
         Top             =   2940
         Width           =   1350
      End
      Begin VB.TextBox Text18 
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
         Height          =   285
         Index           =   0
         Left            =   3255
         TabIndex        =   20
         Text            =   " "
         Top             =   2940
         Width           =   1350
      End
      Begin VB.TextBox Text16 
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
         Height          =   285
         Index           =   0
         Left            =   3255
         TabIndex        =   16
         Top             =   1050
         Width           =   1350
      End
      Begin VB.TextBox Text15 
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
         Height          =   285
         Index           =   0
         Left            =   3255
         TabIndex        =   19
         Text            =   " "
         Top             =   2625
         Width           =   1350
      End
      Begin VB.TextBox Text14 
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
         Height          =   285
         Index           =   0
         Left            =   3255
         TabIndex        =   18
         Text            =   " "
         Top             =   1995
         Width           =   1350
      End
      Begin VB.TextBox Text13 
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
         Height          =   285
         Index           =   0
         Left            =   3255
         TabIndex        =   17
         Text            =   " "
         Top             =   1365
         Width           =   1350
      End
      Begin VB.Label Label74 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   9135
         TabIndex        =   164
         Top             =   420
         Width           =   1350
      End
      Begin VB.Label Label74 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   7665
         TabIndex        =   163
         Top             =   420
         Width           =   1350
      End
      Begin VB.Label Label74 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   6195
         TabIndex        =   162
         Top             =   420
         Width           =   1350
      End
      Begin VB.Label Label74 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   4725
         TabIndex        =   161
         Top             =   420
         Width           =   1350
      End
      Begin VB.Label Label74 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3255
         TabIndex        =   160
         Top             =   420
         Width           =   1350
      End
      Begin VB.Label Label24 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Index           =   4
         Left            =   9135
         TabIndex        =   153
         Top             =   1680
         Width           =   1350
      End
      Begin VB.Label Label25 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Index           =   4
         Left            =   9135
         TabIndex        =   152
         Top             =   735
         Width           =   1350
      End
      Begin VB.Label Label51 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   9135
         TabIndex        =   151
         Top             =   2310
         Width           =   1350
      End
      Begin VB.Label Label52 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Index           =   4
         Left            =   9135
         TabIndex        =   150
         Top             =   3255
         Width           =   1350
      End
      Begin VB.Label Label53 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Index           =   4
         Left            =   9135
         TabIndex        =   149
         Top             =   3570
         Width           =   1350
      End
      Begin VB.Label Label67 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   9135
         TabIndex        =   148
         Top             =   3885
         Width           =   1350
      End
      Begin VB.Label Label68 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   9135
         TabIndex        =   147
         Top             =   4200
         Width           =   1350
      End
      Begin VB.Label Label69 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   9135
         TabIndex        =   146
         Top             =   4515
         Width           =   1350
      End
      Begin VB.Label Label24 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Index           =   3
         Left            =   7665
         TabIndex        =   145
         Top             =   1680
         Width           =   1350
      End
      Begin VB.Label Label25 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Index           =   3
         Left            =   7665
         TabIndex        =   144
         Top             =   735
         Width           =   1350
      End
      Begin VB.Label Label51 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   7665
         TabIndex        =   143
         Top             =   2310
         Width           =   1350
      End
      Begin VB.Label Label52 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Index           =   3
         Left            =   7665
         TabIndex        =   142
         Top             =   3255
         Width           =   1350
      End
      Begin VB.Label Label53 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Index           =   3
         Left            =   7665
         TabIndex        =   141
         Top             =   3570
         Width           =   1350
      End
      Begin VB.Label Label67 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   7665
         TabIndex        =   140
         Top             =   3885
         Width           =   1350
      End
      Begin VB.Label Label68 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   7665
         TabIndex        =   139
         Top             =   4200
         Width           =   1350
      End
      Begin VB.Label Label69 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   7665
         TabIndex        =   138
         Top             =   4515
         Width           =   1350
      End
      Begin VB.Label Label24 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Index           =   2
         Left            =   6195
         TabIndex        =   137
         Top             =   1680
         Width           =   1350
      End
      Begin VB.Label Label25 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Index           =   2
         Left            =   6195
         TabIndex        =   136
         Top             =   735
         Width           =   1350
      End
      Begin VB.Label Label51 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   6195
         TabIndex        =   135
         Top             =   2310
         Width           =   1350
      End
      Begin VB.Label Label52 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Index           =   2
         Left            =   6195
         TabIndex        =   134
         Top             =   3255
         Width           =   1350
      End
      Begin VB.Label Label53 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Index           =   2
         Left            =   6195
         TabIndex        =   133
         Top             =   3570
         Width           =   1350
      End
      Begin VB.Label Label67 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   6195
         TabIndex        =   132
         Top             =   3885
         Width           =   1350
      End
      Begin VB.Label Label68 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   6195
         TabIndex        =   131
         Top             =   4200
         Width           =   1350
      End
      Begin VB.Label Label69 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   6195
         TabIndex        =   130
         Top             =   4515
         Width           =   1350
      End
      Begin VB.Label Label24 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Index           =   1
         Left            =   4725
         TabIndex        =   129
         Top             =   1680
         Width           =   1350
      End
      Begin VB.Label Label25 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Index           =   1
         Left            =   4725
         TabIndex        =   128
         Top             =   735
         Width           =   1350
      End
      Begin VB.Label Label51 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   4725
         TabIndex        =   127
         Top             =   2310
         Width           =   1350
      End
      Begin VB.Label Label52 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Index           =   1
         Left            =   4725
         TabIndex        =   126
         Top             =   3255
         Width           =   1350
      End
      Begin VB.Label Label53 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Index           =   1
         Left            =   4725
         TabIndex        =   125
         Top             =   3570
         Width           =   1350
      End
      Begin VB.Label Label67 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   4725
         TabIndex        =   124
         Top             =   3885
         Width           =   1350
      End
      Begin VB.Label Label68 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   4725
         TabIndex        =   123
         Top             =   4200
         Width           =   1350
      End
      Begin VB.Label Label69 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   4725
         TabIndex        =   122
         Top             =   4515
         Width           =   1350
      End
      Begin VB.Label Label69 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3255
         TabIndex        =   121
         Top             =   4515
         Width           =   1350
      End
      Begin VB.Label Label68 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3255
         TabIndex        =   120
         Top             =   4200
         Width           =   1350
      End
      Begin VB.Label Label67 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3255
         TabIndex        =   119
         Top             =   3885
         Width           =   1350
      End
      Begin VB.Line Line19 
         X1              =   1890
         X2              =   1890
         Y1              =   4515
         Y2              =   4935
      End
      Begin VB.Line Line18 
         X1              =   525
         X2              =   525
         Y1              =   4515
         Y2              =   4935
      End
      Begin VB.Line Line17 
         X1              =   0
         X2              =   10395
         Y1              =   4515
         Y2              =   4515
      End
      Begin VB.Line Line16 
         X1              =   0
         X2              =   10395
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Label Label65 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "bomba"
         Height          =   225
         Left            =   630
         TabIndex        =   114
         Top             =   3675
         Width           =   1185
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "placa"
         Height          =   225
         Left            =   630
         TabIndex        =   113
         Top             =   3360
         Width           =   1185
      End
      Begin VB.Line Line15 
         X1              =   525
         X2              =   525
         Y1              =   3255
         Y2              =   3885
      End
      Begin VB.Line Line14 
         X1              =   0
         X2              =   10395
         Y1              =   3885
         Y2              =   3885
      End
      Begin VB.Line Line13 
         X1              =   0
         X2              =   10395
         Y1              =   3570
         Y2              =   3570
      End
      Begin VB.Line Line12 
         X1              =   0
         X2              =   10395
         Y1              =   3255
         Y2              =   3255
      End
      Begin VB.Line Line11 
         X1              =   0
         X2              =   10395
         Y1              =   2940
         Y2              =   2940
      End
      Begin VB.Line Line10 
         X1              =   0
         X2              =   10395
         Y1              =   2625
         Y2              =   2625
      End
      Begin VB.Line Line9 
         X1              =   1890
         X2              =   10395
         Y1              =   2310
         Y2              =   2310
      End
      Begin VB.Line Line8 
         X1              =   0
         X2              =   10395
         Y1              =   1995
         Y2              =   1995
      End
      Begin VB.Line Line7 
         X1              =   0
         X2              =   10395
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line6 
         X1              =   0
         X2              =   10395
         Y1              =   1365
         Y2              =   1365
      End
      Begin VB.Line Line5 
         X1              =   630
         X2              =   10395
         Y1              =   735
         Y2              =   735
      End
      Begin VB.Line Line4 
         X1              =   630
         X2              =   630
         Y1              =   420
         Y2              =   1050
      End
      Begin VB.Line Line3 
         X1              =   1890
         X2              =   1890
         Y1              =   420
         Y2              =   3885
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   10395
         Y1              =   1050
         Y2              =   1050
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   10395
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Label Label61 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         Height          =   225
         Left            =   7875
         TabIndex        =   111
         Top             =   210
         Width           =   870
      End
      Begin VB.Label Label60 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         Height          =   225
         Left            =   9345
         TabIndex        =   110
         Top             =   210
         Width           =   870
      End
      Begin VB.Label Label59 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         Height          =   225
         Left            =   6405
         TabIndex        =   109
         Top             =   210
         Width           =   870
      End
      Begin VB.Label Label58 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         Height          =   225
         Left            =   4935
         TabIndex        =   108
         Top             =   210
         Width           =   870
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         Height          =   225
         Left            =   3465
         TabIndex        =   107
         Top             =   210
         Width           =   870
      End
      Begin VB.Label Label57 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "[l/min]"
         Height          =   225
         Left            =   1995
         TabIndex        =   106
         Top             =   4620
         Width           =   1080
      End
      Begin VB.Label Label56 
         BackStyle       =   0  'Transparent
         Caption         =   " Qg         garant."
         Height          =   225
         Left            =   105
         TabIndex        =   105
         Top             =   4620
         Width           =   1500
      End
      Begin VB.Label Label55 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ct"
         Height          =   225
         Left            =   105
         TabIndex        =   104
         Top             =   4305
         Width           =   3075
      End
      Begin VB.Label Label54 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cv"
         Height          =   225
         Left            =   105
         TabIndex        =   103
         Top             =   3990
         Width           =   3075
      End
      Begin VB.Label Label53 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Index           =   0
         Left            =   3255
         TabIndex        =   102
         Top             =   3570
         Width           =   1350
      End
      Begin VB.Label Label52 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Index           =   0
         Left            =   3255
         TabIndex        =   101
         Top             =   3255
         Width           =   1350
      End
      Begin VB.Label Label51 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3255
         TabIndex        =   100
         Top             =   2310
         Width           =   1350
      End
      Begin VB.Label Label50 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "[°C]"
         Height          =   225
         Left            =   1995
         TabIndex        =   99
         Top             =   3045
         Width           =   1080
      End
      Begin VB.Label Label49 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "[l/min]"
         Height          =   225
         Left            =   1995
         TabIndex        =   98
         Top             =   3675
         Width           =   1080
      End
      Begin VB.Label Label48 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "[l/min]"
         Height          =   225
         Left            =   1995
         TabIndex        =   97
         Top             =   3360
         Width           =   1080
      End
      Begin VB.Label Label47 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Q1"
         Height          =   225
         Left            =   105
         TabIndex        =   96
         Top             =   3675
         Width           =   345
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Qab "
         Height          =   225
         Left            =   105
         TabIndex        =   95
         Top             =   3360
         Width           =   345
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "t1"
         Height          =   225
         Left            =   105
         TabIndex        =   94
         Top             =   3045
         Width           =   1710
      End
      Begin VB.Label Label44 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "[mm agua]"
         Height          =   225
         Left            =   1995
         TabIndex        =   93
         Top             =   2730
         Width           =   1080
      End
      Begin VB.Label Label43 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Delta Pab"
         Height          =   225
         Left            =   105
         TabIndex        =   92
         Top             =   2730
         Width           =   1710
      End
      Begin VB.Label Label42 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ctte"
         Height          =   225
         Left            =   1995
         TabIndex        =   91
         Top             =   2415
         Width           =   1080
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "N°"
         Height          =   225
         Left            =   1995
         TabIndex        =   90
         Top             =   2100
         Width           =   1080
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " Ramal Usado"
         Height          =   435
         Left            =   525
         TabIndex        =   89
         Top             =   2100
         Width           =   765
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "[°C]"
         Height          =   225
         Left            =   1995
         TabIndex        =   88
         Top             =   1785
         Width           =   1080
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "tl"
         Height          =   225
         Left            =   105
         TabIndex        =   87
         Top             =   1785
         Width           =   1710
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "[l/min]"
         Height          =   225
         Left            =   1995
         TabIndex        =   86
         Top             =   1470
         Width           =   1080
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Caudal"
         Height          =   225
         Left            =   105
         TabIndex        =   85
         Top             =   1470
         Width           =   1710
      End
      Begin VB.Label Label40 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "[mmHg]"
         Height          =   240
         Left            =   1995
         TabIndex        =   84
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "P1"
         Height          =   225
         Left            =   210
         TabIndex        =   83
         Top             =   630
         Width           =   660
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "absoluta"
         Height          =   195
         Left            =   945
         TabIndex        =   82
         Top             =   525
         Width           =   735
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "descenso"
         Height          =   225
         Left            =   525
         TabIndex        =   81
         Top             =   840
         Width           =   1185
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Intensidad"
         Height          =   225
         Left            =   105
         TabIndex        =   80
         Top             =   1155
         Width           =   1710
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "[A]"
         Height          =   225
         Left            =   1995
         TabIndex        =   79
         Top             =   1155
         Width           =   1080
      End
      Begin VB.Label Label25 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Index           =   0
         Left            =   3255
         TabIndex        =   78
         Top             =   735
         Width           =   1350
      End
      Begin VB.Label Label24 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Index           =   0
         Left            =   3255
         TabIndex        =   77
         Top             =   1680
         Width           =   1350
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "[mmHg]"
         Height          =   285
         Left            =   1995
         TabIndex        =   76
         Top             =   525
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Condiciones Ambientales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   6510
      TabIndex        =   50
      Top             =   105
      Width           =   4215
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1470
         TabIndex        =   15
         Text            =   "    15.0"
         Top             =   2145
         Width           =   1065
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1470
         TabIndex        =   14
         Text            =   "     Agua "
         Top             =   1840
         Width           =   1065
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1470
         TabIndex        =   13
         Text            =   "      50"
         Top             =   1230
         Width           =   1065
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1470
         TabIndex        =   12
         Text            =   "    20.0"
         Top             =   945
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1470
         TabIndex        =   11
         Text            =   "    1013 "
         Top             =   620
         Width           =   1065
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "15.0"
         Height          =   285
         Left            =   2520
         TabIndex        =   66
         Top             =   2205
         Width           =   1590
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Agua"
         Height          =   285
         Left            =   2520
         TabIndex        =   65
         Top             =   1890
         Width           =   1590
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "1.2"
         Height          =   285
         Left            =   2520
         TabIndex        =   64
         Top             =   1575
         Width           =   1590
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "50"
         Height          =   285
         Left            =   2520
         TabIndex        =   63
         Top             =   1260
         Width           =   1590
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "20"
         Height          =   285
         Left            =   2520
         TabIndex        =   62
         Top             =   945
         Width           =   1590
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   " S.M. 4514-8062"
         Height          =   285
         Left            =   2520
         TabIndex        =   61
         Top             =   630
         Width           =   1590
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   " 760"
         Height          =   285
         Left            =   2520
         TabIndex        =   60
         Top             =   315
         Width           =   1590
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  1.2049 "
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1470
         TabIndex        =   59
         Top             =   1530
         Width           =   1065
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   760.0 "
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1470
         TabIndex        =   58
         Top             =   315
         Width           =   1065
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "t l   [°C]:"
         Height          =   225
         Left            =   -105
         TabIndex        =   57
         Top             =   2205
         Width           =   1500
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Liquido sello:"
         Height          =   225
         Left            =   -105
         TabIndex        =   56
         Top             =   1890
         Width           =   1500
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "dens [kg/m3]:"
         Height          =   225
         Left            =   -105
         TabIndex        =   55
         Top             =   1575
         Width           =   1500
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "hum rel [%]:"
         Height          =   225
         Left            =   -105
         TabIndex        =   54
         Top             =   1260
         Width           =   1500
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "to   [ºC]:"
         Height          =   225
         Left            =   -105
         TabIndex        =   53
         Top             =   945
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "     [mb]:"
         Height          =   225
         Left            =   -105
         TabIndex        =   52
         Top             =   630
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Po [mm Hg]:"
         Height          =   225
         Left            =   -105
         TabIndex        =   51
         Top             =   315
         Width           =   1500
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "New"
      Height          =   855
      Left            =   10920
      Picture         =   "ENSABVAC.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Nuevo Modelo de Bomba de Vacío"
      Top             =   3045
      Width           =   855
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
      Left            =   10920
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   48
      Top             =   5145
      Width           =   855
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
         Left            =   735
         TabIndex        =   49
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton command1 
      Caption         =   "Exit"
      Height          =   855
      Left            =   10920
      Picture         =   "ENSABVAC.frx":0A68
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Cierra y Abandona la Aplicación"
      Top             =   210
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Help"
      Height          =   855
      Left            =   10920
      Picture         =   "ENSABVAC.frx":0BB2
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1155
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Print"
      Height          =   855
      Left            =   10920
      Picture         =   "ENSABVAC.frx":0EBC
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Guarda e Imprime Informe de Ensayo"
      Top             =   6615
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Setup"
      Height          =   855
      Left            =   10920
      Picture         =   "ENSABVAC.frx":1526
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Configuración - Formularios de Impresión Informes"
      Top             =   2100
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   555
      Left            =   10710
      Picture         =   "ENSABVAC.frx":1968
      Stretch         =   -1  'True
      Top             =   7630
      Width           =   4725
   End
   Begin VB.Label Label64 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      Height          =   225
      Left            =   210
      TabIndex        =   112
      Top             =   2100
      Width           =   1710
   End
End
Attribute VB_Name = "ENSABVAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CONSTARAMA#(5)
'
Private Sub BOTSEL_Click(Index As Integer)
   OBX$ = REGSEL$("MOBOMVAC")
   If Len(OBX$) > 4 Then
      Label66 = TRIM$(Mid$(OBX$, 1, 16))
   Text10 = Mid$(OBX$, 17, 32)
   Text9 = ""
   Text8 = TRIM$(FORMATNUM$(CVS(Mid$(OBX$, 65, 4)), "F8.1"))
   Text7 = TRIM$(FORMATNUM$(CVS(Mid$(OBX$, 69, 4)), "F8.0"))
   Text6 = TRIM$(FORMATNUM$(CVS(Mid$(OBX$, 73, 4)), "F8.0"))
   Text11 = TRIM$(FORMATNUM$(CVS(Mid$(OBX$, 77, 4)), "F8.1"))
   Text12 = TRIM$(FORMATNUM$(CVS(Mid$(OBX$, 81, 4)), "F8.0"))
   End If
End Sub

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Hide
End Sub

Private Sub Command13_Click()
  '
  Command13.Enabled = False
  Call DESHASFECHA(DES%, HAS%, PERIO$)
  If PERIO$ <> "" Then
    JJ& = 0
    FIN& = ULTREG&("ENSAVAC")
    For J& = 1 To FIN&
      X$ = REGLEIDO$("ENSAVAC", J&)
      FE1% = CVI(Mid$(X$, 5, 2))
      If FE1% >= DES% Then
        If FE1% <= HAS% Then
          JJ& = JJ& + 1
          Call GRAREG("LIENSAV", X$, JJ&)
        End If
      End If
    Next J&
    Call SETULTREG("LIENSAV", JJ&)
    Call CIERRARCH("LIENSAV")
    If JJ& > 0 Then
      Call HEADERS("LIENSAV", "")
      Call HEADERS("LIENSAV", "Periodo: " + PERIO$)
      Call FINAL("*LIENSAV")
    End If
  End If
  Command13.Enabled = True
  Hide
  '
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
3  X$ = REGSEL$("ENCENSVA")
   If Len(X$) > 4 Then
     REFE2$ = Mid$(X$, 65, 60)
     Screen.MousePointer = 11
     '
     For UI& = ULTREG&("PDOCLST") To 1 Step -1
       XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
       If InStr(XX$, REFE2$) > 4 Then
          DOCUNU& = CVS(Left$(XX$, 4))
          GoTo 5
       End If
    Next UI&
    '
    Screen.MousePointer = 1
    Call COMUNI("Documento no Registrado")
    GoTo 3
    '
5   For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    '
    REDOCU$ = REGACT$("PDOCLST")
    URE& = ULTREG&("PDOCSPL")
    LI& = 0: LS& = URE&
6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
    GoTo 8
    '
7   Screen.MousePointer = 1
    Call MENSERR(24, "Documento no Registrado")
    GoTo 3
    '
8   TPSP$ = ""
    Screen.MousePointer = 11
    While L& > 1
      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
      If DCC& < DOCUNU& Then GoTo 9
      L& = L& - 1
    Wend
    '
9   CAPAGINAS% = 0
    For KKL& = L& To URE&
      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
      End If
    Next KKL&
    '
95  SPOOLSTRING$ = TPSP$
    For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).Top = 0
      EDITFORM.Picture1(KKI%).Refresh
      EDITFORM.Picture1(KKI%).Height = 7170
      EDITFORM.Picture1(KKI%).Width = 11700
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    ALTUPAGINA = 7170
    TOPEPAGINA = 0
    PAGINACTIVA% = 0
    CAPAGINAS% = 0
    '
    FIN& = Len(TPSP$)
10  PPXX% = InStr(TPSP$, Chr$(255))
    If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call SHOWCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
    End If
    '
    Screen.MousePointer = 1
    PAGINACTIVA% = 0
    For KKI% = 0 To 5
       EDITFORM.Picture1(KKI%).Visible = False
    Next KKI%
    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
    EDITFORM.Picture1(PAGINACTIVA%).Refresh
    EDITFORM.VScroll1.Min = 0
    EDITFORM.VScroll1.Value = 0
    EDITFORM.VScroll1.Max = 1
    If ALTUPAGINA > EDITFORM.Frame1.Height Then
       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
    End If
    EDITFORM.Command3.Enabled = False
    EDITFORM.Command5.Enabled = False
    EDITFORM.Show 1
    '
    GoTo 3
    '
   End If
   Command2.Enabled = True
   '
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   Call PRIENSAVAC
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    CONFENSA.Show 1
    Command4.Enabled = True
End Sub

Private Sub Command6_Click()
   '
   Command6.Enabled = False
   VDEF$ = REGBLAN$("MOBOMVAC")
   '
10 Screen.MousePointer = 1
   OBX$ = OBJPLA$("NUEMOBOMVAC", VDEF$)
   If OBX$ = "" Then
     GoTo 99
   End If
   '
   VDEF$ = OBX$
   CODINS$ = UCase$(TRIM$(Left$(OBX$, 16)))
   If CODINS$ = "" Then
      Call MENSERR(24, "Falta Modelo de Bomba")
      GoTo 10
   End If
   Call REPLA(OBX$, CODINS$, 1, 16)
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("MOBOMVAC")
   For UI& = 1 To FIN&
      XX$ = REGLEIDO$("MOBOMVAC", UI&)
      If UCase$(TRIM$(Left$(XX$, 16))) = CODINS$ Then
         Call MENSERR(24, "Modelo Pre-Existente")
         GoTo 10
      End If
   Next UI&
   '
   MARCAMOT$ = TRIM$(Mid$(OBX$, 17, 32))
   If MARCAMOT$ = "" Then
      Call MENSERR(24, "Falta Marca del Motor.")
      GoTo 10
   End If
   '
   POTMOTOR = CVS(Mid$(OBX$, 65, 4))
   If POTMOTOR < 0.01 Then
      Call MENSERR(24, "Falta Potencia del Motor.")
      GoTo 10
   End If
   '
   POTMOTOR = CVS(Mid$(OBX$, 69, 4))
   If POTMOTOR < 1 Then
      Call MENSERR(24, "Falta Velocidad.")
      GoTo 10
   End If
   '
   POTMOTOR = CVS(Mid$(OBX$, 73, 4))
   If POTMOTOR < 1 Then
      Call MENSERR(24, "Falta Tensión.")
      GoTo 10
   End If
   '
   POTMOTOR = CVS(Mid$(OBX$, 77, 4))
   If POTMOTOR < 0.01 Then
      Call MENSERR(24, "Falta Intensidad.")
      GoTo 10
   End If
   '
   POTMOTOR = CVS(Mid$(OBX$, 81, 4))
   If POTMOTOR < 1 Then
      Call MENSERR(24, "Falta Frecuencia.")
      GoTo 10
   End If
   '
   Call CIERRARCH("MOBOMVAC")
   Call BLOQARCH("MOBOMVAC")
   Call REGAPP("MOBOMVAC", OBX$)
   Call CIERRARCH("MOBOMVAC")
   Call BAJALDISCO
   '
   Label66 = TRIM$(Mid$(OBX$, 1, 16))
   Text10 = Mid$(OBX$, 17, 32)
   Text9 = ""
   Text8 = TRIM$(FORMATNUM$(CVS(Mid$(OBX$, 65, 4)), "F8.1"))
   Text7 = TRIM$(FORMATNUM$(CVS(Mid$(OBX$, 69, 4)), "F8.0"))
   Text6 = TRIM$(FORMATNUM$(CVS(Mid$(OBX$, 73, 4)), "F8.0"))
   Text11 = TRIM$(FORMATNUM$(CVS(Mid$(OBX$, 77, 4)), "F8.1"))
   Text12 = TRIM$(FORMATNUM$(CVS(Mid$(OBX$, 81, 4)), "F8.0"))
   Screen.MousePointer = 1
   '
99 Command6.Enabled = True
   '
End Sub

Private Sub Form_Load()
    '
    Call CENTRAFORM(Me)
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    For KKI% = 0 To 4
       Label24(KKI%) = FORMATNUM$(Val(Text5), "F10.1")
    Next KKI%
    '
    For KKI% = 0 To 4
       Text18(KKI%) = FORMATNUM$(Val(Text2), "F10.1")
    Next KKI%
    '
    Label74(0) = FORMATNUM(30, "F10.1")
    Label74(1) = FORMATNUM(40, "F10.1")
    Label74(2) = FORMATNUM(60, "F10.1")
    Label74(3) = FORMATNUM(200, "F10.1")
    Label74(4) = FORMATNUM(460, "F10.1")
    '
    CONSTARAMA#(1) = 0.00082729
    CONSTARAMA#(2) = 0.0002458
    CONSTARAMA#(3) = 8.1326 / 100000
    CONSTARAMA#(4) = 2.5224 / 100000
    CONSTARAMA#(5) = 1.0428 / 100000
    '
    HOII% = HOY(HOS$)
    Label71 = HOS$
    Label73 = TRIM$(Str$(1 + ULTREG&("ENCENSVA")))
    '
    Show
    Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
End Sub

Private Sub Label24_Change(Index As Integer)
   Label67(Index) = CVAC$(Val(Label74(Index)), Val(Label24(Index)))
   Call CALCT(Index)
End Sub

Private Sub Label52_Change(Index As Integer)
   Call CALQ1(Index)
End Sub

Private Sub Label53_Change(Index As Integer)
   Call CALQG(Index)
End Sub

Sub CALQG(Index)
   On Error Resume Next
   QGE = Val(Label53(Index)) / (Val(Label67(Index)) * Val(Label68(Index)))
   On Error GoTo 0
   Label69(Index) = FORMATNUM$(QGE, "F10.1")
End Sub

Private Sub Label67_Change(Index As Integer)
   Call CALQG(Index)
End Sub

Private Sub Label68_Change(Index As Integer)
   Call CALQG(Index)
End Sub

Private Sub Label8_Change()
   For Index = 0 To 4
     Call CALQ1(Index)
   Next Index
End Sub

Private Sub Label9_Change()
   For Index = 0 To 4
     Call CALQAB(Index)
   Next Index
End Sub

Private Sub Text1_Change()
   Label8 = FORMATNUM$(Val(Text1) * 760 / 1013, "F8.1")
   Call CALDENSAIRE
End Sub

Private Sub Text1_GotFocus()
   Text1.SelStart = 0
   Text1.SelLength = Len(Text1)
End Sub

Private Sub Text1_LostFocus()
   Text1 = FORMATNUM(Text1, "F8.0")
End Sub

Private Sub Text10_GotFocus()
   Text10.SelStart = 0
   Text10.SelLength = Len(Text10)
End Sub

Private Sub Text11_GotFocus()
   Text11.SelStart = 0
   Text11.SelLength = Len(Text11)
End Sub

Private Sub Text12_GotFocus()
   Text12.SelStart = 0
   Text12.SelLength = Len(Text12)
End Sub

Private Sub Text13_GotFocus(Index As Integer)
   Text13(Index).SelStart = 0
   Text13(Index).SelLength = Len(Text13(Index))
End Sub

Private Sub Text13_LostFocus(Index As Integer)
   Text13(Index) = FORMATNUM$(Val(Text13(Index)), "F10.1")
End Sub

Private Sub Text14_Change(Index As Integer)
   Label51(Index) = ""
   RAMAUTI& = Val(Text14(Index))
   If RAMAUTI& >= 1 Then
     If RAMAUTI& <= 5 Then
      CONRAMA# = 100000 * CONSTARAMA#(RAMAUTI&)
      Label51(Index) = FORMATNUM$(CONRAMA#, "F7.4") + "E-5"
     End If
   End If
   Call CALQAB(Index)
End Sub

Private Sub Text14_GotFocus(Index As Integer)
   Text14(Index).SelStart = 0
   Text14(Index).SelLength = Len(Text14(Index))
End Sub

Private Sub Text14_LostFocus(Index As Integer)
   Text14(Index) = FORMATNUM$(Val(Text14(Index)), "F6.0")
End Sub

Private Sub Text15_Change(Index As Integer)
   Call CALQAB(Index)
End Sub

Sub CALQAB(Index)
   Label52(Index) = ""
   RAMAUTI& = Val(Text14(Index))
   If RAMAUTI& >= 1 Then
     If RAMAUTI& <= 5 Then
       QAB = CONSTARAMA#(RAMAUTI&) * Sqr((2 * Val(Text15(Index)) * 9.808) / Val(Label9))
       QAB = QAB * 60 * 1000
       Label52(Index) = FORMATNUM$(QAB, "F10.1")
     End If
   End If
End Sub

Sub CALQ1(Index)
   Label53(Index) = ""
   On Error Resume Next
   Q1 = Val(Label52(Index)) * (Val(Label8) / Val(Label74(Index)))
   Q1 = Q1 * ((273.15 + Val(Text18(Index))) / (273.15 + Val(Text2)))
   Label53(Index) = FORMATNUM$(Q1, "F10.1")
   On Error GoTo 0
End Sub

Private Sub Text15_GotFocus(Index As Integer)
   Text15(Index).SelStart = 0
   Text15(Index).SelLength = Len(Text15(Index))
End Sub

Private Sub Text15_LostFocus(Index As Integer)
   Text15(Index) = FORMATNUM$(Val(Text15(Index)), "F10.0")
End Sub

Private Sub Text16_GotFocus(Index As Integer)
   Text16(Index).SelStart = 0
   Text16(Index).SelLength = Len(Text16(Index))
End Sub

Private Sub Text16_LostFocus(Index As Integer)
   Text16(Index) = FORMATNUM$(Val(Text16(Index)), "F10.1")
End Sub

Private Sub label74_Change(Index As Integer)
   Label25(Index) = FORMATNUM$(760 - Val(Label74(Index)), "F10.1")
   Call CALQ1(Index)
   Label67(Index) = CVAC$(Val(Label74(Index)), Val(Label24(Index)))
End Sub

Private Sub Text18_Change(Index As Integer)
   Call CALQ1(Index)
End Sub

Private Sub Text18_GotFocus(Index As Integer)
   Text18(Index).SelStart = 0
   Text18(Index).SelLength = Len(Text18(Index))
End Sub

Private Sub Text18_LostFocus(Index As Integer)
   Text18(Index) = FORMATNUM$(Val(Text18(Index)), "F10.1")
End Sub

Private Sub Text2_Change()
   Call CALDENSAIRE
   '
   For KKI% = 0 To 4
      Text18(KKI%) = FORMATNUM$(Val(Text2), "F10.1")
   Next KKI%
   '
   For Index = 0 To 4
     Call CALQ1(Index)
     Call CALCT(Index)
   Next Index
End Sub
'
Sub CALCT(Index)
   VACT = 1 + (0.66 * (Val(Text2) - 20) / (Val(Label24(Index)) - 273.15))
   Label68(Index) = FORMATNUM(VACT, "F10.3")
End Sub

Private Sub Text2_gotfocus()
   Text2.SelStart = 0
   Text2.SelLength = Len(Text2)
End Sub

Private Sub Text2_LostFocus()
   Text2 = FORMATNUM(Text2, "F8.1")
End Sub

Private Sub Text22_GotFocus()
   Text22.SelStart = 0
   Text22.SelLength = Len(Text22)
End Sub

Private Sub Text23_GotFocus()
   Text23.SelStart = 0
   Text23.SelLength = Len(Text23)
End Sub

Private Sub Text3_Change()
   Call CALDENSAIRE
End Sub

Private Sub Text3_GotFocus()
   Text3.SelStart = 0
   Text3.SelLength = Len(Text3)
End Sub

Private Sub Text3_LostFocus()
   Text3 = FORMATNUM(Text3, "F8.0")
End Sub

Private Sub Text4_GotFocus()
   Text4.SelStart = 0
   Text4.SelLength = Len(Text4)
End Sub

Private Sub Text5_Change()
   For KKI% = 0 To 4
      Label24(KKI%) = FORMATNUM$(Val(Text5), "F10.1")
   Next KKI%
End Sub

Private Sub Text5_GotFocus()
   Text5.SelStart = 0
   Text5.SelLength = Len(Text5)
End Sub

Private Sub Text5_LostFocus()
   Text5 = FORMATNUM(Text5, "F8.1")
End Sub

Sub CALDENSAIRE()
   DENSI = 1.2049 * (Val(Label8) / 760) * ((273.15 + 20) / (273.15 + Val(Text2)))
   Label9 = FORMATNUM$(DENSI, "F8.4")
End Sub

Function CVAC$(PASP&, TAGUA&)
   '
   Static CTX%
   Static CV(32, 4)
   '
   If CTX% = 0 Then
     CTX% = 1
     TA& = 10: CV(TA& - 10, 0) = 1.16
               CV(TA& - 10, 1) = 1.09
               CV(TA& - 10, 2) = 1.05
               CV(TA& - 10, 3) = 1.01
               CV(TA& - 10, 4) = 1#
     TA& = 11: CV(TA& - 10, 0) = 1.13
               CV(TA& - 10, 1) = 1.08
               CV(TA& - 10, 2) = 1.04
               CV(TA& - 10, 3) = 1.01
               CV(TA& - 10, 4) = 1#
     TA& = 12: CV(TA& - 10, 0) = 1.1
               CV(TA& - 10, 1) = 1.06
               CV(TA& - 10, 2) = 1.03
               CV(TA& - 10, 3) = 1.01
               CV(TA& - 10, 4) = 1#
     TA& = 13: CV(TA& - 10, 0) = 1.07
               CV(TA& - 10, 1) = 1.04
               CV(TA& - 10, 2) = 1.02
               CV(TA& - 10, 3) = 1#
               CV(TA& - 10, 4) = 1#
     TA& = 14: CV(TA& - 10, 0) = 1.04
               CV(TA& - 10, 1) = 1.02
               CV(TA& - 10, 2) = 1.01
               CV(TA& - 10, 3) = 1#
               CV(TA& - 10, 4) = 1#
     TA& = 15: CV(TA& - 10, 0) = 1#
               CV(TA& - 10, 1) = 1#
               CV(TA& - 10, 2) = 1#
               CV(TA& - 10, 3) = 1#
               CV(TA& - 10, 4) = 1#
     TA& = 16: CV(TA& - 10, 0) = 0.96
               CV(TA& - 10, 1) = 0.98
               CV(TA& - 10, 2) = 0.99
               CV(TA& - 10, 3) = 1#
               CV(TA& - 10, 4) = 1#
     TA& = 17: CV(TA& - 10, 0) = 0.92
               CV(TA& - 10, 1) = 0.96
               CV(TA& - 10, 2) = 0.98
               CV(TA& - 10, 3) = 0.99
               CV(TA& - 10, 4) = 1#
     TA& = 18: CV(TA& - 10, 0) = 0.88
               CV(TA& - 10, 1) = 0.93
               CV(TA& - 10, 2) = 0.96
               CV(TA& - 10, 3) = 0.99
               CV(TA& - 10, 4) = 1#
     TA& = 19: CV(TA& - 10, 0) = 0.84
               CV(TA& - 10, 1) = 0.9
               CV(TA& - 10, 2) = 0.95
               CV(TA& - 10, 3) = 0.99
               CV(TA& - 10, 4) = 1#
     TA& = 20: CV(TA& - 10, 0) = 0.79
               CV(TA& - 10, 1) = 0.88
               CV(TA& - 10, 2) = 0.93
               CV(TA& - 10, 3) = 0.99
               CV(TA& - 10, 4) = 1#
     TA& = 21: CV(TA& - 10, 0) = -0.76
               CV(TA& - 10, 1) = 0.85
               CV(TA& - 10, 2) = 0.92
               CV(TA& - 10, 3) = 0.98
               CV(TA& - 10, 4) = 0.99
     TA& = 22: CV(TA& - 10, 0) = -0.71
               CV(TA& - 10, 1) = 0.82
               CV(TA& - 10, 2) = 0.9
               CV(TA& - 10, 3) = 0.98
               CV(TA& - 10, 4) = 0.99
     TA& = 23: CV(TA& - 10, 0) = -0.67
               CV(TA& - 10, 1) = 0.79
               CV(TA& - 10, 2) = 0.89
               CV(TA& - 10, 3) = 0.98
               CV(TA& - 10, 4) = 0.99
     TA& = 24: CV(TA& - 10, 0) = -0.64
               CV(TA& - 10, 1) = 0.75
               CV(TA& - 10, 2) = 0.87
               CV(TA& - 10, 3) = 0.97
               CV(TA& - 10, 4) = 0.99
     TA& = 25: CV(TA& - 10, 0) = -0.58
               CV(TA& - 10, 1) = 0.72
               CV(TA& - 10, 2) = 0.85
               CV(TA& - 10, 3) = 0.97
               CV(TA& - 10, 4) = 0.99
     TA& = 26: CV(TA& - 10, 0) = -0.53
               CV(TA& - 10, 1) = 0.68
               CV(TA& - 10, 2) = 0.83
               CV(TA& - 10, 3) = 0.96
               CV(TA& - 10, 4) = 0.99
     TA& = 27: CV(TA& - 10, 0) = -0.47
               CV(TA& - 10, 1) = 0.64
               CV(TA& - 10, 2) = 0.81
               CV(TA& - 10, 3) = 0.96
               CV(TA& - 10, 4) = 0.99
     TA& = 28: CV(TA& - 10, 0) = -0.41
               CV(TA& - 10, 1) = 0.59
               CV(TA& - 10, 2) = 0.79
               CV(TA& - 10, 3) = 0.96
               CV(TA& - 10, 4) = 0.99
     TA& = 29: CV(TA& - 10, 0) = -0.36
               CV(TA& - 10, 1) = -0.55
               CV(TA& - 10, 2) = 0.76
               CV(TA& - 10, 3) = 0.95
               CV(TA& - 10, 4) = 0.99
     TA& = 30: CV(TA& - 10, 0) = -0.3
               CV(TA& - 10, 1) = -0.5
               CV(TA& - 10, 2) = 0.74
               CV(TA& - 10, 3) = 0.95
               CV(TA& - 10, 4) = 0.98
   End If
   '
   If TAGUA& < 10 Or TAGUA& > 30 Then
      'Call MENSERR(24, "Temperatura de Líquido de Sello\Fuera de Rango")
      CVA = 0
      GoTo 90
   End If
   '
   If PASP& = 30 Then
       CVA = CV(TAGUA& - 10, 0)
     ElseIf PASP& = 40 Then
       CVA = CV(TAGUA& - 10, 1)
     ElseIf PASP& = 60 Then
       CVA = CV(TAGUA& - 10, 2)
     ElseIf PASP& = 200 Then
       CVA = CV(TAGUA& - 10, 3)
     ElseIf PASP& = 460 Then
       CVA = CV(TAGUA& - 10, 4)
     Else
       'Call MENSERR(24, "Presión de Aspiración\no Figura en Tablas")
       CVA = 0
   End If
   '
90 CVAC$ = TRIM$(FORMATNUM$(Abs(CVA), "F10.3"))
   If CVA < 0 Then CVAC$ = CVAC$ + "(cav)"
   While Len(CVAC$) < 10: CVAC$ = " " + CVAC$: Wend
   '
End Function

Sub PRIENSAVAC()
   '
   If TRIM$(Label66) = "" Then
        MERRO$ = "Falta Modelo de Bomba": GoTo 90
     ElseIf TRIM$(Text23) = "" Then
        MERRO$ = "Falta Número de Bomba": GoTo 90
     ElseIf TRIM$(Text10) = "" Then
        MERRO$ = "Falta Marca de Motor": GoTo 90
     ElseIf TRIM$(Text9) = "" Then
        MERRO$ = "Falta Número de Motor": GoTo 90
     ElseIf TRIM$(Text8) = "" Then
        MERRO$ = "Falta Potencia de Motor": GoTo 90
     ElseIf TRIM$(Text7) = "" Then
        MERRO$ = "Falta Velocidad de Motor": GoTo 90
     ElseIf TRIM$(Text6) = "" Then
        MERRO$ = "Falta Tensión de Motor": GoTo 90
     ElseIf TRIM$(Text11) = "" Then
        MERRO$ = "Falta Intensidad de Motor": GoTo 90
     ElseIf TRIM$(Text12) = "" Then
        MERRO$ = "Falta Frecuencia de Motor": GoTo 90
        '
     ElseIf Val(Text1) < 980 Or Val(Text1) > 1030 Then
        MERRO$ = "Presión Ambiente Fuera de Rango": GoTo 90
     ElseIf Val(Text2) < 1 Or Val(Text2) > 60 Then
        MERRO$ = "Temperatura Ambiente Fuera de Rango": GoTo 90
     ElseIf Val(Text3) < 20 Or Val(Text3) > 99 Then
        MERRO$ = "Humedad Ambiente Fuera de Rango": GoTo 90
     ElseIf Val(Text5) < 1 Or Val(Text5) > 60 Then
        MERRO$ = "Temperatura Liquido de Sello Fuera de Rango": GoTo 90
        '
     ElseIf TRIM$(Text4) = "" Then
        MERRO$ = "Falta Líquido de Sello": GoTo 90
   End If
   '
   HAYINFO% = 0
   For KKI% = 0 To 4
       If Val(Text14(KKI%)) < 0 Or Val(Text14(KKI%)) > 5 Then
          MERRO$ = "Ramal Incorrecto": GoTo 90
       End If
       If Val(Text14(KKI%)) > 0 Then
         HAYINFO% = 1
         If Val(Text15(KKI%)) < 1 Then
             MERRO$ = "Falta Delta Pab en Punto " + Str$(KKI% + 1): GoTo 90
           ElseIf Val(Text18(KKI%)) < 1 Then
             MERRO$ = "Falta t1 en Punto " + Str$(KKI% + 1): GoTo 90
         End If
       End If
   Next KKI%
   '
   If HAYINFO% < 1 Then
      MERRO$ = "No hay Información Cargada\Puede Imprimir el Formulario en Blanco"
      If COMALTER%(MERRO$ + "\\Cancelar\Formulario en Blanco") = 2 Then GoTo 10
      Exit Sub
   End If
   '
   Call BLOQARCH("ENCENSVA")
   TTXX$ = REGBLAN$("ENCENSVA")
   REFE1$ = Space$(62)
   Call REPLA(REFE1$, Label66, 1, 12)
   Call REPLA(REFE1$, Text23, 13, 12)
   Call REPLA(REFE1$, "Motor " + TRIM$(Text9), 26, 20)
   Call REPLA(REFE1$, Text10, 46, 15)
   '
   REFE2$ = AJUSTI$("EBV " + TRIM$(Label66) + " " + TRIM$(Text23) + " Mot." + TRIM$(Text9) + " " + TRIM$(Text10), 60)
   IENSA% = 1 + ULTREG&("ENCENSVA")
   '
   Call REPLA(TTXX$, MKI$(IENSA%), 1, 2)
   Call REPLA(TTXX$, REFE1$, 3, 62)
   Call REPLA(TTXX$, REFE2$, 65, 60)
   HOII% = HOY(HOS$)
   Call REPLA(TTXX$, MKI$(HOII%), 125, 2)
   Call REPLA(TTXX$, MKI$(USNBR% Mod 100), 127, 2)
   Call REGAPP("ENCENSVA", TTXX$)
   Call CIERRARCH("ENCENSVA")
   Call BAJALDISCO
   '\\\OT-06-0004-WAR-06/01/06///
    'guardo registro en ENSAVAC.rfs para listado.
    NUEN = Val(Label73)
    FEN% = FECHANUM(Label71)
    MODEL$ = Label66
    NUMODEL$ = Text23.TEXT
    MARMOTOR$ = Text10.TEXT
    NUMOTOR$ = Text9.TEXT
    POMOTOR = Val(Text8.TEXT)
    IP1 = Val(Text16(0).TEXT)
    CP1 = Val(Label69(0))
    IP2 = Val(Text16(1).TEXT)
    CP2 = Val(Label69(1))
    IP3 = Val(Text16(2).TEXT)
    CP3 = Val(Label69(2))
    IP4 = Val(Text16(3).TEXT)
    CP4 = Val(Label69(3))
    IP5 = Val(Text16(4))
    CP5 = Val(Label69(4))
    '
    XZ$ = REGBLAN$("ENSAVAC")
    Call REPLA(XZ$, MKS$(NUEN), 1, 4)
    Call REPLA(XZ$, MKI$(FEN%), 5, 2)
    Call REPLA(XZ$, MODEL$, 7, 12)
    Call REPLA(XZ$, NUMODEL$, 19, 12)
    Call REPLA(XZ$, MARMOTOR$, 31, 12)
    Call REPLA(XZ$, NUMOTOR$, 43, 12)
    Call REPLA(XZ$, MKS$(POMOTOR), 55, 4)
    Call REPLA(XZ$, MKS$(IP1), 80, 4)
    Call REPLA(XZ$, MKS$(CP1), 84, 4)
    Call REPLA(XZ$, MKS$(IP2), 96, 4)
    Call REPLA(XZ$, MKS$(CP2), 100, 4)
    Call REPLA(XZ$, MKS$(IP3), 112, 4)
    Call REPLA(XZ$, MKS$(CP3), 116, 4)
    Call REPLA(XZ$, MKS$(IP4), 128, 4)
    Call REPLA(XZ$, MKS$(CP4), 132, 4)
    Call REPLA(XZ$, MKS$(IP5), 144, 4)
    Call REPLA(XZ$, MKS$(CP5), 148, 4)
    Call REGAPP("ENSAVAC", XZ$)
    Call CIERRARCH("ENSAVAC")
    '
   '\\\OT-06-0004-WAR-FIN///
10 FORIPRE$ = TRIM$(CONTROL$("FORENSA", "FORENSA0"))
   If FORIPRE$ = "" Then
     Call MENSERR(24, "Falta Definir Formulario de Impresión")
     CONFENSA.Show 1
     Exit Sub
   End If
   '
   Screen.MousePointer = 11
   Call FRATEXTO(Text25, 64)
   JJ& = 0
   For U& = 1 To CARETEX%
     Z$ = RETEX$(U&)
     JJ& = JJ& + 1
     Call GRAREG("!OBSERVOF", Z$, JJ&)
   Next U&
   CARE& = CARETEX%
   Call SETULTREG("!OBSERVOF", CARE&)
   Call CIERRARCH("!OBSERVOF")
   '
   FECHDOC$ = Label71
   NUMEDOC$ = TRIM$(Text23)
   TIPODOC$ = REFE2$
   If TRIM$(TIPODOC$) = "" Then
     TIPODOC$ = "Formulario Carga Ensayo"
   End If
   '
   CODPARAM$(1) = "COD-CONJ"
   CODPARAM$(2) = "REF-MAT1"
   CODPARAM$(3) = "MEDIDA-1"
   CODPARAM$(4) = "MEDIDA-2"
   CODPARAM$(5) = "MEDIDA-3"
   CODPARAM$(6) = "N-PLANO"
   CODPARAM$(7) = "R-PLANO"
   CODPARAM$(8) = "F-PLANO"
   CODPARAM$(9) = "Z-PLANO"
   CODPARAM$(10) = "N-HPROF"
   CODPARAM$(11) = "N-HINSP"
   CODPARAM$(12) = "N-ORIGIN"
   CODPARAM$(13) = "N-OCOMPR"
   CODPARAM$(14) = "N-AUTENT"
   CODPARAM$(15) = "N-REMITO"
   CODPARAM$(16) = "N-BOLREC"
   CAPARDOC% = 16
   '
   DOCPARAM$(1) = TRIM$(Label66)
   DOCPARAM$(2) = TRIM$(Text22)
   DOCPARAM$(3) = Label8
   DOCPARAM$(4) = Text1
   DOCPARAM$(5) = Text2
   DOCPARAM$(6) = Text3
   DOCPARAM$(7) = Label9
   DOCPARAM$(8) = Text4
   DOCPARAM$(9) = Text5
   DOCPARAM$(10) = Text10
   DOCPARAM$(11) = Text9
   DOCPARAM$(12) = Text8
   DOCPARAM$(13) = Text7
   DOCPARAM$(14) = Text6
   DOCPARAM$(15) = Text11
   DOCPARAM$(16) = Text12
   '
   CODTABU1$(1) = "COD-MAT"
   CODTABU1$(2) = "DES-MAT"
   CODTABU1$(3) = "PRE-LIST"
   CODTABU1$(4) = "NET-ITEM"
   CODTABU1$(5) = "IMP-TOTA"
   CACOLTAB1% = 5
   '
   CAITAB1% = 0
   For U& = 1 To 5
     TETABU1$(U&, 1) = Label74(U& - 1)
     TETABU1$(U&, 2) = Label25(U& - 1)
     TETABU1$(U&, 4) = Text16(U& - 1)
     TETABU1$(U&, 6) = Text13(U& - 1)
     TETABU1$(U&, 7) = Label24(U& - 1)
     TETABU1$(U&, 9) = Text14(U& - 1)
     TETABU1$(U&, 10) = Label51(U& - 1)
     TETABU1$(U&, 11) = Text15(U& - 1)
     TETABU1$(U&, 12) = Text18(U& - 1)
     TETABU1$(U&, 13) = Label52(U& - 1)
     TETABU1$(U&, 14) = Label53(U& - 1)
     TETABU1$(U&, 15) = Label67(U& - 1)
     TETABU1$(U&, 16) = Label68(U& - 1)
     TETABU1$(U&, 17) = Label69(U& - 1)
   Next U&
   CAITAB1% = 17
   '
   NOCONFIRM% = 1
   Call PRINTDOC("@" + FORIPRE$)
   Screen.MousePointer = 1
   '
   HOII% = HOY(HOS$)
   Label71 = HOS$
   Label73 = TRIM$(Str$(1 + ULTREG&("ENCENSVA")))
   Text9 = ""
   Text22 = ""
   Text9 = ""
   '
   For KKI% = 0 To 4
      Text13(KKI%) = ""
      Text14(KKI%) = ""
      Text15(KKI%) = ""
   Next KKI%
   Exit Sub
   '
90 Call MENSERR(24, MERRO$)
   '
End Sub

Private Sub Text6_GotFocus()
   Text6.SelStart = 0
   Text6.SelLength = Len(Text6)
End Sub

Private Sub Text7_GotFocus()
   Text7.SelStart = 0
   Text7.SelLength = Len(Text7)
End Sub

Private Sub Text8_GotFocus()
   Text8.SelStart = 0
   Text8.SelLength = Len(Text8)
End Sub

Private Sub Text9_GotFocus()
   Text9.SelStart = 0
   Text9.SelLength = Len(Text9)
End Sub
