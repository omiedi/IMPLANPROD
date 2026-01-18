VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form CONFREPA09 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reparaciones - Configuración de Funcionamiento"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   10665
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   10665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame15 
      BackColor       =   &H00CDDADA&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   5460
      TabIndex        =   116
      Top             =   4400
      Width           =   2580
      Begin VB.TextBox Text8 
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
         Height          =   280
         Left            =   1560
         TabIndex        =   117
         Text            =   " "
         Top             =   160
         Width           =   915
      End
      Begin VB.Label FORMUDES 
         BackStyle       =   0  'Transparent
         Caption         =   "N° de Inicialización:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   19
         Left            =   120
         TabIndex        =   118
         Top             =   240
         Width           =   1605
      End
   End
   Begin VB.Frame Frame14 
      BackColor       =   &H00CDDADA&
      Caption         =   "Ctas.Contables de Repuestos y Contrapartida"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1250
      Left            =   5400
      TabIndex        =   100
      Top             =   4920
      Width           =   5175
      Begin VB.CommandButton BOTSEL 
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
         Left            =   1440
         TabIndex        =   104
         Top             =   840
         Width           =   175
      End
      Begin VB.TextBox CTACBLENUM 
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
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   103
         Text            =   " "
         Top             =   840
         Width           =   1035
      End
      Begin VB.TextBox CTACBLENUM 
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
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   102
         Text            =   " "
         Top             =   360
         Width           =   1035
      End
      Begin VB.CommandButton BOTSEL 
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
         Left            =   1440
         TabIndex        =   101
         Top             =   375
         Width           =   175
      End
      Begin VB.Label FORMUDES 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Repuestos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   17
         Left            =   0
         TabIndex        =   109
         Top             =   640
         Width           =   880
      End
      Begin VB.Label FORMUDES 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Contrapartida"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   16
         Left            =   -240
         TabIndex        =   108
         Top             =   165
         Width           =   1320
      End
      Begin VB.Label FORMUDES 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "M.Obra"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   15
         Left            =   0
         TabIndex        =   107
         Top             =   4440
         Width           =   600
      End
      Begin VB.Label DESCRIECO 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   1680
         TabIndex        =   106
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label DESCRIECO 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   1680
         TabIndex        =   105
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1845
      Index           =   0
      Left            =   9765
      ScaleHeight     =   1815
      ScaleWidth      =   2790
      TabIndex        =   62
      Top             =   0
      Width           =   2820
      Begin VB.CommandButton command1 
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
         Height          =   650
         Left            =   120
         Picture         =   "CONFREPA09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   30
         Width           =   650
      End
      Begin VB.PictureBox Picture8 
         BackColor       =   &H00808080&
         Height          =   495
         Index           =   3
         Left            =   1680
         ScaleHeight     =   435
         ScaleWidth      =   915
         TabIndex        =   91
         Top             =   720
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.PictureBox Picture8 
         BackColor       =   &H00808080&
         Height          =   495
         Index           =   2
         Left            =   1440
         ScaleHeight     =   435
         ScaleWidth      =   915
         TabIndex        =   90
         Top             =   0
         Width           =   975
      End
      Begin VB.PictureBox Picture8 
         BackColor       =   &H00808080&
         Height          =   495
         Index           =   1
         Left            =   1680
         ScaleHeight     =   435
         ScaleWidth      =   915
         TabIndex        =   89
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
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
         Picture         =   "CONFREPA09.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   960
         Width           =   650
      End
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00CDDADA&
      Caption         =   "ACCESO A TABLAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   7920
      TabIndex        =   76
      Top             =   1200
      Width           =   1750
      Begin VB.OptionButton Option3 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Tipo de Falla"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   50
         TabIndex        =   78
         ToolTipText     =   "Doble Click Para Abrir Tabla"
         Top             =   480
         Width           =   1640
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Causa de Falla"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   40
         TabIndex        =   77
         ToolTipText     =   "Doble Click Para Abrir Tabla"
         Top             =   240
         Width           =   1640
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00CDDADA&
      Caption         =   "DEPOSITO DE CONSUMO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7440
      TabIndex        =   65
      Top             =   6240
      Width           =   3130
      Begin VB.TextBox Text6 
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
         Left            =   780
         MaxLength       =   4
         TabIndex        =   72
         Top             =   600
         Width           =   420
      End
      Begin VB.CommandButton Command6 
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
         Height          =   300
         Left            =   1185
         Picture         =   "CONFREPA09.frx":0454
         TabIndex        =   71
         Top             =   600
         Width           =   175
      End
      Begin VB.CommandButton command15 
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
         Height          =   300
         Left            =   1185
         Picture         =   "CONFREPA09.frx":075E
         TabIndex        =   67
         Top             =   240
         Width           =   175
      End
      Begin VB.TextBox Text5 
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
         Left            =   780
         MaxLength       =   4
         TabIndex        =   66
         Top             =   240
         Width           =   420
      End
      Begin VB.Label FORMUDES 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Destino"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   11
         Left            =   50
         TabIndex        =   75
         Top             =   645
         Width           =   645
      End
      Begin VB.Label FORMUDES 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Origen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   10
         Left            =   50
         TabIndex        =   74
         Top             =   285
         Width           =   645
      End
      Begin VB.Label Label7 
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
         Height          =   300
         Left            =   1320
         TabIndex        =   73
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label5 
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
         Height          =   300
         Left            =   1320
         TabIndex        =   68
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00CDDADA&
      Caption         =   "COSTO DE REPUESTOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2760
      TabIndex        =   56
      Top             =   6240
      Width           =   4575
      Begin VB.CommandButton Command4 
         Caption         =   "."
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
         Left            =   2985
         TabIndex        =   59
         Top             =   280
         Width           =   175
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Tabla de Costos Calculados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   57
         Top             =   320
         Width           =   1455
      End
      Begin VB.Label Label4 
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
         Left            =   2520
         TabIndex        =   61
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Left            =   2520
         TabIndex        =   60
         Top             =   285
         Width           =   450
      End
      Begin VB.Label Label2 
         BackColor       =   &H00CDDADA&
         Caption         =   "Lista de Precios Nro:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1800
         TabIndex        =   58
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00CDDADA&
      Caption         =   "INGRESO DE EQUIPOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2520
      Left            =   120
      TabIndex        =   47
      Top             =   105
      Width           =   5160
      Begin VB.CommandButton BOTOPENFORM 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   620
         Index           =   7
         Left            =   4500
         Picture         =   "CONFREPA09.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   113
         Top             =   840
         Width           =   540
      End
      Begin VB.CommandButton BOTSELFORM 
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
         Index           =   7
         Left            =   4110
         TabIndex        =   112
         Top             =   840
         Width           =   175
      End
      Begin VB.TextBox FORMUNAME 
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
         Height          =   285
         Index           =   7
         Left            =   2850
         TabIndex        =   111
         Text            =   " "
         Top             =   840
         Width           =   1275
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00CDDADA&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   900
         Index           =   7
         Left            =   120
         TabIndex        =   92
         Top             =   1500
         Width           =   4950
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   840
            TabIndex        =   98
            Text            =   "Combo1"
            Top             =   405
            Width           =   1635
         End
         Begin VB.CommandButton BOTOPENFORM 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   620
            Index           =   6
            Left            =   4400
            Picture         =   "CONFREPA09.frx":0D72
            Style           =   1  'Graphical
            TabIndex        =   95
            Top             =   30
            Width           =   540
         End
         Begin VB.CommandButton BOTSELFORM 
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
            Index           =   6
            Left            =   2280
            TabIndex        =   94
            Top             =   30
            Width           =   175
         End
         Begin VB.TextBox FORMUNAME 
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
            Height          =   285
            Index           =   6
            Left            =   1200
            TabIndex        =   93
            Text            =   " "
            Top             =   30
            Width           =   1155
         End
         Begin VB.Label FORMUDES 
            BackStyle       =   0  'Transparent
            Caption         =   "Impresora"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   14
            Left            =   90
            TabIndex        =   99
            Top             =   495
            Width           =   1125
         End
         Begin VB.Label FORMUECO 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Height          =   645
            Index           =   6
            Left            =   2520
            TabIndex        =   97
            Top             =   30
            Width           =   1695
         End
         Begin VB.Label FORMUDES 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Alta de Equipo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   13
            Left            =   0
            TabIndex        =   96
            Top             =   105
            Width           =   1125
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00CDDADA&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   735
         Index           =   5
         Left            =   105
         TabIndex        =   48
         Top             =   195
         Width           =   4950
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   0
            Top             =   120
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.TextBox FORMUNAME 
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
            Height          =   285
            Index           =   5
            Left            =   2745
            TabIndex        =   51
            Text            =   " "
            Top             =   0
            Width           =   1275
         End
         Begin VB.CommandButton BOTSELFORM 
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
            Index           =   5
            Left            =   4005
            TabIndex        =   50
            Top             =   0
            Width           =   175
         End
         Begin VB.CommandButton BOTOPENFORM 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   620
            Index           =   5
            Left            =   4400
            Picture         =   "CONFREPA09.frx":107C
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   -15
            Width           =   540
         End
         Begin VB.Label FORMUDES 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nota de Recepción"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   9
            Left            =   330
            TabIndex        =   53
            Top             =   60
            Width           =   2325
         End
         Begin VB.Label FORMUECO 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   135
            TabIndex        =   52
            Top             =   315
            Width           =   4095
         End
      End
      Begin VB.Label FORMUECO 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   240
         TabIndex        =   115
         Top             =   1170
         Width           =   4095
      End
      Begin VB.Label FORMUDES 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nota de Recepción al Cliente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   18
         Left            =   435
         TabIndex        =   114
         Top             =   900
         Width           =   2325
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00CDDADA&
      Caption         =   "CODIGO GENERICO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   5460
      TabIndex        =   46
      Top             =   1200
      Width           =   2415
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
         Left            =   2120
         TabIndex        =   55
         Top             =   300
         Width           =   175
      End
      Begin VB.Label Label1 
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
         Left            =   120
         TabIndex        =   54
         Top             =   300
         Width           =   2010
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00CDDADA&
      Caption         =   "REPARACIONES TIPICAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   5460
      TabIndex        =   42
      Top             =   105
      Width           =   4215
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   3240
         TabIndex        =   43
         Top             =   300
         Width           =   3300
         Begin VB.CommandButton Command8 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "CONFREPA09.frx":1386
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Reparaciones Normalizadas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Actualizar B.Datos de Reparaciones Normales"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   735
            TabIndex        =   45
            Top             =   105
            Width           =   2220
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00CDDADA&
      Caption         =   "CONDICIONES COMERCIALES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2760
      TabIndex        =   37
      Top             =   7320
      Width           =   7820
      Begin VB.TextBox Text4 
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
         Left            =   1470
         TabIndex        =   39
         Top             =   360
         Width           =   6000
      End
      Begin VB.TextBox Text3 
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
         Left            =   1470
         TabIndex        =   38
         Top             =   975
         Width           =   6000
      End
      Begin VB.Label FORMUDES 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Plazo de Entrega"
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
         Index           =   8
         Left            =   -960
         TabIndex        =   41
         Top             =   465
         Width           =   2325
      End
      Begin VB.Label FORMUDES 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Forma de Pago"
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
         Index           =   7
         Left            =   -360
         TabIndex        =   40
         Top             =   1065
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "COSTO M. OBRA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   105
      TabIndex        =   13
      Top             =   6240
      Width           =   2550
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "M.Obra s/ Lista de Costo por Artículo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   360
         TabIndex        =   110
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox Text2 
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
         Left            =   1245
         TabIndex        =   17
         Top             =   1815
         Width           =   1065
      End
      Begin VB.TextBox Text1 
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
         Left            =   1245
         TabIndex        =   15
         Top             =   585
         Width           =   1065
      End
      Begin VB.Label FORMUDES 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Coefic. Presupuesto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   5
         Left            =   -15
         TabIndex        =   16
         Top             =   1680
         Width           =   1170
      End
      Begin VB.Label FORMUDES 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "C-H Mano de Obra"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   4
         Left            =   450
         TabIndex        =   14
         Top             =   465
         Width           =   705
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Caption         =   "FORMULARIOS COTIZACION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3525
      Left            =   105
      TabIndex        =   0
      Top             =   2640
      Width           =   5160
      Begin VB.Frame Frame13 
         BackColor       =   &H00CDDADA&
         Caption         =   "MEDIO COBRANZA DEFAULT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   2640
         TabIndex        =   85
         ToolTipText     =   "M.Cobranza x Defecto de Presupuesto de Reparación"
         Top             =   2685
         Width           =   2400
         Begin VB.TextBox Text7 
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
            Height          =   285
            Left            =   150
            MaxLength       =   4
            TabIndex        =   88
            Top             =   300
            Width           =   300
         End
         Begin VB.CommandButton Command7 
            Caption         =   "."
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
            Left            =   480
            TabIndex        =   86
            ToolTipText     =   "M.Cobranza x Defecto de Presupuesto de Reparación"
            Top             =   300
            Width           =   175
         End
         Begin VB.Label Label8 
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
            Left            =   720
            TabIndex        =   87
            Top             =   300
            Width           =   1515
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00CDDADA&
         Caption         =   "IMPRESION PRESUPUESTO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   180
         TabIndex        =   69
         Top             =   2685
         Width           =   2415
         Begin VB.ComboBox Combo2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   70
            Top             =   280
            Width           =   2205
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00CDDADA&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   780
         Index           =   4
         Left            =   105
         TabIndex        =   18
         Top             =   1830
         Width           =   4950
         Begin VB.TextBox FORMUNAME 
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
            Height          =   285
            Index           =   4
            Left            =   2745
            TabIndex        =   21
            Text            =   " "
            Top             =   60
            Width           =   1275
         End
         Begin VB.CommandButton BOTSELFORM 
            Caption         =   "."
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
            Left            =   4005
            TabIndex        =   20
            Top             =   60
            Width           =   175
         End
         Begin VB.CommandButton BOTOPENFORM 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   620
            Index           =   4
            Left            =   4440
            Picture         =   "CONFREPA09.frx":1690
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   0
            Width           =   540
         End
         Begin VB.Label FORMUDES 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Secuencia Operaciones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   6
            Left            =   330
            TabIndex        =   23
            Top             =   165
            Width           =   2325
         End
         Begin VB.Label FORMUECO 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   90
            TabIndex        =   22
            Top             =   480
            Width           =   4095
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00CDDADA&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   705
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   1050
         Width           =   4960
         Begin VB.CommandButton BOTOPENFORM 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   620
            Index           =   3
            Left            =   4400
            Picture         =   "CONFREPA09.frx":199A
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton BOTSELFORM 
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
            Index           =   3
            Left            =   4005
            TabIndex        =   9
            Top             =   60
            Width           =   175
         End
         Begin VB.TextBox FORMUNAME 
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
            Height          =   285
            Index           =   3
            Left            =   2745
            TabIndex        =   8
            Text            =   " "
            Top             =   60
            Width           =   1275
         End
         Begin VB.Label FORMUECO 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   90
            TabIndex        =   12
            Top             =   390
            Width           =   4095
         End
         Begin VB.Label FORMUDES 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Presupuesto Cliente"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   225
            TabIndex        =   11
            Top             =   120
            Width           =   2430
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00CDDADA&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   750
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   195
         Width           =   4950
         Begin VB.TextBox FORMUNAME 
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
            Height          =   285
            Index           =   1
            Left            =   2745
            TabIndex        =   4
            Text            =   " "
            Top             =   105
            Width           =   1275
         End
         Begin VB.CommandButton BOTSELFORM 
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
            Index           =   1
            Left            =   4005
            TabIndex        =   3
            Top             =   105
            Width           =   175
         End
         Begin VB.CommandButton BOTOPENFORM 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   620
            Index           =   1
            Left            =   4400
            Picture         =   "CONFREPA09.frx":1CA4
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   105
            Width           =   540
         End
         Begin VB.Label FORMUECO 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   90
            TabIndex        =   5
            Top             =   435
            Width           =   4095
         End
         Begin VB.Label FORMUDES 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Presupuesto Interno"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   360
            TabIndex        =   6
            Top             =   165
            Width           =   2325
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00CDDADA&
      Caption         =   "ORDEN DE REPARACION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      Left            =   5460
      TabIndex        =   24
      Top             =   1935
      Width           =   5100
      Begin VB.Frame Frame4 
         BackColor       =   &H00CDDADA&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   750
         Index           =   6
         Left            =   105
         TabIndex        =   79
         Top             =   1680
         Width           =   4950
         Begin VB.CommandButton COMMAND11 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   620
            Left            =   4200
            Picture         =   "CONFREPA09.frx":1FAE
            Style           =   1  'Graphical
            TabIndex        =   82
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton COMMAND10 
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
            Left            =   3885
            TabIndex        =   81
            Top             =   0
            Width           =   175
         End
         Begin VB.TextBox TEXT10 
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
            Height          =   285
            Left            =   2640
            TabIndex        =   80
            Text            =   " "
            Top             =   0
            Width           =   1275
         End
         Begin VB.Label LABEL11 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   210
            TabIndex        =   84
            Top             =   360
            Width           =   3855
         End
         Begin VB.Label FORMUDES 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Registro de Reparación"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   12
            Left            =   210
            TabIndex        =   83
            Top             =   80
            Width           =   2325
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00CDDADA&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   855
         Index           =   0
         Left            =   105
         TabIndex        =   31
         Top             =   960
         Width           =   4950
         Begin VB.TextBox FORMUNAME 
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
            Height          =   285
            Index           =   0
            Left            =   2640
            TabIndex        =   34
            Text            =   " "
            Top             =   0
            Width           =   1275
         End
         Begin VB.CommandButton BOTSELFORM 
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
            Index           =   0
            Left            =   3885
            TabIndex        =   33
            Top             =   0
            Width           =   175
         End
         Begin VB.CommandButton BOTOPENFORM 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   620
            Index           =   0
            Left            =   4200
            Picture         =   "CONFREPA09.frx":22B8
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   0
            Width           =   540
         End
         Begin VB.Label FORMUDES 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hoja de Procedimientos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   210
            TabIndex        =   36
            Top             =   80
            Width           =   2325
         End
         Begin VB.Label FORMUECO 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   210
            TabIndex        =   35
            Top             =   360
            Width           =   3855
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00CDDADA&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   780
         Index           =   2
         Left            =   105
         TabIndex        =   25
         Top             =   240
         Width           =   4950
         Begin VB.CommandButton BOTOPENFORM 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   620
            Index           =   2
            Left            =   4200
            Picture         =   "CONFREPA09.frx":25C2
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton BOTSELFORM 
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
            Left            =   3885
            TabIndex        =   27
            Top             =   20
            Width           =   175
         End
         Begin VB.TextBox FORMUNAME 
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
            Height          =   285
            Index           =   2
            Left            =   2640
            TabIndex        =   26
            Text            =   " "
            Top             =   0
            Width           =   1275
         End
         Begin VB.Label FORMUDES 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Formulario de O/Rep"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   120
            TabIndex        =   29
            Top             =   80
            Width           =   2430
         End
         Begin VB.Label FORMUECO 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   210
            TabIndex        =   30
            Top             =   390
            Width           =   3855
         End
      End
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuTablasAuxiliares 
         Caption         =   "Tablas Auxiliares"
      End
   End
   Begin VB.Menu mnuPresupuesto 
      Caption         =   "Presupuesto"
      Begin VB.Menu mnuGeneraAutoImporteMinimo 
         Caption         =   "Generación Automática por Importe Mínimo"
      End
      Begin VB.Menu linea1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCostoReparaciones 
         Caption         =   "Costo de Mano de Obra por Artículo"
      End
   End
   Begin VB.Menu mnuControlFinal 
      Caption         =   "Control Final"
      Begin VB.Menu mnuCantidadControlesPorArticulo 
         Caption         =   "Cantidad de Controles por Artículo"
      End
      Begin VB.Menu mnuDocumentosExternos 
         Caption         =   "Documentos Externos"
         Begin VB.Menu mnuAgregarDocumentosArticulo 
            Caption         =   "Agregar Documentos por Artículo"
         End
         Begin VB.Menu mnuTablaGeneralDocumentos 
            Caption         =   "Tabla General de Documentos"
         End
      End
   End
End
Attribute VB_Name = "CONFREPA09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BOTOPENFORM_Click(Index As Integer)
    Call CONECRUN("FLFORMS/" + TRIM$(FORMUNAME(Index).TEXT), "Formularios de Impresion")
    If EXISTE%("FLFORMS.EXE") < 1 Then
        Call FINAL("")
    End If
End Sub

Private Sub BOTSEL_Click(Index As Integer)
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        CTACBLENUM(Index) = VDEVU$
        DESCRIECO(Index) = VALACT2$("ECUECON")
    End If

End Sub

Private Sub BOTSELFORM_Click(Index As Integer)
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        FORMUNAME(Index).TEXT = VDEVU$
        FORMUECO(Index).Caption = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Check1_Click()
   If Check1.Value = 1 Then
      Label3 = "0"
   End If
End Sub


Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command10_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text10.TEXT = VDEVU$
        Label11 = VALACT2$("INDIFRM")
    End If

End Sub

Private Sub Command11_Click()
    Call CONECRUN("FLFORMS/" + TRIM$(Text10), "Formularios de Impresion")
    If EXISTE%("FLFORMS.EXE") < 1 Then
        Call FINAL("")
    End If

End Sub

Private Sub Command15_Click()
    Call SELECHO("TADEPOSI")
    If VALACT1$("TADEPOSI") <> "" Then
        Text5 = TRIM$(VALACT1$("TADEPOSI"))
    End If

End Sub

Private Sub Command2_Click()
    '
    If CODINT%(Label1) < 1 Then
       Call MENSERR(24, "Falta Código Genérico de Reparaciones")
       Exit Sub
    End If
    '
    If XVALO(Text8) > 0 Then
       APARTIRNUEVONUM& = XVALO(Text8)
       ValorMáximoEnBD& = XVALO(VALOBADA("REPARA", "MAX(NUMREPA)", CONDI$, "NOMESS"))
       If APARTIRNUEVONUM& <= ValorMáximoEnBD& Then
          Call COMUNI("Atención !!! Número de inicialización Existente\Debe Ingresar Número Igual o Mayor a " & ValorMáximoEnBD&)
          Call PINTATEXT(Text8)
          Exit Sub
       End If
    End If
    '
    
    For KK% = 0 To FORMUNAME.UBound
       On Error Resume Next
       Call GRACONTROL("FOREPAR", "DOCUREP" + TRIM$(Str$(KK%)), FORMUNAME(KK%).TEXT)
    Next KK%
    
    Call GRACONTROL("REPARA", "NUEVANUM", SAFETEXT$(XVALO(Text8)))
    '
    Call GRACONTROL("REPARA", "VALHORA", Text1.TEXT)
    Call GRACONTROL("REPARA", "COEFPREA", Text2.TEXT)
    Call GRACONTROL("REPARA", "CODGENER", Label1)
    Call GRACONTROL("REPARA", "LIPREREP", Label3)
    '
    Call GRACONTROL$("", "PLAENREP", Text4)
    Call GRACONTROL$("", "CONPAREP", Text3)
    Call GRACONTROL$("", "DEPRECEN", Text5)
    Call GRACONTROL$("", "DEPREDES", Text6)
    Call GRACONTROL$("REPARA", "FOREPAR", Text7)
    '
    Call GRACONTROL("REPARA", "PUERTOIM", Combo2.TEXT) 'GRABA LA IMPRESORA SELECCIONADA
    Call GRACONTROL$("REPARA", "FOREPAR", Text7) 'MEDIO DE COBRANZA POR DEFECTO.
    '
    Call GRACONTROL("", "RECONRE", Text10.TEXT)
    
    Call GRACONTROL(IDENTER$, "PORETIRE", Combo1.TEXT)
    Call GRACONTROL("*", "PORETIRE", Combo1.TEXT)
    
    Call GRACONTROL("REPARA", "CTCONREP", CTACBLENUM(0).TEXT)
    Call GRACONTROL("REPARA", "CTBREPUE", CTACBLENUM(1).TEXT)
    
    valor$ = "NO"
    If Check2.Value = 1 Then valor$ = "SI"
    Call GRACONTROL("REPARA", "MOCOSREP", valor$)
    
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   Call SELECHO("MASTER")
   VDEVU$ = TRIM$(VALACT1$("MASTER"))
   If VDEVU$ <> "" Then
     Label1 = VDEVU$
   End If
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Call SELECHO("LISTAS")
    VDEVU$ = VALACT1$("LISTAS")
    If VDEVU$ <> "" Then
        Label3 = VDEVU$
        Check1.Value = 0
    End If
End Sub



Private Sub Command6_Click()
    Call SELECHO("TADEPOSI")
    If VALACT1$("TADEPOSI") <> "" Then
        Text6 = TRIM$(VALACT1$("TADEPOSI"))
    End If

End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   MECOBSEL% = XVALO(PRESURE09.TablaCobranzaSel(1))
   If MECOBSEL% < 1 Then GoTo 99
   Text7 = MECOBSEL%
99 Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   On Error Resume Next
   REPABOMB07.Show 1
   If Err Then
      Call MENSERR(24, "No Disponible en la Presente Versión")
   End If
   Command8.Enabled = True
End Sub

Private Sub CTACBLENUM_Change(Index As Integer)
    DESCRIECO(Index).Caption = ""
    VDEVU$ = TRIM$(CTACBLENUM(Index))
    
    DENOCUES$ = ECOARCH$("ECUECON", VDEVU$)
    DESCRIECO(Index).Caption = DENOCUES$

End Sub

Private Sub Form_Load()
    '
    For KK% = 0 To FORMUNAME.UBound
       On Error Resume Next
       FORMUNAME(KK%).TEXT = Control$("FOREPAR", "DOCUREP" + TRIM$(Str$(KK%)))
       FORMUECO(KK%).Caption = ECOARCH$("INDIFRM", FORMUNAME(KK%).TEXT)
    Next KK%
    Text10 = Control("", "RECONRE")
    Text8 = XVALO(Control("REPARA", "NUEVANUM"))
    '
    CUHOR = Val(Control$("REPARA", "VALHORA"))
    Text1 = CSTRING$(MKS$(CUHOR), 4, 8, 2, 2)
    COEFI = Val(Control$("REPARA", "COEFPREA"))
    Text2 = CSTRING$(MKS$(COEFI), 4, 8, 2, 2)
    Label1 = Control$("REPARA", "CODGENER")
    Label3 = TRIM$(Control$("REPARA", "LIPREREP"))
      If Val(Label3) < 1 Then Check1.Value = 1
      '
    PLAENREP$ = TRIM$(Control$("", "PLAENREP"))
    If PLAENREP$ = "" Then PLAENREP$ = "5 Días hábiles de recibida su órden de compra."
    Text4 = PLAENREP$
    '
    CONPAREP$ = TRIM$(Control$("", "CONPAREP"))
    If CONPAREP$ = "" Then CONPAREP$ = "Contado contra entrega"
    Text3 = CONPAREP$
    '
    Text5 = ""
    DEPOSI% = XVALO(Control$("", "DEPRECEN"))
    If DEPOSI% > 0 Then
      Text5 = XVALO(Control$("", "DEPRECEN"))
    End If
    '
    Text6 = ""
    DEPOSI% = XVALO(Control$("", "DEPREDES"))
    If DEPOSI% > 0 Then
      Text6 = XVALO(Control$("", "DEPREDES"))
    End If
    '
    CTACBLENUM(0).TEXT = Control("REPARA", "CTCONREP")
    CTACBLENUM(1).TEXT = Control("REPARA", "CTBREPUE")
    
    If Control("REPARA", "MOCOSREP") = "SI" Then Check2.Value = 1

    'MEDIO DE COBRANZA POR DEFECTO.
    Text7 = XVALO(Control$("REPARA", "FOREPAR"))
    PREP$ = Control$("REPARA", "PUERTOIM")
    Combo2.Clear ' CARGA EL COMBO DE LAS IMPRESORAS
    Dim PRX As Printer
    On Error Resume Next
    Combo2.TEXT = Printer.DeviceName
    On Error GoTo 0
    For Each PRX In Printers
      Combo2.AddItem PRX.DeviceName
      If PRX.DeviceName = PREP$ Then
        Combo2.TEXT = PREP$
      End If
    Next
    
    PRINTERPORT$ = Control$(IDENTER$, "PORETIRE") ' IMPRESORA CONFIGURADA
    If PRINTERPORT$ = "" Then PRINTERPORT$ = Control$("*", "PORETIRE")
    Combo1.Clear ' CARGA EL COMBO DE LAS IMPRESORAS
    Dim PRX1 As Printer
    On Error Resume Next
    Combo1.TEXT = Printer.DeviceName
    On Error GoTo 0
    For Each PRX1 In Printers
      Combo1.AddItem PRX1.DeviceName
      If PRX1.DeviceName = PRINTERPORT$ Then
        Combo1.TEXT = PRINTERPORT$
      End If
    Next

End Sub

Private Sub FORMUNAME_DblClick(Index As Integer)
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        FORMUNAME(Index).TEXT = VDEVU$
        FORMUECO(Index).Caption = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Label1_DblClick()
   Call SELECHO("MASTER")
   VDEVU$ = TRIM$(VALACT1$("MASTER"))
   If VDEVU$ <> "" Then
     Label1 = VDEVU$
   End If
End Sub

Private Sub Label3_Change()
   If Val(Label3) < 1 Then
       Label4 = "Costo Calculado / Informado"
     Else
       Label4 = ECOARCH("LISTAS", Chr$(Val(Label3)))
   End If
End Sub

Private Sub Label3_DblClick()
   Call Command4_Click
End Sub

Private Sub mnuAgregarDocumentosArticulo_Click()

10  COD$ = ListaSeleccionArticulos(DescripcionArticulo$, False)
    
    If COD$ = "" Then GoTo 99
    
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
            CodigoExterno$ = COD$
            DescripcionProducto$ = DescripcionArticulo$
            
            IDLIST& = ObtenerIdTablaAuxiliar("DOCEXCONFIN")
            
            ConsultaSql$ = "INSERT INTO DATASQL (ID_LISTA, CAMPO1, CAMPO2, CAMPO3, CAMPO4, CAMPO5) VALUES ('" & CStr(IDLIST&) & "','','" & CodigoExterno$ & "','" & DescripcionProducto$ & "','" & NombreDocumento$ & "','" & RutaDocumento$ & "')"
            FLEXCONN.Execute ConsultaSql$
        End If
    End If
    
    GoTo 10

99  Exit Sub

End Sub

Private Sub mnuCantidadControlesPorArticulo_Click()

    Call GENERAR_ESTRUCTURAS_SQL
    IDLIST& = ObtenerIdTablaAuxiliar("CANCONART")

    Call ABM_TABLA(IDLIST&)

End Sub

Private Sub mnuCostoReparaciones_Click()
    
    Opcion% = COMALTER%("Costo de Reparación\\Listado de Costo M. Obra por Artículo\Actualización por Porcentaje")
  
    If Opcion% = 1 Then
        Call GENERAR_ESTRUCTURAS_SQL
        IDLIST& = ObtenerIdTablaAuxiliar("LISCOSREPA")
  
        Call ABM_TABLA(IDLIST&)
        
    ElseIf Opcion% = 2 Then

    End If

End Sub

Private Sub mnuGeneraAutoImporteMinimo_Click()

    ImporteActual$ = TRIM$(Control$("REPARA", "GENAUMIN"))
    ImporteMinimoGeneraAuto# = XVALO(InputBox("Importe", "Generación Automática por Importe Mínimo", ImporteActual$))
    
    ImporteNuevo$ = TRIM$(FORMATNUM$(ImporteMinimoGeneraAuto#, "F10.2"))
    
    If ImporteNuevo$ = "" Then
        RESPUESTA% = COMALTER%("¿Está Seguro de Quitar el Importe Mínimo\Para la Emisión Automática de Presupuestos?\\No, Mantener\Si, Eliminar")
        If RESPUESTA% = 1 Then Exit Sub
    End If
    
    Call GRACONTROL("REPARA", "GENAUMIN", ImporteNuevo$)

End Sub

Private Sub mnuTablaGeneralDocumentos_Click()

    Call GENERAR_ESTRUCTURAS_SQL
    IDLIST& = ObtenerIdTablaAuxiliar("DOCEXCONFIN")
  
    Call ABM_TABLA(IDLIST&)
    
End Sub

Sub mnuTablasAuxiliares_Click()
  Call GENERAR_ESTRUCTURAS_SQL
  Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  FRMZELECHO.Caption = "LISTA DE ESTRUCTURAS PRF Y RFS"

  Dim obj As New RECORXET
  Set RS1 = obj.RS_ENCAPRFLISTAS()
  Call Carga_MSF_Recordset(RS1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO.msf2)
  Set obj = Nothing
  RS1.Close
  Set RS1 = Nothing
  EPAC$ = UCase$(TRIM$(EMPREAC$))
  '

  FRMZELECHO.Width = 8500
  FRMZELECHO.Show 1
  IDLIST& = XVALO(RESP_ZELE_VECT(1))
  If InStr(EPAC$, "ENVAPLAST") > 0 Then
    LIXTA$ = TRIM$(RESP_ZELE_VECT(2))
    If LIXTA$ = "MATERIALES ENVAPLAST" Then
       Call COMUNI("Atención !!! Editar Tabla de Familias Desde Maestro de Artículos")
       Exit Sub
    End If
  End If

  Call ABM_TABLA(IDLIST&)

End Sub

Sub Option1_DblClick()
    '
    Call COPYSTRU("INDIPEDI", "TACAUFA")
    '
    Call BLANARCH("!TACAUFA")
    For i& = 1 To ULTREG&("TACAUFA")
        X$ = REGLEIDO$("TACAUFA", i&)
        If CVS(Left$(X$, 4)) > 0 Then
            Call REGAPP("!TACAUFA", X$)
        End If
    Next i&
    Call CIERRARCH("TACAUFA")
    '
    VTB% = VENTABU%("TACAUFA/NC/TITUPAN=Causas de Fallas")
    If VTB% < 0 Then Exit Sub
    J& = 0
    For i& = 1 To ULTREG&("!TACAUFA")
        X$ = REGLEIDO$("!TACAUFA", i&)
        If CVS(Left$(X$, 4)) > 0 Then
            J& = J& + 1
            Call GRAREG("TACAUFA", X$, J&)
        End If
    Next i&
    '
    Call SETULTREG("TACAUFA", J&)
    Call CIERRARCH("TACAUFA")

End Sub

Sub Option3_DblClick()
    '
    Call COPYSTRU("INDIPEDI", "TATIPFA")
    Call BLANARCH("!TATIPFA")
    For i& = 1 To ULTREG&("TATIPFA")
        X$ = REGLEIDO$("TATIPFA", i&)
        If CVS(Left$(X$, 4)) > 0 Then
            Call REGAPP("!TATIPFA", X$)
        End If
    Next i&
    Call CIERRARCH("TATIPFA")
    '
    VTB% = VENTABU%("TATIPFA/NC/TITUPAN=Tipo de Falla")
    If VTB% < 0 Then Exit Sub
    J& = 0
    For i& = 1 To ULTREG&("!TATIPFA")
        X$ = REGLEIDO$("!TATIPFA", i&)
        If CVS(Left$(X$, 4)) > 0 Then
            J& = J& + 1
            Call GRAREG("TATIPFA", X$, J&)
        End If
    Next i&
    '
    Call SETULTREG("TATIPFA", J&)
    Call CIERRARCH("TATIPFA")


End Sub

Private Sub TEXT10_Change()
  If ECOARCH$("INDIFRM", Text10.TEXT) <> "" Then
    Label11 = ECOARCH$("INDIFRM", Text10.TEXT)
  Else
    Label11 = ""
  End If
  
End Sub

Private Sub TEXT10_DblClick()
    '
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text10.TEXT = VDEVU$
        Label11 = VALACT2$("INDIFRM")
    End If

End Sub

Private Sub Text5_Change()
   If XVALO(Text5) > 0 And XVALO(Text5) < 256 Then
      Label5 = ECOARCH$("TADEPOSI", Chr$(Val(Text5)))
   Else
      Label5 = ""
   End If

End Sub

Private Sub Text6_Change()
   If XVALO(Text6) > 0 And XVALO(Text6) < 256 Then
      Label7 = ECOARCH$("TADEPOSI", Chr$(Val(Text6)))
   Else
      Label7 = ""
   End If

End Sub


Private Sub Text7_Change()
   Label8 = VALOBADA("DATASQL", "CAMPO2", "LISTA = 'MEDIOS COBRANZA-MDT' AND CAMPO1 = '" & TRIM$(FORMATOCON0$(XVALO(Text7), 7)) & "'")

End Sub
