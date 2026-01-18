VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form ACRE_MAQ04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0D0F0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "A-B-M-L Base de Datos de Proveedores "
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11460
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ACRE_MAQ04.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   11460
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0D0F0&
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5400
      Left            =   0
      TabIndex        =   94
      Top             =   1470
      Width           =   10450
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0D0F0&
         Caption         =   "Domicilio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   105
         TabIndex        =   123
         Top             =   0
         Width           =   10295
         Begin VB.TextBox Text29 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7035
            TabIndex        =   129
            Text            =   " "
            Top             =   1155
            Width           =   3080
         End
         Begin VB.TextBox Text27 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3885
            TabIndex        =   128
            Text            =   " "
            Top             =   735
            Width           =   6230
         End
         Begin VB.TextBox Text26 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1365
            TabIndex        =   127
            Text            =   " "
            Top             =   735
            Width           =   1305
         End
         Begin VB.TextBox Text25 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1365
            TabIndex        =   126
            Text            =   " "
            Top             =   315
            Width           =   8750
         End
         Begin VB.TextBox COTEXT 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   1365
            TabIndex        =   125
            Text            =   "  "
            Top             =   1155
            Width           =   465
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
            Height          =   315
            Index           =   0
            Left            =   1785
            TabIndex        =   124
            Top             =   1155
            Width           =   175
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "País:"
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
            Left            =   5040
            TabIndex        =   135
            Top             =   1260
            Width           =   1905
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cod.Postal:"
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
            Left            =   -630
            TabIndex        =   134
            Top             =   840
            Width           =   1905
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Calle y Nro:"
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
            Left            =   -630
            TabIndex        =   133
            Top             =   420
            Width           =   1905
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Provincia:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   105
            TabIndex        =   132
            Top             =   1260
            Width           =   1185
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Localidad:"
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
            Left            =   1890
            TabIndex        =   131
            Top             =   840
            Width           =   1905
         End
         Begin VB.Label DETEXT 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   2100
            TabIndex        =   130
            Top             =   1155
            Width           =   4005
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00C0D0F0&
         Caption         =   "Agenda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2180
         Left            =   105
         TabIndex        =   112
         Top             =   1785
         Width           =   4845
         Begin VB.TextBox Text32 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   1365
            TabIndex        =   117
            Text            =   " "
            Top             =   1035
            Width           =   3300
         End
         Begin VB.TextBox Text31 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   1365
            TabIndex        =   116
            Text            =   " "
            Top             =   675
            Width           =   3300
         End
         Begin VB.TextBox Text30 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   1365
            TabIndex        =   115
            Text            =   " "
            Top             =   315
            Width           =   3300
         End
         Begin VB.TextBox Text19 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   1365
            TabIndex        =   114
            Text            =   " "
            Top             =   1395
            Width           =   3300
         End
         Begin VB.TextBox Text36 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   1365
            TabIndex        =   113
            Text            =   " "
            Top             =   1755
            Width           =   3300
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "E-Mail:"
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
            Left            =   -630
            TabIndex        =   122
            Top             =   1120
            Width           =   1905
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Fax:"
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
            Left            =   -630
            TabIndex        =   121
            Top             =   760
            Width           =   1905
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono:"
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
            Left            =   -630
            TabIndex        =   120
            Top             =   400
            Width           =   1905
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Contacto:"
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
            Left            =   -630
            TabIndex        =   119
            Top             =   1480
            Width           =   1905
         End
         Begin VB.Label Label71 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Horario:"
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
            Left            =   -630
            TabIndex        =   118
            Top             =   1840
            Width           =   1905
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00C0D0F0&
         Caption         =   "Impuestos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   105
         TabIndex        =   105
         Top             =   3990
         Width           =   4845
         Begin VB.TextBox Text33 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2310
            MaxLength       =   24
            TabIndex        =   108
            Top             =   315
            Width           =   2355
         End
         Begin VB.TextBox COTEXT 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   945
            MaxLength       =   2
            TabIndex        =   107
            Text            =   " "
            Top             =   315
            Width           =   360
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
            Height          =   315
            Index           =   1
            Left            =   1260
            TabIndex        =   106
            Top             =   315
            Width           =   175
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "CUIT:"
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
            Left            =   315
            TabIndex        =   111
            Top             =   420
            Width           =   1905
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "P-IVA:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   -525
            TabIndex        =   110
            Top             =   420
            Width           =   1380
         End
         Begin VB.Label DETEXT 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   945
            TabIndex        =   109
            Top             =   700
            Width           =   3720
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00C0D0F0&
         Caption         =   "Comerciales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2550
         Left            =   5145
         TabIndex        =   98
         Top             =   2625
         Width           =   5255
         Begin VB.TextBox COTEXT 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   1155
            TabIndex        =   101
            Text            =   " "
            Top             =   315
            Width           =   570
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
            Height          =   315
            Index           =   4
            Left            =   1730
            TabIndex        =   100
            Top             =   315
            Width           =   175
         End
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1665
            Left            =   1155
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   99
            Text            =   "ACRE_MAQ04.frx":0442
            Top             =   735
            Width           =   3920
         End
         Begin VB.Label Label40 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cond.Pago:"
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
            Left            =   40
            TabIndex        =   104
            Top             =   420
            Width           =   1080
         End
         Begin VB.Label DETEXT 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   1995
            TabIndex        =   103
            Top             =   315
            Width           =   3080
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Observac:"
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
            Left            =   45
            TabIndex        =   102
            Top             =   700
            Width           =   1080
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0D0F0&
         Caption         =   "Otros Datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   760
         Left            =   5145
         TabIndex        =   95
         Top             =   1785
         Width           =   5255
         Begin VB.TextBox Text24 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2100
            TabIndex        =   96
            Text            =   " "
            Top             =   315
            Width           =   2975
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre de Fantasía:"
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
            Left            =   105
            TabIndex        =   97
            Top             =   405
            Width           =   1905
         End
      End
   End
   Begin VB.Frame Frame19 
      BackColor       =   &H00C0D0F0&
      Caption         =   "Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1300
      Left            =   8085
      TabIndex        =   89
      Top             =   135
      Width           =   2325
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0D0F0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   880
         Left            =   150
         TabIndex        =   90
         Top             =   270
         Width           =   1990
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
            Height          =   500
            Left            =   725
            Picture         =   "ACRE_MAQ04.frx":0444
            Style           =   1  'Graphical
            TabIndex        =   93
            ToolTipText     =   "Restablece los Valores Anteriores - Descarta Cambios Efectuados"
            Top             =   250
            Width           =   550
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
            Height          =   500
            Left            =   155
            Picture         =   "ACRE_MAQ04.frx":074E
            Style           =   1  'Graphical
            TabIndex        =   92
            ToolTipText     =   "Guarda los Cambios Realizados"
            Top             =   250
            Width           =   550
         End
         Begin VB.CommandButton Command20 
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
            Left            =   1280
            Picture         =   "ACRE_MAQ04.frx":0A58
            Style           =   1  'Graphical
            TabIndex        =   91
            ToolTipText     =   "Búsqueda por CUIT, Domicilio, Nombre de Fantasía, etc"
            Top             =   250
            Width           =   550
         End
      End
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00C0D0F0&
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1300
      Left            =   120
      TabIndex        =   80
      Top             =   135
      Width           =   7890
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
         Height          =   330
         Left            =   2205
         Picture         =   "ACRE_MAQ04.frx":1DBE
         TabIndex        =   85
         Top             =   360
         Width           =   175
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1470
         MaxLength       =   5
         TabIndex        =   84
         Text            =   " "
         Top             =   360
         Width           =   780
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1470
         TabIndex        =   83
         Top             =   780
         Width           =   6180
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5565
         TabIndex        =   82
         Top             =   360
         Width           =   1965
      End
      Begin VB.CommandButton Command10 
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
         Height          =   315
         Left            =   7500
         Picture         =   "ACRE_MAQ04.frx":20C8
         TabIndex        =   81
         Top             =   360
         Width           =   175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro de Cuenta:"
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
         Left            =   210
         TabIndex        =   88
         Top             =   285
         Width           =   1140
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R.Social:"
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
         Left            =   0
         TabIndex        =   87
         Top             =   990
         Width           =   1350
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Código de Cuenta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   4410
         TabIndex        =   86
         Top             =   285
         Width           =   1110
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6675
      Left            =   0
      TabIndex        =   4
      Top             =   8160
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   11774
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   12636400
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "ACRE_MAQ04.frx":23D2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label17"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "DETEXT(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label39"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label35"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame12"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame7"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "BOTSEL(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "COTEXT(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text34"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Command12"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text4"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Frame1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Anotaciones"
      TabPicture(1)   =   "ACRE_MAQ04.frx":23EE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture2"
      Tab(1).Control(1)=   "Text28"
      Tab(1).Control(2)=   "Command17"
      Tab(1).Control(3)=   "Command8"
      Tab(1).Control(4)=   "Command6"
      Tab(1).Control(5)=   "Command5"
      Tab(1).Control(6)=   "Command7"
      Tab(1).Control(7)=   "Picture4"
      Tab(1).Control(8)=   "Line3"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Cuenta Corriente"
      TabPicture(2)   =   "ACRE_MAQ04.frx":240A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame18"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0D0F0&
         Caption         =   "Tipo de Cuenta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   880
         Left            =   1320
         TabIndex        =   77
         Top             =   4320
         Width           =   2115
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0D0F0&
            Caption         =   "C.Corriente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   315
            TabIndex        =   79
            Top             =   280
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0D0F0&
            Caption         =   "Cta.de Gastos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   215
            Left            =   315
            TabIndex        =   78
            Top             =   525
            Width           =   1695
         End
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1665
         TabIndex        =   75
         Text            =   " "
         Top             =   3840
         Width           =   3930
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Gan"
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
         Left            =   600
         TabIndex        =   74
         ToolTipText     =   "Datos Retención Ganancias Cliente Activo"
         Top             =   720
         Width           =   540
      End
      Begin VB.TextBox Text34 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3195
         MaxLength       =   15
         TabIndex        =   70
         Text            =   " "
         Top             =   2760
         Width           =   2355
      End
      Begin VB.TextBox COTEXT 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   1830
         MaxLength       =   2
         TabIndex        =   69
         Text            =   " "
         Top             =   2760
         Width           =   360
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
         Height          =   315
         Index           =   2
         Left            =   2160
         TabIndex        =   68
         Top             =   2760
         Width           =   175
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   4905
         TabIndex        =   66
         Top             =   3135
         Width           =   750
         Begin VB.CommandButton Command22 
            Caption         =   "Cfg"
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
            Left            =   90
            TabIndex        =   67
            ToolTipText     =   "Datos Retención I.Brutos Proveedor Activo"
            Top             =   30
            Width           =   540
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Contable"
         Height          =   1625
         Left            =   360
         TabIndex        =   57
         Top             =   960
         Width           =   5255
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
            Height          =   315
            Index           =   6
            Left            =   4880
            TabIndex        =   61
            Top             =   735
            Width           =   175
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
            Height          =   315
            Index           =   7
            Left            =   1725
            TabIndex        =   60
            Top             =   315
            Width           =   175
         End
         Begin VB.TextBox COTEXT 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   7
            Left            =   1155
            TabIndex        =   59
            Text            =   " "
            Top             =   315
            Width           =   570
         End
         Begin VB.TextBox COTEXT 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   2625
            TabIndex        =   58
            Text            =   " "
            Top             =   735
            Width           =   2270
         End
         Begin VB.Label DETEXT 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   7
            Left            =   1995
            TabIndex        =   65
            Top             =   315
            Width           =   3080
         End
         Begin VB.Label DETEXT 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   420
            TabIndex        =   64
            Top             =   1155
            Width           =   4655
         End
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Grupo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   315
            TabIndex        =   63
            Top             =   420
            Width           =   765
         End
         Begin VB.Label Label42 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta Madre:"
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
            Left            =   630
            TabIndex        =   62
            Top             =   840
            Width           =   1905
         End
      End
      Begin VB.Frame Frame18 
         BorderStyle     =   0  'None
         Caption         =   "Frame18"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6000
         Left            =   -74685
         TabIndex        =   12
         Top             =   440
         Width           =   9990
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2800
            Left            =   105
            ScaleHeight     =   2745
            ScaleWidth      =   9615
            TabIndex        =   18
            Top             =   3120
            Width           =   9675
            Begin VB.Frame Frame14 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Caption         =   "Frame3"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   5630
               Left            =   0
               TabIndex        =   21
               Top             =   315
               Width           =   10515
               Begin VB.ListBox List2 
                  BeginProperty Font 
                     Name            =   "Fixedsys"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   2085
                  Left            =   -30
                  TabIndex        =   22
                  Top             =   0
                  Width           =   9675
               End
               Begin VB.Line Line19 
                  X1              =   7350
                  X2              =   7350
                  Y1              =   2100
                  Y2              =   2415
               End
               Begin VB.Line Line18 
                  X1              =   6195
                  X2              =   6195
                  Y1              =   2100
                  Y2              =   2415
               End
               Begin VB.Line Line17 
                  X1              =   4200
                  X2              =   4200
                  Y1              =   2100
                  Y2              =   2415
               End
               Begin VB.Line Line16 
                  X1              =   3045
                  X2              =   3045
                  Y1              =   2100
                  Y2              =   2415
               End
               Begin VB.Line Line15 
                  X1              =   0
                  X2              =   9660
                  Y1              =   2420
                  Y2              =   2420
               End
               Begin VB.Line Line14 
                  BorderWidth     =   2
                  X1              =   0
                  X2              =   9660
                  Y1              =   2085
                  Y2              =   2085
               End
               Begin VB.Label Label22 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Total General"
                  Height          =   225
                  Left            =   25
                  TabIndex        =   25
                  Top             =   2130
                  Width           =   1695
               End
               Begin VB.Label Label25 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
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
                  Left            =   4330
                  TabIndex        =   24
                  Top             =   2130
                  Width           =   1695
               End
               Begin VB.Label Label28 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
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
                  Left            =   7455
                  TabIndex        =   23
                  Top             =   2130
                  Width           =   1695
               End
               Begin VB.Line Line12 
                  X1              =   2205
                  X2              =   2205
                  Y1              =   1575
                  Y2              =   3360
               End
            End
            Begin VB.Frame Frame15 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Caption         =   "Frame4"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   0
               TabIndex        =   19
               Top             =   0
               Width           =   10515
               Begin VB.Line Line11 
                  X1              =   6195
                  X2              =   6195
                  Y1              =   0
                  Y2              =   1050
               End
               Begin VB.Line Line10 
                  X1              =   4200
                  X2              =   4200
                  Y1              =   0
                  Y2              =   1050
               End
               Begin VB.Line Line9 
                  X1              =   3045
                  X2              =   3045
                  Y1              =   0
                  Y2              =   1050
               End
               Begin VB.Line Line8 
                  X1              =   2205
                  X2              =   2205
                  Y1              =   0
                  Y2              =   1050
               End
               Begin VB.Label Label36 
                  BackStyle       =   0  'Transparent
                  Caption         =   " Documento                     Nro      Emision                    Importe      Vencim.                      Saldo              "
                  Height          =   330
                  Left            =   0
                  TabIndex        =   20
                  Top             =   50
                  Width           =   10410
               End
               Begin VB.Line Line4 
                  X1              =   7350
                  X2              =   7350
                  Y1              =   0
                  Y2              =   1050
               End
            End
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2640
            Left            =   105
            ScaleHeight     =   2580
            ScaleWidth      =   9615
            TabIndex        =   13
            Top             =   210
            Width           =   9675
            Begin VB.Frame Frame16 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Caption         =   "Frame4"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   0
               TabIndex        =   14
               Top             =   0
               Width           =   10515
               Begin VB.Line Line5 
                  X1              =   7560
                  X2              =   7560
                  Y1              =   0
                  Y2              =   1050
               End
               Begin VB.Label Label37 
                  BackStyle       =   0  'Transparent
                  Caption         =   "   Fecha       Documento                      Nro                Debe                   Haber                  Saldo              "
                  Height          =   330
                  Left            =   0
                  TabIndex        =   15
                  Top             =   50
                  Width           =   10410
               End
               Begin VB.Line Line6 
                  X1              =   1155
                  X2              =   1155
                  Y1              =   0
                  Y2              =   1050
               End
               Begin VB.Line Line7 
                  X1              =   3360
                  X2              =   3360
                  Y1              =   0
                  Y2              =   1050
               End
               Begin VB.Line Line13 
                  X1              =   4200
                  X2              =   4200
                  Y1              =   0
                  Y2              =   1050
               End
               Begin VB.Line Line20 
                  X1              =   5775
                  X2              =   5775
                  Y1              =   0
                  Y2              =   1050
               End
            End
            Begin VB.Frame Frame17 
               BackColor       =   &H00C0F0FF&
               BorderStyle     =   0  'None
               Caption         =   "Frame3"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   5580
               Left            =   0
               TabIndex        =   16
               Top             =   -1050
               Width           =   10515
               Begin VB.ListBox List3 
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
                  Left            =   -30
                  TabIndex        =   17
                  Top             =   1365
                  Width           =   9670
               End
               Begin VB.Line Line21 
                  X1              =   1155
                  X2              =   1155
                  Y1              =   1470
                  Y2              =   3255
               End
            End
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -74650
         ScaleHeight     =   375
         ScaleWidth      =   9720
         TabIndex        =   10
         Top             =   1155
         Width           =   9780
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Anotaciones - Novedades - Comentarios"
            Height          =   225
            Left            =   105
            TabIndex        =   11
            Top             =   105
            Width           =   9570
         End
      End
      Begin VB.TextBox Text28 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4215
         Left            =   -74650
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   1575
         Width           =   9780
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Cortar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   -69800
         TabIndex        =   9
         Top             =   5790
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Copiar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   -68615
         TabIndex        =   1
         Top             =   5790
         Width           =   1170
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Pegar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   -67445
         TabIndex        =   2
         Top             =   5790
         Width           =   1275
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Selecc Todo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   -66155
         TabIndex        =   3
         Top             =   5790
         Width           =   1290
      End
      Begin VB.CommandButton Command7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   -74650
         TabIndex        =   8
         Top             =   5790
         Width           =   4830
      End
      Begin VB.PictureBox Picture4 
         AutoRedraw      =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1590
         Left            =   -72690
         ScaleHeight     =   1530
         ScaleWidth      =   2055
         TabIndex        =   7
         Top             =   1875
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ord.Cheque:"
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
         Left            =   240
         TabIndex        =   76
         Top             =   3930
         Width           =   1395
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cat I-B:"
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
         Left            =   240
         TabIndex        =   73
         Top             =   2880
         Width           =   1380
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nro I-B:"
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
         Left            =   1200
         TabIndex        =   72
         Top             =   2865
         Width           =   1905
      End
      Begin VB.Label DETEXT 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   1830
         TabIndex        =   71
         Top             =   3150
         Width           =   3720
      End
      Begin VB.Line Line3 
         X1              =   -64875
         X2              =   -64875
         Y1              =   1155
         Y2              =   6000
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   120
         Left            =   3795
         TabIndex        =   6
         Top             =   4605
         Width           =   15
      End
      Begin VB.Label Label17 
         Caption         =   "Label8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   120
         Left            =   4215
         TabIndex        =   5
         Top             =   3660
         Width           =   15
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   10540
      ScaleHeight     =   8355
      ScaleWidth      =   1215
      TabIndex        =   26
      Top             =   0
      Width           =   1275
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   136
         Top             =   5880
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   137
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command11 
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
         Height          =   640
         Left            =   105
         Picture         =   "ACRE_MAQ04.frx":2426
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   2240
         Width           =   650
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Baja"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "ACRE_MAQ04.frx":2730
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Baja y Restauración de Cuentas de Proveedores"
         Top             =   2880
         Width           =   650
      End
      Begin VB.CommandButton Command21 
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "ACRE_MAQ04.frx":2A3A
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Alta de Nuevo Proveedor"
         Top             =   1600
         Width           =   650
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Sort"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "ACRE_MAQ04.frx":2B84
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Ordena Base de Datos de Proveedores y Gastos"
         Top             =   9500
         Width           =   650
      End
      Begin VB.CommandButton Command25 
         Caption         =   "Print"
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
         Left            =   105
         Picture         =   "ACRE_MAQ04.frx":2E8E
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Imprime Ficha del Proveedor"
         Top             =   5050
         Width           =   650
      End
      Begin VB.CommandButton Command13 
         Caption         =   "List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "ACRE_MAQ04.frx":34F8
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Listado General y Clasificado"
         Top             =   4410
         Width           =   650
      End
      Begin VB.CommandButton Command14 
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
         Height          =   640
         Left            =   105
         Picture         =   "ACRE_MAQ04.frx":3802
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Configuración - Tablas de Validación"
         Top             =   3675
         Width           =   650
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
         Height          =   640
         Left            =   105
         Picture         =   "ACRE_MAQ04.frx":3C44
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   240
         Width           =   650
      End
      Begin VB.CommandButton Command4 
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
         Left            =   105
         Picture         =   "ACRE_MAQ04.frx":3D8E
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Ayuda en Línea"
         Top             =   880
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "ACRE_MAQ04.frx":4098
         Top             =   6195
         Width           =   1515
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Controles no visibles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Left            =   1320
      TabIndex        =   35
      Top             =   8040
      Visible         =   0   'False
      Width           =   6105
      Begin VB.TextBox COTEXT 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   2100
         TabIndex        =   41
         Text            =   " "
         Top             =   1260
         Width           =   570
      End
      Begin VB.TextBox COTEXT 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   2100
         TabIndex        =   40
         Text            =   " "
         Top             =   420
         Width           =   570
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
         Height          =   315
         Index           =   3
         Left            =   2670
         TabIndex        =   39
         Top             =   420
         Width           =   175
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
         Height          =   315
         Index           =   5
         Left            =   2670
         TabIndex        =   38
         Top             =   1275
         Width           =   175
      End
      Begin VB.TextBox Text21 
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
         Left            =   1365
         TabIndex        =   37
         Text            =   " "
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox Text20 
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
         Left            =   4095
         TabIndex        =   36
         Text            =   " "
         Top             =   2040
         Width           =   1710
      End
      Begin VB.Label LTEXT23 
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
         Left            =   780
         TabIndex        =   55
         Top             =   3675
         Width           =   1500
      End
      Begin VB.Label label 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Disp:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   105
         TabIndex        =   54
         Top             =   3720
         Width           =   645
      End
      Begin VB.Label LTEXT37 
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
         Left            =   780
         TabIndex        =   53
         Top             =   3360
         Width           =   1500
      End
      Begin VB.Label label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "C.Cte:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   105
         TabIndex        =   52
         Top             =   3405
         Width           =   645
      End
      Begin VB.Label LTEXT36 
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
         Left            =   780
         TabIndex        =   51
         Top             =   3045
         Width           =   1500
      End
      Begin VB.Label label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "L.Cre:"
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
         Index           =   2
         Left            =   105
         TabIndex        =   50
         Top             =   3090
         Width           =   645
      End
      Begin VB.Label LTEXT35 
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
         Left            =   1830
         TabIndex        =   49
         Top             =   2520
         Width           =   1500
      End
      Begin VB.Label label 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cheq:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   1155
         TabIndex        =   48
         Top             =   2565
         Width           =   645
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   315
         TabIndex        =   47
         Top             =   1365
         Width           =   1710
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de Precios:"
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
         Left            =   840
         TabIndex        =   46
         Top             =   315
         Width           =   1185
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desc.Std:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   315
         TabIndex        =   45
         Top             =   1770
         Width           =   1710
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Límite de Crédito:"
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
         Left            =   2835
         TabIndex        =   44
         Top             =   1770
         Width           =   1905
      End
      Begin VB.Label DETEXT 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   2940
         TabIndex        =   43
         Top             =   420
         Width           =   2880
      End
      Begin VB.Label DETEXT 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   2940
         TabIndex        =   42
         Top             =   1260
         Width           =   2880
      End
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   12000
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu Archiv 
      Caption         =   "Archivos"
      Begin VB.Menu NueCli 
         Caption         =   "Nueva Cuenta"
      End
      Begin VB.Menu EditCli 
         Caption         =   "Abrir Existente"
      End
      Begin VB.Menu Baj_Cli 
         Caption         =   "Baja / Restauración Cuenta Activa"
      End
      Begin VB.Menu PLh2 
         Caption         =   "-"
      End
      Begin VB.Menu Print 
         Caption         =   "Imprimir Ficha"
      End
      Begin VB.Menu php31 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Terminar"
      End
   End
   Begin VB.Menu Edi 
      Caption         =   "Edición"
      Begin VB.Menu DeshCam 
         Caption         =   "Deshacer Cambios"
      End
      Begin VB.Menu PLh5 
         Caption         =   "-"
      End
      Begin VB.Menu Busc 
         Caption         =   "Buscar"
      End
   End
   Begin VB.Menu Cons 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu ConGen 
         Caption         =   "Consultas y Listados General"
         Visible         =   0   'False
      End
      Begin VB.Menu plh6 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu LiGe 
         Caption         =   "Listado General y Clasificado"
      End
      Begin VB.Menu php7 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu LidePo 
         Caption         =   "Listado de Direcciones Postales"
         Visible         =   0   'False
      End
      Begin VB.Menu LiDiMa 
         Caption         =   "Listado de Direcciones de Mails"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Conf 
      Caption         =   "Configuración"
      Begin VB.Menu Confgnr 
         Caption         =   "Configuración General"
      End
      Begin VB.Menu PLh8 
         Caption         =   "-"
      End
      Begin VB.Menu CondPag 
         Caption         =   "Condiciones de Pago"
      End
      Begin VB.Menu ProZonVta 
         Caption         =   "Provincias y Zonas de Venta"
      End
      Begin VB.Menu CuenCon 
         Caption         =   "Cuentas Contables"
         Visible         =   0   'False
      End
      Begin VB.Menu GruProv 
         Caption         =   "Grupo de Proveedores"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "ACRE_MAQ04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHEX$(16)
Dim YACAR%(8), MODIFIC%(8)
Dim GRABATODO%
Dim AGRECLI%
Dim RECONTANT$
Dim UPDATPRO%

Private Sub Baj_Cli_Click()
   Call Command24_Click
End Sub

Private Sub Busc_Click()
Call Command20_Click
End Sub



Private Sub Command10_Click()
   '
   Call GRAGENPRO
   Call GRANOTAPRO
   '
   Screen.MousePointer = 1
   RSLX$ = REGSEL$("INDIPROV")
   If Len(RSLX$) > 4 Then
     NCLIE% = CVI(Mid$(RSLX$, 63, 2))
     If NCLIE% > 0 Then
       NC% = NCLIE%
       Text1.TEXT = TRIM$(Str$(NC%))
     End If
   End If
   '
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call Command1_Click
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   '
   If ULTREG&("DARGAPRO") < 1 Then
     Call GENDAREGAPRO
   End If
   '
   NPRO% = Val(TRIM$(Text1.TEXT))
   If NPRO% > 0 Then
     If ECOARCH$("PROVED1", MKI$(NPRO%)) <> "" Then
       '
       If ULTREG&("TACRETG") < 1 Then
         Call GRAREG("TACRETG", Chr$(0) + "Exento ", 1)
         Call GRAREG("TACRETG", Chr$(1) + "Profesionales", 2)
         Call GRAREG("TACRETG", Chr$(1) + "Bienos de Uso y de Cambio", 3)
         Call SETULTREG("TACRETG", 3)
       End If
       '
       CATMAX% = ULTREG&("TACRETG")
       RECORD$ = FILTERGETRECORD$("DARGAPRO", 1, MKI$(NPRO%))
       Call REPLA(RECORD$, MKI$(NPRO%), 1, 2)
       CATIB1% = Asc(Mid$(RECORD$, 3, 1))
       If CATIB1% < 0 Then
         Call REPLA(RECORD$, Chr$(2), 3, 1)
         Call REPLA(RECORD$, Space$(32), 4, 32)
       End If
       '
10     OBX$ = OBJPLA$("CARDAREGAN", RECORD$)
       If OBX$ <> "" Then
         RECORD$ = OBX$
         CATIB1% = Asc(Mid$(RECORD$, 3, 1))
         If CATIB1% < 0 Or CATIB1% > CATMAX% Then
           Call MENSERR(24, "Mal Categoría de Ret. de Ganancias")
           GoTo 10
         End If
         If CATIB1% = 0 Then
           If TRIM$(Mid$(RECORD$, 4, 32)) = "" Then
             Call MENSERR(24, "Mandatorio Indicar Motivo de Exención")
             GoTo 10
           End If
         End If
         '
         If CATIB1% > 0 Then
           Call REPLA(RECORD$, Space$(32), 4, 32)
         End If
         '
         Call REPLA(RECORD$, MKI$(NPRO%), 1, 2)
         Call REPLA(RECORD$, USERNAME$, 36, 20)
         Call REPLA(RECORD$, MKI$(HOY(HOS$)), 56, 2)
         Call REPLA(RECORD$, MKI$(HORANUM(Time$)), 58, 2)
         Call FILTERPUTRECORD("DARGAPRO", 1, MKI$(NPRO%), RECORD$)
         Call CIERRARCH("DARGAPRO")
         '
       End If
     End If
   End If
   Command12.Enabled = True
End Sub

Private Sub Command22_Click()
   Static CTXX%
   Command22.Enabled = False
   '
   If EXISTE%(LUDAT$ + "DARIBPRO.RFS") < 1 Then
     Call COPYSTRU("DAPIBCLI", "DARIBPRO")
   End If
   '
   If ULTREG&("DARIBPRO") < 1 Then
     Call GENDARIBPRO
   End If
   '
   NCLIE% = Val(TRIM$(Text1.TEXT))
   If NCLIE% > 0 Then
     If ECOARCH$("PROVED1", MKI$(NCLIE%)) <> "" Then
       If TICUEPRO%((-1) * NCLIE%) = 0 Then ' solo cuentas corrientes
         '
         If ULTREG&("TACATIBP") < 5 Then
           Call GRAREG("TACATIBP", Chr$(0) + AJUSTI$("EXENTO", 31) + String$(32, 0), 1)
           Call GRAREG("TACATIBP", Chr$(1) + AJUSTI$("BIENES MUEBLES", 31) + String$(32, 0), 2)
           Call GRAREG("TACATIBP", Chr$(2) + AJUSTI$("LOCACIONES SERVICIOS", 31) + String$(32, 0), 3)
           Call GRAREG("TACATIBP", Chr$(3) + AJUSTI$("EVENTUALES", 31) + String$(32, 0), 4)
           Call GRAREG("TACATIBP", Chr$(4) + AJUSTI$("PARTIDAS ESPECIALES", 31) + String$(32, 0), 5)
           Call CIERRARCH("TACATIBP")
         End If
         '
         RECORD$ = FILTERGETRECORD$("DARIBPRO", 1, MKI$(NCLIE%))
         Call REPLA(RECORD$, MKI$(NCLIE%), 1, 2)
         '
10       OBX$ = OBJPLA$("CARDARETIB", RECORD$)
         If OBX$ <> "" Then
           RECORD$ = OBX$
           CATIB1% = Asc(Mid$(RECORD$, 3, 1))
           If ECOARCH$("TACATIBP", Chr$(CATIB1%)) = "" Then
             Call MENSERR(24, "Mal Categoría I.Brutos")
             GoTo 10
           End If
           If CATIB1% < 1 Then
             If TRIM$(Mid$(RECORD$, 9, 32)) = "" Then
               Call MENSERR(24, "Mandatorio Indicar Motivo de Exención")
               GoTo 10
             End If
           End If
           If CATIB1% > 0 Then
             Call REPLA(RECORD$, Space$(32), 9, 32)
           End If
           PORCONVE = CVS(Mid$(RECORD$, 5, 4))
           If PORCONVE < 0 Or PORCONVE > 100 Then
             Call MENSERR(24, "% Convenio Multilateral\Fuera de Rango")
             GoTo 10
           End If
           '
           Call REPLA(RECORD$, MKI$(NCLIE%), 1, 2)
           Call REPLA(RECORD$, USERNAME$, 41, 20)
           Call REPLA(RECORD$, MKI$(HOY(HOS$)), 61, 2)
           Call REPLA(RECORD$, MKI$(HORANUM(Time$)), 63, 2)
           Call FILTERPUTRECORD("DARIBPRO", 1, MKI$(NCLIE%), RECORD$)
           Call CIERRARCH("DARIBPRO")
           '
         End If
       End If
     End If
   End If
   Command22.Enabled = True
End Sub

Sub GENDARIBPRO()
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("PROVED1")
    REBLA$ = REGBLAN$("DARIBPRO")
    Call REPLA(REBLA$, USERNAME$, 41, 20)
    Call REPLA(REBLA$, MKI$(HOY(HOS$)), 61, 2)
    Call REPLA(REBLA$, MKI$(HORANUM(Time$)), 63, 2)
    '
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("PROVED1", U&)
      NCLIX% = CVI(Left$(XX$, 2))
      If NCLIX% > 0 Then
        If TICUEPRO%((-1) * NCLIX%) = 0 Then
          Call REPLA(RECORD$, MKI$(NCLIX%), 1, 2)
          Call REPLA(RECORD$, Chr$(1), 3, 1)
          Call REPLA(RECORD$, MKS$(100), 5, 4)
          Call FILTERPUTRECORD("DARIBPRO", 1, MKI$(NCLIX%), RECORD$)
        End If
      End If
    Next U&
    Call CIERRARCH("DARIBPRO")
    Screen.MousePointer = 1
    '
End Sub

Private Sub Command4_Click()
  '
  Call HELPONLINE("ACREMAQ")
  '
End Sub

Private Sub CondPag_Click()
SETPROMA04.Option1(6).Value = True
Call SETPROMA04.Command16_Click
End Sub

Private Sub Confgnr_Click()
Call Command14_Click
End Sub

Private Sub ConGen_Click()
Call Command13_Click
End Sub

Private Sub CtaCte_Click()

End Sub

Private Sub CtaGas_Click()
Option2.Value = True
Call Command21_Click
End Sub

Private Sub CtaInac_Click()
Call AMAPRLIS04.Command5_Click
End Sub

Private Sub CuenCon_Click()
  OPMASPR04.Option1(8).Value = True
  Call OPMASPR04.Command16_Click
End Sub

Private Sub DeshCam_Click()
Call Command2_Click
End Sub

Private Sub EditCli_Click()
   Call Command9_Click
End Sub

Private Sub Exit_Click()
   Call Command3_Click
End Sub

'
Private Sub Form_Load()
    '
    SSTab1.Tab = 0
    Call CENTRAFORM(Me)
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    ARCHEX$(0) = "PROVINC"
    ARCHEX$(1) = "CAPOSIV"
    ARCHEX$(2) = "CARETIB"
    ARCHEX$(3) = "LISTAS"
    ARCHEX$(4) = "CONPAG"
    ARCHEX$(5) = "VENDEDOR"
    ARCHEX$(6) = "ECUECON"
    ARCHEX$(7) = "GRUCOPRO" 'NUEVO
    'ARCHEX$(7) = "GRUCOCLI"
    '
    Call CARTAVALI
    '
    If ULTREG&("BKDATCLI") < 1 Then
      LU$ = LUDAT$
      If EXISTE%(LU$ + "DEUDOR1.BAK") > 0 Then
        If EXISTE%(LU$ + "DEUDOR2.BAK") > 0 Then
          Screen.MousePointer = 11
          JJ& = 0
          Dim D1 As String * 32
          Dim D2 As String * 128
          N1% = FreeFile
          Open LU$ + "DEUDOR1.BAK" For Random Access Read Write Shared As #N1% Len = 32
          N2% = FreeFile
          Open LU$ + "DEUDOR2.BAK" For Random Access Read Write Shared As #N2% Len = 128
          For K& = 1 To LOF(N1%) / 32
            Get #N1%, K&, D1$
            NC% = CVI(Left$(D1$, 2))
            If NC% > 0 Then
              Get #N2%, K&, D2$
              D3$ = D1$ + D2$
              JJ& = JJ& + 1
              Call GRAREG("BKDATCLI", D3$, JJ&)
            End If
          Next K&
          Call SETULTREG("BKDATCLI", JJ&)
          Call CIERRARCH("BKDATCLI")
          Screen.MousePointer = 1
        End If
      End If
    End If
    '
End Sub
'
Private Sub Form_Activate()
   WindowState = 0
   OPENFORMS = DoEvents
End Sub
'
Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub
'
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub BOTSEL_Click(Index As Integer)
   Call SELECHO(ARCHEX$(Index))
   VDEVU$ = VALACT1$(ARCHEX$(Index))
   If TRIM$(VDEVU$) <> "" Then
      COTEXT(Index).TEXT = VDEVU$
   End If
End Sub

Private Sub Command1_Click()
   GRABATODO% = 1
   Call GRAGENPRO
   Call GRANOTAPRO
   GRABATODO% = 0
End Sub

Private Sub Command13_Click()
    '
    Command13.Enabled = False
    Call GELISPRO
    Command13.Enabled = True
    '
End Sub
'
Sub GELISPRO()
    '
    Screen.MousePointer = 11
    '
    Call BLANARCH("LIPROVE")
    '
    Call ABREPROVE
    MODOFILT% = 2
    MODORDE% = 2
    'CONSULTA SQL
    SQLX$ = "SELECT * FROM Proved12 "
    If MODFILT% = 0 Then SQLX$ = SQLX$ + "WHERE Tcpr_Cli= " & MODFILT% & ""
    If MODFILT% = 1 Then SQLX$ = SQLX$ + "AND Tcpr_Cli= " & MODFILT% & ""
    If MODFILT% = 2 Then SQLX$ = SQLX$
    'Agrego Clausula order by a la consulta SQL
    Select Case MODORDE%
        Case 1
        SQLX$ = SQLX$ + " ORDER BY Nume_Cli ASC "
        Case 2
        SQLX$ = SQLX$ + " ORDER BY Raso_Cli ASC "
        Case 3
        SQLX$ = SQLX$ + " ORDER BY Cpos_Cli ASC "
    End Select
    Set PROVTEMP = Badaprov.OpenRecordset(SQLX$, dbOpenSnapshot)
    With PROVTEMP
       On Error Resume Next
       .MoveLast
       'Valida que la tabla no este vacia
       If Err Then
         Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
         Screen.MousePointer = 1
         On Error GoTo 0
         GoTo 90
       End If
       On Error GoTo 0
       FIN& = .RecordCount
    End With
       '
    With PROVTEMP
        .MoveFirst
        U& = 0
        Do While (.EOF = False)
            U& = U& + 1
            Call TRACE(20, U&, FIN&)
            NP% = !Nume_Cli
            NC% = (-1) * NP%
             '
            Y$ = REGBLAN$("LIPROVE")
            Call REPLA(Y$, MKI$(NP%), 1, 2)
            Call REPLA(Y$, RASOCLI$(NC%), 3, 32)
            Call REPLA(Y$, TEL1CLI$(NC%), 35, 16)
            Call REPLA(Y$, CUITCLI$(NC%), 51, 16)
            Call REPLA(Y$, POSIVAC$(NC%), 67, 10)
            CUECOX% = 0
            Call REPLA(Y$, MKI$(CUECOX%), 77, 2)
            Call REPLA(Y$, " ", 79, 32)
            Call REPLA(Y$, CUEMADRE$(NC%), 111, 16)
            ULI& = ULI& + 1
            Call GRAREG("LIPROVE", Y$, ULI&)
            '
            Y$ = REGBLAN$("LIPROVE")
            Call REPLA(Y$, MKI$(NP%), 1, 2)
            Call REPLA(Y$, DOMICLI$(NC%), 3, 32)
            'Call REPLA(Y$, FAXCLI$(NC%), 35, 16)
            Call REPLA(Y$, NIBRCLI$(NC%), 51, 16)
            CTRX$ = TRIM$(CATIBRU$(NC%))
            CATIBX$ = "No Def.": If CTRX$ = "0" Then CATIBX$ = "No Resp."
                                 If CTRX$ = "1" Then CATIBX$ = "Cont.Puro"
                                 If CTRX$ = "2" Then CATIBX$ = "C.Multil."
            Call REPLA(Y$, CATIBX$, 67, 10)
            CPAG% = CPAGCLI%(NC%)
            Call REPLA(Y$, MKI$(CPAG%), 77, 2)
            Call REPLA(Y$, DECONPA$(CPAG%), 79, 32)
            ULI& = ULI& + 1
            Call GRAREG("LIPROVE", Y$, ULI&)
            '
            Y$ = REGBLAN$("LIPROVE")
            Call REPLA(Y$, MKI$(NP%), 1, 2)
            DOMI$ = TRIM$(CPOSCLI$(NC%))
            If DOMI$ <> "" Then DOMI$ = DOMI$ + " - "
            DOMI$ = DOMI$ + LOCACLI$(NC%)
            Call REPLA(Y$, DOMI$, 3, 32)
            'Call REPLA(Y$, EMAICLI$(NC%), 35, 16)
            ULI& = ULI& + 1
            Call GRAREG("LIPROVE", Y$, ULI&)
            '
            Y$ = REGBLAN$("LIPROVE")
            ULI& = ULI& + 1
            Call GRAREG("LIPROVE", Y$, ULI&)
            '
            .MoveNext
        Loop
    End With
    '
    Call SETULTREG("LIPROVE", ULI&)
    Call CIERRARCH("LIPROVE")
    Screen.MousePointer = 1
    Call FINAL("*LIPROV1")
    '
90 End Sub
Sub Command14_Click()
   Command14.Enabled = False
   SETPROMA04.Show 1
   Command14.Enabled = True
End Sub

Private Sub COMMAND15_CLICK()
   Command15.Enabled = False
   Call MENSERR(24, "No Disponible en la Presente Versión.\ \Los Ordenamientos son Automáticos\al Generar los Listados.")
   Command15.Enabled = True
End Sub

Private Sub Command17_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Clipboard.SetText Mid$(Text28.TEXT, Text28.SelStart + 1, Text28.SelLength)
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub

Private Sub Command2_Click()
   Call GRAGENPRO
   Call GRANOTAPRO
   For Index = 0 To 8: YACAR%(Index) = 0: Next Index
   Call CARDATPROV
End Sub

Private Sub Command20_Click()
   Static ARGU$
   Command20.Enabled = False
10 ARGU$ = UCase$(TRIM$(InputBox$("Buscar", "Búsqueda Proveedor por Texto", ARGU$)))
   If ARGU$ <> "" Then
     Screen.MousePointer = 11
     JJ& = 0
     Call COPYSTRU(" PROVED1", "SELECLI")
     '
     Call ABREPROVE
     On Error Resume Next
     With Proved12
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
           For KKI% = 1 To .Fields.Count
             On Error Resume Next
             VATEX$ = ""
             VATEX$ = .Fields(KKI% - 1).Value
             On Error GoTo 0
             If InStr(VATEX$, ARGU$) > 0 Then
               NCLIE% = !Nume_Cli
               RACLI$ = !Raso_Cli
               XX$ = MKI$(NCLIE%) + RACLI$
               JJ& = JJ& + 1
               Call GRAREG("!SELECLI", XX$, JJ&)
               Exit For
             End If
           Next KKI%
         .MoveNext
         Loop
       End If
     End With
     Call SETULTREG("!SELECLI", JJ&)
     '
     Screen.MousePointer = 1
     If JJ& < 1 Then
         Call MENSERR(24, "No se Encontraron Coincidencias")
         GoTo 10
       ElseIf JJ& = 1 Then
         Text1 = TRIM$(Str$(CVI(Left$(REGLEIDO$("!SELECLI", 1), 2))))
       Else
         Call FRESHECHO("!SELECLI")
         Call SELECHO("!SELECLI")
         NCLIX$ = VALACT1$("!SELECLI")
         If NCLIX$ <> "" Then
           Text1 = NCLIX$
         End If
     End If
   End If
   Command20.Enabled = True
End Sub
 Private Sub Command21_Click()
   '
   DRXX% = DERACCE%("ALTAPRO", "Alta de Nuevo Proveedor")
   If DRXX% < 2 Then
     If DRXX% = 1 Then
       Call MENSERR(24, "Usuario sin Derecho de Escritura\para la Aplicación Solicitada.\  \Consulte a su Supervisor del Sistema.")
     End If
     Exit Sub
   End If
   '
   Call ABREPROVE
   '
   Command21.Enabled = False
   Call GRAGENPRO
   Call GRANOTAPRO
   '
   Screen.MousePointer = 11
   PRONUCLI& = Val(CONTROL$("", "PRONUPRO"))
   If PRONUCLI& > 9999 Then PRONUCLI& = 0
   NROPRO& = PRONUCLI&
   If NROPRO& < 0 Then NROPRO& = 0
   '
   Set Tranprov = Badaprov.OpenRecordset("SELECT Nume_Cli FROM Proved12 WHERE Nume_Cli >= " & PRONUPRO& & " ORDER BY Nume_Cli DESC;", dbOpenSnapshot)
   On Error Resume Next
   '
3  NROPRO& = NROPRO& + 1
   Tranprov.FindFirst "Nume_Cli = " & NROPRO&
   If Tranprov.NoMatch = False Then GoTo 3
   '
   NCLIE% = NROPRO&
   VDEF$ = String$(128, 0)
   Call REPLA(VDEF$, MKI$(NCLIE%), 13, 2)
   Call REPLA(VDEF$, Space$(62), 15, 62)
   Call REPLA(VDEF$, Chr$(1), 77, 1)
   Call REPLA(VDEF$, Space$(15), 78, 15)
   '
10 Screen.MousePointer = 1
   If Option1.Value = True Then
       OBX$ = OBJPLA$("ALTAPROVACC", VDEF$)
     Else
       Call REPLA(VDEF$, Chr$(0), 77, 1)   ' IVA Consumidor Final
       OBX$ = OBJPLA$("ALTACUEGAST", VDEF$)
   End If
   If OBX$ = "" Then GoTo 99
   VDEF$ = OBX$
   '
   COCLINU$ = TRIM$(Left$(OBX$, 12))
   If COCLINU$ = "" Then
     Call MENSERR(24, "Codigo de Cuenta no Válido")
     GoTo 10
   End If
   '
   NCLINUE% = CVI(Mid$(OBX$, 13, 2))
   If NCLINUE% < 1 Or NCLINUE% > 9999 Then
      Call MENSERR(24, "Número de Cuenta no Válido.")
      GoTo 10
   End If
   '
   Set Tranprov = Badaprov.OpenRecordset("SELECT Codi_Cli FROM Proved12;", dbOpenSnapshot)
   On Error Resume Next
   Tranprov.FindFirst "Codi_Cli = '" & COCLINU$ & "'"
   If Tranprov.NoMatch = False Then
      Call MENSERR(24, "Ingreso no Válido.\Código de Cuenta Pre-existente.")
      GoTo 10
   End If
   '
   Set Tranprov = Badaprov.OpenRecordset("SELECT Nume_Cli FROM Proved12;", dbOpenSnapshot)
   On Error Resume Next
   Tranprov.FindFirst "Nume_Cli = " & NCLINUE
   If Tranprov.NoMatch = False Then
      Call MENSERR(24, "Ingreso no Válido.\Número de Cuenta Pre-existente.")
      GoTo 10
   End If
   '
   DCLINUE$ = TRIM$(Mid$(OBX$, 15, 62))
   If Len(DCLINUE$) < 1 Then
      Call MENSERR(24, "Ingreso no Válido.\Falta Razón Social del Proveedor")
      GoTo 10
   End If
   '
   PIVCLI% = Asc(Mid$(OBX$, 77, 1))
   If PIVCLI% < 0 Or PIVCLI% > 7 Then
      Call MENSERR(24, "Ingreso no Válido.\Posición IVA no Figura en Tablas.")
      GoTo 10
   End If
   '
   CUTT$ = TRIM$(Mid$(OBX$, 78, 15))
   If VALICUIT%(CUTT$) <> 1 Then
     If PIVCLI% <> 0 Then
       If PIVCLI% <> 5 And PIVCLI% <> 7 Then
         Call MENSERR(24, "Ingreso no Válido.\Número de C.U.I.T. Incorrecto.")
         GoTo 10
       End If
     End If
   End If
   '
   If CUTT$ <> "" Then
     Set Tranclie = Badaprov.OpenRecordset("SELECT Nume_Cli FROM Proved12 WHERE Cuit_Cli = '" & CUTT$ & "';", dbOpenSnapshot)
     CAREI% = 0
     On Error Resume Next
     Tranclie.MoveLast
     CAREI% = Tranclie.RecordCount
     '
     If CAREI% > 0 Then
       If Tranclie!Nume_Cli > 0 Then
         NCLIX% = (-1) * Tranclie!Nume_Cli
         TTALTER$ = "Existe Registrada una Cuenta con el Mismo C.U.I.T.\  \'" + TRIM$(RASOCLI$(NCLIX%)) + "' (" + TRIM$(Str$(Abs(NCLIX%))) + ")"
         If COMALTER%(TTALTER$ + "\\Cancelar\Registrar Igual") <> 2 Then
           GoTo 10
         End If
       End If
     End If
   End If
   '
   Screen.MousePointer = 11
   TICPRO% = 0
   If Option2.Value = True Then
     TICPRO% = 1
     PIVCLI% = 0
     CUTT$ = ""
   End If
   Screen.MousePointer = 11
   
20 Proved12.Close
   On Error Resume Next
   Call ABREPROVE
   Proved12.MoveFirst
   If Err < 1 Then
     On Error GoTo 0
     ABUS$ = "Nume_Cli = " & NCLINUE%
     Proved12.FindFirst ABUS$
     If Proved12.NoMatch = False Then
       NCLINUE% = NCLINUE% + 1
       GoTo 20
     End If
   End If
   '
   With Proved12
      .AddNew
      !Nume_Cli = TRIM$(Str$(NCLINUE%))
      !Raso_Cli = DCLINUE$
      !Codi_Cli = COCLINU$
      !Cuit_Cli = CUTT$
      !Piva_Cli = PIVCLI%
      !Stat_Cli = 1
      !Tcpr_Cli = TICPRO%
      '
      On Error Resume Next
      .Update
      If Err Then
         Call MENSERR(24, "Alta no Procesada Correctamente")
      End If
      On Error GoTo 0
      .Bookmark = .LastModified
      .Close
      UPDATPRO% = 1
      Call GENECOPROV
   End With
   '
   AGRECLI% = 1
   Text1.TEXT = TRIM$(Str$(NCLINUE%))
   SSTab1.Tab = 0
   Call CARDATPROV
   Call CARTAVALI
   '
   For Index = 0 To 7
     If Index <> 1 Then
       VAPRE$ = TRIM$(CONTROL$("AMAPRO", ARCHE$(Index)))
       COTEXT(Index) = "": DETEXT(Index) = ""
       If VAPRE$ <> "" Then
         TV11% = TIPFIELD(ARCHE$(Index), 1)
         LF11% = LENFIELD(ARCHE$(Index), 1)
         ARGU$ = CBIN$(VAPRE$, TV11%, LF11%)
         ECO$ = TRIM$(ECOARCH$(ARCHE$(Index), ARGU$))
         If ECO$ <> "" Then
             COTEXT(Index).TEXT = VAPRE$
             DETEXT(Index).Caption = ECO$
           Else
             COTEXT(Index).TEXT = ""
             DETEXT(Index).Caption = ""
         End If
       End If
     End If
   Next Index
   '
   Text3 = DCLINUE$
   COTEXT(1) = TRIM$(Str$(PIVCLI%))
   Text33 = TRIM$(CUTT$)
   '
   Screen.MousePointer = 1
   '
99 Command21.Enabled = True
   '
End Sub

Private Sub Command24_Click()
   '
   Command24.Enabled = False
   Screen.MousePointer = 11
   '
   Call GRAGENPRO
   Call GRANOTAPRO
   '
   If DERACCE%("BAJAPRO", "Baja de Cuenta de Proveedor") < 2 Then
     GoTo 99
   End If
   '
   NCLIE& = (-1) * Abs(Val(TRIM$(Text1.TEXT)))
   If Abs(NCLIE&) < 1 Or REGICLI(NCLIE&) < 1 Then
      RESP1% = 2
      GoTo 90
   End If
   '
   If RESP1% = 0 Then
    RESP1% = ALTERNA%("Dar de Baja Cuenta de Proveedor.\Restauración de Cuenta de Proveedor.")
    If RESP1% < 1 Or RESP1% > 2 Then
        GoTo 99
    End If
   End If
   '
   If RESP1% = 1 Then 'SI QUIERE ANULAR
    If Val(Text1.TEXT) > 0 Then
     If Val(Text1.TEXT) <= 32767 Then
       NCLIE& = Val(Text1.TEXT)
       Call ABREPROVE
       With Proved12
       On Error Resume Next
       .MoveFirst
       .FindFirst "Nume_Cli = " & NCLIE&
       If .NoMatch = False Then
         '
         On Error GoTo 0
         NC% = NCLIE
'         '
'         If Abs(SALDCLI#((-1) * NC%)) >= 0.005 Then
'           Call MENSERR(24, "Baja Imposible - Cuenta con Saldo.")
'           GoTo 99
'         End If
'         'WAR 07/04/05
'         Call ABRECABAN
'         With CajaBanc
'            .FindFirst "NuProve = " & NC% & " "
'            If .NoMatch = False Then
'              Call MENSERR(24, "Baja Imposible - Cuenta con Movimiento.")
'              GoTo 99
'            End If
'         End With
'         '
''         RFF$ = RECFILT$("CUECORR", 4, MKI$(NC%))
''         If RFF$ <> "" Then
''           Call MENSERR(24, "Baja Imposible - Cuenta con Movimiento.")
''           GoTo 99
''         End If
         '
         TTXX$ = "La Baja de un Proveedor\Puede Producir Inconsistencias en la Base de Datos\en Caso de que la Cuenta Registre\Facturación, Pagos o Compras."
         TTXX$ = TTXX$ + "\  \Puede Ahora Cancelar la Operación\o Confirmar la Baja de la Cuenta:\'" + TRIM$(Text3.TEXT) + "'"
         If COMALTER%(TTXX$ + "\\Cancelar Operación\Confirmar Baja") = 2 Then
           '
           .Edit
           !Stat_Cli = (-1)
           .Update
           'GENERO EL INDICE DEL PROVE ANULADOS
           Call GENPROVANU
           '
           YACAR%(1) = 1
           Text28.TEXT = ""
           MODIFIC%(1) = 1
           GRABATODO% = 1
           Call GRANOTAPRO
           '
           UPDATPRO% = 1
           Call GENECOPROV
           Text1.TEXT = ""
           Call BLANFORMU
           For KKI% = 0 To 4: MODIFIC%(KKI%) = 0: YACAR%(KKI%) = 0: Next KKI%
           GRABATODO% = 0
           Call BAJALDISCO
           '
         End If
         '
       End If
       End With
     End If
    End If
   End If
   '
90 If RESP1% = 2 Then 'SI QUIERE RESTAURAR
     RPA$ = ""
     Call BLANARCH("INDIPROA")
     Call GENPROVANU
     Call FRESHECHO("INDIPROA")
     RPA$ = REGSEL$("INDIPROA") 'PROV ANULADOS
     If Len(RPA$) > 4 Then
       NCLIE1% = CVI(Mid$(RPA$, 63, 2))
       If NCLIE1% > 0 Then
         NC% = NCLIE1%
         Text1.TEXT = TRIM$(Str$(NC%))
         Call Command1_Click
       End If
     End If
   End If
   '
99 Screen.MousePointer = 1
   Command24.Enabled = True
   '
End Sub

Private Sub Command25_Click()
   PRINTFORM
End Sub

Private Sub Command3_Click()
   '
   Call GRAGENPRO
   Call GRANOTAPRO
   Call CARDATPROV
   '
   Call CIERRARCH("*.*")
   If AGRECLI% > 0 Then
      Call SORTCLI(1)
      AGRECLI% = 0
   End If
   Call BAJALDISCO
   '
   Hide
   'If APLINVO$ <> "" Then
   '    Call FINAL("")
   '  Else
   '    Hide
   'End If
   '
End Sub

Private Sub Command5_Click()
   Text28.SelStart = 0
   Text28.SelLength = Len(Text28.TEXT)
   Text28.SetFocus
End Sub

Private Sub Command6_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Clipboard.GetText + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub

Private Sub Command8_Click()
   Clipboard.SetText Mid$(Text28.TEXT, Text28.SelStart + 1, Text28.SelLength)
End Sub

Private Sub Command9_Click()
   '
   Call GRAGENPRO
   Call GRANOTAPRO
   '
   Screen.MousePointer = 1
   Call SELECHO("PROVED1")
   NCLIE% = Val(TRIM$(VALACT1$("PROVED1")))
   If NCLIE% < 1 Then Exit Sub
   '
   NC% = NCLIE%
   Text1.TEXT = TRIM$(Str$(NC%))
   '
End Sub

Private Sub COTEXT_Change(Index As Integer)
   '
   ECO$ = ""
   If TRIM$(COTEXT(Index).TEXT) <> "" Then
      TV11% = TIPFIELD(ARCHEX$(Index), 1)
      LF11% = LENFIELD(ARCHEX$(Index), 1)
      ARGU$ = CBIN$(TRIM$(COTEXT(Index).TEXT), TV11%, LF11%)
      ECO$ = TRIM$(ECOARCH$(ARCHEX$(Index), ARGU$))
   End If
   '
   If ECO$ <> "" Then
        DETEXT(Index).Caption = ECO$
      Else
        DETEXT(Index).Caption = "No Definido"
   End If
   '
   MODIFIC%(0) = 1
End Sub

Private Sub COTEXT_DblClick(Index As Integer)
   Call SELECHO(ARCHEX$(Index))
   VDEVU$ = VALACT1$(ARCHEX$(Index))
   If TRIM$(VDEVU$) <> "" Then
      COTEXT(Index).TEXT = VDEVU$
   End If
End Sub

Sub CARDATPROV()
    '
    NCLIE% = (-1) * Val(TRIM$(Text1.TEXT))
    STCL% = STATCLI%(NCLIE)
    If NCLIE >= 0 Or STCL% < 1 Then
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    Text3.TEXT = RASOCLI$(NCLIE)
    Text24.TEXT = TEL2CLI$(NCLIE)
    Text6 = CODICLI$(NCLIE)
    MODIFIC(0) = 0
    '
    Select Case SSTab1.Tab
      Case 0
        Call CARGENPRO
      Case 1
        Call CARNOTAPRO
      Case 2
     '   Call CARESCUENTA
    End Select
    Text1.SelStart = Len(Text1)
    Screen.MousePointer = 1
    '
End Sub
'
Sub CARGENPRO()
    '
    If YACAR%(0) > 0 Then Exit Sub
    NC% = (-1) * Abs(Val(TRIM$(Text1.TEXT)))
    '
    Text2 = CONDICOM$(NC%)
    Text6 = CODICLI$(NC%)
    Text25.TEXT = DOMICLI$(NC%)
    Text25.Refresh
    Text26.TEXT = CPOSCLI$(NC%)
    Text26.Refresh
    Text27.TEXT = LOCACLI$(NC%)
    Text27.Refresh
    If LIMICRE(NC%) < 0 Then
        Text20.TEXT = " Suspendido"
      Else
        Text20.TEXT = CSTRING$(MKS$(LIMICRE(NC%)), 3, 12, 2, 2)
    End If
    Text20.Refresh
    'Text23.TEXT = NPROCLI$(NC%)
    'Text23.Refresh
    Text29.TEXT = PAISCLI$(NC%)
    Text29.Refresh
    Text30.TEXT = TEL1CLI$(NC%)
    Text30.Refresh
    Text31.TEXT = FAXCLI$(NC%)
    Text31.Refresh
    Text32.TEXT = EMAILCLI$(NC%)
    Text32.Refresh
    Text19.TEXT = CONTACTCLI$(NC%)
    Text19.Refresh
    Text36.TEXT = HORARCLI$(NC%)
    Text36.Refresh
    Text33.TEXT = CUITCLI$(NC%)
    Text33.Refresh
    Text34.TEXT = NIBRCLI$(NC%)
    Text34.Refresh
    Text21.TEXT = TRIM$(DESCUCLI$(NC%))
    Text21.Refresh
    Text4.TEXT = TRIM$(ORDECHEQ(NC%))
    Text4.ForeColor = vbBlack
          If Text4 <> RASOCLI(NC%) Then Text4.ForeColor = vbRed
    Text4.Refresh
    '
    If ULTREG&("CAPOSIV") < 8 Then
      IQ& = 0
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(0) + AJUSTI$("IVA Consumidor Final            ", 31) + AJUSTI$("Cons.Fin", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(1) + AJUSTI$("IVA Responsable Inscripto   ", 31) + AJUSTI$("Resp.Ins", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(2) + AJUSTI$("IVA Responsable no Inscripto", 31) + AJUSTI$("R.No Ins", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(3) + AJUSTI$("IVA No Responsable          ", 31) + AJUSTI$("No Resp.", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(4) + AJUSTI$("IVA Responsable Exento      ", 31) + AJUSTI$("Exento  ", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(5) + AJUSTI$("IVA Operac.Comercio Exterior", 31) + AJUSTI$("Com.Ext. ", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(6) + AJUSTI$("IVA Contrib.Monotributo     ", 31) + AJUSTI$("Monotrib", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(7) + AJUSTI$("IVA Exento Zona Promovida   ", 31) + AJUSTI$("Monotrib", 8) + String$(24, 0), IQ&)
      Call SETULTREG("CAPOSIV", IQ&)
      Call CIERRARCH("CAPOSIV")
    End If
    '
    COTEXT(0).TEXT = TRIM$(PROVCLI$(NC%))
    COTEXT(0).Refresh
    COTEXT(1).TEXT = TRIM$(Str$(PIVACLI%(NC%)))
    COTEXT(1).Refresh
    COTEXT(2).TEXT = TRIM$(CATIBRU$(NC%))
    COTEXT(2).Refresh
    COTEXT(3).TEXT = TRIM$(Str$(LIPRCLI%(NC%)))
    COTEXT(3).Refresh
    COTEXT(4).TEXT = TRIM$(Str$(CPAGCLI%(NC%)))
    COTEXT(4).Refresh
    COTEXT(5).TEXT = TRIM$(Str$(VENDCLI%(NC%)))
    COTEXT(5).Refresh
    COTEXT(6).TEXT = TRIM$(CUEMADRE$(NC%))
    COTEXT(6).Refresh
    COTEXT(7).TEXT = TRIM$(Str$(GRUCOCLI%(NC%)))
    COTEXT(7).Refresh
    '
    If TICUEPRO%(NC%) = 0 Then ' cuenta corriente
         Option1.Value = True
       Else
         Option2.Value = True
    End If
    '
    YACAR%(0) = 1: MODIFIC%(0) = 0
    '
End Sub
'
Sub GRAGENPRO()
    '
    NCLIE& = (-1) * Abs(Val(TRIM$(Text1.TEXT)))
    If Abs(NCLIE&) < 1 Or REGICLI(NCLIE&) < 1 Then
      If GRABATODO% = 1 Then
        Call MENSERR(24, "Imposible Registrar !!!\  \Número de Proveedor\Inexistente o no Válido.")
      End If
      Screen.MousePointer = 1
      Exit Sub
    End If
    '
    If Abs(STATCLI%(NCLIE)) < 1 Then
      If GRABATODO% = 1 Then
        Call MENSERR(24, "Imposible Registrar !!!\Numero de Proveedor Inexistente")
        Screen.MousePointer = 1
        Exit Sub
      End If
    End If
    '
    If STATCLI(NCLIE&) < 1 Then
      If GRABATODO% = 1 Then
        If COMALTER%("El Proveedor Elegido Figura\Marcado Como 'Dado de Baja'.\  \Desea Re-Activar Esta Cuenta ?\\No, Dejar de Baja\Si, Re-Activar") = 2 Then
          Call ABREPROVE
          With Proved12
            On Error Resume Next
            .FindFirst "Nume_Cli = " & Abs(NCLIE&)
            If .NoMatch = False Then
             .Edit
             !Stat_Cli = 1
             .Update
             Call GENPROVANU 'aca lo agregue para que actualice
             'el indice de prov. anulados.
            End If
          End With
          UPDATPRO% = 1
          Call GENECOPROV
          Call CARDATPROV
        End If
      End If
      Screen.MousePointer = 1
      Exit Sub
    End If
    '
    If MODIFIC%(0) < 1 Then
      If GRABATODO% <> 1 Then
        Screen.MousePointer = 1
        Exit Sub
      End If
    End If
    '
    If GRABATODO% = 1 Then
      If SSTab1.Tab <> 0 Then
        Exit Sub
      End If
    End If
    '
    If DERACCE%("MODATPRO", "Modificar Datos de Proveedores") < 2 Then
       Exit Sub
    End If
    '
    If GRABATODO% <> 1 Then
      On Error Resume Next
      Command9.SetFocus
      On Error GoTo 0
      If TRIM$(Text3) = "" Then
        Screen.MousePointer = 11
        Exit Sub
      End If
      If COMALTER%("Se han Modificado Datos del Proveedor\'" + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
        YACAR%(0) = 0
        Screen.MousePointer = 11
        Text3.TEXT = RASOCLI$(NCLIE)
        Text24.TEXT = TEL2CLI$(NCLIE)
        MODIFIC(0) = 0
        Call CARGENPRO
        Screen.MousePointer = 1
        Exit Sub
      End If
    End If
    '
    Screen.MousePointer = 11
    '
    CUIT$ = TRIM$(Text33.TEXT)
    If CUIT$ <> "" Then
      PIVA% = VALIMI(COTEXT(1).TEXT, 0, 7)
      If PIVA% <> 5 Then  ' EXCLUYE COMERCIO EXTERIOR
        If VALICUIT%(CUIT$) <> 1 Then
          Screen.MousePointer = 1
          Call MENSERR(24, "Imposible Registrar !!!\Número de C.U.I.T. no Válido")
          SSTab1.Tab = 0
          On Error Resume Next
          Text33.SetFocus
          On Error GoTo 0
          Screen.MousePointer = 1
          Exit Sub
        End If
      End If
    End If
    '
    If NCLIE = 0 Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Imposible Registrar Cuenta Número '0'")
        Exit Sub
    End If
    '
    If TRIM$(Text6) = "" Then
      Call MENSERR(24, "Imposible Registrar !!!\Falta Código de Cuenta")
      Screen.MousePointer = 1
      Exit Sub
    End If
    '
    NPRII% = NUMEPRO%(Text6)
    If NPRII% > 0 Then
      If NPRII% <> Abs(NCLIE) Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Imposible Registrar !!!\Codigo de Cuenta Repetido o no Válido.")
        Exit Sub
      End If
    End If
    '
    If TRIM$(Text3.TEXT) = "" Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Imposible Registrar !!!\Falta Razón Social de Proveedor / Cuenta.")
        Exit Sub
    End If
    '
    Call ABREPROVE
    With Proved12
      On Error Resume Next
      .MoveFirst
      .FindFirst "Nume_Cli = " & Abs(NCLIE&)
      If .NoMatch = False Then GoTo 10
      On Error GoTo 0
      Call MENSERR(24, "Imposible Registrar")
      Exit Sub
    End With
    '
10  On Error GoTo 0
    With Proved12
      .Edit
        If TRIM$(Text3) <> TRIM$(!Raso_Cli) Or UCase$(Text6) <> !Codi_Cli Then UPDATPRO% = 1
      !Codi_Cli = Left$(UCase$(Text6), 12)
      !Raso_Cli = Left$(Text3, 64)
      !Stat_Cli = 1
      !Domi_Cli = Left$(Text25, 64)
      !Cpos_Cli = Left$(Text26, 16)
      !Loca_Cli = Left$(Text27, 48)
      !Tele_Cli = Left$(Text30, 32)
      !Fant_Cli = Left$(Text24, 32)
      !Cuit_Cli = Left$(TRIM$(CUIT$), 24)
      !Nibr_Cli = Left$(TRIM$(Text34), 24)
      !Caib_Cli = Left$(TRIM$(COTEXT(2)), 2)
      !Prov_Cli = Left$(TRIM$(COTEXT(0)), 2)
      '
      !Ord_Cheq = Left$(Text4, 48) 'ORDEN DEL CHEQUE
      '
      PIVA% = VALIMI(COTEXT(1).TEXT, 0, 7)
      !Piva_Cli = PIVA%
      '
      CATIB% = VALIMI(COTEXT(2).TEXT, 0, 2)
      !Caib_Cli = CATIB%
      '
      COPAGI% = VALIMI(COTEXT(4).TEXT, 0, 255)
      If COPAGI% > 0 Then
        If ECOARCH("CONPAG", Chr$(COPAGI%)) = "" Then
          Call MENSERR(24, "Condición de Pago no Válida")
          COPAGI% = 0
        End If
      End If
      !Cpag_Cli = COPAGI%
      !CCom_Cli = RTrim$(Text2)
     '
      CUECO% = VALIMI(COTEXT(7).TEXT, 0, 32767)
      If ECOARCH$("GRUCOPRO", MKI$(CUECO%)) = "" Then
        CUECO% = 0
      End If
      !GrCo_Cli = CUECO%
      '
      CCX% = CUEINT%(COTEXT(6).TEXT)
      !CuMa_Cli = CCX%
      '
      DCC$ = TRIM$(Text21.TEXT)
      !Dsta_Cli = Left$(TRIM$(Text21), 32)
      '
      !Pais_Cli = Left$(Text29, 48)
      !Faxi_Cli = Left$(Text31, 32)
      !Mail_Cli = Left$(Text32, 64)
      !Ctac_Cli = Left$(Text19, 48)
      !Prvd_Cli = Left$(Text23, 24)
      !Hora_Cli = Left$(Text36, 32)
      '
      TICPRO% = 0
      If Option2.Value = True Then TICPRO% = 1
      !Tcpr_Cli = TICPRO%
      '
      .Update
    End With
    '
    Proved12.Close
    Call GENECOPROV
    '
    MODIFIC%(0) = 0
    Screen.MousePointer = 1
    '
End Sub
'
Sub CARNOTAPRO()
   '
   If YACAR%(1) > 0 Then Exit Sub
   NCLIE% = Val(Text1.TEXT)
   Text28 = ""
   '
   If NCLIE% > 0 Then
     If REGICLI((-1) * NCLIE%) > 0 Then
       Call ABREPROVE
       Set ClienTemp = Badaprov.OpenRecordset("SELECT * FROM Proved12 WHERE Nume_Cli = " & Abs(NCLIE%), dbOpenSnapshot)
       '
       On Error Resume Next '
       ClienTemp.MoveFirst
       If Err < 1 Then
         Text28 = ClienTemp!Nota_Cli
       End If
     End If
   End If
   Text28.Refresh
   YACAR%(1) = 1: MODIFIC%(1) = 0
   '
End Sub
'
Sub GRANOTAPRO()
   '
   NCLIE% = Val(TRIM$(Text1.TEXT))
   If NCLIE% < 1 Then Exit Sub
   '
   If REGICLI((-1) * NCLIE%) < 1 Then Exit Sub
   If YACAR%(1) < 1 Then Exit Sub
   '
   If NCLIE% > 0 Then
     '
     If MODIFIC%(1) < 1 Then
       Exit Sub
     End If
     '
     If DERACCE%("ANOTAPRO", "Anotaciones de Proveedores") < 2 Then
        Exit Sub
     End If
     '
     If GRABATODO% <> 1 Then
       Command9.SetFocus
       If COMALTER%("Se han Modificado las Notas del Proveedor\'" + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
         YACAR%(1) = 0
         Call CARNOTAPRO
         Exit Sub
       End If
       Command9.SetFocus
     End If
     '
     Call ABREPROVE
     Screen.MousePointer = 11
     TEXTO$ = Text28.TEXT + Chr$(13) + Chr$(10)
     '
     With Proved12
       On Error Resume Next
       .MoveFirst
       .FindFirst "Nume_Cli = " & Abs(NCLIE%)
       If .NoMatch = False Then GoTo 10
       On Error GoTo 0
       Call MENSERR(24, "Imposible Registrar")
       Screen.MousePointer = 1
       Exit Sub
     End With
     '
10   On Error GoTo 0
     With Proved12
       .Edit
       !Nota_Cli = TEXTO$
       .Update
     End With
     Proved12.Close
     MODIFIC%(1) = 0
     '
   End If
   Screen.MousePointer = 1
   '
End Sub
'
Sub BLANFORMU()
   '
   Text2 = ""
   Text3 = ""
   '\\\OT-05-0271-WAR-18/05/05///
   Text4 = ""
   '\\\OT-05-0271-WAR-FIN///
   Text6 = ""
   Text24 = ""
   Text25.TEXT = ""
   Text26.TEXT = ""
   Text27.TEXT = ""
   Text20.TEXT = ""
   Text29.TEXT = ""
   Text30.TEXT = ""
   Text31.TEXT = ""
   Text32.TEXT = ""
   Text19.TEXT = ""
   Text33.TEXT = ""
   Text34.TEXT = ""
   Text36.TEXT = ""
   Text21.TEXT = ""
   '
   For KK% = 0 To 7
     COTEXT(KK%).TEXT = ""
   Next KK%
   '
   Text28.TEXT = ""
   '
   Text23 = ""
   Text35 = ""
   Text36 = ""
   Text37 = ""
   List2.Clear
   List3.Clear
   '
End Sub

'
Private Sub Form_Unload(Cancel As Integer)
   '
   Call GRAGENPRO
   Call GRANOTAPRO
   Call CARDATPROV
   '
   Call CIERRARCH("*.*")
   If AGRECLI% > 0 Then
      Call SORTCLI(1)
      AGRECLI% = 0
   End If
   Call BAJALDISCO
   '
   On Error Resume Next
   Hide
   On Error GoTo 0
   '
End Sub

Private Sub GruProv_Click()
SETPROMA04.Option1(9).Value = True
Call SETPROMA04.Command16_Click

End Sub


Private Sub Help_Click()
'Call Command4_Click
End Sub

Private Sub LiDePo_Click()
Call AMAPRLIS04.Command4_Click
End Sub

Private Sub LidePo1_Click()
Call AMAPRLIS04.Command4_Click
End Sub

Private Sub LiDiMa_Click()
Call AMAPRLIS04.Command2_Click
End Sub

Private Sub LiDiMa1_Click()
Call AMAPRLIS04.Command2_Click
End Sub

Private Sub LiGe_Click()
  Call Command13_Click
End Sub

Private Sub LiInac_Click()
Call AMAPRLIS04.Command5_Click
End Sub

Private Sub List1_Click()
Call Command13_Click
End Sub

Private Sub List3_DblClick()
   '
   VUELBUS% = 1
   '
3  RELIS$ = List3.TEXT
   TIDOCUM$ = TRIM$(Mid$(RELIS$, 11, 16))
   If TIDOCUM$ = "Ajuste de Saldo" Then TIDOCUM$ = "Ajuste Saldo Cliente"
   NUDOCU$ = " " + TRIM$(Mid$(RELIS$, 27, 8)) + " "
   If InStr(UCase$(TIDOCUM$), "FACTURA") > 0 Or InStr(UCase$(TIDOCUM$), "NOTA DE") > 0 Then
     NUDOCU$ = TRIM$(NUDOCU$)
     While Len(NUDOCU$) < 8: NUDOCU$ = "0" + NUDOCU$: Wend
   End If
   RACLI$ = TRIM$(Left$(RASOCLI$(Val(Text1.TEXT)), 8))
   Screen.MousePointer = 11
   For UI& = ULTREG&("PDOCLST") To 1 Step -1
     XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
     If InStr(XX$, TIDOCUM$) > 4 Then
       If InStr(XX$, NUDOCU$) > 4 Then
         If InStr(XX$, RACLI$) > 16 Or VUELBUS% = 2 Then
           DOCUNU& = CVS(Left$(XX$, 4))
           GoTo 5
         End If
       End If
     End If
   Next UI&
   '
   If VUELBUS% < 2 Then
     VUELBUS% = 2
     GoTo 3
   End If
   '
   Screen.MousePointer = 1
   Call COMUNI("Documento no Encontrado")
   Exit Sub
   '
5  For KKI% = 0 To 5
     EDITFORM.Picture1(KKI%).Cls
     EDITFORM.Picture1(KKI%).DrawWidth = 3
     EDITFORM.Picture1(KKI%).DrawStyle = 0
     EDITFORM.Picture1(KKI%).FillStyle = 1
   Next KKI%
   '
   REDOCU$ = REGACT$("PDOCLST")
   URE& = ULTREG&("PDOCSPL")
   LI& = 0: LS& = URE&
6  E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
   L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
   DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
   If DCC& < DOCUNU& Then LI& = L&: GoTo 6
   If DCC& > DOCUNU& Then LS& = L&: GoTo 6
   GoTo 8
   '
7  Screen.MousePointer = 1
   Call MENSERR(24, "Documento no Registrado")
   Exit Sub
   '
8  TPSP$ = ""
   Screen.MousePointer = 11
   While L& > 1
     DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
     If DCC& < DOCUNU& Then GoTo 9
     L& = L& - 1
   Wend
   '
9  CAPAGINAS% = 0
   For KKL& = L& To URE&
     TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
     If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
     If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
       TPSP$ = TPSP$ + Mid$(TBUF$, 5)
     End If
   Next KKL&
   '
95 SPOOLSTRING$ = TPSP$
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
10 PPXX% = InStr(TPSP$, Chr$(255))
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
     If KKI% = PAGINACTIVA% Then EDITFORM.Picture1(KKI%).Visible = True
   Next KKI%
   '
   EDITFORM.Picture1(PAGINACTIVA%).Refresh
   EDITFORM.VScroll1.Min = 0
   EDITFORM.VScroll1.Value = 0
   EDITFORM.VScroll1.Max = 1
   If ALTUPAGINA > EDITFORM.Frame1.Height Then
      EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
   End If
   EDITFORM.Command3.Enabled = False
   EDITFORM.Command1.Enabled = False
   EDITFORM.Show 1
   '
End Sub

Private Sub ListGral_Click()
  FORLIPRO04.Show 1
End Sub

Private Sub LTEXT36_Change()
   DISPONI = Val(LTEXT36.Caption) - Val(LTEXT37.Caption) - Val(LTEXT35.Caption)
   LTEXT23.ForeColor = &H80000008
   If DISPONI < 0 Then LTEXT23.ForeColor = &HFF&
   LTEXT23.Caption = CSTRING$(MKS$(DISPONI), 4, 11, 2, 2)
   LTEXT23.Refresh
End Sub



Private Sub NueCli_Click()
  Option1.Value = True
  Call Command21_Click
End Sub

Private Sub Option1_Click()
   If Option1.Value = True Then
     MODIFIC(0) = 1
     Frame5.Enabled = True
     Frame9.Enabled = True
     Frame10.Enabled = True
     '
     For KKI% = 0 To 2
       COTEXT(KKI%).Enabled = True
       DETEXT(KKI%).Enabled = True
     Next KKI%
     '
     Text19.Enabled = True
     Text25.Enabled = True
     Text26.Enabled = True
     Text27.Enabled = True
     Text29.Enabled = True
     Text30.Enabled = True
     Text31.Enabled = True
     Text32.Enabled = True
     Text33.Enabled = True
     Text34.Enabled = True
     Text36.Enabled = True
     '
     Command21.ToolTipText = "Alta de Nuevo Proveedor"
     '
  End If
End Sub

Private Sub Option2_Click()
   If Option2.Value = True Then
     MODIFIC(0) = 1
     Frame5.Enabled = False
     Frame9.Enabled = False
     Frame10.Enabled = False
     '
     For KKI% = 0 To 2
       COTEXT(KKI%).Enabled = False
       DETEXT(KKI%).Enabled = False
     Next KKI%
     '
     Text19.Enabled = False
     Text25.Enabled = False
     Text26.Enabled = False
     Text27.Enabled = False
     Text29.Enabled = False
     Text30.Enabled = False
     Text31.Enabled = False
     Text32.Enabled = False
     Text33.Enabled = False
     Text34.Enabled = False
     Text36.Enabled = False
     '
     Command21.ToolTipText = "Alta de Nueva Cuenta de Gastos"
     '
   End If
End Sub

Private Sub Print_Click()
   Call Command25_Click
End Sub

Private Sub ProZonVta_Click()
SETPROMA04.Option1(5).Value = True
Call SETPROMA04.Command16_Click

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Tab <> PreviousTab Then
      Call GRAGENPRO
      Call GRANOTAPRO
      Call FRESHALL
      Call CARDATPROV
   End If
End Sub

Private Sub Text1_Change()
   For Index = 0 To 8: YACAR%(Index) = 0: Next Index
   Call BLANFORMU
   If Val(Text1.TEXT) > 0 Then
     If Val(Text1.TEXT) <= 32767 Then
       NCLIE = (-1) * Val(Text1.TEXT)
       If STATCLI(NCLIE) > 0 Then
         'NC% = NCLIE%
         Call CARDATPROV
       End If
     End If
   End If
End Sub
'
Private Sub Text1_DblClick()
   '
   Call Command9_Click
   '
End Sub

Private Sub Text1_GotFocus()
   Call GRAGENPRO
   Call GRANOTAPRO
End Sub

Private Sub Text19_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text2_change()
  '
  MODIFIC%(0) = 1
  '
End Sub

Private Sub Text20_Change()
   MODIFIC%(0) = 1
   If TRIM$(UCase$(Text20.TEXT)) = "SUSPENDIDO" Then
       LTEXT36.Caption = Text20.TEXT
     Else
       LTEXT36.Caption = CSTRING$(MKS$(Val(Text20.TEXT)), 4, 11, 2, 2)
   End If
End Sub

Private Sub Text20_LostFocus()
  If TRIM$(UCase$(Text20.TEXT)) <> "SUSPENDIDO" Then
    Text20.TEXT = CSTRING$(MKS$(Val(Text20.TEXT)), 3, 12, 2, 2)
  End If
End Sub

Private Sub Text21_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text23_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text23_GotFocus()
   If SSTab1.Tab <> 0 Then
      SSTab1.Tab = 0
   End If
End Sub

Private Sub Text24_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text24_GotFocus()
   If SSTab1.Tab <> 0 Then
      SSTab1.Tab = 0
   End If
End Sub

Private Sub Text25_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text26_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text27_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text28_Change()
   MODIFIC(1) = 1
End Sub

Private Sub Text29_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text3_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text3_GotFocus()
   If SSTab1.Tab <> 0 Then
      SSTab1.Tab = 0
   End If
End Sub

Private Sub Text30_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text31_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text32_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text33_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text34_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub LText35_DblClick()
   Screen.MousePointer = 11
   If Val(Text1.TEXT) > 0 Then
     If Val(Text1.TEXT) <= 32767 Then
       NC% = Val(Text1.TEXT)
       If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
         Call COPYSTRU("CHECAR", "LICHECLI")
         Call HEADERS("LICHECLI", "")
         Call HEADERS("LICHECLI", "Cliente: " + RASOCLI$(NC%))
         '
         JJL& = 0
         RFF$ = RECFILT$("CHECAR", 3, MKI$(NC%))
         HOII% = HOY(HOS$)
         For UI& = 1 To Len(RFF$) Step 4
            RECHE& = CVS(Mid$(RFF$, UI&, 4))
            If RECHE& > 0 Then
               If RECHE& <= ULTREG&("CHECAR") Then
                  XX$ = REGLEIDO$("CHECAR", RECHE&)
                  If CVI(Mid$(XX$, 3, 2)) = NC% Then
                     If CVI(Mid$(XX$, 41, 2)) >= HOII% Then
                        JJL& = JJL& + 1
                        Call GRAREG("LICHECLI", XX$, JJL&)
                     End If
                  End If
               End If
            End If
         Next UI&
         Call SETULTREG("LICHECLI", JJL&)
         Call CIERRARCH("LICHECLI")
         Call CIERRARCH("CHECAR")
         '
         Call CONECRUN("*LICHECLI", "Listado de Cheques por Cliente/WAIT")
         '
       End If
     End If
   End If
   '
   Screen.MousePointer = 1
   '
End Sub
Sub CARESCUENTA()
    '
    If YACAR%(2) > 0 Then Exit Sub
    '
50  Screen.MousePointer = 11
    NC% = Val(Text1.TEXT)
    If NC% > 0 Then
        List3.Clear
        List3.Refresh
        List2.Clear
        List2.Refresh
        LTEXT23.Caption = ""
        LTEXT37.Caption = ""
        LTEXT36.Caption = ""
        LTEXT35.Caption = ""
        LTEXT37.Caption = CSTRING$(MKD$(SALDCLI#(NC%)), 4, 11, 2, 2)
        LMCR = LIMICRE(NC%)
        LTEXT36.Caption = CSTRING$(MKS$(LMCR), 4, 11, 2, 2)
        If LMCR < 0 Then LTEXT36.Caption = " Suspendido"
        Refresh
        LTEXT35.Caption = CSTRING$(MKD$(CHEPEACRE#(NC%)), 4, 11, 2, 2)
        DISPONI = Val(LTEXT36.Caption) - Val(LTEXT37.Caption) - Val(LTEXT35.Caption)
        LTEXT23.ForeColor = &H80000008
        If DISPONI < 0 Then LTEXT23.ForeColor = &HFF&
        LTEXT23.Caption = CSTRING$(MKS$(DISPONI), 4, 11, 2, 2)
        Refresh
    End If
    '
100 MAXDIAC& = ULTREG&("COBRAN")
    MAXE& = ULTREG&("FACTUR")
    '
    RFF$ = RECFILT$("FACTUR", 1, MKI$(NC%))
    RFC$ = RECFILT$("COBRAN", 1, MKI$(NC%))
    '
    LINDICE.Clear
    For U& = 1 To Len(RFF$) Step 4
        REG& = CVS(Mid$(RFF$, U&, 4))
        X$ = REGLEIDO$("FACTUR", REG&)
        If CVI(Left$(X$, 2)) <> NC% Then
            X$ = REGLEIDO$("FACTUR", 1)
            Call GRAREG("FACTUR", X$, 1)
            Call CIERRARCH("FACTUR")
            GoTo 100
        End If
        FF% = CVI(Mid$(X$, 7, 2))
        Z$ = CSTRING$(MKI$(FF%), 1, 6, 0, 2) + CSTRING$(MKS$(REG&), 4, 8, 0, 2)
        LINDICE.AddItem Z$
    Next U&
    '
    For U& = 1 To Len(RFC$) Step 4
        REG& = CVS(Mid$(RFC$, U&, 4))
        X$ = REGLEIDO$("COBRAN", REG&)
        If CVI(Left$(X$, 2)) <> NC% Then
            X$ = REGLEIDO$("COBRAN", 1)
            Call GRAREG("COBRAN", X$, 1)
            Call CIERRARCH("COBRAN")
            GoTo 100
        End If
        FF% = CVI(Mid$(X$, 7, 2))
        Z$ = CSTRING$(MKI$(FF%), 1, 6, 0, 2) + CSTRING$(MKS$((-1) * REG&), 4, 8, 0, 2)
        LINDICE.AddItem Z$
    Next U&
    '
    SALDO# = 0: SARRAS# = 0
    For U& = 1 To LINDICE.ListCount
        LINDICE.ListIndex = U& - 1
        REG& = Val(Mid$(LINDICE.TEXT, 7, 8))
        If REG& > 0 Then
            If REG& > MAXE& Then GoTo 990
            X$ = REGLEIDO$("FACTUR", REG&)
            TI% = Asc(Mid$(X$, 13, 1))
            TIPOS$ = "Factura": If TI% = 0 Then TIPOS$ = "Ajuste de Saldo"
                               If TI% = 1 Then TIPOS$ = "Venta Contado"
                               If TI% = 3 Then TIPOS$ = "Nota de Debito"
                               If TI% = 5 Then TIPOS$ = "Descuento Caja"
                               If TI% = 6 Then TIPOS$ = "Nota de Credito"
                               If TI% = 7 Then TIPOS$ = "Ajuste de Saldo"
                               '
            NROL& = CVS(Mid$(X$, 3, 4))
            If NROL& < 1 Then
               If TI% = 0 Or TI% = 7 Then
                  Call REPLA(X$, MKS$(REG&), 3, 4)
                  Call GRAREG("FACTUR", X$, REG&)
               End If
            End If
            '
            SG% = 1: If TI% >= 4 Then If TI% <= 6 Then SG% = -1
            IMPO# = CVD(Mid$(X$, 15, 8))
            IMPO# = IMPO# + CVD(Mid$(X$, 23, 8))
            IMPO# = IMPO# + CVD(Mid$(X$, 31, 8))
            IMPO# = IMPO# + CVD(Mid$(X$, 49, 8))
            IMPO# = IMPO# + CVD(Mid$(X$, 57, 8))
            IMPO# = SG% * IMPO#
            GoTo 890
            '
         ElseIf REG& < 0 Then
            REG1& = Abs(REG&)
            If REG1& > MAXDIAC& Then GoTo 990
            X$ = REGLEIDO$("COBRAN", REG1&)
            TI% = 4
            TIPOS$ = "Recibo"
            SG% = (-1)
            IMPO# = (-1) * CVD(Mid$(X$, 15, 8))
            GoTo 890
        End If
        GoTo 990
        '
890     SALDO# = SALDO# + (Int(100 * IMPO# + 0.5)) / 100
        FEX = CVI(Mid$(X$, 7, 2))
        NROL& = CVS(Mid$(X$, 3, 4))
        Z$ = Space$(80)
        '
        Call REPLA(Z$, FECHATEX$(FEX), 1, 8)
        Call REPLA(Z$, TIPOS$, 11, 16)
        Call REPLA(Z$, CSTRING$(MKS$(NROL&), 3, 8, 0, 2), 27, 8)
        IMPTEX$ = CSTRING$(MKD$(Abs(IMPO#)), 3, 13, 2, 2)
        SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 14, 2, 2)
        If IMPO# > 0 Then
            Call REPLA(Z$, IMPTEX$, 35, 13)
          ElseIf IMPO# < 0 Then
            Call REPLA(Z$, IMPTEX$, 50, 13)
        End If
        Call REPLA(Z$, SALTEX$, 63, 14)
        List3.AddItem Z$
        '
990 Next U&
    '
991 SARRAS# = SALDO#
    List3.ListIndex = List3.ListCount - 1
    List3.Refresh
    '
    MAXDEU& = ULTREG&("SVD202")
    RFD$ = RECFILT$("SVD202", 1, MKI$(NC%))
    '
    LINDICE.Clear
    For U& = 1 To Len(RFD$) Step 4
        REG& = CVS(Mid$(RFD$, U&, 4))
        X$ = REGLEIDO$("SVD202", REG&)
        NCII% = CVI(Left$(X$, 2))
        If CVI(Left$(X$, 2)) <> NC% Then
            X$ = REGLEIDO$("SVD202", 1)
            Call GRAREG("SVD202", X$, 1)
            Call CIERRARCH("SVD202")
            GoTo 100
        End If
        IMPO# = CVD(Mid$(X$, 21, 8))
        SALDO# = CVD(Mid$(X$, 11, 8))
        If Abs(SALDO#) < 0.005 Then
          If Abs(SALDO#) > 0.0005 Then
            SALDO# = 0
            Call REPLA(X$, MKD$(0), 11, 8)
            Call GRAREG("SVD202", X$, REG&)
          End If
        End If
        '
        If Abs(SALDO#) >= 0.005 Then
          If Sgn(IMPO#) <> Sgn(SALDO#) Or Abs(SALDO#) > Abs(IMPO#) + 0.005 Then
            ABUS$ = Left$(X$, 8)
            For UU& = 1 To Len(RFD$) Step 4
                REG& = CVS(Mid$(RFD$, UU&, 4))
                X$ = REGLEIDO$("SVD202", REG&)
                If Left$(X$, 8) = ABUS$ Then
                    Call GRAREG("SVD202", String$(32, 0), REG&)
                End If
            Next UU&
            GoTo 100
          End If
        End If
        '
        If Abs(SALDO#) > 0 Then
            FF% = CVI(Mid$(X$, 9, 2))
            Z$ = CSTRING$(MKI$(FF%), 1, 6, 0, 2) + CSTRING$(MKS$(REG&), 4, 8, 0, 2)
            LINDICE.AddItem Z$
        End If
    Next U&
    '
    SUMADEU# = 0: SUMAFAC# = 0
    For U& = 1 To LINDICE.ListCount
        LINDICE.ListIndex = U& - 1
        REG& = Val(Mid$(LINDICE.TEXT, 7, 8))
        If REG& > 0 Then
            If REG& <= MAXDEU& Then
                X$ = REGLEIDO$("SVD202", REG&)
                TI% = Asc(Mid$(X$, 3, 1))
                TIPOS$ = "Factura": If TI% = 0 Then TIPOS$ = "Ajuste de Saldo"
                                   If TI% = 1 Then TIPOS$ = "Venta Contado"
                                   If TI% = 3 Then TIPOS$ = "Nota de Debito"
                                   If TI% = 4 Then TIPOS$ = "Anticipo"
                                   If TI% = 5 Then TIPOS$ = "Descuento Caja"
                                   If TI% = 6 Then TIPOS$ = "Nota de Credito"
                                   If TI% = 7 Then TIPOS$ = "Ajuste de Saldo"
                                   '
                SG% = 1: If TI% >= 4 Then If TI% <= 6 Then SG% = -1
                IMPO# = CVD(Mid$(X$, 21, 8))
                If TI% <> 0 Then IMPO# = Abs(IMPO#)
                IMPO# = SG% * IMPO#
                SUMAFAC# = SUMAFAC# + IMPO#
                '
                SALDO# = CVD(Mid$(X$, 11, 8))
                If TI% <> 0 Then SALDO# = Abs(SALDO#)
                SALDO# = SG% * SALDO#
                SUMADEU# = SUMADEU# + SALDO#
                '
                NROL& = CVS(Mid$(X$, 5, 4))
                FEMI = CVI(Mid$(X$, 19, 2))
                FEVE = CVI(Mid$(X$, 9, 2))
                '
1234            FEVEX$ = FECHATEX$(FEVE)
                If Val(Mid$(FEVEX$, 4, 2)) = 2 Then
                  If Val(Mid$(FEVEX$, 1, 2)) > 28 Then
                    FEVE = DIPOST(FEVE)
                    GoTo 1234
                  End If
                End If
                If CVI(Mid$(X$, 9, 2)) <> FEVE Then
                  FEVEI% = FEVE
                  Call REPLA(X$, MKI$(FEVEI%), 9, 2)
                  Call GRAREG("SVD202", X$, REG&)
                End If
                '
                IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 14, 2, 2)
                SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 14, 2, 2)
                '
If Abs(Abs(IMPO#) - Abs(SALDO#)) < 0.005 Then ' NO ESTA CANCELADA
  MONEMX% = Asc(Mid$(X$, 31, 1))
  If MONEMX% > 2 Then
    MONEMX% = 0
    Call REPLA(X$, Chr$(0), 31, 1)
    Call GRAREG("SVD202", X$, REG&)
  End If
  If MONEMX% = 0 Then
    RFG$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
    For KUL& = Len(RFG$) - 3 To 1 Step -4
      RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
      XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
      If CVS(Mid$(XXC$, 39, 4)) = NROL& Then
        If Asc(Mid$(XXC$, 37, 1)) = TI% Then
          MONEMX% = Asc(Mid$(X$, 31, 1))
          GoTo 15
        End If
      End If
    Next KUL&
15 If MONEMX% <> 2 Then MONEMX% = 1
    Call REPLA(X$, Chr$(MONEMX%), 31, 1)
    Call GRAREG("SVD202", X$, REG&)
  End If
  If MONEMX% = 2 Then
    RFG$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
    For KUL& = Len(RFG$) - 3 To 1 Step -4
      RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
      XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
      If CVS(Mid$(XXC$, 39, 4)) = NROL& Then
        If Asc(Mid$(XXC$, 37, 1)) = TI% Then
          SALTEX$ = "u$s" + FORMATNUM$(CVD(Mid$(XXC$, 97, 8)) - CVD(Mid$(XXC$, 105, 8)), "F11.2")
          GoTo 20
        End If
      End If
    Next KUL&
  End If
End If
'
20              If TRIM$(SALTEX$) <> "" Then
                  Z$ = Space$(80)
                  Call REPLA(Z$, TIPOS$, 1, 16)
                  Call REPLA(Z$, CSTRING$(MKS$(NROL&), 3, 8, 0, 2), 17, 8)
                  Call REPLA(Z$, FECHATEX$(FEMI), 27, 8)
                  Call REPLA(Z$, IMPTEX$, 37, 14)
                  Call REPLA(Z$, FECHATEX$(FEVE), 53, 8)
                  Call REPLA(Z$, SALTEX$, 63, 14)
                  '
                  List2.AddItem Z$
                End If
                '
            End If
        End If
        '
    Next U&
    '
    List2.ListIndex = List2.ListCount - 1
    List2.Refresh
    Label28.Caption = CSTRING$(MKD$(SUMADEU#), 4, 14, 2, 2)
    Label28.Refresh
    Label25.Caption = CSTRING$(MKD$(SUMADEU#), 4, 14, 2, 2)
    Label25.Refresh
    '
    If Abs(SARRAS# - SUMADEU#) >= 0.005 Then
       ATENTI% = 1
       TTXX$ = "Se ha Detectado una Diferencia\en la Composición de la Deuda\de la Cuenta"
       If Visible = True Then
           TTXX$ = TTXX$ + " Elegida."
         Else
           TTXX$ = TTXX$ + ":\   \'" + TRIM$(RASOCLI$(NC%)) + "'"
       End If
       TTXX$ = TTXX$ + "\   \Puede Intentarse la Reconstucción\Mediante un Proceso de Re-Aplicación.\  \Este Proceso Reconstruye\la Composición de la Deuda de la Cuenta Elegida\en Base a la Aplicación de Créditos y Cobranzas\Registrada en el Sistema.\  \Desea Cancelar o Continuar el Proceso ?\\Cancelar\Continuar"
       '
       If COMALTER%(TTXX$) = 2 Then
           Call REAPLICRE(NC%)
           GoTo 50
         Else
           ATENTI% = 0
           Call CIERRARCH("*.*")
           Call FINAL("")
       End If
    End If
    '
    If Abs(SARRAS# - SALDCLI#(NC%)) >= 0.005 Then
        RGXC& = REGICLI&(NC%)
        XC$ = REGLEIDO$("DEUDOR2", RGXC&)
        Call REPLA(XC$, MKD$(SARRAS#), 109, 8)
        Call GRAREG("DEUDOR2", XC$, RGXC&)
        LTEXT37.Caption = CSTRING$(MKD$(SARRAS#), 4, 11, 2, 2)
        Refresh
    End If
    YACAR%(2) = 1
    Screen.MousePointer = 1
    '
End Sub
'
Sub REAPLICRE(NC%)
    '
    ATENTI% = 1
    Dim DOCU$(8192), IMPO#(8192), REGDO&(8192), FFDOC%(8192), IORI#(8192)
    '
100 Screen.MousePointer = 11
    '
    SUMARRAS# = 0
    If List3.ListCount > 0 Then
      SUMARRAS# = Val(Mid$(List3.List(List3.ListCount - 1), 63, 14))
    End If
    SUMADEUP# = Val(Label28.Caption)
    CADOCPEN% = List2.ListCount
    '
    MAXDIAC& = ULTREG&("COBRAN")
    MAXE& = ULTREG&("FACTUR")
    '
    RFF$ = RECFILT$("FACTUR", 1, MKI$(NC%))
    RFC$ = RECFILT$("COBRAN", 1, MKI$(NC%))
    '
    NUDOCU% = 0
    FIN& = Len(RFF$)
    For U& = 1 To Len(RFF$) Step 4
        Call TRACE(20, U&, FIN&)
        REG& = CVS(Mid$(RFF$, U&, 4))
        X$ = REGLEIDO$("FACTUR", REG&)
        If CVI(Left$(X$, 2)) <> NC% Then
            X$ = REGLEIDO$("FACTUR", 1)
            Call GRAREG("FACTUR", X$, 1)
            Call CIERRARCH("FACTUR")
            GoTo 100
        End If
110     NUDO& = CVS(Mid$(X$, 3, 4))
        TI% = Asc(Mid$(X$, 13, 1))
        If NUDO& < 1 Then
           If TI% = 0 Or TI% = 7 Then
              Call REPLA(X$, MKS$(REG&), 3, 4)
              Call GRAREG("FACTUR", X$, REG&)
              GoTo 110
           End If
        End If
        VA% = Asc(Mid$(X$, 14, 1))
        SG% = 1: If TI% >= 4 Then If TI% <= 6 Then SG% = -1
        NUDOCU% = NUDOCU% + 1
        If NUDOCU% > 8192 Then
           Call MENSERR(24, "Imposible Continuar Proceso.\Demasiados Movimientos en la Cuenta Elegida.\Consulte a su Soporte de Sistemas FLEXOFT.")
           Call CIERRARCH("*.*")
           Call FINAL("")
        End If
115     IMPO#(NUDOCU%) = CVD(Mid$(X$, 15, 8))
        IMPO#(NUDOCU%) = IMPO#(NUDOCU%) + CVD(Mid$(X$, 23, 8))
        IMPO#(NUDOCU%) = IMPO#(NUDOCU%) + CVD(Mid$(X$, 31, 8))
        IMPO#(NUDOCU%) = IMPO#(NUDOCU%) + CVD(Mid$(X$, 49, 8))
        IMPO#(NUDOCU%) = IMPO#(NUDOCU%) + CVD(Mid$(X$, 57, 8))
        IMPO#(NUDOCU%) = SG% * (Int(100 * IMPO#(NUDOCU%) + 0.5)) / 100
        NUSUCUVE% = CVI(Mid$(X$, 47, 2))
        DOCU$(NUDOCU%) = Left$(FORMDOC$(NUDO&, TI%, VA%, NUSUCUVE%), 2) + "-" + TRIM$(Str$(NUDO&))
        REGDO&(NUDOCU%) = REG&
        FFDOC%(NUDOCU%) = CVI(Mid$(X$, 7, 2))
        '
        If TI% > 0 Then
          If TI% < 7 Then
            If CVD(Mid$(X$, 15, 8)) < (-0.005) Then
              Call REPLA(X$, MKD$(Abs(CVD(Mid$(X$, 15, 8)))), 15, 8)
              Call REPLA(X$, MKD$(Abs(CVD(Mid$(X$, 23, 8)))), 23, 8)
              Call REPLA(X$, MKD$(Abs(CVD(Mid$(X$, 31, 8)))), 31, 8)
              Call REPLA(X$, MKD$(Abs(CVD(Mid$(X$, 49, 8)))), 49, 8)
              Call REPLA(X$, MKD$(Abs(CVD(Mid$(X$, 57, 8)))), 57, 8)
              Call GRAREG("FACTUR", X$, REG&)
              GoTo 115
            End If
          End If
        End If
    Next U&
    '
    FIN& = Len(RFC$)
    For U& = 1 To Len(RFC$) Step 4
        Call TRACE(20, U&, FIN&)
        REG& = CVS(Mid$(RFC$, U&, 4))
        X$ = REGLEIDO$("COBRAN", REG&)
        If CVI(Left$(X$, 2)) <> NC% Then
            X$ = REGLEIDO$("COBRAN", 1)
            Call GRAREG("COBRAN", X$, 1)
            Call CIERRARCH("COBRAN")
            GoTo 100
        End If
        NUDO& = CVS(Mid$(X$, 3, 4))
        TI% = 4
        VA% = 1
        SG% = (-1)
        NUDOCU% = NUDOCU% + 1
        If NUDOCU% > 8192 Then
           Call MENSERR(24, "Imposible Continuar Proceso.\Demasiados Movimientos en la Cuenta Elegida.\Consulte a su Soporte de Sistemas FLEXOFT.")
           Call CIERRARCH("*.*")
           Call FINAL("")
        End If
        IMPO#(NUDOCU%) = CVD(Mid$(X$, 15, 8))
        IMPO#(NUDOCU%) = SG% * (Int(100 * IMPO#(NUDOCU%) + 0.5)) / 100
        DOCU$(NUDOCU%) = "RB-" + TRIM$(Str$(NUDO&))
        REGDO&(NUDOCU%) = (-1) * REG&
        FFDOC%(NUDOCU%) = CVI(Mid$(X$, 7, 2))
    Next U&
    '
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
       U& = KKI%
       Call TRACE(20, U&, FIN&)
       For KKJ% = KKI% + 1 To NUDOCU%
         If FFDOC%(KKI%) > FFDOC%(KKJ%) Then
           DOCX$ = DOCU$(KKI%)
           FFX% = FFDOC%(KKI%)
           IMX# = IMPO#(KKI%)
           RGX& = REGDO&(KKI%)
           DOCU$(KKI%) = DOCU$(KKJ%)
           FFDOC%(KKI%) = FFDOC%(KKJ%)
           IMPO#(KKI%) = IMPO#(KKJ%)
           REGDO&(KKI%) = REGDO&(KKJ%)
           DOCU$(KKJ%) = DOCX$
           FFDOC%(KKJ%) = FFX%
           IMPO#(KKJ%) = IMX#
           REGDO&(KKJ%) = RGX&
         End If
       Next KKJ%
    Next KKI%
    '
    SUMAPEN# = 0
    For KKI% = 1 To NUDOCU%
       SUMAPEN# = SUMAPEN# + IMPO#(KKI%)
       If Abs(SUMAPEN#) < 0.005 Then
          For KKJ% = 1 To KKI%
             IMPO#(KKJ%) = 0
          Next KKJ%
       End If
    Next KKI%
    '
    If NUDOCU% >= 1 Then
       If Left$(DOCU$(1), 2) = "AS" Then
          FIN& = ULTREG&("SALINID")
          For U& = 1 To FIN&
             X$ = REGLEIDO$("SALINID", U&)
             If CVI(Left$(X$, 2)) = NC% Then
               TI% = Asc(Mid$(X$, 3, 1))
               VA% = Asc(Mid$(X$, 4, 1))
               NUSUCUVE% = CVI(Mid$(X$, 29, 2))
               SG% = 1: If TI% >= 4 Then If TI% <= 6 Then SG% = -1
               IPEN# = SG% * CVD(Mid$(X$, 11, 8))
               If Sgn(IPEN#) = Sgn(IMPO#(1)) Then
                 If Abs(IPEN#) <= Abs(IMPO#(1)) Then
                   IMPO#(1) = IMPO#(1) - IPEN#
                   NUDO& = CVS(Mid$(X$, 5, 4))
                   NUDOCU% = NUDOCU% + 1
                   If NUDOCU% > 8192 Then
                      Call MENSERR(24, "Imposible Continuar Proceso.\Demasiados Movimientos en la Cuenta Elegida.\Consulte a su Soporte de Sistemas FLEXOFT.")
                      Call CIERRARCH("*.*")
                      Call FINAL("")
                   End If
                   IMPO#(NUDOCU%) = IPEN#
                   DOCU$(NUDOCU%) = Left$(FORMDOC$(NUDO&, TI%, VA%, NUSUCUVE%), 2) + "-" + TRIM$(Str$(NUDO&))
                   REGDO&(NUDOCU%) = 900000 + U&
                   FFDOC%(NUDOCU%) = CVI(Mid$(X$, 19, 2))
                 End If
               End If
             End If
          Next U&
       End If
    End If
    '
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
       U& = KKI%
       Call TRACE(20, U&, FIN&)
       For KKJ% = KKI% + 1 To NUDOCU%
         If FFDOC%(KKI%) > FFDOC%(KKJ%) Then
           DOCX$ = DOCU$(KKI%)
           FFX% = FFDOC%(KKI%)
           IMX# = IMPO#(KKI%)
           RGX& = REGDO&(KKI%)
           DOCU$(KKI%) = DOCU$(KKJ%)
           FFDOC%(KKI%) = FFDOC%(KKJ%)
           IMPO#(KKI%) = IMPO#(KKJ%)
           REGDO&(KKI%) = REGDO&(KKJ%)
           DOCU$(KKJ%) = DOCX$
           FFDOC%(KKJ%) = FFX%
           IMPO#(KKJ%) = IMX#
           REGDO&(KKJ%) = RGX&
         End If
       Next KKJ%
    Next KKI%
    '
' HASTA AQUI O.K
    KKI% = 0: SUMAPEX# = 0
    For KKJ% = 1 To NUDOCU%
       If Abs(IMPO#(KKJ%)) >= 0.005 Then
          KKI% = KKI% + 1
          DOCU$(KKI%) = DOCU$(KKJ%)
          FFDOC%(KKI%) = FFDOC%(KKJ%)
          IMPO#(KKI%) = IMPO#(KKJ%)
          SUMAPEX# = SUMAPEX# + IMPO#(KKI%)
          REGDO&(KKI%) = REGDO&(KKJ%)
       End If
    Next KKJ%
    NUDOCU% = KKI%
    '
    FIN& = ULTREG&("APLICOB")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("APLICOB", U&)
       If CVI(Mid$(X$, 3, 2)) = NC% Then
          RECI$ = Mid$(X$, 5, 18)
          RECI$ = Left$(RECI$, 2) + "-" + TRIM$(Val(Mid$(RECI$, 11)))
          FACU$ = Mid$(X$, 55, 18)
          PINI% = 11: If Mid$(FACU$, 17, 1) = "-" Then PINI% = 9
          FACU$ = Left$(FACU$, 2) + "-" + TRIM$(Val(Mid$(FACU$, PINI%, 8)))
          '
210       IAPLI# = CVD(Mid$(X$, 23, 8))
          IRECIMA# = 0: IFACMAX# = 0
          For KKI% = 1 To NUDOCU%
             If DOCU$(KKI%) = RECI$ Or (RECI$ = "AS-0" And Left$(DOCU$(KKI%), 2) = "AS") Then
               If IMPO#(KKI%) < 0 Then
                 IRECIMA# = IRECIMA# - IMPO#(KKI%)
               End If
             End If
             If DOCU$(KKI%) = FACU$ Then
               If IMPO#(KKI%) > 0 Then
                 IFACMAX# = IFACMAX# + IMPO#(KKI%)
               End If
             End If
          Next KKI%
          '
          If IAPLI# > IRECIMA# Then IAPLI# = IRECIMA#
          If IAPLI# > IFACMAX# Then IAPLI# = IFACMAX#
          '
          IRECIMA# = IAPLI#
          IFACMAX# = IAPLI#
          '
          If IAPLI# > 0 Then
             For KKI% = 1 To NUDOCU%
                If DOCU$(KKI%) = RECI$ Or (RECI$ = "AS-0" And Left$(DOCU$(KKI%), 2) = "AS") Then
                   If IMPO#(KKI%) < 0 Then
                      IAP1# = IRECIMA#
                      If IAP1# > Abs(IMPO#(KKI%)) Then IAP1# = Abs(IMPO#(KKI%))
                      IMPO#(KKI%) = IMPO#(KKI%) + IAP1#
                      IRECIMA# = IRECIMA# - IAP1#
                   End If
                End If
                If DOCU$(KKI%) = FACU$ Then
                   If IMPO#(KKI%) > 0 Then
                      IAP2# = IFACMAX#
                      If IAP2# > Abs(IMPO#(KKI%)) Then IAP2# = Abs(IMPO#(KKI%))
                      IMPO#(KKI%) = IMPO#(KKI%) - IAP2#
                      IFACMAX# = IFACMAX# - IAP2#
                   End If
                End If
             Next KKI%
          End If
          '
       End If
    Next U&
    '
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
       U& = KKI%
       Call TRACE(20, U&, FIN&)
       For KKJ% = KKI% + 1 To NUDOCU%
         If FFDOC%(KKI%) > FFDOC%(KKJ%) Or (FFDOC%(KKI%) = FFDOC%(KKJ%) And Sgn(REGDO&(KKI%)) = Sgn(REGDO&(KKJ%)) > 0 And Abs(REGDO&(KKI%)) > Abs(REGDO&(KKJ%))) Then
           DOCX$ = DOCU$(KKI%)
           FFX% = FFDOC%(KKI%)
           IMX# = IMPO#(KKI%)
           RGX& = REGDO&(KKI%)
           DOCU$(KKI%) = DOCU$(KKJ%)
           FFDOC%(KKI%) = FFDOC%(KKJ%)
           IMPO#(KKI%) = IMPO#(KKJ%)
           REGDO&(KKI%) = REGDO&(KKJ%)
           DOCU$(KKJ%) = DOCX$
           FFDOC%(KKJ%) = FFX%
           IMPO#(KKJ%) = IMX#
           REGDO&(KKJ%) = RGX&
         End If
       Next KKJ%
    Next KKI%
    '
    SUMAPEN# = 0
    For KKI% = 1 To NUDOCU%
       SUMAPEN# = SUMAPEN# + IMPO#(KKI%)
       If Abs(SUMAPEN#) < 0.005 Then
          For KKJ% = 1 To KKI%
             IMPO#(KKJ%) = 0
          Next KKJ%
       End If
    Next KKI%
    '
    KKI% = 0
    For KKJ% = 1 To NUDOCU%
       If Abs(IMPO#(KKJ%)) >= 0.005 Then
          KKI% = KKI% + 1
          DOCU$(KKI%) = DOCU$(KKJ%)
          FFDOC%(KKI%) = FFDOC%(KKJ%)
          IMPO#(KKI%) = IMPO#(KKJ%)
          REGDO&(KKI%) = REGDO&(KKJ%)
       End If
    Next KKJ%
    NUDOCU% = KKI%
    '
    For KKI% = 1 To NUDOCU%
       For KKJ% = KKI% + 1 To NUDOCU%
          If Abs(IMPO#(KKI%) + IMPO#(KKJ%)) < 0.005 Then
             IMPO#(KKI%) = 0
             IMPO#(KKJ%) = 0
          End If
       Next KKJ%
    Next KKI%
    '
    KKI% = 0: SUMAPEX# = 0
    For KKJ% = 1 To NUDOCU%
       If Abs(IMPO#(KKJ%)) >= 0.005 Then
          KKI% = KKI% + 1
          DOCU$(KKI%) = DOCU$(KKJ%)
          FFDOC%(KKI%) = FFDOC%(KKJ%)
          IMPO#(KKI%) = IMPO#(KKJ%)
          SUMAPEX# = SUMAPEX# + IMPO#(KKJ%)
          REGDO&(KKI%) = REGDO&(KKJ%)
       End If
    Next KKJ%
    NUDOCU% = KKI%
    '
    Screen.MousePointer = 1
    If COMALTER%("Se ha Completado la Reconstrucción\de la Composición de la Deuda de\   \'" + TRIM$(RASOCLI$(NC%)) + "'\   \Desea Grabar la Nueva Composición de la Deuda ?\\Cancelar\Grabar") <> 2 Then
       GoTo 999
    End If
    '
    ' para no enquilombar las cuentas que mas o menos estan
    If Abs(SUMARRAS# - SUMADEUP#) < 0.005 Then
      If NUDOCU% >= CADOCPEN% Then
        GoTo 999
      End If
    End If
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("SVD202")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("SVD202", U&)
       If CVI(Left$(X$, 2)) = NC% Then
          Call REPLA(X$, String$(30, 0), 3, 30)
          Call GRAREG("SVD202", X$, U&)
       End If
    Next U&
    '
    LU$ = LUDAT$
    ARCHIDEU$ = "SVD202"
    RDX% = 1
    FIN& = NUDOCU%
    YABLO% = 0
    For KKI% = 1 To NUDOCU%
       U& = KKI%
       Call TRACE(20, U&, FIN&)
       If Abs(IMPO#(KKI%)) >= 0.005 Then
          IORIG# = IMPO#(KKI%)
          If REGDO&(KKI%) > 0 Then
             If REGDO&(KKI%) > 900000 Then
                X$ = REGLEIDO$("SALINID", REGDO&(KKI%) - 900000)
                NRO = CVS(Mid$(X$, 5, 4))
                TIPO% = Asc(Mid$(X$, 3, 1))
                VA% = Asc(Mid$(X$, 4, 1))
                FFAC% = CVI(Mid$(X$, 9, 2))
                COPAGI% = 0
                GoTo 290
             End If
             X$ = REGLEIDO$("FACTUR", REGDO&(KKI%))
             NRO = CVS(Mid$(X$, 3, 4))
             TIPO% = Asc(Mid$(X$, 13, 1))
             VA% = Asc(Mid$(X$, 14, 1))
             FFAC% = CVI(Mid$(X$, 7, 2))
             COPAGI% = CPAGCLI%(NC%)
             If TIPO% < 1 Or TIPO% > 3 Then COPAGI% = 0
             SG% = 1: If TIPO% >= 4 Then If TIPO% <= 6 Then SG% = -1
             IORIG# = CVD(Mid$(X$, 15, 8))
             IORIG# = IORIG# + CVD(Mid$(X$, 23, 8))
             IORIG# = IORIG# + CVD(Mid$(X$, 31, 8))
             IORIG# = IORIG# + CVD(Mid$(X$, 49, 8))
             IORIG# = IORIG# + CVD(Mid$(X$, 57, 8))
             IORIG# = SG% * IORIG#
           Else
             X$ = REGLEIDO$("COBRAN", (-1) * REGDO&(KKI%))
             NRO = CVS(Mid$(X$, 3, 4))
             TIPO% = 4
             VA% = 0
             FFAC% = CVI(Mid$(X$, 7, 2))
             COPAGI% = 0
          End If
290       TOTFAC# = IMPO#(KKI%)
          If YABLO% = 0 Then
             YABLO% = 1
             Call BLOQARCH(ARCHIDEU$)
          End If
          ARDEX$ = ARCHIDEU$ + "/" + TRIM$(CSTRING$(MKD$(IORIG#), 4, 12, 2, 2))
          Call GRADEU(LU$, ARDEX$, RDX%, TIPO%, VA%, NRO, NC%, FFAC%, COPAGI%, TOTFAC#)
       End If
    Next KKI%
    '
999 Call CIERRARCH(ARCHIDEU$)
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Screen.MousePointer = 1
    ATENTI% = 0
End Sub

Sub GENECOPROV()
   If UPDATPRO% > 0 Then
      Screen.MousePointer = 11
      Call ABREPROVE
      Set CLIENTECO = Badaprov.OpenRecordset("SELECT Codi_Cli,Nume_Cli,Raso_Cli FROM Proved12 WHERE Stat_Cli > 0 ORDER BY Raso_Cli ASC;", dbOpenDynaset, dbReadOnly)
      On Error Resume Next
      CLIENTECO.MoveLast
      If Err Then Exit Sub
      '
      CAREC% = CLIENTECO.RecordCount
      If CAREC% < 1 Then Exit Sub
      '
      FIN& = CAREC%
      REBLA$ = REGBLAN$("PROVED1")
      REBLB$ = REGBLAN$("INDIPROV")
      JJ& = 0: J1& = 0
      With CLIENTECO
        .MoveFirst
        '
10      Call TRACE(20, JJ&, FIN&)
        EL1$ = REBLA$
        Call REPLA(EL1$, MKI$(!Nume_Cli), 1, 2)
        Call REPLA(EL1$, !Raso_Cli, 3, 30)
        JJ& = JJ& + 1
        Call GRAREG("PROVED1", EL1$, JJ&)
        '
        EL2$ = REBLB$
        Call REPLA(EL2$, !Codi_Cli, 1, 12)
        Call REPLA(EL2$, !Raso_Cli, 13, 50)
        Call REPLA(EL2$, MKI$(!Nume_Cli), 63, 2)
        Call GRAREG("INDIPROV", EL2$, JJ&)
        '
        If TICUEPRO((-1) * !Nume_Cli) = 0 Then
          J1& = J1& + 1
          Call GRAREG("ACREDOR", EL1$, J1&)
          Call GRAREG("INDIACRE", EL2$, J1&)
       End If
        '
        If JJ& < CAREC% Then
          CLIENTECO.MoveNext
          GoTo 10
        End If
      End With
      '
      CLIENTECO.Close
      Call SETULTREG("PROVED1", JJ&)
      Call CIERRARCH("PROVED1")
      Call SETULTREG("ACREDOR", J1&)
      Call CIERRARCH("ACREDOR")
      Call FILESORT("INDIPROV", "", 1, 1, "ASC")
      Call FILESORT("INDIACRE", "", 1, 1, "ASC")
      Call FRESHECHO("PROVED1")
      Call FRESHECHO("INDIPROV")
      Call FRESHECHO("INDIACRE")
      Call BAJALDISCO
      UPDATCLI% = 0
      Screen.MousePointer = 1
   End If
End Sub
Sub GENPROVANU()
   '
   Call ABREPROVE
   SQLX$ = "SELECT * FROM PROVED12 "
   SQLX$ = SQLX$ + " WHERE Stat_Cli = -1 "
   Set APRO1TMP = Badaprov.OpenRecordset(SQLX$, 4)
   With APRO1TMP
     On Error Resume Next
     .MoveFirst
     If Err Then
         On Error GoTo 0
         GoTo 99
     End If
     On Error GoTo 0
     REBLB$ = REGBLAN("INDIPROA")
     JJ& = 0
     Do While .EOF = False
       EL2$ = REBLB$
       Call REPLA(EL2$, !Codi_Cli, 1, 12)
       Call REPLA(EL2$, !Raso_Cli, 13, 50)
       Call REPLA(EL2$, MKI$(!Nume_Cli), 63, 2)
       JJ& = JJ& + 1
       Call GRAREG("INDIPROA", EL2$, JJ&)
       .MoveNext
      Loop
      Call SETULTREG("INDIPROA", JJ&)
      Call FILESORT("INDIPROA", "", 1, 1, "ASC")
      Call FRESHECHO("INDIPROA")
   '
99 End With
   '
End Sub

Private Sub Text36_Change()
  '
  MODIFIC%(0) = 1
  '
End Sub
'
Private Sub Text6_Change()
   If NUMEPRO%(Text6) > 0 Then
     Text1 = TRIM$(Str$(NUMEPRO%(Text6)))
   End If
End Sub

Private Sub Text6_DblClick()
   Call Command10_Click
End Sub

Sub GENDAREGAPRO()
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("PROVED1")
    REBLA$ = REGBLAN$("DARGAPRO")
    Call REPLA(REBLA$, USERNAME$, 36, 20)
    Call REPLA(REBLA$, MKI$(HOY(HOS$)), 56, 2)
    Call REPLA(REBLA$, MKI$(HORANUM(Time$)), 58, 2)
    '
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("PROVED1", U&)
      NCLIX% = CVI(Left$(XX$, 2))
      If NCLIX% > 0 Then
        RECORD$ = REBLA$
        Call REPLA(RECORD$, MKI$(NCLIX%), 1, 2)
        Call REPLA(RECORD$, Chr$(2), 3, 1)
        Call FILTERPUTRECORD("DARGAPRO", 1, MKI$(NCLIX%), RECORD$)
      End If
    Next U&
    Call CIERRARCH("DARGAPRO")
    Screen.MousePointer = 1
    '
End Sub

