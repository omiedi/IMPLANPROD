VERSION 5.00
Begin VB.Form FORFACLI 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturación en Cuenta Corriente"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11910
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3165
      Left            =   0
      TabIndex        =   30
      Top             =   5145
      Width           =   10800
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Entrega"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   1530
         Left            =   210
         TabIndex        =   94
         Top             =   1575
         Width           =   5330
         Begin VB.TextBox LENTREGA 
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
            Index           =   3
            Left            =   1575
            MaxLength       =   32
            TabIndex        =   98
            Top             =   1800
            Visible         =   0   'False
            Width           =   3750
         End
         Begin VB.TextBox LENTREGA 
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
            Index           =   0
            Left            =   1155
            MaxLength       =   32
            TabIndex        =   97
            Top             =   345
            Width           =   3990
         End
         Begin VB.TextBox LENTREGA 
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
            Index           =   1
            Left            =   1155
            MaxLength       =   32
            TabIndex        =   96
            Top             =   715
            Width           =   3990
         End
         Begin VB.TextBox LENTREGA 
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
            Index           =   2
            Left            =   1155
            MaxLength       =   32
            TabIndex        =   95
            Top             =   1080
            Width           =   3990
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Dirección:"
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
            Index           =   12
            Left            =   -840
            TabIndex        =   100
            Top             =   435
            Width           =   1905
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Transporte:"
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
            Index           =   11
            Left            =   -840
            TabIndex        =   99
            Top             =   1155
            Width           =   1905
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Resumen Final"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   3120
         Left            =   5670
         TabIndex        =   50
         Top             =   0
         Width           =   5055
         Begin VB.TextBox Text16 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1785
            TabIndex        =   116
            TabStop         =   0   'False
            Top             =   800
            Width           =   1065
         End
         Begin VB.TextBox Total 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   7
            Left            =   3255
            TabIndex        =   115
            TabStop         =   0   'False
            Text            =   " "
            Top             =   800
            Width           =   1650
         End
         Begin VB.TextBox Text15 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3255
            TabIndex        =   110
            Top             =   2405
            Width           =   1650
         End
         Begin VB.TextBox TICAMBIS 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   525
            MaxLength       =   14
            TabIndex        =   105
            Top             =   2775
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.TextBox Total 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   6
            Left            =   3255
            TabIndex        =   80
            TabStop         =   0   'False
            Text            =   " "
            Top             =   2775
            Width           =   1650
         End
         Begin VB.TextBox Total 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   3
            Left            =   3255
            TabIndex        =   69
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1440
            Width           =   1650
         End
         Begin VB.TextBox Total 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   3255
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   225
            Width           =   1650
         End
         Begin VB.TextBox Total 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   3255
            TabIndex        =   67
            TabStop         =   0   'False
            Text            =   " "
            Top             =   510
            Width           =   1650
         End
         Begin VB.TextBox Total 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   2
            Left            =   3255
            TabIndex        =   66
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1155
            Width           =   1650
         End
         Begin VB.TextBox PORDESCU 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1785
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   510
            Width           =   1065
         End
         Begin VB.TextBox Total 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   4
            Left            =   3255
            TabIndex        =   64
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1725
            Width           =   1650
         End
         Begin VB.TextBox Total 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   5
            Left            =   3255
            TabIndex        =   63
            TabStop         =   0   'False
            Text            =   " "
            Top             =   2010
            Width           =   1650
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Bonf.:"
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
            Index           =   14
            Left            =   -210
            TabIndex        =   118
            Top             =   845
            Width           =   1905
         End
         Begin VB.Label Ltota 
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   13
            Left            =   2940
            TabIndex        =   117
            Top             =   845
            Width           =   225
         End
         Begin VB.Label Label18 
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
            Height          =   270
            Left            =   420
            TabIndex        =   113
            Top             =   615
            Width           =   435
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Items:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   105
            TabIndex        =   112
            Top             =   375
            Width           =   645
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Percepcion I.Brutos:"
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
            Left            =   1050
            TabIndex        =   111
            ToolTipText     =   "Doble Click para Configurar Cliente"
            Top             =   2465
            Width           =   2115
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "T-C:"
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
            Left            =   60
            TabIndex        =   107
            Top             =   2820
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Label MONEMIS 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Pesos"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00005500&
            Height          =   330
            Left            =   0
            TabIndex        =   106
            ToolTipText     =   "Doble-Click para 'Dólares'"
            Top             =   960
            Visible         =   0   'False
            Width           =   1380
         End
         Begin VB.Label BASIVA 
            Alignment       =   1  'Right Justify
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
            Height          =   270
            Index           =   2
            Left            =   1785
            TabIndex        =   93
            Top             =   2010
            Width           =   1380
         End
         Begin VB.Label BASIVA 
            Alignment       =   1  'Right Justify
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
            Height          =   270
            Index           =   1
            Left            =   1785
            TabIndex        =   92
            Top             =   1725
            Width           =   1380
         End
         Begin VB.Label BASIVA 
            Alignment       =   1  'Right Justify
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
            Height          =   270
            Index           =   0
            Left            =   1785
            TabIndex        =   91
            Top             =   1440
            Width           =   1380
         End
         Begin VB.Label PORIVA 
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
            Height          =   330
            Index           =   2
            Left            =   735
            TabIndex        =   90
            Top             =   2025
            Width           =   645
         End
         Begin VB.Label PORIVA 
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
            Height          =   330
            Index           =   1
            Left            =   735
            TabIndex        =   89
            Top             =   1755
            Width           =   645
         End
         Begin VB.Label PORIVA 
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
            Height          =   330
            Index           =   0
            Left            =   735
            TabIndex        =   88
            Top             =   1470
            Width           =   645
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ImporteTotal Neto:"
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
            Index           =   10
            Left            =   105
            TabIndex        =   81
            Top             =   1215
            Width           =   3060
         End
         Begin VB.Label Ltota 
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   9
            Left            =   1470
            TabIndex        =   79
            Top             =   2025
            Width           =   225
         End
         Begin VB.Label Ltota 
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   8
            Left            =   1470
            TabIndex        =   78
            Top             =   1755
            Width           =   225
         End
         Begin VB.Label Ltota 
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   7
            Left            =   1470
            TabIndex        =   77
            Top             =   1470
            Width           =   225
         End
         Begin VB.Label Ltota 
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   6
            Left            =   2940
            TabIndex        =   76
            Top             =   555
            Width           =   225
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Bruto sin IVA:"
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
            Index           =   0
            Left            =   105
            TabIndex        =   75
            Top             =   270
            Width           =   3060
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Desc:"
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
            Index           =   1
            Left            =   -210
            TabIndex        =   74
            Top             =   555
            Width           =   1905
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "IVA"
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
            Left            =   -210
            TabIndex        =   73
            Top             =   1470
            Width           =   855
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Documento:"
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
            Index           =   3
            Left            =   1260
            TabIndex        =   72
            Top             =   2835
            Width           =   1905
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "IVA"
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
            Index           =   4
            Left            =   -420
            TabIndex        =   71
            Top             =   1755
            Width           =   1065
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "IVA-NI"
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
            Index           =   5
            Left            =   -315
            TabIndex        =   70
            Top             =   2025
            Width           =   960
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Condiciones Comerciales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   1545
         Left            =   210
         TabIndex        =   43
         Top             =   0
         Width           =   5330
         Begin VB.TextBox Eco21 
            BackColor       =   &H00E0E0E0&
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
            Index           =   2
            Left            =   1890
            TabIndex        =   46
            Text            =   " "
            Top             =   1130
            Width           =   3270
         End
         Begin VB.TextBox TEXT21 
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
            Index           =   2
            Left            =   1155
            TabIndex        =   17
            Text            =   " "
            Top             =   1130
            Width           =   435
         End
         Begin VB.CommandButton Command21 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   1575
            TabIndex        =   18
            Top             =   1130
            Width           =   180
         End
         Begin VB.TextBox Eco21 
            BackColor       =   &H00E0E0E0&
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
            Index           =   1
            Left            =   1890
            TabIndex        =   45
            Text            =   " "
            Top             =   775
            Width           =   3270
         End
         Begin VB.TextBox TEXT21 
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
            Index           =   1
            Left            =   1155
            TabIndex        =   15
            Text            =   " "
            Top             =   775
            Width           =   435
         End
         Begin VB.CommandButton Command21 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   1575
            TabIndex        =   16
            Top             =   775
            Width           =   180
         End
         Begin VB.TextBox Eco21 
            BackColor       =   &H00E0E0E0&
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
            Index           =   0
            Left            =   1890
            TabIndex        =   44
            Text            =   " "
            Top             =   420
            Width           =   3270
         End
         Begin VB.TextBox TEXT21 
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
            Index           =   0
            Left            =   1155
            TabIndex        =   13
            Text            =   " "
            Top             =   420
            Width           =   435
         End
         Begin VB.CommandButton Command21 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   1575
            TabIndex        =   14
            Top             =   420
            Width           =   180
         End
         Begin VB.Label Label21 
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
            Height          =   330
            Index           =   2
            Left            =   -3045
            TabIndex        =   49
            Top             =   1205
            Width           =   4110
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "C.Pago:"
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
            Index           =   1
            Left            =   -3045
            TabIndex        =   48
            Top             =   850
            Width           =   4110
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "L.Precios:"
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
            Index           =   0
            Left            =   -525
            TabIndex        =   47
            Top             =   495
            Width           =   1590
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8520
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   12000
      Begin VB.PictureBox MARSELPENDE 
         BackColor       =   &H0080C0FF&
         Height          =   3180
         Left            =   210
         ScaleHeight     =   3120
         ScaleWidth      =   10455
         TabIndex        =   56
         Top             =   1785
         Visible         =   0   'False
         Width           =   10520
         Begin VB.ListBox List2 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2760
            Left            =   105
            Style           =   1  'Checkbox
            TabIndex        =   61
            Top             =   420
            Width           =   9650
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H0080C0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   3165
            Left            =   0
            TabIndex        =   85
            Top             =   105
            Width           =   960
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H0080C0FF&
            Height          =   470
            Left            =   -105
            ScaleHeight     =   405
            ScaleWidth      =   9795
            TabIndex        =   59
            Top             =   -30
            Width           =   9850
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Selección de Pedidos Pendientes de Entrega"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404000&
               Height          =   260
               Left            =   0
               TabIndex        =   60
               Top             =   105
               Width           =   9825
            End
         End
         Begin VB.CommandButton Command19 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   9820
            Picture         =   "FORFACLI.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   58
            ToolTipText     =   "Continúa Generación de la Factura"
            Top             =   420
            Width           =   570
         End
         Begin VB.CommandButton Command18 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   9820
            Picture         =   "FORFACLI.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   57
            ToolTipText     =   "Cancela la Factura Activa"
            Top             =   1050
            Width           =   570
         End
         Begin VB.Line Line5 
            X1              =   9750
            X2              =   9750
            Y1              =   0
            Y2              =   3200
         End
         Begin VB.Line Line4 
            X1              =   10575
            X2              =   10575
            Y1              =   0
            Y2              =   3150
         End
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00E0E0E0&
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   10920
         Picture         =   "FORFACLI.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Aprueba, Graba e Imprime Documento"
         Top             =   7300
         Width           =   855
      End
      Begin VB.CommandButton Command17 
         BackColor       =   &H00FFFF80&
         Height          =   645
         Left            =   6640
         Picture         =   "FORFACLI.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   104
         ToolTipText     =   "Pedidos pendientes de Despacho, Remito o Factura"
         Top             =   945
         Width           =   670
      End
      Begin VB.TextBox Text14 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8625
         MaxLength       =   16
         TabIndex        =   102
         Top             =   1280
         Width           =   2115
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   4830
         TabIndex        =   101
         Text            =   "Text13"
         Top             =   5500
         Width           =   960
      End
      Begin VB.PictureBox MARCONOTA 
         BackColor       =   &H00FFFFC0&
         Height          =   3180
         Left            =   210
         ScaleHeight     =   3120
         ScaleWidth      =   10455
         TabIndex        =   51
         Top             =   1785
         Visible         =   0   'False
         Width           =   10520
         Begin VB.CommandButton Command16 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   9820
            Picture         =   "FORFACLI.frx":0D60
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Imprime Ficha de Anotaciones"
            Top             =   2470
            Width           =   570
         End
         Begin VB.CommandButton Command15 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   9820
            Picture         =   "FORFACLI.frx":13CA
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Cancela la Factura Activa"
            Top             =   1050
            Width           =   570
         End
         Begin VB.CommandButton Command6 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   9820
            Picture         =   "FORFACLI.frx":16D4
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Continúa Generación de la Factura"
            Top             =   420
            Width           =   570
         End
         Begin VB.TextBox Text6 
            BackColor       =   &H00C0FFFF&
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
            Left            =   -30
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   54
            Top             =   390
            Width           =   9780
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FFFFC0&
            Height          =   435
            Left            =   -30
            ScaleHeight     =   375
            ScaleWidth      =   9720
            TabIndex        =   52
            Top             =   -30
            Width           =   9780
            Begin VB.Label Label8 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Anotaciones de Clientes"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808000&
               Height          =   225
               Left            =   0
               TabIndex        =   53
               Top             =   105
               Width           =   9825
            End
         End
         Begin VB.Line Line2 
            X1              =   10575
            X2              =   10575
            Y1              =   0
            Y2              =   3150
         End
         Begin VB.Line Line1 
            X1              =   9750
            X2              =   9750
            Y1              =   0
            Y2              =   3200
         End
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   0
         Top             =   0
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   6950
         Left            =   10920
         TabIndex        =   35
         Top             =   105
         Width           =   850
         Begin VB.CommandButton Command13 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Opción"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   770
            Left            =   0
            Picture         =   "FORFACLI.frx":19DE
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Anulaciones - Diferencia de Cambio"
            Top             =   4620
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton Command20 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Anul"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   770
            Left            =   0
            Picture         =   "FORFACLI.frx":1B28
            Style           =   1  'Graphical
            TabIndex        =   108
            ToolTipText     =   "Anulaciones - Débitos y Créditos"
            Top             =   4642
            Width           =   855
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00E0E0E0&
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
            Height          =   770
            Left            =   0
            Picture         =   "FORFACLI.frx":1E32
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Cierra y Abandona la Aplicación"
            Top             =   0
            Width           =   855
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "C.Cte."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   770
            Left            =   0
            Picture         =   "FORFACLI.frx":1F7C
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Resumen de Cuenta Corriente"
            Top             =   3090
            Width           =   855
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H00E0E0E0&
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
            Height          =   770
            Left            =   0
            Picture         =   "FORFACLI.frx":2286
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Configuración de Funcionamiento"
            Top             =   5410
            Width           =   855
         End
         Begin VB.CommandButton Command7 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Comerc."
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   770
            Left            =   0
            Picture         =   "FORFACLI.frx":26C8
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Acceso a Listas de Precio, Condiciones de Pago y Cotizacion Dolar"
            Top             =   6170
            Width           =   855
         End
         Begin VB.CommandButton Command8 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Articulos"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   770
            Left            =   0
            Picture         =   "FORFACLI.frx":29D2
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Acceso a Maestro de Artículos"
            Top             =   1540
            Width           =   855
         End
         Begin VB.CommandButton Command11 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Clientes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   770
            Left            =   0
            Picture         =   "FORFACLI.frx":2CDC
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Acceso a Base de Datos de Clientes"
            Top             =   2315
            Width           =   855
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Help"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   770
            Left            =   0
            Picture         =   "FORFACLI.frx":311E
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Ayuda FLEXOFT en Línea"
            Top             =   770
            Width           =   855
         End
         Begin VB.CommandButton Command12 
            BackColor       =   &H00E0E0E0&
            Caption         =   "B-Code"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   770
            Left            =   0
            Picture         =   "FORFACLI.frx":3428
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Por Codigo de Barras"
            Top             =   3860
            Width           =   855
         End
      End
      Begin VB.CommandButton Command10 
         Height          =   260
         Left            =   420
         TabIndex        =   34
         Top             =   6855
         Width           =   2825
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1470
         MaxLength       =   40
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1280
         Width           =   3480
      End
      Begin VB.TextBox Text11 
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
         Height          =   300
         Left            =   9615
         TabIndex        =   11
         Text            =   " "
         Top             =   490
         Width           =   1120
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9615
         MaxLength       =   16
         TabIndex        =   12
         Text            =   " "
         Top             =   890
         Width           =   1120
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00E0E0E0&
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
         Left            =   9615
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   " "
         Top             =   105
         Width           =   960
      End
      Begin VB.CommandButton Command9 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10560
         TabIndex        =   10
         Top             =   105
         Width           =   180
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00E0E0E0&
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
         Left            =   5145
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   " "
         Top             =   490
         Width           =   1380
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00E0E0E0&
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
         Left            =   5145
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1280
         Width           =   1380
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00E0E0E0&
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
         Left            =   2310
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   " "
         Top             =   890
         Width           =   2640
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1470
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   " "
         Top             =   890
         Width           =   750
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1470
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   " "
         Top             =   490
         Width           =   3480
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1470
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   " "
         Top             =   105
         Width           =   3480
      End
      Begin VB.CommandButton Command1 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1050
         TabIndex        =   1
         Top             =   1280
         Width           =   180
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
         Left            =   210
         MaxLength       =   5
         TabIndex        =   0
         Text            =   " "
         Top             =   1280
         Width           =   855
      End
      Begin VB.ListBox List1 
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
         Left            =   210
         TabIndex        =   55
         Top             =   2205
         Width           =   10515
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   210
         ScaleHeight     =   375
         ScaleWidth      =   10455
         TabIndex        =   32
         Top             =   1785
         Width           =   10515
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "  Cant Código          Descripción                  Lista  Desc  P.Unit     Importe  R."
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   225
            Left            =   30
            TabIndex        =   33
            Top             =   105
            Width           =   10515
         End
      End
      Begin VB.Label FVENCIM 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F-VENCIM"
         Height          =   330
         Left            =   7560
         TabIndex        =   114
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label DOCUDIF 
         BackColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   5040
         TabIndex        =   109
         Top             =   0
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "O.Compra:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   330
         Left            =   6720
         TabIndex        =   103
         Top             =   1320
         Width           =   1800
      End
      Begin VB.Line Line9 
         BorderWidth     =   2
         X1              =   6975
         X2              =   6975
         Y1              =   -45
         Y2              =   110
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         X1              =   6975
         X2              =   6975
         Y1              =   825
         Y2              =   1790
      End
      Begin VB.Shape Shape4 
         BorderWidth     =   2
         Height          =   705
         Left            =   6660
         Top             =   105
         Width           =   645
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   750
         Left            =   6720
         TabIndex        =   87
         Top             =   30
         Width           =   540
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   7665
         TabIndex        =   86
         Top             =   885
         Width           =   1275
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Factura:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   7245
         TabIndex        =   84
         Top             =   450
         Width           =   1590
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   9510
         TabIndex        =   83
         Top             =   460
         Width           =   150
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0004"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   8940
         TabIndex        =   82
         Top             =   490
         Width           =   585
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Fc"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   270
         TabIndex        =   62
         Top             =   150
         Width           =   1380
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00C0C000&
         FillColor       =   &H00FFFF80&
         FillStyle       =   0  'Solid
         Height          =   250
         Left            =   420
         Shape           =   4  'Rounded Rectangle
         Top             =   470
         Width           =   540
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H0000C000&
         FillColor       =   &H0080FF80&
         FillStyle       =   0  'Solid
         Height          =   390
         Left            =   315
         Shape           =   4  'Rounded Rectangle
         Top             =   395
         Width           =   750
      End
      Begin VB.Line Line7 
         X1              =   210
         X2              =   10710
         Y1              =   4970
         Y2              =   4970
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Disponible:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   330
         Left            =   5145
         TabIndex        =   31
         Top             =   975
         Width           =   1380
      End
      Begin VB.Line Line8 
         X1              =   10730
         X2              =   10730
         Y1              =   1785
         Y2              =   4950
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Rto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   330
         Left            =   7770
         TabIndex        =   29
         Top             =   940
         Width           =   1800
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   330
         Left            =   8820
         TabIndex        =   28
         Top             =   150
         Width           =   750
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Lim.Crédito:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   330
         Left            =   5145
         TabIndex        =   27
         Top             =   195
         Width           =   1380
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   330
         Left            =   315
         TabIndex        =   26
         Top             =   945
         Width           =   960
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000C0C0&
         FillColor       =   &H0080FFFF&
         FillStyle       =   0  'Solid
         Height          =   540
         Left            =   210
         Shape           =   4  'Rounded Rectangle
         Top             =   315
         Width           =   960
      End
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1050
   End
End
Attribute VB_Name = "FORFACLI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TIFADOC%    ' 2=FACTURA / 3=N.DEBITO / 6=N.CREDITO
Dim MODODOC%    ' 0=MAT.CODIFICADO / 1=TEXTO LIBRE
Dim VARIDOC%
Dim RNIDOC%
Dim DATEDOC%
Dim PIVADOC%
Dim VERLICRE%
Dim MODESCU%
'
Dim CONSICLI%
'
Dim NUOCOMPRA$
Dim NROPEDIDO$
Dim NROREMICO$
'
Dim ARCHIPRE$
Dim DATCLI$(32)
'
Dim MONECO$, TICAMBX#, TIPOVENTA$, PUERTO$

Private Sub Command1_Click()
   Call SELECHO("DEUDOR1")
   NCLIE% = Val(TRIM$(VALACT1$("DEUDOR1")))
   If NCLIE% > 0 Then
     Text1.TEXT = TRIM$(Str$(NCLIE%))
     If Val(Text1.TEXT) > 0 Then
       If Val(Text1.TEXT) <= 32767 Then
         NC% = Val(Text1)
         If REGICLI&(NC%) > 0 Then
           Call CARDACLI
         End If
       End If
     End If
   End If
   End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call CIERRARCH("*.*")
   ACONEC$ = "AMACLI02"
   If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "AMACLI01"
   Call CONECRUN(ACONEC$, "Padron Maestro de Clientes")
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   If TIFADOC% = 2 Then
     If MODODOC% = 0 Then
       NCLIE% = Val(Text1)
       If NCLIE% > 0 Then
         If RASOCLI$(NCLIE%) <> "" Then
           '
           NUOCOMPRA$ = ""
           NROPEDIDO$ = ""
           MARCONOTA.Visible = False
           MARSELPENDE.Visible = False
           '
           CARCOBARR.Label9 = Text1
           CARCOBARR.Label10 = Text2
           CARCOBARR.Label13 = Text3
           CARCOBARR.Label14 = Text4
           CARCOBARR.Label15 = Text5
           CARCOBARR.Label16 = Text12
           CARCOBARR.Label17 = Text9
           CARCOBARR.Label12 = Label12
           CARCOBARR.Label18 = Text11
           CARCOBARR.List1.Clear
           '
           UREPRE& = ULTREG&("PEDDETA")
           Screen.MousePointer = 11
           CARCOBARR.List2.Clear
           '
           RFF$ = RECFILT$("PEDDETA", 3, MKI$(NCLIE%))
           For KKI% = 1 To Len(RFF$) Step 4
             REPRESU& = CVS(Mid$(RFF$, KKI%, 4))
             If REPRESU& > 0 Then
               If REPRESU& <= UREPRE& Then
12               X$ = REGLEIDO$("PEDDETA", REPRESU&)
                 '
                 If Asc(Mid$(X$, 27, 1)) <> 9 Then ' no esta anulado
                   CAN = CVS(Mid$(X$, 79, 4))
                   CADESPACH = CVS(Mid$(X$, 95, 4))
                   CAPENDENT = CAN - CADESPACH
                   If CAPENDENT > 0.5 Then
                     If CAN >= 0.05 Then
                       CIXI% = CVI(Mid$(X$, 33, 2))
                       If CIXI% > 0 Then
                         NRODES& = CVS(Left$(X$, 4))
                         If Mid$(X$, 51, 1) <> "R" Then
                           FENTRE% = CVI(Mid$(X$, 25, 2))
                           '
                           TTXX$ = Space$(128)
                           Call REPLA(TTXX$, TRIM$(Str$(10000 + FENTRE%)), 1, 8)
                           Call REPLA(TTXX$, TRIM$(CODEXT$(CIXI%)), 9, 16)
                           Call REPLA(TTXX$, TRIM$(Str$(CIXI%)), 25, 8)
                           Call REPLA(TTXX$, TRIM$(FORMATNUM$(CAPENDENT, "F8.1")), 33, 8)
                           Call REPLA(TTXX$, TRIM$(Str$(NRODES&)), 41, 8)
                           Call REPLA(TTXX$, TRIM$(Str$(REPRESU&)), 49, 8) ' REGISTRO DE PEDDETA
                           CARCOBARR.List2.AddItem (TTXX$)
                         End If
                       End If
                     End If
                   End If
                 End If
               End If
             End If
19         Next KKI%
           Screen.MousePointer = 1
           '
           CARCOBARR.Show 1
           '
           If CARCOBARR.List1.ListCount < 1 Then
             On Error Resume Next
             Text1.SetFocus
             On Error GoTo 0
             GoTo 99
           End If
           Call CARMATBARR
           Call CARDETFAC
           '
         End If
       End If
     End If
   End If
99 Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   OPVARFAC.Show 1
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   '
   If Val(Text15) >= 0.005 Then
     If Val(Total(4)) >= 0.005 Then
       Call MENSERR(24, "Imposible Registrar este Documento.\Incompatibilidad entre Percepcion IB\y Tasa IVA Secundaria - Consulte.")
       Command14.Enabled = True
       Exit Sub
     End If
   End If
   '
   FORIPRE$ = TRIM$(CONTROL$("", "FORFAC-" + Label13.Caption))
   ITMAXX = Val(ATRIFORM$(FORIPRE$, "MAX-ITEMS"))
   If ITMAXX >= 1 Then
     If Val(Label18) > ITMAXX Then
       Call MENSERR(24, "Imposible Registrar este Documento.\Demasiados Renglones - Máximo " + TRIM$(Str$(Int(ITMAXX))))
       Command14.Enabled = True
       Exit Sub
     End If
   End If
     
18 EPAC$ = TRIM$(UCase$(EMPREAC$))
   '
   'PARA OBLIGAR QUE SI ES FACTURA DEL EXTERIOR EL IMPORTE QUE VIENE DEL PEDIDO ES EN DOLARES
   If InStr(EPAC$, "FERVI") > 0 And PIVADOC% = 5 Then MONEMI% = 2
   '
   FINFACEX$ = ""
   TICAMBX# = 1: MONECO$ = ""
   If PIVADOC% = 5 Then     ' FACTURAS DE EXPORTACION
     If MONEMI% > 1 Then
        FINFACEX$ = String$(88, 0)
        Call REPLA(FINFACEX$, MKD$(Val(Total(6))), 1, 8)
        Call REPLA(FINFACEX$, MKD$(Val(Total(6))), 41, 8)
        Call REPLA(FINFACEX$, "DOLARES     ", 49, 12)
        Call REPLA(FINFACEX$, MKD$(XVALO(TICAMBIS)), 61, 8)
        Call REPLA(FINFACEX$, "FOB                 ", 69, 20)
        '
        VDEF$ = FINFACEX$
19      OBX$ = OBJPLA$("TICAEXPOR", VDEF$)
        If OBX$ = "" Then
           Call COMUNI("Operación Cancelada !!!")
           Command14.Enabled = True
           Exit Sub
        End If
        '
        VDEF$ = OBX$
        TICAMBX# = CVD(Mid$(OBX$, 61, 8))
        If TICAMBX# < 0.1 Then
           Call MENSERR(24, "Ingrese Tipo de Cambio")
           GoTo 19
        End If
        '
        MONECO$ = TRIM$(Mid$(OBX$, 49, 12))
        If MONECO$ = "" Then
           Call MENSERR(24, "Ingrese Moneda")
           GoTo 19
        End If
        TICADOL = TICAMBX#
        TIPOVENTA$ = TRIM$(Mid$(VDEF, 69, 20))
        If TIPOVENTA$ = "" Then Call MENSERR(24, "Datos Incompletos!!!"): GoTo 19
        PUERTO$ = TRIM$(Mid$(VDEF, 89, 40))
        If PUERTO$ = "" Then Call MENSERR(24, "Datos Incompletos!!!"): GoTo 19
        '
        If InStr(EPAC$, "DOSIVA") > 0 Or InStr(EPAC$, "FERVI-AIR") > 0 Then
           TICADOLAR = FORMATNUM$(TICAMBX#, "F10.4")
           FINFACTURA.Text4.TEXT = "Tipo de Cambio: U$S 1 - $" + TICADOLAR
        End If
        '
     End If
   End If
   '
20 NCLIE% = Val(Text1.TEXT)
   If VERLICRE% = 1 Then
     If TIFADOC% < 4 Then ' EXCLUYE CREDITOS
       '
       For U& = 1 To ULTREG&("!CARDEFAC")
         X$ = REGLEIDO$("!CARDEFAC", U&)
         NROPEDXX& = Val(TRIM$(Mid$(X$, 371, 16)))
         If NROPEDXX& < 1 Then GoTo 22
         RFF$ = RECFILT$("AUTORENT", 1, MKS$(NROPEDXX&))
         If Len(RFF$) < 4 Then GoTo 22 ' NO AUTORIZADO
         REGAUT& = CVS(RFF$)
         XAUT$ = REGLEIDO$("AUTORENT", REGAUT&)
         If UCase$(Mid$(XAUT$, 31, 2)) <> "SI" Then GoTo 22
       Next U&
       GoTo 25
       '
22     Call CIERRARCH("DEUDOR2")
       RCLI& = REGICLI&(NCLIE%)
       If RCLI& > 0 Then
         DD2$ = REGLEIDO$("DEUDOR2", RCLI&)
         LMCR# = CVS(Mid$(DD2$, 121, 4)) ' LIMICRE(NCLIE%)
       End If
       SDEU# = SALDCLI#(NCLIE%)
       CHPA# = CHEPEACRE#(NCLIE%)
       TPED# = Val(Total(6).TEXT) * TICAMBX#
       If SDEU# + CHPA# + TPED# > LMCR# Then
         TLIMI$ = TRIM$(CSTRING$(MKD$(LMCR#), 3, 12, 2, 2))
         If TLIMI$ = "" Then TLIMI$ = "0.00"
         Call COMUNI("Imposible Emitir Este Documento.\  \Supera el Límite de Crédito Asignado\para esta Cuenta de $ " + TLIMI$ + ".\  \Informe al Supervisor de Ventas.")
         Command14.Enabled = True
         Call CIERRARCH("*.*")
         Exit Sub
       End If
     End If
   End If
   '
25 LENTREG$ = Space$(96)
   Call REPLA(LENTREG$, TRIM$(LENTREGA(0).TEXT), 1, 32)
   Call REPLA(LENTREG$, TRIM$(LENTREGA(1).TEXT), 33, 32)
   Call REPLA(LENTREG$, TRIM$(LENTREGA(2).TEXT), 65, 32)
   '
   RELENT& = 0: RFEX$ = RECFILT$("LENTRECL", 1, MKI$(NCLIE%))
   If Len(RFEX$) >= 4 Then
     RELENT& = CVS(Left$(RFEX$, 4))
     If RELENT& < 0 Or RELENT& > ULTREG&("LENTRECL") Then RELENT& = 0
   End If
   If RELENT& > 0 Then
       'Call GRAREG("LENTRECL", MKI$(NCLIE%) + LENTREG$, RELENT&)
     Else
       Call REGAPP("LENTRECL", MKI$(NCLIE%) + LENTREG$)
   End If
   Call CIERRARCH("LENTRECL")
   '
   NUOCOMPRA$ = ""
   NROPEDIDO$ = ""
   NROREMICO$ = ""
   '
   For U& = 1 To ULTREG&("!CARDEFAC")
     X$ = REGLEIDO$("!CARDEFAC", U&)
     NUOCOPAR$ = TRIM$(Mid$(X$, 355, 16))
     NROPEDIPAR$ = TRIM$(Mid$(X$, 371, 16))
     NROREMIPAR$ = TRIM$(Mid$(X$, 390, 13))
     '
     If NUOCOPAR$ <> "" Then
       If NUOCOPAR$ <> NUOCOPARA$ Then
         If NUOCOMPRA$ <> "" Then NUOCOMPRA$ = NUOCOMPRA$ + ", "
         NUOCOMPRA$ = NUOCOMPRA$ + NUOCOPAR$
         NUOCOPARA$ = NUOCOPAR$
       End If
     End If
     '
     If NROPEDIPAR$ <> "" Then
       If NROPEDIPAR$ <> NROPEDIPARA$ Then
         If NROPEDIDO$ <> "" Then NROPEDIDO$ = NROPEDIDO$ + ", "
         NROPEDIDO$ = NROPEDIDO$ + NROPEDIPAR$
         NROPEDIPARA$ = NROPEDIPAR$
       End If
     End If
     '
     If CONSICLI% = 1 Then
       If NROREMIPAR$ <> "" Then
         If NROREMIPAR$ <> NROREMIPARA$ Then
           If NROREMICO$ <> "" Then NROREMICO$ = NROREMICO$ + ", "
           NROREMICO$ = NROREMICO$ + NROREMIPAR$
           NROREMIPARA$ = NROREMIPAR$
         End If
       End If
     End If
     '
   Next U&
   '
   NPEDID& = Val(NROPEDIDO$)
   If NPEDID& > 0 Then
     For UU& = ULTREG&("PEDENCA") To 1 Step -1
       XPED$ = REGLEIDO$("PEDENCA", UU&)
       If CVS(Left$(XPED$, 4)) = NPEDID& Then
         LENTREG$ = Mid$(XPED$, 73, 96)
         LENTREGA(0).TEXT = Mid$(LENTREG$, 1, 32)
         LENTREGA(1).TEXT = Mid$(LENTREG$, 33, 32)
         LENTREGA(2).TEXT = Mid$(LENTREG$, 65, 32)
         GoTo 30
       End If
     Next UU&
   End If
   '
30 If CONSICLI% = 1 Then
     REMI = 0
     '
     If TRIM$(Text14.TEXT) = "" Then
       Text14.TEXT = NUOCOMPRA$
     End If
     '
   End If
   '
   Call GRAFACTURA(OKX%)
   '\\\OT-05-0524-WAR-25/07/05///
   If OKX% = 1 Then
     Call GRAFABON
   End If
   '\\\OT-05-0524-WAR-FIN///
   Call CIERRARCH("*.*")
   Call LIBARCH("*.*")
   Call BAJALDISCO
   '
   If OKX% < 0 Then GoTo 10     ' ABORTADA
   '
   If TIFADOC% <= 2 Then ' SOLO FACTURAS
     COPIASREM% = Val(FINFACTURA.Text2(1))
     If COPIASREM% > 0 Then
       For U& = 1 To ULTREG&("!CARDEFAC")
         X$ = REGLEIDO$("!CARDEFAC", U&)
         CADESPX = CVS(Left$(X$, 4))
         If CADESPX >= 0.05 Then
           NROPEDXX& = Val(TRIM$(Mid$(X$, 371, 16)))
           If Abs(NROPEDXX&) < 1 Then
             CIKKI% = CVI(Mid$(X$, 5, 2))
             If CIKKI% > 0 Then
               If Mid$(X$, 461, 1) <> "R" Then
                 PDC$ = REGBLAN$("DESPAPED")
                 Call REPLA(PDC$, MKS$(900000 + Val(Text10)), 1, 4)
                 Call REPLA(PDC$, MKI$(Int(FECHANUM(Text9))), 5, 2)
                 Call REPLA(PDC$, MKI$(NCLIE%), 7, 2)
                 Call REPLA(PDC$, TRIM$(Text14.TEXT), 9, 16)
                   FENTRESOL% = FECHANUM(Text9)
                 Call REPLA(PDC$, MKI$(FENTRESOL%), 25, 2)
                 Call REPLA(PDC$, Chr$(1), 27, 1)
                 Call REPLA(PDC$, Chr$(Val(TEXT21(0).TEXT)), 28, 1)
                 Call REPLA(PDC$, Chr$(Val(TEXT21(1).TEXT)), 29, 1)
                 Call REPLA(PDC$, Chr$(Val(TEXT21(2).TEXT)), 30, 1)
                 Call REPLA(PDC$, MKI$(0), 31, 2)
                 Call REPLA(PDC$, MKI$(CIKKI%), 33, 2)
                 Call REPLA(PDC$, Mid$(X$, 291, 12), 35, 16)
                 Call REPLA(PDC$, Mid$(X$, 283, 8), 51, 8)
                 Call REPLA(PDC$, Mid$(X$, 303, 12), 59, 12)
                 Call REPLA(PDC$, String$(58, 0), 71, 58)
                 Call REPLA(PDC$, MKS$(CADESPX), 75, 4)
                 Call REPLA(PDC$, MKS$(CADESPXN), 99, 4)
                 Call REGAPP("DESPAPED", PDC$)
               End If
             End If
           End If
         End If
       Next U&
     End If
   End If
   Call CIERRARCH("DESPAPED")
   '
   If OKX% > 0 Then
       FF% = FECHANUM(Text9.TEXT)
       REMI = Val(Text10.TEXT)
       Call BLANPEFAC(REMI, FF%)
       Call PRIFACTURA
   End If
   Call CIERRARCH("*.*")
   Call LIBARCH("*.*")
   Call BAJALDISCO
   '
Call CIERRARCH("*.*")
Reset
Call CONECRUN(App.EXEName, "Facturacion en Cuenta Corriente")
Close: End
   '
10 Call BLANFORMU
   Call CIERRARCH("*.*")
   Text1.TEXT = ""
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
   '
   Command14.Enabled = True
End Sub

Private Sub Command15_Click()
   MARCONOTA.Visible = False
   Text1.TEXT = ""
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Command16_Click()
   Call MENSERR(24, "No Disponible por el Momento")
End Sub

Private Sub Command16_LostFocus()
   On Error Resume Next
   Command6.SetFocus
   On Error GoTo 0
End Sub


Private Sub Command17_Click()
   Command17.Enabled = False
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "DOSIVA") > 0 Then
       CONPEDID.Show 1
     Else
       Call SHOWPEDIPEN
   End If
   Command17.Enabled = True
End Sub

Private Sub Command18_Click()
   MARSELPENDE.Visible = False
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Command19_Click()
   '
   NUOCOMPRA$ = ""
   NROPEDIDO$ = ""
   '
   PREFACTU$ = CONTROL("FACTURA", "PREFACTU")
   If PREFACTU$ = "ORDESPA" Then
       Call CARMATODES
     Else
       Call CARMATPEDI
   End If
   '
   If ULTREG&("!CARDEFAC") > 0 Then
     Call CALTOFAC
     If Val(Total(0).TEXT) < 0.005 Then
       Call COMUNI("ATENCION: Los Pedidos Seleccionados\Parecen Ser SIN CARGO !!!")
     End If
   End If
   '
   MARSELPENDE.Visible = False
   Call CARDETFAC
   '
End Sub

Private Sub Command2_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command2_LostFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub


Private Sub Command20_Click()
   Command20.Enabled = False
   If DERACCE%("ANULAFAC", "Anulación de Comprobantes de Facturación") = 2 Then
     TTXX$ = "Pueden Anularse Comprobantes Emitidos\o Formularios Inutilizados"
     MODO% = COMALTER%(TTXX$ + "\\Comprobantes Emitidos\Formularios Arruinados\Cancelar")
     If MODO% = 1 Then Call ANUFAC
     If MODO% = 2 Then Call ANUFORM
   End If
   Command20.Enabled = True
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call HELPONLINE("PRESUP01")
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    RESCUDEU.Text1.TEXT = Text1.TEXT
    PROLLAMA$ = "FORFACLI"
    Hide
    RESCUDEU.Show
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    OPFACTU01.Show 1
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   '
   MARCONOTA.Visible = False
   DOCUDIF = ""
   '
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   '
   CONSICLI% = 0
   If TIFADOC% = 2 Then    ' 2=FACTURA / 3=N.DEBITO / 6=N.CREDITO
     If MODODOC% = 0 Then  ' 0=MAT.CODIFICADO / 1=TEXTO LIBRE
       Call CARCONSIPEN
       If CONSICLI% < 1 Then
           Call SELPEDIPEN
         Else
           Call CARDETFAC
       End If
       Exit Sub
     End If
   End If
   '
   If TIFADOC% = 3 Or TIFADOC% = 6 Then ' débitos o créditos
     If MODODOC% = 1 Then        ' TEXTO LIBRE
       Call CARDICAMBIO
     End If
   End If
   Call CARDETFAC
   '
End Sub
'
Sub SELORDESPEN()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   '
   UREPRE& = ULTREG&("ORDESPA")
10 Screen.MousePointer = 11
   List2.Clear
   NCLIE% = Val(Text1.TEXT)
   NRODESA& = 0: CAPED% = 0
   FINF = 999999: FSUP = 0
   RFF$ = RECFILT$("ORDESPA", 5, MKI$(NCLIE%))
   For KKI% = 1 To Len(RFF$) Step 4
      REPRESU& = CVS(Mid$(RFF$, KKI%, 4))
      If REPRESU& > 0 Then
        If REPRESU& <= UREPRE& Then
12        X$ = REGLEIDO$("ORDESPA", REPRESU&)
          If CVI(Mid$(X$, 13, 2)) <> NCLIE% Then
            X$ = REGLEIDO$("ORDESPA", 1)
            Call GRAREG("ORDESPA", X$, 1)
            Call CIERRARCH("ORDESPA")
            GoTo 10
          End If
          '
          CAN = CVS(Mid$(X$, 79, 4))
          CADESPACH = CVS(Mid$(X$, 83, 4))
          If CADESPACH < 0.05 Then
            If CAN >= 0.05 Then
              If TIDOC% <> 3 And VERISTO% = 1 Then
                CIXI% = CVI(Mid$(X$, 37, 2))
                If CIXI% > 0 Then
                  STACT = STOCKACT(CI%)
                  STDIS = STACT - CAN
                  If STACT < 0.05 Then
                    GoTo 19
                  End If
                End If
              End If
              '
              NRODES& = CVS(Left$(X$, 4))
              If NRODES& <> NRODESA& Then
                 TTXX$ = "Or.Desp." + TRIM$(CSTRING$(MKS$(NRODES&), 3, 8, 0, 2))
                 FEX = CVI(Mid$(X$, 11, 2))
                 TTXX$ = TTXX$ + " - Pedido " + TRIM$(CSTRING$(Mid$(X$, 7, 4), 3, 8, 0, 2)) + " del " + FECHATEX$(FEX)
                 FEX = CVI(Mid$(X$, 31, 2))
                 TTXX$ = TTXX$ + " - Entrega Solic. " + FECHATEX$(FEX)
                 TTXX$ = AJUSTI$(TTXX$, 128)
                 Call REPLA(TTXX$, Str$(REPRESU&), 121, 8)
                 List2.AddItem TTXX$
                 NRODESA& = NRODES&
                 CAPED% = CAPED% + 1
                 FENTRE% = CVI(Mid$(X$, 31, 2))
                 If FENTRE% < FINF Then FINF = FENTRE%
                 If FENTRE% > FSUP Then FSUP = FENTRE%
              End If
            End If
          End If
        End If
      End If
19 Next KKI%
   Screen.MousePointer = 1
   '
   MARSELPENDE.Visible = True
   '
End Sub
'
Sub SELPEDIPEN()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   '
   PREFACTU$ = CONTROL("FACTURA", "PREFACTU")
   If PREFACTU$ = "ORDESPA" Then
      Call SELORDESPEN
      Exit Sub
   End If
   '
   UREPRE& = ULTREG&("PEDDETA")
10 Screen.MousePointer = 11
   List2.Clear
   NCLIE% = Val(Text1.TEXT)
   NROPEDA& = 0: CAPED% = 0
   FINF = 999999: FSUP = 0
   RFF$ = RECFILT$("PEDDETA", 3, MKI$(NCLIE%))
   For KKL& = 1 To Len(RFF$) Step 4
      REPRESU& = CVS(Mid$(RFF$, KKL&, 4))
      If REPRESU& > 0 Then
        If REPRESU& <= UREPRE& Then
12        X$ = REGLEIDO$("PEDDETA", REPRESU&)
          If CVI(Mid$(X$, 7, 2)) <> NCLIE% Then
            X$ = REGLEIDO$("PEDDETA", 1)
            Call GRAREG("PEDDETA", X$, 1)
            Call CIERRARCH("PEDDETA")
            GoTo 10
          End If
          '
          If Asc(Mid$(X$, 27, 1)) <> 9 Then ' no esta anulado
            CAN = CVS(Mid$(X$, 79, 4))
            CADESPACH = CVS(Mid$(X$, 95, 4))
            CAPENDENT = CAN - CADESPACH
            If CAPENDENT > 0.5 Then
              If CAN >= 0.05 Then
                If TIDOC% < 3 And VERISTO% = 1 Then
                  CIXI% = CVI(Mid$(X$, 37, 2))
                  If CIXI% > 0 Then
                    STACT = STOCKACT(CI%)
                    STDIS = STACT - CAPENDENT
                    If STDIS < 0.05 Then
                      GoTo 19
                    End If
                  End If
                End If
                '
                NRODES& = CVS(Left$(X$, 4))
                If NRODES& <> NRODESA& Then
                   TTXX$ = "Pedido Cliente " + TRIM$(CSTRING$(MKS$(NRODES&), 3, 8, 0, 2))
                   If Mid$(X$, 51, 1) = "R" Then TTXX$ = TTXX$ + " (Reparación)"
                   FEX = CVI(Mid$(X$, 5, 2))
                   TTXX$ = TTXX$ + " del " + FECHATEX$(FEX)
                   FEX = CVI(Mid$(X$, 25, 2))
                   TTXX$ = TTXX$ + " - Entrega Solic. " + FECHATEX$(FEX)
                   TTXX$ = AJUSTI$(TTXX$, 128)
                   Call REPLA(TTXX$, Str$(REPRESU&), 121, 8)
                   List2.AddItem TTXX$
                   NRODESA& = NRODES&
                   CAPED% = CAPED% + 1
                   FENTRE% = CVI(Mid$(X$, 25, 2))
                   If FENTRE% < FINF Then FINF = FENTRE%
                   If FENTRE% > FSUP Then FSUP = FENTRE%
                End If
              End If
            End If
          End If
        End If
      End If
19 Next KKL&
   Screen.MousePointer = 1
   '
   MARSELPENDE.Visible = True
   '
End Sub
'
Sub CARCONSIPEN()
   '
   LPRE% = Val(TEXT21(0).TEXT)
   NCLIE% = Val(Text1.TEXT)
   If LPRE% < 1 Then
        ARCHIPRE$ = "MASTER"
      Else
        ARCHIPRE$ = TRIM$(Str$(LPRE%))
        While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
        ARCHIPRE$ = "LIPRE" + ARCHIPRE$
        If LPRE% < 1 Then ARCHIPRE$ = "LIPRE001"
   End If
   CONSICLI% = 0
   '
   RREMX$ = RECFILT$("CONSICLI", 1, MKI$(NCLIE%))
   REGODES& = 0
   Call SETULTREG("!CARDEFAC", 0)
   If Len(RREMX$) >= 4 Then
     For UU& = 1 To Len(RREMX$) Step 4
       U& = CVS(Mid$(RREMX$, UU&, 4))
       X$ = REGLEIDO$("CONSICLI", U&)
       If CVI(Left$(X$, 2)) = NCLIE% Then
         CODIMAT% = CVI(Mid$(X$, 35, 2))
         CAN = CVD(Mid$(X$, 37, 8)) - CVD(Mid$(X$, 45, 8))
         NPEDID& = CVS(Mid$(X$, 109, 4))
         If CAN > 0.01 Then GoTo 5
         '
         VUELTA% = 0
2        For JJ& = 1 To ULTREG&("!CARDEFAC")
           Z$ = REGLEIDO$("!CARDEFAC", JJ&)
           If CVI(Mid$(Z$, 5, 2)) = CODIMAT% Then
             If Val(TRIM$(Mid$(Z$, 371, 16))) = NPEDID& Or VUELTA% = 1 Then
               CAPENCO = CVS(Mid$(Z$, 1, 4))
               If Abs(CAN) <= CAPENCO + 0.01 Then
                   CAPENCO = CAN + CAPENCO
                   Call REPLA(Z$, MKS$(CAPENCO), 1, 4)
                   Call GRAREG("!CARDEFAC", Z$, JJ&)
                   CAN = 0
                   GoTo 12
                 Else
                   'esto se suprimio por el problema de ana de dosivac rto 10883 el 01/04/05
                   'CAN = CAN + CAPENCO
                   'CAPENCO = 0
                   'Call REPLA(Z$, MKS$(CAPENCO), 1, 4)
                   'Call GRAREG("!CARDEFAC", Z$, JJ&)
               End If
             End If
           End If
         Next JJ&
         '
         If CAN < 0.001 Then
           If VUELTA% = 0 Then
             VUELTA% = 1
             GoTo 2
           End If
         End If
         '
5        Z$ = REGBLAN$("!CARDEFAC")
         Call REPLA(Z$, MKS$(CAN), 1, 4)      ' cantidad
           COD$ = CODEXT$(CODIMAT%)
         Call REPLA(Z$, MKI$(CODIMAT%), 5, 2)
         Call REPLA(Z$, DESCRIT$(CODIMAT%), 7, 64)
         Call REPLA(Z$, Mid$(X$, 5, 18), 387, 16) ' numero de remito
           '
           MONEMA% = Asc(Mid$(X$, 121, 1))
           If CAITCAR% < 1 Then
              MONEMI% = Asc(Mid$(X$, 121, 1))
              If MONEMI% < 1 Then MONEMI% = 1
              If MONEMI% = 2 Then MONEMIS = "Dólares" Else MONEMIS = "Pesos"
           End If
           COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
           '
           PRECIO = CVD(Mid$(X$, 113, 8)) * COEFCON
           PRELI = PRELISTA#(ARCHIPRE$, CODIMAT%)
           If PRELI < PRECIO Then PRELI = PRECIO
           PORDESCUEN = 0: If PRELI >= 0.0001 Then PORDESCUEN = 100 * (PRELI - PRECIO) / PRELI
           '
         Call REPLA(Z$, MKS$(PRELI), 421, 4)
         Call REPLA(Z$, MKS$(PORDESCUEN), 425, 4)
         Call REPLA(Z$, MKS$(PRECIO), 429, 4)
           IMPO# = CDbl(PRECIO) * CAN
         Call REPLA(Z$, MKD$(IMPO#), 433, 8)
         Call REPLA(Z$, MKS$(PORCEIVA(CODIMAT%)), 441, 4)
         Call REPLA(Z$, MKS$(0), 505, 4)
         Call REPLA(Z$, MKS$(0), 509, 4)
           NPEDID& = CVS(Mid$(X$, 109, 4))
         Call REPLA(Z$, TRIM$(Str$(NPEDID&)), 371, 16)
           NUOCOX$ = ""
           Y$ = FILTERGETRECORD$("PEDENCA", 1, MKS$(NPEDID&))
           NUOCOX$ = TRIM$(Mid$(Y$, 57, 16))
'           For UK& = ULTREG&("PEDENCA") To 1 Step -1
'             Y$ = REGLEIDO$("PEDENCA", UK&)
'             If CVS(Left$(Y$, 4)) = NPEDID& Then
'               NUOCOX$ = TRIM$(Mid$(Y$, 57, 16))
'               GoTo 10
'             End If
'           Next UK&
10       Call REPLA(Z$, NUOCOX$, 355, 16)
         '
         Call REGAPP("!CARDEFAC", Z$)
         CAITCAR% = CAITCAR% + 1
         '
       End If
12   Next UU&
     '
     JJ& = 0
     For U& = 1 To ULTREG&("!CARDEFAC")
       Z$ = REGLEIDO$("!CARDEFAC", U&)
       If CVS(Mid$(Z$, 1, 4)) >= 0.01 Then
         JJ& = JJ& + 1
         Call GRAREG("!CARDEFAC", Z$, JJ&)
       End If
     Next U&
     Call SETULTREG("!CARDEFAC", JJ&)
     Call CIERRARCH("!CARDEFAC")
     '
     If JJ& > 0 Then
       If COMALTER%("Hay Material en Consignación\Pendiente para este Cliente\\Facturar Consignaciones\Continuar Normalmente") = 1 Then
         CONSICLI% = 1
         Exit Sub
       End If
     End If
     '
   End If
   Call SETULTREG("!CARDEFAC", 0)
   Call CIERRARCH("!CARDEFAC")
   '
End Sub
'
Sub CARDICAMBIO()
   '
   Dim DOCABX$(2048), SALPES#(2048), SALDOL#(2048), REDOCA&(2048)
   '
   NC% = Val(Text1)
   Call BLANARCH("!CARDEFAC")
   '
   Screen.MousePointer = 11
   CADOCPE% = 0
   RFG$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
   For KUL& = 1 To Len(RFG$) Step 4
     RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
     XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
     DOCABY$ = Mid$(XXC$, 13, 18)
     SALPESY# = CVD(Mid$(XXC$, 65, 8)) - CVD(Mid$(XXC$, 73, 8))
     SALDOLY# = CVD(Mid$(XXC$, 97, 8)) - CVD(Mid$(XXC$, 105, 8))
     '
     For KKI% = 1 To CADOCPE%
       If DOCABX$(KKI%) = DOCABY$ Then GoTo 20
     Next KKI%
     CADOCPE% = CADOCPE% + 1
     DOCABX$(CADOCPE%) = DOCABY$
     REDOCA&(CADOCPE%) = RECABIN&
     KKI% = CADOCPE%
     '
20   SALPES#(KKI%) = SALPES#(KKI%) + SALPESY#
     SALDOL#(KKI%) = SALDOL#(KKI%) + SALDOLY#
     '
   Next KUL&
   '
   Call BLANARCH("SELDOCAP")
   For KKI% = 1 To CADOCPE%
     If TIFADOC% = 6 Then    ' N.CREDITO
       If SALDOL#(KKI%) <= 0.005 Then
         If SALPES#(KKI%) > 0 Then
           XXC$ = REGLEIDO$("CARBIMON", REDOCA&(KKI%))
           FEMI = CVI(Mid$(XXC$, 1, 2))
           Z$ = Space$(60) + MKS$(0)
           Call REPLA(Z$, DOCABX$(KKI%), 1, 18)
           Call REPLA(Z$, FECHATEX$(FEMI), 19, 8)
           Call REPLA(Z$, "$" + FORMATNUM$(SALPES#(KKI%), "F11.2"), 29, 12)
           Call REPLA(Z$, "u$s" + FORMATNUM$(SALDOL#(KKI%), "F11.2"), 44, 14)
           Call REPLA(Z$, MKS$(REDOCA&(KKI%)), 61, 4)
           Call REGAPP("SELDOCAP", Z$)
         End If
       End If
     ElseIf TIFADOC% = 3 Then  ' N.DEBITO
       If SALDOL#(KKI%) > 0 Then
         If SALPES#(KKI%) <= 0.005 Then
           XXC$ = REGLEIDO$("CARBIMON", REDOCA&(KKI%))
           FEMI = CVI(Mid$(XXC$, 1, 2))
           Z$ = Space$(60) + MKS$(0)
           Call REPLA(Z$, DOCABX$(KKI%), 1, 18)
           Call REPLA(Z$, FECHATEX$(FEMI), 19, 8)
           Call REPLA(Z$, "$" + FORMATNUM$(SALPES#(KKI%), "F11.2"), 29, 12)
           Call REPLA(Z$, "u$s" + FORMATNUM$(SALDOL#(KKI%), "F11.2"), 44, 14)
           Call REPLA(Z$, MKS$(REDOCA&(KKI%)), 61, 4)
           Call REGAPP("SELDOCAP", Z$)
         End If
       End If
     End If
   Next KKI%
   Call CIERRARCH("SELDOCAP")
   '
   Screen.MousePointer = 1
   If ULTREG&("SELDOCAP") < 1 Then
     Exit Sub
   End If
   '
   Call SELECHO("SELDOCAP")
   DOCABY$ = Left$(VALACT1$("SELDOCAP"), 18)
   If TRIM$(DOCABY$) = "" Then Exit Sub
   DOCUDIF = DOCABY$
   BBXX$ = VALACT2$("SELDOCAP")
   '
   MONEMIS = "Pesos"
   IDIFE = Val(Mid$(BBXX$, 12, 11))
   If TIFADOC% = 3 Then
     IDIFE = Val(Mid$(BBXX$, 29, 11)) * Val(TICAMBIS)
   End If
   If PIVADOC% = 1 Then
     TASA1 = Val(PORIVA(0).Caption)
     IDIFE = IDIFE / (1 + TASA1 / 100)
   End If
   '
   ZZ$ = REGBLAN$("CARDEFAC")
   Call REPLA(ZZ$, MKS$(1), 1, 4)
   Call REPLA(ZZ$, "Por Diferencia de Cambio", 7, 64)
   Call REPLA(ZZ$, MKS$(IDIFE), 429, 4)
   Call REPLA(ZZ$, MKD$(CDbl(IDIFE)), 433, 8)
   Call REGAPP("!CARDEFAC", ZZ$)
   ZZ$ = REGBLAN$("CARDEFAC")
   Call REPLA(ZZ$, "s/Factura " + DOCABY$, 7, 64)
   Call REGAPP("!CARDEFAC", ZZ$)
   Call CIERRARCH("!CARDEFAC")
   '
End Sub
'
Private Sub Command7_Click()
    Command7.Enabled = False
    Call CIERRARCH("*.*")
    MODX% = ALTERNA%("Listas de Precios\Condiciones de Pago\Cotización Dolar\Números Originales")
    If MODX% = 1 Then
         ACONEC$ = "LISPRE02": If EXISTE%(ACONEC$ + ".exe") < 1 Then ACONEC$ = "ADMIPRE"
         Call CONECRUN(ACONEC$, "Listas de Precios de Venta")
       ElseIf MODX% = 2 Then
         Call CARCONPAGOS
       ElseIf MODX% = 3 Then
         Call CARTAMONED
       ElseIf MODX% = 4 Then
         Call FINAL("*CODORCL")  ' Tabla de Conversion de Numeros Originales
    End If
    Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    Call CIERRARCH("*.*")
    Call CONECRUN("AMASTO01", "Maestro de Artículos de Stock")
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   Call SELECHO("SELFECHA")
   Text9.TEXT = VALACT1$("SELFECHA")
   FF% = FECHANUM(Text9.TEXT)
End Sub

Private Sub Eco21_GotFocus(Index As Integer)
   On Error Resume Next
   TEXT21(Index).SetFocus
   On Error GoTo 0
End Sub

Private Sub Form_Load()
    '
    FVENCIM = ""
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    If InStr(EPAC$, "VDR") > 0 Then
       Command12.Enabled = True
    End If
    '
    If InStr(UCase$(EPAC$), "DOSIVA") > 0 Then
      If TRIM$(CONTROL$("SVD451S", "CUEVE-11")) = "" Then
        Call GRACONTROL("SVD451S", "CUEVE-8", "111600")
        Call GRACONTROL("SVD451S", "CUEVE-9", "415340")
        Call GRACONTROL("SVD451S", "CUEVE-11", "400110")
      End If
    End If
    '
    abc$ = REGBLAN$("LEYFAC")
    '
    HOII% = HOY(HOS$)
    Text9.TEXT = HOS$: DATEDOC% = HOII%
    '
    MONEMI% = 1  ' PESOS
    CONVERPRECIO% = 0
    'If InStr(EPAC$, "DOSIVA") > 0 Or InStr(EPAC$, "FERVI-AIR") > 0 Then
    If InStr(EPAC$, "DOSIVA") > 0 Then
       MONEMIS.Visible = True
       CONVERPRECIO% = 1
       TICAMBIS.Visible = True
       Label15.Visible = True
       Command13.Visible = True
    End If
    TICAMBIS = FORMATNUM$(TICAMONE(2), "F9.4")
    '
    Call SETULTREG("!CARDETPE", 0)
    '
    Label7.Caption = "Factura:": TIFADOC% = 2
    Label6.Caption = "(Mat.Codif)": MODODOC% = 0
    Label13.Caption = "A": PIVADOC% = 1: VARIDOC% = 1
    Label12.Caption = "0004"
    Show
    Refresh
    '
    Printer.TrackDefault = False
    '
    Call CARCHCONSI    ' recarga archivo material en consignacion
    '
    VERLICRE% = 0
    If CONTROL$("FACTURA", "VERLICRE") = "SI" Then VERLICRE% = 1
    Call LEETASIVA
    Call NUFACTURA
    '
    MODESCU% = 0
    If CONTROL$("FACTUREM", "MODESCUE") = "PORCIENTO" Then
       MODESCU% = 1
    End If
    '
End Sub
'

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call BAJALDISCO
   End If
End Sub


Private Sub Label16_DblClick()
   Static CTXX%
   '
   If ULTREG&("DAPIBCLI") < 1 Then
     Call MENSERR(24, "Imposible desde esta Aplicación")
     Exit Sub
   End If
   '
   NCLIE% = Val(TRIM$(Text1.TEXT))
   If NCLIE% > 0 Then
     If ECOARCH$("DEUDOR1", MKI$(NCLIE%)) <> "" Then
       '
       If CTXX% < 1 Then
         Call GRAREG("TACATIBC", Chr$(0) + "Exento Percepcion IB", 1)
         Call GRAREG("TACATIBC", Chr$(1) + "Sujeto Percepcion IB", 2)
         Call SETULTREG("TACATIBC", 2)
         CTXX% = 1
       End If
       '
       RECORD$ = FILTERGETRECORD$("DAPIBCLI", 1, MKI$(NCLIE%))
       Call REPLA(RECORD$, MKI$(NCLIE%), 1, 2)
       CATIB1% = Asc(Mid$(RECORD$, 3, 1))
       If CATIB1% < 0 Or CATIB1% > 1 Then
         Call REPLA(RECORD$, Chr$(1), 3, 1)
         Call REPLA(RECORD$, Space$(32), 9, 32)
       End If
       '
10     OBX$ = OBJPLA$("CARDAPERIB", RECORD$)
       If OBX$ <> "" Then
         RECORD$ = OBX$
         CATIB1% = Asc(Mid$(RECORD$, 3, 1))
         If CATIB1% < 0 Or CATIB1% > 1 Then
           Call MENSERR(24, "Mal Categoría I.Brutos")
           GoTo 10
         End If
         If CATIB1% = 0 Then
           If TRIM$(Mid$(RECORD$, 9, 32)) = "" Then
             Call MENSERR(24, "Mandatorio Indicar Motivo de Exención")
             GoTo 10
           End If
         End If
         If CATIB1% = 1 Then
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
         Call REPLA(RECORD$, MKI$(Int(HORANUM(Time$))), 63, 2)
         Call FILTERPUTRECORD("DAPIBCLI", 1, MKI$(NCLIE%), RECORD$)
         Call CIERRARCH("DAPIBCLI")
         '
         abc1 = ALIPERBRU#(0, MINIPER#, VARIDOC%)    ' PARA RESETEAR
         Call CALPERIB
         '
       End If
     End If
   End If

End Sub

Private Sub Label6_Click()
   If ULTREG&("!CARDEFAC") > 0 Then Exit Sub
   If TIFADOC% = 2 Or TIFADOC% = 6 Then
      If MODODOC% = 0 Then
          Label6.Caption = "(T.Libre)": MODODOC% = 1
        ElseIf MODODOC% = 1 Then
          Label6.Caption = "(Mat.Codif)": MODODOC% = 0
      End If
   End If
   '
   NCLIE% = Val(Text1.TEXT)
   If NCLIE% > 0 Then
     If MODODOC% > 0 Then
         PORDESCU.TEXT = ""
       Else
         PORDESCU.TEXT = DESCUCLI$(NCLIE%)
     End If
   End If
End Sub


Private Sub Label7_Click()
   If ULTREG&("!CARDEFAC") > 0 Then Exit Sub
   If Label7.Caption = "Factura:" Then
       Label7.Caption = "N.Débito:": TIFADOC% = 3
       Label6.Caption = "(T.Libre)": MODODOC% = 1
     ElseIf Label7.Caption = "N.Débito:" Then
       Label7.Caption = "N.Crédito:": TIFADOC% = 6
       Label6.Caption = "(Mat.Codif)": MODODOC% = 0
     Else
       Label7.Caption = "Factura:": TIFADOC% = 2
       Label6.Caption = "(Mat.Codif)": MODODOC% = 0
   End If
   Call NUFACTURA
   '
   NCLIE% = Val(Text1.TEXT)
   If NCLIE% > 0 Then
     If MODODOC% > 0 Then
         PORDESCU.TEXT = ""
       Else
         PORDESCU.TEXT = DESCUCLI$(NCLIE%)
     End If
   End If
End Sub

Private Sub Label8_Click()
   On Error Resume Next
   Text13.SetFocus
   On Error GoTo 0
End Sub

Private Sub List1_DblClick()
   Call CARDETFAC
End Sub

Private Sub List2_DblClick()
   PREFACTU$ = CONTROL("FACTURA", "PREFACTU")
   If PREFACTU$ = "ORDESPA" Then
      Call MENSERR(24, "Operacion no Permitida")
      Exit Sub
   End If
   '
   DPEDI$ = List2.TEXT
   Call CARPEPEN(DPEDI$)
   '
   If SELITFAC.List2.ListCount < 1 Then
      Exit Sub
   End If
   '
   Call CARITSEL
   NROPEDIDO$ = ""
   '
   If ULTREG&("!CARDEFAC") > 0 Then
     Call CALTOFAC
     If Val(Total(0).TEXT) < 0.005 Then
       Call COMUNI("ATENCION: Los Pedidos Seleccionados\Parecen Ser SIN CARGO !!!")
     End If
   End If
   '
   MARSELPENDE.Visible = False
   Call CARDETFAC
   '
End Sub



Private Sub MONEMIS_Change()
   MONEMA% = MONEMI%
   If Left$(MONEMIS, 1) = "D" Then
        MONEMI% = 2
        MONEMIS.ForeColor = RGB(150, 0, 0)
        MONEMIS.ToolTipText = "Doble-Click para 'Pesos'"
        'Label12.Visible = True
        If TICAMBIS.Visible = True Then TICAMBIS.Enabled = True
        TICAMBIS = FORMATNUM$(TICAMONE(2), "F9.4")
        TICAMBX# = Val(TICAMBIS)
      Else
        MONEMI% = 1
        MONEMIS.ForeColor = RGB(0, 100, 0)
        MONEMIS.ToolTipText = "Doble-Click para 'Dólares'"
        'Label12.Visible = False
        If TICAMBIS.Visible = True Then TICAMBIS.Enabled = False
        TICAMBX# = 1
   End If
   '
   COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
   FIN& = ULTREG&("!CARDEFAC")
   Text15 = FORMATNUM$(COEFCON * Val(Text15), "F12.2")
   For U& = 1 To FIN&
     X$ = REGLEIDO$("!CARDEFAC", U&)
     CAN = CVS(Mid$(X$, 1, 4))
     PRELI = COEFCON * CVS(Mid$(X$, 421, 4))
     PRECIO = COEFCON * CVS(Mid$(X$, 429, 4))
     IMPO# = CDbl(CAN * PRECIO)
     '
     Call REPLA(X$, MKS$(PRELI), 421, 4)
     Call REPLA(X$, MKS$(PRECIO), 429, 4)
     Call REPLA(X$, MKD$(IMPO#), 433, 8)
     Call GRAREG("!CARDEFAC", X$, U&)
   Next U&
   '
   Call CARLISFAC
   Call CIERRARCH("!CARDEFAC")
   Call CALTOFAC
   '
End Sub

Private Sub MONEMIS_DblClick()
   If TRIM$(UCase$(MONEMIS)) = "PESOS" Then
       MONEMIS = "Dólares"
     Else
       MONEMIS = "Pesos"
   End If
End Sub

Private Sub PORDESCU_Change()
   Call CALTOFAC
End Sub

Private Sub PORDESCU_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
End Sub

Private Sub PORIVA_Click(Index As Integer)
   PORDESCU.SetFocus
End Sub

Private Sub Text1_Change()
   Call BLANFORMU
End Sub

Private Sub Text1_DblClick()
  Call SELECHO("DEUDOR1")
  NCLIE% = Val(TRIM$(VALACT1$("DEUDOR1")))
  If NCLIE% > 0 Then
    Text1.TEXT = TRIM$(Str$(NCLIE%))
    If Val(Text1.TEXT) > 0 Then
      If Val(Text1.TEXT) <= 32767 Then
        NC% = Val(Text1)
        If REGICLI&(NC%) > 0 Then
          Call CARDACLI
        End If
      End If
    End If
  End If
End Sub

Private Sub Text1_GotFocus()
   Text1.TEXT = TRIM$(Text1.TEXT)
   Text1.SelStart = 0
   Text1.SelLength = Len(Text1.TEXT)
   MARCONOTA.Visible = False
   MARSELPENDE.Visible = False
End Sub

Private Sub Text1_KEYPRESS(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If VALICLIE% > 0 Then
      Call BLANFORMU
      If Val(Text1.TEXT) > 0 Then
        If Val(Text1.TEXT) <= 32767 Then
          NC% = Val(Text1)
          If REGICLI&(NC%) > 0 Then
            If MARCONOTA.Visible = False Then
               Call CARDACLI
            End If
            On Error Resume Next
            Command6.SetFocus
            On Error GoTo 0
          End If
        End If
      End If
    End If
  End If
End Sub

Private Sub Text10_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   Text10.TEXT = TRIM$(Text10.TEXT)
   Text10.SelStart = 0
   Text10.SelLength = Len(Text10.TEXT)
End Sub

Private Sub Text11_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   If Val(Text11.TEXT) > 9 Then
      On Error Resume Next
      Text10.SetFocus
      On Error GoTo 0
   End If
End Sub

Private Sub Text11_LostFocus()
   NUMEPRE$ = ""
   For KK% = 1 To Len(Text11.TEXT)
     CCX% = Asc(Mid$(Text11.TEXT, KK%, 1))
     If CCX% >= 48 Then
       If CCX% <= 57 Then
         NUMEPRE$ = NUMEPRE$ + Chr$(CCX%)
       End If
     End If
   Next KK%
   '
   Call NUFACTURA
   'NPP& = NUFACTURA&
   'If Val(NUMEPRE$) < NPP& Then
   '  Call MENSERR(24, "Número no Válido")
   '  Text11.TEXT = TRIM$(Str$(NPP&))
   '  Exit Sub
   'End If
   '
   If NUMEPRE$ <> TRIM$(Text11.TEXT) Then
     Text11.TEXT = NUMEPRE$
     Call MENSERR(24, "Número de Factura Ajustado")
   End If
   '
End Sub

Private Sub Text12_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text13_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
   End If
End Sub

Private Sub Text15_Change()
   Call CALTOFA2
End Sub

Private Sub Text15_GotFocus()
   If VALICLIE% < 1 Then GoTo 99
   If TIFADOC% < 3 Then GoTo 98
   'If Val(Total(2)) < 0.5 Then GoTo 98
   Exit Sub
   '
98 On Error Resume Next
   PORDESCU.SetFocus
   On Error GoTo 0
   Exit Sub
   '
99 On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Text15 = FORMATNUM$(Val(Text15), "F12.2")
  End If
End Sub

Private Sub Text15_LostFocus()
   Text15 = FORMATNUM$(Val(Text15), "F12.2")
End Sub

Private Sub Text16_Change()
   Call CALTOFAC
End Sub

Private Sub Text2_gotfocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub TEXT21_GotFocus(Index As Integer)
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      Exit Sub
   End If
   TEXT21(Index).TEXT = TRIM$(TEXT21(Index).TEXT)
   TEXT21(Index).SelStart = 0
   TEXT21(Index).SelLength = Len(TEXT21(Index).TEXT)
End Sub

Private Sub Text3_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text4_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text5_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text21_Change(Index As Integer)
    '
    If Val(TEXT21(Index).TEXT) < 0 Or Val(TEXT21(Index).TEXT) > 255 Then
       TEXT21(Index).TEXT = ""
    End If
    '
    If Index = 0 Then
        If Val(TEXT21(Index).TEXT) < 1 Then
            Eco21(Index).TEXT = "CONCERTADOS FIJOS"
            ARCHIPRE$ = "MASTER"
          Else
            Eco21(Index).TEXT = ECOARCH$("LISTAS", MKI$(Val(TEXT21(Index).TEXT)))
            ARCHIPRE$ = TRIM$(TEXT21(0).TEXT)
            While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
            ARCHIPRE$ = "LIPRE" + ARCHIPRE$
        End If
        If Val(TEXT21(Index)) > 0 Then
          If MONELISTA%(Val(TEXT21(Index))) = 2 Then
               MONEMIS = "Dólares"
             Else
               MONEMIS = "Pesos"
          End If
        End If
      ElseIf Index = 1 Then
        Eco21(Index).TEXT = ECOARCH$("CONPAG", MKI$(Val(TEXT21(Index).TEXT)))
      ElseIf Index = 2 Then
        Eco21(Index).TEXT = ECOARCH$("VENDEDOR", Chr$(Val(TEXT21(Index).TEXT)))
    End If
    '
End Sub

Private Sub Command21_Click(Index As Integer)
    If VALICLIE% < 1 Then
       On Error Resume Next
       Text1.SetFocus
       On Error GoTo 0
       Exit Sub
    End If
    '
    If Index = 0 Then
        ARCHEC$ = "LISTAS": LOCA% = 2
      ElseIf Index = 1 Then
        ARCHEC$ = "CONPAG": LOCA% = 2
      ElseIf Index = 2 Then
        ARCHEC$ = "VENDEDOR": LOCA% = 1
    End If
    ARGU$ = CBIN$(TRIM$(TEXT21(Index).TEXT), 1, LOCA%)
    Call SELECHO(ARCHEC$)
    TEXT21(Index).TEXT = VALACT1$(ARCHEC$)
    '
End Sub
   '
Function VALICLIE%()
   VCC% = 0
   If Val(Text1.TEXT) > 0 Then
     If Val(Text1.TEXT) <= 32767 Then
       NCLIE% = Val(Text1.TEXT)
       If ECOARCH$("DEUDOR1", MKI$(NCLIE%)) <> "" Then
         VCC% = 1
       End If
     End If
   End If
   '
99 VALICLIE% = VCC%
   '
End Function

Sub BLANFORMU()
   Text2.TEXT = ""
   Text3.TEXT = ""
   Text4.TEXT = ""
   Text5.TEXT = ""
   Text7.TEXT = ""
   Text8.TEXT = ""
   Text10.TEXT = ""
   Text12.TEXT = ""
   Text14.TEXT = ""
   Text6.TEXT = ""
   FVENCIM = ""
   List1.Clear
   DOCUDIF = ""
   Label18 = ""
   '
   For KKI% = 0 To 2
      TEXT21(KKI%).TEXT = ""
      Eco21(KKI%) = ""
      BASIVA(KKI%) = ""
   Next KKI%
   '
   PORDESCU.TEXT = ""
   For KK% = 0 To 7
      Total(KK%).TEXT = ""
   Next KK%
   Text15 = ""
   Text16 = ""
   '
   MARCONOTA.Visible = False
   HOII% = HOY(HOS$)
   Text9.TEXT = HOS$
   Call SETULTREG("!CARDEFAC", 0)
   'Call NUFACTURA
   '
   On Error Resume Next
   Unload FINFACTURA
   On Error GoTo 0
   '
End Sub
'
Sub CARDACLI()
    '
    Screen.MousePointer = 11
    '
    NC% = Val(Text1.TEXT): Text1.Refresh
    If NC% > 0 Then
      If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
        Text2.TEXT = RASOCLI$(NC%): Text2.Refresh
        Text3.TEXT = DOMICLI$(NC%): Text3.Refresh
        Text4.TEXT = CPOSCLI$(NC%): Text4.Refresh
        Text5.TEXT = LOCACLI$(NC%): Text5.Refresh
        Text7.TEXT = CSTRING$(MKS$(LIMICRE(NC%)), 3, 10, 2, 2)
          Text7.Refresh
        Text12.TEXT = TRIM$(POSIVAC$(NC%)) + " - CUIT " + CUITCLI$(NC%)
          Text12.Refresh
        CREDISPO# = LIMICRE(NC%) - SALDCLI#(NC%) - CHEPEACRE#(NC%)
        Text8.ForeColor = RGB(0, 0, 0)
        If CREDISPO# < 0 Then Text8.ForeColor = RGB(255, 0, 0)
        Text8.TEXT = CSTRING$(MKD$(Abs(CREDISPO#)), 3, 10, 2, 2)
          Text8.Refresh
        '
        TEXT21(0).TEXT = Str$(LIPRCLI%(NC%))
        TEXT21(1).TEXT = Str$(CPAGCLI%(NC%))
        TEXT21(2).TEXT = Str$(VENDCLI%(NC%))
        '
        If Val(TEXT21(2).TEXT) < 1 Then
           Eco21(2).TEXT = "GERENCIA"
        End If
        '
        NCLIE% = Val(Text1.TEXT)
        If NCLIE% > 0 Then
          If MODODOC% > 0 Then
              PORDESCU.TEXT = ""
            Else
              PORDESCU.TEXT = DESCUCLI$(NCLIE%)
          End If
        End If
        PIVADOC% = PIVACLI%(NC%)
        Call CARLETDOC: Call NUFACTURA
        Refresh
        '
        LENTREG$ = Space$(96)
        RFEX$ = RECFILT$("LENTRECL", 1, MKI$(NC%))
        If Len(RFEX$) >= 4 Then
          RELENT& = CVS(Left$(RFEX$, 4))
          If RELENT& > 0 Then
            If RELENT& <= ULTREG&("LENTRECL") Then
              LENTREG$ = Mid$(REGLEIDO$("LENTRECL", RELENT&), 3)
            End If
          End If
        End If
        LENTREGA(0) = TRIM$(Left$(LENTREG$, 32))
        If TRIM$(LENTREGA(0)) = "" Then
           LENTREGA(0) = TRIM$(Text3.TEXT)
        End If
        LENTREGA(1) = TRIM$(Mid$(LENTREG$, 33, 32))
        If TRIM$(LENTREGA(1)) = "" Then
           TRPPXX$ = TRIM$(Text4.TEXT)
           If TRPPXX$ = "0" Then TRPPXX$ = ""
           If TRPPXX$ <> "" Then TRPPXX$ = TRPPXX$ + " - "
           LENTREGA(1) = TRPPXX$ + TRIM$(Text5.TEXT)
        End If
        LENTREGA(2) = TRIM$(Mid$(LENTREG$, 65, 32))
        '
        RNC$ = RECFILT$("NOTACLI", 1, MKI$(NC%))
        URN& = ULTREG&("NOTACLI")
        Text6.TEXT = ""
        '
        For U& = 1 To Len(RNC$) Step 4
          RENOTA& = CVS(Mid$(RNC$, U&, 4))
          If RENOTA& > 0 Then
            If RENOTA& <= URN& Then
              TTYY$ = RTrim$(Mid$(REGLEIDO$("NOTACLI", RENOTA&), 3))
              If Text6.TEXT <> "" Then
                Text6.TEXT = Text6.TEXT + Chr$(13) + Chr$(10)
              End If
              Text6.TEXT = Text6.TEXT + RTrim$(TTYY$)
            End If
          End If
        Next U&
        MARCONOTA.Visible = True
        MARCONOTA.Refresh
        '
        If VERLICRE% = 1 Then
          'LMCR# = LIMICRE(NC%)
          Call CIERRARCH("DEUDOR2")
          RCLI& = REGICLI&(NCLIE%)
          If RCLI& > 0 Then
            DD2$ = REGLEIDO$("DEUDOR2", RCLI&)
            LMCR# = CVS(Mid$(DD2$, 121, 4)) ' LIMICRE(NCLIE%)
          End If
          SDEU# = SALDCLI#(NC%)
          CHPA# = CHEPEACRE#(NC%)
          If SDEU# + CHPA# >= LMCR# Then
            Call COMUNI("Imposible Facturar a este Cliente.\  \Límite de Crédito Asignado ya está Superado.\  \Informe al Supervisor de Ventas.")
            Call BLANFORMU
            On Error Resume Next
            Text1.SetFocus
            On Error GoTo 0
          End If
        End If
        '
      End If
    End If
    Screen.MousePointer = 1
    '
End Sub
    
Sub CARLETDOC()
   '
   Label13.Caption = "A"
   Label12.Caption = "0004"
   VARIDOC% = 1
   RNIDOC% = 0
   '
   If PIVADOC% = 0 Or PIVADOC% = 3 Or PIVADOC% = 4 Or PIVADOC% = 6 Then
        Label13.Caption = "B"
        VARIDOC% = 2
        '
      ElseIf PIVADOC% = 2 Then
        RNIDOC% = 1
        '
      ElseIf PIVADOC% = 5 Or PIVADOC% = 7 Then
        Label13.Caption = "X"
        Label12.Caption = PREFEXPO$
        VARIDOC% = 3
        If TRIM$(LETREXPO$) <> "" Then
          Label13.Caption = LETREXPO$
        End If
        If LETREXPO$ = "A" Then
          VARIDOC% = 1
          Label13.Caption = "A"
        End If
    End If
    '
End Sub
   '
Sub NUFACTURA()
   '
   Static FLINUMERA%
   Static NUMECORR%
   '
   Screen.MousePointer = 11
   If FLINUMERA% = 0 Then
      FLINUMERA% = FECHANUM("01/01/99")
      If ULTREG&("FORMFAC") > 0 Then
         FFS$ = REGLEIDO$("FORMFAC", 2&)
         FLINUMERA% = CVI(Mid$(FFS$, 19, 2))
      End If
   End If
   '
   FEX = FLINUMERA%
   abc$ = FECHATEX$(FEX)
   DATEDOC% = FECHANUM(Text9.TEXT)
   If DATEDOC% < FLINUMERA% Then
      Call MENSERR(24, "Imposible Facturar en la Fecha Elegida.\  \La Fecha de Factura Debe Ser POSTERIOR\a la Fecha de Nueva Numeración.")
      Call CIERRARCH("*.*")
      Call FINAL("")
   End If
   '
   If NUMECORR% = 0 Then
      NUMECORR% = 1
      If CONTROL$("NUMEFACT", "NUMECORR") = "INDEPEN" Then
         NUMECORR% = (-1)
      End If
   End If
   '
   If MONUMERA$ = "" Then
      MONUMERA$ = "AUTO"
      If TRIM$(CONTROL$("", "NUMFACTU")) = "MANUAL" Then
         MONUMERA$ = "MANUAL"
      End If
   End If
   '
20 NUMAXI& = 0: KKK% = 0
   For MXC& = ULTREG&("FACTUR") To 1 Step -1
       F$ = REGLEIDO$("FACTUR", MXC&)
       TIX% = Asc(Mid$(F$, 13, 1)): VAX% = Asc(Mid$(F$, 14, 1))
       SUCUVEX% = CVI(Mid$(F$, 47, 2))
       '
       If TIX% > 0 And TIX% < 7 Then
           If TIX% = TIFADOC% Or NUMECORR% = 1 Then
             If VAX% = VARIDOC% Or (VARIDOC% = 3 And VAX% = 0) Then
               If CVI(Mid$(F$, 7, 2)) < FLINUMERA% And KKK% > 32 Then GoTo 99
               If CVI(Mid$(F$, 7, 2)) >= FLINUMERA% Then
                 If CVS(Mid$(F$, 3, 4)) > NUMAXI& Then
                   NUMAXI& = CVS(Mid$(F$, 3, 4))
                 End If
                 KKK% = KKK% + 1: If KKK% > 128 Then GoTo 99
               End If
             End If
           End If
       End If
       '
   Next MXC&
   '
99 NPP& = NUMAXI& + 1
   Text11.TEXT = TRIM$(Str$(NPP&))
   Text11.Refresh
   '
   If TIFADOC% = 2 Then
      If TRIM$(Text10.TEXT) = "" Then
        Call NUREMIFA(NROREMIT&)
        If NROREMIT& > Val(Text10.TEXT) Then
          Text10.TEXT = TRIM$(Str$(NROREMIT&))
        End If
      End If
      Screen.MousePointer = 1
      Exit Sub
     Else
      Text10.TEXT = ""
   End If
   '
   Screen.MousePointer = 1
   '
End Sub
'

Private Sub Text6_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text7_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text8_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text9_GotFocus()
   On Error Resume Next
   Command9.SetFocus
   On Error GoTo 0
End Sub

Private Sub TICAMBIS_Change()
   If TICAMBIS.Enabled = True Then
     VH(MONEMI%) = Val(TICAMBIS)
   End If
   TICAMBX# = 1
   If MONEMI% = 2 Then
     TICAMBX# = Val(TICAMBIS)
   End If
End Sub

Private Sub TICAMBIS_LostFocus()
   TICAMBIS = FORMATNUM(Val(TICAMBIS), "F9.4")
End Sub

Private Sub Timer1_Timer()
   '
   Call CALTOFAC
   '
End Sub

Sub PRIFACTURA()
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FORFAC-" + Label13.Caption))
    If FORIPRE$ = "" Then
       Call MENSERR(24, "Falta Definir Controlador de Impresión")
       Exit Sub
    End If
    '
    FORTELIB$ = TRIM$(UCase$(ATRIFORM$(FORIPRE$, "DES-MAT/FORMATO")))
    ATELIB% = XVALO(Mid$(FORTELIB$, 2))
    FORCODI$ = TRIM$(UCase$(ATRIFORM$(FORIPRE$, "COD-MAT/FORMATO")))
    '
    If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
       Call MENSERR(24, "Imposible Imprimir Documento.\No Existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    FECHDOC$ = Text9.TEXT
    NUMEDOC$ = TRIM$(Text11.TEXT)
    While Len(NUMEDOC$) < 8
      NUMEDOC$ = "0" + NUMEDOC$
    Wend
    NUMEDOC$ = Label12.Caption + "-" + NUMEDOC$
    '
    TIPODOC$ = TRIM$(REPLACAR$(Label7.Caption, ":", ""))
    If TIFADOC% = 3 Then TIPODOC$ = "Nota de Debito"
    If TIFADOC% = 6 Then TIPODOC$ = "Nota de Credito"
    CPAG% = Val(TEXT21(1).TEXT)
    '
    REPLAORI% = 0
    If CONTROL$("REMSTAN", "REPLAORI") = "SI" Then
       REPLAORI% = 1
    End If
    '
    If InStr(UCase$(TIPODOC$), "FACTURA") > 0 Then
      If InStr(UCase$(TIPODOC$), "CONTADO") < 1 Then
        If CPAG% < 1 Or InStr(UCase$(DECONPA$(CPAG%)), "CONTADO") > 0 Then
          TIPODOC$ = TIPODOC$ + " CONTADO"
        End If
      End If
    End If
    '
    DESTIDOC% = Val(Text1.TEXT)
    NCX12% = DESTIDOC%
    '
    CODTABU1$(1) = "COD-MAT"
    CODTABU1$(2) = "DES-MAT"
    CODTABU1$(3) = "CANTIDAD"
    'CODTABU1$(4) = "PRE-LIST"
    'CODTABU1$(5) = "PORDESCU"
    'CODTABU1$(6) = "PRE-NETO"
    'CODTABU1$(7) = "NET-ITEM"
    'CODTABU1$(8) = "FECH-REC"
    CODTABU1$(9) = "ORD-ITEM"
    CODTABU1$(10) = "UNIMED"
    CODTABU1$(11) = "REF-MAT1"
    CODTABU1$(12) = "REF-MAT2"
    CODTABU1$(13) = "REF-MAT3"
    CODTABU1$(14) = "N-PLANO"
    CODTABU1$(15) = "N-ORIGIN"
    CODTABU1$(16) = "N-OCOMPR"
    CODTABU1$(17) = "N-REMITO"
    CODTABU1$(18) = "N-BOLREC"
    CODTABU1$(19) = "FECH-REC"
    CODTABU1$(20) = "PRE-LIST"
    CODTABU1$(21) = "PORDESCU"
    CODTABU1$(22) = "PRE-NETO"
    CODTABU1$(23) = "NET-ITEM"
    CODTABU1$(24) = "ALICUIVA"
    CODTABU1$(25) = "IVA-ITEM"
    CODTABU1$(26) = "BRU-ITEM"
    
    CACOLTAB1% = 26
    '
    CAITAB1% = 0
    CAITCRU% = 0
    '
    For U& = 1 To ULTREG&("!CARDEFAC")
      XXX$ = REGLEIDO$("!CARDEFAC", U&)
      CIPRE% = CVI(Mid$(XXX$, 5, 2))
      CAPEX$ = Str$(CVS(Left$(XXX$, 4)))       ' cantidad
      '
      CODDX$ = CODEXT$(CIPRE%)
      If REPLAORI% = 1 Then
        CODDX$ = CODORIG$(CIPRE%, NCX12%)
      End If
      '
      CAITAB1% = CAITAB1% + 1
      TETABU1$(1, CAITAB1%) = CODDX$
      TETABU1$(2, CAITAB1%) = Mid$(XXX$, 7, 64)
      TETABU1$(3, CAITAB1%) = CAPEX$
      'TETABU1$(4, CAITAB1%) = PRELI$
      'TETABU1$(5, CAITAB1%) = DESCX$
      'TETABU1$(6, CAITAB1%) = PREUX$
      'TETABU1$(7, CAITAB1%) = ITOTX$
      'TETABU1$(8, CAITAB1%) = FENTX$
      TETABU1$(9, CAITAB1%) = Str$(U&)
      TETABU1$(10, CAITAB1%) = Mid$(XXX$, 71, 4)                           ' unidad de medida
      TETABU1$(11, CAITAB1%) = Mid$(XXX$, 75, 64)                ' referencia 1
      TETABU1$(12, CAITAB1%) = Mid$(XXX$, 139, 64)                ' referencia 1
      TETABU1$(13, CAITAB1%) = Mid$(XXX$, 203, 64)                ' referencia 1
      TETABU1$(14, CAITAB1%) = Mid$(XXX$, 315, 24)  ' numero de plano
      TETABU1$(15, CAITAB1%) = Mid$(XXX$, 339, 16)  ' numero original
      TETABU1$(16, CAITAB1%) = Mid$(XXX$, 355, 16)  ' numero de o.compra
      TETABU1$(17, CAITAB1%) = Mid$(XXX$, 387, 16)  ' numero de remito
      TETABU1$(18, CAITAB1%) = Mid$(XXX$, 403, 16)  ' numero de doc-recepcion
      TETABU1$(19, CAITAB1%) = FECHATEX$(CVI(Mid$(XXX$, 419, 2))) ' fecha de recepcion
      '
      FACIVITEM = 1: If Label13.Caption = "B" Then FACIVITEM = 1 + (CVS(Mid$(XXX$, 441, 4))) / 100
      LPRE% = Val(TEXT21(0).TEXT)
      If LPRE% > 0 Then
        TETABU1$(20, CAITAB1%) = Str$(FACIVITEM * CVS(Mid$(XXX$, 421, 4))) ' precio de lista
        TETABU1$(21, CAITAB1%) = Str$(CVS(Mid$(XXX$, 425, 4))) ' PORCENTAJE DE DESCUENTO
      End If
      '
      TETABU1$(22, CAITAB1%) = Str$(FACIVITEM * CVS(Mid$(XXX$, 429, 4))) ' precio unitario
      TETABU1$(23, CAITAB1%) = Str$(CVD(Mid$(XXX$, 433, 8))) ' total neto
      TETABU1$(24, CAITAB1%) = Str$(CVS(Mid$(XXX$, 441, 4))) ' alicuota iva
      TETABU1$(25, CAITAB1%) = Str$(CVD(Mid$(XXX$, 445, 8))) ' importe iva del item
      TETABU1$(26, CAITAB1%) = Str$(CVD(Mid$(XXX$, 453, 8))) ' total bruto
      '
      If COP% = 1 Then
         TRANSPORTE# = TRANSPORTE# + Val(TETABU1$(23, CAITAB1%))
      End If
      '
      If PIVADOC% = 5 Or PIVADOC% = 7 Then ' FACTURA DE EXPORTACION
        If CIPRE% > 0 Then                 ' MATERIAL CODIFICADO
          DESCLAR$ = TRIM$(DESCRILA$(CIPRE%))
          If DESCLAR$ <> "" Then
            Call FRATEXTO(DESCLAR$, ATELIB%)
            TETABU1$(2, CAITAB1%) = RETEX$(1)
            For KKU1% = 2 To CARETEX%
              CAITAB1% = CAITAB1% + 1
              TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
            Next KKU1%
          End If
        End If
      End If
      '
      If Left$(FORCODI$, 1) = "K" Then
        CAITAB1% = CAITAB1% + 1
      End If
      '
109 Next U&
    '
    SUMATOTA# = Val(Total(0).TEXT)
    IDESCUEN# = Val(Total(1).TEXT)
' AQUI HAY QUE HACER ALGO SIMILAR PARA LA BONIFICACION
    '\\\OT-05-0524-WAR-18/07/05///
    BONIF# = Val(Total(7).TEXT)
    '\\\OT-05-0524-WAR-FIN///
    NROREMI$ = TRIM$(Text10.TEXT)
    If NROREMI$ <> "" Then
      While Len(NROREMI$) < 8: NROREMI$ = "0" + NROREMI$: Wend
      PREFIREM$ = CONTROL$("", "PREFIREM")
      If PREFIREM$ = "" Then PREFIREM$ = "0001"
      NROREMI$ = PREFIREM$ + "-" + NROREMI$
    End If
    '
    If CONSICLI% = 1 Then
      NROREMI$ = NROREMICO$
    End If
    '
    COEFIVAINCLU# = 1
    If Val(Total(2).TEXT) > 0.005 Then
      COEFIVAINCLU# = (Val(Total(2).TEXT) + Val(Total(3).TEXT) + Val(Total(4).TEXT) + Val(Total(5).TEXT)) / Val(Total(2).TEXT)
    End If
    '
    If TRIM$(Text14.TEXT) <> "" Then
      NUOCOMPRA$ = TRIM$(Text14.TEXT)
    End If
    '
    CODPARAM$(1) = "CIVA-CLI": DOCPARAM$(1) = Label13.Caption
    CODPARAM$(2) = "NRO-REMI": DOCPARAM$(2) = NROREMI$
    CODPARAM$(3) = "NRO-OCOM":  DOCPARAM$(3) = NUOCOMPRA$
    CODPARAM$(4) = "NRO-AENT":  DOCPARAM$(4) = NROPEDIDO$
    CODPARAM$(5) = "IMP-NETO": DOCPARAM$(5) = Total(0).TEXT
    CODPARAM$(6) = "VENDEDOR": DOCPARAM$(6) = Eco21(2).TEXT
    CODPARAM$(7) = "COND-PAG": DOCPARAM$(7) = Eco21(1).TEXT
    CODPARAM$(8) = "POR-IVA": DOCPARAM$(8) = PORIVA(0).Caption
    CODPARAM$(9) = "IMP-IVA": DOCPARAM$(9) = Total(3).TEXT
    CODPARAM$(10) = "POR-IVA1": DOCPARAM$(10) = PORIVA(1).Caption
    CODPARAM$(11) = "IMP-IVA1": DOCPARAM$(11) = Total(4).TEXT
    CODPARAM$(12) = "IMP-DESC": DOCPARAM$(12) = Total(1).TEXT
    CODPARAM$(13) = "IMPNEDES": DOCPARAM$(13) = Total(2).TEXT
    'CODPARAM$(4) = "NRO-OCOM": DOCPARAM$(4) = Text10.TEXT
    CODPARAM$(14) = "IMP-TOTA": DOCPARAM$(14) = Total(6).TEXT
    CODPARAM$(15) = "POR-DESC": DOCPARAM$(15) = TRIM$(PORDESCU.TEXT)
    CODPARAM$(16) = "IMP-RNI": DOCPARAM$(16) = Total(5).TEXT
    CODPARAM$(23) = "TOTA-DOL": DOCPARAM$(23) = ""
    If TICAMBIO > 0 Then
       VADOLA# = Val(Total(6).TEXT) / TICAMBIO
       DOCPARAM$(23) = TRIM$(CSTRING$(MKD$(VADOLA#), 4, 12, 2, 2))
    End If
    '
    If Label13.Caption = "B" Then
      DOCPARAM$(5) = Str$(COEFIVAINCLU# * Val(Total(0).TEXT))
      DOCPARAM$(12) = Str$(COEFIVAINCLU# * Val(Total(1).TEXT))
      DOCPARAM$(13) = Str$(COEFIVAINCLU# * Val(Total(2).TEXT))
    End If
    '
    CODFOR$ = FORIPRE$
    LOIMPLET% = Val(ATRIFORM$(CODFOR$, "LOIMPLET"))
    If LOIMPLET% < 16 Then LOIMPLET% = 32
    TTTT# = Val(Total(6).TEXT)
    TOTFAC$ = NUMTEX$(TTTT#)
    If MONECO$ <> "" Then
      PPXX% = InStr(TOTFAC$, "PESOS")
      If PPXX% > 0 Then
        TOTFAC$ = Left$(TOTFAC$, PPXX% - 1) + MONECO$ + Mid$(TOTFAC$, PPXX% + 5)
      End If
    End If
    Call FRATEX(TOTFAC$, NRX1$, NRX2$, NRX3$, LOIMPLET%)
    CODPARAM$(17) = "IMP-LET1": DOCPARAM$(17) = NRX1$
    CODPARAM$(18) = "IMP-LET2": DOCPARAM$(18) = NRX2$
    CODPARAM$(19) = "IMP-LET3": DOCPARAM$(19) = NRX3$
    '
    CODPARAM$(20) = "DOMI-TRA": DOCPARAM$(20) = LENTREGA(0).TEXT
    CODPARAM$(21) = "DESTINO": DOCPARAM$(21) = LENTREGA(1).TEXT
    CODPARAM$(22) = "TRANSPOR": DOCPARAM$(22) = LENTREGA(2).TEXT
    CODPARAM$(24) = "MONEDEMI": DOCPARAM$(24) = MONEMIS
    CODPARAM$(25) = "PERC-IVA": DOCPARAM$(25) = Text15
    CODPARAM$(26) = "FECH-VEN": DOCPARAM$(26) = TRIM$(FVENCIM)
    '\\\OT-05-0524-WAR-18/07/05///
    CODPARAM$(27) = "TIPO-VAL": DOCPARAM$(27) = TRIM$(Text16)
    'CODPARAM$(28) = "IMPO-CHE": DOCPARAM$(28) = BONIF#  '''' WALTER !!!
    CODPARAM$(28) = "IMPO-CHE": DOCPARAM$(28) = FORMATNUM$(BONIF#, "F12.2")
    CODPARAM$(29) = "REF-ASIE": DOCPARAM$(29) = TIPOVENTA$
    CODPARAM$(30) = "NRO-PASE": DOCPARAM$(30) = PUERTO$
    
    '\\\OT-05-0524-WAR-FIN///
    CAPARDOC% = 30
    '
'    For KKIL& = 0 To 5
'       Call GRAREG("!OBSERVOF", FINFACTURA.Text1(KKIL&).TEXT, KKIL& + 1)
'    Next KKIL&
'    Call SETULTREG("!OBSERVOF", 6)
'    Call CIERRARCH("!OBSERVOF")
'  \\\OT-07-0508-FL-14-12-07///
   Call SETULTREG("!OBSERVOF", 0)
    If FINFACTURA.Text4 <> "" Then
      JR& = 0
      FORMOBSE$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "OBSERVA/FORMATO")))
      ANCHOB% = XVALO(Mid$(FORMOBSE$, 2)): If ANCHOB% < 32 Then ANCHOB% = 32
      Call FRATEXTO(FINFACTURA.Text4, ANCHOB%)
      For KKII% = 1 To CARETEX%
        JR& = JR& + 1
        Call GRAREG("!OBSERVOF", RETEX$(KKII%), JR&)
      Next KKII%
      Call SETULTREG("!OBSERVOF", JR&)
      Call CIERRARCH("!OBSERVOF")
    End If
'   \\\OT-07-0508-FL-FIN///
     
    
    '
    PRINTERPORT$ = CONTROL$(IDENTER$, "PORTFAC")
    CACOPDOC% = Val(FINFACTURA.Text2(0).TEXT)
    '
    NOCONFIRM% = 0
    If TRIM$(UCase$(CONTROL$("FACTUREM", "IMPRENUM"))) <> "AUTONUME" Then
      If CONTROL$("FACTUREM", "FORMULAR") <> "HOJASUEL" Then
        NOCONFIRM% = 1
      End If
    End If
    '
    HOJASUEL% = 0
    If CONTROL$("FACTUREM", "FORMULAR") = "HOJASUEL" Then
       HOJASUEL% = 1
    End If
    '
    If CACOPDOC% > 0 Then
      Call PRINTDOC("@" + FORIPRE$)
    End If
    '
    CACOPDOC% = Val(FINFACTURA.Text2(1).TEXT)
    If CACOPDOC% > 0 Then
      If TRIM$(Text10.TEXT) <> "" Then
        Call PRIREMIFA
      End If
    End If
    '
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    '
    If NOCONFIRM% > 0 Then
       Call COMUNI("Se han Emitido los Documentos Solicitados")
    End If
    '
End Sub

Sub PRIREMIFA()
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FOREMIFA"))
    If TRIM$(FORIPRE$) <> "" Then
       '
'       For KKIL& = 0 To 2
'          Call GRAREG("!OBSERVOF", FINFACTURA.Text1(KKIL& + 6).TEXT, KKIL& + 1)
'       Next KKIL&
'       Call SETULTREG("!OBSERVOF", 3)
'       Call CIERRARCH("!OBSERVOF")
'   \\\OT-07-0508-FL-14-12-07///
       Call SETULTREG("!OBSERVOF", 0)
       If FINFACTURA.Text5 <> "" Then
         JR& = 0
         CODFOR$ = FORIPRE$
         FORMOBSE$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "OBSERVA/FORMATO")))
         ANCHOB% = XVALO(Mid$(FORMOBSE$, 2)): If ANCHOB% < 32 Then ANCHOB% = 32
         Call FRATEXTO(FINFACTURA.Text5, ANCHOB)
         For KKII% = 1 To CARETEX%
           JR& = JR& + 1
           Call GRAREG("!OBSERVOF", RETEX$(KKII%), JR&)
         Next KKII%
         Call SETULTREG("!OBSERVOF", JR&)
         Call CIERRARCH("!OBSERVOF")
       End If
'   \\\OT-07-0508-FL-FIN///
       '
       PREFIREM$ = CONTROL$("", "PREFIREM")
       If PREFIREM$ = "" Then PREFIREM$ = "0001"
       NROREMI$ = TRIM$(Text10.TEXT)
       While Len(NROREMI$) < 8: NROREMI$ = "0" + NROREMI$: Wend
       NROREMI$ = PREFIREM$ + "-" + NROREMI$
       '
       NC% = Val(Text1.TEXT)
       FECHDOC$ = Text9.TEXT
       NUMEDOC$ = NROREMI$
       TIPODOC$ = "Remito"
       '
       If TRIM$(Text14.TEXT) <> "" Then
         NUOCOMPRA$ = TRIM$(Text14.TEXT)
       End If
       '
       CODPARAM$(1) = "IMP-VARS": DOCPARAM$(1) = TRIM$(Mid$(TRXP$, 1, 32))
       CODPARAM$(2) = "NUBULTOS": DOCPARAM$(2) = TRIM$(FINFACTURA.Text2(2).TEXT)
       CODPARAM$(3) = "CUIT-TRA": DOCPARAM$(3) = TRIM$(Mid$(TRXP$, 65, 16))
       CODPARAM$(4) = "VALO-DEC": DOCPARAM$(4) = TRIM$(FINFACTURA.Text2(3).TEXT)
       CODPARAM$(5) = "COND-TRA": DOCPARAM$(5) = TRIM$(Mid$(TRXP$, 97, 32))
       CODPARAM$(6) = "DOCUCOND": DOCPARAM$(6) = TRIM$(Mid$(TRXP$, 129, 24))
       CODPARAM$(7) = "DEPOSAL ": 'DOCPARAM$(7) = TRIM$(DEPOSAL$)
       CODPARAM$(8) = "NRO-REMI": DOCPARAM$(8) = TRIM$(NROREMI$)
       CODPARAM$(9) = "NRO-OCOM":  DOCPARAM$(9) = NUOCOMPRA$
       CODPARAM$(10) = "PATE-TRA":  DOCPARAM$(10) = TRIM$(Mid$(TRXP$, 81, 16))
       CODPARAM$(11) = "DOMI-TRA":  DOCPARAM$(11) = LENTREGA(0).TEXT
       CODPARAM$(12) = "LOCA-TRA":  DOCPARAM$(12) = TRIM$(Mid$(TRXP$, 159, 32))
       CODPARAM$(13) = "COND-PAG": DOCPARAM$(13) = Eco21(1).TEXT
       CODPARAM$(14) = "NRO-AENT":  DOCPARAM$(14) = NROPEDIDO$
       CODPARAM$(15) = "DESTINO": DOCPARAM$(15) = LENTREGA(1).TEXT
       CODPARAM$(16) = "TRANSPOR": DOCPARAM$(16) = LENTREGA(2).TEXT
       CODPARAM$(17) = "VENDEDOR": DOCPARAM$(17) = Eco21(2).TEXT
       '\\\OT-05-0523-WAR-18/7/05///
       CODPARAM$(23) = "PES-CONJ": DOCPARAM$(23) = TRIM$(FINFACTURA.Label8)
       ''\\\OT-05-0523-WAR-FIN///
       CAPARDOC% = 23 'CAPARDOC% = 17'\\\OT-05-0523-WAR-18/7/05///
       '
       ABUXX$ = TRIM$(UCase$(LENTREGA(2)))
       For UKU& = 1 To ULTREG&("TRANSPOR")
         AUXX$ = REGLEIDO$("TRANSPOR", UKU&)
         If UCase$(TRIM$(Mid$(AUXX$, 3, 32))) = ABUXX$ Then
           CAPARDOC% = 23 'CAPARDOC% = 22'\\\OT-05-0523-WAR-18/7/05///
           CODPARAM$(18) = "PIE-COMP": DOCPARAM$(18) = TRIM$(Mid$(ABUXX$, 3, 32))
           CODPARAM$(19) = "DIB-COMP": DOCPARAM$(19) = TRIM$(Mid$(ABUXX$, 35, 48))
           CODPARAM$(20) = "MAQ-COMP": DOCPARAM$(20) = TRIM$(Mid$(ABUXX$, 83, 48))
           CODPARAM$(21) = "LAR-COMP": DOCPARAM$(21) = TRIM$(Mid$(ABUXX$, 131, 48))
           CODPARAM$(22) = "ACH-COMP": DOCPARAM$(22) = TRIM$(Mid$(ABUXX$, 179, 15))
           GoTo 20
         End If
       Next UKU&
       '
20     CODTABU1$(1) = "COD-MAT"
       CODTABU1$(2) = "DES-MAT"
       CODTABU1$(3) = "CANTIDAD"
       CODTABU1$(4) = "ORD-ITEM"
       CACOLTAB1% = 4
       '
       DESTIDOC% = NC%
       CAITAB1% = 0
       '
       REPLAORI% = 0
       If CONTROL$("REMSTAN", "REPLAORI") = "SI" Then
         REPLAORI% = 1
       End If
       '
       For U& = 1 To ULTREG&("!CARDEFAC")
          XXX$ = REGLEIDO$("!CARDEFAC", U&)
          CANS$ = Str$(CVS(Left$(XXX$, 4)))       ' cantidad
          If Val(CANS$) >= 0.05 Then
             If TRIM$(Mid$(XXX$, 387, 16)) = "" Then
                CAITAB1% = CAITAB1% + 1
                CIX11% = CVI(Mid$(XXX$, 5, 2))   ' codigo
                CDXXS$ = CODEXT$(CIX11%)
                If REPLAORI% = 1 Then
                  CDXXS$ = CODORIG$(CIX11%, NC%)
                End If
                '
                DDTXT$ = Mid$(XXX$, 7, 64)                  ' descripcion
                '
                TETABU1$(1, CAITAB1%) = CDXXS$
                TETABU1$(2, CAITAB1%) = DDTXT$
                TETABU1$(3, CAITAB1%) = CANS$
                TETABU1$(4, CAITAB1%) = Str$(CAITAB1%)
                '
             End If
          End If
       Next U&
       '
       CACOPDOC% = Val(FINFACTURA.Text2(1).TEXT)
       PRINTERPORT$ = CONTROL$(IDENTER$, "PORTREM")
       Call PRINTDOC("FOREMIFA")   ' Remito de Facturacion
       '
    End If
    '
End Sub

Private Sub Total_Change(Index As Integer)
   If VALICLIE% < 1 Then Exit Sub
   If TIFADOC% < 3 Then GoTo 10
   '
   If Index >= 3 Then
     If Index <= 4 Then
       Call CALTOFA2
     End If
   End If
   '
10 If Index = 6 Then
     If Val(Total(6)) > 0.005 Then
       If VERLICRE% = 1 Then
         If TIFADOC% < 4 Then ' EXCLUYE CREDITOS
           NCLIE% = Val(Text1)
           'LMCR# = LIMICRE(NCLIE%)
           Call CIERRARCH("DEUDOR2")
           RCLI& = REGICLI&(NCLIE%)
           If RCLI& > 0 Then
             DD2$ = REGLEIDO$("DEUDOR2", RCLI&)
             LMCR# = CVS(Mid$(DD2$, 121, 4)) ' LIMICRE(NCLIE%)
           End If
           SDEU# = SALDCLI#(NCLIE%)
           CHPA# = CHEPEACRE#(NCLIE%)
           TPED# = Val(Total(6).TEXT) * TICAMBX#
           If SDEU# + CHPA# + TPED# > LMCR# Then
             TLIMI$ = TRIM$(CSTRING$(MKD$(LMCR# - SDEU# - CHPA#), 3, 12, 2, 2))
             If TLIMI$ = "" Or Val(TLIMI$) < 0 Then TLIMI$ = "0.00"
             Call COMUNI("Imposible Emitir Este Documento.\  \Supera el Límite de Crédito Disponible\para esta Cuenta de $ " + TLIMI$ + ".\  \Informe al Supervisor de Ventas.")
             Call CIERRARCH("*.*")
             Exit Sub
           End If
         End If
       End If
     End If
   End If
   '
20 If Index = 2 Or (VARIDOC% = 2 And Index < 6) Then
     Call CALPERIB
   End If
   '
End Sub

Sub CALPERIB()
   '
   NCLIE% = XVALO(Text1)
   TASAX = ALIPERBRU#(NCLIE%, MINIPER#, VARIDOC%)
   '
   Text15 = ""
   BIMPON# = XVALO(Total(2))
   If BIMPON# * TICAMONE(MONEMI%) >= MINIPER# Then
      PERSALTA# = BIMPON# * (TASAX / 100)
      Text15 = FORMATNUM$(PERSALTA#, "F12.2")
   End If
   
   Call CALTOFA2
   '
End Sub
'
Private Sub Total_GotFocus(Index As Integer)
   '
   If VALICLIE% < 1 Then GoTo 99
   If Index < 3 Or Index > 4 Then GoTo 98
   If TIFADOC% < 3 Then GoTo 98
   Exit Sub
   '
98 On Error Resume Next
   PORDESCU.SetFocus
   On Error GoTo 0
   Exit Sub
   '
99 On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Sub GRAFACTURA(OKX%)
   '
   OKX% = 0
   NCLIE% = Val(Text1.TEXT)
   If MARCONOTA.Visible = True Then Exit Sub
   If MARSELPENDE.Visible = True Then Exit Sub
   If ULTREG&("!CARDEFAC") < 1 Then Exit Sub
   '
   Call APROCOPIAS(OKX%)
   COPIASFAC% = Val(FINFACTURA.Text2(0))
   COPIASREM% = Val(FINFACTURA.Text2(1))
   If OKX% < 0 Then Exit Sub
   If COPIASFAC% + COPIASREM% < 1 Then Exit Sub
   '
   If COPIASFAC% > 0 Then
     If Val(Total(6).TEXT) < 0.01 Then
       Call MENSERR(24, "Imposible Emitir Factura SIN CARGO !!!")
       Exit Sub
     End If
   End If
   '
   Call CIERRARCH("SVD202")
   Call CIERRARCH("DEUDOR2")
   Call CIERRARCH("FACTUR")
   Call CIERRARCH("MOVISTO")
   Call CIERRARCH("SALSTOCK")
   '
   If COPIASFAC% < 1 Then
      Call GRAREMITO(OKX%)
      Screen.MousePointer = 1
      Exit Sub
   End If
   '
   If TIFADOC% = 3 Or TIFADOC% = 6 Then
     '
     If TRIM$(DOCUDIF) <> "" Then
       OPIMPUTA.Option2.Value = True
     End If
     '
     OPIMPUTA.Show 1
     VDEF$ = "ND": If TIFADOC% = 6 Then VDEF$ = "NC"
     VDEF$ = VDEF$ + String$(94, 0)
     NRO& = Val(Text11)
     Call REPLA(VDEF$, MKS$(NRO&), 3, 4)
     Call REPLA(VDEF$, MKI$(NCLIE%), 7, 2)
     FFI% = FECHANUM(Text9)
     Call REPLA(VDEF$, MKI$(FFI%), 9, 2)
     Call REPLA(VDEF$, MKD$(Val(Total(2)) * TICAMBX#), 19, 8)
     '
     CUEVE$ = CONTROL$("SVD451S", "CUEVE-8")
     If OPIMPUTA.Option2.Value = True Then CUEVE$ = CONTROL$("SVD451S", "CUEVE-9")
     If OPIMPUTA.Option3.Value = True Then CUEVE$ = CONTROL$("SVD451S", "CUEVE-11")
     QQQ% = CUEINT%(CUEVE$)
     Call REPLA(VDEF$, MKI$(QQQ%), 17, 2)
     XD1$ = VDEF$
     '
   End If
   '
   Call BLOQARCH("SVD202")
   Call BLOQARCH("DEUDOR2")
   Call BLOQARCH("FACTUR")
   Call BLOQARCH("MOVISTO")
   Call BLOQARCH("SALSTOCK")
   '
   NUPANTALLA& = Val(Text11.TEXT)
   Call NUFACTURA
   If Val(Text11.TEXT) > NUPANTALLA& Then
      Call CIERRARCH("SVD202")
      Call CIERRARCH("DEUDOR2")
      Call CIERRARCH("FACTUR")
      Call CIERRARCH("MOVISTO")
      Call CIERRARCH("SALSTOCK")
      Call COMUNI("FLEXOFT ha Re-Numerado el Documento\como Número '" + TRIM$(Text11.TEXT) + "'.\  \Verifique y Vuelva a Grabar.")
      Exit Sub
   End If
   Text11.TEXT = TRIM$(Str$(NUPANTALLA&))
   '
   COPIASREM% = Val(FINFACTURA.Text2(1))
   If COPIASREM% < 1 Then
      Text10.TEXT = ""
   End If
   '
   Screen.MousePointer = 11
   If VALICLIE% < 1 Then
      MERRO$ = "Número de Cliente no Válido"
      GoTo 98
   End If
   NCLIE% = Val(Text1.TEXT)
   '
   If Val(TEXT21(0).TEXT) > 255 Then
      MERRO$ = "Lista de Precios no Válida"
      GoTo 98
   End If
   '
   If Val(TEXT21(1).TEXT) > 255 Then
      MERRO$ = "Condición de Pago no Válida"
      GoTo 98
   End If
   '
   If Val(TEXT21(2).TEXT) > 255 Then
      MERRO$ = "Vendedor no Válido"
      GoTo 98
   End If
   '
   If Val(Total(6).TEXT) * TICAMBX# < 0.01 Then
      MERRO$ = "Importe no Válido"
      GoTo 98
   End If
   '
   HOII% = HOY(HOS$): FEACT = HOII%
   FEMI% = FECHANUM(Text9.TEXT)
   If FEMI% < DIANTE(FEACT) Or FEMI% > DIPOST(FEACT) Then
      MERRO$ = "Fecha de Emisión no Válida"
      GoTo 98
   End If
   '
   If VERLICRE% = 1 Then
     If TIFADOC% < 4 Then ' EXCLUYE CREDITOS
       '
       For UKK& = 1 To ULTREG&("!CARDEFAC")
         XKK$ = REGLEIDO$("!CARDEFAC", UKK&)
         NROPEDXX& = Val(TRIM$(Mid$(XKK$, 371, 16)))
         If NROPEDXX& < 1 Then GoTo 22
         RFFKK$ = RECFILT$("AUTORENT", 1, MKS$(NROPEDXX&))
         If Len(RFFKK$) < 4 Then GoTo 22 ' NO AUTORIZADO
         REGAUTKK& = CVS(RFFKK$)
         XAUT$ = REGLEIDO$("AUTORENT", REGAUTKK&)
         If UCase$(Mid$(XAUT$, 31, 2)) <> "SI" Then GoTo 22
       Next UKK&
       GoTo 25
       '
'22     LMCR# = LIMICRE(NCLIE%)
22     Call CIERRARCH("DEUDOR2")
       RCLI& = REGICLI&(NCLIE%)
       If RCLI& > 0 Then
         DD2$ = REGLEIDO$("DEUDOR2", RCLI&)
         LMCR# = CVS(Mid$(DD2$, 121, 4)) ' LIMICRE(NCLIE%)
       End If
       SDEU# = SALDCLI#(NCLIE%)
       CHPA# = CHEPEACRE#(NCLIE%)
       TPED# = Val(Total(6).TEXT) * TICAMBX#
       If SDEU# + CHPA# + TPED# > LMCR# Then
         TLIMI$ = TRIM$(CSTRING$(MKD$(LMCR#), 3, 12, 2, 2))
         If TLIMI$ = "" Then TLIMI$ = "0.00"
         Call COMUNI("Imposible Emitir Este Documento.\  \Supera el Límite de Crédito Asignado\para esta Cuenta de $ " + TLIMI$ + ".\  \Informe al Supervisor de Ventas.")
         Exit Sub
       End If
     End If
   End If
   '
25 DESC$ = TRIM$(PORDESCU.TEXT)
   PODES = COEFDESCU(DESC$)
   '
   For KK% = 1 To 10: TOTAFAC#(1, KK%) = 0: Next KK%
   TOTAFAC#(1, 1) = Val(BASIVA(0).Caption) / (1 - PODES / 100)
   TOTAFAC#(1, 2) = Val(BASIVA(1).Caption) / (1 - PODES / 100)
   TOTAFAC#(1, 3) = Val(Total(0).TEXT)
   TOTAFAC#(1, 4) = Val(Total(1).TEXT)
   TOTAFAC#(1, 5) = Val(BASIVA(0).Caption)
   TOTAFAC#(1, 6) = Val(Total(3).TEXT)
   TOTAFAC#(1, 7) = Val(Total(5).TEXT)
   TOTAFAC#(1, 8) = Val(Total(6).TEXT)
   TOTAFAC#(1, 9) = Val(BASIVA(1).Caption)
   TOTAFAC#(1, 10) = Val(Total(4).TEXT)
   '
   If Val(Text15) >= 0.005 Then
     TOTAFAC#(1, 9) = 0
     TOTAFAC#(1, 10) = Val(Text15)
   End If
   '
   TICA = 1
   For KK% = 1 To 10
     TOTAFAC#(1, KK%) = TOTAFAC#(1, KK%) * TICAMBX#
     TOTAFAC#(2, KK%) = TOTAFAC#(1, KK%) / TICA
   Next KK%
   '
   LU$ = LUDAT$
   ARCHIDEU$ = "SVD202"
   ARCHICLI$ = "DEUDOR"
   ARCHIFAC$ = "FACTUR"
   RD% = 1
   TIPO% = TIFADOC%
   VA% = VARIDOC%
   NRO = Val(Text11.TEXT)
   NC% = NCLIE%
   FFI% = FECHANUM(Text9.TEXT)
   CPAG% = Val(TEXT21(1).TEXT)
   TOTAGEN# = Val(Total(6).TEXT) * TICAMBX#
   '
   'Call BLIMESS("Calculando Vencimientos ...")
   If TIFADOC% <> 6 Or TRIM$(DOCUDIF) = "" Then
     Call GRADEU(LU$, ARCHIDEU$, RD%, TIPO%, VA%, NRO, NC%, FFI%, CPAG%, TOTAGEN#)
     Call CIERRARCH("SVD202")
   End If
   '
   'Call BLIMESS("Actualizando Saldo de Cuenta ...")
   SIGX% = 1: If TIPO% = 6 Then SIGX% = (-1)
   TTF# = TOTAGEN# * SIGX%
   Call GRACLI(LU$, ARCHICLI$, NC%, TTF#, DATCLI$(), 2)
   Call CIERRARCH("DEUDOR2")
   '
   FLIM% = FECHANUM("01/01/99")
   If ULTREG&("FORMFAC") > 0 Then
      FFS$ = REGLEIDO$("FORMFAC", 2&)
      FLIM% = CVI(Mid$(FFS$, 19, 2))
   End If
   '
   'Call BLIMESS("Grabando Factura ...")
   If CAFOR% < 1 Then CAFOR% = 1
   If CONTROL$("NUMEFACT", "NUMECORR") = "INDEPEN" Then CAFOR% = 3
   Call GRAFAC98(LU$, ARCHIFAC$, 3, RD%, FLIM%, TIPO%, VA%, CAFOR%, NRO, NC%, CLI$, IVAS$, FFI%, FVI%, LPRE%, VENDEX%, TOTAFAC#())
   Call GRAFAC98(LU$, ARCHIFAC$, 4, RD%, FLIM%, TIPO%, VA%, CAFOR%, NRO, NC%, CLI$, IVAS$, FFI%, FVI%, LPRE%, VENDEX%, TOTAFAC#())
   '
   'Call BLIMESS("Movimiento de Stock ...")
   NRO1 = Val(Text11.TEXT)
   If TIFADOC% = 6 Then NRO1 = (-1) * Val(Text11.TEXT)
   '
   REMI = Val(Text10.TEXT)
   REMI1 = REMI: If TIFADOC% = 3 Then REMI1 = (-1) * NRO1
   If CONSICLI% < 1 Then
       Call GRAM2000(NC%, NRO1, REMI1, FFI%)
     Else
       Call GRACONSI
   End If
   '
   If TIFADOC% = 3 Or TIFADOC% = 6 Then
     NRO& = Val(Text11)
     Call REPLA(XD1$, MKS$(NRO&), 3, 4)
     MMXX$ = "ENEFEBMARABRMAYJUNJULAGOSETOCTNOVDIC"
     FEX = FFI%: FEFA$ = FECHATEX$(FEX)
     MESI% = Val(Mid$(FEFA$, 4, 2))
     ARCHICO$ = "ICV" + Mid$(MMXX$, 3 * MESI% - 2, 3) + Mid$(FEFA$, 7)
     Call COPYSTRU("ICOVEN", ARCHICO$)
     Call REGAPP(ARCHICO$, XD1$)
     '
     If TRIM$(DOCUDIF) <> "" Then
       Call APLIDICAM
     End If
     '
   End If
   '
   If MONEMI% > 1 Then
     If TIFADOC% < 3 Then   ' graba noviniento cabeza
       Call GRAMOCABE
     End If
   End If
   '
   OKX% = 1
   Screen.MousePointer = 1
   Exit Sub

98 Call CIERRARCH("PEDENCA")
   Call CIERRARCH("PEDDETA")
   Screen.MousePointer = 1
   Call MENSERR(24, MERRO$)
   '
End Sub
'
Sub GRAMOCABE()
   '
   FEMI% = FECHANUM(Text9)
   NCLIE% = Val(Text1)
   TIPO% = TIFADOC%
   VA% = VARIDOC%
   NRO& = Val(Text11.TEXT)
   NUDOCU$ = FORMDOC$(NRO&, TIPO%, VA%, SUCUVEN%)
   '
   SUCUVEN% = 0
   On Error Resume Next
   SUCUVEN% = Val(Label12)
   On Error GoTo 0
   NUDOCU$ = FORMDOC$(NRO&, TIPO%, VA%, SUCUVEN%)
   '
   TOTAPESOS# = TOTAFAC#(1, 8)
   TOTADOLAR# = TOTAPESOS# / TICAMBX#
   '
   RAPLI$ = REGBLAN$("CARBIMON")
   ' documento cabeza
   Call REPLA(RAPLI$, MKI$(FEMI%), 1, 2)
   Call REPLA(RAPLI$, MKI$(NCLIE%), 3, 2)
   Call REPLA(RAPLI$, Chr$(TIPO%), 5, 1)
   Call REPLA(RAPLI$, MKI$(VA%), 6, 1)
   Call REPLA(RAPLI$, MKS$(NRO), 7, 4)
   Call REPLA(RAPLI$, MKI$(SUCUVEN%), 11, 2)
   Call REPLA(RAPLI$, NUDOCU$, 13, 18)
   Call REPLA(RAPLI$, MKI$(MONEMI%), 31, 2)
   '
   ' documento aplicado
   Call REPLA(RAPLI$, MKI$(FEMI%), 33, 2)
   Call REPLA(RAPLI$, MKI$(NCLIE%), 35, 2)
   Call REPLA(RAPLI$, Chr$(TIPO%), 37, 1)
   Call REPLA(RAPLI$, MKI$(VA%), 38, 1)
   Call REPLA(RAPLI$, MKS$(NRO), 39, 4)
   Call REPLA(RAPLI$, MKI$(SUCUVEN%), 43, 2)
   Call REPLA(RAPLI$, NUDOCU$, 45, 18)
   Call REPLA(RAPLI$, MKI$(MONEMI%), 63, 2)
   '
   'importe en pesos
   Call REPLA(RAPLI$, MKD$(TOTAPESOS#), 65, 8)
   Call REPLA(RAPLI$, MKD$(TICAMBX#), 89, 8)
   '
   'importe en dolares
   Call REPLA(RAPLI$, MKD$(TOTADOLAR#), 97, 8)
   '
   Call REGAPP("CARBIMON", RAPLI$)
   Call CIERRARCH("CARBIMON")
   '
End Sub
'
Sub APLIDICAM()
   '
   Screen.MousePointer = 11
   NCI% = Val(Text1)
   TIDOCAPLI% = 2
   NUDOCAPLI& = Val(Mid$(DOCUDIF, 9, 8))
   CANCE# = Val(Total(6)) * TICAMONE(MONEMI%)
   TICAMBX# = Val(TICAMBIS): If TICAMBX# < 1 Then TICAMBX# = 1
   '
   If TIFADOC% = 6 Then
     FIN& = ULTREG&("SVD202")
     ABUSQUE$ = ""
     For U& = 1 To FIN&
       SS$ = REGLEIDO$("SVD202", U&)
       If CVI(Left$(SS$, 2)) = NCI% Then
         If Asc(Mid$(SS$, 3, 1)) = TIDOCAPLI% Then
           If CVS(Mid$(SS$, 5, 4)) = NUDOCAPLI& Then
             ABUSQUE$ = Mid$(SS$, 3, 6)
             NUSUCUVE% = CVI(Mid$(SS$, 29, 2))
             DOCORI$ = FORMDOC$(CVS(Mid$(SS$, 5, 4)), Asc(Mid$(SS$, 3, 1)), Asc(Mid$(SS$, 4, 1)), NUSUCUVE%)
             ITOTORI# = CVD(Mid$(SS$, 21, 8))
             IVAORI# = (Int(ALIVA * ITOTORI# / (1 + ALIVA / 100) + 0.5)) / 100
             '
             SALANT# = CVD(Mid$(SS$, 11, 8))
             SALDEU# = CVD(Mid$(SS$, 11, 8)) - CANCE#
             If SALDEU# < 0 Then SALDEU# = 0
             Call REPLA(SS$, MKD$(SALDEU#), 11, 8)
             Call GRAREG("SVD202", SS$, U&)
             '
             HOII% = HOY(HO$)
             RefApli$ = "N/C Diferencia de Cambio " + HO$
             FEDOAPLI% = FECHANUM(Text9)
             DOCAPLI$ = FORMDOC$(Val(Text11), TIFADOC%, VARIDOC%, Val(Label12))
             '
             Z$ = REGBLAN$("APLICOB")
             Call REPLA(Z$, MKI$(FEDOAPLI%), 1, 2)
             Call REPLA(Z$, MKI$(NCI%), 3, 2)
             Call REPLA(Z$, DOCAPLI$, 5, 18)
             Call REPLA(Z$, MKD$(CANCE#), 23, 8)
             Call REPLA(Z$, RefApli$, 31, 24)
             Call REPLA(Z$, DOCUDIF, 55, 18)
             Call REPLA(Z$, MKD$(ITOTORI#), 73, 8)
             Call REPLA(Z$, MKD$(IVAORI#), 81, 8)
             '
             Call REGAPP("APLICOB", Z$)
             GoTo 20
             '
           End If
         End If
       End If
     Next U&
     '
20   Call CIERRARCH("APLICOB")
     Call CIERRARCH("SVD202")
     '
     RFF$ = RECFILT$("CARBIMON", 2, MKI$(NCI%))
     If Len(RFF$) > 0 Then
       For JKLX& = 1 To Len(RFF$) Step 4
         RECABI& = CVS(Mid$(RFF$, JKLX&, 4))
         YCB$ = REGLEIDO$("CARBIMON", RECABI&)
         If Mid$(YCB$, 5, 6) = ABUSQUE$ Then
           Call REPLA(YCB$, MKI$(FEDOAPLI%), 33, 2)
           Call REPLA(YCB$, MKI$(NCI%), 35, 2)
           Call REPLA(YCB$, Chr$(TIFADOC%), 37, 1)
           Call REPLA(YCB$, Chr$(VARIDOC%), 38, 1)
           Call REPLA(YCB$, MKS$(Val(Text11)), 39, 4)
           Call REPLA(YCB$, MKI$(Val(Label12)), 43, 2)
           Call REPLA(YCB$, DOCAPLI$, 45, 18)
           Call REPLA(YCB$, MKI$(1), 63, 2) ' RECIBO SIEMPRE EN PESOS
           Call REPLA(YCB$, MKD$(0), 65, 8)
           Call REPLA(YCB$, MKD$(CANCE#), 73, 8)
           Call REPLA(YCB$, MKD$(0), 81, 8)
           Call REPLA(YCB$, MKD$(0), 89, 8)  ' NO PONE TIPO DE CAMBIO
           Call REPLA(YCB$, MKD$(0), 97, 8)
           Call REPLA(YCB$, MKD$(0), 105, 8)
           Call REPLA(YCB$, MKD$(0), 113, 8)
           '
           Call REGAPP("CARBIMON", YCB$)
           GoTo 29
         End If
       Next JKLX&
     End If
     '
29   Call CIERRARCH("CARBIMON")
     '
   ElseIf TIFADOC% = 3 Then ' nota de debito
     '
     RFF$ = RECFILT$("CARBIMON", 2, MKI$(NCI%))
     ABUSQUE$ = TRIM$(DOCUDIF)
     If Len(RFF$) > 0 Then
       For JKLX& = 1 To Len(RFF$) Step 4
         RECABI& = CVS(Mid$(RFF$, JKLX&, 4))
         YCB$ = REGLEIDO$("CARBIMON", RECABI&)
         If Mid$(YCB$, 13, 18) = ABUSQUE$ Then
             FEDOAPLI% = FECHANUM(Text9)
           Call REPLA(YCB$, MKI$(FEDOAPLI%), 33, 2)
           Call REPLA(YCB$, MKI$(NCI%), 35, 2)
           Call REPLA(YCB$, Chr$(TIFADOC%), 37, 1)
           Call REPLA(YCB$, Chr$(VARIDOC%), 38, 1)
           Call REPLA(YCB$, MKS$(Val(Text11)), 39, 4)
           Call REPLA(YCB$, MKI$(Val(Label12)), 43, 2)
             DOCAPLI$ = FORMDOC$(Val(Text11), TIFADOC%, VARIDOC%, Val(Label12))
           Call REPLA(YCB$, DOCAPLI$, 45, 18)
           Call REPLA(YCB$, MKI$(1), 63, 2) ' RECIBO SIEMPRE EN PESOS
           Call REPLA(YCB$, MKD$(CANCE#), 65, 8)
           Call REPLA(YCB$, MKD$(0), 73, 8)
           Call REPLA(YCB$, MKD$(0), 81, 8)
           Call REPLA(YCB$, MKD$(Val(TICAMBIS)), 89, 8)  ' NO PONE TIPO DE CAMBIO
           Call REPLA(YCB$, MKD$(0), 97, 8)
           Call REPLA(YCB$, MKD$(CANCE# / TICAMBX#), 105, 8)
           Call REPLA(YCB$, MKD$(0), 113, 8)
           '
           Call REGAPP("CARBIMON", YCB$)
           GoTo 39
         End If
       Next JKLX&
     End If
     '
39   Call CIERRARCH("CARBIMON")
     '
   End If
   DOCUDIF = ""
   '
End Sub
'
Sub GRAREMITO(OKX%)
   '
   Screen.MousePointer = 11
   If VALICLIE% < 1 Then
      MERRO$ = "Número de Cliente no Válido"
      GoTo 98
   End If
   NCLIE% = Val(Text1.TEXT)
   '
   Call BLOQARCH("MOVISTO")
   Call BLOQARCH("CONSICLI")
   '
   AINDI% = 0
   CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
   If Len(CMO$) <> 2 Then
     CMO$ = "RT"
   End If
   '
   REMI$ = "RT." + TRIM$(Text10.TEXT)
   NROO& = Val(Mid$(REMI$, 4))
   FF% = FECHANUM(Text9.TEXT)
   DEPOSALI% = DEPOEXPE%
   '
   PREFIREM$ = CONTROL$("", "PREFIREM")
   If PREFIREM$ = "" Then PREFIREM$ = "0001"
   NROREMI$ = TRIM$(Mid$(REMI$, 4))
   While Len(NROREMI$) < 8: NROREMI$ = "0" + NROREMI$: Wend
   NROREMI$ = CMO$ + "." + PREFIREM$ + "-" + NROREMI$
   '
   HO% = HOY(HOS$)
   HP = 100 * Val(Left$(Time$, 2)) + Val(Mid$(Time$, 4, 2))
   User% = USNBR% Mod 100
   CTRXX$ = MKI$(HO%) + MKI$(Int(HP)) + Chr$(User%) + String$(2, 0)
   '
   RECONSI$ = REGBLAN$("CONSICLI")
   Call REPLA(RECONSI$, String$(12, 0), 109, 12)
   Call REPLA(RECONSI$, MKI$(NCLIE%), 1, 2)
   Call REPLA(RECONSI$, MKI$(FF%), 3, 2)
   Call REPLA(RECONSI$, NROREMI$, 5, 18)
   Call REPLA(RECONSI$, String$(12, 0), 109, 12)
   Call REPLA(RECONSI$, CTRXX$, 122, 7)
   '
   REFE$ = RASOCLI$(NCLIE%)
   '
   YXI% = 0
   FIN& = ULTREG&("!CARDEFAC")
   For I& = 1 To FIN&
     Call TRACE(20, I&, FIN&)
     X$ = REGLEIDO$("!CARDEFAC", I&)
     CANTAA = CVS(Left$(X$, 4))
     If CANTAA > 0 Then
       CI% = CVI(Mid$(X$, 5, 2))
       If CI% > 0 Then
         NPEDID& = Val(Mid$(X$, 371, 8))
         COD$ = CODEXT$(CI%)
         PRE# = CDbl(CVS(Mid$(X$, 429, 4)))
         CMOV$ = "RT"
         If Mid$(X$, 461, 1) = "R" Then CMOV$ = "*R" ' REPARACION
         DOPRI$ = REMI$
         DOSEC$ = ""
         LOTE$ = ""
         SIG% = (-1)
         '
         Call MOVISTO(FF%, CI%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, PRE#, "$", NCLIE%, DEPOSALI%, SIG% * CANTAA)
         EPAC$ = TRIM$(UCase$(EMPREAC$))
         If InStr(EPAC$, "DOSIVA") > 0 Then
           If Left$(DOPRI$, 2) = "RT" Then
             DOPRIQ$ = "PL." + TRIM$(Mid$(X$, 371, 16))
             If TRIM$(DOPRIQ$) = "PL." Then   ' REMITO DIRECTO SIN O/VENTA
               NRXLX& = 900000 + Val(Text10)
               DOPRIQ$ = DOPRIQ$ + TRIM$(Str$(NRXLX&))
             End If
             REFEQ$ = "Packing-List " + REFE$
             If CMOV$ = "*R" Then
                  CMOVQ$ = "*D"
               Else
                  CMOVQ$ = "RP"
             End If
             Call MOVISTO(FF%, CI%, LOTE$, CMOVQ$, DOPRIQ$, DOSEC$, REFEQ$, 0, "$", 0, DEPOSALI%, (-1) * SIG% * CANTAA)
           End If
         End If
         '
         YXI% = YXI% + 1
         ZZ$ = RECONSI$
         Call REPLA(ZZ$, TRIM$(REMI$) + "-" + TRIM$(Str$(YXI%)), 23, 12)
         Call REPLA(ZZ$, MKI$(CI%), 35, 2)
         Call REPLA(ZZ$, MKD$(CDbl(Abs(CANTAA))), 37, 8)
         Call REPLA(ZZ$, MKS$(NPEDID&), 109, 4)
         Call REPLA(ZZ$, MKD$(PRE#), 113, 8)
         Call REPLA(ZZ$, Chr$(MONEMI%), 121, 1)
         Call REGAPP("CONSICLI", ZZ$)
         '
       End If
     End If
   Next I&
   '
   Call CIERRARCH("MOVISTO")
   Call CIERRARCH("CONSICLI")
   Screen.MousePointer = 1
   OKX% = 1
   Exit Sub
   '
   '
98 Call CIERRARCH("MOVISTO")
   Call CIERRARCH("CONSICLI")
   Screen.MousePointer = 1
   Call MENSERR(24, MERRO$)
   '
End Sub
'
Sub APROCOPIAS(OKX%)
   '
   TIPCOM$ = TRIM$(REPLACAR$(Label7.Caption, ":", ""))
   CPAG% = Val(TEXT21(1).TEXT)
   CATIVAS$ = Label13.Caption
   '
   TCMP$ = TRIM$(TIPCOM$)
   If InStr(UCase$(TIPCOM$), "FACTURA") > 0 Then
     If InStr(UCase$(TIPCOM$), "CONTADO") < 1 Then
       If CPAG% < 1 Or InStr(UCase$(DECONPA$(CPAG%)), "CONTADO") > 0 Then
         TCMP$ = TCMP$ + " CONTADO"
       End If
     End If
   End If
   '
   TICOMPROBAN$ = TCMP$
   FINFACTURA.Label2.Caption = TCMP$
   FINFACTURA.Label1(0) = TCMP$ + ":"
   FINFACTURA.Label4.Caption = CATIVAS$
   '
   CODFOR$ = TRIM$(CONTROL$("", "FORFAC-" + Label13.Caption))
   NROCON& = Val(Text11.TEXT)
   NCN$ = TRIM$(Str$(NROCON&))
   While Len(NCN$) < 8: NCN$ = "0" + NCN$: Wend
   PREFI$ = Label12.Caption
   NCN$ = PREFI$ + "-" + NCN$
   NUCOMPROBAN$ = NCN$
   FINFACTURA.Label3.Caption = NCN$
   '
5  COPIASFAC% = 2: COPIASREM% = 2
   If InStr(TCMP$, "CONTADO") > 0 Then
      COPIASFAC% = 3: COPIASREM% = 0
   End If
   If TIFADOC% > 2 Then
      COPIASFAC% = 2: COPIASREM% = 0
   End If
   '
   If InStr(UCase$(EMPREAC$), "FUINY") > 0 Then
      COPIASFAC% = 3: COPIASREM% = 4
      If TIFADOC% > 2 Then COPIASREM% = 0
   End If
   '
   If InStr(EMPREAC$, "NEUMANN") > 0 Then
      SUCUVEN% = 0
      On Error Resume Next
      SUCUVEN% = Val(Label12.Caption)
      If SUCUVEN% > 1 Then
         COPIASREM% = 0
      End If
      If CONSICLI% > 0 Then
         COPIASREM% = 0
      End If
   End If
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "DOSIVAC") > 0 Then
      COPIASFAC% = 2: COPIASREM% = 3
      If TIFADOC% > 2 Then
         COPIASREM% = 0
      End If
   End If
   '
   If CONSICLI% = 1 Then
      COPIASREM% = 0
   End If
   '
   FINFACTURA.Text2(0) = TRIM$(Str$(COPIASFAC%))
   FINFACTURA.Text2(1) = TRIM$(Str$(COPIASREM%))
   FINFACTURA.Text2(2) = " "
   FINFACTURA.Text2(3) = TRIM$(Str$(Int(TICAMBX# * Val(FORFACLI.Total(2).TEXT)))) + ".="
   FINFACTURA.Label5.Caption = ""
   For U& = 0 To 8
      FINFACTURA.Text1(U&).TEXT = ""
   Next U&
   '
   FLEYE% = VARIDOC% - 1
   RFF$ = RECFILT$("LEYFAC", 1, MKI$(FLEYE%))
   JJ& = (-1)
   For U& = 1 To Len(RFF$) Step 4
     RELE& = CVS(Mid$(RFF$, U&, 4))
     LEYY$ = REGLEIDO$("LEYFAC", RELE&)
     JJ& = JJ& + 1
     If JJ& <= 5 Then
       FINFACTURA.Text1(JJ&) = Mid$(LEYY$, 3)
     End If
   Next U&
   '
   RFF$ = RECFILT$("LEYFAC", 1, MKI$(3))
   JJ& = 5
   For U& = 1 To Len(RFF$) Step 4
     RELE& = CVS(Mid$(RFF$, U&, 4))
     LEYY$ = REGLEIDO$("LEYFAC", RELE&)
     JJ& = JJ& + 1
     If JJ& <= 8 Then
       FINFACTURA.Text1(JJ&) = Mid$(LEYY$, 3)
     End If
   Next U&
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   TTXX$ = ""
   If InStr(EPAC$, "DOSIVA") > 0 Or InStr(EPAC$, "FERVI-AIR") > 0 Then
      If MONEMI% = 2 Then  ' clausula dolar
          BASEIMPO = TICAMBX# * (Val(Total(2)))
          TOTAIVAX = TICAMBX# * (Val(Total(3)) + Val(Total(4)) + Val(Total(5)))
          TTXX$ = "TIPO DE CAMBIO UTILIZADO $ " + TRIM$(FORFACLI.TICAMBIS)
          TTXX$ = TTXX$ + " POR CADA U$S, AL SOLO EFECTO IMPOSITIVO EL IVA DE"
          TTXX$ = TTXX$ + " ESTA FACTURA EQUIVALE A: $ "
          TTXX$ = TTXX$ + TRIM$(FORMATNUM$(TOTAIVAX, "F12.2"))
          TTXX$ = TTXX$ + ". BASE IMPONIBLE $ " + TRIM$(FORMATNUM$(BASEIMPO, "F12.2")) + "."
      End If
      Call FRATEXTO(TTXX$, 54)
      For KKII% = 1 To CARETEX%
        FINFACTURA.Text1(KKII% - 1) = RETEX$(KKII%)
      Next KKII%
   End If
   '
   FINFACTURA.Label6.Enabled = False
   FINFACTURA.CLAREDOL = "0"
   If MONEMI% = 1 Then
     FINFACTURA.Label6.Enabled = True
   End If
   '\\\OT-05-0523-WAR-18/07/05///
   PST = 0
   P1 = 0
   PT = 0
   For UKK& = 1 To ULTREG&("!CARDEFAC")
     XKK$ = REGLEIDO$("!CARDEFAC", UKK&)
     CI% = CVI(Mid$(XKK$, 5, 2))
     CANTI1 = CVS(Mid$(XKK$, 1, 4))
     P1 = PESUNI(CI%)
     PST = P1 * CANTI1
     PT = PT + PST
   Next UKK&
   FINFACTURA.Label8 = FORMATNUM$(PT, "F8.1")
   '\\\OT-05-0523-WAR-FIN///
10 FINFACTURA.Show 1
   FORFACLI.Refresh
   '
   If FINFACTURA.Label5.Caption <> "APRO" Then
      If COMALTER%("Realmente Quiere CANCELAR este Documento\\No, Continuar\Si, Cancelar") <> 2 Then GoTo 10
      OKX% = (-1)
      Exit Sub
   End If
   '
   COPIASFAC% = Val(FINFACTURA.Text2(0))
   If COPIASFAC% < 0 Or COPIASFAC% > 6 Then GoTo 5
   COPIASREM% = Val(FINFACTURA.Text2(1))
   If COPIASREM% > 6 Then GoTo 5
   If COPIASREM% < 1 Then
      Text10.TEXT = ""
   End If
   '
End Sub
'
Private Sub Total_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
     If Index = 0 Then
       On Error Resume Next
       PORDESCU.SetFocus
       On Error GoTo 0
      Else
       On Error Resume Next
       Command14.SetFocus
       On Error GoTo 0
     End If
   End If
End Sub

Private Sub Total_LostFocus(Index As Integer)
   'If Index = 0 Then
     TTT1 = Val(Total(Index).TEXT)
     Total(Index) = CSTRING$(MKS$(TTT1), 3, 12, 2, 2)
   'End If
End Sub

Sub CARDETFAC()
   '
   HOII% = HOY(HOS$)
   '
   ATRI$ = "/NC/ARCHECO(2)=" + ARCHIPRE$
   ATRI$ = ATRI$ + "/TITUPAN=Detalle de Facturación (Mat.Codificado)"
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Frame1.Top + Picture3.Top))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture3.Left + 50))
   '
   Timer1.Enabled = True
10 PRF$ = "FACDOSI"
   If MODESCU% = 1 Then PRF$ = "FACPORDE"
   If MODODOC% = 1 Then
      PRF$ = "FAC-LIBR"
      ATRI$ = "/NC"
      ATRI$ = ATRI$ + "/TITUPAN=Detalle de Facturación (Texto Libre)"
      ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Frame1.Top + Picture3.Top))
      ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture3.Left + 50))
   End If
   List1.Clear
   '
   'If MODODOC% = 1 Then
   '  If TIFADOC% > 2 Then
   '    Call SELCREDEPEN
   '  End If
   'End If
   '
   XXX% = VENTABU%(PRF$ + ATRI$)
   Call CALTOFAC
   Timer1.Enabled = False
   '
   NPEDID& = 0
   If XXX% < 0 Then
      Call CIERRARCH("!CARDEFAC")
      Call BLANARCH("!CARDEFAC")
      Call BLANFORMU
      Total(0).TEXT = ""
      Text1.TEXT = ""
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   '
   If ULTREG&("!CARDEFAC") < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   '
   Screen.MousePointer = 11   ' Registrando FACTURA ...
   List1.Clear
   FIN& = ULTREG&("!CARDEFAC")
   For U& = 1 To FIN&
      Y$ = REGLEIDO$("!CARDEFAC", U&)
      If MODODOC% < 1 Then
        CI% = CVI(Mid$(Y$, 5, 2))
        If CI% < 1 Or CI% > NUMAS% Then
          Call MENSERR(24, "Código no Válido")
          Call REPLA(Y$, MKI$(0), 5, 2)
          Call GRAREG("!CARDEFAC", Y$, U&)
          GoTo 10
        End If
        '
        PRECIO = CVS(Mid$(Y$, 429, 4))
        PRELI = PRELISTA#(ARCHIPRE$, CI%)
        If PRELI < PRECIO Then PRELI = PRECIO
        PORDESCUEN = 0: If PRELI >= 0.0001 Then PORDESCUEN = 100 * (PRELI - PRECIO) / PRELI
        Call REPLA(Y$, MKS$(PRELI), 421, 4)
        Call REPLA(Y$, MKS$(PORDESCUEN), 425, 4)
        MAREPA$ = UCase$(Mid$(Y$, 461, 1)): If MAREPA$ <> "R" Then MAREPA$ = " "
        Call REPLA(Y$, MAREPA$, 461, 2)
        Call GRAREG("!CARDEFAC", Y$, U&)
        '
        If NPEDID& = 0 Then
           NPEDID& = Val(Mid$(Y$, 371, 8))
        End If
        '
        CAN = CVS(Mid$(Y$, 1, 4))
        If CAN > 999999 Then
            Call MENSERR(24, "Cantidad Excesiva")
            GoTo 10
          ElseIf CAN < 0.1 Then
            Call MENSERR(24, "Falta Cantidad")
            Call REPLA(Y$, MKS$(0), 1, 4)
            Call GRAREG("!CARDEFAC", X$, U&)
            GoTo 10
        End If
        '
      End If
   Next U&
   '
   Call CARLISFAC
   '
   If NPEDID& > 0 Then
      For UK& = ULTREG&("PEDENCA") To 1 Step -1
        YX$ = REGLEIDO$("PEDENCA", UK&)
        If CVS(Left$(YX$, 4)) = NPEDID& Then GoTo 1300
      Next UK&
       GoTo 99
      '
1300  TEXT21(0).TEXT = Str$(Asc(Mid$(YX$, 268, 1)))
      TEXT21(1).TEXT = Str$(Asc(Mid$(YX$, 269, 1)))
      TEXT21(2).TEXT = Str$(Asc(Mid$(YX$, 270, 1)))
      '
      If Val(TEXT21(2).TEXT) < 1 Then
        Eco21(2).TEXT = "GERENCIA"
      End If
      '
      LENTREG$ = Mid$(YX$, 73, 96)
      If TRIM$(LENTREG$) <> "" Then
        LENTREGA(0).TEXT = Mid$(LENTREG$, 1, 32)
        LENTREGA(1).TEXT = Mid$(LENTREG$, 33, 32)
        LENTREGA(2).TEXT = Mid$(LENTREG$, 65, 32)
      End If
      '
      If Mid$(YX$, 275, 14) <> String$(14, 0) Then
         PORDESCU.TEXT = TRIM$(Mid$(YX$, 275, 14))
      End If
   End If
   '
99 On Error Resume Next
   Command14.SetFocus
   On Error GoTo 0
   Screen.MousePointer = 1
   '
End Sub
'
Sub CARLISFAC()
    List1.Clear
   FIN& = ULTREG&("!CARDEFAC")
   For U& = 1 To FIN&
      Y$ = REGLEIDO$("!CARDEFAC", U&)
      TTXX$ = Space$(128)
      CI% = CVI(Mid$(Y$, 5, 2))
      PRECIO = CVS(Mid$(Y$, 429, 4))
      PRELI = PRELISTA#(ARCHIPRE$, CI%)
      PORDESCUEN = 0: If PRELI >= 0.0001 Then PORDESCUEN = 100 * (PRELI - PRECIO) / PRELI
      MAREPA$ = UCase$(Mid$(Y$, 461, 1)): If MAREPA$ <> "R" Then MAREPA$ = " "
      NPEDID& = Val(Mid$(Y$, 371, 8))
      CAN = CVS(Mid$(Y$, 1, 4))
      '
      Call REPLA(TTXX$, CSTRING$(MKS$(CAN), 3, 6, 0, 2), 1, 6)
      Call REPLA(TTXX$, CODEXT$(CI%), 8, 15)
      Call REPLA(TTXX$, Mid$(Y$, 7, 64), 24, 26)
      PRELI = CVS(Mid$(Y$, 421, 4))
      Call REPLA(TTXX$, CSTRING$(MKS$(PRELI), 3, 8, 2, 2), 50, 8)
      PORDESCUIN = CVS(Mid$(Y$, 425, 4))
      Call REPLA(TTXX$, CSTRING$(MKS$(PORDESCUIN), 3, 5, 1, 2), 59, 5)
      PREUNI = CVS(Mid$(Y$, 429, 4))
      Call REPLA(TTXX$, CSTRING$(MKS$(PREUNI), 3, 8, 2, 2), 64, 8)
      TOT# = CVD(Mid$(Y$, 433, 8))
      Call REPLA(TTXX$, CSTRING$(MKD$(TOT#), 3, 10, 2, 2), 74, 10)
      Call REPLA(TTXX$, Mid$(Y$, 461, 2), 86, 2) ' MARCA REPARACION
      List1.AddItem TTXX$
      '
   Next U&
   List1.Refresh
   '
End Sub
'
Sub CALTOFAC()
   '
   Static CLIA$, RNI%, APLIVA%
   If Text1.TEXT <> CLIA$ Then
     CLIA$ = Text1.TEXT: NCLIE% = Val(CLIA$)
     RNI% = 0: APLIVA% = 1
     If NCLIE% > 0 Then
       PPVI% = PIVACLI%(NCLIE%)
       If PPVI% = 2 Then RNI% = 1
       If PPVI% = 5 Or PPVI% = 7 Then APLIVA% = 0
     End If
   End If
   '
   TASA1 = Val(PORIVA(0).Caption)
   TASA2 = Val(PORIVA(1).Caption)
   TASA3 = Val(PORIVA(2).Caption)
   SUNEFAC# = 0: SUNEFA1# = 0: SUNEFA2# = 0
   '
   SUIVA# = 0: SUIVARNI# = 0: TOIVA# = 0: TOTOTA# = 0
   TOITEMS% = 0
   For U& = 1 To ULTREG&("!CARDEFAC")
     '
     X$ = REGLEIDO$("!CARDEFAC", U&)
     CIXXI% = CVI(Mid$(X$, 5, 2))
     CAN = CVS(Mid$(X$, 1, 4))
     PORDESCUIN = CVS(Mid$(X$, 425, 4))
     If PORDESCUIN < 0 Or PORDESCUIN > 100 Then PORDESCUIN = 0
     '
     If CAN > 0.05 Or MODODOC% > 0 Then
       TASARE = TASA1
       If CIXXI% > 0 Then
         CILOCAL% = CIXXI%
         TASARE = PORCEIVA(CILOCAL%)
       End If
       If TASARE <> CVS(Mid$(X$, 441, 4)) Then
         Call REPLA(X$, MKS$(TASARE), 441, 4)
         Call GRAREG("!CARDEFAC", X$, U&)
       End If
       '
       PRELI = CVS(Mid$(X$, 421, 4))
       PREUNI = CVS(Mid$(X$, 429, 4)) '
       '
       If MODESCU% = 1 Then
         If PRELI < 0.00005 Then PRELI = PREUNI
         PREUNI = PRELI * (100 - PORDESCUIN) / 100
         Call REPLA(X$, MKS$(PREUNI), 429, 4)
       End If
       '
       TOT# = CDbl(Int(10 * CAN + 0.5)) / 10 * CDbl(Int(10000 * PREUNI + 0.5)) / 10000
       If MODODOC% = 1 Then
          PREUNI = CVS(Mid$(X$, 429, 4))
          TOT# = CVD(Mid$(X$, 433, 8))
       End If
       Call REPLA(X$, MKD$(TOT#), 433, 8)
       SUNEFAC# = SUNEFAC# + TOT#
       '
       TASARE = CVS(Mid$(X$, 441, 4))
       IVAITEM1# = (Int(TOT# * TASARE * APLIVA% + 0.5)) / 100
       Call REPLA(X$, MKD$(IVAITEM1#), 445, 8)
       Call REPLA(X$, MKD$(TOT# + IVAITEM1#), 453, 8)
       Call GRAREG("!CARDEFAC", X$, U&)
       TOITEMS% = TOITEMS% + 1
       '
       If Abs(TASARE - TASA2) < 0.01 Then
           SUNEFA2# = SUNEFA2# + TOT#
         Else
           SUNEFA1# = SUNEFA1# + TOT#
       End If
       '
     End If
   Next U&
   '
   Label18 = TRIM$(Str$(TOITEMS%))
   '
   Total(0).TEXT = CSTRING$(MKD$(SUNEFAC#), 3, 12, 2, 2)
   Total(0).Refresh
   '
15 DESC$ = TRIM$(PORDESCU.TEXT)
   PODES = COEFDESCU(DESC$)
   TTT1# = Val(Total(0).TEXT) * PODES / 100
   If TTT1# > Val(Total(0)) Then
     Call MENSERR(24, "Descuento no Procesable")
     PORDESCU = ""
     GoTo 15
   End If
   Total(1).TEXT = CSTRING$(MKD$(TTT1#), 3, 12, 2, 2)
   '
20 DESC3$ = TRIM$(Text16.TEXT)
   PODES3 = COEFDESCU(DESC3$)
   TTT13# = (Val(Total(0).TEXT) - Val(Total(1))) * PODES3 / 100
   If TTT13# > (Val(Total(0)) - Val(Total(1))) Then
     Call MENSERR(24, "Bonificación no Procesable")
     Text16 = ""
     GoTo 20
   End If
   Total(7).TEXT = CSTRING$(MKD$(TTT13#), 3, 12, 2, 2)
   '
   TTT2# = Val(Total(0).TEXT) - Val(Total(1).TEXT) - Val(Total(7).TEXT)
   Total(2).TEXT = CSTRING$(MKD$(TTT2#), 3, 12, 2, 2)
   '
   TTTX# = SUNEFA2# * (100 - PODES) / 100
   BASIVA(1).Caption = TRIM$(CSTRING$(MKD$(TTTX#), 3, 12, 2, 2))
   Total(4).TEXT = CSTRING$(MKD$(Val(BASIVA(1).Caption) * Val(PORIVA(1).Caption) * APLIVA% / 100), 3, 12, 2, 2)
   '
   TTTX# = (Val(Total(2).TEXT) - Val(BASIVA(1).Caption))
   BASIVA(0).Caption = TRIM$(CSTRING$(MKD$(TTTX#), 3, 12, 2, 2))
   Total(3).TEXT = CSTRING$(MKD$(Val(BASIVA(0).Caption) * Val(PORIVA(0).Caption) * APLIVA% / 100), 3, 12, 2, 2)
   '
   Call CALTOFA2
   '
End Sub
'
Sub CALTOFA2()
   '
   Static CLIA$, RNI%
   If Text1.TEXT <> CLIA$ Then
     CLIA$ = Text1.TEXT: NCLIE% = Val(CLIA$)
     RNI% = 0
     If NCLIE% > 0 Then
       PPVI% = PIVACLI%(NCLIE%)
       If PPVI% = 2 Then RNI% = 1
     End If
   End If
   '
   TTTX# = (Val(Total(3).TEXT) + Val(Total(4).TEXT)) * RNI%
   BASIVA(2).Caption = TRIM$(CSTRING$(MKD$(TTTX#), 3, 12, 2, 2))
   Total(5).TEXT = CSTRING$(MKD$(Val(BASIVA(2).Caption) * Val(PORIVA(2).Caption) / 100), 3, 12, 2, 2)
   '
   TTTX# = 0: For KKI% = 2 To 5
                 TTTX# = TTTX# + Val(Total(KKI%).TEXT)
              Next KKI%
              TTTX# = TTTX# + Val(Text15)
   Total(6).TEXT = CSTRING$(MKD$(TTTX#), 3, 12, 2, 2)
   '
End Sub

Sub RECODOSI(CIXI%, CIXJ%, NPX&)
   '
   CIBIN$ = MKI$(CIXI%)
   CIXJ% = CIXI%
   If Asc(Mid$(CIBIN$, 2, 1)) = 0 Then
     If NPX& > 1600 Then
       If NPX& <= 1614 Then
         If NPX& = 1601 Then CIXJ% = CODINT%("041130M-PVT0")
         If NPX& = 1602 Then CIXJ% = CODINT%("042121T-PHH0")
         If NPX& = 1603 Then CIXJ% = CODINT%("041015M-PVT0")
         If NPX& = 1604 Then CIXJ% = CODINT%("041015M-PVT0")
         If NPX& = 1605 Then CIXJ% = CODINT%("041130M-PVT0")
         If NPX& = 1606 Then CIXJ% = CODINT%("041130M-PVT0")
         If NPX& = 1607 Then CIXJ% = CODINT%("042501T-PHH0")
         If NPX& = 1608 Then CIXJ% = CODINT%("44234")
         If NPX& = 1609 Then CIXJ% = CODINT%("44234")
         If NPX& = 1610 Then CIXJ% = CODINT%("LD54S2-PHCA-D13")
         If NPX& = 1611 Then CIXJ% = CODINT%("044501T-PTT0")
         If NPX& = 1612 Then CIXJ% = CODINT%("044251T-PTT0")
         If NPX& = 1613 Then CIXJ% = CODINT%("041130M-PVT0")
         If NPX& = 1614 Then CIXJ% = CODINT%("A4190")
       End If
     End If
   End If
   If CIXJ% < 1 Then CIXJ% = CIXI%
   '
End Sub

Sub CARMATODES()
   '
   CAPESEL% = 0
   UREODES& = ULTREG&("ORDESPA")
   Call CIERRARCH("!CARDEFAC")
   Call BLANARCH("!CARDEFAC")
   '   If TIDOC% <> 1 Then GoTo 1211
   '
   For U& = 1 To List2.ListCount
     If List2.Selected(U& - 1) = True Then
       CAPESEL% = CAPESEL% + 1
       TTXX$ = List2.List(U& - 1)
       NUODSEL& = Val(Mid$(TTXX$, 9, 10))
       REPRESU& = Val(Mid$(TTXX$, 121, 8))
       '
1205   TINI = Timer
       '
       For UJ& = REPRESU& To UREODES&
          X1$ = REGLEIDO$("ORDESPA", UJ&)
          NODES& = CVS(Left$(X1$, 4))
          If NODES& <> NUODSEL& Then GoTo 1209
          '
          REGODES& = UJ&
          REGPEDA& = CVS(Mid$(X1$, 111, 4))
          X$ = REGLEIDO$("PEDDETA", REGPEDA&)
          '
          If Asc(Mid$(X$, 27, 1)) <> 9 Then ' no esta anulado
            CAN = CVS(Mid$(X1$, 79, 4))
            If CAN >= 0.05 Then
              CADESPACH = CVS(Mid$(X1$, 83, 4))
              If CADESPACH < 0.1 Then
                CI% = CVI(Mid$(X$, 33, 2))
                If VERISTO% = 1 Then
                  If CI% > 0 Then
                    STACT = STOCKACT(CI%)
                    STDIS = STACT - CAN
                    If STDIS < 0 Then
                      If TIDOC% <> 3 Then
                        CAN = STACT
                        If CAN < 0.05 Then GoTo 1208
                      End If
                    End If
                  End If
                End If
                '
                Z$ = REGBLAN$("!CARDEFAC")
                Call REPLA(Z$, MKS$(CAN), 1, 4)      ' cantidad
                   COD$ = CODEXT$(CI%)
                Call REPLA(Z$, MKI$(CI%), 5, 2)
                Call REPLA(Z$, DESCRIT$(CI%), 7, 64)
                   TACOMED$ = Mid$(X$, 51, 8) + Mid$(X$, 35, 12) + Mid$(X$, 59, 12)
                Call REPLA(Z$, TACOMED$, 283, 32)
                   PRECIO = CVS(Mid$(X$, 71, 4))
                   If PRECIO < 0.0001 Then PRECIO = CVS(Mid$(X$, 83, 4))
                   PRELI = PRELISTA#(ARCHIPRE$, CI%)
                   If PRELI < PRECIO Then PRELI = PRECIO
                   PORDESCUEN = 0: If PRELI >= 0.0001 Then PORDESCUEN = 100 * (PRELI - PRECIO) / PRELI
                Call REPLA(Z$, MKS$(PRELI), 421, 4)
                Call REPLA(Z$, MKS$(PORDESCUEN), 425, 4)
                Call REPLA(Z$, MKS$(PRECIO), 429, 4)
                   IMPO# = CDbl(PRECIO) * CAN
                Call REPLA(Z$, MKD$(IMPO#), 433, 8)
                Call REPLA(Z$, MKS$(PORCEIVA(CI%)), 441, 4)
                Call REPLA(Z$, MKS$(REGODES&), 505, 4)
                Call REPLA(Z$, MKS$(REGPEDA&), 509, 4)
                '
                Call REGAPP("!CARDEFAC", Z$)
                CAITCAR% = CAITCAR% + 1
                '
                If NUOCOPAR$ <> "" Then
                  If NUOCOPAR$ <> NUOCOPARA$ Then
                    If NUOCOMPRA$ <> "" Then NUOCOMPRA$ = NUOCOMPRA$ + ", "
                    NUOCOMPRA$ = NUOCOMPRA$ + NUOCOPAR$
                    NUOCOPARA$ = NUOCOPAR$
                  End If
                End If
                '
                If NROPEDIPAR$ <> "" Then
                  If NROPEDIPAR$ <> NROPEDIPARA$ Then
                    If NROPEDIDO$ <> "" Then NROPEDIDO$ = NROPEDIDO$ + ", "
                    NROPEDIDO$ = NROPEDIDO$ + NROPEDIPAR$
                    NROPEDIPARA$ = NROPEDIPAR$
                  End If
                End If
                '
              End If
            End If
          End If
1208   Next UJ&
       '
     End If
1209 Next U&
     '
1210 Z$ = ""
     Call CIERRARCH("!CARDEFAC")
     '
     Text14.TEXT = NUOCOMPRA$
     '
End Sub
        
Sub CARMATPEDI()
   '
   CAPESEL% = 0
   UREODES& = ULTREG&("PEDDETA")
   Call CIERRARCH("!CARDEFAC")
   Call BLANARCH("!CARDEFAC")
   JJKK& = 0
   '   If TIDOC% <> 1 Then GoTo 1211
   '
   NUPRIPE& = 0: REPRIPE& = 0
   For U& = 1 To List2.ListCount
     If List2.Selected(U& - 1) = True Then
       CAPESEL% = CAPESEL% + 1
       TTXX$ = List2.List(U& - 1)
       NUODSEL& = Val(Mid$(TTXX$, 15, 10))
       REPRESU& = Val(Mid$(TTXX$, 121, 8))
       '
1205   TINI = Timer
       '
       For UJ& = REPRESU& To UREODES&
          REGPEDA& = UJ&
          X$ = REGLEIDO$("PEDDETA", REGPEDA&)
          NODES& = CVS(Left$(X$, 4))
          If NODES& <> NUODSEL& Then GoTo 1209
          '
          If Asc(Mid$(X$, 27, 1)) <> 9 Then ' no esta anulado
            CAN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
            If CAN >= 0.05 Then
              CI% = CVI(Mid$(X$, 33, 2))
              If VERISTO% = 1 Then
                If CI% > 0 Then
                  STACT = STOCKACT(CI%)
                  STDIS = STACT - CAN
                  If STDIS < 0 Then
                    If TIDOC% <> 3 Then
                      CAN = STDIS
                      If CAN < 0.05 Then GoTo 1208
                    End If
                  End If
                End If
              End If
              '
              If REPRIPE& < 1 Then
                NUPRIPE& = NODES&
                REPRIPE& = UJ&
              End If
              '
              MONEMA% = Asc(Mid$(X$, 52, 1))
              If JJKK& < 1 Then
                MONEMI% = Asc(Mid$(X$, 52, 1))
                If MONEMI% < 1 Then MONEMI% = 1
                If MONEMI% = 2 Then MONEMIS = "Dólares" Else MONEMIS = "Pesos"
              End If
              COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
              '
              Z$ = REGBLAN$("!CARDEFAC")
              Call REPLA(Z$, MKS$(CAN), 1, 4)      ' cantidad
                 COD$ = CODEXT$(CI%)
              Call REPLA(Z$, MKI$(CI%), 5, 2)
              Call REPLA(Z$, DESCRIT$(CI%), 7, 64)
                 TACOMED$ = Mid$(X$, 51, 8) + Mid$(X$, 35, 12) + Mid$(X$, 59, 12)
              Call REPLA(Z$, TACOMED$, 283, 32)
                 PRECIO = COEFCON * CVS(Mid$(X$, 83, 4))
                 'If PRECIO < 0.0001 Then PRECIO = CVS(Mid$(X$, 83, 4))
                 PRELI = PRELISTA#(ARCHIPRE$, CI%)
                 If PRELI < PRECIO Then PRELI = PRECIO
                 PORDESCUEN = 0: If PRELI >= 0.0001 Then PORDESCUEN = 100 * (PRELI - PRECIO) / PRELI
              Call REPLA(Z$, MKS$(PRELI), 421, 4)
              Call REPLA(Z$, MKS$(PORDESCUEN), 425, 4)
              Call REPLA(Z$, MKS$(PRECIO), 429, 4)
                 IMPO# = CDbl(PRECIO) * CAN
              Call REPLA(Z$, MKD$(IMPO#), 433, 8)
              Call REPLA(Z$, MKS$(PORCEIVA(CI%)), 441, 4)
              '
              Call REPLA(Z$, " ", 461, 1)
              If Mid$(X$, 51, 1) = "R" Then Call REPLA(Z$, "R", 461, 1)
              '
              Call REPLA(Z$, MKS$(REGODES&), 505, 4)
              Call REPLA(Z$, MKS$(REGPEDA&), 509, 4)
                 NPEDID& = CVS(Left$(X$, 4))
              Call REPLA(Z$, TRIM$(Str$(NPEDID&)), 371, 16)
              '
              JJKK& = JJKK& + 1
              Call REGAPP("!CARDEFAC", Z$)
              CAITCAR% = CAITCAR% + 1
              '
              NUOCOPAR$ = TRIM$(Mid$(X$, 9, 16))
              NROPEDIPAR$ = TRIM$(Str$(CVS(Left$(X$, 4))))
              '
              If NUOCOPAR$ <> "" Then
                If NUOCOPAR$ <> NUOCOPARA$ Then
                  If NUOCOMPRA$ <> "" Then NUOCOMPRA$ = NUOCOMPRA$ + ", "
                  NUOCOMPRA$ = NUOCOMPRA$ + NUOCOPAR$
                  NUOCOPARA$ = NUOCOPAR$
                End If
              End If
              '
              If NROPEDIPAR$ <> "" Then
                If NROPEDIPAR$ <> NROPEDIPARA$ Then
                  If NROPEDIDO$ <> "" Then NROPEDIDO$ = NROPEDIDO$ + ", "
                  NROPEDIDO$ = NROPEDIDO$ + NROPEDIPAR$
                  NROPEDIPARA$ = NROPEDIPAR$
                End If
              End If
              '
            End If
          End If
1208   Next UJ&
       '
     End If
1209 Next U&
     '
1210 Z$ = ""
     Call CIERRARCH("!CARDEFAC")
     '
     Text14.TEXT = NUOCOMPRA$
     '
     If REPRIPE& > 0 Then
       REGENCA& = CVS(Mid$(X$, 109, 4))
       If REGENCA& > 0 Then
         If REGENCA& <= ULTREG&("PEDENCA") Then
           YX$ = REGLEIDO$("PEDENCA", REGENCA&)
           If CVS(Left$(YX$, 4)) = NUPRIPE& Then GoTo 1300
         End If
       End If
       '
       For UK& = ULTREG&("PEDENCA") To 1 Step -1
         YX$ = REGLEIDO$("PEDENCA", UK&)
         If CVS(Left$(YX$, 4)) = NUPRIPE& Then GoTo 1300
       Next UK&
       GoTo 999
       '
1300   MONEMISA = MONEMIS
       TEXT21(0).TEXT = Str$(Asc(Mid$(YX$, 268, 1)))
       TEXT21(1).TEXT = Str$(Asc(Mid$(YX$, 269, 1)))
       TEXT21(2).TEXT = Str$(Asc(Mid$(YX$, 270, 1)))
       MONEMIS = MONEMISA
       '
       If Val(TEXT21(2).TEXT) < 1 Then
         Eco21(2).TEXT = "GERENCIA"
       End If
       '
       If Mid$(YX$, 275, 14) <> String$(14, 0) Then
          PORDESCU.TEXT = TRIM$(Mid$(YX$, 275, 14))
       End If
     End If
     '
999  End Sub
'
Sub CARMATBARR()
   '
   UREODES& = ULTREG&("PEDDETA")
   Call CIERRARCH("!CARDEFAC")
   Call BLANARCH("!CARDEFAC")
   JJKK& = 0
   '
   For U& = 1 To CARCOBARR.List1.ListCount
     TTXX$ = CARCOBARR.List1.List(U& - 1)
     CODE$ = TRIM$(Left$(TTXX$, 16))
     CANEF = Val(Mid$(TTXX$, 52, 6))
     '
     For J& = 1 To CARCOBARR.List2.ListCount
       TTYY$ = CARCOBARR.List2.List(J& - 1)
       If TRIM$(Mid$(TTYY$, 9, 16)) = CODE$ Then
         NODES& = Val(Mid$(TTYY$, 41, 8))
         REGPEDA& = Val(Mid$(TTYY$, 49, 8))
         X$ = REGLEIDO$("PEDDETA", REGPEDA&)
         '
         If Asc(Mid$(X$, 27, 1)) <> 9 Then ' no esta anulado
           CAN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
           If CAN >= 0.05 Then
             MONEMA% = Asc(Mid$(X$, 52, 1))
             If JJKK& < 1 Then
               MONEMI% = Asc(Mid$(X$, 52, 1))
               If MONEMI% < 1 Then MONEMI% = 1
               If MONEMI% = 2 Then MONEMIS = "Dólares" Else MONEMIS = "Pesos"
             End If
             COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
             '
             Z$ = REGBLAN$("!CARDEFAC")
             Call REPLA(Z$, MKS$(CANEF), 1, 4)      ' cantidad
                CIXI% = CODINT%(CODE$)
             Call REPLA(Z$, MKI$(CIXI%), 5, 2)
             Call REPLA(Z$, DESCRIT$(CIXI%), 7, 64)
                TACOMED$ = Mid$(X$, 51, 8) + Mid$(X$, 35, 12) + Mid$(X$, 59, 12)
             Call REPLA(Z$, TACOMED$, 283, 32)
                PRECIO = COEFCON * CVS(Mid$(X$, 83, 4))
                'If PRECIO < 0.0001 Then PRECIO = CVS(Mid$(X$, 83, 4))
                PRELI = PRELISTA#(ARCHIPRE$, CIXI%)
                If PRELI < PRECIO Then PRELI = PRECIO
                PORDESCUEN = 0: If PRELI >= 0.0001 Then PORDESCUEN = 100 * (PRELI - PRECIO) / PRELI
             Call REPLA(Z$, MKS$(PRELI), 421, 4)
             Call REPLA(Z$, MKS$(PORDESCUEN), 425, 4)
             Call REPLA(Z$, MKS$(PRECIO), 429, 4)
                IMPO# = CDbl(PRECIO) * CAN
             Call REPLA(Z$, MKD$(IMPO#), 433, 8)
             Call REPLA(Z$, MKS$(PORCEIVA(CIXI%)), 441, 4)
             '
             Call REPLA(Z$, " ", 461, 1)
             If Mid$(X$, 51, 1) = "R" Then Call REPLA(Z$, "R", 461, 1)
             '
             Call REPLA(Z$, MKS$(0), 505, 4)
             Call REPLA(Z$, MKS$(REGPEDA&), 509, 4)
                NPEDID& = CVS(Left$(X$, 4))
             Call REPLA(Z$, TRIM$(Str$(NPEDID&)), 371, 16)
             '
             JJKK& = JJKK& + 1
             Call REGAPP("!CARDEFAC", Z$)
             CAITCAR% = CAITCAR% + 1
             '
             NUOCOPAR$ = TRIM$(Mid$(X$, 9, 16))
             NROPEDIPAR$ = TRIM$(Str$(CVS(Left$(X$, 4))))
             '
             If NUOCOPAR$ <> "" Then
               If NUOCOPAR$ <> NUOCOPARA$ Then
                 If NUOCOMPRA$ <> "" Then NUOCOMPRA$ = NUOCOMPRA$ + ", "
                 NUOCOMPRA$ = NUOCOMPRA$ + NUOCOPAR$
                 NUOCOPARA$ = NUOCOPAR$
               End If
             End If
             '
             If NROPEDIPAR$ <> "" Then
               If NROPEDIPAR$ <> NROPEDIPARA$ Then
                 If NROPEDIDO$ <> "" Then NROPEDIDO$ = NROPEDIDO$ + ", "
                 NROPEDIDO$ = NROPEDIDO$ + NROPEDIPAR$
                 NROPEDIPARA$ = NROPEDIPAR$
               End If
             End If
             '
             CARCOBARR.List2.RemoveItem (J& - 1)
             GoTo 19
             '
           End If
         End If
1208   End If
       '
     Next J&
     '
     If MONEMI% < 1 Then MONEMI% = 1: COEFCON = 1
     Z$ = REGBLAN$("!CARDEFAC")
     Call REPLA(Z$, MKS$(CANEF), 1, 4)      ' cantidad
          CIXI% = CODINT%(CODE$)
     Call REPLA(Z$, MKI$(CIXI%), 5, 2)
     Call REPLA(Z$, DESCRIT$(CIXI%), 7, 64)
          PRELI = PRELISTA#(ARCHIPRE$, CIXI%)
          PRECIO = PRELI
          PORDESCUEN = 0
     Call REPLA(Z$, MKS$(PRELI), 421, 4)
     Call REPLA(Z$, MKS$(PORDESCUEN), 425, 4)
     Call REPLA(Z$, MKS$(PRECIO), 429, 4)
          IMPO# = CDbl(PRECIO) * CAN
     Call REPLA(Z$, MKD$(IMPO#), 433, 8)
     Call REPLA(Z$, MKS$(PORCEIVA(CIXI%)), 441, 4)
     '
     Call REPLA(Z$, " ", 461, 1)
     Call REPLA(Z$, MKS$(0), 505, 4)
     Call REPLA(Z$, MKS$(0), 509, 4)
          NPEDID& = CVS(Left$(X$, 4))
     Call REPLA(Z$, TRIM$(Str$(NPEDID&)), 371, 16)
     '
     JJKK& = JJKK& + 1
     Call REGAPP("!CARDEFAC", Z$)
     CAITCAR% = CAITCAR% + 1
     '
19 Next U&
   '
20 Z$ = ""
   Call CIERRARCH("!CARDEFAC")
   '
   Text14.TEXT = NUOCOMPRA$
   '
     If REPRIPE& > 0 Then
       REGENCA& = CVS(Mid$(X$, 109, 4))
       If REGENCA& > 0 Then
         If REGENCA& <= ULTREG&("PEDENCA") Then
           YX$ = REGLEIDO$("PEDENCA", REGENCA&)
           If CVS(Left$(YX$, 4)) = NUPRIPE& Then GoTo 1300
         End If
       End If
       '
       For UK& = ULTREG&("PEDENCA") To 1 Step -1
         YX$ = REGLEIDO$("PEDENCA", UK&)
         If CVS(Left$(YX$, 4)) = NUPRIPE& Then GoTo 1300
       Next UK&
       GoTo 999
       '
1300   MONEMISA = MONEMIS
       TEXT21(0).TEXT = Str$(Asc(Mid$(YX$, 268, 1)))
       TEXT21(1).TEXT = Str$(Asc(Mid$(YX$, 269, 1)))
       TEXT21(2).TEXT = Str$(Asc(Mid$(YX$, 270, 1)))
       MONEMIS = MONEMISA
       '
       If Val(TEXT21(2).TEXT) < 1 Then
         Eco21(2).TEXT = "GERENCIA"
       End If
       '
       If Mid$(YX$, 275, 14) <> String$(14, 0) Then
          PORDESCU.TEXT = TRIM$(Mid$(YX$, 275, 14))
       End If
     End If
     '
999 End Sub
'
Sub CARPEPEN(DPEDI$)
     '
     UREODES& = ULTREG&("PEDDETA")
     JJKK& = 0
     '   If TIDOC% <> 1 Then GoTo 1211
     '
     CAPESEL% = 1
     TTXX$ = DPEDI$
     NUODSEL& = Val(Mid$(TTXX$, 15, 10))
     REPRESU& = Val(Mid$(TTXX$, 121, 8))
     '
1205 TINI = Timer
     '
     SELITFAC.List2.Clear
     For UJ& = REPRESU& To UREODES&
       REGPEDA& = UJ&
       X$ = REGLEIDO$("PEDDETA", REGPEDA&)
       NODES& = CVS(Left$(X$, 4))
       If NODES& <> NUODSEL& Then GoTo 1209
       '
       If Asc(Mid$(X$, 27, 1)) <> 9 Then ' no esta anulado
         CAN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
         If CAN >= 0.05 Then
           CI% = CVI(Mid$(X$, 33, 2))
           TTYY$ = Space$(128)
           Call REPLA(TTYY$, CODEXT$(CI%), 1, 16)
           Call REPLA(TTYY$, DESCRIT0$(CI%), 19, 46)
           Call REPLA(TTYY$, FORMATNUM$(CAN, "F10.1"), 66, 10)
           Call REPLA(TTYY$, FORMATNUM$(REGODES&, "F8.0"), 113, 8)
           Call REPLA(TTYY$, FORMATNUM$(REGPEDA&, "F8.0"), 121, 8)
           SELITFAC.List2.AddItem TTYY$
           '
         End If
       End If
1208 Next UJ&
     '
1209 SELITFAC.Top = Top + 2140
     SELITFAC.Left = Left + 240
     SELITFAC.Show 1
     '
End Sub
'
Sub CARITSEL()
   '
   CAPESEL% = 1
   UREODES& = ULTREG&("PEDDETA")
   Call CIERRARCH("!CARDEFAC")
   Call BLANARCH("!CARDEFAC")
   JJKK& = 0
   '
   NUPRIPE& = 0: REPRIPE& = 0
   For U& = 1 To SELITFAC.List2.ListCount
     If SELITFAC.List2.Selected(U& - 1) = True Then
       CAPESEL% = CAPESEL% + 1
       TTXX$ = SELITFAC.List2.List(U& - 1)
       'NUODSEL& = Val(Mid$(TTXX$, 15, 10))
       REGPEDA& = Val(Mid$(TTXX$, 121, 8))
       REGODES& = Val(Mid$(TTXX$, 113, 8))
       '
1205   TINI = Timer
       '
       X$ = REGLEIDO$("PEDDETA", REGPEDA&)
       NODES& = CVS(Left$(X$, 4))
       '
       If Asc(Mid$(X$, 27, 1)) <> 9 Then ' no esta anulado
         CAN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
         If CAN >= 0.05 Then
           CI% = CVI(Mid$(X$, 33, 2))
           If VERISTO% = 1 Then
             If CI% > 0 Then
               STACT = STOCKACT(CI%)
               STDIS = STACT - CAN
               If STDIS < 0 Then
                 If TIDOC% <> 3 Then
                   CAN = STDIS
                   If CAN < 0.05 Then GoTo 28
                 End If
               End If
             End If
           End If
           '
           If REPRIPE& < 1 Then
             NUPRIPE& = NODES&
             REPRIPE& = UJ&
           End If
           '
           MONEMA% = Asc(Mid$(X$, 52, 1))
           If JJKK& < 1 Then
             MONEMI% = Asc(Mid$(X$, 52, 1))
             If MONEMI% < 1 Then MONEMI% = 1
             If MONEMI% = 2 Then MONEMIS = "Dólares" Else MONEMIS = "Pesos"
           End If
           COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
           '
           Z$ = REGBLAN$("!CARDEFAC")
           Call REPLA(Z$, MKS$(CAN), 1, 4)      ' cantidad
              COD$ = CODEXT$(CI%)
           Call REPLA(Z$, MKI$(CI%), 5, 2)
           Call REPLA(Z$, DESCRIT$(CI%), 7, 64)
              TACOMED$ = Mid$(X$, 51, 8) + Mid$(X$, 35, 12) + Mid$(X$, 59, 12)
           Call REPLA(Z$, TACOMED$, 283, 32)
              PRECIO = COEFCON * CVS(Mid$(X$, 83, 4))
              'If PRECIO < 0.0001 Then PRECIO = CVS(Mid$(X$, 83, 4))
              PRELI = PRELISTA#(ARCHIPRE$, CI%)
              If PRELI < PRECIO Then PRELI = PRECIO
              PORDESCUEN = 0: If PRELI >= 0.0001 Then PORDESCUEN = 100 * (PRELI - PRECIO) / PRELI
           Call REPLA(Z$, MKS$(PRELI), 421, 4)
           Call REPLA(Z$, MKS$(PORDESCUEN), 425, 4)
           Call REPLA(Z$, MKS$(PRECIO), 429, 4)
              IMPO# = CDbl(PRECIO) * CAN
           Call REPLA(Z$, MKD$(IMPO#), 433, 8)
           Call REPLA(Z$, MKS$(PORCEIVA(CI%)), 441, 4)
           '
           Call REPLA(Z$, " ", 461, 1)
           If Mid$(X$, 51, 1) = "R" Then Call REPLA(Z$, "R", 461, 1)
           '
           Call REPLA(Z$, MKS$(REGODES&), 505, 4)
           Call REPLA(Z$, MKS$(REGPEDA&), 509, 4)
              NPEDID& = CVS(Left$(X$, 4))
           Call REPLA(Z$, TRIM$(Str$(NPEDID&)), 371, 16)
           '
           JJKK& = JJKK& + 1
           Call REGAPP("!CARDEFAC", Z$)
           CAITCAR% = CAITCAR% + 1
           '
           NUOCOPAR$ = TRIM$(Mid$(X$, 9, 16))
           NROPEDIPAR$ = TRIM$(Str$(CVS(Left$(X$, 4))))
           '
           If NUOCOPAR$ <> "" Then
             If NUOCOPAR$ <> NUOCOPARA$ Then
               If NUOCOMPRA$ <> "" Then NUOCOMPRA$ = NUOCOMPRA$ + ", "
               NUOCOMPRA$ = NUOCOMPRA$ + NUOCOPAR$
               NUOCOPARA$ = NUOCOPAR$
             End If
           End If
           '
           If NROPEDIPAR$ <> "" Then
             If NROPEDIPAR$ <> NROPEDIPARA$ Then
               If NROPEDIDO$ <> "" Then NROPEDIDO$ = NROPEDIDO$ + ", "
               NROPEDIDO$ = NROPEDIDO$ + NROPEDIPAR$
               NROPEDIPARA$ = NROPEDIPAR$
             End If
           End If
           '
         End If
       End If
     End If
28 Next U&
     '
1210 Z$ = ""
     Call CIERRARCH("!CARDEFAC")
     '
     Text14.TEXT = NUOCOMPRA$
     '
     If REPRIPE& > 0 Then
       REGENCA& = CVS(Mid$(X$, 109, 4))
       If REGENCA& > 0 Then
         If REGENCA& <= ULTREG&("PEDENCA") Then
           YX$ = REGLEIDO$("PEDENCA", REGENCA&)
           If CVS(Left$(YX$, 4)) = NUPRIPE& Then GoTo 1300
         End If
       End If
       '
       For UK& = ULTREG&("PEDENCA") To 1 Step -1
         YX$ = REGLEIDO$("PEDENCA", UK&)
         If CVS(Left$(YX$, 4)) = NUPRIPE& Then GoTo 1300
       Next UK&
       GoTo 999
       '
1300   MONEMISA = MONEMIS
       TEXT21(0).TEXT = Str$(Asc(Mid$(YX$, 268, 1)))
       TEXT21(1).TEXT = Str$(Asc(Mid$(YX$, 269, 1)))
       TEXT21(2).TEXT = Str$(Asc(Mid$(YX$, 270, 1)))
       MONEMIS = MONEMISA
       '
       If Val(TEXT21(2).TEXT) < 1 Then
         Eco21(2).TEXT = "GERENCIA"
       End If
       '
       If Mid$(YX$, 275, 14) <> String$(14, 0) Then
          PORDESCU.TEXT = TRIM$(Mid$(YX$, 275, 14))
       End If
     End If
     '
999 End Sub
'
Sub BLANPEFAC(REMI, FF%)  '  blanqueo pedidos pendientes - asociado a FAC2000
    '
    If CONSICLI% = 1 Then Exit Sub
    '
    Call BLOQARCH("PEDDETA")
    Call BLOQARCH("ORDESPA")
    Call BLOQARCH("REMPEN")
    '
    Screen.MousePointer = 11
    '
PEDIANT& = 0
    FIN& = ULTREG&("!CARDEFAC")
    For U& = 1 To FIN&
       X$ = REGLEIDO$("!CARDEFAC", U&)
       CIENT% = CVI(Mid$(X$, 5, 2))
       CAENTRE = CVS(Left$(X$, 4))
       If CAENTRE > 0 Then
         REGPEDA& = CVS(Mid$(X$, 509, 4))
         If REGPEDA& > 0 Then
           If REGPEDA& <= ULTREG&("PEDDETA") Then
             Y$ = REGLEIDO$("PEDDETA", REGPEDA&)
             If CVI(Mid$(Y$, 33, 2)) = CIENT% Then ' debe coincidir codigo
               TOENTRE = CAENTRE + CVS(Mid$(Y$, 95, 4))
               SALPEN = CVS(Mid$(Y$, 79, 4)) - TOENTRE: If SALPEN < 0 Then SALPEN = 0
               '
               CASIG = 0: If Mid$(Y$, 75, 4) <> "    " Then CASIG = CVS(Mid$(Y$, 75, 4))
               CASIG = CASIG - CAENTRE: If CASIG < 0 Then CASIG = 0
               Call REPLA(Y$, MKS$(CASIG), 75, 4)
               '
               Call REPLA(Y$, MKS$(TOENTRE), 95, 4)
               Call REPLA(Y$, MKS$(SALPEN), 99, 4)
               Call REPLA(Y$, MKI$(FF%), 103, 2)
               Call REPLA(Y$, MKS$(REMI), 105, 4)
               Call GRAREG("PEDDETA", Y$, REGPEDA&)
               '
               NPEDIDX& = CVS(Left$(Y$, 4))
               If NPEDIDX& <> PEDIANT& Then
                 If PEDIANT& <> 0 Then
                   Call REGAPP("ANOTAPED", Left$(MKS$(PEDIANT&) + Space$(10) + String$(56, "-"), 96))
                   Call REGAPP("ANOTAPED", MKS$(PEDIANT&) + String$(88, "="))
                 End If
                 PEDIANT& = NPEDIDX&
                 COPIASFAC% = Val(FINFACTURA.Text2(0))
                 COPIASREM% = Val(FINFACTURA.Text2(1))
                 TTXX$ = ""
                 If COPIASREM% > 0 Then TTXX$ = "Remito " + TRIM$(Text10)
                 If COPIASFAC% > 0 Then TTXX$ = TTXX$ + " - Factura " + TRIM$(Label12) + "-" + TRIM$(Text11)
                 TTXX$ = TTXX$ + ":"
                 '
                 TTXYZ$ = Space$(76)
                 HOII% = HOY(HOS$)
                 Call REPLA(TTXYZ$, HOS$, 1, 8)
                 Call REPLA(TTXYZ$, TTXX$, 11, 56)
                 Call REPLA(TTXYZ$, USERTER$, 69, 24)
                 Call REGAPP("ANOTAPED", MKS$(NPEDIDX&) + TTXYZ$)
                 Call REPLA(TTXYZ$, Space$(88), 1, 88)
                 Call REGAPP("ANOTAPED", MKS$(NPEDIDX&) + TTXYZ$)
                 '
               End If
               '
               TTXX$ = Space$(76)
               Call REPLA(TTXX$, "- " + CODEXT$(CIENT%), 11, 18)
               Call REPLA(TTXX$, DESCRIT0$(CIENT%), 29, 38)
               Call REPLA(TTXX$, CSTRING$(MKS$(CAENTRE), 3, 10, 1, 2), 57, 10)
               Call REGAPP("ANOTAPED", Left$(MKS$(NPEDIDX&) + TTXX$, 96))
               '
             End If
           End If
         End If
         '
         REGODES& = CVS(Mid$(X$, 505, 4))
         If REGODES& > 0 Then
           If REGODES& <= ULTREG&("ORDESPA") Then
             Y$ = REGLEIDO$("ORDESPA", REGODES&)
             If CVI(Mid$(Y$, 37, 2)) = CIENT% Then ' debe coincidir codigo
               TOENTRE = CAENTRE + CVS(Mid$(Y$, 83, 4))
               SALPEN = CVS(Mid$(Y$, 79, 4)) - TOENTRE
               If SALPEN < 0 Then SALPEN = 0
               '
               Call REPLA(Y$, MKS$(TOENTRE), 83, 4)
               Call REPLA(Y$, MKS$(SALPEN), 87, 4)
               Call REPLA(Y$, MKI$(FF%), 91, 2)
               Call REPLA(Y$, MKS$(REMI), 99, 4)
               Call REPLA(Y$, MKS$(REMI), 103, 4)
               Call GRAREG("ORDESPA", Y$, REGODES&)
               '
             End If
           End If
         End If
       End If
    Next U&
    '
    Call REGAPP("ANOTAPED", Left$(MKS$(PEDIANT&) + Space$(10) + String$(56, "-"), 96))
    Call REGAPP("ANOTAPED", MKS$(PEDIANT&) + String$(88, "="))
    '
    For U& = 1 To FIN&
       X$ = REGLEIDO$("!CARDEFAC", U&)
       CIENT% = CVI(Mid$(X$, 5, 2))
       CAENTRE = CVS(Left$(X$, 4))
       If CAENTRE > 0 Then
         REGPEDPE& = CVS(Mid$(X$, 509, 4))
         If REGPEDPE& < 0 Then
           REGREMPE& = Abs(REGPEDPE&)
           If REGREMPE& <= ULTREG&("REMPEN") Then
             Y$ = REGLEIDO$("REMPEN", REGREMPE&)
             '
             Call REPLA(Y$, String$(96, 0), 1, 96)
             Call GRAREG("REMPEN", Y$, REGREMPE&)
             '
           End If
         End If
       End If
    Next U&
    '
    Call CIERRARCH("PEDDETA")
    Call CIERRARCH("ORDESPA")
    Call CIERRARCH("REMPEN")
    Call CIERRARCH("ANOTAPED")
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    '
    Screen.MousePointer = 1
    '
End Sub
'
Sub LEETASIVA()
   '
   URETA& = ULTREG&("GRUPIVA")
   If URETA& < 2 Then
     Call MENSERR(24, "Falta Definir Grupos de I.V.A.")
     Call CIERRARCH("*.*")
     Call FINAL("")
   End If
   '
   For U& = 1 To 2
      X$ = REGLEIDO$("GRUPIVA", U&)
      TAIVA(U& - 1) = CVS(Mid$(X$, 13, 4))
      PORIVA(U& - 1) = TRIM$(CSTRING$(MKS$(TAIVA(U& - 1)), 3, 6, 2, 2))
   Next U&
   PORIVA(2) = "50.00" ' responsable no inscripto
   Call CIERRARCH("GRUPIVA")
   ALIIVA = ALIVA
   '
End Sub

Sub CARCHCONSI()        ' carga inicial archivo material consignado
   '
   If ULTREG&("REMPEN") > 0 Then
     Screen.MousePointer = 11
     FIN& = ULTREG&("REMPEN")
     UREPE& = ULTREG&("PEDDETA")
     REMIA& = 0
     For U& = 1 To FIN&
       XX$ = REGLEIDO$("REMPEN", U&)
       NCLIE% = CVI(Left$(XX$, 2))
       If NCLIE% > 0 Then
         NREMIT& = CVS(Mid$(XX$, 3, 4))
         FEREMIT% = CVI(Mid$(XX$, 7, 2))
         CIREMI% = CVI(Mid$(XX$, 9, 2))
         If CIREMI% > 0 Then
           If CIREMI% <= NUMAS% Then
             CAREMI# = CVS(Mid$(XX$, 27, 4))
             If CAREMI# > 0.05 Then
               '
               If NREMIT& <> REMIA& Then
                 REMIA& = NREMIT&
                 YXI% = 0
               End If
               CMO$ = "RT"
               PREFIREM$ = CONTROL$("", "PREFIREM")
               If PREFIREM$ = "" Then PREFIREM$ = "0001"
               NROREMI$ = TRIM$(Str$(NREMIT&))
               While Len(NROREMI$) < 8: NROREMI$ = "0" + NROREMI$: Wend
               NROREMI$ = CMO$ + "." + PREFIREM$ + "-" + NROREMI$
               '
               HO% = HOY(HOS$)
               HP = 100 * Val(Left$(Time$, 2)) + Val(Mid$(Time$, 4, 2))
               User% = USNBR% Mod 100
               CTRXX$ = MKI$(HO%) + MKI$(Int(HP)) + Chr$(User%) + String$(2, 0)
               '
               NPEDI& = 0: PREUNID# = 0
               For KU& = 1 To UREPE&
                 YY$ = REGLEIDO$("PEDDETA", KU&)
                 If CVS(Mid$(YY$, 105, 4)) = NREMIT& Then
                   If CVI(Mid$(YY$, 33, 2)) = CIREMI% Then
                     NPEDI& = CVS(Left$(YY$, 4))
                     PREUNID# = CVS(Mid$(YY$, 83, 4))
                   End If
                 End If
               Next KU&
               '
               RECONSI$ = REGBLAN$("CONSICLI")
               Call REPLA(RECONSI$, MKI$(NCLIE%), 1, 2)
               Call REPLA(RECONSI$, MKI$(FEREMIT%), 3, 2)
               Call REPLA(RECONSI$, NROREMI$, 5, 18)
               YXI% = YXI% + 1
               Call REPLA(RECONSI$, "RT." + TRIM$(Str$(NREMIT&)) + "-" + TRIM$(Str$(YXI%)), 23, 12)
               Call REPLA(RECONSI$, MKI$(CIREMI%), 35, 2)
               Call REPLA(RECONSI$, MKD$(CAREMI#), 37, 8)
               Call REPLA(RECONSI$, MKS$(NPEDI&), 109, 4)
               Call REPLA(RECONSI$, MKD$(PREUNID#), 113, 8)
               Call REPLA(RECONSI$, CTRXX$, 122, 7)
               Call REGAPP("CONSICLI", RECONSI$)
               '
               Call GRAREG("REMPEN", String$(128, 0), U&)
               '
             End If
           End If
         End If
       End If
     Next U&
   End If
   '
   Call CIERRARCH("CONSICLI")
   Call CIERRARCH("REMPEN")
   Screen.MousePointer = 1
   '
End Sub

Sub GRACONSI()
   '
   SUCUVEN% = 0
   DEPOSALI% = 1
   '
   NC1% = Val(Text1.TEXT): If NC% < 1 Then NC1% = 10001
   REFE$ = ECOARCH$("DEUDOR1", MKI$(NC%))
   FIN& = ULTREG&("!CARDEFAC")
   '
   For I& = 1 To FIN&
     Call TRACE(20, I&, FIN&)
     X$ = REGLEIDO$("!CARDEFAC", I&)
     CANTAA = CVS(Left$(X$, 4))
     If CANTAA > 0 Then
       CI% = CVI(Mid$(X$, 5, 2))
       If CI% > 0 Then
         '
         COD$ = CODEXT$(CI%)
         PRE# = CDbl(CVS(Mid$(X$, 429, 4)))
         NPEDI& = Val(Mid$(X$, 371, 8))
         CMOV$ = "RT"
         If Mid$(X$, 461, 1) = "R" Then CMOV$ = "*R" ' REPARACION
         NROCON& = Val(Text11.TEXT)
         DOSEC$ = "FC." + TRIM$(Str$(NROCON&))
         '
         REMIRX$ = Mid$(X$, 387, 16)
         DOPRI$ = "RT." + TRIM$(Str$(Val(Mid$(REMIRX$, 9, 8))))
         '
         FF% = FECHANUM(Text9.TEXT)
         REFE$ = "Factura " + DOSEC$
         Call MOVISTO(FF%, CI%, LOTE$, CMOV$, DOPRI$, "", REFE$, PRE#, "$", NC%, DEPOSALI%, CANTAA)
         REFE$ = ECOARCH$("DEUDOR1", MKI$(NC%))
         Call MOVISTO(FF%, CI%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, PRE#, "$", NC%, DEPOSALI%, (-1) * CANTAA)
         '
         NCN$ = TRIM$(Str$(NROCON&))
         While Len(NCN$) < 8: NCN$ = "0" + NCN$: Wend
         PREFI$ = Label12.Caption
         NCN$ = PREFI$ + "-" + NCN$
         '
         HO% = HOY(HOS$)
         HP = 100 * Val(Left$(Time$, 2)) + Val(Mid$(Time$, 4, 2))
         User% = USNBR% Mod 100
         CTRXX$ = MKI$(HO%) + MKI$(Int(HP)) + Chr$(User%) + String$(2, 0)
         '
         RECONSI$ = REGBLAN$("CONSICLI")
         Call REPLA(RECONSI$, MKI$(NC%), 1, 2)
         Call REPLA(RECONSI$, MKI$(FF%), 3, 2)
         Call REPLA(RECONSI$, NCN$, 5, 18)
         Call REPLA(RECONSI$, MKI$(CI%), 35, 2)
         Call REPLA(RECONSI$, MKD$(CDbl(CANTAA)), 45, 8)
         Call REPLA(RECONSI$, MKS$(NPEDI&), 109, 4)
         Call REPLA(RECONSI$, MKD$(PRE#), 113, 8)
         Call REPLA(RECONSI$, CTRXX$, 122, 7)
         Call REGAPP("CONSICLI", RECONSI$)
       End If
     End If
   Next I&
   '
End Sub

Sub CARTAMONED()
   If DERACCE%("TICAMBIO", "Ajustar Tipo de Cambio Central") < 2 Then
     Exit Sub
   End If
   '
   Call TRALOCAL("TAMONED", "")
   VTB% = VENTABU%("TAMONED/NC/TITUPAN=Monedas y Tipo de Cambio")
   If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!TAMONED")
         X$ = REGLEIDO$("!TAMONED", U&)
         Call REPLA(X$, Chr$(U&), 1, 1)
         Call GRAREG("TAMONED", X$, U&)
      Next U&
      FIN& = ULTREG&("!TAMONED")
      Call SETULTREG("TAMONED", FIN&)
   End If
   Call CIERRARCH("TAMONED")
   TICAMBIS = FORMATNUM(TICAMONE(2), "F9.4")
End Sub

Sub SHOWPEDIPEN()
     '
     Static UFEPEDET#
     If CDbl(FileDateTime(LUDAT$ + "PEDDETA.DAT")) <> UFEPEDET# Then
        UFEPEDET# = CDbl(FileDateTime(LUDAT$ + "PEDDETA.DAT"))
        '
        JJ& = 0
        UREPRE& = ULTREG&("PEDDETA")
        Screen.MousePointer = 11
        JJ& = 0
        NROPEDA& = 0: CAPED% = 0
        FINF = 999999: FSUP = 0
        For REPRESU& = 1 To UREPRE&
          Call TRACE(20, REPRESU&, UREPRE&)
          X$ = REGLEIDO$("PEDDETA", REPRESU&)
          If Asc(Mid$(X$, 27, 1)) <> 9 Then ' no esta anulado
            CAN = CVS(Mid$(X$, 79, 4))
            CADESPACH = CVS(Mid$(X$, 95, 4))
            CAPENDENT = CAN - CADESPACH
            If CAPENDENT > 0.5 Then
              NRODES& = CVS(Left$(X$, 4))
              If NRODES& <> NRODESA& Then
                PRIREDETA& = REPRESU&
                NRODESA& = NRODES&
                TTXX$ = MKS$(NRODES&)
                NOCO$ = TRIM$(Mid$(X$, 9, 16))
                If NOCO$ <> "" Then
                  TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
                End If
                FEX = CVI(Mid$(X$, 5, 2))
                TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
                NCLIE% = CVI(Mid$(X$, 7, 2))
                TTXX$ = TTXX$ + TRIM$(RASOCLI$(NCLIE%)) + Space$(52)
                REGENCA& = CVS(Mid$(X$, 109, 4))
                JJ& = JJ& + 1
                Call GRAREG("INDIPEP", Left$(TTXX$, 52) + MKS$(REGENCA&), JJ&)
              End If
            End If
          End If
        Next REPRESU&
        Call SETULTREG("INDIPEP", JJ&)
        Call CIERRARCH("INDIPEP")
        '
        Screen.MousePointer = 1
        YAINDI% = 1
     End If
     '
2112 If ULTREG&("INDIPEP") < 1 Then
        Call COMUNI("No hay O/V's Pendientes\que Puedan Consultarse")
        GoTo 9999
     End If
     '
2110 X$ = REGSEL$("INDIPEP/Ordenes de Venta Pendientes de Entrega")
     If Len(X$) <= 4 Then
        GoTo 9999
     End If
     '
     NPX& = CVS(Left$(X$, 4))
     If NPX& < 1 Then GoTo 9999
     '
     TIDOCUM$ = "Orden de Despacho"
     If InStr(UCase$(EMPREAC$), "VDR") > 0 Then TIDOCUM$ = "Pedido de Cliente"
     NUDOCU$ = TRIM$(Str$(NPX&))
     While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
     NUDOCU$ = NUDOCU$ + " "
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
     Screen.MousePointer = 1
     Call COMUNI("Documento no Encontrado")
     Exit Sub
     '
     '16/03/07 (OT-07-0113)
5    Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
     '16/03/07 (FIN)
'5    For KKI% = 0 To 5
'       EDITFORM.Picture1(KKI%).Cls
'       EDITFORM.Picture1(KKI%).DrawWidth = 3
'       EDITFORM.Picture1(KKI%).DrawStyle = 0
'       EDITFORM.Picture1(KKI%).FillStyle = 1
'     Next KKI%
'     '
'     REDOCU$ = REGACT$("PDOCLST")
'     URE& = ULTREG&("PDOCSPL")
'     LI& = 0: LS& = URE&
'6    E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
'     L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
'     DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'     If DCC& < DOCUNU& Then LI& = L&: GoTo 6
'     If DCC& > DOCUNU& Then LS& = L&: GoTo 6
'     GoTo 8
'     '
'7    Screen.MousePointer = 1
'     Call MENSERR(24, "Documento no Registrado")
'     GoTo 9999
'     '
'8    TPSP$ = ""
'     Screen.MousePointer = 11
'     While L& > 1
'       DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'       If DCC& < DOCUNU& Then GoTo 9
'       L& = L& - 1
'     Wend
'     '
'9    CAPAGINAS% = 0
'     For KKL& = L& To URE&
'       TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
'       If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
'       If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
'         TPSP$ = TPSP$ + Mid$(TBUF$, 5)
'       End If
'     Next KKL&
'     '
'95   SPOOLSTRING$ = TPSP$
'     For KKI% = 0 To 5
'       EDITFORM.Picture1(KKI%).Cls
'       EDITFORM.Picture1(KKI%).Top = 0
'       EDITFORM.Picture1(KKI%).Refresh
'       EDITFORM.Picture1(KKI%).Height = 7170
'       EDITFORM.Picture1(KKI%).Width = 11700
'       EDITFORM.Picture1(KKI%).DrawWidth = 3
'       EDITFORM.Picture1(KKI%).DrawStyle = 0
'       EDITFORM.Picture1(KKI%).FillStyle = 1
'     Next KKI%
'     ALTUPAGINA = 7170
'     TOPEPAGINA = 0
'     PAGINACTIVA% = 0
'     CAPAGINAS% = 0
'     '
'10   PPXX% = InStr(TPSP$, Chr$(255))
'     If PPXX% > 0 Then
'       COPRIAT$ = Left$(TPSP$, PPXX% - 1)
'       TPSP$ = Mid$(TPSP$, PPXX% + 1)
'       Call TRACE(20, FIN& - Len(TPSP$), FIN&)
'       '
'       If Len(COPRIAT$) > 0 Then
'         COPRI% = Asc(Left$(COPRIAT$, 1))
'         ATRIP$ = Mid$(COPRIAT$, 2)
'         Call SHOWCOMMAND(COPRI%, ATRIP$)
'       End If
'       GoTo 10
'       '
'     End If
'     '
'     Screen.MousePointer = 1
'     PAGINACTIVA% = 0
'     For KKI% = 0 To 5
'       EDITFORM.Picture1(KKI%).Visible = False
'     Next KKI%
'     EDITFORM.Picture1(PAGINACTIVA%).Visible = True
'     EDITFORM.Picture1(PAGINACTIVA%).Refresh
'     EDITFORM.VScroll1.Min = 0
'     EDITFORM.VScroll1.Value = 0
'     EDITFORM.VScroll1.Max = 1
'     If ALTUPAGINA > EDITFORM.Frame1.Height Then
'       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
'     End If
'     EDITFORM.Command3.Enabled = False
'     EDITFORM.Command5.Enabled = False
'     EDITFORM.Show 1
     '
     GoTo 2110
     '
9999 End Sub
'\\\OT-05-0524-WAR-25/07/05///
Sub GRAFABON()
    '
    Call BLOQARCH("FACBONIF")
    NCLIE% = Val(Text1.TEXT)
    FEMI% = FECHANUM(Text9.TEXT)
    NUFA1 = Val(Text11.TEXT)
    TIPO% = TIFADOC%
    VA% = VARIDOC%
    NRO& = NUFA1
    'NUDOCU1$ = FORMDOC$(NRO&, TIPO%, VA%, SUCUVEN%)
    '
    SUCUVEN% = 0
    On Error Resume Next
    SUCUVEN% = Val(Label12)
    On Error GoTo 0
    NUDOCU1$ = FORMDOC$(NRO&, TIPO%, VA%, SUCUVEN%)
    '
    SIG% = 1: If TIFADOC% = 6 Then SIG% = (-1)
    '
    IBRUTO1# = Val(Total(0)) * SIG%
    PORDES1$ = TRIM(PORDESCU)
    IDESC1# = Val(Total(1)) * SIG%
    PORBON1$ = TRIM(Text16)
    IBON1# = Val(Total(7)) * SIG%
    TOTNET1# = Val(Total(2)) * SIG%
    COEFI1# = 1
    If Abs(IBRUTO1#) >= 0.01 Then COEFI1# = TOTNET1# / IBRUTO1# * 100
    '
    X$ = REGBLAN("FACBONIF")
    UREG1& = ULTREG("FACBONIF")
    PREG1& = UREG1& + 1
    Call REPLA(X$, MKI$(NCLIE%), 1, 2)
    Call REPLA(X$, MKS$(NUFA1), 3, 4)
    Call REPLA(X$, MKI$(FEMI%), 7, 2)
    Call REPLA(X$, Chr$(TIPO%), 9, 1)
    Call REPLA(X$, Chr$(VA%), 10, 1)
    Call REPLA(X$, NUDOCU1$, 11, 18)
    Call REPLA(X$, MKI$(NCLIE%), 29, 2)
    Call REPLA(X$, MKD$(IBRUTO1#), 33, 8)
    Call REPLA(X$, PORDES1$, 41, 16)
    Call REPLA(X$, MKD$(IDESC1#), 57, 8)
    Call REPLA(X$, PORBON1$, 65, 16)
    Call REPLA(X$, MKD$(IBON1#), 81, 8)
    Call REPLA(X$, MKD$(TOTNET1#), 89, 8)
    Call REPLA(X$, MKD$(COEFI1#), 97, 8)
    Call GRAREG("FACBONIF", X$, PREG1&)
    Call SETULTREG("FACBONIF", PREG1&)
    '
    Call CIERRARCH("FACBONIF")
    '
End Sub
'\\\OT-05-0524-WAR-FIN///
