VERSION 5.00
Begin VB.Form FORCONFA04 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FORIMPRE"
   ClientHeight    =   13230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   13230
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   13600
      Left            =   0
      ScaleHeight     =   13545
      ScaleWidth      =   10980
      TabIndex        =   0
      Top             =   0
      Width           =   11040
      Begin VB.Frame Frame5 
         Caption         =   "Varios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1800
         Left            =   210
         TabIndex        =   163
         Top             =   11235
         Width           =   10620
         Begin VB.Label Label12 
            Caption         =   "Varios(5)"
            Height          =   225
            Index           =   5
            Left            =   7560
            TabIndex        =   187
            Top             =   1365
            Width           =   750
         End
         Begin VB.Label Label12 
            Caption         =   "Varios(4)"
            Height          =   225
            Index           =   4
            Left            =   7560
            TabIndex        =   186
            Top             =   1155
            Width           =   750
         End
         Begin VB.Label Label12 
            Caption         =   "Varios(3)"
            Height          =   225
            Index           =   3
            Left            =   7560
            TabIndex        =   185
            Top             =   945
            Width           =   750
         End
         Begin VB.Label Label12 
            Caption         =   "Varios(2)"
            Height          =   225
            Index           =   2
            Left            =   7560
            TabIndex        =   184
            Top             =   735
            Width           =   750
         End
         Begin VB.Label Label12 
            Caption         =   "Varios(1)"
            Height          =   225
            Index           =   1
            Left            =   7560
            TabIndex        =   183
            Top             =   525
            Width           =   750
         End
         Begin VB.Label Label12 
            Caption         =   "Varios(0)"
            Height          =   225
            Index           =   0
            Left            =   7560
            TabIndex        =   182
            Top             =   315
            Width           =   750
         End
         Begin VB.Label Varios 
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
            Height          =   225
            Index           =   5
            Left            =   3885
            TabIndex        =   181
            Top             =   1365
            Width           =   3420
         End
         Begin VB.Label Varios 
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
            Height          =   225
            Index           =   4
            Left            =   3885
            TabIndex        =   180
            Top             =   1155
            Width           =   3420
         End
         Begin VB.Label Varios 
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
            Height          =   225
            Index           =   3
            Left            =   3885
            TabIndex        =   179
            Top             =   945
            Width           =   3420
         End
         Begin VB.Label Varios 
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
            Height          =   225
            Index           =   2
            Left            =   3885
            TabIndex        =   178
            Top             =   735
            Width           =   3420
         End
         Begin VB.Label Varios 
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
            Height          =   225
            Index           =   1
            Left            =   3885
            TabIndex        =   177
            Top             =   525
            Width           =   3420
         End
         Begin VB.Label Label117 
            Caption         =   "DEN-CUEN"
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
            Left            =   8610
            TabIndex        =   176
            Top             =   1155
            Width           =   1665
         End
         Begin VB.Label Label115 
            Caption         =   "Guia de Importacion"
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
            Left            =   210
            TabIndex        =   175
            Top             =   1155
            Width           =   3990
         End
         Begin VB.Label Label114 
            Caption         =   "NUME-CHE"
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
            Left            =   8610
            TabIndex        =   174
            Top             =   1365
            Width           =   1665
         End
         Begin VB.Label Label112 
            Caption         =   "Credito"
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
            Left            =   210
            TabIndex        =   173
            Top             =   1380
            Width           =   3360
         End
         Begin VB.Label Label111 
            Caption         =   "DOCU-APL"
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
            Left            =   8610
            TabIndex        =   172
            Top             =   945
            Width           =   1665
         End
         Begin VB.Label Label109 
            Caption         =   "Banco Local"
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
            Left            =   210
            TabIndex        =   171
            Top             =   960
            Width           =   3045
         End
         Begin VB.Label Label108 
            Caption         =   "BANC-CHE"
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
            Left            =   8610
            TabIndex        =   170
            Top             =   735
            Width           =   1665
         End
         Begin VB.Label Label106 
            Caption         =   "Banco Exterior"
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
            Left            =   210
            TabIndex        =   169
            Top             =   750
            Width           =   2940
         End
         Begin VB.Label Label123 
            Caption         =   "DET-PASE"
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
            Left            =   8610
            TabIndex        =   168
            Top             =   525
            Width           =   1665
         End
         Begin VB.Label Label121 
            Caption         =   "Destino"
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
            Left            =   210
            TabIndex        =   167
            Top             =   525
            Width           =   3465
         End
         Begin VB.Label Label120 
            Caption         =   "REF-ASIE"
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
            Left            =   8610
            TabIndex        =   166
            Top             =   315
            Width           =   1665
         End
         Begin VB.Label Varios 
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
            Height          =   225
            Index           =   0
            Left            =   3885
            TabIndex        =   165
            Top             =   315
            Width           =   3420
         End
         Begin VB.Label Label118 
            Caption         =   "Puerto de Embarque"
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
            Left            =   210
            TabIndex        =   164
            Top             =   330
            Width           =   3465
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Comerciales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1380
         Left            =   210
         TabIndex        =   151
         Top             =   9765
         Width           =   10620
         Begin VB.Label Label11 
            Caption         =   "Comerc(3)"
            Height          =   225
            Index           =   4
            Left            =   7560
            TabIndex        =   191
            Top             =   945
            Width           =   750
         End
         Begin VB.Label Comerc 
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
            Height          =   225
            Index           =   3
            Left            =   3885
            TabIndex        =   190
            Top             =   945
            Width           =   3420
         End
         Begin VB.Label Label96 
            Caption         =   "CTAC-CLI"
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
            Left            =   8610
            TabIndex        =   189
            Top             =   945
            Width           =   1665
         End
         Begin VB.Label Label94 
            Caption         =   "Contacto en Cliente"
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
            Left            =   105
            TabIndex        =   188
            Top             =   960
            Width           =   3360
         End
         Begin VB.Label Label11 
            Caption         =   "Comerc(2)"
            Height          =   225
            Index           =   2
            Left            =   7560
            TabIndex        =   162
            Top             =   735
            Width           =   750
         End
         Begin VB.Label Label11 
            Caption         =   "Comerc(1)"
            Height          =   225
            Index           =   1
            Left            =   7560
            TabIndex        =   161
            Top             =   525
            Width           =   750
         End
         Begin VB.Label Label11 
            Caption         =   "Comerc(0)"
            Height          =   225
            Index           =   0
            Left            =   7560
            TabIndex        =   160
            Top             =   315
            Width           =   750
         End
         Begin VB.Label Comerc 
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
            Height          =   225
            Index           =   2
            Left            =   3885
            TabIndex        =   159
            Top             =   735
            Width           =   3420
         End
         Begin VB.Label Comerc 
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
            Height          =   225
            Index           =   1
            Left            =   3885
            TabIndex        =   158
            Top             =   525
            Width           =   3420
         End
         Begin VB.Label Label5 
            Caption         =   "Lista de Precios"
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
            Left            =   105
            TabIndex        =   157
            Top             =   315
            Width           =   3570
         End
         Begin VB.Label Label18 
            Caption         =   "COND-PAG"
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
            Left            =   8610
            TabIndex        =   156
            Top             =   735
            Width           =   1665
         End
         Begin VB.Label Label16 
            Caption         =   "Condicion de Pago"
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
            Left            =   105
            TabIndex        =   155
            Top             =   750
            Width           =   3990
         End
         Begin VB.Label Label15 
            Caption         =   "VENDEDOR"
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
            Left            =   8610
            TabIndex        =   154
            Top             =   525
            Width           =   1665
         End
         Begin VB.Label Comerc 
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
            Height          =   225
            Index           =   0
            Left            =   3885
            TabIndex        =   153
            Top             =   315
            Width           =   3420
         End
         Begin VB.Label Label13 
            Caption         =   "Vendedor"
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
            Left            =   105
            TabIndex        =   152
            Top             =   540
            Width           =   3990
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Transporte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   210
         TabIndex        =   126
         Top             =   7980
         Width           =   10620
         Begin VB.Label Label2 
            Caption         =   "Trans(5)"
            Height          =   225
            Index           =   5
            Left            =   7560
            TabIndex        =   150
            Top             =   1365
            Width           =   645
         End
         Begin VB.Label Label2 
            Caption         =   "Trans(4)"
            Height          =   225
            Index           =   4
            Left            =   7560
            TabIndex        =   149
            Top             =   1155
            Width           =   645
         End
         Begin VB.Label Label2 
            Caption         =   "Trans(3)"
            Height          =   225
            Index           =   3
            Left            =   7560
            TabIndex        =   148
            Top             =   945
            Width           =   645
         End
         Begin VB.Label Label2 
            Caption         =   "Trans(2)"
            Height          =   225
            Index           =   2
            Left            =   7560
            TabIndex        =   147
            Top             =   735
            Width           =   645
         End
         Begin VB.Label Label2 
            Caption         =   "Trans(1)"
            Height          =   225
            Index           =   1
            Left            =   7560
            TabIndex        =   146
            Top             =   525
            Width           =   645
         End
         Begin VB.Label Label2 
            Caption         =   "Trans(0)"
            Height          =   225
            Index           =   0
            Left            =   7560
            TabIndex        =   145
            Top             =   315
            Width           =   645
         End
         Begin VB.Label Trans 
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
            Height          =   225
            Index           =   5
            Left            =   3885
            TabIndex        =   144
            Top             =   1365
            Width           =   3420
         End
         Begin VB.Label Trans 
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
            Height          =   225
            Index           =   4
            Left            =   3885
            TabIndex        =   143
            Top             =   1155
            Width           =   3420
         End
         Begin VB.Label Trans 
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
            Height          =   225
            Index           =   3
            Left            =   3885
            TabIndex        =   142
            Top             =   945
            Width           =   3420
         End
         Begin VB.Label Trans 
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
            Height          =   225
            Index           =   2
            Left            =   3885
            TabIndex        =   141
            Top             =   735
            Width           =   3420
         End
         Begin VB.Label Trans 
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
            Height          =   225
            Index           =   1
            Left            =   3885
            TabIndex        =   140
            Top             =   525
            Width           =   3420
         End
         Begin VB.Label Trans 
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
            Height          =   225
            Index           =   0
            Left            =   3885
            TabIndex        =   139
            Top             =   315
            Width           =   3420
         End
         Begin VB.Label Label105 
            Caption         =   "PES-COMP"
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
            Left            =   8610
            TabIndex        =   138
            Top             =   1365
            Width           =   1665
         End
         Begin VB.Label Label103 
            Caption         =   "Peso Bruto"
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
            Left            =   105
            TabIndex        =   137
            Top             =   1380
            Width           =   3990
         End
         Begin VB.Label Label102 
            Caption         =   "PES-CONJ"
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
            Left            =   8610
            TabIndex        =   136
            Top             =   1155
            Width           =   1665
         End
         Begin VB.Label Label100 
            Caption         =   "Peso Neto"
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
            Left            =   105
            TabIndex        =   135
            Top             =   1170
            Width           =   3990
         End
         Begin VB.Label Label99 
            Caption         =   "NUBULTOS"
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
            Left            =   8610
            TabIndex        =   134
            Top             =   945
            Width           =   1665
         End
         Begin VB.Label Label97 
            Caption         =   "Numero de Bultos"
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
            Left            =   105
            TabIndex        =   133
            Top             =   945
            Width           =   3990
         End
         Begin VB.Label Label63 
            Caption         =   "DOMI-TRA"
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
            Left            =   8610
            TabIndex        =   132
            Top             =   525
            Width           =   1665
         End
         Begin VB.Label Label61 
            Caption         =   "Domicilio Transportista"
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
            Left            =   105
            TabIndex        =   131
            Top             =   525
            Width           =   3990
         End
         Begin VB.Label Label69 
            Caption         =   "TRANSPOR"
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
            Left            =   8610
            TabIndex        =   130
            Top             =   315
            Width           =   1665
         End
         Begin VB.Label Label67 
            Caption         =   "Transportista"
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
            Left            =   105
            TabIndex        =   129
            Top             =   315
            Width           =   3990
         End
         Begin VB.Label Label66 
            Caption         =   "DESTINO"
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
            Left            =   8610
            TabIndex        =   128
            Top             =   735
            Width           =   1665
         End
         Begin VB.Label Label64 
            Caption         =   "Localidad"
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
            Left            =   105
            TabIndex        =   127
            Top             =   750
            Width           =   3990
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Importes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4740
         Left            =   210
         TabIndex        =   25
         Top             =   3150
         Width           =   10620
         Begin VB.Label Impot 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   19
            Left            =   3885
            TabIndex        =   113
            Top             =   4305
            Width           =   3480
         End
         Begin VB.Label Impot 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   18
            Left            =   3885
            TabIndex        =   112
            Top             =   4095
            Width           =   3480
         End
         Begin VB.Label Label164 
            Caption         =   "Impot(18)"
            Height          =   225
            Index           =   19
            Left            =   7560
            TabIndex        =   111
            Top             =   4095
            Width           =   855
         End
         Begin VB.Label Label164 
            Caption         =   "Impot(19)"
            Height          =   225
            Index           =   18
            Left            =   7560
            TabIndex        =   110
            Top             =   4305
            Width           =   855
         End
         Begin VB.Label Label93 
            Caption         =   "TICAMBIO"
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
            Left            =   8610
            TabIndex        =   109
            Top             =   4305
            Width           =   1665
         End
         Begin VB.Label Label91 
            Caption         =   "Tipo de Cambio"
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
            Left            =   105
            TabIndex        =   108
            Top             =   4320
            Width           =   3465
         End
         Begin VB.Label Label72 
            Caption         =   "MONEDEMI"
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
            Left            =   8610
            TabIndex        =   107
            Top             =   4095
            Width           =   1665
         End
         Begin VB.Label Label70 
            Caption         =   "Moneda de Emision"
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
            Left            =   105
            TabIndex        =   106
            Top             =   4110
            Width           =   3675
         End
         Begin VB.Label Label164 
            Caption         =   "Impot(17)"
            Height          =   225
            Index           =   17
            Left            =   7560
            TabIndex        =   105
            Top             =   3885
            Width           =   855
         End
         Begin VB.Label Impot 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   17
            Left            =   3885
            TabIndex        =   104
            Top             =   3885
            Width           =   3480
         End
         Begin VB.Label Label78 
            Caption         =   "SUB-TOTAL"
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
            Left            =   8610
            TabIndex        =   103
            Top             =   3255
            Width           =   1665
         End
         Begin VB.Label Label76 
            Caption         =   "Subtotal"
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
            Left            =   105
            TabIndex        =   102
            Top             =   3255
            Width           =   3675
         End
         Begin VB.Label Label181 
            Caption         =   "IMP-LET3"
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
            Left            =   8610
            TabIndex        =   93
            Top             =   3885
            Width           =   1665
         End
         Begin VB.Label Label180 
            Caption         =   "IMP-LET2"
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
            Left            =   8610
            TabIndex        =   92
            Top             =   3675
            Width           =   1665
         End
         Begin VB.Label Label179 
            Caption         =   "IMP-LET1"
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
            Left            =   8610
            TabIndex        =   91
            Top             =   3465
            Width           =   1665
         End
         Begin VB.Label Label178 
            Caption         =   "TOTA-DOL"
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
            Left            =   8610
            TabIndex        =   90
            Top             =   3045
            Width           =   1665
         End
         Begin VB.Label Label177 
            Caption         =   "IMP-TOTA"
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
            Left            =   8610
            TabIndex        =   89
            Top             =   2835
            Width           =   1665
         End
         Begin VB.Label Label176 
            Caption         =   "TIPO-VAL"
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
            Left            =   8610
            TabIndex        =   88
            Top             =   2625
            Width           =   1665
         End
         Begin VB.Label Label175 
            Caption         =   "PERC-IVA"
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
            Left            =   8610
            TabIndex        =   87
            Top             =   2415
            Width           =   1665
         End
         Begin VB.Label Label174 
            Caption         =   "PERC-BRU"
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
            Left            =   8610
            TabIndex        =   86
            Top             =   2205
            Width           =   1665
         End
         Begin VB.Label Label173 
            Caption         =   "IMP-RNI"
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
            Left            =   8610
            TabIndex        =   85
            Top             =   1995
            Width           =   1665
         End
         Begin VB.Label Label172 
            Caption         =   "IMP-IVA1"
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
            Left            =   8610
            TabIndex        =   84
            Top             =   1785
            Width           =   1665
         End
         Begin VB.Label Label171 
            Caption         =   "POR-IVA1"
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
            Left            =   8610
            TabIndex        =   83
            Top             =   1575
            Width           =   1665
         End
         Begin VB.Label Label170 
            Caption         =   "IMP-IVA"
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
            Left            =   8610
            TabIndex        =   82
            Top             =   1365
            Width           =   1665
         End
         Begin VB.Label Label169 
            Caption         =   "POR-IVA"
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
            Left            =   8610
            TabIndex        =   81
            Top             =   1155
            Width           =   1665
         End
         Begin VB.Label Label168 
            Caption         =   "IMPNEDES"
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
            Left            =   8610
            TabIndex        =   80
            Top             =   945
            Width           =   1665
         End
         Begin VB.Label Label167 
            Caption         =   "IMP-DESC"
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
            Left            =   8610
            TabIndex        =   79
            Top             =   735
            Width           =   1665
         End
         Begin VB.Label Label166 
            Caption         =   "POR-DESC"
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
            Left            =   8610
            TabIndex        =   78
            Top             =   525
            Width           =   1665
         End
         Begin VB.Label Label165 
            Caption         =   "IMP-NETO"
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
            Left            =   8610
            TabIndex        =   77
            Top             =   315
            Width           =   1665
         End
         Begin VB.Label Label164 
            Caption         =   "Impot(16)"
            Height          =   225
            Index           =   16
            Left            =   7560
            TabIndex        =   76
            Top             =   3675
            Width           =   855
         End
         Begin VB.Label Label164 
            Caption         =   "Impot(15)"
            Height          =   225
            Index           =   15
            Left            =   7560
            TabIndex        =   75
            Top             =   3465
            Width           =   855
         End
         Begin VB.Label Label164 
            Caption         =   "Impot(14)"
            Height          =   225
            Index           =   14
            Left            =   7560
            TabIndex        =   74
            Top             =   3255
            Width           =   855
         End
         Begin VB.Label Label164 
            Caption         =   "Impot(13)"
            Height          =   225
            Index           =   13
            Left            =   7560
            TabIndex        =   73
            Top             =   3045
            Width           =   855
         End
         Begin VB.Label Label164 
            Caption         =   "Impot(12)"
            Height          =   225
            Index           =   12
            Left            =   7560
            TabIndex        =   72
            Top             =   2835
            Width           =   855
         End
         Begin VB.Label Label164 
            Caption         =   "Impot(11)"
            Height          =   225
            Index           =   11
            Left            =   7560
            TabIndex        =   71
            Top             =   2625
            Width           =   855
         End
         Begin VB.Label Label164 
            Caption         =   "Impot(10)"
            Height          =   225
            Index           =   10
            Left            =   7560
            TabIndex        =   70
            Top             =   2415
            Width           =   855
         End
         Begin VB.Label Label164 
            Caption         =   "Impot(9)"
            Height          =   225
            Index           =   9
            Left            =   7560
            TabIndex        =   69
            Top             =   2205
            Width           =   855
         End
         Begin VB.Label Label164 
            Caption         =   "Impot(8)"
            Height          =   225
            Index           =   8
            Left            =   7560
            TabIndex        =   68
            Top             =   1995
            Width           =   855
         End
         Begin VB.Label Label164 
            Caption         =   "Impot(7)"
            Height          =   225
            Index           =   7
            Left            =   7560
            TabIndex        =   67
            Top             =   1785
            Width           =   855
         End
         Begin VB.Label Label164 
            Caption         =   "Impot(6)"
            Height          =   225
            Index           =   6
            Left            =   7560
            TabIndex        =   66
            Top             =   1575
            Width           =   855
         End
         Begin VB.Label Label164 
            Caption         =   "Impot(5)"
            Height          =   225
            Index           =   5
            Left            =   7560
            TabIndex        =   65
            Top             =   1365
            Width           =   855
         End
         Begin VB.Label Label164 
            Caption         =   "Impot(4)"
            Height          =   225
            Index           =   4
            Left            =   7560
            TabIndex        =   64
            Top             =   1155
            Width           =   855
         End
         Begin VB.Label Label164 
            Caption         =   "Impot(3)"
            Height          =   225
            Index           =   3
            Left            =   7560
            TabIndex        =   63
            Top             =   945
            Width           =   855
         End
         Begin VB.Label Label164 
            Caption         =   "Impot(2)"
            Height          =   225
            Index           =   2
            Left            =   7560
            TabIndex        =   62
            Top             =   735
            Width           =   855
         End
         Begin VB.Label Label164 
            Caption         =   "Impot(1)"
            Height          =   225
            Index           =   1
            Left            =   7560
            TabIndex        =   61
            Top             =   525
            Width           =   855
         End
         Begin VB.Label Label164 
            Caption         =   "Impot(0)"
            Height          =   225
            Index           =   0
            Left            =   7560
            TabIndex        =   60
            Top             =   315
            Width           =   855
         End
         Begin VB.Label Impot 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   16
            Left            =   3885
            TabIndex        =   59
            Top             =   3675
            Width           =   3480
         End
         Begin VB.Label Impot 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   15
            Left            =   3885
            TabIndex        =   58
            Top             =   3465
            Width           =   3480
         End
         Begin VB.Label Impot 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   14
            Left            =   3885
            TabIndex        =   57
            Top             =   3255
            Width           =   3480
         End
         Begin VB.Label Impot 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   13
            Left            =   3885
            TabIndex        =   56
            Top             =   3045
            Width           =   3480
         End
         Begin VB.Label Impot 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   12
            Left            =   3885
            TabIndex        =   55
            Top             =   2835
            Width           =   3480
         End
         Begin VB.Label Impot 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   11
            Left            =   3885
            TabIndex        =   54
            Top             =   2625
            Width           =   3480
         End
         Begin VB.Label Impot 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   10
            Left            =   3885
            TabIndex        =   53
            Top             =   2415
            Width           =   3480
         End
         Begin VB.Label Impot 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   9
            Left            =   3885
            TabIndex        =   52
            Top             =   2205
            Width           =   3480
         End
         Begin VB.Label Impot 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   8
            Left            =   3885
            TabIndex        =   51
            Top             =   1995
            Width           =   3480
         End
         Begin VB.Label Impot 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   7
            Left            =   3885
            TabIndex        =   50
            Top             =   1785
            Width           =   3480
         End
         Begin VB.Label Impot 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   6
            Left            =   3885
            TabIndex        =   49
            Top             =   1575
            Width           =   3480
         End
         Begin VB.Label Impot 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   5
            Left            =   3885
            TabIndex        =   48
            Top             =   1365
            Width           =   3480
         End
         Begin VB.Label Impot 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   4
            Left            =   3885
            TabIndex        =   47
            Top             =   1155
            Width           =   3480
         End
         Begin VB.Label Impot 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   3
            Left            =   3885
            TabIndex        =   46
            Top             =   945
            Width           =   3480
         End
         Begin VB.Label Impot 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   2
            Left            =   3885
            TabIndex        =   45
            Top             =   735
            Width           =   3480
         End
         Begin VB.Label Impot 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   1
            Left            =   3885
            TabIndex        =   44
            Top             =   525
            Width           =   3480
         End
         Begin VB.Label Impot 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   0
            Left            =   3885
            TabIndex        =   43
            Top             =   315
            Width           =   3480
         End
         Begin VB.Label Label163 
            Caption         =   "Importe en Letras 3"
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
            Left            =   105
            TabIndex        =   42
            Top             =   3885
            Width           =   3990
         End
         Begin VB.Label Label162 
            Caption         =   "Importe en Letras 2"
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
            Left            =   105
            TabIndex        =   41
            Top             =   3675
            Width           =   3990
         End
         Begin VB.Label Label161 
            Caption         =   "Importe en Letras 1"
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
            Left            =   105
            TabIndex        =   40
            Top             =   3465
            Width           =   3990
         End
         Begin VB.Label Label160 
            Caption         =   "Total en Dólares"
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
            Left            =   105
            TabIndex        =   39
            Top             =   3045
            Width           =   3990
         End
         Begin VB.Label Label159 
            Caption         =   "Importe Total"
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
            Left            =   105
            TabIndex        =   38
            Top             =   2835
            Width           =   3990
         End
         Begin VB.Label Label158 
            Caption         =   "Valor FOB"
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
            Left            =   105
            TabIndex        =   37
            Top             =   2625
            Width           =   3990
         End
         Begin VB.Label Label157 
            Caption         =   "Importe Percepción Iva"
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
            Left            =   105
            TabIndex        =   36
            Top             =   2415
            Width           =   3990
         End
         Begin VB.Label Label156 
            Caption         =   "Importe de Precepción Ingresos Brutos"
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
            Left            =   105
            TabIndex        =   35
            Top             =   2205
            Width           =   3990
         End
         Begin VB.Label Label155 
            Caption         =   "Importe de Iva RNI"
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
            Left            =   105
            TabIndex        =   34
            Top             =   1995
            Width           =   3990
         End
         Begin VB.Label Label154 
            Caption         =   "Importe de Iva - Tasa Secundaria"
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
            Left            =   105
            TabIndex        =   33
            Top             =   1785
            Width           =   3990
         End
         Begin VB.Label Label153 
            Caption         =   "Porcentaje de Iva - Tasa Secundaria"
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
            Left            =   105
            TabIndex        =   32
            Top             =   1575
            Width           =   3990
         End
         Begin VB.Label Label152 
            Caption         =   "Importe de Iva Tasa Básica"
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
            Left            =   105
            TabIndex        =   31
            Top             =   1365
            Width           =   3990
         End
         Begin VB.Label Label151 
            Caption         =   "Porcentaje de Iva Tasa Básica"
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
            Left            =   105
            TabIndex        =   30
            Top             =   1155
            Width           =   3990
         End
         Begin VB.Label Label148 
            Caption         =   "Importe Neto de Descuento"
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
            Left            =   105
            TabIndex        =   29
            Top             =   945
            Width           =   3990
         End
         Begin VB.Label Label145 
            Caption         =   "Importe de Descuento"
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
            Left            =   105
            TabIndex        =   28
            Top             =   735
            Width           =   3990
         End
         Begin VB.Label Label142 
            Caption         =   "Porcentaje de Descuento"
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
            Left            =   105
            TabIndex        =   27
            Top             =   525
            Width           =   3990
         End
         Begin VB.Label Label127 
            Caption         =   "Total Neto Factura"
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
            Left            =   105
            TabIndex        =   26
            Top             =   315
            Width           =   3990
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Identificación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2955
         Left            =   210
         TabIndex        =   1
         Top             =   105
         Width           =   10620
         Begin VB.Label Ident 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   11
            Left            =   3885
            TabIndex        =   125
            Top             =   2520
            Width           =   3480
         End
         Begin VB.Label Ident 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   10
            Left            =   3885
            TabIndex        =   124
            Top             =   2310
            Width           =   3480
         End
         Begin VB.Label Label124 
            Caption         =   "Ident(10)"
            Height          =   225
            Index           =   10
            Left            =   7560
            TabIndex        =   123
            Top             =   2520
            Width           =   1065
         End
         Begin VB.Label Label124 
            Caption         =   "Ident(9)"
            Height          =   225
            Index           =   9
            Left            =   7560
            TabIndex        =   122
            Top             =   2310
            Width           =   1065
         End
         Begin VB.Label Label87 
            Caption         =   "FECH-LAN"
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
            Left            =   8610
            TabIndex        =   121
            Top             =   2520
            Width           =   1665
         End
         Begin VB.Label Label85 
            Caption         =   "Cantidad dias para Vencimiento"
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
            Left            =   105
            TabIndex        =   120
            Top             =   2535
            Width           =   3990
         End
         Begin VB.Label Label84 
            Caption         =   "FECH-VEN"
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
            Left            =   8610
            TabIndex        =   119
            Top             =   2310
            Width           =   1665
         End
         Begin VB.Label Label82 
            Caption         =   "Fecha de Vencimiento"
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
            Left            =   105
            TabIndex        =   118
            Top             =   2325
            Width           =   3990
         End
         Begin VB.Label Label124 
            Caption         =   "Ident(8)"
            Height          =   225
            Index           =   8
            Left            =   7560
            TabIndex        =   117
            Top             =   2100
            Width           =   1065
         End
         Begin VB.Label Ident 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   8
            Left            =   3885
            TabIndex        =   116
            Top             =   2100
            Width           =   3480
         End
         Begin VB.Label Label51 
            Caption         =   "CIVA-CLI"
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
            Left            =   8610
            TabIndex        =   115
            Top             =   840
            Width           =   1665
         End
         Begin VB.Label Label49 
            Caption         =   "Letra"
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
            Left            =   105
            TabIndex        =   114
            Top             =   840
            Width           =   3360
         End
         Begin VB.Label Label124 
            Caption         =   "Ident(7)"
            Height          =   225
            Index           =   7
            Left            =   7560
            TabIndex        =   101
            Top             =   1890
            Width           =   1065
         End
         Begin VB.Label Label124 
            Caption         =   "Ident(6)"
            Height          =   225
            Index           =   6
            Left            =   7560
            TabIndex        =   100
            Top             =   1680
            Width           =   1065
         End
         Begin VB.Label Ident 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   7
            Left            =   3885
            TabIndex        =   99
            Top             =   1890
            Width           =   3480
         End
         Begin VB.Label Ident 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   6
            Left            =   3885
            TabIndex        =   98
            Top             =   1680
            Width           =   3480
         End
         Begin VB.Label Label9 
            Caption         =   "NRO-AENT"
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
            Left            =   8610
            TabIndex        =   97
            Top             =   2100
            Width           =   1665
         End
         Begin VB.Label Label7 
            Caption         =   "Numero de Pedido"
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
            Left            =   105
            TabIndex        =   96
            Top             =   2115
            Width           =   3990
         End
         Begin VB.Label Label6 
            Caption         =   "NRO-OCOM"
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
            Left            =   8610
            TabIndex        =   95
            Top             =   1890
            Width           =   1665
         End
         Begin VB.Label Label4 
            Caption         =   "Numero de Orden de Compra"
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
            Left            =   105
            TabIndex        =   94
            Top             =   1905
            Width           =   3990
         End
         Begin VB.Label Label124 
            Caption         =   "Ident(5)"
            Height          =   225
            Index           =   5
            Left            =   7560
            TabIndex        =   24
            Top             =   1470
            Width           =   1065
         End
         Begin VB.Label Label124 
            Caption         =   "Ident(4)"
            Height          =   225
            Index           =   4
            Left            =   7560
            TabIndex        =   23
            Top             =   1260
            Width           =   1065
         End
         Begin VB.Label Label124 
            Caption         =   "Ident(3)"
            Height          =   225
            Index           =   3
            Left            =   7560
            TabIndex        =   22
            Top             =   1050
            Width           =   1065
         End
         Begin VB.Label Label124 
            Caption         =   "Ident(2)"
            Height          =   225
            Index           =   2
            Left            =   7560
            TabIndex        =   21
            Top             =   840
            Width           =   1065
         End
         Begin VB.Label Label124 
            Caption         =   "Ident(1)"
            Height          =   225
            Index           =   1
            Left            =   7560
            TabIndex        =   20
            Top             =   630
            Width           =   1065
         End
         Begin VB.Label Label124 
            Caption         =   "Ident(0)"
            Height          =   225
            Index           =   0
            Left            =   7560
            TabIndex        =   19
            Top             =   420
            Width           =   1065
         End
         Begin VB.Label Ident 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   5
            Left            =   3885
            TabIndex        =   18
            Top             =   1470
            Width           =   3480
         End
         Begin VB.Label Label150 
            Caption         =   "Numero de Remito Corto"
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
            Left            =   105
            TabIndex        =   17
            Top             =   1695
            Width           =   3990
         End
         Begin VB.Label Ident 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   4
            Left            =   3885
            TabIndex        =   16
            Top             =   1260
            Width           =   3480
         End
         Begin VB.Label Label147 
            Caption         =   "Numero de Remito Largo"
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
            Left            =   105
            TabIndex        =   15
            Top             =   1485
            Width           =   3990
         End
         Begin VB.Label Label146 
            Caption         =   "NRO-REMI"
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
            Left            =   8610
            TabIndex        =   14
            Top             =   1470
            Width           =   1665
         End
         Begin VB.Label Ident 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   3
            Left            =   3885
            TabIndex        =   13
            Top             =   1050
            Width           =   3480
         End
         Begin VB.Label Label144 
            Caption         =   "Numero de Factura Corto"
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
            Left            =   105
            TabIndex        =   12
            Top             =   1275
            Width           =   3990
         End
         Begin VB.Label Ident 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   2
            Left            =   3885
            TabIndex        =   11
            Top             =   840
            Width           =   3480
         End
         Begin VB.Label Label141 
            Caption         =   "Numero de Factura Largo"
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
            Left            =   105
            TabIndex        =   10
            Top             =   1065
            Width           =   3990
         End
         Begin VB.Label Ident 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   1
            Left            =   3885
            TabIndex        =   9
            Top             =   630
            Width           =   3480
         End
         Begin VB.Label Label126 
            Caption         =   "Fecha de Emisión"
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
            Left            =   105
            TabIndex        =   8
            Top             =   645
            Width           =   3990
         End
         Begin VB.Label Ident 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   0
            Left            =   3885
            TabIndex        =   3
            Top             =   420
            Width           =   3480
         End
         Begin VB.Label Label1 
            Caption         =   "Numero de Cliente"
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
            Left            =   105
            TabIndex        =   2
            Top             =   435
            Width           =   3990
         End
      End
   End
   Begin VB.Label Label137 
      Caption         =   "NRO-REMI"
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
      Left            =   7665
      TabIndex        =   7
      Top             =   0
      Width           =   1665
   End
   Begin VB.Label Label134 
      Caption         =   "NRO-REMI"
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
      Left            =   7665
      TabIndex        =   6
      Top             =   0
      Width           =   1665
   End
   Begin VB.Label Label131 
      Caption         =   "NRO-REMI"
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
      Left            =   7665
      TabIndex        =   5
      Top             =   0
      Width           =   1665
   End
   Begin VB.Label Label128 
      Caption         =   "NRO-REMI"
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
      Left            =   7665
      TabIndex        =   4
      Top             =   0
      Width           =   1665
   End
End
Attribute VB_Name = "FORCONFA04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'\\\OT-06-0405-MH-20/09/06///
Private Sub Form_Load()
      '
'      Label50.Caption = FORFACCS.Label13.Caption 'Cond Iva
'      '
'      REMI$ = "RT." + TRIM$(FORFACCS.Text10.TEXT)
'      NroRemi$ = TRIM$(Mid$(REMI$, 4))
'      While Len(NroRemi$) < 8: NroRemi$ = "0" + NroRemi$: Wend
'      CMO$ = "RT"
'      PREFIREM$ = CONTROL$("", "PREFIREM")
'      If PREFIREM$ = "" Then PREFIREM$ = "0001"
'      NroRemi$ = CMO$ + "." + PREFIREM$ + "-" + NroRemi$
'      '
'      'Label2.Caption = NroRemi$                   'nº remito
'      'Label5.Caption = NUOCOMPRA$                 'nº Orden de Compra
'      'Label8.Caption = NROPEDIDO$                 'nº Pedido
'      IMPO(0).Caption = FORFACCS.Total(0).TEXT    'Total Neto
'      'Label14.Caption = FORFACCS.Eco21(2).TEXT    'Vendedor
'      Label17.Caption = FORFACCS.Eco21(1).TEXT    'Cond Pago
'      IMPO(4).Caption = FORFACCS.PORIVA(0).Caption 'Porc IVA Basica
'      IMPO(5).Caption = FORFACCS.Total(3).TEXT    'Impo IVA Basica
'      IMPO(6).Caption = FORFACCS.PORIVA(1).Caption 'Porc IVA Secun
'      IMPO(7).Caption = FORFACCS.Total(4).TEXT    'Impo IVA Secun
'      IMPO(2).Caption = FORFACCS.Total(1).TEXT    'Importe dto
'      IMPO(3).Caption = FORFACCS.Total(2).TEXT  'Importe Neto dto
'      IMPO(12).Caption = FORCACCS.Total(6).TEXT  'Importe Total
'      IMPO(1).Caption = TRIM$(FORFACCS.PORDESCU.TEXT) 'Porc Dto
'      IMPO(8).Caption = FORFACCS.Total(5).TEXT 'Importe iva RNI
'      '
'      TTTT# = XVALO(FORFACCS.Total(6).TEXT)
'      TOTFAC$ = NUMTEX$(TTTT#)
'      Call FRATEX(TOTFAC$, NRX1$, NRX2$, NRX3$, LOIMPLET%)
'      '
'      IMPO(14).Caption = NRX1$  'Importe en letras 1
'      IMPO(15).Caption = NRX2$  'Importe en letras 2
'      IMPO(16).Caption = NRX3$  'Importe en letras 3
'      '
'      IMPO(13).Caption = ""  'Total en dolares
'      Label62.Caption = FORFACCS.LENTREGA(0).TEXT  'Domic transportista
'      Label65.Caption = FORFACCS.LENTREGA(1).TEXT  'Localidad
'      Label68.Caption = FORFACCS.LENTREGA(2).TEXT  'Transportista
'      Label71.Caption = FORFACCS.MONEMIS.Caption   'Moneda emision
'      IMPO(9).Caption = FORFACCS.Text15.TEXT  'Importe percep IB
'      Label77.Caption = ""  ' Subtotal
'      IMPO(10).Caption = FORFACCS.Text17.TEXT 'Importe percep IVA
'      Label83.Caption = Format$(FVENCIM, "dd/mm/yy") 'Fecha vencim
'      On Error Resume Next
'      FECHADOC$ = FORFACCS.Text9.TEXT
'      FEXO = FECHANUM(FECHDOC$)
'      FEXU = CVINT(FVENCIM)
'      CADIVEN = DIENTRE(FEXO, FEXU)
'      On Error GoTo 0
'      Label86.Caption = FORMATNUM$(CADIVEN, "F5.0") 'cant dias para venc
'      IMPO(11).Caption = TRIM(FORFACCS.Text18) ' Valor FOB
'      Label92.Caption = FORMATNUM$(TICAMONE(2), "F7.4") ' Tipo cambio
'      '
'      DESTIDOC% = XVALO(FORFACCS.Text1.TEXT)
'      Label95.Caption = CONTACTCLI$(DESTIDOC%) 'Contacto en cliente
'      Label98.Caption = TRIM$(FINFACTURA.Text2(2)) 'Num bultos
'      Label101.Caption = TRIM$(FINFACTURA.Text2(11)) + " Kg" 'Peso neto
'      Label104.Caption = TRIM$(FINFACTURA.Text2(4)) + " Kg" 'Peso bruto
'      Label107.Caption = TRIM$(FINFACTURA.Text2(5)) 'Banco Exterior
'      Label110.Caption = TRIM$(FINFACTURA.Text2(8)) 'Banco Local
'      Label113.Caption = TRIM$(FINFACTURA.Text2(10)) 'Credito
'      Label116.Caption = TRIM$(FINFACTURA.Text2(9)) 'Guia Importacion
'      Label119.Caption = TRIM$(FINFACTURA.Text2(6)) 'Puerto Embarque
'      Label122.Caption = TRIM$(FINFACTURA.Text2(7)) 'Destino
'      '
End Sub
'\\\OT-06-0405-MH-FIN///

