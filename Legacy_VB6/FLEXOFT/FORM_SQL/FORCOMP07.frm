VERSION 5.00
Begin VB.Form FORCOMP07 
   BackColor       =   &H00C0C0A0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ORDENES DE COMPRAS - PROVEEDORES EXTERNOS"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11805
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FORCOMP07.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   11805
   Begin VB.PictureBox Picture999 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   0
      Left            =   240
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   69
      Top             =   2520
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
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
      Height          =   8520
      Left            =   0
      TabIndex        =   10
      Top             =   -60
      Width           =   12000
      Begin VB.Frame Frame2 
         BackColor       =   &H00CDDADA&
         Caption         =   "DESTINO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2400
         Left            =   120
         TabIndex        =   111
         Top             =   6600
         Visible         =   0   'False
         Width           =   9525
         Begin VB.TextBox TEXT25 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7580
            MaxLength       =   16
            TabIndex        =   121
            Text            =   " "
            Top             =   795
            Width           =   2850
         End
         Begin VB.Frame Frame16 
            BackColor       =   &H00CDDADA&
            Height          =   510
            Left            =   600
            TabIndex        =   118
            Top             =   240
            Width           =   9265
            Begin VB.OptionButton Option1 
               Appearance      =   0  'Flat
               BackColor       =   &H00CDDADA&
               Caption         =   "Pedido de Reposición de Stocks"
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
               Height          =   255
               Left            =   480
               TabIndex        =   120
               Top             =   200
               Width           =   3255
            End
            Begin VB.OptionButton Option2 
               Appearance      =   0  'Flat
               BackColor       =   &H00CDDADA&
               Caption         =   "Solicitud por Pedido de Cliente"
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
               Height          =   255
               Left            =   6120
               TabIndex        =   119
               Top             =   200
               Value           =   -1  'True
               Width           =   2895
            End
         End
         Begin VB.CommandButton Command25 
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
            Left            =   2040
            TabIndex        =   115
            Top             =   795
            Width           =   180
         End
         Begin VB.TextBox TEXT24 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1155
            TabIndex        =   114
            Text            =   " "
            Top             =   795
            Width           =   915
         End
         Begin VB.TextBox Eco24 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   2400
            TabIndex        =   113
            Top             =   795
            Width           =   4095
         End
         Begin VB.TextBox Text22 
            Appearance      =   0  'Flat
            Height          =   1110
            Left            =   1155
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   112
            Top             =   1185
            Width           =   9285
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "O/C Cliente"
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
            Index           =   6
            Left            =   6400
            TabIndex        =   122
            Top             =   880
            Width           =   1110
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente"
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
            Index           =   9
            Left            =   330
            TabIndex        =   117
            Top             =   880
            Width           =   735
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Observa- ciones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Index           =   7
            Left            =   0
            TabIndex        =   116
            Top             =   1150
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00CDDADA&
         Caption         =   "CONDICIONES COMERCIALES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2300
         Left            =   180
         TabIndex        =   93
         Top             =   3840
         Width           =   6165
         Begin VB.TextBox Text12 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1155
            TabIndex        =   99
            ToolTipText     =   "Doble Clic Ampliar Casillero"
            Top             =   1320
            Width           =   4845
         End
         Begin VB.TextBox Text13 
            Appearance      =   0  'Flat
            Height          =   560
            Left            =   1155
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   100
            Top             =   1680
            Width           =   4845
         End
         Begin VB.TextBox Eco21 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Index           =   1
            Left            =   1890
            TabIndex        =   102
            Text            =   " "
            Top             =   600
            Width           =   4110
         End
         Begin VB.TextBox TEXT21 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   1155
            TabIndex        =   96
            Text            =   " "
            Top             =   600
            Width           =   435
         End
         Begin VB.CommandButton Command21 
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
            Left            =   1575
            TabIndex        =   101
            Top             =   600
            Width           =   180
         End
         Begin VB.TextBox Eco21 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Index           =   0
            Left            =   1890
            TabIndex        =   97
            Top             =   250
            Width           =   4110
         End
         Begin VB.TextBox TEXT21 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   1155
            TabIndex        =   95
            Text            =   " "
            Top             =   250
            Width           =   435
         End
         Begin VB.CommandButton Command21 
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
            Left            =   1560
            TabIndex        =   94
            Top             =   250
            Width           =   180
         End
         Begin VB.TextBox Text10 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1155
            TabIndex        =   98
            ToolTipText     =   "Doble Clic Ampliar Casillero"
            Top             =   960
            Width           =   4845
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Solicita"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   240
            TabIndex        =   107
            Top             =   1410
            Width           =   825
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Observa- ciones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Index           =   3
            Left            =   0
            TabIndex        =   106
            Top             =   1640
            Width           =   1095
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Embalaje"
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
            Index           =   1
            Left            =   -3045
            TabIndex        =   105
            Top             =   680
            Width           =   4110
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cond. Pago"
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
            Index           =   0
            Left            =   -525
            TabIndex        =   104
            Top             =   350
            Width           =   1590
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Lugar Entrega"
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
            Index           =   2
            Left            =   -3045
            TabIndex        =   103
            Top             =   1050
            Width           =   4110
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00CDDADA&
         Caption         =   "RESUMEN FINAL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2300
         Left            =   6420
         TabIndex        =   70
         Top             =   3840
         Width           =   4335
         Begin VB.TextBox TXTPESO 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   3165
            Locked          =   -1  'True
            TabIndex        =   109
            Top             =   1920
            Width           =   1020
         End
         Begin VB.TextBox Total 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Index           =   3
            Left            =   2685
            TabIndex        =   78
            Top             =   1575
            Width           =   1500
         End
         Begin VB.TextBox Total 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Index           =   2
            Left            =   2685
            TabIndex        =   77
            Top             =   945
            Width           =   1500
         End
         Begin VB.TextBox XText12 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            MaxLength       =   14
            TabIndex        =   76
            Top             =   1440
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox Total 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Index           =   0
            Left            =   2685
            TabIndex        =   75
            Top             =   315
            Width           =   1500
         End
         Begin VB.TextBox Total 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Index           =   1
            Left            =   2685
            TabIndex        =   74
            TabStop         =   0   'False
            Text            =   " "
            Top             =   630
            Width           =   1500
         End
         Begin VB.TextBox PORDESCU 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1725
            MaxLength       =   14
            TabIndex        =   73
            Top             =   630
            Width           =   750
         End
         Begin VB.TextBox Total 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Index           =   4
            Left            =   2685
            TabIndex        =   72
            Top             =   1260
            Width           =   1500
         End
         Begin VB.TextBox Text17 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1830
            MaxLength       =   14
            TabIndex        =   71
            Top             =   1260
            Width           =   645
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Kilos"
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
            Index           =   8
            Left            =   2040
            TabIndex        =   110
            ToolTipText     =   "Haga un Doble-Click para Emitir CON / SIN Precios"
            Top             =   1990
            Width           =   1035
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Pesos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   80
            TabIndex        =   92
            ToolTipText     =   "T.C.: 1.0000 - Doble-Click para Cambiar Moneda"
            Top             =   1560
            Width           =   1275
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "%:"
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
            Left            =   2520
            TabIndex        =   91
            Top             =   1020
            Width           =   225
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total O/Compra"
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
            Index           =   4
            Left            =   600
            TabIndex        =   90
            ToolTipText     =   "Haga un Doble-Click para Emitir CON / SIN Precios"
            Top             =   1650
            Width           =   1995
         End
         Begin VB.Label Ltota 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "IVA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   1080
            TabIndex        =   89
            ToolTipText     =   "Doble Click para Editar Detalle"
            Top             =   945
            Width           =   645
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Subtotal"
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
            Index           =   0
            Left            =   1635
            TabIndex        =   88
            Top             =   400
            Width           =   960
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Descuento"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   900
            TabIndex        =   87
            Top             =   705
            Width           =   780
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "%:"
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
            Left            =   2505
            TabIndex        =   86
            Top             =   720
            Width           =   225
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Percepcion"
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
            Left            =   -1180
            TabIndex        =   85
            Top             =   1320
            Width           =   2955
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "%:"
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
            Left            =   2505
            TabIndex        =   84
            Top             =   1320
            Width           =   225
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "T.Cambio"
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
            Index           =   5
            Left            =   45
            TabIndex        =   83
            Top             =   1965
            Width           =   690
         End
         Begin VB.Label TICAMBIS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   800
            TabIndex        =   82
            Top             =   1920
            Width           =   1005
         End
         Begin VB.Label Label118 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   1770
            TabIndex        =   81
            Top             =   945
            Width           =   705
         End
         Begin VB.Label Label18 
            Height          =   255
            Left            =   75
            TabIndex        =   80
            Top             =   1155
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "C. Imp:"
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
            Index           =   7
            Left            =   -60
            TabIndex        =   79
            Top             =   930
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00CDDADA&
         Caption         =   "ENTREGA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1005
         Left            =   7440
         TabIndex        =   18
         Top             =   480
         Width           =   1050
         Begin VB.TextBox Text19 
            Appearance      =   0  'Flat
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
            Left            =   600
            TabIndex        =   62
            Top             =   240
            Width           =   400
         End
         Begin VB.CommandButton Command7 
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
            Left            =   840
            TabIndex        =   20
            Top             =   600
            Width           =   160
         End
         Begin VB.TextBox Text14 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   120
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   600
            Width           =   720
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Dias"
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
            Left            =   0
            TabIndex        =   63
            Top             =   320
            Width           =   435
         End
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   9450
         TabIndex        =   64
         Top             =   210
         Width           =   810
      End
      Begin VB.TextBox Text18 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5955
         TabIndex        =   60
         Top             =   540
         Width           =   1340
      End
      Begin VB.CommandButton Command15 
         Caption         =   "T.Lib"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   6645
         TabIndex        =   59
         Top             =   905
         Width           =   645
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   5640
         TabIndex        =   58
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1180
         Width           =   1650
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5955
         TabIndex        =   56
         Top             =   225
         Width           =   1170
      End
      Begin VB.CommandButton Command18 
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
         Left            =   7110
         TabIndex        =   55
         Top             =   225
         Width           =   180
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   9000
         Left            =   10965
         ScaleHeight     =   8970
         ScaleWidth      =   1035
         TabIndex        =   43
         Top             =   0
         Width           =   1065
         Begin VB.PictureBox MARCOBARRA 
            Appearance      =   0  'Flat
            BackColor       =   &H002B9973&
            ForeColor       =   &H80000008&
            Height          =   200
            Left            =   120
            ScaleHeight     =   165
            ScaleWidth      =   615
            TabIndex        =   140
            Top             =   7680
            Width           =   650
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H0048CC9D&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   0
               TabIndex        =   141
               Top             =   0
               Width           =   12000
            End
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Ver"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   105
            Picture         =   "FORCOMP07.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Ver por Pantalla y Re-Imprimir"
            Top             =   3105
            Width           =   650
         End
         Begin VB.CommandButton Command6 
            BackColor       =   &H00E0E0E0&
            Caption         =   "List"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   105
            Picture         =   "FORCOMP07.frx":074C
            Style           =   1  'Graphical
            TabIndex        =   52
            ToolTipText     =   "Consultas y Listados"
            Top             =   3795
            Width           =   650
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   640
            Left            =   105
            Picture         =   "FORCOMP07.frx":0A56
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "Ayuda FLEXOFT en Línea"
            Top             =   790
            Width           =   650
         End
         Begin VB.CommandButton Command13 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Option"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   105
            Picture         =   "FORCOMP07.frx":0D60
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Opciones Varias - Modificaciones y Consultas"
            Top             =   5935
            Width           =   650
         End
         Begin VB.CommandButton Command11 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Files"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   105
            Picture         =   "FORCOMP07.frx":0EAA
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Acceso a Base de Datos Maestras"
            Top             =   1550
            Width           =   650
         End
         Begin VB.CommandButton Command8 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Direct"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   105
            Picture         =   "FORCOMP07.frx":11B4
            Style           =   1  'Graphical
            TabIndex        =   48
            ToolTipText     =   "Accesos Directos a Otras Aplicaciones"
            Top             =   5250
            Width           =   650
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Setup"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   105
            Picture         =   "FORCOMP07.frx":14BE
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Configuración de Funcionamiento de Compras"
            Top             =   2310
            Width           =   650
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Exit"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   105
            Picture         =   "FORCOMP07.frx":1900
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Cierra y Abandona la Aplicación"
            Top             =   100
            Width           =   650
         End
         Begin VB.CommandButton Command22 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Check"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   105
            Picture         =   "FORCOMP07.frx":1A4A
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Aprobar Ordenes de Compra"
            Top             =   4485
            Width           =   650
         End
         Begin VB.CommandButton Command14 
            BackColor       =   &H00E0E0E0&
            Caption         =   "O.K."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   105
            Picture         =   "FORCOMP07.frx":1D54
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Aprueba, Graba e Imprime Orden de Compra"
            Top             =   6720
            Width           =   650
         End
      End
      Begin VB.CommandButton Command10 
         Caption         =   "!"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1155
         TabIndex        =   42
         ToolTipText     =   "Lista de Proveedores con Entregas Programadas"
         Top             =   870
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Frame Frame15 
         Caption         =   "Frame15"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1830
         Left            =   10450
         TabIndex        =   41
         Top             =   2010
         Width           =   15
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00CDDADA&
         Caption         =   "PENDIENTE DE RECEPCION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1845
         Left            =   160
         TabIndex        =   33
         Top             =   6200
         Width           =   10580
         Begin VB.Frame Frame14 
            Caption         =   "Frame14"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1550
            Left            =   9345
            TabIndex        =   40
            Top             =   210
            Width           =   15
         End
         Begin VB.Frame Frame13 
            Caption         =   "Frame13"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1550
            Left            =   8715
            TabIndex        =   39
            Top             =   210
            Width           =   15
         End
         Begin VB.Frame Frame12 
            Caption         =   "Frame12"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1550
            Left            =   7665
            TabIndex        =   38
            Top             =   210
            Width           =   15
         End
         Begin VB.Frame Frame11 
            Caption         =   "Frame11"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1550
            Left            =   5985
            TabIndex        =   37
            Top             =   210
            Width           =   15
         End
         Begin VB.Frame Frame10 
            Caption         =   "Frame10"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1550
            Left            =   1890
            TabIndex        =   36
            Top             =   210
            Width           =   15
         End
         Begin VB.ListBox List2 
            Appearance      =   0  'Flat
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
            Height          =   1155
            Left            =   30
            MultiSelect     =   1  'Simple
            TabIndex        =   35
            Top             =   600
            Width           =   10500
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   435
            Left            =   30
            ScaleHeight     =   405
            ScaleWidth      =   10470
            TabIndex        =   34
            Top             =   200
            Width           =   10500
            Begin VB.Label Label34 
               BackStyle       =   0  'Transparent
               Caption         =   "O/Compra"
               Height          =   255
               Left            =   9360
               TabIndex        =   139
               Top             =   120
               Width           =   975
            End
            Begin VB.Label Label32 
               BackStyle       =   0  'Transparent
               Caption         =   "Atraso"
               Height          =   255
               Left            =   8700
               TabIndex        =   138
               Top             =   120
               Width           =   975
            End
            Begin VB.Label Label31 
               BackStyle       =   0  'Transparent
               Caption         =   "F. Entrega"
               Height          =   255
               Left            =   7740
               TabIndex        =   137
               Top             =   120
               Width           =   975
            End
            Begin VB.Label Label30 
               BackStyle       =   0  'Transparent
               Caption         =   "Cantidad"
               Height          =   255
               Left            =   6600
               TabIndex        =   136
               Top             =   120
               Width           =   975
            End
            Begin VB.Label Label29 
               BackStyle       =   0  'Transparent
               Caption         =   "Descripción"
               Height          =   255
               Left            =   1920
               TabIndex        =   135
               Top             =   120
               Width           =   1215
            End
            Begin VB.Label Label22 
               BackStyle       =   0  'Transparent
               Caption         =   "Código"
               Height          =   255
               Left            =   80
               TabIndex        =   134
               Top             =   120
               Width           =   975
            End
         End
         Begin VB.Line Line1 
            X1              =   0
            X2              =   10560
            Y1              =   1940
            Y2              =   1940
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1830
         Index           =   7
         Left            =   6250
         TabIndex        =   32
         Top             =   2010
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1830
         Index           =   6
         Left            =   8665
         TabIndex        =   31
         Top             =   2010
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1830
         Index           =   5
         Left            =   7465
         TabIndex        =   30
         Top             =   2010
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1830
         Index           =   4
         Left            =   9400
         TabIndex        =   29
         Top             =   2010
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1830
         Index           =   3
         Left            =   5725
         TabIndex        =   28
         Top             =   2010
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1830
         Index           =   2
         Left            =   4675
         TabIndex        =   27
         Top             =   2010
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1830
         Index           =   1
         Left            =   4100
         TabIndex        =   26
         Top             =   2010
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1830
         Index           =   0
         Left            =   2050
         TabIndex        =   25
         Top             =   2010
         Width           =   15
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00CDDADA&
         Caption         =   "VALORIZAR EN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   8520
         TabIndex        =   21
         Top             =   870
         Width           =   2220
         Begin VB.TextBox Text16 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   720
            TabIndex        =   24
            TabStop         =   0   'False
            Text            =   " "
            Top             =   290
            Width           =   1395
         End
         Begin VB.TextBox Text15 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   90
            MaxLength       =   2
            TabIndex        =   23
            TabStop         =   0   'False
            Text            =   " "
            Top             =   290
            Width           =   435
         End
         Begin VB.CommandButton Command12 
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
            Left            =   510
            TabIndex        =   22
            Top             =   280
            Width           =   150
         End
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2310
         MaxLength       =   30
         TabIndex        =   17
         Text            =   " "
         Top             =   1180
         Width           =   2985
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1830
         Left            =   160
         TabIndex        =   16
         Top             =   1995
         Width           =   10565
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   0
         Top             =   1080
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00404040&
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
         Height          =   6950
         Left            =   11025
         TabIndex        =   15
         Top             =   105
         Width           =   850
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   160
         ScaleHeight     =   405
         ScaleWidth      =   10530
         TabIndex        =   14
         Top             =   1570
         Width           =   10565
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "F.Entrega"
            Height          =   375
            Left            =   9360
            TabIndex        =   133
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "% Dto"
            Height          =   375
            Left            =   8700
            TabIndex        =   132
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "P.Unitario"
            Height          =   375
            Left            =   7570
            TabIndex        =   131
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad"
            Height          =   375
            Left            =   6490
            TabIndex        =   130
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "U/C"
            Height          =   375
            Left            =   5640
            TabIndex        =   129
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad"
            Height          =   255
            Left            =   4450
            TabIndex        =   128
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "U/S"
            Height          =   255
            Left            =   3980
            TabIndex        =   126
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción"
            Height          =   255
            Left            =   1960
            TabIndex        =   125
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label76 
            BackStyle       =   0  'Transparent
            Caption         =   "Código"
            Height          =   255
            Left            =   80
            TabIndex        =   124
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   9615
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " "
         Top             =   525
         Width           =   960
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
         Height          =   285
         Left            =   10560
         TabIndex        =   9
         Top             =   525
         Width           =   195
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   2310
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   " "
         Top             =   870
         Width           =   2985
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1470
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   " "
         Top             =   870
         Width           =   750
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1470
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   " "
         Top             =   540
         Width           =   3825
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1470
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   " "
         Top             =   210
         Width           =   3825
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
         Left            =   920
         TabIndex        =   3
         Top             =   870
         Width           =   195
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   160
         MaxLength       =   5
         TabIndex        =   0
         Text            =   " "
         Top             =   870
         Width           =   750
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Orden de Compra"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   585
         Left            =   0
         TabIndex        =   123
         Top             =   120
         Width           =   1500
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Contacto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   1060
         TabIndex        =   108
         Top             =   1280
         Width           =   1185
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "PQ"
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
         Left            =   4750
         TabIndex        =   2
         Top             =   1275
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5040
         TabIndex        =   1
         Top             =   1180
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Número"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   8860
         TabIndex        =   66
         Top             =   285
         Width           =   570
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   10440
         TabIndex        =   65
         ToolTipText     =   "Número de Revisión"
         Top             =   210
         Width           =   315
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   10290
         X2              =   10395
         Y1              =   345
         Y2              =   345
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Sector"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   5355
         TabIndex        =   61
         ToolTipText     =   "Dobe-Click para Asignar Sector por Item"
         Top             =   630
         Width           =   465
      End
      Begin VB.Label Label14 
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
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   5355
         TabIndex        =   57
         ToolTipText     =   "Dobe-Click para Asignar Destino por Item"
         Top             =   300
         Width           =   615
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   165
         TabIndex        =   54
         Top             =   1185
         Width           =   1170
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   12000
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "C.U.I.T."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5625
         TabIndex        =   13
         Top             =   945
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   8880
         TabIndex        =   12
         Top             =   580
         Width           =   690
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   160
         TabIndex        =   11
         Top             =   660
         Width           =   1200
      End
   End
   Begin VB.PictureBox SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   1530
      TabIndex        =   68
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.PictureBox Skin1 
      Height          =   480
      Left            =   3600
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   142
      Top             =   1200
      Width           =   1200
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      Height          =   255
      Left            =   0
      TabIndex        =   127
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Ltota 
      Alignment       =   1  'Right Justify
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
      Index           =   6
      Left            =   0
      TabIndex        =   67
      Top             =   0
      Width           =   330
   End
   Begin VB.Menu Archima 
      Caption         =   "Archivos Maestros"
      Begin VB.Menu Amastock 
         Caption         =   "Articulos de Stock"
      End
      Begin VB.Menu Provedor 
         Caption         =   "Proveedores y Contratistas"
      End
      Begin VB.Menu plh3 
         Caption         =   "-"
      End
      Begin VB.Menu Nuorig 
         Caption         =   "Números Originales"
         Enabled         =   0   'False
      End
      Begin VB.Menu plh4 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Terminar"
      End
   End
   Begin VB.Menu Orcom 
      Caption         =   "Ordenes de Compra"
      Begin VB.Menu MnuNueOCAbi 
         Caption         =   "Ordenes de Compra Abiertas"
      End
      Begin VB.Menu plh41 
         Caption         =   "-"
      End
      Begin VB.Menu Nue_Orco 
         Caption         =   "Nueva Orden de Compra"
         Begin VB.Menu Directa 
            Caption         =   "O/Compra Directa"
         End
         Begin VB.Menu Por_MRP 
            Caption         =   "O/Compra por MRP"
         End
      End
      Begin VB.Menu Modi_Ocom 
         Caption         =   "Modificar O/Compra"
      End
      Begin VB.Menu phl1 
         Caption         =   "-"
      End
      Begin VB.Menu Ver_Ocom 
         Caption         =   "Ver y Re-Imprimir"
      End
      Begin VB.Menu plh2 
         Caption         =   "-"
      End
      Begin VB.Menu MarItCum 
         Caption         =   "Marcar Items Cumplidos"
      End
      Begin VB.Menu plh22 
         Caption         =   "-"
      End
      Begin VB.Menu Anul_Ocom 
         Caption         =   "Anular O/Compra"
      End
      Begin VB.Menu Desanuoc 
         Caption         =   "Revertir Anulación"
      End
      Begin VB.Menu AnulacionMasivaPorTolerancia 
         Caption         =   "Anulacion Masiva por Dias de Tolerancia"
      End
      Begin VB.Menu plh6 
         Caption         =   "-"
      End
      Begin VB.Menu Maconsi1 
         Caption         =   "Materiales Consignados"
         Begin VB.Menu Entrecon 
            Caption         =   "Entregar Materiales"
         End
         Begin VB.Menu DevoConsi 
            Caption         =   "Devoluciones"
            Enabled         =   0   'False
         End
      End
   End
   Begin VB.Menu ConLis 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu CoCoPen 
         Caption         =   "Control de Compras Pendientes"
         Begin VB.Menu ConuOC 
            Caption         =   "Por Numero de O/C"
            Begin VB.Menu Screen1 
               Caption         =   "Por Pantalla"
            End
            Begin VB.Menu Print1 
               Caption         =   "Imprimir / Exportar"
            End
         End
         Begin VB.Menu CoProv 
            Caption         =   "Por Proveedor"
            Begin VB.Menu Screen2 
               Caption         =   "Por Pantalla"
            End
            Begin VB.Menu Print2 
               Caption         =   "Imprimir / Exportar"
            End
         End
         Begin VB.Menu CoMat 
            Caption         =   "Por Material"
            Begin VB.Menu Screen3 
               Caption         =   "Por Pantalla"
            End
            Begin VB.Menu Print3 
               Caption         =   "Imprimir / Exportar"
            End
         End
         Begin VB.Menu plh7 
            Caption         =   "-"
         End
         Begin VB.Menu Maconsi 
            Caption         =   "Material en Consignación"
            Begin VB.Menu Screen4 
               Caption         =   "Por Pantalla"
            End
            Begin VB.Menu Print4 
               Caption         =   "Imprimir / Exportar"
            End
         End
      End
      Begin VB.Menu CoUlCom 
         Caption         =   "Control Ultimas Compras"
         Begin VB.Menu UcoPro 
            Caption         =   "Control por Proveedor"
         End
         Begin VB.Menu UcoMat 
            Caption         =   "Control por Material"
         End
      End
      Begin VB.Menu MnuEstad 
         Caption         =   "Estadística de Costos"
      End
      Begin VB.Menu repocliente 
         Caption         =   "Reporte Pedidos de Reposición x Cliente"
         Visible         =   0   'False
      End
      Begin VB.Menu pedreppendiente 
         Caption         =   "Reporte Pedidos de Reposición Pendientes de Entrega"
         Visible         =   0   'False
      End
      Begin VB.Menu ocompendaprobastocksypedidos 
         Caption         =   "Pedidos de Reposición No Aprobados x Valida Stocks "
         Visible         =   0   'False
      End
      Begin VB.Menu MnuOCVSAV 
         Caption         =   "Estados de Pedidos de Reposición Vs Notas de Venta Avellaneda"
         Visible         =   0   'False
      End
      Begin VB.Menu s10 
         Caption         =   "-"
      End
      Begin VB.Menu Stockminymax 
         Caption         =   "Stocks Míinimos y Máximos x Artículo"
      End
   End
   Begin VB.Menu ocom 
      Caption         =   "Orden de Compra x Pedido"
      Begin VB.Menu cargadetped 
         Caption         =   "Cargar Detalle en Base a un Pedido"
      End
   End
   Begin VB.Menu Config 
      Caption         =   "Configuracion "
      Begin VB.Menu Setup 
         Caption         =   "Configuración General"
      End
      Begin VB.Menu plh1 
         Caption         =   "-"
      End
      Begin VB.Menu Nuorig1 
         Caption         =   "Números Originales"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "FORCOMP07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHIPRE$
Dim MUECOPEN%
Dim EMICONSI%
Dim MODOULPRE%
Dim RAI$ '\\\OT-05-0676-WAR-20/10/05///'REGLAMENTO DE COMPRA ANEXOS

Sub CambiosEnPantallaCompraInterna(NC)
   'compra entre sucursales de la misma empresa
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    IDdeTablaDatasql& = ExisteNumeroProveedorEnConfiguraciónCompraInterna&(NC)
    If IDdeTablaDatasql& > 0 Then
           Label10.Caption = "PR"
           Label10.ToolTipText = "Pedido de Reposición"
           Caption = REPLACAR(Caption, "Ordenes de Compra", "Pedido de Reposición")
           Dim CONTROLX As Control ' declaro el objeto control
           Dim OBJFRAME As Frame
           For Each CONTROLX In Me.Controls ' RECORRO LOS CONTROLES DEL FORM ACTIVO
            On Error Resume Next
            If TypeOf CONTROLX Is Frame Then
               Set OBJFRAME = CONTROLX
               OBJFRAME.BackColor = &HB0D0E0
            End If
            On Error GoTo 0
            Frame2.Left = 180
            
            Frame2.Top = 3840
            Frame2.Height = 2300
            Frame2.Width = 10600
            Frame2.Visible = True
            Label10.ForeColor = &HB0D0E0
            Option1.BackColor = &HB0D0E0
            Option2.BackColor = &HB0D0E0
            
           Next
           
           TEXT24 = NroClientePropioCompraInterna&()
      End If
End Sub

Sub DEPURACOMPRA_orig(Optional CantidadDiasTolerancia%)
    '
    'CANTIDAD DE DIAS CONFIGURADOS EN SETUP DE COMPRAS
    CantiDias% = XVALO(Control("ORCOMPRA", "DIOCVEN"))
    '
    If CantidadDiasTolerancia% > 0 Then CantiDias% = CantidadDiasTolerancia%
    If CantiDias% < 1 Then Exit Sub
    
    'TOMA LA FECHA DE HOY Y LE RESTO LOS DIAS DE TOLERANCIA (CANTIDIAS%)
    FEHOY% = HOY(HO$)
    FEHOY% = DIANTE(FEHOY%, CantiDias%)
    FECHOYTOLERA = FETEXCOR1$(FEHOY%) 'FECHA DE HOY - LOS DIAS DE TOLERANCIA
    '
    Screen.MousePointer = 11
    'CONSULTO ORDENES DE COMPRA DONDE LA FECHA DE ENTREGA SEA < A LA FECHA DE HOY - TOLERANCIA
    SQLX$ = "Select * from OCOMDETA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " Where Stat_Ocom < 8 "
    SQLX$ = SQLX$ + " AND Stat_Ocom <> -1 "
    SQLX$ = SQLX$ + " AND Cant_Rec < Cant_Ocom - 0.005 "
    SQLX$ = SQLX$ + " AND FENTRE_SOL < '" & FECHOYTOLERA & "'"
    SQLX$ = SQLX$ + " ORDER BY Nume_Ocom, Cod_Exte, FEntre_Sol "
    Dim OCOMTMP As ADODB.Recordset
    Set OCOMTMP = READSET(SQLX$)
    '
    FEHOY% = HOY(HO$)
    '
    ORCOPEND07.List1.Clear
    '
    With OCOMTMP
      Do While Not .EOF
      '
        TX$ = ""
        FESOL% = CVINT(!Fentre_Sol)
        FEMI% = CVINT(!Femi_Ocom)
        NP% = XVALO(!Npro_Ocom)
        ROZOPRO$ = SAFETEXT$(!Raso_Prov)
        NOCOMPRA& = XVALO(!nume_ocom)
        CIXI% = XVALO(!Cod_Inte)
        CANTICOMPRA = XVALO(!Cant_Ocom)
        CANTIRECIBI = XVALO(!Cant_Rec)
        NUOR% = XVALO(!NuOrd_Item)
        
        '
        'ARMO EL STRING PARA MOSTRAR EN EL LIST
        Call REPLA(TX$, FORMATNUM$(NOCOMPRA&, "F9.0"), 1, 9)
        Call REPLA(TX$, FORMATNUM$(NP%, "F5.0"), 11, 5)
        Call REPLA(TX$, ROZOPRO$, 17, 17)
        Call REPLA(TX$, CODEXT$(CIXI%), 35, 16)
        Call REPLA(TX$, DESCRIT$(CIXI%), 52, 24)
        Call REPLA(TX$, FECHATEX$(FESOL%), 77, 8)
        Call REPLA(TX$, FORMATNUM$(CANTICOMPRA, "F8.1"), 86, 8)
        Call REPLA(TX$, FORMATNUM$(CANTIRECIBI, "F8.1"), 95, 8)
        Call REPLA(TX$, FORMATNUM$(CANTICOMPRA - CANTIRECIBI, "F8.1"), 103, 8)
        Call REPLA(TX$, FORMATNUM$((NUOR%), "F4.0"), 124, 4)
        'AGREGO AL LIST
        ORCOPEND07.List1.AddItem TX$
'        Debug.Print TX$
     .MoveNext
     Loop
     End With
     '
     OCOMTMP.Close
     Set OCOMTMP = Nothing
     '
     Screen.MousePointer = 1
     'MUESTRO LA VENTANA CON LOS ITEMS
     If ORCOPEND07.List1.ListCount > 0 Then
        TX$ = "Detalle de Ordenes de Compras con Fecha Entrega Solicitada Vencida (Tolerancia: " & CantiDias% & " Días)"
        ORCOPEND07.Caption = TX$
        ORCOPEND07.Show 1
     Else
         Screen.MousePointer = 1
         Exit Sub
     End If
     '
     'VALIDO DERECHOS PARA EJECUTAR TRANSACCION EN EL CASO QUE HAYA APROBADO
     If ORCOPEND07.List1.ListCount > 0 Then
        If DERACCE%("DEPOCVE", "Depuración de Ordenes de Compra Vencidas") < 2 Then
           Exit Sub
        End If
     End If
     '
     ITEMANUL% = 0
     Screen.MousePointer = 11
     '
     'BORRO LOS REGISTROS SELECCIONADOS
     If ORCOPEND07.List1.ListCount > 0 Then
       For i = 1 To ORCOPEND07.List1.ListCount
          If ORCOPEND07.List1.Selected(i - 1) = True Then
             ITEMANUL% = ITEMANUL% + 1
             NOCOMPRA& = XVALO(Left$(ORCOPEND07.List1.List(i - 1), 9))
             NP% = XVALO(Mid$(ORCOPEND07.List1.List(i - 1), 11, 5))
             CIXI% = CODINT%(Mid$(ORCOPEND07.List1.List(i - 1), 35, 16))
             NUOR% = XVALO(Mid$(ORCOPEND07.List1.List(i - 1), 124, 4))
             CONDI$ = " NUME_OCOM = " & NOCOMPRA& & " AND COD_INTE = " & CIXI%
             CONDI$ = CONDI$ + " AND NPRO_OCOM = " & NP%
             CONDI$ = CONDI$ + " AND NUORD_ITEM = " & NUOR%
             CONDI$ = CONDI$ + " AND Stat_Ocom < 8"
             '
             Call ACTUREGISTRO("OCOMDETA", "Stat_Ocom", CONDI$, 9, REGAF&, "NOMESS")
             '
          End If
       Next i
     End If
     '
     Screen.MousePointer = 1
     If ORCOPEND07.List1.ListCount > 0 Then
       Call COMUNI("Se Han Anulado " & ITEMANUL% & " Items Seleccionados")
     End If
     
End Sub
Sub DEPURACOMPRA(Optional CantidadDiasTolerancia%)
    '
    'CANTIDAD DE DIAS CONFIGURADOS EN SETUP DE COMPRAS
    CantiDias% = XVALO(Control("ORCOMPRA", "DIOCVEN"))
    '
    If Abs(CantidadDiasTolerancia%) > 0 Then CantiDias% = CantidadDiasTolerancia%
    If Abs(CantiDias%) < 1 Then Exit Sub
    
20  Campos$ = "N° Compra/F9;Cta./F7;R.Social/A24;Codigo/A16;Descripción/A32;F.O.C./D12;F.E.S./D12;C.Comp./F9.1;C.Rec./F9.1;Saldo/F9.1;Orden/F5"
    
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
    
    
    GRILLAGRAL.mnuMenu1.Caption = "Anular Ordenes de Compra Seleccionadas"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.mnuMenu2.Caption = "Enviar Por Mail"
    GRILLAGRAL.mnuMenu2.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.TXTBUSCAR = ""

    
    'TOMA LA FECHA DE HOY Y LE RESTO LOS DIAS DE TOLERANCIA (CANTIDIAS%)
    FEHOY% = HOY(HO$)
    If CantiDias% > 0 Then
       FEHOY% = DIANTE(FEHOY%, CantiDias%)
    Else
       FEHOY% = DIASPOST(FEHOY%, Abs(CantiDias%)) 'CONSIDERA CANTIDAD DE DIAS EN NEGATIVO
    End If
    
    FECHOYTOLERA = FETEXCOR1$(FEHOY%) 'FECHA DE HOY - LOS DIAS DE TOLERANCIA
    '
    'CARGO EL RECORDSET
    Dim obj As New RECORXET
    Set RS1 = obj.RS_DEPURACOMP(FECHOYTOLERA)
    Call Carga_MSF_Recordset(RS1, Campos$, GRILLAGRAL.GRID, 1)
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    '*******************************************
    Screen.MousePointer = 1
    If GRILLAGRAL.GRID.Rows < 2 Then Exit Sub
    
    TX$ = "Detalle de Ordenes de Compras con Fecha Entrega Solicitada Vencida (Tolerancia: " & CantiDias% & " Días)"
    GRILLAGRAL.Caption = TX$
    '*******************************************
30  GRILLAGRAL.Show 1
    ITEMANUL% = 0
    Dim MSF As msFlexGrid

    If GRILLAGRAL.MENU1 > 0 Then
       If DERACCE%("DEPOCVE", "Depuración de Ordenes de Compra Vencidas") >= 2 Then
            Set MSF = GRILLAGRAL.GRID
            If TRIM$(GRILLAGRAL.TXTBUSCAR) <> "" Then Set MSF = GRILLAGRAL.MSF_2
            For i& = 1 To MSF.Rows - 1
               If MSF.TextMatrix(i&, 0) = "*" Then
                 ITEMANUL% = ITEMANUL% + 1
                 NOCOMPRA& = XVALO(MSF.TextMatrix(i&, 1))
                 NP% = XVALO(MSF.TextMatrix(i&, 2))
                 CIXI% = CODINT%(MSF.TextMatrix(i&, 4))
                 NUOR% = XVALO(MSF.TextMatrix(i&, 11))
                 CONDI$ = " NUME_OCOM = " & NOCOMPRA& & " AND COD_INTE = " & CIXI%
                 CONDI$ = CONDI$ + " AND NPRO_OCOM = " & NP%
                 CONDI$ = CONDI$ + " AND NUORD_ITEM = " & NUOR%
                 CONDI$ = CONDI$ + " AND Stat_Ocom < 8"
                 '
                 Call ACTUREGISTRO("OCOMDETA", "Stat_Ocom", CONDI$, 8, REGAF&, "NOMESS")
                 GRILLAGRAL.MENU1 = 0
               End If
            Next i&
            '
            If ITEMANUL% < 1 Then
               Call COMUNI("No se Ha Seleccionado Ningún Registro")
               GRILLAGRAL.MENU1 = 0 ' lo vuelvo a poner en 0 la bandera para saber si se cliqueo el menu1
               GoTo 30 'presento la ventana nuevamente
            Else
               Call COMUNI("Se Han Anulado " & ITEMANUL% & " Items Seleccionados")
               GoTo 20
            End If
       End If
    ElseIf GRILLAGRAL.MENU2 > 0 Then  'ENVIO DE MAIL
            GRILLAGRAL.MENU2 = 0
            Set MSF = GRILLAGRAL.GRID
            If TRIM$(GRILLAGRAL.TXTBUSCAR) <> "" Then Set MSF = GRILLAGRAL.MSF_2
            Call CARGA_ENCABEZADO(MAILNEW.MSF, Campos$, MAILNEW)
            Call CopiarMsfAOtroMsf(MSF, MAILNEW.MSF, 1, 0, MAILNEW)
            
            MAILNEW.Show 1
    End If
       
    
    

     '
     
End Sub

Function Excepcion(CI1%, COMPRA&, Optional NPROV)
  'cliente es el que esta recibiendo
  CONDI$ = "NUME_OCOM = " & COMPRA& & " AND CODINT = " & CI1%
  Excepcion = XVALO(VALOBADA("EXCEPCION", "CANTIEXCEPCION", CONDI$, "NOMESS"))
  
End Function

Sub GENERAR_TRANSFERENCIA_OCOMPRA(Optional MODO%, Optional MENSAXE$)
        
        Ruta$ = ""
        If MODO% = 1 Then 'ES MODIFICACION
          VERSIONX% = XVALO(FOMODOC07.Label16)
          OCOMPRA& = XVALO(FOMODOC07.Text11)
          NROCLIENTE = XVALO(FOMODOC07.TEXT24)
          ORCOMCLIENTE$ = TRIM$(FOMODOC07.TEXT25)
          RAZO_NC_PEDIDO$ = Left$(TRIM$(FOMODOC07.Eco24), 32)
        Else
          VERSIONX% = XVALO(Label16)
          OCOMPRA& = XVALO(Text11)
          NROCLIENTE = XVALO(TEXT24)
          ORCOMCLIENTE$ = TRIM$(TEXT25)
          RAZO_NC_PEDIDO$ = Left$(Me.Eco24, 32)
        End If
        '
        EPAC$ = TRIM$(UCase$(EMPREAC$))
        If InStr(EPAC$, "METALURGICASIAM-ARCON") > 0 Then NC_CONFIG1 = 999
        If InStr(EPAC$, "NEUQUEN") > 0 Then NC_CONFIG1 = 998
        

        
''        NC_CONFIG1 = XVALO(CONTROL("RETRASLA", "SUC1CLIE"))
''        NC_CONFIG2 = XVALO(CONTROL("RETRASLA", "SUC2CLIE"))
''        If NC_CONFIG1 = NCLIE Then
'        RUTA$ = CONTROL("RETRASLA", "SUC1RUTA")
''         ElseIf NC_CONFIG2 = NCLIE Then
''             RUTA$ = CONTROL("RETRASLA", "SUC2RUTA")
''        End If
        '
        'CARPETASINCRO$ = TRIM$(Control("", "CARPEINT"))
        CARPETASINCRO$ = RutaCarpetaLocalCompraInterna$
        Ruta$ = CARPETASINCRO$
        MsgBox Ruta$
        If Ruta$ = "" Then
           Call COMUNI("Imposible Continuar.\Falta Definir Carpeta de Intercambio")
           Exit Sub
        End If
        
        A$ = Date
        A$ = REPLACAR(A$, "/", "_")
        B$ = Time
        B$ = REPLACAR(B$, ".", "")
        B$ = REPLACAR(B$, ":", "")
        
        NARCHI2% = FILEOPEN%(Ruta$ + "\Ocompra_Sucursal " + A$ + " " + B$ + ".TXT", 2, 1024)    ' ABRE EL ARCHIVO DE ESCRITURA
        TX$ = AJUSTI$("SUCURSAL", 8): TX$ = TX$ + ";" 'NCLIENTE ARCON O COMODORO
        TX$ = TX$ + AJUSTI$("CODIGO", 8): TX$ = TX$ + ";"  'CODIGO
        TX$ = TX$ + AJUSTI$("NRO.PED.", 8): TX$ = TX$ + ";" 'NRO DE PEDIDO
        TX$ = TX$ + AJUSTI$("CTE.PED.", 8): TX$ = TX$ + ";" ' CLIENTE DEL PEDIDO
        TX$ = TX$ + AJUSTI$("USUARIO", 8): TX$ = TX$ + ";" 'USUARIO
        TX$ = TX$ + AJUSTI$("OCOMPRA", 8): TX$ = TX$ + ";" 'O.COMPRA
        TX$ = TX$ + AJUSTI$("CANTIDAD", 10): TX$ = TX$ + ";" 'CANTIDAD
        TX$ = TX$ + AJUSTI$("SOLENTRE", 10): TX$ = TX$ + ";" 'FECHA SOLICITUD DE ENTREGA
        TX$ = TX$ + AJUSTI$("VERS-COMP", 8): TX$ = TX$ + ";" 'VERISION
        TX$ = TX$ + AJUSTI$("FEMISION", 8): TX$ = TX$ + ";" 'FECHA DE EMISION
        TX$ = TX$ + AJUSTI$("O.COMPRA CLIENTE", 16): TX$ = TX$ + ";"
        TX$ = TX$ + AJUSTI$("R.SOCIAL CTE.PEDIDO", 32)
        Print #NARCHI2%, TX$ ' ESCRIBE EN EL ARCHIVO EL NUEVO STRING DE CONEXION
        '
        USER1% = USNBR% Mod 100
        FEHOY% = HOY(HO$)
        For i& = 1 To ULTREG&("!OCOMDETA")
          X$ = REGLEIDO$("!OCOMDETA", i&)
          '
          CI1% = CVI(Mid$(X$, 9, 2))
          FENTRESOL% = CVI(Mid$(X$, 95, 2))
          If FENTRESOL% < FEHOY% Then FENTRESOL% = FEHOY%
          CANSTOK = CVS(Mid$(X$, 51, 4))
          NSUCURSAL = NC_CONFIG1
'          NPX& = XVALO(Mid$(x$, 31, 24))
'          If NPX& > 0 And NCPEDIDO < 1 Then
'            NCPEDIDO = XVALO(VALOBADA("PEDDETA", "NROCLIE", "NROPED = " & NPX&, "NOMESS"))
'          End If

          TX$ = AJUSTI$(SAFETEXT$(NC_CONFIG1), 8): TX$ = TX$ + ";" 'NCLIENTE ARCON O COMODORO
          TX$ = TX$ + AJUSTI$(SAFETEXT$(CI1%), 8): TX$ = TX$ + ";"  'CODIGO
          TX$ = TX$ + AJUSTI$(SAFETEXT$(NPX&), 8): TX$ = TX$ + ";" ''NRO DE PEDIDO
          TX$ = TX$ + AJUSTI$(SAFETEXT$(NROCLIENTE), 8): TX$ = TX$ + ";" ' CLIENTE DEL PEDIDO
          TX$ = TX$ + AJUSTI$(SAFETEXT$(USER1%), 8): TX$ = TX$ + ";" 'USUARIO
          TX$ = TX$ + AJUSTI$(SAFETEXT$(OCOMPRA&), 8): TX$ = TX$ + ";" 'O.COMPRA
          TX$ = TX$ + AJUSTI$(SAFETEXT$(CANSTOK), 10): TX$ = TX$ + ";" 'CANTIDAD
          TX$ = TX$ + AJUSTI$(FECHATEX$(FENTRESOL%), 10): TX$ = TX$ + ";" 'FECHA SOLICITUD DE ENTREGA
          TX$ = TX$ + AJUSTI$(SAFETEXT$(VERSIONX%), 8): TX$ = TX$ + ";" 'VERISION
          TX$ = TX$ + AJUSTI$(FECHATEX$(FEHOY%), 8): TX$ = TX$ + ";" ' fecha de emision
          TX$ = TX$ + AJUSTI$(ORCOMCLIENTE$, 16): TX$ = TX$ + ";" 'ORDEN DE COMPRA DEL CLIENTE
          TX$ = TX$ + AJUSTI$(RAZO_NC_PEDIDO$, 32): TX$ = TX$ + ";"
          Print #NARCHI2%, TX$ ' ESCRIBE EN EL ARCHIVO EL NUEVO STRING DE CONEXION
        Next i&
        'SI SE GENERO MODIFICACION GUARDA EL MENSAJE
        If TRIM$(MENSAXE$) <> "" Then
           TX$ = "||MENSAJE||;" & Left$(MENSAXE$, 1000)
           Print #NARCHI2%, TX$
        End If
        '
        Close #NARCHI2%
End Sub

Sub GENERAR_TRANSFERENCIA_OCOMPRA_INTERNAS_CONTEPLAST(NPROV, NC_SOLICITANTE, Optional MODO%, Optional MENSAXE$)
        'CONTEPLAST
        Ruta$ = ""
        
        If MODO% = 1 Then 'ES MODIFICACION
          VERSIONX% = XVALO(FOMODOC07.Label16)
          OCOMPRA& = XVALO(FOMODOC07.Text11)
          NROCLIENTE = XVALO(FOMODOC07.TEXT24)
          ORCOMCLIENTE$ = TRIM$(FOMODOC07.TEXT25)
          RAZO_NC_PEDIDO$ = Left$(TRIM$(FOMODOC07.Eco24), 32)
        Else
          VERSIONX% = XVALO(Label16)
          OCOMPRA& = XVALO(Text11)
          NROCLIENTE = XVALO(TEXT24)
          ORCOMCLIENTE$ = TRIM$(TEXT25)
          RAZO_NC_PEDIDO$ = Left$(Me.Eco24, 32)
        End If
        '
        
'        EPAC$ = TRIM$(UCase$(EMPREAC$))
'        If InStr(EPAC$, "METALURGICASIAM-ARCON") > 0 Then NC_CONFIG1 = 999
'        If InStr(EPAC$, "NEUQUEN") > 0 Then NC_CONFIG1 = 998
        

        
''        NC_CONFIG1 = XVALO(CONTROL("RETRASLA", "SUC1CLIE"))
''        NC_CONFIG2 = XVALO(CONTROL("RETRASLA", "SUC2CLIE"))
''        If NC_CONFIG1 = NCLIE Then
'        RUTA$ = CONTROL("RETRASLA", "SUC1RUTA")
''         ElseIf NC_CONFIG2 = NCLIE Then
''             RUTA$ = CONTROL("RETRASLA", "SUC2RUTA")
''        End If
        '
        CARPETASINCRO$ = RutaCarpetaSegunNProvCompraInterna$(NPROV)
        Ruta$ = CARPETASINCRO$
        If Ruta$ = "" Then
           Call COMUNI("Imposible Continuar.\Falta Definir Carpeta de Intercambio")
           Exit Sub
        End If
        
        A$ = Date
        A$ = REPLACAR(A$, "/", "_")
        B$ = Time
        B$ = REPLACAR(B$, ".", "")
        B$ = REPLACAR(B$, ":", "")
        
        NARCHI2% = FILEOPEN%(Ruta$ + "\Ocompra_Sucursal " + A$ + " " + B$ + ".TXT", 2, 1024)    ' ABRE EL ARCHIVO DE ESCRITURA
        TX$ = AJUSTI$("SUCURSAL", 8): TX$ = TX$ + ";" 'NCLIENTE ARCON O COMODORO
        TX$ = TX$ + AJUSTI$("CODIGO", 8): TX$ = TX$ + ";"  'CODIGO
        TX$ = TX$ + AJUSTI$("NRO.PED.", 8): TX$ = TX$ + ";" 'NRO DE PEDIDO
        TX$ = TX$ + AJUSTI$("CTE.PED.", 8): TX$ = TX$ + ";" ' CLIENTE DEL PEDIDO
        TX$ = TX$ + AJUSTI$("USUARIO", 8): TX$ = TX$ + ";" 'USUARIO
        TX$ = TX$ + AJUSTI$("OCOMPRA", 8): TX$ = TX$ + ";" 'O.COMPRA
        TX$ = TX$ + AJUSTI$("CANTIDAD", 10): TX$ = TX$ + ";" 'CANTIDAD
        TX$ = TX$ + AJUSTI$("SOLENTRE", 10): TX$ = TX$ + ";" 'FECHA SOLICITUD DE ENTREGA
        TX$ = TX$ + AJUSTI$("VERS-COMP", 8): TX$ = TX$ + ";" 'VERISION
        TX$ = TX$ + AJUSTI$("FEMISION", 8): TX$ = TX$ + ";" 'FECHA DE EMISION
        TX$ = TX$ + AJUSTI$("O.COMPRA CLIENTE", 16): TX$ = TX$ + ";"
        TX$ = TX$ + AJUSTI$("R.SOCIAL CTE.PEDIDO", 32): TX$ = TX$ + ";"
        TX$ = TX$ + AJUSTI$("CODIGO_EXT", 24)
        X$ = TX$ + AJUSTI$(CODIGOEXT$, 24): TX$ = TX$ + ";"
        Print #NARCHI2%, TX$ ' ESCRIBE EN EL ARCHIVO EL NUEVO STRING DE CONEXION
        '
        USER1% = USNBR% Mod 100
        FEHOY% = HOY(HO$)
        For i& = 1 To ULTREG&("!OCOMDETA")
          X$ = REGLEIDO$("!OCOMDETA", i&)
          '
          CI1% = CVI(Mid$(X$, 9, 2))
          CODIGOEXT$ = CODEXT$(CI1%)
          FENTRESOL% = CVI(Mid$(X$, 95, 2))
          If FENTRESOL% < FEHOY% Then FENTRESOL% = FEHOY%
          'CANSTOK = CVS(Mid$(X$, 51, 4)) 'CAMBIO PARA QUE TOME LA CANTIDAD DE COMPRA
          CANSTOK = CVS(Mid$(X$, 55, 4))
          'UNIST$ = UCase$(TRIM$(Mid$(X$, 11, 4)))
          UNICO$ = UCase$(TRIM$(Mid$(X$, 59, 4)))
          TX$ = AJUSTI$(SAFETEXT$(NC_SOLICITANTE), 8): TX$ = TX$ + ";" 'NCLIENTE ARCON O COMODORO
          TX$ = TX$ + AJUSTI$(SAFETEXT$(CI1%), 8): TX$ = TX$ + ";"  'CODIGO
          TX$ = TX$ + AJUSTI$(SAFETEXT$(NPX&), 8): TX$ = TX$ + ";" ''NRO DE PEDIDO
          TX$ = TX$ + AJUSTI$(SAFETEXT$(NROCLIENTE), 8): TX$ = TX$ + ";" ' CLIENTE DEL PEDIDO
          TX$ = TX$ + AJUSTI$(SAFETEXT$(USER1%), 8): TX$ = TX$ + ";" 'USUARIO
          TX$ = TX$ + AJUSTI$(SAFETEXT$(OCOMPRA&), 8): TX$ = TX$ + ";" 'O.COMPRA
          TX$ = TX$ + AJUSTI$(SAFETEXT$(CANSTOK), 10): TX$ = TX$ + ";" 'CANTIDAD
          TX$ = TX$ + AJUSTI$(FECHATEX$(FENTRESOL%), 10): TX$ = TX$ + ";" 'FECHA SOLICITUD DE ENTREGA
          TX$ = TX$ + AJUSTI$(SAFETEXT$(VERSIONX%), 8): TX$ = TX$ + ";" 'VERISION
          TX$ = TX$ + AJUSTI$(FECHATEX$(FEHOY%), 8): TX$ = TX$ + ";" ' fecha de emision
          TX$ = TX$ + AJUSTI$(ORCOMCLIENTE$, 16): TX$ = TX$ + ";" 'ORDEN DE COMPRA DEL CLIENTE
          TX$ = TX$ + AJUSTI$(RAZO_NC_PEDIDO$, 32): TX$ = TX$ + ";"
          TX$ = TX$ + AJUSTI$(CODIGOEXT$, 24): TX$ = TX$ + ";"
          Print #NARCHI2%, TX$ ' ESCRIBE EN EL ARCHIVO EL NUEVO STRING DE CONEXION
        Next i&
        'SI SE GENERO MODIFICACION GUARDA EL MENSAJE
        If TRIM$(MENSAXE$) <> "" Then
           TX$ = "||MENSAJE||;" & Left$(MENSAXE$, 1000)
           Print #NARCHI2%, TX$
        End If
        '
        Close #NARCHI2%
End Sub


Sub IMPRLISTMAT()
    '
    'IMPRESION DE FROMULARIO DE LISTA DE MATERIALES PARA ARMADORES
    If Control$("OCOMPRA", "LISMATE") = "SI" Then
        OCOMP& = XVALO(Text11.TEXT)
        For J& = 1 To ULTREG&("!OCOMDETA")
           ZZ$ = REGLEIDO$("!OCOMDETA", J&)
           Ci% = CVI(Mid$(ZZ$, 9, 2))
           CANTU = CVS(Mid$(ZZ$, 55, 4))
           CAIT% = CantRegistros&("FORMULAS", "CodIConj =" & Ci%)
           If CAIT% > 0 Then
              FILTX$ = TRIM$(Str$(Ci%)) & ";" & TRIM$(Str$(OCOMP&)) & ";" & TRIM$(Str$(CANTU))
              Call STDFORM("LISTAMAT", FILTX$, "PRINT")
           End If
        Next J&
    End If

End Sub



'Function MaximoCompra&(CI1%, ValidaStockYPedidos%, ValidaStockYPedidosStockMin%, CANTI, TX$, Optional MODO$)
'
'    MaximoCompra& = 0
'    TX$ = ""
'    'CRITERIO
'    'SI POSEE EXCEPCION ESTA MANDA SOBRE TODO O SEA QUE EL MAXIMO DE COMPRA VA A SER IGUAL A LA EXCEPCION.
'    If MODO$ <> "MODIFICACION" Then
'      NPROV = XVALO(Text1)
'      NUOC& = XVALO(Text11)
'    Else
'      NPROV = XVALO(FOMODOC07.LText1) 'MODIFICACION
'      NUOC& = XVALO(FOMODOC07.Text11) 'MODIFICACION
'    End If
'    '
'    'ASIGNO VALORES
'    ComprasPendientes = CompraPendientesAcumuladas(CI1%, NUOC&)
'    PedidosPendientes = CANTIDAD_PEDIDA_PENDIENTE(CI1%)
'    EnRecepcion = CANRECEP(CI1%)
'
'    NecesidadPorPedidosMenosCompra = PedidosPendientes - (ComprasPendientes + EnRecepcion)
'    StockFisico = STOCKACT(CI1%)
'    StockMaximo = LOREPOMA(CI1%)
'    stockMinimo = STOMIN(CI1%)
'    '
'    EXCECION& = Excepcion(CI1%, NUOC&, NPROV)
'    If EXCECION& > 0 Then
'       MaximoCompra& = EXCECION&
'       GoTo 90
'    End If
'
'    'SI CONSIDERA STOCK MINIMO SOLO PUEDE PEDIR SI ESTA POR DEBAJO DEL STOCK MINIMO
'    'POR LO TANTO
'    '
'    Necesidad = NecesidadPorPedidosMenosCompra - StockFisico
'
'    'la necesidad ya esta considerando el stock si es positivo falta
'    'si es negativo sobra
'    If ValidaStockYPedidosStockMin% > 0 Then
'        'si el stock es >= al minimo se considera como tope el minimo
'        If StockFisico + Abs(NecesidadPorPedidosMenosCompra) >= stockMinimo Then
'           TopeStockMinyMax = stockMinimo
'        Else 'si es < se considera como topo el maximo
'           TopeStockMinyMax = StockMaximo
'        End If
'
'        If Necesidad > 0 Then
'              MaximoCompra& = Necesidad + TopeStockMinyMax
'        Else
'           Sobra = Abs(Necesidad)
'           MaximoCompra& = TopeStockMinyMax - Sobra
'        End If
'        GoTo 90
'    End If
'    '
'    'SI LLEGA ACA ES POR QUE NO ESTA CONFIGRUADO CONSIDERAR STOCK MINIMO.
'    If Necesidad > 0 Then
'        MaximoCompra& = Necesidad + StockMaximo
'    Else
'        Sobra = Abs(Necesidad)
'        MaximoCompra& = StockMaximo - Sobra
'    End If
'
'
'90  If MaximoCompra& < 0 Then MaximoCompra& = 0
'
'    'ESTO ES UN RETORONO
'    TX$ = CODEXT$(CI1%)
'    TX$ = TX$ + ";" + DESCRIT0$(CI1%)
'    TX$ = TX$ + ";" + FORMATNUM$(CANTI, "F10.1")
'    TX$ = TX$ + ";" + FORMATNUM$(stockMinimo, "F10.1")
'    TX$ = TX$ + ";" + FORMATNUM$(StockMaximo, "F10.1")
'    TX$ = TX$ + ";" + FORMATNUM$(StockFisico, "F10.1")
'    TX$ = TX$ + ";" + FORMATNUM$(PedidosPendientes, "F10.1")
'    TX$ = TX$ + ";" + FORMATNUM$(ComprasPendientes, "F10.1")
'    TX$ = TX$ + ";" + FORMATNUM$(EnRecepcion, "F10.1")
'    TX$ = TX$ + ";" + FORMATNUM$(EXCECION&, "F10.1")
'    TX$ = TX$ + ";" + FORMATNUM$(MaximoCompra&, "F10.1")
'
'End Function
'
Function MaximoCompra&(CI1%, ValidaStockYPedidos%, ValidaStockYPedidosStockMin%, CANTI, TX$, Optional MODO$)
    MaximoCompra& = 0
    TX$ = ""
    'CRITERIO
    'SI POSEE EXCEPCION ESTA MANDA SOBRE TODO O SEA QUE EL MAXIMO DE COMPRA VA A SER IGUAL A LA EXCEPCION.
    If MODO$ <> "MODIFICACION" Then
      NPROV = XVALO(Text1)
      NUOC& = XVALO(Text11)
    Else
      NPROV = XVALO(FOMODOC07.LTEXT1) 'MODIFICACION
      NUOC& = XVALO(FOMODOC07.Text11) 'MODIFICACION
    End If
    '
    'ASIGNO VALORES
    ComprasPendientes = CompraPendientesAcumuladas(CI1%, NUOC&, 1)
    PedidosPendientes = CANTIDAD_PEDIDA_PENDIENTE(CI1%)
    EnRecepcion = CANRECEP(CI1%)
    
'    NecesidadPorPedidosMenosCompra = PedidosPendientes - (ComprasPendientes + EnRecepcion)
    StockFisico = STOCKACT(CI1%)
    StockMaximo = LOREPOMA(CI1%)
    stockMinimo = STOMIN(CI1%)
    '
    EXCECION& = Excepcion(CI1%, NUOC&, NPROV)
    If EXCECION& > 0 Then
       MaximoCompra& = EXCECION&
       GoTo 90
    End If
    
    'SI CONSIDERA STOCK MINIMO SOLO PUEDE PEDIR SI ESTA POR DEBAJO DEL STOCK MINIMO
    'POR LO TANTO
    '
    StocksComprasPendientes = ComprasPendientes + StockFisico + EnRecepcion
    
    
    If ValidaStockYPedidosStockMin% > 0 Then
        If StocksComprasPendientes >= stockMinimo Then
           MaximoCompra& = stockMinimo - StocksComprasPendientes
        Else
           MaximoCompra& = StockMaximo - StocksComprasPendientes
        End If
        MaximoCompra& = PedidosPendientes + MaximoCompra&
        GoTo 90
    End If
    
    MaximoCompra& = StockMaximo - StocksComprasPendientes
    MaximoCompra& = PedidosPendientes + MaximoCompra&
    '
    'SI LLEGA ACA ES POR QUE NO ESTA CONFIGRUADO CONSIDERAR STOCK MINIMO.

    
90  If MaximoCompra& < 0 Then MaximoCompra& = 0
    
    'ESTO ES UN RETORONO
    TX$ = CODEXT$(CI1%)
    TX$ = TX$ + ";" + DESCRIT0$(CI1%)
    TX$ = TX$ + ";" + FORMATNUM$(CANTI, "F10.1")
    TX$ = TX$ + ";" + FORMATNUM$(stockMinimo, "F10.1")
    TX$ = TX$ + ";" + FORMATNUM$(StockMaximo, "F10.1")
    TX$ = TX$ + ";" + FORMATNUM$(StockFisico, "F10.1")
    TX$ = TX$ + ";" + FORMATNUM$(PedidosPendientes, "F10.1")
    TX$ = TX$ + ";" + FORMATNUM$(ComprasPendientes, "F10.1")
    TX$ = TX$ + ";" + FORMATNUM$(EnRecepcion, "F10.1")
    TX$ = TX$ + ";" + FORMATNUM$(EXCECION&, "F10.1")
    TX$ = TX$ + ";" + FORMATNUM$(MaximoCompra&, "F10.1")
    
End Function


Sub CreatablaExcepcion()
   Call CREACAMPO("", "Excepcion", "Nume_Ocom", 4, 0, 1)
   Call CREACAMPO("", "Excepcion", "NUCLIEN", 4, 0, 1)
   Call CREACAMPO("", "Excepcion", "CODINT", 4, 0, 1)
   Call CREACAMPO("", "Excepcion", "CODIGO", 10, 32)
   Call CREACAMPO("", "Excepcion", "FECHA", 8, 0)
   Call CREACAMPO("", "Excepcion", "USUARIO", 3, 0)
   Call CREACAMPO("", "Excepcion", "CantiExcepcion", 7, 0)
   Call CREACAMPO("", "Excepcion", "ID", 4, 0, 1)
End Sub

Function NROCOM$(NRO&)
      NOCOM$ = TRIM$(Str$(NRO&))
      While Len(NOCOM$) < 6: NOCOM$ = "0" + NOCOM$: Wend
      NOCOM$ = "OC-" + Left$(NOCOM$, 6) + "-"
      NROCOM$ = NOCOM$
End Function


Private Sub Anul_Ocom_Click()
   Call SETULTREG("!OCOMDETA", 0)
   Call BLANFORMU
   Call OPCICOM07.Command14_Click
End Sub

Private Sub AnulacionMasivaPorTolerancia_Click()

    CantidadDias% = XVALO(InputBox("Dias de tolerancia", "Cantidad de dias de tolerancia", ""))
    DEPURACOMPRA CantidadDias%
    
End Sub

Private Sub cargadetped_Click()
    TAPLI$ = Mid$(APLIXYZ$, 5)
    On Error Resume Next
    '
    Call BLANFORMU
    
    Call SELECHO("PROVED1")
    NPROV = XVALO(TRIM$(VALACT1$("PROVED1")))
    If NPROV < 1 Then Exit Sub

    Text1 = ""
    Text1.TEXT = TRIM$(Str$(NPROV))
    Call CARDACLI
    Call CARCOANT
    '
3   Call SELECHO("INDIPEP/Indice General de Pedidos (Pendientes)")
    NPX& = Val(VALACT1$("INDIPEP"))
    If NPX& < 1 Then Exit Sub

    Call ABRECOMPRAS
    '
    
    SQLX$ = "SELECT CodiInt,SUM(CanPed) AS CANTIPED FROM PEDDETA WITH(NOLOCK) WHERE Nroped = " & NPX&
    SQLX$ = SQLX$ + " and STATUS <= 8 "
    SQLX$ = SQLX$ + " GROUP BY CODIINT"
    Set COTRE = READSET(SQLX$)
    '
    JJ& = 0
    REBLA$ = REGBLAN$("OCOMDETA")
    Call SETULTREG("!OCOMDETA", 0)
    ARCHILIB$ = "!1OCOMDET"
    Call SETULTREG(ARCHILIB$, 0)
    '
    With COTRE
       Do While Not .EOF
         '
         CI1% = XVALO(!CODIINT)
         DECRI$ = rasocli$(-1 * NPROV)
         UMECO$ = UNICOM$(CI1%)
         PRUNI = 0 'carga precio unitario del articulo
         CANTU = XVALO(!CANTIPED)
         FESOL% = CVINT(!FeSolEnt)
         '
         ZZ$ = REBLA$
         Call REPLA(ZZ$, MKI$(CI1%), 9, 2)
         Call REPLA(ZZ$, Unimed$(CI1%), 11, 4)
         Call REPLA(ZZ$, MKS$(CANTU), 51, 4)
         Call REPLA(ZZ$, UMECO$, 59, 4)
         Call REPLA(ZZ$, MKS$(1), 63, 4)
         Call REPLA(ZZ$, MKD$(0), 67, 8)
         Call REPLA(ZZ$, 0, 75, 4)
         Call REPLA(ZZ$, MKD$(0), 79, 8)
         Call REPLA(ZZ$, MKD$(0), 87, 8)
         Call REPLA(ZZ$, MKI$(FESOL%), 95, 2)
            SoliciYNPres$ = SAFETEXT$(NPX&)
         Call REPLA(ZZ$, SoliciYNPres$, 31, 20)
         '
         JJ& = JJ& + 1
         Call GRAREG("!OCOMDETA", ZZ$, JJ&)
         Call GRAREG(ARCHILIB$, DECRI$, JJ&)
       .MoveNext
       Loop
    End With
    Call SETULTREG("!OCOMDETA", JJ&)
    Call SETULTREG(ARCHILIB$, JJ&)
    '
    Call CARDETCOM
    '

End Sub

Private Sub Command1_Click()
   Call SELECHO("PROVED1")
   NCLIE% = XVALO(TRIM$(VALACT1$("PROVED1")))
   If NCLIE% > 0 Then
     If REGICLI&((-1) * NCLIE%) > 0 Then
       Call BLANFORMU
       Text1.TEXT = TRIM$(Str$(NCLIE%))
       Call CARDACLI
       Call CARCOPEN
       Call CARCOANT
       If ULTREG&("!OCOMDETA") > 0 Then
            Call CARLISCO
            Call CALTOCOM
         Else
            Call CARDETCOM
       End If
     End If
   End If
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   CONDI$ = "Npro_Ocom > 0 AND Cantidad > 0"
   If CantRegistros("ocoemi", CONDI$) > 0 Then
      Call CARGALISEL("PROVEPLA", "OCOEMI", "Npro_Ocom/F6+Proved1/A30", CONDI$, "DISTINCT")
   End If
   
   Call SELECHO("PROVEPLA")
   NCLIE% = XVALO(TRIM$(VALACT1$("PROVEPLA")))
   If NCLIE% > 0 Then
     If REGICLI&((-1) * NCLIE%) > 0 Then
       Call BLANFORMU
       Text1.TEXT = TRIM$(Str$(NCLIE%))
       Call CARDACLI
       Call CARCOPEN
       Call CARCOANT
       If ULTREG&("!OCOMDETA") > 0 Then
            Call CARLISCO
            Call CALTOCOM
         Else
            Call CARDETCOM
       End If
     End If
   End If
   
99 Command10.Enabled = True
   
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call CIERRARCH("*.*")
   OPX% = ALTERNA%("Artículos de Stock\Padrón de Proveedores")
   If OPX% = 1 Then
        ACONEC$ = "ARCHIG07/AMASTO"
        Call CONECRUN(ACONEC$, "Maestro de Artículos de Stock")
      ElseIf OPX% = 2 Then
        ACONEC$ = "ARCHIG07/AMAPRO"
        Call CONECRUN(ACONEC$, "Padron Maestro de Proveedores")
   End If
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
    '
    Command12.Enabled = False
    '
3   Call SELECHO("TAMONED")
    MMI% = XVALO(VALACT1$("TAMONED"))
    If MMI% > 0 Then
      Text15 = TRIM$(Str$(MMI%))
    End If
    '
99  Command12.Enabled = True
    '
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   List1.Clear
   Call SETULTREG("!OCOMDETA", 0)
   OPCICOM07.Show 1
   Refresh
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
    '
    Command14.Enabled = False
    '
    If EmpresaNoRealizaCambios% > 0 Then
      GoTo 99
    End If
    XXX$ = rasocli$((-1) * (XVALO(Text1) + 1)) ' para que refresque

    DESTICOMPRAMINEPAR% = 0
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "AHP") > 0 Then
      If XVALO(TXTPESO) > 50 Then
       OPX% = COMALTER%("Atención Peso Total Supera los 50 Kg.\\Cancelar\Continuar")
       If OPX% <= 1 Then GoTo 99
      End If
      If InStr(EPAC$, "MINEPARTS") < 1 And InStr(EPAC$, "AHP") > 0 Then
         OPX% = COMALTER%("INDICAR DESTINO DE COMPRA\\Proveedor Activo\Proveedor Mineparts")
         If OPX% = 2 Then
            NprovMineparts = XVALO(Control("CFGMINEP", "PRFEMINE"))
            If NprovMineparts > 0 Then
               DESTICOMPRAMINEPAR% = 1
            Else
               Call COMUNI("Falta Configurar Proveedor Mineparts\Configure desde el Setup de Compras (Menú) y Vuelva a Intentar")
               GoTo 99
            End If
         End If
      End If

    End If
    '
    'validaciones de solicitud de pedidos de reposicion de neuquen
    If Frame2.Visible = True Then
       If Option2.Value = True Then
        If TRIM$(TEXT24) = "" Then
           Call COMUNI("Imposible Continuar !!! \ Se ha Seleccionado 'Solicitud por Pedido de Cliente'\Debe Seleccionar Un Cliente o Grabar por 'Pedido de Reposición de Stocks'")
           GoTo 99
        End If
        If TRIM$(TEXT25) = "" Then
           Call COMUNI("Falta Ingresar Orden de Compra del Cliente")
           GoTo 99
        End If
        '
        If OCOMEXISTENEUQUEN%(XVALO(TEXT24), TRIM$(TEXT25)) > 0 Then
          Call COMUNI("Atención Orden de Compra del Cliente Informado es Existente")
          GoTo 99
        End If
       End If
    End If
    '
    'CONSIDERAR VALIDACIONES DE STOCK MINIMO O MAXIMO Y PEDIDOS 'ARCON'
    Resultado$ = ""
    ValidaStockYPedidos% = 0: ValidaStockYPedidosStockMin% = 0:
    If Control("OCOMPRA", "VALISTPE") = "SI" Then
      ValidaStockYPedidos% = 1
      If Control("OCOMPRA", "CONSSTMI") = "SI" Then
        ValidaStockYPedidosStockMin% = 1
      End If
      If Frame2.Visible = True And Option2.Value = False Then 'SI ES POR ORDEN DE COMPARA DEL PEDIDO NO VALIDA
        Call ValidarCompraArcon%(ValidaStockYPedidos%, ValidaStockYPedidosStockMin%, Resultado$, "")
        If Resultado$ = "CANCELA" Then GoTo 99
      End If
    End If
    '
    
10  FORSELREG.Label3 = Text11 '\\**//
    Call SELEC_REGLACOM(OKX%)
    If OKX% < 1 Then GoTo 99
    '
    
    PORPRIN% = 1: PORMAIL% = 0
    NCI% = XVALO(Text1)
    NCIX% = NCI% * (-1)
    EMAI$ = TRIM$(EMAILCLI$(NCIX%))
    If EMAI$ = "" Or InStr(EMAI$, "@") < 1 Then
       GoTo 20 'NO TIENE DIRECCION O DIRECCION NO VALIDA
    End If
    '
    If DESTICOMPRAMINEPAR% < 1 Then
12      SALIOCOM.Show 1
        Call FRESHALL
        If SALIOCOM.Check1.Value < 1 Then
           If SALIOCOM.Check2.Value < 1 Then
              Call MENSERR(24, "Debe Marcar al Menos\una Dirección de Salida.")
              GoTo 12
           End If
        End If
    Else
        PORPRIN% = 1
        Call COMUNI("Orden de Compra Con Destino Mineparts\Generar Impresión o Envios Por Mail Desde La Pantalla de Modificación\Cada Uno Desde La Empresa Correspondiente")
    End If
    
    'Si esta activo enviar adjuntos
    'ENVIO DE MEMORIA TECNICA (ARCHIVO CARGADO EN ARCHIVO MAESTRO DE ARTICULOS SOLAPA   ESPECIFICACIONES)
    If SALIOCOM.Check3.Value > 0 Then 'OPCION ENVIO DE PDF
      For i = 1 To 256
        MAI_ADJUN$(i) = ""
      Next i
      FIN& = ULTREG("!OCOMDETA")
      For AX& = 1 To FIN&
        X$ = REGLEIDO$("!OCOMDETA", AX&)
        Ci% = CVI(Mid$(X$, 9, 2))
        If ARMEMOTEC$(Ci%) <> "" Then
           MAI_ADJUN$(AX&) = ARMEMOTEC$(Ci%)
        End If
      Next AX&
    End If
    '
    PORPRIN% = 0: If SALIOCOM.Check1.Value = 1 Then PORPRIN% = 1
    PORMAIL% = 0: If SALIOCOM.Check2.Value = 1 Then PORMAIL% = 1
    '
15  SENDBYMAIL$ = ""
    '
20  PCTJE# = XVALO(Control$("", "SOBREPRE")) ' VALIDACION DE PRECIO INGRESADO CON COSTO MAS PORCENTAJE LIMITE
    If PCTJE# > 0 Then
      For U& = 1 To ULTREG&("!OCOMDETA")
        ZZ$ = REGLEIDO$("!OCOMDETA", U&)
        Ci% = CVI(Mid$(ZZ$, 9, 2))
        '
        If Ci% > 0 Then
          PREUNIT# = CVD(Mid$(ZZ$, 67, 8)) * TICAMONE(MONEMI%) ' PRECIO INGRESADO en pesos
          If PREUNIT# > COSPESOS(Ci%) * (1 + PCTJE# / 100) Then
            Call COMUNI("Precio de Compra Supera el Límite Establecido.\Para el Código: " & CODEXT(Ci%) & "\Consulte a su Supervisor")
            GoTo 99
          End If
        End If
      Next U&
    End If
    '


    Call BLOQUEAR_CONTROLES(Me, 2, 1, 1, 1, 0, 1, 1) 'DESHABILITO TODOS LOS OBJETOS
    '
    Call GRACOMPRA(OKX%, Resultado$, NrodeOrdenDeCompraGenerado&)
    If OKX% > 0 Then
      If DESTICOMPRAMINEPAR% > 0 Then
         Call DUPLICARCOMPRAMINEPARTS(NrodeOrdenDeCompraGenerado&, NCI%)
      End If
      'ACTUALIZA INDICE DE ORDENES DE COMPRA
      CONDI$ = "Nume_Ocom > 0 ORDER BY Nume_Ocom"
      Call CARGALISEL("INDIOCOM", "OCOMENCA", "Nume_Ocom/F7;Refe_OcoM/A44", CONDI$, "DISTINCT")
      Call FRESHECHO("INDIOCOM")
      '
'      If DESTICOMPRAMINEPAR% > 0 Then 'SI EL DESTION ES MINPARTS
'         MEMOFORM$ = Control("", "FORMOCOM") 'GUARDA EL CONTROL
'         Call GRACONTROL("", "FORMOCOM", "OCOM-MIN") 'LE GRABA EL DE MINEPARTS
'      End If
      If PORPRIN% > 0 Then      ' EMITE POR IMPRESORA
          '\\\OT-08-01047-FL-21/11/08///Si la aprobacion es posterior y la impresión "Al Aprobar" realiza unicamente la impresion fantasma
          If Control$("OCOMPRA", "MODOAPRO") = "POSTERIOR" Then
             If Control$("OCOMPRA", "IMPRIME") = "ALAPROBA" Then
                CANCELPRINT% = 1
             End If
          End If
          '
          If DESTICOMPRAMINEPAR% > 0 Then
            CANCELPRINT% = 1
          End If
          
          Call PRICOMPRA

          If RAI$ <> "" Then Call PRIREGLACOM   ' REGLAMENTO DE COMPRA
          Call IMPRLISTMAT                      ' LISTA DE MATERIALES PARA ARMADORES (STDSCREEN- GABAL)
          '
       End If
       '
       If PORMAIL% > 0 Then      ' ENVIA POR MAIL
          'PERMITE ELEGIR MAIL DE ENVIA A PARTIR DE LA LISTA DE MAILS DE CONTACTOS
          Call CONTACTO_MAIL(Label8, NCIX%) 'AGREGA A SENDBYMAIL LOS MAIL DE LOS CONTACTOS SELECCIONADOS
          If SENDBYMAIL$ = "" Then
            SENDBYMAIL$ = TRIM$(EMAILCLI$(NCIX%))
          End If
          '
          For Each PRX1 In Printers
       If InStr(UCase$(PRX1.DeviceName), "PDFWRITER") > 0 Or InStr(UCase$(PRX1.DeviceName), "BIOPDF") > 0 Or InStr(UCase$(PRX1.DeviceName), "BULLZIP") > 0 Then GoTo 25
          Next
          Call MENSERR(24, "Imposible Enviar por Mail.\   \Falta Impresora Virtual PDF.")
          SENDBYMAIL = ""
          GoTo 30
          '

25        Call PRICOMPRA
          If RAI$ <> "" Then Call PRIREGLACOM   ' REGLAMENTO DE COMPRA
          Call IMPRLISTMAT                      ' LISTA DE MATERIALES PARA ARMADORES (STDSCREEN- GABAL)
          '
       End If
       
'       If DESTICOMPRAMINEPAR% > 0 Then
'         Call GRACONTROL("", "FORMOCOM", MEMOFORM$) 'VUELVE A GRABAR EL ORIGINAL DE AHP
'       End If
       
       Call BLOQUEAR_CONTROLES(Me, 1, 1, 1, 1, 0, 1, 1) 'HABILITO TODOS LOS OBJETOS
       '
       
30     'AGREGO ESTO PARA QUE SIEMPRE EN NEUQUEN ENTRE EN A GENERAR EL ARCHIVO
       If InStr(EPAC$, "NEUQUEN") > 0 Then
          Call GENERAR_TRANSFERENCIA_OCOMPRA(0, TRIM$(Text13))
          GoTo 40
       End If
       If Left$(Control$("OCOMPRA", "MODOAPRO"), 6) <> "DIRECT" Or Resultado$ = "GRABA_PENDIENTE_APROBACION" Then ' VALIDACION SI LA APROBACION ES DIRECTA O POSTERIOR
          Call COMUNI("O/compra Registrada como\PENDIENTE DE APROBACION")
       ElseIf Frame2.Visible = True And XVALO(TEXT24) > 0 Then 'CONTEPLAST COMPRAS INTERNAS
          NC_SOLICITANTE = NClteSegunNProvCompraInterna&(XVALO(Text1))
          Call GENERAR_TRANSFERENCIA_OCOMPRA_INTERNAS_CONTEPLAST(XVALO(Text1), NC_SOLICITANTE, 0, TRIM$(Text13))
       Else
          If Control("OCOMPRA", "ENVARSUC") = "SI" Then 'CONFIGURACION SETUP DE COMPRA DE ENVIO DE ARCHIVO DE O.C. A CENTRAL
             Call GENERAR_TRANSFERENCIA_OCOMPRA(0, TRIM$(Text13))
          End If
       End If
       '
40     Call BLANFORMU
       Call SETULTREG("!OCOMDETA", 0)
       Call CIERRARCH("*.*")
       Text1.TEXT = ""
       If Left$(APLINVO$, 4) = "CREC" Or Left$(APLINVO$, 4) = "CREM" Then
          Call FINAL("")
       End If
          
        'AGREGADO PARA ENVIAR POR MAIL
       If SENDBYMAIL$ <> "" Then
            Call RECONECTA
            Exit Sub
       End If
       'HASTA ACA AGREGADO
       On Error Resume Next
       Text1.SetFocus
       On Error GoTo 0
       '
       If CODCARPE$ <> "" Then
             Call FINAL("")
       End If
       '
       CANCELPRINT% = 0
       '
    End If
    '
99  Command14.Enabled = True
    '
End Sub

Sub ValidarCompraArcon(ValidaStockYPedidos%, ValidaStockYPedidosStockMin%, Resultado$, MODO$)
      VUELTA% = 0
      For U& = 1 To ULTREG&("!OCOMDETA")
        ZZ$ = REGLEIDO$("!OCOMDETA", U&)
        CI1% = CVI(Mid$(ZZ$, 9, 2))
        If CI1% > 0 Then
           CANSTOK = CVS(Mid$(ZZ$, 51, 4))
           If MaximoCompra&(CI1%, ValidaStockYPedidos%, ValidaStockYPedidosStockMin%, CANSTOK, TX$, MODO$) < CANSTOK Then
             If VUELTA% = 0 Then
                VUELTA% = 1
                Campos$ = "Código/A16;Descripción/A24;Cant.P.R./F10.1;St.Min./F10.1;St.Max./F10.1;St.Físico/F10.1;Ped.Clien./F10.1;O.Repos.Pend./F11.1;En Recep./F10.1;Excepción/F10.1;Máx.Repos./F11.1"
                Call CARGA_ENCABEZADO(MUESTRADATOS7.GRID, Campos$, MUESTRADATOS7)
                If ValidaStockYPedidosStockMin% > 0 Then
                   TXTT$ = "Configurado para Que Considere Stocks Mínimo." & vbCrLf
                   TXTT$ = TXTT$ + "Solo se Podrá Pedir en el Caso que el Stock sea Menor al Stock Mínimo. " + vbCrLf
                Else
                   TXTT$ = "Solo Es Posible Pedir por Debajo del Stock Máximo " & vbCrLf
                End If
                   TXTT$ = TXTT$ + "En este Caso Permite Completar hasta el Stock Máximo. " + vbCrLf
                   TXTT$ = TXTT$ + "No se Realizará Validación en los Siguientes Casos: " + vbCrLf
                   TXTT$ = TXTT$ + "1.- Si el Pedido de Reposición se Realiza para un Cliente con su Orden de Compra Correspondiente. " + vbCrLf
                   TXTT$ = TXTT$ + "2.- Que Haya una Excepción Generada Desde Central."
                   MUESTRADATOS7.Label1 = TXTT$
             End If
             
             Call TEXTOLOTES$(TX$, ";")
             With MUESTRADATOS7.GRID
                .Rows = .Rows + 1
                For i& = 1 To UBound(CADENATEX$)
                   .TextMatrix(.Rows - 1, i&) = CADENATEX$(i&)
                Next i&
             End With
           End If
        End If
      Next U&
      'SI DETECTO ALGUNA VALIDACION ADVERSA
      
      If VUELTA% > 0 Then
80       OPX% = COMALTER("Atención Se Han Detectado Cantidades de Pedidos de Reposición de Stock\Que Superan Máximos Según Stock Mínimo o Máximo\\Cancelar\Ver Detalle\Grabar Como Pendiente de Aprobación")
         If OPX% <= 1 Then Resultado$ = "CANCELA"
         If OPX% = 2 Then Resultado$ = "VER DETALLE"
         If OPX% = 3 Then Resultado$ = "GRABA_PENDIENTE_APROBACION"
      Else
         Resultado$ = "GRABAR"
      End If
      If Resultado$ = "VER DETALLE" Then
         MUESTRADATOS7.Caption = "Items que Superan el Máximo de Pedidos de Reposición de Stocks"
         MUESTRADATOS7.Show 1
         GoTo 80
      End If
      

End Sub


Private Sub Command15_Click()
   Command15.Enabled = False
   FORCOLIB07.Show 1
   Command15.Enabled = True
End Sub




Private Sub Command16_Click()
Form1.Show
End Sub



Private Sub Command18_Click()
   If VALICLIE% < 1 Then
     On Error Resume Next
     Text1.SetFocus
     On Error GoTo 0
     Exit Sub
   End If
   '
   If ULTREG&("ECUECON") < 1 Then
      Call MENSERR(24, "Falta Definir Plan de Cuentas Contables")
      Call CARCUEINGAS
      Exit Sub
   End If
   '
   RG1$ = REGSEL$("ECUECON")
   If Len(RG1$) > 4 Then
      Text7.TEXT = Left$(RG1$, 12)
   End If
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

Private Sub Command22_Click()
    Command22.Enabled = False
  FOAPROC07.Show 1
    'Screen.MousePointer = 11
    '
    'JJ& = 0
    'For U& = 1 To ULTREG&("PEDENCA")
2   '  XX$ = REGLEIDO$("PEDENCA", U&)
    '  If CVS(Left$(XX$, 4)) > 0 Then
    '    NPX& = CVS(Left$(XX$, 4))
    '    NCLIE% = CVI(Mid$(XX$, 55, 2))
    '    FEX = CVI(Mid$(XX$, 53, 2))
    '    ANULX$ = "": If Mid$(XX$, 267, 1) = Chr$(9) Then ANULX$ = "(Anul) "
    '    YY$ = MKS$(NPX&) + ANULX$ + TRIM$(Mid$(XX$, 5, 48))
    '    JJ& = JJ& + 1
    '    Call GRAREG("INDIPEDI", YY$, JJ&)
    '  End If
    'Next U&
    'Call SETULTREG("INDIPEDI", JJ&)
    'Screen.MousePointer = 1
    '
    'CONPEDID.Show 1
    '
    Command22.Enabled = True
End Sub

Private Sub Command25_Click()
   NCLIE = ListaSeleccionClientes("", True)
   If NCLIE < 1 Then Exit Sub
   '
   NC = NCLIE
   TEXT24.TEXT = TRIM$(Str$(NC))

'   Call SELECHO("DEUDOR1")
'   NCLIEN = XVALO(TRIM$(VALACT1$("DEUDOR1")))
'   TEXT24.TEXT = TRIM$(Str$(NCLIEN))
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call HELPONLINE("PRESUP01")
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    Call OPLISCO07.Command22_Click
99  Command4.Enabled = True
   End Sub

Sub Command5_Click()
    Command5.Enabled = False
    Call CIERRARCH("*.*")
    If DERACCE%("STPOCOM", "Configuracion de Funcionamiento Compras") = 2 Then
      SETUPCOM.Show 1
'      Call CONECRUN("FSETUP04/SETUPCOM", "Configuración de Funcionamiento")
      Call FINAL("")
    End If
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   OPLISCO07.Show 1
   Command6.Enabled = True
End Sub
'
Private Sub Command7_Click()
    Command7.Enabled = False
    If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
    End If
    Call SELECHO("SELFECHA")
    VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
    If VDEVU$ <> "" Then
      Text14.TEXT = VALACT1$("SELFECHA")
    End If
    Command7.Enabled = True
End Sub

Sub Command8_Click()
    Command8.Enabled = False
    OPX% = ALTERNA%("Ordenes de Compra Abiertas\Actualizar Precios Proveedores\Otros Accesos Directos")
    If OPX% = 1 Then
        Call CIERRARCH("*.*")
        ACONEC$ = "OCOABI07"
        Call CONECRUN(ACONEC$, "Ordenes de Compra Abiertas")
        Call FINAL("")
      ElseIf OPX% = 2 Then
        Call ACOSPROV
      ElseIf OPX% = 3 Then
        Call ACCDIR("OCOMPRA")
    End If
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
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text9.TEXT = VALACT1$("SELFECHA")
   End If
   FF% = FECHANUM(Text9.TEXT)
End Sub

Private Sub Desanuoc_Click()
   Call SETULTREG("!OCOMDETA", 0)
   Call BLANFORMU
   Call OPCICOM07.Command10_Click
End Sub

Private Sub DevoConsi_Click()
   OPCICOMP.Command15_Click
End Sub

Private Sub Directa_Click()
   Call Command1_Click
End Sub

Private Sub Eco21_GotFocus(Index As Integer)
   On Error Resume Next
   TEXT21(Index).SetFocus
   On Error GoTo 0
End Sub

Private Sub Entrecon_Click()
   OPCICOM07.Command2_Click
End Sub

Private Sub Exit_Click()
   Call Command2_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  While WindowState <> 0
    WindowState = 0
    Openforms = DoEvents
    AppActivate Caption
    SendKeys "%" & Chr$(32) & Chr$(13), True
  Wend
End Sub

Private Sub Form_Load()
    '
    VERANTER$ = "ORCOMP02"
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(1, EMPREAC$, "AHP") > 0 Then
        Ltota(8).Visible = True
        TXTPESO.Visible = True
    End If
    '
    If InStr(EPAC$, "MINEPARTS") > 0 Then
      UPDATMASTER% = 1
      Call GENECOMAS
      UPDATMASTER% = 0
    End If
    
    Call CONAPLI(APLINVO$, TERMINA%) ' Busca la aplicacion a mostrar
    If TERMINA% > 0 Then Call FINAL("")
    '
    'DESHABILITO LOS BOTONES DE SELECCION DE PROVEEDOR PARA QUE  NO INTERRRUMPAN LA EJECUCION
    'DE LA CARGA DE DATOS
    '****************************************************************************************
    Command1.Enabled = False
    Command10.Enabled = False
    '****************************************************************************************
    Call ABRECONEXION
    Call CreatablaExcepcion
    '
    EPAC$ = UCase$(TRIM$(EMPREAC$))
    If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
10     Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
       If InStr(EPAC$, "DOSIVAC") > 0 Or InStr(EPAC$, "ARCON") > 0 Then
         Text6.Width = 2320
         Label15.Visible = True
         Label17.Visible = True
       End If
    End If
    '
    If InStr(EPAC$, "METALURGICASIAM-ARCON") > 0 Or InStr(EPAC$, "NEUQUEN") > 0 Then
       Label10.Caption = "PR"
       Label10.ToolTipText = "Pedido de Reposición"
       Caption = REPLACAR(Caption, "Ordenes de Compra", "Pedido de Reposición")
       Dim CONTROLX As Control ' declaro el objeto control
       Dim OBJFRAME As Frame
       For Each CONTROLX In Me.Controls ' RECORRO LOS CONTROLES DEL FORM ACTIVO
        On Error Resume Next
        If TypeOf CONTROLX Is Frame Then
           Set OBJFRAME = CONTROLX
           OBJFRAME.BackColor = &H8000000F
        End If
        On Error GoTo 0
        Frame2.Left = 180
        
        Frame2.Top = 3840
        Frame2.Height = 2300
        Frame2.Width = 10600
        Frame2.Visible = True
        Label10.ForeColor = vbBlue
       Next
    End If
       
    Call GRATITFOR(Caption)
    '
    If ULTREG&("PROVEPLA") > 0 Then
      Command10.Visible = True
    End If
    '
    If Control$("OCOMPRA", "REPLAORI") = "SI" Then
      Nuorig.Enabled = True
      Nuorig1.Enabled = True
    End If
    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me, Picture2)

    If EXISTE%("OCOABI07.EXE") < 1 Then
      MnuNueOCAbi.Enabled = False
    End If
    '
    Call SETULTREG("!OCOMDETA", 0)
    '
    If InStr(1, EMPREAC$, "MINEPARTS") > 0 Then
      CONDI$ = "Nume_Ocom > 0 ORDER BY Nume_Ocom"
      Call CARGALISEL("INDIOCOM", "OCOMENCA", "Nume_Ocom/F7;Refe_OcoM/A44", CONDI$, "DISTINCT")
      Call FRESHECHO("INDIOCOM")
    End If
    'ESTAS ASIGNACION DE CONTROL SE SUBIO AQUI POR QUE SI SE INGRESA RAPIDO
    'EL NUMERO DE PROVEEDOR Y EL FRM NO SE TERMINO DE CARGAR NO CONSIDERA ESTA
    'CONFIGURACION  AL TRAER EL PRECIO
    MODOULPRE% = 1
    If Control("OCOMPRA", "PRECIDEF") = "COSTOREP" Then MODOULPRE% = 2
    Call GENERAR_ESTRUCTURAS_SQL
    Show
    '
    Call RECLAPAG
    '
    HOII% = HOY(HOS$)
    Text9.TEXT = HOS$
    Text14.TEXT = HOS$
    Text11.TEXT = TRIM$(Str$(PRONUOCO&))
    Label118.Caption = FORMATNUM$(ALIVA, "F5.2")
    '
    Text15.TEXT = TRIM$(Control$("OCOMPRA", "MONECOS"))
    If Text15 = "" Then Text15 = "1"      ' MONEDA DE VALUACIóN
    MONEMI% = 1
    MONEDAXDEFECTO% = XVALO(Control("ORCOMPRA", "MOCODEFE"))
    If MONEDAXDEFECTO% <> MONEMI% Then
       MONEMI% = MONEDAXDEFECTO%
    End If
    '
    'SE MODIFICO PARA QUE NO MUESTRE EL TEXTO ALINEADO A LA IZQUIERDA , DE ESTE MODO LO MUESTRA = Q' CON EL DOBLE CLICK
    DMX$ = DEMONECO$(MONEMI%)
    PPXX% = InStr(DMX$, " ")
    If PPXX% > 0 Then DMX$ = Left$(DMX$, PPXX% - 1)
    Label11 = DMX$
    '
    If MONEMI% > 1 Then
      Label11.ForeColor = RGB(255, 0, 0)
    End If
    '
    Call APERBASDAT("COMPRAS")
    '
    Screen.MousePointer = 11
    If Left$(APLINVO$, 4) = "CREC" Then
         Call CARGALACOMPRA
       ElseIf Left$(APLINVO$, 4) = "CREM" Then
         Call CARGACOMUL
       ElseIf Left$(APLINVO$, 4) = "CPRE" Then 'AHP, GENERACION DE O.C. A PARTIR DE PRESUPUESTO
         APLIXYZ$ = APLINVO$
         Call CARGAOCPRE
    End If
    Screen.MousePointer = 1
    '
    If Mid$(APLINVO$, 1, 7) = "NUEVAOC" Then
      If CODCARPE$ <> "" Then
        CONDI$ = "CodcarpImp ='" & CODCARPE$ & "' AND STATUS < 9"
        If CantRegistros("CARPEIMP", CONDI$) > 0 Then
          NPRO% = XVALO(VALOBADA("CARPEIMP", "Npro_Ocom", CONDI$, "NOMESS"))
          Text1 = NPRO%
          Label18 = CODCARPE$
          Call Text1_KeyPress(13)
          If Label18 <> "" Then
             Command1.Enabled = False
             Text1.Enabled = False
             Label8.Enabled = False
          End If
        End If
      End If
    End If
    '
    If InStr(1, EMPREAC$, "ALEJPA") < 1 Then
        Call DEPURACOMPRA
    End If
    '
    If Control("OCOMPRA", "VALISTPE") = "SI" Then
      CREAR_TABLA_DetaOVenta
      Call ACTUALIZAR_ARCHIVOS_DE_CENTRAL
      Call NOVEDADESAVELLANEDA(1, 0)
      'PONGO LOS MENU VISIBLES SI ES NEUQUEN
      repocliente.Visible = True
      pedreppendiente.Visible = True
      ocompendaprobastocksypedidos.Visible = True
      MnuOCVSAV.Visible = True
    End If

    If InStr(1, EMPREAC$, "FERVI") > 0 Then
        CONDI$ = "NRO_OCOM LIKE '%5073%' OR NRO_OCOM LIKE '%5074%'"
        Call ACTUREGISTRO("MACONSIG", "CANT_ASIG", CONDI$, "CANT_ENTREG", REGAF&)
        Call ACTUREGISTRO("MACONSIG", "CANT_RENDID", CONDI$, "CANT_ENTREG", REGAF&)
        Call ACTUREGISTRO("MACONSIG", "CANT_DEVU", CONDI$, "0", REGAF&)
        CONDI$ = "NUME_OCOM=" & 11091 & " AND COD_EXTE='AM-0010'"
        Call ACTUREGISTRO("OCOMDETA", "CANT_REC", CONDI$, 16, REGAF&)
    End If

    'VUELVO A HABILITAR LOS BOTONES DE SELECCION DE PROVEEDOR
    Command1.Enabled = True
    Command10.Enabled = True
    '
End Sub
Sub NOVEDADESAVELLANEDA(Optional ValidarDiferencia%, Optional PRESENTAFORMULARIO%)
     If ValidarDiferencia% > 0 Then
       SQLX$ = "select nume_ocom, Cod_Inte,cant_ocom   from ocomdeta as o WITH(NOLOCK) "            'SELECT CAMPOS DEL OCOMDETA
       SQLX$ = SQLX$ + " where  stat_ocom < 9 "                                                     'QUE NO ESTEN ANULADOS
       SQLX$ = SQLX$ + " and exists(select * from detaoventa as d where o.nume_ocom = d.nro_ocom)"  'QUE EXISTAN LOS NUMEROS DE COMPRA DE OCOMDETA EN ODETAVENTA
       SQLX$ = SQLX$ + " EXCEPT select nro_ocom, Codint,canped from DetaOVenta WITH(NOLOCK)"        'FUNCION EXCEPT COMPARA CON LOS MISMOS CAMPOS EN DETAOVENTA
                                                                                                     'SI NO SON IGUALES LOS TRAE EN EL RECORDSET.
       Set RS = READSET(SQLX$)
       If RS.EOF = True Then Exit Sub                                                               'SI NO HAY DIFERENCIAS SE VA
     End If
     
     'SI LEVANTA LA PANTALLA MUESTRA TODAS LAS ORDENES DE COMPRA  PENDIENTES Y PINTADAS EN AMARILLO LAS QUE TIENEN DIFERENCIAS
     '***************************************************************************************************************************
    'recorro y cargo las ordenes de compra pendientes
    'recorro la grilla y miro mediante la consulta veo las que difiere
    'en esta consulta no van a aparecer el caso que se haya agrgado algo o sea si pedi un item y agregaron uno mas PERO DEBERIA SER IRRELEVANTE
    'LO QUE IMPORTA ES SI LO QUE PIDEN SE LO MANDAN O NO.
    HAYNOVEDADES% = 0
    Campos$ = "P.Repos./f9;Cta.Prov/F8;Razon Social/A20;Referencia/A20;N.V.Central/F12;Suceso/A20;Cambio/F0;Clte. Suc./F11;R.Social/A20;O.C.Clte./A16"
    Call CARGA_ENCABEZADO(NovAve09.GRID, Campos$, NovAve09, , , ANTOT)
    Call CARGA_ENCABEZADO(NovAve09.MSF_2, Campos$, NovAve09, , , ANTOT)

    Campos$ = "N°.: Ped.Reposición/f20;N.de Venta/F12;Código/A20;Descripcion/A24;Solicitud./F10.1;Cant.N.Venta/F12.1;Diferencia/F10.1"
    Call CARGA_ENCABEZADO(NovAve09.MSFDET, Campos$, NovAve09, , , ANTOT)
    '
    SQLX$ = " SELECT  DISTINCT OCOMENCA.Nume_Ocom, OCOMENCA.Refe_Ocom ,OCOMENCA.NUME_PEDI,OCOMENCA.Npro_Ocom,OCOMENCA.RaSo_Prov,OCOMENCA.OComCliente,OCOMENCA.Nume_Cli"
    SQLX$ = SQLX$ + " FROM OCOMENCA "
    SQLX$ = SQLX$ + " INNER JOIN  OCOMDETA "
    SQLX$ = SQLX$ + " ON (OCOMENCA.Nume_Ocom = OCOMDETA.Nume_Ocom) "
    SQLX$ = SQLX$ + " WHERE OCOMDETA.Stat_Ocom < 8 " ' Excluyo las que estan anuladas O COMPLIDAS
    SQLX$ = SQLX$ + " AND OCOMDETA.Stat_Ocom >= 0 " 'EXCLUYO LAS PENDIENTES DE APROBACION
    SQLX$ = SQLX$ + " AND OCOMDETA.Cant_Rec < ( OCOMDETA.Cant_Ocom - 0.05 ) "
    SQLX$ = SQLX$ + " AND OCOMDETA.Cod_Inte > 0 "
    SQLX$ = SQLX$ + " AND OCOMENCA.Stat_Ocom < 9 "
    SQLX$ = SQLX$ + " AND OCOMENCA.Nume_Ocom > 0 "
    SQLX$ = SQLX$ + " AND OCOMENCA.Nume_Pedi > 0 "
    SQLX$ = SQLX$ + " AND OCOMENCA.CODIEMPR = " & CodiEmp%
    SQLX$ = SQLX$ + " AND OCOMDETA.CODIEMPR = " & CodiEmp%
    SQLX$ = SQLX$ + " ORDER BY OCOMENCA.Nume_Ocom DESC "
    Dim OCCUMTEMP As ADODB.Recordset
    Set OCCUMTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With OCCUMTEMP
     If .EOF = True Then
        Call MENSERR(24, "No hay Ordenes de Compra\que Puedan Marcarse como cumplidas.")
        Screen.MousePointer = 1
        Exit Sub
     Else
      '
        J& = 0
        Do While (.EOF = False)
           NUOC& = XVALO(!nume_ocom)
           J& = J& + 1
           NPEDX& = XVALO(!Nume_Pedi)
           NovAve09.GRID.Rows = J& + 1
           NovAve09.GRID.TextMatrix(J&, 1) = NUOC&
           NovAve09.GRID.TextMatrix(J&, 2) = XVALO(!Npro_Ocom)
           NovAve09.GRID.TextMatrix(J&, 3) = SAFETEXT$(!Raso_Prov)
           NovAve09.GRID.TextMatrix(J&, 4) = SAFETEXT$(!Refe_OcoM)
            NovAve09.GRID.TextMatrix(J&, 5) = SAFETEXT$(NPEDX&)
            NCNEUEQUEN = XVALO(!nume_cli)
            OComCliente = SAFETEXT$(!OComCliente)
              CONDI$ = " nroped = " & NPEDX&
              SUCESOX$ = VALOBADA("DetaOVenta", "ACCION", CONDI$, "NOMESS")
              SUCESOX1$ = ""
              Select Case SUCESOX$
                Case "MODIPEDI"
                  SUCESOX1$ = "N.Venta Modificada"
                Case "PEDIDOANUL"
                  SUCESOX1$ = "N.Venta Anulada"
               Case "GENEPEDI"
                  SUCESOX1$ = "N.Venta Generada"
              End Select
           NovAve09.GRID.TextMatrix(J&, 6) = SUCESOX1$
           NovAve09.GRID.TextMatrix(J&, 7) = "0"
           NovAve09.GRID.TextMatrix(J&, 8) = NCNEUEQUEN
           NovAve09.GRID.TextMatrix(J&, 9) = rasocli$(NCNEUEQUEN)
           NovAve09.GRID.TextMatrix(J&, 10) = OComCliente
           '
           'VERIFICO SI HUBO CAMBIOS CON ESTA INSTRUCCION EXCEPT QUE COMPARA LOS REGSISTROS
           
           SQLX$ = "select Cod_Inte,cant_ocom from ocomdeta WITH(NOLOCK) where STAT_OCOM < 9 AND NUME_OCOM = " & NUOC&
           SQLX$ = SQLX$ + " EXCEPT "
           SQLX$ = SQLX$ + " select Codint,canped from DetaOVenta WITH(NOLOCK)  where nroped = " & NPEDX&
           Set RS1 = READSET(SQLX$)
           With RS1
             If Not RS1.EOF Then
                Call COLOREAR_GRILLA_SOLO_UNA_AFILA(NovAve09.GRID, vbYellow, J&)
                NovAve09.GRID.TextMatrix(J&, 7) = "1"
                HAYNOVEDADES% = 1
             End If
           End With
           RS1.Close
           Set RS1 = Nothing
           .MoveNext
         Loop
     End If
    End With
    OCCUMTEMP.Close
    Set OCCUMTEMP = Nothing
    
    If PRESENTAFORMULARIO% > 0 Or HAYNOVEDADES% > 0 Then
       NovAve09.Show 1
    End If
End Sub
Sub ACTUALIZAR_ARCHIVOS_DE_CENTRAL()

    Archivo$ = "ARCHIVO_TRANSF_EXCEPCION"
    Call OPCICOM07.RECIBIRARCHIVO(Archivo$, RETORNOOK%, "NOMESS")
    If RETORNOOK% < 0 Then Exit Sub
    '
    Archivo$ = "ARCHIVO_TRANSF_MIN_MAX"
    Call OPCICOM07.RECIBIRARCHIVO(Archivo$, RETORNOOK%, "NOMESS")
    If RETORNOOK% < 0 Then Exit Sub

    Archivo$ = "ARCHIVO_GENERO_PEDIDO"
    Call OPCICOM07.RECIBIRARCHIVO(Archivo$, RETORNOOK%, "NOMESS")
    If RETORNOOK% < 0 Then Exit Sub

End Sub
'
Sub CARGAOCPRE()
    TAPLI$ = Mid$(APLIXYZ$, 5)
    On Error Resume Next
    Call TEXTOLOTES$(TAPLI$, "_")
    NPROV = XVALO(CADENATEX$(1))
    NPRES& = XVALO(CADENATEX$(2))
    On Error GoTo 0
    NSOLICITUD = NumeroSolicitud(NPROV, NPRES&)
    Text1.TEXT = SAFETEXT$(NPROV)
    '
    Call BLANFORMU
    Text1.TEXT = TRIM$(Str$(NPROV))
    Call CARDACLI
    Call CARCOANT
    '
    Call ABRECOMPRAS
    Label11 = DEMONECO$(MONEMI%)
    TICAMBIS = TRIM$(FORMATNUM(TICAMONE(MONEMI%), "F10.4"))
    '
    SQLX$ = "SELECT * FROM PRESUCOMPRA WITH(NOLOCK) WHERE NroPres = " & NPRES&
    SQLX$ = SQLX$ + " AND Num_Prov = " & NPROV
    SQLX$ = SQLX$ + " ORDER BY ORDEN"
    Set COTRE = READSET(SQLX$)
    '
    JJ& = 0
    REBLA$ = REGBLAN$("OCOMDETA")
    Call SETULTREG("!OCOMDETA", 0)
    ARCHILIB$ = "!1OCOMDET"
    Call SETULTREG(ARCHILIB$, 0)
    '
    With COTRE
       Do While Not .EOF
         '
         If VUELTA% < 1 Then
            VUELTA% = 1
            MONEMI% = XVALO(!MONE_COS) ' carga la moneda de emision de la cotizacion
            Label11 = DEMONECO$(MONEMI%)
            TICAMBIS = TRIM$(FORMATNUM(TICAMONE(MONEMI%), "F10.4"))
         End If
         '
         CI1% = XVALO(!CODINT)
         DECRI$ = rasocli$(-1 * NPROV)
         UMECO$ = UNICOM$(CI1%)
         PreUnid# = XVALO(!costo) 'carga precio unitario del articulo
         CANTU = XVALO(!Cantidad)
         '
         ZZ$ = REBLA$
         Call REPLA(ZZ$, MKI$(CI1%), 9, 2)
         Call REPLA(ZZ$, Unimed$(CI1%), 11, 4)
         Call REPLA(ZZ$, MKS$(CANTU), 51, 4)
         Call REPLA(ZZ$, UMECO$, 59, 4)
         Call REPLA(ZZ$, MKS$(1), 63, 4)
         Call REPLA(ZZ$, MKD$(PreUnid#), 67, 8)
         Call REPLA(ZZ$, 0, 75, 4)
         Call REPLA(ZZ$, MKD$(PreUnid#), 79, 8)
         Call REPLA(ZZ$, MKD$(PreUnid# * CANTU), 87, 8)
            SoliciYNPres$ = SAFETEXT$(NSOLICITUD)
         Call REPLA(ZZ$, SoliciYNPres$, 31, 20)
         '
         JJ& = JJ& + 1
         Call GRAREG("!OCOMDETA", ZZ$, JJ&)
         Call GRAREG(ARCHILIB$, DECRI$, JJ&)
       .MoveNext
       Loop
    End With
    Call SETULTREG("!OCOMDETA", JJ&)
    Call SETULTREG(ARCHILIB$, JJ&)
    '
    Call CARDETCOM
    '
End Sub
Sub CARGALACOMPRA()
    TAPLI$ = Mid$(APLINVO$, 5)
    NUSOLI& = XVALO(Left$(TAPLI$, 8))
    ORSOLI& = XVALO(Mid$(TAPLI$, 9, 4))
    NCLIE% = XVALO(TRIM$(Mid$(TAPLI$, 13, 6))) ' NUMERO DE PROVEEDOR
    ' Se agregar para carpeta de importacion
    Label18 = TRIM$(Mid$(TAPLI$, 19))
    If Label18 <> "" Then
        Command1.Enabled = False
        Text1.Enabled = False
        Label8.Enabled = False
    End If
    '
    Call BLANFORMU
    Text1.TEXT = TRIM$(Str$(NCLIE%))
    Call CARDACLI
    '  Call CARCOPEN     ' DESHABILITADO PARA QUE NO MUESTRE LAS COMPRAS PENDIENTES
    Call CARCOANT
    '
    Call ABRECOMPRAS
    CONDISO$ = "NUME_SOLINT = " & NUSOLI& & " AND NUORD_ITEM = " & ORSOLI& & " "
    CONDI$ = "NUME_SOLINT = " & NUSOLI& & " AND NUORD_ITEM = " & ORSOLI& & " AND NPRO_OCOM = " & NCLIE%
    '
    CANTU = XVALO(VALOBADA("DETASOLINT", "Cant_Solint", CONDISO$))
    FENTREIT% = CVINT(VALOBADA("DETASOLINT", "Fentre_Sol", CONDISO$))
    Text7 = VALOBADA("DETASOLINT", "Iconta_Solint", CONDISO$)
    Text18 = VALOBADA("ENCASOLINT", "Usua_Gen", "NUME_SOLINT = " & NUSOLI& & " ")
   
    MONEMI% = XVALO(VALOBADA("COTIRECI", "MONE_EMIS", CONDI$)) ' carga la moneda de emision de la cotizacion
    Label11 = DEMONECO$(MONEMI%)
    TICAMBIS = TRIM$(FORMATNUM(TICAMONE(MONEMI%), "F10.4"))
    '
    Text19 = XVALO(VALOBADA("COTIRECI", "PLAZ_ENTRE", CONDI$)) ' carga el plazo de vigencia de la cotizacion
    Text13 = VALOBADA("COTIRECI", "OBSERVA", CONDI$)
    '
    CI0% = XVALO(VALOBADA("COTIRECI", "Cod_Inte", CONDI$))
    DECRI$ = VALOBADA("COTIRECI", "Descritem", CONDI$)
    UMECO$ = VALOBADA("COTIRECI", "Uni_Sol", CONDI$)

    PRUNI = XVALO(VALOBADA("COTIRECI", "PRE_UNIT", CONDI$)) 'carga precio unitario del articulo
    PorcDescuento$ = VALOBADA("COTIRECI", "DESCUENTO", CONDI$)
    If MONEMI% > 1 Then
       PRUNI = PRUNI * TICAMONE(MONEMI%)
    End If
    JJ& = 0
    REBLA$ = REGBLAN$("OCOMDETA")
    Call SETULTREG("!OCOMDETA", 0)
    ARCHILIB$ = "!1OCOMDET"
    Call SETULTREG(ARCHILIB$, 0)
    '
    ZZ$ = REBLA$
    Call REPLA(ZZ$, MKI$(CI0%), 9, 2)
    Call REPLA(ZZ$, Unimed$(CI0%), 11, 4)
    Call REPLA(ZZ$, MKS$(CANTU), 51, 4)
    Call REPLA(ZZ$, UMECO$, 59, 4)
    Call REPLA(ZZ$, MKS$(1), 63, 4)
    Call REPLA(ZZ$, MKI$(FENTREIT%), 95, 2)
    Call REPLA(ZZ$, Text18, 15, 12)
    Call REPLA(ZZ$, Text7, 105, 12)
       NUSI$ = TRIM$(Str$(NUSOLI&)) + "/" + TRIM$(Str$(ORSOLI&))
    Call REPLA(ZZ$, NUSI$, 31, 20)
    '
    PreUnid# = PRUNI / TICAMONE(MONEMI%)
    Call REPLA(ZZ$, MKD$(PreUnid#), 67, 8)
    COEFDESC = COEFDESCU(PorcDescuento$)
    Call REPLA(ZZ$, MKS$(COEFDESC), 75, 4)
    Call REPLA(ZZ$, MKD$(PreUnid#), 79, 8)
    Call REPLA(ZZ$, MKD$(PreUnid# * CANTU), 87, 8)
    '
    JJ& = JJ& + 1
    Call GRAREG("!OCOMDETA", ZZ$, JJ&)
    Call GRAREG(ARCHILIB$, DECRI$, JJ&)
    Call SETULTREG("!OCOMDETA", JJ&)
    '
    Call CARDETCOM
    '
End Sub

Sub CARGACOMUL()
    TAPLI$ = Mid$(APLINVO$, 5)
    NroProveedor% = XVALO(TRIM$(Mid$(TAPLI$, 1, 6))) ' NUMERO DE PROVEEDOR
    NroSolicitud& = XVALO(TRIM$(Mid$(TAPLI$, 7, 6))) ' NUMERO DE SIM
    '
    Command1.Enabled = False
    Text1.Enabled = False
    Label8.Enabled = False
    '
    Call BLANFORMU
    Text1.TEXT = TRIM$(Str$(NroProveedor%))
    Call CARDACLI
    Call CARCOANT
    '
    Call ABRECOMPRAS
    Label11 = DEMONECO$(MONEMI%)
    TICAMBIS = TRIM$(FORMATNUM(TICAMONE(MONEMI%), "F10.4"))
    '
    Dim COTRE As ADODB.Recordset
    Set COTRE = New ADODB.Recordset
    SQLX$ = "SELECT * FROM COTIRECI WHERE NUME_OCOM = 0 AND NPRO_OCOM = " & CStr(NroProveedor%) & " "
    If NroSolicitud& > 0 Then SQLX$ = SQLX$ & " AND NUME_SOLINT=" & CStr(NroSolicitud&)
    Set COTRE = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    JJ& = 0
    REBLA$ = REGBLAN$("OCOMDETA")
    Call SETULTREG("!OCOMDETA", 0)
    ARCHILIB$ = "!1OCOMDET"
    Call SETULTREG(ARCHILIB$, 0)
    '
    With COTRE
       Do While Not .EOF
         CI0% = XVALO(!Cod_Inte)
         DECRI$ = SAFETEXT$(!DESCRITEM)
         UMECO$ = SAFETEXT$(!Uni_Sol)
         PRUNI = XVALO(!Pre_Unit) 'carga precio unitario del articulo
         If MONEMI% > 1 Then
            PRUNI = PRUNI * TICAMONE(MONEMI%)
         End If
         '
         NUSOLI& = XVALO(!NUME_SOLINT)
         ORSOLI& = XVALO(!NuOrd_Item)
         PorcDescuento$ = SAFETEXT$(!DESCUENTO)
         CONDISO$ = "NUME_SOLINT = " & NUSOLI& & " AND NUORD_ITEM = " & ORSOLI& & " "
         CONDI$ = "NUME_SOLINT = " & NUSOLI& & " AND NUORD_ITEM = " & ORSOLI& & " AND NPRO_OCOM = " & NroProveedor%
         '
         ZZ$ = REBLA$
         Call REPLA(ZZ$, MKI$(CI0%), 9, 2)
         Call REPLA(ZZ$, Unimed$(CI0%), 11, 4)
             CANTU = XVALO(VALOBADA("DETASOLINT", "Cant_Solint", CONDISO$))
         Call REPLA(ZZ$, MKS$(CANTU), 51, 4)
         Call REPLA(ZZ$, UMECO$, 59, 4)
         Call REPLA(ZZ$, MKS$(1), 63, 4)
             FENTREIT% = CVINT(VALOBADA("DETASOLINT", "Fentre_Sol", CONDISO$))
         Call REPLA(ZZ$, MKI$(FENTREIT%), 95, 2)
             Text18 = VALOBADA("ENCASOLINT", "Usua_Gen", "NUME_SOLINT = " & NUSOLI& & " ")
         Call REPLA(ZZ$, Text18, 15, 12)
             Text7 = VALOBADA("DETASOLINT", "Iconta_Solint", CONDISO$)
         Call REPLA(ZZ$, Text7, 105, 12)
               NUSI$ = TRIM$(Str$(NUSOLI&)) + "/" + TRIM$(Str$(ORSOLI&))
         Call REPLA(ZZ$, NUSI$, 31, 20)
         '
         PreUnid# = PRUNI / TICAMONE(MONEMI%)
         Call REPLA(ZZ$, MKD$(PreUnid#), 67, 8)
         COEFDESC = COEFDESCU(PorcDescuento$)
         Call REPLA(ZZ$, MKS$(COEFDESC), 75, 4)
'         Call REPLA(ZZ$, 0, 75, 4)
         Call REPLA(ZZ$, MKD$(PreUnid#), 79, 8)
         Call REPLA(ZZ$, MKD$(PreUnid# * CANTU), 87, 8)
         '
         JJ& = JJ& + 1
         Call GRAREG("!OCOMDETA", ZZ$, JJ&)
         Call GRAREG(ARCHILIB$, DECRI$, JJ&)
       .MoveNext
       Loop
    End With
    Call SETULTREG("!OCOMDETA", JJ&)
    Call SETULTREG(ARCHILIB$, JJ&)
    Call CIERRARCH("!OCOMDETA")
    Call CIERRARCH(ARCHILIB$)
    '
    Call CARDETCOM
    '
End Sub

'
Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call LIBARCH("*.*")
   Call FINAL("")
End Sub



Private Sub Label11_DblClick()
   '
   Static CAMON%, DMONE$(), TCMONE()
   If CAMON% < 1 Then
     CAMON% = ULTREG&("TAMONED")
     If CAMON% < 1 Then
         CAMON% = 1
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         DMONE$(1) = "Pesos"
         TCMONE(1) = 1
       Else
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         For U& = 1 To CAMON%
           XX$ = REGLEIDO$("TAMONED", U&)
           DMX$ = TRIM$(Left$(XX$, 8)) + " "
           PPXX% = InStr(DMX$, " ")
           DMX$ = Left$(DMX$, PPXX% - 1)
           '
           VHX = CVS(Mid$(XX$, 29, 4)): If VHX < 0.005 Then VHX = 1
           DMONE$(U&) = TRIM$(DMX$)
           TCMONE(U) = VHX
         Next U&
     End If
   End If
   '
   MONEMIA% = MONEMI%
15 MONEMI% = MONEMI% + 1
   If MONEMI% > CAMON% Then MONEMI% = 1
   If DMONE$(MONEMI%) = "" Then GoTo 15
   Label11 = DMONE$(MONEMI%)
   Label11.ForeColor = RGB(0, 0, 255)
   If MONEMI% > 1 Then
      Label11.ForeColor = RGB(255, 0, 0)
   End If
   Label11.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
   TICAMBIS = TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4"))
   '
   COEFCON = TICAMONE(MONEMIA%) / TICAMONE(MONEMI%)
   TICAMBX# = TICAMONE(MONEMI%)
   '
   For U& = 1 To ULTREG&("!OCOMDETA")
     YX$ = REGLEIDO$("!OCOMDETA", U&)
     PREDOL# = CVD(Mid$(YX$, 67, 8)) * COEFCON
     Call REPLA(YX$, MKD$(PREDOL#), 67, 8)
     Call GRAREG("!OCOMDETA", YX$, U&)
   Next U&
   Call CIERRARCH("!OCOMDETA")
   Call CARLISCO
   Call CALTOCOM
   '
End Sub

Private Sub Label14_dblClick()
   If VALICLIE% < 1 Then
     On Error Resume Next
     Text1.SetFocus
     On Error GoTo 0
     Exit Sub
   End If
   '
   Call CARDETCOM1
End Sub

Private Sub Label2_DblClick()
   Call Label14_dblClick
End Sub

Private Sub Label22_DblClick()
   List1.Clear
   Call SETULTREG("!OCOMDETA", 0)
   Unload FORVENTB
   Call CARDETCOM
End Sub

Private Sub Label8_Click()
   Call SELECHO("INDIACRE")
   CPROX$ = TRIM$(VALACT1$("INDIACRE"))
   If CPROX$ <> "" Then
     NCLIE% = NUMEPRO%(CPROX$)
     If NCLIE% > 0 Then
       If REGICLI&((-1) * NCLIE%) > 0 Then
         Call BLANFORMU
         Text1.TEXT = TRIM$(Str$(NCLIE%))
         Call CARDACLI
         Call CARCOPEN
         Call CARCOANT
         If ULTREG&("!OCOMDETA") > 0 Then
              Call CARLISCO
              Call CALTOCOM
           Else
              Call CARDETCOM
         End If
       End If
     End If
   End If
End Sub

Private Sub Label8_DblClick()
   Call Label8_Click
End Sub

Private Sub List1_Click()
   List1.ToolTipText = DESCRIT$(CODINT%(Left$(List1.TEXT, 15)))
   MUECOPEN% = 1
   For U& = 1 To List2.ListCount
     List2.Selected(U& - 1) = False
     If Left$(List2.List(U& - 1), 15) = Left$(List1.TEXT, 15) Then
       List2.Selected(U& - 1) = True
     End If
   Next U&
   MUECOPEN% = 0
End Sub

Private Sub List1_DblClick()
    Call CARDETCOM
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
     Call MUESTRADATCOD(Left$(List1.TEXT, 15))
   End If
   '
End Sub

Private Sub List2_Click()
   If MUECOPEN% < 1 Then
     For U& = 1 To List2.ListCount
       List2.Selected(U& - 1) = False
     Next U&
   End If
End Sub

Private Sub List2_DblClick()
    NPX& = XVALO(Mid$(List2.TEXT, 79, 6))
    If NPX& > 0 Then
      Call SHOWOCOM(NPX&)
    End If
End Sub

Private Sub Ltota_DblClick(Index As Integer)
   NumeroProveedor& = XVALO(Text1) * (-1)
   PosicionIva% = PIVACLI%(NumeroProveedor&)
   If Index = 2 And PosicionIva% < 4 Then
      RIVACOM07.Show 1
      Label118 = TRIM$(RIVACOM07.Label2)
      Total(2) = TRIM$(RIVACOM07.Label3)
      Call CALTOCOM
   End If
'   For KKI% = 0 To 4
'     Total(KKI%).Enabled = Not (Total(KKI%).Enabled)
'   Next KKI%
End Sub

Private Sub MarItCum_Click()
  Call OPCICOM07.Command4_Click
End Sub

Private Sub Amastock_Click()
    ACONEC$ = "ARCHIG07/AMASTO"
    Call CONECRUN(ACONEC$, "Maestro de Artículos de Stock")
End Sub

Private Sub MnuEstad_Click()
   Call AMASTOK07.Command6_Click
End Sub

Private Sub MnuNueOCAbi_Click()
   Call CIERRARCH("*.*")
   ACONEC$ = "OCOABI07"
   Call CONECRUN(ACONEC$, "Ordenes de Compra Abiertas")
   Call FINAL("")
  '\\\OT-06-0149-WAR-FIN///
End Sub

Private Sub MnuOCVSAV_Click()
  Call NOVEDADESAVELLANEDA(0, 1)
End Sub

Private Sub Modi_Ocom_Click()
   Call SETULTREG("!OCOMDETA", 0)
   Call BLANFORMU
   Call OPCICOM07.Command13_Click
End Sub

Private Sub Nuorig_Click()
   Call OPCICOM07.Command19_Click
End Sub

Private Sub Nuorig1_Click()
   Call OPCICOM07.Command19_Click
End Sub

Private Sub ocompendaprobastocksypedidos_Click()
   Call FINAL("?OCOMP-ARC")
End Sub

Private Sub Option1_Click()
   If TEXT24.Enabled = True Then
    TEXT24.Enabled = Not (TEXT24.Enabled)
    Eco24.Enabled = Not (Eco24.Enabled)
    TEXT25.Enabled = Not (TEXT25.Enabled)
    Command25.Enabled = Not (Command25.Enabled)
   End If

End Sub

Private Sub Option2_Click()
   If TEXT24.Enabled = False Then
    TEXT24.Enabled = Not (TEXT24.Enabled)
    Eco24.Enabled = Not (Eco24.Enabled)
    TEXT25.Enabled = Not (TEXT25.Enabled)
    Command25.Enabled = Not (Command25.Enabled)
   End If
End Sub

Private Sub pedreppendiente_Click()
   Call FINAL("?OCOMP-NEU2")
End Sub

Private Sub Por_MRP_Click()
   Call Command10_Click
End Sub

Private Sub PORDESCU_Change()
   Call CALTOCOM
End Sub

Private Sub PORDESCU_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
End Sub

Private Sub Print1_Click()
   OPLISCO07.Option2.Value = True
   Call OPLISCO07.Command3_Click
End Sub

Private Sub Print2_Click()
   OPLISCO07.Option2.Value = True
   Call OPLISCO07.Command4_Click
End Sub

Private Sub Print3_Click()
   OPLISCO07.Option2.Value = True
   Call OPLISCO07.Command5_Click
End Sub

Private Sub Print4_Click()
   OPLISCO07.Option2.Value = True
   Call OPLISCO07.Command9_Click
End Sub

Private Sub Provedor_Click()
   ACONEC$ = "ARCHIG07/AMAPRO"
   Call CONECRUN(ACONEC$, "Padron Maestro de Proveedores")
End Sub

Private Sub repocliente_Click()
   Call FINAL("?OCOMP-NEU1")
End Sub

Private Sub Screen1_Click()
   OPLISCO07.Option1.Value = True
   Call OPLISCO07.Command3_Click
End Sub

Private Sub Screen2_Click()
   OPLISCO07.Option1.Value = True
   Call OPLISCO07.Command4_Click
End Sub

Private Sub Screen3_Click()
   OPLISCO07.Option1.Value = True
   Call OPLISCO07.Command5_Click
End Sub

Private Sub Screen4_Click()
   OPLISCO07.Option1.Value = True
   Call OPLISCO07.Command9_Click
End Sub

Private Sub Setup_Click()
  OPCICOM07.Command18_Click
End Sub



Private Sub Stockminymax_Click()
  Call FINAL("?stockMinmax")
End Sub

Private Sub Text1_DblClick()
  Call SELECHO("PROVEPLA")
  NCLIE% = XVALO(TRIM$(VALACT1$("PROVEPLA")))
  If NCLIE% > 0 Then
    If REGICLI&((-1) * NCLIE%) > 0 Then
      Call BLANFORMU
      Text1.TEXT = TRIM$(Str$(NCLIE%))
      Call CARDACLI
      Call CARCOPEN
      Call CARCOANT
      If ULTREG&("!OCOMDETA") > 0 Then
           Call CARLISCO
           Call CALTOCOM
        Else
           Call CARDETCOM
      End If
    End If
  End If
End Sub

Private Sub Text1_GotFocus()
   Text1.TEXT = TRIM$(Text1.TEXT)
   Text1.SelStart = 0
   Text1.SelLength = Len(Text1.TEXT)
   '
   If MONEMI% < 1 Then
     Label11.ForeColor = RGB(0, 0, 255)
     Label11 = "Pesos"
     Label11.Refresh
     MONEMI% = 1
     Label11.ToolTipText = "Doble-Click para 'Dólares'"
   End If
   '
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If VALICLIE% > 0 Then
      Call BLANFORMU
      Call CARDACLI
      Call CARCOPEN
      Call CARCOANT
      If ULTREG&("!OCOMDETA") > 0 Then
           Call CARLISCO
           Call CALTOCOM
         Else
           Call CARDETCOM
      End If
    End If
  End If
End Sub

Private Sub Text10_DblClick()
   OBSERVAC.Caption = "Datos de Entrega"
   OBSERVAC.Label24.Caption = ""
   OBSERVAC.Label2 = 128
   OBSERVAC.Text28.MaxLength = 128
   OBSERVAC.Text28 = Text10

   OBSERVAC.Show 1
   
   If OBSERVAC.Label1 = "OK" Then
     Text10 = OBSERVAC.Text28
   End If


End Sub

Private Sub Text10_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
   End If
End Sub

Private Sub Text11_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   If ULTREG&("PEDENCA") > 2 Then
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
   NPP& = PRONUOCO&
   If XVALO(NUMEPRE$) < NPP& Then
     Call MENSERR(24, "Número no Válido")
     Text11.TEXT = TRIM$(Str$(NPP&))
     Exit Sub
   End If
   '
   If NUMEPRE$ <> TRIM$(Text11.TEXT) Then
     Text11.TEXT = NUMEPRE$
     Call MENSERR(24, "Número de Presupuesto Ajustado")
   End If
   '
End Sub

Private Sub Text12_Change()
   Call CALTOCOM
End Sub

Private Sub Text12_DblClick()
   OBSERVAC.Caption = "Datos de Contacto"
   OBSERVAC.Label24.Caption = ""
   OBSERVAC.Label2 = 128
   OBSERVAC.Text28.MaxLength = 128
   OBSERVAC.Text28 = Text12

   OBSERVAC.Show 1
   
   If OBSERVAC.Label1 = "OK" Then
     Text12 = OBSERVAC.Text28
   End If


End Sub


Private Sub Text12_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
   End If
End Sub


Private Sub Text13_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
   End If
End Sub

Private Sub Text15_Change()
   Text16 = ECOARCH$("TAMONED", Chr$(XVALO(Text15)))
End Sub

Private Sub Text17_Change()
   Call CALTOCOM
End Sub

Private Sub Text17_GotFocus()
   If VALICLIE% < 1 Then
     On Error Resume Next
     Text1.SetFocus
     On Error GoTo 0
     Exit Sub
   End If
End Sub

Private Sub Text19_Change()
    '
    If VALICLIE% < 1 Then
       On Error Resume Next
       Text19 = ""
       Text1.SetFocus
       On Error GoTo 0
       Exit Sub
    End If
    '
'    FEMIS% = FECHANUM(Text9)
'    FENTREG% = FEMIS%
'    '
'    If Text19 <> "" Then
'     If IsNumeric(Text19) = True Then
'      While DIENTRE%(FEMIS%, FENTREG%) < CInt(Text19)
'         FENTREG% = DIPOST(FENTREG)
'      Wend
'      Text14 = FECHATEX$(FENTREG%)
'     End If
'    End If
    '\\\OT-05-0690-WAR-24/10/05///
    '
    FEMIS% = FECHANUM(Text9)
    FENTREG% = FEMIS%
    '
    If Text19 <> "" Then
      If IsNumeric(Text19) = True And XVALO(Text19) < 7300 Then 'MENOR A 20 AÑOS EN DIAS
        While DIENTRE%(FEMIS%, FENTREG%) < CInt(Text19)
          FENTREG% = DIPOST(FENTREG)
        Wend
        Text14 = FECHATEX$(FENTREG%)
      Else
        Call MENSERR(24, "Cantidad de Días no Válida.")
      End If
    Else
      HOII% = HOY(HOS$)
      Text14.TEXT = HOS$
    End If
    '\\\OT-05-0690-WAR-FIN///
    '
End Sub
Private Sub Text2_gotfocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub TICAMBIS_LostFocus()
   TICAMBIS = FORMATNUM(TICAMBIS, "F7.4")
End Sub

Private Sub TEXT21_DblClick(Index As Integer)
    '
    If VALICLIE% < 1 Then
       On Error Resume Next
       Text1.SetFocus
       On Error GoTo 0
       Exit Sub
    End If
    '
    If Index = 0 Then
        ARCHEC$ = "CONPAG": LOCA% = 1
      ElseIf Index = 1 Then
        ARCHEC$ = "TAEMBALA": LOCA% = 1
    End If
    ARGU$ = CBIN$(TRIM$(TEXT21(Index).TEXT), 1, LOCA%)
    Call SELECHO(ARCHEC$)
    VDEVU$ = TRIM$(VALACT1$(ARCHEC$))
    If VDEVU$ <> "" Then
       TEXT21(Index).TEXT = VDEVU$
    End If
    '
End Sub

Private Sub Text21_GotFocus(Index As Integer)
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      Exit Sub
   End If
   TEXT21(Index).TEXT = TRIM$(TEXT21(Index).TEXT)
   TEXT21(Index).SelStart = 0
   TEXT21(Index).SelLength = Len(TEXT21(Index).TEXT)
End Sub

Private Sub Text22_Change()
   Text13 = Text22
End Sub

Private Sub Text24_Change()
   Eco24 = rasocli$(XVALO(TEXT24))
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
    If XVALO(TEXT21(Index).TEXT) < 0 Or XVALO(TEXT21(Index).TEXT) > 255 Then
       TEXT21(Index).TEXT = ""
    End If
    '
    If Index = 0 Then
        Eco21(Index).TEXT = ECOARCH$("CONPAG", Chr$(XVALO(TEXT21(Index).TEXT)))
      ElseIf Index = 1 Then
        Eco21(Index).TEXT = ECOARCH$("TAEMBALA", Chr$(XVALO(TEXT21(Index).TEXT)))
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
        ARCHEC$ = "CONPAG": LOCA% = 1
      ElseIf Index = 1 Then
        ARCHEC$ = "TAEMBALA": LOCA% = 1
    End If
    ARGU$ = CBIN$(TRIM$(TEXT21(Index).TEXT), 1, LOCA%)
    Call SELECHO(ARCHEC$)
    VDEVU$ = TRIM$(VALACT1$(ARCHEC$))
    If VDEVU$ <> "" Then
       TEXT21(Index).TEXT = VDEVU$
    End If
    '
End Sub
   '
Function VALICLIE%()
   VCC% = 0
   If XVALO(Text1.TEXT) > 0 Then
     If XVALO(Text1.TEXT) <= 32767 Then
       NCLIE% = XVALO(Text1.TEXT)
       NPPX% = XVALO(VALOBADA("PROVED12", "Nume_Cli", "stat_cli >= 0 and Nume_Cli = " & NCLIE%))
       If NPPX% > 0 Then
         VCC% = 1
       End If
     End If
   End If
   '
99 VALICLIE% = VCC%
   '
End Function

Sub BLANFORMU()
   EPAC$ = TRIM$(UCase$(EMPREAC$))

   Call SETULTREG("!OCOMDETA", 0)
   Call CIERRARCH("!OCOMDETA")
   Call BLANARCH("!OCOMDETA")
   If EXISTE%(LUDAT$ + "1OCOMDET.RFS") > 0 Then
     Call SETULTREG("!1OCOMDET", 0)
   End If
   Call CIERRARCH("!1OCOMDET")
   Call BLANARCH("!1OCOMDET")
   On Error Resume Next
   Unload FORVENTB
   On Error GoTo 0
   Label15 = ""
   Label8 = ""
   Text2.TEXT = ""
   Text3.TEXT = ""
   Text4.TEXT = ""
   Text5.TEXT = ""
   'Label2 = ""
   Text8.TEXT = ""
   Text10 = ""
   Text12.TEXT = ""
   Label118 = FORMATNUM$(ALIVA, "F5.2")
   Text13.TEXT = ""
   Text6.TEXT = ""
   Text7.TEXT = ""
   Text18.TEXT = ""
   Text19.TEXT = ""
   List1.Clear: List1.ToolTipText = ""
   List2.Clear
   '
   TEXT21(0).TEXT = "": Eco21(0) = ""
   TEXT21(1).TEXT = "": Eco21(1) = ""
   'TEXT21(2).TEXT = "": Eco21(2) = ""
   '
   PORDESCU.TEXT = ""
   Text17 = ""
   For KKI% = 0 To 4
      Total(KKI%).Enabled = True
      Total(KKI%).TEXT = ""
   Next KKI%
   '
   'MARCONOTA.Visible = False
   HOII% = HOY(HOS$)
   Text9.TEXT = HOS$
   Text11.TEXT = TRIM$(Str$(PRONUOCO&))
   '
   Text22 = ""
   TEXT24 = ""
   TEXT25 = ""
   '
   If InStr(EPAC$, "NEUQUEN") < 1 Then 'SI NO ES NEUQUEN QUITA LA VENTNA DE PEDIDO DE REPOSICION, SI NO LA MANTIENE
      Frame2.Visible = False
   End If
End Sub
'
Sub CARDACLI()
    '
    Screen.MousePointer = 11
    '
    NC% = (-1) * XVALO(Text1.TEXT)
    If NC% < 0 Then
      If STATCLI%(NC%) > 0 Then
        Label8 = CODICLI$(NC%)
        Text2.TEXT = rasocli$(NC%)
        Text3.TEXT = DOMICLI$(NC%)
        Text4.TEXT = CPOSCLI$(NC%)
        Text5.TEXT = LOCACLI$(NC%)
        Text6 = CONTACTCLI$(NC%)
        'Label2 = TRIM$(POSIVAC$(NC%))
        Text8.TEXT = TRIM$(CuitCli$(NC%))
        '
        TEXT21(0).TEXT = Str$(CPAGCLI%(NC%))
        TEXT21(1).TEXT = ""
        '
        'POR DEFECTO EL LUGAR DE ENTREGA EL CARGADO EN EL PROVEEDOR SI DEVUELVE VACIO
        'CARGA EL DEL CONTROL
        Text10 = LUENTRECOMPRA(NC%)
        If TRIM$(Text10) = "" Then
          Text10 = Control$("OCOMPRA", "LENTREGA")
        End If
        '
        'OBSERVACIONES DE COMPRA CARGADA PARA ESTE PROVEEDOR INGRESADO EN MAESTRO DE PROVEEDORES
        Text13 = OBSERVACOMPRA(NC%)
        '
        Text15 = TRIM$(Str$(MONEVACO%(NC%)))
        If MONEMI% < 1 Then
          Label11.ForeColor = RGB(0, 0, 255)
          Label11 = "Pesos"
          MONEMI% = 1
          Label11.ToolTipText = "Doble-Click para 'Dólares'"
        End If
        '
        If MONEMICO%(NC%) = 2 Then
          Label11.ForeColor = RGB(255, 0, 0)
          Label11 = "Dólares"
          MONEMI% = 2
          Label11.ToolTipText = "Doble-Click para 'Pesos'"
        End If
        '
        Call APERBASDAT("COMPRAS")
        SQLX$ = "SELECT Nume_Ocom, Femi_Ocom, AtSr_Ocom, Pdes_Ocom FROM OcomEnca "
        SQLX$ = SQLX$ + "WHERE Npro_Ocom = " & Abs(NC%) & " "
        SQLX$ = SQLX$ + "ORDER BY Femi_Ocom DESC;"
        'Set OEncaTemp = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
        Dim OEncaTemp As ADODB.Recordset
        Set OEncaTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
        With OEncaTemp
          On Error Resume Next
          .MoveFirst
          If Err < 1 Then
            ATEN$ = !AtSr_Ocom: If TRIM$(ATEN$) <> "" Then Text6 = ATEN$
            PODESC$ = !Pdes_Ocom: PORDESCU = PODESC$
          End If
        End With
        PORDESCUSTA$ = TRIM$(DESCUCLI(NC%))
        If PORDESCUSTA$ <> "" Then PORDESCU = PORDESCUSTA
        '
        If EXISTE%(LUDAT$ + "CALINTEG.RFS") > 0 Then
          EPAC$ = TRIM$(UCase$(EMPREAC$))
          If InStr(EPAC$, "DOSIVA") > 0 Or InStr(EPAC$, "ARCON") > 0 Then
            XX11$ = FILTERGETRECORD$("CALINTEG", 1, MKI$((-1) * NC%))
            PUNTAPRO = CVS(Mid$(XX11$, 35, 4))
            Label15 = FORMATNUM$(PUNTAPRO, "F4.1")
            GoTo 10
          End If
          '
          ' SI NO ENTRO POR DOSIVAC PRESENTA LA CLASE "A,B,C" EN EL LABEL15
          PTJE# = INDIC#(NC%)
          Label15 = CLASE$(PTJE#)
        End If
        '
        
        
        Call CambiosEnPantallaCompraInterna(NC)
        'CARGO ULTIMA IMP.CONTABLE
10      If NC% < 0 Then
          NCXX1% = NC% * (-1)
        End If
        Call APERBASDAT("COMPRAS")
        SQLX$ = " SELECT IContab  FROM OCOMDETA "
        SQLX$ = SQLX$ + " WHERE Npro_Ocom = " & NCXX1% & " "
        SQLX$ = SQLX$ + " AND Stat_Ocom < 9 "
        SQLX$ = SQLX$ + " AND IContab <> 0 "
        SQLX$ = SQLX$ + " ORDER BY Femi_Ocom DESC "
        'Set ICONTMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
        Dim ICONTMP As ADODB.Recordset
        Set ICONTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
        On Error Resume Next
        With ICONTMP
            .MoveFirst
            If Err Then
              On Error GoTo 0
              GoTo 20
            End If
            On Error GoTo 0
            Text7 = CODCUE(!IContab)
        End With
        'CARGO ULTIMO SECTOR
20      Call APERBASDAT("COMPRAS")
        SQLX$ = " SELECT Sect_Ocom  FROM OCOMDETA "
        SQLX$ = SQLX$ + " WHERE Npro_Ocom = " & NCXX1% & " "
        SQLX$ = SQLX$ + " AND Stat_Ocom < 9 "
        SQLX$ = SQLX$ + " AND Sect_Ocom <> ' ' "
        SQLX$ = SQLX$ + " ORDER BY Femi_Ocom DESC "
        'Set SECTORTMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
        Dim SECTORTMP As ADODB.Recordset
        Set SECTORTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
        With SECTORTMP
            On Error Resume Next
            .MoveFirst
            If Err Then
              On Error GoTo 0
              GoTo 99
            End If
            On Error GoTo 0
            Text18 = !Sect_Ocom
        End With
        '
      End If
    End If
    
99  Screen.MousePointer = 1
    '
    On Error Resume Next
    OEncaTemp.Close
    Set OEncaTemp = Nothing
    ICONTMP.Close
    Set ICONTMP = Nothing
    
    SECTORTMP.Close
    Set SECTORTMP = Nothing
    On Error GoTo 0
    
End Sub
   '
Function PRONUOCO&()
   
   PRONU& = 1
   Call APERBASDAT("COMPRAS")
   
   SQLX$ = "Select max(Nume_ocom) as maximo from OcomEnca"
   Dim OEncaTemp As ADODB.Recordset
   Set OEncaTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
   If Not IsNull(OEncaTemp!Maximo) Then
      PRONU& = 1 + OEncaTemp!Maximo
   End If
   OEncaTemp.Close
   Set OEncaTemp = Nothing
   PRONUOCO& = PRONU&
   '
End Function

Private Sub Text6_GotFocus()
   If VALICLIE% < 1 Then
     On Error Resume Next
     Text1.SetFocus
     On Error GoTo 0
     Exit Sub
   End If
   '
   On Error Resume Next
   Text6.SelStart = 0
   Text6.SelLength = Len(Text6.TEXT)
   On Error GoTo 0
End Sub

Private Sub Text7_DblClick()
   Call Command18_Click
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
     TICACH(MONEMI%) = XVALO(TICAMBIS)
   End If
   TICAMBX# = 1
   If MONEMI% > 1 Then
     TICAMBX# = XVALO(TICAMBIS)
   End If

End Sub

Private Sub TICAMBIS_DblClick()
      ENTRATICA.Frame4.Caption = "Cotización Local para " '+ MONEMIS.Caption
      ENTRATICA.Text4 = TICAMBIS.Caption
      ENTRATICA.Text4 = FORMATNUM(XVALO(ENTRATICA.Text4), "F10.4")
      ENTRATICA.Show 1
      If TRIM$(ENTRATICA.Text4) <> "" Then
        TICAMBIS = TRIM$(ENTRATICA.Text4)
      End If

End Sub

Private Sub Timer1_Timer()
   '
   'Call CALTOCOM
   '
   '
   If CAMBIOVALGRILLA% > 0 Then
      Screen.MousePointer = 11
      Call CALTOCOM
      CAMBIOVALGRILLA% = 0
      Screen.MousePointer = 1
   End If
   '

End Sub

Private Sub Total_Change(Index As Integer)
   If Index = 0 Then
     Call CALTOCOM
'      DESC$ = PORDESCU.TEXT
'      TTT1 = XVALO(Total(0).TEXT) * COEFDESCU(DESC$) / 100
'      Total(1).TEXT = CSTRING$(MKS$(TTT1), 3, 11, 2, 2)
'      TTT3 = XVALO(Total(0).TEXT) - XVALO(Total(1).TEXT)
'      Total(2).TEXT = CSTRING$(MKS$(TTT3), 3, 11, 2, 2)
   End If
'   If Index < 3 Then
'      TTT1 = XVALO(Total(2).TEXT)
'      TTT2 = XVALO(TOTIVA(3).TEXT)
'      Total(3).TEXT = Str$(TTT1 + TTT2)
'   End If
  If Index < 2 Then
     RIVACOM07.TONECOM = FORMATNUM$(XVALO(Total(0)) - XVALO(Total(1)), "F13.2")
     Label118 = TRIM$(RIVACOM07.Label2)
     Total(2) = TRIM$(RIVACOM07.Label3)
  End If
End Sub

Private Sub Total_DblClick(Index As Integer)
   For KKI% = 0 To 4
     Total(KKI%).Enabled = Not (Total(KKI%).Enabled)
   Next KKI%
End Sub

'
Private Sub Total_GotFocus(Index As Integer)
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Sub GRACOMPRA(OKX%, Optional VALIDACIONSTOCKMINIMO$, Optional RetNroCompra&)
   '
   
   OKX% = 0
   If ULTREG&("!OCOMDETA") < 1 Then Exit Sub
   '
   COPRECIO = 1
   If Total(0).Enabled = False Then COPRECIO = 0
   '
   If InStr(1, EMPREAC$, "AHP") > 0 Then
       Call SELECHO("VIASENTR")
       VIA% = XVALO(VALACT1$("VIASENTR"))
   End If

   Call CIERRARCH("OCOMENCA")
   Call CIERRARCH("OCOMDETA")
   '
'   Call BLOQARCH("OCOMENCA")
'   TINI = Int(Timer)
'   While Int(Timer) < TINI + 5: Wend
'   Call BLOQARCH("OCOMENCA")
   '
   NR1$ = XVALO(Text11)
   NR1$ = NR1$ & "-" & (Label16)

' VERIFICA QUE EL NÚMERO PROPUESTO ESTA LIBRE EN LA TABLA OCOMENCA
   Call APERBASDAT("COMPRAS")
   STRSQL = "Select Nume_Ocom from OcomEnca where Nume_Ocom=" & TRIM$(Text11.TEXT)
   Dim MIRS As ADODB.Recordset
   Set MIRS = FLEXCONN.Execute(FILTSQL$(STRSQL))
   If Not (MIRS.EOF And MIRS.BOF) Then ' si lo encuentra
15    NPP& = XVALO(Text11.TEXT) + 1
      Call CIERRARCH("OCOMENCA")
      Call CIERRARCH("OCOMDETA")
      Call CIERRARCH("MACONSIG")
      Call COMUNI("FLEXOFT ha Re-Numerado la O/Compra\como O/C Número '" + TRIM$(Str$(NPP&)) + "'.\  \Verifique y Vuelva a Grabar.")
      Text11.TEXT = TRIM$(Str$(NPP&))
      Exit Sub
   End If
   
   MONECO% = XVALO(Text15)
   If MONECO% < 1 Then MONECO% = XVALO(TRIM$(Control$("OCOMPRA", "MONECOS")))
   If MONECO% < 1 Then MONECO% = 1
   '
   Screen.MousePointer = 11
   If VALICLIE% < 1 Then
      MERRO$ = "Número de Proveedor no Válido"
      GoTo 98
   End If
   NPRO% = XVALO(Text1.TEXT)
   '
   If XVALO(TEXT21(0).TEXT) > 255 Then
      MERRO$ = "Condición de Pago no Válida"
      GoTo 98
   End If
   '
   If XVALO(TEXT21(1).TEXT) > 255 Then
      MERRO$ = "Embalaje no Válido"
      GoTo 98
   End If
   '
   HOII% = HOY(HOS$)
   FEMI% = FECHANUM(Text9.TEXT)
   If FEMI% < 1 Or FEMI% > HOII% Then
      MERRO$ = "Fecha de Emisión no Válida"
      Call COMUNI(MERRO$)
      GoTo 98
   End If
   '
123 If COPRECIO < 0.99 Then
      PORDESCU.TEXT = ""
      Text17 = ""
      For KK% = 0 To 4
        Total(KK%).TEXT = ""
      Next KK%
    End If
    '
' VUELVE A VERIFICAR QUE EL NÚMERO DE O/C ESTÉ LIBRE
    STRSQL = "Select * from OcomEnca where Nume_Ocom=" & XVALO(Text11.TEXT)
    Set MIRS = FLEXCONN.Execute(FILTSQL$(STRSQL))
    If Not (MIRS.EOF And MIRS.BOF) Then ' si lo encuentra
       GoTo 15
    End If
    '
' AQUI AGREGA UN UNICO REGISTRO DE ENCABEZADO EN LA TABLA OCOMENCA
    STATUOC% = (-1)
    If Left$(Control$("OCOMPRA", "MODOAPRO"), 6) = "DIRECT" Then STATUOC% = 0
    If VALIDACIONSTOCKMINIMO$ = "GRABA_PENDIENTE_APROBACION" Then STATUOC% = -2

    Call APERBASDAT("COMPRAS")
    Dim OCOMENCA As ADODB.Recordset
    Set OCOMENCA = New ADODB.Recordset
    OCOMENCA.Open FILTSQL$(STRSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With OCOMENCA
          .AddNew   ' YA CONVERTIDO
          On Error Resume Next
          !CodiEmpr = CodiEmp%
          On Error GoTo 0
             NRO& = XVALO(Text11)
          !nume_ocom = NRO&
             FEX = FEMI%
             REFE$ = FECHATEX$(FEX) + " - " + rasocli$((-1) * NPRO%)
          !Refe_OcoM = Left$(REFE$, 64)
          !Femi_Ocom = CVDAT(FEMI%)
          !Npro_Ocom = NPRO%
          !Raso_Prov = Left$(rasocli$((-1) * NPRO%), 48) 'OT-06-0227-WAR-05/06/06///
          !TSum_Ocom = 0               ' materiales
'             STATUOC% = (-1)           ' SIGNIFICA PENDIENTE DE APROBACIÓN  'lo puse arriba
'             If Left$(CONTROL$("OCOMPRA", "MODOAPRO"), 6) = "DIRECT" Then STATUOC% = 0
'          If VALIDACIONSTOCKMINIMO$ = "GRABA_PENDIENTE_APROBACION" Then STATUOC% = -2
          !Stat_Ocom = STATUOC%        ' PASA A 0 (CERO) SI APROBACION DIRECTA
          '
          !Fgen_Ocom = Now
          !Usua_Gen = USNBR%
             MONECOI% = XVALO(Text15)
          !MONE_COS = MONECOI%        ' MONEDA PARA COSTEO
          !Mone_Emis = MONEMI%        ' MONEDA DE EMISIÓN
             TICAMBX = 1
             If MONEMI% <= 1 Then
               If MONECOI% > 1 Then
                 TICAMBX = TICAMONE(MONECOI%)
               End If
             End If
          !Tcam_Ocom = TICAMBX       ' TIPO DE CAMBIO / COTIZACION
             CPAG% = XVALO(TEXT21(0))
          !Cpag_Ocom = CPAG%
             EMBAL% = XVALO(TEXT21(1))
          On Error Resume Next    ' POR SI NO ESTA DEFINIDO
          !Embala_Ocom = EMBAL%
          On Error GoTo 0
          !AtSr_Ocom = TRIM$(Text6)
             '
             PORDESCUENTO$ = TRIM$(PORDESCU)
             DESC$ = PORDESCUENTO$
             TOTANE# = XVALO(Total(0)) * COPRECIO
             IDESCUEN# = XVALO(Total(1)) * COPRECIO
             TOTANEDES# = TOTANE# - IDESCUEN#
             TOTIVA# = XVALO(Total(2)) * COPRECIO
             TOTPER# = XVALO(Total(4)) * COPRECIO
             TOTATO# = TOTANEDES# + TOTIVA# + TOTPER
             '
          !ToBru_Ocom = TOTANE#
          !Pdes_Ocom = PORDESCU
          !IDes_Ocom = IDESCUEN#
          !Iiva_Ocom = TOTIVA#
          !Iotr1_Ocom = TOTPER#
          !Total_Ocom = TOTATO#
          !Lentrega = Left$(Text10, 128) 'OT-06-0227-WAR-05/06/06///
          !Obse_Ocom = Text13
          !Id_Comprob = "OC." + SAFETEXT(NRO&)
          On Error Resume Next
          !Anex_Ocom = RAI$
          On Error GoTo 0
          '\\\OT-08-0887-OD-24/09/08///
          !PorcIva = Label118 ' PORCENTAJE DE IVA PARA ESTA ORDEN DE COMPRA
          '\\\OT-08-0887-OD-FIN///
          '
          !CodCarpImp = Label18 ' Guarda Nº de Carpeta de Importaciones
          If Label18 <> "" Then ' Anotador de Carpeta de importaciones
            Call REPLA(TTXX$, HOS$, 1, 8)
            Call REPLA(TTXX$, USERNAME$, 69, 24)
            NR$ = SAFETEXT(Text11)
            Call REPLA(TTXX$, "Generacion de Orden de Compra Nº: " + NR$, 11, 56)
            Call AGREGA_ANOTACARPI(CODCARPE$, TTXX$)
            TTXX$ = Space(76)
            Call REPLA(TTXX$, "Proveedor: " + Left$(rasocli$((-1) * NPRO%), 64), 11, 56)
            Call AGREGA_ANOTACARPI(CODCARPE$, TTXX$)
            Call AGREGA_ANOTACARPI(CODCARPE$, String$(88, "="))
          End If
          '
          !Respon_Ocompra = Left$(Text12, 128)
          !Via_Entrega = XVALO(VIA%)
          !COTIZA_DOLAR = TICAMONE(2)
          !OComCliente = Left(TEXT25, 16)
          !nume_cli = XVALO(TEXT24)
          .Update
    End With  'fin de OcomEnca
    OCOMENCA.Close
    Set OCOMENCA = Nothing
    '
    ' aqui graba el detalle del IVA
    STRSQL = "Select * from DETAIVA where Nume_Ocom = " & NRO&
    Dim DETIVA As ADODB.Recordset
    Set DETIVA = New ADODB.Recordset
    DETIVA.Open FILTSQL$(STRSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With RIVACOM07
      For KKII% = 0 To 5
        If XVALO(.IMPOGRA(KKII%)) >= 0.005 Or XVALO(.IMPOIVA(KKII%)) >= 0.005 Then
          DETIVA.AddNew   ' YA CONVERTIDO
                On Error Resume Next
          DETIVA!CodiEmpr = CodiEmp%
                On Error GoTo 0
          DETIVA!nume_ocom = NRO&
          DETIVA!NORIT = KKII%
          DETIVA!Tasa_Iva = .TASIVA(KKII%)
          DETIVA!Base_Impo = XVALO(.IMPOGRA(KKII%).TEXT)
          DETIVA!Impo_Iva = XVALO(.IMPOIVA(KKII%))
          DETIVA.Update
        End If
      Next KKII%
    End With
    DETIVA.Close
    Set DETIVA = Nothing
    '
    EMICONSI% = 0
    IDENOC$ = MKS$(NRO&) + MKI$(FEMI%) + MKI$(NPRO%)
    '
    ORDIT% = 0
'    APRODIR% = 0
'    If Left$(CONTROL$("OCOMPRA", "MODOAPRO"), 6) = "DIRECT" Then
'      APRODIR% = 1
'    End If
    ARCHILIB$ = "!1OCOMDET"
    '
' A PARTIR DE AQUI GRABA LOS REGISTROS DE DETALLE DE LA ORDEN DE COMPRA (VARIOS)
      STRSQL = "Select * from Ocomdeta where Nume_ocom=" & NRO&
      Dim OCOMDETA As ADODB.Recordset
      Set OCOMDETA = New ADODB.Recordset
      OCOMDETA.Open FILTSQL$(STRSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    
    For J& = 1 To ULTREG&("!OCOMDETA") '(grilla q llena el usuario)
      '
      ZZ$ = REGLEIDO$("!OCOMDETA", J&)
      If COPRECIO < 0.99 Then
        Call REPLA(ZZ$, String$(28, 0), 67, 28)  ' BORRA PRECIO UNITARIO E IMPORTE
        Call GRAREG("!OCOMDETA", ZZ$, J&)
      End If
      '
      ZZ$ = REGLEIDO$("!OCOMDETA", J&)
      If TRIM$(Mid$(ZZ$, 105, 12)) = "" Then Call REPLA(ZZ$, Text7, 105, 12)
      If TRIM$(Mid$(ZZ$, 15, 12)) = "" Then Call REPLA(ZZ$, Text18, 15, 12)
      Call GRAREG("!OCOMDETA", ZZ$, J&)
      '
      CI0% = CVI(Mid$(ZZ$, 9, 2))
      PedInt$ = Mid$(ZZ$, 31, 20)
      COEFUNI = CVS(Mid$(ZZ$, 63, 4)): If COEFUNI < 0.0001 Then COEFUNI = 1
      PREUNPE# = CVD(Mid$(ZZ$, 67, 8)) * TICAMONE(MONEMI%) / COEFUNI
      PREUNCO# = PREUNPE# / TICAMONE(MONECOI%)
      Monecos$ = "$ ": If MONECOI% > 1 Then Monecos$ = Left$(Text16, 2)
      VAUNI$ = Monecos$ + Mid$(FORMATNUM$(PREUNCO#, "F12.4"), 3)
      '
      '
      ORDIT% = ORDIT% + 1
      With OCOMDETA
        .AddNew   ' YA CONVERTIDO
        On Error Resume Next
        !CodiEmpr = CodiEmp%
        On Error GoTo 0
         !nume_ocom = NRO&
         !NuOrd_Item = ORDIT%
         !Femi_Ocom = CVDAT(FEMI%)
         !Npro_Ocom = NPRO%
         !Raso_Prov = Left$(rasocli$((-1) * NPRO%), 64) 'OT-06-0227-WAR-05/06/06///
         !Medi_Ocom = PedInt$
         !Cod_Inte = CI0%
         !Cod_Exte = Left$(CODEXT$(CI0%), 24)
         !DESCRITEM = DESCRIT0$(CI0%)
         !DELIBRE = ""
         If COLTELIB% >= 0 Then
           !DELIBRE = TRIM$(REGLEIDO$(ARCHILIB$, J&))
         End If
         If CI0% < 1 Then
           If TRIM$(!DELIBRE) <> "" Then
             !DESCRITEM = Left$(!DELIBRE, 64)
           End If
         End If
         !Valo_Unid = VAUNI$
         !Uni_Mas = Mid$(ZZ$, 11, 4)
           CAUNIS = CVS(Mid$(ZZ$, 51, 4))  ' CANT. EN UNIDADES DE STOCK
         !UNIDS_OCOM = CAUNIS
           CANTU = CVS(Mid$(ZZ$, 55, 4))   ' CANT EN UNIDADES DE COMPRA
         !Cant_Ocom = CANTU
         !Uni_Com = Mid$(ZZ$, 59, 4) ' unidad de compra
         !Coef_Unids = COEFUNI
         !Pre_Unit = CVD(Mid$(ZZ$, 67, 8))
         !Desc_Item = CVS(Mid$(ZZ$, 75, 4))
         !prun_neto = CVD(Mid$(ZZ$, 79, 8))
         !INet_Item = CVD(Mid$(ZZ$, 87, 8))
         !PrUn_Neto_Pes = PREUNCO# 'Se agrega este campo para guardar el valor en pesos
         !Fentre_Sol = CVDAT(CVI(Mid$(ZZ$, 95, 2)))
         !CONSIMAT = TRIM$(Mid$(ZZ$, 97, 2))
         !Cant_Rec = 0
         !Cod_Oper = 0
         !DOC_ORIG = ""
         !N_OrServ = 0
         !MONE_COS = MONECOI%        ' antes pos 83,1
         !Mone_Emis = MONEMI%        ' antes pos 84,1
         !TSum_Ocom = 0              ' materiales
'         APRODIR% = -1
'         If VALIDACIONSTOCKMINIMO$ = "GRABA_PENDIENTE_APROBACION" Then APRODIR% = -2
'         !stat_ocom = APRODIR%   ' status Pendiente de aprobacion
         '
         !Stat_Ocom = STATUOC%
         !IContab = CUECOINT(TRIM$(Mid$(ZZ$, 105, 12))) ' Imputacion contable
         '!IContab = 0
         !Sect_Ocom = TRIM$(Mid$(ZZ$, 15, 12))
        .Update
      End With
      RetNroCompra& = NRO&
      '
' EN ESTE SUB-BLOQUE GENERA LAS NECESIDADES DE MATERIALES A CONSIGNAR AL PRO-
' VEEDOR SEGÚN LA ESTRUCTURA DE PRODUCTO DE CADA REGISTRO DE LA O/COMPRA
      NOCOM$ = TRIM$(Str$(NRO&))
        While Len(NOCOM$) < 6: NOCOM$ = "0" + NOCOM$: Wend
      ITNU$ = TRIM$(Str$(ORDIT%))
        While Len(ITNU$) < 2: ITNU$ = "0" + ITNU$: Wend
      NOCOM$ = "OC-" + Left$(NOCOM$, 6) + "-" + ITNU$
      Call LEEEXPLO(CODEXT$(CI0%), CAUNIS)
      '
      Call APERBASDAT("STOCKS")
      STRSQL = "Select * from Maconsig where Nro_ocom='" & NOCOM$ & "'"
      Dim MaConsig As ADODB.Recordset
      Set MaConsig = New ADODB.Recordset
      MaConsig.Open FILTSQL$(STRSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
      NUORD% = 0
      For i% = 1 To CAIT%
        CII% = CIJ%(i%)
        CANTI# = CDbl(CAUNIS * USOI(i%))
        '
        If TRIM$(CODEXT$(CII%)) <> "" Then
          If CANTI# > 0.05 Then
            With MaConsig
               .AddNew   ' YA CONVERTIDO
               On Error Resume Next
               !CodiEmpr = CodiEmp%
               On Error GoTo 0
               !Femi_Ocom = CVDAT(FEMI%)
               !Npro_Ocom = NPRO%
               !Raso_Prov = Left$(rasocli$((-1) * NPRO%), 64) 'OT-06-0227-WAR-05/06/06///
               !nro_ocom = NOCOM$
                '
               !Cod_Inte = CII%
               !Cod_Exte = CODEXT$(CII%)
               !DESCRITEM = DESCRIT0$(CII%)
               !Unidad = Unimed$(CII%)
               !cant_asig = CANTI#
               !Secc_Ejec = 0
               !Deno_Secc = ""
               !NRO_ORFA = ""
               !LOTE_OCOM = ""
               !cant_entreg = 0
               !FECHA_ENTREG = CVDAT(0)
               !NREMI_CONSIG = ""
               !Valo_Unit = COSUNI(CII%)
               !Moneda = MONECOSTO%(CII%)
               !Valo_Tot = CANTI# * COSUNI(CII%) * TICAMONE(MONECOSTO(CII%))
               !CANT_RENDID = 0
               !cant_devu = 0
               !Saldo_Pend = 0
               !Observa = ""
               NUORD% = NUORD% + 1
               !NUORDITEM = NUORD%
              .Update
              EMICONSI% = 1
            End With
          End If
        End If
      Next i%
      MaConsig.Close
      Set MaConsig = Nothing
      '
   Next J&
   OCOMDETA.Close
   Set OCOMDETA = Nothing
   
   Call BAJALDISCO
   '
   ' AQUI GRABA EL NUMERO DE O/c EN LAS COTIZACIONES RECIBIDAs
   If Left$(APLINVO$, 4) = "CREC" Then
       TAPLI$ = Mid$(APLINVO$, 5)
       NUSOLI& = XVALO(Left$(TAPLI$, 8))
       ORSOLI& = XVALO(Mid$(TAPLI$, 9, 4))
       NCLIE% = XVALO(TRIM$(Mid$(TAPLI$, 13, 6))) ' NUMERO DE PROVEEDOR
       '
       CONDI$ = "NUME_SOLINT = " & NUSOLI& & " AND NUORD_ITEM = " & ORSOLI& & " AND NPRO_OCOM = " & NCLIE%
       Call ACTUREGISTRO("COTIRECI", "Nume_Ocom", CONDI$, NRO&, RAFE&, "NOMESS")
       Call ANOTASIM07.ANOTASIM(NUSOLI&, "Emision O/C - " + rasocli$((-1) * NPRO%))
       '
       For R& = 1 To ULTREG&("!OCOMDETA") ' ANOTA EN EL RESERVA
            CI0% = CVI(Mid$(ZZ$, 9, 2))
            Condi1$ = "NUME_SOLINT = " & NUSOLI& & " AND COD_INTE = " & CI0% ' GRABO EN EL RESERVA NRO SOLICI
            Call ACTUREGISTRO("RESERVA", "Nume_Ocom", Condi1$, NRO&, RAFE&, "NOMESS")
       Next R&
   End If
   '
   If Left$(APLINVO$, 4) = "CREM" Then
      For J& = 1 To ULTREG&("!OCOMDETA") '(grilla q llena el usuario)
         ZZ$ = REGLEIDO$("!OCOMDETA", J&)
         TAPLI$ = TRIM$(Mid$(ZZ$, 31, 20))
         PPXX% = InStr(TAPLI$, "/")
         If PPXX% > 1 Then
            NUSOLI& = XVALO(Left$(TAPLI$, PPXX% - 1))
            ORSOLI& = XVALO(Mid$(TAPLI$, PPXX% + 1))
            NCLIE% = XVALO(Text1) ' NUMERO DE PROVEEDOR
            '
            CONDI$ = "NUME_SOLINT = " & NUSOLI& & " AND NUORD_ITEM = " & ORSOLI& & " AND NPRO_OCOM = " & NCLIE%
            Call ACTUREGISTRO("COTIRECI", "Nume_Ocom", CONDI$, NRO&, RAFE&, "NOMESS")
            Call ANOTASIM07.ANOTASIM(NUSOLI&, "Emision O/C - " + rasocli$((-1) * NPRO%))
            '
            CI0% = CVI(Mid$(ZZ$, 9, 2))
            Condi1$ = "NUME_SOLINT = " & NUSOLI& & " AND COD_INTE = " & CI0% ' GRABO EN EL RESERVA NRO SOLICI
            Call ACTUREGISTRO("RESERVA", "Nume_Ocom", Condi1$, NRO&, RAFE&, "NOMESS")
         End If
      Next J&
   End If
   '
' AQUI TERMINA LA GRABACIÓN DE LOS REGISTROS DE DETALLE EN OCOMDETA
   '
   If COPRECIO < 0.99 Then
      PORDESCU.TEXT = ""
      Text17 = ""
      For KK% = 0 To 4
        Total(KK%).TEXT = ""
      Next KK%
   End If
   '
' DESDE AQUI actualiza costo standard y kilos por pieza
   ACOSREP$ = UCase$(Control$("OCOMPRA", "ACOSREP"))
   '
   HOII% = HOY(HOS$)
   Dim Master As ADODB.Recordset
   '
   For J& = 1 To ULTREG&("!OCOMDETA")
     ZZ$ = REGLEIDO$("!OCOMDETA", J&)
     NUORSE& = CVS(Mid$(ZZ$, 117, 4))
     If NUORSE& < 1 Then   ' EXCLUYE ORDEN DE SERVICIO
        CI0% = CVI(Mid$(ZZ$, 9, 2))
        COEFUNI = CVS(Mid$(ZZ$, 63, 4)): If COEFUNI < 0.0001 Then COEFUNI = 1
        'AGREGADO PARA TENER EN CUENTA SI SE CONSIDERA EL DESCUENTO PARA EL COSTO DE REPOSCION ALEJPA
        If Control("OCOMPRA", "NODESCUE") <> "SI" Then ' ACA ENTRA COMO ESTABA
           COSPESOSD# = CVD(Mid$(ZZ$, 79, 8)) * TICAMONE(MONEMI%) / COEFUNI
        Else ' SI ESTA MARCADO EN SETUP DE COMPRAS NO CONSIDERAR DESCUENTO
           COSPESOSD# = CVD(Mid$(ZZ$, 67, 8)) * TICAMONE(MONEMI%) / COEFUNI
        End If
        '
        'SQL
        If CI0% < 1 Or CI0% > NUMAS% Then GoTo 170
        STRSQL = "Select * from Master where codint=" & CI0%
        Set Master = New ADODB.Recordset
        Master.Open FILTSQL$(STRSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

        PRHBL% = PROVHABI%(CI0%)
        If PRHBL% < 1 Or PRHBL% = NPRO% Or ACOSREP$ = "SIEMPRE" Then
            COSMONCO = COSPESOSD# / TICAMONE(MONECO%)
            If COSMONCO >= 0.000005 Or PRHBL% < 1 Then  ' PARA QUE ACTUALICE P.HABIT AUN SI COSTO REPOS. ES CERO
               If ACOSREP$ <> "NUNCA" Then
                  '
                  With Master
                        If COSMONCO >= 0.000005 Then
                           COSTOANTER# = XVALO(!COSUNI)
                           !COSUNI = COSMONCO
                           !FECOSTO = CVDAT(FEMI%)
                           PRENUE# = COSMONCO
                           Call ADD_NOVEDAD(CI0%, NPRO%, COSTOANTER#, PRENUE#, "Actualización Por Orden de Compra: " & NR1$, MONECO%, "COSTO")
                        End If
                        If PRHBL% < 1 Then
                           !PROVHABI = NPRO%
                        End If
                        !Monecos = MONECO%
                        .Update
                  End With
               End If
            End If
            '
        Else
            '
            If PRHBL% < 1 Then GoTo 164 ' no hay proveedor habitual anterior
            '
            TTXX$ = "Para el Artículo " + TRIM$(CODEXT$(CI0%)) + "\el Proveedor Habitual Registrado es\"
            TTXX$ = TTXX$ + TRIM$(rasocli$((-1) * PRHBL%)) + "\  \"
            TTXX$ = TTXX$ + "Desea Cambiar el nuevo Proveedor Habitual a\"
            TTXX$ = TTXX$ + rasocli$((-1) * NPRO%)
            If COMALTER%(TTXX$ + "\\No, Conservar Anterior\Si, Cambiar") = 2 Then
              '
164           With Master
                !PROVHABI = NPRO%
                COSMONCO = COSPESOSD# / TICAMONE(MONECO%)
                If COSMONCO >= 0.000005 Then
                   COSTOANTER# = XVALO(!COSUNI)
                  !COSUNI = COSMONCO
                  !FECOSTO = CVDAT(FEMI%)
                  !Monecos = MONECO%
                  PRENUE# = COSMONCO
                  Call ADD_NOVEDAD(CI0%, NPRO%, COSTOANTER#, PRENUE#, "Actualización Por Orden de Compra : " & NR1$, MONECO%, "COSTO")
                End If
                .Update
              End With
              
              ZZX$ = REGBLAN$("RECAMPH")
              Call REPLA(ZZX$, MKI$(CI0%), 1, 2)
              HOII% = HOY(HOS$)
              Call REPLA(ZZX$, MKI$(HOII%), 3, 2)
              Call REPLA(ZZX$, MKI$(PRHBL%), 5, 2)
              Call REPLA(ZZX$, MKI$(NPRO%), 7, 2)
              Call REPLA(ZZX$, USERTER$, 9, 24)
              Call REPLA(ZZX$, "Emis.O/C " + TRIM$(Str$(NRO&)), 33, 32)
              Call REGAPP("RECAMPH", ZZX$)
              Call CIERRARCH("RECAMPH")
              GoTo 169
              '
              End If
            '
            CUNIREG# = COSUNI(CI0%) * TICAMONE(MONECOSTO(CI0%))
            If CUNIREG# >= 0.00005 Then
              If COSPESOSD# > CUNIREG# + 0.00005 Then
                Call COMUNI("ATENCION: Precio de Compra\es Superior al Costo Standard\en Código '" + TRIM$(CODEXT$(CI0%)) + "' !!!\  \Informe al Supervisor de Compras.")
              End If
            End If
169     End If
        '
' EN ESTE BLOQUE SE ACTUALIZA EL PESO POR UNIDAD DE LA PIEZA
        If PESUNI(CI0%) < 0.000005 Then
           UNIST$ = UCase$(TRIM$(Mid$(ZZ$, 11, 4)))
           If Left$(UNIST$, 1) = "U" Then
             UNICO$ = UCase$(TRIM$(Mid$(ZZ$, 59, 4)))
             If UNICO$ = "KG" Then
               On Error Resume Next
               PSXX = 1 / COEFUNI
               On Error GoTo 0
               With Master
                   !PESUNI = PSXX
                   .Update
               End With
             End If
           End If
        End If
        '
        Master.Close
        Set Master = Nothing
     End If
170 Next J&
   '
   'termina la actualización
   'EN ESTE BLOQUE TOMA LOS ITEMS PLANEADOS EN PLACOMP Y LOS NETEA POR LA ORDEN
   ' DE COMPRA RECIEN EMITIDA - EN PRINCIPIO NO HAY NADA QUE MODIFICAR
   For J& = 1 To ULTREG&("!OCOMDETA")
      ZZ$ = REGLEIDO$("!OCOMDETA", J&)
      NUORSE& = CVS(Mid$(ZZ$, 117, 4))
      If NUORSE& < 1 Then   ' EXCLUYE ORDEN DE SERVICIO
        CI0% = CVI(Mid$(ZZ$, 9, 2))
        COEFUNI = CVS(Mid$(ZZ$, 63, 4)): If COEFUNI < 0.0001 Then COEFUNI = 1
        CANOCOM = CVS(Mid$(ZZ$, 51, 4))
        FENTRESO% = CVI(Mid$(ZZ$, 95, 2))
        If FENTRESO% < FEMI% Then FENTRESO% = FEMI%
        '
        If CantRegistros&("OCOEMI") > 0 Then
           CANTOCO# = CANOCOM
           Dim OCOSE As ADODB.Recordset
           Set OCOSE = New ADODB.Recordset
           SQLX$ = "SELECT * FROM OCOEMI WHERE COD_INTE = " & CI0%
           OCOSE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
           '
           With OCOSE
             Do While Not .EOF
               CAPENO# = XVALO(!Cantidad)
               If CAPENO# < CANTOCO# Then
                   !Cantidad = 0
                   CANTOCO# = CANTOCO# - CAPENO#
                 Else
                   !Cantidad = CAPENO# - CANTOCO#
                   CANTOCO# = 0
               End If
               .Update
             .MoveNext
             Loop
           End With
           OCOSE.Close
           Set OCOSE = Nothing
           GoTo 179
        End If
        '
        RFF$ = RECFILT$("PLACOMP", 1, MKI$(CI0%))
        For UJ& = 1 To Len(RFF$) Step 4
          RPLA& = CVS(Mid$(RFF$, UJ&, 4))
          XX$ = REGLEIDO$("PLACOMP", RPLA&)
          FENTREPLA% = CVI(Mid$(XX$, 3, 2))
            If FENTREPLA% < FEMI% Then FENTREPLA% = FEMI%
          CANPLA = CVS(Mid$(XX$, 21, 4))
          '
          If FENTREPLA% >= FENTRESO% Then
            If CANPLA <= CANOCOM Then
                CANOCOM = CANOCOM - CANPLA
                CANPLA = 0
              Else
                CANPLA = CANPLA - CANOCOM
                CANOCOM = 0
            End If
            Call REPLA(XX$, MKS$(CANPLA), 21, 4)
            ' depurar registro de plan de compras
            If CANPLA < 0.005 Then XX$ = String$(128, 0)
            Call GRAREG("PLACOMP", XX$, RPLA&)
          End If
        Next UJ&
      End If
179 Next J&
    Call CIERRARCH("PLACOMP")
    '
    JJ& = 0
    For U& = 193 To 512 Step 64
      TTXX$ = Mid$(ENCAO$, U&, 64)
      JJ& = JJ& + 1
      Call GRAREG("!OBSERVOF", TTXX$, JJ&)
    Next U&
    Call SETULTREG("!OBSERVOF", JJ&)
    Call CIERRARCH("!OBSERVOF")
    '
    CPAG% = CVI(Mid$(ENCAO$, 153, 2))
    FORPAG$ = TRIM$(ECOARCH$("CONPAG", Chr$(CPAG%)))
    If FORPAG$ = "" Then FORPAG$ = "Contado Contra Entrega"
    '
   
    OKX% = 1
98  Screen.MousePointer = 1
 End Sub

Private Sub Total_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
     If Index = 0 Then
       On Error Resume Next
       PORDESCU.SetFocus
       On Error GoTo 0
     End If
   End If
End Sub

Private Sub Total_LostFocus(Index As Integer)
   If Index = 0 Then
     TTT1 = XVALO(Total(0).TEXT)
     Total(0) = CSTRING$(MKS$(TTT1), 3, 11, 2, 2)
   End If
End Sub
'\\\OT-05-0688-WAR-20/10/05///
Sub CARDESCOM()
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))

   Call CIERRARCH("!OCOMDETA")
   List2.ListIndex = (-1)
   HOII% = HOY(HOS$)
   '
10 ATRI$ = "/NC"
   ATRI$ = ATRI$ + "/TITUPAN=Detalle de Orden de Compra en " + TRIM$(Label11)
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(140 + Top + Frame1.Top + Picture3.Top))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture3.Left + 50))
   '
   Timer1.Enabled = True
   PRF$ = "CARDESCO"
   XXX% = VENTABU%(PRF$ + ATRI$)
   List1.Clear
   Call CALTOCOM
   Timer1.Enabled = False
   '
   If XXX% < 0 Or ULTREG&("!OCOMDETA") < 1 Then
      Call CIERRARCH("!OCOMDETA")
      Call BLANARCH("!OCOMDETA")
      Call BLANFORMU
      Total(0).TEXT = ""
      Text1.TEXT = ""
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   '
   Screen.MousePointer = 11
   'ATK% = 0
   'DESACTIVE LA BANDERA DE CAMBIOS XQ SIEMPRE CAMBIA
   'YA QUE SE ESCRIBE EL REGISTRO DE UNIDAD/C X EJEMPLO.
   NPRO% = XVALO(Text1)
   For U& = 1 To ULTREG&("!OCOMDETA")
      ZZ$ = REGLEIDO$("!OCOMDETA", U&)
      'Z0$ = ZZ$
      CI0% = CVI(Mid$(ZZ$, 9, 2))
      '
      TIECOD% = VALCODPROV%(NPRO%, CI0%)    ' Esta funcion valida que el codigo tenga relacion con el proveedor
      If TIECOD% = 0 Then GoTo 10           ' Si devuelve 0 el codigo no esta relacionado
      '
      ' validacion que no tenga marca en proceso de articulos con salidas menores a 5000 e
      ' importes por cantidad salida menores a 5000
      If STATMOVI%(CI0%) < 0 Then
         Call COMUNI("El Artículo " & CODEXT$(CI0%) & " Tiene Marca de Baja Para Realizar Compras ")
         GoTo 10
      End If

      If TRIM$(Mid$(ZZ$, 11, 4)) = "" Then Call REPLA(ZZ$, Unimed$(CI0%), 11, 4)
      Call REPLA(ZZ$, Unimed$(CI0%), 59, 4)
      Call REPLA(ZZ$, MKS$(1), 63, 4)
      '
      CANSTOK = CVS(Mid$(ZZ$, 51, 4))
      CANCOMP = CANSTOK
      '
      'VALIDAR SI TIENE LOTE INDIVISIBLE- VALIDA QUE LA CANTIDAD SEA MULTIPLO DE L.INDIV.

      If InStr(EPAC$, "NEUQUEN") < 1 Then 'NEUQUEN NO VALIDA N.SERIE
        If LINDIVIS(CI0%) > 0.05 Then
          CANTID# = ROUND#(XVALO(CANSTOK), 1)
          CANLINDIV# = ROUND#(REDON_INDIVIS(CI0%, CANTID#), 1)
          
          If CANLINDIV# <> CANTID# Then
             Call COMUNI("Cantidad del Código: '" & CODEXT$(CI0%) & "'\Ajustada Según Lote Indivisible")
             CANSTOK = CANLINDIV#
             CANCOMP = CANSTOK: If COEFI > 0 Then CANCOMP = CANSTOK / COEFI
             Call REPLA(ZZ$, MKS$(CANSTOK), 51, 4)
             Call REPLA(ZZ$, MKS$(CANCOMP), 55, 4)
             Call GRAREG("!OCOMDETA", ZZ$, U&)
             GoTo 10
          End If
      End If
      End If
      UNIST$ = UCase$(TRIM$(Mid$(ZZ$, 11, 4)))
      UNICO$ = UNIST$
      Call REPLA(ZZ$, MKS$(CANCOMP), 55, 4)
      Call REPLA(ZZ$, UNICO$, 59, 4)
      '
      PreUnid# = CVD(Mid$(ZZ$, 67, 8))
      If PreUnid# < 0.00005 Then
        PORDESCUEI = 0
        If CI0% > 0 Then    ' para que no lo traiga con material libre
          If UC$ = "" Then UC$ = ULTICOM$(NPRO%, CI0%)
          MONULCO% = Asc(Mid$(UC$, 122, 1)): If MONULCO% < 1 Then MONULCO% = 1
          PreUnid# = CVD(Mid$(UC$, 67, 8)) * TICAMONE(MONULCO%)
          PORDESCUEI = CVS(Mid$(UC$, 75, 4))
          If PreUnid# < 0.00005 Or MODOULPRE% = 2 Then
            PreUnid# = COSUNI(CI0%) * TICAMONE(MONECOSTO%(CI0%)) * COEFI
            PORDESCUEI = 0
          End If
          PreUnid# = PreUnid# / TICAMONE(MONEMI%)
        End If
        Call REPLA(ZZ$, MKD$(PreUnid#), 67, 8)
        Call REPLA(ZZ$, MKS$(PORDESCUEI), 75, 4)
      End If
      '
      If TRIM$(Mid$(ZZ$, 105, 12)) = "" Then Call REPLA(ZZ$, Text7, 105, 12)
      If TRIM$(Mid$(ZZ$, 15, 12)) = "" Then Call REPLA(ZZ$, Text18, 15, 12)
    '  If ZZ$ <> Z0$ Then ATK% = 1
      Call GRAREG("!OCOMDETA", ZZ$, U&)
   Next U&
   '
   Screen.MousePointer = 1
   'If ATK% = 1 Then GoTo 10
   '
   Call CARLISCO
   Call CALTOCOM
   Screen.MousePointer = 1
   '
  '
End Sub
'\\\OT-05-0688-WAR-FIN///
Sub CARDETCOM()
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))

   '\\\OT-05-0688-WAR-20/10/05///
   If Control$("OCOMPRA", "MODOIMPU") = "DES" Then
     Call CARDESCOM
     Screen.MousePointer = 1
     Exit Sub
   End If
   '\\\OT-05-0688-WAR-FIN///
   '
   Call CIERRARCH("!OCOMDETA")
   List2.ListIndex = (-1)
   HOII% = HOY(HOS$)
   '
10 ATRI$ = "/NC"
   ATRI$ = ATRI$ + "/TITUPAN=Detalle de Orden de Compra en " + TRIM$(Label11)
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(140 + Top + Frame1.Top + Picture3.Top))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture3.Left + 50))
   '
   Timer1.Enabled = True
   PRF$ = "CARDETCO"
   XXX% = VENTABU%(PRF$ + ATRI$)
   
   'AGREGADO PARA QUE NO SUPERA LA CANTIDAD DE FILAS 255
   CANTITEM = ULTREG&("!OCOMDETA")
   If CANTITEM > 255 And XXX% >= 0 Then
      Call COMUNI("Se Ha Superado El Máximo (255) de Filas Permitido Para Una Orden de Compra\Actualmente ha Ingresado: " & CANTITEM & "\Suprima el Excedente y Carguelo En otra Orden de Compra")
      GoTo 10
   End If
   
   List1.Clear
   Call CALTOCOM
   Timer1.Enabled = False
   '
   If XXX% < 0 Or ULTREG&("!OCOMDETA") < 1 Then
      Call CIERRARCH("!OCOMDETA")
      Call BLANARCH("!OCOMDETA")
      Call BLANFORMU
      Total(0).TEXT = ""
      Text1.TEXT = ""
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   '
   Screen.MousePointer = 11
   ATK% = 0
   NPRO% = XVALO(Text1)
   For U& = 1 To ULTREG&("!OCOMDETA")
      ZZ$ = REGLEIDO$("!OCOMDETA", U&)
      Z0$ = ZZ$
      CI0% = CVI(Mid$(ZZ$, 9, 2))
      '
      TIECOD% = VALCODPROV%(NPRO%, CI0%)    ' Esta funcion valida que el codigo tenga relacion con el proveedor
      If TIECOD% = 0 Then GoTo 10           ' Si devuelve 0 el codigo no esta relacionado
      '
      ' validacion que no tenga marca en proceso de articulos con salidas menores a 5000 e
      ' importes por cantidad salida menores a 5000
      If STATMOVI%(CI0%) < 0 Then
         Call COMUNI("El Artículo " & CODEXT$(CI0%) & " Tiene Marca de Baja Para Realizar Compras ")
         GoTo 10
      End If
      '
      If TRIM$(Mid$(ZZ$, 11, 4)) = "" Then Call REPLA(ZZ$, Unimed$(CI0%), 11, 4)
      If TRIM$(Mid$(ZZ$, 59, 4)) = "" Then Call REPLA(ZZ$, UNICOM$(CI0%, NPRO%), 59, 4)
      If TRIM$(Mid$(ZZ$, 11, 4)) = TRIM$(Mid$(ZZ$, 59, 4)) Then Call REPLA(ZZ$, MKS$(1), 63, 4)
      '
      ' COEFICIENTE CALCULADO A PARTIR DE LAS CANTIDADES DE COMPRA Y STOCK INGRESADAS
      UC$ = ""
      COEFI = CVS(Mid$(ZZ$, 63, 4))
      CANCOMP = CVS(Mid$(ZZ$, 55, 4))
      CANSTOK = CVS(Mid$(ZZ$, 51, 4))
      'If CANCOMP < 0.05 Or CANSTOK < 0.05 Then COEFI = 0
      If CANCOMP > 0.05 And CANSTOK > 0.05 Then COEFI = 0
      If COEFI < 0.0001 Then
        If CANCOMP > 0.05 Then
          If CANSTOK > 0.05 Then
            COEFI = CANSTOK / CANCOMP
          End If
        End If
      End If
      '
      ' COEFICIENTE OBTENIDO DEL PESO POR UNIDAD DEL MAESTRO
      If COEFI < 0.0001 Then
        Peso# = PESUNI(CI0%)
        If Peso# > 0.00005 Then
            COEFI = 1 / Peso#
        End If
      End If
      '
      ' COEFICIENTE OBTENIDO A PARTIR DE LA ULTIMA COMPRA
      If COEFI < 0.0001 Then
        UC$ = ULTICOM$(NPRO%, CI0%)
        If TRIM$(Mid$(ZZ$, 11, 4)) = TRIM$(Mid$(UC$, 11, 4)) Then
          If TRIM$(Mid$(ZZ$, 59, 4)) = TRIM$(Mid$(UC$, 59, 4)) Then
            COEFI = CVS(Mid$(UC$, 63, 4))
          End If
        End If
      End If
      '
      If COEFI < 0.0001 Then COEFI = 1
      
      '
      'VALIDACION PARA EL CASO QUE POSEA COEFICIENTE Y TENGA MISMA UNIDAD DE MEDIDA
      If TRIM$(Mid$(ZZ$, 11, 4)) = TRIM$(Mid$(ZZ$, 59, 4)) Then
         If COEFI <> 1 Then
           Call COMUNI("Imposible Continuar. El Artículo '" & CODEXT$(CI0%) & "'\Posee Cantidades Distintas para Unidades de Medida Iguales")
           GoTo 10
         End If
      End If
      '
      
      'VALIDAR SI TIENE LOTE INDIVISIBLE- VALIDA QUE LA CANTIDAD SEA MULTIPLO DE L.INDIV.
      If InStr(EPAC$, "NEUQUEN") < 1 Then 'NEUQUEN NO VALIDA N.SERIE
        If LINDIVIS(CI0%) > 0.05 Then
          CANTID# = ROUND#(XVALO(CANSTOK), 1)
          CANLINDIV# = ROUND#(REDON_INDIVIS(CI0%, CANTID#), 1)
          
          If CANLINDIV# <> CANTID# Then
             Call COMUNI("Cantidad del Código: '" & CODEXT$(CI0%) & "'\Ajustada Según Lote Indivisible")
             CANSTOK = CANLINDIV#
             CANCOMP = CANSTOK / COEFI
             Call REPLA(ZZ$, MKS$(CANSTOK), 51, 4)
             Call REPLA(ZZ$, MKS$(CANCOMP), 55, 4)
             Call GRAREG("!OCOMDETA", ZZ$, U&)
             GoTo 10
          End If
        End If
      End If
      '
      If CANCOMP < 0.05 Then
          CANCOMP = CANSTOK / COEFI
        Else
          CANSTOK = CANCOMP * COEFI
      End If
      CANCOMP = CANSTOK / COEFI
      '
'      If InStr(EPAC$, "CONTEPLAST") > 0 And TIMATE%(CI0%) = 41 Then 'si es conteplast y tipomat es bobinas
'         StringNroCtes$ = StringDeNroProveedoresComprasInternas$
'         'SI EL NUMERO DE CLIENTE ESTA DENTRO DE LA CADENA DE NRO DE PROVEEDORES CONFIGURADOS COMO PARTICIPANTES DE LA COMPRA INTERNA.
'         If NumeroEstaEnCadena(StringNroCtes$, NPRO%) Then
'            Peso# = PESUNI(CI0%)
'            If Peso# > 0.00005 Then
'               coefi = 1 / Peso#
'               CANCOMP = CANSTOK * coefi 'SI SON BOBINAS
'            End If
'         End If
'      End If
       If InStr(EPAC$, "CONTEPLAST") > 0 Or InStr(EPAC$, "CAMBACERES") > 0 Or InStr(EPAC$, "QUILMES") > 0 Then
            TMT% = TIMATE%(CI0%)
            UME$ = TRIM$(UCase$(Unimed$(CI0%)))
            If TMT% = 41 And UME$ = "M." Then ' tipomat es bobinas Y Unimed$ ES M.
                StringNroCtes$ = StringDeNroProveedoresComprasInternas$
               'SI EL NUMERO DE CLIENTE ESTA DENTRO DE LA CADENA DE NRO DE PROVEEDORES CONFIGURADOS COMO PARTICIPANTES DE LA COMPRA INTERNA.
               If NumeroEstaEnCadena(StringNroCtes$, NPRO%) Then
                  Peso# = PESUNI(CI0%)
                  If Peso# > 0.00005 Then
                     COEFI = 1 / Peso#
                     CANCOMP = CANSTOK * COEFI 'SI SON BOBINAS
                     CANCOMP = XVALO(FORMATNUM$(CANCOMP, "F12.0"))
                  End If
               End If
            ElseIf TMT% = 41 And UME$ = "U." Then 'ESTOY DENTRO DE CONTEP,CAMBA,QUILMES
                StringNroCtes$ = StringDeNroProveedoresComprasInternas$
               'SI EL NUMERO DE CLIENTE ESTA DENTRO DE LA CADENA DE NRO DE PROVEEDORES CONFIGURADOS COMO PARTICIPANTES DE LA COMPRA INTERNA.
               If NumeroEstaEnCadena(StringNroCtes$, NPRO%) Then
                  Peso# = PESUNI(CI0%)
                  If Peso# > 0.00005 Then
                     COEFI = 1 / Peso#
                     CANCOMP = CANSTOK / COEFI 'SI SON BOBINAS
                     CANCOMP = XVALO(FORMATNUM$(CANCOMP, "F12.0"))
                  End If
               End If
            End If
        End If
        
      Call REPLA(ZZ$, MKS$(CANSTOK), 51, 4)
      Call REPLA(ZZ$, MKS$(CANCOMP), 55, 4)
      Call REPLA(ZZ$, MKS$(COEFI), 63, 4)
      '
      PreUnid# = CVD(Mid$(ZZ$, 67, 8))
      If DERACCE%("VERCOSTO/NOMESS", "Visualizar Costos y Precios") > 1 Then
        If PreUnid# < 0.00005 Then
          If UC$ = "" Then UC$ = ULTICOM$(NPRO%, CI0%)
          MONULCO% = Asc(Mid$(UC$, 122, 1)): If MONULCO% < 1 Then MONULCO% = 1
          PreUnid# = CVD(Mid$(UC$, 67, 8)) * TICAMONE(MONULCO%)
          PORDESCUEI = CVS(Mid$(UC$, 75, 4))
          If PreUnid# < 0.00005 Or MODOULPRE% = 2 Then
            PreUnid# = COSUNI(CI0%) * TICAMONE(MONECOSTO%(CI0%)) * COEFI
            PORDESCUEI = 0
          End If
          PreUnid# = PreUnid# / TICAMONE(MONEMI%)
          Call REPLA(ZZ$, MKD$(PreUnid#), 67, 8)
          Call REPLA(ZZ$, MKS$(PORDESCUEI), 75, 4)
        End If
      End If
      '
      If ZZ$ <> Z0$ Then ATK% = 1
      Call REPLA(ZZ$, Text7, 105, 12)
      Call REPLA(ZZ$, Text18, 15, 12)
      Call GRAREG("!OCOMDETA", ZZ$, U&)
   Next U&
   '
   Screen.MousePointer = 1
   If ATK% = 1 Then GoTo 10
   '
   Call CARLISCO
   Call CALTOCOM
   Screen.MousePointer = 1
   FORCOMP07.SetFocus
   '
End Sub

Sub CALTOCOM()
   '1
   EPAC$ = TRIM$(UCase$(EMPREAC$))

   NC1% = (-1) * XVALO(Text1)
   NPRO% = XVALO(Text1)
   TOTANE# = 0: JJ& = 0:
   J& = 0
   ACUMULAPESO# = 0
   For J& = 1 To ULTREG&("!OCOMDETA")
       X$ = REGLEIDO$("!OCOMDETA", J&)
       CI0% = CVI(Mid$(X$, 9, 2))
       UNIMAS$ = Mid$(X$, 11, 4)
       CANTU = CVS(Mid$(X$, 55, 4))
       PreUnid# = CVD(Mid$(X$, 67, 8))
       PORDE = CVS(Mid$(X$, 75, 4))
       FENTRE% = CVI(Mid$(X$, 95, 2))
       CANSTXX = CVS(Mid$(X$, 51, 4))
       '
       'SI NO TIENE CANTIDAD LO BORRA
       'SE COMENTO POR QUE AL ENTRAR POR LA CARGA DE CADA CAMPO  EN LA PRIMERA VEZ COMO NO TENIA LA CANTIDAD ASIGNADA SALTEABA EL REGISTRO.
       'If CANSTXX < 0.05 And CANTU < 0.05 Then GoTo 30
       'FIN (15-0348)
       PREUNINE# = (Int(PreUnid# * 100 * (100 - PORDE) + 0.5)) / 10000
       INETITEM# = CDbl((Int(100 * PREUNINE# * CANTU + 0.5)) / 100)
       TOTANE# = TOTANE# + INETITEM#
       '
       Call REPLA(X$, MKD$(PREUNINE#), 79, 8)
       Call REPLA(X$, MKD$(INETITEM#), 87, 8)
       Call GRAREG("!OCOMDETA", X$, J&)
       '
       UNIST$ = UCase$(TRIM$(Mid$(X$, 11, 4)))
       UNICO$ = UCase$(TRIM$(Mid$(X$, 59, 4)))
       CANSTXX = CVS(Mid$(X$, 51, 4))
       '
       ' COEFICIENTE CALCULADO A PARTIR DE LAS CANTIDADES DE COMPRA Y STOCK INGRESADAS
       UC$ = ""
       COEFI = CVS(Mid$(X$, 63, 4))
       CANCOMP = CVS(Mid$(X$, 55, 4))
       CANSTOK = CVS(Mid$(X$, 51, 4))
       If CANCOMP < 0.05 Or CANSTOK < 0.05 Then COEFI = 0
       If COEFI < 0.0001 Then
         If CANCOMP > 0.05 Then
           If CANSTOK > 0.05 Then
             COEFI = CANSTOK / CANCOMP
           End If
         End If
       End If
       '
       
       ' COEFICIENTE OBTENIDO DEL PESO POR UNIDAD DEL MAESTRO
       LEYOPESO% = 0
       If COEFI < 0.0001 Then
         Peso# = PESUNI(CI0%)
         If Peso# > 0.00005 Then
             COEFI = 1 / Peso#
         End If
         LEYOPESO% = 1
       End If
       '
       'AGREGO ESTAS LINEAS PARA QUE NO LEA SI NO ES NECESARIO Y NO PROVOQUE LENTITUD EN OTROS CLIENTES
       If LEYOPESO% = 0 And InStr(EPAC$, "AHP") > 0 Then
         Peso# = PESUNI(CI0%)
       End If
       '
       ACUMULAPESO# = ACUMULAPESO# + (Peso# * CANSTOK)
       ' COEFICIENTE OBTENIDO A PARTIR DE LA ULTIMA COMPRA
       If COEFI < 0.0001 Then
         UC$ = ULTICOM$(NPRO%, CI0%)
         If TRIM$(Mid$(X$, 11, 4)) = TRIM$(Mid$(UC$, 11, 4)) Then
           If TRIM$(Mid$(X$, 59, 4)) = TRIM$(Mid$(UC$, 59, 4)) Then
             COEFI = CVS(Mid$(UC$, 63, 4))
           End If
         End If
       End If
       '
       If COEFI < 0.0001 Then COEFI = 1
       Call REPLA(X$, MKS$(COEFI), 63, 4)
       JJ& = JJ& + 1
       Call GRAREG("!OCOMDETA", X$, JJ&)
       '
30 Next J&
   '
   Call SETULTREG("!OCOMDETA", JJ&)
   Call CIERRARCH("!OCOMDETA")
   '
' presentar registro de totales y condicion de pago
   '
   If MONEMI% > 1 Then
     MONECOI% = MONEMI%
     TICAMBX = TICAEMI
   End If
   '
   PORDESCUENTO$ = TRIM$(PORDESCU)
   DESC$ = PORDESCUENTO$
   IDESCUEN# = TOTANE# * COEFDESCU(DESC$) / 100
   TOTANEDES# = TOTANE# - IDESCUEN#
   TOTIVA# = 0: APLIVA% = 0
'   ALIVX = XVALO(Text12)
'   If ALIVX < 0 Or ALIVX > 99 Then
'      Text12 = TRIM$(Str$(ALIVA))
'      ALIVX = ALIVA
'   End If
   If Abs(NC1%) > 0 Then
     PVCL% = PIVACLI%(NC1%)
     If PVCL% <> 2 Then
       If PVCL% <> 4 Then
         If PVCL% <> 5 Then
'           TOTIVA# = (Int(ALIVX * TOTANEDES# + 0.5)) / 100
           APLIVA% = 1
         End If
       End If
     End If
   End If
   '
   TOTPER# = 0
   PERCX = XVALO(Text17)
   If PERCX < 0 Or PERCX > 99 Then
      Text17 = ""
      PERCX = 0
   End If
   If Abs(NC1%) > 0 Then
     PVCL% = PIVACLI%(NC1%)
     If PVCL% <> 2 Then
       If PVCL% <> 4 Then
         If PVCL% <> 5 Then
           TOTPER# = (Int(PERCX * TOTANEDES# + 0.5)) / 100
         End If
       End If
     End If
   End If
   '
   Total(0) = FORMATNUM$(TOTANE#, "F11.2")
   Total(1) = FORMATNUM$(IDESCUEN#, "F11.2")
     TOTIVA# = APLIVA% * XVALO(RIVACOM07.Label3)
   Total(2) = FORMATNUM$(TOTIVA#, "F11.2")
   Total(4) = FORMATNUM$(TOTPER#, "F11.2")
     TOTATO# = TOTANEDES# + TOTIVA# + TOTPER
   Total(3) = FORMATNUM$(TOTATO#, "F13.2")
   TXTPESO = FORMATNUM$(ACUMULAPESO#, "F10.2")
   '
End Sub

Sub CARCOPEN()
   '
   Screen.MousePointer = 11
   NPRO% = XVALO(Text1)
   HO% = HOY(HOS$)
   '
   JJ& = 0
   REBLA$ = REGBLAN$("OCOMDETA")
   Call SETULTREG("!OCOMDETA", 0)
   ARCHILIB$ = "!1OCOMDET"
   Call SETULTREG(ARCHILIB$, 0)
   '
   PCI$ = "": UFECHA = HOY(HOS$)
   '
   FORSELMAT.LIMATSEL.Clear
   CONDI$ = "Npro_Ocom = " & NPRO% & " AND Cantidad > 0"
   If CantRegistros("OCOEMI", CONDI$, "NOMESS") > 0 Then
      If Label18 = "" Then
'         Call CARGALISTA(FORSELMAT.LIMATSEL, "OCOEMI", "COD_INTE", CONDI$, "DISTINCT")
         'CAMBIE EL CARGALISTA POR LA CONSULTA, POR QUE NO ESTA FUNCIONANDO BIEN LA RUTINA PRESENTABA LA VENTANA VACIA
         'ESTO APARECEIO EN MEDITEA
         SQLX$ = "SELECT DISTINCT OCOEMI.COD_INTE , MASTER.CODIGO, MASTER.DESCRIPCION FROM OCOEMI  INNER JOIN MASTER WITH(NOLOCK)  "
         SQLX$ = SQLX$ + " ON MASTER.CODINT = OCOEMI.COD_INTE "
         SQLX$ = SQLX$ + " WHERE OCOEMI.Npro_Ocom = " & NPRO%
         SQLX$ = SQLX$ + " AND OCOEMI.Cantidad > 0"
         Set RS = READSET(SQLX$)
         With RS
            Do While Not .EOF
              CI1% = XVALO(!Cod_Inte)
              CODIX$ = SAFETEXT$(!Codigo)
              DESX$ = SAFETEXT$(!Descripcion)
              Call REPLA(TX$, CODIX$, 1, 16)
              Call REPLA(TX$, DESX$, 18, 48)
              FORSELMAT.LIMATSEL.AddItem TX$
              .MoveNext
            Loop
         End With
         GoTo 14
      End If
   End If
   '
   ' ESTE ES EL VIEJO QUE SELECCIONA DE LA TABLA DAT
   ' NO COMENTAR ESTE CODIGO YA QUE ES UTILIZADO POR ASEMBLI EN EL MODULO SIMULA07
   RFF$ = RECFILT$("PLACOMP", 8, MKI$(NPRO%))
   If RFF$ = "" Then Exit Sub
   '
   For KKI% = 1 To Len(RFF$) Step 4
      REPX& = CVS(Mid$(RFF$, KKI%, 4))
      XX$ = REGLEIDO$("PLACOMP", REPX&)
      CI0% = CVI(Left$(XX$, 2))
      MCI0$ = MKI$(CI0%)
      FENTREIT% = CVI(Mid$(XX$, 3, 2))
      If FENTREIT% > UFECHA Then UFECHA = FENTREIT%
      For UKI% = 1 To Len(PCI$) Step 2
        If Mid$(PCI$, UKI%, 2) = MCI0$ Then GoTo 9
      Next UKI%
      PCI$ = PCI$ + MCI0$
9  Next KKI%
   '
   FORSELMAT.LIMATSEL.Clear
   For UKI% = 1 To Len(PCI$) Step 2
     CI0% = CVI(Mid$(PCI$, UKI%, 2))
     TTXX$ = AJUSTI$(CODEXT$(CI0%), 16) + DESCRIT0$(CI0%)
     FORSELMAT.LIMATSEL.AddItem TTXX$
   Next UKI%
   '
14 FORSELMAT.Caption = "Materiales a Comprar - " + Text2
   FORSELMAT.Label2 = FECHATEX$(UFECHA)
   For U& = 1 To FORSELMAT.LIMATSEL.ListCount
     FORSELMAT.LIMATSEL.Selected(U& - 1) = True
   Next U&
   '
   Screen.MousePointer = 1
   On Error Resume Next
   FORSELMAT.BOTREC.SetFocus
   FORSELMAT.LIMATSEL.TopIndex = 0
   On Error GoTo 0
   FORSELMAT.Show 1
   '
   DoEvents
   If TRIM$(FORSELMAT.Label2 = "") Then Exit Sub
   '
   Screen.MousePointer = 11
   PCI$ = "":
   UFECHAI% = FECHANUM(FORSELMAT.Label2)
   For U& = 1 To FORSELMAT.LIMATSEL.ListCount
     If FORSELMAT.LIMATSEL.Selected(U& - 1) = True Then
       ZZZ = CODINT%(TRIM$(Left$(FORSELMAT.LIMATSEL.List(U& - 1), 15)))
       PCI$ = PCI$ + MKI$(CODINT%(TRIM$(Left$(FORSELMAT.LIMATSEL.List(U& - 1), 15))))
     End If
   Next U&
   '
   CONDI$ = "Npro_Ocom = " & NPRO%
   If CantRegistros("OCOEMI") > 0 Then
       GoSub LEE_OCOEMI
       GoTo 40
   End If
   '
   ' ACA LEE DE LA VIEJA TABLA DAT
   RFF$ = RECFILT$("PLACOMP", 8, MKI$(NPRO%))
   For KKI% = 1 To Len(RFF$) Step 4
      REPX& = CVS(Mid$(RFF$, KKI%, 4))
      XX$ = REGLEIDO$("PLACOMP", REPX&)
      CI0% = CVI(Left$(XX$, 2))
      MCI0$ = MKI$(CI0%)
      For UKI% = 1 To Len(PCI$) Step 2
        If Mid$(PCI$, UKI%, 2) = MCI0$ Then GoTo 24
      Next UKI%
      GoTo 29
      
24    FENTREIT% = CVI(Mid$(XX$, 3, 2))
      If FENTREIT% < HO% Then FENTREIT% = HO%
      If FENTREIT% <= UFECHAI% Then
        '
        CANTU = CVS(Mid$(XX$, 21, 4))
        '
        For KKJ& = 1 To JJ&
          ZZ$ = REGLEIDO$("!OCOMDETA", KKJ&)
          If CVI(Mid$(ZZ$, 9, 2)) = CI0% Then
            If CVI(Mid$(ZZ$, 95, 2)) = FENTREIT% Then
              CANTU = CANTU + CVS(Mid$(ZZ$, 51, 4))
              Call REPLA(ZZ$, MKS$(CANTU), 51, 4)
              Call GRAREG("!OCOMDETA", ZZ$, KKJ&)
              GoTo 29
            End If
          End If
        Next KKJ&
        '
        ZZ$ = REBLA$
        Call REPLA(ZZ$, MKI$(CI0%), 9, 2)
        Call REPLA(ZZ$, Unimed$(CI0%), 11, 4)
        Call REPLA(ZZ$, MKS$(CANTU), 51, 4)
        Call REPLA(ZZ$, Unimed$(CI0%), 59, 4)
        Call REPLA(ZZ$, MKS$(1), 63, 4)
        Call REPLA(ZZ$, MKI$(FENTREIT%), 95, 2)
        '
        UC$ = ULTICOM$(NPRO%, CI0%)
        MONULCO% = Asc(Mid$(UC$, 122, 1)): If MONULCO% < 1 Then MONULCO% = 1
        PreUnid# = CVD(Mid$(UC$, 67, 8)) * TICAMONE(MONULCO%)
        COEFIX = CVS(Mid$(UC$, 63, 4)): If COEFIX < 0.0005 Then COEFIX = 1
        If PreUnid# < 0.00005 Or MODOULPRE% = 2 Then
          PreUnid# = COSUNI(CI0%) * TICAMONE(MONECOSTO%(CI0%)) * COEFIX
        End If
        PreUnid# = PreUnid# / TICAMONE(MONEMI%)
        '
        UNICOMPRA$ = TRIM$(Mid$(UC$, 59, 4))
        If UNICOMPRA$ <> "" Then Call REPLA(ZZ$, UNICOMPRA$, 59, 4)
        COEFUNI = CVS(Mid$(UC$, 63, 4)): If COEFUNI < 0.0001 Then COEFUNI = 1
        If UNICOMPRA$ = TRIM$(Unimed$(CI0%)) Then COEFUNI = 1
        Call REPLA(ZZ$, MKS$(COEFUNI), 63, 4)
        'SI TIENE DERECHOS MUESTRA EL PRECIO
        If DERACCE%("VERCOSTO/NOMESS", "Visualizar Costos y Precios") > 1 Then
          Call REPLA(ZZ$, MKD$(PreUnid#), 67, 8)
        End If
        '
        JJ& = JJ& + 1
        Call GRAREG("!OCOMDETA", ZZ$, JJ&)
          DELIBX$ = DESCRIT0$(CI0%)
        Call GRAREG(ARCHILIB$, DELIBX$, JJ&)
        '
      End If
29 Next KKI%
   '
40 Call SETULTREG("!OCOMDETA", JJ&)
   Screen.MousePointer = 1
Exit Sub
   '
   '
LEE_OCOEMI:
   '
   SQLX$ = "SELECT * FROM OCOEMI WHERE Npro_Ocom = " & NPRO%
   Dim OCOSEL As ADODB.Recordset
   Set OCOSEL = New ADODB.Recordset
   Set OCOSEL = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   JJ& = 0
   With OCOSEL
      Do While Not .EOF
         CI0% = XVALO(!Cod_Inte)
         MCI0$ = MKI$(CI0%)
         For UKI% = 1 To Len(PCI$) Step 2
           If Mid$(PCI$, UKI%, 2) = MCI0$ Then GoTo 34
         Next UKI%
         GoTo 39
         
34       FENTREIT% = CVINT(!Fentre_Sol)
         If FENTREIT% < HO% Then FENTREIT% = HO%
         If FENTREIT% <= UFECHAI% Then
           '
           CANTU = XVALO(!Cantidad)
           '
           For KKJ& = 1 To JJ&
             ZZ$ = REGLEIDO$("!OCOMDETA", KKJ&)
             If CVI(Mid$(ZZ$, 9, 2)) = CI0% Then
               If CVI(Mid$(ZZ$, 95, 2)) = FENTREIT% Then
                 CANTU = CANTU + CVS(Mid$(ZZ$, 51, 4))
                 Call REPLA(ZZ$, MKS$(CANTU), 51, 4)
                 Call GRAREG("!OCOMDETA", ZZ$, KKJ&)
                 GoTo 39
               End If
             End If
           Next KKJ&
           '
           ZZ$ = REBLA$
           Call REPLA(ZZ$, MKI$(CI0%), 9, 2)
           Call REPLA(ZZ$, Unimed$(CI0%), 11, 4)
           Call REPLA(ZZ$, MKS$(CANTU), 51, 4)
           Call REPLA(ZZ$, Unimed$(CI0%), 59, 4)
           Call REPLA(ZZ$, MKS$(1), 63, 4)
           Call REPLA(ZZ$, MKI$(FENTREIT%), 95, 2)
           '
           UC$ = ULTICOM$(NPRO%, CI0%)
           MONULCO% = Asc(Mid$(UC$, 122, 1)): If MONULCO% < 1 Then MONULCO% = 1
           PreUnid# = CVD(Mid$(UC$, 67, 8)) * TICAMONE(MONULCO%)
           COEFIX = CVS(Mid$(UC$, 63, 4)): If COEFIX < 0.0005 Then COEFIX = 1
           If PreUnid# < 0.00005 Or MODOULPRE% = 2 Then
'           aca es donde se puede producir el problema de ambest el codigo comentado lo agregue yo pero hay
'           que analizar si estaria bien
'             If MONECOSTO%(CI0%) = MONULCO% Then
                PreUnid# = COSUNI(CI0%) * TICAMONE(MONECOSTO%(CI0%)) * COEFIX
'             End If
           End If
           AA = COSUNI(CI0%)
           A1 = TICAMONE(MONECOSTO%(CI0%))
           PreUnid# = PreUnid# / TICAMONE(MONEMI%)
           '
           UNICOMPRA$ = TRIM$(Mid$(UC$, 59, 4))
           If UNICOMPRA$ <> "" Then Call REPLA(ZZ$, UNICOMPRA$, 59, 4)
           COEFUNI = CVS(Mid$(UC$, 63, 4)): If COEFUNI < 0.0001 Then COEFUNI = 1
           If UNICOMPRA$ = TRIM$(Unimed$(CI0%)) Then COEFUNI = 1
           Call REPLA(ZZ$, MKS$(COEFUNI), 63, 4)
           'SI TIENE DERECHOS MUESTRA EL PRECIO
           If DERACCE%("VERCOSTO/NOMESS", "Visualizar Costos y Precios") > 1 Then
             Call REPLA(ZZ$, MKD$(PreUnid#), 67, 8)
           End If
           '
           JJ& = JJ& + 1
           Call GRAREG("!OCOMDETA", ZZ$, JJ&)
             DELIBX$ = DESCRIT0$(CI0%)
           Call GRAREG(ARCHILIB$, DELIBX$, JJ&)
           '
         End If
39    .MoveNext
      Loop
   End With
   '
Return
'
End Sub

Sub CARCOANT()
   '
   Screen.MousePointer = 11
   List2.Clear
   NPRO% = XVALO(Text1)
   HOU = HOY(HOS$)
   '
   Call APERBASDAT("COMPRAS")
   SQLX$ = " SELECT * FROM  OCOMDETA O INNER JOIN MASTER M WITH(NOLOCK)"
   SQLX$ = SQLX$ + " ON O.COD_INTE = M.CODINT "
   SQLX$ = SQLX$ + " WHERE O.Stat_Ocom < 8 "
   SQLX$ = SQLX$ + " AND O.Stat_Ocom >=0 "
   SQLX$ = SQLX$ + " AND O.Npro_Ocom = " & NPRO% & " "
   SQLX$ = SQLX$ + " AND O.N_OrServ < 1 "
   'Set CARCOANTMP = Compras.OpenRecordset(sqlx$, dbOpenSnapshot)
   Set CARCOANTMP = READSET(SQLX$)
   With CARCOANTMP
     Do While Not .EOF
'       NUORSE& = !N_OrServ
'       If NUORSE& < 1 Then
         COEFCAN = !Coef_Unids
         If COEFCAN < 0.005 Then COEFCAN = 1
         CACOM = !Cant_Ocom * COEFCAN
         CAREC = !Cant_Rec * COEFCAN
         CAPEN = CACOM - CAREC
         If CAPEN / COEFCAN >= 0.05 Then
           CIXI% = !Cod_Inte
           NROPED& = !nume_ocom
           UNICO$ = !Uni_Com
           FEX = CVINT(!Fentre_Sol)
           ATRAU = 0
           If FEX < HOU Then
             ATRAU = (DIENTRE(FEX, HOU)) / 7
             If ATRAU > 99.9 Then ATRAU = 99.9
           End If
           ATRAX$ = FORMATNUM$(ATRAU, "F4.1")
           '
           LI2TEX$ = Space$(80)
'           Call REPLA(LI2TEX$, CODEXT$(CIXI%), 1, 16)
'           Call REPLA(LI2TEX$, DESCRIT0$(CIXI%), 17, 33)
           Call REPLA(LI2TEX$, SAFETEXT$(!Cod_Exte), 1, 16)
           Call REPLA(LI2TEX$, SAFETEXT$(!Descripcion), 17, 33)

             If CIXI% < 1 Then        ' MATERIAL LIBRE
               DECRI$ = TRIM$(!DELIBRE)
               If DECRI$ <> "" Then
                 Call REPLA(LI2TEX$, DECRI$, 17, 33)
               End If
             End If
           Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN / COEFCAN), 4, 8, 1, 2), 51, 8)
           Call REPLA(LI2TEX$, UNICO$, 60, 4)
           Call REPLA(LI2TEX$, FECHATEX$(FEX), 65, 8)
           Call REPLA(LI2TEX$, ATRAX$, 74, 4)
           Call REPLA(LI2TEX$, FORMATNUM$(NROPED&, "I6"), 79, 6)
           List2.AddItem LI2TEX$
           '
         End If
'       End If
       .MoveNext
     Loop
   End With
   CARCOANTMP.Close
   Set CARCOANTMP = Nothing
   Screen.MousePointer = 1
   '
'   '
'   Screen.MousePointer = 11
'   List2.Clear
'   NPRO% = XVALO(Text1)
'   HOU = HOY(HOS$)
'   '
'   RFF$ = RECFILT$("OCOMDETA", 3, MKI$(NPRO%))
'   For UK& = 1 To Len(RFF$) Step 4
'     U& = CVS(Mid$(RFF$, UK&, 4))
'     X$ = REGLEIDO$("OCOMDETA", U&)
'     STAT% = Asc(Mid$(X$, 124, 1))
'     If STAT% < 8 Then      ' no esta anulada ni marcada cumplida
'       NUORSE& = CVS(Mid$(X$, 117, 4))
'       If NUORSE& < 1 Then
'         COEFCAN = CVS(Mid$(X$, 63, 4))
'         If COEFCAN < 0.005 Then COEFCAN = 1
'         CACOM = CVS(Mid$(X$, 55, 4)) * COEFCAN
'         CAREC = CVS(Mid$(X$, 99, 4)) * COEFCAN
'         CAPEN = CACOM - CAREC
'         If CAPEN / COEFCAN >= 0.05 Then
'           CIXI% = CVI(Mid$(X$, 9, 2))
'           NroPed& = CVS(Left$(X$, 4))
'           UNICO$ = TRIM$(Mid$(X$, 59, 4))
'           FEX = CVI(Mid$(X$, 95, 2))
'           ATRAU = 0
'           If FEX < HOU Then
'             ATRAU = (DIENTRE(FEX, HOU)) / 7
'             If ATRAU > 99.9 Then ATRAU = 99.9
'           End If
'           ATRAX$ = FORMATNUM$(ATRAU, "F4.1")
'           '
'           LI2TEX$ = Space$(80)
'           Call REPLA(LI2TEX$, CODEXT$(CIXI%), 1, 16)
'           Call REPLA(LI2TEX$, DESCRIT0$(CIXI%), 17, 33)
'           Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN / COEFCAN), 4, 8, 1, 2), 51, 8)
'           Call REPLA(LI2TEX$, UNICO$, 60, 4)
'           Call REPLA(LI2TEX$, FECHATEX$(FEX), 65, 8)
'           Call REPLA(LI2TEX$, ATRAX$, 74, 4)
'           Call REPLA(LI2TEX$, FORMATNUM$(NroPed&, "I6"), 79, 6)
'           List2.AddItem LI2TEX$
'           '
'         End If
'       End If
'     End If
'   Next UK&
'   '
'   Screen.MousePointer = 1
'   '
End Sub

Sub CARLISCO()
   '
   Screen.MousePointer = 11   ' cargando Lista de Compra
   NPRO% = XVALO(Text1)
   'ESTO AGREGADO PARA QUE BORRE SI ALGUN ITEM NO TIENE CANTIDAD, ESTA VALIDACION ESTABA EN LA RUTINA CALTOCOM PERO SE QUITO
   'POR QUE TRAIA PROBELMAS CUANDO RECORRIA EL ARCHIVO OCOMDETA PARA ACTUALIZAR EN EL TIMER1 (15-0348)
   FIN& = ULTREG&("!OCOMDETA")
   For U& = 1 To FIN&
     X$ = REGLEIDO$("!OCOMDETA", U&)
     Ci% = CVI(Mid$(X$, 9, 2))
     '
     CANSTOK = CVS(Mid$(X$, 51, 4))
     If CANSTOK > 0.05 Then
       U1& = U1& + 1
       X$ = REGLEIDO$("!OCOMDETA", U1&)
     End If
   Next U&
   Call SETULTREG&("!OCOMDETA", U1&)
   'FIN (15-0348)
   
   HOII% = HOY(HOS$)
   FENTREGEN% = FECHANUM(Text14)
   List1.Clear: List1.ToolTipText = ""
   FIN& = ULTREG&("!OCOMDETA")
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   ARCHILIB$ = "!1OCOMDET"
   For U& = 1 To FIN&
     X$ = REGLEIDO$("!OCOMDETA", U&)
     Ci% = CVI(Mid$(X$, 9, 2))
     FENTRESOL% = CVI(Mid$(X$, 95, 2))
     If FENTRESOL% < FENTREGEN% Then FENTRESOL% = FENTREGEN%
     If FENTRESOL% < HOII% Then FENTRESOL% = HOII%
     If FENTRESOL% <> CVI(Mid$(X$, 95, 2)) Then
        Call REPLA(X$, MKI$(FENTRESOL%), 95, 2)
        Call GRAREG("!OCOMDETA", X$, U&)
     End If
     '
     UNIST$ = UCase$(TRIM$(Mid$(X$, 11, 4)))
     UNICO$ = UCase$(TRIM$(Mid$(X$, 59, 4)))
     CANSTOK = CVS(Mid$(X$, 51, 4))
     CANCOMP = CVS(Mid$(X$, 55, 4))
     'If CANSTOK < 0.05 Then
     '   Call MENSERR(24, "Falta Cantidad en " + UNIST$ + "\en Código '" + TRIM$(CODEXT$(CI0%)) + "'.")
     'End If
     '
     COEFI = CVS(Mid$(X$, 63, 4))
     If UNIST$ = UNICO$ Then COEFI = 1
     If COEFI < 0.0001 Then COEFI = 1
     '
     If CANCOMP < 0.05 Then
         CANCOMP = CANSTOK / COEFI
       Else
         CANSTOK = CANCOMP * COEFI
     End If
     CANCOMP = CANSTOK / COEFI
'     If InStr(EPAC$, "CONTEPLAST") > 0 And TIMATE%(Ci%) = 41 Then 'si es conteplast y tipomat es bobinas
'         StringNroCtes$ = StringDeNroProveedoresComprasInternas$
'         'SI EL NUMERO DE CLIENTE ESTA DENTRO DE LA CADENA DE NRO DE PROVEEDORES CONFIGURADOS COMO PARTICIPANTES DE LA COMPRA INTERNA.
'         If NumeroEstaEnCadena(StringNroCtes$, XVALO(Text1)) Then
'            Peso# = PESUNI(Ci%)
'            If Peso# > 0.00005 Then
'               coefi = 1 / Peso#
'               CANCOMP = CANSTOK / coefi 'SI SON BOBINAS
'               CANSTOK = CANCOMP * Peso#
'            End If
'         End If
'
'      End If
    If InStr(EPAC$, "CONTEPLAST") > 0 Or InStr(EPAC$, "CAMBACERES") > 0 Or InStr(EPAC$, "QUILMES") > 0 Then
         TMT% = TIMATE%(Ci%)
         UME$ = TRIM$(UCase$(Unimed$(Ci%)))
         If TMT% = 41 And UME$ = "M." Then ' tipomat es bobinas Y Unimed$ ES M.
            StringNroCtes$ = StringDeNroProveedoresComprasInternas$
            'SI EL NUMERO DE CLIENTE ESTA DENTRO DE LA CADENA DE NRO DE PROVEEDORES CONFIGURADOS COMO PARTICIPANTES DE LA COMPRA INTERNA.
            If NumeroEstaEnCadena(StringNroCtes$, XVALO(Text1)) Then
               Peso# = PESUNI(Ci%)
               If Peso# > 0.00005 Then
                  COEFI = 1 / Peso#
                  CANCOMP = CANSTOK / COEFI 'SI SON BOBINAS
                  CANCOMP = XVALO(FORMATNUM$(CANCOMP, "F12.0"))
                  CANSTOK = CANCOMP * Peso#
               End If
            End If
         ElseIf TMT% = 41 And UME$ = "U." Then 'ESTOY DENTRO DE CONTEP,CAMBA,QUILMES     Call REPLA(X$, MKS$(CANSTOK), 51, 4)
            StringNroCtes$ = StringDeNroProveedoresComprasInternas$
            'SI EL NUMERO DE CLIENTE ESTA DENTRO DE LA CADENA DE NRO DE PROVEEDORES CONFIGURADOS COMO PARTICIPANTES DE LA COMPRA INTERNA.
            If NumeroEstaEnCadena(StringNroCtes$, XVALO(Text1)) Then
               Peso# = PESUNI(Ci%)
               If Peso# > 0.00005 Then
                  COEFI = 1 / Peso#
                  CANCOMP = CANSTOK / COEFI 'SI SON BOBINAS
                  CANCOMP = XVALO(FORMATNUM$(CANCOMP, "F12.0"))
                  CANSTOK = CANCOMP * Peso#
               End If
            End If
         End If
     End If
     Call REPLA(X$, MKS$(CANCOMP), 55, 4)
     Call REPLA(X$, MKS$(COEFI), 63, 4)
     Call GRAREG("!OCOMDETA", X$, U&)
     '
     UNISTO$ = Mid$(X$, 11, 4)
     UNICOX$ = Mid$(X$, 59, 4)
     COEFUNI = CVS(Mid$(X$, 63, 4))
     PREUNI = CVD(Mid$(X$, 67, 8))
     'ACA CONTROLA SI LA CANTIDADES SON 0 NO LO PEGA AL LIST
     CANSTO = CVS(Mid$(X$, 51, 4))
     CANCOM = CVS(Mid$(X$, 55, 4))
     If CANSTO < 0.05 And CANCOM < 0.05 Then GoTo 30
     '
     PORDESCUIN = CVS(Mid$(X$, 75, 4))
     If PORDESCUIN < 0 Or PORDESCUIN > 99.9 Then PORDESCUIN = 0
     '
     TTXX$ = Space$(96)
     Call REPLA(TTXX$, CODEXT$(Ci%), 1, 15)
     Call REPLA(TTXX$, DESCRIT0$(Ci%), 17, 15)
       If COLTELIB% > 0 Then
         Call REPLA(TTXX$, RTrim$(REGLEIDO$(ARCHILIB$, U&)), 17, 15)
       End If
     Call REPLA(TTXX$, UNISTO$, 34, 4)
     Call REPLA(TTXX$, FORMATNUM$(CANSTO, "F8.1"), 38, 8)
     Call REPLA(TTXX$, UNICOX$, 48, 4)
     Call REPLA(TTXX$, FORMATNUM$(CANCOM, "F9.1"), 52, 9)
     MASCA$ = "F10.4"
     If PREUNI > 9999 Then MASCA$ = "F10.2"
     Call REPLA(TTXX$, FORMATNUM$(PREUNI, MASCA$), 61, 10)
     Call REPLA(TTXX$, FORMATNUM$(PORDESCUIN, "F5.1"), 72, 5)
     Call REPLA(TTXX$, FORMATNUM$(FENTRESOL%, "D8"), 78, 8)
     'List1.AddItem TRIM$(TTXX$)
     List1.AddItem TTXX$
     '
     'If CAN > 999999 Then
     '  Call MENSERR(24, "Cantidad Excesiva - Max 999999")
     '  Call REPLA(X$, MKS$(0), 47, 4)
     '  Call GRAREG("!CARDETPE", X$, U&)
     '  GoTo 10
     'End If
     'If CAN < 0.5 Then
     '  Call MENSERR(24, "Falta Cantidad")
     '  Call REPLA(X$, MKS$(0), 47, 4)
     '  Call GRAREG("!CARDETPE", X$, U&)
     '  GoTo 10
     'End If
     '
     '
30 Next U&
   'Call CIERRARCH("!CARDETPE")
   'Call CALTOPED
   List1.Refresh
   '
   
   Screen.MousePointer = 1
   If List1.ListCount = 0 Then Call BLANFORMU: Text1 = ""
End Sub

Sub SHOWOCOM(NPX&)
    '
    Screen.MousePointer = 11
    TIDOCUM$ = "Nota de Pedido"
    EPAC$ = UCase$(EMPREAC$)
    If InStr(EPAC$, "AHP") > 0 Then TIDOCUM$ = "ORDEN DE COMPRA"

    NUDOCU$ = TRIM$(Str$(NPX&))
    DOCUNU& = 0
    '
    NoPuedeImprimir% = 0
    MENSAJEDOC$ = ""
    CONDI$ = "NUME_OCOM = " & NPX&
    ST% = XVALO(VALOBADA("OCOMENCA", "STAT_OCOM", CONDI$, "NOMESS"))
    If ST% < 0 Then
        NoPuedeImprimir% = 1
        MENSAJEDOC$ = "Imposible Imprimir Esta Orden de Compra, La Misma Está Pendiente de Aprobación"
    End If
    Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, NUDOCU$, , , MENSAJEDOC$, NoPuedeImprimir%)
    NoPuedeImprimir% = 0
    MENSAJEDOC$ = ""

'    TIDOCUM$ = "Nota de Pedido"
'    NUDOCU$ = TRIM$(Str$(NPX&))
'    While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
'    NUDOCU1$ = " " + NUDOCU$ + " "
'    NUDOCU2$ = " " + NUDOCU$ + "-"
'    '
'    Screen.MousePointer = 11
'    '
'    For UI& = ULTREG&("PDOCLST") To 1 Step -1
'      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
'      If InStr(XX$, TIDOCUM$) > 4 Then
'        PPXX% = InStr(XX$, "Nro.")
'        If PPXX% > 0 Then
'          XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
'        End If
'        If InStr(XX$, NUDOCU1$) > 4 Or InStr(XX$, NUDOCU2$) > 4 Then
'          DOCUNU& = CVS(Left$(XX$, 4))
'          GoTo 5
'        End If
'      End If
'    Next UI&
'    Screen.MousePointer = 1
'    Call COMUNI("Documento no Encontrado")
'    Exit Sub
'    '
'5   Call MUESTRADOC(DOCUNU&)
    '
End Sub

Sub PRICOMPRA()
    '
    MONEXA$ = TRIM$(Label11)
    CPAG% = XVALO(TEXT21(0))
    FORPAG$ = TRIM$(Eco21(0))
    If FORPAG$ = "" Then FORPAG$ = "Contado Contra Entrega"
    TIEM% = XVALO(TEXT21(1))
    EMBALA$ = TRIM$(Eco21(1))
    '
    Call SETULTREG("!OBSERVOF", 0)
    Call FRATEXTO(Text13, 52)
    For KKU& = 1 To CARETEX%
      Call REGAPP("!OBSERVOF", RETEX$(KKU&))
    Next KKU&
    Call CIERRARCH("!OBSERVOF")
    '
    PRIPLANOI% = 0
    If Control$("OCOMPRA", "PRIPLANO") = "SI" Then PRIPLANOI% = 1
    PRIESPEC% = 0
    If Control$("OCOMPRA", "PRIESPEC") = "SI" Then PRIESPEC% = 1
    PRIHRINS% = 0
    If Control$("OCOMPRA", "PRIHRINS") = "SI" Then PRIHRINS% = 1
    '
    ' IMPRESION DE LA O/COMPRA
    '
    Call LIMPIATETAB
    FORIPRE$ = Control$("", "FORMOCOM")
    '
    If TRIM$(FORIPRE$) <> "" Then
      '
      'SI EXISTEN LAS TABLAS PLAFABR Y OCOEMIHIS VA A TRAER INFORMACION DEL CALCULO MRP (SOLICITUD DE AMBEST)
      INFOMRP% = 0
      If Control("OCOMPRA", "IMPRCMRP") = "SI" Then ' configuracion en setup de compras
        If EXISTE_TABLA%("COBERFA") Then
          If EXISTE_TABLA%("COBERFA") Then
            INFOMRP% = 1
          End If
        End If
      End If
      
      Screen.MousePointer = 11
      FECHEMII% = FECHANUM(Text9)
      FECHDOC$ = FETEXTO$(FECHEMII%)
      NUMEDOC$ = TRIM$(Text11)
      While Len(NUMEDOC$) < 6
        NUMEDOC$ = "0" + NUMEDOC$
      Wend
      '\\\OT-06-0082-WAR-06/03/06///
'      NUMEDOC$ = NUMEDOC$ + "-" + "0"
      'PARA EVITAR QUE EN AHP IMPRIMA LA VERSION
      EPAC$ = UCase$(EMPREAC$)
      If InStr(EPAC$, "AHP") < 1 Then
         NUMEDOC$ = NUMEDOC$ + "-" + TRIM$(Label16)
      End If

      '\\\OT-06-0082-WAR-FIN///
      TIPODOC$ = "Nota de Pedido"
      If InStr(EPAC$, "AHP") > 0 Then TIPODOC$ = "ORDEN DE COMPRA"
      '
      CODPARAM$(1) = "COND-PAG"
      CODPARAM$(2) = "NUBULTOS"
      CODPARAM$(3) = "IMP-NETO"
      CODPARAM$(4) = "POR-IVA"
      CODPARAM$(5) = "IMP-IVA"
      CODPARAM$(6) = "IMP-TOTA"
      CODPARAM$(7) = "VENDEDOR"
      CODPARAM$(8) = "LIST-PRE"
      CODPARAM$(9) = "ZONA-VEN"
      CODPARAM$(10) = "FECH-LAN"
      CODPARAM$(11) = "IMPNEDES"
      CODPARAM$(12) = "TRANSPOR"
      CODPARAM$(13) = "POR-DESC"
      CODPARAM$(14) = "IMP-DESC"
      CODPARAM$(15) = "FECH-REC"
      CODPARAM$(16) = "TIPO-VAL"
      CODPARAM$(17) = "IMPO-CHE"
      CODPARAM$(18) = "REF-CONJ"
      CODPARAM$(19) = "DOMI-TRA"
      CODPARAM$(20) = "PERC-IVA"
      '\\\OT-05-0676-WAR-20/10/05///
      CODPARAM$(21) = "DOCU-APL" 'REGLAMENTOS DE COMPRA
      '\\\OT-05-0676-WAR-FIN///
      CODPARAM$(22) = "REF-MAT2"
      CODPARAM$(26) = "ACH-CUCO"
      CODPARAM$(27) = "PES-CONJ"
      CODPARAM$(28) = "LAR-CUTE"
      CODPARAM$(29) = "LAR-MATE"
      
     '\\\OT-08-0969-OD-24/10/08/// se modifico capaardoc% a 25 para no sobreescribir
     'importe en letras
      CAPARDOC% = 29     ' para importe en letras
      '\\\OT-08-0969-OD-FIN///
      '
      DOCPARAM$(1) = FORPAG$
      DOCPARAM$(2) = EMBALA$
      DOCPARAM$(3) = Total(0)
      DOCPARAM$(4) = Label118    '  Text12
         If UCase$(Left$(Label118, 3)) = "VAR" Then DOCPARAM$(4) = "Ver Nota"
      DOCPARAM$(5) = Total(2)
      DOCPARAM$(6) = Total(3)
      DOCPARAM$(7) = "" 'Mid$(X1$, 147, 20)
      DOCPARAM$(8) = "" 'Mid$(X1$, 67, 40)
      DOCPARAM$(9) = "" 'Mid$(X1$, 167, 40)
      DOCPARAM$(10) = "" ' Mid$(X1$, 107, 40)
      DOCPARAM$(11) = Str$(XVALO(Total(0)) - XVALO(Total(1)))
      DOCPARAM$(12) = TRIM$(Text6) ' ATENCION SR.
      DOCPARAM$(13) = TRIM$(PORDESCU)
      DOCPARAM$(14) = Total(1)
      HOII% = HOY(HOS$)
      FECHENT% = FECHANUM(Text14)
      DOCPARAM$(15) = Text14
      If FECHENT% = HOII% Then
          DOCPARAM$(15) = DOCPARAM$(15) + " (inmediato)"
        ElseIf FECHENT% < HOII% + 4 Then
          DOCPARAM$(15) = DOCPARAM$(15) + " (urgente!!)"
      End If
      DOCPARAM$(16) = MONEXA$
      DOCPARAM$(17) = ""
      If MONEMI% > 1 Then DOCPARAM$(17) = Str$(TICAMONE(MONEMI%))
      DOCPARAM$(18) = Eco21(1)
      DOCPARAM$(19) = Text10
      DOCPARAM$(20) = Total(4)
      '\\\OT-05-0676-WAR-20/10/05///
      DOCPARAM$(21) = RAI$ 'REGLAMENTO COMPRA ANEXO
      '\\\OT-05-0676-WAR-FIN///
      '\\\OT-06-0483-RG-14/11/06///
      '\\\OT-08-0969-OD-24/10/08/// el label decia label 5 se modifico a label 15
      'que es la calificacion del proveedor
      DOCPARAM$(22) = Label15.Caption
      DOCPARAM$(26) = TRIM$(Text12)
      DOCPARAM$(27) = TXTPESO
      'AGREGADO PARA ARCON SOLO PARA LAS SUCURSALES QUE IMPRIMA EL CLIENTE Y LA O.C. DEL MISMO
      DOCPARAM$(28) = Left$(TRIM$(TEXT24) & " - " & TRIM$(Eco24), 45)
      DOCPARAM$(29) = Left$(TEXT25, 16)

      '\\\OT-08-0969-OD-fin///
      '\\\OT-06-0483-RG-FIN///
      '
      ' esto es para poner el importe en letras
      TTTT# = XVALO(Total(3))
      TOTFAC$ = NUMTEX$(TTTT#)
      If Label11 <> "" And Label11 <> "Pesos" Then
         PPXX% = InStr(TOTFAC$, "PESOS")
         If PPXX% > 0 Then
           TOTFAC$ = Left$(TOTFAC$, PPXX% - 1) + UCase$(Label11) + Mid$(TOTFAC$, PPXX% + 5)
         End If
      End If
      Call FRATEXTO(TOTFAC$, 40)
      '\\\OT-08-0969-OD-24/10/08///SE RENUMERO (+1)POR QUE SE UTILIZO EL 22 PARA LA
      'CALIFICACION DEL PROVEEDOR

      CODPARAM$(23) = "IMP-LET1": DOCPARAM$(23) = RETEX$(1)
      CODPARAM$(24) = "IMP-LET2": DOCPARAM$(24) = RETEX$(2)
      CODPARAM$(25) = "IMP-LET3": DOCPARAM$(25) = RETEX$(3)
      '\\\OT-08-0969-OD-FIN
      '
      CODTABU2$(1) = "REF-MAT3"
      CACOLTAB2% = 1
      '
      CODTABU1$(1) = "COD-MAT"
      CODTABU1$(2) = "DES-MAT"
      CODTABU1$(3) = "UNIMED"
      CODTABU1$(4) = "CANTIDAD"
      CODTABU1$(5) = "FECH-VEN"
      CODTABU1$(6) = "REF-MAT1"
      CODTABU1$(7) = "PRE-LIST"
      CODTABU1$(8) = "PORDESCU"
      CODTABU1$(9) = "PRE-NETO"
      CODTABU1$(10) = "NET-ITEM"
      CODTABU1$(11) = "CAKILOS"
      CODTABU1$(12) = "F-PLANO"
      CODTABU1$(13) = "ORD-ITEM"
      CODTABU1$(14) = "DESTINO"
      CODTABU1$(15) = "EQUIPO1"   ' para sector solicitante
      CODTABU1$(16) = "MEDIDA-1"
      CODTABU1$(17) = "MEDIDA-2"
      
      CODTABU1$(18) = "REF-MAT4"
      CODTABU1$(19) = "NUBULTOS"
      CODTABU1$(20) = "REF-MAT5"
      CODTABU1$(21) = "ACH-MACO"
      CODTABU1$(22) = "ORD-ITEM"
      CODTABU1$(23) = "REF-MAT6"
      
      CACOLTAB1% = 23
      '
      DESTIDOC% = (-1) * XVALO(Text1)
      CAITAB1% = 0: CAITAB2% = 0
      CAITCRU% = 0
      DESCPRI% = 0: DESCOMP% = 0: INCLUSPE% = 0
      If Control("OCOMPRA", "IMPDESPR") = "SI" Then DESCPRI% = 1
      If Control("OCOMPRA", "IMPDESCO") = "SI" Then DESCOMP% = 1
      If Control$("OCOMPRA", "INCLUSPE") = "SI" Then INCLUSPE% = 1
      If DESCOMP% + INCLUSPE% < 1 Then DESCPRI% = 1 'POR SI NO SE SELECCIONO NADA

      '
      REPLAORI% = 0
      If Control$("OCOMPRA", "REPLAORI") = "SI" Then
         REPLAORI% = 1
      End If
      '
      RECOPROP% = 0
      If Control$("OCOMPRA", "RECOPROP") = "SI" Then
         RECOPROP% = 1
      End If
      '
      REVICODX% = 0
      If Control$("OCOMPRA", "REVICOD") = "SI" Then
         REVICODX% = 1
      End If
      '
      NC1% = DESTIDOC%
      ARCHILIB$ = "!1OCOMDET"
      '\\\OT-08-1040-OD-13/11/08/// ' SE INTRODUCE PARA LA IMPRESION DE LA
      ' ESPECIFICACION DEL ARTICULO
      FORTELIB$ = TRIM$(UCase$(ATRIFORM$(FORIPRE$, "DES-MAT/FORMATO")))
      ATELIB% = XVALO(Mid$(FORTELIB$, 2))
      '\\\OT-08-1040-OD-FIN///
      If ATELIB% < 16 Then ATELIB% = 16
      ORD% = 0
      For J& = 1 To ULTREG&("!OCOMDETA")
         X$ = REGLEIDO$("!OCOMDETA", J&)
         CI0% = CVI(Mid$(X$, 9, 2))
         CANTU = CVS(Mid$(X$, 55, 4))
         PreUnid# = CVD(Mid$(X$, 67, 8))
         PORDE = CVS(Mid$(X$, 75, 4))
         FEX = CVI(Mid$(X$, 95, 2))
         CONSIMAT% = 0: If TRIM$(Mid$(X$, 97, 2)) <> "" Then CONSIMAT% = 1
         '
         CAITAB1% = CAITAB1% + 1
         CODDX$ = TRIM$(CODEXT$(CI0%))
         CPROP$ = CODDX$
         If REPLAORI% = 1 Then
            CODDX$ = TRIM$(CODORIG$(CI0%, NC1%))
         End If
         FEREVX = FEREVI%(CI0%)
         RVCDDX$ = TRIM$(REVICOD$(CI0%))
         FECHYREVI$ = ""
         If REVICODX% = 1 Then
            If RVCDDX$ <> "" Then
              CODDX$ = CODDX$ + " (Letra: " + RVCDDX$ + " - " + FECHATEX$(FEREVX) + ")"
            End If
         Else 'esto se agrega por que eym pide que aparezca en otra columna la revision
            FECHYREVI$ = ""
            If RVCDDX$ <> "" Then
               'FECHYREVI$ = "Rev.: " & RVCDDX$
               FECHYREVI$ = TRIM$(RVCDDX$)
               SOLOREVI$ = FECHYREVI$
               If FEREVX > 33 Then FECHYREVI$ = FECHYREVI$ & "-(" & FECHATEX$(FEREVX) & ")"
            End If
         End If
         '
         'TRES DESCRIPCIONES POSIBLES
         DESPRIN$ = TRIM$(DESCRIT0$(CI0%))
         DESCOMPLEM$ = DESCRIT2$(CI0%)
         ESPECIFICAC$ = ESPECIFIT$(CI0%)
         '
         TETABU1$(1, CAITAB1%) = CODDX$
         If REPLAORI% = 1 Then
             If CODDX$ <> CPROP$ Then
               If RECOPROP% > 0 Then
                 DESPRIN$ = DESPRIN$ + " (" + CPROP$ + ")"
               End If
             End If
         End If
         '
         TETABU1$(3, CAITAB1%) = Mid$(X$, 59, 4)  'UNIDAD DE COMPRA
         If TRIM$(Mid$(X$, 59, 4)) = "" Then TETABU1$(3, CAITAB1%) = Mid$(X$, 11, 4) 'UNIDAD DE STOCK
         '
         TETABU1$(4, CAITAB1%) = TRIM$(Str$(CANTU))
         TETABU1$(5, CAITAB1%) = FECHATEX$(FEX)
         TETABU1$(6, CAITAB1%) = Mid$(X$, 97, 2)  'consigna m.prima
         TETABU1$(7, CAITAB1%) = Str$(PreUnid#)
         TETABU1$(8, CAITAB1%) = Str$(PORDE)
         TETABU1$(9, CAITAB1%) = Str$(PreUnid# * (1 - PORDE / 100))
         TETABU1$(10, CAITAB1%) = Str$(CANTU * PreUnid# * (1 - PORDE / 100))
           KILOX$ = FORMATNUM$(CVS(Mid$(X$, 51, 4)) * PESUNI(CI0%), "F12.3")
         TETABU1$(11, CAITAB1%) = KILOX$
         TETABU1$(12, CAITAB1%) = FEPLANO$(CI0%)
         TETABU1$(13, CAITAB1%) = TRIM$(Str$(J&))
         TETABU1$(14, CAITAB1%) = TRIM$(Mid$(X$, 105, 12))   ' Imputacion contable
         TETABU1$(15, CAITAB1%) = TRIM$(Mid$(X$, 15, 12))   ' Sector
         TETABU1$(16, CAITAB1%) = TRIM$(Mid$(X$, 31, 20))
         TETABU1$(17, CAITAB1%) = TRIM$(CODORIG$(CI0%, NC1%))  '08/10/09 - Proveedor
         TETABU1$(18, CAITAB1%) = TRIM$(PRESENVASE$(CI0%)) 'PRESENTACION DEL BULTO
         CANTID# = CANTU
         TETABU1$(19, CAITAB1%) = TRIM$(CANBULTOS(CI0%, CANTID#)) ' CANTIDAD DE BULTOS
         TETABU1$(20, CAITAB1%) = TRIM$(FECHYREVI$)
           PORNA = PORNACIO(CI0%)
         TETABU1$(21, CAITAB1%) = TRIM$(FORMATNUM$(PORNA, "F7.2")) 'PORCENTAJE DE NACIONALIZACION
         TETABU1$(22, CAITAB1%) = TRIM$(ORD% = ORD% + 1) 'ORDEN DE ITEM
         TETABU1$(23, CAITAB1%) = SOLOREVI$ ' SOLO EL ID DE REVISION DEL ARTICULO
         '
         POCODI% = XVALO(ATRIFORM$(FORIPRE$, "COD-MAT/POSINHOR"))
         PODESI% = XVALO(ATRIFORM$(FORIPRE$, "DES-MAT/POSINHOR"))
         If PODESI% <= POCODI% + 3 Then
           If TRIM$(CODDX$) <> "" Then
             CAITAB1% = CAITAB1% + 1
           End If
         End If
         '
'        TETABU1$(2, CAITAB1%) = DECRIT$
         If COLTELIB% > 0 Then
               DESPRIN$ = RTrim$(REGLEIDO$(ARCHILIB$, J&))
               '
               'LO SIGUIENTE SE AGREGA PORQUE EL RTRIM NO QUITA LOS CARACTERES
               'DE CONTROL QUE SE ENCUENTRAN EN DESPRIN$ A LA DERECHA DEL TEXTO
               DESPRIN$ = REPLACAR(DESPRIN$, Chr$(0), Chr$(32))
               DESPRIN$ = RTrim$(DESPRIN$)
               '
               If REPLAORI% = 1 Then
                 If CODDX$ <> CPROP$ Then
                   If RECOPROP% > 0 Then
                     DESPRIN$ = DESPRIN$ + " (" + CPROP$ + ")"
                   End If
                 End If
               End If
         End If
         '''
         '''
         UNIST$ = Mid$(X$, 11, 4)
         UNICO$ = UCase$(TRIM$(Mid$(X$, 59, 4)))
         DEADI$ = ""
         If TRIM$(Left$(UCase$(UNIST$), 2)) <> UNICO$ Then
           If CI0% > 0 Then  ' que no lo haga en texto libre
              DEADI$ = "(equiv. " + TRIM$(FORMATNUM$(CVS(Mid$(X$, 51, 4)), "F12.1")) + " " + TRIM$(ECOARCH$("UNIMED", UNIST$)) + ".)"
           End If
         End If
         '
         'impresion de datos del mrp
         If INFOMRP% > 0 Then
         
         
         End If
         'DESCRIPCION
         DIMPRE$ = ""
         If DESCPRI% > 0 Then
               If DESPRIN$ <> "" Then
                  DIMPRE$ = DESPRIN$
               End If
         End If
            '
        'DESCRIPCION ADICIONAL
        If DEADI$ <> "" Then
             If DIMPRE$ <> "" Then
               DIMPRE$ = DIMPRE$ + vbCrLf
             End If
             DIMPRE$ = DIMPRE$ + DEADI$
         End If

        'DESCRIPCION COMPLEMENTARIA
        If DESCOMP% > 0 Then
           If DESCOMPLEM$ <> "" Then
             If DIMPRE$ <> "" Then
               DIMPRE$ = DIMPRE$ + vbCrLf
             End If
             DIMPRE$ = DIMPRE$ + DESCOMPLEM$
           End If
        End If
        '
        'ESPECIFICACION
        If INCLUSPE% > 0 Then
           If ESPECIFICAC$ <> "" Then
             If DIMPRE$ <> "" Then
               DIMPRE$ = DIMPRE$ + vbCrLf
             End If
             DIMPRE$ = DIMPRE$ + ESPECIFICAC$
           End If
        End If
         
         
''''         '\\\OT-08-1040-OD-13/11/08///
''''         If INCLUSPE > O Then
'''''           INTERLI = XVALO(ATRIFORM$(FORIPRE$, "INTERLIN"))
''''           XX$ = ESPECIFIT$(CI0%)
''''           If TRIM$(XX$) <> "" Then
'''''              COLUTA% = 0
'''''              If DESCPRI% = 1 Or DESCOMP% = 1 Then CAITAB1% = CAITAB1% + 1
''''               If DIMPRE$ <> "" Then DIMPRE$ = DIMPRE$ + vbCrLf
''''               DIMPRE$ = DIMPRE$ & XX$
'''''              Call FRATEXTO(XX$, ATELIB%)
'''''              For JJ& = 1 To CARETEX%
'''''                If INTERLI <= 6 Or InStr(UCase$(FORIPRE$), "MZZ") > 0 Then
'''''                  TETABU1$(2, CAITAB1%) = RETEX$(JJ&)
'''''                  CAITAB1% = CAITAB1% + 1
'''''                Else
'''''                  COLUTA% = COLUTA% + 1
'''''                  If COLUTA% > 2 Then COLUTA% = 1
'''''                  If COLUTA% = 1 Then CAITAB1% = CAITAB1% + 1
'''''                  TETABU1$(COLUTA%, CAITAB1%) = RETEX$(JJ&)
'''''                End If
'''''              Next JJ&
'''''              CAITAB1% = CAITAB1% - 1
''''           End If
''''         End If
         '\\\OT-08-1040-OD-FIN///
         '
         'IMPRESION DE LA DESCRIPCION CON TODAS SUS VARIANTES CONFIGURADAS
         FORTELIB$ = TRIM$(UCase$(ATRIFORM$(FORIPRE$, "DES-MAT/FORMATO")))
         ATELIB% = XVALO(Mid$(FORTELIB$, 2))
         Call FRATEXTO(DIMPRE$, ATELIB%)
         For KKU1% = 1 To CARETEX%
                  TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
                  CAITAB1% = CAITAB1% + 1
         Next KKU1%
         'SE RESTA CAITAB1% PORQUE DEL FOR SALE EN LA PROXIMA LINEA
         'LO CUAL YA SE HACE CUANDO PASA AL PROXIMO CODIGO (VER MAS ARRIBA)
         CAITAB1% = CAITAB1% - 1
         '
         'LINEA DE SEPARACION
         Call CARLINSEP  'Carga linea de separacion entre codigo y codigo

'         If INCLUSPE% > 0 Then
'           RNC$ = RECFILT$("ESPECIFI", 1, MKI$(CI0%))
'           If Len(RNC$) > 0 Then
'             INTERLI = XVALO(ATRIFORM$(FORIPRE$, "INTERLIN"))
'             '
'             COLUTA% = 0
'             URN& = ULTREG&("ESPECIFI")
'             '
'             For U& = 1 To Len(RNC$) Step 4
'               RENOTA& = CVS(Mid$(RNC$, U&, 4))
'               If RENOTA& > 0 Then
'                 If RENOTA& <= URN& Then
'                   TTXY$ = RTrim$(Mid$(REGLEIDO$("ESPECIFI", RENOTA&), 3))
'                   TTYY$ = REPLACAR$(TTXY$, Chr$(9), "     ")
'                   If INTERLI <= 6 Or InStr(UCase$(FORIPRE$), "MZZ") > 0 Then
'                       CAITAB1% = CAITAB1% + 1
'                       TETABU1$(2, CAITAB1%) = TTYY$
'                     Else
'                       COLUTA% = COLUTA% + 1
'                       If COLUTA% > 2 Then COLUTA% = 1
'                       If COLUTA% = 1 Then CAITAB1% = CAITAB1% + 1
'                       TETABU1$(COLUTA%, CAITAB1%) = TTYY$
'                   End If
'                 End If
'               End If
'             Next U&
'             '
'           End If
'         End If
         '
109   Next J&
      '
      ' IMPRIME DETALLE DE IVA
      If UCase$(Left$(Label118, 3)) = "VAR" Then
         CAITAB2% = 0
         With RIVACOM07
           For KKII% = 0 To 5
             If XVALO(.IMPOGRA(KKII%)) >= 0.005 Or XVALO(.IMPOIVA(KKII%)) >= 0.005 Then
               TTXX$ = AJUSTI$(.TASIVA(KKII%), 8)
               TTXX$ = TTXX$ + FORMATNUM$(XVALO(.IMPOGRA(KKII%).TEXT), "F11.2")
               TTXX$ = TTXX$ + FORMATNUM$(XVALO(.IMPOIVA(KKII%)), "F10.2")
               CAITAB2% = CAITAB2% + 1
               TETABU2$(1, CAITAB2%) = TTXX$
             End If
           Next KKII%
         End With
      End If
      '
      If InStr(1, EMPREAC$, "CONTEPLAST") < 1 Then
        Call PRINTDOC("FORMOCOM")   ' Orden de Compra
      Else
        Opcion% = COMALTER%("Elija Formulario de Impresión\\Conteplast\Envaplast")
        If Opcion% = 1 Then
            Call PRINTDOC("@OCOM-CTP")
        Else
            Call PRINTDOC("FORMOCOM")
        End If
      End If
      '
      For KK1% = 1 To CACOLTAB1%
        For KK2% = 1 To CAITAB1%
          TETABU1$(KK1%, KK2%) = ""
        Next KK2%
      Next KK1%
      '
      DOCUORIG$ = "N/Ped-" + NUMEDOC$
      NUMEDOC$ = ""
      '
      EPAC$ = TRIM$(UCase$(EMPREAC$))
      If InStr(EPAC$, "ARCON") > 0 Then
         If InStr(Me.Caption, "Pedido de Reposi") < 1 Then
           Call CIERRARCH("*.*")
           Call BAJALDISCO
           ANEXOCOM.Show 1
           GoTo 199
         End If
      End If
      '
      For J& = 1 To ULTREG&("!OCOMDETA")
         X$ = REGLEIDO$("!OCOMDETA", J&)
         CI0% = CVI(Mid$(X$, 9, 2))
         If CI0% > 0 Then
           If CI0% <= NUMAS% Then
             If PRIPLANOI% = 1 Then Call PRIPLANO(CI0%)
             If PRIESPEC% = 1 Then Call PRIHOJESP(CI0%)
             If PRIHRINS% = 1 Then Call PRIHOJRINSP(CI0%)
           End If
         End If
      Next J&
      '
    End If
199 Call CIERRARCH("*.*")
    Call BAJALDISCO
    Screen.MousePointer = 1
    '
    If EMICONSI% = 1 Then
      If SENDBYMAIL$ = "" Then   ' solo si es por impresion
         If COMALTER%("Puede Ahora Generar\el Remito de Consignacion de Materiales\\Consignar Materiales\Continuar") = 1 Then
           NPRO% = XVALO(Text1)
           'Call ENTRECONSIACC(NPRO%)
           NCOMX$ = NROCOM$(XVALO(Text11))
           Call NUENTRECONSI(NPRO%, NCOMX$)
         End If
      End If
    End If
    Call CIERRARCH("*.*")
    '
    Exit Sub
    '
999 Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
End Sub
Sub GENERALISTAPICK(NROCOMX$, Optional MODO$)
    '
    Screen.MousePointer = 11
    Call CREATABLA_PICK
    '
    'BORRO LOS REGISTROS MARCADOS CON MARBORRA
    CONDI$ = "COD_INTE > 0"
    Call BORRAREGISTROS("PICKCOMPRA", CONDI$, REGAF&, "NOMESS")
    '
    DEPOCON% = XVALO(Control$("", "DEPOCONS"))
    '
    SQLX$ = "SELECT * FROM MACONSIG WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE Nro_Ocom Like '" & NROCOMX$ & "%'"
    SQLX$ = SQLX$ + " ORDER BY NUORDITEM ASC "
    Dim MACONTMP As ADODB.Recordset
    Set MACONTMP = READSET(SQLX$)
    '
    Dim MOVTMP As ADODB.Recordset
    '
    With MACONTMP
      Do While Not .EOF
        'LEO EL MACONSIG
        Ci% = XVALO(!Cod_Inte)
        If Ci% <= 0 Then GoTo 40
        '
        NORD% = XVALO(!NUORDITEM)
        Asig = XVALO(!cant_asig)
        ENTREG = XVALO(!cant_entreg)
        Remito$ = TRIM$(SAFETEXT$(!NREMI_CONSIG))
        UBICA$ = Left$(UBISTO$(Ci%), 64)
        'STOCK1 = STOCKDIS(CI%)
        STOCK1 = STODEPOS(Ci%, DEPOCON%) - STOCKRES(Ci%)
        LOTE$ = ""
           'COMPRAS PENDIENTES DE ENTREGA
           CONDI$ = "CANT_REC < CANT_OCOM  AND STAT_OCOM < 3 AND COD_INTE = " & Ci%
        COMP = XVALO(VALOBADA("OCOMDETA", "SUM(CANT_OCOM) - sum(CANT_REC)", CONDI$, "NOMESS"))
           'RECEPCI0NES PENDIENTES DE APROBACION
           CONDI$ = "CANTNOM > CANTOAPRO AND COD_INTE = " & Ci%
        RECEP = XVALO(VALOBADA("BOLRECEP", "SUM(CANTNOM) - sum(CANTOAPRO)", CONDI$, "NOMESS"))
        
        If Remito$ = "" Then
            Call AGREGAREGISTROPICK(Ci%, NORD%, Asig, ENTREG, "", UBICA$, STOCK1, COMP, RECEP, 0, LOTE$, 0)
        Else
            'RECORRO EL MOVISTO
            SQLX$ = "SELECT * FROM MOVISTO WITH(NOLOCK)"
            SQLX$ = SQLX$ + " WHERE DOSECLAR LIKE  '" & NROCOMX$ & "%'"
            SQLX$ = SQLX$ + " AND COD_INTE = " & Ci%
           
            Set MOVTMP = READSET(SQLX$)
            '
            With MOVTMP
              Do While Not .EOF
                 CANTI = XVALO(!CANTSALID) - XVALO(!CANTINGRE)
                 LOTE$ = SAFETEXT$(!idenlote)
                 FECH1% = CVINT(!FechMov)
                 Remito$ = SAFETEXT$(!DOPRILAR)
                 Call AGREGAREGISTROPICK(Ci%, NORD%, Asig, ENTREG, Remito$, UBICA$, STOCK1, COMP, RECEP, FECH1%, LOTE$, CANTI)
                 .MoveNext
              Loop
            End With
            MOVTMP.Close
            Set MOVTMP = Nothing
            '
        End If
40      .MoveNext
      Loop
    End With
    MACONTMP.Close
    Set MACONTMP = Nothing
    '
    Screen.MousePointer = 1
    NORCOMP& = XVALO(Mid$(NROCOMX$, 4, 6))
    FILTX$ = TRIM$(Str$(NORCOMP&))
    Call STDFORM("MACO-MDT", FILTX$, MODO$)

End Sub

Sub AGREGAREGISTROPICK(CI1%, NUMEORDEN%, CANTIASIG, CANTIENTRE, REMITOCONS$, UBICA1$, STOCKDISPONIBLE, COMPRADO, RECIBIDO, FECHAMOVIMIENTO%, LOTE1$, CANTI)
   
   Dim Rst As ADODB.Recordset
   Set Rst = New ADODB.Recordset
   '
   Rst.Open "select * from pickcompra where cod_inte<0", FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   With Rst
        .AddNew
        !UBICACION = UBICA1$
        !Cod_Inte = CI1%
        !NUORDITEM = NUMEORDEN%
        !cant_asig = XVALO(FORMATNUM$(CANTIASIG, "F14." & CNTDC$))
        !cant_entreg = XVALO(FORMATNUM$(CANTIENTRE, "F14." & CNTDC$))
        !Remito = REMITOCONS$
        !stock = XVALO(FORMATNUM$(STOCKDISPONIBLE, "F14." & CNTDC$))
        !compras = XVALO(FORMATNUM$(COMPRADO, "F14." & CNTDC$))
        !recepcion = XVALO(FORMATNUM$(RECIBIDO, "F14." & CNTDC$))
        Faltante# = 0
        SaldoEntrega# = !cant_asig - !cant_entreg
        If SaldoEntrega# > 0.05 Then
            Existencia# = !stock
            If SaldoEntrega# > Existencia# Then
               Faltante# = SaldoEntrega# - Existencia#
            End If
        End If
        !Faltante = Faltante#
        !FechMov = CVDAT(FECHAMOVIMIENTO%)
        !LOTE = LOTE1$
        !Cantidad = CANTI
        .Update
   End With
   
   Rst.Close
   Set Rst = Nothing
   
'10 RNDX$ = RANDSTRING$(-8)
'   If CantRegistros("PICKCOMPRA", "UBICACION = '" & RNDX$ & "'", "NOMESS") > 0 Then GoTo 10
'
'   Call CreaRegistro("PICKCOMPRA", "UBICACION", RNDX$, "NOMESS")
'   CONDI$ = "UBICACION = '" & RNDX$ & "'"
'   Call ACTUREGISTRO("PICKCOMPRA", "COD_INTE", CONDI$, CI1%, REGAF&, "NOMESS")
'   Call ACTUREGISTRO("PICKCOMPRA", "NUORDITEM", CONDI$, NUMEORDEN%, REGAF&, "NOMESS")
'   Call ACTUREGISTRO("PICKCOMPRA", "CANT_ASIG", CONDI$, FORMATNUM$(CANTIASIG, "F14." & CNTDC$), REGAF&, "NOMESS")
'   Call ACTUREGISTRO("PICKCOMPRA", "CANT_ENTREG", CONDI$, FORMATNUM$(CANTIENTRE, "F14." & CNTDC$), REGAF&, "NOMESS")
'   Call ACTUREGISTRO("PICKCOMPRA", "REMITO", CONDI$, REMITOCONS$, REGAF&, "NOMESS")
'   Call ACTUREGISTRO("PICKCOMPRA", "STOCK", CONDI$, FORMATNUM$(STOCKDISPONIBLE, "F14." & CNTDC$), REGAF&, "NOMESS")
'   Call ACTUREGISTRO("PICKCOMPRA", "Compras", CONDI$, FORMATNUM$(COMPRADO, "F14." & CNTDC$), REGAF&, "NOMESS")
'   Call ACTUREGISTRO("PICKCOMPRA", "RECEPCION", CONDI$, FORMATNUM$(RECIBIDO, "F14." & CNTDC$), REGAF&, "NOMESS")
'   Call ACTUREGISTRO("PICKCOMPRA", "FECHMOV", CONDI$, FETEXCOR1$(FECHAMOVIMIENTO%), REGAF&, "NOMESS")
'   Call ACTUREGISTRO("PICKCOMPRA", "LOTE", CONDI$, LOTE1$, REGAF&, "NOMESS")
'   Call ACTUREGISTRO("PICKCOMPRA", "CANTIDAD", CONDI$, CANTI, REGAF&, "NOMESS")
'
'   '
'   Call ACTUREGISTRO("PICKCOMPRA", "UBICACION", CONDI$, UBICA1$, REGAF&, "NOMESS")
   '
End Sub
Sub CREATABLA_PICK()

    Call CREACAMPO("COMPRAS", "PICKCOMPRA", "NUORDITEM", 3, 0)
    Call CREACAMPO("COMPRAS", "PICKCOMPRA", "COD_INTE", 3, 0)
    Call CREACAMPO("COMPRAS", "PICKCOMPRA", "CANT_ASIG", 7, 0)
    Call CREACAMPO("COMPRAS", "PICKCOMPRA", "CANT_ENTREG", 7, 0)
    Call CREACAMPO("COMPRAS", "PICKCOMPRA", "FECHMOV", 8, 0)
    Call CREACAMPO("COMPRAS", "PICKCOMPRA", "REMITO", 10, 16)
    Call CREACAMPO("COMPRAS", "PICKCOMPRA", "UBICACION", 10, 64)
    Call CREACAMPO("COMPRAS", "PICKCOMPRA", "STOCK", 7, 0)
    Call CREACAMPO("COMPRAS", "PICKCOMPRA", "COMPRAS", 7, 0)
    Call CREACAMPO("COMPRAS", "PICKCOMPRA", "RECEPCION", 7, 0)
    Call CREACAMPO("COMPRAS", "PICKCOMPRA", "FALTANTE", 7, 0)
    Call CREACAMPO("COMPRAS", "PICKCOMPRA", "LOTE", 10, 18)
    Call CREACAMPO("COMPRAS", "PICKCOMPRA", "CANTIDAD", 7, 0)
    '
End Sub
'\\\OT-05-0676-WAR-14/10/05///
Sub PRIREGLACOM() 'IMPRIME REGLAMENTO DE COMPRA
  '
  With FORSELREG
    '
    For JJ% = 0 To .LIMATSEL.ListCount - 1
      If .LIMATSEL.Selected(JJ%) = True Then
        NREGLA& = Mid$(.LIMATSEL.List(JJ%), 1, 5)
        If NREGLA& > 0 Then
          'IMPRESIÓN DEL REGLAMENTO
          FORIPRE$ = Control$("", "REGLACOM")
          If TRIM$(FORIPRE$) <> "" Then
           '
           Screen.MousePointer = 11
           FECHDOC$ = Text9
           NUMOC$ = TRIM$(FORSELREG.Label3) '\\**//
           While Len(NUMOC$) < 6
             NUMOC$ = "0" + NUMOC$
           Wend
           '
           TIPODOC$ = "Reglamento de Compra"
           '
           CODPARAM$(1) = "NUME-COM"
           CAPARDOC% = 1
           '
           DOCPARAM$(1) = NUMOC$
           '
           CODTABU1$(1) = "REF-MAT1"
           CACOLTAB1% = 1
           '
           CAITAB1% = 0
           DESREGLA$ = DESCRIREGLA$(NREGLA&) 'DESCRIPCION DEL REGLAMENTO
           Call FRATEXTO(DESREGLA$, 92)
           For U& = 1 To CARETEX%
             Z$ = "." + RETEX$(U&)
             CAITAB1% = CAITAB1% + 1
             TETABU1$(1, CAITAB1%) = Z$
           Next U&
           '
           Call PRINTDOC("REGLACOM")
           '
           ' LIMPIA EL REGLAMENTO DE COMPRAS
           For KK1% = 1 To CACOLTAB1%
             For KK2% = 1 To CAITAB1%
               TETABU1$(KK1%, KK2%) = ""
             Next KK2%
           Next KK1%
           '
          End If
        End If
      End If
    Next JJ%
    Screen.MousePointer = 1
  End With
  '
  Unload FORSELREG
  '
End Sub
  
Sub SELEC_REGLACOM(OKX%) 'selecciona REGLAMENTO DE COMPRA
   '
   OKX% = 0
   RAI$ = "" 'REGLAMENTO COMPRA ANEXO
   FOREGLA$ = TRIM$(Control$("", "REGLACOM"))
   If FOREGLA$ <> "" Then
     If ECOARCH$("INDIFRM", FOREGLA$) <> "" Then
       If ULTREG&("INDIRCOM") > 0 Then
         GoTo 10
       End If
     End If
   End If
   OKX% = 1 'SI BASE DE DATOS NO TIENE  REGLAMENTOS O NO USA
   Exit Sub
   '
10 With FORSELREG
    .LIMATSEL.Clear
    FIN& = ULTREG&("INDIRCOM")
    For J& = 1 To FIN&
      X$ = REGLEIDO$("INDIRCOM", J&)
      NR& = CVI(Mid$(X$, 1, 2))
      DRC$ = Mid$(X$, 3, 46)
      RC$ = FORMATNUM(NR&, "F5.0") & Chr(9) & DRC$
      .LIMATSEL.AddItem RC$
    Next J&
    .Show 1
    If .Label1 = "OK" Then
      OKX% = 1
    End If
    Call FRESHALL
    DoEvents
    '
    RAI$ = "" 'REGLAMENTOS ADJUNTOS IMPRESOS
    For JJ% = 0 To .LIMATSEL.ListCount - 1
      If .LIMATSEL.Selected(JJ%) = True Then
        NREGLA& = XVALO(Mid$(.LIMATSEL.List(JJ%), 1, 5))
        If NREGLA& > 0 Then
          RAI$ = RAI$ + TRIM$(FORMATNUM(NREGLA&, "F6.0")) + "-"
        End If
      End If
    Next JJ%
    Screen.MousePointer = 1
    '
    If Len(RAI$) > 0 Then
      RAI$ = Mid$(RAI$, 1, Len(RAI$) - 1)
    End If
    '
  End With
  '
End Sub
Private Function DESCRIREGLA$(NREGLA&)
   '
   CIIINSTRU% = XVALO(NREGLA&)
   If CIIINSTRU% < 1 Then Exit Function
   '
   REGI& = CIIINSTRU%
   URN& = ULTREG&("REGLACOM")
   '
   DESCRIREGLA$ = "": TTYZ$ = ""
   If CIIINSTRU% > 0 Then
     If CIIINSTRU% <= URN& Then
       Screen.MousePointer = 11
       RNC$ = RECFILT$("REGLACOM", 1, MKI$(CIIINSTRU%))
       '
       For U& = 1 To Len(RNC$) Step 4
         RENOTA& = CVS(Mid$(RNC$, U&, 4))
         If RENOTA& > 0 Then
           If RENOTA& <= URN& Then
             TTYY$ = RTrim$(Mid$(REGLEIDO$("REGLACOM", RENOTA&), 3))
             If TTYZ$ <> "" Then
               TTYZ$ = TTYZ$ + Chr$(13) + Chr$(10)
             End If
             TTYZ$ = TTYZ$ + RTrim$(TTYY$)
           End If
         End If
       Next U&
       Screen.MousePointer = 1
     End If
   End If
   '
   LOCADE& = Len(TTYZ$)
   For U& = LOCADE& To 1 Step -1
     If Asc(Mid$(TTYZ$, U&, 1)) > 32 Then GoTo 30
   Next U&
   U& = 0
   '
30 DESCRIREGLA$ = Left$(TTYZ$, U&)
  '
End Function
'\\\OT-05-0676-WAR-FIN///
Private Sub UcoMat_Click()
Call OPLISCO07.Command20_Click
End Sub

Private Sub UcoPro_Click()
Call OPLISCO07.Command12_Click
End Sub

Private Sub Ver_Ocom_Click()
    Call OPLISCO07.Command22_Click
End Sub
'///////******RUTINA VIEJA DE  CONSIGNACION
'Sub ENTRECONSIACC(NPROVE%)
'        '
'        LLAMACOMPRA% = 0
'        If NPROVE% > 0 Then LLAMACOMPRA% = 1
'        '
'        FORIPRE$ = TRIM$(CONTROL$("", "FOREMICO"))
'        If FORIPRE$ <> "" Then
'            LU$ = LUDAT$
'            If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
'                RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
'                Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Remito en Consignación:\  \'" + RMCC$ + "'.")
'                GoTo 9999
'            End If
'        End If
'        '
'        DEPOCON% = XVALO(CONTROL$("", "DEPOCONS"))
'        DEPORIG% = DEPOCON%
'        DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DEPORIG%))
'        '
'        'Call BLIMESS("Revisando Archivo de Movimientos de Stock")
'1100    Screen.MousePointer = 11
'        CMO$ = TRIM$(CONTROL$("", "CODRECON"))
'        If CMO$ = "" Then CMO$ = "RC"
'        FIN& = ULTREG&("MOVISTO")
'        For U& = FIN& To 1 Step -1
'            Call TRACE(20, FIN& - U& + 1, FIN&)
'            XX$ = REGLEIDO$("MOVISTO", U&)
'            If UCase$(Mid$(XX$, 21, 2)) = CMO$ Then
'                UNRO& = XVALO(Mid$(XX$, 26, 7))
'                GoTo 1002
'            End If
'        Next U&
'        UNRO& = 0
'        '
'1002    REMI$ = CMO$ + "-" + TRIM$(Str$(UNRO& + 1))
'        Call FRESHALL '¿?
'        '
'1105    Call CIERRARCH("*.*")
'        '
'        Screen.MousePointer = 1
'        If NPROVE% > 0 Then
'             EBOL$ = MKI$(NPROVE%)
'          Else
'             EBOL$ = OBJPLA$("SELPROVE", "")
'             If EBOL$ = "" Then
'                GoTo 9999
'             End If
'        End If
'        '
'        NC% = CVI(EBOL$) 'Obtengo el numero de prov
'        If NC% < 1 Or ECOARCH$("PROVED1", MKI$(NC%)) = "" Then
'            Call MENSERR(24, "Proveedor Inexistente")
'            NPROVE% = 0
'            GoTo 1105
'        End If
'        '
'        NC1% = (-1) * NC%
'        '
'        Y$ = PEDY$: If Y$ = "" Then Y$ = Space$(128)
'        Call REPLA(Y$, MKS$(NROBOL&), 1, 4)
'        Call REPLA(Y$, EBOL$, 7, 2)
'        Call REPLA(Y$, MKI$(HOY(HO$)), 5, 2)
'        Call REPLA(Y$, RASOCLI$(NC1%), 9, 32)
'        Call REPLA(Y$, DOMICLI$(NC1%), 41, 32)
'        CLOCA$ = TRIM$(CPOSCLI$(NC1%)): If CLOCA$ <> "" Then CLOCA$ = CLOCA$ + " "
'        Call REPLA(Y$, CLOCA$ + LOCACLI$(NC1%), 73, 32)
'1110    Call REPLA(Y$, REMI$, 105, 16)
'        Call REPLA(Y$, Chr$(DEPORIG%), 121, 1)
'        '
'1111    PEDY$ = OBJPLA$("REMICON-ENCB", Y$)
'        If PEDY$ = "" Then GoTo 9998
'        FF% = CVI(Mid$(PEDY$, 5, 2))
'        REMI$ = TRIM$(Mid$(PEDY$, 105, 16))
'        If TRIM$(REMI$) = "" Then
'            Y$ = PEDY$
'            Call MENSERR(24, "Falta Numero de Remito")
'            GoTo 1111
'        End If
'        '
'        DEPORIG% = Asc(Mid$(PEDY$, 121, 1))
'        If DEPOCON% > 0 Then
'          If DEPORIG% < 1 Or ECOARCH$("TADEPOSI", Chr$(DEPORIG%)) = "" Then
'            Call MENSERR(24, "Depósito de Salida no Válido")
'            DEPORIG% = XVALO(CONTROL$("", "DEPOCONS"))
'            GoTo 1110
'          End If
'        End If
'        DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DEPORIG%))
'        '
'1112    Call CIERRARCH("*.*")
'        Call BLANARCH("!BOLREDET")
'        '
'        Screen.MousePointer = 11
'        HAYOC% = 0
'        For I& = 1 To ULTREG&("MACONSIG") 'Aca empieza con maconsig
'            X$ = REGLEIDO$("MACONSIG", I&)
'            If CVI(Mid$(X$, 3, 2)) = NC% Then
'                CAPENE# = CVD(Mid$(X$, 55, 8)) - CVD(Mid$(X$, 63, 8)) ' Cant asignada - Cant entregada
'                If CAPENE# > 0 Then
'                    Nuoc& = XVALO(Mid$(X$, 24, 6))
'                    CI% = CVI(Mid$(X$, 33, 2))
'                    STODISPO# = CDbl(STOCKACT(CI%))
'                    If STODISPO# < 0 Then STODISPO# = 0
'                    If STODISPO# < CAPENE# Then CAPENE# = STODISPO# ' Chequea el stock disponible
'                    '
'                    Y$ = REGBLAN$("!BOLREDET") 'Nueva boleta recep. prov
'                    Call REPLA(Y$, MKI$(CI%), 83, 2) 'cod mat
'                    Call REPLA(Y$, UNIMED$(CI%), 85, 4) 'unidad de medida
'                    Call REPLA(Y$, MKD$(CAPENE#), 89, 8) 'Cantidad nominal
'                    Call REPLA(Y$, MKS$(Nuoc&), 151, 4) 'Nº O/C
'                    Call REPLA(Y$, MKS$(I&), 189, 4) ' Relleno
'                    HAYOC% = 1
'                    Call REGAPP("!BOLREDET", Y$)
'                End If
'            End If
'        Next I&
'        '
'        If HAYOC% = 0 Then
'            Call COMUNI("No hay Consignaciones Pendientes\para este Proveedor.")
'        End If
'        '
'1113    VTB% = VENTABU%("ECONSIG/NC/TITUPAN=Consignación de Materiales - " + TRIM$(RASOCLI$(NC1%)))
'        If VTB% < 0 Or ULTREG&("!BOLREDET") < 1 Then
'            X$ = "": Y$ = "": YY$ = ""
'            EBOL$ = "": PEDY$ = "": DESCRI$ = ""
'            GoTo 1105
'        End If
'        '
'        For YX& = 1 To ULTREG&("!BOLREDET")
'          YY$ = REGLEIDO$("!BOLREDET", YX&)
'          CI% = CVI(Mid$(YY$, 83, 2)) 'Asigna el cod mat
'          CAN# = CVD(Mid$(YY$, 89, 8)) ' Asigna la cant nominal
'          If CI% > 0 Then
'            If CI% <= NUMAS% Then
'              If CAN# > 0.005 Then
'                GoTo 1114
'              End If
'            End If
'          End If
'        Next YX&
'        Call MENSERR(24, "Consignación Cancelada !!!\Faltan Cantidades o Códigos.")
'        GoTo 1113
'        '
'1114    If TRIM$(FORIPRE$) <> "" Then
'           DTP$ = OBJPLA$("TRANSPOR", DATRANSPO$)
'           If DTP$ = "" Then GoTo 1113         ' CANCELADO
'           DATRANSPO$ = DTP$
'        End If
'        '
'11141   CONTA% = ULTREG&("!BOLREDET")
'        '
'1115    If CONTA% > 0 Then
'          '
'          Call CIERRARCH("MOVISTO")
'          Call BLOQARCH("MOVISTO")
'          Call BLOQARCH("MACONSIG")
'          Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
'          '
'          AINDI% = 0
'          For YX& = 1 To ULTREG&("!BOLREDET")
'             YY$ = REGLEIDO$("!BOLREDET", YX&)
'             CI% = CVI(Mid$(YY$, 83, 2))  'Asigna cod mat de boleta recep. prov.
'             CAN# = CVD(Mid$(YY$, 89, 8)) 'Asigna cant nominal
'             LOTE$ = Mid$(YY$, 69, 12)    'Asigna identif. lote
'             If CI% > 0 Then
'               If CI% <= NUMAS% Then
'                 If CAN# > 0.005 Then
'                    Nuoc& = CVS(Mid$(YY$, 151, 4)) ' Asigna Nº O/C
'                    REGOC& = CVS(Mid$(YY$, 189, 4)) 'Asigna filler
'                    DOSEC$ = ""
'                    If REGOC& > 0 Then
'                       If REGOC& <= ULTREG&("MACONSIG") Then
'                          ZZ$ = REGLEIDO$("MACONSIG", REGOC&)
'                          If CVI(Mid$(ZZ$, 3, 2)) = NC% Then ' Si son = PROV
'                             If CVI(Mid$(ZZ$, 33, 2)) = CI% Then ' Si son = MAT
'                                DOSEC$ = Mid$(ZZ$, 21, 12)       ' asigna Nº O/C
'                                TOENTRE# = CVD(Mid$(ZZ$, 63, 8)) + CAN# 'cant entregada + cant nominal
'                                Call REPLA(ZZ$, MKD$(TOENTRE#), 63, 8)  'Escribe como: nueva cantidad entregada
'                                If FF% > CVI(Mid$(ZZ$, 71, 2)) Then
'                                    Call REPLA(ZZ$, MKI$(FF%), 71, 2)   'Escribe la fecha de entrega
'                                End If
'                                Call GRAREG("MACONSIG", ZZ$, REGOC&) ' Graba el nuevo registro
'                                GoTo 1120   ' Va al 1120
'                             End If
'                          End If
'                       End If
'                    End If
'                    '
'                    ZZ$ = REGBLAN$("MACONSIG") 'Hace un blanco en el archivo maconsig
'                    Call REPLA(ZZ$, MKI$(FF%), 1, 2) 'fecha de reserva
'                    Call REPLA(ZZ$, MKI$(NC%), 3, 2) ' Nº de prov.
'                       NOCOM$ = TRIM$(Str$(Nuoc&)) ' Calcula Nº O/C.
'                       While Len(NOCOM$) < 6: NOCOM$ = "0" + NOCOM$: Wend
'                       NOCOM$ = "OC-" + Left$(NOCOM$, 6) + "-00"
'                    Call REPLA(ZZ$, NOCOM$, 21, 12) ' Nº O/C
'                    Call REPLA(ZZ$, MKI$(CI%), 33, 2) ' Codigo de Mat
'                    Call REPLA(ZZ$, UNIMED$(CI%), 51, 4) 'Unidad de medida
'                    Call REPLA(ZZ$, MKD$(CAN#), 55, 8) 'Cant asignada
'                    Call REPLA(ZZ$, MKD$(CAN#), 63, 8) ' Cant entregada
'                    Call REPLA(ZZ$, MKI$(FF%), 71, 2)   'Fecha de entrega
'                    Call REGAPP("MACONSIG", ZZ$)
'                    '
'1120                CANTI = (-1) * CAN# ' Hace negativa la cantidad nominal
'                    'LLAMA A LA FUNCION MOVISTO
'                    'Call Movisto(FF%, CI%, LOTE$, CMO$, REMI$, DOSEC$, RASOCLI$(NC1%), 0, "$", NC1%, DEPORIG%, CANTI)
'                 End If
'               End If
'             End If
'             '
'          Next YX&
'          Call CIERRARCH("MOVISTO")
'          Call CIERRARCH("MACONSIG")
'          '
'          Call FRESHECHO("IREPRIPE")
'          Screen.MousePointer = 1
'          '
'          If TRIM$(FORIPRE$) <> "" Then
'            '
'            FEX = FF%
'            FECHDOC$ = FECHATEX$(FEX)
'            NUMEDOC$ = TRIM$(REMI$)
'            TIPODOC$ = "Remito de Traslado"
'            '
'            CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = TRIM$(Mid$(DTP$, 1, 32))
'            CODPARAM$(2) = "DOMI-TRA": DOCPARAM$(2) = TRIM$(Mid$(DTP$, 33, 32))
'            CODPARAM$(3) = "CUIT-TRA": DOCPARAM$(3) = TRIM$(Mid$(DTP$, 65, 16))
'            CODPARAM$(4) = "PATE-TRA": DOCPARAM$(4) = TRIM$(Mid$(DTP$, 81, 16))
'            CODPARAM$(5) = "COND-TRA": DOCPARAM$(5) = TRIM$(Mid$(DTP$, 97, 32))
'            CODPARAM$(6) = "DOCUCOND": DOCPARAM$(6) = TRIM$(Mid$(DTP$, 129, 24))
'            CODPARAM$(7) = "LOCA-TRA": DOCPARAM$(7) = TRIM$(Mid$(DTP$, 159, 32))
'            CODPARAM$(8) = "NUBULTOS": DOCPARAM$(8) = TRIM$(Mid$(DTP$, 153, 6))
'            CAPARDOC% = 8
'            '
'            CODTABU1$(1) = "COD-MAT"
'            CODTABU1$(2) = "DES-MAT"
'            CODTABU1$(3) = "CANTIDAD"
'            CODTABU1$(4) = "CAKILOS"
'            CODTABU1$(5) = "UNIMED"
'            CODTABU1$(6) = "N-OCOMPR"
'            CACOLTAB1% = 6
'            '
'            DESTIDOC% = NC1%
'            CAITAB1% = 0
'            '
'            For YX& = 1 To ULTREG&("!BOLREDET")
'               YY$ = REGLEIDO$("!BOLREDET", YX&)
'               CI% = CVI(Mid$(YY$, 83, 2))
'               CAN# = CVD(Mid$(YY$, 89, 8))
'               If CI% > 0 Then
'                  If CI% <= NUMAS% Then
'                     If CAN# > 0.005 Then
'                        UNID$ = TRIM$(Mid$(YY$, 85, 4))
'                        If UNID$ = "" Then UNID$ = UNIMED$(CI%)
'                        CANS$ = CSTRING$(MKD$(CAN#), 3, 9, 1, 2)
'                        KILO$ = CSTRING$(MKD$(CAN# * PESUNI(CI%)), 3, 12, 3, 2)
'                        NULOTE$ = TRIM$(Mid$(YY$, 69, 12))
'                        NUOCO$ = CSTRING$(Mid$(YY$, 151, 4), 3, 7, 0, 2)
'                        '
'                        CAITAB1% = CAITAB1% + 1
'                        TETABU1$(1, CAITAB1%) = CODEXT$(CI%)
'                        TETABU1$(2, CAITAB1%) = DESCRIT$(CI%)
'                        TETABU1$(3, CAITAB1%) = CANS$
'                        TETABU1$(4, CAITAB1%) = KILO$
'                        TETABU1$(5, CAITAB1%) = UNID$
'                        TETABU1$(6, CAITAB1%) = NUOCO$
'                        '
'                     End If
'                  End If
'               End If
'            Next YX&
'            '
'            Call PRINTDOC("FOREMICO")   ' Remito en Consignacion
'            '
'          End If
'          '
'1199    End If
'        '
'9998    Call CIERRARCH("*.*")
'        '
'        X$ = "": Y$ = "": YY$ = ""
'        EBOL$ = "": PEDY$ = "": DESCRI$ = ""
'        '
'        If LLAMACOMPRA% = 0 Then
'           GoTo 1100
'        End If
'        '
'9999    Call CIERRARCH("*.*")
'        '
'End Sub
''

Sub CARDETCOM1()
   '
   Call CIERRARCH("!OCOMDETA")
   List2.ListIndex = (-1)
   HOII% = HOY(HOS$)
   '
   For U& = 1 To ULTREG&("!OCOMDETA")
     ZZ$ = REGLEIDO$("!OCOMDETA", U&)
     If TRIM$(Mid$(ZZ$, 105, 12)) = "" Then Call REPLA(ZZ$, Text7, 105, 12)
     If TRIM$(Mid$(ZZ$, 15, 12)) = "" Then Call REPLA(ZZ$, Text18, 15, 12)
     Call GRAREG("!OCOMDETA", ZZ$, U&)
   Next U&
   '
10 ATRI$ = "/NC"
   ATRI$ = ATRI$ + "/TITUPAN=Detalle de Orden de Compra en " + TRIM$(Label11)
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(140 + Top + Frame1.Top + Picture3.Top))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture3.Left + 50))
   '
   Timer1.Enabled = True
   PRF$ = "CARDETC1"
   XXX% = VENTABU%(PRF$ + ATRI$)
   '
   If XXX% < 0 Or ULTREG&("!OCOMDETA") < 1 Then
      Call CIERRARCH("!OCOMDETA")
      Call BLANARCH("!OCOMDETA")
      Call BLANFORMU
      Total(0).TEXT = ""
      Text1.TEXT = ""
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   '
End Sub

Sub ACOSPROV()
    '
    If DERACCE%("ACOSPROV", "Actualizacion de Costos por Proveedore") >= 2 Then
      HOII% = HOY(HOS$)
3     VDEF$ = String$(6, 0) + MKI$(2) + MKI$(HOII%) + MKI$(HOII%)
5     OBX$ = OBJPLA$("AJUCOSPRO", VDEF$)
      If OBX$ = "" Then Exit Sub
      '
10    VDEF$ = OBX$
      NCX1% = CVI(Left$(OBX$, 2))
      If NCX1% < 1 Then GoTo 5
      If ECOARCH$("PROVED1", MKI$(NCX1%)) = "" Then
        Call MENSERR(24, "Numero de Proveedor Inexistente o no Válido")
        GoTo 5
      End If
      '
      PORCE = CVS(Mid$(OBX$, 3, 4))
      CADECI% = CVI(Mid$(OBX$, 7, 2))
      FCOR% = CVI(Mid$(OBX$, 9, 2))
      FVIG% = CVI(Mid$(OBX$, 11, 2))
      '
      Screen.MousePointer = 11
      FIN& = NUMAS%
      Call BLANARCH("ACTUCOS")
      i& = 0
      For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        'X$ = REGLEIDO$("MASTER", U& + 1)
        CIXI% = U&
        If PROVHABI%(CIXI%) = NCX1% Then
           If STATITEM%(CIXI%) > 0 Then    ' PARA SUPROMIR ITEMS DE BAJA
             CODE$ = CODEXT$(CIXI%)
             MON% = MONECOSTO%(CIXI%)
             PODES1 = 0
             COS1# = COSUNI(CIXI%)
             COS2# = COS1#
             FCOS% = FECOSTO%(CIXI%)
             If FCOS% <= FCOR% Then
                COS2# = ROUND#(COS1# * (1 + PORCE / 100), CADECI%)
                PODES1 = PORCE
             End If
             '
             Y$ = REGBLAN$("ACTUCOS")
             Call REPLA(Y$, MKI$(CIXI%), 1, 2)
             Call REPLA(Y$, MKI$(MON%), 3, 2)
             Call REPLA(Y$, MKD$(COS1#), 5, 8)
             Call REPLA(Y$, MKS$(PODES1), 13, 4)
             Call REPLA(Y$, MKD$(COS2#), 17, 8)
             Call REPLA(Y$, MKI$(FCOS%), 25, 2)
             i& = i& + 1
             Call GRAREG("ACTUCOS", Y$, i&)
           End If
           '
        End If
      Next U&
      Call SETULTREG("ACTUCOS", i&)
      If i& < 1 Then
         Call MENSERR(24, "Proveedor sin Códigos Pre-Asignados")
         GoTo 5
      End If
      '
      Call FILESORT("ACTUCOS", "", 1, 1)
      '
      OPX% = COMALTER%("Opciones de Trabajo:\\Actualizar\Listar")
      If OPX% < 1 Or OPX% > 2 Then GoTo 3
      If OPX% = 2 Then
         Call HEADERS("ACTUCOS", "")
         Call HEADERS("ACTUCOS", "Proveedor: " + rasocli$((-1) * NCX1%))
         Call FINAL("*LACTUCOS")
         GoTo 5
      End If
      '
      Call TRALOCAL("ACTUCOS", "")
      Screen.MousePointer = 1
      '
      JUN% = VENTABU%("ACOSPROV/TITUPAN=Costo de Reposición - " + rasocli$((-1) * NCX1%))
      If JUN% < 0 Then GoTo 5
      '
      Screen.MousePointer = 11
      FEVIGE$ = FETEXCOR1$(FVIG%)
      FIN& = ULTREG&("!ACTUCOS")
      For J& = 1 To FIN&
         Call TRACE(20, J&, FIN&)
         Z$ = REGLEIDO$("!ACTUCOS", J&)
         CIXI% = CVI(Mid$(Z$, 1, 2))
         COUNUE = CVD(Mid$(Z$, 17, 8))
         '
         '\\\OT-08-0692-OD-28/08/08///
         'SE AGREGO EL FORMATNUM AL COUNUE PARA QUE NO DE ERROR POR QUE _
         EL SQL CUANDO LOS NUMEROS SON CON COMA PARA LOS DECIMALES NO LOS SOPORTA
         SQLX$ = "UPDATE MASTER SET Cosuni = " & FORMATNUM$(COUNUE, "F12.4") & " ,  Fecosto = '" & FEVIGE$ & "' "
         SQLX$ = SQLX$ + " WHERE Codint = " & CIXI%
         On Error Resume Next
         FLEXCONN.Execute SQLX$
         On Error GoTo 0
         '\\\OT-08-0692-OD-FIN///
         '
      Next J&
      '
      'ACTUALIZA TABLA NOVEDADES
      For J& = 1 To FIN&
         Call TRACE(20, J&, FIN&)
         Z$ = REGLEIDO$("!ACTUCOS", J&)
         CIXI% = CVI(Mid$(Z$, 1, 2))
         COUVIE# = CVD(Mid$(Z$, 5, 8))
         COUNUEXX# = CVD(Mid$(Z$, 17, 8))
         Call ADD_NOVEDAD(CIXI%, NCX1%, COUVIE#, COUNUEXX#, "Actualización Masiva Por Porcentaje", MONECOSTO%(CIXI%), "COSTO")
         '
      Next J&

      Screen.MousePointer = 1
      '
      GoTo 3
    End If
End Sub

Sub CONAPLI(APLINVO$, TERMINA%)
   '
   CODCARPE$ = BUSCAVALOR(PARAPLI$, "CODCARPE =") ' codigo de carpeta de import / export
   '
   TERMINA% = 1
   Select Case APLINVO$
       '
       'OPCIONES VARIAS DE PRESUPUESTOS
       '
       Case "OPVARCOM"
           Call Command13_Click                 ' OPCIONES VARIAS - COMPRAS
       Case "OCOABI07"
           Call CIERRARCH("*.*")                ' ORDENES DE COMPRAS ABIERTAS
           ACONEC$ = "OCOABI07"
           Call CONECRUN(ACONEC$, "Ordenes de Compra Abiertas")
           Call FINAL("")
       Case "APRORCOMP"
           Call Command22_Click                 ' APRUEBA ORDEN DE COMPRA
       Case "MODORCOMP"
           Call OPCICOM07.Command13_Click       ' MODIFICAR ORDEN DE COMPRA
       Case "ITECUMCOMP"
           Call OPCICOM07.Command4_Click        ' ITEM CUMPLIDO ORDEN DE COMPRA
       Case "ANUORCOMP"
           Call OPCICOM07.Command14_Click       ' ANULACION DE ORDEN DE COMPRA
       Case "REVANUOCOM"
           Call OPCICOM07.Command10_Click       ' REVERTIR ANULACION DE ORDEN DE COMPRA
       Case "CONSMAT"
           Call OPCICOM07.Command2_Click        ' CONSIGANCION DE MATERIALES
       Case "COMATCONS"
           Call OPCICOM07.Command15_Click       ' CONSUMO DE MATERIALES CONSIGNADOS
       '
       'CONSULTAS Y LISTADOS DE PEDIDOS
       '
       Case "CONLISCOM"
           Call Command6_Click                  ' CONSULTAS Y LISTADOS
       Case "CONSULOC"
           Call Command6_Click
       Case "ORCOMNUM"
           OPLISCO07.Option2.Value = True
           Call OPLISCO07.Command3_Click        ' ORDEN DE COMPRA POR NUMERO
       Case "ORCOMPRO"
           OPLISCO07.Option2.Value = True
           Call OPLISCO07.Command4_Click        ' ORDEN DE COMPRA POR PROVEEDOR
       Case "ORCOMART"
           OPLISCO07.Option2.Value = True
           Call OPLISCO07.Command5_Click        ' ORDEN DE COMPRA POR ARTICULO
       Case "MATENCONS"
           OPLISCO07.Option2.Value = True
           Call OPLISCO07.Command9_Click        ' MATERIAL EN CONSIGANCION
       Case "VERIMPORCOM"
           Call OPLISCO07.Command22_Click       ' VER Y REIMPRIMIR ORDEN DE COMPRA
       Case "ULCOMART"
           Call OPLISCO07.Command20_Click       ' CONTROL ULTIMAS COMPRAS POR ARTICULO
       Case "ULCOMPROV"
           Call OPLISCO07.Command12_Click       ' CONTROL ULTIMAS COMPRAS POR PROVEEDOR
       Case "ESTCOMPRAS"
           Call OPCICOM07.Command20_Click       ' ESTADISTICAS DE COMPRAS
       '
       'CONFIGURACION DE PRESUPUESTOS
       '
       Case "CONGENORCOM"
           Call Command5_Click                  ' CONFIGURACION GENERAL DE PRESUPUESTOS
      
       Case Else
           TERMINA% = 0
   End Select
End Sub

