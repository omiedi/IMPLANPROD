VERSION 5.00
Begin VB.Form FOMODOC07 
   BackColor       =   &H00B0D0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ordenes de Compra - Proveedores Externos"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11820
   Icon            =   "FOMODOC07.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   64
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00B0D0E0&
      BorderStyle     =   0  'None
      Height          =   8520
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   12000
      Begin VB.Frame Frame2 
         BackColor       =   &H00B0D0E0&
         Caption         =   "DESTINO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2055
         Left            =   240
         TabIndex        =   109
         Top             =   5880
         Visible         =   0   'False
         Width           =   6045
         Begin VB.TextBox Text22 
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
            Height          =   675
            Left            =   1155
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   117
            Top             =   1200
            Width           =   9285
         End
         Begin VB.TextBox Eco24 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   2400
            TabIndex        =   116
            Top             =   795
            Width           =   4095
         End
         Begin VB.TextBox TEXT24 
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
            Left            =   1155
            TabIndex        =   115
            Text            =   " "
            Top             =   795
            Width           =   915
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
            TabIndex        =   114
            Top             =   795
            Width           =   180
         End
         Begin VB.Frame Frame16 
            BackColor       =   &H00B0D0E0&
            Height          =   510
            Left            =   1170
            TabIndex        =   111
            Top             =   200
            Width           =   9265
            Begin VB.OptionButton Option2 
               Appearance      =   0  'Flat
               BackColor       =   &H00B0D0E0&
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
               Left            =   6000
               TabIndex        =   113
               Top             =   200
               Value           =   -1  'True
               Width           =   3015
            End
            Begin VB.OptionButton Option1 
               Appearance      =   0  'Flat
               BackColor       =   &H00B0D0E0&
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
               Left            =   600
               TabIndex        =   112
               Top             =   200
               Width           =   3255
            End
         End
         Begin VB.TextBox TEXT25 
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
            Left            =   7580
            MaxLength       =   16
            TabIndex        =   110
            Text            =   " "
            Top             =   795
            Width           =   2850
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
            Index           =   8
            Left            =   6360
            TabIndex        =   141
            Top             =   880
            Width           =   1110
         End
         Begin VB.Label Label21 
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
            Height          =   450
            Index           =   10
            Left            =   360
            TabIndex        =   119
            Top             =   1150
            Width           =   735
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
            Left            =   300
            TabIndex        =   118
            Top             =   880
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00B0D0E0&
         Caption         =   "RESUMEN FINAL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2265
         Left            =   6410
         TabIndex        =   83
         Top             =   3840
         Width           =   4335
         Begin VB.TextBox TXTPESO 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Height          =   285
            Left            =   3165
            Locked          =   -1  'True
            TabIndex        =   107
            Top             =   1920
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.TextBox Text17 
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
            Left            =   1830
            MaxLength       =   14
            TabIndex        =   91
            Top             =   1260
            Width           =   645
         End
         Begin VB.TextBox Total 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   2685
            TabIndex        =   90
            Top             =   1260
            Width           =   1500
         End
         Begin VB.TextBox PORDESCU 
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
            Left            =   1725
            MaxLength       =   14
            TabIndex        =   89
            Top             =   630
            Width           =   750
         End
         Begin VB.TextBox Total 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   2685
            TabIndex        =   88
            TabStop         =   0   'False
            Text            =   " "
            Top             =   600
            Width           =   1500
         End
         Begin VB.TextBox Total 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   2685
            TabIndex        =   87
            Top             =   315
            Width           =   1500
         End
         Begin VB.TextBox LText12 
            Alignment       =   1  'Right Justify
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
            Left            =   120
            MaxLength       =   14
            TabIndex        =   86
            Top             =   1200
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox Total 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   2685
            TabIndex        =   85
            Top             =   945
            Width           =   1500
         End
         Begin VB.TextBox Total 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   2685
            TabIndex        =   84
            Top             =   1575
            Width           =   1500
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
            TabIndex        =   108
            ToolTipText     =   "Haga un Doble-Click para Emitir CON / SIN Precios"
            Top             =   2010
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Label Ltota 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "I.V.A."
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
            Left            =   1125
            TabIndex        =   106
            ToolTipText     =   "Doble Click para Editar Detalle"
            Top             =   945
            Width           =   600
         End
         Begin VB.Label TICAMBIS 
            Alignment       =   1  'Right Justify
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
            Height          =   270
            Left            =   960
            TabIndex        =   102
            Top             =   1950
            Width           =   795
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "T. Cambio"
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
            TabIndex        =   101
            Top             =   2010
            Width           =   810
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   230
            Left            =   2520
            TabIndex        =   100
            Top             =   1320
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
            Left            =   840
            TabIndex        =   99
            Top             =   1350
            Width           =   900
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2520
            TabIndex        =   98
            Top             =   720
            Width           =   225
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
            Height          =   225
            Index           =   1
            Left            =   720
            TabIndex        =   97
            Top             =   715
            Width           =   900
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
            Height          =   225
            Index           =   0
            Left            =   1635
            TabIndex        =   96
            Top             =   400
            Width           =   960
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
            Height          =   225
            Index           =   4
            Left            =   480
            TabIndex        =   95
            ToolTipText     =   "Haga un Doble-Click para Emitir CON / SIN Precios"
            Top             =   1660
            Width           =   2115
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   230
            Left            =   2520
            TabIndex        =   94
            Top             =   1000
            Width           =   225
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
            Height          =   330
            Left            =   120
            TabIndex        =   93
            ToolTipText     =   "T.C.: 1.0000 - Doble-Click para Cambiar Moneda"
            Top             =   1560
            Width           =   1305
         End
         Begin VB.Label Label118 
            Alignment       =   2  'Center
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
            Height          =   270
            Left            =   1725
            TabIndex        =   92
            Top             =   945
            Width           =   750
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00B0D0E0&
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
         ForeColor       =   &H00000000&
         Height          =   2265
         Left            =   160
         TabIndex        =   66
         Top             =   3840
         Width           =   6165
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
            Height          =   285
            Left            =   1155
            TabIndex        =   103
            Top             =   1200
            Width           =   4950
         End
         Begin VB.TextBox Text13 
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
            Height          =   435
            Left            =   1155
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   74
            Top             =   1800
            Width           =   4950
         End
         Begin VB.TextBox Eco21 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   1890
            TabIndex        =   73
            Text            =   " "
            Top             =   600
            Width           =   4215
         End
         Begin VB.TextBox TEXT21 
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
            Left            =   1155
            TabIndex        =   72
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
            Height          =   250
            Index           =   1
            Left            =   1575
            TabIndex        =   71
            Top             =   600
            Width           =   180
         End
         Begin VB.TextBox Eco21 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   1890
            TabIndex        =   70
            Top             =   300
            Width           =   4215
         End
         Begin VB.TextBox TEXT21 
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
            Left            =   1155
            TabIndex        =   69
            Text            =   " "
            Top             =   300
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
            Height          =   250
            Index           =   0
            Left            =   1575
            TabIndex        =   68
            Top             =   315
            Width           =   180
         End
         Begin VB.TextBox Text10 
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
            Left            =   1155
            TabIndex        =   67
            Top             =   900
            Width           =   4950
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Solicito"
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
            Left            =   -3045
            TabIndex        =   104
            Top             =   1280
            Width           =   4110
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
            Height          =   585
            Index           =   3
            Left            =   360
            TabIndex        =   82
            Top             =   1755
            Width           =   735
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
            Height          =   225
            Index           =   1
            Left            =   -3045
            TabIndex        =   81
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
            Height          =   225
            Index           =   0
            Left            =   -525
            TabIndex        =   80
            Top             =   375
            Width           =   1590
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "L. Entrega"
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
            Left            =   -3045
            TabIndex        =   79
            Top             =   970
            Width           =   4110
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Anexos"
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
            Index           =   4
            Left            =   -3045
            TabIndex        =   78
            Top             =   1575
            Width           =   4110
         End
         Begin VB.Label Label13 
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
            Left            =   1155
            TabIndex        =   77
            Top             =   1500
            Width           =   3765
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
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
            Left            =   5670
            TabIndex        =   76
            Top             =   1500
            Width           =   420
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
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
            Index           =   5
            Left            =   5070
            TabIndex        =   75
            Top             =   1560
            Width           =   510
         End
      End
      Begin VB.TextBox Text18 
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
         Left            =   6195
         TabIndex        =   58
         Top             =   525
         Width           =   1340
      End
      Begin VB.CommandButton Command15 
         Caption         =   "T.Lib"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   6880
         TabIndex        =   57
         Top             =   840
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   5880
         TabIndex        =   56
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1180
         Width           =   1650
      End
      Begin VB.TextBox Text7 
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
         Left            =   6195
         TabIndex        =   54
         Top             =   210
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
         Left            =   7350
         TabIndex        =   53
         Top             =   210
         Width           =   180
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   9000
         Left            =   10965
         ScaleHeight     =   8970
         ScaleWidth      =   1035
         TabIndex        =   42
         Top             =   0
         Width           =   1065
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
            Left            =   120
            Picture         =   "FOMODOC07.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "Ver por Pantalla y Re-Imprimir"
            Top             =   4155
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
            Picture         =   "FOMODOC07.frx":074C
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Consultas y Listados"
            Top             =   4845
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
            Picture         =   "FOMODOC07.frx":0A56
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Ayuda FLEXOFT en Línea"
            Top             =   790
            Width           =   650
         End
         Begin VB.CommandButton Command11 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Prove."
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
            Picture         =   "FOMODOC07.frx":0D60
            Style           =   1  'Graphical
            TabIndex        =   48
            ToolTipText     =   "Acceso a Base de Datos de Proveedores"
            Top             =   1650
            Width           =   650
         End
         Begin VB.CommandButton Command8 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Art."
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
            Picture         =   "FOMODOC07.frx":106A
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Acceso a Maestro de Artículos"
            Top             =   2340
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
            Picture         =   "FOMODOC07.frx":1374
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Configuración de Funcionamiento de Compras"
            Top             =   3255
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
            Picture         =   "FOMODOC07.frx":17B6
            Style           =   1  'Graphical
            TabIndex        =   45
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
            Picture         =   "FOMODOC07.frx":1900
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Aprobar Ordenes de Compra"
            Top             =   5535
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
            Height          =   1005
            Left            =   105
            Picture         =   "FOMODOC07.frx":1C0A
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Aprueba, Graba e Imprime Orden de Compra"
            Top             =   6405
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
         TabIndex        =   41
         ToolTipText     =   "Lista de Proveedores con Entregas Programadas"
         Top             =   870
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Frame Frame15 
         Caption         =   "Frame15"
         Height          =   1830
         Left            =   10450
         TabIndex        =   40
         Top             =   2010
         Width           =   15
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00B0D0E0&
         Caption         =   "PENDIENTES DE RECEPCION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1845
         Left            =   160
         TabIndex        =   32
         Top             =   6120
         Width           =   10580
         Begin VB.Frame Frame14 
            Caption         =   "Frame14"
            Height          =   1730
            Left            =   9345
            TabIndex        =   39
            Top             =   210
            Width           =   15
         End
         Begin VB.Frame Frame13 
            Caption         =   "Frame13"
            Height          =   1730
            Left            =   8715
            TabIndex        =   38
            Top             =   210
            Width           =   15
         End
         Begin VB.Frame Frame12 
            Caption         =   "Frame12"
            Height          =   1730
            Left            =   7665
            TabIndex        =   37
            Top             =   210
            Width           =   15
         End
         Begin VB.Frame Frame11 
            Caption         =   "Frame11"
            Height          =   1730
            Left            =   5985
            TabIndex        =   36
            Top             =   210
            Width           =   15
         End
         Begin VB.Frame Frame10 
            Caption         =   "Frame10"
            Height          =   1730
            Left            =   1890
            TabIndex        =   35
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
            Left            =   0
            MultiSelect     =   1  'Simple
            TabIndex        =   34
            Top             =   620
            Width           =   10565
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   435
            Left            =   0
            ScaleHeight     =   405
            ScaleWidth      =   10530
            TabIndex        =   33
            Top             =   195
            Width           =   10565
            Begin VB.Label Label38 
               BackStyle       =   0  'Transparent
               Caption         =   "Código"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   140
               Top             =   120
               Width           =   975
            End
            Begin VB.Label Label37 
               BackStyle       =   0  'Transparent
               Caption         =   "Descripción"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1965
               TabIndex        =   139
               Top             =   120
               Width           =   1215
            End
            Begin VB.Label Label36 
               BackStyle       =   0  'Transparent
               Caption         =   "Cantidad"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   6645
               TabIndex        =   138
               Top             =   120
               Width           =   975
            End
            Begin VB.Label Label35 
               BackStyle       =   0  'Transparent
               Caption         =   "F. Entrega"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   7785
               TabIndex        =   137
               Top             =   120
               Width           =   975
            End
            Begin VB.Label Label20 
               BackStyle       =   0  'Transparent
               Caption         =   "Atraso"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   8745
               TabIndex        =   136
               Top             =   120
               Width           =   975
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "O/Compra"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   9405
               TabIndex        =   135
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
         Height          =   1830
         Index           =   7
         Left            =   6250
         TabIndex        =   31
         Top             =   2010
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   1830
         Index           =   6
         Left            =   8665
         TabIndex        =   30
         Top             =   2010
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   1830
         Index           =   5
         Left            =   7465
         TabIndex        =   29
         Top             =   2010
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   1830
         Index           =   4
         Left            =   9400
         TabIndex        =   28
         Top             =   2010
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   1830
         Index           =   3
         Left            =   5725
         TabIndex        =   27
         Top             =   2010
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   1830
         Index           =   2
         Left            =   4675
         TabIndex        =   26
         Top             =   2010
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   1830
         Index           =   1
         Left            =   4100
         TabIndex        =   25
         Top             =   2010
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   1830
         Index           =   0
         Left            =   2050
         TabIndex        =   24
         Top             =   2010
         Width           =   15
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00B0D0E0&
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
         Height          =   645
         Left            =   7665
         TabIndex        =   20
         Top             =   840
         Width           =   3060
         Begin VB.TextBox Text16 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   900
            TabIndex        =   23
            TabStop         =   0   'False
            Text            =   " "
            Top             =   280
            Width           =   2010
         End
         Begin VB.TextBox Text15 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   210
            MaxLength       =   2
            TabIndex        =   22
            TabStop         =   0   'False
            Text            =   " "
            Top             =   280
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
            Height          =   240
            Left            =   630
            TabIndex        =   21
            Top             =   280
            Width           =   150
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00B0D0E0&
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
         Height          =   645
         Left            =   7680
         TabIndex        =   17
         Top             =   120
         Width           =   1380
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
            Left            =   1050
            TabIndex        =   19
            Top             =   300
            Width           =   195
         End
         Begin VB.TextBox Text14 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   105
            TabIndex        =   18
            TabStop         =   0   'False
            Text            =   " "
            Top             =   300
            Width           =   960
         End
      End
      Begin VB.TextBox Text6 
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
         Left            =   2310
         MaxLength       =   30
         TabIndex        =   16
         Text            =   " "
         Top             =   1180
         Width           =   3225
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
         TabIndex        =   15
         Top             =   1995
         Width           =   10565
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   -120
         Top             =   240
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   6950
         Left            =   11025
         TabIndex        =   14
         Top             =   105
         Width           =   850
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   160
         ScaleHeight     =   405
         ScaleWidth      =   10530
         TabIndex        =   13
         Top             =   1575
         Width           =   10565
         Begin VB.Label Label76 
            BackStyle       =   0  'Transparent
            Caption         =   "Código"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   129
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1995
            TabIndex        =   128
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "U/S"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4020
            TabIndex        =   127
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4485
            TabIndex        =   126
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "U/C"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5685
            TabIndex        =   125
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6525
            TabIndex        =   124
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "P.Unitario"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7605
            TabIndex        =   123
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "% Dto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8745
            TabIndex        =   122
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "F.Entrega"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9405
            TabIndex        =   121
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   9480
         TabIndex        =   7
         Top             =   160
         Width           =   810
      End
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   9615
         TabIndex        =   5
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
         TabIndex        =   6
         Top             =   525
         Width           =   195
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   2310
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   " "
         Top             =   820
         Width           =   3225
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   1470
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   " "
         Top             =   820
         Width           =   750
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   1470
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   " "
         Top             =   490
         Width           =   4065
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   1470
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   " "
         Top             =   160
         Width           =   4065
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
         TabIndex        =   0
         Top             =   870
         Width           =   195
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
         TabIndex        =   120
         Top             =   50
         Width           =   1500
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   7
         Left            =   -1400
         TabIndex        =   105
         Top             =   1280
         Width           =   3615
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
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
         Left            =   5040
         TabIndex        =   63
         Top             =   1180
         Visible         =   0   'False
         Width           =   495
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
         TabIndex        =   62
         Top             =   1280
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   10290
         X2              =   10450
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Label Label16 
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
         Left            =   10455
         TabIndex        =   61
         ToolTipText     =   "Número de Revisión"
         Top             =   160
         Width           =   315
      End
      Begin VB.Label LText1 
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
         Left            =   165
         TabIndex        =   60
         Top             =   870
         Width           =   750
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
         Left            =   5600
         TabIndex        =   59
         ToolTipText     =   "Dobe-Click para Asignar Sector por Item"
         Top             =   610
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
         Left            =   5595
         TabIndex        =   55
         ToolTipText     =   "Dobe-Click para Asignar Destino por Item"
         Top             =   300
         Width           =   615
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
         Left            =   165
         TabIndex        =   52
         Top             =   1185
         Width           =   1170
      End
      Begin VB.Line Line5 
         X1              =   180
         X2              =   10720
         Y1              =   3850
         Y2              =   3850
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "C.U.I.T."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   5880
         TabIndex        =   12
         Top             =   945
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nro."
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
         Left            =   9050
         TabIndex        =   11
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
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
         Height          =   240
         Left            =   9050
         TabIndex        =   10
         Top             =   600
         Width           =   510
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
         Height          =   230
         Left            =   160
         TabIndex        =   9
         Top             =   630
         Width           =   960
      End
   End
   Begin VB.PictureBox SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   1425
      TabIndex        =   65
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox Skin1 
      Height          =   480
      Left            =   120
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   142
      Top             =   1560
      Width           =   1200
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción"
      Height          =   255
      Left            =   1840
      TabIndex        =   134
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      Height          =   255
      Left            =   6520
      TabIndex        =   133
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "F. Entrega"
      Height          =   255
      Left            =   7660
      TabIndex        =   132
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "Atraso"
      Height          =   255
      Left            =   8620
      TabIndex        =   131
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "O/Compra"
      Height          =   255
      Left            =   9280
      TabIndex        =   130
      Top             =   0
      Visible         =   0   'False
      Width           =   975
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
      Begin VB.Menu Nue_Orco 
         Caption         =   "Nueva Orden de Compra"
         Begin VB.Menu Directa 
            Caption         =   "O/C Directa"
         End
         Begin VB.Menu Por_MRP 
            Caption         =   "O/C por MRP"
         End
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
Attribute VB_Name = "FOMODOC07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHIPRE$
Dim MUECOPEN%
Dim EMICONSI%
Dim MODOULPRE%
Dim RAI$
Dim ACUMUMENSAX$ 'MENSAJE DE LOS ITEMS QUE SE MOODIFICARON PARA ENVIAR DE SUCURSAL A ARCON




Function MENSAJEDECAMBIO$(CI1%, CANTU, NRO&, ORDIT%, FEENT%)
    MENSAXE$ = ""
    SQLX$ = "SELECT * FROM OCOMDETA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE Cod_Inte =" & CI1%
    SQLX$ = SQLX$ + " AND NuOrd_Item = " & ORDIT%
    Set RS = READSET(SQLX$)
    With RS
    If .EOF Then
      MENSAXE$ = "Art.Agregado: " & CODEXT$(CI1) & " Cant.: " & TRIM$(FORMATNUM$(CANTU, "F9.1")) + " - "
      MENSAXE$ = MENSAXE$ + "Sol.Entrega: " & FECHATEX$(FEENT%)
    Else
      CantOrig = ROUND(XVALO(!Cant_Ocom), 1)
      FEORIG% = CVINT(!Fentre_Sol)
      If CantOrig <> ROUND(XVALO(CANTU), 1) Then MENSAXE$ = MENSAXE$ + "Modifica Cantidad: " & CODEXT$(CI1%) & " Nueva Cant.: " & TRIM$(FORMATNUM$(CANTU, "F9.1")) + " - "
      If FEORIG% <> FEENT% Then MENSAXE$ = MENSAXE$ + "Modifica Fecha Entrega: " & CODEXT$(CI1%) & " Nueva Fecha: " & FECHATEX$(FEENT%) + " - "
    End If
    End With
    MENSAJEDECAMBIO$ = MENSAXE$
    
End Function

Private Sub Anul_Ocom_Click()
   Call SETULTREG("!OCOMDETA", 0)
   Call BLANFORMU
   Call OPCICOM07.Command14_Click
End Sub

Private Sub Command1_Click()
   Call MENSERR(24, "No Válido Cambiar Proveedor")
End Sub

Private Sub Command10_Click()
   Call MENSERR(24, "No Válido Cambiar Proveedor")
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call CIERRARCH("*.*")
   ACONEC$ = "ARCHIG07/AMAPRO"
   Call CONECRUN(ACONEC$, "Padron Maestro de Proveedores")
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

Private Sub Command14_Click()
    Command14.Enabled = False
    '
    ACUMUMENSAX$ = ""
    ' esto es para prevenir el fallo de GUIDI DEL 27/08/2007 de que al llamar para
    ' modificar y luego aprobar sin cambios cambiaba la descripcion a "Material No
    ' Codificado"
    If COLTELIB% = 0 Then
       Call MENSERR(24, "Debe Editar Antes de Guardar.")
       Call Label22_DblClick
       Command14.Enabled = True
       Exit Sub
    End If
    
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "MINEPARTS") > 0 Then
       OPX% = COMALTER%("Atención !! Recuerde que los Cambios Realizados en Orden de Compra\Impactan de Manera Independiente\En Caso de Querer Realizar el Mismo Cambio\Deberá Modificar la Misma en AHP.\\Cancelar\Continuar")
       If OPX% < 2 Then
         Command14.Enabled = True
         Exit Sub
       End If
    End If
    
    '
    If InStr(EPAC$, "AHP") > 0 Then
      If XVALO(TXTPESO) > 50 Then
       OPX% = COMALTER%("Atención Peso Total Supera los 50 Kg.\\Cancelar\Continuar")
       If OPX% <= 1 Then Command14.Enabled = True: Exit Sub
      End If
      If InStr(EPAC$, "MINEPARTS") < 1 And InStr(EPAC$, "AHP") > 0 Then
        NprovMineparts = XVALO(Control("CFGMINEP", "PRFEMINE"))
        If NprovMineparts > 0 And XVALO(LText1) = NprovMineparts Then
             DESTICOMPRAMINEPAR% = 1
        End If
      End If
    End If
    '
    'validaciones de sucursales arcon
    If Frame2.Visible = True Then
       If Option2.Value = True And XVALO(TEXT24) < 1 Then
          Call COMUNI("Imposible Continuar !!! \ Se ha Seleccionado 'Solicitud por Pedido de Cliente'\Debe Seleccionar Un Cliente o Grabar por 'Pedido de Reposición de Stocks'")
          Exit Sub
       End If
       If Option2.Value = True And TRIM$(TEXT25) = "" Then
          Call COMUNI("Falta Ingresar Orden de Compra del Cliente")
          Exit Sub
       End If
       '
       If Option2.Value = True And OCOMEXISTENEUQUEN%(XVALO(TEXT24), TRIM$(TEXT25), TRIM$(Text11)) > 0 Then
          Call COMUNI("Atención Orden de Compra del Cliente Informado es Existente")
          Exit Sub
       End If
    End If
    
    'CONSIDERAR VALIDACIONES DE STOCK MINIMO O MAXIMO Y PEDIDOS 'ARCON'
    Resultado$ = ""
    ValidaStockYPedidos% = 0: ValidaStockYPedidosStockMin% = 0:
    If Control("OCOMPRA", "VALISTPE") = "SI" Then
      ValidaStockYPedidos% = 1
      If Control("OCOMPRA", "CONSSTMI") = "SI" Then
        ValidaStockYPedidosStockMin% = 1
      End If
      If Frame2.Visible = True And Option2.Value = False Then 'SI ES POR ORDEN DE COMPARA DEL PEDIDO NO VALIDA
        Call FORCOMP07.ValidarCompraArcon%(ValidaStockYPedidos%, ValidaStockYPedidosStockMin%, Resultado$, "MODIFICACION")
        If Resultado$ = "CANCELA" Then Command14.Enabled = True: Exit Sub
      End If
    End If
    
10  FORSELREG.Label3 = Text11 '\\**//
    Call SELEC_REGLACOM(OKX%)
    If OKX% < 1 Then
      Command14.Enabled = True
      Exit Sub
    End If
    '
    PORPRIN% = 1: PORMAIL% = 0
    NCI% = XVALO(LText1)
    NCIX% = NCI% * (-1)
    EMAI$ = TRIM$(EMAILCLI$(NCIX%))
    If EMAI$ = "" Or InStr(EMAI$, "@") < 1 Then
       GoTo 20 'NO TIENE DIRECCION O DIRECCION NO VALIDA
    End If
    '
12  SALIOCOM.Show 1
    Call FRESHALL
    If SALIOCOM.Check1.Value < 1 Then
       If SALIOCOM.Check2.Value < 1 Then
          Call MENSERR(24, "Debe Marcar al Menos\una Dirección de Salida.")
          GoTo 12
       End If
    End If
    
    'Si esta activo enviar adjuntos (SE AGREGA TAMBIEN EN LA MODIFICACION )
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
             Command14.Enabled = True
             Exit Sub
          End If
        End If
      Next U&
    End If
    '
    Call BORRACONSIG(XVALO(Text11))   'BORRA LOS MATERIALES CONSIGNADOS SEGUN LA O/C
    '
'    Call GRACOMPRA(OKX%)
    Call GRACOMPRA(OKX%, Resultado$, NrodeOrdenDeCompraGenerado&)
    If OKX% > 0 Then
       '
       ' ACTUALIZA LISTA INDICE DE ORDENES DE COMPRA
       CONDI$ = "Nume_Ocom > 0 ORDER BY Nume_Ocom"
       Call CARGALISEL("INDIOCOM", "OCOMENCA", "Nume_Ocom/F7;Refe_OcoM/A44", CONDI$, "DISTINCT")
       Call FRESHECHO("INDIOCOM")
       '
       If PORPRIN% > 0 Then      ' EMITE POR IMPRESORA
          '\\\OT-08-01047-FL-21/11/08///Si la aprobacion es "Posterior" y la impresión "Al Aprobar" realiza unicamente la impresion fantasma
          If Control("OCOMPRA", "MODOAPRO") = "POSTERIOR" Then
             If Control("OCOMPRA", "IMPRIME") = "ALAPROBA" Then
                CANCELPRINT% = 1
             End If
          End If
          Call PRICOMPRA
          If RAI$ <> "" Then Call FORCOMP07.PRIREGLACOM         'IMPRIME REGLAMENTO DE COMPRA
       End If
       '
       If PORMAIL% > 0 Then      ' ENVIA POR MAIL
          'PERMITE ELEGIR MAIL DE ENVIA A PARTIR DE LA LISTA DE MAILS DE CONTACTOS
          Call CONTACTO_MAIL(Label8, NCIX%) 'AGREGA A SENDBYMAIL LOS MAIL DE LOS CONTACTOS SELECCIONADOS
          If SENDBYMAIL$ = "" Then
            SENDBYMAIL$ = TRIM$(EMAILCLI$(NCIX%))
          End If
          '
          'SENDBYMAIL$ = EMAI$
          For Each PRX1 In Printers
       If InStr(UCase$(PRX1.DeviceName), "PDFWRITER") > 0 Or InStr(UCase$(PRX1.DeviceName), "BIOPDF") > 0 Or InStr(UCase$(PRX1.DeviceName), "BULLZIP") > 0 Then GoTo 25
          Next
          Call MENSERR(24, "Imposible Enviar por Mail.\   \Falta Impresora Virtual PDF.")
          SENDBYMAIL = ""
          GoTo 30
          '
25        Call PRICOMPRA
          If RAI$ <> "" Then Call FORCOMP07.PRIREGLACOM         'IMPRIME REGLAMENTO DE COMPRA
       End If
       '
'30     If Left$(CONTROL$("OCOMPRA", "MODOAPRO"), 6) <> "DIRECT" Then ' VALIDACION SI LA APROBACION ES DIRECTA O POSTERIOR
'          Call COMUNI("O/compra Registrada como\PENDIENTE DE APROBACION")
'       End If
       '
30     If Left$(Control$("OCOMPRA", "MODOAPRO"), 6) <> "DIRECT" Or Resultado$ = "GRABA_PENDIENTE_APROBACION" Then ' VALIDACION SI LA APROBACION ES DIRECTA O POSTERIOR
          Call COMUNI("O/compra Registrada como\PENDIENTE DE APROBACION")
       ElseIf Frame2.Visible = True And XVALO(TEXT24) > 0 Then 'CONTEPLAST COMPRAS INTERNAS
          NC_SOLICITANTE = NClteSegunNProvCompraInterna&(XVALO(LText1))
          If ACUMUMENSAX$ <> "" Then 'IF ACUMENSAX$ ESTA VACIO ES POR QUE NO CAMBIO NADA
                Call FORCOMP07.GENERAR_TRANSFERENCIA_OCOMPRA(1, TRIM$(Text13) & ";" & ACUMUMENSAX$)
                Call FORCOMP07.GENERAR_TRANSFERENCIA_OCOMPRA_INTERNAS_CONTEPLAST(XVALO(LText1), NC_SOLICITANTE, 1, TRIM$(Text22) & ";" & ACUMUMENSAX$)
             Else
                OPX% = COMALTER%("Opciones de Trabajo:\Pedido de Reposición Modificado - Desea Forzar el Envío Hacia el Proveedor\\No Enviar Modificación\Si Enviar Modificación")
                If OPX% > 1 Then
                  Call FORCOMP07.GENERAR_TRANSFERENCIA_OCOMPRA_INTERNAS_CONTEPLAST(XVALO(LText1), NC_SOLICITANTE, 1, TRIM$(Text22) & ";" & ACUMUMENSAX$)
                End If
                
             End If
       Else
          If Control("OCOMPRA", "ENVARSUC") = "SI" Then 'CONFIGURACION SETUP DE COMPRA DE ENVIO DE ARCHIVO DE O.C. A CENTRAL
             If ACUMUMENSAX$ <> "" Then 'IF ACUMENSAX$ ESTA VACIO ES POR QUE NO CAMBIO NADA
                Call FORCOMP07.GENERAR_TRANSFERENCIA_OCOMPRA(1, TRIM$(Text22) & ";" & ACUMUMENSAX$)
             Else
                OPX% = COMALTER%("Opciones de Trabajo:\Pedido de Reposición Modificado - Desea Forzar el Envío Hacia Avellaneda\\No Enviar Modificación\Si Enviar Modificación")
                If OPX% > 1 Then
                  Call FORCOMP07.GENERAR_TRANSFERENCIA_OCOMPRA(1, TRIM$(Text22) & ";" & ACUMUMENSAX$)
                End If
             End If
          End If
       End If
       '
' ESTO SE SACÓ DE AQUI PORQUE DEBE IR INMEDIATAMENTE
' DESPUES DE CALL GRACOMPRA - EPB 01/06/2010
'       'ACTAULIZA INDICE DE ORDENES DE COMPRA
'       CONDI$ = "Nume_Ocom > 0 ORDER BY Nume_Ocom"
'       Call CARGALISEL("INDIOCOM", "OCOMENCA", "Nume_Ocom/F7;Refe_OcoM/A44", CONDI$, "DISTINCT")

       Call BLANFORMU
       Call SETULTREG("!OCOMDETA", 0)
       Call CIERRARCH("*.*")
       Command14.Enabled = True
        'AGREGADO PARA ENVIAR POR MAIL
       If SENDBYMAIL$ <> "" Then
            Call RECONECTA
            Exit Sub
       End If
       'HASTA ACA AGREGADO
    End If
    '
    Unload Me
    '
End Sub

Private Sub Command15_Click()
   Command15.Enabled = False
   FORCOLIB07.Show 1
   Command15.Enabled = True
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
    Unload Me
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

Private Sub Command3_Click()
    Command3.Enabled = False
    Call HELPONLINE("PRESUP01")
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
     'agrego todas las O/C de ocomenca
     Call OPLISCO07.Command22_Click
     GoTo 99
     
     
      Call APERBASDAT("COMPRAS")
      SQLX$ = " SELECT Nume_Ocom, Refe_Ocom FROM OCOMENCA"
      SQLX$ = SQLX$ + " ORDER BY Nume_Ocom  DESC "
      'Set VEROCTEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
      Dim VEROCTEMP As ADODB.Recordset
      Set VEROCTEMP = New ADODB.Recordset
      VEROCTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

      With VEROCTEMP
        On Error Resume Next ' si hay algun error msgbox err
         .MoveFirst
         If Err Then
           Call MENSERR(24, "No hay Ordenes de Compra Emitidas")
           On Error GoTo 0
           GoTo 99
         End If
        On Error GoTo 0
         '
         JJ& = 0
         Do While (.EOF = False)
             XXY$ = REGBLAN$("INDIOCOM")
             Call REPLA(XXY$, MKS(!nume_ocom), 1, 4)
             Call REPLA(XXY$, !Refe_OcoM, 5, 44)
             JJ& = JJ& + 1
             Call GRAREG("!INDIOCOM", XXY$, JJ&)
             .MoveNext
         Loop
      End With
      VEROCTEMP.Close
      Set VEROCTEMP = Nothing
      
      Call SETULTREG("!INDIOCOM", JJ&)
      Call CIERRARCH("!INDIOCOM")
      Call FRESHECHO("!INDIOCOM")
      '
3     Call SELECHO("!INDIOCOM/Indice General de Ordenes de Compra (Pendientes-Cumplidas-Anuladas)")
      
      NPX& = XVALO(VALACT1$("!INDIOCOM"))
      If NPX& > 0 Then
        Call SHOWOCOM(NPX&)
        GoTo 3
      End If
      '
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

Private Sub Command8_Click()
    Command8.Enabled = False
    Call CIERRARCH("*.*")
    ACONEC$ = "ARCHIG07/AMASTO"
    Call CONECRUN(ACONEC$, "Maestro de Artículos de Stock")
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   Call MENSERR(24, "No Válido Cambiar Fecha de Emisión")
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
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(1, EMPREAC$, "AHP") > 0 Then
        Ltota(8).Visible = True
        TXTPESO.Visible = True
    End If

    If InStr(EPAC$, "METALURGICASIAM-ARCON") > 0 Or InStr(EPAC$, "NEUQUEN") > 0 Then
       Label10.Caption = "PR"
       Label10.ToolTipText = "Modificación de Pedido de Reposición"
       Me.BackColor = &HFBCA9D
       Caption = REPLACAR(Caption, "Ordenes de Compra", "Modificación de Pedido de Reposición")
       Dim CONTROLX As Control ' declaro el objeto control
       Dim OBJFRAME As Frame
       For Each CONTROLX In Me.Controls ' RECORRO LOS CONTROLES DEL FORM ACTIVO
        On Error Resume Next
        If TypeOf CONTROLX Is Frame Then
           Set OBJFRAME = CONTROLX
           OBJFRAME.BackColor = &HFBCA9D
        End If
        On Error GoTo 0
        Frame2.Left = 170
        
        Frame2.Top = 3840
        Frame2.Height = 2300
        Frame2.Width = 10600
        Frame2.Visible = True
        TEXT24.Locked = True
        TEXT25.Locked = True
        Command25.Enabled = False
        Option1.Enabled = False
        Option2.Enabled = False
        Eco24.Locked = True
        Label10.ForeColor = vbBlue
       Next
    End If
    If Control$("OCOMPRA", "REPLAORI") = "SI" Then
      Nuorig.Enabled = True
      Nuorig1.Enabled = True
    End If
    '
    
    Label118 = FORMATNUM$(ALIVA, "F5.2")
    EPAC$ = UCase$(TRIM$(EMPREAC$))
    If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
       If InStr(EPAC$, "DOSIVAC") > 0 Then
         Text6.Width = 2320
         Label17.Visible = True
         Label18.Visible = True
       End If
    End If

    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)

    MODOULPRE% = 1
    If Control("OCOMPRA", "PRECIDEF") = "COSTOREP" Then MODOULPRE% = 2
    '
    Call RECLAPAG
    '
End Sub
'

Sub CambiosEnPantallaCompraInterna(NC)
   'compra entre sucursales de la misma empresa
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    IDdeTablaDatasql& = ExisteNumeroProveedorEnConfiguraciónCompraInterna&(NC)
    If IDdeTablaDatasql& > 0 Then
       Label10.Caption = "PR"
       Label10.ToolTipText = "Modificación de Pedido de Reposición"
       Me.BackColor = &HB0D0E0
       Caption = REPLACAR(Caption, "Ordenes de Compra", "Modificación de Pedido de Reposición")
       Dim CONTROLX As Control ' declaro el objeto control
       Dim OBJFRAME As Frame
       For Each CONTROLX In Me.Controls ' RECORRO LOS CONTROLES DEL FORM ACTIVO
        On Error Resume Next
        If TypeOf CONTROLX Is Frame Then
           Set OBJFRAME = CONTROLX
           OBJFRAME.BackColor = &HB0D0E0
        End If
        On Error GoTo 0
        Frame2.Left = 170
        
        Frame2.Top = 3840
        Frame2.Height = 2300
        Frame2.Width = 10600
        Frame2.Visible = True
        TEXT24.Locked = True
        TEXT25.Locked = True
        Command25.Enabled = False
        Option1.Enabled = False
        Option2.Enabled = False
        Eco24.Locked = True
        Label10.ForeColor = vbBlue
       Next
       Option1.BackColor = &HB0D0E0
       Option2.BackColor = &HB0D0E0

    
    
    End If
    
    
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call LIBARCH("*.*")
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
   Call CARDETCOM
End Sub

Private Sub Label8_Click()
   '
   Call MENSERR(24, "No Válido Cambiar Proveedor")
   Exit Sub
   '
   Call SELECHO("INDIACRE")
   CPROX$ = TRIM$(VALACT1$("INDIACRE"))
   If CPROX$ <> "" Then
     NCLIE% = NUMEPRO%(CPROX$)
     If NCLIE% > 0 Then
       If REGICLI&((-1) * NCLIE%) > 0 Then
         Call BLANFORMU
         Text1.TEXT = TRIM$(Str$(NCLIE%))
         Call CARDACLI
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
   NumeroProveedor& = XVALO(LText1) * (-1)
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

Private Sub LText1_Change()
   Call CARDACLI
End Sub

Private Sub MarItCum_Click()
  Call OPCICOM07.Command4_Click
End Sub

Private Sub Amastock_Click()
  Call Command8_Click
End Sub

Private Sub Nuorig_Click()
   Call OPCICOM07.Command19_Click
End Sub

Private Sub Nuorig1_Click()
   Call OPCICOM07.Command19_Click
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
   Command11_Click
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




Private Sub Text1_DblClick()
   OBSERVAC.Caption = "Datos de Contacto"
   OBSERVAC.Label24.Caption = ""
   OBSERVAC.Label2 = 128
   OBSERVAC.Text28.MaxLength = 128
   OBSERVAC.Text28 = Text1

   OBSERVAC.Show 1
   
   If OBSERVAC.Label1 = "OK" Then
     Text1 = OBSERVAC.Text28
   End If


End Sub

Private Sub Text10_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
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
Private Sub Text11_GotFocus()
   Text6.SetFocus
End Sub

Private Sub Text12_Change()
   Call CALTOCOM
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



Private Sub Text16_Click()
   On Error Resume Next
   Text6.SetFocus
   On Error GoTo 0
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

Private Sub Text2_gotfocus()
   On Error Resume Next
   Text6.SetFocus
   On Error GoTo 0
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

Private Sub Text3_GotFocus()
   On Error Resume Next
   Text6.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text4_GotFocus()
   On Error Resume Next
   Text6.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text5_GotFocus()
   On Error Resume Next
   Text6.SetFocus
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
   If XVALO(LText1) > 0 Then
     If XVALO(LText1) <= 32767 Then
       NCLIE% = XVALO(LText1)
       NPPX% = XVALO(VALOBADA("PROVED12", "Nume_Cli", "Nume_Cli = " & NCLIE%))
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
   Call SETULTREG("!OCOMDETA", 0)
   Call CIERRARCH("!OCOMDETA")
   Call BLANARCH("!OCOMDETA")
   On Error Resume Next
   Unload FORVENTB
   On Error GoTo 0
   '\\\OT-08-0969-OD-24/10/08///
   Label18 = ""
   '\\\OT-08-0969-OD-FIN///
   Label8 = ""
   Text2.TEXT = ""
   Text3.TEXT = ""
   Text4.TEXT = ""
   Text5.TEXT = ""
   'Label2 = ""
   Text8.TEXT = ""
   Label118.Caption = FORMATNUM$(ALIVA, "F5.2")
   Text13.TEXT = ""
   Text6.TEXT = ""
   Text7.TEXT = ""
   Text18.TEXT = ""
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
   Text9.TEXT = ""
   Text11.TEXT = ""
   '
   Text22 = ""
   TEXT24 = ""
   TEXT25 = ""

End Sub
Private Sub Option2_Click()
   If TEXT24.Enabled = False Then
    TEXT24.Enabled = Not (TEXT24.Enabled)
    Eco24.Enabled = Not (Eco24.Enabled)
    TEXT25.Enabled = Not (TEXT25.Enabled)
    Command25.Enabled = Not (Command25.Enabled)
   End If
End Sub

Private Sub Option1_Click()
   If TEXT24.Enabled = True Then
    TEXT24.Enabled = Not (TEXT24.Enabled)
    Eco24.Enabled = Not (Eco24.Enabled)
    TEXT25.Enabled = Not (TEXT25.Enabled)
    Command25.Enabled = Not (Command25.Enabled)
   End If

End Sub

Private Sub Text24_Change()
   Eco24 = rasocli$(XVALO(TEXT24))
End Sub

Private Sub Command25_Click()
   Call SELECHO("DEUDOR1")
   NCLIEN = XVALO(TRIM$(VALACT1$("DEUDOR1")))
   TEXT24.TEXT = TRIM$(Str$(NCLIEN))
End Sub

'
Sub CARDACLI()
    '
    Screen.MousePointer = 11
    '
    NC% = (-1) * XVALO(LText1)
    If NC% < 0 Then
      If STATCLI%(NC%) > 0 Then
        Label8 = CODICLI$(NC%)
        Text2.TEXT = rasocli$(NC%)
        Text3.TEXT = DOMICLI$(NC%)
        Text4.TEXT = CPOSCLI$(NC%)
        Text5.TEXT = LOCACLI$(NC%)
        'Label2 = TRIM$(POSIVAC$(NC%))
        Text8.TEXT = TRIM$(CuitCli$(NC%))
      End If
    End If
    '
    Call CambiosEnPantallaCompraInterna(XVALO(LText1))

    Screen.MousePointer = 1
    '
End Sub
   '
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
   Text6.SetFocus
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
     RIVACOM07.TONECOM = FORMATNUM$(XVALO(Total(0)) - XVALO(Total(1)), "F12.2")
     Label118 = TRIM$(RIVACOM07.Label2)
     Total(2) = TRIM$(RIVACOM07.Label3)
  End If
If Index = 2 Then
AAA = BBB
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

'
Sub GRACOMPRA(OKX%, Optional VALIDACIONSTOCKMINIMO$, Optional RetNroCompra&)
   '
   OKX% = 0
   If ULTREG&("!OCOMDETA") < 1 Then Exit Sub
   '
   COPRECIO = 1
   If Total(0).Enabled = False Then COPRECIO = 0
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))

   If InStr(1, EMPREAC$, "AHP") > 0 Then
       Call SELECHO("VIASENTR")
       VIA% = XVALO(VALACT1$("VIASENTR"))
   End If

   Call CIERRARCH("OCOMENCA")
   Call CIERRARCH("OCOMDETA")
   '
   Call BLOQARCH("OCOMENCA")
   Call BLOQARCH("OCOMDETA")
   Call BLOQARCH("MACONSIG")
   '
   NRO& = XVALO(Text11)
   '

   MONECO% = XVALO(Text15)
   If MONECO% < 1 Then MONECO% = XVALO(TRIM$(Control$("OCOMPRA", "MONECOS")))
   If MONECO% < 1 Then MONECO% = 1
   '
   Screen.MousePointer = 11
   If VALICLIE% < 1 Then
      MERRO$ = "Número de Proveedor no Válido"
      GoTo 98
   End If
   NPRO% = XVALO(LText1)
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
      GoTo 98
   End If
   '
   NR1$ = XVALO(Text11)
   NR1$ = NR1$ & "-" & (Label16)
   '
   EsProvInterno% = 0
   StringNroCtes$ = StringDeNroProveedoresComprasInternas$
   'SI EL NUMERO DE CLIENTE ESTA DENTRO DE LA CADENA DE NRO DE PROVEEDORES CONFIGURADOS COMO PARTICIPANTES DE LA COMPRA INTERNA.
   If NumeroEstaEnCadena(StringNroCtes$, NPRO%) Then EsProvInterno% = 1

123 If COPRECIO < 0.99 Then
      PORDESCU.TEXT = ""
      Text17 = ""
      For KK% = 0 To 4
        Total(KK%).TEXT = ""
      Next KK%
    End If
    '
    STATUOC% = 0
    If Left$(Control$("OCOMPRA", "MODOAPRO"), 6) <> "DIRECT" Then STATUOC% = -1
    If VALIDACIONSTOCKMINIMO$ = "GRABA_PENDIENTE_APROBACION" Then STATUOC% = -2

    Call APERBASDAT("COMPRAS")
    SQLX$ = " SELECT * FROM OCOMENCA "
    SQLX$ = SQLX$ + "Where Nume_Ocom = " & NRO&
    SQLX$ = SQLX$ + " And Stat_Ocom < 9 "   ' EXCLUYE ANULADAS
    Dim OCOMENCA As ADODB.Recordset
    Set OCOMENCA = New ADODB.Recordset
    OCOMENCA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With OCOMENCA
       If (.EOF And .BOF) Then
         MERRO$ = "No se Ha Encontrado Registro\de Encabezado de esta O/Compra."
         GoTo 98
       End If
       '
       FEX = FEMI%
       REFE$ = FECHATEX$(FEX) + " - " + rasocli$((-1) * NPRO%)
       !Refe_OcoM = Left$(REFE$, 64)
       !Femi_Ocom = CVDAT(FEMI%)
       !Npro_Ocom = NPRO%
       !Raso_Prov = Left$(rasocli$((-1) * NPRO%), 48) 'OT-06-0227-WAR-05/06/06///
       !TSum_Ocom = 0               ' materiales
       '
       !FMod_Ocom = Now
       !Usua_Mod = USNBR%
         '
         MONECOI% = XVALO(Text15)
       !MONE_COS = MONECOI%        ' antes pos 83,1
       !Mone_Emis = MONEMI%        ' antes pos 84,1
         TICAMBX = 1
         If MONEMI% <= 1 Then
           If MONECOI% > 1 Then
             TICAMBX = TICAMONE(MONECOI%)
           End If
         End If
       !Tcam_Ocom = TICAMBX       ' antes pos 85, 4
'       COTIZA_DOLAR QUEDA IGUAL NO SE MODIFICA

         '
         CPAG% = XVALO(TEXT21(0))
       !Cpag_Ocom = CPAG%
         EMBAL% = XVALO(TEXT21(1))
       On Error Resume Next    ' POR SI NO ESTA DEFINIDO
       !Embala_Ocom = EMBAL%
       '\\\OT-06-0082-WAR-06/03/06///
       !OCOMREVIS = XVALO(Label16)
       '\\\OT-06-0082-WAR-FIN///
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
       '\\\OT-08-0887-OD-24/09/08///
       !PorcIva = Label118
       '\\\OT-08-0887-OD-fin///
       '
'       STATUOC% = 0
'       If Left$(CONTROL$("OCOMPRA", "MODOAPRO"), 6) <> "DIRECT" Then STATUOC% = -1
'       If VALIDACIONSTOCKMINIMO$ = "GRABA_PENDIENTE_APROBACION" Then STATUOC% = -2
       !Stat_Ocom = STATUOC%       ' status Pendiente de aprobacion
       
       On Error Resume Next
       !Anex_Ocom = RAI$
       On Error GoTo 0
       '
       If CODCARPE$ <> "" Then
            Call REPLA(TTXX$, HOS$, 1, 8)
            Call REPLA(TTXX$, USERNAME$, 69, 24)
            NR$ = SAFETEXT(Text11)
            Call REPLA(TTXX$, "Modificacion de Orden de Compra Nº: " + NR$, 11, 56)
            Call AGREGA_ANOTACARPI(CODCARPE$, TTXX$)
            TTXX$ = Space(76)
            Call REPLA(TTXX$, "Proveedor: " + Left$(rasocli$((-1) * NPRO%), 64), 11, 56)
            Call AGREGA_ANOTACARPI(CODCARPE$, TTXX$)
            Call AGREGA_ANOTACARPI(CODCARPE$, String$(88, "="))
       End If
       !Respon_Ocompra = Left$(Text1, 128)
       !Via_Entrega = XVALO(VIA%)
       !OComCliente = Left(TEXT25, 16)
       !nume_cli = XVALO(TEXT24)
       .Update
    End With
    OCOMENCA.Close
    Set OCOMENCA = Nothing
    '
    RetNroCompra& = NRO&
    Label13 = RAI$
    
    'anular items en sql
125 CONDI$ = "Stat_Ocom < 9 and Nume_Ocom = " & NRO&
    Call ACTUREGISTRO("OCOMDETA", "Stat_Ocom", CONDI$, 10, RAFE&, , 99)
    If CantRegistros&("OCOMDETA", CONDI$) > 0 Then GoTo 125
    '
    'aqui graba el detalle del IVA
    STRSQL = "DELETE DETAIVA where Nume_Ocom = " & NRO&
    FLEXCONN.Execute (FILTSQL$(STRSQL))
    '
    STRSQL = "Select * from DETAIVA where Nume_Ocom = " & NRO&
    Dim DETIVA As ADODB.Recordset
    Set DETIVA = New ADODB.Recordset
    DETIVA.Open FILTSQL$(STRSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With RIVACOM07
      For KKII% = 0 To 5
        If XVALO(.IMPOGRA(KKII%)) >= 0.005 Or XVALO(.IMPOIVA(KKII%)) >= 0.005 Then
          DETIVA.AddNew   ' YA CONVERTIDO
                On Error Resume Next
          DETIVA!CODIEMPR = CodiEmp%
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
126 EMICONSI% = 0
    IDENOC$ = MKS$(NRO&) + MKI$(FEMI%) + MKI$(NPRO%)
    '
    ORDITMA% = 0
    ARCHILIB$ = "!1OCOMDET"
    '
    Dim ACURECI(32767)
    ' aca
    
      STRSQL = "Select * from Ocomdeta where Nume_ocom=" & NRO&
      Dim OCOMDETA As ADODB.Recordset
      Set OCOMDETA = New ADODB.Recordset
      
      ACUMUMENSAX$ = "": ENVIOASUCURSAL% = 0
      If Control("OCOMPRA", "ENVARSUC") = "SI" Then ENVIOASUCURSAL% = 1 'CONFIGURACION SETUP DE COMPRA DE ENVIO DE ARCHIVO DE O.C. A CENTRAL
      
      OCOMDETA.Open FILTSQL$(STRSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
  For J& = 1 To ULTREG&("!OCOMDETA")
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
      ORDIT% = ORDIT% + 1
      If ENVIOASUCURSAL% = 1 Then
        If VALIDACIONSTOCKMINIMO$ <> "GRABA_PENDIENTE_APROBACION" Then
           CANTU = CVS(Mid$(ZZ$, 55, 4)): FEENT% = CVI(Mid$(ZZ$, 95, 2))
           MENSAX$ = MENSAJEDECAMBIO$(CI0%, CANTU, NRO&, ORDIT%, FEENT%)
           ACUMUMENSAX$ = ACUMUMENSAX$ + MENSAX$
        End If
      End If
      '
      If EsProvInterno% > 0 Then 'para conteplast, cambaceres,quilmes
           CANTU = CVS(Mid$(ZZ$, 55, 4)): FEENT% = CVI(Mid$(ZZ$, 95, 2))
           MENSAX$ = MENSAJEDECAMBIO$(CI0%, CANTU, NRO&, ORDIT%, FEENT%)
           ACUMUMENSAX$ = ACUMUMENSAX$ + MENSAX$
      End If

      With OCOMDETA
        .AddNew   ' YA CONVERTIDO
        On Error Resume Next
        !CODIEMPR = CodiEmp%
        On Error GoTo 0
         !nume_ocom = NRO&
           ORDIT% = CVS(Mid$(XX$, 125, 4))
           If ORDIT% < 1 Then ORDIT% = ORDITMA% + 1
           If ORDIT% > ORDITMA% Then ORDITMA% = ORDIT%
         !NuOrd_Item = ORDIT%
         !Femi_Ocom = CVDAT(FEMI%)
         !Medi_Ocom = SAFETEXT$(PedInt$)
         !Npro_Ocom = NPRO%
         !Raso_Prov = Left$(rasocli$((-1) * NPRO%), 64) 'OT-06-0227-WAR-05/06/06///
         !cod_inte = CI0%
         !Cod_Exte = Left$(CODEXT$(CI0%), 24)
         !DESCRITEM = Left$(DESCRIT0$(CI0%), 24)
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
           CANTU = CVS(Mid$(ZZ$, 55, 4))
         !Cant_Ocom = CANTU
         !Uni_Com = Mid$(ZZ$, 59, 4) ' unidad de compra
         !Coef_Unids = COEFUNI
         !Pre_Unit = CVD(Mid$(ZZ$, 67, 8))
         !Desc_Item = CVS(Mid$(ZZ$, 75, 4))
         !prun_neto = CVD(Mid$(ZZ$, 79, 8))
         !INet_Item = CVD(Mid$(ZZ$, 87, 8))
         !PrUn_Neto_Pes = PREUNCO# 'Se agrega este campo para guardar el valor en pesos
         !Fentre_Sol = CVDAT(CVI(Mid$(ZZ$, 95, 2)))
         !CONSIMAT = Left$(TRIM$(Mid$(ZZ$, 97, 2)), 2)
            CFXUNI# = !Coef_Unids
            If CFXUNI# < 0.005 Then CFXUNI# = 1
         CANRECI = CANTORECI(CI0%, NRO&) - ACURECI(CI0%)
         If CANRECI > CAUNIS Then CANRECI = CAUNIS
         ' !Cant_Rec = CANTORECI(CI0%, NRO&) / CFXUNI#
         !Cant_Rec = CANRECI / CFXUNI#
         ACURECI(CI0%) = ACURECI(CI0%) + CANRECI
            ' AQUI AGREGO PARA TMAR LOS VALORES ANTERIORES
            CONDI$ = "Nume_Ocom = " & NRO& & " "
            CONDI$ = CONDI$ + " AND NuOrd_Item = " & ORDIT% & " "
            CONDI$ = CONDI$ + " AND COD_INTE = " & CI0% & " "
            CONDI$ = CONDI$ + " AND N_ORSERV > 0 "
            CODOPI% = XVALO(VALOBADA("OCOMDETA", "Cod_Oper", CONDI$))
         !Cod_Oper = CODOPI%
            DOCORI$ = SAFETEXT$(VALOBADA("OCOMDETA", "Doc_Orig", CONDI$))
         !DOC_ORIG = DOCORI$
            NOSERVI& = XVALO(VALOBADA("OCOMDETA", "N_OrServ", CONDI$))
         !N_OrServ = NOSERVI&
         ' FIN DE INTERVENCION PARA MELINA - sI ESTO FUNCIONA, TENDRÍA QUE
         ' 1.- Anular esas 4 BR's
         ' 2.- Volve a modificar las OC`s correspondientes
         ' 3.- Volver procesar las BR's.
         ' Testear e laboratorio ates de libera c on el cliente
         '
         '
         '
         !MONE_COS = MONECOI%        ' antes pos 83,1
         !Mone_Emis = MONEMI%        ' antes pos 84,1
         !TSum_Ocom = 0              ' materiales
'          STATUOC% = XVALO(Label15)   'SE PASO ARRIBA PARA LEERLO SOLO UNA VEZ
'          If Left$(CONTROL$("OCOMPRA", "MODOAPRO"), 6) <> "DIRECT" Then STATUOC% = (-1)
'          If VALIDACIONSTOCKMINIMO$ = "GRABA_PENDIENTE_APROBACION" Then STATUOC% = -2
          !Stat_Ocom = STATUOC%
            
            NUECUEIMPO$ = TRIM$(Mid$(ZZ$, 105, 12))
         !IContab = CUECOINT(TRIM$(NUECUEIMPO$)) ' Imputacion contable
            SECTOROC$ = TRIM$(Mid$(ZZ$, 15, 12))
         !Sect_Ocom = SECTOROC$   ' Sector
            '
        .Update
      End With
      
      ' EN ESTE SUB-BLOQUE GENERA LAS NECESIDADES DE MATERIALES A CONSIGNAR AL PRO-
      ' VEEDOR SEGÚN LA ESTRUCTURA DE PRODUCTO DE CADA REGISTRO DE LA O/COMPRA
      NOCOM$ = TRIM$(Str$(NRO&))
        While Len(NOCOM$) < 6: NOCOM$ = "0" + NOCOM$: Wend
      ITNU$ = TRIM$(Str$(ORDIT%))
        While Len(ITNU$) < 2: ITNU$ = "0" + ITNU$: Wend
      NOCOM$ = "OC-" + Left$(NOCOM$, 6) + "-" + ITNU$
      '
      'SE AGREGAN ESTAS LINEAS PARA QUE NO GENERE LA CONSIGNACION DE UNA O/C POR SERVICIO DE TERCEROS
      CONDI$ = " NUME_OCOM = " & NRO& & " AND Nuord_Item = " & ORDIT%
      NORFA$ = TRIM$(VALOBADA("OCOMDETA", "DOC_ORIG", CONDI$, "NOMESS"))
      If NORFA$ <> "" Then GoTo 135
      '
      Call LEEEXPLO(CODEXT$(CI0%), CAUNIS)
      '
      Call APERBASDAT("STOCKS")
      STRSQL = "Select * from Maconsig where Nro_ocom='" & NOCOM$ & "'"
      Dim MaConsig As ADODB.Recordset
      Set MaConsig = New ADODB.Recordset
      MaConsig.Open FILTSQL$(STRSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
          
      For i% = 1 To CAIT%
        CII% = CIJ%(i%)
        CANTIASIGN# = CDbl(CAUNIS * USOI(i%))
        '
        If TRIM$(CODEXT$(CII%)) <> "" Then
          If CANTIASIGN# > 0.05 Then
            '
            'EN CASO QUE EXISTA MATERIAL YA CONSIGNADO SE GENERA UNA CONSIGNACION POR EL RESTO ENTRE LA NUEVA CANTIDAD Y LO CONSIGNADO
            CONDI$ = "Nro_ocom = '" & NOCOM$ & "' AND CANT_ASIG > 0 AND COD_INTE = " & CII%
            CANYASIG# = XVALO(VALOBADA("MACONSIG", "SUM(CANT_ASIG)", CONDI$, "NOMESS"))
            If CANYASIG# > 0 Then
                CANTIASIGN# = CANTIASIGN# - CANYASIG#
                If CANTIASIGN# = 0 Then GoTo 130
            End If
            '
            With MaConsig
              .AddNew   ' YA CONVERTIDO
               On Error Resume Next
               !CODIEMPR = CodiEmp%
               On Error GoTo 0
               !Femi_Ocom = CVDAT(FEMI%)
               !Npro_Ocom = NPRO%
               !Raso_Prov = Left$(rasocli$((-1) * NPRO%), 64) 'OT-06-0227-WAR-05/06/06///
               !nro_ocom = NOCOM$
                '
               !cod_inte = CII%
               !Cod_Exte = Left$(CODEXT$(CII%), 24)
               !DESCRITEM = Left$(DESCRIT0$(CII%), 64)
               !Unidad = Unimed$(CII%)
               !cant_asig = CANTIASIGN#
               !Secc_Ejec = 0
               !Deno_Secc = ""
               !NRO_ORFA = ""
               !LOTE_OCOM = ""
               !cant_entreg = 0
               !FECHA_ENTREG = CVDAT(0)
               !NREMI_CONSIG = ""
               !Valo_Unit = COSUNI(CII%)
               !Moneda = MONECOSTO%(CII%)
               !Valo_Tot = CANTIASIGN# * COSUNI(CII%) * TICAMONE(MONECOSTO(CII%))
               !CANT_RENDID = 0
               !Saldo_Pend = 0
               !Observa = ""
              .Update
              EMICONSI% = 1
            End With
          End If
        End If
130   Next i%
      MaConsig.Close
      Set MaConsig = Nothing
      '
    Next J&
   '
135
   OCOMDETA.Close
   Set OCOMDETA = Nothing
   Call BAJALDISCO
   '
   If COPRECIO < 0.99 Then
     PORDESCU.TEXT = ""
     Text17 = ""
     For KK% = 0 To 4
       Total(KK%).TEXT = ""
     Next KK%
   End If
   '
   ' aqui actualiza costo y peso por unidad
   ACOSREP$ = UCase$(Control$("OCOMPRA", "ACOSREP"))
   '
   HOII% = HOY(HOS$)
   '
   Dim Master As ADODB.Recordset
   '
   ORDIT% = 0
   For J& = 1 To ULTREG&("!OCOMDETA")
     ZZ$ = REGLEIDO$("!OCOMDETA", J&)
     ORDIT% = ORDIT% + 1
     CI0% = CVI(Mid$(ZZ$, 9, 2))
     CONDI$ = "Nume_Ocom = " & NRO& & " " ' MODIFICACION PARA QUE ACTUALCE EL PRECIO DE COSTO CUANDO ES ORDEN DE SERVICIO
     CONDI$ = CONDI$ + " AND NuOrd_Item = " & ORDIT% & " "
     CONDI$ = CONDI$ + " AND COD_INTE = " & CI0% & " "
     CONDI$ = CONDI$ + " AND N_ORSERV > 0 "
     '
     NUORSE& = XVALO(VALOBADA("OCOMDETA", "N_OrServ", CONDI$))
     If NUORSE& > 0 Then
       CODOPI% = XVALO(VALOBADA("OCOMDETA", "Cod_Oper", CONDI$))
       CONDI1$ = "CodIConj = " & CI0% & " AND NumeOper = " & CODOPI% & ""
       PREUNINE = FORMATNUM(CVD(Mid$(ZZ$, 79, 8)), "F10.2")
       Call ACTUREGISTRO("SECOPER", "PREUNID", CONDI1$, PREUNINE, REGAF&)
       Call ACTUREGISTRO("SECOPER", "PROTEROP", CONDI1$, HOY(HO$), REGAF&)
     End If
     '
     If NUORSE& < 1 Then   ' EXCLUYE ORDEN DE SERVICIO
        CI0% = CVI(Mid$(ZZ$, 9, 2))
        If CI0% < 1 Then GoTo 170   ' CÓDIGO NULO
        '
        COEFUNI = CVS(Mid$(ZZ$, 63, 4)): If COEFUNI < 0.0001 Then COEFUNI = 1
        If Control("OCOMPRA", "NODESCUE") <> "SI" Then ' ACA ENTRA COMO ESTABA
           COSPESOSD# = CVD(Mid$(ZZ$, 79, 8)) * TICAMONE(MONEMI%) / COEFUNI
        Else ' SI ESTA MARCADO EN SETUP DE COMPRAS NO CONSIDERAR DESCUENTO
           COSPESOSD# = CVD(Mid$(ZZ$, 67, 8)) * TICAMONE(MONEMI%) / COEFUNI
        End If
           
        '
        'SQL
        STRSQL = "Select * from Master where codint=" & CI0%
        Set Master = New ADODB.Recordset
        Master.Open FILTSQL$(STRSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        
        If Master.BOF And Master.EOF Then
           Master.Close
           Set Master = Nothing
           GoTo 170
        End If
        '
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
                  Call ADD_NOVEDAD(CI0%, NPRO%, COSTOANTER#, PRENUE#, "Actualización Por Orden de Compra: " & NR1$, MONECO%, "COSTO")
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
        Master.Close ' se corrio por que si quedaba fuera del if daba error 91 si no habia entrado
        Set Master = Nothing
     End If
170 Next J&
   '
   'termina la actualización
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
    Screen.MousePointer = 1
    
    OKX% = 1
    Exit Sub
    '
98  If OCOMENCA Is Nothing Then GoTo 101
    OCOMENCA.Close
    Set OCOMENCA = Nothing
    OCOMDETA.Close
101 If OCOMDETA Is Nothing Then GoTo 102
    Set OCOMDETA = Nothing
    Call CIERRARCH("*.*")
102 Screen.MousePointer = 1
    Call MENSERR(24, MERRO$)
    '
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

Sub CARDETCOM()
   '
   'ESTO AGREGADO PARA QUE BORRE SI ALGUN ITEM NO TIENE CANTIDAD, ESTA VALIDACION ESTABA EN LA RUTINA CALTOCOM PERO SE QUITO
   'POR QUE TRAIA PROBELMAS CUANDO RECORRIA EL ARCHIVO OCOMDETA PARA ACTUALIZAR EN EL TIMER1 (15-0348)
'   FIN& = ULTREG&("!OCOMDETA")
'   For U& = 1 To FIN&
'     X$ = REGLEIDO$("!OCOMDETA", U&)
'     CI% = CVI(Mid$(X$, 9, 2))
'     '
'     CANSTOK = CVS(Mid$(X$, 51, 4))
'     If CANSTOK > 0.05 Then
'       U1& = U1& + 1
'       X$ = REGLEIDO$("!OCOMDETA", U&)
'       Call GRAREG("!OCOMDETA", X$, U1&)
'     End If
'   Next U&
'   Call SETULTREG&("!OCOMDETA", U1&)
   'FIN (15-0348)

   '\\\OT-05-0690-WAR-24/10/05///
   'PRIMER SWTICH = SI FORMATO = DESTINO
   If Control$("OCOMPRA", "MODOIMPU") = "DES" Then
     Call CARDESCOM
     Screen.MousePointer = 1
     Exit Sub
   End If
   '\\\OT-05-0690-WAR-FIN///
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
   CONDI$ = "NUME_OCOM = " & XVALO(Text11)
   ORDSER& = XVALO(VALOBADA("OCOMDETA", "N_ORSERV", CONDI$))

   EPAC$ = TRIM$(UCase$(EMPREAC$))
   MODOGABAL% = 0: If InStr(EPAC$, "GABAL") > 0 And ORDSER& > 0 Then MODOGABAL% = 1
   If MODOGABAL% > 0 Then
    ATRI$ = ATRI$ + "/POSEDIT(1)=2"
    ATRI$ = ATRI$ + "/POSEDIT(2)=2"
    ATRI$ = ATRI$ + "/POSEDIT(3)=2"
    ATRI$ = ATRI$ + "/POSEDIT(4)=2"
    ATRI$ = ATRI$ + "/POSEDIT(5)=2"
    ATRI$ = ATRI$ + "/POSEDIT(6)=2"
    ATRI$ = ATRI$ + "/POINDEL=NO" ' DESHABILITA EL INSERTAR Y EL DEL
   End If
   '
   Timer1.Enabled = True
   PRF$ = "CARMODCO"
   If XVALO(Label15) < 0 Then ' no esta aprobada - puede editarse todo
     PRF$ = "CARDETCO"
   End If
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
      LText1 = ""
      Hide
      Exit Sub
   End If
   '
   Screen.MousePointer = 11
   ATK% = 0
   NPRO% = XVALO(LText1)                    'Se corrigio asignacion de variable el nombre del objeto es LTEXT1
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

      If TRIM$(Mid$(ZZ$, 11, 4)) = "" Then Call REPLA(ZZ$, Unimed$(CI0%), 11, 4)
      If TRIM$(Mid$(ZZ$, 59, 4)) = "" Then Call REPLA(ZZ$, UNICOM$(CI0%), 59, 4)
      If TRIM$(Mid$(ZZ$, 11, 4)) = TRIM$(Mid$(ZZ$, 59, 4)) Then Call REPLA(ZZ$, MKS$(1), 63, 4)
      '
      UC$ = ""
      COEFI = CVS(Mid$(ZZ$, 63, 4))
      CANCOMP = CVS(Mid$(ZZ$, 55, 4))
      CANSTOK = CVS(Mid$(ZZ$, 51, 4))
      If COEFI < 0.0001 Then
        If CANCOMP > 0.05 Then
          If CANSTOK > 0.05 Then
            COEFI = CANSTOK / CANCOMP
          End If
        End If
      End If
      '
      If COEFI < 0.0001 Then
        UC$ = ULTICOM$(NPRO%, CI0%)
        If TRIM$(Mid$(ZZ$, 11, 4)) = TRIM$(Mid$(UC$, 11, 4)) Then
          If TRIM$(Mid$(ZZ$, 59, 4)) = TRIM$(Mid$(UC$, 59, 4)) Then
            COEFI = CVS(Mid$(UC$, 63, 4))
          End If
        End If
      End If
      If COEFI < 0.0001 Then COEFI = 1
      '
      'VALIDAR SI TIENE LOTE INDIVISIBLE- VALIDA QUE LA CANTIDAD SEA MULTIPLO DE L.INDIV.
      EPAC$ = TRIM$(UCase$(EMPREAC$))
      MODOGABAL% = 0: If InStr(EPAC$, "GABAL") > 0 Then MODOGABAL% = 1
      LINDIVISBLE = LINDIVIS(CI0%)
      If MODOGABAL% > 0 Then LINDIVISBLE = LINDIVEN(CI0%)
      If InStr(EPAC$, "NEUQUEN") < 1 Then 'NEUQUEN NO VALIDA N.SERIE
        If LINDIVISBLE > 0.05 Then
          CANTID# = ROUND#(XVALO(CANSTOK), 1)
          If MODOGABAL% < 1 Then ' SI NO ES GABAL
             CANLINDIV# = ROUND#(REDON_INDIVIS(CI0%, CANTID#))
          Else
             'SI ES GABAL
             CANLINDIV# = ROUND#(REDON_INDIVIS(CI0%, CANTID#, "VENTA"))
          End If
          
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
      '
      If CANCOMP < 0.05 Then
          CANCOMP = CANSTOK / COEFI
        Else
          CANSTOK = CANCOMP * COEFI
      End If
      
      CANCOMP = CANSTOK / COEFI
      Call REPLA(ZZ$, MKS$(CANSTOK), 51, 4)
      Call REPLA(ZZ$, MKS$(CANCOMP), 55, 4)
      Call REPLA(ZZ$, MKS$(COEFI), 63, 4)
      '
'      If InStr(EPAC$, "CONTEPLAST") > 0 And TIMATE%(CI0%) = 41 Then 'si es conteplast y tipomat es bobinas
'         StringNroCtes$ = StringDeNroProveedoresComprasInternas$
'        'SI EL NUMERO DE CLIENTE ESTA DENTRO DE LA CADENA DE NRO DE PROVEEDORES CONFIGURADOS COMO PARTICIPANTES DE LA COMPRA INTERNA.
'        If NumeroEstaEnCadena(StringNroCtes$, NPRO%) Then
'        Peso# = PESUNI(CI0%)
'        If Peso# > 0.00005 Then
'            coefi = 1 / Peso#
'            CANCOMP = ROUND(CANSTOK * coefi)
'        End If
'       End If
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
   '
End Sub
'\\\OT-05-0690-WAR-24/10/05///
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
   PRF$ = "CAMOCODE"
   If XVALO(Label15) < 0 Then ' no esta aprobada - puede editarse todo
     PRF$ = "CARDESCO"
   End If
   XXX% = VENTABU%(PRF$ + ATRI$) 'aca se presenta
   List1.Clear
   Call CALTOCOM
   Timer1.Enabled = False
   '
   If XXX% < 0 Or ULTREG&("!OCOMDETA") < 1 Then
      Call CIERRARCH("!OCOMDETA")
      Call BLANARCH("!OCOMDETA")
      Call BLANFORMU
      Total(0).TEXT = ""
      LText1 = ""
      Hide
      Exit Sub
   End If
   '
   Screen.MousePointer = 11
   NPRO% = XVALO(LText1)              'Se corrigio asignacion de variable el nombre del objeto es LTEXT1
   For U& = 1 To ULTREG&("!OCOMDETA")
      ZZ$ = REGLEIDO$("!OCOMDETA", U&)
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
      CANSTOK = CVS(Mid$(ZZ$, 51, 4))
      CANCOMP = CANSTOK
      UNIST$ = UCase$(TRIM$(Mid$(ZZ$, 11, 4)))
      UNICO$ = UNIST$
      Call REPLA(ZZ$, MKS$(CANSTOK), 51, 4)
      Call REPLA(ZZ$, MKS$(CANCOMP), 55, 4)
      '
      'VALIDAR SI TIENE LOTE INDIVISIBLE- VALIDA QUE LA CANTIDAD SEA MULTIPLO DE L.INDIV.
      If InStr(EPAC$, "NEUQUEN") < 1 Then 'NEUQUEN NO VALIDA N.SERIE
        If LINDIVIS(CI0%) > 0.05 Then
          CANTID# = ROUND#(XVALO(CANSTOK), 1)
          CANLINDIV# = ROUND#(REDON_INDIVIS(CI0%, CANTID#))
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
      '
      Call REPLA(ZZ$, MKS$(CANCOMP), 55, 4)
      Call REPLA(ZZ$, UNICO$, 59, 4)
      Call REPLA(ZZ$, MKS$(COEFI), 63, 4)
      '
      PreUnid# = CVD(Mid$(ZZ$, 67, 8))
      If PreUnid# < 0.00005 Then
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
          Call REPLA(ZZ$, MKD$(PreUnid#), 67, 8)
          Call REPLA(ZZ$, MKS$(PORDESCUEI), 75, 4)
        End If
      End If
      '
      Call GRAREG("!OCOMDETA", ZZ$, U&)
   Next U&
   '
   Screen.MousePointer = 1
   '
   Call CARLISCO
   Call CALTOCOM
   Screen.MousePointer = 1
   '
End Sub
'\\\OT-05-0690-WAR-FIN///
Sub CALTOCOM()
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   NC1% = (-1) * XVALO(LText1)
   TOTANE# = 0:: JJ& = 0
   For J& = 1 To ULTREG&("!OCOMDETA")
       X$ = REGLEIDO$("!OCOMDETA", J&)
       AX$ = REGLEIDO$("!1OCOMDET", J&)
       
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

       'SI NO TIENE CANTIDAD LO BORRA
'       If CANSTXX < 0.05 And CANTU < 0.05 Then GoTo 30       '
       PREUNINE# = (Int(PreUnid# * 100 * (100 - PORDE) + 0.5)) / 10000
       INETITEM# = CDbl((Int(100 * PREUNINE# * CANTU + 0.5)) / 100)
       TOTANE# = TOTANE# + INETITEM#
       '
       'AGREGO ESTAS LINEAS PARA QUE NO LEA SI NO ES NECESARIO Y NO PROVOQUE LENTITUD EN OTROS CLIENTES
       If InStr(EPAC$, "AHP") > 0 Then
         Peso# = PESUNI(CI0%)
         ACUMULAPESO# = ACUMULAPESO# + (Peso# * CANSTXX)
       End If
       '
       Call REPLA(X$, MKD$(PREUNINE#), 79, 8)
       Call REPLA(X$, MKD$(INETITEM#), 87, 8)
       Call GRAREG("!OCOMDETA", X$, J&)
       '
       UNIST$ = UCase$(TRIM$(Mid$(X$, 11, 4)))
       UNICO$ = UCase$(TRIM$(Mid$(X$, 59, 4)))
       CANSTXX = CVS(Mid$(X$, 51, 4))
       '
       COEFI = 1
       If CANTU > 0.05 Then COEFI = CANSTXX / CANTU
       Call REPLA(X$, MKS$(COEFI), 63, 4)
       JJ& = JJ& + 1
       Call GRAREG("!OCOMDETA", X$, JJ&)
       Call GRAREG("!1OCOMDET", AX$, JJ&)
       '
30 Next J&

'       If CANTU > 0.05 Then
'         COEFI = CANSTXX / CANTU
'         Call REPLA(X$, MKS$(COEFI), 63, 4)
'         JJ& = JJ& + 1
'         Call GRAREG("!OCOMDETA", X$, JJ&)
'       End If
       '
'   Next J&
   Call SETULTREG("!OCOMDETA", JJ&)
   Call CIERRARCH("!OCOMDETA")
   Call SETULTREG("!1OCOMDET", JJ&)
   Call CIERRARCH("!1OCOMDET")
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
   
   TOTATO# = TOTANEDES# + TOTIVA# + TOTPER#
   '
   Total(0) = FORMATNUM$(TOTANE#, "F11.2")
   Total(1) = FORMATNUM$(IDESCUEN#, "F11.2")
     TOTIVA# = APLIVA% * XVALO(RIVACOM07.Label3)
   Total(2) = FORMATNUM$(TOTIVA#, "F11.2")
   Total(4) = FORMATNUM$(TOTPER#, "F11.2")
     TOTATO# = TOTANEDES# + TOTIVA# + TOTPER
   Total(3) = FORMATNUM$(TOTATO#, "F11.2")
   
   If InStr(EPAC$, "AHP") > 0 Then TXTPESO = FORMATNUM$(ACUMULAPESO#, "F10.2")
   '
End Sub

Sub CARCOANT()
   '
   Screen.MousePointer = 11
   List2.Clear
   NPRO% = XVALO(LText1)
   NOCOA& = XVALO(Text11)
   HOU = HOY(HOS$)
   '
   Call APERBASDAT("COMPRAS")
   SQLX$ = " SELECT * FROM OCOMDETA "
   SQLX$ = SQLX$ + " WHERE Stat_Ocom < 8 "
   SQLX$ = SQLX$ + " AND Stat_Ocom >=0 "
   SQLX$ = SQLX$ + " AND Npro_Ocom = " & NPRO% & " "
   SQLX$ = SQLX$ + " AND Nume_Ocom <> " & NOCOA& & " "
   'Set CARCOANTMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim CARCOANTMP As ADODB.Recordset
   Set CARCOANTMP = New ADODB.Recordset
   CARCOANTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   With CARCOANTMP
     On Error Resume Next
     .MoveFirst
     If Err Then
       On Error GoTo 0
       Exit Sub
     End If
     Do While Not .EOF
       NUORSE& = !N_OrServ
             If NUORSE& < 1 Then
         COEFCAN = !Coef_Unids
         If COEFCAN < 0.005 Then COEFCAN = 1
         CACOM = !Cant_Ocom * COEFCAN
         CAREC = !Cant_Rec * COEFCAN
         CAPEN = CACOM - CAREC
         If CAPEN / COEFCAN >= 0.05 Then
           CIXI% = !cod_inte
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
           Call REPLA(LI2TEX$, CODEXT$(CIXI%), 1, 16)
           Call REPLA(LI2TEX$, DESCRIT0$(CIXI%), 17, 33)
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
       End If
       .MoveNext
     Loop
   End With

   CARCOANTMP.Close
   Set CARCOANTMP = Nothing
   Screen.MousePointer = 1

'   esto ya estaba comentado'
'
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

   HOII% = HOY(HOS$)
   FENTREGEN% = FECHANUM(Text14)
   List1.Clear: List1.ToolTipText = ""
   FIN& = ULTREG&("!OCOMDETA")
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   For U& = 1 To FIN&
     X$ = REGLEIDO$("!OCOMDETA", U&)
     Ci% = CVI(Mid$(X$, 9, 2))
     FENTRESOL% = CVI(Mid$(X$, 95, 2))
     If FENTRESOL% < FENTREGEN% Then FENTRESOL% = FENTREGEN%
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
     If InStr(EPAC$, "CONTEPLAST") > 0 Or InStr(EPAC$, "CAMBACERES") > 0 Or InStr(EPAC$, "QUILMES") > 0 Then
         TMT% = TIMATE%(Ci%)
         UME$ = TRIM$(UCase$(Unimed$(Ci%)))
         If TMT% = 41 And UME$ = "M." Then ' tipomat es bobinas Y Unimed$ ES M.
            StringNroCtes$ = StringDeNroProveedoresComprasInternas$
            'SI EL NUMERO DE CLIENTE ESTA DENTRO DE LA CADENA DE NRO DE PROVEEDORES CONFIGURADOS COMO PARTICIPANTES DE LA COMPRA INTERNA.
            If NumeroEstaEnCadena(StringNroCtes$, XVALO(LText1)) Then
               Peso# = PESUNI(Ci%)
               If Peso# > 0.00005 Then
                  COEFI = 1 / Peso#
                  'CANCOMP = CANSTOK / coefi 'SI SON BOBINAS
                  CANSTOK = ROUND(CANCOMP * Peso#)
               End If
            End If
         ElseIf TMT% = 41 And UME$ = "U." Then 'ESTOY DENTRO DE CONTEP,CAMBA,QUILMES
              Peso# = PESUNI(Ci%)
              If Peso# > 0.00005 Then
                  COEFI = 1 * Peso#
                  'CANCOMP = CANSTOK / coefi 'SI SON BOBINAS
                  CANSTOK = ROUND(CANCOMP / Peso#)
              End If
         End If
      End If

     Call REPLA(X$, MKS$(CANSTOK), 51, 4)
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
     '\\\OT-05-0690-WAR-24/10/05///
     Call REPLA(TTXX$, DESCRIT0$(Ci%), 17, 15)
     If Ci% <= 0 Then
       DESCRILI$ = TRIM$(REGLEIDO$("!1OCOMDET", U&))
       Call REPLA(TTXX$, DESCRILI$, 17, 15)
     End If
     '\\\OT-05-0690-WAR-FIN///
     Call REPLA(TTXX$, UNISTO$, 34, 4)
     Call REPLA(TTXX$, FORMATNUM$(CANSTO, "F8.1"), 38, 8)
     Call REPLA(TTXX$, UNICOX$, 48, 4)
     Call REPLA(TTXX$, FORMATNUM$(CANCOM, "F9.1"), 52, 9)
     MASCA$ = "F10.4"
     If PREUNI > 9999 Then MASCA$ = "F10.2"
     Call REPLA(TTXX$, FORMATNUM$(PREUNI, MASCA$), 61, 10)
     Call REPLA(TTXX$, FORMATNUM$(PORDESCUIN, "F5.1"), 72, 5)
     Call REPLA(TTXX$, FORMATNUM$(FENTRESOL%, "D8"), 78, 8)
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
   Screen.MousePointer = 1
End Sub

Sub SHOWOCOM(NPX&)
    '
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

''
''
''
''
''
''
''    NUDOCU$ = TRIM$(Str$(NPX&))
''    While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
''    NUDOCU1$ = " " + NUDOCU$ + " "
''    NUDOCU2$ = " " + NUDOCU$ + "-"
''    '
''    Screen.MousePointer = 11
''    '
''    For UI& = ULTREG&("PDOCLST") To 1 Step -1
''      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
''      If InStr(XX$, TIDOCUM$) > 4 Then
''        PPXX% = InStr(XX$, "Nro.")
''        If PPXX% > 0 Then
''          XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
''        End If
''        If InStr(XX$, NUDOCU1$) > 4 Or InStr(XX$, NUDOCU2$) > 4 Then
''          DOCUNU& = CVS(Left$(XX$, 4))
''          GoTo 5
''        End If
''      End If
''    Next UI&
''    Screen.MousePointer = 1
''    Call COMUNI("Documento no Encontrado")
''    Exit Sub
''    '
''    NoPuedeImprimir% = 0
''    MENSAJEDOC$ = ""
''    CONDI$ = "NUME_OCOM = " & NPX&
''    ST% = XVALO(VALOBADA("OCOMENCA", "STAT_OCOM", CONDI$, "NOMESS"))
''    If ST% < 0 Then
''        NoPuedeImprimir% = 1
''        MENSAJEDOC$ = "Imposible Imprimir Esta Orden de Compra, La Misma Está Pendiente de Aprobación"
''    End If
''    Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, SAFETEXT$(NPX&), , , MENSAJEDOC$, NoPuedeImprimir%)
''    NoPuedeImprimir% = 0
''    MENSAJEDOC$ = ""
''
''5   Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, , SAFETEXT$(NPX&), , MENSAJEDOC$, NoPuedeImprimir%)
    '
End Sub

Sub PRICOMPRA()
    '
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
    FORIPRE$ = Control$("", "FORMOCOM")
    '
    If TRIM$(FORIPRE$) <> "" Then
      '
      Screen.MousePointer = 11
      FECHEMII% = FECHANUM(Text9)
      FECHDOC$ = FETEXTO$(FECHEMII%)
      NUMEDOC$ = TRIM$(Text11)
      While Len(NUMEDOC$) < 6
        NUMEDOC$ = "0" + NUMEDOC$
      Wend
      '\\\OT-06-0082-WAR-06/03/06///
      ' NUMEDOC$ = NUMEDOC$ + "-" + TRIM$(Label16)
      '\\\OT-06-0082-WAR-FIN///
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
      '\\\OT-08-0969-OD-24/10/08///
      CODPARAM$(22) = "REF-MAT2"
      '\\\OT-08-0969-OD-FIN///
      CODPARAM$(26) = "ACH-CUCO"
      CODPARAM$(27) = "PES-CONJ"
      CODPARAM$(28) = "LAR-CUTE"
      CODPARAM$(29) = "LAR-MATE"

      CAPARDOC% = 29
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
      If FECHENT% <= HOII% Then
          DOCPARAM$(15) = DOCPARAM$(15) + " (inmediato)"
        ElseIf FECHENT% < HOII% + 4 Then
          DOCPARAM$(15) = DOCPARAM$(15) + " (urgente!!)"
      End If
      DOCPARAM$(16) = TRIM$(Label11)
      DOCPARAM$(17) = ""
      If MONEMI% > 1 Then DOCPARAM$(17) = Str$(TICAMONE(MONEMI%))
      DOCPARAM$(18) = Eco21(1)
      DOCPARAM$(19) = Text10
      DOCPARAM$(20) = Total(4)
      '\\\OT-05-0676-WAR-20/10/05///
      DOCPARAM$(21) = RAI$ 'REGLAMENTO COMPRA ANEXO
      '\\\OT-05-0676-WAR-FIN///
       '\\\OT-08-0969-OD-24/10/08/// el label decia label5 se modifico a label 15
      'que es la calificacion del proveedor
      DOCPARAM$(22) = Label18.Caption
      DOCPARAM$(26) = TRIM$(Text1)
      DOCPARAM$(27) = TXTPESO
      'AGREGADO PARA ARCON SOLO PARA LAS SUCURSALES QUE IMPRIMA EL CLIENTE Y LA O.C. DEL MISMO
      DOCPARAM$(28) = Left$(TRIM$(TEXT24) & " - " & TRIM$(Eco24), 45)
      DOCPARAM$(29) = Left$(TEXT25, 16)
      '\\\OT-08-0969-OD-fin///

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
      '
      '\\\OT-08-0969-OD-FIN///
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
      CODTABU1$(20) = "ACH-MACO"
      CODTABU1$(21) = "REF-MAT5"
      CODTABU1$(22) = "REF-MAT6"
      CACOLTAB1% = 22
      '
      DESTIDOC% = (-1) * XVALO(LText1)
      CAITAB1% = 0
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
      For J& = 1 To ULTREG&("!OCOMDETA")
         X$ = REGLEIDO$("!OCOMDETA", J&)
         CI0% = CVI(Mid$(X$, 9, 2))
         CANTU = CVS(Mid$(X$, 55, 4))
         If CANTU < 0.0005 Then GoTo 109 'NEXT
         PreUnid# = CVD(Mid$(X$, 67, 8))
         PORDE = CVS(Mid$(X$, 75, 4))
         FEX = CVI(Mid$(X$, 95, 2))
         CONSIMAT% = 0: If TRIM$(Mid$(X$, 97, 2)) <> "" Then CONSIMAT% = 1
         '
         DIMPRE$ = ""
         CAITAB1% = CAITAB1% + 1
         CODDX$ = TRIM$(CODEXT$(CI0%))
         CPROP$ = CODDX$
         If REPLAORI% = 1 Then
            CODDX$ = TRIM$(CODORIG$(CI0%, NC1%))
         End If
'
'         If REVICODX% = 1 Then
'            FEREVX = FEREVI%(CI0%)
'            RVCDDX$ = TRIM$(REVICOD$(CI0%))
'            If RVCDDX$ <> "" Then
'              CODDX$ = CODDX$ + " (Letra: " + RVCDDX$ + " - " + FECHATEX$(FEREVX) + ")"
'            End If
'         End If

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
               FECHYREVI$ = TRIM$(RVCDDX$)
               SOLOREVI$ = FECHYREVI$
               If FEREVX > 33 Then FECHYREVI$ = FECHYREVI$ & "-(" & FECHATEX$(FEREVX) & ")"
            End If
         End If
         TETABU1$(1, CAITAB1%) = CODDX$
         'TRES DESCRIPCIONES POSIBLES
         DESPRIN$ = DESCRIT0$(CI0%)
         DESCOMPLEM$ = DESCRIT2$(CI0%)
         ESPECIFICAC$ = ESPECIFIT$(CI0%)
         
         If REPLAORI% = 1 Then
             If CODDX$ <> CPROP$ Then
               If RECOPROP% > 0 Then
                 DESPRIN$ = DESPRIN$ + " (" + CPROP$ + ")"
               End If
             End If
         End If
           '
         TETABU1$(3, CAITAB1%) = Mid$(X$, 59, 4) 'UNIDAD DE COMPRA
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
           PORNA = PORNACIO(CI0%)
         TETABU1$(20, CAITAB1%) = TRIM$(FORMATNUM$(PORNA, "F7.2")) 'PORCENTAJE DE NACIONALIZACION
         TETABU1$(21, CAITAB1%) = TRIM$(FECHYREVI$)
         TETABU1$(22, CAITAB1%) = SOLOREVI$ ' SOLO EL ID DE REVISION DEL ARTICULO
         '
         POCODI% = XVALO(ATRIFORM$(FORIPRE$, "COD-MAT/POSINHOR"))
         PODESI% = XVALO(ATRIFORM$(FORIPRE$, "DES-MAT/POSINHOR"))
         If PODESI% <= POCODI% + 3 Then
           If TRIM$(CODDX$) <> "" Then
             CAITAB1% = CAITAB1% + 1
           End If
         End If
         '
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
         '
         UNIST$ = Mid$(X$, 11, 4)
         UNICO$ = UCase$(TRIM$(Mid$(X$, 59, 4)))
         DEADI$ = ""
         If TRIM$(Left$(UCase$(UNIST$), 2)) <> UNICO$ Then
           If CI0% > 0 Then  ' que no lo haga en texto libre
             DEADI$ = "(equiv. " + TRIM$(FORMATNUM$(CVS(Mid$(X$, 51, 4)), "F12.1")) + " " + TRIM$(ECOARCH$("UNIMED", UNIST$)) + ".)"
           End If
         End If
         '
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

'''''         If INCLUSPE > O Then
''''''           INTERLI = XVALO(ATRIFORM$(FORIPRE$, "INTERLIN"))
'''''           XX$ = ESPECIFIT$(CI0%)
'''''           If TRIM$(XX$) <> "" Then
'''''              If DIMPRE$ <> "" Then DIMPRE$ = DIMPRE$ + vbCrLf
'''''              DIMPRE$ = DIMPRE$ & XX$
'''''
''''''              COLUTA% = 0
''''''              If DESCPRI% = 1 Or DESCOMP% = 1 Then CAITAB1% = CAITAB1% + 1
''''''              Call FRATEXTO(XX$, ATELIB%)
''''''              For JJ& = 1 To CARETEX%
''''''                If INTERLI <= 6 Or InStr(UCase$(FORIPRE$), "MZZ") > 0 Then
''''''                  TETABU1$(2, CAITAB1%) = RETEX$(JJ&)
''''''                  CAITAB1% = CAITAB1% + 1
''''''                Else
''''''                  COLUTA% = COLUTA% + 1
''''''                  If COLUTA% > 2 Then COLUTA% = 1
''''''                  If COLUTA% = 1 Then CAITAB1% = CAITAB1% + 1
''''''                  TETABU1$(COLUTA%, CAITAB1%) = RETEX$(JJ&)
''''''                End If
''''''              Next JJ&
''''''              CAITAB1% = CAITAB1% - 1
'''''           End If
'''''         End If
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
         Call CARLINSEP  'Carga linea de separacion entre codigo y codigo
         '
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
         CANTU = CVS(Mid$(X$, 55, 4))
         If CI0% > 0 And CANTU > 0 Then
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
           NPRO% = XVALO(LText1)
           'Call ENTRECONSIACC(NPRO%)
           Call NUENTRECONSI(NPRO%)
         End If
      End If
    End If
    Exit Sub
    '
999 Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
End Sub

Private Sub UcoMat_Click()
Call OPLISCO07.Command20_Click
End Sub

Private Sub UcoPro_Click()
Call OPLISCO07.Command12_Click
End Sub

Private Sub Ver_Ocom_Click()
    Call OPLISCO07.Command22_Click
End Sub

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
      'Text1.TEXT = ""
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   '
End Sub

Function CANTORECI(CIXI%, NUOCO&)
    '
    CANTO = 0
    '
    Call APERBASDAT("STOCKS")
    SQLX$ = "SELECT * FROM BOLRECEP "
    SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIXI% & " "
    SQLX$ = SQLX$ + "AND NumOC = " & NUOCO& & " "
    'Set BOLDET = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    Dim BOLDET As ADODB.Recordset
    Set BOLDET = FLEXCONN.Execute(FILTSQL$(SQLX$))
    
    With BOLDET
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        GoTo 126
      End If
      On Error GoTo 0
      Do While Not .EOF
        CANTO = CANTO + XVALO(!CantNom) - XVALO(!CanToRech) - XVALO(!CanToFalta)
      .MoveNext
      Loop
    End With
    '
126 CANTORECI = CANTO
    '
End Function

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
10 RAI$ = "-" + Label13 + "-"
   With FORSELREG
     .LIMATSEL.Clear
     FIN& = ULTREG&("INDIRCOM")
     For J& = 1 To FIN&
       X$ = REGLEIDO$("INDIRCOM", J&)
       NR& = CVI(Mid$(X$, 1, 2))
       DRC$ = Mid$(X$, 3, 46)
       RC$ = FORMATNUM(NR&, "F5.0") & Chr(9) & DRC$
       .LIMATSEL.AddItem RC$
       NR1$ = "-" + TRIM$(Str$(NR&)) + "-"
       If InStr(RAI$, NR1$) > 0 Then
         .LIMATSEL.Selected(J& - 1) = True
       End If
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

Sub BORRACONSIG(NUOC&)
    '
    NOCOM$ = TRIM$(Str$(NUOC&))
    While Len(NOCOM$) < 6: NOCOM$ = "0" + NOCOM$: Wend
    NOCOM$ = "OC-" + Left$(NOCOM$, 6)
    CONDI$ = "NRO_OCOM LIKE '" & NOCOM$ & "%' and Cant_Entreg = 0"
    'Elimino del MACONSIG todas las consignaciones de la orden de compra que no tienen entregas
    Call BORRAREGISTROS("MACONSIG", CONDI$, REGS&, "NOMESS")
    '
End Sub
