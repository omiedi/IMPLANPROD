VERSION 5.00
Begin VB.Form MODSIMAT 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modificación de Pedidos de Clientes - Formato Lineal"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3165
      Left            =   0
      TabIndex        =   35
      Top             =   5145
      Width           =   10800
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   1275
         Left            =   5355
         TabIndex        =   82
         Top             =   1785
         Visible         =   0   'False
         Width           =   1380
         Begin VB.TextBox TOTIVA 
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
            Index           =   0
            Left            =   0
            TabIndex        =   85
            Text            =   " "
            Top             =   0
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.TextBox TOTIVA 
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
            Index           =   1
            Left            =   0
            TabIndex        =   84
            Text            =   " "
            Top             =   315
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.TextBox TOTIVA 
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
            Index           =   2
            Left            =   0
            TabIndex        =   83
            Text            =   " "
            Top             =   630
            Visible         =   0   'False
            Width           =   870
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFC0&
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
         ForeColor       =   &H00808000&
         Height          =   3120
         Left            =   5670
         TabIndex        =   59
         Top             =   0
         Width           =   5055
         Begin VB.TextBox TOTIVA 
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
            Height          =   315
            Index           =   3
            Left            =   3400
            TabIndex        =   92
            Text            =   " "
            Top             =   1680
            Width           =   1500
         End
         Begin VB.TextBox Total 
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
            Height          =   315
            Index           =   3
            Left            =   3400
            TabIndex        =   91
            Text            =   " "
            Top             =   2625
            Width           =   1500
         End
         Begin VB.TextBox Text15 
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
            Height          =   315
            Left            =   3400
            TabIndex        =   90
            Top             =   2100
            Width           =   1485
         End
         Begin VB.TextBox TICAMBIS 
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
            Height          =   300
            Left            =   630
            MaxLength       =   14
            TabIndex        =   86
            Top             =   1125
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.TextBox Text14 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   3400
            TabIndex        =   79
            TabStop         =   0   'False
            Text            =   "    *******"
            Top             =   1125
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox Text14 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   3400
            TabIndex        =   78
            TabStop         =   0   'False
            Text            =   "    *******"
            Top             =   720
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox Text14 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   3400
            TabIndex        =   77
            TabStop         =   0   'False
            Text            =   "    *******"
            Top             =   315
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox Total 
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
            Index           =   0
            Left            =   3405
            TabIndex        =   62
            Text            =   " "
            Top             =   315
            Width           =   1500
         End
         Begin VB.TextBox Total 
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
            Index           =   1
            Left            =   3405
            TabIndex        =   61
            TabStop         =   0   'False
            Text            =   " "
            Top             =   720
            Width           =   1500
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
            Height          =   300
            Left            =   1890
            MaxLength       =   14
            TabIndex        =   22
            Top             =   735
            Width           =   1065
         End
         Begin VB.TextBox Total 
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
            Index           =   2
            Left            =   3405
            TabIndex        =   60
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1125
            Width           =   1500
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Bruto Pedido:"
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
            Index           =   7
            Left            =   1365
            TabIndex        =   95
            Top             =   2730
            Width           =   1905
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Percepción I.Brutos Bs.As.:"
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
            Index           =   6
            Left            =   210
            TabIndex        =   94
            Top             =   2205
            Width           =   3060
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Impuestos (I.V.A.):"
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
            Left            =   1365
            TabIndex        =   93
            Top             =   1785
            Width           =   1905
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   210
            Picture         =   "MODSIMAT.frx":0000
            Top             =   315
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label Label12 
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
            Left            =   165
            TabIndex        =   89
            Top             =   1200
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
            Left            =   315
            TabIndex        =   88
            ToolTipText     =   "Doble-Click para 'Dólares'"
            Top             =   735
            Visible         =   0   'False
            Width           =   1380
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Descuento:"
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
            Index           =   4
            Left            =   315
            TabIndex        =   87
            Top             =   495
            Width           =   1905
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Neto:"
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
            Left            =   1365
            TabIndex        =   65
            Top             =   390
            Width           =   1905
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Sub-Total Neto:"
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
            Left            =   1365
            TabIndex        =   64
            Top             =   1200
            Width           =   1905
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "%:"
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
            Left            =   3045
            TabIndex        =   63
            Top             =   825
            Width           =   225
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFC0&
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
         ForeColor       =   &H00808000&
         Height          =   3120
         Left            =   210
         TabIndex        =   51
         Top             =   0
         Width           =   5330
         Begin VB.TextBox Text44 
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
            Left            =   4620
            TabIndex        =   75
            Top             =   1230
            Width           =   540
         End
         Begin VB.TextBox Text13 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1170
            Left            =   1155
            MaxLength       =   96
            MultiLine       =   -1  'True
            TabIndex        =   18
            Top             =   1785
            Width           =   4005
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
            Index           =   2
            Left            =   1890
            TabIndex        =   54
            Text            =   " "
            Top             =   1230
            Width           =   1800
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
            TabIndex        =   16
            Text            =   " "
            Top             =   1230
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
            Index           =   2
            Left            =   1575
            TabIndex        =   17
            Top             =   1230
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
            TabIndex        =   53
            Text            =   " "
            Top             =   825
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
            TabIndex        =   14
            Text            =   " "
            Top             =   825
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
            Index           =   1
            Left            =   1575
            TabIndex        =   15
            Top             =   825
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
            TabIndex        =   52
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
            TabIndex        =   12
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
            TabIndex        =   13
            Top             =   420
            Width           =   180
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
            Left            =   210
            TabIndex        =   97
            Top             =   2415
            Width           =   645
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
            TabIndex        =   96
            Top             =   2670
            Width           =   540
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "% Comis:"
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
            Left            =   3675
            TabIndex        =   76
            Top             =   1305
            Width           =   855
         End
         Begin VB.Label Label21 
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
            Height          =   330
            Index           =   3
            Left            =   -3045
            TabIndex        =   58
            Top             =   1785
            Width           =   4110
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
            TabIndex        =   57
            Top             =   1305
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
            TabIndex        =   56
            Top             =   900
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
            TabIndex        =   55
            Top             =   495
            Width           =   1590
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFC0&
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
         Height          =   1550
         Left            =   5670
         TabIndex        =   40
         Top             =   1575
         Width           =   5055
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
            TabIndex        =   25
            Top             =   1125
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
            Index           =   1
            Left            =   1155
            MaxLength       =   32
            TabIndex        =   24
            Top             =   720
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
            TabIndex        =   23
            Top             =   315
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
            Index           =   3
            Left            =   1155
            MaxLength       =   32
            TabIndex        =   71
            Top             =   1125
            Visible         =   0   'False
            Width           =   3750
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
            Index           =   5
            Left            =   -840
            TabIndex        =   50
            Top             =   1210
            Width           =   1905
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Domicilio:"
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
            Left            =   -840
            TabIndex        =   41
            Top             =   400
            Width           =   1905
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8520
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   12000
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
         TabIndex        =   73
         Top             =   1260
         Width           =   180
      End
      Begin VB.PictureBox MARCONOTA 
         BackColor       =   &H00FFFFC0&
         Height          =   3180
         Left            =   210
         ScaleHeight     =   3120
         ScaleWidth      =   10455
         TabIndex        =   66
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
            Picture         =   "MODSIMAT.frx":030A
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
            Picture         =   "MODSIMAT.frx":0974
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Cancela Pedido de Cliente"
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
            Picture         =   "MODSIMAT.frx":0C7E
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Continúa Generación del Pedido de Cliente"
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
            TabIndex        =   69
            Top             =   390
            Width           =   9780
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FFFFC0&
            Height          =   435
            Left            =   -30
            ScaleHeight     =   375
            ScaleWidth      =   9720
            TabIndex        =   67
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
               TabIndex        =   68
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
         TabIndex        =   70
         Top             =   2205
         Width           =   10515
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
         TabIndex        =   42
         Top             =   105
         Width           =   850
         Begin VB.CommandButton Command22 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Notas"
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
            Picture         =   "MODSIMAT.frx":0F88
            Style           =   1  'Graphical
            TabIndex        =   81
            ToolTipText     =   "Anotador de Pedidos"
            Top             =   3860
            Visible         =   0   'False
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
            Picture         =   "MODSIMAT.frx":1292
            Style           =   1  'Graphical
            TabIndex        =   28
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
            Picture         =   "MODSIMAT.frx":13DC
            Style           =   1  'Graphical
            TabIndex        =   49
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
            Picture         =   "MODSIMAT.frx":16E6
            Style           =   1  'Graphical
            TabIndex        =   48
            ToolTipText     =   "Listas de Precios de Venta"
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
            Picture         =   "MODSIMAT.frx":1B28
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Acceso a Condiciones Comerciales"
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
            Picture         =   "MODSIMAT.frx":1E32
            Style           =   1  'Graphical
            TabIndex        =   47
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
            Picture         =   "MODSIMAT.frx":213C
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Acceso a Base de Datos de Clientes"
            Top             =   2315
            Width           =   855
         End
         Begin VB.CommandButton Command13 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Direct"
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
            Picture         =   "MODSIMAT.frx":257E
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Accesos Directos"
            Top             =   4642
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
            Picture         =   "MODSIMAT.frx":2888
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Ayuda FLEXOFT en Línea"
            Top             =   770
            Width           =   855
         End
         Begin VB.CommandButton Command12 
            BackColor       =   &H00E0E0E0&
            Caption         =   "View"
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
            Picture         =   "MODSIMAT.frx":2B92
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Consultas y Re-Impresión"
            Top             =   3860
            Width           =   855
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
         Picture         =   "MODSIMAT.frx":2E9C
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Aprueba, Graba e Imprime Pedido"
         Top             =   7300
         Width           =   855
      End
      Begin VB.CommandButton Command10 
         Height          =   260
         Left            =   420
         TabIndex        =   39
         Top             =   6855
         Width           =   2825
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   210
         ScaleHeight     =   375
         ScaleWidth      =   10455
         TabIndex        =   37
         Top             =   1785
         Width           =   10515
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "  Cant Código          Descripción            Fecha   Lista  Desc  P.Unit     Importe"
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
            TabIndex        =   38
            Top             =   105
            Width           =   10515
         End
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
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1280
         Width           =   3690
      End
      Begin VB.TextBox Text11 
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
         TabIndex        =   10
         Text            =   " "
         Top             =   890
         Width           =   1120
      End
      Begin VB.TextBox Text10 
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
         Left            =   8820
         MaxLength       =   16
         TabIndex        =   11
         Top             =   1280
         Width           =   1920
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
         TabIndex        =   8
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
         TabIndex        =   9
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
         Left            =   5565
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   " "
         Top             =   490
         Width           =   1695
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
         Left            =   5565
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1280
         Width           =   1695
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
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   " "
         Top             =   890
         Width           =   2850
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
         TabIndex        =   3
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
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   " "
         Top             =   490
         Width           =   3690
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
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   " "
         Top             =   105
         Width           =   3690
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
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
         Top             =   1260
         Width           =   855
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Venta"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   9030
         TabIndex        =   80
         ToolTipText     =   "Doble-Click para 'Reparación'"
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label9 
         Height          =   225
         Left            =   7455
         TabIndex        =   74
         Top             =   0
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "OV"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   855
         Left            =   150
         TabIndex        =   72
         Top             =   180
         Width           =   1380
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00C0C000&
         FillColor       =   &H0080FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   420
         Shape           =   4  'Rounded Rectangle
         Top             =   465
         Width           =   540
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H0000C000&
         FillColor       =   &H00FFFF80&
         FillStyle       =   0  'Solid
         Height          =   390
         Left            =   315
         Shape           =   4  'Rounded Rectangle
         Top             =   390
         Width           =   750
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000C0C0&
         FillColor       =   &H00C0E0FF&
         FillStyle       =   0  'Solid
         Height          =   540
         Left            =   210
         Shape           =   4  'Rounded Rectangle
         Top             =   315
         Width           =   960
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
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   5565
         TabIndex        =   36
         Top             =   975
         Width           =   1695
      End
      Begin VB.Line Line8 
         X1              =   10730
         X2              =   10730
         Y1              =   1785
         Y2              =   4950
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pedido Nro:"
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
         Height          =   330
         Left            =   7770
         TabIndex        =   34
         Top             =   945
         Width           =   1800
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "O/C o N/P:"
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
         Height          =   330
         Left            =   7035
         TabIndex        =   33
         Top             =   1340
         Width           =   1695
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
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   8820
         TabIndex        =   32
         Top             =   150
         Width           =   750
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Límite Crédito:"
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
         Height          =   330
         Left            =   5565
         TabIndex        =   31
         Top             =   200
         Width           =   1695
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
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   315
         TabIndex        =   30
         Top             =   945
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
Attribute VB_Name = "MODSIMAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHIPRE$

Private Sub Command1_Click()
   If List1.ListCount < 1 Then Exit Sub
   Call SELECHO("DEUDOR1")
   NCLIE% = Val(TRIM$(VALACT1$("DEUDOR1")))
   If NCLIE% > 0 Then
     If REGICLI&(NCLIE%) > 0 Then
       Text1.TEXT = TRIM$(Str$(NCLIE%))
       Call CARDACLI
     End If
   End If
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call CIERRARCH("*.*")
   ACONEC$ = "AMACLI01"
   If EXISTE%("AMACLI02.EXE") > 0 Then ACONEC$ = "AMACLI02"
   Call CONECRUN(ACONEC$, "Padron Maestro de Clientes")
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
    '
    Command12.Enabled = False
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
    '    'YY$ = MKS$(NPX&) + ANULX$ + TRIM$(RASOCLI$(NCLIE%)) + " - O.C." + TRIM$(Mid$(XX$, 57, 16)) + " - " + FECHATEX$(FEX)
    '    YY$ = MKS$(NPX&) + ANULX$ + TRIM$(Mid$(XX$, 5, 48))
    '    JJ& = JJ& + 1
    '    Call GRAREG("INDIPEDI", YY$, JJ&)
    '  End If
    'Next U&
    'Call SETULTREG("INDIPEDI", JJ&)
    'Screen.MousePointer = 1
    '
'3   Call SELECHO("INDIPEDI/Indice General de Pedidos (Pendientes-Cumplidos-Anulados)")
'    NPX& = Val(VALACT1$("INDIPEDI"))
3   Call SELECHO("PEDENCA/Indice General de Pedidos (Pendientes-Cumplidos-Anulados)")
    NPX& = Val(VALACT1$("PEDENCA"))
    If NPX& < 1 Then GoTo 99
    '
    TIDOCUM$ = "Pedido de Cliente"
    NUDOCU$ = TRIM$(Str$(NPX&))
    While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
    NUDOCU$ = " " + NUDOCU$ + " "
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
    '
    Screen.MousePointer = 1
    Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a este Pedido.\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
    GoTo 99
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
    GoTo 99
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
    UI& = UI& + 1
    XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
    If InStr(XX$, TIDOCUM$) > 4 Then
      If InStr(XX$, NUDOCU$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 7
      End If
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
99  Command12.Enabled = True
    '
End Sub

Private Sub Command13_Click()
   Command3.Enabled = False
   Call ACCDIR("PEDIDO02")
   Command3.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "DOSIVA") > 0 Then
     If TRIM$(TEXT21(2)) = "" Or TRIM$(Text44) = "" Then
       Call MENSERR(24, "Falta Vendedor / Comisión")
       GoTo 99
     End If
   End If
   '
   NROPE& = Val(Text11.TEXT)
   RFF$ = RECFILT$("PEDDETA", 1, MKS$(NROPE&))
   For U& = 1 To Len(RFF$) Step 4
     REDE& = CVS(Mid$(RFF$, U&, 4))
     XX$ = REGLEIDO$("PEDDETA", REDE&)
     If Asc(Mid$(XX$, 27, 1)) <> 9 Then ' no esta anulado
       CADESPACH = CVS(Mid$(XX$, 95, 4))
       If CADESPACH >= 0.05 Then
         Call MENSERR(24, "Imposible Modificar este Pedido.\  \Ya hay Entregas Parciales por Remito.")
         GoTo 99
       End If
     End If
   Next U&
   '
   Call GRAOBRA(OKX%)
   If OKX% > 0 Then
     Call PRIPEDIDO
     Call GENERODES      ' genera orden de despacho
     Call BLANFORMU
     Call SETULTREG("!CARDETPE", 0)
     Call CIERRARCH("*.*")
     Hide
   End If
99 Command14.Enabled = True
End Sub

Private Sub COMMAND15_CLICK()
   MARCONOTA.Visible = False
   Text1.TEXT = ""
   Call BLANFORMU
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command16_Click()
   Call MENSERR(24, "No Disponible por el Momento")
End Sub

Private Sub Command16_LostFocus()
   On Error Resume Next
   Command6.SetFocus
   On Error GoTo 0
End Sub

Private Sub Command2_Click()
    Call SETULTREG("!CARDETPE", 0)
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command22_Click()
    Command22.Enabled = False
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
    CONPEDID.Show 1
    '
    Command22.Enabled = True
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call HELPONLINE("PEDIDO01")
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Call MENSERR(24, "Imposible Acceder al Resumen de Cuenta Corriente\Desde esta Pantalla.")
    Exit Sub
    'Command4.Enabled = False
    'RESCUDEU.Text1.TEXT = Text1.TEXT
    'PROLLAMA$ = "MODPEDCLI"
    'Hide
    'RESCUDEU.Show
    'Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    Call CIERRARCH("*.*")
    If (USNBR% Mod 100) <> 1 Then
      Call COMUNI("La Opción Elegida Requiere\Privilegios de Supervisor.")
      Exit Sub
    End If
    OPORDESPA.Show 1
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   MARCONOTA.Visible = False
   'Call CARDETPEDI(1)
End Sub
'
Private Sub Command7_Click()
    Command7.Enabled = False
    Call CIERRARCH("*.*")
    MODX% = ALTERNA%("Listas de Precios\Condiciones de Pago")
    If MODX% = 1 Then
         Call CONECRUN("ADMIPRE", "Listas de Precios de Venta")
       ElseIf MODX% = 2 Then
         Call CARCONPAGOS
    End If
    Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    Call CIERRARCH("*.*")
    ACONEC$ = "AMASTO01"
    If EXISTE%("AMASTO02.EXE") > 0 Then ACONEC$ = "AMASTO02"
    Call CONECRUN(ACONEC$, "Maestro de Artículos de Stock")
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
End Sub

Private Sub Eco21_GotFocus(Index As Integer)
   On Error Resume Next
   TEXT21(Index).SetFocus
   On Error GoTo 0
End Sub

Private Sub Form_Load()
    '
    Call CENTRAFORM(Me)
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Frame4.Height = 1545
    '
    If InStr(UCase$(EPAC$), "DOSIVA") > 0 Then
      Label11.Visible = True
      Command22.Visible = True
      Command12.Visible = False
    End If
    '
    If EXISTE%("REPARA02.EXE") > 0 Then
      Label11.Visible = True
    End If
    '
    MONEMI% = 1  ' PESOS
    CONVERPRECIO% = 0
    If InStr(EPAC$, "DOSIVA") > 0 Then
       MONEMIS.Visible = True
       CONVERPRECIO% = 1
       TICAMBIS.Visible = True
       Label12.Visible = True
       Image1.Visible = True
    End If
    TICAMBIS = FORMATNUM$(TICAMONE(2), "F9.4")
    '
    CMO$ = TRIM$(CONTROL$("", "CODPECLI"))
    If CMO$ <> "" Then Label10.Caption = CMO$
    '
    ECOPRECIO% = 0
    If CONTROL$("ADMIPED", "MUEPREC") <> "SI" Then
      For KKI% = 0 To 2      ' TAPA IMPORTES TOTALES
        Text14(KKI%).Visible = True
      Next KKI%
      ECOPRECIO% = (-1)
    End If
    '
    If ECOPRECIO% < 0 Then
      Label22.Caption = "  Cant Código          Descripción                                           Entrega "
      Image1.Visible = False
    End If
    '
End Sub
'

Private Sub Image1_DblClick()
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
         Call REPLA(RECORD$, MKI$(HORANUM%(Time$)), 63, 2)
         Call FILTERPUTRECORD("DAPIBCLI", 1, MKI$(NCLIE%), RECORD$)
         Call CIERRARCH("DAPIBCLI")
         '
         Call CALPERIB
         '
       End If
     End If
   End If
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If ECOPRECIO >= 0 Then
     Frame4.Height = 3120
   End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Frame4.Height = 1545
End Sub

Private Sub Label11_DblClick()
   If TRIM$(UCase$(Label11)) = "VENTA" Then
       Label11.ForeColor = RGB(255, 0, 0)
       Label11 = "Reparación"
       Label11.ToolTipText = "Doble-Click para 'Venta'"
     Else
       Label11.ForeColor = RGB(0, 0, 255)
       Label11 = "Venta"
       Label11.ToolTipText = "Doble-Click para 'Reparación'"
   End If
End Sub

Private Sub Label8_Click()
   On Error Resume Next
   Text13.SetFocus
   On Error GoTo 0
End Sub

Private Sub LENTREGA_GotFocus(Index As Integer)
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   LENTREGA(Index).TEXT = TRIM$(LENTREGA(Index).TEXT)
   LENTREGA(Index).SelStart = 0
   LENTREGA(Index).SelLength = Len(LENTREGA(Index).TEXT)
End Sub

Private Sub List1_DblClick()
   'Call CARDETPEDI(2)
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
      Else
        MONEMI% = 1
        MONEMIS.ForeColor = RGB(0, 100, 0)
        MONEMIS.ToolTipText = "Doble-Click para 'Dólares'"
        'Label12.Visible = False
        If TICAMBIS.Visible = True Then TICAMBIS.Enabled = False
   End If
   '
   COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
   FIN& = ULTREG&("!CARDETPE")
   For U& = 1 To FIN&
     X$ = REGLEIDO$("!CARDETPE", U&)
     CAN = CVS(Mid$(X$, 47, 4))
     PORDESCUIN = CVS(Mid$(X$, 69, 4))
     PRECARGA = COEFCON * CVS(Mid$(X$, 39, 4))
     PREUNI = PRECARGA * (100 - PORDESCUIN) / 100
     TOT# = CDbl(CAN * PREUNI)
     '
     Call REPLA(X$, MKS$(PRECARGA), 39, 4)
     Call REPLA(X$, MKS$(PRECARGA), 51, 4)
     Call REPLA(X$, MKD$(TOT#), 55, 8)
     Call REPLA(X$, MKS$(PRECARGA), 65, 4)
     Call REPLA(X$, MKS$(PREUNI), 73, 4)
     Call REPLA(X$, MKD$(TOT#), 77, 8)
     Call GRAREG("!CARDETPE", X$, U&)
   Next U&
   '
   Call CARLISPE
   Call CIERRARCH("!CARDETPE")
   Call CALTOPED
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
   Call CALTOPED
End Sub

Private Sub PORDESCU_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
End Sub

Private Sub Text1_Change()
   NCLIE% = Val(Text1.TEXT)
   If NCLIE% > 0 Then
     If REGICLI&(NCLIE%) > 0 Then
       Call CARDACLI
     End If
   End If
End Sub

Private Sub Text1_DblClick()
   If List1.ListCount < 1 Then Exit Sub
   Call SELECHO("DEUDOR1")
   NCLIE% = Val(TRIM$(VALACT1$("DEUDOR1")))
   If NCLIE% > 0 Then
     If REGICLI&(NCLIE%) > 0 Then
       Text1.TEXT = TRIM$(Str$(NCLIE%))
       Call CARDACLI
     End If
   End If
End Sub

Private Sub Text1_GotFocus()
   If List1.ListCount < 1 Then
      On Error Resume Next
      Command2.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
End Sub

Private Sub Text10_GotFocus()
   Text10.TEXT = TRIM$(Text10.TEXT)
   Text10.SelStart = 0
   Text10.SelLength = Len(Text10.TEXT)
End Sub

Private Sub Text11_GotFocus()
   On Error Resume Next
   Command2.SetFocus
   On Error GoTo 0
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

Private Sub Text15_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text2_GotFocus()
   On Error Resume Next
   Text1.SetFocus
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
        ARCHEC$ = "LISTAS": LOCA% = 2
      ElseIf Index = 1 Then
        ARCHEC$ = "CONPAG": LOCA% = 2
      ElseIf Index = 2 Then
        ARCHEC$ = "VENDEDOR": LOCA% = 1
    End If
    ARGU$ = CBIN$(TRIM$(TEXT21(Index).TEXT), 1, LOCA%)
    Call SELECHO(ARCHEC$)
    VDEVU$ = TRIM$(VALACT1$(ARCHEC$))
    If VDEVU$ <> "" Then
       TEXT21(Index).TEXT = VDEVU$
    End If
    '
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

Private Sub Text44_LostFocus()
   If Val(Text44) > 99.9 Then Text42 = "99.9"
   If TRIM$(Text44) <> "" Then Text44 = Format$(Val(Text44), "#0.0")
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
      ElseIf Index = 1 Then
        Eco21(Index).TEXT = ECOARCH$("CONPAG", MKI$(Val(TEXT21(Index).TEXT)))
      ElseIf Index = 2 Then
        Eco21(Index).TEXT = ECOARCH$("VENDEDOR", Chr$(Val(TEXT21(Index).TEXT)))
    End If
    '
End Sub

Private Sub COMMAND21_CLICK(Index As Integer)
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
    VDEVU$ = TRIM$(VALACT1$(ARCHEC$))
    If VDEVU$ <> "" Then
       TEXT21(Index).TEXT = VDEVU$
    End If
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
   Text13.TEXT = ""
   Text15 = ""
   Text6.TEXT = ""
   List1.Clear
   Label18 = ""
   '
   TEXT21(0).TEXT = "": Eco21(0) = ""
   TEXT21(1).TEXT = "": Eco21(1) = ""
   TEXT21(2).TEXT = "": Eco21(2) = ""
   Text44 = ""
   '
   PORDESCU.TEXT = ""
   For KK% = 0 To 3
      Total(KK%).TEXT = ""
      TOTIVA(KK%).TEXT = ""
      LENTREGA(KK%).TEXT = ""
   Next KK%
   '
   Label11.ForeColor = RGB(0, 0, 255)
   Label11 = "Venta"
   Label11.ToolTipText = "Doble-Click para 'Reparación'"
   '
   MARCONOTA.Visible = False
   '
End Sub
'
Sub CARDACLI()
    '
    Screen.MousePointer = 11
    NC% = Val(Text1.TEXT)
    If NC% > 0 Then
      If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
        Text2.TEXT = RASOCLI$(NC%)
        Text3.TEXT = DOMICLI$(NC%)
        Text4.TEXT = CPOSCLI$(NC%)
        Text5.TEXT = LOCACLI$(NC%)
        Text7.TEXT = CSTRING$(MKS$(LIMICRE(NC%)), 3, 13, 2, 2)
        Text12.TEXT = TRIM$(POSIVAC$(NC%)) + " - CUIT " + CUITCLI$(NC%)
        CREDISPO# = LIMICRE(NC%) - SALDCLI#(NC%) - CHEPEACRE#(NC%)
        Text8.ForeColor = RGB(0, 0, 0)
        If CREDISPO# < 0 Then Text8.ForeColor = RGB(255, 0, 0)
        Text8.TEXT = CSTRING$(MKD$(Abs(CREDISPO#)), 3, 13, 2, 2)
      End If
    End If
    Screen.MousePointer = 1
    '
End Sub
   '
Function NUOBRA&()
   '
   NPP& = 1
   '
   FIN& = ULTREG&("PEDENCA")
   If FIN& > 0 Then
      For U& = FIN& To FIN& - 16 Step -1
         If U& < 1 Then GoTo 99
         XX$ = REGLEIDO$("PEDENCA", U&)
         NPPX& = CVS(Mid$(XX$, 1, 4))
         If NPPX& >= NPP& Then
            NPP& = 1 + NPPX&
         End If
      Next U&
   End If
   '
99 NUOBRA& = NPP&
   '
End Function
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
End Sub

Private Sub Timer1_Timer()
   '
   Call CALTOPED
   '
End Sub

Private Sub Total_Change(Index As Integer)
   If Index = 0 Then
     Call CALTOPED
'      DESC$ = PORDESCU.TEXT
'      TTT1 = Val(Total(0).TEXT) * COEFDESCU(DESC$) / 100
'      Total(1).TEXT = CSTRING$(MKS$(TTT1), 3, 11, 2, 2)
'      TTT3 = Val(Total(0).TEXT) - Val(Total(1).TEXT)
'      Total(2).TEXT = CSTRING$(MKS$(TTT3), 3, 11, 2, 2)
   End If
'   If Index < 3 Then
'      TTT1 = Val(Total(2).TEXT)
'      TTT2 = Val(TOTIVA(3).TEXT)
'      Total(3).TEXT = Str$(TTT1 + TTT2)
'   End If
End Sub
'
Sub PRIPEDIDO()
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FORPECLI"))
    If FORIPRE$ = "" Then Exit Sub
    '
    If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
       Call MENSERR(24, "Imposible Imprimir Pedido de Cliente.\No Existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    CANCELPRINT = 0
    If CONTROL$("PEDCLIEN", "SUPRINT") = "SI" Then
       CANCELPRINT% = 1
    End If
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "DOSIVAC") > 0 Then
      CANCELPRINT% = 0
      Screen.MousePointer = 1
      OPRIPEDI.Show 1
      If OPRIPEDI.Check1(0).Value <> 1 Then
        CANCELPRINT% = 1
      End If
    End If
    Screen.MousePointer = 11
    '
    FECHDOC$ = Text9.TEXT
    NUMEDOC$ = TRIM$(Text11.TEXT)
    While Len(NUMEDOC$) < 6
      NUMEDOC$ = "0" + NUMEDOC$
    Wend
    '
    TIPODOC$ = "Pedido de Cliente"
    If NC% < 1 Then
      TIPODOC$ = "Pedido Interno"
    End If
    '
    CODTABU1$(1) = "COD-MAT"
    CODTABU1$(2) = "DES-MAT"
    CODTABU1$(3) = "CANTIDAD"
    CODTABU1$(4) = "PRE-LIST"
    CODTABU1$(5) = "PORDESCU"
    CODTABU1$(6) = "PRE-NETO"
    CODTABU1$(7) = "NET-ITEM"
    CODTABU1$(8) = "FECH-REC"
    CODTABU1$(9) = "ORD-ITEM"
    CACOLTAB1% = 9
    '
    DESTIDOC% = NC%
    CAITAB1% = 0
    CAITCRU% = 0
    '
    For U& = 1 To ULTREG&("!CARDETPE")
      XX$ = REGLEIDO$("!CARDETPE", U&)
      CIPRE% = CVI(Left$(XX$, 2))
      CAPEX$ = Str$(CVS(Mid$(XX$, 47, 4)))
      '
      If CIPRE% > 0 Then
        If CIPRE% <= NUMAS% Then
          If Val(CAPEX$) > 0 Then
            '
            PRELIX = PRELISTA#(ARCHIPRE$, CIPRE%)
            PRECARGA = CVS(Mid$(XX$, 39, 4))
            PREUNI = CVS(Mid$(XX$, 73, 4))
            If PRELIX < PRECARGA Then PRELIX = PRECARGA
            If PRELIX < PREUNI Then PRELIX = PREUNI
            '
            PRELI$ = Str$(PRELIX)
            PREUX$ = Str$(CVS(Mid$(XX$, 65, 4)))
            DESCX$ = "0": On Error Resume Next
                          DESCX$ = Str$(100 * (Val(PRELI$) - Val(PREUX$)) / Val(PRELI$))
                          On Error GoTo 0
            ITOTX$ = Str$((Val(CAPEX$)) * (Val(PREUX$)))
            FEX = CVI(Mid$(XX$, 63, 2))
            FENTX$ = FECHATEX$(FEX)
            '
            CAITAB1% = CAITAB1% + 1
            TETABU1$(1, CAITAB1%) = CODEXT$(CIPRE%)
            TETABU1$(2, CAITAB1%) = DESCRIT0$(CIPRE%)
            TETABU1$(3, CAITAB1%) = CAPEX$
            TETABU1$(4, CAITAB1%) = PRELI$
            TETABU1$(5, CAITAB1%) = DESCX$
            TETABU1$(6, CAITAB1%) = PREUX$
            TETABU1$(7, CAITAB1%) = ITOTX$
            TETABU1$(8, CAITAB1%) = FENTX$
            TETABU1$(9, CAITAB1%) = Str$(CAITAB1%)
            '
          End If
        End If
      End If
109 Next U&
    '
    SUMATOTA# = Val(Total(0).TEXT)
    IDESCUEN# = Val(Total(1).TEXT)
    '
    CODPARAM$(1) = "IMP-NETO": DOCPARAM$(1) = Str$(SUMATOTA#)
    CODPARAM$(2) = "VENDEDOR":     DOCPARAM$(2) = Eco21(2).TEXT
    CODPARAM$(3) = "COND-PAG":     DOCPARAM$(3) = Eco21(1).TEXT
    CODPARAM$(4) = "NRO-OCOM":     DOCPARAM$(4) = Text10.TEXT
    CODPARAM$(5) = "DOMI-TRA": DOCPARAM$(5) = LENTREGA(0).TEXT
    CODPARAM$(6) = "DESTINO": DOCPARAM$(6) = LENTREGA(1).TEXT
    CODPARAM$(7) = "TRANSPOR": DOCPARAM$(7) = LENTREGA(2).TEXT
    CODPARAM$(8) = "IMP-DESC": DOCPARAM$(8) = Total(1).TEXT
    CODPARAM$(9) = "IMPNEDES": DOCPARAM$(9) = Total(2).TEXT
    CODPARAM$(10) = "IMP-IVA": DOCPARAM$(10) = TOTIVA(3).TEXT
    CODPARAM$(11) = "IMP-TOTA": DOCPARAM$(11) = Total(3).TEXT
    CODPARAM$(12) = "POR-DESC": DOCPARAM$(12) = TRIM$(PORDESCU.TEXT)
    CODPARAM$(13) = "MONEDEMI": DOCPARAM$(13) = MONEMIS
    CODPARAM$(14) = "PERC-IVA": DOCPARAM$(14) = Text15
    CAPARDOC% = 14
    '
    Call BLANARCH("!OBSERVOF")
    Call FRATEXTO(Text13.TEXT, 32)
    If CARETEX% >= 1 Then Call GRAREG("!OBSERVOF", RETEX$(1), 1)
    If CARETEX% >= 2 Then Call GRAREG("!OBSERVOF", RETEX$(2), 2)
    If CARETEX% >= 3 Then Call GRAREG("!OBSERVOF", RETEX$(3), 3)
    URETE& = CARETEX%
    Call SETULTREG("!OBSERVOF", URETE&)
    Call CIERRARCH("!OBSERVOF")
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FORPECLI"))
    If FORIPRE$ <> "" Then
      Call PRINTDOC("FORPECLI")   ' PEDIDO DE CLIENTE
    End If
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FORDESPA"))
    If FORIPRE$ <> "" Then
      If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") > 0 Then
        TIPODOC$ = "Orden de Despacho"
        CANCELPRINT% = 0
        If CONTROL$("ORDESPA", "SUPRINT") = "SI" Then
           CANCELPRINT% = 1
        End If
        '
        If InStr(EPAC$, "DOSIVAC") > 0 Then
          CANCELPRINT% = 0
          If OPRIPEDI.Check1(1).Value <> 1 Then
            CANCELPRINT% = 1
          End If
        End If
        '
        If CARETEX% >= 1 Then Call GRAREG("!OBSERVOF", RETEX$(1), 1)
        If CARETEX% >= 2 Then Call GRAREG("!OBSERVOF", RETEX$(2), 2)
        If CARETEX% >= 3 Then Call GRAREG("!OBSERVOF", RETEX$(3), 3)
        URETE& = CARETEX%
        Call SETULTREG("!OBSERVOF", URETE&)
        Call CIERRARCH("!OBSERVOF")
        '
        Call PRINTDOC("FORDESPA")   ' ORDEN DE DESPACHO
      End If
    End If
    '
End Sub

Private Sub Total_GotFocus(Index As Integer)
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Sub GRAOBRA(OKX%)
   '
   If CONTROL$("ADMIPED", "VERLICRE") = "SI" Then VERLICRE% = 1
   '
   ' este bloque es para inicializar los valores de orden de despacho para SABO
   DIANTIDES% = Val(CONTROL$("", "ANTIDESP"))
   If DIANTIDES% < 1 Then DIANTIDES% = 3
   KKK% = 0
   OKEY% = 0
   '
   FORIPRE$ = TRIM$(CONTROL$("", "FORDESPA"))
   If FORIPRE$ = "" Then
       FORIPRE$ = "ODES-SAB"
   End If
   CODFOR$ = TRIM$(UCase$(FORIPRE$))
   MAXIT% = Val(ATRIFORM$(CODFOR$, "MAX-ITEM"))
   If MAXIT% < 1 Or MAXIT% > 256 Then MAXIT% = 10
   '
   REBLA$ = REGBLAN$("CARODES")
   'aqui termina el bloque de inicializacion para SABO
   '
   OKX% = 0
   If MARCONOTA.Visible = True Then Exit Sub
   If Val(Total(2).TEXT) < 0.01 Then
     If COMALTER%("Este Pedido es SIN CARGO !!!\Está Seguro de Registrarlo Igual ?\\Cancelar Pedido\Registrar Igual") <> 2 Then
       Exit Sub
     End If
   End If
   If ULTREG&("!CARDETPE") < 1 Then Exit Sub
   '
   Call CIERRARCH("PEDENCA")
   Call CIERRARCH("PEDDETA")
   Call CIERRARCH("PEDDETP")
   '
   Call BLOQARCH("PEDENCA")
   Call BLOQARCH("PEDDETA")
   Call BLOQARCH("PEDDETP")
   '
   REGENCA& = Val(Label9.Caption)
   If REGENCA& < 1 Or REGENCA& > ULTREG&("PEDENCA") Then
     Call MENSERR(24, "Imposible Modificar este Pedido.")
     Call FINAL("")
   End If
   '
   PDC$ = REGLEIDO$("PEDENCA", REGENCA&)
   If CVS(Left$(PDC$, 4)) <> Val(Text11.TEXT) Then
     Call MENSERR(24, "Imposible Modificar este Pedido.")
     Call FINAL("")
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
   'If Val(Total(2).TEXT) < 0.01 Then  VERIFICADO MAS ARRIBA
   '   MERRO$ = "Importe no Válido"
   '   GoTo 98
   'End If
   '
   HOII% = HOY(HOS$)
   FEMI% = FECHANUM(Text9.TEXT)
   If FEMI% < 1 Or FEMI% > HOII% Then
      MERRO$ = "Fecha de Emisión no Válida"
      GoTo 98
   End If
   '
   REFE$ = " - " + TRIM$(Text10.TEXT)
   If REFE$ = " - " Then REFE$ = ""
   FEX = FEMI%
   REFE$ = REFE$ + " - " + FECHATEX$(FEX)
   LOMA% = 48 - Len(REFE$)
   REFE$ = Left$(TRIM$(ECOARCH$("DEUDOR1", MKI$(NCLIE%))), LOMA%) + REFE$
   '
   PDC$ = REGLEIDO$("PEDENCA", REGENCA&)
   NROPE& = Val(Text11.TEXT)
     Call REPLA(PDC$, MKS$(NROPE&), 1, 4)
     Call REPLA(PDC$, REFE$, 5, 48)
     Call REPLA(PDC$, MKI$(FEMI%), 53, 2)
     Call REPLA(PDC$, MKI$(NCLIE%), 55, 2)
     Call REPLA(PDC$, TRIM$(Text10.TEXT), 57, 16)
     Call REPLA(PDC$, TRIM$(LENTREGA(0).TEXT), 73, 32)
     Call REPLA(PDC$, TRIM$(LENTREGA(1).TEXT), 105, 32)
     Call REPLA(PDC$, TRIM$(LENTREGA(2).TEXT), 137, 32)
   Call FRATEXTO(Text13.TEXT, 32)
     If CARETEX% >= 1 Then Call REPLA(PDC$, RETEX$(1), 169, 32)
     If CARETEX% >= 2 Then Call REPLA(PDC$, RETEX$(2), 201, 32)
     If CARETEX% >= 3 Then Call REPLA(PDC$, RETEX$(3), 233, 32)
     Call REPLA(PDC$, MKI$(0), 265, 2)
     Call REPLA(PDC$, Chr$(1), 267, 1)
     Call REPLA(PDC$, Chr$(Val(TEXT21(0).TEXT)), 268, 1)
     Call REPLA(PDC$, Chr$(Val(TEXT21(1).TEXT)), 269, 1)
     Call REPLA(PDC$, Chr$(Val(TEXT21(2).TEXT)), 270, 1)
     Call REPLA(PDC$, MKS$(1 + ULTREG&("PEDDETA")), 271, 4)
     Call REPLA(PDC$, TRIM$(PORDESCU.TEXT), 275, 14)
     Call REPLA(PDC$, CAVERI$, 289, 16)
   Call GRAREG("PEDENCA", PDC$, REGENCA&)
   Call CIERRARCH("PEDENCA")
   '
   RFFX$ = RECFILT$("PEDDETP", 1, MKS$(NROPE&))
   For UKF& = 1 To Len(RFFX$) Step 4
     UU& = CVS(Mid$(RFFX$, UKF&, 4))
     XX$ = REGLEIDO$("PEDDETP", UU&)
     If CVS(Left$(XX$, 4)) = NROPE& Then
       ZZ$ = REGBLAN$("PEDDETP")
       Call GRAREG("PEDDETP", ZZ$, UU&)
     End If
   Next UKF&
   '
   RFFX$ = RECFILT$("DESPAPED", 1, MKS$(NROPE&))
   For UKF& = 1 To Len(RFFX$) Step 4
     UU& = CVS(Mid$(RFFX$, UKF&, 4))
     XX$ = REGLEIDO$("DESPAPED", UU&)
     If CVS(Left$(XX$, 4)) = NROPE& Then
       ZZ$ = REGBLAN$("DESPAPED")
       Call GRAREG("DESPAPED", ZZ$, UU&)
     End If
   Next UKF&
   Call CIERRARCH("DESPAPED")
   '
   IKAMA% = 0
   RFFX$ = RECFILT$("PEDDETA", 1, MKS$(NROPE&))
   For UKF& = 1 To Len(RFFX$) Step 4
     UU& = CVS(Mid$(RFFX$, UKF&, 4))
     XX$ = REGLEIDO$("PEDDETA", UU&)
     If CVS(Left$(XX$, 4)) = NROPE& Then
       If Val(Mid$(XX$, 113 + 7, 3)) > IKAMA% Then
         IKAMA% = Val(Mid$(XX$, 113 + 7, 3))
       End If
       CAN = CVS(Mid$(XX$, 79, 4))
       CADESPACH = CVS(Mid$(XX$, 95, 4))
       CAPENDENT = CAN - CADESPACH
       If CAPENDENT > 0.5 Then
           ZZ$ = REGBLAN$("PEDDETA")
           Call GRAREG("PEDDETA", ZZ$, UU&)
         Else
           Call REGAPP("PEDDETA", XX$)
           UREPA& = ULTREG&("PEDDETA")
           Call REPLA(XX$, MKS$(UREPA&), 109, 4)
           Call REGAPP("PEDDETP", XX$)
       End If
     End If
   Next UKF&
   '
   TTXYZ$ = Space$(76)
   HOII% = HOY(HOS$)
   Call REPLA(TTXYZ$, HOS$, 1, 8)
   Call REPLA(TTXYZ$, "Modificación de Pedido:", 11, 56)
   Call REPLA(TTXYZ$, USERTER$, 69, 24)
   Call REGAPP("ANOTAPED", MKS$(NROPE&) + TTXYZ$)
   '
   CAVER$ = TRIM$(Str$(NROPE&))
   While Len(CAVER$) > 6: CAVER$ = Mid$(CAVER$, 2): Wend
   While Len(CAVER$) < 6: CAVER$ = "0" + CAVER$: Wend
   IKITEM% = IKAMA%
   JJ& = 0
   '
   For U& = 1 To ULTREG&("!CARDETPE")
     X$ = REGLEIDO$("!CARDETPE", U&)
     CAN = CVS(Mid$(X$, 47, 4))
     CAVERI$ = TRIM$(Mid$(X$, 113, 10))
     If CAN > 0.05 Then
       PDC$ = REGBLAN$("PEDDETA")
       Call REPLA(PDC$, MKS$(NROPE&), 1, 4)
       Call REPLA(PDC$, MKI$(FEMI%), 5, 2)
       Call REPLA(PDC$, MKI$(NCLIE%), 7, 2)
       Call REPLA(PDC$, TRIM$(Text10.TEXT), 9, 16)
         FENTRESOL% = CVI(Mid$(X$, 63, 2))
       Call REPLA(PDC$, MKI$(FENTRESOL%), 25, 2)
       Call REPLA(PDC$, Chr$(1), 27, 1)
       Call REPLA(PDC$, Chr$(Val(TEXT21(0).TEXT)), 28, 1)
       Call REPLA(PDC$, Chr$(Val(TEXT21(1).TEXT)), 29, 1)
       Call REPLA(PDC$, Chr$(Val(TEXT21(2).TEXT)), 30, 1)
       Call REPLA(PDC$, MKI$(10 * Val(Text44.TEXT)), 31, 2)
       Call REPLA(PDC$, Left$(X$, 62), 33, 62)
       '
       Call REPLA(PDC$, String$(8, 0), 51, 8)
       Call REPLA(PDC$, Left$(Label11, 1), 51, 1) ' venta o reparacion
       Call REPLA(PDC$, Chr$(MONEMI%), 52, 1) ' MONEDA DE EMISION
       Call REPLA(PDC$, MKS$(Val(TICAMBIS)), 53, 4)
       '
       Call REPLA(PDC$, String$(18, 0), 95, 18)
       Call REPLA(PDC$, MKS$(CAN), 99, 4)
       Call REPLA(PDC$, MKS$(REGENCA&), 109, 4)  ' puntero a ENCABEZADO
       '
       If Left$(CAVERI$, 6) <> CAVER$ Then
         IKITEM% = IKITEM% + 1
         IK$ = TRIM$(Str$(IKITEM%))
         CAVERI$ = CAVER$ + "-" + IK$
       End If
       Call REPLA(PDC$, CAVERI$, 113, 10)
       '
       Call REGAPP("PEDDETA", PDC$)
       RGX& = ULTREG&("PEDDETA")
       Call REPLA(PDC$, MKS$(RGX&), 109, 4)  ' puntero a PEDDETA
       Call REGAPP("PEDDETP", PDC$)
       '
       '
' a partir de aca graba !CARODES para la orden de despacho sabo
       FEPED% = CVI(Mid$(PDC$, 5, 2))
       CLIPED% = CVI(Mid$(PDC$, 7, 2))
       CIPED% = CVI(Mid$(PDC$, 33, 2))
       REGPEDA& = RGX&
       REGPEDP& = ULTREG&("PEDDETP")
       MOTINOAS$ = ""
       '
       Call VESARRAS(CIPED%)
       STODISPO = STODEPOS(CIPED%, DEPOEXPE%)
       STODISPO = STODISPO - STOCKASI(CIPED%)
       For J1& = 1 To JJ&  ' netea de preasignacion anterior
          Z$ = REGLEIDO$("!CARODES", J1&)
          If CVI(Left$(Z$, 2)) = CIPED% Then
            STODISPO = STODISPO - CVS(Mid$(Z$, 33, 4))
          End If
       Next J1&
       If STODISPO < 0 Then STODISPO = 0
       '
       CAPEDORI = CVS(Mid$(PDC$, 79, 4))
       If CAPEDORI < 0.5 Then
         MOTINOAS$ = "Pedido Nulo"
       End If
       '
       CAPENPED = CAPEDORI - CVS(Mid$(PDC$, 95, 4)) 'Q.entregada
       CAPENPED = CAPENPED - CVS(Mid$(PDC$, 75, 4)) 'Q.ya asignada
       If CAPENPED < 0 Then CAPENPED = 0
       '
       CANASIG = CAPENPED
       If CANASIG > STODISPO Then
         MOTINOAS$ = "Stock Dispon."
         CANASIG = STODISPO
       End If
       '
       FENTRESO% = CVI(Mid$(PDC$, 25, 2))
       FE1 = HOY(HOS$): FE2 = FENTRESO%
       If DIENTRE%(FE1, FE2) > DIANTIDES% Then
         MOTINOAS$ = "Fech.Ent.Sol"
         CANASIG = 0
       End If
       '
       Z$ = REBLA$
       Call REPLA(Z$, MKI$(CIPED%), 1, 2)
       Call REPLA(Z$, Space$(16), 3, 16)   ' lote / color
       Call REPLA(Z$, MKI$(FENTRESO%), 19, 2)
       Call REPLA(Z$, MKS$(CAPEDORI), 21, 4)
       Call REPLA(Z$, MKS$(CAPENPED), 25, 4)
       Call REPLA(Z$, MKS$(STODISPO), 29, 4)
       Call REPLA(Z$, MKS$(CANASIG), 33, 4)
       Call REPLA(Z$, MKS$(REGPEDA&), 37, 4)
       Call REPLA(Z$, MKS$(REGPEDP&), 41, 4)
       Call REPLA(Z$, MKS$(REGPEDP&), 45, 4)   ' REGISTRO DE PEDDETP
       Call REPLA(Z$, MOTINOAS$, 49, 16)       ' MOTIVO NO ASIGNACION
       JJ& = JJ& + 1
       Call GRAREG("!CARODES", Z$, JJ&)
' hasta aqui grabo archivo !CARODES para SABO
       '
       '
       If Mid$(PDC$, 51, 1) <> "R" Then
         Call REPLA(PDC$, Chr$(1), 27, 1)
         Call REPLA(PDC$, String$(38, 0), 71, 38)
         Call REPLA(PDC$, MKS$(CAN), 75, 4)
         Call REPLA(PDC$, MKS$(CAN), 99, 4)
         Call REGAPP("DESPAPED", PDC$)
       End If
       '
       TTXX$ = Space$(76)
       CIXI% = CVI(Mid$(PDC$, 33, 2))
       Call REPLA(TTXX$, "- " + CODEXT$(CIXI%), 11, 18)
       Call REPLA(TTXX$, DESCRIT0$(CIXI%), 29, 38)
       Call REPLA(TTXX$, CSTRING$(MKS$(CAN), 3, 10, 1, 2), 57, 10)
       Call REGAPP("ANOTAPED", Left$(MKS$(NROPE&) + TTXX$, 96))
       '
     End If
   Next U&
   Call REGAPP("ANOTAPED", Left$(MKS$(NROPE&) + Space$(10) + String$(56, "-"), 96))
   '
   Call SETULTREG("!CARODES", JJ&)
   Call CIERRARCH("!CARODES")
   '
   Call CIERRARCH("PEDDETA")
   Call CIERRARCH("PEDDETP")
   Call CIERRARCH("DESPAPED")
   '
   Call FRATEXTO(Text13.TEXT, 32)
   For KKI% = 1 To CARETEX%
     TTXX$ = Space$(92)
     Call REPLA(TTXX$, RETEX$(KKI%), 16, 32)
     If KKI% = 1 Then Call REPLA(TTXX$, "Observaciones:", 1, 15)
     Call REGAPP("ANOTAPED", Left$(MKS$(NROPE&) + Space$(10) + TTXX$, 96))
   Next KKI%
   '
   Call REGAPP("ANOTAPED", MKS$(NROPE&) + String$(88, "="))
   Call CIERRARCH("ANOTAPED")
   '
   LENTREG$ = Space$(96)
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
       XYX$ = REGLEIDO$("LENTRECL", RELENT&)
       If TRIM$(Mid$(XYX$, 3, 96)) = "" Then
         Call GRAREG("LENTRECL", MKI$(NCLIE%) + LENTREG$, RELENT&)
       End If
     Else
       Call REGAPP("LENTRECL", MKI$(NCLIE%) + LENTREG$)
   End If
   Call CIERRARCH("LENTRECL")
   '
   Call CIERRARCH("PEDDETA")
   Call CIERRARCH("PEDDETP")
   Call CIERRARCH("DESPAPED")
   Call CIERRARCH("*.*")
   Call LIBARCH("*.*")
   Call BAJALDISCO
   '
   OKX% = 1
   If VERLICRE% = 1 Then
     LMCR# = LIMICRE(NCLIE%)
     SDEU# = SALDCLI#(NCLIE%)
     CHPA# = CHEPEACRE#(NCLIE%)
     TPED# = Val(Total(3).TEXT)
     If SDEU# + CHPA# + TPED# > LMCR# Then
       TLIMI$ = TRIM$(CSTRING$(MKD$(LMCR#), 3, 12, 2, 2))
       If TLIMI$ = "" Then TLIMI$ = "0.00"
       ATTX$ = "ATENCION: Este Pedido Supera\el Límite de Crédito Asignado\para esta Cuenta de $ " + TLIMI$
       If DERACCE%("AUTORENT/NOMESS", "Autorizar Pedidos con Exceso Credito") < 2 Then
           Call COMUNI(ATTX$ + ".\  \Informe al Supervisor de Ventas\para Autorizar Facturación y Despacho.")
         Else
           AUTORENT$ = "NO"
           If COMALTER%(ATTX$ + "\   \Puede Autorizar la Entrega\con Crédito Excedido.\\Continuar\Autorizar Entrega") = 2 Then AUTORENT$ = "SI"
           RFF$ = RECFILT$("AUTORENT", 1, MKS$(NROPE&))
           REGAUT& = 0: If Len(RFF$) >= 4 Then REGAUT& = CVS(Left$(RFF$, 4))
           AUTORI$ = MKS$(NROPE&) + AJUSTI$(USERNAME$, 26) + AUTORENT$
           If REGAUT& > 0 Then
               Call GRAREG("AUTORENT", AUTORI$, REGAUT&)
             Else
               Call REGAPP("AUTORENT", AUTORI$)
           End If
       End If
     End If
   End If
   Screen.MousePointer = 1
   Exit Sub
   '
98 Call CIERRARCH("PEDENCA")
   Call CIERRARCH("PEDDETA")
   Call CIERRARCH("PEDDETP")
   Screen.MousePointer = 1
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
     TTT1 = Val(Total(0).TEXT)
     Total(0) = CSTRING$(MKS$(TTT1), 3, 11, 2, 2)
   End If
End Sub

Sub CARDETPEDI(MODO%)
   '
   HOII% = HOY(HOS$)
   '
   ATRI$ = "/NC/ARCHECO(1)=" + ARCHIPRE$
   ATRI$ = ATRI$ + "/TITUPAN=Detalle Pedido de Cliente"
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Frame1.Top + Picture3.Top))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture3.Left + 50))
   '
   Timer1.Enabled = True
10 VENCARPE$ = CONTROL$("", "VENCARPE")
   If TRIM$(VENCARPE$) = "" Then VENCARPE$ = "CARDETP"
   PRF$ = VENCARPE$    ' If FORMAPED% = 2 Then PRF$ = "CARDETQ"
   XXX% = VENTABU%(PRF$ + ATRI$)
   Call CALTOPED
   If ECOPRECIO% >= 0 Then
     If Val(Total(2).TEXT) < 0.005 Then
       If TRIM$(Text13.TEXT) = "" Then
         Text13.TEXT = "SIN CARGO !!!"
       End If
     End If
   End If
   Timer1.Enabled = False
   '
   If XXX% < 0 Then
      Call CIERRARCH("!CARDETPE")
      Call BLANARCH("!CARDETPE")
      Call BLANFORMU
      Text1.TEXT = ""
      Call CIERRARCH("*.*")
      Hide
      Exit Sub
   End If
   '
20 If ULTREG&("!CARDETPE") < 1 Then
      Call BLANFORMU
      Text1.TEXT = ""
      Call CIERRARCH("*.*")
      Hide
      Exit Sub
   End If
   '
   Screen.MousePointer = 11   ' Registrando Pedido ...
   List1.Clear
   FIN& = ULTREG&("!CARDETPE")
   If FIN& > 0 Then
     Y$ = REGLEIDO$("!CARDETPE", FIN&)
     CI% = CVI(Left$(Y$, 2))
     If CI% < 1 Or CI% > NUMAS% Then
       Call SETULTREG("!CARDETPE", FIN& - 1)
       Call CIERRARCH("!CARDETPE")
       GoTo 20
     End If
   End If
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   For U& = 1 To FIN&
     Y$ = REGLEIDO$("!CARDETPE", U&)
     CI% = CVI(Left$(Y$, 2))
     If CI% < 1 Or CI% > NUMAS% Then
       Call MENSERR(24, "Código no Válido")
       Call REPLA(Y$, MKI$(0), 1, 2)
       Call GRAREG("!CARDETPE", Y$, U&)
       GoTo 10
     End If
     '
     FENTRESOL% = CVI(Mid$(Y$, 63, 2))
     If FENTRESOL% < HOII% Then
        FENTRESOL% = HOII%
        Call REPLA(Y$, MKI$(FENTRESOL%), 63, 2)
        Call GRAREG("!CARDETPE", Y$, U&)
     End If
     '
     ' para DOSIVAC no verifica codigos repetidos
     If InStr(EPAC$, "DOSIVA") < 1 Then
       For W& = 1 To U& - 1
         Y$ = REGLEIDO$("!CARDETPE", W&)
         If CVI(Left$(Y$, 2)) = CI% Then
           If CVI(Mid$(Y$, 63, 2)) = FENTRESOL% Then
             Screen.MousePointer = 1
             TTXX$ = "Código " + TRIM$(CODEXT$(CI%)) + "\Duplicado en el Pedido."
             If COMALTER%(TTXX$ + "\\Corregir\Continuar") <> 2 Then
               GoTo 10
             End If
           End If
         End If
       Next W&
     End If
     '
     X$ = REGLEIDO$("!CARDETPE", U&)
     CI% = CVI(Left$(X$, 2))
     TTXX$ = Space$(96)
     FENTRESOL% = CVI(Mid$(X$, 63, 2))
     If FENTRESOL% < HOII% Then
        FENTRESOL% = HOII%
        Call REPLA(X$, MKI$(FENTRESOL%), 63, 2)
        Call GRAREG("!CARDETPE", X$, U&)
     End If
     '
     CAN = CVS(Mid$(X$, 47, 4))
     If CAN > 999999 Then
       Call MENSERR(24, "Cantidad Excesiva - Max 999999")
       Call REPLA(X$, MKS$(0), 47, 4)
       Call GRAREG("!CARDETPE", X$, U&)
       GoTo 10
     End If
     If CAN < 0.5 Then
       Call MENSERR(24, "Falta Cantidad")
       Call REPLA(X$, MKS$(0), 47, 4)
       Call GRAREG("!CARDETPE", X$, U&)
       GoTo 10
     End If
     '
     PORDESCUIN = CVS(Mid$(X$, 69, 4))
     If PORDESCUIN < 0 Or PORDESCUIN > 99.9 Then
       Call MENSERR(24, "Descuento no Válido")
       Call REPLA(X$, MKS$(0), 69, 4)
       Call GRAREG("!CARDETPE", X$, U&)
       GoTo 10
     End If
     '
     If ECOPRECIO% < 0 Then
       PRECARGA = PRELISTA#(ARCHIPRE$, CI%)
       Call REPLA(X$, MKS$(PRECARGA), 39, 4)
       Call GRAREG("!CARDETPE", X$, U&)
     End If
     PRECARGA = CVS(Mid$(X$, 39, 4))
     PREUNI = PRECARGA * (100 - PORDESCUIN) / 100
     TOT# = CDbl(CAN * PREUNI)
     '
     If PRECARGA < 0.005 Or PRECARGA > 99999.99 Then
       If MODO% < 2 Then
         If ECOPRECIO% >= 0 Then
           Call MENSERR(24, "Artículo sin Precio")
           Call REPLA(X$, MKS$(0), 39, 4)
           Call GRAREG("!CARDETPE", X$, U&)
           MODO% = 2
           GoTo 10
         End If
       End If
     End If
     '
     Call REPLA(X$, MKS$(PRECARGA), 39, 4)
     Call REPLA(X$, MKS$(PRECARGA), 51, 4)
     Call REPLA(X$, MKD$(TOT#), 55, 8)
     Call REPLA(X$, MKS$(PRECARGA), 65, 4)
     Call REPLA(X$, MKS$(PREUNI), 73, 4)
     Call REPLA(X$, MKD$(TOT#), 77, 8)
     Call GRAREG("!CARDETPE", X$, U&)
     '
   Next U&
   '
   Call CARLISPE
   Call CIERRARCH("!CARDETPE")
   Call CALTOPED
   '
   On Error Resume Next
   Command14.SetFocus
   On Error GoTo 0
   Screen.MousePointer = 1
   '
End Sub

Sub CARLISPE()
   List1.Clear
   FIN& = ULTREG&("!CARDETPE")
   For U& = 1 To FIN&
     X$ = REGLEIDO$("!CARDETPE", U&)
     TTXX$ = Space$(96)
     CI% = CVI(Left$(X$, 2))
     FENTRESOL% = CVI(Mid$(X$, 63, 2))
     CAN = CVS(Mid$(X$, 47, 4))
     PORDESCUIN = CVS(Mid$(X$, 69, 4))
     PRECARGA = CVS(Mid$(X$, 39, 4))
     PREUNI = PRECARGA * (100 - PORDESCUIN) / 100
     TOT# = CDbl(CAN * PREUNI)
     '
     Call REPLA(TTXX$, CSTRING$(MKS$(CAN), 3, 6, 0, 2), 1, 6)
     Call REPLA(TTXX$, CODEXT$(CI%), 8, 15)
     FEX = FENTRESOL%
     If ECOPRECIO% < 0 Then
          Call REPLA(TTXX$, DESCRIT0$(CI%), 24, 52)
          Call REPLA(TTXX$, FECHATEX$(FEX), 78, 8)
        Else
          Call REPLA(TTXX$, DESCRIT0$(CI%), 24, 22)
          Call REPLA(TTXX$, Left$(FECHATEX$(FEX), 5), 47, 5)
     End If
     PRELIX = PRELISTA#(ARCHIPRE$, CI%)
     If PRELIX < PRECARGA Then PRELIX = PRECARGA
     If PRELIX < PREUNI Then PRELIX = PREUNI
     PORDESCUIN = 0: If PRELIX > 0.005 Then PORDESCUIN = 100 * (PRELIX - PREUNI) / PRELIX
     If ECOPRECIO% >= 0 Then Call REPLA(TTXX$, CSTRING$(MKS$(PRELIX), 3, 8, 2, 2), 52, 8)
     If ECOPRECIO% >= 0 Then Call REPLA(TTXX$, CSTRING$(MKS$(PORDESCUIN), 3, 5, 1, 2), 61, 5)
     If ECOPRECIO% >= 0 Then Call REPLA(TTXX$, CSTRING$(MKS$(PREUNI), 3, 8, 2, 2), 66, 8)
     If ECOPRECIO% >= 0 Then Call REPLA(TTXX$, CSTRING$(MKD$(TOT#), 3, 10, 2, 2), 76, 10)
     List1.AddItem TTXX$
     '
   Next U&

End Sub
'
Sub CALTOPED()
   '
   Static CLIA$, RNI%, APLIVA%
   If Text1.TEXT <> CLIA$ Then
     CLIA$ = Text1.TEXT: NCLIE% = Val(CLIA$)
     RNI% = 0: APLIVA% = 1
     If NCLIE% > 0 Then
       PVCLI% = PIVACLI%(NCLIE%)
       If PVCLI% = 2 Then RNI% = 1
       If PVCLI% = 5 Or PVCLI% = 7 Then APLIVA% = 0
     End If
   End If
   '
   SUTOTPED# = 0: SUIVA# = 0: SUIVARNI# = 0: TOIVA# = 0: TOTOTA# = 0
   TOTITEMS% = 0
   For U& = 1 To ULTREG&("!CARDETPE")
     X$ = REGLEIDO$("!CARDETPE", U&)
     CIXXI% = CVI(Left$(X$, 2))
     CAN = CVS(Mid$(X$, 47, 4))
     PORDESCUIN = CVS(Mid$(X$, 69, 4))
     If PORDESCUEIN < 0 Or PORDESCUEIN > 99 Then PORDESCUEIN = 0
     If CAN > 0.05 Then
       PRELI = CVS(Mid$(X$, 39, 4))
       'If PRELI < 0.01 Then
       '  PRELI = PRELISTA(ARCHIPRE$, CIXXI%)
       'End If
       PREUNI = PRELI * (100 - PORDESCUIN) / 100
       TOT# = CDbl(CAN * PREUNI)
       SUTOTPED# = SUTOTPED# + TOT#
       SUIVA# = SUIVA# + (Int(TOT# * PORCEIVA(CIXXI%) + 0.5)) / 100
       Call REPLA(X$, MKS$(PREUNI) + MKD$(TOT#), 51, 12)
       Call GRAREG("!CARDETPE", X$, U&)
       TOTITEMS% = TOTITEMS% + 1
     End If
   Next U&
   '
   Label18 = TRIM$(Str$(TOTITEMS%))
   '
   Total(0).TEXT = CSTRING$(MKD$(SUTOTPED#), 3, 11, 2, 2)
   Total(0).Refresh
   '
   DESC$ = TRIM$(PORDESCU.TEXT)
   TTT1 = Val(Total(0).TEXT) * COEFDESCU(DESC$) / 100
   Total(1).TEXT = CSTRING$(MKS$(TTT1), 3, 11, 2, 2)
   '
   TTT3 = Val(Total(0).TEXT) - Val(Total(1).TEXT)
   Total(2).TEXT = CSTRING$(MKS$(TTT3), 3, 11, 2, 2)
   '
   On Error Resume Next
   TOTIVA(0).TEXT = CSTRING$(MKD$(APLIVA% * SUIVA# * Val(Total(2).TEXT) / Val(Total(0).TEXT)), 4, 11, 2, 2)
   On Error GoTo 0
   TOTIVA(2).TEXT = CSTRING$(MKD$(CDbl(Val(TOTIVA(0).TEXT) * 0.5 * RNI%)), 4, 11, 2, 2)
   TOTIVA(3).TEXT = CSTRING$(MKD$(CDbl(Val(TOTIVA(0).TEXT) + Val(TOTIVA(2).TEXT))), 4, 11, 2, 2)
   '
   Call CALPERIB
   '
   Total(3).TEXT = CSTRING$(MKD$(CDbl(Val(Total(2).TEXT) + Val(TOTIVA(3).TEXT)) + Val(Text15)), 4, 11, 2, 2)
   '
End Sub

Sub GENERODES()
   '
   If CONTROL$("ADMIPED", "GENDESPA") = "REVISTOCK" Then
     Call GENODESPA(Val(Text11), NODES&, 1, OKKY%)
     If OKKY% = 1 Then
        Call PRIODESPA(NODES&)
     End If
   End If
   '
End Sub
'
Sub CALPERIB()
   '
   Static CTTXX%, TASAA, TASAB, TASAEXP, MINIPER
   '
   'If TIFADOC% > 2 Then
   '   Exit Sub
   'End If
   '
   If CTTXX% < 1 Then
     TASAA = Val(CONTROL$("PERIBRU", "TASA-A"))
     TASAB = Val(CONTROL$("PERIBRU", "TASA-B"))
     TASAEXP = Val(CONTROL$("PERIBRU", "TASA-EXP"))
     MINIPER = Val(CONTROL$("PERIBRU", "MINIPER"))
     CTTXX% = 1
   End If
   '
   TASAX = TASAA
   NCLIE% = Val(Text1)
   If NCLIE% < 1 Then Exit Sub
   '
   VARIDC% = 1: PVCLX% = PIVACLI%(NCLIE%)
   If PVCLX% = 0 Or PVCLX% = 4 Then VARIDC% = 2
   If PVCLX% = 5 Or PVCLX% = 7 Then VARIDC% = 3
   '
   If VARIDC% = 2 Then TASAX = TASAB
   If VARIDC% = 3 Then TASAX = TASAEXP
   '
   Text15 = ""
   REPERCI$ = FILTERGETRECORD$("DAPIBCLI", 1, MKI$(NCLIE%))
   CATIB1% = Asc(Mid$(REPERCI$, 3, 1))
   PORCONVE = CVS(Mid$(REPERCI$, 5, 4))
   If PORCONVE > 100 Then PORCONVE = 100
   '
   If CATIB1% > 0 Then         ' NO ESTA EXENTO
     If PORCONVE >= 10.1 Then  ' PARTICIPACION CONVENIO MULTILATERAL
       BIMPON# = Val(Total(2))
       If VARIDC% = 2 Then BIMPON# = BIMPON# + Val(TOTIVA(3))
       If BIMPON# * TICAMONE(MONEMI%) >= MINIPER Then
         PERSALTA# = (BIMPON# * PORCONVE / 100) * (TASAX / 100)
         Text15 = CSTRING$(MKD$(PERSALTA#), 3, 11, 2, 2)
       End If
     End If
   End If
   '
End Sub
'


Private Sub TOTIVA_GotFocus(Index As Integer)
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub
