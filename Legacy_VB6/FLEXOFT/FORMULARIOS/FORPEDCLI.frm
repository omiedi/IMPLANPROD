VERSION 5.00
Begin VB.Form FORPEDCLI 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pedidos de Clientes - Formato Lineal"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   Icon            =   "FORPEDCLI.frx":0000
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
      TabIndex        =   36
      Top             =   5145
      Width           =   10800
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
         ForeColor       =   &H00808000&
         Height          =   3120
         Left            =   5670
         TabIndex        =   60
         Top             =   0
         Width           =   5055
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
            Height          =   300
            Index           =   2
            Left            =   3405
            TabIndex        =   104
            TabStop         =   0   'False
            Text            =   "    *******"
            Top             =   1125
            Visible         =   0   'False
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
            Left            =   3405
            TabIndex        =   98
            Top             =   2100
            Width           =   1485
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
            Left            =   3405
            TabIndex        =   97
            Text            =   " "
            Top             =   2625
            Width           =   1500
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
            Height          =   315
            Index           =   3
            Left            =   3405
            TabIndex        =   96
            Text            =   " "
            Top             =   1680
            Width           =   1500
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
            Left            =   840
            MaxLength       =   14
            TabIndex        =   88
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
            Height          =   300
            Index           =   1
            Left            =   3405
            TabIndex        =   84
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
            Height          =   300
            Index           =   0
            Left            =   3405
            TabIndex        =   83
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
            TabIndex        =   63
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
            Left            =   3420
            TabIndex        =   62
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
            Left            =   1995
            MaxLength       =   14
            TabIndex        =   23
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
            TabIndex        =   61
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1125
            Width           =   1500
         End
         Begin VB.Label Label14 
            BackColor       =   &H0000C000&
            Caption         =   "IVA 2"
            Height          =   255
            Left            =   120
            TabIndex        =   108
            Top             =   2640
            Width           =   1215
         End
         Begin VB.Label Label13 
            BackColor       =   &H0000C000&
            Caption         =   "IVA 1"
            Height          =   255
            Left            =   120
            TabIndex        =   107
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   210
            Picture         =   "FORPEDCLI.frx":0442
            Top             =   315
            Visible         =   0   'False
            Width           =   480
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
            TabIndex        =   101
            Top             =   1785
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
            TabIndex        =   100
            Top             =   2205
            Width           =   3060
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
            TabIndex        =   99
            Top             =   2730
            Width           =   1905
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
            TabIndex        =   91
            Top             =   525
            Width           =   1905
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
            TabIndex        =   90
            ToolTipText     =   "Doble-Click para 'Dólares'"
            Top             =   730
            Visible         =   0   'False
            Width           =   1380
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
            Left            =   375
            TabIndex        =   89
            Top             =   1200
            Visible         =   0   'False
            Width           =   435
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
            TabIndex        =   66
            Top             =   390
            Width           =   1905
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Pedido:"
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
            TabIndex        =   65
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
            Left            =   3150
            TabIndex        =   64
            Top             =   840
            Width           =   225
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   1590
         Left            =   4095
         TabIndex        =   92
         Top             =   1575
         Visible         =   0   'False
         Width           =   2220
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
            Left            =   210
            TabIndex        =   95
            Text            =   " "
            Top             =   210
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
            Left            =   210
            TabIndex        =   94
            Text            =   " "
            Top             =   525
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
            Left            =   210
            TabIndex        =   93
            Text            =   " "
            Top             =   840
            Visible         =   0   'False
            Width           =   870
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
         ForeColor       =   &H00808000&
         Height          =   3120
         Left            =   210
         TabIndex        =   52
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
            TabIndex        =   82
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
            TabIndex        =   19
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
            TabIndex        =   55
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
            TabIndex        =   17
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
            TabIndex        =   18
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
            TabIndex        =   54
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
            TabIndex        =   15
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
            TabIndex        =   16
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
            TabIndex        =   53
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
            TabIndex        =   103
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
            TabIndex        =   102
            Top             =   2680
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
            TabIndex        =   81
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
            TabIndex        =   59
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
            TabIndex        =   58
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
            TabIndex        =   57
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
            TabIndex        =   56
            ToolTipText     =   "Doble-Click para Agregar Artículo"
            Top             =   495
            Width           =   1590
         End
      End
      Begin VB.Frame Frame3 
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
         Height          =   1500
         Left            =   5670
         TabIndex        =   41
         Top             =   1620
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
            TabIndex        =   26
            Top             =   1085
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
            TabIndex        =   25
            Top             =   700
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
            TabIndex        =   24
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
            TabIndex        =   72
            Top             =   1085
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
            TabIndex        =   51
            Top             =   1155
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
            TabIndex        =   42
            Top             =   400
            Width           =   1905
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8520
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   12000
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   11000
         ScaleHeight     =   75
         ScaleWidth      =   630
         TabIndex        =   105
         Top             =   7140
         Width           =   695
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   106
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   735
         Top             =   0
      End
      Begin VB.CommandButton Command20 
         Height          =   720
         Left            =   4995
         Picture         =   "FORPEDCLI.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   86
         ToolTipText     =   "Búsqueda por CUIT, Domicilio, Nombre de Fantasía, etc"
         Top             =   890
         Width           =   540
      End
      Begin VB.PictureBox MARSELPRESU 
         BackColor       =   &H00FFFF00&
         Height          =   3180
         Left            =   210
         ScaleHeight     =   3120
         ScaleWidth      =   10455
         TabIndex        =   73
         Top             =   1785
         Visible         =   0   'False
         Width           =   10520
         Begin VB.ListBox List2 
            BackColor       =   &H00F0FFC0&
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
            Left            =   -30
            Style           =   1  'Checkbox
            TabIndex        =   79
            Top             =   420
            Width           =   9780
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00FFFF00&
            Height          =   470
            Left            =   -30
            ScaleHeight     =   405
            ScaleWidth      =   9720
            TabIndex        =   77
            Top             =   -30
            Width           =   9780
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Selección de Presupuestos en Curso"
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
               TabIndex        =   78
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
            Picture         =   "FORPEDCLI.frx":1AB2
            Style           =   1  'Graphical
            TabIndex        =   76
            ToolTipText     =   "Continúa Generación del Pedido de Cliente"
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
            Picture         =   "FORPEDCLI.frx":1DBC
            Style           =   1  'Graphical
            TabIndex        =   75
            ToolTipText     =   "Cancela Pedido de Cliente"
            Top             =   1050
            Width           =   570
         End
         Begin VB.CommandButton Command17 
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
            Picture         =   "FORPEDCLI.frx":20C6
            Style           =   1  'Graphical
            TabIndex        =   74
            ToolTipText     =   "Imprime Ficha de Anotaciones"
            Top             =   2470
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
      Begin VB.PictureBox MARCONOTA 
         BackColor       =   &H00FFFFC0&
         Height          =   3180
         Left            =   210
         ScaleHeight     =   3120
         ScaleWidth      =   10455
         TabIndex        =   67
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
            Picture         =   "FORPEDCLI.frx":2730
            Style           =   1  'Graphical
            TabIndex        =   22
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
            Picture         =   "FORPEDCLI.frx":2D9A
            Style           =   1  'Graphical
            TabIndex        =   21
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
            Picture         =   "FORPEDCLI.frx":30A4
            Style           =   1  'Graphical
            TabIndex        =   20
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
            TabIndex        =   70
            Top             =   390
            Width           =   9780
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FFFFC0&
            Height          =   435
            Left            =   -30
            ScaleHeight     =   375
            ScaleWidth      =   9720
            TabIndex        =   68
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
               TabIndex        =   69
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
         TabIndex        =   71
         Top             =   2205
         Width           =   10520
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
         TabIndex        =   43
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
            Picture         =   "FORPEDCLI.frx":33AE
            Style           =   1  'Graphical
            TabIndex        =   87
            ToolTipText     =   "Anotador de Pedidos"
            Top             =   3885
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
            Picture         =   "FORPEDCLI.frx":36B8
            Style           =   1  'Graphical
            TabIndex        =   29
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
            Picture         =   "FORPEDCLI.frx":3802
            Style           =   1  'Graphical
            TabIndex        =   50
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
            Picture         =   "FORPEDCLI.frx":3B0C
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Configuración de Funcionamiento de Pedidos"
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
            Picture         =   "FORPEDCLI.frx":3F4E
            Style           =   1  'Graphical
            TabIndex        =   28
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
            Picture         =   "FORPEDCLI.frx":4258
            Style           =   1  'Graphical
            TabIndex        =   48
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
            Picture         =   "FORPEDCLI.frx":4562
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Acceso a Base de Datos de Clientes"
            Top             =   2315
            Width           =   855
         End
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
            Picture         =   "FORPEDCLI.frx":49A4
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Opciones Varias - Modificaciones y Consultas"
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
            Picture         =   "FORPEDCLI.frx":4AEE
            Style           =   1  'Graphical
            TabIndex        =   45
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
            Picture         =   "FORPEDCLI.frx":4DF8
            Style           =   1  'Graphical
            TabIndex        =   44
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
         Height          =   855
         Left            =   10935
         Picture         =   "FORPEDCLI.frx":5102
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Aprueba, Graba e Imprime Pedido"
         Top             =   7410
         Width           =   855
      End
      Begin VB.CommandButton Command10 
         Height          =   260
         Left            =   420
         TabIndex        =   40
         Top             =   6855
         Width           =   2825
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   210
         ScaleHeight     =   375
         ScaleWidth      =   10455
         TabIndex        =   38
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
            TabIndex        =   39
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
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1280
         Width           =   3435
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
         TabIndex        =   11
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
         TabIndex        =   12
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
         Left            =   5775
         TabIndex        =   7
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
         Left            =   5775
         TabIndex        =   8
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
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   " "
         Top             =   890
         Width           =   2595
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
         Width           =   4065
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
         Width           =   4065
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
      Begin VB.Line Line8 
         X1              =   10730
         X2              =   10730
         Y1              =   1785
         Y2              =   4950
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   7875
         Picture         =   "FORPEDCLI.frx":540C
         ToolTipText     =   "Recepción Pedidos Electrónicos"
         Top             =   210
         Visible         =   0   'False
         Width           =   480
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
         TabIndex        =   85
         ToolTipText     =   "Doble-Click para 'Reparación'"
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
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
         TabIndex        =   80
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
         Left            =   5775
         TabIndex        =   37
         Top             =   975
         Width           =   1695
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
         TabIndex        =   35
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         Left            =   5775
         TabIndex        =   32
         Top             =   195
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
         TabIndex        =   31
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
Attribute VB_Name = "FORPEDCLI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHIPRE$

Private Sub command1_Click()
   Call SELECHO("DEUDOR1")
   NCLIE% = Val(TRIM$(VALACT1$("DEUDOR1")))
   Text1.TEXT = TRIM$(Str$(NCLIE%))
   If NCLIE% > 0 Then
     If REGICLI&(NCLIE%) > 0 Then
       Call CARDACLI
       On Error Resume Next
       Command6.SetFocus
       On Error GoTo 0
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
3   Call SELECHO("PEDENCA/Indice General de Pedidos (Pendientes-Cumplidos-Anulados)")
    NPX& = Val(VALACT1$("PEDENCA"))
    If NPX& < 1 Then GoTo 99
    '
    TIDOCUM$ = "Pedido de Cliente"
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
    '
    Screen.MousePointer = 1
    Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a este Pedido.\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
    On Error Resume Next
    OPPEDIDO.Command46.SetFocus
    On Error GoTo 0
    OPPEDIDO.Show 1
    GoTo 99
    '
    '16/03/07 (OT-07-0113)
5   Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
    '16/03/07 (FIN)
'5   For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Cls
'      EDITFORM.Picture1(KKI%).DrawWidth = 3
'      EDITFORM.Picture1(KKI%).DrawStyle = 0
'      EDITFORM.Picture1(KKI%).FillStyle = 1
'    Next KKI%
'    '
'    REDOCU$ = REGACT$("PDOCLST")
'    URE& = ULTREG&("PDOCSPL")
'    LI& = 0: LS& = URE&
'6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
'    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
'    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
'    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
'    GoTo 8
'    '
'7   Screen.MousePointer = 1
'    Call MENSERR(24, "Documento no Registrado")
'    GoTo 99
'    '
'8   TPSP$ = ""
'    Screen.MousePointer = 11
'    While L& > 1
'      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'      If DCC& < DOCUNU& Then GoTo 9
'      L& = L& - 1
'    Wend
'    '
'9   CAPAGINAS% = 0
'    For KKL& = L& To URE&
'      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
'      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
'      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
'        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
'      End If
'    Next KKL&
'    '
'95  SPOOLSTRING$ = TPSP$
'    For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Cls
'      EDITFORM.Picture1(KKI%).Top = 0
'      EDITFORM.Picture1(KKI%).Refresh
'      EDITFORM.Picture1(KKI%).Height = 7170
'      EDITFORM.Picture1(KKI%).Width = 11700
'      EDITFORM.Picture1(KKI%).DrawWidth = 3
'      EDITFORM.Picture1(KKI%).DrawStyle = 0
'      EDITFORM.Picture1(KKI%).FillStyle = 1
'    Next KKI%
'    ALTUPAGINA = 7170
'    TOPEPAGINA = 0
'    PAGINACTIVA% = 0
'    CAPAGINAS% = 0
'    '
'10  PPXX% = InStr(TPSP$, Chr$(255))
'    If PPXX% > 0 Then
'      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
'      TPSP$ = Mid$(TPSP$, PPXX% + 1)
'      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
'      '
'      If Len(COPRIAT$) > 0 Then
'        COPRI% = Asc(Left$(COPRIAT$, 1))
'        ATRIP$ = Mid$(COPRIAT$, 2)
'        Call SHOWCOMMAND(COPRI%, ATRIP$)
'      End If
'      GoTo 10
'      '
'    End If
'    '
'    Screen.MousePointer = 1
'    PAGINACTIVA% = 0
'    For KKI% = 0 To 5
'       EDITFORM.Picture1(KKI%).Visible = False
'    Next KKI%
'    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
'    EDITFORM.Picture1(PAGINACTIVA%).Refresh
'    EDITFORM.VScroll1.Min = 0
'    EDITFORM.VScroll1.Value = 0
'    EDITFORM.VScroll1.Max = 1
'    If ALTUPAGINA > EDITFORM.Frame1.Height Then
'       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
'    End If
'    EDITFORM.Command3.Enabled = False
'    EDITFORM.Command5.Enabled = False
'    EDITFORM.Show 1
    '
    GoTo 3
    '
99  Command12.Enabled = True
    '
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   OPPEDIDO.Show 1
   NC% = Val(Text1)
   If MONEMI% < 1 Or MONEMI% > 2 Then MONEMI% = 1
   If MONEMI% = 2 Then
       MONEMIS = "Dólares"
     Else
       MONEMIS = "Pesos"
   End If
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "DOSIVA") > 0 Then
     If TRIM$(TEXT21(2)) = "" Or TRIM$(Text44) = "" Then
       Call MENSERR(24, "Falta Vendedor / Comisión")
       On Error Resume Next
       TEXT21(2).SetFocus
       On Error GoTo 0
       GoTo 99
     End If
   End If
   '
   Call GRAPEDIDO(OKX%)
   If OKX% > 0 Then
     CONPRIPE.Show (1)
     If CONPRIPE.Label1.Caption = "OK" Then
        '
        CANCELPRINT% = 0
        If CONPRIPE.Check2.Value = 1 Then
           CANCELPRINT% = 1  ' IMPRESION FANTASMA
        End If
        '
        If CONPRIPE.Check1.Value = 1 Then
           COMUNI ("La Opción REPLICA POR MAIL No Esta Disponible Para Esta Versión")
        End If
        '
     End If
     Call PRIPEDIDO
     Call GENERODES      ' genera orden de despacho
     CANCELPRINT% = 0
     Call BLANFORMU
     Call SETULTREG("!CARDETPE", 0)
     Call CIERRARCH("*.*")
     Text1.TEXT = ""
     On Error Resume Next
     Text1.SetFocus
     On Error GoTo 0
   End If
99 Command14.Enabled = True
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
   Call MENSERR(24, "No Disponible por el Momento")
End Sub

Private Sub Command18_Click()
   MARSELPRESU.Visible = False
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Command19_Click()
   '
   NOFRESHPRE% = 0
   If Left$(UCase$(Label11), 6) = "REPARA" Then GoTo 50
   '
   UREPRE& = ULTREG&("PRESUPU")
   HOII% = HOY(HOS$)
   JJ& = 0: ATX% = 0: COEFCON = 1
   For KKI% = 1 To List2.ListCount
     If List2.Selected(KKI% - 1) = True Then
       REPRESU& = Val(Mid$(List2.List(KKI% - 1), 121, 8))
       If REPRESU& > 0 Then
         If REPRESU& <= UREPRE& Then
           NPRESX& = CVS(Left$(REGLEIDO$("PRESUPU", REPRESU&), 4))
           For REPRES1& = REPRESU& To UREPRE&
             XX$ = REGLEIDO$("PRESUPU", REPRES1&)
             If CVS(Left$(XX$, 4)) <> NPRESX& Then GoTo 19
             If CVI(Mid$(XX$, 113, 2)) > 0 Then GoTo 19 ' FECHA O.VENTA
             '
             ZZ$ = REGBLAN$("!CARDETPE")
             CIXI% = CVI(Mid$(XX$, 69, 2))
             Call REPLA(ZZ$, MKI$(CIXI%), 1, 2)
             PRELI = CVS(Mid(XX$, 71, 4))
             DESCUTEX$ = TRIM$(Mid$(XX$, 75, 14))
             CPPP = (100 - COEFDESCU(DESCUTEX$)) / 100
             If CPPP < 1 Then ATX% = 1
             '
             MONEMA% = Asc(Mid$(XX$, 68, 1))
             If JJ& < 1 Then
               MONEMI% = Asc(Mid$(XX$, 68, 1))
               If MONEMI% < 1 Then MONEMI% = 1
               If MONEMI% = 2 Then MONEMIS = "Dólares" Else MONEMIS = "Pesos"
             End If
             COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
             CANTIPRES = (Int(10 * CVS(Mid$(XX$, 101, 4)) + 0.5)) / 10
             If CANTIPRES < 0.05 Then CANTIPRES = 1
             '
             Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 39, 4)
             Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 51, 4)
             Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 65, 4)
             Call REPLA(ZZ$, MKS$(CANTIPRES), 47, 4)   ' CANTIDAD
             DESCUTEX$ = TRIM$(Mid$(XX$, 75, 14))
             'Call REPLA(ZZ$, MKS$(COEFDESCU(DESCUTEX$)), 69, 4)
             Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 73, 4)
             Call REPLA(ZZ$, MKI$(HOII%), 63, 2)
             Call REPLA(ZZ$, MKS$(REPRES1&), 125, 4)
             JJ& = JJ& + 1
             Call GRAREG("!CARDETPE", ZZ$, JJ&)
             NOFRESPRE% = 1
           Next REPRES1&
         End If
       End If
     End If
19 Next KKI%
   Call SETULTREG("!CARDETPE", JJ&)
   Call CIERRARCH("!CARDETPE")
   GoTo 90
   '
50 UREPRE& = ULTREG&("PRESREPA")
   HOII% = HOY(HOS$)
   TEXT21(0) = "0"
   JJ& = 0: ATX% = 0
   For KKI% = 1 To List2.ListCount
     If List2.Selected(KKI% - 1) = True Then
       REPRESU& = Val(Mid$(List2.List(KKI% - 1), 121, 8))
       If REPRESU& > 0 Then
         If REPRESU& <= UREPRE& Then
           XX$ = REGLEIDO$("PRESREPA", REPRESU&)
           ZZ$ = REGBLAN$("!CARDETPE")
           CIXI% = CVI(Mid$(XX$, 69, 2))
           Call REPLA(ZZ$, MKI$(CIXI%), 1, 2)
           PRELI = CVS(Mid(XX$, 117, 4))
           'DESCUTEX$ = TRIM$(Mid$(XX$, 75, 14))
           CPPP = 1  ' (100 - COEFDESCU(DESCUTEX$)) / 100
           'If CPPP < 1 Then ATX% = 1
           '
           MONEMA% = 1   ' REPARACIONES SIEMPRE EN PESOS
           MONEMI% = 1
           MONEMIS = "Pesos"
           COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
           '
           Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 39, 4)
           Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 51, 4)
           Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 65, 4)
           Call REPLA(ZZ$, MKS$(1), 47, 4)   ' CANTIDAD
           DESCUTEX$ = TRIM$(Mid$(XX$, 75, 14))
           'Call REPLA(ZZ$, MKS$(COEFDESCU(DESCUTEX$)), 69, 4)
           Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 73, 4)
           Call REPLA(ZZ$, MKI$(HOII%), 63, 2)
           Call REPLA(ZZ$, MKS$((-1) * REPRESU&), 125, 4)
           JJ& = JJ& + 1
           Call GRAREG("!CARDETPE", ZZ$, JJ&)
         End If
       End If
     End If
   Next KKI%
   Call SETULTREG("!CARDETPE", JJ&)
   Call CIERRARCH("!CARDETPE")
   '
90 MARSELPRESU.Visible = False
   '
   If ATX% > 0 Then
      PORDESCU.TEXT = ""
      Call COMUNI("ATENCION: Descuento Aplicado según Presupuesto Previo !!!")
   End If
   Call CARDETPEDI(1)
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
   Static ARGU$
   Command20.Enabled = False
10 ARGU$ = UCase$(TRIM$(InputBox$("Buscar", "Búsqueda Cliente por Texto", ARGU$)))
   If ARGU$ <> "" Then
     Screen.MousePointer = 11
     PUNCLI$ = ""
     Call COPYSTRU("DEUDOR1", "SELECLI")
     FIN& = ULTREG&("DEUDOR2")
     For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = UCase$(REGLEIDO$("DEUDOR2", U&))
       If InStr(X$, ARGU$) > 0 Then
         X$ = REGLEIDO$("DEUDOR1", U&)
         PUNCLI$ = PUNCLI$ + Left$(X$, 2)
         JJ& = JJ& + 1
         Call GRAREG("!SELECLI", X$, JJ&)
       End If
     Next U&
     '
     FIN& = ULTREG&("BADACLIE")
     For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = UCase$(REGLEIDO$("BADACLIE", U&))
       If InStr(X$, ARGU$) > 0 Then
         NCX$ = Left$(REGLEIDO$("BADACLIE", U&), 2)
         For KKI% = 1 To Len(PUNCLI$) Step 2
           If Mid$(PUNCLI$, KKI%, 2) = NCX$ Then GoTo 19 ' ya esta
         Next KKI%
         UK& = REGBUS&("DEUDOR1", 1, NCX$)
         If UK& > 0 Then
           X$ = REGLEIDO$("DEUDOR1", UK&)
           JJ& = JJ& + 1
           Call GRAREG("!SELECLI", X$, JJ&)
         End If
       End If
19   Next U&
     '
     Call SETULTREG("!SELECLI", JJ&)
     '
     Screen.MousePointer = 1
     If JJ& < 1 Then
         Call MENSERR(24, "No se Encontraron Coincidencias")
         GoTo 10
       ElseIf JJ& = 1 Then
         Text1 = TRIM$(Str$(CVI(Left$(REGLEIDO$("!SELECLI", 1), 2))))
         Call CARDACLI
       Else
         Call FRESHECHO("!SELECLI")
         Call SELECHO("!SELECLI")
         NCLIX$ = VALACT1$("!SELECLI")
         If NCLIX$ <> "" Then
           Text1 = NCLIX$
           Call CARDACLI
         End If
     End If
   End If
   Command20.Enabled = True
End Sub

Private Sub Command22_Click()
    Command22.Enabled = False
    Screen.MousePointer = 11
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
    Screen.MousePointer = 1
    '
    CONPEDID.Show 1
    NC% = Val(Text1) '\\\OT-06-0242-WAR-20/06/06///
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
    RESCUDEU.Text1.TEXT = Text1.TEXT
    PROLLAMA$ = "FORPEDCLI"
    Hide
    RESCUDEU.Show
    Command4.Enabled = True
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
   Call SELPRESUPEN
End Sub
'
Sub SELPRESUPEN()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   '
   UREPRE& = ULTREG&("PRESUPU")
   Screen.MousePointer = 11
   List2.Clear
   NCLIE% = Val(Text1.TEXT)
   If UCase$(Left$(Label11, 6)) = "REPARA" Then GoTo 50
   '
10 RFF$ = RECFILT$("PRESUPU", 3, MKI$(NCLIE%))
   For KKI% = 1 To Len(RFF$) Step 4
      REPRESU& = CVS(Mid$(RFF$, KKI%, 4))
      If REPRESU& > 0 Then
        If REPRESU& <= UREPRE& Then
12        XX$ = REGLEIDO$("PRESUPU", REPRESU&)
          If CVI(Mid$(XX$, 7, 2)) <> NCLIE% Then
            XX$ = REGLEIDO$("PRESUPU", 1)
            Call GRAREG("PRESUPU", XX$, 1)
            Call CIERRARCH("PRESUPU")
            GoTo 10
          End If
          If CVI(Mid$(XX$, 113, 2)) < 1 Then
            TTXX$ = "Presup.No: " + TRIM$(CSTRING$(Left$(XX$, 4), 3, 8, 0, 2))
            FEX = CVI(Mid$(XX$, 5, 2))
            TTXX$ = TTXX$ + " - " + FECHATEX$(FEX)
            If TRIM$(Mid$(XX$, 9, 16)) <> "" Then
              TTXX$ = TTXX$ + " - Ref: " + TRIM$(Mid$(XX$, 9, 16))
            End If
            CIXI% = CVI(Mid$(XX$, 69, 2))
            '
            If CVS(Left$(XX$, 4)) = NPX& Then GoTo 19
            '
            NPX& = CVS(Left$(XX$, 4))
            CAPRESU& = CVS(Mid$(XX$, 101, 4))
            If CAPRESU& < 1 Then CAPRESU& = 1
            Call RECODOSI(CIXI%, CIXJ%, NPX&)
            If CIXJ% <> CIXI% Then
              Call REPLA(XX$, MKI$(CIXJ%), 69, 2)
              Call GRAREG("PRESUPU", XX$, REPRESU&)
              GoTo 12
            End If
            '
            If CVS(Left$(REGLEIDO$("PRESUPU", REPRESU& + 1), 4)) = NPX& Then
                TTXX$ = TTXX$ + " - " + "VARIOS"
               Else
                TTXX$ = TTXX$ + " - " + TRIM$(CODEXT$(CIXI%)) + " (" + TRIM$(Str$(CAPRESU&)) + ") " + DESCRIT0$(CIXI%)
            End If
            TTXX$ = AJUSTI$(TTXX$, 128)
            Call REPLA(TTXX$, Str$(REPRESU&), 121, 8)
            List2.AddItem TTXX$
          End If
        End If
      End If
19 Next KKI%
   GoTo 90
   '
50 UREPRE& = ULTREG&("PRESREPA")
   Screen.MousePointer = 11
   RFF$ = RECFILT$("PRESREPA", 4, MKI$(NCLIE%))
   For KKI% = 1 To Len(RFF$) Step 4
      REPRESU& = CVS(Mid$(RFF$, KKI%, 4))
      If REPRESU& > 0 Then
        If REPRESU& <= UREPRE& Then
          XX$ = REGLEIDO$("PRESREPA", REPRESU&)
          If CVI(Mid$(XX$, 121, 2)) > 0 Then
            If Abs(CVS(Mid$(XX$, 125, 4))) < 0.5 Then
              TTXX$ = "Reparación No: " + TRIM$(CSTRING$(Left$(XX$, 4), 3, 8, 0, 2))
              TTXX$ = TTXX$ + " - " + Mid$(XX$, 5, 60)
              TTXX$ = AJUSTI$(TTXX$, 128)
              Call REPLA(TTXX$, Str$(REPRESU&), 121, 8)
              List2.AddItem TTXX$
            End If
          End If
        End If
      End If
   Next KKI%
   '
90 Screen.MousePointer = 1
   '
   If List2.ListCount < 1 Then
      Call SETULTREG("!CARDETPE", 0)
      Call CIERRARCH("!CARDETPE")
      Call CARDETPEDI(1)
      Exit Sub
   End If
   '
   MARSELPRESU.Visible = True
   Command19.SetFocus
   '
End Sub
'
Private Sub Command7_Click()
    Command7.Enabled = False
    Call CIERRARCH("*.*")
    MODX% = ALTERNA%("Listas de Precios\Condiciones de Pago")
    If MODX% = 1 Then
         LPREU% = Val(TEXT21(0))
         If LPREU% < 1 Then LPRE% = 1
         ARCHIPRU$ = TRIM$(Str$(LPREU%))
         While Len(ARCHIPRU$) < 3: ARCHIPRU$ = "0" + ARCHIPRU$: Wend
         ARCHIPRU$ = "LIPRE" + ARCHIPRU$
         Call FRESHECHO(ARCHIPRU$)
         '
         If EXISTE%("LISPRE02.EXE") > 0 Then
              Call CONECRUN("LISPRE02", "Listas de Precios de Venta")
            Else
              Call CONECRUN("ADMIPRE", "Listas de Precios de Venta")
         End If
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
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text9.TEXT = VALACT1$("SELFECHA")
   End If
   FF% = FECHANUM(Text9.TEXT)
End Sub

Private Sub Eco21_GotFocus(Index As Integer)
   On Error Resume Next
   TEXT21(Index).SetFocus
   On Error GoTo 0
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  While WindowState <> 0
    WindowState = 0
    OPENFORMS = DoEvents
    AppActivate Caption
    SendKeys "%" & Chr$(32) & Chr$(13), True
  Wend
End Sub


Private Sub Form_Load()
    '
    Frame4.Height = 1545
    OPENFORMS = DoEvents
    VERANTER$ = "PEDIDO01"
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    
    If InStr(UCase$(EPAC$), "DOSIVA") > 0 Then
      Label11.Visible = True
      Command22.Visible = True
      Command12.Visible = False
    End If
    '
    If EXISTE%("REPARA02.EXE") > 0 Then
      Label11.Visible = True
    End If
    Call GRATITFOR(Caption)
    '
    If IDENTER$ <> "" Then
      If CONTROL$("CAPRINT", IDENTER$) = "SI" Then
        Timer2.Enabled = True
      End If
    End If
    '
'    If EXISTE%("REPEDREM.EXE") > 0 Then
'      Image2.Visible = True
'    End If
    '
    If InStr(UCase$(EPAC$), "VDR") > 0 Then
      Image2.Visible = True
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
    HOII% = HOY(HOS$)
    Text9.TEXT = HOS$
    Text11.TEXT = TRIM$(Str$(NUPEDIDO&))
    '
    CMO$ = TRIM$(CONTROL$("", "CODPECLI"))
    If CMO$ <> "" Then Label10.Caption = CMO$
    DoEvents
    '
    If EXISTE%(LUDAT$ + "INDIPEDI.DAT") > 0 Then
      Screen.MousePointer = 11
      FIN& = ULTREG&("PEDENCA")
      For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        XX$ = REGLEIDO$("PEDENCA", U&)
        If Mid$(XX$, 267, 1) = Chr$(9) Then
          DEPED$ = Mid$(XX$, 5, 48)
          If Left$(DEPED$, 6) <> "(Anul)" Then
            DEPED$ = "(Anul) " + DEPED$
          End If
          Call REPLA(XX$, DEPED$, 5, 48)
          Call GRAREG("PEDENCA", XX$, U&)
        End If
      Next U&
      Call CIERRARCH("PEDENCA")
      Call CIERRARCH("*.*")
      On Error Resume Next
      Kill LUDAT$ + "INDIPEDI.DAT"
      On Error GoTo 0
      Screen.MousePointer = 1
    End If
    '
    Call SETULTREG("!CARDETPE", 0)
    '
    If InStr(UCase$(EPAC$), "DOSIVA") > 0 Then
      If CONTROL$("PEDCLIEN", "MARPREPA") <> "SI" Then
        Show
        DoEvents
        Call PONEMARPREP
      End If
    End If
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
         Call CALTOPED
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

Private Sub Image2_DblClick()
   Call CIERRARCH("*.*")
   'Call CONECRUN("REPEDREM", "Recepción Pedidos Remotos por Mail")
   '\\\OT-05-0553-WAR-01/08/05///
   If VALICLIE% > 0 Then
     Call MENSERR(2, "Imposible Importar Pedidos\con Cliente Activo.")
     Exit Sub
   End If
   FOSEARIM.Show
   '\\\OT-05-0553-WAR-FIN///
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


Private Sub Label21_dblClick(Index As Integer)
   If Index <> 0 Then Exit Sub
   If DERACCE%("ALTACOPE", "Alta de Artículo desde Pedidos") < 2 Then Exit Sub
   If TRIM$(ECOARCH$("LISTAS", MKI$(Val(TEXT21(Index).TEXT)))) = "" Then Exit Sub
        '
        VDEF$ = REGBLAN$("MASTER") + String$(128, 0)
        Call REPLA(VDEF$, MKI$(1 + NUMAS%), 1, 2)
        Call REPLA(VDEF$, "U.", 109, 2)
        Call REPLA(VDEF$, Chr$(1), 93, 1)
        Call REPLA$(VDEF$, Eco21(0), 137, 25)
10      OBX$ = OBJPLA$("ALTACOPED", VDEF$)
        If OBX$ = "" Then GoTo 99
        VDEF$ = OBX$
        CODINUE$ = TRIM$(Mid$(OBX$, 77, 16))
        '
        If Len(CODINUE$) < 3 Then
           Call MENSERR(24, "Ingreso no Válido.\Longitud Código Menor que 3 Caracteres.")
           GoTo 10
        End If
        '
        If Len(CODINUE$) > 15 Then
           Call MENSERR(24, "Ingreso no Válido.\Longitud Código Mayor que 15 Caracteres.")
           GoTo 10
        End If
        '
        If CODINT%(CODINUE$) <> 0 Then
           Call MENSERR(24, "Ingreso no Válido.\Código Pre-Existente.")
           GoTo 10
        End If
        '
        If TRIM$(Mid$(OBX$, 5, 64)) = "" Then
           Call MENSERR(24, "Ingreso no Válido.\Falta Descripcion")
           GoTo 10
        End If
        DENUE$ = TRIM$(Mid$(OBX$, 5, 64))
        '
        TIMA% = Asc(Mid$(OBX$, 93, 1))
        If ECOARCH$("TATIMAT", Chr$(TIMA%)) = "" Then
           Call MENSERR(24, "Tipo de Material no Válido")
           GoTo 10
        End If
        '
        UNID$ = Mid$(OBX$, 109, 2)
        If ECOARCH$("UNIMED", UNID$) = "" Then
           Call MENSERR(24, "Unidad de Medida no Válida")
           GoTo 10
        End If
        '
        PRE1# = CVD(Mid$(OBX$, 129, 8))
        If PRE1# <= 0.005 Then
           Call MENSERR(24, "Precio no Válido")
           GoTo 10
        End If
        '
        Screen.MousePointer = 11
        'ACA CODIGO PARA GRABACIÓN
        '
        Call BLOQARCH("MASTER")
        NMSI% = NUMAS% + 1
        REMAS& = NUMAS% + 2
        Call GRAREG("MASTER", OBX$, REMAS&)
        '
        X$ = REGLEIDO$("MASTER", 1&)
        Call REPLA(X$, MKI$(NMSI%), 1, 2)
        Call GRAREG("MASTER", X$, 1)
        '
        Call CIERRARCH("MASTER")
        Call BAJALDISCO
        Call REPUNCODIG
        '
        EL1$ = REGBLAN$("ECOMAST")
        Call REPLA(EL1$, CODINUE$, 1, 16)
        Call REPLA(EL1$, Mid$(OBX$, 5, 64), 17, 46)
        Call REPLA(EL1$, MKI$(NMSI%), 63, 2)
        Call REGAPP("ECOMAST", EL1$)
        Call CIERRARCH("ECOMAST")
        Call FILESORT("ECOMAST", "", 1, 1, "ASC")
        Call BAJALDISCO
        '
        NVII1% = 1 + NUINV%
        EL1$ = REGLEIDO$("INVERT", 1)
        Call REPLA(EL1$, String$(16, 0), 1, 16)
        Call REPLA(EL1$, MKI$(NVII1%), 17, 2)
        Call GRAREG("INVERT", EL1$, 1)
        EL1$ = REGBLAN$("INVERT")
        Call REPLA(EL1$, CODINUE$, 1, 16)
        Call REPLA(EL1$, MKI$(NMSI%), 17, 2)
        Call REGAPP("INVERT", EL1$)
        Call CIERRARCH("INVERT")
        Call FILESORT("INVERT", "", 1, 1, "ASC")
        Call BAJALDISCO
        '
        Call FRESHECHO("ECOMAST")
        Call FRESHECHO("MASTER")
        Call BAJALDISCO
        OKX% = 1
        '
        CI% = NMSI%
        LISTACTIVA& = XVALO(TEXT21(0))
        XX$ = REGLEIDO$("LISTAS", LISTACTIVA&)
        CADECI% = Asc(Mid$(XX$, 49, 1)) 'SACO LA CANTIDAD DE DECIMALES DE LA LISTA
        PRELI# = ROUND(PRE1#, CADECI%)
        PREX$ = "  " + TRIM$(CSTRING$(MKD$(PRELI#), 3, 12, CADECI%, 2))
        DE$ = DESCRIT0$(CI%)
        NXX$ = REGBLAN$("LIPREVEN")
        Call REPLA$(NXX$, MKI$(CI%), 1, 2)
        Call REPLA$(NXX$, DE$, 3, 42)
        Call REPLA(NXX$, PREX$ + "  ", 43 - Len(PREX$), 2 + Len(PREX$))
        Call REPLA$(NXX$, MKD$(PRELI#), 62, 8)
        Call REGAPP(ARCHIPRE$, NXX$)
        '
        Call INDEXARCH(ARCHIPRE$, 1&)
        Call CIERRARCH(ARCHIPRE$)
        Call CIERRARCH("*.*")
        Call BAJALDISCO
        Call FRESHECHO(ARCHIPRE$)
        '
        Call COMUNI("Alta Código '" + CODINUE$ + "' en Maestro\y Lista de Precios Activa - Completo")
        '
99      Screen.MousePointer = 1
        '
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
   Call CARDETPEDI(2)
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
   Call BLANFORMU
End Sub

Private Sub Text1_DblClick()
  Call SELECHO("DEUDOR1")
  NCLIE% = Val(TRIM$(VALACT1$("DEUDOR1")))
  Text1.TEXT = TRIM$(Str$(NCLIE%))
  If NCLIE% > 0 Then
    If REGICLI&(NCLIE%) > 0 Then
      Call CARDACLI
      Command6.SetFocus
    End If
  End If
End Sub

Private Sub Text1_GotFocus()
   If MARCONOTA.Visible = True Then
      Command6.SetFocus
      Exit Sub
    ElseIf MARSELPRESU.Visible = True Then
      Command19.SetFocus
      Exit Sub
   End If
   Text1.TEXT = TRIM$(Text1.TEXT)
   Text1.SelStart = 0
   Text1.SelLength = Len(Text1.TEXT)
   MARCONOTA.Visible = False
   MARSELPRESU.Visible = False
   '
   If TRIM$(Text1) = "" Then
     Label11.ForeColor = RGB(0, 0, 255)
     Label11 = "Venta"
     Label11.ToolTipText = "Doble-Click para 'Reparación'"
   End If
   '
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If VALICLIE% > 0 Then
      Call BLANFORMU
      If Val(Text1.TEXT) > 0 Then
        If Val(Text1.TEXT) <= 32767 Then
          NC% = Val(Text1)
          If REGICLI&(NC%) > 0 Then
            If MARCONOTA.Visible = False Then
              If MARSELPRESU.Visible = False Then
                Call CARDACLI
              End If
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
   NPP& = NUPEDIDO&
   If Val(NUMEPRE$) < NPP& Then
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
        If MONELISTA%(Val(TEXT21(Index))) = 2 Then
             MONEMIS = "Dólares"
           Else
             MONEMIS = "Pesos"
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
   Label18 = ""
   List1.Clear
   Call SETULTREG("!CARDETPE", 0)
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
   MARCONOTA.Visible = False
   HOII% = HOY(HOS$)
   Text9.TEXT = HOS$
   Text11.TEXT = TRIM$(Str$(NUPEDIDO&))
   '
   NOFRESHPRE% = 0
   '
End Sub
'
Sub CARDACLI()
    '
    Screen.MousePointer = 11
    MARCONOTA.Visible = False
    MARSELPRESU.Visible = False
    If CONTROL$("ADMIPED", "VERLICRE") = "SI" Then VERLICRE% = 1
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    NC% = Val(Text1.TEXT)
    If NC% > 0 Then
      '
      If InStr(EPAC$, "DOSIVA") > 0 Then
        PVCL% = PIVACLI%(NC%)
        If PVCL% <> 0 And PVCL% <> 5 Then
          PRVC$ = TRIM$(PROVCLI$(NC%))
          If PRVC$ = "" Or DEPROVI$(PRVC$) = "" Then
            Call COMUNI("Imposible Registrar O/Venta.\Falta Código de Provincia del Cliente.")
            Text1 = ""
            GoTo 99
          End If
        End If
      End If
      '
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
        '
        TEXT21(0).TEXT = Str$(LIPRCLI%(NC%))
        TEXT21(1).TEXT = Str$(CPAGCLI%(NC%))
        TEXT21(2).TEXT = Str$(VENDCLI%(NC%))
        '
        If Val(TEXT21(2).TEXT) < 1 Then
           Eco21(2).TEXT = "GERENCIA"
        End If
        '
        PORDESCU.TEXT = DESCUCLI$(NC%)
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
        LENTREGA(0) = Left$(LENTREG$, 32)
        LENTREGA(1) = Mid$(LENTREG$, 33, 32)
        LENTREGA(2) = Mid$(LENTREG$, 65, 32)
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
        '
        If VERLICRE% = 1 Then
          LMCR# = LIMICRE(NC%)
          SDEU# = SALDCLI#(NC%)
          CHPA# = CHEPEACRE#(NC%)
          If SDEU# + CHPA# > LMCR# Then
            Call COMUNI("Este Pedido no Podrá Facturarse.\  \Límite de Crédito Asignado ya está Superado.\  \Informe al Supervisor de Ventas.")
          End If
        End If
        '
      End If
    End If
99  Screen.MousePointer = 1
    '
End Sub
   '
Function NUPEDIDO&()
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
99 NUPEDIDO& = NPP&
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

Private Sub TICAMBIS_LostFocus()
   TICAMBIS = FORMATNUM(Val(TICAMBIS), "F9.4")
End Sub

Private Sub Timer1_Timer()
   '
   Call CALTOPED
   '
End Sub

Private Sub Timer2_Timer()
  '
  Static FIDAPRIN#
  '
  FDP# = FIDATIM#("IPEDPRIN")
  If FDP# <> FIDAPRIN# Then
    FIDAPRIN# = FDP#
    RFF$ = RECFILT$("IPEDPRIN", 3, MKI$(0))
    '
    Screen.MousePointer = 11
    For UL& = 1 To Len(RFF$) Step 4
      U& = CVS(Mid$(RFF$, UL&, 4))
      YY$ = REGLEIDO$("IPEDPRIN", U&)
      TDOC$ = UCase$(Left$(YY$, 2))
      NPX& = Val(Mid$(YY$, 3, 6))
      '
      If NPX& > 0 Then
        TIDOCUM$ = "Pedido de Cliente"
        If TDOC$ = "OD" Then TIDOCUM$ = "Orden de Despacho"
        '
        NUDOCU$ = TRIM$(Str$(NPX&))
        '
        ' esto para forzar que haya bajado al disco
        Call BLOQARCH("PDOCSPL")
        Call LIBARCH("PDOCSPL")
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
        GoTo 49
        '
5       REDOCU$ = REGACT$("PDOCLST")
        URE& = ULTREG&("PDOCSPL")
        LI& = 0: LS& = URE&
6       E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 49
        L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 49
        DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
        If DCC& < DOCUNU& Then LI& = L&: GoTo 6
        If DCC& > DOCUNU& Then LS& = L&: GoTo 6
        '
8       TPSP$ = ""
        Screen.MousePointer = 11
        While L& > 1
          DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
          If DCC& < DOCUNU& Then GoTo 9
          L& = L& - 1
        Wend
        '
9       For KKL& = L& To URE&
          TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
          If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 10
          If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
            TPSP$ = TPSP$ + Mid$(TBUF$, 5)
          End If
        Next KKL&
        '
10      PPXX% = InStr(TPSP$, Chr$(255))
        If PPXX% > 0 Then
          COPRIAT$ = Left$(TPSP$, PPXX% - 1)
          TPSP$ = Mid$(TPSP$, PPXX% + 1)
          '
          If Len(COPRIAT$) > 0 Then
            COPRI% = Asc(Left$(COPRIAT$, 1))
            ATRIP$ = Mid$(COPRIAT$, 2)
            Call PRINTCOMMAND(COPRI%, ATRIP$)
          End If
          GoTo 10
        End If
        '
      End If
      '
      YY$ = REGLEIDO$("IPEDPRIN", U&)
      HOII% = HOY(HOS$)
      Call REPLA(YY$, MKI$(HOII%), 57, 2)
      Call REPLA(YY$, Time$, 59, 6)
      Call GRAREG("IPEDPRIN", YY$, U&)
      Call CIERRARCH("IPEDPRIN")
      '
49  Next UL&
    Screen.MousePointer = 1
    '
  End If
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
    If CANCELPRINT% <> 1 Then  'Se verifica Impresion Fantasma
       If CONTROL$("PEDCLIEN", "SUPRINT") = "SI" Then
          CANCELPRINT% = 1
       End If
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
    '
    Screen.MousePointer = 11
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
            PRELI$ = Str$(CVS(Mid$(XX$, 39, 4)))
            PREUX$ = Str$(CVS(Mid$(XX$, 51, 4)))
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
    CODPARAM$(10) = "IMP-IVA": DOCPARAM$(10) = Label13 'TOTIVA(3).TEXT
    CODPARAM$(11) = "IMP-TOTA": DOCPARAM$(11) = Total(3).TEXT
    CODPARAM$(12) = "POR-DESC": DOCPARAM$(12) = TRIM$(PORDESCU.TEXT)
    CODPARAM$(13) = "PERC-IVA": DOCPARAM$(13) = Text15
    '\\\OT-05-0560-WAR-03/08/05///
    CODPARAM$(14) = "IMP-IVA1": DOCPARAM$(14) = Label14
    '\\\OT-05-0560-WAR-FIN///
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
        If CANCELPRINT% <> 1 Then    'Se Verifica Impresion Fantasma
           If CONTROL$("ORDESPA", "SUPRINT") = "SI" Then
              CANCELPRINT% = 1
           End If
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
    Screen.MousePointer = 1
    '
End Sub

Private Sub Total_GotFocus(Index As Integer)
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Sub GRAPEDIDO(OKX%)
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
   '
   OKX% = 0
   If MARCONOTA.Visible = True Then Exit Sub
   If MARSELPRESU.Visible = True Then Exit Sub
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
   Call CIERRARCH("DESPAPED")
   '
   Call BLOQARCH("PEDENCA")
   Call BLOQARCH("PEDDETA")
   Call BLOQARCH("PEDDETP")
   Call BLOQARCH("DESPAPED")
   '
   NPP& = NUPEDIDO&
   If NPP& > Val(Text11.TEXT) Then
      Call CIERRARCH("PEDENCA")
      Call CIERRARCH("PEDDETA")
      Call CIERRARCH("PEDDETP")
      Call CIERRARCH("DESPAPED")
      Call COMUNI("FLEXOFT ha Re-Numerado el Pedido\como Pedido Número '" + TRIM$(Str$(NPP&)) + "'.\  \Verifique y Vuelva a Grabar.")
      Text11.TEXT = TRIM$(Str$(NPP&))
      Exit Sub
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
   'If Val(Total(2).TEXT) < 0.01 Then  validado más arriba
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
   PDC$ = REGBLAN$("PEDENCA")
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
     Call REPLA(PDC$, Space$(16), 289, 16)
   Call REGAPP("PEDENCA", PDC$)
   REGENCA& = ULTREG&("PEDENCA")
   Call CIERRARCH("PEDENCA")
   '
   TTXYZ$ = Space$(76)
   HOII% = HOY(HOS$)
   Call REPLA(TTXYZ$, HOS$, 1, 8)
   Call REPLA(TTXYZ$, "Generación de Orden de Venta:", 11, 56)
   Call REPLA(TTXYZ$, USERTER$, 69, 24)
   Call REGAPP("ANOTAPED", MKS$(NROPE&) + TTXYZ$)
   '
   UREPRE& = ULTREG&("PRESUPU")
   UREPREP& = ULTREG&("PRESREPA")
   CAVER$ = TRIM$(Str$(NROPE&))
   While Len(CAVER$) > 6: CAVER$ = Mid$(CAVER$, 2): Wend
   While Len(CAVER$) < 6: CAVER$ = "0" + CAVER$: Wend
   IKITEM% = 0: JJ& = 0
   For U& = 1 To ULTREG&("!CARDETPE")
     X$ = REGLEIDO$("!CARDETPE", U&)
     CAN = CVS(Mid$(X$, 47, 4))
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
       IKITEM% = IKITEM% + 1
       IK$ = TRIM$(Str$(IKITEM%))
       CAVERI$ = CAVER$ + "-" + IK$
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
       CANASIG = CAPEDORI
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
       Call REPLA(Z$, MKS$(CAPEDORI), 25, 4)
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
         'VERIFICADO AL 31/10/05
         PRUNI = CVS(Mid$(PDC$, 83, 4))
         Call REPLA(PDC$, Chr$(1), 27, 1)
         Call REPLA(PDC$, String$(38, 0), 71, 38)
         Call REPLA(PDC$, MKS$(CAN), 75, 4)
         Call REPLA(PDC$, MKS$(CAN), 99, 4)
         Call REPLA(PDC$, MKS$(PRUNI), 83, 4)
         Call REGAPP("DESPAPED", PDC$)    ' aqui graba
       End If
       '
       REPRESU& = CVS(Mid$(X$, 125, 4))
       If REPRESU& > 0 Then
         If REPRESU& <= UREPRE& Then
           YY$ = REGLEIDO$("PRESUPU", REPRESU&)
           If CVI(Mid$(YY$, 7, 2)) = CVI(Mid$(PDC$, 7, 2)) Then   ' Coincide cliente
             If CVI(Mid$(YY$, 69, 2)) = CVI(Mid$(PDC$, 33, 2)) Then   ' Coincide Articulo
               Call REPLA(YY$, MKI$(FEMI%), 113, 2)
               Call REPLA(YY$, MKS$(NROPE&), 115, 4)
               Call GRAREG("PRESUPU", YY$, REPRESU&)
             End If
           End If
         End If
       End If
       '
       If REPRESU& < 0 Then
         REPRESU& = Abs(REPRESU&)
         If REPRESU& <= UREPREP& Then
           YY$ = REGLEIDO$("PRESREPA", REPRESU&)
           If CVI(Mid$(YY$, 67, 2)) = CVI(Mid$(PDC$, 7, 2)) Then   ' Coincide cliente
             If CVI(Mid$(YY$, 69, 2)) = CVI(Mid$(PDC$, 33, 2)) Then   ' Coincide Articulo
               Call REPLA(YY$, MKS$((-1) * NROPE&), 125, 4)
               Call GRAREG("PRESREPA", YY$, REPRESU&)
             End If
           End If
         End If
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
   Call CIERRARCH("PRESUPU")
   Call CIERRARCH("PRESREPA")
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
   Call CIERRARCH("PRESUPU")
   Call CIERRARCH("PRESREPA")
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
   Call CIERRARCH("DESPAPED")
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

Sub CARDETPEDI(modo%)
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
      Total(0).TEXT = ""
      Text1.TEXT = ""
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   '
20 If ULTREG&("!CARDETPE") < 1 Then
      Screen.MousePointer = 1
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   '
   Screen.MousePointer = 11   ' Registrando Pedido ...
   FIN& = ULTREG&("!CARDETPE")
   If FIN& > 0 Then
     Y$ = REGLEIDO$("!CARDETPE", FIN&)
     CI% = CVI(Left$(Y$, 2))
     If CI% < 1 Or CI% > NUMAS% Then
       Call SETULTREG("!CARDETPE", FIN& - 1)
       Call CIERRARCH("!CARDETPE")
       Call MENSERR(24, "Código No Reconocido o No Válido\Debe Correr Control de Consistencia de Maestro.")
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
       If modo% < 2 Then
         If ECOPRECIO% >= 0 Then
           Call MENSERR(24, "Artículo sin Precio")
           Call REPLA(X$, MKS$(0), 39, 4)
           Call GRAREG("!CARDETPE", X$, U&)
           modo% = 2
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
   List1.Refresh
   On Error Resume Next
   Command14.SetFocus
   On Error GoTo 0
   Screen.MousePointer = 1
   '
End Sub

Sub CARLISPE()
   FORMACAN$ = "I6"
10 List1.Clear
   FIN& = ULTREG&("!CARDETPE")
   For U& = 1 To FIN&
     X$ = REGLEIDO$("!CARDETPE", U&)
     TTXX$ = Space$(96)
     CI% = CVI(Left$(X$, 2))
     FENTRESOL% = CVI(Mid$(X$, 63, 2))
     CAN = CVS(Mid$(X$, 47, 4))
     If FORMACAN$ = "I6" Then
       If (Abs(CAN - Int(CAN))) >= 0.05 Then
         FORMACAN$ = "F6.1"
         GoTo 10
       End If
     End If
     PORDESCUIN = CVS(Mid$(X$, 69, 4))
     PRECARGA = CVS(Mid$(X$, 39, 4))
     PREUNI = PRECARGA * (100 - PORDESCUIN) / 100
     TOT# = CDbl(CAN * PREUNI)
     '
     Call REPLA(TTXX$, FORMATNUM$(CAN, FORMACAN$), 1, 6)
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
   '\\\OT-05-0560-WAR-03/08/05///
     Dim IMPTASIVA#()
     ReDim IMPTASIVA#(ULTREG&("GRUPIVA"))
     DESC0$ = TRIM$(PORDESCU.TEXT)
     COEF1# = 1 - COEFDESCU(DESC0$) / 100
   '\\\OT-05-0560-WAR-FIN///
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
       '\\\OT-05-0560-WAR-03/08/05///
       REG& = CIXXI% + 1
       REMA$ = REGLEIDO$("MASTER", REG&)
       GRIVA% = Asc(Mid$(REMA$, 3, 1))
       TODE1# = TOT# * COEF1#
       IMPTASIVA#(GRIVA%) = IMPTASIVA#(GRIVA%) + (Int(TODE1# * PORCEIVA(CIXXI%) + 0.5)) / 100
       Label13 = CSTRING$(MKD$(IMPTASIVA#(0)), 3, 11, 2, 2)
       Label14 = CSTRING$(MKD$(IMPTASIVA#(1)), 3, 11, 2, 2)
       '\\\OT-05-0560-WAR-FIN///
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

Private Sub TOTIVA_Change(Index As Integer)
   'If Index = 3 Then
   '  Total(3).TEXT = Str$(Val(Total(2).TEXT) + Val(TOTIVA(3).TEXT))
   'End If
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

Sub CALPERIB()
   '
   NCLIE% = Val(Text1)
   If NCLIE% < 1 Then Exit Sub
   '
   VARIDC% = 1: PVCLX% = PIVACLI%(NCLIE%)
   If PVCLX% = 0 Or PVCLX% = 4 Then VARIDC% = 2
   If PVCLX% = 5 Or PVCLX% = 7 Then VARIDC% = 3
   '
   NCLIE% = XVALO(Text1)
   TASAX = ALIPERBRU#(NCLIE%, MINIPER#, VARIDC%)
   '
   Text15 = ""
   BIMPON# = XVALO(Total(2))
   If VARIDC% = 2 Then BIMPON# = BIMPON# + Val(TOTIVA(3))
   If BIMPON# * TICAMONE(MONEMI%) >= MINIPER# Then
      PERSALTA# = BIMPON# * (TASAX / 100)
      Text15 = FORMATNUM$(PERSALTA#, "F11.2")
   End If
   '
End Sub
'
Sub PONEMARPREP()
   '
   If COMALTER%("FLEXOFT Debe Ahora Reorganizar\la Base de Datos de Pedidos\  \Todos los Usuarios Deben Salir\del Módulo de Pedidos de Clientes\\Cancelar\Re-Organizar Ahora") <> 2 Then
     Call FINAL("")
   End If
   '
   Call INDEXARCH("PEDDETA", 1)
   Call BLOQARCH("PEDDETA")
   Call BLOQARCH("PEDDETP")
   Call BLOQARCH("DESPAPED")
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("PEDDETA")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("PEDDETA", U&)
     NROPE& = CVS(Left$(XX$, 4))
     If NROPE& > 0 Then
       CAVER$ = TRIM$(Str$(NROPE&))
       While Len(CAVER$) > 6: CAVER$ = Mid$(CAVER$, 2): Wend
       While Len(CAVER$) < 6: CAVER$ = "0" + CAVER$: Wend
       '
       If Mid$(XX$, 113, 6) <> CAVER$ Then
         IKITEM% = 0
         RFF$ = RECFILT$("PEDDETA", 1, MKS$(NROPE&))
         For UL& = 1 To Len(RFF$) Step 4
           REPED& = CVS(Mid$(RFF$, UL&, 4))
           YY$ = REGLEIDO$("PEDDETA", REPED&)
           '
           IKITEM% = IKITEM% + 1
           IK$ = TRIM$(Str$(IKITEM%))
           CAVERI$ = CAVER$ + "-" + IK$
           Call REPLA(YY$, CAVERI$, 113, 10)
           Call GRAREG("PEDDETA", YY$, REPED&)
         Next UL&
       End If
     End If
   Next U&
   Call CIERRARCH("PEDDETA")
   '
   FIN& = ULTREG&("PEDDETP")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("PEDDETP", U&)
     NROPE& = CVS(Left$(XX$, 4))
     REPEDA& = CVS(Mid$(XX$, 109, 4))
     '
     If REPEDA& > 0 Then
       YY$ = REGLEIDO$("PEDDETA", REPEDA&)
       If CVS(Left$(YY$, 4)) = NROPE& Then
         CAVERI$ = Mid$(YY$, 113, 10)
         Call REPLA(XX$, CAVERI$, 113, 10)
         Call GRAREG("PEDDETP", XX$, U&)
       End If
     End If
   Next U&
   Call CIERRARCH("PEDDETP")
   '
   FIN& = ULTREG&("DESPAPED")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("DESPAPED", U&)
     NROPE& = CVS(Left$(XX$, 4))
     REPEDA& = CVS(Mid$(XX$, 109, 4))
     '
     If REPEDA& > 0 Then
       YY$ = REGLEIDO$("PEDDETA", REPEDA&)
       If CVS(Left$(YY$, 4)) = NROPE& Then
         CAVERI$ = Mid$(YY$, 113, 10)
         Call REPLA(XX$, CAVERI$, 113, 10)
         Call GRAREG("DESPAPED", XX$, U&)
       End If
     End If
   Next U&
   Call CIERRARCH("DESPAPED")
   '
   Call GRACONTROL("PEDCLIEN", "MARPREPA", "SI")
   Screen.MousePointer = 1
   Call COMUNI("Re-Organización Completada")
   '
End Sub
