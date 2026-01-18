VERSION 5.00
Begin VB.Form FORPED04 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pedidos de Clientes - Formato Lineal"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11910
   Icon            =   "FORPED04.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   11910
   Begin VB.PictureBox Picture6 
      Height          =   8300
      Left            =   11025
      ScaleHeight     =   8235
      ScaleWidth      =   1215
      TabIndex        =   124
      Top             =   0
      Visible         =   0   'False
      Width           =   1275
      Begin VB.CommandButton Command23 
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
         Height          =   720
         Left            =   120
         Picture         =   "FORPED04.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   125
         ToolTipText     =   "Aprueba, Graba e Imprime Pedido"
         Top             =   7080
         Width           =   650
      End
      Begin VB.Image Image7 
         Height          =   8415
         Left            =   0
         Picture         =   "FORPED04.frx":074C
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   1095
      End
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Por Medidas"
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
      Left            =   9155
      TabIndex        =   123
      Top             =   525
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3375
      Left            =   0
      TabIndex        =   31
      Top             =   4650
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
         Height          =   1440
         Left            =   5670
         TabIndex        =   47
         Top             =   105
         Width           =   5055
         Begin VB.Frame Frame8 
            BackColor       =   &H00C0C0C0&
            Caption         =   "I.V.A."
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
            Left            =   105
            TabIndex        =   106
            Top             =   2205
            Width           =   1485
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               BackColor       =   &H0000C000&
               BackStyle       =   0  'Transparent
               Caption         =   "IVA 2"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   420
               TabIndex        =   112
               Top             =   620
               Width           =   1005
            End
            Begin VB.Label Label19 
               BackColor       =   &H0000C000&
               BackStyle       =   0  'Transparent
               Caption         =   "T-2"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   60
               TabIndex        =   111
               Top             =   620
               Width           =   405
            End
            Begin VB.Label Label15 
               BackColor       =   &H0000C000&
               BackStyle       =   0  'Transparent
               Caption         =   "T-0"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   60
               TabIndex        =   110
               Top             =   190
               Width           =   400
            End
            Begin VB.Label Label16 
               BackColor       =   &H0000C000&
               BackStyle       =   0  'Transparent
               Caption         =   "T-1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   60
               TabIndex        =   109
               Top             =   405
               Width           =   400
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               BackColor       =   &H0000C000&
               BackStyle       =   0  'Transparent
               Caption         =   "IVA 1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   420
               TabIndex        =   108
               Top             =   405
               Width           =   1005
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               BackColor       =   &H0000C000&
               BackStyle       =   0  'Transparent
               Caption         =   "IVA 0"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   420
               TabIndex        =   107
               Top             =   190
               Width           =   1005
            End
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
            Index           =   2
            Left            =   3405
            TabIndex        =   89
            TabStop         =   0   'False
            Text            =   "    *******"
            Top             =   970
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
            TabIndex        =   83
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
            TabIndex        =   82
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
            TabIndex        =   81
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
            Height          =   270
            Left            =   840
            MaxLength       =   14
            TabIndex        =   73
            Top             =   970
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
            TabIndex        =   71
            TabStop         =   0   'False
            Text            =   "    *******"
            Top             =   640
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
            TabIndex        =   70
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
            TabIndex        =   50
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
            TabIndex        =   49
            TabStop         =   0   'False
            Text            =   " "
            Top             =   640
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
            Height          =   270
            Left            =   1995
            MaxLength       =   14
            TabIndex        =   22
            Top             =   640
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
            TabIndex        =   48
            TabStop         =   0   'False
            Text            =   " "
            Top             =   970
            Width           =   1500
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   210
            Picture         =   "FORPED04.frx":2AE4
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
            TabIndex        =   86
            Top             =   1785
            Width           =   1905
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Percep.I.Brutos:"
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
            TabIndex        =   85
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
            TabIndex        =   84
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
            TabIndex        =   76
            Top             =   420
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
               Weight          =   900
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00005500&
            Height          =   330
            Left            =   525
            TabIndex        =   75
            ToolTipText     =   "Doble-Click para 'Dólares'"
            Top             =   570
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
            TabIndex        =   74
            Top             =   1050
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
            TabIndex        =   53
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
            TabIndex        =   52
            Top             =   1100
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
            TabIndex        =   51
            Top             =   735
            Width           =   225
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   1590
         Left            =   3480
         TabIndex        =   77
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
            TabIndex        =   80
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
            TabIndex        =   79
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
            TabIndex        =   78
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
         TabIndex        =   39
         Top             =   105
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
            TabIndex        =   69
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
            TabIndex        =   42
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
            TabIndex        =   41
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
            TabIndex        =   40
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
            TabIndex        =   88
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
            TabIndex        =   87
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
            TabIndex        =   68
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
            TabIndex        =   46
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
            TabIndex        =   45
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
            TabIndex        =   44
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
            TabIndex        =   43
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
         Height          =   1600
         Left            =   5670
         TabIndex        =   36
         Top             =   1615
         Width           =   5055
         Begin VB.CommandButton Command24 
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
            Left            =   1800
            TabIndex        =   119
            Top             =   1200
            Width           =   195
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
            Left            =   2040
            MaxLength       =   48
            TabIndex        =   118
            Top             =   1200
            Width           =   2865
         End
         Begin VB.TextBox Text19 
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
            Left            =   1155
            TabIndex        =   117
            Top             =   1200
            Width           =   645
         End
         Begin VB.CommandButton Command25 
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
            Left            =   1785
            TabIndex        =   114
            Top             =   870
            Width           =   195
         End
         Begin VB.TextBox Text20 
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
            Left            =   1155
            TabIndex        =   113
            Top             =   870
            Width           =   645
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
            Top             =   555
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
            Top             =   240
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
            TabIndex        =   59
            Top             =   555
            Visible         =   0   'False
            Width           =   3750
         End
         Begin VB.Label Label24 
            Height          =   225
            Left            =   105
            TabIndex        =   121
            Top             =   840
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label Label23 
            Height          =   225
            Left            =   105
            TabIndex        =   120
            Top             =   525
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Sucursal:"
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
            Index           =   13
            Left            =   -840
            TabIndex        =   116
            Top             =   975
            Width           =   1905
         End
         Begin VB.Label Label27 
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
            Height          =   300
            Left            =   2040
            TabIndex        =   115
            Top             =   870
            Width           =   2865
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
            TabIndex        =   38
            Top             =   1305
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
            TabIndex        =   37
            Top             =   360
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
      TabIndex        =   25
      Top             =   0
      Width           =   12000
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
         Height          =   360
         Left            =   1080
         TabIndex        =   122
         Top             =   850
         Width           =   180
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
         Picture         =   "FORPED04.frx":2DEE
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "Búsqueda por CUIT, Domicilio, Nombre de Fantasía, etc"
         Top             =   890
         Width           =   540
      End
      Begin VB.PictureBox MARSELPRESU 
         BackColor       =   &H00FFFF00&
         Height          =   2940
         Left            =   210
         ScaleHeight     =   2880
         ScaleWidth      =   10455
         TabIndex        =   60
         Top             =   1700
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
            Height          =   2490
            Left            =   -30
            Style           =   1  'Checkbox
            TabIndex        =   66
            Top             =   420
            Width           =   9780
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00FFFF00&
            Height          =   470
            Left            =   -30
            ScaleHeight     =   405
            ScaleWidth      =   9720
            TabIndex        =   64
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
               TabIndex        =   65
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
            Picture         =   "FORPED04.frx":4154
            Style           =   1  'Graphical
            TabIndex        =   63
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
            Picture         =   "FORPED04.frx":445E
            Style           =   1  'Graphical
            TabIndex        =   62
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
            Picture         =   "FORPED04.frx":4768
            Style           =   1  'Graphical
            TabIndex        =   61
            ToolTipText     =   "Imprime Ficha de Anotaciones"
            Top             =   1890
            Width           =   570
         End
         Begin VB.Line Line5 
            X1              =   9750
            X2              =   9750
            Y1              =   0
            Y2              =   2940
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
         Height          =   2940
         Left            =   210
         ScaleHeight     =   2880
         ScaleWidth      =   10455
         TabIndex        =   54
         Top             =   1700
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
            Picture         =   "FORPED04.frx":4DD2
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Imprime Ficha de Anotaciones"
            Top             =   2205
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
            Picture         =   "FORPED04.frx":543C
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
            Picture         =   "FORPED04.frx":5746
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
            Height          =   2540
            Left            =   -30
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   57
            Top             =   370
            Width           =   9780
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FFFFC0&
            Height          =   435
            Left            =   -30
            ScaleHeight     =   375
            ScaleWidth      =   9720
            TabIndex        =   55
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
               TabIndex        =   56
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
         Height          =   2535
         Left            =   210
         TabIndex        =   58
         Top             =   2100
         Width           =   10515
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   0
         Top             =   0
      End
      Begin VB.CommandButton Command10 
         Height          =   260
         Left            =   420
         TabIndex        =   35
         Top             =   6855
         Width           =   2825
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   210
         ScaleHeight     =   375
         ScaleWidth      =   10455
         TabIndex        =   33
         Top             =   1700
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
            TabIndex        =   34
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
         TabIndex        =   10
         Text            =   " "
         Top             =   930
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
         Top             =   120
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
         Height          =   330
         Left            =   5775
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   " "
         Top             =   670
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
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   " "
         Top             =   105
         Width           =   4065
      End
      Begin VB.TextBox Text1 
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
         Left            =   210
         MaxLength       =   5
         TabIndex        =   0
         Text            =   " "
         Top             =   890
         Width           =   855
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00400040&
         Height          =   8100
         Left            =   11025
         ScaleHeight     =   8040
         ScaleWidth      =   945
         TabIndex        =   90
         Top             =   -20
         Width           =   1000
         Begin VB.PictureBox MARCOBARRA 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   105
            ScaleHeight     =   75
            ScaleWidth      =   585
            TabIndex        =   102
            Top             =   6300
            Width           =   650
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   0
               TabIndex        =   103
               Top             =   0
               Width           =   12000
            End
         End
         Begin VB.CommandButton Command14 
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
            Height          =   640
            Left            =   105
            Picture         =   "FORPED04.frx":5A50
            Style           =   1  'Graphical
            TabIndex        =   101
            ToolTipText     =   "Aprueba, Graba e Imprime Pedido"
            Top             =   7200
            Width           =   650
         End
         Begin VB.CommandButton Command3 
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
            Height          =   590
            Left            =   105
            Picture         =   "FORPED04.frx":5D5A
            Style           =   1  'Graphical
            TabIndex        =   99
            ToolTipText     =   "Ayuda FLEXOFT en Línea"
            Top             =   685
            Width           =   650
         End
         Begin VB.CommandButton Command13 
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
            Height          =   580
            Left            =   105
            Picture         =   "FORPED04.frx":6064
            Style           =   1  'Graphical
            TabIndex        =   98
            ToolTipText     =   "Opciones Varias - Modificaciones y Consultas"
            Top             =   3960
            Width           =   650
         End
         Begin VB.CommandButton Command11 
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
            Height          =   580
            Left            =   105
            Picture         =   "FORPED04.frx":61AE
            Style           =   1  'Graphical
            TabIndex        =   97
            ToolTipText     =   "Acceso a Archivos Maestros"
            Top             =   1420
            Width           =   650
         End
         Begin VB.CommandButton Command8 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   590
            Left            =   105
            Picture         =   "FORPED04.frx":64B8
            Style           =   1  'Graphical
            TabIndex        =   96
            ToolTipText     =   "Consulta Rápida de Pedidos Pendientes"
            Top             =   2000
            Width           =   650
         End
         Begin VB.CommandButton Command7 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   105
            Picture         =   "FORPED04.frx":67C2
            Style           =   1  'Graphical
            TabIndex        =   95
            ToolTipText     =   "Acceso a Condiciones Comerciales"
            Top             =   5380
            Width           =   650
         End
         Begin VB.CommandButton Command5 
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
            Height          =   580
            Left            =   105
            Picture         =   "FORPED04.frx":6ACC
            Style           =   1  'Graphical
            TabIndex        =   94
            ToolTipText     =   "Configuración de Funcionamiento de Pedidos"
            Top             =   4800
            Width           =   650
         End
         Begin VB.CommandButton Command4 
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
            Height          =   580
            Left            =   105
            Picture         =   "FORPED04.frx":6F0E
            Style           =   1  'Graphical
            TabIndex        =   93
            ToolTipText     =   "Resumen de Cuenta Corriente"
            Top             =   2790
            Width           =   650
         End
         Begin VB.CommandButton Command2 
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
            Height          =   580
            Left            =   105
            Picture         =   "FORPED04.frx":7218
            Style           =   1  'Graphical
            TabIndex        =   92
            ToolTipText     =   "Cierra y Abandona la Aplicación"
            Top             =   105
            Width           =   650
         End
         Begin VB.CommandButton Command22 
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
            Height          =   580
            Left            =   315
            Picture         =   "FORPED04.frx":7362
            Style           =   1  'Graphical
            TabIndex        =   91
            ToolTipText     =   "Anotador de Pedidos"
            Top             =   3380
            Visible         =   0   'False
            Width           =   650
         End
         Begin VB.CommandButton Command12 
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
            Height          =   580
            Left            =   105
            Picture         =   "FORPED04.frx":766C
            Style           =   1  'Graphical
            TabIndex        =   100
            ToolTipText     =   "Consultas y Re-Impresión"
            Top             =   3380
            Width           =   650
         End
         Begin VB.Image Image3 
            Height          =   390
            Left            =   -310
            Picture         =   "FORPED04.frx":7976
            Top             =   6615
            Width           =   1515
         End
      End
      Begin VB.Image Image5 
         BorderStyle     =   1  'Fixed Single
         Height          =   420
         Left            =   8190
         Picture         =   "FORPED04.frx":9898
         Stretch         =   -1  'True
         ToolTipText     =   "Doble Click Para Importar desde B.Datos MDB"
         Top             =   220
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Image Image6 
         BorderStyle     =   1  'Fixed Single
         Height          =   420
         Left            =   8190
         Picture         =   "FORPED04.frx":99E2
         Stretch         =   -1  'True
         ToolTipText     =   "Doble Click Para Importar desde Planilla Excel"
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Venta"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   5500
         TabIndex        =   105
         Top             =   20
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label Label26 
         BackColor       =   &H00EEEEEE&
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
         Height          =   300
         Left            =   210
         TabIndex        =   104
         Top             =   1280
         Width           =   1035
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   7665
         Picture         =   "FORPED04.frx":AA2C
         ToolTipText     =   "Doble Click para Acceso a Pedidos Abiertos"
         Top             =   210
         Width           =   480
      End
      Begin VB.Line Line6 
         X1              =   0
         X2              =   11865
         Y1              =   0
         Y2              =   0
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
         TabIndex        =   67
         Top             =   0
         Width           =   1380
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00C0C000&
         FillColor       =   &H0080FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   420
         Shape           =   4  'Rounded Rectangle
         Top             =   315
         Width           =   540
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H0000C000&
         FillColor       =   &H00FFFF80&
         FillStyle       =   0  'Solid
         Height          =   390
         Left            =   315
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   750
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000C0C0&
         FillColor       =   &H00C0E0FF&
         FillStyle       =   0  'Solid
         Height          =   540
         Left            =   210
         Shape           =   4  'Rounded Rectangle
         Top             =   165
         Width           =   960
      End
      Begin VB.Line Line7 
         X1              =   210
         X2              =   10710
         Y1              =   4640
         Y2              =   4640
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
         TabIndex        =   32
         Top             =   1025
         Width           =   1695
      End
      Begin VB.Line Line8 
         X1              =   10725
         X2              =   10725
         Y1              =   1700
         Y2              =   4640
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
         TabIndex        =   30
         Top             =   975
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
         Top             =   420
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   50
         TabIndex        =   26
         Top             =   680
         Width           =   960
      End
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   105
      Picture         =   "FORPED04.frx":AE6E
      ToolTipText     =   "Recepción Pedidos Electrónicos"
      Top             =   8715
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1050
   End
   Begin VB.Menu Archi 
      Caption         =   "Archivos"
      Begin VB.Menu nueped 
         Caption         =   "Nuevo Pedido"
      End
      Begin VB.Menu MoPeCur 
         Caption         =   "Modificar Pedido en Curso"
      End
      Begin VB.Menu php0 
         Caption         =   "-"
      End
      Begin VB.Menu LiPrVe 
         Caption         =   "Listas de Precios de Venta"
      End
      Begin VB.Menu php0a 
         Caption         =   "-"
      End
      Begin VB.Menu maeart 
         Caption         =   "Consulta Rapida de Pedidos Pendientes"
      End
      Begin VB.Menu maecli 
         Caption         =   "Maestro de Clientes"
      End
      Begin VB.Menu ss1 
         Caption         =   "-"
      End
      Begin VB.Menu grabarpedido 
         Caption         =   "Grabar Pedido"
      End
      Begin VB.Menu php1 
         Caption         =   "-"
      End
      Begin VB.Menu sali 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu Edic 
      Caption         =   "Edición"
      Begin VB.Menu MoPeCu 
         Caption         =   "Modificación de Pedido en Curso"
      End
      Begin VB.Menu AnuPed 
         Caption         =   "Anulación de Pedidos"
      End
      Begin VB.Menu RevAnu 
         Caption         =   "Revertir Anulación de Pedidos"
      End
      Begin VB.Menu php1a 
         Caption         =   "-"
      End
      Begin VB.Menu aprobarpedidos1 
         Caption         =   "Aprobar Pedidos"
         Begin VB.Menu verpendientesaprobacion 
            Caption         =   "Ver Pedidos Pendientes de Aprobación"
         End
         Begin VB.Menu aprobarpedidos 
            Caption         =   "Aprobación de Pedidos"
         End
      End
      Begin VB.Menu ss2 
         Caption         =   "-"
      End
      Begin VB.Menu AnuItem 
         Caption         =   "Anulación por Items de Pedidos"
      End
      Begin VB.Menu ss3 
         Caption         =   "-"
      End
      Begin VB.Menu pedidoxcopia 
         Caption         =   "Pedidos por Copia"
      End
   End
   Begin VB.Menu ConLis 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu PePeCli 
         Caption         =   "Pedidos Pendientes por Cliente"
         Begin VB.Menu PeClPan 
            Caption         =   "Por Pantalla"
         End
         Begin VB.Menu PeCliEx 
            Caption         =   "Imprimir/Exportar"
         End
      End
      Begin VB.Menu PePeCo 
         Caption         =   "Pedidos Pendientes por Código de Producto"
         Begin VB.Menu PeCoPa 
            Caption         =   "Por Pantalla"
         End
         Begin VB.Menu PeCoEx 
            Caption         =   "Imprimir/Exportar"
         End
      End
      Begin VB.Menu php1b 
         Caption         =   "-"
      End
      Begin VB.Menu EncaPe 
         Caption         =   "Encabezados de Pedidos"
      End
      Begin VB.Menu DeNumPe 
         Caption         =   "Detalle por Número de Pedido"
      End
      Begin VB.Menu php2 
         Caption         =   "-"
      End
      Begin VB.Menu ConRimp 
         Caption         =   "Consultas y Re-Impresión"
      End
      Begin VB.Menu ss4 
         Caption         =   "-"
      End
      Begin VB.Menu listadopendienteaprobar 
         Caption         =   "Listado de Pedidos Pendiente de Aprobación"
      End
      Begin VB.Menu ss6 
         Caption         =   "-"
      End
      Begin VB.Menu estaventaypedi 
         Caption         =   "Listados de Estadisticas de Ventas y Pedidos"
      End
      Begin VB.Menu histo 
         Caption         =   "Listado Histórico de Precios por Pedido"
      End
      Begin VB.Menu listcomplementarios 
         Caption         =   "Listados Complementarios"
      End
   End
   Begin VB.Menu AccDirec 
      Caption         =   "Accesos Directos"
      Begin VB.Menu LiPreVen 
         Caption         =   "Listas de Precios de Ventas"
      End
      Begin VB.Menu php2a 
         Caption         =   "-"
      End
      Begin VB.Menu MaesArt 
         Caption         =   "Maestro de Articulos"
      End
      Begin VB.Menu MaesCli 
         Caption         =   "Maestro de Clientes"
      End
      Begin VB.Menu php3 
         Caption         =   "-"
      End
      Begin VB.Menu RCtaCte 
         Caption         =   "Resumen de Cuenta Corriente"
      End
      Begin VB.Menu php4 
         Caption         =   "-"
      End
      Begin VB.Menu anotadorpedidos 
         Caption         =   "Anotador de Pedidos"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu otr 
         Caption         =   "Otros Accesos Directos"
      End
   End
   Begin VB.Menu Configura 
      Caption         =   "Configuración"
      Begin VB.Menu ConGral 
         Caption         =   "Configuración General"
      End
      Begin VB.Menu php5 
         Caption         =   "-"
      End
      Begin VB.Menu ConCom 
         Caption         =   "Condiciones Comerciales"
         Begin VB.Menu CoPag 
            Caption         =   "Condiciones de Pago"
         End
      End
   End
   Begin VB.Menu utiles 
      Caption         =   "Útiles"
      Begin VB.Menu pedidoafactura 
         Caption         =   "Rastreabilidad de Pedido a Factura"
      End
      Begin VB.Menu ss7 
         Caption         =   "-"
      End
      Begin VB.Menu factuaapedido 
         Caption         =   "Rastreabilidad de Factura a Pedido"
      End
   End
End
Attribute VB_Name = "FORPED04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim ARCHIPRE$
Dim STATUPC%


Private Sub anotadorpedidos_Click()
    Call Command22_Click
End Sub

Private Sub AnuItem_Click()
Call OPPED04.Command4_Click
End Sub

Private Sub AnuPed_Click()
Call OPPED04.Command45_Click
End Sub

Private Sub aprobarpedidos_Click()
   Call APRONP04.Command14_Click
End Sub

Private Sub ayuda_Click()
Call Command3_Click
End Sub

Private Sub Command1_Click()
   Call SELECHO("DEUDOR1")
   NCLIE% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
   Text1.TEXT = TRIM$(Str$(NCLIE%))
   If NCLIE% > 0 Then
     If REGICLI&(NCLIE%) > 0 Then
       Call CARDACLI
       '\\\OT-08-0179-OD-21/04/08///
       FAVEN% = HAYFACVEN%(NCLIE%)
       If FAVEN% > 0 Then
        Text8 = "Fac.Vencidas"
        Text8.ForeColor = RGB(255, 0, 0)
       End If
       '\\\OT-08-0179-OD-FIN///
       '\\\OT-05-0399-WAR-14/06/05///
       Image6.Enabled = True 'MODULO IMPORTADOR
       ''\\\OT-05-0399-WAR-FIN///
       On Error Resume Next
       Command6.SetFocus
       On Error GoTo 0
     Else
        Image6.Enabled = False 'MODULO IMPORTADOR
     End If
   End If
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call CIERRARCH("*.*")
   ACONEC$ = "ARCHIG04"
   Call CONECRUN(ACONEC$, "Archivos Maestros")
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
    '
    Command12.Enabled = False
    Screen.MousePointer = 11
    '
    SQLX$ = "SELECT * FROM PEDENCA ORDER BY NroPed ASC"
    Set PEDENCTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    JJ& = 0
    With PEDENCTEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        Do While Not .EOF
          'ANTIGUO CODIGO ANTES DE OT-05-0682-WAR-18/10/05///
'          NPX& = !NroPed
'          REFEPED$ = !Referencia
'          NCLIE% = !NroClie
'          FEX = CVINT(!FECHEMIS)
'          ANULX$ = "": If !Status = 9 Then ANULX$ = "(Anul) "
'          YY$ = MKS$(NPX&) + ANULX$ + AJUSTI$(REFEPED$, 60)
'          JJ& = JJ& + 1
'          Call GRAREG("INDIPEDI", YY$, JJ&)
            '\\\OT-05-0682-WAR-18/10/05///
             NRODES& = !NroPed
             TTXX$ = MKS$(NRODES&)
               NOCO$ = !NroOcom
             If NOCO$ <> "" Then
               TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
             End If
               FEX = CVINT(!FECHEMIS)
             TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
               NCLIE% = !NroClie
               ANULX$ = "": If !Status = 9 Then ANULX$ = "(Anul) "
             TTXX$ = TTXX$ + TRIM$(RASOCLI$(NCLIE%)) + ANULX$ 'Space$(52)
             JJ& = JJ& + 1
             Call GRAREG("INDIPEDI", TTXX$, JJ&)
             '\\\OT-05-0682-WAR-FIN///
        .MoveNext
        Loop
      End If
    End With
    '
    Call SETULTREG("INDIPEDI", JJ&)
    Call CIERRARCH("INDIPEDI")
    Call FRESHECHO("INDIPEDI")
    Screen.MousePointer = 1
    '
3   Call SELECHO("INDIPEDI/Indice General de Pedidos (Pendientes-Cumplidos-Anulados)")
    NPX& = XVALO(VALACT1$("INDIPEDI"))
    If NPX& < 1 Then GoTo 99
    '
    TIDOCUM$ = "Pedido de Cliente"
    TIDOCUX$ = "Pedido ABIERTO"
    TIDOCUY$ = "Pedido Interno"
    NUDOCU$ = TRIM$(Str$(NPX&))
    While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
    NUDOCU$ = NUDOCU$ + " "
    Screen.MousePointer = 11
    '
4   For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Or InStr(XX$, TIDOCUX$) > 4 Or InStr(XX$, TIDOCUY$) > 4 Then
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
    OPPED04.Command46.SetFocus
    On Error GoTo 0
    OPPED04.Show 1
    GoTo 99
    '
5   Call MUESTRADOC(DOCUNU&)
    GoTo 3
    '
99  Command12.Enabled = True
    '
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   PROLLAMA$ = "CERR"
   OPPED04.Show 1
   If MONEMI% < 1 Then MONEMI% = 1
   MONEMIS = DEMONECO$(MONEMI%)
   TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
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
   If TRIM$(Text10) <> "" Then
    If CONTROL$("", "VALEXIOC") = "SI" Then ' VALIDACION DE NUMERO DE ORDEN DE COMPRA EN PEDIDO
     NCC& = XVALO(Text1)
     SQLX$ = "SELECT * FROM PEDENCA "
     SQLX$ = SQLX$ + " WHERE NROOCOM = '" & TRIM$(Text10) & "'"
     SQLX$ = SQLX$ + " AND NroClie =  " & NCC& & ""
     SQLX$ = SQLX$ + " ORDER BY NroPed DESC "
     Set PEDENCTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     With PEDENCTEMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
          On Error GoTo 0
          NPED& = XVALO(!NroPed)
          Call MENSERR(24, "Orden de Compra " & TRIM$(Text10) & " Procesada con Pedido Nro. " & NPED&)
          GoTo 99
       End If
       On Error GoTo 0
     End With
    End If
   End If
   '
   Call GRAPEDIDO(OKX%)
   If OKX% > 0 Then
     If CONTROL$("PEDIDO", "PRINOVEN") <> "NO" Then
        Call PRIPEDIDO(1)     ' Imprime Orden de Venta
     End If
     If CONTROL$("PEDIDO", "PRINODES") = "SI" Then
        Call PRIPEDIDO(2)     ' Imprime Orden de Despacho
     End If
     Call GENERODES      ' genera orden de despacho
     '
     GENORFAB% = 0
     If CONTROL$("ADMIPED", "GENORFAB") = "SI" Then GENORFAB% = 1
     If STATUPC% = (-1) Then
          Call COMUNI("Pedido Registado como\'PENDIENTE DE APROBACIÓN'")
        Else
          ' GENERA ORDEN DE FABRICACION SOLO SI PEDIDO CONFIRMADO
          If GENORFAB% = 1 Then
             Call GENORFA
          End If
          '
          If CONTROL$("ADMIPED", "REPLICAM") = "SI" Then
            NCXX% = XVALO(Text1)
            If EMAILCLI$(NCXX%) <> "" Then
              If COMALTER%("¿ Enviar por Mail Al Cliente ?\\Si, Enviar\No, Continuar") = 1 Then
                SENDBYMAIL$ = EMAILCLI$(NCXX%)
                Call PRIPEDIDO(1)     ' Imprime Orden de Venta
                SENDBYMAIL$ = ""
              End If
            End If
          End If
     End If
     Screen.MousePointer = 1
     '
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

 Sub Command19_Click()
   '
   If Left$(UCase$(Label11), 6) = "REPARA" Then GoTo 50
   '
   NPX& = 0
   Call ABREPRESVEN
   HOII% = HOY(HOS$)
   NCLIE% = XVALO(Text1)
   JJ& = 0: ATX% = 0: COEFCON = 1
   For KKI% = 1 To List2.ListCount
     If List2.Selected(KKI% - 1) = True Then
       NUPRESU& = XVALO(Mid$(List2.List(KKI% - 1), 251, 6))
       If NUPRESU& > 0 Then
         SQLX$ = "SELECT * FROM PRESUVEN "
         SQLX$ = SQLX$ + "WHERE NroClie = " & NCLIE% & " "
         SQLX$ = SQLX$ + "AND NroPres = " & NUPRESU& & " "
         SQLX$ = SQLX$ + "AND NroPedi < 1 "
         SQLX$ = SQLX$ + "ORDER BY NroPres ASC"
         '
         Set IPRESPEN = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
         With IPRESPEN
           On Error Resume Next
           .MoveFirst
           If Err < 1 Then
             Do While Not .EOF
               If !NroPedi < 1 Then
                 ZZ$ = REGBLAN$("!CARDETPE")
                   CIXI% = !CodiInt        ' CVI(Mid$(XX$, 69, 2))
                   CLRBASE$ = TRIM$(SAFETEXT(!CBase))
                   TTELA$ = TRIM$(SAFETEXT(!TTalle))
                   CLRTELA$ = TRIM$(SAFETEXT(!TColor))

                 Call REPLA(ZZ$, MKI$(CIXI%), 1, 2)
                 Call REPLA(ZZ$, CLRBASE$, 3, 12)
                 Call REPLA(ZZ$, TTELA$, 15, 12)
                 Call REPLA(ZZ$, CLRTELA$, 27, 12)
                   PRELI = !PreList        ' CVS(Mid(XX$, 71, 4))
                   DESCUTEX$ = !TDESCUEN   ' TRIM$(Mid$(XX$, 75, 14))
                   CPPP = (100 - COEFDESCU(DESCUTEX$)) / 100
                   If CPPP < 1 Then ATX% = 1
                   '
                   MONEMA% = !MonedEmi     ' Asc(Mid$(XX$, 68, 1))
                   If JJ& < 1 Then
                     MONEMI% = MONEMA%     ' Asc(Mid$(XX$, 68, 1))
                     If MONEMI% < 1 Then MONEMI% = 1
                     MONEMIS = DEMONECO$(MONEMI%)
                     TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
                   End If
                   COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
                   CAPRESU& = !Cantit      ' CVS(Mid$(XX$, 101, 4))
                   If CAPRESU& < 1 Then CAPRESU& = 1
                   '
                 Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 39, 4)
                 Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 51, 4)
                 Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 65, 4)
                 Call REPLA(ZZ$, MKS$(CAPRESU&), 47, 4)   ' CANTIDAD
                   DESCUTEX$ = !TDESCUEN   ' TRIM$(Mid$(XX$, 75, 14))
                   'Call REPLA(ZZ$, MKS$(COEFDESCU(DESCUTEX$)), 69, 4)
                 Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 73, 4)
                 Call REPLA(ZZ$, MKI$(HOII%), 63, 2)
                   NUORIT% = NuOrdIt
                 Call REPLA(ZZ$, MKI$(NUORIT%), 123, 2)
                 Call REPLA(ZZ$, MKS$(NUPRESU&), 125, 4)
                 JJ& = JJ& + 1
                 Call GRAREG("!CARDETPE", ZZ$, JJ&)
               End If
             .MoveNext
             Loop
           End If
         End With
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
       REPRESU& = XVALO(Mid$(List2.List(KKI% - 1), 121, 8))
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
           MONEMIS = DEMONECO$(MONEMI%)
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
     '
     Call ABRECLIEN
     On Error Resume Next
     With Deudor12
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
    '
    Command22.Enabled = False
    Screen.MousePointer = 11
    '
    SQLX$ = "SELECT * FROM PEDENCA ORDER BY NroPed DESC"
    Set PEDENCTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    JJ& = 0
    With PEDENCTEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        Do While Not .EOF
          NPX& = !NroPed
          NCLIE% = !NroClie
          FEX = CVINT(!FECHEMIS)
          ANULX$ = "": If !Status = 9 Then ANULX$ = "(Anul) "
          YY$ = MKS$(NPX&) + FECHATEX$(FEX) + "  " + ANULX$ + TRIM$(RASOCLI$(NCLIE%))
          JJ& = JJ& + 1
          Call GRAREG("INDIPEDI", YY$, JJ&)
        .MoveNext
        Loop
      End If
    End With
    '
    Call SETULTREG("INDIPEDI", JJ&)
    Screen.MousePointer = 1
    CONPED04.Show 1
    '
    Command22.Enabled = True
    '
End Sub



Private Sub Command23_Click()
   Call Command14_Click
End Sub

Private Sub Command24_Click()
  '\\\OT-6-422-FG-05/10/06
   Call SELECHO("TRANSPOR")
   NTRAN% = XVALO(VALACT1$("TRANSPOR"))
   Text19 = TRIM$(FORMATNUM(NTRAN%, "I4"))
  '\\\OT-6-422-FG-FIN
End Sub

Private Sub Command25_Click()
    Call SELECHO("!SUCENTRE")
    VDEVU$ = VALACT1$("!SUCENTRE")
    Text20 = VDEVU$
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call HELPONLINE("PRESUP01")
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    RECUEC04.Text1.TEXT = Text1.TEXT
    RECUEC04.Show 1
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    Call CIERRARCH("*.*")
    If (USNBR% Mod 100) <> 1 Then
      Call COMUNI("La Opción Elegida Requiere\Privilegios de Supervisor.")
      Exit Sub
    End If
    CFGPED04.Show 1
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
   'UREPRE& = ULTREG&("PRESUPU")
   Screen.MousePointer = 11
   List2.Clear
   NCLIE% = XVALO(Text1.TEXT)
   If UCase$(Left$(Label11, 6)) = "REPARA" Then GoTo 50
   '
   SQLX$ = "SELECT * FROM PRESUVEN "
   SQLX$ = SQLX$ + "WHERE NroClie = " & NCLIE% & " "
   SQLX$ = SQLX$ + "AND NroPedi < 1 "
   SQLX$ = SQLX$ + "ORDER BY NroPres ASC"
   '
   NPX& = 0
   Call ABREPRESVEN
   Set IPRESPEN = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   With IPRESPEN
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       Do While Not .EOF
         If IsNull(!MarNoVig) = True Or !MarNoVig = "" Then
           NRO& = !NroPres
             TTXX$ = "Presup.No: " + TRIM$(Str$(NRO&))
           FEX = CVINT(!FECHEMIS)
             TTXX$ = TTXX$ + " - " + FECHATEX$(FEX)
           If TRIM$(!Referen) <> "" Then
             TTXX$ = TTXX$ + " - Ref: " + TRIM$(!Referen)
           End If
          'End If
           CIXI% = !CodiInt
           '
           If NRO& = NPX& Then
             TTXX$ = TTXX$ + " - VARIOS"
             List2.RemoveItem (List2.ListCount - 1)
             GoTo 15
           End If
           '
           NPX& = NRO&
           CAPRESU& = !Cantit
           If CAPRESU& < 1 Then CAPRESU& = 1
           TTXX$ = TTXX$ + " - " + TRIM$(CODEXT$(CIXI%)) + " (" + TRIM$(Str$(CAPRESU&)) + ") " + DESCRIT0$(CIXI%)
           '
15         TTXX$ = AJUSTI$(TTXX$, 256)
           Call REPLA(TTXX$, TRIM$(Str$(NRO&)), 251, 6)
           List2.AddItem TTXX$
         End If
         '
19     .MoveNext
       Loop
     End If
   End With
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
    MODX% = ALTERNA%("Listas de Precios\Condiciones de Pago\Alícuotas de I.V.A.\Descuentos Combinados")
    If MODX% = 1 Then
         Call CONECRUN("LISPRE04", "Listas de Precios de Venta")
       ElseIf MODX% = 2 Then
         Call CARCONPAGOS
       ElseIf MODX% = 3 Then
         Call CARTASIVA
         Call COMUNI("Re-Inicie Ahora la Aplicación.")
         Call CIERRARCH("*.*")
         Call FINAL("")
       ElseIf MODX% = 4 Then
         DEFDESCO.Show 1
    End If
    Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    Call CIERRARCH("*.*")
    CONPEPEN.Show 1
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



Private Sub ConGral_Click()
Call Command5_Click
End Sub

Private Sub ConRimp_Click()
Call Command12_Click
End Sub

Private Sub CoPag_Click()
MODO% = 2
Call MODOCON(MODO%)
End Sub

Private Sub DeNumPe_Click()
Call OPPED04.Command63_Click
End Sub

Private Sub Eco21_GotFocus(Index As Integer)
   On Error Resume Next
   TEXT21(Index).SetFocus
   On Error GoTo 0
End Sub


Private Sub EncaPe_Click()
Call OPPED04.Command61_Click
End Sub
Private Sub conlis_Click()
   PROLLAMA$ = "CERR"
   If MONEMI% < 1 Then MONEMI% = 1
   MONEMIS = DEMONECO$(MONEMI%)
   TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
End Sub
Private Sub edic_Click()
   PROLLAMA$ = "CERR"
   If MONEMI% < 1 Then MONEMI% = 1
   MONEMIS = DEMONECO$(MONEMI%)
   TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
End Sub

Private Sub estaventaypedi_Click()
   Call OPPED04.Command12_Click
End Sub

Private Sub factuaapedido_Click()
     Call OPPED04.FILARTCLI
End Sub

Private Sub Form_Activate()
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture6.Visible = True
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame3.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Frame5.BackColor = &HFCD7B6
      Me.Frame7.BackColor = &HFCD7B6
      Me.Frame8.BackColor = &HFCD7B6
      Me.MARSELPRESU.BackColor = &HFCD7B6
      Me.MARCONOTA.BackColor = &HFCD7B6
      Me.Picture1.BackColor = &HFCD7B6
      Me.Picture2.BackColor = &HFCD7B6
      Me.Picture3.BackColor = &HFCD7B6
      Me.Picture4.BackColor = &HFCD7B6
   End If

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
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    '
    Frame4.Height = 1440
    OPENFORMS = DoEvents
    VERANTER$ = "PEDIDO01"

    Call TRANSLABELS(Name)
    DoEvents
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    If InStr(EPAC$, "DOSIVA") > 0 Then
       Label11.Visible = True
    End If
    '
    Command22.Visible = True
    Command22.Left = 105
    Command12.Visible = False
    '
    ORIPED$ = TRIM$(UCase$(CONTROL$("PEDIDO", "ORIDATOS")))
    If InStr(ORIPED$, "ORDENEMPAQUE.MDB") > 0 Then
       Image5.Visible = True
       Image6.Visible = False
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
    If EXISTE%("REPEDREM.EXE") > 0 Then
      Image2.Visible = True
    End If
    '
    MONEMI% = 1  ' PESOS
    MONEMIS.Caption = DEMONECO(MONEMI%)
    MONEMIS.Visible = True
    CONVERPRECIO% = 1
    TICAMBIS.Visible = True
    Label12.Visible = True
    Image1.Visible = True
    TICAMBIS = FORMATNUM$(TICAMONE(1), "F9.4")
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
    If EXISTE%(LUDAT$ + "PEDMEDID.RFS") > 0 Then
       Check1.Visible = True
       Call SETULTREG("!PEDMEDID", 0)
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Cancel = 0 Then
       Call Command2_Click
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
   Call FINAL("")
End Sub



Private Sub grabarpedido_Click()
    Call Command14_Click
End Sub

Private Sub histo_Click()
     Call OPPED04.Command10_Click
End Sub

Private Sub Image1_DblClick()
   Static CTXX%
   '
   If ULTREG&("DAPIBCLI") < 1 Then
     Call MENSERR(24, "Imposible desde esta Aplicación")
     Exit Sub
   End If
   '
   NCLIE% = XVALO(TRIM$(Text1.TEXT))
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
   Frame4.Height = 1440
End Sub

Private Sub Image2_DblClick()
   Call CIERRARCH("*.*")
   Call CONECRUN("REPEDREM", "Recepción Pedidos Remotos por Mail")
End Sub

Private Sub Image4_DblClick()
   PEDABIE04.Show 1
End Sub

Private Sub Image5_DblClick()
  '
  ORIPED$ = TRIM$(UCase$(CONTROL$("PEDIDO", "ORIDATOS")))
  If InStr(ORIPED$, "ORDENEMPAQUE.MDB") > 0 Then
     Call IMPORZET(ORIPED$)
     Exit Sub
  End If
  '
End Sub

Private Sub Image6_DblClick()
  '
  With IMPPED04
    .Label26 = Text1
    .Text2 = Text2
    .Text3 = Text3
    .Text4 = Text4
    .Text5 = Text5
    .Text12 = Text12
    .Text10 = Text10
    .Text9 = Text9
    .Label11 = XVALO(TEXT21(0)) 'LISTA DE PRECIOS
    .Show 1
  End With
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

Private Sub Label26_DblClick()
   Call SELECHO("INDICLIE")
   CPROX$ = TRIM$(VALACT1$("INDICLIE"))
   If CPROX$ <> "" Then
     NCXI% = NUMECLI%(CPROX$)
     If NCXI% > 0 Then
       If REGICLI&(NCXI%) > 0 Then
         Label26 = CPROX$
         NC% = NCXI%
         Text1.TEXT = TRIM$(Str$(NC%))
         Call CARDACLI
         On Error Resume Next
         Command6.SetFocus
         On Error GoTo 0
       End If
     End If
   End If
End Sub

Private Sub listadopendienteaprobar_Click()
    Call APRONP04.Command6_Click
End Sub

Private Sub listcomplementarios_Click()
    Call OPPED04.Command8_Click
End Sub

Private Sub pedidoafactura_Click()
     
    Call OPPED04.PEDAFACT
End Sub

Private Sub pedidoxcopia_Click()
    Call OPPED04.Command11_Click
End Sub

Private Sub Text8_DblClick()

    HOII% = HOY(HOS$)
    DIAHOY = Int(CDbl(CVDAT(HOII%)))
    If Text8 <> "Fac.Vencidas" Then GoTo 99
    NC% = XVALO(Text1)
    If NC% > 0 Then
      Call APERBASDAT("CABAN")
      SQLX$ = "SELECT * FROM SADEUPEN "
      SQLX$ = SQLX$ + " Where NuClien = " & NC%
      SQLX$ = SQLX$ + " AND SaldDebe > SaldAcre "
      SQLX$ = SQLX$ + " AND INT(CDBL(FEVENCIM)) < " & DIAHOY & " "
      JJ& = 0
      Set DEUDATMP = CueCorri.OpenRecordset(SQLX$)
      With DEUDATMP
        On Error Resume Next
        .MoveFirst
         If Err > 0 Then
           Call COMUNI("No Registra Facturas Vencidas ")
           GoTo 99
         Else
           On Error GoTo 0
           Do While Not .EOF
              SALDEB# = XVALO(!SaldDebe)
              SALACR# = XVALO(!SaldAcre)
              DEUDA$ = FORMATNUM$(SALDEB# - SALACR#, "F12.2")
              IMPORIG# = XVALO(!IBruComp)
              FEVE = CVINT(!FEVENCIM)
              NUMFACT$ = SAFETEXT(!CODICOM_LAR)
              FECHOR = CVINT(!FECHEMISI)
              JJ& = JJ& + 1
              Z$ = Space$(128)
              Call REPLA(Z$, NUMFACT$, 1, 18)
              Call REPLA(Z$, FECHATEX$(FECHOR), 19, 8)
              Call REPLA(Z$, FORMATNUM(IMPORIG#, "F11.2"), 27, 11)
              Call REPLA(Z$, FECHATEX$(FEVE), 40, 5)
              Call REPLA(Z$, DEUDA$, 45, 12)
              Call GRAREG("!SELDOCAP", Z$, JJ&)
              '
              .MoveNext
           Loop
         End If
         On Error GoTo 0
      End With
      Call SETULTREG("!SELDOCAP", JJ&)
      Call FRESHECHO("!SELDOCAP")
      SPDX$ = REGSEL$("!SELDOCAP")
    End If
99
    
End Sub




Private Sub Label8_Click()
   On Error Resume Next
   Text13.SetFocus
   On Error GoTo 0
End Sub

Private Sub LENTREGA_GotFocus(Index As Integer)
   '\\\OT-6-422-FG-05/10/06
   If Index = 2 Then
     If XVALO(Text19) > 0 Then
       Call MENSERR(24, "Campo no Editable")
       On Error Resume Next
       Text19.SetFocus
       On Error GoTo 0
       Exit Sub ' LE AGREGUE ESTA LINEA PARA QUE NO SE EJECUTE EL RESTO DEL CODIGO
     End If
   End If
   '\\\OT-6-422-FG-FIN
   '
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

Private Sub LiPreVen_Click()
    Call CIERRARCH("*.*")
    Call CONECRUN("LISPRE04", "Listas de Precios de Venta")
End Sub

Private Sub LiPrVe_Click()
    Call CIERRARCH("*.*")
    Call CONECRUN("LISPRE04", "Listas de Precios de Venta")
End Sub

Private Sub LiPrVta_Click()
   MODO% = 1
   Call MODOCON(MODO%)
End Sub

Private Sub List1_DblClick()
   Call CARDETPEDI(2)
End Sub



Private Sub maeart_Click()
'\\\OT-5-869-28/12/05///
   Call Command8_Click
'\\\OT-5-869-FIN///
End Sub

Private Sub maecli_Click()
'\\\OT-5-869-28/12/05///
Call Command11_Click
'\\\OT-5-869-FIN///
End Sub

Private Sub MaesArt_Click()
'\\\OT-5-869-28/12/05///
  Call Command8_Click
'\\\OT-5-869-FIN///
End Sub

Private Sub MaesCli_Click()
'\\\OT-5-869-28/12/05///
Call Command11_Click
'\\\OT-5-869-///
End Sub

Private Sub MONEMIS_Change()
   MONEMIS.ForeColor = RGB(0, 0, 255)
   If MONEMI% > 1 Then
      MONEMIS.ForeColor = RGB(255, 0, 0)
   End If
   MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
Exit Sub
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
'   If TRIM$(UCase$(MONEMIS)) = "PESOS" Then
'       MONEMIS = "Dólares"
'     Else
'       MONEMIS = "Pesos"
'   End If
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
           DMONE$(U&) = DMX$
           TCMONE(U) = VHX
         Next U&
     End If
   End If
   '
   MONEMIA% = MONEMI%
15 MONEMI% = MONEMI% + 1
   If MONEMI% > CAMON% Then MONEMI% = 1
   If DMONE$(MONEMI%) = "" Then GoTo 15
   '
   MONEMIS = DMONE$(MONEMI%)
   MONEMIS.ForeColor = RGB(0, 0, 255)
   If MONEMI% > 1 Then
      MONEMIS.ForeColor = RGB(255, 0, 0)
   End If
   MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
   '
   COEFCON = TICAMONE(MONEMIA%) / TICAMONE(MONEMI%)
   TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
   '
   If MONEMI% > 1 Then
     TICAMBIS.Visible = True
     TICAMBIS.Enabled = False
   End If
   '
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

Private Sub MoPeCu_Click()
Call OPPED04.Command46_Click
End Sub

Private Sub MoPeCur_Click()
'\\\OT-5-869-28/12/05///
Call MoPeCu_Click
'\\\OT-5-869-FIN///
End Sub

Private Sub nueped_Click()
Call Command1_Click
End Sub

Private Sub otr_Click()
Call MENSERR(24, "No Disponible por Ahora")
End Sub

Private Sub PeCliEx_Click()
OPPED04.Option2 = True
Call OPPED04.Command62_Click
End Sub

Private Sub PeClPan_Click()
OPPED04.Option1 = True
Call OPPED04.Command62_Click
End Sub

Private Sub PeCoEx_Click()
OPPED04.Option2 = True
Call OPPED04.Command64_Click
End Sub

Private Sub PeCoPa_Click()
OPPED04.Option1 = True
Call OPPED04.Command64_Click
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

Private Sub RCtaCte_Click()
Call Command4_Click
End Sub

Private Sub RevAnu_Click()
Call OPPED04.Command13_Click
End Sub

Private Sub sali_Click()
Call Command2_Click
End Sub

Private Sub Text1_Change()
   Call BLANFORMU
   Text19 = Str$(NUTRANCLI%(XVALO(Text1)))
End Sub

Private Sub Text1_DblClick()
  Call SELECHO("DEUDOR1")
  NCLIE% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
  Text1.TEXT = TRIM$(Str$(NCLIE%))
  If NCLIE% > 0 Then
    If REGICLI&(NCLIE%) > 0 Then
      Call CARDACLI
      '\\\OT-08-0179-OD-21/04/08///
      FAVEN% = HAYFACVEN%(NCLIE%)
      If FAVEN% > 0 Then
        Text8 = "Fac.Vencidas"
        Text8.ForeColor = RGB(255, 0, 0)
      End If
      '\\\OT-08-0179-OD-FIN///
      '\\\OT-05-0399-WAR-14/06/05///
      Image6.Enabled = True 'MODULO IMPORTADOR
      '\\\OT-05-0399-WAR-FIN///
      On Error Resume Next
      Command6.SetFocus
      On Error GoTo 0
    Else
      Image6.Enabled = False 'MODULO IMPORTADOR
    End If
  End If
End Sub

Private Sub Text1_GotFocus()
   If MARCONOTA.Visible = True Then
      On Error Resume Next
        Command6.SetFocus
      On Error GoTo 0
      Exit Sub
    ElseIf MARSELPRESU.Visible = True Then
      On Error Resume Next
        Command19.SetFocus
      On Error GoTo 0
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
      If XVALO(Text1.TEXT) > 0 Then
        If XVALO(Text1.TEXT) <= 32767 Then
          NC% = XVALO(Text1)
          If REGICLI&(NC%) > 0 Then
            If MARCONOTA.Visible = False Then
              If MARSELPRESU.Visible = False Then
                Call CARDACLI
                Image6.Enabled = True 'MODULO IMPORTADOR
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
   For kk% = 1 To Len(Text11.TEXT)
     CCX% = Asc(Mid$(Text11.TEXT, kk%, 1))
     If CCX% >= 48 Then
       If CCX% <= 57 Then
         NUMEPRE$ = NUMEPRE$ + Chr$(CCX%)
       End If
     End If
   Next kk%
   '
   NPP& = NUPEDIDO&
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

Private Sub Text19_Change()
   '\\\OT-07-0040-RG-02/02/07///
   Label23.Caption = ""
   Label24.Caption = ""
   LENTREGA(2) = ""
   TRANACT% = 0
   '\\\OT-07-0040-RG-FIN///
   '\\\OT-6-422-FG-05/10/06
   NTRA = XVALO(Text19)
   If NTRA > 0 Then
      If NTRA < 32767 Then
         NTRAI% = NTRA
         LENTREGA(2) = ECOARCH$("TRANSPOR", MKI$(NTRAI%))
         If LENTREGA(2) <> "" Then
            TRANACT% = NTRAI%
         End If
         '
         '\\\OT-07-0040-RG-02/02/07///
         Label23.Caption = DOMITRANS$(NTRAI%)
         Label24.Caption = LOCATRANS$(NTRAI%)
         '\\\OT-07-0040-RG-FIN///
         '
      End If
   End If
  '\\\OT-6-422-FG-FIN
End Sub

Private Sub Text19_DblClick()
   '
   Call Command24_Click
   '
End Sub

Private Sub Text2_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text20_Change()
   'AGREGADO MH04/10
  NSUC& = XVALO(Text20)
  Label27 = ""
  If EXISTE%(LUDAT$ + "SUCENTRE.RFS") > 0 Then
    LENTREGA(0) = ""
    LENTREGA(1) = ""
    LENTREGA(2) = ""
    Text19 = ""
      If NSUC& > 0 Then
        If NSUC& <= ULTREG&("!SUCENTRE") Then
          XX$ = REGLEIDO$("!SUCENTRE", NSUC&)
          Label27 = TRIM$(Mid$(XX$, 3, 62))
          LENTREGA(0) = TRIM$(Mid$(XX$, 65, 48))
          LENTREGA(1) = TRIM$(Mid$(XX$, 113, 48))
          NUSSUC% = NSUC&
          Text19 = TRIM$(NUTRANCLI%(XVALO(Text1), NUSSUC%))
        End If
      End If
  End If
End Sub

Private Sub Text20_DblClick()
    '
    Call Command25_Click
    '
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
   If XVALO(Text44) > 99.9 Then Text42 = "99.9"
   If TRIM$(Text44) <> "" Then Text44 = Format$(XVALO(Text44), "#0.0")
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
        If XVALO(TEXT21(Index).TEXT) < 1 Then
            Eco21(Index).TEXT = "CONCERTADOS FIJOS"
            ARCHIPRE$ = "MASTER"
          Else
            Eco21(Index).TEXT = ECOARCH$("LISTAS", MKI$(XVALO(TEXT21(Index).TEXT)))
            ARCHIPRE$ = TRIM$(TEXT21(0).TEXT)
            While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
            ARCHIPRE$ = "LIPRE" + ARCHIPRE$
        End If
        If MONELISTA%(XVALO(TEXT21(Index))) = 2 Then
             MONEMI% = 2
             TICAMBIS = FORMATNUM$(TICAMONE(2), "F9.4")
             MONEMIS = DEMONECO$(2)
           Else
             MONEMI% = 1
             TICAMBIS = FORMATNUM$(TICAMONE(1), "F9.4")
             MONEMIS = DEMONECO$(1)
        End If
      ElseIf Index = 1 Then
        Eco21(Index).TEXT = ECOARCH$("CONPAG", MKI$(XVALO(TEXT21(Index).TEXT)))
      ElseIf Index = 2 Then
        Eco21(Index).TEXT = ECOARCH$("VENDEDOR", Chr$(XVALO(TEXT21(Index).TEXT)))
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
   If XVALO(Text1.TEXT) > 0 Then
     If XVALO(Text1.TEXT) <= 32767 Then
       NCLIE% = XVALO(Text1.TEXT)
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
   Text19 = ""
   Text20 = ""
   Text8.ForeColor = &H808000
   List1.Clear
   Call SETULTREG("!CARDETPE", 0)
   If Check1.Visible = True Then   ' PARA QUE NO SALTE EN INSTALACIONES SIN PEDMEDID
      Call SETULTREG("!PEDMEDID", 0)
   End If
   'AGREGADOMH04/10
   Call SETULTREG("!SUCENTRE", 0)
   Call FRESHECHO("!SUCENTRE")
   '
   TEXT21(0).TEXT = "": Eco21(0) = ""
   TEXT21(1).TEXT = "": Eco21(1) = ""
   TEXT21(2).TEXT = "": Eco21(2) = ""
   Text44 = ""
   '
   PORDESCU.TEXT = ""
   For kk% = 0 To 3
      Total(kk%).TEXT = ""
      TOTIVA(kk%).TEXT = ""
      LENTREGA(kk%).TEXT = ""
   Next kk%
   Label13 = ""
   Label14 = ""
   Label20 = ""
   '
   MARCONOTA.Visible = False
   HOII% = HOY(HOS$)
   Text9.TEXT = HOS$
   Text11.TEXT = TRIM$(Str$(NUPEDIDO&))
   '\\\OT-05-0399-WAR-14/06/05///
   Image6.Enabled = False 'MODULO IMPORTADOR
   '\\\OT-05-0399-WAR-FIN///
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
   NC% = XVALO(Text1.TEXT)
   If NC% > 0 Then
     '
     If LIMICRE(NC%) < (-0.5) Then
        Call MENSERR(24, "Cuenta Suspendida.")
        Text1 = ""
        Call BLANFORMU
        GoTo 99
     End If
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
     If RASOCLI$(NC%) <> "" Then
       Text2.TEXT = RASOCLI$(NC%)
       Text3.TEXT = DOMICLI$(NC%)
       Text4.TEXT = CPOSCLI$(NC%)
       Text5.TEXT = LOCACLI$(NC%)
       '
       '27/03/07 (OT-07-0133)
       LMCR# = LIMICRE(NC%)  'PONE EN 'LIMITE DE CREDITO' LA LEYENDA 'SUSPENDIDO'
        If LMCR# < 0 Then
           Text7 = "Suspendido"
        Else
           Text7.TEXT = CSTRING$(MKS$(LIMICRE(NC%)), 3, 13, 2, 2)
        End If
       '27/03/07 (FIN)
       '
       Text12.TEXT = TRIM$(POSIVAC$(NC%)) + " - CUIT " + CUITCLI$(NC%)
       CREDISPO# = LIMICRE(NC%) - SALDCLI#(NC%) - CHEPEACRE#(NC%) - VALOPEDPEN1#(NC%)
       Text8.ForeColor = RGB(0, 0, 0)
       If CREDISPO# < 0 Then
         Text8.ForeColor = RGB(255, 0, 0)
         Call COMUNI("Atención : Cuenta sin Crédito Disponible ")
       End If
       Text8.TEXT = CSTRING$(MKD$(Abs(CREDISPO#)), 3, 13, 2, 2)
       '
       TEXT21(0).TEXT = Str$(LIPRCLI%(NC%))
       TEXT21(1).TEXT = Str$(CPAGCLI%(NC%))
       TEXT21(2).TEXT = Str$(VENDCLI%(NC%))
       '
       If XVALO(TEXT21(2).TEXT) < 1 Then
          Eco21(2).TEXT = "GERENCIA"
       End If
       '
       PORDESCU.TEXT = DESCUCLI$(NC%)
       '
       'AGREGUE MH02-10
       '\\\OT-07-0010-RG-12/01/07///
       Call FILTERFILE("SUCENTRE", "", 5, MKI$(NC%))
       '\\\OT-07-0010-RG-FIN///
       '
         For KKI% = 0 To 2
           LENTREGA(KKI%) = LUENTRECLI$(NC%, KKI%)
         Next KKI%
         Text19 = NUTRANCLI%(NC%)
       '
       If REGICLI(NC%) > 0 Then
         Call ABRECLIEN
         Set ClienTemp = Badaclie.OpenRecordset("SELECT * FROM Deudor12 WHERE Nume_Cli = " & NC%, dbOpenSnapshot)
         '
         On Error Resume Next '
         ClienTemp.MoveFirst
         If Err < 1 Then
           Text6 = ClienTemp!Nota_Cli
         End If
       End If
       MARCONOTA.Visible = True
       '
       '23/03/07 (OT-07-0121)'NO PERMITE GENERAR UN PEDIDO A UN CLIENTE CON LA
       LMCR# = LIMICRE(NC%)  'CUENTA SUSPENDIDA
       If LMCR# < 0 Then
          Call COMUNI("Imposible Generar Pedido.\ \La Cuenta esta Suspendida")
          Text1 = ""
          Text1.SetFocus
          GoTo 99
       End If
       '23/03/07 (FIN)
       '
       If VERLICRE% = 1 Then
         SDEU# = SALDCLI#(NC%)
         CHPA# = CHEPEACRE#(NC%)
         If SDEU# + CHPA# > LMCR# Then
           Call COMUNI("Este Pedido no Podrá Facturarse.\  \Límite de Crédito Asignado ya está Superado.\  \Informe al Supervisor de Ventas.")
         End If
       End If
       '
     End If
   End If
99 Screen.MousePointer = 1
   '
End Sub
   '
Function NUPEDIDO&()
   '
   NPP& = 1
   Call ABREPEDCLI
   '
   SQLX$ = "SELECT max(NroPed) as NUMAXI FROM Pedenca "
   Set PEDENCATEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   On Error Resume Next
   PEDENCATEMP.MoveFirst
   If Err < 1 Then NPP& = 1 + PEDENCATEMP!NUMAXI
   PEDENCATEMP.Close
   On Error GoTo 0
   '
   NUPEDIDO& = NPP&
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
     VH(MONEMI%) = XVALO(TICAMBIS)
   End If
End Sub

Private Sub TICAMBIS_LostFocus()
   TICAMBIS = FORMATNUM(XVALO(TICAMBIS), "F9.4")
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
      NPX& = XVALO(Mid$(YY$, 3, 6))
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
End Sub
'
Sub PRIPEDIDO(OPCI%)
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FORPECLI"))
    If FORIPRE$ = "" Then Exit Sub
    '
    If OPCI% = 2 Then ' ORDEN DE DESPACHO
       FORODES$ = TRIM$(CONTROL$("", "FORDESPA"))
       If FORODES$ <> "" Then
          FORIPRE$ = FORODES$
       End If
    End If
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
    If STATUPC% = (-1) Then
       If CONTROL$("PEDIDO", "IMPRIME") = "ALAPROBA" Then
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
    NC% = XVALO(Text1)
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
    If OPCI% = 2 Then
       TIPODOC$ = "Orden de Despacho"
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
    CODTABU1$(10) = "R-PLANO"
    CODTABU1$(11) = "F-PLANO"
    CACOLTAB1% = 11
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
          If XVALO(CAPEX$) > 0 Then
            '
            '\\\OT-07-0107-RG-15-03-07///
            PRELIX = PRELISTA#(ARCHIPRE$, CIPRE%)
            PRELI$ = Str$(PRELIX)
            '\\\OT-07-0107-RG-FIN///
            'PRELI$ = Str$(CVS(Mid$(XX$, 39, 4)))
            PREUX$ = Str$(CVS(Mid$(XX$, 51, 4)))
            DESCX$ = "0": On Error Resume Next
             DESCX$ = Str$(100 * (XVALO(PRELI$) - XVALO(PREUX$)) / XVALO(PRELI$))
             On Error GoTo 0
            ITOTX$ = FORMATNUM$((XVALO(CAPEX$)) * (XVALO(PREUX$)), "F13.2")
            FEX = CVI(Mid$(XX$, 63, 2))
            FENTX$ = FECHATEX$(FEX)
            DESXLXTA$ = FORMATNUM$(CVS(Mid$(XX$, 109, 4)), "F4.1")
            DESXGPO$ = FORMATNUM$(CVS(Mid$(XX$, 113, 4)), "F4.1")
               FORDESGRU$ = TRIM$(UCase$(ATRIFORM$(FORIPRE$, "F-PLANO/FORMATO")))
               If UCase$(Left$(FORDESGRU$, 1)) = "A" Then
                 DESXGPO$ = TRIM$(Mid$(XX$, 93, 16))
               End If
               '
            CAITAB1% = CAITAB1% + 1
            TETABU1$(1, CAITAB1%) = CODEXT$(CIPRE%)
            TETABU1$(3, CAITAB1%) = CAPEX$
            TETABU1$(4, CAITAB1%) = PRELI$
            TETABU1$(5, CAITAB1%) = DESCX$
            TETABU1$(6, CAITAB1%) = PREUX$
            TETABU1$(7, CAITAB1%) = ITOTX$
            If XVALO(ITOTX$) < 0.005 Then
              TETABU1$(7, CAITAB1%) = "S/Cargo"
            End If
            TETABU1$(8, CAITAB1%) = FENTX$
            TETABU1$(9, CAITAB1%) = Str$(CAITAB1%)
            TETABU1$(10, CAITAB1%) = DESXLXTA$
            TETABU1$(11, CAITAB1%) = DESXGPO$
            '
            ' LO QUE SIGUE ES PARA IMPRIMIR O/DESPACHO
            If OPCI% = 2 Then
               TETABU1$(4, CAITAB1%) = ""
               TETABU1$(5, CAITAB1%) = ""
               TETABU1$(6, CAITAB1%) = ""
               TETABU1$(7, CAITAB1%) = ""
            End If
            ' FIN O/DESPACHO
            '
            PHCODI% = XVALO(ATRIFORM$(FORIPRE$, "COD-MAT/POSINHOR"))
            PHDESI% = XVALO(ATRIFORM$(FORIPRE$, "DES-MAT/POSINHOR"))
            PVCODI% = XVALO(ATRIFORM$(FORIPRE$, "COD-MAT/POSINVER"))
            PVDESI% = XVALO(ATRIFORM$(FORIPRE$, "DES-MAT/POSINVER"))
            If PHDESI% <= PHCODI% + 3 Then
               If PVDESI% <= PVCODI% + 1 Then
                  CAITAB1% = CAITAB1% + 1
               End If
            End If
            DECRI$ = DESCRIT0$(CIPRE%)
            TETABU1$(2, CAITAB1%) = DECRI$
            '
            If PVDESI% <= PVCODI% + 1 Then
               FORTELIB$ = TRIM$(UCase$(ATRIFORM$(FORIPRE$, "DES-MAT/FORMATO")))
               ATELIB% = XVALO(Mid$(FORTELIB$, 2))
               Call FRATEXTO(DECRI$, ATELIB%)
               TETABU1$(2, CAITAB1%) = RETEX$(1)
               For KKU1% = 2 To CARETEX%
                 CAITAB1% = CAITAB1% + 1
                 TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
               Next KKU1%
            End If
            '
            COMPDESCRI$ = ""
            If CONTROL$("PEDIDO", "VARCOLOR") = "SI" Then
               COLORBA$ = TRIM(Mid$(XX$, 3, 12))
               TELA$ = TRIM(Mid$(XX$, 15, 12))
               COLTELA$ = TRIM(Mid$(XX$, 27, 12))
               COMPDESCRI$ = "CB: " + COLORBA$ + " - T: " + TELA$ + " - CT: " + COLTELA$
               CAITAB1% = CAITAB1% + 1
               TETABU1$(2, CAITAB1%) = COMPDESCRI$
            End If
            '
          End If
        End If
      End If
109 Next U&
    '
    SUMATOTA# = XVALO(Total(0).TEXT)
    IDESCUEN# = XVALO(Total(1).TEXT)
    '\\\OT-05-0853-WAR-24/11/05///
      If PIVACLI%(NC%) = 5 Or PIVACLI%(NC%) = 7 Then
        Label13 = "" 'SI CLIENTE EXTERIOR O PROMOVIDO
        Label14 = "" 'NO DISCRIMINA IVA
        Label20 = ""
      End If
    '\\\OT-05-0853-WAR-FIN///
    CODPARAM$(1) = "IMP-NETO": DOCPARAM$(1) = Str$(SUMATOTA#)
    CODPARAM$(2) = "VENDEDOR":     DOCPARAM$(2) = Eco21(2).TEXT
    CODPARAM$(3) = "COND-PAG":     DOCPARAM$(3) = Eco21(1).TEXT
    CODPARAM$(4) = "NRO-OCOM":     DOCPARAM$(4) = Text10.TEXT
    CODPARAM$(5) = "DOMI-TRA": DOCPARAM$(5) = LENTREGA(0).TEXT
    CODPARAM$(6) = "DESTINO": DOCPARAM$(6) = LENTREGA(1).TEXT
    CODPARAM$(7) = "TRANSPOR": DOCPARAM$(7) = LENTREGA(2).TEXT
    CODPARAM$(8) = "IMP-DESC": DOCPARAM$(8) = Total(1).TEXT
    CODPARAM$(9) = "IMPNEDES": DOCPARAM$(9) = Total(2).TEXT
    CODPARAM$(10) = "IMP-IVA": DOCPARAM$(10) = Label13 ' TOTIVA(3).TEXT
    CODPARAM$(11) = "IMP-TOTA": DOCPARAM$(11) = Total(3).TEXT
    CODPARAM$(12) = "POR-DESC": DOCPARAM$(12) = TRIM$(PORDESCU.TEXT)
    '\\\OT-5-869-28/12/05///
    'CODPARAM$(13) = "MONEDEMI": DOCPARAM$(13) = MONEMIS
    '\\\OT-5-869-FIN///
    CODPARAM$(13) = "PERC-IVA": DOCPARAM$(13) = Text15
    CODPARAM$(14) = "IMP-IVA1": DOCPARAM$(14) = Label14
    'AGREGADOMH04-10
    CODPARAM$(15) = "NRO-PASE": DOCPARAM$(15) = Label27
    '\\\OT-07-0007-RG-08/01/07///
    HORA$ = HORAENT$(NC%)
    CODPARAM$(16) = "Z-PLANO": DOCPARAM$(16) = TRIM$(HORA$)
    '\\\OT-07-0007-RG-FIN///
    '
    '\\\OT-07-0010-RG-12/01/07///
    AISUC% = XVALO(Text20)
    TELSUC$ = TELESUC$(NC%, AISUC%)
    CODPARAM$(17) = "REF-MAT1": DOCPARAM$(17) = TRIM$(TELSUC$)
    '\\\OT-07-0010-RG-FIN///
    '
    CODPARAM$(18) = "REF-MAT2": DOCPARAM$(18) = Format(Now, "dd/mm/yy - hh:mm")
    CODPARAM$(19) = "MEDIDA-1": DOCPARAM$(19) = FORMATNUM$(DEUCUECO#(NC%), "F12.2")
    CODPARAM$(20) = "MEDIDA-2": DOCPARAM$(20) = FORMATNUM$(CHEPEACRE#(NC%), "F12.2")
    CODPARAM$(21) = "MEDIDA-3": DOCPARAM$(21) = FORMATNUM$(VALOPEDPEN#(NC%), "F12.2")
    '
    ' TRANSPORTE
    NTRANI% = XVALO(Text19)
    CODPARAM$(22) = "COD-CUEN": DOCPARAM$(22) = DOMITRANS$(NTRANI%)
    CODPARAM$(23) = "DEN-CUEN": DOCPARAM$(23) = LOCATRANS$(NTRANI%)
    CODPARAM$(24) = "DET-PASE": DOCPARAM$(24) = TELETRANS$(NTRANI%)
    '
    CAPARDOC% = 24
    '
    ' LO QUE SIGUE ES PARA QUE EN CASO DE QUE HAYA UN SOLO CASILLERO
    ' DE IVA, TOTALICE TODOS LOS IMPORTES PARCIALES DE IVA EN UNO SOLO
    FORIVASEC$ = TRIM$(UCase$(ATRIFORM$(FORIPRE$, "IMP-IVA1/FORMATO")))
    If FORIVASEC$ = "" Then
      DOCPARAM$(10) = TOTIVA(3).TEXT
      DOCPARAM$(14) = ""
    End If
    '
    ' LO QUE SIGUE ES PARA ORDEN DE DESPACHO
    If OPCI% = 2 Then
       CODPARAM$(1) = ""
       CODPARAM$(8) = ""
       CODPARAM$(9) = ""
       CODPARAM$(10) = ""
       CODPARAM$(11) = ""
       CODPARAM$(12) = ""
       CODPARAM$(13) = ""
       CODPARAM$(14) = ""
       CODPARAM$(18) = ""
       CODPARAM$(19) = ""
       CODPARAM$(20) = ""
       CODPARAM$(21) = ""
    End If
    ' FIN ORDEN DE DESPACHO
    '
    Call BLANARCH("!OBSERVOF")
    ' NUEVO PARA LA IMPRESION DE OBSERVACIONES
    CODFOR$ = FORIPRE$
    FORMATOBS$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "OBSERVA/FORMATO")))
    ANCHOMA = XVALO(Mid$(FORMATOBS$, 2))
    ANCHOMAXU% = ANCHOMA
    Call FRATEXTO(Text13.TEXT, ANCHOMAXU%)
    For KU& = 1 To CARETEX%
       Call GRAREG("!OBSERVOF", RETEX$(KU&), KU&)
    Next KU&
    Call SETULTREG("!OBSERVOF", KU& - 1)
    
'    Call FRATEXTO(Text13.TEXT, 32)
'    If CARETEX% >= 1 Then Call GRAREG("!OBSERVOF", RETEX$(1), 1)
'    If CARETEX% >= 2 Then Call GRAREG("!OBSERVOF", RETEX$(2), 2)
'    If CARETEX% >= 3 Then Call GRAREG("!OBSERVOF", RETEX$(3), 3)
'    URETE& = CARETEX%
'    Call SETULTREG("!OBSERVOF", URETE&)
'    Call CIERRARCH("!OBSERVOF")
    '
    If FORIPRE$ <> "" Then
      'Call PRINTDOC("FORPECLI")   ' PEDIDO DE CLIENTE
      Call PRINTDOC("@" + FORIPRE$) ' PEDIDO DE CLIENTE
    End If
    '
    For KK1% = 1 To CACOLTAB1%
      For KK2% = 1 To CAITAB1%
        TETABU1$(KK1%, KK2%) = ""
      Next KK2%
    Next KK1%
    '
    'FORIPRE$ = TRIM$(CONTROL$("", "FORDESPA"))
    'If FORIPRE$ <> "" Then
    '  If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") > 0 Then
    '    TIPODOC$ = "Orden de Despacho"
    '    CANCELPRINT = 0
    '    If CONTROL$("ORDESPA", "SUPRINT") = "SI" Then
    '       CANCELPRINT% = 1
    '    End If
    '    '
    '    If InStr(EPAC$, "DOSIVAC") > 0 Then
    '      CANCELPRINT% = 0
    '      If OPRIPEDI.Check1(1).Value <> 1 Then
    '        CANCELPRINT% = 1
    '      End If
    '    End If
    '    '
    '    If CARETEX% >= 1 Then Call GRAREG("!OBSERVOF", RETEX$(1), 1)
    '    If CARETEX% >= 2 Then Call GRAREG("!OBSERVOF", RETEX$(2), 2)
    '    If CARETEX% >= 3 Then Call GRAREG("!OBSERVOF", RETEX$(3), 3)
    '    URETE& = CARETEX%
    '    Call SETULTREG("!OBSERVOF", URETE&)
    '    Call CIERRARCH("!OBSERVOF")
    '    '
    '    Call PRINTDOC("FORDESPA")   ' ORDEN DE DESPACHO
    '  End If
    'End If
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
   OKX% = 0
   If MARCONOTA.Visible = True Then Exit Sub
   If MARSELPRESU.Visible = True Then Exit Sub
   If XVALO(Total(2).TEXT) < 0.01 Then
     If COMALTER%("Este Pedido es SIN CARGO !!!\Está Seguro de Registrarlo Igual ?\\Cancelar Pedido\Registrar Igual") <> 2 Then
       Exit Sub
     End If
   End If
   If ULTREG&("!CARDETPE") < 1 Then Exit Sub
   '
   'VUELTA% = 0
'24 Call BLOQARCH("CARDETPE")
'   TINI = Timer: While Timer < TINI + 2: Wend
'   SQLX$ = "SELECT NroPed FROM Pedenca "
'   SQLX$ = SQLX$ + "WHERE NroPed = " & XVALO(Text11.TEXT)
'   Set PEDENCATEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
'   On Error Resume Next
'   With PEDENCATEMP
'     .MoveFirst
'     If Err < 1 Then
'        Call COMUNI("Número Repetido !!!")
'        NPP& = NUPEDIDO&
'        GoTo 25
'     End If
'   End With
'   PEDENCATEMP.Close

   ' VERIFICA NUMERO DE PEDIDO
   Call CIERRARCH("PEDENCA")
   Call CIERRARCH("PEDDETA")
   '
   Call BLOQARCH("PEDENCA")
   TINI = Int(Timer)
   While Int(Timer) < TINI + 5: Wend
   Call BLOQARCH("PEDENCA")
   '
   PedEnca.Close
   Call ABREPEDCLI
   PedEnca.FindFirst "NROPED = " & TRIM$(Text11.TEXT)
   If PedEnca.NoMatch = False Then
15    NPP& = XVALO(Text11.TEXT) + 1
      Call LIBARCH("PEDENCA")
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
   NCLIE% = XVALO(Text1.TEXT)
   '
   If XVALO(TEXT21(0).TEXT) > 255 Then
      MERRO$ = "Lista de Precios no Válida"
      GoTo 98
   End If
   '
   If XVALO(TEXT21(1).TEXT) > 255 Then
      MERRO$ = "Condición de Pago no Válida"
      GoTo 98
   End If
   '
   If XVALO(TEXT21(2).TEXT) > 255 Then
      MERRO$ = "Vendedor no Válido"
      GoTo 98
   End If
   '
   '\\\OT-06-0512-RG-27/11/06///
   If CONTROL$("ADMIPED", "OBLICOM") = "SI" Then
   '\\\OT-06-0512-RG-FIN///
    If TRIM$(Text10) = "" Then
       MERRO$ = "Falta Orden de Compra Cliente"
       GoTo 98
    End If
   End If
   '
   HOII% = HOY(HOS$)
   FEMI% = FECHANUM(Text9.TEXT)
   If FEMI% < 1 Or FEMI% > HOII% Then
      MERRO$ = "Fecha de Emisión no Válida"
      GoTo 98
   End If
   '
   NROPE& = XVALO(Text11.TEXT)
   REFE$ = " - " + TRIM$(Text10.TEXT)
   If REFE$ = " - " Then REFE$ = ""
   FEX = FEMI%
   REFE$ = REFE$ + " - " + FECHATEX$(FEX)
   LOMA% = 48 - Len(REFE$)
   REFE$ = Left$(TRIM$(ECOARCH$("DEUDOR1", MKI$(NCLIE%))), LOMA%) + REFE$
   '
   STATUPC% = (-1)
   If Left$(CONTROL$("PEDCLI", "MODOAPRO"), 6) = "DIRECT" Then
     STATUPC% = 0
   ElseIf Left$(CONTROL$("PEDCLI", "MODOAPRO"), 8) = "FACVENCI" Then
      CREDISPO# = LIMICRE(NCLIE%) - SALDCLI#(NCLIE%) - CHEPEACRE#(NCLIE%) - VALOPEDPEN1#(NCLIE%)
      If (CREDISPO# > 0) And (InStr(Text8, "Fac.Vencidas") < 1) Then
        STATUPC% = 0
      End If
   End If
   '
   PedEnca.FindFirst "NROPED = " & TRIM$(Text11.TEXT)
   If PedEnca.NoMatch = False Then GoTo 15
   '
   With PedEnca
     .AddNew
     !NroPed = NROPE&
     !Referencia = REFE$
     !FECHEMIS = CVDAT(FEMI%)
     !NroClie = NCLIE%
     !RasoClie = RASOCLI$(NCLIE%)
     !NroOcom = TRIM$(Text10.TEXT)
     !DOMIENT = LENTREGA(0)
     !LOCAENT = LENTREGA(1)
     !TranspEnt = LENTREGA(2)
     !NTranspo = XVALO(Text19)
     !NSucent = XVALO(Text20)
     !Observa = Text13
     !FeSolEnt = Now
       '\\\OT-05-0236-WAR-02/05/05///
     !Status = STATUPC%              ' Status  de Aprobación
       '\\\OT-05-0236-FIN///
     !ListPre = XVALO(TEXT21(0))
     !CondPag = XVALO(TEXT21(1))
     !VENDED = XVALO(TEXT21(2))
     !TDescApli = TRIM$(PORDESCU.TEXT)
     .Update
   End With
   '
   ANOTA$ = ""
   TTXYZ$ = Space$(76)
   HOII% = HOY(HOS$)
   Call REPLA(TTXYZ$, HOS$, 1, 8)
   Call REPLA(TTXYZ$, "Generación de Orden de Venta:", 11, 56)
   Call REPLA(TTXYZ$, USERTER$, 69, 24)
   ANOTA$ = ANOTA$ + TTXYZ$ + Chr$(10) + Chr$(13)
   '
   CAVER$ = TRIM$(Str$(NROPE&))
   While Len(CAVER$) > 6: CAVER$ = Mid$(CAVER$, 2): Wend
   While Len(CAVER$) < 6: CAVER$ = "0" + CAVER$: Wend
   IKITEM% = 0
   For U& = 1 To ULTREG&("!CARDETPE")
     
     X$ = REGLEIDO$("!CARDETPE", U&)
     CAN = CVS(Mid$(X$, 47, 4))
     If CAN > 0.05 Then
       '
       IKITEM% = IKITEM% + 1
       IK$ = TRIM$(Str$(IKITEM%))
       CAVERI$ = CAVER$ + "-" + IK$
       '
       ' ESTO ES PARA PONER EL PRECIO DE LISTA EN LA B.DATOS
       CIXI% = CVI(Left$(X$, 2))
       If CIXI% > 0 Then
          If Left$(UCase$(ARCHIPRE$), 5) = "LIPRE" Then
             PLISTA = PRELISTA#(ARCHIPRE$, CIXI%)
             Call REPLA(X$, MKS$(PLISTA), 39, 4)
          End If
       End If
       '
       COLORBA$ = "": TELA$ = "": COLTELA$ = "": SUBCODX$ = ""
       If CONTROL$("PEDIDO", "VARCOLOR") = "SI" Then
            COLORBA$ = TRIM(Mid$(X$, 3, 12))
            TELA$ = TRIM(Mid$(X$, 15, 12))
            COLTELA$ = TRIM(Mid$(X$, 27, 12))
          Else
            SUBCODX$ = TRIM(Mid$(X$, 3, 8))
       End If
     
       With PedDeta
         .AddNew
           !NroPed = NROPE&
           !FECHEMIS = CVDAT(FEMI%)
           !NroClie = NCLIE%
           !RasoClie = RASOCLI$(NCLIE%)
           !NroOcom = TRIM$(Text10.TEXT)
             FENTRESOL% = CVI(Mid$(X$, 63, 2))
           !FeSolEnt = CVDAT(FENTRESOL%)
           !Status = STATUPC%              ' Status  de Aprobación
           !ListPre = XVALO(TEXT21(0).TEXT)
           !CondPag = XVALO(TEXT21(1).TEXT)
           !VENDED = XVALO(TEXT21(2).TEXT)
           !PorComVend = XVALO(Text44.TEXT)
           !NuorItem = IKITEM%
           !CodiInt = CIXI%
           !CodiExt = CODEXT$(CIXI%)
           !SubCod = SUBCODX$
           !Descrip = DESCRIT0$(CIXI%)
           !TLargo = CVS(Mid$(X$, 27, 4))
           !TAncho = CVS(Mid$(X$, 31, 4))
           !TEspesor = CVS(Mid$(X$, 35, 4))
           !CBase = Left$(COLORBA$, 16)
           !TTalle = Left$(TELA$, 8)
           !TColor = Left$(COLTELA$, 12)
           !VenRep = Left$(Label11, 1)   ' Venta o Reparacion
           !MoneEmis = MONEMI%
           !TipoCam = XVALO(TICAMBIS)
           !PreList = CVS(Mid$(X$, 39, 4))
           !DescLista = CVS(Mid$(X$, 109, 4))
           !DescCombi = TRIM$(Mid$(X$, 93, 16))
           !DescRubro = CVS(Mid$(X$, 113, 4))
           !DescTotal = CVS(Mid$(X$, 69, 4))
           !CanAsig = CVS(Mid$(X$, 43, 4))
           !CANPED = CAN
           !PREUNID = CVS(Mid$(X$, 51, 4))
           !ImpoTot = CVD(Mid$(X$, 55, 8))
           !CantEnt = 0
           !CanSaldo = CAN
           !FeUltEnt = CVDAT(0)
           !NroRemi = ""
           !TPedItem = CAVERI$
           '
           TTXYZ$ = Space$(76)
           Call REPLA(TTXYZ$, "- " + CODEXT$(CIXI%), 11, 18)
           Call REPLA(TTXYZ$, DESCRIT0$(CIXI%), 29, 38)
           Call REPLA(TTXYZ$, CSTRING$(MKS$(CAN), 3, 10, 1, 2), 57, 10)
           ANOTA$ = ANOTA$ + TTXYZ$ + Chr$(10) + Chr$(13)
           '
         .Update
       End With
       '
       NUORIT% = CVI(Mid$(X$, 123, 2))
       NUPRESU& = CVS(Mid$(X$, 125, 4))
       If NUPRESU& > 0 Then
         Call ABREPRESVEN
         With PresuVen
           TBXX$ = "NroPres = " & NUPRESU&
           If NUORIT% > 0 Then
                TBXX$ = TBXX$ & " AND NuOrdIt = " & NUORIT%
              Else
                TBXX$ = TBXX$ & " AND CodiInt = " & CIXI%
           End If
           .FindFirst TBXX$
           If .NoMatch = False Then
             .Edit
             !FechVent = CVDAT(FEMI%)
             !NroPedi = NROPE&
             .Update
           End If
         End With
       End If
       '
     End If
   Next U&
   '
   ANOTA$ = ANOTA$ + Space$(10) + String$(56, "-") + Chr$(10) + Chr$(13)
   Call FRATEXTO(Text13.TEXT, 32)
   For KKI% = 1 To CARETEX%
     TTXX$ = Space$(92)
     Call REPLA(TTXX$, RETEX$(KKI%), 26, 32)
     If KKI% = 1 Then Call REPLA(TTXX$, "Observaciones:", 11, 15)
     ANOTA$ = ANOTA$ + TTXX$ + Chr$(10) + Chr$(13)
   Next KKI%
   ANOTA$ = ANOTA$ + String$(88, "=") + Chr$(10) + Chr$(13)
   '
   Call APERBASDAT("PEDCLI")
   SQLX$ = "SELECT ANOTAPED FROM PEDENCA"
   SQLX$ = SQLX$ + " WHERE NroPed = " & NROPE& & ""
   Set PEDDDENCA = Ventas.OpenRecordset(SQLX$, dbOpenDynaset)
   ANOTB$ = SAFETEXT$(PEDDDENCA!AnotaPed)
   ANOTA$ = ANOTB$ + ANOTA$
   '\\\OT-08-0912-OD-30/09/08/// PARA EVITAR QUE SE GUARDEN DATOS CON COMILLAS SIMPLES
   ANOTA$ = REPLACAR$(ANOTA$, "'", "")
   '\\\OT-08-0912-OD-FIN///
   SQLX$ = "UPDATE PEDENCA SET AnotaPed=" & "'" & ANOTA$ & "' WHERE NroPed=" & NROPE&
   Ventas.Execute (SQLX$)
   '
   Call ABRECLIEN
   With Deudor12
     .FindFirst "Nume_Cli = " & NCLIE%
     If .NoMatch = False Then
       .Edit
       !Lentre0_Cli = TRIM$(LENTREGA(0).TEXT)
       !Lentre1_Cli = TRIM$(LENTREGA(1).TEXT)
       !Lentre2_Cli = TRIM$(LENTREGA(2).TEXT)
       .Update
       STX% = STATCLI%(0)
     End If
   End With
   '
   Call LIBARCH("PEDENCA")
   Call CIERRARCH("*.*")
   Call LIBARCH("*.*")
   Call BAJALDISCO
   '
   OKX% = 1
   If VERLICRE% = 1 Then
     LMCR# = LIMICRE(NCLIE%)
     SDEU# = SALDCLI#(NCLIE%)
     CHPA# = CHEPEACRE#(NCLIE%)
     TPED# = XVALO(Total(3).TEXT)
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
98 Call LIBARCH("PEDENCA")
   '
   PedEnca.Close
   PedDeta.Close
   '
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
     TTT1 = XVALO(Total(0).TEXT)
     Total(0) = CSTRING$(MKS$(TTT1), 3, 11, 2, 2)
   End If
End Sub

Sub CARDETPEDI(MODO%)

   If Check1.Value = 1 Then ' si la opcion de pedidos x medida esta en true
      Call CargaPorMedida(OKEY2%)
      If OKEY2% < 1 Then Exit Sub
      GoTo 6
   End If
   '
   CLIEN% = XVALO(Text1.TEXT)
   NULISTA% = XVALO(TEXT21(Index).TEXT)
   '
   '\\\OT-08-0133-FL-3-4-08///
   If CONTROL$("PEDIDO", "VARCOLOR") = "SI" Then  ' si la opcion de pedidos por variante esta en true
      Call CIERRARCH("!CARDETP1")
      LFCDP$ = LOADFILE$(LUCOM$ + "CARDETPE.WKF")
      Call SAVEFILE(LUCOM$ + "CARDETP1.WKF", LFCDP$)
      If List1.ListCount < 1 And MODO% <> 1 Then ' si viene de presupuestos MODO% = 1
         Call SETULTREG("!CARDETP1", 0)
         Call SETULTREG("!CARDETPE", 0)
      End If
      ATRI$ = "/NC/ARCHECO(1)=" + ARCHIPRE$
      ATRI$ = ATRI$ + "/TITUPAN=Detalle Pedido de Cliente"
      ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Frame1.Top + Picture3.Top + 400))
      ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture3.Left))
3     DETPED% = VENTABU%("CARDETP1" + ATRI$)
      If DETPED% = -1 Then
         Text1.TEXT = ""
         Exit Sub 'Si se cancela vuelve a la ventana FORPED04 BLANQUEADA
      End If
      FIN& = ULTREG&("!CARDETP1")
      If FIN& < 1 Then
         Text1.TEXT = ""     'Si no se cargan registros al precionar ok
         Exit Sub            'vuelve a la ventana FORPED04 BLANQUEADA
      End If
      '
      For KKU& = 1 To FIN&
         X$ = REGLEIDO$("!CARDETP1", KKU&)
         CI% = CVI(Left$(X$, 2))
         PRECARGA = PRELISTA#(ARCHIPRE$, CI%)
         ' Trae el descuento de lista
         DESCULISTA# = ROUND(DESCULISPRE(NULISTA%), 1)
         KOEFILISTA# = 1 - DESCULISTA# / 100 ' Calcula el coeficiente de Dto. de lista
         ' Trae la combinacion del descuento que tiene el articulo/Cliente
         DESCOMBIN$ = DESCOMBINA$(CLIEN%, CI%)
         DESCUCOMBI# = ROUND(PORCEDESCU(DESCOMBIN$), 1)
         KOEFICOMBI# = 1 - DESCUCOMBI# / 100 'Calcula el coeficiente
         ' Determina el descuento general
         DESCUGENER# = ROUND#(100 * (1 - KOEFILISTA# * KOEFICOMBI#), 1)
         Call REPLA(X$, DESCOMBIN$, 93, 16)
         Call REPLA(X$, MKS$(DESCULISTA#), 109, 4)
         Call REPLA(X$, MKS$(DESCUCOMBI#), 113, 4)
         Call REPLA(X$, MKS$(DESCUGENER#), 69, 4)
         '
         If CVS(Mid$(X$, 39, 4)) < 0.005 Then
            Call REPLA(X$, MKS$(PRECARGA), 39, 4)
         End If
         Call GRAREG("!CARDETP1", X$, KKU&)
      Next KKU&
      Call CIERRARCH("!CARDETP1")
      Call CIERRARCH("!CARDETPE")
      '
      ' TRASLADA A !CARDETPE
      Call CIERRARCH("!CARDETPE")
      LFCDP$ = LOADFILE$(LUCOM$ + "CARDETP1.WKF")
      Call SAVEFILE(LUCOM$ + "CARDETPE.WKF", LFCDP$)
      '
   End If
   '\\\OT-08-0133-FL-FIN///
6  HOII% = HOY(HOS$)
   '
   ATRI$ = "/NC/ARCHECO(1)=" + ARCHIPRE$
   ATRI$ = ATRI$ + "/TITUPAN=Detalle Pedido de Cliente"
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Frame1.Top + Picture3.Top + 400))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture3.Left))
   '
   Timer1.Enabled = True
10 VENCARPE$ = CONTROL$("", "VENCARPE")
   If TRIM$(VENCARPE$) = "" Then VENCARPE$ = "CARDETP"
   PRF$ = VENCARPE$    ' If FORMAPED% = 2 Then PRF$ = "CARDETQ"
   XXX% = VENTABU%(PRF$ + ATRI$)
   
   If XXX% = -1 Then    ' no se aprobo la segunda ventana de carga
     If CONTROL$("PEDIDO", "VARCOLOR") = "SI" Then
          Call CIERRARCH("!CARDETPE")
          Call CIERRARCH("!CARDETP1")
          LFCDP$ = LOADFILE$(LUCOM$ + "CARDETPE.WKF")
          Call SAVEFILE(LUCOM$ + "CARDETP1.WKF", LFCDP$)
          GoTo 3
       Else
          Text1.TEXT = ""
          Exit Sub
     End If
   End If
   '
   If CONTROL$("PEDIDO", "VARCOLOR") = "SI" Then  ' si la opcion de pedidos por variante esta en true
      HUBOCA% = 0
      FIN& = ULTREG&("!CARDETPE")
      If ULTREG&("!CARDETP1") <> FIN& Then HUBOCA% = 1
      If HUBOCA% = 0 Then
         For KKU& = 1 To FIN&
            XXE$ = REGLEIDO$("!CARDETPE", KKU&)
            XXF$ = REGLEIDO$("!CARDETP1", KKU&)
            If CVI(Mid$(XXE$, 1, 2)) <> CVI(Mid$(XXF$, 1, 2)) Then HUBOCA% = 1 ' CAMBIÓ CODIGO
         Next KKU&
      End If
      Call CIERRARCH("!CARDETPE")
      Call CIERRARCH("!CARDETP1")
      If HUBOCA% > 0 Then
         Call CIERRARCH("!CARDETP1")
         Call COMUNI("Cambios Admisibles solo en Primera Ventana de Carga.\  \Debe Validar Cambios Introducidos.")
         LFCDP$ = LOADFILE$(LUCOM$ + "CARDETPE.WKF")
         Call SAVEFILE(LUCOM$ + "CARDETP1.WKF", LFCDP$)
         GoTo 3
      End If
   End If
   '
   Call CALTOPED
   If ECOPRECIO% >= 0 Then
     If XVALO(Total(2).TEXT) < 0.005 Then
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
     If CONTROL$("PEDIDO", "SUPRIREP") <> "SI" Then
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
   List1.Refresh
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
     NLIS% = XVALO(Mid$(ARCHIPRE$, 6, 3))
     CANDECI% = DECILISTA%(NLIS%)
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
     If ECOPRECIO% >= 0 Then Call REPLA(TTXX$, CSTRING$(MKS$(PRELIX), 3, 8, CANDECI%, 2), 52, 8)
     If ECOPRECIO% >= 0 Then Call REPLA(TTXX$, CSTRING$(MKS$(PORDESCUIN), 3, 5, 1, 2), 61, 5)
     If ECOPRECIO% >= 0 Then Call REPLA(TTXX$, CSTRING$(MKS$(PREUNI), 3, 8, CANDECI%, 2), 66, 8)
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
     CLIA$ = Text1.TEXT: NCLIE% = XVALO(CLIA$)
     RNI% = 0: APLIVA% = 1
     If NCLIE% > 0 Then
       PVCLI% = PIVACLI%(NCLIE%)
       If PVCLI% = 2 Then RNI% = 1
       If PVCLI% = 5 Or PVCLI% = 7 Then APLIVA% = 0
     End If
   End If
   '
   SUTOTPED# = 0: SUIVA# = 0: SUIVARNI# = 0: TOIVA# = 0: TOTOTA# = 0
     Dim IMPTASIVA#()
     ReDim IMPTASIVA#(ULTREG&("GRUPIVA"))
     DESC0$ = TRIM$(PORDESCU.TEXT)
     COEF1# = 1 - COEFDESCU(DESC0$) / 100
   TOTITEMS% = 0
   For U& = 1 To ULTREG&("!CARDETPE")
     X$ = REGLEIDO$("!CARDETPE", U&)
     CIXXI% = CVI(Left$(X$, 2))
     CAN = CVS(Mid$(X$, 47, 4))
     PORDESCUEIN = CVS(Mid$(X$, 69, 4))
     If PORDESCUEIN < 0 Or PORDESCUEIN > 99 Then PORDESCUEIN = 0
     '
     COEFITO# = (100 - PORDESCUEIN) / 100
     COEFILI# = (100 - CVS(Mid$(X$, 109, 4))) / 100
     PDESCGRU = PORDESCUEIN
     If COEFILI# > 0 Then
       COEFIGR# = COEFITO# / COEFILI#
       PDESCGRU = 100 * (1 - COEFIGR#)
     End If
     Call REPLA(X$, MKS$(PDESCGRU), 113, 4)
     Call GRAREG("!CARDETPE", X$, U&)
     '
     If CAN > 0.05 Then
       PRELI = CVS(Mid$(X$, 39, 4))
       'If PRELI < 0.01 Then
       '  PRELI = PRELISTA(ARCHIPRE$, CIXXI%)
       'End If
       PREUNI = PRELI * (100 - PORDESCUEIN) / 100
       TOT# = CDbl(CAN * PREUNI)
       SUTOTPED# = SUTOTPED# + TOT#
       SUIVA# = SUIVA# + (Int(TOT# * PORCEIVA(CIXXI%) + 0.5)) / 100
       '\\\OT-05-0560-WAR-03/08/05///
       TODE1# = TOT# * COEF1#
       GRIVA% = GRUPIVA%(CIXXI%)
       IMPTASIVA#(GRIVA%) = IMPTASIVA#(GRIVA%) + (Int(TODE1# * PORCEIVA(CIXXI%) + 0.5)) / 100
       Label13 = CSTRING$(MKD$(IMPTASIVA#(0)), 3, 11, 2, 2)
       Label14 = CSTRING$(MKD$(IMPTASIVA#(1)), 3, 11, 2, 2)
       Label20 = CSTRING$(MKD$(IMPTASIVA#(2)), 3, 11, 2, 2)
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
   TTT1 = XVALO(Total(0).TEXT) * COEFDESCU(DESC$) / 100
   Total(1).TEXT = CSTRING$(MKS$(TTT1), 3, 11, 2, 2)
   '
   TTT3 = XVALO(Total(0).TEXT) - XVALO(Total(1).TEXT)
   Total(2).TEXT = CSTRING$(MKS$(TTT3), 3, 11, 2, 2)
   '
   On Error Resume Next
   TOTIVA(0).TEXT = CSTRING$(MKD$(APLIVA% * SUIVA# * XVALO(Total(2).TEXT) / XVALO(Total(0).TEXT)), 4, 11, 2, 2)
   On Error GoTo 0
   TOTIVA(2).TEXT = CSTRING$(MKD$(CDbl(XVALO(TOTIVA(0).TEXT) * 0.5 * RNI%)), 4, 11, 2, 2)
   TOTIVA(3).TEXT = CSTRING$(MKD$(CDbl(XVALO(TOTIVA(0).TEXT) + XVALO(TOTIVA(2).TEXT))), 4, 11, 2, 2)
   '
   Call CALPERIB
   '
   Total(3).TEXT = CSTRING$(MKD$(CDbl(XVALO(Total(2).TEXT) + XVALO(TOTIVA(3).TEXT)) + XVALO(Text15)), 4, 11, 2, 2)
   '
End Sub

Private Sub TOTIVA_Change(Index As Integer)
   'If Index = 3 Then
   '  Total(3).TEXT = Str$(XVALO(Total(2).TEXT) + XVALO(TOTIVA(3).TEXT))
   'End If
End Sub

Sub GENERODES()
   '
   If CONTROL$("ADMIPED", "GENDESPA") = "REVISTOCK" Then
     Call GENODESPA(XVALO(Text11), NODES&, 1, OKKY%)
     If OKKY% = 1 Then
        Call PRIODESPA(NODES&)
     End If
   End If
   '
End Sub
'
Sub CALPERIB()
   '
   NCLIE% = XVALO(Text1)
   TASAX = ALIPERBRU#(NCLIE%, MINIPER#, VARIDOC%)
   '
   Text15 = ""
   BIMPON# = XVALO(Total(2))
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
'NUEVO SUB
Private Sub MODOCON(MODO%)
    Call CIERRARCH("*.*")
    MODX% = MODO%
    If MODX% = 1 Then
         Call CONECRUN("LISPRE04", "Listas de Precios de Venta")
       ElseIf MODX% = 2 Then
         Call CARCONPAGOS
    End If
End Sub

Sub IMPORZET(ORIPED$)
   '
   Dim NUPEDI&(), CLIPEDI%(), FEPEDI%()
   Dim NPEDIT&(), NORIT%(), CODIT%()
   CAENCA% = 0
   CADETA% = 0
   '
   If EXISTE%(ORIPED$) < 1 Then
     MERRO$ = REPLACAR$(ORIPED$, "\", "/")
     Call MENSERR(24, "Imposible Importar.\No Existe la B.D. Referenciada\'" + MERRO$ + "'.")
     Exit Sub
   End If
   '
   FECORTE% = FECHANUM(CONTROL$("PEDIDO", "FECORTE"))
   '
   Dim OEMPAQ As Database
   Dim ENCAOEM As Recordset
   Dim DETAOEM As Recordset
   Dim PARCIOEM As Recordset
   '
   Set OEMPAQ = OpenDatabase(ORIPED$)
   Set ENCAOEM = OEMPAQ.OpenRecordset("ORDENESEMPAQUE", dbOpenDynaset)
   Set DETAOEM = OEMPAQ.OpenRecordset("DETALLESOE", dbOpenDynaset)
   Set PARCIOEM = OEMPAQ.OpenRecordset("OEPARCIALES", dbOpenDynaset)
   '
   On Error Resume Next
   ENCAOEM.MoveLast
   If Err Then
     On Error GoTo 0
     Call MENSERR(24, "Imposible Importar. Tabla 'OrdenesEmpaque' Vacía.")
     Exit Sub
   End If
   On Error GoTo 0
   FIN& = ENCAOEM.RecordCount
   '
   Screen.MousePointer = 11
   With ENCAOEM
      .MoveFirst
      KUK& = 0
      PRIPEDI& = 999999
      Do While Not .EOF
        KUK& = KUK& + 1
        Call TRACE(20, KUK&, FIN&)
        If CVINT(!FechaHoraEmisión) >= FECORTE% Then
           ' validar cliente
           If !CodClien < 1 Or !CodClien > 32767 Then
              MERRO$ = "OE " & !NumOE & " - Cliente Número " & !CodClien & " no Válido."
              Call ADDMERR(MERRO$)
              GoTo 12
           End If
           If STATCLI%(!CodClien) < 1 Then
              MERRO$ = "OE " & !NumOE & " - Cliente Número " & !CodClien & " Inexistente o Inactivo (de Baja)."
              Call ADDMERR(MERRO$)
           End If
           ' fin validar cliente
           '
12         CAENCA% = CAENCA% + 1
           ReDim Preserve NUPEDI&(CAENCA%), CLIPEDI%(CAENCA%), FEPEDI%(CAENCA%)
           NUPEDI&(CAENCA%) = !NumOE
             If !NumOE < PRIPEDI& Then
               PRIPEDI& = !NumOE
             End If
           CLIPEDI%(CAENCA%) = !CodClien
           FEPEDI%(CAENCA%) = CVINT(!FechaHoraEmisión)
        End If
      .MoveNext
      Loop
   End With
   '
   On Error Resume Next
   DETAOEM.MoveLast
   If Err Then
     On Error GoTo 0
     Call MENSERR(24, "Imposible Importar. Tabla 'DetallesOE' Vacía.")
     Exit Sub
   End If
   On Error GoTo 0
   FIN& = DETAOEM.RecordCount
   '
   With DETAOEM
      .MoveFirst
      KUK& = 0
      Do While Not .EOF
        KUK& = KUK& + 1
        Call TRACE(20, KUK&, FIN&)
        NUPED& = !NumOE
        If NUPED& >= PRIPEDI& Then
           NUIT% = !NumItem
           codex$ = !CodProd
           CIXI% = CODINT%(codex$)
           ' validar articulo
           If CIXI% < 1 Then
              MERRO$ = "Codigo '" & codex$ & "' Inexistente o no Válido."
              Call ADDMERR(MERRO$)
           End If
           ' fin validar articulo
           '
           CANTU = !CANTIDAD
           NCLIE% = 0: FPED% = 0
           For KKU& = 1 To CAENCA%
             If NUPEDI&(KKU&) = NUPED& Then
               NCLIE% = CLIPEDI%(KKU&)
               FPED% = FEPEDI%(KKU&)
               GoTo 15
             End If
           Next KKU&
           MERRO$ = "OE " & NUPED& & " Código '" & codex$ & "' Falta Encabezado de OE."
           Call ADDMERR(MERRO$)
           GoTo 19      ' NO ENCONTRO PEDIDO
           '
15         If NCLIE% > 0 Then
             CADETA% = CADETA% + 1
             ReDim Preserve NPEDIT&(CADETA%), NORIT%(CADETA%), CODIT%(CADETA%)
             NPEDIT&(CADETA%) = NUPED&
             NORIT%(CADETA%) = NUIT%
             CODIT%(CADETA%) = CIXI%
             Call AGREPED(NCLIE%, NUPED&, NUIT%, FPED%, CIXI%, CANTU)
           End If
        End If
        '
19    .MoveNext
      Loop
   End With
   '
   ' FIN DE IMPORTACION DE PEDIDOS
   On Error Resume Next
   PARCIOEM.MoveLast
   If Err Then
     On Error GoTo 0
     Call MENSERR(24, "Imposible Importar. Tabla 'OEParciales' Vacía.")
     Exit Sub
   End If
   On Error GoTo 0
   FIN& = PARCIOEM.RecordCount
   '
   With PARCIOEM
      .MoveFirst
      KUK& = 0
      Do While Not .EOF
        KUK& = KUK& + 1
        Call TRACE(20, KUK&, FIN&)
        NUPED& = !NumOE
        FENTREG% = CVINT(!FechaEntrega)
        CANTU = !CANTIDAD
        '
        If NUPED& >= PRIPEDI& Then
           NUIT% = !NumItem
           NORE% = !NumeroEntrega
           '
           CIXI% = 0
           For KKU& = 1 To CADETA%
             If NPEDIT&(KKU&) = NUPED& Then
               If NORIT%(KKU&) = NUIT% Then
                 CIXI% = CODIT%(KKU&)
                 GoTo 25
               End If
             End If
           Next KKU&
           MERRO$ = "Parcial OE " & NUPED& & " Item " & NUIT% & " Orden " & NORE% & " Falta Encabezado de OE."
           Call ADDMERR(MERRO$)
           GoTo 29      ' NO ENCONTRO PEDIDO
           '
25         If CIXI% > 0 Then
             Call AGREMOVST(NUPED&, NUIT%, NORE%, FENTREG%, CIXI%, CANTU)
           End If
        End If
        '
29    .MoveNext
      Loop
   End With
   '
   Call CIERRARCH("LOGIMPOR")
   Screen.MousePointer = 1
   '
   If ULTREG&("LOGIMPOR") < 1 Then
       Call COMUNI("Fin de Importación de Ordenes de Entrega.")
     Else
       OPX% = COMALTER%("Fin de Importación de OE's con Errores.\¿ Desea Consultar LOG-File Correspondiente ?\\Si, Consultar\No, Terminar")
       If OPX% = 1 Then
         Call FINAL("*LOGIMPOR")
       End If
   End If
   '
End Sub

Sub AGREPED(NCLIE%, NUPED&, NUIT%, FPED%, CIXI%, CANTU)
   '
   Call ABREPEDCLI
   With PedEnca
     .FindFirst "NroPed = " & NUPED&
     If .NoMatch = False Then
        ' ENCONTRO ENCABEZADO DE PEDIDO
        If !NroClie <> NCLIE% Then
          MERRO$ = "OE Número " + TRIM$(Str$(NUPED&)) + " no Coincide Cliente"
          Call ADDMERR(MERRO$)
        End If
        GoTo 20
     End If
     '
     STATUPC% = (-1)
     If Left$(CONTROL$("PEDCLI", "MODOAPRO"), 6) = "DIRECT" Then STATUPC% = 0
     '
     .AddNew
     !NroPed = NUPED&
     '!Referencia = REFE$
     !FECHEMIS = CVDAT(FPED%)
     !NroClie = NCLIE%
     !RasoClie = RASOCLI$(NCLIE%)
     !NroOcom = "OE " + TRIM$(Str$(NUPED&))
     !DOMIENT = LUENTRECLI$(NCLIE%, 0)
     !LOCAENT = LUENTRECLI$(NCLIE%, 1)
     !TranspEnt = LUENTRECLI$(NCLIE%, 2)
     !NTranspo = 0
     !NSucent = 0
     !Observa = ""
     !FeSolEnt = CVDAT(FPED%)
     !Status = STATUPC%              ' Status  de Aprobación
     !ListPre = LIPRCLI%(NCLIE%)
     !CondPag = CPAGCLI%(NCLIE%)
     !VENDED = VENDCLI%(NCLIE%)
     !TDescApli = DESCUCLI$(NCLIE%)
     .Update
   End With
   '
20 With PedDeta
      .FindFirst "NroPed = " & NUPED& & " AND NuOrItem = " & NUIT%
      If .NoMatch = False Then
         ' VERIFICAR QUE COINCIDAN DATOS
         If !NroClie = NCLIE% Then
           If !FECHEMIS = CVDAT(FPED%) Then
             If !CodiInt = CIXI% Then
               If Abs(!CANPED - CANTU) < 0.05 Then
                 GoTo 99
               End If
             End If
           End If
         End If
         MERRO$ = "OE " & NUPED& & " Código '" & codex$ & "' Pre-Importada no Coinciden Datos."
         Call ADDMERR(MERRO$)
         GoTo 99
      End If
'   '
'   ANOTA$ = ""
'   TTXYZ$ = Space$(76)
'   HOII% = HOY(HOS$)
'   Call REPLA(TTXYZ$, HOS$, 1, 8)
'   Call REPLA(TTXYZ$, "Modificación de Orden de Venta:", 11, 56)
'   Call REPLA(TTXYZ$, USERTER$, 69, 24)
'   ANOTA$ = ANOTA$ + TTXYZ$ + Chr$(10) + Chr$(13)
'   '
'
      .AddNew
         !NroPed = NUPED&
         !FECHEMIS = CVDAT(FPED%)
         !NroClie = NCLIE%
         !RasoClie = RASOCLI$(NCLIE%)
         !NroOcom = "OE " + TRIM$(Str$(NUPED&))
         !FeSolEnt = CVDAT(FPED%)
         !Status = STATUPE%   'idem anterior.
         !ListPre = LIPRCLI%(NCLIE%)
         !CondPag = CPAGCLI%(NCLIE%)
         !VENDED = VENDCLI%(NCLIE%)
         !PorComVend = 0
         !NuorItem = NUIT%
         !CodiInt = CIXI%
         !CodiExt = CODEXT$(CIXI%)
         !SubCod = ""
         !Descrip = DESCRIT0$(CIXI%)
         !TColor = ""
         !TTalle = ""
         !TLargo = 0
         !TAncho = 0
         !TEspesor = 0
         !VenRep = "V"
         !MoneEmis = 1
         !TipoCam = 1
         !PreList = 0 ' CVS(Mid$(X$, 39, 4))
         !CanAsig = 0
         !CANPED = CANTU
         !PREUNID = 0  ' CVS(Mid$(X$, 51, 4))
         !ImpoTot = 0  ' CVD(Mid$(X$, 55, 8))
         '   CAENTAN = CVS(Mid$(X$, 117, 4))   ' CANTIDAD ENTREGAS ANTERIORES
         '   FULENTAN% = CVI(Mid$(X$, 121, 2)) ' FECHA ULTIMA ENTREGA
         !CantEnt = 0 ' CAENTAN
         !CanSaldo = CANTU   ' CAN - CAENTAN
         !FeUltEnt = CVDAT(0)
         !NroRemi = ""
            CAVER$ = TRIM$(Str$(NUPED&))
            While Len(CAVER$) > 6: CAVER$ = Mid$(CAVER$, 2): Wend
            While Len(CAVER$) < 6: CAVER$ = "0" + CAVER$: Wend
            IK$ = TRIM$(Str$(NUIT%))
            CAVERI$ = CAVER$ + "-" + IK$
         !TPedItem = CAVERI$
         '
      .Update
   End With
   '
99 End Sub

Sub AGREMOVST(NUPED&, NUIT%, NORE%, FECHA%, CICONJ%, CANTU)
   '
   CMOV$ = "OE"
   IDMO$ = TRIM$(Str$(10 * NUIT% + NORE%))
   While Len(IDMO$) < 3: IDMO$ = "0" + IDMO$: Wend
   DOPRILA$ = "OE-" + TRIM$(Str$(NUPED&)) + "-" + IDMO$
   NUPED1& = NUPED&: While NUPED1& > 9999: NUPED1& = NUPED1& - 10000: Wend
   DOPRI$ = "OE" + TRIM$(Str$(NUPED1&)) + "-" + IDMO$
   DOSELA$ = DOPRILA$
   DOSEC$ = DOPRI$
   LOTE$ = DOPRILA$
   REFE$ = "Orden Empaque " + TRIM$(Str$(NUPED&))
   VUNI# = 0
   MONEII% = 1
   NCX% = 0
   DEPX% = DEPOEXPE%
   '
   Call ABRESTOCKS
   With Movisto
     .FindFirst "CodiMovi = '" & CMOV$ & "' AND DoPriLar = '" & DOPRILA$ & "'"
     If .NoMatch = False Then
        ' ENCONTRO MOVIMIENTO - FALTA VALIDAR
        If CVINT(!fechMov) = FECHA% Then
          If !Cod_Inte = CICONJ% Then
            If Abs(!CantIngre - CANTU) < 0.05 Then
              GoTo 20
            End If
          End If
        End If
        MERRO$ = "Parciales OE " & NUPED& & " Código '" & CODEXT$(CICONJ%) & "' Pre-Importada no Coinciden Datos."
        Call ADDMERR(MERRO$)
        GoTo 20
     End If
     '
     Call MOVISTOK(FECHA%, CICONJ%, LOTE$, CMOV$, DOPRILA$, DOPRI$, DOSELA$, DOSEC$, REFE$, 0, VUNI#, MONEII%, NCX%, DEPX%, CANTU)
   End With
   '
20 End Sub

Sub ADDMERR(MERRO$)

   Static CTX1%
   If CTX1% < 1 Then
     Call SETULTREG("LOGIMPOR", 0)
     CTX1% = 1
   End If
   '
   Call REGAPP("LOGIMPOR", MERRO$)
   '
End Sub

Sub GENORFA()
   '
   For U& = 1 To ULTREG&("!CARDETPE")
     X$ = REGLEIDO$("!CARDETPE", U&)
     CAN = CVS(Mid$(X$, 47, 4))
     If CAN > 0.05 Then
        CIXI% = CVI(Left$(X$, 2))
        CANSTO = STOCKACT(CIXI%)
        If CAN > CANSTO + 0.05 Then
           CODE$ = TRIM$(CODEXT$(CIXI%))
           Call LEEEXPLO(CODE$, 1)
           If CAIT% > 0 Then
             TTXX$ = "Stock Insuficiente para\'" + CODE$ + " - " + TRIM$(DESCRIT0$(CIXI%)) + "'"
             TTXX$ = TTXX$ + "\   \¿ Desea Generar una Orden de Fabricación ?"
             If COMALTER%(TTXX$ + "\\Si, Generar O/F\No, Continuar") = 1 Then
             '\\\OF-08-0600-FL-31/07/08///
                ACONEC$ = "ORFABR04/ORFNUE" + Chr$(124) + CODE$ + Chr$(124) + TRIM$(Str$(CAN)) + Chr$(124) + "PC " + TRIM$(Text11.TEXT) + " - " + FANTCLI$(XVALO(Text1.TEXT))
             '\\\OF-08-0600-FL-FIN///
                Call CONECRUN(ACONEC$, "Ordenes de Fabricación")
             End If
           End If
        End If
     End If
   Next U&
   '
End Sub
'
Sub CargaPorMedida(OKEY2%) '\\OT-07-0420
'
    OKEY2% = 0
    HOII% = HOY(HOS$)
    '***********************************************************************************************
    'VTB% = VENTABU%("VENCICAI/NC/TITUPAN=Tabla de Vencimientos de CAI")
    ATRI$ = "PEDMEDID/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Carga de Presupuesto por Medida"""
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Frame1.Top + Picture3.Top + 600))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture3.Left + 50))
    
    'VTB% = VENTABU("PEDMEDID/NC/TITUPAN=Carga de Presupuesto por Medida")  'LA VENTANA QUE SE ABRE con los datos y al confirmala seguimos aca abajo
10  VTB% = VENTABU(ATRI$)  'LA VENTANA QUE SE ABRE con los datos y al confirmala seguimos aca abajo
    If VTB% >= 0 Then
        FIN& = ULTREG&("!PEDMEDID") 'NUMERO DE ULTIMO REGISTRO
        List1.Clear
         
        ' VALIDACIONES
        For U& = 1 To FIN& ' dentro de este form se van leyendo los campos que se ingresaron en la ventana
           Y$ = REGLEIDO$("!PEDMEDID", U&) 'NOS DEVUELVE EL REGISTRO LEIDO EN Y$
           CODFAM$ = TRIM$(Mid$(Y$, 1, 2)) 'MID$ le pasamos como arg.en y$ el registro en 1 donde comienza y en 2 los bytes de acuerdo a la estructura
           M1% = CVI(Mid$(Y$, 3, 2)) ' aqui con cvi paara que devuel. un numero decimal
           M2% = CVI(Mid$(Y$, 5, 2))
           M3% = CVI(Mid$(Y$, 7, 2))
           VARIANTE$ = TRIM$(Mid$(Y$, 9, 4))
           CAN = CVS(Mid$(Y$, 13, 4))
           FECHA% = CVI(Mid$(Y$, 17, 2)) 'CVI pasaba la fecha a un numero
           '
           If CODFAM$ = "" Then
                 Call MENSERR(24, "Falta Código de Familia")
                 GoTo 10
              ElseIf ECOARCH$("TAFAMFIL", CODFAM$) = "" Then
                 Call MENSERR(24, "Código de Familia no Válido Según Tabla")
                 GoTo 10
           End If
           '
           If CDbl(M1) * CDbl(M2) < 1 Then
               Call MENSERR(24, "Faltan Medidas Básicas")
               GoTo 10
           End If
           '
           If CAN < 0.5 Then
               Call MENSERR(24, "Falta Cantidad")
               GoTo 10
           End If
           '
           CODIGO$ = CODFAM$ & M1 & M2   ' concateno el nuevo codigo
           If M3 > 0 Then CODIGO$ = CODIGO & M3   ' concateno el nuevo codigo
           If VARIANTE$ <> "" Then CODIGO$ = CODIGO$ + "-" + VARIANTE$
           If Len(CODIGO$) > 15 Then
              Call MENSERR(24, "Imposible Registrar Código '" + CODIGO$ + "'")
              GoTo 10
           End If
        Next U&
        '
        '
        For U& = 1 To FIN&
           Y$ = REGLEIDO$("!PEDMEDID", U&) 'NOS DEVUELVE EL REGISTRO LEIDO EN Y$
           CODFAM$ = TRIM$(Mid$(Y$, 1, 2)) 'MID$ le pasamos como arg.en y$ el registro en 1 donde comienza y en 2 los bytes de acuerdo a la estructura
           M1% = CVI(Mid$(Y$, 3, 2)) ' aqui con cvi paara que devuel. un numero decimal
           M2% = CVI(Mid$(Y$, 5, 2))
           M3% = CVI(Mid$(Y$, 7, 2))
           VARIANTE$ = TRIM$(Mid$(Y$, 9, 4))
           CAN = CVS(Mid$(Y$, 13, 4))
           FECHA% = CVI(Mid$(Y$, 17, 2)) 'CVI pasaba la fecha a un numero
           '
           CODIGO$ = CODFAM$ & M1 & M2   ' concateno el nuevo codigo
           If M3 > 0 Then CODIGO$ = CODIGO & M3   ' concateno el nuevo codigo
           If VARIANTE$ <> "" Then CODIGO$ = CODIGO$ + "-" + VARIANTE$
           '
30         If CODINT%(CODIGO$) < 1 Then
              OKEY% = 0
              Call CREACODIGO(CODFAM$, M1%, M2%, M3%, VARIANTE$, OKEY%)
              If OKEY% > 0 Then GoTo 30
              Call FINAL("")
           End If
           CII% = CODINT%(CODIGO$)
           '
           Call LEEEXPLO(CODIGO$, 1)
           If CAIT% < 1 Then
              Call CREAFORMULA(CII%, CODFAM$, M1%, M2%, M3%, VARIANTE$)
           End If
           '
           NLISTA% = XVALO(TEXT21(0).TEXT)
           If NLISTA% > 0 Then
              If PRELISTA#(ARCHIPRE$, CII%) < 0.005 Then
                 ' CALCULA EL PRECIO DE VENTA
                 PREUNI# = PRECALCULADO(CODFAM$, M1, M2)
                 TTXX$ = REGBLAN$("LIPREVEN") 'EL ESPACIO DESTINADO A LA VARIABLE
                 'ARMAMOS EL REGISTRO
                 Call REPLA(TTXX$, MKI$(CII%), 1, 2)
                 Call REPLA(TTXX$, DENOMI$, 3, 42)
                 Call REPLA(TTXX$, MKD$(PREUNI#), 62, 8)
                 Call FILTERPUTRECORD(ARCHIPRE$, 1, MKI$(CII%), TTXX$)
                 Call CIERRARCH(ARCHIPRE$)
                 PREUNI2 = PRELISTA(ARCHIPRE$ + "/NEW", CII%)
              End If
           End If
           '
           YY$ = REGBLAN$("CARDETPE")
           Call REPLA(YY$, MKI$(CII%), 1, 2)
           PREUNI = PRELISTA(ARCHIPRE$ + "/NEW", CII%)
           Call REPLA(YY$, MKS$(PREUNI), 39, 4)
           Call REPLA(YY$, MKS$(CAN), 47, 4)
           Call REPLA(YY$, MKI$(FECHA%), 63, 2)
           Call GRAREG("!CARDETPE", YY$, U&)
        Next U&
        Call SETULTREG("!CARDETPE", FIN&)
        OKEY2% = 1
        '
    End If
    Call CIERRARCH("PEDMEDID")
    '\\OT-07-0420
End Sub


Sub CREACODIGO(CODFAM$, M1%, M2%, M3%, VARIANTE$, OKEY%)
    '
    'en este Procedimiento se genero el codigo por las medidas
    'se verifico si existia ,si no (1) o (2)si existe
    '1) se creaba la descripcion en base  a la familia y a las medidas
    'y se graba y se lista en el list
    '2)si no se listaba  directamente
    '3)Si no existe el codigo en la lista se agrega de acuerdo  a la Nro.de lista del cliente
    '
    OKEY% = 0
    '
    ' ARMA EL NUEVO CODIGO
    CODIGO$ = CODFAM$ & M1 & M2   ' concateno el nuevo codigo
    If M3 > 0 Then CODIGO$ = CODIGO & M3   ' concateno el nuevo codigo
    If VARIANTE$ <> "" Then CODIGO$ = CODIGO$ + "-" + VARIANTE$
    '
    ' ARMA LA NUEVA DESCRIPCION
    DENOMI$ = TRIM(ECOARCH$("TAFAMFIL", CODFAM$)) 'dev.el segundo campo pasandole el primero como argumento
    DENOMI$ = DENOMI$ & " " & M1 & " x " & M2
    If M3 > 0 Then DENOMI$ = DENOMI$ & " x " & M3
    DENOMI$ = DENOMI$ & " mm." 'concateno la descripcion
    If VARIANTE$ <> "" Then DENOMI$ = DENOMI$ + " " & VARIANTE$
    '
    ' AGREGA EL CODIGO AL MAESTRO DE ARTICULOS
    XX$ = FILTERGETRECORD$("TAFAMFIL", 1, CODFAM$)
    TIMA% = Asc(Mid$(XX$, 30, 1))
    Call AGRECOD(CODIGO$, DENOMI$, TIMA%, "U.", OKX%) 'se graban los datos
    If OKX% < 1 Then     ' FALLO LA CREACIÓN
       Exit Sub
    End If
    CII% = CODINT%(CODIGO$)
    '
    OKEY% = 1
    '
End Sub
           '
Sub CREAFORMULA(CII%, CODFAM$, M1%, M2%, M3%, VARIANTE$)
    '
    'ARMAMOS LA ESTRUCTURA DE PRODUCTO
20  Call FILTERFILE("STRUCFAM", "!STRUCFAM", 1, CODFAM$)
    CAIT% = ULTREG&("!STRUCFAM")
    For U& = 1 To CAIT%
       XX$ = REGLEIDO$("!STRUCFAM", U&)
       CODCOMPO% = CVI(Mid$(XX$, 3, 2))
       CANCOMPO = CVD(Mid$(XX$, 5, 8))
       MODOCALC$ = TRIM$(UCase$(Mid$(XX$, 13, 32)))
       MODOCALC$ = REPLACAR$(MODOCALC$, " ", "")
       If MODOCALC$ = "M1XM2" Or MODOCALC$ = "M1*M2" Then
            CANTU = CANCOMPO * (M1% / 1000) * (M2% / 1000)
          ElseIf MODOCALC$ = "M1+M2" Then
            CANTU = CANCOMPO * ((M1% / 1000) + (M2% / 1000))
          '\\\OT-08-679-FL-27/08/08///
          ElseIf XVALO(MODOCALC$) > 0 Then
            CANTU = CANCOMPO * XVALO(MODOCALC$)
          '\\\OT-08-679-FL-FIN///
          Else
            Call MENSERR(24, "Modo de Cálculo '" + MODOCALC$ + "' no Procesable.")
            Call FINAL("")
       End If
       '
       EXPLOLIN$(U&) = Space$(128)
       Mid$(EXPLOLIN$(U&), 1, 16) = CODEXT$(CODCOMPO%)
       Mid$(EXPLOLIN$(U&), 49, 16) = FORMATNUM$(CANTU, "F12.4")
    Next U&
    Call AREXPLO(CII%, 1, 2)
    '
End Sub
           
Function PRECALCULADO(CODFAM$, M1, M2)
    RECORD$ = FILTERGETRECORD$("TAFAMFIL", 1, CODFAM$) 'lee en la tabla tafamfil en la posicion 1 con el argumento codigo de familia
    PRECITAFAMFIL = CVD(Mid$(RECORD$, 33, 8)) 'devuelve el precio que viene de la tabla tafamfil
    PRECIFIJO = CVD(Mid$(RECORD$, 47, 8))
    MEDIDA = (M1 / 1000) * (M2 / 1000) 'calculamos la medida en M2 (dividimos * 1000 por que son mm.
    PRECALCULADO = PRECIFIJO + PRECITAFAMFIL * MEDIDA 'precio unitario surge del precio que trajimos de la tabla tafamfil por los M2
End Function

           

Private Sub verpendientesaprobacion_Click()
    Call APRONP04.Command4_Click
End Sub
