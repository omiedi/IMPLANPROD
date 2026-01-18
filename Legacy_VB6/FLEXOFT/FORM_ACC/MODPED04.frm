VERSION 5.00
Begin VB.Form MODPED04 
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
      TabIndex        =   29
      Top             =   5145
      Width           =   10800
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   1275
         Left            =   3990
         TabIndex        =   64
         Top             =   1470
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
            TabIndex        =   67
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
            TabIndex        =   66
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
            TabIndex        =   65
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
         Height          =   1440
         Left            =   5670
         TabIndex        =   43
         Top             =   0
         Width           =   5055
         Begin VB.Frame Frame8 
            BackColor       =   &H00C0FFC0&
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
            TabIndex        =   95
            Top             =   2205
            Width           =   1485
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
               TabIndex        =   101
               Top             =   190
               Width           =   1005
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
               TabIndex        =   100
               Top             =   405
               Width           =   1005
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
               TabIndex        =   99
               Top             =   405
               Width           =   400
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
               TabIndex        =   98
               Top             =   190
               Width           =   400
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
               TabIndex        =   97
               Top             =   620
               Width           =   405
            End
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
               TabIndex        =   96
               Top             =   620
               Width           =   1005
            End
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
            Left            =   3400
            TabIndex        =   74
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
            TabIndex        =   73
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
            TabIndex        =   72
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
            TabIndex        =   68
            Top             =   1000
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
            Index           =   2
            Left            =   3400
            TabIndex        =   62
            TabStop         =   0   'False
            Text            =   "    *******"
            Top             =   970
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
            Index           =   1
            Left            =   3400
            TabIndex        =   61
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
            Left            =   3400
            TabIndex        =   60
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
            TabIndex        =   46
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
            TabIndex        =   45
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
            Height          =   300
            Left            =   1890
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
            TabIndex        =   44
            TabStop         =   0   'False
            Text            =   " "
            Top             =   970
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
            TabIndex        =   77
            Top             =   2730
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
            TabIndex        =   76
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
            TabIndex        =   75
            Top             =   1785
            Width           =   1905
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   210
            Picture         =   "MODPED04.frx":0000
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
            TabIndex        =   71
            Top             =   1060
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
               Weight          =   900
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00005500&
            Height          =   330
            Left            =   315
            TabIndex        =   70
            ToolTipText     =   "Doble-Click para 'Dólares'"
            Top             =   650
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
            TabIndex        =   69
            Top             =   420
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
            TabIndex        =   49
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
            TabIndex        =   48
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
            Left            =   3045
            TabIndex        =   47
            Top             =   735
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
         TabIndex        =   35
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
            TabIndex        =   58
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
            TabIndex        =   38
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
            TabIndex        =   37
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
            TabIndex        =   36
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
            TabIndex        =   79
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
            TabIndex        =   78
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
            TabIndex        =   59
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
            TabIndex        =   42
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
            TabIndex        =   41
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
            TabIndex        =   40
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
            TabIndex        =   39
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
         Height          =   1600
         Left            =   5670
         TabIndex        =   34
         Top             =   1540
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
            Index           =   1
            Left            =   1150
            MaxLength       =   32
            TabIndex        =   113
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
            Index           =   3
            Left            =   1150
            MaxLength       =   32
            TabIndex        =   112
            Top             =   555
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
            Left            =   1150
            MaxLength       =   32
            TabIndex        =   107
            Top             =   240
            Width           =   3750
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
            Left            =   1150
            TabIndex        =   106
            Top             =   870
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
            Left            =   1800
            TabIndex        =   105
            Top             =   870
            Width           =   195
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
            Left            =   1150
            TabIndex        =   104
            Top             =   1200
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
            Index           =   2
            Left            =   2035
            MaxLength       =   48
            TabIndex        =   103
            Top             =   1200
            Width           =   2865
         End
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
            TabIndex        =   102
            Top             =   1200
            Width           =   195
         End
         Begin VB.Label Label24 
            Height          =   225
            Left            =   105
            TabIndex        =   115
            Top             =   780
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label Label23 
            Height          =   225
            Left            =   105
            TabIndex        =   114
            Top             =   525
            Visible         =   0   'False
            Width           =   960
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
            Left            =   -800
            TabIndex        =   111
            Top             =   360
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
            Index           =   5
            Left            =   -800
            TabIndex        =   110
            Top             =   1305
            Width           =   1905
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
            Left            =   -800
            TabIndex        =   109
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
            Left            =   2035
            TabIndex        =   108
            Top             =   870
            Width           =   2865
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8520
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   12000
      Begin VB.PictureBox Picture6 
         Height          =   9000
         Left            =   11025
         ScaleHeight     =   8940
         ScaleWidth      =   1215
         TabIndex        =   116
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
            Picture         =   "MODPED04.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   117
            ToolTipText     =   "Aprueba, Graba e Imprime Pedido"
            Top             =   7440
            Width           =   650
         End
         Begin VB.Image Image7 
            Height          =   8415
            Left            =   0
            Picture         =   "MODPED04.frx":0614
            Stretch         =   -1  'True
            Top             =   -120
            Width           =   1095
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00400040&
         Height          =   8415
         Left            =   11025
         ScaleHeight     =   8355
         ScaleWidth      =   945
         TabIndex        =   80
         Top             =   0
         Width           =   1000
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
            Picture         =   "MODPED04.frx":29AC
            Style           =   1  'Graphical
            TabIndex        =   92
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
            Picture         =   "MODPED04.frx":2CB6
            Style           =   1  'Graphical
            TabIndex        =   93
            ToolTipText     =   "Consultas y Re-Impresión"
            Top             =   3380
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
            Picture         =   "MODPED04.frx":2FC0
            Style           =   1  'Graphical
            TabIndex        =   91
            ToolTipText     =   "Cierra y Abandona la Aplicación"
            Top             =   105
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
            Picture         =   "MODPED04.frx":310A
            Style           =   1  'Graphical
            TabIndex        =   90
            ToolTipText     =   "Resumen de Cuenta Corriente"
            Top             =   2790
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
            Picture         =   "MODPED04.frx":3414
            Style           =   1  'Graphical
            TabIndex        =   89
            ToolTipText     =   "Configuración de Funcionamiento de Pedidos"
            Top             =   4800
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
            Picture         =   "MODPED04.frx":3856
            Style           =   1  'Graphical
            TabIndex        =   88
            ToolTipText     =   "Acceso a Condiciones Comerciales"
            Top             =   5380
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
            Picture         =   "MODPED04.frx":3B60
            Style           =   1  'Graphical
            TabIndex        =   87
            ToolTipText     =   "Acceso a Maestro de Artículos"
            Top             =   2000
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
            Picture         =   "MODPED04.frx":3E6A
            Style           =   1  'Graphical
            TabIndex        =   86
            ToolTipText     =   "Acceso a Base de Datos de Clientes"
            Top             =   1420
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
            Picture         =   "MODPED04.frx":42AC
            Style           =   1  'Graphical
            TabIndex        =   85
            ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
            Top             =   3960
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
            Picture         =   "MODPED04.frx":45B6
            Style           =   1  'Graphical
            TabIndex        =   84
            ToolTipText     =   "Ayuda FLEXOFT en Línea"
            Top             =   685
            Width           =   650
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
            Height          =   960
            Left            =   105
            Picture         =   "MODPED04.frx":48C0
            Style           =   1  'Graphical
            TabIndex        =   83
            ToolTipText     =   "Aprueba, Graba e Imprime Pedido"
            Top             =   7250
            Width           =   650
         End
         Begin VB.PictureBox MARCOBARRA 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   105
            ScaleHeight     =   75
            ScaleWidth      =   585
            TabIndex        =   81
            Top             =   6300
            Width           =   650
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   0
               TabIndex        =   82
               Top             =   0
               Width           =   12000
            End
         End
         Begin VB.Image Image3 
            Height          =   390
            Left            =   -310
            Picture         =   "MODPED04.frx":4BCA
            Top             =   6615
            Width           =   1515
         End
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
         TabIndex        =   56
         Top             =   1260
         Width           =   180
      End
      Begin VB.PictureBox MARCONOTA 
         BackColor       =   &H00FFFFC0&
         Height          =   3180
         Left            =   210
         ScaleHeight     =   3120
         ScaleWidth      =   10455
         TabIndex        =   50
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
            Picture         =   "MODPED04.frx":6AEC
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
            Picture         =   "MODPED04.frx":7156
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
            Picture         =   "MODPED04.frx":7460
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
            TabIndex        =   53
            Top             =   390
            Width           =   9780
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FFFFC0&
            Height          =   435
            Left            =   -30
            ScaleHeight     =   375
            ScaleWidth      =   9720
            TabIndex        =   51
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
               TabIndex        =   52
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
         TabIndex        =   54
         Top             =   2205
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
         TabIndex        =   33
         Top             =   6855
         Width           =   2825
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   210
         ScaleHeight     =   375
         ScaleWidth      =   10455
         TabIndex        =   31
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
            TabIndex        =   32
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
      Begin VB.Label LTEXT1 
         Caption         =   "Label19"
         Height          =   330
         Left            =   525
         TabIndex        =   94
         Top             =   0
         Visible         =   0   'False
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
            Weight          =   900
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   9030
         TabIndex        =   63
         ToolTipText     =   "Doble-Click para 'Reparación'"
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label9 
         Height          =   225
         Left            =   7455
         TabIndex        =   57
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
         TabIndex        =   55
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
         TabIndex        =   30
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
   Begin VB.Menu Archi 
      Caption         =   "Archivos"
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
End
Attribute VB_Name = "MODPED04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim ARCHIPRE$

Private Sub anotadorpedidos_Click()
   Call Command22_Click
End Sub

Private Sub Command1_Click()
   If List1.ListCount < 1 Then Exit Sub
   Call SELECHO("DEUDOR1")
   NCLIE% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
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
   ACONEC$ = "ARCHIG04/AMACLI"
   Call CONECRUN(ACONEC$, "Padron Maestro de Clientes")
   Command11.Enabled = True
'
'   Command11.Enabled = False
'   Call CIERRARCH("*.*")
'   ACONEC$ = "AMACLI01"
'   If EXISTE%("AMACLI02.EXE") > 0 Then ACONEC$ = "AMACLI02"
'   Call CONECRUN(ACONEC$, "Padron Maestro de Clientes")
'   Command11.Enabled = True
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
5   Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
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
   Call GRAPEDIDO(OKX%)
   '
   If OKX% > 0 Then
     '
     CONPRIPE.Show 1
     If CONPRIPE.Label1.Caption = "OK" Then
        CANCELPRINT% = 0
        If CONPRIPE.Check2.Value = 1 Then CANCELPRINT% = 1  ' IMPRESION FANTASMA
        '
        If CONPRIPE.Check7.Value = 1 Then
           Call PRIPEDIDO(1)     ' Imprime Orden de Venta
        End If
        If CONPRIPE.Check8.Value = 1 Then
           Call PRIPEDIDO(2)     ' Imprime Orden de Despacho
        End If
        '
        'ENVIO POR MAIL
        If CONPRIPE.Check1.Value = 1 Then
            NCXX% = XVALO(Text1)
            If EMAILCLI$(NCXX%) <> "" Then
                SENDBYMAIL$ = EMAILCLI$(NCXX%)
                Call PRIPEDIDO(1)     ' Imprime Orden de Venta
                SENDBYMAIL$ = ""
            End If
        End If
        '
        CONPRIPE.Label1 = ""
        CANCELPRINT% = 0
        
     End If
     '
     Call GENERODES      ' genera orden de despacho
     Call BLANFORMU
     Call SETULTREG("!CARDETPE", 0)
     Call CIERRARCH("*.*")
     Hide
   End If
99 Command14.Enabled = True
End Sub

Private Sub Command15_Click()
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
    '\\\OT-6-422-FG-05/10/06
    Call SELECHO("!SUCENTRE")
    VDEVU$ = VALACT1$("!SUCENTRE")
    Text20 = VDEVU$
    '\\\OT-6-422-FG-FIN
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
    CFGPED04.Show 1
    Command5.Enabled = True
    '
'    Command5.Enabled = False
'    Call CIERRARCH("*.*")
'    If (USNBR% Mod 100) <> 1 Then
'      Call COMUNI("La Opción Elegida Requiere\Privilegios de Supervisor.")
'      Exit Sub
'    End If
'    OPORDESPA.Show 1
'    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   MARCONOTA.Visible = False
   Call CARDETPEDI(1)
End Sub
'
Private Sub Command7_Click()
    Command7.Enabled = False
    Call CIERRARCH("*.*")
    MODX% = ALTERNA%("Listas de Precios\Condiciones de Pago")
    If MODX% = 1 Then
         Call CONECRUN("LISPRE04", "Listas de Precios de Venta")
         'Call CONECRUN("ADMIPRE", "Listas de Precios de Venta")
       ElseIf MODX% = 2 Then
         Call CARCONPAGOS
    End If
    Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    Call CIERRARCH("*.*")
    ACONEC$ = "ARCHIG04/AMASTO"
    Call CONECRUN(ACONEC$, "Maestro de Artículos de Stock")
    Command8.Enabled = True
    '
'    Command8.Enabled = False
'    Call CIERRARCH("*.*")
'    ACONEC$ = "AMASTO01"
'    If EXISTE%("AMASTO02.EXE") > 0 Then ACONEC$ = "AMASTO02"
'    Call CONECRUN(ACONEC$, "Maestro de Artículos de Stock")
'    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
End Sub

Private Sub CoPag_Click()
    Call CARCONPAGOS
End Sub

Private Sub Eco21_GotFocus(Index As Integer)
   On Error Resume Next
   TEXT21(Index).SetFocus
   On Error GoTo 0
End Sub

Private Sub Form_Load()
    '
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
      Me.MARCONOTA.BackColor = &HFCD7B6
      Me.Picture1.BackColor = &HFCD7B6
    End If

    Call CENTRAFORM(Me)
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Frame4.Height = 1440
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
    MONEMIS.Caption = DEMONECO$(MONEMI%)
    MONEMIS.Visible = True
    CONVERPRECIO% = 1
    TICAMBIS.Visible = True
    Label12.Visible = True
    Image1.Visible = True
    TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
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

Private Sub grabarpedido_Click()
   Call Command14_Click
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
    Call CONECRUN("LISPRE04", "Listas de Precios de Venta")
End Sub

Private Sub List1_DblClick()
   Call CARDETPEDI(2)
End Sub

Private Sub MaesArt_Click()
   Call Command8_Click
End Sub

Private Sub MaesCli_Click()
   Call Command11_Click
End Sub

Private Sub MONEMIS_Change()
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

Private Sub otr_Click()
    Call Command13_Click
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

Private Sub sali_Click()
    Call Command2_Click
End Sub

Private Sub Text1_Change()
   If Text1 = "" Then
        LTEXT1 = ""
      Else
        If Text1 <> LTEXT1 Then
          Call MENSERR(24, "No Válido Cambiar Cliente")
          Text1 = LTEXT1
        End If
   End If
   '
   NCLIE% = XVALO(Text1.TEXT)
   If NCLIE% > 0 Then
     If REGICLI&(NCLIE%) > 0 Then
       Call CARDACLI
     End If
   End If
End Sub

Private Sub Text1_DblClick()
   If List1.ListCount < 1 Then Exit Sub
   Call SELECHO("DEUDOR1")
   NCLIE% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
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
      '\\\OT-07-0040-RG-02/02/07///
      Label23.Caption = DOMITRANS$(NTRAI%)
      Label24.Caption = LOCATRANS$(NTRAI%)
      '\\\OT-07-0040-RG-FIN///
    End If
  End If
  '\\\OT-6-422-FG-FIN
End Sub

Private Sub Text2_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text20_Change()
    '\\\OT-6-422-FG-05/10/06
    NSUC& = XVALO(Text20)
      Label27 = ""
      If NSUC& > 0 Then
        If NSUC& <= ULTREG&("!SUCENTRE") Then
          XX$ = REGLEIDO$("!SUCENTRE", NSUC&)
          Label27 = TRIM$(Mid$(XX$, 3, 62))
          LENTREGA(0) = TRIM$(Mid$(XX$, 65, 48))
          LENTREGA(1) = TRIM$(Mid$(XX$, 113, 48))
          Text19 = CVI(Mid$(XX$, 195, 2))
        End If
      End If
      '\\\OT-6-422-FG-FIN
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
   '\\\OT-7-010-RG-12/01/07
   Text1.TEXT = ""
   '\\\OT-7-010-RG-FIN
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
   '\\\OT-6-422-FG-05/10/06
   Text19 = ""
   Text20 = ""
   Call SETULTREG("!SUCENTRE", 0)
   Call FRESHECHO("!SUCENTRE")
   '\\\OT-6-422-FG-FIN
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
   '\\\OT-05-0560-WAR-24/11/05///
   Label13 = ""
   Label14 = ""
   Label20 = ""
   '\\\OT-05-0560-WAR-FIN///
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
    NC% = XVALO(Text1.TEXT)
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
        '
        '\\\OT-07-0010-RG-12/01/07///
        '\\\OT-6-422-FG-04/10/06
        ' CARGA LISTA DE SUCURSALES DEL CLIENTE ACTIVO
        Call FILTERFILE("SUCENTRE", "", 5, MKI$(NC%))
        '\\\OT-6-422-FG-FIN
        '\\\OT-07-0010-RG-FIN///
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
        '
      End If
    End If
    Screen.MousePointer = 1
    '
End Sub
'
Private Sub Text6_GOTFOCUS()

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

Private Sub Timer1_Timer()
   '
   Call CALTOPED
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
    '
    If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
       Call MENSERR(24, "Imposible Imprimir Pedido de Cliente.\No Existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    '
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
            PRELIX = PRELISTA#(ARCHIPRE$, CIPRE%)
            PRECARGA = CVS(Mid$(XX$, 39, 4))
            PREUNI = CVS(Mid$(XX$, 73, 4))
            If PRELIX < PRECARGA Then PRELIX = PRECARGA
            If PRELIX < PREUNI Then PRELIX = PREUNI
            '
            PRELI$ = Str$(PRELIX)
            PREUX$ = Str$(CVS(Mid$(XX$, 73, 4)))
            DESCX$ = "0": On Error Resume Next
                          DESCX$ = Str$(100 * (XVALO(PRELI$) - XVALO(PREUX$)) / XVALO(PRELI$))
                          On Error GoTo 0
            ITOTX$ = Str$((XVALO(CAPEX$)) * (XVALO(PREUX$)))
            FEX = CVI(Mid$(XX$, 63, 2))
            FENTX$ = FECHATEX$(FEX)
            DESXLXTA$ = FORMATNUM$(CVS(Mid$(XX$, 109, 4)), "F4.1")
            DESXGPO$ = FORMATNUM$(CVS(Mid$(XX$, 113, 4)), "F4.1")
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
          End If
        End If
      End If
109 Next U&
    '
    SUMATOTA# = XVALO(Total(0).TEXT)
    IDESCUEN# = XVALO(Total(1).TEXT)
    '
    '\\\OT-05-0853-WAR-24/11/05///
    If PIVACLI%(NC%) = 5 Or PIVACLI%(NC%) = 7 Then
      Label13 = "" 'SI CLIENTE EXTERIOR O PROMOVIDO
      Label14 = "" 'NO DISCRIMINA IVA
      Label20 = ""
    End If
    '\\\OT-05-0853-WAR-FIN///
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
    CODPARAM$(10) = "IMP-IVA": DOCPARAM$(10) = Label13 ' TOTIVA(3).TEXT
    CODPARAM$(11) = "IMP-TOTA": DOCPARAM$(11) = Total(3).TEXT
    CODPARAM$(12) = "POR-DESC": DOCPARAM$(12) = TRIM$(PORDESCU.TEXT)
    '\\\OT-5-869-28/12/05///
    'CODPARAM$(13) = "MONEDEMI": DOCPARAM$(13) = MONEMIS
    '\\\OT-5-869-FIN///
    CODPARAM$(13) = "PERC-IVA": DOCPARAM$(13) = Text15
    CODPARAM$(14) = "IMP-IVA1": DOCPARAM$(14) = Label14
    '
    '\\\OT-6-422-FG-05/10/06
    CODPARAM$(15) = "NRO-PASE": DOCPARAM$(15) = Label27
    '\\\OT-6-422-FG-FIN
    '
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
    'FORIPRE$ = TRIM$(CONTROL$("", "FORPECLI"))
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
    Call BLANARCH("!OBSERVOF")
    ' NUEVO PARA IMPRESION DE OBSERVACIONES
    CODFOR$ = FORIPRE$
    FORMATOBS$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "OBSERVA/FORMATO")))
    ANCHOMA = XVALO(Mid$(FORMATOBS$, 2))
    If ANCHOMA > RECLEN("OBSERVOF") Then ANCHOMA = RECLEN("OBSERVOF")
    ANCHOMAXU% = ANCHOMA
    Call FRATEXTO(Text13.TEXT, ANCHOMAXU%)
    For KU& = 1 To CARETEX%
       Call GRAREG("!OBSERVOF", RETEX$(KU&), KU&)
    Next KU&
    Call SETULTREG("!OBSERVOF", KU& - 1)
    '
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
    '\\\OT-05-0854-WAR-24/11/05///
'    'IMPRESIÓN DEL FORMULARIO DE ORDEN DE DESPACHO
      
'    FORIPRE$ = TRIM$(CONTROL$("", "FORDESPA"))
'    If FORIPRE$ <> "" Then
'      If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") > 0 Then
'        TIPODOC$ = "Orden de Despacho"
'        CANCELPRINT% = 0
'        If CONTROL$("ORDESPA", "SUPRINT") = "SI" Then
'           CANCELPRINT% = 1
'        End If
'        '
'        If InStr(EPAC$, "DOSIVAC") > 0 Then
'          CANCELPRINT% = 0
'          If OPRIPEDI.Check1(1).Value <> 1 Then
'            CANCELPRINT% = 1
'          End If
'        End If
'        '
'        If CARETEX% >= 1 Then Call GRAREG("!OBSERVOF", RETEX$(1), 1)
'        If CARETEX% >= 2 Then Call GRAREG("!OBSERVOF", RETEX$(2), 2)
'        If CARETEX% >= 3 Then Call GRAREG("!OBSERVOF", RETEX$(3), 3)
'        URETE& = CARETEX%
'        Call SETULTREG("!OBSERVOF", URETE&)
'        Call CIERRARCH("!OBSERVOF")
'        '
'        Call PRINTDOC("FORDESPA")   ' ORDEN DE DESPACHO
'      End If
'    End If
'    '
'\\\OT-05-0854-WAR-FIN///
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
   If XVALO(Total(2).TEXT) < 0.01 Then
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
   'REGENCA& = XVALO(Label9.Caption)
   'If REGENCA& < 1 Or REGENCA& > ULTREG&("PEDENCA") Then
   '  Call MENSERR(24, "Imposible Modificar este Pedido.")
   '  Call FINAL("")
   'End If
   '
   'PDC$ = REGLEIDO$("PEDENCA", REGENCA&)
   'If CVS(Left$(PDC$, 4)) <> XVALO(Text11.TEXT) Then
   '  Call MENSERR(24, "Imposible Modificar este Pedido.")
   '  Call FINAL("")
   'End If
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
   'If XVALO(Total(2).TEXT) < 0.01 Then  VERIFICADO MAS ARRIBA
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
   '\\\OT-06-0512-RG-27/11/06///
   If CONTROL$("ADMIPED", "OBLICOM") = "SI" Then
   '\\\OT-06-0512-RG-FIN///
    If TRIM$(Text10) = "" Then
       MERRO$ = "Falta Orden de Compra Cliente"
       GoTo 98
    End If
   End If
   '
   REFE$ = " - " + TRIM$(Text10.TEXT)
   If REFE$ = " - " Then REFE$ = ""
   FEX = FEMI%
   REFE$ = REFE$ + " - " + FECHATEX$(FEX)
   LOMA% = 48 - Len(REFE$)
   REFE$ = Left$(TRIM$(ECOARCH$("DEUDOR1", MKI$(NCLIE%))), LOMA%) + REFE$
   '
   NROPE& = XVALO(Text11.TEXT)
   Call ABREPEDCLI
   With PedEnca
     .FindFirst "NroPed = " & NROPE&
     If .NoMatch = True Then
       MERRO$ = "Imposible Modificar Este Pedido"
       GoTo 98
     End If
     '
     STATUPE% = !Status
     .Edit
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
     !ListPre = XVALO(TEXT21(0))
     !CondPag = XVALO(TEXT21(1))
     !VENDED = XVALO(TEXT21(2))
     !TDescApli = TRIM$(PORDESCU.TEXT)
     .Update
   End With
   '
   SQLX$ = "SELECT * FROM PEDDETA "
   SQLX$ = SQLX$ + "WHERE NroPed = " & NROPE& & " "
   Set PEDDETTEMP = Ventas.OpenRecordset(SQLX$, dbOpenDynaset)
   With PEDDETTEMP
     .MoveFirst
     Do While Not .EOF
        STATUPRO% = 100 + !Status
        .Edit
        !Status = STATUPRO
        .Update
     .MoveNext
     Loop
   End With
   

'   With PedDeta
'25   .FindFirst "NroPed = " & NROPE& & " And Status < 8 " '\\\OT-06-0520-RG-01/12/06///
'     If .NoMatch = False Then
'       .Delete
'       GoTo 25
'     End If
'   End With
   '
   ANOTA$ = ""
   TTXYZ$ = Space$(76)
   HOII% = HOY(HOS$)
   Call REPLA(TTXYZ$, HOS$, 1, 8)
   Call REPLA(TTXYZ$, "Modificación de Orden de Venta:", 11, 56)
   Call REPLA(TTXYZ$, USERTER$, 69, 24)
   ANOTA$ = ANOTA$ + TTXYZ$ + Chr$(10) + Chr$(13)
   '
   CAVER$ = TRIM$(Str$(NROPE&))
   While Len(CAVER$) > 6: CAVER$ = Mid$(CAVER$, 2): Wend
   While Len(CAVER$) < 6: CAVER$ = "0" + CAVER$: Wend
   '
   For U& = 1 To ULTREG&("!CARDETPE")
     X$ = REGLEIDO$("!CARDETPE", U&)
     CAN = CVS(Mid$(X$, 47, 4))
     If CAN > 0.05 Then
       '
26     IKITEM% = CVI(Mid$(X$, 123, 2))
       If IKITEM% < 1 Then
27        CAVERI$ = CAVER$ + "-" + TRIM$(Str$(IKITEM%))
          PEDDETTEMP.FindFirst "TPedItem = '" + CAVERI$ + "'"
          If PEDDETTEMP.NoMatch = False Then
             IKITEM% = IKITEM% + 1
             GoTo 27
          End If
       End If
       CAVERI$ = CAVER$ + "-" + TRIM$(Str$(IKITEM%))
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
     
       With PEDDETTEMP
           .FindFirst "CodiInt = " & CIXI% & " AND TPedItem = '" + CAVERI$ + "'"
           If .NoMatch = False Then
                .Edit
              Else
                .AddNew
           End If
           !NroPed = NROPE&
           !FECHEMIS = CVDAT(FEMI%)
           !NroClie = NCLIE%
           !RasoClie = RASOCLI$(NCLIE%)
           !NroOcom = TRIM$(Text10.TEXT)
             FENTRESOL% = CVI(Mid$(X$, 63, 2))
           !FeSolEnt = CVDAT(FENTRESOL%)
           If !Status > 90 Then
                STATUI% = !Status - 100
                !Status = STATUI%
              Else
                !Status = STATUPE%   'idem anterior.
           End If
           !ListPre = XVALO(TEXT21(0).TEXT)
           !CondPag = XVALO(TEXT21(1).TEXT)
           !VENDED = XVALO(TEXT21(2).TEXT)
           !PorComVend = XVALO(Text44.TEXT)
           !NuorItem = IKITEM%
             CIXI% = CVI(Left$(X$, 2))
           !CodiInt = CIXI%
           !CodiExt = CODEXT$(CIXI%)
           !SubCod = SUBCODX$
           !Descrip = DESCRIT0$(CIXI%)
           !TColor = Mid$(X$, 3, 16)
           !TTalle = Mid$(X$, 19, 8)
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
             CAENTAN = CVS(Mid$(X$, 117, 4))   ' CANTIDAD ENTREGAS ANTERIORES
             FULENTAN% = CVI(Mid$(X$, 121, 2)) ' FECHA ULTIMA ENTREGA
           !CantEnt = CAENTAN
           !CanSaldo = CAN - CAENTAN
           !FeUltEnt = CVDAT(FULENTAN%)
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
     End If
   Next U&
   '
   With PEDDETTEMP
35   .MoveFirst
     Do While Not .EOF
        STATUPRO% = !Status
        If STATUPRO% >= 99 Then
           If STATUPRO% < 103 Then
             .Delete
             GoTo 35
           End If
        End If
        If STATUPRO% = 108 Or STATUPRO% = 109 Then
           .Edit
           !Status = STATUPRO% - 100
           .Update
        End If
     .MoveNext
     Loop
   End With
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
   ANOTB$ = PEDDDENCA!AnotaPed
   ANOTA$ = ANOTB$ + ANOTA$
   '\\\OT-08-0912-OD-30/09/08/// PARA EVITAR QUE SE GUARDEN DATOS CON COMILLAS SIMPLES
   ANOTA$ = REPLACAR$(ANOTA$, "'", "")
   '\\\OT-08-0912-OD-FIN///
   '
   SQLX$ = "UPDATE PEDENCA SET AnotaPed=" & "'" & ANOTA$ & "' WHERE NroPed=" & NROPE&
   '
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
   ' RUTINA REVISA Y MODIFICA EN EL CASO QUE DADO EL NUMERO DE PEDIDO
   ' SI HAY REPETICIONES DE NUMERO DE SERIE O Nº ORDEN EN 0
   '
   Call REVISANUMERO(NROPE&)
   '
   Call CIERRARCH("*.*")
   Call LIBARCH("*.*")
   Call BAJALDISCO
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
     TTT1 = XVALO(Total(0).TEXT)
     Total(0) = CSTRING$(MKS$(TTT1), 3, 11, 2, 2)
   End If
End Sub

Sub CARDETPEDI(MODO%)
   '
   CLIEN% = XVALO(Text1.TEXT)
   NULISTA% = XVALO(TEXT21(Index).TEXT)
   NUPEDXX& = XVALO(Text11)
   '
   '\\\OT-08-0133-FL-3-4-08///
   If CONTROL$("PEDIDO", "VARCOLOR") = "SI" Then  ' si la opcion de pedidos por variante esta en true
      Call CIERRARCH("!CARDETP1")
      LFCDP$ = LOADFILE$(LUCOM$ + "CARDETPE.WKF")
      Call SAVEFILE(LUCOM$ + "CARDETP1.WKF", LFCDP$)
3     ATRI$ = "/NC/ARCHECO(1)=" + ARCHIPRE$
      ATRI$ = ATRI$ + "/TITUPAN=Detalle Pedido de Cliente"
      ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Frame1.Top + Picture3.Top + 400))
      ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture3.Left))
      DETPED% = VENTABU%("CARDETP1" + ATRI$)
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
      'AQUI VALIDAR CAMBIOS DE CODIGO
      HUBOCA% = 0
      SQLX$ = "SELECT * FROM PEDDETA "
      SQLX$ = SQLX$ + "WHERE NroPed = " & NUPEDXX& & " "
      SQLX$ = SQLX$ + "AND Status < 8 "
      Set PEDDETTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
      For KKU& = 1 To FIN&
         X$ = REGLEIDO$("!CARDETP1", KKU&)
         CI% = CVI(Left$(X$, 2))
         NORIT% = CVI(Mid$(X$, 123, 2))
         '
         PEDDETTEMP.FindFirst "CodiInt = " & CI% '& " AND NuorItem = " & NORIT%
         If PEDDETTEMP.NoMatch = True Then    ' NO LO ENCONTRO - CAMBIÓ EL CÓDIGO
            PRECARGA = PRELISTA#(ARCHIPRE$, CI%)
            Call REPLA(X$, MKS$(PRECARGA), 39, 4)
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
            Call GRAREG("!CARDETP1", X$, KKU&)
            HUBOCA% = 1
         End If
      Next KKU&
      '
      If HUBOCA% > 0 Then
         Call COMUNI("Precios y Descuentos Ajustados por Cambio de Códigos.")
      End If
      '
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
   '
   If XXX% = -1 Then    ' no se aprobo la segunda ventana de carga
     If CONTROL$("PEDIDO", "VARCOLOR") = "SI" Then
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
     If FENTRESOL% < 1 Then
        FENTRESOL% = UFENTRE%
        If FENTRESOL% < 1 Then FENTRESOL% = HOII%
        Call REPLA(Y$, MKI$(FENTRESOL%), 63, 2)
        Call GRAREG("!CARDETPE", Y$, U&)
     End If
     UFENTRE% = FENTRESOL%
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
   '\\\OT-05-0560-WAR-24/11/05///
   Dim IMPTASIVA#()
   ReDim IMPTASIVA#(ULTREG&("GRUPIVA"))
   DESC0$ = TRIM$(PORDESCU.TEXT)
   COEF1# = 1 - COEFDESCU(DESC0$) / 100
   '\\\OT-05-0560-WAR-FIN///
   TOTITEMS% = 0
   For U& = 1 To ULTREG&("!CARDETPE")
     X$ = REGLEIDO$("!CARDETPE", U&)
     CIXXI% = CVI(Left$(X$, 2))
     CAN = CVS(Mid$(X$, 47, 4))
     PORDESCUIN = CVS(Mid$(X$, 69, 4))
     If PORDESCUIN < 0 Or PORDESCUIN > 99 Then PORDESCUIN = 0
     '
     COEFITO# = (100 - PORDESCUIN) / 100
     COEFILI# = (100 - CVS(Mid$(X$, 109, 4))) / 100
     PDESCGRU = PORDESCUIN
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
       PREUNI = PRELI * (100 - PORDESCUIN) / 100
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


Private Sub TOTIVA_GotFocus(Index As Integer)
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub
