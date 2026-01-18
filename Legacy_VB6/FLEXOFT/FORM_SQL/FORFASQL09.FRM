VERSION 5.00
Begin VB.Form FORFASQL09 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturación en Cuenta Corriente"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11805
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command17 
      BackColor       =   &H00FFFF80&
      Height          =   545
      Left            =   6720
      Picture         =   "FORFASQL09.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   130
      ToolTipText     =   "Pedidos pendientes de Despacho, Remito o Factura"
      Top             =   1045
      Width           =   550
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3165
      Left            =   0
      TabIndex        =   26
      Top             =   5145
      Width           =   10800
      Begin VB.Frame Frame5 
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
         Height          =   990
         Left            =   210
         TabIndex        =   31
         Top             =   0
         Width           =   5325
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
            TabIndex        =   34
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
            Height          =   270
            Index           =   1
            Left            =   1890
            TabIndex        =   33
            Text            =   " "
            Top             =   600
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
            Height          =   285
            Index           =   1
            Left            =   1155
            TabIndex        =   15
            Text            =   " "
            Top             =   600
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
            Height          =   270
            Index           =   1
            Left            =   1575
            TabIndex        =   16
            Top             =   600
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
            Height          =   270
            Index           =   0
            Left            =   1890
            TabIndex        =   32
            Text            =   " "
            Top             =   315
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
            Height          =   270
            Index           =   0
            Left            =   1155
            TabIndex        =   13
            Text            =   " "
            Top             =   315
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
            Height          =   270
            Index           =   0
            Left            =   1575
            TabIndex        =   14
            Top             =   315
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
            TabIndex        =   37
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
            TabIndex        =   36
            Top             =   675
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
            TabIndex        =   35
            Top             =   390
            Width           =   1590
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Com."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   960
         Left            =   4830
         TabIndex        =   149
         Top             =   35
         Width           =   705
         Begin VB.TextBox Text22 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   55
            MaxLength       =   5
            TabIndex        =   155
            Top             =   350
            Width           =   570
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
            Index           =   4
            Left            =   1575
            TabIndex        =   152
            Top             =   1130
            Width           =   180
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
            Index           =   4
            Left            =   1155
            TabIndex        =   151
            Text            =   " "
            Top             =   1130
            Width           =   435
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
            Index           =   4
            Left            =   1890
            TabIndex        =   150
            Text            =   " "
            Top             =   1130
            Width           =   3270
         End
         Begin VB.Label Label28 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "A"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   750
            Left            =   80
            TabIndex        =   154
            Top             =   280
            Width           =   540
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
            Index           =   4
            Left            =   -3045
            TabIndex        =   153
            Top             =   1205
            Width           =   4110
         End
      End
      Begin VB.Frame Frame4 
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
         TabIndex        =   94
         Top             =   0
         Width           =   5055
         Begin VB.TextBox Text18 
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
            Left            =   2110
            TabIndex        =   134
            TabStop         =   0   'False
            Top             =   2760
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Frame Frame9 
            BorderStyle     =   0  'None
            Height          =   1120
            Left            =   105
            TabIndex        =   109
            Top             =   1100
            Width           =   4900
            Begin VB.CheckBox Check1 
               Caption         =   "Modo Test"
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
               Left            =   3240
               TabIndex        =   148
               Top             =   0
               Visible         =   0   'False
               Width           =   1590
            End
            Begin VB.TextBox BASIVA 
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
               Left            =   1680
               TabIndex        =   147
               Top             =   840
               Width           =   1395
            End
            Begin VB.TextBox BASIVA 
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
               Left            =   1680
               TabIndex        =   146
               Top             =   560
               Width           =   1395
            End
            Begin VB.TextBox BASIVA 
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
               Left            =   1680
               TabIndex        =   145
               Top             =   280
               Width           =   1395
            End
            Begin VB.TextBox TLABEL28 
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
               Left            =   1680
               TabIndex        =   144
               Top             =   0
               Width           =   1395
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
               Height          =   250
               Index           =   3
               Left            =   3150
               TabIndex        =   112
               TabStop         =   0   'False
               Text            =   " "
               Top             =   280
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
               Height          =   250
               Index           =   4
               Left            =   3150
               TabIndex        =   111
               TabStop         =   0   'False
               Text            =   " "
               Top             =   560
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
               Height          =   250
               Index           =   5
               Left            =   3150
               TabIndex        =   110
               TabStop         =   0   'False
               Text            =   " "
               Top             =   840
               Width           =   1650
            End
            Begin VB.Label Label29 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "IVA No Gravado:"
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
               TabIndex        =   143
               Top             =   30
               Width           =   1590
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
               Left            =   630
               TabIndex        =   121
               Top             =   870
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
               Left            =   630
               TabIndex        =   120
               Top             =   590
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
               Left            =   630
               TabIndex        =   119
               Top             =   280
               Width           =   645
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
               Left            =   1365
               TabIndex        =   118
               Top             =   870
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
               Left            =   1365
               TabIndex        =   117
               Top             =   580
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
               Left            =   1365
               TabIndex        =   116
               Top             =   280
               Width           =   225
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
               Left            =   -315
               TabIndex        =   115
               Top             =   280
               Width           =   855
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
               Left            =   -525
               TabIndex        =   114
               Top             =   580
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
               Left            =   -410
               TabIndex        =   113
               Top             =   870
               Width           =   960
            End
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
            TabIndex        =   95
            TabStop         =   0   'False
            Text            =   " "
            Top             =   2760
            Width           =   1650
         End
         Begin VB.Frame Frame8 
            BorderStyle     =   0  'None
            Height          =   1170
            Left            =   105
            TabIndex        =   98
            Top             =   190
            Width           =   4845
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
               Height          =   250
               Index           =   0
               Left            =   3150
               TabIndex        =   102
               TabStop         =   0   'False
               Top             =   60
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
               Height          =   250
               Index           =   1
               Left            =   3150
               TabIndex        =   101
               TabStop         =   0   'False
               Text            =   " "
               Top             =   340
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
               Height          =   250
               Index           =   2
               Left            =   3150
               TabIndex        =   100
               TabStop         =   0   'False
               Text            =   " "
               Top             =   620
               Width           =   1650
            End
            Begin VB.TextBox PORDESCU 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   250
               Left            =   1680
               TabIndex        =   99
               TabStop         =   0   'False
               Top             =   340
               Width           =   1065
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
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
               Height          =   250
               Left            =   420
               TabIndex        =   108
               Top             =   340
               Width           =   435
            End
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
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
               Height          =   250
               Left            =   0
               TabIndex        =   138
               Top             =   340
               Visible         =   0   'False
               Width           =   645
            End
            Begin VB.Label IEMBALA 
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
               Height          =   250
               Left            =   210
               TabIndex        =   135
               Top             =   620
               Visible         =   0   'False
               Width           =   1170
            End
            Begin VB.Label Label15 
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
               TabIndex        =   107
               Top             =   110
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
               TabIndex        =   106
               Top             =   660
               Width           =   2955
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
               Left            =   2835
               TabIndex        =   105
               Top             =   435
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
               Left            =   0
               TabIndex        =   104
               Top             =   110
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
               Left            =   -315
               TabIndex        =   103
               Top             =   435
               Width           =   1905
            End
         End
         Begin VB.Frame Frame10 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   110
            TabIndex        =   122
            Top             =   2200
            Width           =   4860
            Begin VB.TextBox Text17 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
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
               Left            =   1630
               TabIndex        =   125
               Top             =   210
               Width           =   1440
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
               Left            =   3360
               TabIndex        =   123
               Top             =   210
               Width           =   1440
            End
            Begin VB.Label TICAMBIS 
               Alignment       =   1  'Right Justify
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
               Height          =   270
               Left            =   160
               TabIndex        =   133
               Top             =   210
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label Label24 
               BackStyle       =   0  'Transparent
               Caption         =   "T-C:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   0
               TabIndex        =   132
               ToolTipText     =   "Doble Click para Configurar Cliente"
               Top             =   0
               Visible         =   0   'False
               Width           =   1065
            End
            Begin VB.Label Label23 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Perc. I.V.A.:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   1260
               TabIndex        =   126
               Top             =   0
               Width           =   1065
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Perc. I-B: "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   2835
               TabIndex        =   124
               ToolTipText     =   "Doble Click para Configurar Cliente"
               Top             =   0
               Width           =   960
            End
         End
         Begin VB.Label Label20 
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
            Left            =   105
            TabIndex        =   96
            Top             =   945
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
            Left            =   120
            TabIndex        =   131
            ToolTipText     =   "T.C.: 1.0000 - Doble-Click para Cambiar Moneda"
            Top             =   2730
            Width           =   1380
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
            TabIndex        =   97
            Top             =   2820
            Width           =   1905
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Contable"
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
         Height          =   660
         Left            =   210
         TabIndex        =   80
         Top             =   2450
         Width           =   5330
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
            Height          =   300
            Left            =   2310
            TabIndex        =   87
            Top             =   240
            Width           =   180
         End
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
            Height          =   285
            Left            =   1155
            TabIndex        =   86
            Top             =   240
            Width           =   1170
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
            Index           =   3
            Left            =   1575
            TabIndex        =   83
            Top             =   1130
            Width           =   180
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
            Index           =   3
            Left            =   1155
            TabIndex        =   82
            Text            =   " "
            Top             =   1130
            Width           =   435
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
            Index           =   3
            Left            =   1890
            TabIndex        =   81
            Text            =   " "
            Top             =   1130
            Width           =   3270
         End
         Begin VB.Label Label16 
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
            Left            =   2600
            TabIndex        =   88
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta:"
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
            Left            =   -525
            TabIndex        =   85
            Top             =   320
            Width           =   1590
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
            Index           =   3
            Left            =   -3045
            TabIndex        =   84
            Top             =   1205
            Width           =   4110
         End
      End
      Begin VB.Frame Frame7 
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
         ForeColor       =   &H00004040&
         Height          =   1485
         Left            =   210
         TabIndex        =   56
         Top             =   970
         Width           =   5330
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
            Height          =   270
            Left            =   1155
            TabIndex        =   140
            Top             =   775
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
            Height          =   270
            Left            =   1785
            TabIndex        =   139
            Top             =   775
            Width           =   180
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
            Height          =   270
            Left            =   1785
            TabIndex        =   137
            Top             =   1100
            Width           =   180
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
            Height          =   270
            Left            =   1155
            TabIndex        =   136
            Top             =   1090
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
            Index           =   3
            Left            =   1575
            MaxLength       =   32
            TabIndex        =   60
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
            Height          =   270
            Index           =   0
            Left            =   1155
            MaxLength       =   48
            TabIndex        =   59
            Top             =   210
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
            Height          =   270
            Index           =   1
            Left            =   1155
            MaxLength       =   48
            TabIndex        =   58
            Top             =   490
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
            Height          =   270
            Index           =   2
            Left            =   2050
            MaxLength       =   48
            TabIndex        =   57
            Top             =   1090
            Width           =   3095
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
            Height          =   270
            Left            =   2055
            TabIndex        =   142
            Top             =   780
            Width           =   3095
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
            TabIndex        =   141
            Top             =   860
            Width           =   1905
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
            TabIndex        =   62
            Top             =   315
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
            TabIndex        =   61
            Top             =   1155
            Width           =   1905
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8520
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   12000
      Begin VB.PictureBox MARCONOTA 
         BackColor       =   &H00FFFFC0&
         Height          =   3180
         Left            =   210
         ScaleHeight     =   3120
         ScaleWidth      =   10455
         TabIndex        =   38
         Top             =   1785
         Visible         =   0   'False
         Width           =   10520
         Begin VB.CommandButton Command26 
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
            Style           =   1  'Graphical
            TabIndex        =   156
            ToolTipText     =   "Facturacion de Bobinas"
            Top             =   1800
            Width           =   570
         End
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
            Picture         =   "FORFASQL09.frx":0442
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
            Picture         =   "FORFASQL09.frx":0AAC
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
            Left            =   9840
            Picture         =   "FORFASQL09.frx":0DB6
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
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   41
            Top             =   390
            Width           =   9780
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FFFFC0&
            Height          =   435
            Left            =   -30
            ScaleHeight     =   375
            ScaleWidth      =   9720
            TabIndex        =   39
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
               TabIndex        =   40
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
      Begin VB.Frame Frame11 
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
         Height          =   435
         Left            =   7455
         TabIndex        =   127
         Top             =   0
         Width           =   1275
         Begin VB.CommandButton Command22 
            Caption         =   "Cred"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   630
            Style           =   1  'Graphical
            TabIndex        =   129
            Top             =   240
            Width           =   650
         End
         Begin VB.CommandButton Command23 
            Caption         =   "N/Deb"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   128
            Top             =   240
            Width           =   650
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00C0C0C0&
         Height          =   3195
         Left            =   210
         ScaleHeight     =   3135
         ScaleWidth      =   10455
         TabIndex        =   91
         Top             =   1785
         Visible         =   0   'False
         Width           =   10520
         Begin VB.ListBox List3 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2760
            Left            =   -30
            TabIndex        =   92
            Top             =   420
            Visible         =   0   'False
            Width           =   10520
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Label17"
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
            Left            =   40
            TabIndex        =   93
            Top             =   105
            Width           =   10590
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
         Picture         =   "FORFASQL09.frx":10C0
         Style           =   1  'Graphical
         TabIndex        =   77
         ToolTipText     =   "Aprueba, Graba e Imprime Documento"
         Top             =   7320
         Width           =   840
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00400000&
         Height          =   7215
         Left            =   10920
         ScaleHeight     =   7155
         ScaleWidth      =   1155
         TabIndex        =   67
         Top             =   0
         Width           =   1215
         Begin VB.PictureBox MARCOBARRA 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   120
            ScaleHeight     =   75
            ScaleWidth      =   540
            TabIndex        =   78
            TabStop         =   0   'False
            Top             =   6510
            Width           =   600
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   0
               TabIndex        =   79
               Top             =   0
               Width           =   12000
            End
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
            Height          =   690
            Left            =   105
            Picture         =   "FORFASQL09.frx":13CA
            Style           =   1  'Graphical
            TabIndex        =   76
            ToolTipText     =   "Ayuda FLEXOFT en Línea"
            Top             =   790
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
            Height          =   640
            Left            =   105
            Picture         =   "FORFASQL09.frx":16D4
            Style           =   1  'Graphical
            TabIndex        =   75
            ToolTipText     =   "Acceso a Bases de Datos Maestras"
            Top             =   1580
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
            Height          =   640
            Left            =   105
            Picture         =   "FORFASQL09.frx":19DE
            Style           =   1  'Graphical
            TabIndex        =   74
            ToolTipText     =   "Remitos Pendientes de Facturación"
            Top             =   2220
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
            Height          =   640
            Left            =   105
            Picture         =   "FORFASQL09.frx":1CE8
            Style           =   1  'Graphical
            TabIndex        =   73
            ToolTipText     =   "Acceso a Listas de Precio, Condiciones de Pago y Tasas de IVA"
            Top             =   4265
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
            Height          =   640
            Left            =   105
            Picture         =   "FORFASQL09.frx":1FF2
            Style           =   1  'Graphical
            TabIndex        =   72
            ToolTipText     =   "Configuración de Funcionamiento"
            Top             =   5670
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
            Height          =   640
            Left            =   105
            Picture         =   "FORFASQL09.frx":2434
            Style           =   1  'Graphical
            TabIndex        =   71
            ToolTipText     =   "Resumen de Cuenta Corriente"
            Top             =   2860
            Width           =   650
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
            Height          =   690
            Left            =   120
            Picture         =   "FORFASQL09.frx":273E
            Style           =   1  'Graphical
            TabIndex        =   70
            ToolTipText     =   "Cierra y Abandona la Aplicación"
            Top             =   100
            Width           =   650
         End
         Begin VB.CommandButton Command20 
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
            Picture         =   "FORFASQL09.frx":2888
            Style           =   1  'Graphical
            TabIndex        =   69
            ToolTipText     =   "Opciones Varias - Anulación / Control Diferencia de Cambio"
            Top             =   4935
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
            Height          =   640
            Left            =   105
            Picture         =   "FORFASQL09.frx":29D2
            Style           =   1  'Graphical
            TabIndex        =   68
            ToolTipText     =   "Acceso al Control General de Gestión Comercial"
            Top             =   3510
            Width           =   650
         End
         Begin VB.Image Image2 
            Height          =   390
            Left            =   -315
            Picture         =   "FORFASQL09.frx":2CDC
            Top             =   6760
            Width           =   1515
         End
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
         TabIndex        =   64
         ToolTipText     =   "Doble-Click para Facturar por Ordenes de Compra de Clientes"
         Top             =   1280
         Width           =   2115
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   4830
         TabIndex        =   63
         Text            =   "Text13"
         Top             =   5500
         Width           =   960
      End
      Begin VB.PictureBox MARSELPENDE 
         BackColor       =   &H0080C0FF&
         Height          =   3180
         Left            =   210
         ScaleHeight     =   3120
         ScaleWidth      =   10455
         TabIndex        =   43
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
            TabIndex        =   48
            Top             =   420
            Width           =   9650
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H0080C0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   3165
            Left            =   0
            TabIndex        =   53
            Top             =   105
            Width           =   960
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H0080C0FF&
            Height          =   470
            Left            =   -105
            ScaleHeight     =   405
            ScaleWidth      =   9795
            TabIndex        =   46
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
               TabIndex        =   47
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
            Picture         =   "FORFASQL09.frx":4BFE
            Style           =   1  'Graphical
            TabIndex        =   45
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
            Picture         =   "FORFASQL09.frx":4F08
            Style           =   1  'Graphical
            TabIndex        =   44
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
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   0
         Top             =   0
      End
      Begin VB.CommandButton Command10 
         Height          =   260
         Left            =   420
         TabIndex        =   30
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
         ToolTipText     =   "Este Número Puede Sufrir Modificaciones segun el Servidor de la A.F.I.P."
         Top             =   480
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
         ToolTipText     =   "Doble Click para Ajustar Limite de Credito con Privilegios"
         Top             =   480
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
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1050
         TabIndex        =   1
         Top             =   890
         Width           =   180
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
         Left            =   240
         MaxLength       =   5
         TabIndex        =   0
         Text            =   " "
         Top             =   890
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
         TabIndex        =   42
         Top             =   2205
         Width           =   10515
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   210
         ScaleHeight     =   375
         ScaleWidth      =   10455
         TabIndex        =   28
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
            TabIndex        =   29
            Top             =   105
            Width           =   10515
         End
      End
      Begin VB.Label Label30 
         Height          =   255
         Left            =   960
         TabIndex        =   157
         Top             =   0
         Visible         =   0   'False
         Width           =   495
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
         ForeColor       =   &H00004040&
         Height          =   330
         Left            =   0
         TabIndex        =   90
         Top             =   660
         Width           =   960
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
         TabIndex        =   89
         Top             =   1280
         Width           =   1035
      End
      Begin VB.Label DOCUDIF 
         BackColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   5040
         TabIndex        =   66
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
         TabIndex        =   65
         ToolTipText     =   "Doble-Click para Facturar por Ordenes de Compra de Clientes"
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
         TabIndex        =   55
         Top             =   30
         Width           =   540
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
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
         TabIndex        =   54
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
         TabIndex        =   52
         Top             =   480
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
         TabIndex        =   51
         Top             =   460
         Width           =   150
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0001"
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
         Left            =   9000
         TabIndex        =   50
         ToolTipText     =   "Doble-Click Para Cambiar Punto de Venta"
         Top             =   480
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
         Left            =   315
         TabIndex        =   49
         Top             =   -45
         Width           =   1380
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00C0C000&
         FillColor       =   &H00FFFF80&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   420
         Shape           =   4  'Rounded Rectangle
         Top             =   255
         Width           =   540
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H0000C000&
         FillColor       =   &H0080FF80&
         FillStyle       =   0  'Solid
         Height          =   390
         Left            =   315
         Shape           =   4  'Rounded Rectangle
         Top             =   180
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
         TabIndex        =   27
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
         Top             =   195
         Width           =   1380
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000C0C0&
         FillColor       =   &H0080FFFF&
         FillStyle       =   0  'Solid
         Height          =   540
         Left            =   210
         Shape           =   4  'Rounded Rectangle
         Top             =   105
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
Attribute VB_Name = "FORFASQL09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TIFADOC%    ' 2=FACTURA / 3=N.DEBITO / 6=N.CREDITO
Dim MODODOC%    ' 0=MAT.CODIFICADO / 1=TEXTO LIBRE
Dim VARIDOC%
Dim PIVADOC%
Dim RNIDOC%
Dim DATEDOC%
Dim VERLICRE%
Dim MODESCU%
Dim FVENCIM
'
Dim CONTROFIS% ' Si 0 --> No Inicializado /// 1 --> Controlador Fiscal /// -1 --> Printer Común
Dim FACELEC%
Dim YAIMPRIMIO%
'
Dim CONSICLI%, PREFACTU$
'
Dim DESELEC As String, HASELEC As String
'
Dim NUOCOMPRA$
Dim NROPEDIDO$
Dim NROREMICO$
'
Dim ARCHIPRE$
Dim DATCLI$(32)
'
Dim NFCAM%    ' flag (default=0) que se pone en 1 si se ha cambiado manualmente en nro de factura
'
Dim MONECO$, TICAMBX#
Dim NOCAMONE%
Dim APLIVA%, RNI%
'
Dim COPIASREM%
'
Dim TASACLICK
Dim FACTLISTAEMPAQUE& ' 0 SI NO SE ESTA FACTURANDO LISTA DE EMPAQUE O CASO CONTRARIO EL NUMERO DE LA MISMA

Sub GRABACONCAE(DOCUM$)
    '
    FACELECOMUN% = XVALO(CONTROL("FACTURA", "PUVENELE"))
    FACELEBIENCAP% = XVALO(CONTROL("FACTURA", "PUVENCAP"))
    FACELEEXPO% = XVALO(CONTROL("FACTURA", "PUVEFAEX"))
    PTOVTA% = XVALO(Label12.Caption)
    '
    Select Case PTOVTA%
        Case FACELECOMUN%: ARCHIVOLECTURA$ = "consulta_cae.txt"
        Case FACELEBIENCAP%: ARCHIVOLECTURA$ = "consulta_cae_WSBFE.txt"
        Case FACELEEXPO%: ARCHIVOLECTURA$ = "consulta_cae_WSFEX.txt"
        Case Else: Exit Sub
    End Select
    '
    TX$ = ""
    CARPETOK$ = "C:\FLEXOFT\FACELEC\"
    NARCHI2% = FILEOPEN%(CARPETOK$ + ARCHIVOLECTURA$, 1, 128) ' ABRE EL ARCHIVO DE ESCRITURA
    Do While Not EOF(NARCHI2%) ' RECORRE
       Line Input #NARCHI2%, XX$ ' LEE CADA REGISTRO
       TX$ = TX$ + XX$ + vbCrLf
    Loop
    Close NARCHI2%
    '
    Call ABRECABAN
    Call ACTUREGISTRO("CAJABANC", "TEXTOCAE", "CODICOM_LAR = '" & DOCUM$ & "' AND TIPOMOVIM='FVEN'", TX$, REGAF&, "NOMESS")
    Call ACTUREGISTRO("CAJABANC", "TEXTOCAE", "CODICOM_LAR = '" & DOCUM$ & "' AND TIPOMOVIM='VENTA'", TX$, REGAF&, "NOMESS")
    Call ACTUREGISTRO("CAJABANC", "TEXTOCAE", "CODICOM_LAR = '" & DOCUM$ & "' AND TIPOMOVIM='IVA'", TX$, REGAF&, "NOMESS")
    '
End Sub
'
Function VALIDASTOCK()
   '
   CAITMO% = 0
   For I& = 1 To ULTREG&("!CARDEFAC")
       XX$ = REGLEIDO$("!CARDEFAC", I&)
       CI% = CVI(Mid$(XX$, 5, 2))
       CAN# = CVS(Mid$(XX$, 1, 4))
       If CAN# > 0.005 Then
         CAITMO% = CAITMO% + 1
         CIMOVIM%(CAITMO%) = CI%
         CAMOVIM(CAITMO%) = (-1) * CAN#
       End If
   Next I&
   '
   DepOrig% = CONTROL$("", "DEPOEXPE")
   AA = PUEDEMOVI%(DepOrig%)
   VALIDASTOCK = AA
   '
End Function
'
Sub CARDETCRE(OKX%)

    OKX% = 0
    NC% = XVALO(Text1)
    Call COPYSTRU("SELCUERE", "SELFCND")
    SQLX$ = " TipoMovim = 'FVEN' "
    SQLX$ = SQLX$ + " AND (OPCIMovim = 'FC' OR OPCIMovim = 'ND') "
    SQLX$ = SQLX$ + " AND NuCLIEN = " & NC% & " "
    SQLX$ = SQLX$ + " ORDER BY FechConta DESC "
    Call CARGALISEL("SELFCND", "CAJABANC", "NUMECOMP/I6;CODICOM_LAR/A20;FechConta/D10", SQLX$, "DISTINCT")
    '
    Call TRALOCAL("SELFCND", "")
    '
    If ULTREG("SELFCND") < 1 Then Exit Sub ' VALIDA QUE TENGA INFORMACION
    Call SELECHO("SELFCND") ' PRESENTO PARA SELECCIONAR FACTUR
    If VALACT1$("SELFCND") = "" Then Exit Sub ' SI NO ELIGIO REGISTRO
    '
    NFACTUR$ = Left$(VALACT2$("SELFCND"), 18) ' FACTURA SELECCIONADA
    '
    CONDI$ = "DOSECLAR = '" & NFACTUR$ & "' AND COD_INTE > 0"
    CANTCI& = CantRegistros("MOVISTO", CONDI$) 'CANTIDAD DE REGISTROS CON CODIGO INTERNO
    If CANTCI& < 1 Then
       Label6.Caption = "(T.Libre)": MODODOC% = 1
    Else
        Label6.Caption = "(Mat.Codif)": MODODOC% = 0
    End If
    
    If MODODOC% = 1 Then ' texto libre
        '
        CONDI$ = "NUMFACTURA = '" & NFACTUR$ & "'"
        CANTTRAZ& = CantRegistros("TANUMSE", CONDI$) '
        If CANTTRAZ& > 0 Then 'VALIDA QUE NO HAYA ART.TRAZABLES EN TANUMSE
            Call MENSERR(24, "Hay Equipos Trazables (Nivel 2)\Que No se Han Reingresado\Deberá Primero Reingresar Los Mismos ")
            OKX% = 0
            Exit Sub
        End If
        'TRAE EL IMPORTE DE LA FACTURA Y LO PRESENTA EN CARDETFAC
        CONDI$ = "TIPOMOVIM = 'VENTA' AND CODICOM_LAR = '" & NFACTUR$ & "'AND NuCLIEN = " & NC% & ""
        TOT# = XVALO(VALOBADA("CAJABANC", "IHABECOMP", CONDI$, "NOMESS"))
        IMPOR = TOT# ' PARA REPETIR EL IMPORTE
        '
        TX$ = REGBLAN$("!CARDEFAC")
        NFACT$ = "Corresponde A Factura: " + NFACTUR$
        '
        Call SETULTREG("!CARDEFAC", 0)
        Call REPLA(TX$, MKS$(1), 1, 4)
        Call REPLA(TX$, NFACT$, 7, 64)
        Call REPLA(TX$, MKS(IMPOR), 429, 4)
        Call REPLA(TX$, MKD(TOT#), 433, 8)
        Call GRAREG("!CARDEFAC", TX$, 1)
        Call REPLA(TX$, MKS$(PORCEIVA(CI%)), 441, 4)
        OKX% = 1
        ' Advertencia por que es posible que el total de la factura difiera segun tasa de iva
        Call COMUNI("Verifique Que la Tasa de Iva \Aplicada a Esta Factura Sea la Correcta  ")

          '
    ElseIf MODODOC% = 0 Then ' material codificado
        '
        '
        SQLX$ = "SELECT * FROM MOVISTO WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE DOSECLAR  = '" & NFACTUR$ & "'"
        SQLX$ = SQLX$ + " AND NUME_CLIE = " & NC% & ""
        SQLX$ = SQLX$ + " AND COD_INTE > 0 "
        Set FACTTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
        JX& = 0
        With FACTTMP
          Do While Not .EOF ' CARGA TODOS LOS ITEMS DEL MOVISTO RELACIONADO CON FACTURA
           CI% = XVALO(!COD_INTE)
           NUSERIE$ = TRIM$(SAFETEXT$(!IdenLote))
           ' If ESTRAZABLE%(CI%) < 2 Then NUSERIE$ = "S/N"
           JX& = JX& + 1
           CAN = XVALO(!Cantsalid)
           PRECIO = XVALO(!VALOUNIT)
           TX$ = REGBLAN$("!CARDEFAC")
           Call REPLA(TX$, MKS$(CAN), 1, 4)
           Call REPLA(TX$, MKI$(CI%), 5, 2)
           Call REPLA(TX$, DESCRIT$(CI%), 7, 64)
           Call REPLA(TX$, MKS$(PRECIO), 421, 4)
           Call REPLA(TX$, MKS$(0), 425, 4)       ' DESCUENTO
           Call REPLA(TX$, MKS$(PRECIO), 429, 4)
           IMPO# = PRECIO * CAN
           Call REPLA(TX$, MKD$(IMPO#), 433, 8)
           Call REPLA(TX$, MKS$(PORCEIVA(CI%)), 441, 4)
           Call REPLA(TX$, NUSERIE$, 139, 64)
           Call GRAREG("!CARDEFAC", TX$, JX&)
           '
           .MoveNext
          Loop
        End With
        FACTTMP.Close
        Set FACTTMP = Nothing
        '
        ' RECORRE EL TANUMSE Y LE APLICA EL NUMERO DE SERIE
        '
        SQLX$ = "SELECT * FROM TANUMSE WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE NUMFACTURA = '" & NFACTUR$ & "'"
        SQLX$ = SQLX$ + " AND NUMECLI = " & NC%
        Dim TANUTMP As ADODB.Recordset
        Set TANUTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
        With TANUTMP
          Do While Not .EOF
            CI_TANUMSE% = XVALO(!COD_INTE)
            NUMESER$ = SAFETEXT$(!NUMEROSERIE)
            For I& = 1 To ULTREG&("!CARDEFAC")
              XX$ = REGLEIDO$("!CARDEFAC", I&)
              CI% = CVI(Mid$(XX$, 5, 2))
              If CI% = CI_TANUMSE% Then
                 NUMSERCFAC$ = TRIM$(Mid$(XX$, 139, 64))
                 If NUMSERFAC$ = NUMESER$ Then GoTo 110    ' YA LO ENCONTRO EN EL CARDEFAC
                 If NUMSERCFAC$ = "" Then
                    Call REPLA(XX$, NUMESER$, 139, 64)
                    Call GRAREG("!CARDEFAC", XX$, I&)
                    GoTo 110
                 End If
              End If
            Next I&
110       .MoveNext
          Loop
        End With
        TANUTMP.Close
        Set TANUTMP = Nothing
        '
'        ' VALIDA QUE LOS TRAZABLES 2 QUE NO TIENEN NUMEROS DE SERIE LOS QUITA
'        J& = 0
'        For I& = 1 To ULTREG&("!CARDEFAC")
'           XX$ = REGLEIDO$("!CARDEFAC", I&)
'           CI% = CVI(Mid$(XX$, 5, 2))
'           NUMESER$ = TRIM$(Mid$(XX$, 139, 64))
'           If ESTRAZABLE%(CI%) > 1 And NUMESER$ = "" Then GoTo 120
'           J& = J& + 1
'           Call GRAREG("!CARDEFAC", XX$, J&)
'120     Next I&
        'VERIFICO SI NO ESTA VACIO EL CARDEFAC
        '
        Call SETULTREG("!CARDEFAC", JX&) ' marca fin de cardefac

        If JX& < 1 Then ' si no encontro ningun ariticulo
           TTXX$ = "No Hay Materiales Para Descontar de Esta Factura.\"
           TTXX$ = TTXX$ + "Posible Cliente con Grupo Divisional.\"
           TTXX$ = TTXX$ + "Para Anular la Factura Debe Realizar una Nota de Crédito de Texto Libre\"
           TTXX$ = TTXX$ + "y Re-ingresar el Material Mediante la Opción de Devolución de Materiales\"
           TTXX$ = TTXX$ + "del Modulo de Remitos o a partir de un Ajuste de Stock"
           Call COMUNI(TTXX$)
           OKX% = 0
           Exit Sub
        Else
           OKX% = 1
        End If
        
    End If


'''    CODIGO ANTERIOR COMENTADO DE RESPALDO

'''    OKX% = 0
'''    NC% = XVALO(Text1)
'''    Call COPYSTRU("SELCUERE", "SELFCND")
'''    SQLX$ = " TipoMovim = 'FVEN' "
'''    SQLX$ = SQLX$ + " AND (OPCIMovim = 'FC' OR OPCIMovim = 'ND') "
'''    SQLX$ = SQLX$ + " AND NuCLIEN = " & NC% & " "
'''    SQLX$ = SQLX$ + " ORDER BY FechConta DESC "
'''    Call CARGALISEL("SELFCND", "CAJABANC", "NUMECOMP/I6;CODICOM_LAR/A20;FechConta/D10", SQLX$, "DISTINCT")
'''    '
'''    Call TRALOCAL("SELFCND", "")
'''    '
'''    If ULTREG("SELFCND") < 1 Then Exit Sub ' VALIDA QUE TENGA INFORMACION
'''    Call SELECHO("SELFCND") ' PRESENTO PARA SELECCIONAR FACTUR
'''    If VALACT1$("SELFCND") = "" Then Exit Sub ' SI NO ELIGIO REGISTRO
'''    '
'''    NFACTUR$ = Left$(VALACT2$("SELFCND"), 18) ' FACTURA SELECCIONADA
'''    '
'''    CONDI$ = "DOSECLAR = '" & NFACTUR$ & "' AND COD_INTE > 0"
'''    CANTCI& = CantRegistros("MOVISTO", CONDI$) 'CANTIDAD DE REGISTROS CON CODIGO INTERNO
'''    If CANTCI& < 1 Then
'''       Label6.Caption = "(T.Libre)": MODODOC% = 1
'''    Else
'''        Label6.Caption = "(Mat.Codif)": MODODOC% = 0
'''    End If
'''
'''    If MODODOC% = 1 Then ' texto libre
'''        '
'''        CONDI$ = "NUMFACTURA = '" & NFACTUR$ & "'"
'''        CANTTRAZ& = CantRegistros("TANUMSE", CONDI$) '
'''        If CANTTRAZ& > 0 Then 'VALIDA QUE NO HAYA ART.TRAZABLES EN TANUMSE
'''            Call MENSERR(24, "Hay Equipos Trazables (Nivel 2)\Que No se Han Reingresado\Deberá Primero Reingresar Los Mismos ")
'''            OKX% = 0
'''            Exit Sub
'''        End If
'''        'TRAE EL IMPORTE DE LA FACTURA Y LO PRESENTA EN CARDETFAC
'''        CONDI$ = "TIPOMOVIM = 'VENTA' AND CODICOM_LAR = '" & NFACTUR$ & "'AND NuCLIEN = " & NC% & ""
'''        TOT# = XVALO(VALOBADA("CAJABANC", "IHABECOMP", CONDI$, "NOMESS"))
'''        IMPOR = TOT# ' PARA REPETIR EL IMPORTE
'''        '
'''        TX$ = REGBLAN$("!CARDEFAC")
'''        NFACT$ = "Corresponde A Factura: " + NFACTUR$
'''        '
'''        Call SETULTREG("!CARDEFAC", 0)
'''        Call REPLA(TX$, MKS$(1), 1, 4)
'''        Call REPLA(TX$, NFACT$, 7, 64)
'''        Call REPLA(TX$, MKS(IMPOR), 429, 4)
'''        Call REPLA(TX$, MKD(TOT#), 433, 8)
'''        Call GRAREG("!CARDEFAC", TX$, 1)
'''        Call REPLA(TX$, MKS$(PORCEIVA(CI%)), 441, 4)
'''        OKX% = 1
'''        ' Advertencia por que es posible que el total de la factura difiera segun tasa de iva
'''        Call COMUNI("Verifique Que la Tasa de Iva \Aplicada a Esta Factura Sea la Correcta  ")
'''
'''          '
'''    ElseIf MODODOC% = 0 Then ' material codificado
'''        '
'''        CONDI$ = "NUMFACTURA = '" & NFACTUR$ & "'"
'''        Call COPYSTRU("INDINUSE", "NCINDINU")
'''
'''        Call SETULTREG("!CARDEFAC", 0)
'''        Call SETULTREG("!NCINDINU", 0)
'''        '
'''        SQLX$ = "SELECT * FROM MOVISTO WITH(NOLOCK)"
'''        SQLX$ = SQLX$ + " WHERE DOSECLAR  = '" & NFACTUR$ & "'"
'''        SQLX$ = SQLX$ + " AND NUME_CLIE = " & NC% & ""
'''        Set FACTTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
'''        JX& = 0
'''        With FACTTMP
'''          Do While Not .EOF
'''           CI% = XVALO(!COD_INTE)
'''           CAN = XVALO(!cantsalid)
'''           'PRECIO = XVALO(!ValoUnit)
'''
'''           If CI% < 1 Then GoTo 100 ' PARA QUE NO MUESTRE ARTICULOS NO CODIFICADOS
'''           '
'''           If ESTRAZABLE%(CI%) > 1 Then GoTo 100 'Valida si estrazable
'''           JX& = JX& + 1
'''
''''           NUSERIE$ = VALOBADA("TANUMSE", "NUMEROSERIE", CONDI$, "NOMESS")
'''           DESCRI$ = DESCRIT$(CI%)
'''           Call REPLA(XX$, "S/N", 1, 16)
'''           Call REPLA(XX$, CODEXT$(CI%), 18, 18)
'''           Call REPLA(XX$, "Cant.= " & CAN, 37, 10)
'''           Call GRAREG("!NCINDINU", XX$, JX&) '*****CARGAR EL FORMULARIO CON LIST DE SELECCION
'''100        .MoveNext
'''          Loop
'''        End With
'''        FACTTMP.Close
'''        Set FACTTMP = Nothing
'''        '
'''        Call CREATABLA_EQUIPOS
'''        SQLX$ = "SELECT * FROM TANUMSE "
'''        SQLX$ = SQLX$ + " WHERE NUMFACTURA = '" & NFACTUR$ & "'"
'''        SQLX$ = SQLX$ + " AND NUMECLI = " & NC% & ""
'''        Dim TANUTMP As ADODB.Recordset
'''        Set TANUTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
'''        With TANUTMP
'''            '
'''            Do While Not .EOF
'''             CI% = XVALO(!COD_INTE)
'''             NUSERIE$ = SAFETEXT$(!NUMEROSERIE)
'''             JX& = JX& + 1
'''             Call REPLA(XX$, NUSERIE$, 1, 16)
'''             Call REPLA(XX$, CODEXT$(CI%), 18, 18)
'''             Call REPLA(XX$, "Cant.= 1", 37, 10)
'''             Call GRAREG("!NCINDINU", XX$, JX&) '*****CARGAR EL FORMULARIO CON LIST DE SELECCION
'''           '
'''110          .MoveNext
'''          Loop
'''        End With
'''        TANUTMP.Close
'''        Set TANUTMP = Nothing
'''        '
'''        Call SETULTREG("!NCINDINU", JX&)
'''        Call FRESHECHO("!NCINDINU")
'''        '
'''        If ULTREG&("!NCINDINU") < 1 Then
'''           Call COMUNI("No Hay Equipos Para Descontar de Esta Factura\Probablemente Se Haya Realizado N.Credito\o Reingeso de Equipos de la Misma")
'''           OKX% = 0
'''           Exit Sub
'''        End If
'''        '
'''        '*****ACA EN VEZ DE PRESENTAR SELECHO PRESENTAR FORMULARIO CON LIST PARA SELECCIONAR ARTICULOS
'''
'''        Call SELECHO("!NCINDINU")
'''        '
'''        If VALACT1("!NCINDINU") = "" Then
'''          OKX% = 0
'''          Exit Sub
'''        End If
'''        '
'''        ' *****LUEGO DE QUE SE ACEPTA LA VENTANA CARGAR EL CARDETFAC CON LA INFORMACION
'''        ' *****DE LOS ARTICULOS SELECCIONADOS EN EL FORMULARIO DE SELECCION
'''
'''        Call SETULTREG("!CARDEFAC", 0)
'''        CAN = TRIM$(Mid$(VALACT2("!NCINDINU"), 26, 10))
'''        CEXT$ = TRIM$(Mid$(VALACT2("!NCINDINU"), 1, 18))
'''        CI% = CODINT%(CEXT$)
'''        NUMESER$ = TRIM$(VALACT1$("!NCINDINU"))
'''        CONDI$ = "DOSECLAR = '" & NFACTUR$ & "'AND COD_INTE = " & CI% & ""
'''        PRECIO = XVALO(VALOBADA("MOVISTO", "VALOUNIT", CONDI$, "NOMESS"))
'''        TX$ = REGBLAN$("!CARDEFAC")
'''        Call REPLA(TX$, MKS$(CAN), 1, 4)
'''        Call REPLA(TX$, MKI$(CI%), 5, 2)
'''        Call REPLA(TX$, DESCRIT$(CI%), 7, 64)
'''        Call REPLA(TX$, MKS$(PRECIO), 421, 4)
'''        Call REPLA(TX$, MKS$(0), 425, 4)       ' DESCUENTO
'''        Call REPLA(TX$, MKS$(PRECIO), 429, 4)
'''        IMPO# = PRECIO * CAN
'''        Call REPLA(TX$, MKD$(IMPO#), 433, 8)
'''        Call REPLA(TX$, MKS$(PORCEIVA(CI%)), 441, 4)
'''        Call REPLA(TX$, NUMESER$, 139, 64)
'''        Call GRAREG("!CARDEFAC", TX$, 1)
'''        Call SETULTREG("!CARDEFAC", 1)
'''        OKX% = 1
'''    End If
'''        '
'''
End Sub


Sub IMPRIMENOTAENTREGA()

       FORIPRE$ = TRIM$(CONTROL$("", "FORNOTEN"))
       If TRIM$(FORIPRE$) = "" Then Exit Sub
       '
       PREFIREM$ = CONTROL$("", "PREFIREM")
       If PREFIREM$ = "" Then PREFIREM$ = "0001"
       NroRemi$ = TRIM$(Text10.TEXT)
       While Len(NroRemi$) < 8: NroRemi$ = "0" + NroRemi$: Wend
       NroRemi$ = PREFIREM$ + "-" + NroRemi$
       '
       NC% = XVALO(Text1.TEXT)
       FECHDOC$ = Text9.TEXT
       NUMEDOC$ = NroRemi$
       TIPODOC$ = "Nota de Entrega"
       '
       CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = TRIM$(Mid$(TRXP$, 1, 32))
       CODPARAM$(2) = "CUIT-TRA": DOCPARAM$(2) = TRIM$(Mid$(TRXP$, 65, 16))
       'CODPARAM$(6) = "DOCUCOND": DOCPARAM$(6) = TRIM$(Mid$(TRXP$, 129, 24))
       CODPARAM$(3) = "NRO-REMI": DOCPARAM$(3) = TRIM$(NroRemi$)
       CODPARAM$(4) = "DOMI-TRA":  DOCPARAM$(4) = LENTREGA(0).TEXT
       CODPARAM$(5) = "LOCA-TRA":  DOCPARAM$(5) = TRIM$(Mid$(TRXP$, 159, 32))
       CODPARAM$(6) = "TRANSPOR": DOCPARAM$(6) = LENTREGA(2).TEXT
         NTRA% = XVALO(Text19)
       '
       CAPARDOC% = 6
       '
       CODTABU1$(1) = "COD-MAT"
       CODTABU1$(2) = "DES-MAT"
       CODTABU1$(3) = "CANTIDAD"
       CODTABU1$(4) = "ORD-ITEM"
       CACOLTAB1% = 4
       '
       DESTIDOC% = NC%
       CAITAB1% = 0
       '
       '
       ARCHILIB$ = "!1CARDEFA"
       For U& = 1 To ULTREG&("!CARDEFAC")
         XXX$ = REGLEIDO$("!CARDEFAC", U&)
         CANS$ = Str$(CVS(Left$(XXX$, 4))) ' cantidad
         NC1% = XVALO(Text1.TEXT)
         CIPRE% = CVI(Mid$(XXX$, 5, 2))
         'If XVALO(CANS$) >= 0.05 Then
           If TRIM$(Mid$(XXX$, 387, 16)) = "" Then
             CAITAB1% = CAITAB1% + 1
             DDTXT$ = Mid$(XXX$, 7, 64)                  ' descripcion
             '
             CDXXS$ = CODEXT$(CVI(Mid$(XXX$, 5, 2)))     ' codigo
             '\\\OT-06-0339-MH-24/08/06///
             If REPLAORI% = 1 Then
               ' se pasa el codigo interno, mas el nº cliente(NC1%) y trae el codigo externo solicitado
               CDXXS$ = CODORIG$(CIPRE%, NC1%) ' + SUBCO$
             End If
             '\\\OT-06-0339-MH-FIN///
             TETABU1$(1, CAITAB1%) = CDXXS$
             TETABU1$(2, CAITAB1%) = DDTXT$
             TETABU1$(3, CAITAB1%) = CANS$
             TETABU1$(4, CAITAB1%) = Str$(CAITAB1%)
             '
             'IMPRIME DETALLE DE LOTES REMITIDOS
             Call DETLOTREM(CIPRE%, U&)
             '
             Call CARLINSEP         'Carga linea de separacion entre codigo y codigo
             '
 '
           End If
         'End If
       Next U&
       '
       '
80     CACOPDOC% = XVALO(FORFINFAC.Text2(1).TEXT)
       Call SETPRINPORT("RESTORE")
       Call PRINTDOC("@" + FORIPRE$) ' Nota de Remito
       '
End Sub





Private Sub BASIVA_Change(Index As Integer)
   If Index < 2 Then
      Total(Index + 3).TEXT = CSTRING$(MKD$(XVALO(BASIVA(Index)) * XVALO(PORIVA(Index).Caption) * APLIVA% / 100), 3, 12, 2, 2)
      INOGRA# = XVALO(Total(2)) - XVALO(BASIVA(0)) - XVALO(BASIVA(1))
      TLABEL28 = TRIM$(FORMATNUM$(INOGRA#, "F12.2"))
      Call CALTOFA2
   End If
End Sub

Private Sub BASIVA_DBLClick(Index As Integer)
   '
   If Index <= 1 Then
     If Label6.Caption = "(T.Libre)" Then
       NCLIE% = XVALO(Text1)
       RNI% = 0: APLIVA% = 1
       If NCLIE% > 0 Then
         PPVI% = PIVACLI%(NCLIE%)
         If PPVI% = 2 Then RNI% = 1
         If PPVI% = 5 Or PPVI% = 7 Then APLIVA% = 0
       End If
       '
       Select Case Index
       Case 0 '21%
          If XVALO(BASIVA(0)) < 0.005 Then
            TASACLICK = XVALO(PORIVA(0).Caption)
            BASIVA(0) = BASIVA(1)
            BASIVA(1) = ""
          End If
          Call CALTOFA2
       Case 1 '10.5%
          If XVALO(BASIVA(1)) < 0.005 Then
            TASACLICK = XVALO(PORIVA(1).Caption)
            BASIVA(1) = BASIVA(0)
            AA1 = XVALO(BASIVA(1))
            aa2 = PORIVA(1)
            aa3 = APLIVA%
            BASIVA(0) = ""
          End If
          Call CALTOFA2
       End Select
     End If
   End If
   '
End Sub

Private Sub BASIVA_GotFocus(Index As Integer)
   If VALICLIE% < 1 Then GoTo 99
   If Index > 1 Then GoTo 98
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

Private Sub BASIVA_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
     Call BASIVA_LostFocus(Index)
   End If
End Sub

Private Sub BASIVA_LostFocus(Index As Integer)
   BASIVA(Index) = TRIM$(FORMATNUM$(BASIVA(Index), "F13.2"))
End Sub

Private Sub BASIVA_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
     '
     If Index <= 1 Then
        If Label6.Caption = "(T.Libre)" Then
            BASIVA(Index).ToolTipText = "Doble Click para Cambiar Tasa IVA"
        Else
            BASIVA(Index).ToolTipText = ""
        End If
     End If
     '
End Sub

Private Sub Check1_Click()
   If MUECOMHAS% <= 1 Then MUECOMHAS% = Check1.Value
End Sub

Private Sub Command1_Click()
   Call SELECHO("DEUDOR1")
   NCLIE% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
   If NCLIE% > 0 Then
     Text1.TEXT = TRIM$(Str$(NCLIE%))
     If XVALO(Text1.TEXT) > 0 Then
       If XVALO(Text1.TEXT) <= 32767 Then
         NC% = XVALO(Text1)
         'If REGICLI&(NC%) > 0 Then
           Call CARDACLI
         'End If
       End If
     End If
   End If
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call CIERRARCH("*.*")
   Call CONECRUN("ARCHIG07", "Archivos Maestros")
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Call SELECHO("ECUECON")
   VDEVU$ = TRIM$(VALACT1$("ECUECON"))
   If VDEVU$ <> "" Then
     Text16 = VDEVU$
   End If
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   COGEVE07.Show 1
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   If VALICLIE% < 1 Then Exit Sub
   If Abs(XVALO(Total(6))) < 0.005 Then Exit Sub
   '
   If VALIDAPUNVEN% < 1 Then Exit Sub
   '
   Command14.Enabled = False
   '
   If TRIM$(CONTROL$("STOCKS", "CONEGATI")) = "SI" And Label6.Caption <> "(T.Libre)" And CONSICLI% < 1 And TIFADOC% <= 3 Then
     If VALIDASTOCK < 1 Then Command14.Enabled = True: Exit Sub
   End If
   '
   If Text22.Visible = True And TRIM$(Text22) <> "" Then
      If IsNumeric(Text22.TEXT) <> True Then
        Call MENSERR(24, "Comisión No Válida")
        Command14.Enabled = True
        Exit Sub
      End If
   End If
   '
'5  If CONTROFIS% = 2 Then ' validacion de la longitud de razon social y el domicilio para impresora fiscal
'     If Len(Text2) > 30 Then
'       RZO$ = InputBox("Razon Social Supera Límite de Caracteres (30)", "Cambiar Manualmente Razon Social", Text2.TEXT)
'       If TRIM$(RZO$) = "" Then GoTo 5
'       Text2 = TRIM$(RZO$)
'       GoTo 5
'     End If
'     '
'     If Len(Text3) > 30 Then
'       RZO$ = InputBox("Domicilio Supera Límite de Caracteres (30)", "Cambiar Manualmente Domicilio", Text3.TEXT)
'       If TRIM$(RZO$) = "" Then GoTo 5
'       Text3 = TRIM$(RZO$)
'       GoTo 5
'     End If
'   End If
   '
   '\\\OT-07-0023-RG-24/01/07///
   FORFINFAC.Text7 = TRIM$(Str$(DEPOEXPE%))
   '
   '\\\OT-07-0023-RG-FIN///
   If XVALO(TLABEL28) < 0 Then
      Call MENSERR(24, "Imposible Registrar este Documento.\Importe No Gravado no Admisible.")
      Command14.Enabled = True
      Exit Sub
   End If
   '
   If XVALO(Text15) >= 0.005 Then
     If XVALO(Total(4)) >= 0.005 Then
       Call MENSERR(24, "Imposible Registrar este Documento.\Incompatibilidad entre Percepcion IB\y Tasa IVA Secundaria - Consulte.")
       Command14.Enabled = True
       Exit Sub
     End If
     CUEPERI$ = CONTROL$("PERIBRU", "CUECONTA")
     CUENPER% = CUECOINT%(CUEPERI$)
     If TRIM$(CUEPERI$) = "" Or CUENPER% < 1 Then
       Call MENSERR(24, "Imposible Registrar este Documento.\Falta Cuenta Contable Correspondiente\a Percepción de Ingresos Brutos.")
       Command14.Enabled = True
       Exit Sub
     End If
   End If
   '
   FECHALIM% = XVALO(CONTROL$("", "CIEMEVEN"))
   If FECHANUM(Text9) <= FECHALIM% Then
     Call MENSERR(24, "Fecha no Válida.\  \Ese Período Mensual Está Cerrado.")
     Command14.Enabled = True
     Exit Sub
   End If
   '
18 EPAC$ = TRIM$(UCase$(EMPREAC$))
   '
   FINFACEX$ = ""
   MONECO$ = MONEMIS
   '
20 NCLIE% = XVALO(Text1.TEXT)
   If VERLICRE% = 1 Then
     If TIFADOC% < 4 Then ' EXCLUYE CREDITOS
       '
       For U& = 1 To ULTREG&("!CARDEFAC")
         X$ = REGLEIDO$("!CARDEFAC", U&)
         NROPEDXX& = XVALO(TRIM$(Mid$(X$, 371, 16)))
         If NROPEDXX& < 1 Then GoTo 22
         RFF$ = RECFILT$("AUTORENT", 1, MKS$(NROPEDXX&))
         If Len(RFF$) < 4 Then GoTo 22 ' NO AUTORIZADO
         REGAUT& = CVS(RFF$)
         XAUT$ = REGLEIDO$("AUTORENT", REGAUT&)
         If UCase$(Mid$(XAUT$, 31, 2)) <> "SI" Then GoTo 22
       Next U&
       GoTo 25
       '
22     LMCR# = XVALO(Text7)     '  LIMICRE(NCLIE%)
       SDEU# = SALDCLI#(NCLIE%)
       CHPA# = CHEPEACRE#(NCLIE%)
       TPED# = XVALO(Total(6).TEXT) * TICAMBX#
       If SDEU# + CHPA# + TPED# > LMCR# Then
         TLIMI$ = TRIM$(CSTRING$(MKD$(LMCR#), 3, 12, 2, 2))
         If TLIMI$ = "" Then TLIMI$ = "0.00"
         Call COMUNI("Imposible Emitir Este Documento.\  \Supera el Límite de Crédito Asignado\para esta Cuenta de $ " + TLIMI$ + ".\  \Informe al Supervisor de Ventas.")
         Command14.Enabled = True
         Exit Sub
       End If
     End If
   End If
   '
25 'VERIFICACION DE LIMITE DE CREDITO EN CUENTA CORRIENTE
   TPED# = XVALO(Total(6).TEXT) * TICAMBX#
   If EXEDELIMCC%(NCLIE%, TPED#, DIFER) > 0 Then
     If TIFADOC% < 4 Then ' EXCLUYE CREDITOS
       LIMCCTE$ = TRIM$(FORMATNUM$(LCred_CCte(NCLIE%), "f12.2"))
       TX$ = "Esta Factura + el Saldo Supera\el Límite de Crédito Asignado en C.Cte.\para esta Cuenta de $ " + LIMCCTE$
       If DERACCE%("AUTOFACT", "Autorizar Facturas con Exceso Credito") < 2 Then
         Call COMUNI("Imposible Emitir Este Documento.\  \" + TX$)
         Command14.Enabled = True
         Exit Sub
       Else
         If COMALTER%(TX$ + "\   \Puede Autorizar la Facturación\con Crédito Excedido.\\Cancelar\Autorizar Factura") < 2 Then
            Command14.Enabled = True
            Exit Sub
         End If
       End If
     End If
   End If

   ' ojo aca grabar lugares de entrega
   Call APERBASDAT("CLIEN")
   CONDICION$ = "Nume_Cli = " & NCLIE% & ""
   '
   TTXXT0$ = Left$(LENTREGA(0) + " ", 48)
   Call ACTUREGISTRO("DEUDOR12", "Lentre0_Cli", CONDICION$, TTXXT0$, REGAF&, "NOMESS")
   '
   TTXXT1$ = Left$(LENTREGA(1) + " ", 48)
   Call ACTUREGISTRO("DEUDOR12", "Lentre1_Cli", CONDICION$, TTXXT1$, REGAF&, "NOMESS")
   '
   TTXXT2$ = Left$(LENTREGA(2) + " ", 48)
   Call ACTUREGISTRO("DEUDOR12", "Lentre2_Cli", CONDICION$, TTXXT2$, REGAF&, "NOMESS")
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
30 If CONSICLI% = 1 Then
     REMI = 0
     '
     If TRIM$(Text14.TEXT) = "" Then
       Text14.TEXT = NUOCOMPRA$
     End If
     '
   End If
   '
   FORIPRE$ = TRIM$(CONTROL$("", "FORFAC-" + Label13.Caption))
   If FORIPRE$ = "" Then
      Call MENSERR(24, "Falta Definir Controlador de Impresión")
      Exit Sub
   End If
   '
   If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
      Call MENSERR(24, "Imposible Imprimir Documento.\No Existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
      Exit Sub
   End If
   '
   If MODODOC% = 0 Then    ' Material Codificado
     ARCHILOC$ = "!CARDEFAC"
     Call CIERRARCH(ARCHILOC$)
     If INTERFI$ = "" Then
       INTERFI$ = "TRAN" + RANDSTRING$(-4)   ' negativo para que sean numeros
     End If
     ALOCDA$ = Mid$(ARCHILOC$, 2)
     Call COPYSTRU(ALOCDA$, INTERFI$)
     XY1% = COPYFILE(LUCOM$ + ALOCDA$ + ".WKF", LUDAT$ + INTERFI$ + ".DAT")
     '
     Call FILESORT(INTERFI$, "", 30, 30, "ASC")   ' POR NUMERO ORDEN ITEM
     Call FILESORT(INTERFI$, "", 18, 18, "ASC")   ' POR REMITO
     Call FILESORT(INTERFI$, "", 20, 20, "ASC")   ' POR FECHA REMITO
     XY1% = COPYFILE(LUDAT$ + INTERFI$ + ".DAT", LUCOM$ + ALOCDA$ + ".WKF")
     Call BAJALDISCO
   End If
   '
   CARENGLO& = ULTREG&("!CARDEFAC")
   ITPORHOJA% = XVALO(ATRIFORM$(FORIPRE$, "MAX-ITEM"))
   If ITPORHOJA% < 1 Then
     ITPORHOJA% = 32767
   End If
   '
   CAHOJNEC% = Int(0.98 + CARENGLO& / ITPORHOJA%)
   Call GRAFACTURA(OKX%)
   Call CIERRARCH("*.*")
   Call LIBARCH("*.*")
   Call BAJALDISCO
   '
   If OKX% < 0 Then
       GoTo 10     ' ABORTADA O SOLO REMITO
     ElseIf OKX% = 0 Then
       GoTo 18
   End If
   '
   If OKX% > 0 Then
       FF% = FECHANUM(Text9.TEXT)
       REMI = XVALO(Text10.TEXT)
       If FACELEC% < 1 Then
          Call PRIFACTURA ' NO IMPRIME (DESPUER VER SI SE VA A IIMPRMIR PROVISORIO)
       Else
          CONDI$ = "CODICOM_LAR = '" & TRIM$(Label30) & "'"
          IDCAE$ = SAFETEXT$(VALOBADA("CAJABANC", "IDCAE", CONDI$))
          TESTCAE09.Text5 = IDCAE$
          Call TESTCAE09.PIDENUEVOCAE
       End If
       '
       Call PRIREMIFA ' IMPRIME SOLO REMITO ' VER POR QUE EL REMITO IMPRMIE EL NUMERO DE FACTURA
   End If
   Call CIERRARCH("*.*")
   Call LIBARCH("*.*")
   Call BAJALDISCO
   '
   If REPRIFAC% = 1 Then
      Call FINAL("")
   End If
   '
   If OKX% > 0 Then
     If InStr(UCase$(Eco21(1)), "CONTADO") > 0 Or EsPorPedido& > 0 Then
       EPAC$ = TRIM$(UCase$(EMPREAC$))
       If InStr(EPAC$, "DOSIVA") < 1 Then
          RECIBO08.Text1 = Text1
          NC% = XVALO(TRIM$(Text1))
          RECIBO08.Text2 = Text2
          RECIBO08.Text3 = Text3
          RECIBO08.Text4 = Text4
          RECIBO08.Text5 = Text5
          RECIBO08.Text9 = Text9
          RECIBO08.Label26 = CODICLI$(NC%) 'CODIGO DEL CLIENTE
          If ULTREG&("COBRADOR") > 0 Then
            XXX$ = REGLEIDO$("COBRADOR", 1)
            RECIBO08.Label22.Caption = Left$(TRIM$(Asc(Left$(XXX$, 1))) + " - " + Mid$(XXX$, 2, 30), 19)
          End If
          RECIBO08.Show 1
       End If
     End If
   End If
   '
10 Call BLANFORMU
   Call NUFACTURA
   Call CIERRARCH("*.*")
   Text1.TEXT = ""
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
   '
   If EsPorPedido& > 0 Then Call FINAL("")
   Command14.Enabled = True
End Sub

Private Sub Command23_Click()
    NDCLIEN07.Show 1
End Sub

Private Sub Command24_Click()
   '\\\OT-06-0570-RG-27/12/06///
   Call CommSel("TRANSPOR", Command24, Text19)
   '\\\OT-06-0570-RG-FIN///
   'Call SELECHO("TRANSPOR")
   'NTRAN% = XVALO(VALACT1$("TRANSPOR"))
   'Text19 = TRIM$(FORMATNUM(NTRAN%, "I4"))
End Sub

Private Sub Command25_Click()
   If EXISTE%(LUDAT$ + "SUCENTRE.RFS") > 0 Then
      '\\\OT-06-0570-RG-27/12/06///
      Call CommSel("!SUCENTRE", Command25, Text20)
      '\\\OT-06-0570-RG-FIN///
     'Call SELECHO("!SUCENTRE")
     'VDEVU$ = VALACT1$("!SUCENTRE")
     'Text20 = VDEVU$
   End If
End Sub

Private Sub Command26_Click()
    '
    Call FACLOT07.LIMPILABEL
    Call FACLOT07.LIMPITEXT
    Call FACLOT07.List2.Clear
    FACLOT07.Label15 = "ACTIVO"
    '
    Call ASIGLOT07.List1.Clear ' LIMPIO INFORMACION DEL ASIGLOT07
    Call ASIGLOT07.List2.Clear
    ASIGLOT07.Label6 = ""
    ASIGLOT07.Label10 = ""
    ASIGLOT07.Label4 = ""
    ASIGLOT07.Label8 = ""
    ASIGLOT07.Label1 = ""
    ASIGLOT07.Label3 = ""
    '
    FACLOT07.Show 1
    '
End Sub

Private Sub Label12_DblClick()
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "DOSIVA") > 0 Then Exit Sub
   '
   If DERACCE%("CAMPREFI/NOMESS", "Modificación de Prefijo") < 2 Then ' derecho de acceso para modificar prefijo sin mensaje
      Exit Sub
   End If
   '
   ' VALIDACION PARA QUE MUESTRE LOS PUNTOS DE VENTA SOLO BIENES CAPITALES O FACT.ELECT.COMUN
   If CONTROL$("FACTURA", "SALIDA") = "FELECT" Then
     If TRIM$(Label12) = TRIM$(CONTROL("FACTURA", "PUVENELE")) Or TRIM$(Label12) = TRIM$(CONTROL("FACTURA", "PUVENCAP")) Or TRIM$(Label12) = TRIM$(CONTROL("FACTURA", "PUVEFAEX")) Then
        PVTANOR% = XVALO(CONTROL("FACTURA", "PUVENELE"))
        PVTABCP% = XVALO(CONTROL("FACTURA", "PUVENCAP"))
        PVTAFAEX% = XVALO(CONTROL("FACTURA", "PUVEFAEX"))  'P.VTA. FACT.EXP.
        If PVTANOR% > 0 And PVTABCP% > 0 Then
            If TRIM$(Label12) = TRIM$(CONTROL("FACTURA", "PUVENELE")) Then
              Label12 = TRIM$(CONTROL("FACTURA", "PUVENCAP"))
              GoTo 99
            End If
            If TRIM$(Label12) = TRIM$(CONTROL("FACTURA", "PUVENCAP")) Then
              If PVTAFAEX% > 0 Then
               Label12 = TRIM$(CONTROL("FACTURA", "PUVEFAEX"))
               GoTo 99
              Else
               Label12 = TRIM$(CONTROL("FACTURA", "PUVENELE"))
               GoTo 99
              End If
            End If
            
            If TRIM$(Label12) = TRIM$(CONTROL("FACTURA", "PUVEFAEX")) Then
               Label12 = TRIM$(CONTROL("FACTURA", "PUVENELE"))
            End If
            GoTo 99
            
        End If
     End If
   End If
   
   If Label12 = "0001" Then
       Label12 = "0002"
     ElseIf Label12 = "0002" Then
       Label12 = "0003"
     ElseIf Label12 = "0003" Then
       Label12 = "0004"
     ElseIf Label12 = "0004" Then
       Label12 = "0005"
     Else
       Label12 = "0001"
   End If
99  CLAVE$ = "PREFAC-" & Label13.Caption '  MODIFICADO PARA GUARDAR PREFIJO DE FACTURA POR ESTACION
    Call GRACONTROL(CLAVE$, IDENTER$, Label12)
   '
End Sub

Private Sub Label12_Change()
   VRFF% = VERIFSALIDA%("NOMESS")
   If VRFF% = 1 Then
        Label12.ToolTipText = "Fact.Elect. - Productos / Servicios Comunes - Doble Click para Cambiar"
     ElseIf VRFF% = 2 Then
        Label12.ToolTipText = "Fact.Elect. - Régimen Bienes de Capital - Doble Click para Cambiar"
     ElseIf VRFF% = 3 Then
        Label12.ToolTipText = "Fact.Elect. - Tipo Exportación - Doble Click para Cambiar"

   End If
   Call NUFACTURA
End Sub

Private Sub Label14_DblClick()
   Call Text14_DBLCLICK
End Sub

Private Sub Label28_DblClick()
   '
   If Label28.Caption = "" Then
       Label28.Caption = "A"
     ElseIf Label28.Caption = "A" Then
       Label28.Caption = "B"
     ElseIf Label28.Caption = "B" Then
       Label28.Caption = "C"
     ElseIf Label28.Caption = "C" Then
       Label28.Caption = "D"
     ElseIf Label28.Caption = "D" Then
       Label28.Caption = "A"
   End If
   '
End Sub

Private Sub Label7_Change()
   BASIVA(0).BackColor = &HE0E0E0
   BASIVA(1).BackColor = &HE0E0E0
   Total(3).BackColor = &HE0E0E0
   Total(4).BackColor = &HE0E0E0
   If TIFADOC% > 2 Then ' DEBITOS Y CRÉDITOS
      BASIVA(0).BackColor = &H80000005
      BASIVA(1).BackColor = &H80000005
      Total(3).BackColor = &H80000005
      Total(4).BackColor = &H80000005
   End If
End Sub

Private Sub LENTREGA_GotFocus(Index As Integer)
   If Index = 2 Then
     If XVALO(Text19) > 0 Then
       Call MENSERR(24, "Campo no Editable")
       On Error Resume Next
       Text19.SetFocus
       On Error GoTo 0
     End If
   End If
End Sub

Private Sub MONEMIS_Change()
   '
   '\\\OT-07-0086-RG-06/03/07///
   ENTRATICA.Label6.Caption = MONEMIS.Caption
   ENTRATICA.Text4 = TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4"))
   '\\\OT-07-0086-RG-FIN///
   '
End Sub

Private Sub PORIVA_DblClick(Index As Integer)
   If Index <= 1 Then
     '
     SUBASIM# = XVALO(BASIVA(0)) + XVALO(BASIVA(1))
     TEBASIM$ = TRIM$(FORMATNUM$(SUBASIM#, "F12.2"))
     If Index = 0 Then
         BASIVA(0) = TEBASIM$
         BASIVA(1) = ""
         
       Else
         BASIVA(0) = ""
         BASIVA(1) = TEBASIM$
     End If
     ALIVAGEN = XVALO(PORIVA(Index))
     '
     Total(4).TEXT = CSTRING$(MKD$(XVALO(BASIVA(1)) * XVALO(PORIVA(1)) * APLIVA% / 100), 3, 12, 2, 2)
     Total(3).TEXT = CSTRING$(MKD$(XVALO(BASIVA(0)) * XVALO(PORIVA(0)) * APLIVA% / 100), 3, 12, 2, 2)
     '
     TTTX# = (XVALO(Total(3).TEXT) + XVALO(Total(4).TEXT)) * RNI%
     BASIVA(2) = TRIM$(CSTRING$(MKD$(TTTX#), 3, 12, 2, 2))
     Total(5).TEXT = CSTRING$(MKD$(XVALO(BASIVA(2)) * XVALO(PORIVA(2).Caption) / 100), 3, 12, 2, 2)
     '
     TTTX# = 0: For KKI% = 2 To 5
                    TTTX# = TTTX# + XVALO(Total(KKI%).TEXT)
                 Next KKI%
                 TTTX# = TTTX# + XVALO(Text15)
                 TTTX# = TTTX# + XVALO(Text17)
     Total(6).TEXT = CSTRING$(MKD$(TTTX#), 3, 12, 2, 2)
     '
     For KU& = 1 To ULTREG&("!CARDEFAC")
       XXX$ = REGLEIDO$("!CARDEFAC", KU&)
       INETO1# = CVD(Mid$(XXX$, 433, 8))
       Call REPLA(XXX$, MKS$(ALIVAGEN), 441, 4)
       Call REPLA(XXX$, MKD$(INETO1# * ALIVAGEN / 100), 445, 8)
       Call REPLA(XXX$, MKD$(INETO1# * (1 + ALIVAGEN / 100)), 453, 8)
       Call GRAREG("!CARDEFAC", XXX$, KU&)
     Next KU&
     Call CIERRARCH("!CARDEFAC")
     '
   End If
End Sub

Private Sub Text14_DBLCLICK()
    '
    If VALICLIE% < 1 Then
       On Error Resume Next
       Text1.SetFocus
       On Error GoTo 0
       Exit Sub
    End If
    '
    If TIFADOC% <> 2 Then
      ' solo para facturas
      Exit Sub
    End If
    '
    If MODODOC% <> 0 Then  ' 0=MAT.CODIFICADO / 1=TEXTO LIBRE
       ' solo mat codificado
       Exit Sub
    End If
    '
    If PREFACTU$ <> "REMITO" Then
      ' solo para forma de guidi
      Exit Sub
    End If
    '
    'ACA PONGO EL INDICE
    NCLIE% = XVALO(Text1)
    Screen.MousePointer = 11
    '
    Call APERBASDAT("STOCKS")
    Call APERBASDAT("PEDCLI")
    SQLX$ = " SELECT DOSECLAR,  NUPEDCLI, COD_INTE FROM MOVISTO "
    SQLX$ = SQLX$ + "WHERE Nume_Clie = " & NCLIE% & " "
    SQLX$ = SQLX$ + " AND (CodiMovi = 'RT' "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
    SQLX$ = SQLX$ + " AND DoPriLar <> '' "
    SQLX$ = SQLX$ + " AND (DoSECLar = '' "
    SQLX$ = SQLX$ + " OR DoSecLar LIKE 'P/F%' ) "
    SQLX$ = SQLX$ + " ORDER BY NUPEDCLI ASC "
    On Error GoTo 0
    'Set pedianuittemp = Stocks.OpenRecordset(sqlx$, dbOpenSnapshot)
    Dim pedianuittemp As ADODB.Recordset
    Set pedianuittemp = New ADODB.Recordset
25  On Error Resume Next
    pedianuittemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0

    '
    JJ& = 0
    With pedianuittemp
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        TTXX$ = REGBLAN("INDIFAPE")
        Do While Not .EOF
         If PEDANTI& <> !NuPedCli Then
            NRODES& = !NuPedCli
            CI% = !COD_INTE
            If NRODES& <= 0 Then GoTo 15
          TTXX$ = MKS$(NRODES&)
            NOCO$ = NOCOMCLI(NRODES&)
          If NOCO$ <> "" Then
            TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
          End If
            'With PedEnca
            '.FindFirst "NroPed = " & NRODES& & ""
            SQLX$ = " SELECT * FROM PedEnca"
            SQLX$ = SQLX$ + " WHERE NroPed =" & NRODES& & ""
            Dim PEDDDENCA As ADODB.Recordset
            Set PEDDDENCA = FLEXCONN.Execute(FILTSQL$(SQLX$))
            On Error Resume Next
            With PEDDDENCA
              'If .NoMatch = False Then
              If Err < 1 Then
                On Error GoTo 0
                FEX = CVINT(PedEnca!FechEmis)
              End If
              On Error GoTo 0
            End With
            Set PEDDDENCA = Nothing
          TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
            NCLIE% = !NroClie
          TTXX$ = TTXX$ + TRIM$(RASOCLI$(NCLIE%)) + Space$(52)
          JJ& = JJ& + 1
          Call GRAREG("INDIFAPE", Left$(TTXX$, 52), JJ&)
          PEDANTI& = NRODES&
         End If
15      .MoveNext
      Loop
      End If
      On Error GoTo 0
    End With
    Set pedianuittemp = Nothing
    Call SETULTREG("INDIFAPE", JJ&)
    Call CIERRARCH("INDIFAPE")
    Call FRESHECHO("INDIFAPE")
    Screen.MousePointer = 1
    '
    If ULTREG&("INDIFAPE") < 1 Then
       Call COMUNI("No hay Remitos Pendientes de este Pedido\que Puedan Facturarse")
       Call CIERRARCH("*.*")
       Exit Sub
    End If
    '
22  X$ = REGSEL$("INDIFAPE/Pedidos Pendientes para Facturación")
    If Len(X$) <= 4 Then GoTo 999
    '
    NUPEDIDO& = CVS(Left$(X$, 4))
    Text14 = NOCOMCLI$(NUPEDIDO&)
    CONSICLI% = 0
    Call CARCONSIPED(NUPEDIDO&)
    If CONSICLI% > 0 Then
      Call CARDETFAC
      MARCONOTA.Visible = False
      Exit Sub
    End If
    '
999  End Sub

Private Sub Text15_Change()
   Call CALTOFA2
End Sub

Private Sub Text15_GotFocus()
   '
   If VALICLIE% < 1 Then GoTo 99
   If TIFADOC% < 3 Then GoTo 98
   'If XVALO(Total(2)) < 0.5 Then GoTo 98
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
    Text15 = FORMATNUM$(XVALO(Text15), "F11.2")
  End If
End Sub

Private Sub Text15_LostFocus()
   Text15 = FORMATNUM$(XVALO(Text15), "F11.2")
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

Sub Command17_Click()
   Command17.Enabled = False
       'EPAC$ = TRIM$(UCase$(EMPREAC$))
   CONPED07.Show 1
       'Call SHOWPEDIPEN
   Command17.Enabled = True
End Sub

Private Sub Command18_Click()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Command19_Click()
   '
   NUOCOMPRA$ = ""
   NROPEDIDO$ = ""
   '
   Call CARDETFAC
   '
End Sub

Private Sub Command2_Click()
    If EsPorPedido& > 0 Then
       Call FINAL("")
       Exit Sub
    End If
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_LostFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Sub Command20_Click()
   Command20.Enabled = False
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "GUIDI") > 0 Then
      OPP% = ALTERNA%("Repetir Impresion de Factura\Otros Complementos de Facturación")
      If OPP% = 1 Then
         Call SELFACREPI
         GoTo 99
      End If
   End If
   Screen.MousePointer = 1
   OPVARFA07.Show 1
99 Command20.Enabled = True
End Sub

Sub SELFACREPI()
    '
    If DERACCE%("REPRIFAC", "Repetir Impresión de Facturas") < 2 Then
      Exit Sub
    End If
    '
10  Call SELECHO("DEUDOR1")
    NC% = XVALO(VALACT1$("DEUDOR1"))
    If NC% < 1 Then Exit Sub
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'FVEN' "
    SQLX$ = SQLX$ + "AND NuCLIEN = " & NC% & " "
    SQLX$ = SQLX$ + "Order by FECHEMISI DESC, CODICOM_LAR DESC "
    '
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    JJ& = 0
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
          NCLIE% = !NUCLIEN
          IMPO# = !ValAbsComp
          If IMPO# >= 0.005 Then
            IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
            If TRIM$(IMPOS$) = "" Then IMPOS$ = AJUSTD$("0.00", 14)
            FEMISI$ = FECHATEX$(CVINT(!fechemisi))
            NPX& = XVALO(Mid$(!codicom_lar, 11, 8))
            If Left$(!codicom_lar, 2) = "CC" Then
              NPX& = XVALO(Mid$(!codicom_lar, 4))
            End If
            Y$ = MKS$(NPX&) + AJUSTI$(!codicom_lar, 18) + AJUSTD$(FEMISI$, 14) + IMPOS$
            JJ& = JJ& + 1
            Call GRAREG("!SELCUERE", Y$, JJ&)
          End If
        .MoveNext
        Loop
      End If
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
 
    Call SETULTREG("!SELCUERE", JJ&)
    Screen.MousePointer = 1
    '
    If ULTREG&("!SELCUERE") < 1 Then
       Call MENSERR(24, "La Cuenta Elegida no tiene Facturas de Venta\que Puedan Re-Imprimirse.")
       GoTo 10
    End If
    '
    Call CIERRARCH("!SELCUERE")
    Call FRESHECHO("!SELCUERE")
    '
    Call SELECHO("!SELCUERE")
    If VALACT1$("!SELCUERE") = "" Then GoTo 10
    REGCUEC& = XVALO(VALACT1$("!SELCUERE"))
    If REGCUEC& < 1 Then GoTo 10
    '
    DOCULA$ = Left$(VALACT2$("!SELCUERE"), 18)
    NUOPANU& = XVALO(Mid$(DOCULA$, 13, 6))
    REPRIFAC% = 1
    '
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'FVEN' "
    SQLX$ = SQLX$ + "AND NuCLIEN = " & NC% & " "
    SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & DOCULA$ & "' "
    Set CABATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    Text1 = TRIM$(Str$(NC%))
    Call CARDACLI
    DOCUCO$ = CABATEMP!CodiCom_Cor
    Text9 = FECHATEX$(CVINT(CABATEMP!fechemisi))
    Label12 = TRIM$(Mid$(DOCULA$, 4, 4))
    Label13 = TRIM$(Mid$(DOCULA$, 18, 1))
    Text11 = TRIM$(Str$(XVALO(Mid$(DOCULA$, 9, 8))))
    '
    MARCONOTA.Visible = False
    Call COMUNI("ADVERTENCIA: Los datos centrales de facturación\tales como lista de precios, condicion de pago\sucursal de entrega, etc son tomados de\las BASES DE DATOS CENTRALES.\   \Recomendamos revisar visualmente si corresponden\con los de la factura y eventualmente\ajustarlos manualmente.")
    '
    Call CARDETAFAC(NC%, DOCULA$)
    Call CARDETFAC
    '
End Sub
'
Sub CARDETAFAC(NCLIE%, DOCULA$)
   '
   REGODES& = 0
   Call SETULTREG("!CARDEFAC", 0)
   If EXISTE%(LUCOM$ + "1CARDEFA.WKF") > 0 Then
     Call SETULTREG("!1CARDEFA", 0)
   End If
   JJ& = 0: CONSICLI% = 0
   '
   Screen.MousePointer = 11
   Call APERBASDAT("STOCKS")
   '
   SQLX$ = " SELECT * FROM MOVISTO "
   SQLX$ = SQLX$ + "WHERE Nume_Clie = " & NCLIE% & " "
   SQLX$ = SQLX$ + " AND DoSecLar = '" + DOCULA$ + "' "
   
   Dim CONSITMP As ADODB.Recordset
   Set CONSITMP = New ADODB.Recordset
   CONSITMP.CursorLocation = adUseClient
25 On Error Resume Next
   CONSITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   With CONSITMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       RECONSI$ = REGBLAN("CONSICLI")
       J& = 0
       Do While .EOF = False
         CODIMAT% = XVALO(!COD_INTE)
         CAN = XVALO(!Cantsalid) - XVALO(!CantIngre)
         NPEDID& = XVALO(!NuPedCli)
         NUMEPED& = 0: PREPED = 0
           If NPEDID& < 1 Then
             FEREMIT% = CVINT%(!fechMov)
             TITUCON$ = TRIM$(SAFETEXT$(!DoPriCor)) + " - Cant: " + TRIM$(Str$(CAN)) + " - Cod: " + TRIM$(CODEXT$(CODIMAT%)) + " - " + DESCRIT$(CODIMAT%)
             Call RECARPED(NCLIE%, CODIMAT%, FEREMIT%, NUMEPED&, PREPED, TITUCON$)
             NPEDID& = NUMEPED&
           End If
5        Z$ = REGBLAN$("!CARDEFAC")
         Call REPLA(Z$, MKS$(CAN), 1, 4)      ' cantidad
           COD$ = CODEXT$(CODIMAT%)
         Call REPLA(Z$, MKI$(CODIMAT%), 5, 2)
         Call REPLA(Z$, DESCRIT$(CODIMAT%), 7, 64)
           NREMIT$ = SAFETEXT$(!DoPriCor)
         Call REPLA(Z$, NREMIT$, 387, 16) ' numero de remito
           FEREMIT% = CVINT%(!fechMov)
         Call REPLA(Z$, MKI$(FEREMIT%), 419, 2)
         Call REPLA(Z$, NUPLANO$(CODIMAT%), 315, 24)
           '
7          MONEMA% = XVALO(!MoneVal)
           MONEMI% = MONEMA%
           PRECIO = XVALO(!VALOUNIT)
             '
         Call REPLA(Z$, MKS$(PRELI), 421, 4)
         Call REPLA(Z$, MKS$(PORDESCUEN), 425, 4)
         Call REPLA(Z$, MKS$(PRECIO), 429, 4)
           IMPO# = CDbl(PRECIO) * CAN
         Call REPLA(Z$, MKD$(IMPO#), 433, 8)
         Call REPLA(Z$, MKS$(PORCEIVA(CODIMAT%)), 441, 4)
         Call REPLA(Z$, " ", 461, 1)    ' NUNCA ES REPARACION
           NORIT% = XVALO(!NuOrItem)
         Call REPLA(Z$, MKI$(NORIT%), 503, 2)
         Call REPLA(Z$, MKS$(0), 505, 4)
         Call REPLA(Z$, MKS$(0), 509, 4)
         Call REPLA(Z$, TRIM$(Str$(NPEDID&)), 371, 16)
           NUOCOX$ = NUOCOCLI$(NPEDID&, CODIMAT%)
10       Call REPLA(Z$, NUOCOX$, 355, 16)
         '
         Call REGAPP("!CARDEFAC", Z$)
         CAITCAR% = CAITCAR% + 1
         '
       .MoveNext
       Loop
12   End If
   End With
   CONSITMP.Close
   Set CONSITMP = Nothing
   '
   JJ& = 0
   For U& = 1 To ULTREG&("!CARDEFAC")
     Z$ = REGLEIDO$("!CARDEFAC", U&)
     If Abs(CVS(Mid$(Z$, 1, 4))) >= 0.01 Then
       JJ& = JJ& + 1
       Call GRAREG("!CARDEFAC", Z$, JJ&)
     End If
   Next U&
   Call SETULTREG("!CARDEFAC", JJ&)
   If EXISTE%(LUCOM$ + "1CARDEFA.WKF") > 0 Then
     Call SETULTREG("!1CARDEFA", 0)
     Call CIERRARCH("!1CARDEFA")
   End If
   Call CIERRARCH("!CARDEFAC")
   '
End Sub
'
Private Sub Command22_Click()
    Command22.Enabled = False
     Call MENSERR(24, "No Disponible.\Debe Emitir Nota de Débito Propia.")
    Command22.Enabled = True
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call HELPONLINE("FACTURA")
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    RECUEC07.Text1.TEXT = Text1.TEXT
    RECUEC07.Show 1
    Command4.Enabled = True
End Sub

Sub Command5_Click()
    Command5.Enabled = False
    If DERACCE%("STPFACTU", "Configuración de Funcionamiento Facturacion") > 1 Then OPFACTU07.Show 1
    Call FINAL("")
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   '
   Static AGRUPAENTREGA% 'ASIGNA A AGRUPAENTREGA% SEGUN LA CONFIGURACION DEL SETUP
   If AGRUPAENTREGA% = 0 Then
      AGRUPAENTREGA% = (-1)
      If CONTROL$("FACTUREM", "AGRUPEDI") = "TRANSUCU" Then AGRUPAENTREGA% = 1
   End If
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
   If EsPorPedido& > 0 Then
      TIFADOC% = 2
      MODODOC% = 0
   End If
   '
   FACLOT07.Label15 = "" ' MARCA COMO NO ACTIVO MODADLIDAD DE CARGA POR BOBINAS
   If TIFADOC% = 3 Then   ' nota de debito
     If MODODOC% = 0 Then ' mat codificado
       Call DIFPREC
       Exit Sub
     End If
   End If
   '
   CONSICLI% = 0
   If TIFADOC% = 2 Then    ' 2=FACTURA / 3=N.DEBITO / 6=N.CREDITO
     If MODODOC% = 0 Then  ' 0=MAT.CODIFICADO / 1=TEXTO LIBRE
        '
       If LISTADEMPAQ%(XVALO(Text1)) > 0 Then ' 1 SI HAY LISTA DE EMPAQUE
             OPX% = COMALTER%("Hay Listas de Empaque Pendientes\\Facturar Ahora\Continuar")
             If OPX% = 1 Then
                Call SELEMPAQUE
                Call CARDETFAC
                Exit Sub
             End If
       End If

       Call CARCONSIPEN
       If CONSICLI% > 0 Then
          Call CARDETFAC
          Exit Sub
       End If
       If PREFACTU$ = "REMITO" Then ' POR REMITO Y NO HAY CONSIGNACION PENDIENTE
          Screen.MousePointer = 1
          Text1 = ""
          Call BLANFORMU
          On Error Resume Next
          Text1.SetFocus
          On Error GoTo 0
          Exit Sub
       End If
       ' SELECCIONAR PEDIDOS PENDIENTES
       Screen.MousePointer = 1
       If PREFACTU$ = "ORDESPA" Then
           Call SELPEDIP
'          'VALIDACION PARA MEDITEA (NO HAYA PEDIDOS CON DIF.MOTIVOS Y QUE ESTEN AUTORIZADOS)
'          If VERIPEDMO% < 1 Then GoTo 10
         Else
           Call CARDETPEDIPEN   ' CARGA DETALLE DE PEDIDOS PENDIENTES
       End If
       '
       MONEM1% = MONEMI%: DESCU1$ = "": TRANSPO1% = 0
       MALCOMBINA% = 0: HAYPESEL% = 0
       
       If VERIPEDMO% < 1 Then MALCOMBINA% = 1   ' DISTINTO MOTIVO DE PEDIDO
       '
40     If AGRUPAENTREGA% > 0 Then  ' VALIDA SI ESTA SELECCIONADA LA OPCION DE AGRUPACION DEL SETUP Y PRESENTA WARNING
          For KKU& = 1 To ULTREG&("!CARDEFAC")
            XXU$ = REGLEIDO$("!CARDEFAC", KKU&)
            NPEDXX& = XVALO(TRIM$(Mid$(XXU$, 371, 16)))
            If NPEDXX& > 0 Then HAYPESEL% = 1
            DESCU2$ = TRIM$(DESCUPEDI$(NPEDXX&)): If DESCU2$ = "" Then DESCU2$ = "0"
            If DESCU1$ = "" Then DESCU1$ = DESCU2$
            If DESCU2$ <> DESCU1$ Then
                MALCOMBINA% = 1
            End If
            MONEM2% = MONEPEDI%(NPEDXX&)
            If MONEM1% < 1 Then MONEM1% = MONEM2%
            If MONEM2% <> MONEM1% Then
                MALCOMBINA% = 1
            End If
            ' OT-8-914 - VALIDA QUE TODOS LOS PEDIDOS TENGAN EL MISMO TRANSPORTE
            Call APERBASDAT("PEDCLI")
            TRANSPO2% = XVALO(VALOBADA("PEDENCA", "NTranspo", "NroPed=" & NPEDXX&))
            If TRANSPO1% = 0 And TRANSPO2% <> 0 Then TRANSPO1% = TRANSPO2%
            If TRANSPO1% <> 0 And TRANSPO2% <> 0 And TRANSPO2% <> TRANSPO1% Then
                MALCOMBINA% = 1
            End If
          Next KKU&
          '
       End If
       '
       If MALCOMBINA% > 0 Then
           Call MENSERR(24, "Se han Detectado Pedidos con Diferencias\que Imposibilitan su Agrupamiento\   \Motivos Posibles:\Diferente Motivo, Moneda, Transportista y/o Descuento.\   \Dada la Configuración Seteada, no Podrá Aprobarse la Factura\que Incluya Ordenes de Venta no Agrupables.") ' MENSAJE PEDIDOS MAL COMBINADOS
       End If
       '
       MONEMI% = MONEM1%
       MONEMIS.Caption = Left$(DEMONECO$(MONEMI%), 8)
       TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
       MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
       MONEMIS.ForeColor = RGB(0, 0, 255)
       TICAMBIS.Visible = False
       Label24.Visible = False
       If MONEMI% > 1 Then
          MONEMIS.ForeColor = RGB(255, 0, 0)
          TICAMBIS.Visible = True
          Label24.Visible = True
       End If
       '
       If HAYPESEL% > 0 Then
            PORDESCU = DESCU1$
          Else
            NCXX = XVALO(Text1)
            PORDESCU = DESCUCLI$(NCXX)      ' TOMA EL STANDARD DEL CLIENTE
       End If
       '
       ' OT-8-914 - ASIGNA EL TRANSPORTE DEL PEDIDO
       If TRANSPO1% > 0 Then Text19 = TRANSPO1%
       '
       If EsPorPedido& > 0 Then
         PORDES$ = VALOBADA("PEDENCA", "TDescApli", "NroPed  = " & EsPorPedido&, "NOMESS")
         PORDESCU = PORDES$
       End If
       '
       Call CARDETFAC       ' POR PEDIDO Y SIN CONSIGNACION PENDIENTE.
       '
       Exit Sub
     End If
   End If
   '
   If TIFADOC% = 3 Or TIFADOC% = 6 Then ' débitos o créditos
     If MODODOC% = 1 Then               ' TEXTO LIBRE
       MODOMONEI% = 0
       MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESVENTA", "MODOMONE"))), 8)
       If MODOMONE$ = "MULTIMON" Then
           MODOMONEI% = 2
       End If
       If MODOMONEI% <> 2 Then Call CARDICAMBIO
     End If
   End If
   '
    ' ASIGNACION DE NOTA DE CREDITO A FACTURA
   If TIFADOC% = 6 Then  ' SI ES NOTA DE CREDITO
    '
     OP% = COMALTER%("Nota de Crédito\\N/Crédito Abierta\Anula Factura")
     '
     If OP% <> 2 Then GoTo 110 ' SI ELIGE LIBRE O CANCELA SE VA A CARDETFAC
     If DERACCE%("CREDEVOL", "Nota de Crédito a Factura") >= 2 Then
       Call CARDETCRE(OKX%)
     Else
       OKX% = 0
     End If
     If OKX% < 1 Then Text1 = "": Exit Sub

   End If
   '
110 Call CARDETFAC
   '
End Sub
Sub SELEMPAQUE()
    'ARMO LA LISTA DE SELECCION
    NC% = XVALO(Text1)
    CONDI$ = "NUMLTAEMP > 0 AND CODI_CLI = " & NC% & " AND EmpaCerra > 0 AND  (NumFactura is null or NumFactura  = '') ORDER BY NumLtaEmp"
    Call CARGALISEL("LEMPEVTA", "ENCEMPAQ", "NUMLTAEMP/I5;REFERENCIA/A48", CONDI$, "DISTINCT")
    Call FRESHECHO("LEMPEVTA")
    '
    'PRESENTO LA LISTA Y VALIDO
    Call SELECHO("LEMPEVTA/LISTAS DE EMPAQUE PENDIENTES")
    NLTAEMPAQ& = XVALO(VALACT1("LEMPEVTA"))
    If NLTAEMPAQ& < 1 Then Exit Sub
    '
    'ARMO CONSULTA PARA CARGAR EL CARDETFAC CON EL DETALLE DE LISTA DE EMPAQUE
    SQLX$ = "SELECT CODINT,NROPED,SUM(PIEZASTOTAL)AS TOTALPIEZAS FROM LISEMPAQ WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE NumLtaEmp = " & NLTAEMPAQ&
    SQLX$ = SQLX$ + " GROUP BY CODINT, NROPED"
    'SQLX$ = SQLX$ + " ORDER BY NUMPALLET, NUORDPALLET"
    Dim LISEMPTMP As ADODB.Recordset
    Set LISEMPTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With LISEMPTMP
     Do While Not .EOF
        CI% = !CODINT
        If CI% > 0 And CI% <= NUMAS% Then
          CAN = XVALO(!TOTALPIEZAS)
          NPEDID& = !NROPED
          CONDI$ = "NROPED = " & NPEDID& & " AND CODIINT = " & CI%
          PRECI1# = XVALO(VALOBADA("PEDDETA", "PreUnid", CONDI$, "NOMESS"))
          '
          'CALCULO DE RECARGO POR CONDICION DE PAGO
          RECXPAGO$ = VALOBADA("PEDENCA", "Rec_ConPa", "NROPED = " & NPEDID&, "NOMESS")
          CODESC = COEFDESCU(RECXPAGO$)
          If Abs(XVALO(CODESC)) < 0.05 Then CODESC = 0
          PRECI1# = PRECI1# + (PRECI1# * CODESC / 100)
          '
          'TIPO DE MONEDA
          MONEMA% = XVALO(VALOBADA("PEDDETA", "MoneEmis", CONDI$, "NOMESS"))
          If MONEMA% > ULTREG&("TAMONED") Then MONEMA% = 1 ' ATAJA FUERA DE RANGO
          COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
          '
          'SI NO TIENE PRECIO LE APLICA EL DE LA LISTA
          PRECIO = PRECI1# * COEFCON
            If PRECIO < 0.00005 Then PRECIO = PRELISTA#(ARCHIPRE$, CI%)
          PRELI = XVALO(VALOBADA("PEDDETA", "prelist", CONDI$, "NOMESS"))  ' MODIFICADO PARA QUE TRAIGA EL PRECIO DE LISTA CON EL CUAL SE CREO EL PEDIDO
            If PRELI < PRECIO Then PRELI = PRECIO
          PORDESCUEN = 0: If PRELI >= 0.0001 Then PORDESCUEN = 100 * (PRELI - PRECIO) / PRELI
          '
          'CARGA EL CARDETFAC
5             Z$ = REGBLAN$("!CARDEFAC")
          Call REPLA(Z$, MKS$(CAN), 1, 4)      ' cantidad
            COD$ = CODEXT$(CI%)
          Call REPLA(Z$, MKI$(CI%), 5, 2)
          Call REPLA(Z$, DESCRIT$(CI%), 7, 64)
            NREMIT$ = ""
          Call REPLA(Z$, NREMIT$, 387, 16) ' numero de remito
            FEREMIT% = 0
          '
          Call REPLA(Z$, MKI$(FEREMIT%), 419, 2)
          Call REPLA(Z$, NUPLANO$(CI%), 315, 24)
          '
          Call REPLA(Z$, MKS$(PRELI), 421, 4)
          Call REPLA(Z$, MKS$(PORDESCUEN), 425, 4)
          Call REPLA(Z$, MKS$(PRECIO), 429, 4)
            IMPO# = CDbl(PRECIO) * CAN
          Call REPLA(Z$, MKD$(IMPO#), 433, 8)
          Call REPLA(Z$, MKS$(PORCEIVA(CI%)), 441, 4)
          '
          Call REPLA(Z$, " ", 461, 1)
            '
          NORITPE% = XVALO(VALOBADA("PEDDETA", "NuOrItem", CONDI$, "NOMESS"))
          Call REPLA(Z$, MKI$(NORITPE%), 503, 2)
          Call REPLA(Z$, MKS$(0), 505, 4)
          Call REPLA(Z$, MKS$(0), 509, 4)
          Call REPLA(Z$, TRIM$(Str$(NPEDID&)), 371, 16)
            NUOCOX$ = NUOCOCLI$(NPEDID&, CI%)
10            Call REPLA(Z$, NUOCOX$, 355, 16)
            If TRIM$(NUOCOX$) <> "" Then
              If InStr(Text14, TRIM$(NUOCOX$) + ",") < 1 Then
                Text14 = Text14 + NUOCOX$ + ","
              End If
            End If
            '
          JJ& = JJ& + 1
          Call GRAREG("!CARDEFAC", Z$, JJ&)
          '
        End If
      .MoveNext
     Loop
    End With
    '
    LISEMPTMP.Close
    Set LISEMPTMP = Nothing
    '
    'ASIGNO NUMERO DE LISTA DE EMPAQUE A VARIABLE A NIVEL DE FRM
    FACTLISTAEMPAQUE& = NLTAEMPAQ&
    '

End Sub


'
Sub CARCONSIPEN()
   '
   Static DES%, HAS%
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   '
   If PREFACTU$ = "PEDIDO" Or PREFACTU$ = "ORDESPA" Then
     DES% = 33: HAS% = 32767: GoTo 3
    End If
   '
   If DES% < 1 Then
2    FEX = HOY(HO$)
     HAS% = FEX
     HASS$ = FECHATEX$(FEX)
     DESS$ = "01" + Mid$(HASS$, 3)
     DES% = FECHANUM(DESS$)
   End If
   '
   VDEF$ = MKI$(DES%) + MKI$(HAS%)
   OBX$ = OBJPLA$("DESHASFECHA/REMITOS Emitidos DESDE - HASTA Fecha", VDEF$)
   If OBX$ = "" Then
     CONSICLI% = 0
     Exit Sub
   End If
   '
   DES% = CVI(Left$(OBX$, 2))
   HAS% = CVI(Mid$(OBX$, 3, 2))
   If DES% > HAS% Then
      GoTo 2
   End If
   FEX = DES%: PERIO$ = FECHATEX$(FEX)
   FEX = HAS%: PERIO$ = PERIO$ + " al " + FECHATEX$(FEX)
   '
'ASI ORIGINAL EN ACCESS
'3  DES1# = CDbl(CVDAT(DES%))
3   DESELEC = FETEXCOR1$(DES%)
'   HAS1# = CDbl(CVDAT(HAS%))
    HASELEC = FETEXCOR1$(HAS%)
   
   
   LPRE% = XVALO(TEXT21(0).TEXT)
   NCLIE% = XVALO(Text1.TEXT)
   ARCHIPRE$ = TRIM$(Str$(LPRE%)): LIPRACTI% = LPRE%
   While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
   ARCHIPRE$ = "LIPRE" + ARCHIPRE$
   If LPRE% < 1 Then ARCHIPRE$ = "MASTER"
   CONSICLI% = 0
   '
   VUELTA% = 0
41 REGODES& = 0
   Call SETULTREG("!CARDEFAC", 0)
   If EXISTE%(LUCOM$ + "1CARDEFA.WKF") > 0 Then
     Call SETULTREG("!1CARDEFA", 0)
   End If
   JJ& = 0: CONSICLI% = 0
   '
   CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
   If Len(CMO$) <> 2 Then
     CMO$ = "RT"
   End If
   Screen.MousePointer = 11
   Call APERBASDAT("STOCKS")
   '
   ORIMOVI$ = "MOVISTO": ESREPA$ = " "
   '
4  SQLX$ = " SELECT * FROM " + ORIMOVI$ + " "
   SQLX$ = SQLX$ + "WHERE Nume_Clie = " & NCLIE% & " "
   SQLX$ = SQLX$ + " AND (CodiMovi = '" & CMO$ & "' "
   SQLX$ = SQLX$ + " OR CodiMovi =  'DV' "
   SQLX$ = SQLX$ + " OR CodiMovi =  '*R') "
   SQLX$ = SQLX$ + " AND DoPriLar <> '' "
   SQLX$ = SQLX$ + " AND ((DoSecLar is null) OR (DoSecLar = '') OR (DoSecLar LIKE 'P/F%')) "
   SQLX$ = SQLX$ + " AND FechMov >= '" & DESELEC & "' "
   SQLX$ = SQLX$ + " AND FechMov <= '" & HASELEC & "' "
   '
   Dim CONSITMP As ADODB.Recordset
   Set CONSITMP = New ADODB.Recordset
   CONSITMP.CursorLocation = adUseClient
25 On Error Resume Next
   CONSITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   MODODOSI% = 0: If InStr(EPAC$, "DOSIVA") > 0 Then MODODOSI% = 1
   '
   If MODODOSI% = 1 Then
     If VUELTA% = 0 Then
       Call RECONCLI(NCLIE%, HUBORECU%)
       VUELTA% = 1
       If HUBORECU% > 0 Then
          CONSITMP.Close
          Set CONSITMP = Nothing
          GoTo 41
       End If
     End If
   End If
   
   With CONSITMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       RECONSI$ = REGBLAN("CONSICLI")
       J& = 0
       Do While .EOF = False
         CODIMAT% = XVALO(!COD_INTE)
         CAN = XVALO(!Cantsalid) - XVALO(!CantIngre)
         NPEDID& = XVALO(!NuPedCli)
         If InStr(UCase$(EPAC$), "GUIDI") > 0 Then
            NUMEPED& = 0: PREPED = 0
            If NPEDID& < 1 Then
              FEREMIT% = CVINT%(!fechMov)
              TITUCON$ = TRIM$(SAFETEXT$(!DoPriCor)) + " - Cant: " + TRIM$(Str$(CAN)) + " - Cod: " + TRIM$(CODEXT$(CODIMAT%)) + " - " + DESCRIT$(CODIMAT%)
              Call RECARPED(NCLIE%, CODIMAT%, FEREMIT%, NUMEPED&, PREPED, TITUCON$)
              NPEDID& = NUMEPED&
            End If
         End If
         NREMIT$ = SAFETEXT$(!DoPriCor)
         '
5        Z$ = REGBLAN$("!CARDEFAC")
         Call REPLA(Z$, MKS$(CAN), 1, 4)      ' cantidad
           COD$ = CODEXT$(CODIMAT%)
         Call REPLA(Z$, MKI$(CODIMAT%), 5, 2)
         Call REPLA(Z$, DESCRIT$(CODIMAT%), 7, 64)
         Call REPLA(Z$, NREMIT$, 387, 16) ' numero de remito
           FEREMIT% = CVINT%(!fechMov)
         Call REPLA(Z$, MKI$(FEREMIT%), 419, 2)
         Call REPLA(Z$, NUPLANO$(CODIMAT%), 315, 24)
           '
           If InStr(EPAC$, "GUIDI") > 0 Then GoTo 7 ' para forzar que traiga el precio del remito
           '
           If NPEDID& > 0 Then
             ' busca precio y moneda del pedido
             PRECI1# = PRECIOPED#(NPEDID&, CODIMAT%, MONEMA%)
             If MONEMA% > ULTREG&("TAMONED") Then MONEMA% = 1 ' ATAJA FUERA DE RANGO
             COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
             PRECIO = PRECI1# * COEFCON
             GoTo 8
           End If
           '
7          MONEMA% = XVALO(!MoneVal)
           If MONEMA% <= 1 Then
             MONEPE% = MONEMIPED%(NPEDID&)    ' RECUPERA MONEDA DE PEDIDO
             If MONEPE% > MONEMA% Then MONEMA% = MONEPE%
           End If
           If MONEMA% > ULTREG&("TAMONED") Then MONEMA% = 1 ' ATAJA FUERA DE RANGO
           '
           COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
           PRECIO = XVALO(!VALOUNIT) * COEFCON
             '
8            If PRECIO < 0.00005 Then PRECIO = PREPED
             If PRECIO < 0.00005 Then PRECIO = PRELISTA#(ARCHIPRE$, CODIMAT%)
           PRELI = PRELISTA#(ARCHIPRE$, CODIMAT%)
           If PRELI < PRECIO Then PRELI = PRECIO
           PORDESCUEN = 0: If PRELI >= 0.0001 Then PORDESCUEN = 100 * (PRELI - PRECIO) / PRELI
         Call REPLA(Z$, MKS$(PRELI), 421, 4)
         Call REPLA(Z$, MKS$(PORDESCUEN), 425, 4)
         Call REPLA(Z$, MKS$(PRECIO), 429, 4)
           IMPO# = CDbl(PRECIO) * CAN
         Call REPLA(Z$, MKD$(IMPO#), 433, 8)
         Call REPLA(Z$, MKS$(PORCEIVA(CODIMAT%)), 441, 4)
         Call REPLA(Z$, ESREPA$, 461, 1)    ' PUEDE SER REPARACION
           '
           NORIT% = XVALO(!NuOrItem)
         Call REPLA(Z$, MKI$(NORIT%), 503, 2)
         Call REPLA(Z$, MKS$(0), 505, 4)
         Call REPLA(Z$, MKS$(0), 509, 4)
         Call REPLA(Z$, TRIM$(Str$(NPEDID&)), 371, 16)
           NUOCOX$ = NUOCOCLI$(NPEDID&, CODIMAT%)
10       Call REPLA(Z$, NUOCOX$, 355, 16)
         '
         AA1 = !CodiMovi
         Call REGAPP("!CARDEFAC", Z$)
         CAITCAR% = CAITCAR% + 1
         '
11     .MoveNext
       Loop
12   End If
   End With
   CONSITMP.Close
   Set CONSITMP = Nothing
   '
   If MODODOSI% = 1 Then
      If ORIMOVI$ = "MOVISTO" Then
         ORIMOVI$ = "MOVIREPA": ESREPA$ = "R"
         GoTo 4
      End If
   End If
   '
   JJ& = 0
   For U& = 1 To ULTREG&("!CARDEFAC")
     Z$ = REGLEIDO$("!CARDEFAC", U&)
     If Abs(CVS(Mid$(Z$, 1, 4))) >= 0.01 Then
       If MODODOSI% = 1 Then
          If CVS(Mid$(Z$, 1, 4)) < 0 Then GoTo 19
       End If
       JJ& = JJ& + 1
       Call GRAREG("!CARDEFAC", Z$, JJ&)
     End If
19 Next U&
   Call SETULTREG("!CARDEFAC", JJ&)
   If EXISTE%(LUCOM$ + "1CARDEFA.WKF") > 0 Then
     Call SETULTREG("!1CARDEFA", 0)
     Call CIERRARCH("!1CARDEFA")
   End If
   Call CIERRARCH("!CARDEFAC")
   '
   If PREFACTU$ = "REMITO" Then
     CONSICLI% = 1
     If JJ& < 1 Then
       Call MENSERR(24, "No hay Remitos Pendientes de Facturación\del " + PERIO$ + " para la Cuenta Elegida.")
       CONSICLI% = 0
     End If
     Exit Sub
   End If
   '
   If JJ& > 0 Then
     If COMALTER%("Hay Material en Consignación\Pendiente para este Cliente\\Facturar Consignaciones\Continuar Normalmente") = 1 Then
       CONSICLI% = 1
       Exit Sub
     End If
   End If
   '
   CONSICLI% = 0
   Call SETULTREG("!CARDEFAC", 0)
   Call CIERRARCH("!CARDEFAC")
   If EXISTE%(LUCOM$ + "1CARDEFA.WKF") > 0 Then
     Call SETULTREG("!1CARDEFA", 0)
     Call CIERRARCH("!1CARDEFA")
   End If
   '
End Sub
'
Sub CARCONSIPED(NPED&)
   '
   Static DES%, HAS%
   '
   DES% = 0: HAS% = 32767
   '
3  DES1# = CDbl(CVDAT(DES%))
   HAS1# = CDbl(CVDAT(HAS%))
   '
   LPRE% = XVALO(TEXT21(0).TEXT)
   NCLIE% = XVALO(Text1.TEXT)
   ARCHIPRE$ = TRIM$(Str$(LPRE%)): LIPRACTI% = LPRE%
   While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
   ARCHIPRE$ = "LIPRE" + ARCHIPRE$
   If LPRE% < 1 Then ARCHIPRE$ = "MASTER"
   CONSICLI% = 0
   '
   REGODES& = 0
   Call SETULTREG("!CARDEFAC", 0)
   If EXISTE%(LUCOM$ + "1CARDEFA.WKF") > 0 Then
     Call SETULTREG("!1CARDEFA", 0)
   End If
   JJ& = 0: CONSICLI% = 0
   '
   CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
   If Len(CMO$) <> 2 Then
     CMO$ = "RT"
   End If
   Screen.MousePointer = 11
   Call APERBASDAT("STOCKS")
   '
   SQLX$ = " SELECT * FROM MOVISTO "
   SQLX$ = SQLX$ + "WHERE Nume_Clie = " & NCLIE% & " "
   SQLX$ = SQLX$ + " AND (CodiMovi = '" & CMO$ & "' "
   SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
   SQLX$ = SQLX$ + " AND DoPriLar <> '' "
   SQLX$ = SQLX$ + " AND (DoSECLar = '' "
   SQLX$ = SQLX$ + " OR DoSecLar LIKE 'P/F%' ) "
   SQLX$ = SQLX$ + " AND NuPedCli = " & NPED& & " "
   '
   On Error GoTo 0
   'Set CONSITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim CONSITMP As ADODB.Recordset
   Set CONSITMP = New ADODB.Recordset
25 On Error Resume Next
   CONSITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   With CONSITMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       
       On Error GoTo 0
       RECONSI$ = REGBLAN("CONSICLI")
       J& = 0
       Do While .EOF = False
         CODIMAT% = !COD_INTE
         CAN = !Cantsalid - !CantIngre
         NPEDID& = 0: If IsNull(!NuPedCli) = False Then NPEDID& = !NuPedCli
         NUMEPED& = 0: PREPED = 0
           If NPEDID& < 1 Then
             FEREMIT% = CVINT%(!fechMov)
             TITUCON$ = TRIM$(!DoPriCor) + " - Cant: " + TRIM$(Str$(CAN)) + " - Cod: " + TRIM$(CODEXT$(CODIMAT%)) + " - " + DESCRIT$(CODIMAT%)
             Call RECARPED(NCLIE%, CODIMAT%, FEREMIT%, NUMEPED&, PREPED, TITUCON$)
             NPEDID& = NUMEPED&
           End If
5        Z$ = REGBLAN$("!CARDEFAC")
         Call REPLA(Z$, MKS$(CAN), 1, 4)      ' cantidad
           COD$ = CODEXT$(CODIMAT%)
         Call REPLA(Z$, MKI$(CODIMAT%), 5, 2)
         Call REPLA(Z$, DESCRIT$(CODIMAT%), 7, 64)
           NREMIT$ = !DoPriCor
         Call REPLA(Z$, NREMIT$, 387, 16) ' numero de remito
           FEREMIT% = CVINT%(!fechMov)
         Call REPLA(Z$, MKI$(FEREMIT%), 419, 2)
         Call REPLA(Z$, NUPLANO$(CODIMAT%), 315, 24)
           '
           MONEMA% = !MoneVal
           If MONEMA% <= 1 Then
             MONEPE% = MONEMIPED%(NPEDID&)    ' RECUPERA MONEDA DE PEDIDO
             If MONEPE% > MONEMA% Then MONEMA% = MONEPE%
           End If
           '
           COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
           PRECIO = !VALOUNIT * COEFCON
             If PRECIO < 0.00005 Then PRECIO = PREPED
             If PRECIO < 0.00005 Then PRECIO = PRELISTA#(ARCHIPRE$, CODIMAT%)
           PRELI = PRELISTA#(ARCHIPRE$, CODIMAT%)
           If PRELI < PRECIO Then PRELI = PRECIO
           PORDESCUEN = 0: If PRELI >= 0.0001 Then PORDESCUEN = 100 * (PRELI - PRECIO) / PRELI
         Call REPLA(Z$, MKS$(PRELI), 421, 4)
         Call REPLA(Z$, MKS$(PORDESCUEN), 425, 4)
         Call REPLA(Z$, MKS$(PRECIO), 429, 4)
           IMPO# = CDbl(PRECIO) * CAN
         Call REPLA(Z$, MKD$(IMPO#), 433, 8)
         Call REPLA(Z$, MKS$(PORCEIVA(CODIMAT%)), 441, 4)
         Call REPLA(Z$, " ", 461, 1)     ' NUNCA ES REPARACION
           '
           NORIT% = !NuOrItem
         Call REPLA(Z$, MKI$(NORIT%), 503, 2)
         Call REPLA(Z$, MKS$(0), 505, 4)
         Call REPLA(Z$, MKS$(0), 509, 4)
         Call REPLA(Z$, TRIM$(Str$(NPEDID&)), 371, 16)
           NUOCOX$ = NUOCOCLI$(NPEDID&, CODIMAT%)
10       Call REPLA(Z$, NUOCOX$, 355, 16)
         '
         Call REGAPP("!CARDEFAC", Z$)
         CAITCAR% = CAITCAR% + 1
         '
       .MoveNext
       Loop
12   End If
     On Error GoTo 0
   End With
   '
   JJ& = 0
   For U& = 1 To ULTREG&("!CARDEFAC")
     Z$ = REGLEIDO$("!CARDEFAC", U&)
     If Abs(CVS(Mid$(Z$, 1, 4))) >= 0.01 Then
       JJ& = JJ& + 1
       Call GRAREG("!CARDEFAC", Z$, JJ&)
     End If
   Next U&
   Call SETULTREG("!CARDEFAC", JJ&)
   If EXISTE%(LUCOM$ + "1CARDEFA.WKF") > 0 Then
     Call SETULTREG("!1CARDEFA", 0)
     Call CIERRARCH("!1CARDEFA")
   End If
   Call CIERRARCH("!CARDEFAC")
   '
   If PREFACTU$ = "REMITO" Then
     CONSICLI% = 1
     If JJ& < 1 Then
       Call MENSERR(24, "No hay Remitos Pendientes de Facturación\del " + PERIO$ + " para la Cuenta Elegida.")
       CONSICLI% = 0
     End If
     Exit Sub
   End If
   '
   If JJ& > 0 Then
     If COMALTER%("Hay Material en Consignación\Pendiente para este Cliente\\Facturar Consignaciones\Continuar Normalmente") = 1 Then
       CONSICLI% = 1
       Exit Sub
     End If
   End If
   '
   CONSICLI% = 0
   Call SETULTREG("!CARDEFAC", 0)
   Call CIERRARCH("!CARDEFAC")
   If EXISTE%(LUCOM$ + "1CARDEFA.WKF") > 0 Then
     Call SETULTREG("!1CARDEFA", 0)
     Call CIERRARCH("!1CARDEFA")
   End If
   '
End Sub
'
Sub RECARPED(NCLIE%, CIPD%, FEREMIT%, NUMEPED&, PREPED, TITUCON$)
   '
   'FEREMIX = CDbl(CVDAT(FEREMIT%))
   FEREMIX = FECHATEX$(FEREMIT%)
   
   
   Call APERBASDAT("PEDCLI")
   SQLX$ = " SELECT * "
   SQLX$ = SQLX$ + " FROM PEDDETA "
   SQLX$ = SQLX$ + " WHERE NroClie = " & NCLIE% & " "
   SQLX$ = SQLX$ + " AND Status < 8 "    ' NO ANULADOS NI CUMPLIDOS
   SQLX$ = SQLX$ + " AND STATUS >= 0 "   ' APROBADOS
   SQLX$ = SQLX$ + "AND CodiInt = " & CIPD% & " "
   'SQLX$ = SQLX$ + "AND CDBL(INT(FechEmis)) <= " & FEREMIX & " "
   'SQLX$ = SQLX$ + " AND (((CanPed - CantEnt) >= 0.05) OR (TipoPed = 'ABIERTO' AND CDBL(FinVigen) >= " & FEREMIX & ")) "
   SQLX$ = SQLX$ + " AND (((CanPed - CantEnt) >= 0.05) OR (TipoPed = 'ABIERTO' AND FinVigen >= '" & FEREMIX & "'))"
'   SQLX$ = SQLX$ + " AND FinVigen >= '" & FEREMIX & "'"
   SQLX$ = SQLX$ + " ORDER BY NroPed ASC"
   '
   'Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim PEDTEMP As ADODB.Recordset
   Set PEDTEMP = New ADODB.Recordset
   PEDTEMP.CursorLocation = adUseClient
25 On Error Resume Next
   PEDTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   NUMEPED& = 0
   PREPED = 0
   ENCUPE% = 0
   '
   With PEDTEMP
     On Error Resume Next
     .MoveLast
     CAREGIS& = .RecordCount
     '
     If CAREGIS& = 1 Then
         NUMEPED& = !NROPED
         PREPED = !PreUnid
         ENCUPE% = 1
       ElseIf CAREGIS& > 1 Then
         .MoveFirst
         JJ& = 0
         Do While (.EOF = False)
           TTXX$ = MKS$(!NROPED)
           TTXX$ = TTXX$ + AJUSTI$(!NroOcom, 16) + "  "
              FEX = CVINT(!FechEmis)
           TTXX$ = TTXX$ + FECHATEX$(FEX) + "  "
              FEX = CVINT(!FeSolEnt)
           TTXX$ = TTXX$ + FECHATEX$(FEX) + "  "
           TTXX$ = TTXX$ + FORMATNUM$(!PreUnid, "F12.4")
              CAPEN = !CanPed - !CantEnt
              If CAPEN < 0 Then CAPEN = 0
              If !TipoPed = "ABIERTO" Then CAPEN = 999999
           TTXX$ = TTXX$ + FORMATNUM$(CAPEN, "F10.1")
           TTXX$ = TTXX$ + FORMATNUM(!NuOrItem, "F6.0")
           JJ& = JJ& + 1
           Call GRAREG("!SELECPE", TTXX$, JJ&)
         .MoveNext
         Loop
         '
         Call SETULTREG("!SELECPE", JJ&)
         '
20       Call FRESHECHO("!SELECPE")
         TTXX$ = REGSEL$("!SELECPE/" + TITUCON$)
         If Len(TTXX$) <= 4 Then GoTo 20
         '
         NUMEPED& = CVS(Left$(TTXX$, 4))
         PREPED = XVALO(Mid$(TTXX$, 45, 12))
         ENCUPE% = 1
      End If
    End With
    PEDTEMP.Close
    Set PEDTEMP = Nothing
    '
End Sub
'
Sub CARDICAMBIO()
   '
   DOCUDIF = ""
   MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESVENTA", "MODOMONE"))), 8)
   If MODOMONE$ <> "BIMONETA" Then
      If MODOMONE$ <> "MULTIMON" Then
         Exit Sub
      End If
   End If
   '
   Dim DOCABX$(2048), SALPES#(2048), SALDOL#(2048), REDOCA&(2048)
   '
   NC% = XVALO(Text1)
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
           Call REPLA(Z$, "u$s" + FORMATNUM$(SALDOL#(KKI%), "F12.2"), 42, 16)
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
           Call REPLA(Z$, "u$s" + FORMATNUM$(SALDOL#(KKI%), "F12.2"), 42, 16)
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
   IDIFE = XVALO(Mid$(BBXX$, 12, 11))
   If TIFADOC% = 3 Then
     IDIFE = XVALO(Mid$(BBXX$, 28, 12)) * TICAMONE(2)
   End If
   If PIVADOC% = 1 Then
     TASA1 = XVALO(PORIVA(0).Caption)
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
    MODX% = ALTERNA%("Listas de Precios\Condiciones de Pago\Alícuotas de I.V.A.")
    If MODX% = 1 Then
         Call CONECRUN("LPREAC07", "Listas de Precios de Venta")
       ElseIf MODX% = 2 Then
         Call CARCONPAGOS
       ElseIf MODX% = 3 Then
         Call CARTASIVA
         Call COMUNI("Re-Inicie Ahora la Aplicación.")
         Call CIERRARCH("*.*")
         Call FINAL("")
    End If
    Command7.Enabled = True
End Sub

Sub Command8_Click()
    '
    Dim RESP%
    Dim NUMEOCXX$
    '
    Command8.Enabled = False
    MODO% = COMALTER%("Opciones de Salida\\Pantalla\Impresora\Cancelar")
    If MODO% < 1 Or MODO% > 2 Then GoTo 99
    '
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
      CMO$ = "RT"
    End If
10  FEX = HOY(HO$)
    HAS% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    VDEF$ = MKI$(DES%) + MKI$(HAS%)
    '
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        GoTo 10
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    'jandy sql
    'DES1# = CDbl(CVDAT(DES%))
    'HAS1# = CDbl(CVDAT(HAS%))
    DES1$ = FETEXCOR1$(DES%)
    HAS1$ = FETEXCOR1$(HAS%)
    
    '
    If MODO% = 1 Then GoTo 56
    '////////******NUEVO FILTRO CLIENTE*******///////
    NCMIN% = 32767: NCMAX% = 0
    FIN& = ULTREG&("DEUDOR1")
    For U& = 1 To FIN&
        NCX% = CVI(Left$(REGLEIDO$("DEUDOR1", U&), 2))
        If NCX% > 0 Then
            If NCX% > NCMAX% Then NCMAX% = NCX%
            If NCX% < NCMIN% Then NCMIN% = NCX%
        End If
    Next U&
    '
    With RACLIRPF
      .Text10 = TRIM$(Str$(NCMIN%))
      .Text9 = RASOCLI$(NCMIN%)
      .Text8 = TRIM$(Str$(NCMAX%))
      .Text7 = RASOCLI$(NCMAX%)
      .Show 1
    End With
    '
    'SI CANCELA EN EL FORMULARIO DE RANGO DE CTAS. DE CLIENTES
    If TRIM$(RACLIRPF.Text10.TEXT) = "" And TRIM$(RACLIRPF.Text8.TEXT) = "" Then
        Command8.Enabled = True
        Exit Sub
    End If
    '
    CLINFE = XVALO(RACLIRPF.Text10)
    CLISUP = XVALO(RACLIRPF.Text8)
    '//////********HASTA ACA NUEVO FILTRO CLIENTE****///////
56  Screen.MousePointer = 11
    Call APERBASDAT("STOCKS")
    '
    'JANDY SQL
    
    Dim CONSITMP As ADODB.Recordset
    Set CONSITMP = New ADODB.Recordset
    CONSITMP.CursorLocation = adUseClient
    
    SQLX$ = " SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE (CodiMovi = '" & CMO$ & "' "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
    SQLX$ = SQLX$ + " AND DoPriLar <> '' "
    SQLX$ = SQLX$ + " AND ((DoSecLar IS NULL) or (DoSecLar = '') OR (DoSecLar LIKE 'P/F%')) "
    SQLX$ = SQLX$ + " AND FechMov >= '" & DES1$ & "' "
    SQLX$ = SQLX$ + " AND FechMov <= '" & HAS1$ & "' "
'
'    SQLX$ = SQLX$ + " AND ((ISNULL(DoSecLar) = true) or (DoSecLar = '') OR (DoSecLar LIKE 'P/F*')) "
'    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) >= " & DES1# & " "
'    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) <= " & HAS1# & " "
'

    'PARA EVITAR COMPROBANTES ARRUINADOS
    'SQLX$ = SQLX$ + " AND LEFT$(UCASE$(REFERCOR),4) <> 'ANUL' "
    SQLX$ = SQLX$ + "AND REFERCOR NOT LIKE 'ANUL%'"
    'FILTRO CLIENTE
    If MODO% <> 1 Then
      SQLX$ = SQLX$ + " AND Nume_Clie >= " & CLINFE & ""
      SQLX$ = SQLX$ + "AND Nume_Clie<= " & CLISUP & ""
    End If
    '
    'Set CONSITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
25  On Error Resume Next
    CONSITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0
    With CONSITMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, "No Existen Remitos Pendientes de Facturación.")  'Generados para este Cliente")
        On Error GoTo 0
        Screen.MousePointer = 1
        Command8.Enabled = True
        Exit Sub
      End If
      On Error GoTo 0
      '
      RECONSI$ = REGBLAN("CONSICLI")
      RECONSJ$ = REGBLAN("REPENFAC")
      J& = 0
      Do While .EOF = False
         NCLIE% = !Nume_Clie
         Call REPLA(RECONSI$, MKI$(NCLIE%), 1, 2)
         Call REPLA(RECONSI$, MKI$(CVINT(!fechMov)), 3, 2)
         Call REPLA(RECONSI$, "", 5, 18)
         Call REPLA(RECONSI$, !DoPriCor, 23, 12)
         Call REPLA(RECONSI$, MKI$(!COD_INTE), 35, 2)
         Call REPLA(RECONSI$, MKD$(!Cantsalid), 37, 8)
         Call REPLA(RECONSI$, MKD$(!CantIngre), 45, 8)
         Call REPLA(RECONSI$, MKD$(!CantSaldo), 53, 8)
         Call REPLA(RECONSI$, !ReferCor, 61, 48)
         Call REPLA(RECONSI$, String$(12, 0), 109, 12)
         Call REPLA(RECONSI$, MKS$(XVALO(Mid(!DoPriCor, 5))), 109, 4)
         Call REPLA(RECONSI$, MKI$(CVINT(!FeReMovi)), 122, 2)
          HOR = Format(!FeReMovi, "HH:MM:SS")
         Call REPLA(RECONSI$, MKI$(CVINT(HOR)), 124, 2)
         Call REPLA(RECONSI$, MKI$(XVALO(!NumUsuar)), 127, 2)
         '
         Call REPLA(RECONSJ$, MKI$(NCLIE%), 1, 2)
         Call REPLA(RECONSJ$, MKI$(!COD_INTE), 3, 2)
         Call REPLA(RECONSJ$, !DoPriCor, 5, 12)
         Call REPLA(RECONSJ$, !DOPRILAR, 17, 18)
         Call REPLA(RECONSJ$, MKI$(CVINT(!fechMov)), 35, 2)
         If IsNull(!NuPedCli) Then
            Call REPLA(RECONSJ$, MKS$(0), 37, 4)
         Else
            Call REPLA(RECONSJ$, MKS$(!NuPedCli), 37, 4)
         End If
         If IsNull(!NuPedCli) Then
             NUMEOCXX$ = ""
           Else
             NUMEOCXX$ = NUOCOCLI$(!NuPedCli, !COD_INTE)
         End If
         Call REPLA(RECONSJ$, NUMEOCXX$, 41, 24)
         Call REPLA(RECONSJ$, MKD$(!Cantsalid), 65, 8)
         Call REPLA(RECONSJ$, MKD$(!CantIngre), 73, 8)
         Call REPLA(RECONSJ$, MKD$(!CantSaldo), 81, 8)
         Call REPLA(RECONSJ$, MKD$(!VALOUNIT), 89, 8)
           TOTI# = !VALOUNIT * (!Cantsalid - !CantIngre)
         Call REPLA(RECONSJ$, MKD$(TOTI#), 97, 8)
         '
         J& = J& + 1
         Call GRAREG("CONSICLI", RECONSI$, J&)
         Call GRAREG("REPENFAC", RECONSJ$, J&)
      .MoveNext
      Loop
    End With
    CONSITMP.Close
    Set CONSITMP = Nothing
    '
    Call SETULTREG("CONSICLI", J&)
    Call FILESORT("CONSICLI", "", 1, 6, "ASC")
    Call CIERRARCH("CONSICLI")
    '
    Call SETULTREG("REPENFAC", J&)
    Call FILESORT("REPENFAC", "", 1, 6, "ASC")
    Call CIERRARCH("REPENFAC")
    '
    If MODO% = 1 Then
        Call CONECRUN("*CREPENFA", "Remitos Sin Facturar (Clientes)")
      Else
        RESP% = COMALTER%("Opciones de Agrupamiento\\Comprobante\Articulo\Cancelar")
        If RESP% = 1 Then
            Call CONECRUN("*LIREPEFA", "Remitos Sin Facturar (Clientes)")
          ElseIf RESP% = 2 Then
            Call HEADERS("REPENFAC", "")
            If CLINFE = CLISUP Then
                NCXX% = CLINFE
                Call HEADERS("REPENFAC", "Cliente: " + RASOCLI$(NCXX%))
                Call CONECRUN("*LREPENFB", "Remitos Sin Facturar (Clientes)")
              Else
                Call CONECRUN("*LREPENFA", "Remitos Sin Facturar (Clientes)")
            End If
        End If
    End If
    '
99  Command8.Enabled = True
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
    Screen.MousePointer = 11
    Call CENTRAFORM(Me)
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Label28 = ""
    Frame5.Width = 4500
    Eco21(0).Width = 2450
    Eco21(1).Width = 2450

    If CONTROL$("ADMIPED", "MOCOMIS") = "LETRA" Then
        Label28.Visible = True
        Text22.Visible = False
      ElseIf CONTROL$("ADMIPED", "MOCOMIS") = "PORCIEN" Then
        Label28.Visible = False
        Text22.Visible = True
    End If
    '
    Refresh
    Frame1.Refresh
    DoEvents
    '\\\OT-05-0419-WAR-23/06/05///
    If CONTROL$("FACTURA", "SUPCONF") <> "SI" Then
       Call CARGAPUERTOS
    End If
    '
    If EXISTE%("PEDIDO07.EXE") > 0 Then
       Command17.Visible = True
    End If
    '
    HOII% = HOY(HOS$)
    Text9.TEXT = HOS$: DATEDOC% = HOII%
    '
    MONEMI% = 1  ' PESOS
    TICAMBIS = "1.0000"
    TICAMBIS.Visible = False
    Label24.Visible = False
    CONVERPRECIO% = 1
    '
    Call SETULTREG("!CARDEFAC", 0)
    If EXISTE%(LUCOM$ + "1CARDEFA.WKF") > 0 Then
      On Error Resume Next
      Kill LUCOM$ + "1CARDEFA.WKF"
      On Error GoTo 0
    End If
    '
    TIFADOC% = 2: Label7.Caption = "Factura:"
    Label6.Caption = "(Mat.Codif)": MODODOC% = 0
    Label13.Caption = "A": PIVADOC% = 1: VARIDOC% = 1
'    UPREFI$ = TRIM$(CONTROL$("FACTURA", "PREFIFA"))
    CLAVE$ = "PREFAC-" & Label13.Caption '  MODIFICADO PARA GUARDAR PREFIJO DE FACTURA POR ESTACION
    UPREFI$ = TRIM$(CONTROL$(CLAVE$, IDENTER$))
    If InStr(EPAC$, "DOSIVA") > 0 Then UPREFI$ = "0001"
    '
    If UPREFI$ = "" Then UPREFI$ = "0001"
    Label12.Caption = UPREFI$
    '
    Printer.TrackDefault = False
    '
    VERLICRE% = 0
    If CONTROL$("FACTURA", "VERLICRE") = "SI" Then VERLICRE% = 1
    Call LEETASIVA
    'Call NUFACTURA ' TRASLADADO MAS ABAJO PARA QUE ESTE INICIALIZADO LA I.FISCAL
    '
    MODESCU% = 0
    If CONTROL$("FACTUREM", "MODESCUE") = "PORCIENTO" Then
       MODESCU% = 1
    End If
    '
    ' TIPO DE CAMBIO POR OMISION
    TICAMBX# = 1
    MONEMI% = 1   ' pesos
    TICAMBIS = "1.0000"
    TICAMBIS.Visible = False
    Label24.Visible = False
    '
    If EsPorPedido& > 0 Then PREFACTU$ = "ORDESPA"
    '
    If PREFACTU$ = "" Then
      PREFACTU$ = TRIM$(CONTROL("FACTURA", "PREFACTU"))
      If PREFACTU$ = "" Then PREFACTU$ = "REMITO"   ' si no ORDESPA / PEDIDO
    End If
    '
    FACELEC% = 0
    If CONTROFIS% = 0 Then
       CONTROFIS% = (-1)
       Set FORFACPRINT = FORFASQL09
       If CONTROL$("FACTURA", "SALIDA") = "CFISCAL" Then
          CONTROFIS% = 1
          FORCARGA$ = "SINREMITO"
          Check1.Visible = True
          ' INICIALIZAR CONTROLADOR FISCAL
          Call FORHASAR.INICIA
          '
       ElseIf CONTROL$("FACTURA", "SALIDA") = "IFISCAL" Then
          CONTROFIS% = 2
          Check1.Visible = True
          ' INICIALIZAR CONTROLADOR FISCAL
          Call FORHAS320.INICIA
          On Error Resume Next
          Call FORHAS320.HASAR1.ObtenerDatosDeInicializacion
          YY$ = FORHAS320.HASAR1.Respuesta(0)
          On Error GoTo 0
          PUNTODEVENTA$ = TRIM$(Mid$(YY$, 92, 4))
          Call COMUNI("El Punto de Venta de la Impresora Fiscal es '" + PUNTODEVENTA$ + "'")
          '
       ElseIf CONTROL$("FACTURA", "SALIDA") = "FELECT" Then ' FACTURA ELECTRONICA
          '
          CONTROFIS% = 0
          FACELEC% = 1
          OPVARFA07.Command5.Visible = True
          OPVARFA07.Command7.Visible = True
          '
          ' si no tiene configurado punto de venta normal
          If XVALO(CONTROL("FACTURA", "PUVENELE")) < 1 Then GoTo 40
          '
          VUELTA% = 0
          CARPETOK$ = "C:\FLEXOFT\FACELEC\"
          ARCHIVO_TOKEN$ = "TOKEN.TXT"
30        If EXISTE%(CARPETOK$ + ARCHIVO_TOKEN$) = 1 Then
             FechArch = FileSystem.FileDateTime(CARPETOK$ + ARCHIVO_TOKEN$)
             FECHARCHIVO% = CVINT(FechArch)
             FECHHOY% = HOY(HO$)
             If FECHARCHIVO% < FECHHOY% Then
                Screen.MousePointer = 1
                Call MENSERR(24, "Token Vencido. Vuelva a Generarlo")
                OPVARFA07.Show 1
                Call FINAL("")
             End If
             'Timer2.Enabled = True
          Else
             ' Verifica la existencia del token por cualquiera de las 3 modalidades
             ' Factura comun, bienes de capital o exportacion
'             If VUELTA% = 0 Then
'                ARCHIVO_TOKEN$ = "TOKEN_WSBFE.TXT": VUELTA% = 1: GoTo 30
'             End If
'             If VUELTA% = 1 Then
'                ARCHIVO_TOKEN$ = "TOKEN_WSFEX.TXT": VUELTA% = 2: GoTo 30
'             End If
             Screen.MousePointer = 1
             Call MENSERR(24, "Debe Generarse el Archivo 'TOKEN.TXT'\al Ingresar a Facturacion.")
             OPVARFA07.Show 1
             Call FINAL("")
          End If
       End If
    End If
    '
40  Call NUFACTURA    ' SE TRASLADÓ AQUÍ PARA QUE YA ESTÉ INICIALIZADO LA IMPRESORA FISCAL
    '
    If EsPorPedido& > 0 Then
      NC% = XVALO(VALOBADA("PEDENCA", "NroClie", "NroPed = " & EsPorPedido& & "", "NOMESS"))
      Text1 = NC%
    End If

    Screen.MousePointer = 1
    '
End Sub
'
Sub CARGAPUERTOS()
    '
    Printer.TrackDefault = False
    PORTUFAC$ = TRIM$(CONTROL$(IDENTER$, "PORTFAC"))
      If PORTUFAC$ = "" Then PORTUFAC$ = CONTROL$("*", "PORTFAC")
      If PORTUFAC$ = "" Then GoTo 20
      NEWPORT$ = PORTUFAC$    ' + "/NOMESS"
      Call SETPRINPORT(NEWPORT$)
      If NEWPORT$ = "" Then GoTo 20
      Call SETPRINPORT("RESTORE")
    PORTUREM$ = TRIM$(CONTROL$(IDENTER$, "PORTREM"))
      If PORTUREM$ = "" Then PORTUREM$ = CONTROL$("*", "PORTREM")
      If PORTUREM$ = "" Then GoTo 20
      NEWPORT$ = PORTUREM$   ' + "/NOMESS"
      Call SETPRINPORT(NEWPORT$)
      If NEWPORT$ = "" Then GoTo 20
      Call SETPRINPORT("RESTORE")
    Exit Sub    ' PUDO ACCEDER A LOS PUERTOS DE IMPRESIóN
    '
20  Screen.MousePointer = 1
    DIREPRIN.Show 1
    Call FRESHALL
    DoEvents
    If DIREPRIN.Label3 <> "OK" Then
      Call FINAL("")
    End If
    '
    PORTUFAC$ = DIREPRIN.Combo1.TEXT
    PORTUREM$ = DIREPRIN.Combo2.TEXT
    '
    If PORTUFAC$ <> "" Then
      Call SETPRINPORT(PORTUFAC$)
      Call SETPRINPORT("RESTORE")
    End If
    
    If PORTUREM$ <> "" Then
      Call SETPRINPORT(PORTUREM$)
      Call SETPRINPORT("RESTORE")
    End If
    
    Call GRACONTROL("*", "PORTFAC", PORTUFAC$)
    'If TRIM$(CONTROL$(IDENTER$, "PORTFAC")) <> "" Then
      Call GRACONTROL(IDENTER$, "PORTFAC", PORTUFAC$)
    'End If
    Call GRACONTROL("*", "PORTREM", PORTUREM$)
    'If TRIM$(CONTROL$(IDENTER$, "PORTREM")) <> "" Then
      Call GRACONTROL(IDENTER$, "PORTREM", PORTUREM$)
    'End If
    '
End Sub
'
Private Sub Form_Resize()
   If WindowState = 1 Then
     Call BAJALDISCO
   End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Label19_DblClick()
   Static CTXX%
   '
   If ULTREG&("DAPIBCLI") < 1 Then
     Call MENSERR(24, "Imposible desde esta Aplicación")
     Exit Sub
   End If
   '
   NCLIE% = XVALO(TRIM$(Text1.TEXT))
   If NCLIE% > 0 Then
     If RASOCLI$(NCLIE%) <> "" Then
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
         '
       End If
     End If
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
       End If
     End If
   End If
End Sub

Private Sub Label6_Click()
   If ULTREG&("!CARDEFAC") > 0 Then Exit Sub
   'If TIFADOC% = 2 Or TIFADOC% = 6 Then
      If MODODOC% = 0 Then
          Label6.Caption = "(T.Libre)": MODODOC% = 1
        ElseIf MODODOC% = 1 Then
          Label6.Caption = "(Mat.Codif)": MODODOC% = 0
      End If
   'End If
   '
   NCLIE% = XVALO(Text1.TEXT)
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
       TIFADOC% = 3: Label7.Caption = "N.Débito:"
       Label6.Caption = "(T.Libre)": MODODOC% = 1
       '
     ElseIf Label7.Caption = "N.Débito:" Then
       TIFADOC% = 6: Label7.Caption = "N.Crédito:"
       Label6.Caption = "(Mat.Codif)": MODODOC% = 0
     Else
       TIFADOC% = 2: Label7.Caption = "Factura:"
       Label6.Caption = "(Mat.Codif)": MODODOC% = 0
   End If
   Call NUFACTURA
   '
   NCLIE% = XVALO(Text1.TEXT)
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
   If Command14.Enabled = True Then
   
     If FACLOT07.Label15 = "ACTIVO" Then
        FACLOT07.Show 1  'CARGA POR BOBINA
     Else
        Call CARDETFAC
     End If
   End If
   
End Sub

Private Sub List3_DblClick()
   Call DIFPREC2
End Sub

Private Sub MONEMIS_DblClick()
   '
   Static CAMON%, DMONE$(), TCMONE()
   If CAMON% < 1 Then
     CAMON% = ULTREG&("TAMONED")
10   If CAMON% < 1 Then
         CAMON% = 1
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         DMONE$(1) = "Pesos"
         TCMONE(1) = 1
       Else
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         U2& = 0
         For U& = 1 To CAMON%
           XX$ = REGLEIDO$("TAMONED", U&)
           DMX$ = TRIM$(Left$(XX$, 8)) + " "
           PPXX% = InStr(DMX$, " ")
           DMX$ = Left$(DMX$, PPXX% - 1)
           '
           If TRIM$(DMX$) <> "" Then
             VHX = CVS(Mid$(XX$, 29, 4)): If VHX < 0.005 Then VHX = 1
             U2& = U2& + 1
             DMONE$(U2&) = DMX$
             TCMONE(U2&) = VHX
           End If
         Next U&
         CAMON% = U2&
         If CAMON% < 1 Then GoTo 10
     End If
   End If
   '
   MONEMIA% = MONEMI%
   MONEMI% = MONEMI% + 1
   If MONEMI% > CAMON% Then MONEMI% = 1
   MONEMIS = DMONE$(MONEMI%)
'   COEFICO = TICAMONE(MONEMIA%) / TICAMONE(MONEMI%)
   MONEMIS.ForeColor = RGB(0, 0, 255)
   If MONEMI% > 1 Then
      MONEMIS.ForeColor = RGB(255, 0, 0)
   End If
   MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
   TICAMBIS.Visible = False: Label24.Visible = False
   If MONEMI% > 1 Then
     TICAMBIS.Visible = True: Label24.Visible = True
   End If
   '
   COEFCON = TICAMONE(MONEMIA%) / TICAMONE(MONEMI%)
   TICAMBX# = TICAMONE(MONEMI%)
   TICAMBIS = FORMATNUM(TICAMBX#, "F9.4")
   FIN& = ULTREG&("!CARDEFAC")
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
End Sub

'Private Sub MONEMIS_dblclick()
'   If TRIM$(UCase$(MONEMIS)) = "PESOS" Then
'      MONEMIS = "Dólares"
'     Else
'       MONEMIS = "Pesos"
'   End If
'End Sub

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
   '
   If DERACCE%("MODESFA/NOMESS", "Modificacion descuento de Factura") < 2 Then
       Text14.SetFocus
       If List1.ListCount > 0 Then AA1% = DERACCE%("MODESFA", "Modificacion descuento de Factura")
   End If
   '
End Sub

Private Sub PORIVA_Click(Index As Integer)
   PORDESCU.SetFocus
End Sub

Private Sub Text1_Change()
   Call BLANFORMU
   CLIPROACTIVO% = XVALO(Text1)
   '
   If EsPorPedido& > 0 Then
       Call Text1_KeyPress(13)
       Call Command6_Click
   End If
   '
End Sub
'
Private Sub Text1_DblClick()
  Call SELECHO("DEUDOR1")
  NCLIE% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
  If NCLIE% > 0 Then
    Text1.TEXT = TRIM$(Str$(NCLIE%))
    If XVALO(Text1.TEXT) > 0 Then
      If XVALO(Text1.TEXT) <= 32767 Then
        NC% = XVALO(Text1)
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

Private Sub TEXT11_DBLCLICK()
    'NUEVO PARA FACTURAR CON OTRO NUMERO
    If VALICLIE% = 1 Then
      If DERACCE%("FANULI", "Cambiar Manualmente Número de Factura") >= 2 Then
        '
10      NUF = InputBox("Ingrese el Número de Factura", "Cambiar Manualmente Número de Factura")
        NUF1$ = CStr(NUF)
        If TRIM$(NUF1$) = "" Then Exit Sub
        If IsNumeric(NUF) = True And XVALO(NUF) > 0 Then
            'MsgBox NUF
            Text11 = NUF
            NFCAM% = 1
        Else
           Call MENSERR(32, "Número Ingresado No Válido.")
           GoTo 10
        End If
        '
      End If
    Else
        On Error Resume Next
        Text1.SetFocus
        On Error GoTo 0
    End If
    '
End Sub
'
Private Sub Text11_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   If XVALO(Text11.TEXT) > 9 Then
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

Private Sub Text16_Change()
   Label16 = Left$(ECOARCH("ECUECON", Text16), 20)
End Sub

Private Sub Text16_DblClick()
   Call Command12_Click
End Sub

''///****NUEVO***//////////
Private Sub TEXT17_CHANGE()
   Call CALTOFA2
End Sub
'
Private Sub Text17_GotFocus()
   Text17.SelStart = 0
   Text17.SelLength = Len(Text17)
End Sub

Private Sub Text17_LostFocus()
    '
    Text17 = FORMATNUM$(XVALO(Text17), "F11.2")
    '
End Sub

Private Sub Text19_Change()
  NTRA = XVALO(Text19)
  If NTRA > 0 Then
    If NTRA < 32767 Then
      NTRAI% = NTRA
      LENTREGA(2) = ECOARCH$("TRANSPOR", MKI$(NTRAI%))
    End If
  Else
    LENTREGA(2) = "" ' MODIFICADO PARA QUE SI EL TRANSPORTE ES 0 LIMPIE EL TEXT
  End If
End Sub

Private Sub Text19_DblClick()
   '\\\OT-06-0570-RG-27/12/06///
   Call Command24_Click
   '\\\OT-06-0570-RG-FIN///
End Sub

Private Sub Text2_gotfocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text20_Change()
  NSUC& = XVALO(Text20)
  Label27 = ""
  If EXISTE%(LUDAT$ + "SUCENTRE.RFS") > 0 Then
    If NSUC& > 0 Then
      If NSUC& <= ULTREG&("!SUCENTRE") Then
        XX$ = REGLEIDO$("!SUCENTRE", NSUC&)
        Label27 = TRIM$(Mid$(XX$, 3, 62))
        LENTREGA(0) = TRIM$(Mid$(XX$, 65, 48))
        LENTREGA(1) = TRIM$(Mid$(XX$, 113, 48))
      End If
    End If
    Text19 = NUTRANCLI%(XVALO(Text1), XVALO(Text20))
  End If
End Sub

Private Sub Text20_DblClick()
   Call Command25_Click
End Sub

Private Sub TEXT21_DblClick(Index As Integer)
   '\\\OT-06-0570-RG-27/12/06///
   If Index = 0 Then
      Call Command21_Click(Index)
   ElseIf Index = 1 Then
      Call Command21_Click(Index)
   End If
   '\\\OT-06-0570-RG-FIN///
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

Private Sub Text22_LostFocus()
    If IsNumeric(Text22) = True Then
       TX# = XVALO(Text22)
       Text22 = FORMATNUM$(TX#, "F5.2")
    End If
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
        If XVALO(TEXT21(Index).TEXT) < 1 Then
            Eco21(Index).TEXT = "CONCERTADOS FIJOS"
            ARCHIPRE$ = "MASTER": LIPRACTI% = 0
          Else
            Eco21(Index).TEXT = ECOARCH$("LISTAS", MKI$(XVALO(TEXT21(Index).TEXT)))
            ARCHIPRE$ = TRIM$(TEXT21(0).TEXT)
            While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
            ARCHIPRE$ = "LIPRE" + ARCHIPRE$: LIPRACTI% = XVALO(Mid$(ARCHIPRE$, 6, 3))
        End If
        If NOCAMONE% < 1 Then
           If MONELISTA%(XVALO(TEXT21(Index))) = 2 Then
                MONEMIS = "Dólares"
                MONEMIS.ForeColor = RGB(255, 0, 0)
                MONEMI% = 2
                TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
                MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
                TICAMBIS.Visible = True
                Label24.Visible = True
              Else
                MONEMIS = "Pesos"
                MONEMIS.ForeColor = RGB(0, 0, 255)
                MONEMI% = 1
                TICAMBIS = "1.0000"
                MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
                TICAMBIS.Visible = False
                Label24.Visible = False
           End If
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
    '\\\OT-06-0570-RG-27/12/06///
    Call CommSel(ARCHEC$, Command21(Index), TEXT21(Index))
    '\\\OT-06-0570-RG-FIN///
    'Call SELECHO(ARCHEC$)
    'TEXT21(Index).TEXT = VALACT1$(ARCHEC$)
    '
End Sub
   '
Function VALICLIE%()
   VCC% = 0
   If XVALO(Text1.TEXT) > 0 Then
     If XVALO(Text1.TEXT) <= 32767 Then
       NCLIE% = XVALO(Text1.TEXT)
       If RASOCLI$(NCLIE%) <> "" Then
         VCC% = 1
       End If
     End If
   End If
   '
99 VALICLIE% = VCC%
   '
End Function

Sub BLANFORMU()
   TASACLICK = 0
   Text2.TEXT = ""
   Text3.TEXT = ""
   Text4.TEXT = ""
   Text5.TEXT = ""
   Text7.TEXT = ""
   Text8.TEXT = ""
   Text10.TEXT = ""
   Text12.TEXT = ""
   Text14.TEXT = ""
   Text16 = ""
   Label16 = ""
   Label26 = ""
   Text17 = "" 'LIMPIA RET-IVA
   IEMBALA = ""
   Text19 = ""
   Text20 = ""
   Text22 = ""
   Label25 = ""
   Label28 = ""
   Ltota(10) = "Importe Total Neto"
   '
   If EXISTE%(LUDAT$ + "SUCENTRE.RFS") > 0 Then
     Call SETULTREG("!SUCENTRE", 0)
     Call FRESHECHO("!SUCENTRE")
   End If
   '
   '\\\OT-5-579-FG-16/08/05
   Ltota(3).Caption = "Total Documento:"
   Ltota(3).Left = 1260
   Text18.Visible = False
   '\\\OT-5-579-FG-FIN
   '\\\OT-08-0672-OD-26/08/08
   TIFADOC% = 2: Label7.Caption = "Factura:"
   Label6.Caption = "(Mat.Codif)": MODODOC% = 0
   Label13.Caption = "A": PIVADOC% = 1: VARIDOC% = 1
   '\\\OT-08-0672-OD-FIN
   '
   ' PONE LETRA POR OMISIÓN
   Label13.Caption = "A"
   '\\\OT-5-506-FG-02/09/05
   ' PONE PREFIJO POR OMISION AL 0001
'   UPREFI$ = TRIM$(CONTROL$("FACTURA", "PREFIFA"))
    CLAVE$ = "PREFAC-" & Label13.Caption '  MODIFICADO PARA GUARDAR PREFIJO DE FACTURA POR ESTACION
    UPREFI$ = TRIM$(CONTROL$(CLAVE$, IDENTER$))
   If UPREFI$ = "" Then UPREFI$ = "0001"
   Label12.Caption = UPREFI$
   '\\\OT-5-506-FG-FIN
   '
   Text6.TEXT = ""
   List1.Clear
   DOCUDIF = ""
   '
   Picture5.Visible = False
   List3.Clear
   '
   '\\\OT-8-0852-OD-02/09/05
   'SE INVIRTIO LA POSICION DE LOS DOS FOR QUE ESTAN A CONTINUACION POR LA ENTRADA
   'AL EVENTO CHANGE DEL TEXT
   PORDESCU.TEXT = ""
   For KK% = 6 To 0 Step -1
      Total(KK%).TEXT = ""
   Next KK%
   TLABEL28 = ""
   'EL TEXT15 QUE SE ENCONTRABA MAS ARRIBA EN EL CODIGO SE BAJO POR
   'QUE PROVOCABA QUE APARECIERA LA VENTANA CON EL MENSAJE
   '"IMPOSIBLE FACTURAR POR EL LIMITE DE CREDITO " ,
   Text15 = ""
   '
   For KKI% = 0 To 2
      TEXT21(KKI%).TEXT = ""
      Eco21(KKI%) = ""
      BASIVA(KKI%) = ""
      LENTREGA(KKI%) = ""
   Next KKI%
  '\\\OT-8-0852-OD-FIN

   '
   MARCONOTA.Visible = False
   HOII% = HOY(HOS$)
   Text9.TEXT = HOS$
   Call SETULTREG("!CARDEFAC", 0)
   If EXISTE%(LUCOM$ + "1CARDEFA.WKF") > 0 Then
     On Error Resume Next
     Kill LUCOM$ + "1CARDEFA.WKF"
     On Error GoTo 0
   End If
   'Call NUFACTURA
   '
   FACTLISTAEMPAQUE& = 0

End Sub
'
Sub CARDACLI()
    '
    Screen.MousePointer = 11
    '
    NC% = XVALO(Text1.TEXT): Text1.Refresh
    If NC% > 0 Then
      If RASOCLI$(NC%) <> "" Then
        If InStr(EMPREAC$, "DOSIVAC") > 0 Then Call MUESTRAVENCIDAS(NC%, 15)
      
        Text2.TEXT = RASOCLI$(NC%): Text2.Refresh
        Text3.TEXT = DOMICLI$(NC%): Text3.Refresh
        Text4.TEXT = CPOSCLI$(NC%): Text4.Refresh
        Text5.TEXT = LOCACLI$(NC%): Text5.Refresh
        Label26 = CODICLI$(NC%)
        Text7.TEXT = CSTRING$(MKS$(LIMICRE(NC%)), 3, 10, 2, 2)
          Text7.Refresh
        Text12.TEXT = TRIM$(POSIVAC$(NC%)) + " - CUIT " + CUITCLI$(NC%)
          Text12.Refresh
        '
        VEND% = VENDCLI%(NC%)
        FIN& = ULTREG&("VENDEDOR")
        For I& = 1 To FIN&
            TX$ = REGLEIDO$("VENDEDOR", I&)
            If Asc(Mid$(TX$, 1, 1)) = VEND% Then
               COMIS# = CVS(Mid$(TX$, 44, 4))
               Text22 = FORMATNUM$(COMIS#, "F5.2")
            End If
        Next I&
        '
        '\\\OT-5-579-FG-16/08/05
        YY$ = TRIM$(POSIVAC$(NC%))
        If YY$ = "COM.EXT." Then
          Ltota(3).Caption = "Total:"
          Ltota(3).Left = 250
          Text18 = "FOB"
          Text18.Visible = True
        End If
        '\\\OT-5-579-FG-FIN
        '
        'CREDISPO# = LIMICRE(NC%) - SALDCLI#(NC%) - CHEPEACRE#(NC%)
        CREDISPO# = XVALO(Text7) - SALDCLI#(NC%) - CHEPEACRE#(NC%)
        Text8.ForeColor = RGB(0, 0, 0)
        If CREDISPO# < 0 Then Text8.ForeColor = RGB(255, 0, 0)
        Text8.TEXT = CSTRING$(MKD$(Abs(CREDISPO#)), 3, 10, 2, 2)
          Text8.Refresh
        '
        TEXT21(0).TEXT = Str$(LIPRCLI%(NC%))
        TEXT21(1).TEXT = Str$(CPAGCLI%(NC%))
        TEXT21(2).TEXT = Str$(VENDCLI%(NC%))
        '
        If XVALO(TEXT21(2).TEXT) < 1 Then
           Eco21(2).TEXT = "GERENCIA"
        End If
        '
        NCLIE% = XVALO(Text1.TEXT)
        If NCLIE% > 0 Then
          If MODODOC% > 0 Then
              PORDESCU.TEXT = ""
            Else
              PORDESCU.TEXT = DESCUCLI$(NCLIE%)
          End If
        End If
        PIVADOC% = PIVACLI%(NCLIE%)
        Call CARLETDOC: 'Call NUFACTURA '\\\OT-06-0051-WAR-15/02/06///
        Refresh
        '
        PPVI% = PIVADOC%
        If PPVI% < 0 Or PPVI% > 7 Then PPVI% = 1
        Text16 = CONTROL$("SVD451S", "CUEVE-" + TRIM$(Str$(PPVI%)))
        '
        'verificacion si es factura de exportacion valida si tiene token
        If (PPVI% = 5 Or PPVI% = 7) And FACELEC% > 0 Then
        CARPETOK$ = "C:\FLEXOFT\FACELEC\"
        ARCHIVO_TOKEN$ = "TOKEN_WSFEX.TXT"
            If EXISTE%(CARPETOK$ + ARCHIVO_TOKEN$) = 1 Then
                 FechArch = FileSystem.FileDateTime(CARPETOK$ + ARCHIVO_TOKEN$)
                 FECHARCHIVO% = CVINT(FechArch)
                 FECHHOY% = HOY(HO$)
                 If FECHARCHIVO% < FECHHOY% Then
                    Screen.MousePointer = 1
                    Call MENSERR(24, "Token Vencido. Vuelva a Generarlo")
                    Text1 = ""
                    Exit Sub
                 End If
            Else
                 Screen.MousePointer = 1
                 Call MENSERR(24, "Debe Generarse el Archivo 'TOKEN_WSFEX.TXT'")
                 Text1 = ""
                 Exit Sub
            End If
        End If
        '
        '\\\OT-5-506-FG-02/09/05
        ' PARA QUE TOME EN CUENTA EL PREFIJO DE LAS FACTURAS
        ' DE EXPORTACION
        If (PPVI% = 5 Or PPVI% = 7) Then
          Label12.Caption = PREFEXPO$()
'          CLAVE$ = "PREFAC-" & Label13.Caption
'          UPREFI$ = TRIM$(CONTROL$(CLAVE$, IDENTER$))
'          If UPREFI$ = "" Then UPREFI$ = PREFEXPO$()
'          If UPREFI$ = "" Then UPREFI$ = "0001"
'          Label12.Caption = UPREFI$
        End If
        '\\\OT-5-506-FG-FIN
        'VALIDO SI TIENE CONFIGURADO P.VTA FACT.ELECT. EXPO, SI NO DEJA EL CONFIGURADO EN MODO HABITUAL
        If PPVI% = 5 And FACELEC% = 1 Then
          PUVEX$ = TRIM$(CONTROL("FACTURA", "PUVEFAEX"))
          If PUVEX$ <> "" Then
            Label12 = PUVEX$
          End If
        End If
        '
        If PPVI% = 0 Or PPVI% = 3 Or PPVI% = 4 Or PPVI% = 6 Then ' PARA EL CASO QUE LA FACTURA SEA B
            CLAVE$ = "PREFAC-" & Label13.Caption
            UPREFI$ = TRIM$(CONTROL$(CLAVE$, IDENTER$))
            If UPREFI$ = "" Then UPREFI$ = "0001"
            Label12.Caption = UPREFI$
        End If
        Call NUFACTURA '\\\OT-06-0051-WAR-15/02/06///
        '
        If EXISTE%(LUDAT$ + "SUCENTRE.RFS") > 0 Then
          '\\\OT-7-010-RG-12/01/07
          Call FILTERFILE("SUCENTRE", "", 5, MKI$(NCLIE%))
          '\\\OT-7-010-RG-FIN
        End If
        '
        For KKI% = 0 To 2
          LENTREGA(KKI%) = LUENTRECLI$(NCLIE%, KKI%)
        Next KKI%
        '
        NTRAN% = NUTRANCLI%(NCLIE%)
        If NTRAN% > 0 Then
          Text19 = TRIM$(Str$(NTRAN%))
        End If
        '
        Text6 = ANOTACLI$(NC%)
        MARCONOTA.Visible = True
        MARCONOTA.Refresh
        '
        If VERLICRE% = 1 Then
          LMCR# = LIMICRE(NC%)
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
        If EXEDELIMCC%(NC%, 0, DIFER) > 0 Then
            LIMCCTE$ = TRIM$(FORMATNUM$(LCred_CCte(NC%), "f12.2"))
            TX$ = "Límite de Crédito Asignado en C.Cte. $ " & LIMCCTE$ & " ya está Superado."
            If DERACCE%("AUTOFACT", "Autorizar Facturas con Exceso Credito") < 2 Then
                Call COMUNI("Imposible Facturar a este Cliente.\  \Límite de Crédito Asignado en C.Cte. ya está Superado.\  \Informe al Supervisor de Ventas.")
                Call BLANFORMU
                On Error Resume Next
                Text1.SetFocus
                On Error GoTo 0
            Else
              If COMALTER%(TX$ + "\   \Puede Autorizar la Facturación\con Crédito Excedido.\\Cancelar\Autorizar Factura") < 2 Then
                Call BLANFORMU
                On Error Resume Next
                Text1.SetFocus
                On Error GoTo 0
              End If
            End If
        End If
        ' SE INTRODUCE UN XVALO PORQUE POREMBAL DEVUELVE UN TEXTO Y LA COMPARACION SE HACE COMO SI FUESE NUMERICO
        If XVALO(POREMBAL(NCLIE%)) > 0 Then
          CUEMBA$ = CONTROL$("SVD451S", "CUEVE-12")
          If CUEMBA$ = "" Or CUECOINT%(CUEMBA$) < 1 Then
            Call MENSERR(24, "Falta Definir Cuenta Contable\Correspondiente a Embalajes.\   \FLEXOFT Utilizará la Cuenta de Ventas Netas.")
          End If
        End If
        '
        PVCL% = PIVADOC%
        If PVCL% = 1 Or PVCL% = 2 Or PVCL% = 3 Or PVCL% = 4 Or PVCL% = 6 Then
           CUTT$ = CUITCLI$(NC%)
           If VALICUIT%(CUTT$) < 1 Then
              Call MENSERR(24, "Imposible Emitir Documento Fiscal.\  \Posición de IVA del Cliente\Requiere C.U.I.T. Válido.")
              Text1 = ""
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
Function FLINUMERA%()         ' FECHA DE NUEVA NUMERACION
   '
   Static FLINU%
   '
   If FLINU% = 0 Then
      FLINU% = FECHANUM("01/01/99")
      If ULTREG&("FORMFAC") > 0 Then
         FFS$ = REGLEIDO$("FORMFAC", 2&)
         FLINU% = CVI(Mid$(FFS$, 19, 2))
      End If
   End If
   '
   FLINUMERA% = FLINU%
   '
End Function
'
Function NUMECORR%()          ' NUMERACION CORRELATIVA O INDEPENDIENTE
   '
   Static NUMECO%
   '
   If NUMECO% = 0 Or NUMECO% = (-1) Then
      NUMECO% = 1
      QTRX$ = CONTROL$("NUMEFACT", "NUMECORR")
      If QTRX$ = "INDEPEN" Then
           NUMECO% = (-1)
         ElseIf QTRX$ = "UNIFICA" Then
           NUMECO% = (-2)
      End If
   End If
   '
   'CASO QUE ES FACTURA ELECTRONICA Y DE EXPORTACION MARCA CORRELATIVIDAD INDEPENDIENTE
   If FACELEC% = 1 And Label13.Caption = "E" Then NUMECO% = (-1)
   '
   NUMECORR% = NUMECO%
   '
End Function
'
Function MONUMERA$()
   '
   Static MONUM$
   '
   If MONUM$ = "" Then
      MONUM$ = "AUTO"
      If TRIM$(CONTROL$("", "NUMFACTU")) = "MANUAL" Then
         MONUM$ = "MANUAL"
      End If
   End If
   '
   MONUMERA$ = MONUM$
   '
End Function
'
Sub NUREMIFA_HAS(NROREMIT&, PUNTODEVENTA$)
   '
   VAMOVIX$ = Label13
   Call TOMANUMDELCONFIS("RT", VAMOVIX$, PUNTODEVENTA$, NUMECOMP$)
   NROREMIT& = XVALO(NUMECOMP$)
   '
End Sub
'
Sub NUFACTURA()
   '
   DATEDOC% = FECHANUM(Text9.TEXT)
   If DATEDOC% < FLINUMERA% Then
      Call MENSERR(24, "Imposible Facturar en la Fecha Elegida.\  \La Fecha de Factura Debe Ser POSTERIOR\a la Fecha de Nueva Numeración.")
      Call CIERRARCH("*.*")
      Call FINAL("")
   End If
   '
   Screen.MousePointer = 11
   '
   OPMOVIX$ = "FC"
   If TIFADOC% = 3 Then OPMOVIX$ = "ND"
   If TIFADOC% = 6 Then OPMOVIX$ = "NC"
   VAMOVIX$ = Label13
   '
   ' SI ES IMPRESORA FISCAL, TOMA EL NUMERO DEL HASAR
   If CONTROFIS% = 2 Then
      If MUECOMHAS% < 2 Then
         If VAMOVIX$ = "A" Or VAMOVIX$ = "B" Then ' PARA EXCLUIR EXPORTACION
           Call TOMANUMDELCONFIS(OPMOVIX$, VAMOVIX$, PUNTODEVENTA$, NUMECOMP$)
           Label12 = PUNTODEVENTA$
           NPP& = XVALO(NUMECOMP$)
           Text11.TEXT = TRIM$(Str$(NPP&))
           '
           Call NUREMIFA_HAS(NROREMIT&, PUNTODEVENTA$)
           If NROREMIT& > XVALO(Text10.TEXT) Then
              Text10.TEXT = TRIM$(Str$(NROREMIT&))
           End If
           GoTo 999
         End If
      End If
   End If
   
   FLINUME = CDbl(CVDAT(FLINUMERA%))
   '
   FEXUL = DATEDOC%
   FLIDIAS = CDbl(CVDAT(DIASANTES(FEXUL, 15)))
   '//////////*******NOTA*******////////
   'Con posterioridad se detecto un Error en la Impòrtacion
   '(VARIMOVIM(A,B), SE IMPORTO EN VEZ DE LA LETRA CON EL OPCIMOVIX$(ND,NC).
   'Por tal motivo el sistema no toma en cuenta los numero
   'de NC-ND inferiores al 01/11/04.
   'Pero si lo hace a futuro. es decir yo puedo reiniciar numeración
   'al 01/11/04
   
  ' FLIMI = FLIDIAS: MODORAPIDO% = 1
  ' If FLINUME > FLIMI Then
  '    FLIMI = FLINUME
  '    MODORAPIDO% = 0
  ' End If
   
   '
'22
   Call APERBASDAT("CABAN")
   'jandy sql esta rutina debe ser igual
   'ASELE$ = "SELECT NumeComp FROM CajaBanc "
   '
   FLIMI = FETEXCOR1$(FLINUMERA)
     
   ASELE$ = "SELECT MAX(NumeComp) AS Maximo FROM CajaBanc WITH(NOLOCK)"
  
   If OPMOVIX$ = "ND" Or OPMOVIX$ = "NC" Then
     If NUMECORR% = (-2) Then    ' Unificada debitos y creditos
       ASELE$ = ASELE$ + "WHERE ((TipoMovim = 'FVEN' AND (OpciMovim = 'ND' OR OpciMovim = 'NC')) "
       ASELE$ = ASELE$ + "OR (TipoMovim = 'FCOM' AND (OpciMovim = 'DP' OR OpciMovim = 'CP'))) "
       GoTo 25
     End If
   End If
   '
   ASELE$ = ASELE$ + "WHERE TipoMovim = 'FVEN' "
   If NUMECORR% <> 1 Then
       ASELE$ = ASELE$ + "AND OpciMovim = '" & OPMOVIX$ & "' "
     Else
       '       ASELE$ = ASELE$ + "AND (OpciMovim = 'FC' OR OpciMovim = 'ND' OR OpciMovim = 'NC') "
       '       //////****NUEVA CONDICION PARA CALCULAR PROX. Nº FAC***////
       '       ESTA CONDICION LA AGREGUE  PARA QUE TOME EN CTA. LOS Nº
       '       DE COMPROBANTES ARRUINADOS: OR OpciMovim = 'ANUL'
       ASELE$ = ASELE$ + "AND (OpciMovim = 'FC' OR OpciMovim = 'ND' OR OpciMovim = 'NC' OR OpciMovim = 'ANUL') "
   End If
   '
25 ASELE$ = ASELE$ + "AND VariMovim = '" & VAMOVIX$ & "' "
   If XVALO(Label12) <= 1 Then
       ASELE$ = ASELE$ + "AND (NumeTalo = 0 OR NumeTalo = 1) "
     Else
       ASELE$ = ASELE$ + "AND NumeTalo = " & XVALO(Label12) & " "
   End If
   ASELE$ = ASELE$ + "AND FechEmisi >= '" & FLIMI & "' "
   
   'ASELE$ = ASELE$ & "ORDER BY NumeComp DESC;"
   '
   'Set NUFAXI = CueCorri.OpenRecordset(ASELE$, dbOpenDynaset, dbReadOnly)
   '
   NUMAXI& = 0
   Dim MIRS As ADODB.Recordset
   Set MIRS = FLEXCONN.Execute(FILTSQL$(ASELE$))
   If Not IsNull(MIRS!maximo) Then
        NUMAXI& = MIRS!maximo
   End If
   MIRS.Close
   Set MIRS = Nothing
   
   
'   On Error Resume Next
'   NUFAXI.MoveFirst
'   If Err < 1 Then
'       On Error GoTo 0
'       NUMAXI& = NUFAXI!NumeComp
'     Else
'       On Error GoTo 0
'       If MODORAPIDO% = 1 Then
'          FLIMI = FLINUME
'          MODORAPIDO% = 0
'          GoTo 22
'       End If
'   End If
'
99 If Visible = True Then
     If NUMAXI& < 1 Then
       NUMAXI& = XVALO(InputBox$("Ingrese Número del Próximo Comprobante del Tipo Elegido", "Incicialización de Talonarios")) - 1
       If NUMAXI& < 0 Then GoTo 99
     End If
   End If
   NPP& = NUMAXI& + 1
   Text11.TEXT = TRIM$(Str$(NPP&))
   Text11.Refresh
   '
   If TIFADOC% = 2 Then
      If TRIM$(Text10.TEXT) = "" Or FACELEC% = 1 Then       'Para facturas electronica actualiza el Nº de remito
        If VERIFSALIDA("NOMESS") < 3 Then NPVX$ = Label12                                     'segun el punto de venta seleccionado
        If PIVADOC% = 5 Or PIVADOC% = 7 Then
          EPAC$ = TRIM$(UCase$(EMPREAC$))
          If InStr(EPAC$, "DOSIVA") > 0 Or InStr(EPAC$, "TANIT") > 0 Then
            NPVX$ = "0001"
          End If
        End If
        '
        Call NUREMIFA(NROREMIT&, NPVX$)
        If NROREMIT& > XVALO(Text10.TEXT) Or FACELEC% = 1 Then      'Para facturas electronica actualiza el Nº de remito
          Text10.TEXT = TRIM$(Str$(NROREMIT&))                      'segun el punto de venta seleccionado
        End If
        '
        ' esto se puso por reclamo de dosivac de que a veces se 'traba'
        ' y repite el numero de remito anterior - Ana el 22 de setiembre de 2008
995     If VERINURE% < 0 Then
           NROREMIT& = NROREMIT& + 1
           Text10.TEXT = TRIM$(Str$(NROREMIT&))
           GoTo 995
        End If
        '
      End If
      Screen.MousePointer = 1
      Exit Sub
     Else
      Text10.TEXT = ""
   End If
   '
999 Screen.MousePointer = 1
   '
End Sub
'
Sub TOMANUMDELCONFIS(TDC$, LTDC$, PREFIFA$, NUMECOMP$)
     Call FORHAS320.TOMANUMDELCONFIS(TDC$, LTDC$, PREFIFA$, NUMECOMP$)
End Sub
'
Function VERINUFA%(NFCY$)
   '
   Static FLINUMERA%
   Static NUMECORR%
   '
   If CONTROFIS% = 2 Then
       VERINUFA% = 1
       Exit Function
   End If
   '
   OPMOVIX$ = Left$(NFCY$, 2)
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
   If NUMECORR% = 0 Then
      NUMECORR% = 1
      QTRX$ = CONTROL$("NUMEFACT", "NUMECORR")
      If QTRX$ = "INDEPEN" Then
           NUMECORR% = (-1)
         ElseIf QTRX$ = "UNIFICA" Then
           NUMECORR% = (-2)
      End If
   End If
   '
   FLIMI = CDbl(CVDAT(FLINUMERA%))
22 Call APERBASDAT("CABAN")
   SQLX$ = "SELECT * FROM CajaBanc "
   SQLX$ = SQLX$ + "WHERE (TipoMovim = 'FVEN' and CODICOM_LAR = '" & NFCY$ & "') "
   If NUMECORR% = 1 Then
      NFCY1$ = "FC" + Mid$(NFCY$, 3)
      NFCY2$ = "ND" + Mid$(NFCY$, 3)
      NFCY3$ = "NC" + Mid$(NFCY$, 3)
      SQLX$ = SQLX$ + " OR (TipoMovim = 'FVEN' and (CODICOM_LAR = '" & NFCY1$ & "' OR CODICOM_LAR = '" & NFCY2$ & "' OR CODICOM_LAR = '" & NFCY3$ & "')) "
   End If
   If OPMOVIX$ = "ND" Or OPMOVIX$ = "NC" Then
     If NUMECORR% = (-2) Then    ' Unificada debitos y creditos
       SQLX$ = SQLX$ + "OR (TipoMovim = 'FCOM' AND CODICOM_LAR = '" & NFCY$ & "') "
     End If
   End If
   
   'jandy sql
   
   'Set NUFAXI = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset, dbReadOnly)
   Dim NUFAXI As ADODB.Recordset
   Set NUFAXI = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   
   On Error Resume Next
   NUFAXI.MoveFirst
   If Err Then
       VERINUFA% = 1
     Else
       Do While Not NUFAXI.EOF
         'If CDbl(NUFAXI.FechEmisi) >= FLIMI Then
         If CDbl(NUFAXI!fechemisi) >= FLIMI Then
            VERINUFA% = (-1)
            GoTo 90
         End If
       NUFAXI.MoveNext
       Loop
       VERINUFA% = 1
   End If

90 On Error GoTo 0
NUFAXI.Close
Set NUFAXI = Nothing

   '
End Function
'
Function VERINURE%()
   '
   Static PREFIREM$, FLIM%, CMO$
   '
   If CONTROFIS% = 2 Then
       VERINURE% = 1
       Exit Function
   End If
   '
   If CMO$ = "" Then
     CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
     If CMO$ = "" Then CMO$ = "RT"
   End If
   '
   If PREFIREM$ = "" Then
     PREFIREM$ = CONTROL$("", "PREFIREM")
     If PREFIREM$ = "" Then PREFIREM$ = "0001"
     'ESTO SE AGREGA PORQUE LOS REMITOS DE IMPRESORA FISCAL TIENEN EL MISMO
     'PUNTO DE VENTA QUE LA FACTURA
     If CONTROL$("FACTURA", "SALIDA") = "IFISCAL" Then PREFIREM$ = Label12
   End If
   '
   'ESTO SE AGREGA PORQUE LOS REMITOS DE LAS
   'FACTURAS ELECTRONICAS TIENEN EL PUNTO DE VENTA DE LA FACTURA
   If FACELEC% = 1 And VERIFSALIDA("NOMESS") < 3 Then PREFIREM$ = Label12
   '
   If FLIM% < 1 Then
     FENUE$ = CONTROL$("", "FENUREFA")
     If FENUE$ = "" Then FENUE$ = "18/04/03"
     FLIM% = FECHANUM(FENUE$)
     If FLIM% > HOY(HO$) Then FLIM% = HOY(HO$)
   End If
   'FLINUME = CDbl(CVDAT(FLIM%))
   FLINUME = FECHATEX(FLIM%)
   
   '
   REMICOR$ = "RT." + TRIM$(Text10)
   REMILAR$ = Text10: While Len(REMILAR$) < 8: REMILAR$ = "0" + REMILAR$: Wend
   REMILAR$ = CMO$ + "-X-" + PREFIREM$ + "-" + REMILAR$
   '
   Call APERBASDAT("STOCKS")
   'JANDY SQL
   SQLX$ = "SELECT * FROM MOVISTO "
   SQLX$ = SQLX$ + " WHERE FechMov >= '" & FLINUME & "' "
   SQLX$ = SQLX$ + " AND CodiMovi = '" & CMO$ & "' "
   SQLX$ = SQLX$ + " AND DoPriLar = '" & REMILAR$ & "' "
   '
   Dim MOVITMP As ADODB.Recordset
   Set MOVITMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With MOVITMP
      If Not (.EOF And .BOF) Then
          VERINURE% = (-1)
          Exit Function
      End If
   End With
   MOVITMP.Close
   Set MOVITMP = Nothing
   '
   SQLX$ = "SELECT * FROM MOVIREPA "
   SQLX$ = SQLX$ + " WHERE FechMov >= '" & FLINUME & "' "
   SQLX$ = SQLX$ + " AND CodiMovi = '*R' "
   SQLX$ = SQLX$ + " AND DoPriLar = '" & REMILAR$ & "' "
   '
   Set MOVITMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With MOVITMP
      If Not (.EOF And .BOF) Then
          VERINURE% = (-1)
          Exit Function
      End If
   End With
   MOVITMP.Close
   Set MOVITMP = Nothing
   '
   VERINURE% = 1
   '
End Function
'
Function VERINUFAVIE%(NFCY$)
    '
    Static NUMECORR%
    '
    OPMOVIX$ = Left$(NFCY$, 2)
    '
    If NUMECORR% = 0 Then
      NUMECORR% = 1
      QTRX$ = CONTROL$("NUMEFACT", "NUMECORR")
      If QTRX$ = "INDEPEN" Then
           NUMECORR% = (-1)
         ElseIf QTRX$ = "UNIFICA" Then
           NUMECORR% = (-2)
      End If
    End If
    '
    Call APERBASDAT("CABAN")
    With CajaBanc
       SQLX$ = "(TipoMovim = 'FVEN' and CODICOM_LAR = '" & NFCY$ & "') "
       If OPMOVIX$ = "ND" Or OPMOVIX$ = "NC" Then
         If NUMECORR% = (-2) Then    ' Unificada debitos y creditos
           SQLX$ = SQLX$ + "OR (TipoMovim = 'FCOM' AND CODICOM_LAR = '" & NFCY$ & "') "
         End If
       End If
       '
       VERINUFAVIE% = 1
       .FindFirst SQLX$
       If .NoMatch = False Then
         VERINUFAVIE = (-1)            'Nº DE FACTURA PREEXISTENTE
       End If
    End With
    '
End Function

Private Sub Text6_GotFocus()
   On Error Resume Next
   'Text1.SetFocus
   Command6.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text7_DblClick()
    '\\\OT-08-00732-OD-29/08/08///
    If DERACCE%("MODLICRE", "Modificar Límite de Crédito") >= 2 Then
       NC% = XVALO(Text1)
       LimCre$ = Text7
       LiCreNue = ROUND#(XVALO(InputBox$("Ingrese el Nuevo Límite", "Límite de Crédito", LimCre$)), 2)
       If LiCreNue < 0.01 Then
            Call MENSERR(20, "Ingreso No Válido")
            Exit Sub
         ElseIf LiCreNue >= 10000000 - 0.01 Then
            Call COMUNI("Importe Supera el Valor Máximo Permitido")
            Exit Sub
         Else
            SQLX$ = "SELECT  Lcre_Cli FROM DEUDOR12 "
            SQLX$ = SQLX$ + "WHERE Nume_Cli = " & NC% & ""
            '
            Dim LIMTMP As ADODB.Recordset
            Set LIMTMP = New ADODB.Recordset
25          On Error Resume Next
            LIMTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
            If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
               On Error GoTo 0
               Call CapturaErrorOpen
               GoTo 25
            End If
            On Error GoTo 0
            '
            With LIMTMP
               If Not .EOF Then
                 !Lcre_Cli = LiCreNue
                 .Update
               End If
            End With
            LIMTMP.Close
            Set LIMTMP = Nothing
            Text7 = FORMATNUM(LiCreNue, "F10.2")
            Text7.Refresh
            '
            CREDISPO# = XVALO(Text7) - SALDCLI#(NC%) - CHEPEACRE#(NC%)
            Text8.ForeColor = RGB(0, 0, 0)
            If CREDISPO# < 0 Then Text8.ForeColor = RGB(255, 0, 0)
            Text8.TEXT = CSTRING$(MKD$(Abs(CREDISPO#)), 3, 10, 2, 2)
            Text8.Refresh
       End If
    End If
End Sub

Private Sub Text7_GotFocus()
   On Error Resume Next
   Text14.SetFocus
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
     TICACH(MONEMI%) = XVALO(TICAMBIS)
   End If
   TICAMBX# = 1
   If MONEMI% > 1 Then
     TICAMBX# = XVALO(TICAMBIS)
   End If
End Sub

Private Sub TICAMBIS_LostFocus()
   TICAMBIS = FORMATNUM(XVALO(TICAMBIS), "F9.4")
End Sub

Private Sub TICAMBIS_DblClick()
   'Call CARTAMONED
   '
   '\\\OT-07-0086-RG-06/03/07///
   If DERACCE%("COTILOCA", "Ajustar Tipo de Cambio Local") >= 2 Then
      ENTRATICA.Frame4.Caption = "Cotización Local para " + MONEMIS.Caption
      ENTRATICA.Text4 = TICAMBIS.Caption
      ENTRATICA.Text4 = FORMATNUM(XVALO(ENTRATICA.Text4), "F10.4")
      ENTRATICA.Show 1
      If TRIM$(ENTRATICA.Text4) <> "" Then
        TICAMBIS = TRIM$(ENTRATICA.Text4)
      End If
   End If
   '\\\OT-07-0086-RG-FIN///
   '
End Sub

Private Sub Timer1_Timer()
   '
   If CAMBIOVALGRILLA% > 0 Then
      Screen.MousePointer = 11
      Call CALTOFAC
      CAMBIOVALGRILLA% = 0
      Screen.MousePointer = 1
   End If
   '
End Sub

Sub PRIFACTURA()
    '
    If (CONTROFIS% = 1 Or CONTROFIS% = 2) Then ' OBVIA LOS COMPROBANTES 'E' DE LA IMP. FISCAL
       'imprime por controlador / impresora fiscal
       Call PRICONFIS
       If VARIDOC% <> 3 Then   ' PARA QUE EN FACTURA DE EXPORTACION SIGA A
          Exit Sub             ' EMITIR FACTURA POR IMPRESORA COMÚN
       End If
    End If
    '
    Static AGREDESP%
    TICAMBX# = XVALO(TICAMBIS)
    '
    If REPRIFAC% = 1 Then GoTo 10
    '
    ' ESTO ES PARA ACELERAR BUSQUEDA DEL NUMERO DE SERIE
    CMO$ = TRIM$(CONTROL$("", "CODREFAC")):   If Len(CMO$) <> 2 Then CMO$ = "RT"
    NC1% = XVALO(Text1.TEXT)
    '
    Call APERBASDAT("STOCKS")
    SQLX$ = " SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + "WHERE Nume_Clie = " & NC1% & " "
    SQLX$ = SQLX$ + " AND (CodiMovi = '" & CMO$ & "' "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
    SQLX$ = SQLX$ + " AND DoPriLar <> '' "
    SQLX$ = SQLX$ + " AND (DoSecLar is null or (DoSecLar = '') OR (DoSecLar LIKE 'P/F*')) "
    SQLX$ = SQLX$ + " AND FechMov >= '" & DESELEC & " ' "
    SQLX$ = SQLX$ + " AND FechMov <= '" & HASELEC & "' "
    Dim CONSITEMP As ADODB.Recordset
    Set CONSITEMP = New ADODB.Recordset
    '
10  CAHOJAS% = 0
    FORIPRE$ = TRIM$(CONTROL$("", "FORFAC-" + Label13.Caption))
    TIPOCOM$ = TRIM$(Label6)
    If TIPOCOM$ = "(T.Libre)" Then
      If Label13.Caption <> "E" Then   ' EXCLUYE FACTURAS DE EXPORTACION
        FORIPRE$ = TRIM$(CONTROL$("", "FACTEXLI"))
        If UCase$(Label13) = "B" Then
          FORIPRE$ = TRIM$(CONTROL$("", "FACTELIB"))
        End If
      End If
    End If
    '\\\OT-06-0038-WAR-01/02/06///
    FORIALT$ = ""
    If TIFADOC% = 3 Or TIFADOC% = 6 Then 'SI NC O ND
      If TIFADOC% = 3 Then ' SI NOTA DE DEBITO
        If UCase$(Label13.Caption) = "A" Then
          FORIALT$ = CONTROL$("", "FORNDE-A") 'ND-A
          ElseIf UCase$(Label13.Caption) = "B" Then
          FORIALT$ = CONTROL$("", "FORNDE-B") 'ND-B
        End If
      End If
      '
      If TIFADOC% = 6 Then ' SI NOTA DE CREDITO
        If UCase$(Label13.Caption) = "A" Then
          FORIALT$ = CONTROL$("", "FORNCR-A") 'NC-A
          ElseIf UCase$(Label13.Caption) = "B" Then
          FORIALT$ = CONTROL$("", "FORNCR-B") 'NC-B
        End If
      End If
    End If
    If TRIM$(FORIALT$) <> "" Then
      FORIPRE$ = FORIALT$
    End If
    ''\\\OT-06-0038-WAR-FIN///
    If FORIPRE$ = "" Then
       Call MENSERR(24, "Falta Definir Controlador de Impresión")
       Exit Sub
    End If
    '
    If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
       Call MENSERR(24, "Imposible Imprimir Documento.\No Existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
       Exit Sub
    End If
    '
    CODFOR$ = FORIPRE$
    '
    Screen.MousePointer = 11
    ITPORHOJA% = XVALO(ATRIFORM$(FORIPRE$, "MAX-ITEM"))
    If ITPORHOJA% < 1 Then
      ITPORHOJA% = 32767
    End If
    '
    CANCELPRINT% = 0
    FECHDOC$ = Text9.TEXT
    NUMEDOC$ = TRIM$(Text11.TEXT)
    While Len(NUMEDOC$) < 8
      NUMEDOC$ = "0" + NUMEDOC$
    Wend
    NUMEDOC$ = Label12.Caption + "-" + NUMEDOC$
    '
    TIPODOC$ = TRIM$(REPLACAR$(Label7.Caption, ":", "")): CODOCUM$ = "FAC"
    If TIFADOC% = 3 Then TIPODOC$ = "Nota de Debito": CODOCUM$ = "DEB"
    If TIFADOC% = 6 Then TIPODOC$ = "Nota de Credito": CODOCUM$ = "CRE"
    CPAG% = XVALO(TEXT21(1).TEXT)
    '
    DESTIDOC% = XVALO(Text1.TEXT)
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
    '
    CACOLTAB1% = 26
    '
    CAITAB1% = 0
    CAITCRU% = 0
    '
    SUMATOTA# = XVALO(Total(0).TEXT)
    IDESCUEN# = XVALO(Total(1).TEXT)
    '
    COPIASREM% = XVALO(FORFINFAC.Text2(1))
    If COPIASREM% < 1 Then Text10.TEXT = ""
    '
    NroRemi$ = TRIM$(Text10.TEXT)
    If NroRemi$ <> "" Then
      While Len(NroRemi$) < 8: NroRemi$ = "0" + NroRemi$: Wend
      PREFIREM$ = CONTROL$("", "PREFIREM")
      If PREFIREM$ = "" Then PREFIREM$ = "0001"
      'ESTO SE AGREGA PORQUE LOS REMITOS DE LAS
      'FACTURAS ELECTRONICAS TIENEN EL PUNTO DE VENTA DE LA FACTURA
      If FACELEC% = 1 And VERIFSALIDA("NOMESS") < 3 Then PREFIREM$ = Label12
      '
      NroRemi$ = PREFIREM$ + "-" + NroRemi$
    End If
    '
    If CONSICLI% = 1 Then
      NroRemi$ = NROREMICO$
    End If
    '
    COEFIVAINCLU# = 1
    If XVALO(Total(2).TEXT) > 0.005 Then
      COEFIVAINCLU# = (XVALO(Total(2).TEXT) + XVALO(Total(3).TEXT) + XVALO(Total(4).TEXT) + XVALO(Total(5).TEXT)) / XVALO(Total(2).TEXT)
    End If
    '
    If TRIM$(Text14.TEXT) <> "" Then
      NUOCOMPRA$ = TRIM$(Text14.TEXT)
    End If
    '
    CODPARAM$(1) = "CIVA-CLI": DOCPARAM$(1) = Label13.Caption
    CODPARAM$(2) = "NRO-REMI": DOCPARAM$(2) = NroRemi$
    CODPARAM$(3) = "NRO-OCOM":  DOCPARAM$(3) = NUOCOMPRA$
    CODPARAM$(4) = "NRO-AENT":  DOCPARAM$(4) = NROPEDIDO$
    CODPARAM$(5) = "IMP-NETO": DOCPARAM$(5) = ""
      NUVENDEX$ = AJUSTI$(FORMATNUM$(XVALO(TEXT21(2)), "I3"), 4)
    CODPARAM$(6) = "VENDEDOR": DOCPARAM$(6) = NUVENDEX$ + Eco21(2).TEXT
    CODPARAM$(7) = "COND-PAG": DOCPARAM$(7) = Eco21(1).TEXT
    CODPARAM$(8) = "POR-IVA": DOCPARAM$(8) = PORIVA(0).Caption
    CODPARAM$(9) = "IMP-IVA": DOCPARAM$(9) = ""
    CODPARAM$(10) = "POR-IVA1": DOCPARAM$(10) = PORIVA(1).Caption
    CODPARAM$(11) = "IMP-IVA1": DOCPARAM$(11) = ""
    CODPARAM$(12) = "IMP-DESC": DOCPARAM$(12) = ""
    CODPARAM$(13) = "IMPNEDES": DOCPARAM$(13) = ""
    'CODPARAM$(4) = "NRO-OCOM": DOCPARAM$(4) = Text10.TEXT
    CODPARAM$(14) = "IMP-TOTA": DOCPARAM$(14) = ""
    CODPARAM$(15) = "POR-DESC": DOCPARAM$(15) = TRIM$(PORDESCU.TEXT)
    CODPARAM$(16) = "IMP-RNI": DOCPARAM$(16) = ""
    CODPARAM$(17) = "IMP-LET1": DOCPARAM$(17) = ""
    CODPARAM$(18) = "IMP-LET2": DOCPARAM$(18) = ""
    CODPARAM$(19) = "IMP-LET3": DOCPARAM$(19) = ""
    CODPARAM$(23) = "TOTA-DOL": DOCPARAM$(23) = ""    ' Valor en PESOS equivalente a dolares
    '
    CODPARAM$(20) = "DOMI-TRA": DOCPARAM$(20) = LENTREGA(0).TEXT
    CODPARAM$(21) = "DESTINO": DOCPARAM$(21) = LENTREGA(1).TEXT
    CODPARAM$(22) = "TRANSPOR": DOCPARAM$(22) = LENTREGA(2).TEXT
    CODPARAM$(24) = "MONEDEMI": DOCPARAM$(24) = MONEMIS
    CODPARAM$(25) = "PERC-BRU": DOCPARAM$(25) = ""
    CODPARAM$(26) = "SUB-TOTA": DOCPARAM$(26) = ""
    CODPARAM$(27) = "PERC-IVA": DOCPARAM$(27) = "" 'PERCIVA
    CODPARAM$(28) = "FECH-VEN": DOCPARAM$(28) = Format$(FVENCIM, "dd/mm/yy") ' FECHA DE VENCIMIENTO
      On Error Resume Next
      FEXO = FECHANUM(FECHDOC$)
      FEXU = CVINT(FVENCIM)
      CADIVEN = DIENTRE(FEXO, FEXU)
      On Error GoTo 0
    CODPARAM$(32) = "FECH-LAN": DOCPARAM$(32) = FORMATNUM$(CADIVEN, "F5.0")
    '\\\OT-5-579-FG-16/08/05
    CODPARAM$(29) = "TIPO-VAL": DOCPARAM$(29) = TRIM$(Text18)          ' VALOR FOB
    CODPARAM$(30) = "TICAMBIO": DOCPARAM$(30) = TICAMBIS               ' TIPO DE CAMBIO
    CODPARAM$(31) = "CTAC-CLI": DOCPARAM$(31) = CONTACTCLI$(DESTIDOC%) ' CONTACTO
    '\\\OT-5-579-FG-FIN
    '\\\OT-06-0125-WAR-03/04/06///
    CODPARAM$(33) = "NUBULTOS": DOCPARAM$(33) = TRIM$(FORFINFAC.Text2(2))          ' BULTOS
    CODPARAM$(34) = "PES-CONJ": DOCPARAM$(34) = TRIM$(FORFINFAC.Text2(11)) + " Kg" ' PESO NETO
    CODPARAM$(35) = "PES-COMP": DOCPARAM$(35) = TRIM$(FORFINFAC.Text2(4)) + " Kg"  ' PESO BRUTO
    CODPARAM$(36) = "BANC-CHE": DOCPARAM$(36) = TRIM$(FORFINFAC.Text2(5))          'BANCO EXTERIOR
    CODPARAM$(37) = "DOCU-APL": DOCPARAM$(37) = TRIM$(FORFINFAC.Text2(8))          ' BANCO LOCAL
    CODPARAM$(38) = "NUME-CHE": DOCPARAM$(38) = TRIM$(FORFINFAC.Text2(10))         ' CREDITO
    CODPARAM$(39) = "DEN-CUEN": DOCPARAM$(39) = TRIM$(FORFINFAC.Text2(9))          ' GUIA DE IMPORTACION
    CODPARAM$(40) = "REF-ASIE": DOCPARAM$(40) = TRIM$(FORFINFAC.Text2(6))          ' PUERTO DE EMBARQUE
    CODPARAM$(41) = "DET-PASE": DOCPARAM$(41) = TRIM$(FORFINFAC.Text2(7))          ' DESTINO
    CODPARAM$(42) = "MEDIDA-1": DOCPARAM$(42) = TRIM$(IEMBALA)                      ' gasto de embalaje
    CODPARAM$(43) = "ZONA-VEN": DOCPARAM$(43) = TRIM$(Str$(GRUCOCLI%(NC1%)))
    '\\\OT-06-0125-WAR-FIN///
    'AGREGUEMH04-10
    CODPARAM$(44) = "NRO-PASE": DOCPARAM$(44) = TRIM$(Label27.Caption) 'Sucursal
    '
    CODPARAM$(45) = "ITO-GRAV": DOCPARAM$(45) = ""
    CODPARAM$(46) = "ITO-NOGR": DOCPARAM$(46) = ""
    CODPARAM$(47) = "ITO-GRA1": DOCPARAM$(47) = ""
    CODPARAM$(48) = "ITO-GRA2": DOCPARAM$(48) = ""
    '
    CODPARAM$(49) = "IMP-DEBE": DOCPARAM$(49) = FORFINFAC.Text2(13)  ' FLETE
    CODPARAM$(50) = "IMP-HABE": DOCPARAM$(50) = FORFINFAC.Text2(12)  ' SEGURO
    CODPARAM$(51) = "TOT-HABE": DOCPARAM$(51) = FORMATNUM(TOTATOTA#, "F13.2")
    CODPARAM$(52) = "REF-MAT4": DOCPARAM$(52) = FINFACTU107.Label12
    '
    CAPARDOC% = 52 '32 '\\\OT-06-0125-WAR-03/04/06///
    '
    Call SETULTREG("!OBSERVOF", 0)
    If FORFINFAC.Text4 <> "" Then
      JR& = 0
      FORMOBSE$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "OBSERVA/FORMATO")))
      ANCHOB% = XVALO(Mid$(FORMOBSE$, 2)): If ANCHOB% < 32 Then ANCHOB% = 32
      Call FRATEXTO(FORFINFAC.Text4, ANCHOB%)
      For KKII% = 1 To CARETEX%
        JR& = JR& + 1
        Call GRAREG("!OBSERVOF", RETEX$(KKII%), JR&)
      Next KKII%
      Call SETULTREG("!OBSERVOF", JR&)
      Call CIERRARCH("!OBSERVOF")
    End If
    POVERTI% = XVALO(TRIM$(UCase$(ATRIFORM$(CODFOR$, "OBSERVA/POSINVER"))))
    '
    PRINTERPORT$ = PORTUFAC$
    CACOPDOC% = XVALO(FORFINFAC.Text2(0).TEXT)
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
    ARCHILIB$ = "!1CARDEFA"
    FIN& = ULTREG&("!CARDEFAC")
    SUBTOTA# = 0: SUBTOTB# = 0
  '
  '\\\OT-06-0339-MH-24/08/06///
    'control es la funcion que devuelve en este caso SI o NO
    'y asi reemplaza el cod orig o no seg lo selecc en configuracion
    'remstan es el modulo en que esta el proyecto
    'replaori es una clave que refleja la configuracion
    REPLAORI% = 0
      If CONTROL$("REMSTAN", "REPLAORI") = "SI" Then
        REPLAORI% = 1
      End If
     'va fuera del for para q no se ejecute innecesariamente
  '\\\OT-06-0339-MH-FIN///
    '
    NOJA% = 0
    If REPRIFAC% = 1 Then
       PROJA% = XVALO(InputBox$("Imprime desde Hoja Número", "Primera Hoja de Factura", "1"))
    End If
    If PROJA% < 1 Then PROJA% = 1
    '
    For U& = 1 To FIN&
      If (U& Mod ITPORHOJA%) = 1 Then
        ' PONE ENCABEZADO
        CAITAB1% = 0
        CAHOJAS% = CAHOJAS% + 1
      End If
      '
      XXX$ = REGLEIDO$("!CARDEFAC", U&)
      CIPRE% = CVI(Mid$(XXX$, 5, 2))
      CAPEX$ = Str$(CVS(Left$(XXX$, 4)))       ' cantidad
      '
      NPEDIDD& = XVALO(TRIM$(Mid$(XXX$, 371, 16)))
      SUBCO$ = SUBCODI$(NPEDIDD&, CIPRE%)
      '
      CAITAB1% = CAITAB1% + 1
      '
      '\\\OT-06-0339-MH-24/08/06///
      CODDX$ = CODEXT$(CIPRE%) + SUBCO$ ' en coddx$ trae el codigo externo a partir de cipre(int)
      If REPLAORI% = 1 Then
        CODDX$ = CODORIG$(CIPRE%, NC1%) ' + SUBCO$
      ' se pasa el codigo interno, mas el nº cliente(NC1%) y trae el codigo externo solicitado
      End If
      '\\\OT-06-0339-MH-FIN///
      '
      '
      PLANOX$ = VALOBADA("MASTER", "NPlano", "CODINT = " & CODINT%(CODDX$))
      '
      TETABU1$(1, CAITAB1%) = CODDX$
      TETABU1$(2, CAITAB1%) = Mid$(XXX$, 7, 64)
        ' SE AGREGA CONDICION PARA QUE ACTUE UNICAMENTE EN DOSIVAC
        ' Y MANTENER LA FUNCIONALIDAD PARA EL RESTO DE LOS USUARIOS
        If PIVADOC% = 5 And InStr(EMPREAC$, "DOSIVA") > 0 Then
          DEALTE$ = TRIM$(DESCRIT2$(CIPRE%))
          If DEALTE$ <> "" Then
             TETABU1$(2, CAITAB1%) = DEALTE$
          End If
        End If
      TETABU1$(3, CAITAB1%) = CAPEX$
      'TETABU1$(4, CAITAB1%) = PRELI$
      'TETABU1$(5, CAITAB1%) = DESCX$
      'TETABU1$(6, CAITAB1%) = PREUX$
      'TETABU1$(7, CAITAB1%) = ITOTX$
      'TETABU1$(8, CAITAB1%) = FENTX$
      TETABU1$(9, CAITAB1%) = Str$(U&)
      TETABU1$(10, CAITAB1%) = Mid$(XXX$, 71, 4)    ' unidad de medida
      TETABU1$(11, CAITAB1%) = Mid$(XXX$, 75, 64)   ' referencia 1
      'TETABU1$(12, CAITAB1%) = Mid$(XXX$, 139, 64)  ' referencia 1 ' SE AGREGO ABAJO
      TETABU1$(13, CAITAB1%) = Mid$(XXX$, 203, 64)  ' referencia 1
      TETABU1$(14, CAITAB1%) = PLANOX$
      'TETABU1$(14, CAITAB1%) = Mid$(XXX$, 315, 24)  ' numero de plano
      TETABU1$(15, CAITAB1%) = Mid$(XXX$, 339, 16)  ' numero original
      TETABU1$(16, CAITAB1%) = Mid$(XXX$, 355, 16)  ' numero de o.compra
      TETABU1$(17, CAITAB1%) = Mid$(XXX$, 387, 16)  ' numero de remito
      TETABU1$(18, CAITAB1%) = Mid$(XXX$, 403, 16)  ' numero de doc-recepcion
      TETABU1$(19, CAITAB1%) = FECHATEX$(CVI(Mid$(XXX$, 419, 2))) ' fecha de recepcion
      '
      FACIVITEM = 1: If Label13.Caption = "B" Then FACIVITEM = 1 + (CVS(Mid$(XXX$, 441, 4))) / 100
      LPRE% = XVALO(TEXT21(0).TEXT)
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
      SUBTOTA# = SUBTOTA# + XVALO(TETABU1$(23, CAITAB1%))
      SUBTOTB# = SUBTOTB# + XVALO(TETABU1$(26, CAITAB1%))
      '
      NUSERRR$ = TRIM$(Mid$(XXX$, 139, 64)) ' SI TIENE NUMERO DE SERIE
      If NUSERRR$ <> "" And UCase$(NUSERRR$) <> "S/N" Then 'IMPRIME NUMERO DE SERIE ABAJO DE LA DESCRIPCION
        CAITAB1% = CAITAB1% + 1
        TETABU1$(12, CAITAB1%) = "S/Nro.: " + NUSERRR$ ' numero de serie
      End If
      '
      If CONTROL("FACTUREM", "INCLUSPE") = "SI" Then ' IMPRIME ESPECIFICACIONES DEL ARTICULO
        XX$ = ESPECIFIT$(CIPRE%)
        If TRIM$(XX$) <> "" Then
          CAITAB1% = CAITAB1% - 1 ' DESCUENTA 1 PARA PISAR LA DESCRIPCION
          Call PRINTEXT(XX$, 2, "DES-MAT")
        End If
      End If
      '
      Call CARLINSEP         'Carga linea de separacion entre codigo y codigo
      '
      If COLTELIB% > 0 Then
          TELIBRE$ = REGLEIDO$(ARCHILIB$, U&)
          FORTELIB$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "DES-MAT/FORMATO")))
          ATELIB% = XVALO(Mid$(FORTELIB$, 2))
          Call FRATEXTO(TELIBRE$, ATELIB%)
          TETABU1$(2, CAITAB1%) = RETEX$(1)
          For KKU1% = 2 To CARETEX%
            CAITAB1% = CAITAB1% + 1
            TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
          Next KKU1%
        Else
          FORTELIB$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "DES-MAT/FORMATO")))
          ATELIB% = XVALO(Mid$(FORTELIB$, 2))
          If ATELIB% > 100 Then
            ATELIB% = ATELIB% Mod 100
            If CIPRE% > 0 Then
                 TELIBRE$ = DESCRIT0$(CIPRE%)
                 If PIVADOC% = 5 Then
                   DEALTE$ = TRIM$(DESCRIT2$(CIPRE%))
                   If DEALTE$ <> "" Then
                      TELIBRE$ = DEALTE$
                   End If
                 End If
               Else
                 TELIBRE$ = Mid$(XXX$, 7, 64)
            End If
            Call FRATEXTO(TELIBRE$, ATELIB%)
            TETABU1$(2, CAITAB1%) = RETEX$(1)
            For KKU1% = 2 To CARETEX%
              CAITAB1% = CAITAB1% + 1
              TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
            Next KKU1%
          End If
      End If
      '
      ' AQUI BUSCA EL NUMERO DE SERIE
      If CIPRE% > 0 Then
        DOPRIC$ = TRIM$(Mid$(XXX$, 387, 16))
        NORIT% = CVI(Mid$(XXX$, 503, 2))
        '
        SQLX$ = "SELECT * from MOVISTO "
        SQLX$ = SQLX$ + "WHERE DoPriCor = '" & DOPRIC$ & "' "
        SQLX$ = SQLX$ + "AND Cod_Inte = " & CIPRE% & " "
        SQLX$ = SQLX$ + "AND NuorItem = " & NORIT% & " "
        
        Set CONSITEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
        With CONSITEMP
          If Not (.EOF And .BOF) Then
            SERNU$ = !ReferCor
            If Left$(SERNU$, 5) = "S.No." Then
              CAITAB1% = CAITAB1% + 1
              TETABU1$(2, CAITAB1%) = SERNU$
            End If
          End If
        End With
      End If
      '
      ' AQUI AGREGA DESPACHO DE IMPORTACION
      If AGREDESP% = 0 Then
        AGREDESP% = (-1)
        If CONTROL$("FACTURA", "AGREDESP") = "SI" Then AGREDESP% = 1
      End If
      If AGREDESP% = 1 Then
        DESPAI$ = TRIM$(ATRIB1$(CIPRE%))
        If DESPAI$ <> "" Then
          CAITAB1% = CAITAB1% + 1
          TETABU1$(2, CAITAB1%) = "Despacho: " + DESPAI$
        End If
      End If
      'ESTO SE LLEVO PARA ARRIBA POR QUE SE GENERA UN DIFERENCIA EN EL CAITAB Y NO REALIZA LA SUMA
'      SUBTOTA# = SUBTOTA# + XVALO(TETABU1$(23, CAITAB1%))
'      SUBTOTB# = SUBTOTB# + XVALO(TETABU1$(26, CAITAB1%))
      '
      If (U& Mod ITPORHOJA%) = 0 Then
        If U& < FIN& Then
          If CACOPDOC% > 0 Then
            DOCPARAM$(26) = TRIM$(CSTRING$(MKD$(SUBTOTA#), 4, 12, 2, 2))
            NOJA% = NOJA% + 1
            If REPRIFAC% = 0 Or NOJA% >= PROJA% Then
                 ' ESTO QUE SIGUE ES PARA IMPRIMIR TRANSPORTE DEL IMPORTE
                 FORSUBTO$ = UCase$(TRIM$(UCase$(ATRIFORM$(CODFOR$, "SUB-TOTA/FORMATO"))))
                 If Left$(FORSUBTO$, 4) = "SUMA" Then
                     CAITAB1% = CAITAB1% + 1
                     For KKII% = 1 To CACOLTAB1%: TETABU1$(KKII%, CAITAB1%) = "": Next KKII%
                     TETABU1$(23, CAITAB1%) = "=========================="
                     TETABU1$(26, CAITAB1%) = "=========================="
                     CAITAB1% = CAITAB1% + 1
                     For KKII% = 1 To CACOLTAB1%: TETABU1$(KKII%, CAITAB1%) = "": Next KKII%
                     TETABU1$(2, CAITAB1%) = "TRANSPORTE .........................................................."
                     TETABU1$(23, CAITAB1%) = FORMATNUM$(SUBTOTA#, "F13.2")
                     TETABU1$(26, CAITAB1%) = FORMATNUM$(SUBTOTB#, "F13.2")
                 End If
                 ' HASTA AQUI AGREGADO EL 22/12/2009 PARA IMPRIMIR TRANSPORTE
                 Call PRINTDOC("@" + FORIPRE$, "NOCANOBS")
              Else
                 For KKI% = 0 To CACOLTAB1%
                   For KKJ% = 0 To CAITAB1%
                     TETABU1$(KKI%, KKJ%) = ""
                   Next KKJ%
                 Next KKI%
                 CAITAB1% = 0
            End If
          End If
        End If
      End If
109 Next U&
    '
    'IMPRIME GASTOS DE EXPORTACION
    If Label13.Caption = "E" And XVALO(FORFINFAC.Text2(15)) > 0 Then
        CAITAB1% = CAITAB1% + 1
        For KKII% = 1 To CACOLTAB1%: TETABU1$(KKII%, CAITAB1%) = "": Next KKII%
        TETABU1$(2, CAITAB1%) = TRIM$(FORFINFAC.Text10)
        TETABU1$(3, CAITAB1%) = 1
        TETABU1$(22, CAITAB1%) = TRIM$(FORFINFAC.Text2(15))
        TETABU1$(26, CAITAB1%) = TRIM$(FORFINFAC.Text2(15))
        SUBTOTB# = SUBTOTB# + XVALO(TETABU1$(26, CAITAB1%))
    End If
    '
    FORSUBTO$ = UCase$(TRIM$(UCase$(ATRIFORM$(CODFOR$, "SUB-TOTA/FORMATO"))))
    If Left$(FORSUBTO$, 3) = "SUM" Then
      CAITAB1% = CAITAB1% + 1
      For KKII% = 1 To CACOLTAB1%: TETABU1$(KKII%, CAITAB1%) = "": Next KKII%
      TETABU1$(23, CAITAB1%) = "=========================="
      TETABU1$(26, CAITAB1%) = "=========================="
      CAITAB1% = CAITAB1% + 1
      For KKII% = 1 To CACOLTAB1%: TETABU1$(KKII%, CAITAB1%) = "": Next KKII%
      TETABU1$(2, CAITAB1%) = "TOTAL .........................................................."
      TETABU1$(23, CAITAB1%) = FORMATNUM$(SUBTOTA#, "F13.2")
      TETABU1$(26, CAITAB1%) = FORMATNUM$(SUBTOTB#, "F13.2")
    End If
    '
    If POVERTI% < 1 Then
      CAITAB1% = CAITAB1% + 1
      For KU& = 1 To ULTREG&("!OBSERVOF")
        TTXX$ = REGLEIDO$("!OBSERVOF", KU&)
        CAITAB1% = CAITAB1% + 1
        TETABU1$(2, CAITAB1%) = RTrim$(TTXX$)
      Next KU&
      Call SETULTREG("!OBSERVOF", 0)
    End If
        
    CODPARAM$(5) = "IMP-NETO": DOCPARAM$(5) = FORMATNUM$(XVALO(Total(0).TEXT) + XVALO(IEMBALA) + XVALO(FORFINFAC.Text2(15)), "F13.2")
    CODPARAM$(9) = "IMP-IVA": DOCPARAM$(9) = Total(3).TEXT
    CODPARAM$(11) = "IMP-IVA1": DOCPARAM$(11) = Total(4).TEXT
    CODPARAM$(12) = "IMP-DESC": DOCPARAM$(12) = Total(1).TEXT
    CODPARAM$(13) = "IMPNEDES": DOCPARAM$(13) = FORMATNUM$(XVALO(Total(2).TEXT) + XVALO(FORFINFAC.Text2(15)), "F13.2")
    CODPARAM$(14) = "IMP-TOTA": DOCPARAM$(14) = FORMATNUM$(XVALO(Total(6).TEXT) + XVALO(FORFINFAC.Text2(15)), "F13.2")
    CODPARAM$(16) = "IMP-RNI": DOCPARAM$(16) = Total(5).TEXT
       VADOLA# = XVALO(Total(6).TEXT) * TICAMBX#
    DOCPARAM$(23) = TRIM$(CSTRING$(MKD$(VADOLA#), 4, 12, 2, 2))
    CODPARAM$(25) = "PERC-BRU": DOCPARAM$(25) = Text15
    DOCPARAM$(26) = ""
    DOCPARAM$(27) = Text17                          'PARA IMPRIMIR PERC-IVA
      ITOGRAVA# = XVALO(Total(2)) - XVALO(TLABEL28)
    DOCPARAM$(45) = TRIM$(FORMATNUM$(ITOGRAVA#, "F13.2"))
      INOGRAVA# = XVALO(TLABEL28)
    DOCPARAM$(46) = TRIM$(FORMATNUM$(INOGRAVA#, "F13.2"))
    DOCPARAM$(47) = TRIM$(BASIVA(0))
    DOCPARAM$(48) = TRIM$(BASIVA(1))
      TOTATOTA# = XVALO(Total(6)) + XVALO(FORFINFAC.Text2(12)) + XVALO(FORFINFAC.Text2(13))
    DOCPARAM$(51) = FORMATNUM$(TOTATOTA#, "F13.2")
    '
    If Label13.Caption = "B" Then
      DOCPARAM$(5) = FORMATNUM$(COEFIVAINCLU# * (XVALO(Total(0).TEXT) + XVALO(IEMBALA)), "F13.2")
      DOCPARAM$(12) = FORMATNUM$(COEFIVAINCLU# * XVALO(Total(1).TEXT), "F13.2")
      DOCPARAM$(13) = FORMATNUM$(COEFIVAINCLU# * XVALO(Total(2).TEXT), "F13.2")
      DOCPARAM$(42) = FORMATNUM$(COEFIVAINCLU# * XVALO(IEMBALA), "F13.2")          ' gasto de embalaje
      DOCPARAM$(45) = TRIM$(FORMATNUM$(COEFIVAINCLU# * ITOGRAVA#, "F13.2"))
      DOCPARAM$(46) = TRIM$(FORMATNUM$(COEFIVAINCLU# * INOGRAVA#, "F13.2"))
      DOCPARAM$(47) = TRIM$(FORMATNUM$(COEFIVAINCLU# * XVALO(BASIVA(0)), "F13.2"))
      DOCPARAM$(48) = TRIM$(FORMATNUM$(COEFIVAINCLU# * XVALO(BASIVA(1)), "F13.2"))
    End If
    '
    CODFOR$ = FORIPRE$
    LOIMPLET% = XVALO(ATRIFORM$(CODFOR$, "LOIMPLET"))
    If LOIMPLET% < 16 Then LOIMPLET% = 32
    TTTT# = XVALO(Total(6).TEXT) + XVALO(FORFINFAC.Text2(15))
    TOTFAC$ = NUMTEX$(TTTT#)
    If MONECO$ <> "" Then
      PPXX% = InStr(TOTFAC$, "PESOS")
      If PPXX% > 0 Then
        TOTFAC$ = Left$(TOTFAC$, PPXX% - 1) + UCase$(MONECO$) + Mid$(TOTFAC$, PPXX% + 5)
      End If
    End If
    Call FRATEX(TOTFAC$, NRX1$, NRX2$, NRX3$, LOIMPLET%)
    CODPARAM$(17) = "IMP-LET1": DOCPARAM$(17) = NRX1$
    CODPARAM$(18) = "IMP-LET2": DOCPARAM$(18) = NRX2$
    CODPARAM$(19) = "IMP-LET3": DOCPARAM$(19) = NRX3$
    '
    If CACOPDOC% > 0 Then
      Call PRINTDOC("@" + FORIPRE$, "NOCANOBS")
    End If
    '
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    '
    If CONTROFIS% < 2 Then      ' OBVIA LOS COMPROBANTES 'E' DE LA IMP. FISCAL
       CACOPDOC% = XVALO(FORFINFAC.Text2(1).TEXT)
       If CACOPDOC% > 0 Then
          If TRIM$(Text10) <> "" Then
           '
           Call LIMPIATETAB
           '
           DOCUORIG$ = NUMEDOC$
           Call PRIREMIFA
           '
           Call LIMPIATETAB
           Call IMPRIMENOTAENTREGA
          End If
       End If
    End If
    '
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    '
    If NOCONFIRM% > 0 Then
       Call COMUNI("Se han Emitido los Documentos Solicitados")
    End If
    For KKI% = 0 To CACOLTAB1%
      For KKJ% = 0 To CAITAB1%
        TETABU1$(KKI%, KKJ%) = ""
      Next KKJ%
    Next KKI%
    Call SETULTREG("!OBSERVOF", 0)
    '
End Sub

Sub PRICONFIS()
' DISTINGUIR:
' A) fACTURA / DEBITO / CREDITO
' B) A O B
   '

   If YAIMPRIMIO% > 0 Then Exit Sub
   '
   If TIFADOC% = 2 Then                ' ticket factura
        CACOPDOCF% = XVALO(FORFINFAC.Text2(0).TEXT) ' LEE CANTIDAD COPIAS FACTURA
        CACOPDOCR% = XVALO(FORFINFAC.Text2(1).TEXT) ' LEE CANTIDAD COPIAS REMITO
        If VARIDOC% = 1 Then           ' factura a
            If CACOPDOCF% > 0 Then Call PRITIFACA
            If CONTROFIS% = 2 And CACOPDOCR% > 0 Then Call PRITIREMI
          ElseIf VARIDOC% = 2 Then     ' factura b
            If CACOPDOCF% > 0 Then Call PRITIFACB
            If CONTROFIS% = 2 And CACOPDOCR% > 0 Then Call PRITIREMI
          ElseIf VARIDOC% = 3 Then     ' FACTURA EXPORTACION
            If CONTROFIS% = 2 And CACOPDOCR% > 0 Then
               Call COMUNI("Se Imprimirá Remito de Exportacion por Impresora Fiscal.")
               Call PRITIREMI
            End If
            If CACOPDOCF% > 0 Then
               Call MENSERR(24, "Cliente de Exportacion.\  \Imposible Emitir Factura Exportacion\por Impresora Fiscal.")
            End If
          Else
            Call MENSERR(24, "Imposible Emitir Factura Variante '" & TRIM$(Str$(VARIDOC%)) & "'\por Impresora Fiscal.")
            GoTo 90
        End If
        Exit Sub
     ElseIf TIFADOC% = 3 Then          ' nota de debito
        If VARIDOC% = 1 Then           ' N/D A
            Call PRINODEBA
          ElseIf VARIDOC% = 2 Then     ' N/D B
            Call PRINODEBB
          Else
            GoTo 90
        End If
        Exit Sub
     ElseIf TIFADOC% = 6 Then           ' nota de credito
        If VARIDOC% = 1 Then            ' N/C A
            Call PRINOCREA
          ElseIf VARIDOC% = 2 Then        ' N/C B
            Call PRINOCREB
          Else
            GoTo 90
        End If
        Exit Sub
     Else
    
90     Call MENSERR(24, "Comprobante no Admisible por Controlador Fiscal. Consulte")
       Call FINAL("")
   End If
   '
End Sub
Sub PRITIREMI()
    '
    If CONTROFIS% = 1 Then
        Call MENSERR(24, "Imposible Emitir Remito por Controlador Fiscal")
      ElseIf CONTROFIS% = 2 Then
        Call FORHAS320.PRITIREMI
    End If
    '
End Sub

Sub PRITIFACA()
    '
    If CONTROFIS% = 1 Then
      Call FORHASAR.PRITIFACA
    ElseIf CONTROFIS% = 2 Then
      Call FORHAS320.PRITIFACA
    End If
    '
End Sub


Sub PRITIFACB()
    '
    If CONTROFIS% = 1 Then
      Call FORHASAR.PRITIFACB
    ElseIf CONTROFIS% = 2 Then
      Call FORHAS320.PRITIFACB
    End If
    '
End Sub
Sub PRINODEBA()
   '
   If CONTROFIS% = 1 Then
     Call FORHASAR.PRINODEBA
   ElseIf CONTROFIS% = 2 Then
     Call FORHAS320.PRINODEBA
   End If
   '
End Sub
Sub PRINODEBB()
   '
   If CONTROFIS% = 1 Then
     Call FORHASAR.PRINODEBB
   ElseIf CONTROFIS% = 2 Then
     Call FORHAS320.PRINODEBB
   End If
   '
End Sub
Sub PRINOCREA()
    '
    If CONTROFIS% = 1 Then
      Call FORHASAR.PRINOCREA
    ElseIf CONTROFIS% = 2 Then
      Call FORHAS320.PRINOCREA
    End If
    '
End Sub

Sub PRINOCREB()
   '
   If CONTROFIS% = 1 Then
     Call FORHASAR.PRINOCREB
   ElseIf CONTROFIS% = 2 Then
     Call FORHAS320.PRINOCREB
   End If
   '
End Sub

Sub PRIREMIFA()
    '
    CACOPDOC% = XVALO(FORFINFAC.Text2(1).TEXT)
    If CACOPDOC% < 1 Then Exit Sub
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FOREMIFA"))
    If TRIM$(FORIPRE$) <> "" Then
       '
       '\\\OT-06-0114-WAR-31/03/06///
       Call SETULTREG("!OBSERVOF", 0)
       If FORFINFAC.Text5 <> "" Then
         JR& = 0
         CODFOR$ = FORIPRE$
         FORMOBSE$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "OBSERVA/FORMATO")))
         ANCHOB% = XVALO(Mid$(FORMOBSE$, 2)): If ANCHOB% < 32 Then ANCHOB% = 32
         Call FRATEXTO(FORFINFAC.Text5, ANCHOB)
         For KKII% = 1 To CARETEX%
           JR& = JR& + 1
           Call GRAREG("!OBSERVOF", RETEX$(KKII%), JR&)
         Next KKII%
         Call SETULTREG("!OBSERVOF", JR&)
         Call CIERRARCH("!OBSERVOF")
       End If
       POVERTI% = XVALO(TRIM$(UCase$(ATRIFORM$(CODFOR$, "OBSERVA/POSINVER"))))
       '\\\OT-06-0114-WAR-31/03/06///
       '
       PREFIREM$ = CONTROL$("", "PREFIREM")
       If PREFIREM$ = "" Then PREFIREM$ = "0001"
       'ESTO SE AGREGA PORQUE LOS REMITOS DE LAS
       'FACTURAS ELECTRONICAS TIENEN EL PUNTO DE VENTA DE LA FACTURA
       If FACELEC% = 1 And VERIFSALIDA("NOMESS") < 3 Then PREFIREM$ = Label12
       '
       NroRemi$ = TRIM$(Text10.TEXT)
       While Len(NroRemi$) < 8: NroRemi$ = "0" + NroRemi$: Wend
       NroRemi$ = PREFIREM$ + "-" + NroRemi$
       '
       NC% = XVALO(Text1.TEXT)
       FECHDOC$ = Text9.TEXT
       NUMEDOC$ = NroRemi$
       TIPODOC$ = "Remito"
       '
       If TRIM$(Text14.TEXT) <> "" Then
         NUOCOMPRA$ = TRIM$(Text14.TEXT)
       End If
       '
       CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = TRIM$(Mid$(TRXP$, 1, 32))
       CODPARAM$(2) = "NUBULTOS": DOCPARAM$(2) = TRIM$(FORFINFAC.Text2(2).TEXT)
       CODPARAM$(3) = "CUIT-TRA": DOCPARAM$(3) = TRIM$(Mid$(TRXP$, 65, 16))
       CODPARAM$(4) = "VALO-DEC": DOCPARAM$(4) = TRIM$(FORFINFAC.Text2(3).TEXT)
       CODPARAM$(5) = "COND-TRA": DOCPARAM$(5) = TRIM$(Mid$(TRXP$, 97, 32))
       CODPARAM$(6) = "DOCUCOND": DOCPARAM$(6) = TRIM$(Mid$(TRXP$, 129, 24))
       CODPARAM$(7) = "DEPOSAL ": 'DOCPARAM$(7) = TRIM$(DEPOSAL$)
       CODPARAM$(8) = "NRO-REMI": DOCPARAM$(8) = TRIM$(NroRemi$)
       CODPARAM$(9) = "NRO-OCOM":  DOCPARAM$(9) = NUOCOMPRA$
       CODPARAM$(10) = "PATE-TRA":  DOCPARAM$(10) = TRIM$(Mid$(TRXP$, 81, 16))
       CODPARAM$(11) = "DOMI-TRA":  DOCPARAM$(11) = LENTREGA(0).TEXT
       CODPARAM$(12) = "LOCA-TRA":  DOCPARAM$(12) = TRIM$(Mid$(TRXP$, 159, 32))
       CODPARAM$(13) = "COND-PAG": DOCPARAM$(13) = Eco21(1).TEXT
       CODPARAM$(14) = "NRO-AENT":  DOCPARAM$(14) = NROPEDIDO$
       CODPARAM$(15) = "DESTINO": DOCPARAM$(15) = LENTREGA(1).TEXT
       CODPARAM$(16) = "TRANSPOR": DOCPARAM$(16) = LENTREGA(2).TEXT
       CODPARAM$(17) = "MEDIDA-1": DOCPARAM$(17) = TRIM$(IEMBALA)             ' gasto de embalaje
       CODPARAM$(18) = "ZONA-VEN": DOCPARAM$(18) = TRIM$(Str$(GRUCOCLI%(NC%)))
       CODPARAM$(19) = "TOTA-ITS": DOCPARAM$(19) = Label25
         NTRA% = XVALO(Text19)
       CODPARAM$(20) = "REF-ASIE": DOCPARAM$(20) = DOMITRANS$(NTRA%)
       CODPARAM$(21) = "DEN-CUEN": DOCPARAM$(21) = LOCATRANS$(NTRA%)
       CODPARAM$(22) = "MEDIDA-2": DOCPARAM$(22) = TELETRANS$(NTRA%)
       'AGREGUE 02-10MH
       CODPARAM$(23) = "NRO-PASE": DOCPARAM$(23) = TRIM$(Label27.Caption)
       '\\\OT-07-0007-RG-08/01/07///
       HORA$ = HORAENT$(NC%)
       CODPARAM$(24) = "Z-PLANO": DOCPARAM$(24) = TRIM$(HORA$)
          NUVENDEX$ = AJUSTI$(FORMATNUM$(XVALO(TEXT21(2)), "I3"), 4)
       CODPARAM$(25) = "VENDEDOR": DOCPARAM$(25) = NUVENDEX$ + Eco21(2).TEXT
       '\\\OT-07-0007-RG-FIN///
       '
       CAPARDOC% = 25
       '
       CODTABU1$(1) = "COD-MAT"
       CODTABU1$(2) = "DES-MAT"
       CODTABU1$(3) = "CANTIDAD"
       CODTABU1$(4) = "ORD-ITEM"
       CACOLTAB1% = 4
       '
       DESTIDOC% = NC%
       CAITAB1% = 0
       '
       If FORFINFAC.Text3 <> "" Then
         Call FRATEXTO(FORFINFAC.Text3, 40)
         For KKII% = 1 To CARETEX%
           CAITAB1% = CAITAB1% + 1
           TETABU1$(1, CAITAB1%) = ""
           TETABU1$(2, CAITAB1%) = RETEX$(KKII%)
           TETABU1$(3, CAITAB1%) = ""
           TETABU1$(4, CAITAB1%) = Str$(CAITAB1%)
         Next KKII%
         FORFINFAC.Text3 = ""
         FORFINFAC.Frame5.Visible = False
         GoTo 80
       End If
       '
       ARCHILIB$ = "!1CARDEFA"
      
      '\\\OT-06-0339-MH-24/08/06///
      REPLAORI% = 0
      If CONTROL$("REMSTAN", "REPLAORI") = "SI" Then
        REPLAORI% = 1
      End If
      '\\\OT-06-0339-MH-FIN///
       
       For U& = 1 To ULTREG&("!CARDEFAC")
         XXX$ = REGLEIDO$("!CARDEFAC", U&)
         CANS$ = Str$(CVS(Left$(XXX$, 4))) ' cantidad
         NC1% = XVALO(Text1.TEXT) ' AGREGADO MH
         CIPRE% = CVI(Mid$(XXX$, 5, 2)) ' AGREGADO
         'If XVALO(CANS$) >= 0.05 Then
           If TRIM$(Mid$(XXX$, 387, 16)) = "" Then
             CAITAB1% = CAITAB1% + 1
             DDTXT$ = Mid$(XXX$, 7, 64)                  ' descripcion
             ' SE AGREGA CONDICION PARA QUE ACTUE UNICAMENTE EN DOSIVAC
             ' Y MANTENER LA FUNCIONALIDAD PARA EL RESTO DE LOS USUARIOS
             If PIVADOC% = 5 And InStr(EMPREAC$, "DOSIVA") > 0 Then
               DEALTE$ = TRIM$(DESCRIT2$(CIPRE%))
               If DEALTE$ <> "" Then
                  DDTXT$ = DEALTE$
               End If
             End If
             '
             CDXXS$ = CODEXT$(CVI(Mid$(XXX$, 5, 2)))     ' codigo
             '\\\OT-06-0339-MH-24/08/06///
             If REPLAORI% = 1 Then
               ' se pasa el codigo interno, mas el nº cliente(NC1%) y trae el codigo externo solicitado
               CDXXS$ = CODORIG$(CIPRE%, NC1%) ' + SUBCO$
             End If
             '\\\OT-06-0339-MH-FIN///
             CINTE% = CODINT%(CDXXS$)
             Cantidad# = XVALO(CANS$)
             '
             If DESCOMPO%(CINTE%) > 0 Then
                'Imprimir componentes en factura
                Call AREXPLO(CINTE%, 1, 1)
                NUORIT% = 0
                For KKJ% = 1 To CAIT%
                  NUORIT% = NUORIT% + 1
                  CICO% = CIJ%(KKJ%)
                  CACICO = USOI(KKJ%)
                  
                  TETABU1$(1, CAITAB1%) = CODEXT$(CICO%)
                  TETABU1$(2, CAITAB1%) = "..." + DESCRIT0$(CICO%)
                  TETABU1$(3, CAITAB1%) = CACICO * Cantidad#
                  TETABU1$(4, CAITAB1%) = Str$(CAITAB1%)
                  If KKJ% < CAIT% Then CAITAB1% = CAITAB1% + 1
                Next KKJ%
              Else 'De lo Contrario Imprimir el Item
                TETABU1$(1, CAITAB1%) = CDXXS$
                TETABU1$(2, CAITAB1%) = DDTXT$
                TETABU1$(3, CAITAB1%) = CANS$
                TETABU1$(4, CAITAB1%) = Str$(CAITAB1%)
                If CONTROL("FACTUREM", "INCLUSPE") = "SI" Then ' IMPRIME ESPECIFICACIONES DEL ARTICULO
                  XX$ = ESPECIFIT$(CIPRE%)
                  If TRIM$(XX$) <> "" Then
                    CAITAB1% = CAITAB1% - 1 ' DESCUENTA 1 PARA PISAR LA DESCRIPCION
                    Call PRINTEXT(XX$, 2, "DES-MAT")
                  End If
                End If

             End If
             '
             NUSERRR$ = TRIM$(Mid$(XXX$, 139, 64)) 'NUMERO DE SERIE

             If NUSERRR$ <> "" And UCase$(NUSERRR$) <> "S/N" Then 'IMPRIME NUMERO DE SERIE ABAJO DE LA DESCRIPCION
               CAITAB1% = CAITAB1% + 1
               TETABU1$(2, CAITAB1%) = "S/Nro.: " + NUSERRR$ ' numero de serie
             End If
             '
             '
             'IMPRIME DETALLE DE LOTES REMITIDOS
'             Call DETLOTREM(CIPRE%, U&)
             '
             Call CARLINSEP         'Carga linea de separacion entre codigo y codigo
             '
             If COLTELIB% > 0 Then
               TELIBRE$ = REGLEIDO$(ARCHILIB$, U&)
               FORTELIB$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "DES-MAT/FORMATO")))
               ATELIB% = XVALO(Mid$(FORTELIB$, 2))
               Call FRATEXTO(TELIBRE$, ATELIB%)
               TETABU1$(2, CAITAB1%) = RETEX$(1)
               For KKU1% = 2 To CARETEX%
                 CAITAB1% = CAITAB1% + 1
                 TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
               Next KKU1%
             End If
             '\\\OT-06-0273-WAR-03/07/06///
             ' AQUI AGREGA DESPACHO DE IMPORTACION
             If AGREDESP% = 0 Then
               AGREDESP% = (-1)
               If CONTROL$("FACTURA", "AGREDESP") = "SI" Then AGREDESP% = 1
             End If
             If AGREDESP% = 1 Then
               CIIY% = CODINT%(CDXXS$)
               DESPAI$ = TRIM$(ATRIB1$(CIIY%))
               If DESPAI$ <> "" Then
                 CAITAB1% = CAITAB1% + 1
                 TETABU1$(2, CAITAB1%) = "Despacho: " + DESPAI$
                 TETABU1$(3, CAITAB1%) = "" 'Para que no Imprima Q
               End If
             End If
             '\\\OT-06-0273-WAR-FIN///
             '
           End If
         'End If
       Next U&
       '
       If POVERTI% < 1 Then
         CAITAB1% = CAITAB1% + 1
         For KU& = 1 To ULTREG&("!OBSERVOF")
           TTXX$ = REGLEIDO$("!OBSERVOF", KU&)
           CAITAB1% = CAITAB1% + 1
           TETABU1$(2, CAITAB1%) = RTrim$(TTXX$)
         Next KU&
         Call SETULTREG("!OBSERVOF", 0)
       End If
       '
80     CACOPDOC% = XVALO(FORFINFAC.Text2(1).TEXT)
       PRINTERPORT$ = PORTUREM$
       Call PRINTDOC("FOREMIFA")   ' Remito de Facturacion
       '
       'SE COMENTA PORQUE AHORA EL CHECK DE REPLICA EN EXPEDICION
       'ESTA EN LA PANTALLA DE FINFACTURA07
       '
'       IMPREOR$ = CONTROL("FACTURA", "REPLIREM")
'       PORTURE$ = TRIM$(CONTROL$("ORDESPA", "PUERTOIM"))
       '
      IMPREOR$ = "NO"
      '
      If FORFINFAC.Check8.Value = 1 Then
         IMPREOR$ = "SI"
         PORTURE$ = TRIM$(CONTROL$("ORDESPA", "PUERTOIM"))
      End If
      '
      If IMPREOR$ = "SI" And PORTURE$ <> "" Then
         PRINTERPORT$ = PORTURE$
         Call SETPRINPORT(PORTURE$)   ' IMPRIME UNA COPIA DEL REMITO SI ESTA MARCADA LA
                                       ' OPCION DEL SETUP 'REPLICA EN  DESPACHO'
         Call PRINTDOC("FOREMIFA")    ' Y EL MISMO SE IMPRIME EN LA IMPRESORA CONFIGURADA
         Call SETPRINPORT("RESTORE")  ' COMO IMPRESION DE ORDEN DE DESPACHO (SETUP DE PEDIDOS)
      End If
    End If
    '
End Sub

Private Sub TLABEL28_Change()
'   NOGRA# = XVALO(TLABEL28)
'   IGRAVA0# = XVALO(Total(2)) - NOGRA#
'   If TRIM$(BASIVA(0)) <> "" Then
'       BASIVA(0) = TRIM$(FORMATNUM$(IGRAVA0#, "F13.2"))
'       BASIVA(1) = ""
'     Else
'       BASIVA(0) = ""
'       BASIVA(1) = TRIM$(FORMATNUM$(IGRAVA0#, "F13.2"))
'   End If
'   Total(3).TEXT = CSTRING$(MKD$(XVALO(BASIVA(0)) * XVALO(PORIVA(0).Caption) * APLIVA% / 100), 3, 12, 2, 2)
'   Total(4).TEXT = CSTRING$(MKD$(XVALO(BASIVA(1)) * XVALO(PORIVA(1).Caption) * APLIVA% / 100), 3, 12, 2, 2)
End Sub

Private Sub TLABEL28_GotFocus()
   '
   If VALICLIE% < 1 Then GoTo 99
   If TIFADOC% < 3 Then GoTo 98
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

Private Sub TLABEL28_LostFocus()
   If XVALO(TLABEL28) > XVALO(Total(2)) Then
      Call MENSERR(24, "Importe no Gravado Supera Neto Comprobante")
      TLABEL28 = ""
      Exit Sub
   End If
   TLABEL28 = TRIM$(FORMATNUM$(TLABEL28, "F13.2"))
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
     If XVALO(Total(6)) > 0.005 Then
       If VERLICRE% = 1 Then
         If TIFADOC% < 4 Then ' EXCLUYE CREDITOS
           NCLIE% = XVALO(Text1)
           LMCR# = XVALO(Text7)
           SDEU# = SALDCLI#(NCLIE%)
           CHPA# = CHEPEACRE#(NCLIE%)
           TPED# = XVALO(Total(6).TEXT) * TICAMBX#
           If SDEU# + CHPA# + TPED# > LMCR# Then
             TLIMI$ = TRIM$(CSTRING$(MKD$(LMCR# - SDEU# - CHPA#), 3, 12, 2, 2))
             If TLIMI$ = "" Or XVALO(TLIMI$) < 0 Then TLIMI$ = "0.00"
             Call COMUNI("Imposible Emitir Este Documento.\  \Supera el Límite de Crédito Disponible\para esta Cuenta de $ " + TLIMI$ + ".\  \Informe al Supervisor de Ventas.")
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

'
Private Sub Total_GotFocus(Index As Integer)
   '
   If VALICLIE% < 1 Then GoTo 99
   If Index < 3 Or Index > 4 Then GoTo 98
   If TIFADOC% < 3 Then
     EPAC$ = TRIM$(UCase$(EMPREAC$))
     If InStr(EPAC$, "FRIOLA") < 1 Then GoTo 98
   End If
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
'
Sub GRAMOCABE()
   '
   MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESVENTA", "MODOMONE"))), 8)
   If MODOMONE$ <> "BIMONETA" Then
      If MODOMONE$ <> "MULTIMON" Then
         Exit Sub
      End If
   End If
   '
   FEMI% = FECHANUM(Text9)
   NCLIE% = XVALO(Text1)
   Tipo% = TIFADOC%
   VA% = VARIDOC%
   NRO& = XVALO(Text11.TEXT)
   NUDOCU$ = CoComLar$
   '
   SUCUVEN% = 0
   On Error Resume Next
   SUCUVEN% = XVALO(Label12)
   On Error GoTo 0
   '
   TOTAPESOS# = TOTAFAC#(1, 8)
   TOTADOLAR# = TOTAPESOS# / TICAMBX#
   '
   RAPLI$ = REGBLAN$("CARBIMON")
   ' documento cabeza
   Call REPLA(RAPLI$, MKI$(FEMI%), 1, 2)
   Call REPLA(RAPLI$, MKI$(NCLIE%), 3, 2)
   Call REPLA(RAPLI$, Chr$(Tipo%), 5, 1)
   Call REPLA(RAPLI$, MKI$(VA%), 6, 1)
   Call REPLA(RAPLI$, MKS$(NRO), 7, 4)
   Call REPLA(RAPLI$, MKI$(SUCUVEN%), 11, 2)
   Call REPLA(RAPLI$, NUDOCU$, 13, 18)
   Call REPLA(RAPLI$, MKI$(MONEMI%), 31, 2)
   '
   ' documento aplicado
   Call REPLA(RAPLI$, MKI$(FEMI%), 33, 2)
   Call REPLA(RAPLI$, MKI$(NCLIE%), 35, 2)
   Call REPLA(RAPLI$, Chr$(Tipo%), 37, 1)
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
   MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESVENTA", "MODOMONE"))), 8)
   If MODOMONE$ <> "BIMONETA" Then
      If MODOMONE$ <> "MULTIMON" Then
         Exit Sub
      End If
   End If
   '
   Screen.MousePointer = 11
   NCI% = XVALO(Text1)
   TIDOCAPLI% = 2
   NUDOCAPLI& = XVALO(Mid$(DOCUDIF, 9, 8))
   CANCE# = XVALO(Total(6)) * TICAMONE(MONEMI%)
   TICAMBX# = XVALO(TICAMBIS): If TICAMBX# < 0.05 Then TICAMBX# = 1
   '
   If TIFADOC% = 6 Then
     Call APERBASDAT("CABAN")
     DOCUCA$ = DOCUDIF
     ICAN# = CANCE#
     '
     NumeClie% = NCI%
     DocOrCor$ = ""
     TotOrig# = 0
     IvaOrig# = 0
     '
     SQLX$ = " SELECT * FROM SaDeuPen "
     SQLX$ = SQLX$ + " WHERE NuClien = " & NCI% & " "
     SQLX$ = SQLX$ + " AND CODICOM_LAR = '" & DOCUCA$ & "' "
     SQLX$ = SQLX$ + " AND SALDdebe > 0"
     Dim SADDDEUPEN As ADODB.Recordset
     Set SADDDEUPEN = New ADODB.Recordset
25   On Error Resume Next
     SADDDEUPEN.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
     End If
     On Error GoTo 0
     '
     With SADDDEUPEN
       '.FindFirst "NuClien = " & NCI% & " AND CODICOM_LAR = '" & DOCUCA$ & "' AND SALDdebe > 0"
       'If .NoMatch = False Then
       On Error Resume Next
       If Not (.EOF And .BOF) Then
         Do While (Not .EOF)
850        PARCAN# = ICAN#: If !salddebe < ICAN# Then PARCAN# = !salddebe
           '.Edit
           !salddebe = !salddebe - PARCAN#
           .Update
           '
           DocOrCor$ = !CodiCom_Cor
           DocOrLar$ = !codicom_lar
           TotOrig# = !ibrucomp
           IvaOrig# = 0
           On Error Resume Next
           IvaOrig# = !IIvaComp
           On Error GoTo 0
           '
           ICAN# = ICAN# - PARCAN#
           If ICAN# < 0.005 Then
              Exit Do
  '           .FindNext "CODICOM_LAR = '" & DOCUCA$ & "' AND SALDdebe > 0"
  '           GoTo 850
           End If
           .MoveNext
         Loop
       End If
     End With
     SADDDEUPEN.Close
     Set SADDDEUPEN = Nothing
     '
     ImpApli# = CANCE#
     RefApli$ = "Diferencia de Cambio " + DOCUDIF
     Call GRAPLICOB
     '
     MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESVENTA", "MODOMONE"))), 8)
     If MODOMONE$ = "BIMONETA" Or MODOMONE$ = "MULTIMON" Then
       NCI% = NumeClie%
       RFF$ = RECFILT$("CARBIMON", 2, MKI$(NCI%))
       If Len(RFF$) > 0 Then
         If CANCE# >= 0.005 Then
           ABUSQUE$ = DOCUDIF
           For JKLX& = 1 To Len(RFF$) Step 4
             RECABI& = CVS(Mid$(RFF$, JKLX&, 4))
             YCB$ = REGLEIDO$("CARBIMON", RECABI&)
             If TRIM$(Mid$(YCB$, 13, 18)) = ABUSQUE$ Or TRIM$(Mid$(YCB$, 45, 18)) = ABUSQUE$ Then
               DOCUCABE$ = TRIM$(Mid$(YCB$, 13, 18))
               For JKLY& = 1 To Len(RFF$) Step 4
                 RECABI& = CVS(Mid$(RFF$, JKLY&, 4))
                 YCB$ = REGLEIDO$("CARBIMON", RECABI&)
                 If TRIM$(Mid$(YCB$, 13, 18)) = DOCUCABE$ Then
                     MONEMORI% = CVI(Mid$(YCB$, 31, 2))
                     TICAMBX# = TICAMONE(MONEMORI%)
                     FEDOAPLI% = FECHANUM(Text9)
                     NROO& = XVALO(Mid$(CoComLar$, 11))
                   Call REPLA(YCB$, MKI$(FEDOAPLI%), 33, 2)
                   Call REPLA(YCB$, MKI$(NCI%), 35, 2)
                   Call REPLA(YCB$, Chr$(TIDOAP%), 37, 1)
                   Call REPLA(YCB$, Chr$(VADOAP%), 38, 1)
                   Call REPLA(YCB$, MKS$(NROO&), 39, 4)
                   Call REPLA(YCB$, MKI$(SUC%), 43, 2)
                   Call REPLA(YCB$, CoComLar$, 45, 18)
                   Call REPLA(YCB$, MKI$(1), 63, 2) ' RECIBO SIEMPRE EN PESOS
                   Call REPLA(YCB$, MKD$(0), 65, 8)
                   Call REPLA(YCB$, MKD$(CANCE#), 73, 8)
                   Call REPLA(YCB$, MKD$(0), 81, 8)
                   Call REPLA(YCB$, MKD$(TICAMBX#), 89, 8)
                   Call REPLA(YCB$, MKD$(0), 97, 8)
                   Call REPLA(YCB$, MKD$(0), 105, 8)
                   Call REPLA(YCB$, MKD$(0), 113, 8)
                   '
                   Call REGAPP("CARBIMON", YCB$)
                   GoTo 835
                 End If
               Next JKLY&
             End If
           Next JKLX&
         End If
       End If
835    Call CIERRARCH("CARBIMON")
     End If
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
           Call REPLA(YCB$, MKS$(XVALO(Text11)), 39, 4)
           Call REPLA(YCB$, MKI$(XVALO(Label12)), 43, 2)
             DOCAPLI$ = CoComLar$
           Call REPLA(YCB$, DOCAPLI$, 45, 18)
           Call REPLA(YCB$, MKI$(1), 63, 2) ' RECIBO SIEMPRE EN PESOS
           Call REPLA(YCB$, MKD$(CANCE#), 65, 8)
           Call REPLA(YCB$, MKD$(0), 73, 8)
           Call REPLA(YCB$, MKD$(0), 81, 8)
           Call REPLA(YCB$, MKD$(XVALO(TICAMBIS)), 89, 8)  ' NO PONE TIPO DE CAMBIO
           Call REPLA(YCB$, MKD$(0), 97, 8)
           Call REPLA(YCB$, MKD$(0), 105, 8)
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
Sub GRADEUFA(Tipo%, VA%, NRO, NC%, FFAC%, COPAGI%, TOTFAC#)
   '
   Dim FVI%(6), PLAII%(6), IM#(6), FEX(3)
   '
   LU$ = LUDAT$
   IORIG# = TOTFAC#
   '
   If Tipo% = 0 Then
     If IORIG# < 0 Then
       Tipo% = 7
     End If
   End If
   '
   If Tipo% > 3 Then
     NUPAG% = 1
     FVI%(1) = FFAC%
     IM#(1) = TOTFAC#     ' IORIG#
     GoTo 41
   End If
   '
   ' busca condicion de pago segun COPAGI%
20 Call BUCONPA(LU$, COPAGI%, CoPag$, FREF%, PLA$)
   '
   ' determina vencimientos segun total factura y condicion de pago
30 FF = FFAC%: If FF < 1 Then FF = FECHANUM("01/01/80")
   FFS$ = FECHATEX$(FF)
   PLA$ = Left$(PLA$ + String$(6, 0), 6)

   If (FREF% < 1 Or FREF% > 4) Then
           MENSERR 24, "Mal Fecha Referencia en Cond.Pago Nro." + Str$(COPAGI%)
           FREF% = 1
   End If

   If FREF% = 2 Then FREF% = 1   ' AUN NO FUNCIONA SEMANA DE FACTURA
   If FREF% <> 3 Then GoTo 35
   If XVALO(Left$(FFS$, 2)) <= 15 Then FEBAS$ = "15" + Mid$(FFS$, 3)
   If XVALO(Left$(FFS$, 2)) > 15 Then FEBAS$ = "31" + Mid$(FFS$, 3)
35 If FREF% = 1 Then FEBAS$ = FFS$
   If FREF% = 4 Then FEBAS$ = "31" + Mid$(FFS$, 3)
   '
   For I% = 1 To 6
     FVI%(I%) = 0
   Next I%
   '
   FEFA = FECHANUM(FEBAS$)
   For I% = 1 To 6
     PLAI% = Asc(Mid$(PLA$, I%, 1))
     PLAII%(I%) = PLAI%
     If I% > 1 Then
       If (PLAII%(I%) < PLAII%(I% - 1) Or PLAII%(I%) <= 0) Then
         GoTo 31
       End If
     End If
     '
     FEXU = FEFA + PLAI%
     While DIENTRE(FEFA, FEXU) < PLAI%
       FEXU = DIPOST(FEXU)
     Wend
     '
     FVI%(I%) = FEXU
     NUPAG% = I%
   Next I%

31 FVIJ% = FVI%(NUPAG%)
   If NUPAG% < 2 Then FVIJ% = FFAC%
   '
   ' determina importes segun numero de pagos
40 If NUPAG% < 1 Then NUPAG% = 1
   IM#(1) = (Int(100 * IORIG# / NUPAG% + 0.5)) / 100

   For H% = 2 To NUPAG% - 1
     IM#(H%) = IM#(1)
   Next H%
   IM#(NUPAG%) = IORIG#

   For H% = 1 To NUPAG% - 1
     IM#(NUPAG%) = IM#(NUPAG%) - IM#(H%)
   Next H%
   '
   ICANCE# = IORIG# - TOTFAC#
   For H% = 1 To NUPAG%
     If ICANCE# > 0 Then
       If ICANCE# >= IM#(H%) Then
           ICANCE# = ICANCE# - IM#(H%)
           IM#(H%) = 0
         Else
           IM#(H%) = IM#(H%) - ICANCE#
           ICANCE# = 0
       End If
     End If
   Next H%
   '
41 NuorVen% = 0
   SIGX% = 1: If Tipo% > 3 Then SIGX% = (-1)
   For I% = 1 To NUPAG%
     If Abs(IM#(I%)) >= 0.005 Then
       FechEmi = CVDAT(FFAC%)
       IBruTot# = IORIG#
       NuorVen% = NuorVen% + 1
       FeVenc = CVDAT(FVI%(I%))
       ImVenc# = (Abs(IM#(I%))) * SIGX%
       Call GRADEUPEN
     End If
   Next I%
   FVENCIM = CVDAT(FVI%(1))
   If IsDate(FVENCIM) = False Then
     FVENCIM = ""
   End If
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
   NCLIE% = XVALO(Text1.TEXT)
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
   NROO& = XVALO(Mid$(REMI$, 4))
   FF% = FECHANUM(Text9.TEXT)
   DEPOSALI% = DEPOEXPE%
   '
   PREFIREM$ = CONTROL$("", "PREFIREM")
   If PREFIREM$ = "" Then PREFIREM$ = "0001"
   'ESTO SE AGREGA PORQUE LOS REMITOS DE LAS
   'FACTURAS ELECTRONICAS TIENEN EL PUNTO DE VENTA DE LA FACTURA
   If FACELEC% = 1 And VERIFSALIDA("NOMESS") < 3 Then PREFIREM$ = Label12
   '
   NroRemi$ = TRIM$(Mid$(REMI$, 4))
   While Len(NroRemi$) < 8: NroRemi$ = "0" + NroRemi$: Wend
   NroRemi$ = CMO$ + "." + PREFIREM$ + "-" + NroRemi$
   '
   HO% = HOY(HOS$)
   HP = 100 * XVALO(Left$(Time$, 2)) + XVALO(Mid$(Time$, 4, 2))
   User% = USNBR% Mod 100
   CTRXX$ = MKI$(HO%) + MKI$(Int(HP)) + Chr$(User%) + String$(2, 0)
   '
   RECONSI$ = REGBLAN$("CONSICLI")
   Call REPLA(RECONSI$, String$(12, 0), 109, 12)
   Call REPLA(RECONSI$, MKI$(NCLIE%), 1, 2)
   Call REPLA(RECONSI$, MKI$(FF%), 3, 2)
   Call REPLA(RECONSI$, NroRemi$, 5, 18)
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
         NPEDID& = XVALO(Mid$(X$, 371, 8))
         COD$ = CODEXT$(CI%)
         PRE# = CDbl(CVS(Mid$(X$, 429, 4)))
         CMOV$ = "RT"
         If Mid$(X$, 461, 1) = "R" Then CMOV$ = "*R" ' REPARACION
         DOPRI$ = REMI$
         DOSEC$ = ""
         LOTE$ = ""
         SIG% = (-1)
         '
         'Call MOVISTO(FF%, CI%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, PRE#, "$", NCLIE%, DEPOSALI%, SIG% * CANTAA)
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
   CPAG% = XVALO(TEXT21(1).TEXT)
   CATIVAS$ = Label13.Caption
   '
   TCMP$ = TRIM$(TIPCOM$)
   For KKI% = 0 To 13
     FORFINFAC.Text2(KKI%) = ""
   Next KKI%
   '
   TICOMPROBAN$ = TCMP$
   FORFINFAC.Label2.Caption = TCMP$
   FORFINFAC.Label1(0) = TCMP$ + ":"
   FORFINFAC.Label4.Caption = CATIVAS$
   '
   CODFOR$ = TRIM$(CONTROL$("", "FORFAC-" + Label13.Caption))
   NROCON& = XVALO(Text11.TEXT)
   NCN$ = TRIM$(Str$(NROCON&))
   While Len(NCN$) < 8: NCN$ = "0" + NCN$: Wend
   PREFI$ = Label12.Caption
   NCN$ = PREFI$ + "-" + NCN$
   NUCOMPROBAN$ = NCN$
   FORFINFAC.Label3.Caption = NCN$
   '
5  COPIASFAC% = XVALO(CONTROL$("", "COPIAFAC"))
   COPIASREM% = XVALO(CONTROL$("", "COPIAREM"))
   If COPIASFAC% < 1 Then COPIASFAC% = 1
   If COPIASREM% < 1 Then COPIASREM% = 1
   '
   If TIFADOC% > 2 Then
      'COPIASFAC% = 2: COPIASREM% = 0
      COPIASREM% = 0           ' POR TEMA CANTIDAD DE COPIAS FRIOLATINA
   End If
   '
   If CONSICLI% = 1 Or REPRIFAC% = 1 Then   ' SI REPITE IMPRESION
      COPIASREM% = 0
   End If
   '
   'VALIDACION DE PEDIDOS CON MOTIVO
   X1$ = REGLEIDO$("!CARDEFAC", 1)
   NPED& = XVALO(TRIM$(Mid$(X1$, 371, 16)))
   If NUMOTIVO%(NPED&) > 0 Then
      COPIASFAC% = 0
   End If
   '
   FORFINFAC.Text2(0) = TRIM$(Str$(COPIASFAC%))
   FORFINFAC.Text2(1) = TRIM$(Str$(COPIASREM%))
   FORFINFAC.Text2(2) = " "
'   'VALIDACION DE PEDIDOS CON MOTIVO
'   X1$ = REGLEIDO$("!CARDEFAC", 1)
'   NPED& = XVALO(TRIM$(Mid$(X1$, 371, 16)))
'   If NUMOTIVO%(NPED&) > 0 Then FORFINFAC.Text2(0) = ""
'   '
   'FORFINFAC.Text2(3) = TRIM$(Str$(Int(TICAMBX# * XVALO(Total(2).TEXT)))) + ".="
   '\\\OT-05-0492-WAR-11/07/05///
     COVADE = XVALO(CONTROL("", "COEVADEC")) / 100 'COEFICIENTE DE VALOR AGREGADO
     If COVADE < 0.01 Then COVADE = 1
     FORFINFAC.Text2(3) = TRIM$(Str$(Int(COVADE * TICAMBX# * XVALO(Total(2).TEXT)))) + ".="
   '\\\OT-05-0492-WAR-FIN
   FORFINFAC.Label5.Caption = ""
   For U& = 0 To 8
      FORFINFAC.Text1(U&).TEXT = ""
   Next U&
   '\\\OT-06-0114-WAR-03/04/06///
   FORFINFAC.Text4.TEXT = ""
   FORFINFAC.Text5.TEXT = ""
   'LIMPIO LOS TEXT DE DATOS DE LA EXPO
   For JX% = 4 To 11
    FORFINFAC.Text2(JX%) = ""
   Next JX%
   '\\\OT-06-0114-WAR-FIN///
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   TTXX$ = ""
   '
   FORFINFAC.Label6.Enabled = False
   FORFINFAC.CLAREDOL = "0"
   If MONEMI% = 1 Then
     FORFINFAC.Label6.Enabled = True
   End If
   '
   FORFINFAC.Label6.Visible = False
   FORFINFAC.Text3 = ""
   FORFINFAC.Frame5.Visible = False
   '\\\OT-06-0125-WAR-03/04/06///
   FORFINFAC.Height = 6690
   'SI FACTURA DE EXPO => SE MUESTRA DATOS DE EXPORTACIÓN
   NC1% = XVALO(FORFASQL09.Text1.TEXT)
   If PIVACLI%(NC1%) = 5 Or PIVACLI%(NC1%) = 7 Then
     FORFINFAC.Height = FORFINFAC.Height + 1935 '(8600)
   End If
   'CACULA EL PESO EN KILOS NETO
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
   FORFINFAC.Text2(11) = FORMATNUM$(PT, "F8.1")
   '\\\OT-06-0125-WAR-FIN///
   '18/04/07 (OT-07-0162) 'PARA VER LA COTIZACION DEL DOLAR CUANDO EMITIMOS UNA FACTURA EN PESOS
   If MONEMI% <= 1 Then
      NLIPRE% = XVALO(FORFASQL09.TEXT21(0))
      If MONELISTA%(NLIPRE%) = 2 Then
         TTXX$ = "Emitido en " + DEMONECO$(MONEMI%) + " a cotización " + DEMONECO$(2) + " de " + TRIM$(FORMATNUM$(TICAMONE(2), "F9.4")) + ". "
          FORFINFAC.Text4 = TTXX$
      End If
   End If
   '18/04/07 (FIN)
   '
'   FORFINFAC.Label11 = ""
'   If FACELEC% = 1 Then
'      If VERIFSALIDA%("NOMESS") > 0 Then
'      If Label13 = "E" Then
'         FORFINFAC.Command1.Top = 5655
'         FORFINFAC.Command1.Height = FORFINFAC.Command2.Height + FORFINFAC.Command2.Top - FORFINFAC.Command1.Top
'         FORFINFAC.Command1.Enabled = False
'         FORFINFAC.Command2.Enabled = True
'
'       End If
'     End If
'   End If
   '
   If CONTROL("FACTURA", "VERIVEN") = "SI" Then
    If XVALO(TEXT21(2).TEXT) < 1 Then ' VALIDACION DE VENDEDOR
     FECH = Text9
     TIPFACT$ = "FC-"
     If Label7 = "N.Débito:" Then TIPFACT$ = "ND-"
     If Label7 = "N.Crédito:" Then TIPFACT$ = "NC-"
     
     NFCZ$ = TRIM$(Text11): While Len(NFCZ$) < 8: NFCZ$ = "0" + NFCZ$: Wend

     NFACT$ = TIPFACT$ & Label12 & "-" & NFCZ$ & "-" & Label13
     NC% = XVALO(Text1)
     RazonSoc$ = TRIM$(Text2)
     Call REPLA(VDEF$, MKI$(CVINT(FECH)), 1, 2) ' INFORMACION PARA EL OBJETO
     Call REPLA(VDEF$, NFACT$, 3, 18)
     Call REPLA(VDEF$, MKI$(NC%), 21, 2)
     Call REPLA(VDEF$, RazonSoc$, 23, 30)
    '
30   OBX$ = OBJPLA$("ACTUVENDE", VDEF$) ' ABRE EL OBJETO
     If OBX$ = "" Then
        Call COMUNI("Debe Asignar un Vendedor a La Factura")
        GoTo 30
     End If
     '
     On Error Resume Next ' POR SI INGRESA UN VENDEDOR > 255
     VDDOR% = CVI(Mid$(OBX$, 53, 2)) ' VALIDA EL NUMERO DE VENDEDOR
     If VDDOR% < 1 Or ECOARCH$("VENDEDOR", Chr$(VDDOR%)) = "" Then
       On Error GoTo 0
       Call COMUNI("Vendedor No Válido")
       VDEF$ = OBX$
       GoTo 30
     End If
     On Error GoTo 0
     TEXT21(2) = VDDOR%
    End If
   End If
   '
   'Limpio campos de gastos de exportacion
   FORFINFAC.Text2(15) = ""
   FORFINFAC.Text10 = ""
   '
10 FORFINFAC.Show 1
   Refresh
   '
   If FORFINFAC.Label5.Caption <> "APRO" Then
      If COMALTER%("Realmente Quiere CANCELAR este Documento\\No, Continuar\Si, Cancelar") <> 2 Then GoTo 10
      OKX% = (-1)
      Exit Sub
   End If
   '
   FACELECOMUN% = XVALO(CONTROL("FACTURA", "PUVENELE"))
   FACELEBIENCAP% = XVALO(CONTROL("FACTURA", "PUVENCAP"))
   FACELEEXPO% = XVALO(CONTROL("FACTURA", "PUVEFAEX"))
   PTOVTA% = XVALO(Label12.Caption)
   '
   If FACELEC% > 0 And PTOVTA% = FACELECOMUN% Then Call SOLICITARCAE1
   '
   If FACELEC% > 0 And PTOVTA% = FACELEBIENCAP% Then Call SOLICITARCAE_WSBFE1
   '
   If FACELEC% > 0 And PTOVTA% = FACELEEXPO% Then
        DATFAEX07.Show 1
        If DATFAEX07.Label17 = "OK" Then
           Call SOLICITARCAE_WSFEX1
        End If
   End If
   '
   COPIASFAC% = XVALO(FORFINFAC.Text2(0))
   If COPIASFAC% < 0 Or COPIASFAC% > 9 Then GoTo 5
   COPIASREM% = XVALO(FORFINFAC.Text2(1))
   If COPIASREM% > 9 Then GoTo 5
   If REPRIFAC% = 1 Then COPIASREM% = 0
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
     ElseIf Index = 3 Then
      ' calcula importe no gravado
       If Label6.Caption = "(T.Libre)" Then
         BASIVA(0) = TRIM$(FORMATNUM(XVALO(Total(3)) / ((XVALO(PORIVA(0))) / 100), "F13.2"))
       End If
     ElseIf Index = 4 Then
      ' calcula importe no gravado
       If Label6.Caption = "(T.Libre)" Then
         BASIVA(1) = TRIM$(FORMATNUM(XVALO(Total(4)) / ((XVALO(PORIVA(1))) / 100), "F13.2"))
       End If
     Else
       On Error Resume Next
       Command14.SetFocus
       On Error GoTo 0
     End If
   End If
End Sub

Private Sub Total_LostFocus(Index As Integer)
   'If Index = 0 Then
     TTT1 = XVALO(Total(Index).TEXT)
     Total(Index) = CSTRING$(MKS$(TTT1), 3, 12, 2, 2)
   'End If
End Sub

Sub CARDETFAC()
   '
   Static AGRUPAENTREGA% 'ASIGNA A AGRUPAENTREGA% SEGUN LA CONFIGURACION DEL SETUP
   If AGRUPAENTREGA% = 0 Then
     AGRUPAENTREGA% = (-1)
     If CONTROL$("FACTUREM", "AGRUPEDI") = "TRANSUCU" Then AGRUPAENTREGA% = 1
   End If

   HOII% = HOY(HOS$)
   '
   ATRI$ = "/NC/ARCHECO(2)=" + ARCHIPRE$
   ATRI$ = ATRI$ + "/TITUPAN=Detalle de Facturación (Mat.Codificado)"
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Frame1.Top + Picture3.Top))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture3.Left - 50))
   '
   If FACTLISTAEMPAQUE& > 0 Then
     ATRI$ = ATRI$ + "/POSEDIT(1)=2"
     ATRI$ = ATRI$ + "/POSEDIT(2)=2"
     ATRI$ = ATRI$ + "/POSEDIT(3)=2"
     ATRI$ = ATRI$ + "/POSEDIT(4)=2"
     ATRI$ = ATRI$ + "/POSEDIT(5)=2"
     ATRI$ = ATRI$ + "/POINDEL=NO" ' DESHABILITA EL INSERTAR Y EL DEL
   End If
   '
10 Timer1.Enabled = True
   PRF$ = "FACONGUI"
   If CONSICLI% = 0 Then
      PRF$ = "FACDOSI"
      EPAC$ = TRIM$(UCase$(EMPREAC$))
      If InStr(EPAC$, "FRIOLA") > 0 Then
        PRF$ = "FACFRIOL"
      End If
      If MODESCU% = 1 Then PRF$ = "FACPORDE"
   End If
   If MODODOC% = 1 Then

      PRF$ = "FAC-LIBR"
      ATRI$ = "/NC"
      ATRI$ = ATRI$ + "/TITUPAN=Detalle de Facturación (Texto Libre)"
      ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Frame1.Top + Picture3.Top))
      ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture3.Left + 50))
   End If
   List1.Clear
   '
   If EsPorPedido& < 1 Then
        XXX% = VENTABU%(PRF$ + ATRI$)
      Else
        XXX% = 0                   ' SI ES POR PEDIDO
   End If
   '
   Call CALTOFAC
   '
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
   If CONSICLI% < 1 Then
     Call ASIGNALOTES("!CARDEFAC", RESUL$)    ' SI NO EXISTEN ARTICULOS TRAZABLE NO INGRESA
     If RESUL$ <> "APRO" Then GoTo 10
   End If
   '
   If InStr(EMPREAC$, "MEDITEA") > 0 And CONSICLI% = 0 Then ' EXCLUYE LA FACTURACION DE CONSIGNACIONES
      Call CARNUMSER(OK%)                    ' CARGA NUMEROS DE SERIE
      If OK% <> 1 Then GoTo 10               ' CANCELA EN VENTANA N/SERIE VUELVE ANTERIOR
   End If
   '

   Screen.MousePointer = 11   ' Registrando FACTURA ...
   List1.Clear
   FIN& = ULTREG&("!CARDEFAC")
   TXPEDI$ = "": Label28 = ""
   PORLETRA% = 0:  If CONTROL$("ADMIPED", "MOCOMIS") = "LETRA" Then PORLETRA% = 1
   
   '
   For U& = 1 To FIN&
      Y$ = REGLEIDO$("!CARDEFAC", U&)
      If MODODOC% < 1 Then
           CI% = CVI(Mid$(Y$, 5, 2))
           '
           CAN = CVS(Mid$(Y$, 1, 4))
           If CAN > 999999 Then
               Call MENSERR(24, "Cantidad Excesiva")
               GoTo 10
             ElseIf Abs(CAN) < 0.1 Then
               If CI% > 0 Then
                 Call MENSERR(24, "Falta Cantidad")
                 Call REPLA(Y$, MKS$(0), 1, 4)
                 Call GRAREG("!CARDEFAC", Y$, U&)
                 GoTo 10
               End If
           End If
           '
           If Abs(CAN - REDON_INDIVIS(CI%, CDbl(CAN), "VENTA")) >= 0.05 Then
              TTXX$ = "Cantidad Ajustada a Lotes Indivisibles\en Código '" + CODEXT$(CI%) + "'"
              Call MENSERR(24, TTXX$)
              Call REPLA(Y$, MKS$(REDON_INDIVIS(CI%, CDbl(CAN), "VENTA")), 1, 4)
              Call GRAREG("!CARDEFAC", Y$, U&)
              GoTo 10
           End If
           '
           NPEDID& = XVALO(Mid$(Y$, 371, 8))
           If NPEDID& > 0 Then
              TYPEDI$ = FORMATNUM$(NPEDID&, "F10.0")
              If InStr(TXPEDI$, TYPEDI$) < 1 Then
                 TXPEDI$ = TXPEDI$ + TYPEDI$
                 If PORLETRA% > 0 Then
                    If TRIM$(Label28) = "" Then
                      LETRAPE$ = Chr$(64 + XVALO(VALOBADA("PEDENCA", "PorComVend", "NroPed = " & NPEDID&)))
                      If Asc(LETRAPE$) > 64 Then Label28 = LETRAPE$
                    End If
                 End If
              End If
           End If
           '
         Else     ' FACTURA DE TEXTO LIBRE
           CAN = CVS(Mid$(Y$, 1, 4))
           If CAN < 0.1 Then
              CAN = 1
              Call REPLA(Y$, MKS$(1), 1, 4)
              Call GRAREG("!CARDEFAC", Y$, U&)
              Call MENSERR(24, "Faltaba Cantidad - Ajustada a '1,0'")
              GoTo 10
           End If
           PREUNI = CVS(Mid$(X$, 429, 4))
           TOT# = CVD(Mid$(X$, 433, 8))
           If Abs(CAN * PREUNI - TOT#) >= 0.005 Then
              TOT# = ROUND(CAN * PREUNI)
              Call REPLA(Y$, MKD$(TOT#), 433, 8)
              Call GRAREG("!CARDEFAC", Y$, U&)
              Call MENSERR(24, "Importe Incorrecto - Ajustado")
              GoTo 10
           End If
      End If
   Next U&
   '
   Call CARLISFAC
   '
   MALCOMBINA% = 0
   If VERIPEDMO% < 1 Then MALCOMBINA% = 1   ' DISTINTO MOTIVO DE PEDIDO

   If AGRUPAENTREGA% > 0 Then ' VALIDA SI ESTA SELECCIONADA LA OPCION DE AGRUPACION DEL SETUP
    If Len(TXPEDI$) > 0 Then
       Dim PEDDDENCA As ADODB.Recordset
       'MALCOMBINA% = 0
       LIPENOCO07.List1.Clear
       PDXA$ = "": LPRXA$ = "": CPGXA$ = "": DMEXA$ = "": LMEXA$ = "": NTREXA$ = "": TREXA$ = "": NSUENA$ = "": NMOEMA$ = "": VENDEA% = (-1)
       For KKI% = 1 To Len(TXPEDI$) Step 10
         TENCA$ = ""
         NPEDID& = XVALO(TRIM$(Mid$(TXPEDI$, KKI%, 10)))
         SQLX$ = " SELECT * FROM PedEnca"
         SQLX$ = SQLX$ + " WHERE NroPed = " & NPEDID& & ""
         Set PEDDDENCA = FLEXCONN.Execute(FILTSQL$(SQLX$))
         '
         With PEDDDENCA
           If Not (.EOF And .BOF) Then
             '
             NPEDI$ = (TRIM$(SAFETEXT$(!NROPED)))
             TENCA$ = AJUSTD$(NPEDI$, 6)
             '
             PDX$ = TRIM$(SAFETEXT$(!TDescApli))
             If PDXA$ <> "" Then
                If PDX$ <> PDXA$ Then
                    MALCOMBINA% = 1
'                  Call MENSERR(24, "No Válido Combinar Pedidos Con Distinto Descuento.")
'                  GoTo 10
                End If
             End If
             Call REPLA(TENCA$, AJUSTI$(PDX$, 12), 103, 12)
             PDXA$ = PDX$
             '
'             LPRX$ = TRIM$(Str$(XVALO(!LISTPRE)))
'             If LPRXA$ <> "" Then
'                If LPRX$ <> LPRXA$ Then
'                  Call MENSERR(24, "No Válido Combinar Pedidos Con Distinta Lista de Precios.")
'                  GoTo 10
'                End If
'             End If
'             LPRXA$ = LPRX$
             '
             CPGX$ = TRIM$(Str$(XVALO(!CondPag)))
             If CPGXA$ <> "" Then
                If CPGX$ <> CPGXA$ Then
                     MALCOMBINA% = 1
'                  Call MENSERR(24, "No Válido Combinar Pedidos Con Distinta Condicion de Pago.")
'                  GoTo 10
                End If
             End If
             Call REPLA(TENCA$, AJUSTD$(CPGX$, 5), 95, 5)
             CPGXA$ = CPGX$
             '
             DMEX$ = TRIM$(SAFETEXT$(!DOMIENT))
             If DMEXA$ <> "" Then
                If DMEX$ <> DMEXA$ Then
                   MALCOMBINA% = 1
'                  Call MENSERR(24, "No Válido Combinar Pedidos Con Distinto Domicilio de Entrega.")
'                  GoTo 10
                End If
             End If
             Call REPLA(TENCA$, AJUSTI$(DMEX$, 25), 9, 37)
             DMEXA$ = DMEX$
             '
             LMEX$ = TRIM$(SAFETEXT$(!LOCAENT))
             If LMEXA$ <> "" Then
                If LMEX$ <> LMEXA$ Then
                   MALCOMBINA% = 1
'                  Call MENSERR(24, "No Válido Combinar Pedidos Con Distinta Localidad de Entrega.")
'                  GoTo 10
                End If
             End If
             Call REPLA(TENCA$, AJUSTI$(LMEX$, 22), 37, 22)
             LMEXA$ = LMEX$
             '
             NTREX$ = TRIM$(Str$(XVALO(!NTranspo)))
             If NTREXA$ <> "" Then
                If NTREX$ <> NTREXA$ Then
                   MALCOMBINA% = 1
'                  Call MENSERR(24, "No Válido Combinar Pedidos Con Distinto Numero de Transporte.")
'                  GoTo 10
                End If
             End If
             Call REPLA(TENCA$, AJUSTD$(NTREX$, 4), 61, 4)
             NTREXA$ = NTREX$
             '
             TREX$ = TRIM$(SAFETEXT$(!TranspEnt))
             If TREXA$ <> "" Then
                If TREX$ <> TREXA$ Then
                   MALCOMBINA% = 1
'                  Call MENSERR(24, "No Válido Combinar Pedidos Con Distinto Transporte.")
'                  GoTo 10
                End If
             End If
             Call REPLA(TENCA$, AJUSTI$(TREX$, 22), 67, 30)
             TREXA$ = TREX$
             '
             VENDE% = XVALO(!VENDED)
             If VENDEA% > (-1) Then
                 If VENDE% <> VENDEA% Then MALCOMBINA% = 1
             End If
             VENDEA% = VENDE%
             
             '
             
             NOCAMONE% = 1            ' BANDERA PARA NO CAMBIAR LA MONEDA DE EMISION
             PORDESCU.TEXT = PDXA$
             If LPRXA$ <> "" Then TEXT21(0).TEXT = LPRXA$
             TEXT21(1).TEXT = CPGXA$
             NOCAMONE% = 0
             '
             If DMEXA$ <> "" Then LENTREGA(0).TEXT = DMEXA$
             If LMEXA$ <> "" Then LENTREGA(1).TEXT = LMEXA$
               If NTREXA$ <> "" Then Text19.TEXT = NTREXA$
             If TREXA$ <> "" Then LENTREGA(2).TEXT = TREXA$
             If VENDEA% > 0 Then TEXT21(2).TEXT = VENDEA%
           End If
         End With
         LIPENOCO07.List1.AddItem (TENCA$)
       Next KKI%
       '
       If MALCOMBINA% > 0 Then
          If COMALTER%("Los Pedidos Seleccionados Tienen Condiciones Comerciales\que no Pueden Combinarse.\\Ver Detalle\Corregir Seleccion") <> 2 Then
             Screen.MousePointer = 1
             LIPENOCO07.Show 1
             GoTo 10
          Else
             GoTo 10
          End If
       End If
       '
'       If CONTROL("FACTURA", "VERIVEN") = "SI" Then ' CONTROL DE VENDEDOR
'          If VENDEA% < 1 Then
'              Call COMUNI("Se Han Encontrado Pedidos Sin Vendedor\Actualice El Vendedor en el Pedido Y Vuelva a Facturar")
'              GoTo 10
'          End If
'       End If
       '
    End If
   End If
   '
   'Se comenta la condicion para que presente la la seleccion de lotes
   'independientemente de la configuracion de facturacion
   'If TRIM$(CONTROL("FACTURA", "PREFACTU")) = "SELELOTE" Then
   'If InStr(EPAC$, "PAPELERA NICARAGUA") > 0 Then
      Screen.MousePointer = 1
      '
      FIN& = ULTREG&("!CARDEFAC")
      For UI& = 1 To FIN&
        XX$ = REGLEIDO$("!CARDEFAC", UI&)
        CCII% = CVI(Mid$(XX$, 5, 2))
        If ESTRAZABLE%(CCII%) = 1 Then
            For I = 0 To ASIGLOT07.List1.ListCount - 1
                YY$ = ASIGLOT07.List1.List(I)
                CANTASI = XVALO(Mid$(YY$, 64, 10))
                REGIST& = XVALO(Mid$(YY$, 80, 5))
                If UI& = REGIST& And CANTASI > 0 Then
                    Call REPLA(XX$, MKS$(CANTASI), 1, 4)
                    Call GRAREG("!CARDEFAC", XX$, UI&)
                End If
            Next I
        End If
      Next UI&

' AQUI AJUSTAR CANTIDADES E IMPORTES
      Call CALTOFAC
      Call CARLISFAC
'   End If
'   End If
   '
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
      '
      If MAREPA$ = "R" Then
        PORDESCU = ""
      End If
      '
      CAN = CVS(Mid$(Y$, 1, 4))
      '
      If CAN <= 9999 Then
        Call REPLA(TTXX$, CSTRING$(MKS$(CAN), 3, 6, 1, 2), 1, 6)
      Else
        Call REPLA(TTXX$, CSTRING$(MKS$(CAN), 3, 6, 0, 2), 1, 6)
      End If
      '
      CANDEC% = DECILISTA%(LIPRACTI%)
      Call REPLA(TTXX$, CODEXT$(CI%), 8, 15)
      Call REPLA(TTXX$, Mid$(Y$, 7, 64), 24, 26)
      PRELI = CVS(Mid$(Y$, 421, 4))
      Call REPLA(TTXX$, CSTRING$(MKS$(PRELI), 3, 8, CANDEC%, 2), 50, 8)
      PORDESCUIN = CVS(Mid$(Y$, 425, 4))
      Call REPLA(TTXX$, CSTRING$(MKS$(PORDESCUIN), 3, 5, 1, 2), 59, 5)
      PREUNI = CVS(Mid$(Y$, 429, 4))
      Call REPLA(TTXX$, CSTRING$(MKS$(PREUNI), 3, 8, CANDEC%, 2), 64, 8)
      TOT# = CVD(Mid$(Y$, 433, 8))
      Call REPLA(TTXX$, CSTRING$(MKD$(TOT#), 3, 10, 2, 2), 74, 10)
      Call REPLA(TTXX$, Mid$(Y$, 461, 2), 86, 2) ' MARCA REPARACION
      List1.AddItem TTXX$
      '
      If GRUPIVA%(CI%) > 1 Then
         Call MENSERR(24, "Grupo de IVA no Procesable\para Código '" + TRIM$(CODEXT$(CI%)) + "'.\   \Se Procesará a la Tasa Básica del " + TRIM$(FORMATNUM$(ALIVA, "F5.2")) + " %.")
      End If
      '
   Next U&
   List1.Refresh
   '
End Sub
'
Sub CALTOFAC()
   '
   Static CLIA$, RNI%
   If Text1.TEXT <> CLIA$ Then
     CLIA$ = Text1.TEXT: NCLIE% = XVALO(CLIA$)
     RNI% = 0: APLIVA% = 1
     If NCLIE% > 0 Then
       PPVI% = PIVACLI%(NCLIE%)
       If PPVI% = 2 Then RNI% = 1
       If PPVI% = 5 Or PPVI% = 7 Then APLIVA% = 0
     End If
   End If
   '
   CLIACT% = XVALO(Text1)
   '
   TASA1 = XVALO(PORIVA(0).Caption)
   TASA2 = XVALO(PORIVA(1).Caption)
   TASA3 = XVALO(PORIVA(2).Caption)
   SUNEFAC# = 0: SUNEFA1# = 0: SUNEFA2# = 0
   '
   Label18 = TRIM$(Str$(ULTREG&("!CARDEFAC")))
   SUIVA# = 0: SUIVARNI# = 0: TOIVA# = 0: TOTOTA# = 0
   SUMACAN = 0
   For U& = 1 To ULTREG&("!CARDEFAC")
     '
     X$ = REGLEIDO$("!CARDEFAC", U&)
     CIXXI% = CVI(Mid$(X$, 5, 2))
     CAN = CVS(Mid$(X$, 1, 4))
     SUMACAN = SUMACAN + CAN
     PORDESCUIN = CVS(Mid$(X$, 425, 4))
     If PORDESCUIN < 0 Or PORDESCUIN > 100 Then PORDESCUIN = 0
     '
     If Abs(CAN) > 0.05 Or MODODOC% > 0 Then
       TASARE = TASA1
       If TASACLICK >= 0.005 Then TASARE = TASACLICK
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
       If CAN = 1 Then
           TOT# = CDbl(PREUNI)
         Else
           TOT# = CDbl(Int(10 * CAN + 0.5)) / 10 * (Int(10000 * CDbl(PREUNI) + 0.5)) / 10000
       End If
       '
       If MODODOC% = 1 Then
          PREUNI = CVS(Mid$(X$, 429, 4))
          TOT# = CVD(Mid$(X$, 433, 8))
       End If
       '
       Call REPLA(X$, MKD$(TOT#), 433, 8)
       SUNEFAC# = SUNEFAC# + TOT#
       '
       TASARE = CVS(Mid$(X$, 441, 4))
       IVAITEM1# = (Int(TOT# * TASARE * APLIVA% + 0.5)) / 100
       Call REPLA(X$, MKD$(IVAITEM1#), 445, 8)
       Call REPLA(X$, MKD$(TOT# + IVAITEM1#), 453, 8)
       Call GRAREG("!CARDEFAC", X$, U&)
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
   Label25 = FORMATNUM(SUMACAN, "F8.1")
   '
   Total(0).TEXT = CSTRING$(MKD$(SUNEFAC#), 3, 12, 2, 2)
   Total(0).Refresh
   '
   DESC$ = TRIM$(PORDESCU.TEXT)
   PODES = COEFDESCU(DESC$)
   TTT1# = XVALO(Total(0).TEXT) * PODES / 100
   Total(1).TEXT = CSTRING$(MKD$(TTT1#), 3, 12, 2, 2)
   '
   TTT2# = XVALO(Total(0).TEXT) - XVALO(Total(1).TEXT)
   '
   ' CALCULO DE EMBALAJE
   PEMBA = POREMBAL(CLIACT%)
   Ltota(10) = "Importe Total Neto"
   IEMBALA = "0.00"
   If TIFADOC% < 3 Then ' SOLO PARA FACTURAS, NO N/CRED Y N/DEB
     If XVALO(PEMBA) > 0 Then 'se agrega xvalo por que devolvia "" y daba error 13
       IEMBALA = TRIM$(FORMATNUM$(PEMBA * TTT2# / 100, "F12.2"))
       Ltota(10) = "Tot.Neto (Inc." + TRIM$(IEMBALA) + " Embal.)"
       TTT2# = TTT2# + XVALO(IEMBALA)
     End If
   End If
   '
   Total(2).TEXT = CSTRING$(MKD$(TTT2#), 3, 12, 2, 2)
   TLABEL28 = ""
   '
   TTTX# = SUNEFA2# * (100 - PODES) / 100
   BASIVA(1) = TRIM$(CSTRING$(MKD$(TTTX#), 3, 12, 2, 2))
   Total(4).TEXT = CSTRING$(MKD$(XVALO(BASIVA(1)) * XVALO(PORIVA(1).Caption) * APLIVA% / 100), 3, 12, 2, 2)
   '
   TTTX# = (XVALO(Total(2).TEXT) - XVALO(BASIVA(1)))
   BASIVA(0) = TRIM$(CSTRING$(MKD$(TTTX#), 3, 12, 2, 2))
   Total(3).TEXT = CSTRING$(MKD$(XVALO(BASIVA(0)) * XVALO(PORIVA(0).Caption) * APLIVA% / 100), 3, 12, 2, 2)
   '
   If APLIVA% < 1 Then
     TLABEL28 = Total(2).TEXT
     BASIVA(0) = ""
     BASIVA(1) = ""
   End If
   '
   Call CALTOFA2
   '
   
End Sub
'
Sub CALTOFA2()
   '
   Static CLIA$
   If Text1.TEXT <> CLIA$ Then
     CLIA$ = Text1.TEXT: NCLIE% = XVALO(CLIA$)
     RNI% = 0
     If NCLIE% > 0 Then
       PPVI% = PIVACLI%(NCLIE%)
       If PPVI% = 2 Then RNI% = 1
     End If
   End If
   '
   TTTX# = (XVALO(Total(3).TEXT) + XVALO(Total(4).TEXT)) * RNI%
   BASIVA(2) = TRIM$(CSTRING$(MKD$(TTTX#), 3, 12, 2, 2))
   Total(5).TEXT = CSTRING$(MKD$(XVALO(BASIVA(2)) * XVALO(PORIVA(2).Caption) / 100), 3, 12, 2, 2)
   '
   TTTX# = 0: For KKI% = 2 To 5
                 TTTX# = TTTX# + XVALO(Total(KKI%).TEXT)
              Next KKI%
              TTTX# = TTTX# + XVALO(Text15)
              TTTX# = TTTX# + XVALO(Text17)
   Total(6).TEXT = CSTRING$(MKD$(TTTX#), 3, 12, 2, 2)
   '
End Sub

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

Sub GRACONSI()
   'jandy
   'ver ernesto LISTO 22/08
   CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
   If Len(CMO$) <> 2 Then
     CMO$ = "RT"
   End If
   '
   NUFACLAR$ = CoComLar$
   NUFACORT$ = CoComCor$
   '
   NC1% = XVALO(Text1.TEXT)
   Call APERBASDAT("STOCKS")
   '
   ' ESTO ES NUEVO
'   SQLX$ = " SELECT * FROM MOVISTO "
'   SQLX$ = SQLX$ + "WHERE Nume_Clie = " & NC1% & " "
'   SQLX$ = SQLX$ + " AND (CodiMovi = '" & CMO$ & "' "
'   SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
'   SQLX$ = SQLX$ + " AND DoPriLar <> '' "
'   SQLX$ = SQLX$ + " AND ((ISNULL(DoSecLar) = true) or (DoSecLar = '') OR (DoSecLar LIKE 'P/F*')) "
'   SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) >= " & DESELEC# & " "
'   SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) <= " & HASELEC# & " "
   '
   
 '  Set CONSITMP = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
   '
   FIN& = ULTREG&("!CARDEFAC")
   '
   DESCX$ = PORDESCU
   PODES = COEFDESCU(DESCX$)
   COEFPRE = (100 - PODES) / 100
   '
   TTT1# = XVALO(Total(0).TEXT) * PODES / 100
   For I& = 1 To FIN&
     Call TRACE(20, I&, FIN&)
     X$ = REGLEIDO$("!CARDEFAC", I&)
     CANTAA = CVS(Left$(X$, 4))
     If Abs(CANTAA) > 0 Then
       CI% = CVI(Mid$(X$, 5, 2))
       DOPRIC$ = TRIM$(Mid$(X$, 387, 16))
       NORIT% = CVI(Mid$(X$, 503, 2))
       PREUNI = CVS(Mid$(X$, 429, 4)) * COEFPRE
       NPEDCAR& = XVALO(Mid$(X$, 371, 16))
       ESREPA$ = Mid$(X$, 461, 1)
       '
       If CI% > 0 Then
         '
         ORIMOVI$ = "MOVISTO"
         If ESREPA$ = "R" Then ORIMOVI$ = "MOVIREPA"
         '
         SQLX$ = " SELECT * FROM " + ORIMOVI$ + " "
         SQLX$ = SQLX$ + "WHERE Nume_Clie = " & NC1% & " "
         SQLX$ = SQLX$ + " AND (CodiMovi = '" & CMO$ & "' "
         SQLX$ = SQLX$ + " OR CodiMovi =  'DV' "
         SQLX$ = SQLX$ + " OR CodiMovi =  '*R') "
         SQLX$ = SQLX$ + " AND DoPriLar <> '' "
         SQLX$ = SQLX$ + " AND ((DoSecLar is null) or (DoSecLar = '') OR (DoSecLar LIKE 'P/F%')) "
         SQLX$ = SQLX$ + " AND FechMov >= '" & DESELEC & "' "
         SQLX$ = SQLX$ + " AND FechMov <= '" & HASELEC & "' "
         SQLX$ = SQLX$ + " AND DoPriCor = '" & DOPRIC$ & "' "
         SQLX$ = SQLX$ + " AND Cod_Inte = " & CI% & " "
         SQLX$ = SQLX$ + " AND NuorItem = " & NORIT% & " "
         ' SQLX$ = SQLX$ + " AND SIGN(Cantsalid - CantIngre)= " & Sgn(CANTAA) & " "
         Dim CONSITMP As ADODB.Recordset
         Set CONSITMP = New ADODB.Recordset
25       On Error Resume Next
         CONSITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
         If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 25
         End If
         On Error GoTo 0
         '
         With CONSITMP   ' Movisto
             If (.EOF And .BOF) Then
                Call MENSERR(24, "Imposible Continuar-Incoherencia")
                Exit Sub
             End If
             '
             ' ESTO REEMPLAZA LA CONDICION SQL POR PROBLEMA GUIDI 4/3/2009
26           If Sgn(XVALO(!Cantsalid) - XVALO(!CantIngre)) <> Sgn(CANTAA) Then
                If .EOF Then GoTo 29
                .MoveNext: GoTo 26
             End If
             '
             If IsNull(!NuPedCli) = False Then NPEDID& = !NuPedCli
             '
             !DoSecCor = NUFACORT$
             !DoSecLar = NUFACLAR$
             !VALOUNIT = PREUNI
             !MoneVal = MONEMI%         '  moneda de emision
             If MONEMI% <= 1 Then
                 !TIPOCAM = TICAMONE(2) '  siempre cotizacion del dolar
               Else
                 !TIPOCAM = TICAMONE(MONEMI%)
             End If
             If CANTAA > 0 Then
                 !Cantsalid = CANTAA
                 !CantIngre = 0
               Else
                 !Cantsalid = 0
                 !CantIngre = Abs(CANTAA)
             End If
             If NPEDID& < 1 Then
               !NuPedCli = NPEDCAR&
             End If
             NUMSERIE$ = SAFETEXT$(!IdenLote)
             .Update
             GoTo 29
         End With
       End If
       '

       FECHA7% = FECHANUM(Text9)
       VUNI7# = PREUNI
       Call MOVISTOK(FECHA7%, CI%, "", CMO$, NUFACLAR$, NUFACORT$, NUFACLAR$, NUFACORT$, RASOCLI$(NC1%), NORIT%, VUNI7#, MONEMI%, NC1%, DEPOEXPE%, (-1) * CANTAA, NPEDCAR&)
       If DESCOMPO%(CI%) > 0 Then
          Call MOVISTOK(FECHA7%, CI%, "", CMO$, NUFACLAR$, NUFACORT$, NUFACLAR$, NUFACORT$, RASOCLI$(NC1%), NORIT%, VUNI7#, MONEMI%, NC1%, DEPOEXPE%, CANTAA, NPEDCAR&)
          Call AREXPLO(CI%, 1, 1)
          For KKJ% = 1 To CAIT%
             CIMO% = CIJ%(KKJ%)
             CAMO = USOI(KKJ%) * CANTAA
             Call MOVISTOK(FECHA7%, CIMO%, "", CMO$, NUFACLAR$, NUFACORT$, NUFACLAR$, NUFACORT$, RASOCLI$(NC1%), NORIT%, VUNI7#, MONEMI%, NC1%, DEPOEXPE%, (-1) * CAMO, NPEDCAR&)
             NORIT% = NORIT% + 1
          Next KKJ%
       End If
       '
     End If
     'SI TIENE NUMERO DE SERIE ESCRIBE EN EL TAMUMSE EL NUMERO DE FACTURA
29   If TRIM$(NUMSERIE$) <> "" And TRIM$(NUMSERIE$) <> "S/N" Then
        CONDI$ = "NUMEROSERIE = '" & NUMSERIE$ & "'"
        Call ACTUREGISTRO("TANUMSE", "NUMFACTURA", CONDI$, NUFACLAR$, REGAF&, "NOMESS")
     End If
   Next I&
   '
   CONSITMP.Close
   Set CONSITMP = Nothing
   '
End Sub

Sub GRAREMIFA()
   '
   CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
   If Len(CMO$) <> 2 Then
     CMO$ = "RT"
   End If
   DEPOSALI% = DEPOEXPE%
   '
   NUFACLAR$ = CoComLar$
   NUFACORT$ = CoComCor$
   COPIASFAC% = XVALO(FORFINFAC.Text2(0))
   If COPIASFAC% < 1 Then
      NUFACLAR$ = ""
      NUFACORT$ = ""
   End If
   '
   PREFIREM$ = CONTROL$("", "PREFIREM")
   If PREFIREM$ = "" Then PREFIREM$ = "0001"
   'ESTO SE AGREGA PORQUE LOS REMITOS DE IMPRESORA FISCAL TIENEN EL MISMO
   'PUNTO DE VENTA QUE LA FACTURA
   If CONTROL$("FACTURA", "SALIDA") = "IFISCAL" Then PREFIREM$ = Label12
   '
   'ESTO SE AGREGA PORQUE LOS REMITOS DE LAS
   'FACTURAS ELECTRONICAS TIENEN EL PUNTO DE VENTA DE LA FACTURA
   If FACELEC% = 1 And VERIFSALIDA("NOMESS") < 3 Then PREFIREM$ = Label12
   '
   REMICOR$ = "RT." + TRIM$(Text10)
   REMILAR$ = Text10: While Len(REMILAR$) < 8: REMILAR$ = "0" + REMILAR$: Wend
   REMILAR$ = CMO$ + "-X-" + PREFIREM$ + "-" + REMILAR$
   If COPIASREM% < 1 Then
     '\\\OT-8-965 - SE MODIFICA PORQUE LA ANULACION DE UNA NC REQUIERE ESTE DATO
     REMICOR$ = NUFACORT$
     REMILAR$ = NUFACLAR$
     '\\\OT-8-965-FIN
   End If
   '
   NC1% = XVALO(Text1.TEXT)
   FF% = FECHANUM(Text9)
   '
   DESCX$ = PORDESCU
   PODES = COEFDESCU(DESCX$)
   COEFPRE = (100 - PODES) / 100
   '
   FIN& = ULTREG&("!CARDEFAC")
   '\\\OT-07-0023-RG-24/01/07///
   DEPOSALI% = XVALO(FORFINFAC.Text7)
   '\\\OT-07-0023-RG-FIN///
   PEDIANT& = 0
   '
   ANOTA$ = ""
   TTXYZ$ = Space$(76)
   HOII% = HOY(HOS$)
   Call REPLA(TTXYZ$, HOS$, 1, 8)
   Call REPLA(TTXYZ$, "Factura: " & NUFACORT$ & " - Remito: " & REMICOR$, 11, 56) 'ASIGNO EL NUMERO DE FACUTURA Y REMITO AL PROSPECT
   Call REPLA(TTXYZ$, USERTER$, 69, 24)
   ANOTA$ = ANOTA$ + TTXYZ$ + Chr$(10) + Chr$(13)
   '
   For I& = 1 To FIN&
     Call TRACE(20, I&, FIN&)
     X$ = REGLEIDO$("!CARDEFAC", I&)
     CANTAA = CVS(Left$(X$, 4))
     If Abs(CANTAA) > 0 Then
       CI% = CVI(Mid$(X$, 5, 2))
       If CI% >= 0 Then                 ' para incluir remito texto libre
         NORIT% = I& - 1
         PREUNI# = CDbl(CVS(Mid$(X$, 429, 4)) * COEFPRE)
         FACIVITEM = 1: If Label13.Caption = "B" Then FACIVITEM = 1 + PORCEIVA(CI%) / 100
         PRECONIVA# = FACIVITEM * CDbl(CVS(Mid$(X$, 429, 4)))
         '\\\OT-05-0562-WAR-02/08/05///
         SIG% = (-1)
         If TIFADOC% = 6 Then SIG% = 1 'NC
         '\\\OT-05-0562-WAR-02/08/05///
         NUPEDIORI& = XVALO(Mid$(X$, 371, 16))
         '
         If Mid$(X$, 461, 1) = "R" Then CMO$ = "*R" ' REPARACION
         TEXLIB$ = Mid$(X$, 7, 64)
         '
         NUSERRR$ = TRIM$(Mid$(X$, 139, 64))
         If TRIM$(NUSERRR$) = "S/N" Then NUSERRR$ = ""
         '
         ' APLICA ASIGNACION DE LOTE
         If ESTRAZABLE%(CI%) = 1 Then
             ' BUSCA EL LOTE DEL ARTICULO
             GoSub MUEVEPORLOTE
         Else
             Call MOVISTOK(FF%, CI%, NUSERRR$, CMO$, REMILAR$, REMICOR$, NUFACLAR$, NUFACORT$, RASOCLI$(NC1%), NORIT%, PREUNI#, MONEMI%, NC1%, DEPOSALI%, SIG% * CANTAA, NUPEDIORI&, , TEXLIB$, , , PRECONIVA#) '(-1) * CANTAA )
             If DESCOMPO%(CI%) > 0 Then
                  Call MOVISTOK(FF%, CI%, LOTE$, CMO$, REMILAR$, REMICOR$, NUFACLAR$, NUFACORT$, RASOCLI$(NC1%), NORIT%, PREUNI#, MONEMI%, NC1%, DEPOSALI%, (-1) * SIG% * CANTAA, NUPEDIORI&, , TEXLIB$, , , PRECONIVA#) '(-1) * CANTAA )
                  Call AREXPLO(CI%, 1, 1)
                  For KKJ% = 1 To CAIT%
                     CIMO% = CIJ%(KKJ%)
                     CAMO = USOI(KKJ%) * CANTAA
                     Call MOVISTOK(FF%, CIMO%, LOTE$, CMO$, REMILAR$, REMICOR$, NUFACLAR$, NUFACORT$, RASOCLI$(NC1%), NORIT%, PREUNI#, MONEMI%, NC1%, DEPOSALI%, SIG% * CAMO, NUPEDIORI&, , TEXLIB$, , , PRECONIVA#) '(-1) * CAMO)
                     NORIT% = NORIT% + 1
                  Next KKJ%
             End If
         End If
         '
10       If NUSERRR$ <> "" And UCase$(NUSERRR$) <> "S/N" And TIFADOC% <> 6 Then  'GRABA NUERO DE FACTURA Y DE REMITO EN TABLA TANUMSE
            NUREM1$ = REMILAR$
            NUFAC1$ = NUFACLAR$
            NUCLI1% = NC1%
            FechEnt = FETEXCOR1$(FF%)
            FE1% = DIASPOST(FF%, 730) 'FECHA DE GARANTIA
            FECHENT1 = CVDAT(FE1%)
            '
            Call ACTUREGISTRO("TANUMSE", "NumRemito", "NUMEROSERIE = '" & NUSERRR$ & "'", NUREM1$, NUAFEC&, "NOMESS")
            Call ACTUREGISTRO("TANUMSE", "NumFACTURA", "NUMEROSERIE = '" & NUSERRR$ & "'", NUFAC1$, NUAFEC&, "NOMESS")
            Call ACTUREGISTRO("TANUMSE", "NumeCli", "NUMEROSERIE = '" & NUSERRR$ & "'", NUCLI1%, NUAFEC&, "NOMESS")
            Call ACTUREGISTRO("TANUMSE", "FechEntrega", "NUMEROSERIE = '" & NUSERRR$ & "'", FechEnt, NUAFEC&, "NOMESS")
            Call ACTUREGISTRO("TANUMSE", "FeVenGar", "NUMEROSERIE = '" & NUSERRR$ & "'", FECHENT1, NUAFEC&, "NOMESS")
            '
            'GRABACION EN EL ANOTADOR (HISTORIAL DE EQUIPOS)
            CANCOPIASFAC% = XVALO(FORFINFAC.Text2(0)) ' cantidad de copias de factura
            CANCOPIASREM% = XVALO(FORFINFAC.Text2(1)) ' cantidad de copias de remitos
            CUERPO$ = "Venta de Equipos "
            If CANCOPIASFAC% > 0 Then
              CUERPO$ = CUERPO$ & "S/Factura: " & NUFAC1$ & vbCrLf
            End If
            If CANCOPIASREM% > 0 Then
               CUERPO$ = CUERPO$ + "Remito: " & NUREM1$
               '
               'ESCRIBO CON EL NUMERO DE REMITO QUE ESTE SALIO SIN CARGO Y EL MOTIVO
               NUMOT% = NUMOTIVO%(NUPEDIORI&)
               If NUMOT% > 0 Then
                 DESCMOTI$ = ECOARCH$("TAMOTPRE", Chr$(NUMOT%))
                 CUERPO$ = CUERPO$ + " - S/Cargo - Motivo: " & NUMOT%
                 CUERPO$ = CUERPO$ + " (" & TRIM$(DESCMOTI$) & ")"
               End If
            End If
            Call ANOTANUSE(NUSERRR$, CUERPO$)
    '
         End If
         ' NOTA DE CREDITO
         If NUSERRR$ <> "" And UCase$(NUSERRR$) <> "S/N" And TIFADOC% = 6 Then
            DEPX% = DEPOSALI% ' deposito configurado en FORFINFAC
            CONDI$ = "NUMEROSERIE = '" & NUSERRR$ & "'"
            Call ACTUREGISTRO("Tanumse", "NumRemito", CONDI$, "", REG&, "NOMESS")
            Call ACTUREGISTRO("Tanumse", "NumFactura", CONDI$, "", REG&, "NOMESS")
            Call ACTUREGISTRO("Tanumse", "FechEntrega", CONDI$, "", REG&, "NOMESS")
            Call ACTUREGISTRO("Tanumse", "NumeCli", CONDI$, 0, REG&, "NOMESS")
            Call ACTUREGISTRO("Tanumse", "DepoNumse", CONDI$, DEPX%, REG&, "NOMESS")
            Call ACTUREGISTRO("Tanumse", "FeVenGar", CONDI$, "", REG&, "NOMESS")
            Call ACTUREGISTRO("PEDDETA", "NUMEROSERIE", CONDI$, "", REG&, "NOMESS")
            
            '
            'GRABACION EN EL ANOTADOR (HISTORIAL DE EQUIPOS)
            CUERPO$ = "Reingreso de Equipo Según: " & NUFACLAR$
            Call ANOTANUSE(NUSERRR$, CUERPO$)
            '

         End If
         '
         'EN EL CASO QUE EL PEDIDO SEA CON MOTIVO LO MARCA COMO SIN CARGO
         If NUMOTIVO%(NUPEDIORI&) > 0 Then
            CONDI$ = "DOPRILAR = '" & REMILAR$ & "'"
            Call ACTUREGISTRO("MOVISTO", "DOSECCOR", CONDI$, "S/Cargo", REGAF&, "NOMESS")
            Call ACTUREGISTRO("MOVISTO", "DoSecLar", CONDI$, "S/Cargo", REGAF&, "NOMESS")
            Call COMUNI("Movimiento " + REMICOR$ + " Marcado 'Sin Cargo'.")
         End If

         TTXYZ$ = Space$(76) ' AGREGO DETALLE AL PROSPECT
         Call REPLA(TTXYZ$, "- " + CODEXT$(CI%), 11, 18)
         Call REPLA(TTXYZ$, DESCRIT0$(CI%), 29, 38)
         Call REPLA(TTXYZ$, CSTRING$(MKS$(CANTAA), 3, 10, 1, 2), 57, 10)
         ANOTA$ = ANOTA$ + TTXYZ$ + Chr$(10) + Chr$(13)

         ' PARA DOSIVAC ESQUEMA DE PACKING-LIST
         EPAC$ = TRIM$(UCase$(EMPREAC$))
         If InStr(EPAC$, "DOSIVA") > 0 Then
           If Left$(REMICOR$, 2) = "RT" Then
             DOPRIQ$ = "PL." + TRIM$(Mid$(X$, 371, 16))
             If TRIM$(DOPRIQ$) = "PL." Then   ' REMITO DIRECTO SIN O/VENTA
               NRXLX& = 900000 + Val(Text10)
               DOPRIQ$ = DOPRIQ$ + TRIM$(Str$(NRXLX&))
             End If
             REFEQ$ = "Packing-List " + RASOCLI$(NC1%)
             If CMO$ = "*R" Then
                  CMOVQ$ = "*D"
               Else
                  CMOVQ$ = "RP"
             End If
             Call MOVISTOK(FF%, CI%, LOTE$, CMOVQ$, DOPRIQ$, DOPRIQ$, NUFACLAR$, NUFACORT$, REFEQ$, NORIT%, PREUNI#, MONEMI%, NC1%, DEPOSALI%, (-1) * SIG% * CANTAA, NUPEDIORI&, , , , , PRECONIVA#) '(-1) * CANTAA)
             If DESCOMPO(CI%) > 0 Then
                  Call MOVISTOK(FF%, CI%, LOTE$, CMOVQ$, DOPRIQ$, DOPRIQ$, NUFACLAR$, NUFACORT$, REFEQ$, NORIT%, PREUNI#, MONEMI%, NC1%, DEPOSALI%, SIG% * CANTAA, NUPEDIORI&, , , , , PRECONIVA#) '(-1) * CANTAA)
                  Call AREXPLO(CI%, 1, 1)
                  For KKJ% = 1 To CAIT%
                     CIMO% = CIJ%(KKJ%)
                     CAMO = USOI(KKJ%) * CANTAA
                     Call MOVISTOK(FF%, CIMO%, LOTE$, CMOVQ$, DOPRIQ$, DOPRIQ$, NUFACLAR$, NUFACORT$, REFEQ$, NORIT%, PREUNI#, MONEMI%, NC1%, DEPOSALI%, (-1) * SIG% * CAMO, NUPEDIORI&, , , , , PRECONIVA#)      '(-1) * CAMO)
                     NORIT% = NORIT% + 1
                 Next KKJ%
             End If
           End If
         End If
         '
         If CANTAA > 0 Then ' PARA EVITAR CON NOTAS DE CREDITO
           CANAX# = CANTAA
           NUORITPE% = CVI(Mid$(X$, 503, 2))
           If NUPEDIORI& > 0 Then
             Call BLANPEDIPE(NC1%, FF%, REMILAR$, CI%, CANAX#, NUPEDIORI&, NUORITPE%)
             GoSub ANOTAPEDI
           End If
         End If
         '
       End If
     End If
   Next I&
   '
   If FACTLISTAEMPAQUE& > 0 Then
     CONDI$ = "NUMLTAEMP = " & FACTLISTAEMPAQUE& & " AND CODI_CLI = " & NC1%
     Call ACTUREGISTRO("ENCEMPAQ", "NUMFACTURA", CONDI$, NUFACLAR$, REGAF&, "NOMESS")
   End If
   '
   If Label13.Caption = "E" And XVALO(FORFINFAC.Text2(15)) > 0 Then
       Call MOVISTOK(FF%, 0, LOTE$, CMO$, REMILAR$, REMICOR$, NUFACLAR$, NUFACORT$, TRIM$(FORFINFAC.Text10), NORIT% + 1, XVALO(FORFINFAC.Text2(15)), MONEMI%, NC1%, DEPOSALI%, -1, NUPEDIORI&)
       GRABGASEXP% = 1
   End If
   '
   If PEDIANT& > 0 Then
      Call AGREGA_ANOTAPED(PEDIANT&, Space$(10) + String$(56, "-"))
      Call AGREGA_ANOTAPED(PEDIANT&, String$(88, "="))
   End If
   '
   Call ABRECLIEN
   '
   NPRO$ = VALOBADA("BSEGCOM", "CODI_CLI", "CODI_CLI = '" & TRIM$(Label26) & "'") 'BUSCO EL NUMERO DE CODIGO DE CLIENTE
   If NPRO$ <> "" Then
     Call AGREGA_ANOTAPROSPECT(NPRO$, Mid$(ANOTA$, 1, Len(ANOTA$) - 2)) ' AGREGA AL ANOTADOR DEL PROSPECT
     Call AGREGA_ANOTAPROSPECT(NPRO$, String$(88, "="))
   End If
   
Exit Sub
   '
ANOTAPEDI:                ' ANOTADOR DE PEDIDOS
   '
   NPEDIDX& = NUPEDIORI&
   If NPEDIDX& <> PEDIANT& Then
      If PEDIANT& <> 0 Then
          Call AGREGA_ANOTAPED(PEDIANT&, Space$(10) + String$(56, "-"))
          Call AGREGA_ANOTAPED(PEDIANT&, String$(88, "="))
      End If
      PEDIANT& = NPEDIDX&
      COPIASFAC% = Val(FORFINFAC.Text2(0))
      COPIASREM% = Val(FORFINFAC.Text2(1))
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
      Call AGREGA_ANOTAPED(NPEDIDX&, TTXYZ$)
      Call REPLA(TTXYZ$, Space$(88), 1, 88)
      Call AGREGA_ANOTAPED(NPEDIDX&, TTXYZ$)
      '
   End If
   '
   TTXX$ = Space$(76)
   Call REPLA(TTXX$, "- " + CODEXT$(CI%), 11, 18)
   Call REPLA(TTXX$, DESCRIT0$(CI%), 29, 38)
   Call REPLA(TTXX$, FORMATNUM$(CANAX#, "F10.1"), 57, 10)
   Call AGREGA_ANOTAPED(NPEDIDX&, TTXX$)
   '
Return

'***********************************************************************************************
MUEVEPORLOTE:
        For II& = 0 To ASIGLOT07.List1.ListCount
            XX$ = ASIGLOT07.List1.List(II&)
            If XVALO(Mid$(XX$, 80, 5)) = I& Then
               If CODINT%(TRIM$(Mid$(XX$, 1, 16))) = CI% Then
                  CANTLOT$ = Mid$(XX$, 85)
                  GoTo 5
               End If
            End If
        Next II&
    '
    ' MUEVE EL STOCK POR CADA LOTE
5   NN% = InStr(CANTLOT$, ";")
    If NN% > 0 Then
        LOTE$ = Mid$(CANTLOT$, 1, NN% - 1)
        If TRIM$(LOTE$) <> "" Then
            CANTAA = XVALO(VALOBADA("LOTINGRE", "CANALTA", "IDENLOTE='" & LOTE$ & "'", "NOMESS"))
            Call MOVISTOK(FF%, CI%, LOTE$, CMO$, REMILAR$, REMICOR$, NUFACLAR$, NUFACORT$, RASOCLI$(NC1%), NORIT%, PREUNI#, MONEMI%, NC1%, DEPOSALI%, SIG% * CANTAA, NUPEDIORI&, , TEXLIB$) '(-1) * CANTAA )
            CANTLOT$ = Mid$(CANTLOT$, NN% + 1)
            GoTo 5
        End If
    End If
    GoTo 10
    '
Return

End Sub

Sub AJUPREVE()
   '
   NC1% = XVALO(Text1.TEXT)
   Call APERBASDAT("STOCKS")
   
   FIN& = ULTREG&("!DIFPRECO")
   '
   For I& = 1 To FIN&
     Call TRACE(20, I&, FIN&)
     X$ = REGLEIDO$("!DIFPRECO", I&)
     PRENUEVO# = CVD(Mid$(X$, 45, 8))
     If PRENUEVO# > 0.00005 Then
       CI% = CVI(Mid$(X$, 1, 2))
       If CI% > 0 Then
         DOPRIL$ = TRIM$(Mid$(X$, 81, 18))
         NORIT% = CVI(Mid$(X$, 117, 2))
         PREUNI = CVD(Mid$(X$, 45, 8))
         '
         SQLX$ = "SELECT * FROM MOVISTO "
         SQLX$ = SQLX$ + " WHERE DoPriLar = '" & DOPRIL$ & "' "
         SQLX$ = SQLX$ + "AND Cod_Inte = " & CI% & " "
         SQLX$ = SQLX$ + "AND Nume_Clie = " & NC1% & " "
         SQLX$ = SQLX$ + "AND NuorItem = " & NORIT% & " "
         Dim MOVVVISTO As ADODB.Recordset
         Set MOVVVISTO = New ADODB.Recordset
24       On Error Resume Next
         MOVVVISTO.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
         If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 24
         End If
         On Error GoTo 0
         '
         With MOVVVISTO
           '.FindFirst SQLX$
25         'If .NoMatch = False Then
             '.Edit
           If Not (.EOF And .BOF) Then
             !VALOUNIT = PREUNI
             .Update
           End If
         End With
       End If
     End If
   Next I&
   MOVVVISTO.Close
   Set MOVVVISTO = Nothing
   '
End Sub

Sub CARTAMONED()
   '
   If DERACCE%("TICAMBIO", "Ajustar Tipo de Cambio Central") < 2 Then
     Exit Sub
   End If
   '
   If ULTREG&("TAMONED") > 6 Then
      Call SETULTREG("TAMONED", 6)
      Call COMUNI("Tabla de Monedas Truncada\al Máximo de 6 Registros")
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
   TICAMBIS = FORMATNUM(TICAMONE(MONEMI%), "F9.4")
   TICAMBX# = TICAMONE(MONEMI%)
   '
End Sub

Sub SHOWPEDIPEN()
     
     If YAINDI% = 0 Then
         '
         Screen.MousePointer = 11
         
         Call ABREPEDCLI
         'jandy sql     '

         Dim PediModifTemp As ADODB.Recordset
         Set PediModifTemp = New ADODB.Recordset
         PediModifTemp.CursorLocation = adUseClient
         
         SQLX$ = " SELECT DISTINCT PEDENCA.NroPed, PEDENCA.NroClie, PEDENCA.FechEmis, PEDENCA.NroOcom "
         SQLX$ = SQLX$ + " FROM PEDENCA "
         SQLX$ = SQLX$ + " INNER JOIN  PEDDETA "
         SQLX$ = SQLX$ + " ON (PEDENCA.NroPed = PEDDETA.NroPed) "
         SQLX$ = SQLX$ + " WHERE PEDDETA.Status < 8 "   ' para que no este anulada ni cumplida
         SQLX$ = SQLX$ + " AND PEDDETA.CantEnt < (PEDDETA.CanPed - 0.005) "
         SQLX$ = SQLX$ + " AND PEDENCA.Status < 9 "
         SQLX$ = SQLX$ + " AND PEDENCA.CODIEMPR = " & CodiEmp% & " "
         SQLX$ = SQLX$ + " AND PEDDETA.CODIEMPR = " & CodiEmp% & " "
         SQLX$ = SQLX$ + " ORDER BY PEDENCA.NroPed ASC "
         
         'Set PediModifTemp = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
25       On Error Resume Next
         PediModifTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
         If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 25
         End If
         On Error GoTo 0
         '
         JJ& = 0
         With PediModifTemp
           On Error Resume Next
           .MoveFirst
           If Err < 1 Then
             Do While Not .EOF
                 NRODES& = !NROPED
               TTXX$ = MKS$(NRODES&)
                 NOCO$ = !NroOcom
               If NOCO$ <> "" Then
                 TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
               End If
                 FEX = CVINT(!FechEmis)
               TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
                 NCLIE% = !NroClie
               TTXX$ = TTXX$ + TRIM$(RASOCLI$(NCLIE%)) + Space$(52)
               JJ& = JJ& + 1
               Call GRAREG("INDIPEP", Left$(TTXX$, 52) + MKS$(REGENCA&), JJ&)
             .MoveNext
             Loop
           End If
         End With
         PediModifTemp.Close
         Set PediModifTemp = Nothing
         '
         Call SETULTREG("INDIPEP", JJ&)
         Call CIERRARCH("INDIPEP")
         Call FRESHECHO("INDIPEP")
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
     TIDOCUM$ = "Pedido de Cliente"
     TIDOCUX$ = "Pedido ABIERTO"
     TIDOCUY$ = "Pedido Interno"
     If DERACCE%("VERCOSTO/NOMESS", "Visualizar Costos y Precios") < 1 Then
        TIDOCUM$ = "Orden de Despacho"
        TIDOCUX$ = "Orden de Despacho"
        TIDOCUY$ = "Orden de Despacho"
     End If
     '
     NUDOCU$ = TRIM$(Str$(NPX&))
     While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
     NUDOCU$ = NUDOCU$ + " "
     Screen.MousePointer = 11
     '
4    For UI& = ULTREG&("PDOCLST") To 1 Step -1
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
     Call COMUNI("Documento no Encontrado")
     Exit Sub
     '
5    Call MUESTRADOC(DOCUNU&)
     '
     GoTo 2110
     '
9999 End Sub
'
Sub DIFPREC()
    '
    CONSICLI% = 0
    NCLIE% = XVALO(Text1.TEXT)
    If NCLIE% <= 0 Then
        Call COMUNI("Numero de Cliente no Válido")
        Exit Sub
    End If
    '
    If DES% < 1 Then
2     FEX = HOY(HO$)
      HAS% = FEX
      HASS$ = FECHATEX$(FEX)
      DESS$ = "01" + Mid$(HASS$, 3)
      DES% = FECHANUM(DESS$)
    End If
    '
    VDEF$ = MKI$(DES%) + MKI$(HAS%)
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
       GoTo 2
    End If
    '
    'DES1# = CDbl(CVDAT(DES%))
    'HAS1# = CDbl(CVDAT(HAS%))
    DESDE1 = CVDAT(DES%)
    HASTA1 = CVDAT(HAS%)
    '
    Call APERBASDAT("STOCKS")
    SQLX$ = " SELECT * FROM MOVISTO  "
    SQLX$ = SQLX$ + " WHERE FechMov >= '" & DESDE1 & "' "
    SQLX$ = SQLX$ + "AND FechMov <= '" & HASTA1 & "' "
    SQLX$ = SQLX$ + "AND CodiMovi= 'RT' "
    SQLX$ = SQLX$ + "AND LEFT(DoSecCor,2)='FC' "
    SQLX$ = SQLX$ + "AND Nume_Clie= " & NCLIE% & " "
    SQLX$ = SQLX$ + "ORDER BY FechMov ASC "
    'Set DIFPRETMP = Stocks.OpenRecordset(SQLX$, 4)
    Dim DIFPRETMP As ADODB.Recordset
    Set DIFPRETMP = New ADODB.Recordset
25  On Error Resume Next
    DIFPRETMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0
    '
    With DIFPRETMP
       On Error Resume Next
       .MoveFirst
       If Err Then
            Call COMUNI("No Existen Remitos Generados")
            On Error GoTo 0
            Exit Sub
       End If
       On Error GoTo 0
       YYY$ = REGBLAN("!DIFPRECO")
       JJ& = 0
       '
       If TRIM$(TEXT21(0)) = "0" Then
         Call COMUNI("Debe Selecionar una Lista de Precio")
         Screen.MousePointer = 1
         Exit Sub
       End If
       NULIST$ = TRIM(TEXT21((0)))
       Do While Len(NULIST$) < 3
         NULIST$ = "0" + NULIST$
       Loop
            '
       NOLIST$ = "LIPRE" + NULIST$
       Do While .EOF = False
            '
            PRENUE# = PRELISTA(NOLIST$, !COD_INTE)
            '
            Call REPLA(YYY$, MKI(!COD_INTE), 1, 2)
            Call REPLA(YYY$, !DoPriCor, 3, 12)
            Call REPLA(YYY$, !DoSecCor, 15, 12)
            Call REPLA(YYY$, MKI(CVINT(!fechMov)), 27, 2)
            Call REPLA(YYY$, MKD(!Cantsalid), 29, 8)
            Call REPLA(YYY$, MKD(!VALOUNIT), 37, 8)
            Call REPLA(YYY$, !DOPRILAR, 81, 18)
            Call REPLA(YYY$, !DoSecLar, 99, 18)
            Call REPLA(YYY$, MKI(!NuOrItem), 117, 2)
            Call REPLA(YYY$, MKD(PRENUE#), 45, 8)
            JJ& = JJ& + 1
            Call GRAREG("!DIFPRECO", YYY$, JJ&)
       .MoveNext
       Loop
    End With
    Set DIFPRETMP = Nothing
    Call SETULTREG("!DIFPRECO", JJ&)
    Call CIERRARCH("!DIFPRECO")
    '
    Call DIFPREC2
End Sub

Sub DIFPREC2()
    '
    List3.Clear
    '
    PRF$ = "DIFPREC"
    ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Diferencia de Cotización"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Frame1.Top + Picture3.Top))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture3.Left + 50))
    '
    VTB% = VENTABU%(PRF$ + ATRI$)
    '
    If VTB% >= 0 Then 'Si esta aprobado
      JJJ& = 0
      For U& = 1 To ULTREG&("!DIFPRECO")
        XX$ = REGLEIDO$("!DIFPRECO", U&)
        PRENUE# = CVD(Mid$(XX$, 45, 8))
        'If PRENUE# >= 0.00005 Then
          If PRENUE# < 0.00005 Then PRENUE# = CVD(Mid$(XX$, 37, 8))
          DIFUNI# = PRENUE - CVD(Mid$(XX$, 37, 8))
          DIFTOT# = DIFUNI# * CVD(Mid$(XX$, 29, 8))
            Call REPLA(XX$, MKD(DIFUNI#), 53, 8)
            Call REPLA(XX$, MKD(DIFTOT#), 61, 8)
            JJJ& = JJJ& + 1
            Call GRAREG("!DIFPRECO", XX$, JJJ&)
         'End If
      Next U&
       Call SETULTREG("!DIFPRECO", JJJ&)
      '
    Else
        Call BLANFORMU ' SI CANCELA EL WINTABU
        Exit Sub
    End If
    '
    Call COMPLETITU
    Call COMPLEDETA
    '
End Sub
Sub COMPLETITU()
    T1$ = "Cant."
    T2$ = "Código"
    T3$ = "Remito "
    T4$ = "Factura"
    T5$ = "Pr.Fact"
    T51$ = "Pr.Ajus"
    T6$ = "Difer."
    T7$ = "Importe"
    '
    T1$ = AJUSTD$(T1$, 7)
    T2$ = AJUSTI$(T2$, 12)
    T3$ = AJUSTI$(T3$, 12)
    T4$ = AJUSTI$(T4$, 12)
    T5$ = AJUSTD$(T5$, 7)
    T51$ = AJUSTD$(T51$, 7)
    T6$ = AJUSTD$(T6$, 7)
    T7$ = AJUSTD$(T7$, 7)
    '
    ADLIS$ = T1 & Space$(2) & T2 & Space$(1) & T3 & Space$(2) & T4 & Space$(1) & T5 & Space$(1) & T51 & Space$(1) & T6 & Space$(3) & T7 & Space$(1)
    Label17 = ADLIS$
    List3.Visible = True
    '
    MARSELPENDE.Visible = False
    '
End Sub
'
Sub COMPLEDETA()
    '
    Z$ = REGBLAN("!CARDEFAC")
    JJ& = 0
    '
    For U& = 1 To ULTREG&("!DIFPRECO")
        XX$ = REGLEIDO$("!DIFPRECO", U&)
        CANTI1# = CVD(Mid$(XX$, 29, 8))
        CODI1% = CVI(Mid$(XX$, 1, 2))
        RTO1$ = TRIM$(Mid$(XX$, 3, 12))
        FACT1$ = TRIM$(Mid$(XX$, 15, 12))
        PRVTA1# = CVD(Mid$(XX$, 37, 8))
        PRVCO1# = CVD(Mid$(XX$, 45, 8))
        DIFCOT1# = CVD(Mid$(XX$, 53, 8))
        IMPORDIF# = CVD(Mid$(XX$, 61, 8))
        FECHFACT% = CVI(Mid$(XX$, 27, 2))
        '
        RTOLAR$ = TRIM$(Mid$(XX$, 81, 18))
        FACTLAR$ = TRIM$(Mid$(XX$, 99, 18))
        NUMORDIT% = CVI(Mid$(XX$, 117, 2))
        '
        D1$ = FORMATNUM$(CANTI1#, "F7.1")
        CODI2$ = CODEXT(CODI1%)
        D2$ = AJUSTI(CODI2$, 16)
        D3$ = AJUSTI(RTO1$, 13)
        D4$ = AJUSTI(FACT1$, 13)
        D5$ = FORMATNUM$(PRVTA1#, "F11.4")
        D6$ = FORMATNUM$(PRVCO1#, "F11.4")
        D7$ = FORMATNUM$(DIFCOT1#, "F11.4") 'DIF
        D8$ = FORMATNUM$(IMPORDIF#, "F11.2") 'IMPORTE
        If Abs(PRVCO1#) < Abs(0.00005) Or TRIM$(D5$) = TRIM$(D6$) Then
            GoTo 3
        Else
            ADDLIS$ = D1$ & "  " & D2$ & D3$ & D4$ & D5$ & D6$ & D7$ & Space(1) & D8$ & Space(3)
            List3.AddItem ADDLIS$
            '
            DESCRI$ = AJUSTI$(FACTLAR$, 18)
            DESCRI$ = DESCRI$ + FORMATNUM(PRVCO1#, "F12.4")
            '
            Call REPLA(Z$, MKI(CODI1%), 5, 5)
            Call REPLA(Z$, DESCRI$, 7, 64)
            Call REPLA(Z$, MKI(CVINT(Text9)), 419, 2)
            Call REPLA(Z$, MKS(CANTI1#), 1, 4)
            Call REPLA(Z$, MKS(DIFCOT1#), 421, 4)
            Call REPLA(Z$, MKS(DIFCOT1#), 429, 4)
            Call REPLA(Z$, MKD(IMPORDIF#), 433, 8)
            Call REPLA(Z$, RTOLAR$, 387, 16)
            Call REPLA(Z$, MKI(NUMORDIT%), 503, 2)
            JJ& = JJ& + 1
            Call GRAREG("!CARDEFAC", Z$, JJ&)
        End If
        '
3   Next U&
    '
    Picture5.Visible = True
    '
    Call CALTOFAC
    Call CALTOFA2
    '
End Sub

Sub CALPERIB()
   '
   '
   NCLIE% = XVALO(Text1)
   TASAX = ALIPERBRU#(NCLIE%, MINIPER#, VARIDOC%)
   '
   Text15 = ""
98 BIMPON# = XVALO(Total(2))
   'Lo que sigue se suprimió porque según M.Pazos (confirmado con Cont.Abadie
   'también en facturas 'B' la percepción se calcula sobre el NETO, no sobre
   'el bruto - Realizado por EPB 5/1/2007
   'If VARIDOC% = 2 Then BIMPON# = BIMPON# + XVALO(Total(3)) + XVALO(Total(4)) + XVALO(Total(5))
   If BIMPON# * TICAMONE(MONEMI%) >= MINIPER Then
       PERSALTA# = BIMPON# * TASAX / 100
       Text15 = FORMATNUM$(PERSALTA#, "F11.2")
   End If
   '
   Call CALTOFA2
   '
End Sub
'
Sub CARDETPEDIPEN()
   '
   MODODOSI% = 0
   If InStr(UCase$(EMPREAC$), "DOSIVA") > 0 Then MODODOSI% = 1
   '
   Screen.MousePointer = 1
   NC% = XVALO(Text1)
   Call SELECHO("SELFECHA/Entrega Solicitada Hasta:")
   FEHAS$ = TRIM$(VALACT1$("SELFECHA"))
   If FEHAS$ <> "" Then
       HAS1# = CDbl(CVDAT(FECHANUM(FEHAS$)))
     Else
       Call COMUNI("Sin Fecha de Corte.\   \Selecciona TODOS los Pedidos\Pendientes de Entrega.")
   End If
   '
   Call APERBASDAT("PEDCLI")
   SQLX$ = " SELECT * "
   SQLX$ = SQLX$ + " FROM PEDDETA "
   SQLX$ = SQLX$ + " WHERE NroClie = " & NC% & " "
   SQLX$ = SQLX$ + " AND Status < 8 "    ' NO ANULADOS NI CUMPLIDOS
   SQLX$ = SQLX$ + " AND STATUS >= 0 "   ' APROBADOS
   '
   If HAS1# > 0 Then
      'sqlx$ = sqlx$ + " AND CDBL(FeSolEnt) <= " & HAS1# & " "
      SQLX$ = SQLX$ + " AND FeSolEnt <= '" & FEHAS$ & "' "
   End If
   '
   'sqlx$ = sqlx$ + " AND (((CanPed - CantEnt) >= 0.05) OR (TipoPed = 'ABIERTO' AND CDBL(FinVigen) >= " & Int(CDbl(Now)) & ")) "
   SQLX$ = SQLX$ + " AND (((CanPed - CantEnt) >= 0.05) OR (TipoPed = 'ABIERTO' AND FinVigen >= getdate()))"
   
   SQLX$ = SQLX$ + " ORDER BY NroPed, NuOrItem;" 'ESTA LINEA LA AGREGUE
   '
   'Set PEDTEMP = Ventas.OpenRecordset(sqlx$, dbOpenSnapshot)
   Dim PEDTEMP As ADODB.Recordset
   Set PEDTEMP = New ADODB.Recordset
   PEDTEMP.CursorLocation = adUseClient
25 On Error Resume Next
   PEDTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   With PEDTEMP
     On Error Resume Next
     .MoveFirst
     '
     JJ& = 0: CAITCAR% = 0
     Do While (.EOF = False)
       CODIMAT% = !CodiInt
       If CODIMAT% > 0 And CODIMAT% <= NUMAS% Then
         CAN = !CanPed - !CantEnt
           If CAPEN < 0 Then CAPEN = 0
         If !TipoPed = "ABIERTO" Then CAPEN = 999999
         NPEDID& = !NROPED
         FEMIPED% = CVINT(!FechEmis)
         FENTRESOL% = CVINT(!FeSolEnt)
         PRECI1# = !PreUnid
         MONEMA% = !MoneEmis
            If MONEMA% > ULTREG&("TAMONED") Then MONEMA% = 1 ' ATAJA FUERA DE RANGO
         COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
         '
         If MODODOSI% = 1 Then COEFCON = 1 ' DOSIVAC CONSERVA LA MONEDA DEL PEDIDO
         '
         PRECIO = PRECI1# * COEFCON
           If PRECIO < 0.00005 Then PRECIO = PRELISTA#(ARCHIPRE$, CODIMAT%)
         PRELI = PRELISTA#(ARCHIPRE$, CODIMAT%)
           If PRELI < PRECIO Then PRELI = PRECIO
         PORDESCUEN = 0: If PRELI >= 0.0001 Then PORDESCUEN = 100 * (PRELI - PRECIO) / PRELI
         '
5        Z$ = REGBLAN$("!CARDEFAC")
         Call REPLA(Z$, MKS$(CAN), 1, 4)      ' cantidad
           COD$ = CODEXT$(CODIMAT%)
         Call REPLA(Z$, MKI$(CODIMAT%), 5, 2)
         Call REPLA(Z$, DESCRIT$(CODIMAT%), 7, 64)
           NREMIT$ = ""
         Call REPLA(Z$, NREMIT$, 387, 16) ' numero de remito
           FEREMIT% = 0
         Call REPLA(Z$, MKI$(FEREMIT%), 419, 2)
         Call REPLA(Z$, NUPLANO$(CODIMAT%), 315, 24)
         '
         Call REPLA(Z$, MKS$(PRELI), 421, 4)
         Call REPLA(Z$, MKS$(PORDESCUEN), 425, 4)
         Call REPLA(Z$, MKS$(PRECIO), 429, 4)
           IMPO# = CDbl(PRECIO) * CAN
         Call REPLA(Z$, MKD$(IMPO#), 433, 8)
         Call REPLA(Z$, MKS$(PORCEIVA(CODIMAT%)), 441, 4)
         '
         Call REPLA(Z$, " ", 461, 1)
         If Left$(UCase$(SAFETEXT$(!VenRep)), 1) = "R" Then
            Call REPLA(Z$, "R", 461, 1)
         End If
           '
           NORITPE% = !NuOrItem
         Call REPLA(Z$, MKI$(NORITPE%), 503, 2)
         Call REPLA(Z$, MKS$(0), 505, 4)
         Call REPLA(Z$, MKS$(0), 509, 4)
         Call REPLA(Z$, TRIM$(Str$(NPEDID&)), 371, 16)
           NUOCOX$ = NUOCOCLI$(NPEDID&, CODIMAT%)
10       Call REPLA(Z$, NUOCOX$, 355, 16)
         '
         JJ& = JJ& + 1
         Call GRAREG("!CARDEFAC", Z$, JJ&)
         CAITCAR% = CAITCAR% + 1
         '
       End If
     .MoveNext
     Loop
   End With
   PEDTEMP.Close
   Set PEDTEMP = Nothing
   '
   Call SETULTREG("!CARDEFAC", JJ&)
   '
   JJ& = 0
   For U& = 1 To ULTREG&("!CARDEFAC")
     Z$ = REGLEIDO$("!CARDEFAC", U&)
     If Abs(CVS(Mid$(Z$, 1, 4))) >= 0.01 Then
       JJ& = JJ& + 1
       Call GRAREG("!CARDEFAC", Z$, JJ&)
     End If
   Next U&
   Call SETULTREG("!CARDEFAC", JJ&)
   If EXISTE%(LUCOM$ + "1CARDEFA.WKF") > 0 Then
     Call SETULTREG("!1CARDEFA", 0)
     Call CIERRARCH("!1CARDEFA")
   End If
   Call CIERRARCH("!CARDEFAC")
   '
End Sub

Sub SELPEDIP()
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   MODODOSI% = 0: If InStr(EPAC$, "DOSIVA") > 0 Then MODODOSI% = 1
   '
   '
   Screen.MousePointer = 11
   NC% = XVALO(Text1)
   Call ABREPEDCLI
   '
   SQLX$ = " SELECT DISTINCT PEDENCA.NroPed, PEDENCA.NroClie, PEDENCA.FechEmis, PEDENCA.NroOcom "
   SQLX$ = SQLX$ + " FROM PEDENCA "
   SQLX$ = SQLX$ + " INNER JOIN  PEDDETA "
   SQLX$ = SQLX$ + " ON (PEDENCA.NroPed = PEDDETA.NroPed) "
   SQLX$ = SQLX$ + " WHERE PEDENCA.NroClie = " & NC% & " "
   SQLX$ = SQLX$ + " AND PEDDETA.Status < 8 "   ' para que no este anulada ni cumplida
   SQLX$ = SQLX$ + " AND PEDDETA.CantEnt < (PEDDETA.CanPed - 0.005) "
   SQLX$ = SQLX$ + " AND PEDENCA.Status < 9 "
   SQLX$ = SQLX$ + " AND PEDENCA.CODIEMPR = " & CodiEmp% & " "
   SQLX$ = SQLX$ + " AND PEDDETA.CODIEMPR = " & CodiEmp% & " "
   If EsPorPedido& > 0 Then
      SQLX$ = SQLX$ + " AND PEDDETA.NroPed = " & EsPorPedido&
'      GoTo 251
   End If

   SQLX$ = SQLX$ + " ORDER BY PEDENCA.NroPed ASC "
   'JANDY SQL
   Dim PediModifTemp As ADODB.Recordset
   Set PediModifTemp = New ADODB.Recordset
   'Set PediModifTemp = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   PediModifTemp.CursorLocation = adUseClient
25 On Error Resume Next
   PediModifTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   U& = 0
   SELPEDIPEN.LICOPEN.Clear
   With PediModifTemp
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
         Do While Not .EOF
           U& = U& + 1
             NRODES& = !NROPED
           TTXX$ = FORMATNUM$(NRODES&, "F6.0") + "  "
             NOCO$ = TRIM$(!NroOcom)
           If NOCO$ <> "" Then
             TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
           End If
             FEX = CVINT(!FechEmis)
           TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
             NCLIE% = !NroClie
           TTXX$ = TTXX$ + TRIM$(RASOCLI$(NCLIE%)) + Space$(52)
           SELPEDIPEN.LICOPEN.AddItem TTXX$
           If EsPorPedido& > 0 Then
             SELPEDIPEN.LICOPEN.Selected(U& - 1) = True
           End If
         .MoveNext
         Loop
      End If
   End With
   PediModifTemp.Close
   Set PediModifTemp = Nothing
   '
   Screen.MousePointer = 1
   '
   If EsPorPedido& < 1 Then
        If SELPEDIPEN.LICOPEN.ListCount < 1 Then ' se modifico el nombre del list por que no era el correcto
           SELPEDIPEN.Label5 = "0"
           GoTo 12                ' para que no muestre el form de seleccion
        End If
        SELPEDIPEN.Show 1
      Else
        Call SELPEDIPEN.BOTREC_Click
   End If
   '
12 CANSELE% = XVALO(SELPEDIPEN.Label5)
   If CANSELE% < 1 Then
      JJ& = 0
      GoTo 20     ' PARA BLANQUEAR ARCHIVO DE CARGA
   End If
   '
   Call APERBASDAT("PEDCLI")
   SQLX$ = " SELECT * "
   SQLX$ = SQLX$ + " FROM PEDDETA "
   SQLX$ = SQLX$ + " WHERE NroClie = " & NC% & " "
   SQLX$ = SQLX$ + " AND Status < 8 "    ' NO ANULADOS NI CUMPLIDOS
   SQLX$ = SQLX$ + " AND STATUS >= 0 "   ' APROBADOS
   'SQLX$ = SQLX$ + " AND (((CanPed - CantEnt) >= 0.05) OR (TipoPed = 'ABIERTO' AND CDBL(FinVigen) >= " & Int(CDbl(Now)) & ")) "
   SQLX$ = SQLX$ + " AND (((CanPed - CantEnt) >= 0.05)"
   SQLX$ = SQLX$ + " OR (TipoPed = 'ABIERTO' AND FinVigen >= GETDATE()))"
   '
   ATXX% = 0
   For KKU& = 1 To SELPEDIPEN.LICOPEN.ListCount
      If SELPEDIPEN.LICOPEN.Selected(KKU& - 1) = True Then
         NPESEL& = XVALO(Left$(SELPEDIPEN.LICOPEN.List(KKU& - 1), 6))
         If ATXX% < 1 Then
             SQLX$ = SQLX$ + "AND (NroPed = " & NPESEL& & " "
             ATXX% = 1
           Else
             SQLX$ = SQLX$ + "OR NroPed = " & NPESEL & " "
         End If
      End If
   Next KKU&
   SQLX$ = SQLX$ + ") "
   '
   SQLX$ = SQLX$ + " ORDER BY NroPed, NuOrItem;" 'ESTA LINEA LA AGREGUE
   '
   'Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   'JANDY SQL
   Dim PEDTEMP As ADODB.Recordset
   Set PEDTEMP = New ADODB.Recordset
   PEDTEMP.CursorLocation = adUseClient
26 On Error Resume Next
   PEDTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 26
   End If
   On Error GoTo 0
   '
   With PEDTEMP
     On Error Resume Next
     .MoveFirst
     '
     JJ& = 0: CAITCAR% = 0
     Do While (.EOF = False)
       CODIMAT% = !CodiInt
       If CODIMAT% > 0 And CODIMAT% <= NUMAS% Then
         CAN = !CanPed - !CantEnt
           If CAPEN < 0 Then CAPEN = 0
         If !TipoPed = "ABIERTO" Then CAPEN = 999999
         NPEDID& = !NROPED
         FEMIPED% = CVINT(!FechEmis)
         FENTRESOL% = CVINT(!FeSolEnt)
         PRECI1# = XVALO(!PreUnid)
         '
         'CALCULO DE RECARGO POR CONDICION DE PAGO
         RECXPAGO$ = VALOBADA("PEDENCA", "Rec_ConPa", "NROPED = " & NPEDID&, "NOMESS")
         CODESC = COEFDESCU(RECXPAGO$)
         If Abs(XVALO(CODESC)) < 0.05 Then CODESC = 0
         PRECI1# = PRECI1# + (PRECI1# * CODESC / 100)
         '
         MONEMA% = !MoneEmis
            If MONEMA% > ULTREG&("TAMONED") Then MONEMA% = 1 ' ATAJA FUERA DE RANGO
         COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
         If MODODOSI% = 1 Then COEFCON = 1 ' DOSIVAC CONSERVA LA MONEDA DEL PEDIDO
         '
         PRECIO = PRECI1# * COEFCON
           If PRECIO < 0.00005 Then PRECIO = PRELISTA#(ARCHIPRE$, CODIMAT%)
         'PRELI = PRELISTA#(ARCHIPRE$, CODIMAT%)
         PRELI = XVALO(!prelist) ' MODIFICADO PARA QUE TRAIGA EL PRECIO DE LISTA CON EL CUAL SE CREO EL PEDIDO
           If PRELI < PRECIO Then PRELI = PRECIO
         PORDESCUEN = 0: If PRELI >= 0.0001 Then PORDESCUEN = 100 * (PRELI - PRECIO) / PRELI
         '
5        Z$ = REGBLAN$("!CARDEFAC")
         Call REPLA(Z$, MKS$(CAN), 1, 4)      ' cantidad
           COD$ = CODEXT$(CODIMAT%)
         Call REPLA(Z$, MKI$(CODIMAT%), 5, 2)
         Call REPLA(Z$, DESCRIT$(CODIMAT%), 7, 64)
           NREMIT$ = ""
         Call REPLA(Z$, NREMIT$, 387, 16) ' numero de remito
           FEREMIT% = 0
         '
         If ESTRAZABLE%(CODIMAT%) = 2 Then
            NUMSE$ = "" ' NUMERO DE SERIE ASIGNADO AL PEDIDO
            ORDENITEM% = XVALO(!NuOrItem)
            CONDICION$ = " NroPed = " & NPEDID& & " AND NuOrItem = " & ORDENITEM% & " AND CodiInt = " & CODIMAT% & ""
            NUMSE$ = VALOBADA("PEDDETA", "NUMEROSERIE", CONDICION$)
            Call REPLA(Z$, NUMSE$, 139, 64)
         End If
         '
         Call REPLA(Z$, MKI$(FEREMIT%), 419, 2)
         Call REPLA(Z$, NUPLANO$(CODIMAT%), 315, 24)
         '
         Call REPLA(Z$, MKS$(PRELI), 421, 4)
         Call REPLA(Z$, MKS$(PORDESCUEN), 425, 4)
         Call REPLA(Z$, MKS$(PRECIO), 429, 4)
           IMPO# = CDbl(PRECIO) * CAN
         Call REPLA(Z$, MKD$(IMPO#), 433, 8)
         Call REPLA(Z$, MKS$(PORCEIVA(CODIMAT%)), 441, 4)
         '
         Call REPLA(Z$, " ", 461, 1)
         If Left$(UCase$(SAFETEXT$(!VenRep)), 1) = "R" Then
            Call REPLA(Z$, "R", 461, 1)
         End If
           '
           NORITPE% = !NuOrItem
         Call REPLA(Z$, MKI$(NORITPE%), 503, 2)
         Call REPLA(Z$, MKS$(0), 505, 4)
         Call REPLA(Z$, MKS$(0), 509, 4)
         Call REPLA(Z$, TRIM$(Str$(NPEDID&)), 371, 16)
           NUOCOX$ = NUOCOCLI$(NPEDID&, CODIMAT%)
10       Call REPLA(Z$, NUOCOX$, 355, 16)
           If TRIM$(NUOCOX$) <> "" Then
             If InStr(Text14, TRIM$(NUOCOX$) + ",") < 1 Then
               Text14 = Text14 + NUOCOX$ + ","
             End If
           End If
           '
         JJ& = JJ& + 1
         Call GRAREG("!CARDEFAC", Z$, JJ&)
         CAITCAR% = CAITCAR% + 1
         '
       End If
     .MoveNext
     Loop
   End With
   PEDTEMP.Close
   Set PEDTEMP = Nothing
   '
20 Call SETULTREG("!CARDEFAC", JJ&)
   '
   JJ& = 0
   For U& = 1 To ULTREG&("!CARDEFAC")
     Z$ = REGLEIDO$("!CARDEFAC", U&)
     If Abs(CVS(Mid$(Z$, 1, 4))) >= 0.01 Then
       JJ& = JJ& + 1
       Call GRAREG("!CARDEFAC", Z$, JJ&)
     End If
   Next U&
   Call SETULTREG("!CARDEFAC", JJ&)
   If EXISTE%(LUCOM$ + "1CARDEFA.WKF") > 0 Then
     Call SETULTREG("!1CARDEFA", 0)
     Call CIERRARCH("!1CARDEFA")
   End If
   Call CIERRARCH("!CARDEFAC")
   '
End Sub
'
Sub GRAFACTURA(OKX%)
   '
   Dim NIVA7#(10), NIVA7DOL#(10), CUEIVX$(10)
   '
   COTIMONEM# = XVALO(TICAMBIS) ' para prever ajuste de moneda de emision
   If COTIMONEM# < 0.05 Then COTIMONEM# = 1
   '
   OKX% = 0
   Tipo% = TIFADOC%
   GoSub APRUEBACOPIAS       ' APROBACION FINAL Y CANTIDAD DE COPIAS
     If REPRIFAC% = 1 Then
        OKX% = 1
        Exit Sub
     End If
   GoSub SOLOREMITO          ' PARA EL CASO DE QUE LAS COPIAS DE FACTURA SEAN CERO
   GoSub ICONTABLE           ' IMPUTACIÓN CONTABLE DE DEBITOS Y CREDITOS
   GoSub VALIDATCLI          ' VALIDACION DE DATOS DEL CLIENTE
   GoSub ARMANUMERO          ' ARMA NUMERO DE FACTURA CORTO Y LARGO
   GoSub VERILIMCRE          ' VERIFICA LIMITE DE CREDITO
   GoSub RECALTOTA           ' RECALCULA Y VERIFICA TOTALES
   GoSub ACSALCUEN           ' ACTUALIZA SALDO DE LA CUENTA
   GoSub RENUMHOJAS          ' RENUMERA SEGUN CANTIDAD DE HOJAS A CONSUMIR
   GoSub VERINUMERO          ' BLOQUEA Y VERIFICA NUMERO DE FACTURA LIBRE
   '
   If CONTROFIS% = 2 Then    ' PARA EL CASO DE CONTROLADOR FISCAL IMPRIME Y LUEGO GRABA
      Call LIBARCH("CUECORR")
      YAIMPRIMIO% = 0
      '
      ' se agrega esto para verificar si grabó en la impresora fiscal
      OPMOVIX$ = "FC"
         If TIFADOC% = 3 Then OPMOVIX$ = "ND"
         If TIFADOC% = 6 Then OPMOVIX$ = "NC"
      VAMOVIX$ = Label13
      Call TOMANUMDELCONFIS(OPMOVIX$, VAMOVIX$, PUNTODEVENTA$, NUMECOMP$)
      ESTAFAC$ = PUNTODEVENTA$ + "-" + NUMECOMP$
      CACOPDOCF% = XVALO(FORFINFAC.Text2(0).TEXT) ' LEE CANTIDAD COPIAS FACTURA
      '
      YAIMPRIMIO% = 0
      Call PRIFACTURA
      YAIMPRIMIO% = 1
      '
      ' ahora verifica que la factura quedó registrada en la impresora fiscal
      If VAMOVIX$ = "E" Then GoTo 29
      Call TOMANUMDELCONFIS(OPMOVIX$, VAMOVIX$, PUNTODEVENTA$, NUMECOMP$)
      PROXIFAC$ = PUNTODEVENTA$ + "-" + NUMECOMP$
      '
      If CACOPDOCF% > 0 Then
        If PROXIFAC$ = ESTAFAC$ Then
           TTXX$ = "Este Comprobante no Pudo Ser Emitido por Impresora Fiscal !!!"
           TTXX$ = TTXX$ + "\   \Causa Probable: Datos Incorrectos del Cliente."
           TTXX$ = TTXX$ + "\   \1.- Razón Social muy Larga (max. 30 pos)"
           TTXX$ = TTXX$ + "\2.- Domicilio muy Largo (max 30 pos)"
           TTXX$ = TTXX$ + "\3.- Posicion IVA o C.U.I.T. no Válidos"
           TTXX$ = TTXX$ + "\4.- Caracteres Extraños (acentos o eñes)"
           TTXX$ = TTXX$ + "\   \El Comprobante NO se Registró en la B.Datos FLEXOFT."
           TTXX$ = TTXX$ + "\   \Corrija los Datos y Re-Intente."
           Call MENSERR(24, TTXX$)
           '
           If MUECOMHAS% < 2 Then
              Call FINAL("")
           End If
        End If
      End If
      '
   End If
   '
29 Call APERBASDAT("CABAN")
   NOCLOSE% = 1
   GoSub GRATOTFAC           ' GRABA CAJABANC PARA FACTURA
   GoSub GRAVENETA           ' GRABA VENTA NETA
   GoSub GRAEMBALA           ' GRABA IMPORTE EMBALAJES
   GoSub GRABAIVA            ' graba iva inscripto Y NO INSCRIPTO
   GoSub GRAPERCEP           ' graba PERCEPCIONES IVA E I.BRUTOS
   GoSub GRAGASEXP           ' graba GASTOS DE EXPORTACION
   '
   If FACELEC% > 0 Then ' GRABA EL ARCHIVO GENERADO A PARTIR DEL CONSULTA CAE
      Call GRABACONCAE(NFCY$)
   End If
   '
   CARPETOK$ = "C:\FLEXOFT\FACELEC\"
   ' Call FILEKILL(LUCOM$ + "IDSOLCAE.TXT")  ' SI LLEGO HASTA ACA ELIMINA EL ARCHIVO IDSOLCAE.TXT
   Call LIBARCH("CUECORR")
   '
   'GRABA LA LETRA CORRESPONDIENTE A LA COMISION DEL VENDEDOR Y EL N. de VENDEDOR
   CONDI$ = "CodiCom_Lar = '" & TRIM$(NFCY$) & "'"
   If Label28.Visible = True Then
     Call ACTUREGISTRO("CAJABANC", "LETRACOMIS", CONDI$, Label28.Caption, REG&)
   End If
   '
   If Text22.Visible = True Then
     Call ACTUREGISTRO("CAJABANC", "LETRACOMIS", CONDI$, TRIM$(Text22), REGAF&, "NOMESS")
   End If
   '
   VENDXX = XVALO(TEXT21(2))
   Call ACTUREGISTRO("CAJABANC", "Vend_Cli", CONDI$, VENDXX, REGAF&, "NOMESS")
   '
   
   'Call BLIMESS("Calculando Vencimientos ...")
   If TIFADOC% <> 6 Or TRIM$(DOCUDIF) = "" Then
     Call GRADEUFA(Tipo%, VA%, NRO, NC%, FFI%, CPAG%, TOTAGEN#)
   End If
   'CajaBanc.Close
   'SaDeuPen.Close
   NOCLOSE% = 0
   '
   'Call BLIMESS("Movimiento de Stock ...")
   NRO1 = XVALO(Text11.TEXT)
   If TIFADOC% = 6 Then NRO1 = (-1) * XVALO(Text11.TEXT)
   '
   REMI = XVALO(Text10.TEXT)
   REMI1 = REMI: If TIFADOC% = 3 Then REMI1 = (-1) * NRO1
   If TIFADOC% = 3 Then   ' nota de debito
     If MODODOC% = 0 Then ' por diferencia de precios
       Call AJUPREVE
       GoTo 80
     End If
   End If
   '
   If CONSICLI% < 1 Then
       Call GRAREMIFA
     Else
       Call GRACONSI
   End If
   '
   If TIFADOC% = 3 Or TIFADOC% = 6 Then
     If TRIM$(DOCUDIF) <> "" Then
       Call APLIDICAM
     End If
   End If
   '
   If MONEMI% > 1 Then
     If TIFADOC% <= 3 Then   ' graba noviniento cabeza en caso de factura o nota
       Call GRAMOCABE        ' de debito en moneda extranjera
     End If
   End If
   '
80 OKX% = 1
   Screen.MousePointer = 1
   Exit Sub
   '
APRUEBACOPIAS:          ' PRIMERA VALIDACIÓN DE DATOS *******************************
   NCLIE% = XVALO(Text1.TEXT)
   If MARCONOTA.Visible = True Then Exit Sub
   If ULTREG&("!CARDEFAC") < 1 Then Exit Sub
   '
   Call APROCOPIAS(OKX%)
   COPIASFAC% = XVALO(FORFINFAC.Text2(0))
   COPIASREM% = XVALO(FORFINFAC.Text2(1))
   If COPIASREM% < 1 Then
      Text10.TEXT = ""
   End If
   If OKX% < 0 Then Exit Sub
   If COPIASFAC% + COPIASREM% < 1 Then Exit Sub
   '
   If COPIASFAC% > 0 Then
     If XVALO(Total(6).TEXT) < 0.01 Then
       Call MENSERR(24, "Imposible Emitir Factura SIN CARGO !!!")
       Exit Sub
     End If
   End If
Return                  ' ***********************************************************
'
SOLOREMITO:             ' CASO DE QUE SE EMITE REMITO SOLO Y NO FACTURA
   If COPIASFAC% < 1 Then
      '
      ' verificar que no sea de texto libre
      FIN& = ULTREG&("!CARDEFAC")
      For I& = 1 To FIN&
         Call TRACE(20, I&, FIN&)
         X$ = REGLEIDO$("!CARDEFAC", I&)
         CI% = CVI(Mid$(X$, 5, 2))
         If CI% < 1 Then
            Call MENSERR(24, "Imposible Emitir Solo Remito\de Material No Codificado.")
            OKX% = 0
            Exit Sub
         End If
      Next I&
      '
      GoSub VERINUMERO
      Call GRAREMIFA
      Call LIBARCH("CUECORR")
      YAIMPRIMIO% = 0
      Call PRIFACTURA
      Screen.MousePointer = 1
      OKX% = (-1)
      Exit Sub
   End If
Return                  ' ***********************************************************
'
ICONTABLE:              ' IMPUTACIÓN CONTABLE DE DEBITOS Y CREDITOS
   PPVI% = PIVACLI%(NCLIE%)
   If PPVI% < 0 Or PPVI% > 7 Then PPVI% = 1
   CUEVE$ = CONTROL$("SVD451S", "CUEVE-" + TRIM$(Str$(PPVI%)))
   QQQ% = CUECOINT%(CUEVE$)
   '
   If TIFADOC% = 3 Or TIFADOC% = 6 Then
      OPIMPUTA.Option1.Value = True
      Call OPIMPUTA.Option1_Click
      '
      If TRIM$(DOCUDIF) <> "" Then
        OPIMPUTA.Option2.Value = True
      End If
      '
      OPIMPUTA.Show 1
   End If
Return                  ' ***********************************************************
'
VALIDATCLI:             ' VALIDACION DE DATOS DE CLIENTE
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
   If XVALO(Total(6).TEXT) * TICAMBX# < 0.01 Then
      MERRO$ = "Importe no Válido"
      GoTo 98
   End If
   '

   HOII% = HOY(HOS$): FEACT = HOII%
   FEMI% = FECHANUM(Text9.TEXT)
   ' LA SIGUIENTE VALIDACION SE SUPRIME A PEDIDO DE S.IGLESIAS DE I.GUIDI
   'If FEMI% < DIANTE(FEACT) Or FEMI% > DIPOST(FEACT) Then
   '   MERRO$ = "Fecha de Emisión no Válida"
   '   GoTo 98
   'End If
   GoTo 99                 ' PASÓ LA VALIDACION DE DATOS
   '
98 Screen.MousePointer = 1
   Call MENSERR(24, MERRO$)
   Exit Sub
   '
99 Return                  ' ***********************************************************
'
ARMANUMERO:                ' ARMA NUMERO DE FACTURA CORTO Y LARGO
   TICOMX$ = "FC"
   If Tipo% = 3 Then TICOMX$ = "ND"
   If Tipo% = 6 Then TICOMX$ = "NC"
   '
   'ARMADO DE NUMERO PROVISORIO DE FACTURA
   If FACELEC% > 0 Then
      PUNTOVENTA$ = "XXX" & XVALO(Label12): PUNTOVENTA$ = Right(PUNTOVENTA$, 4)
20    ALEATO$ = RANDSTRING(-6)
      NFCZ$ = "00000000" & ALEATO$: NFCZ$ = Right(NFCZ$, 8)
      NFCY$ = TICOMX$ & "-" & PUNTOVENTA$ & "-" & NFCZ$ & "-" & Label13
      
      'VALIDO QUE NO EXISTA EN NINGUNA DE LAS TABLAS SIGUIENTES
      CJBNC& = CantRegistros("CAJABANC", "CODICOM_LAR =  '" & NFCY$ & "'", "NOMESS")
      MVST& = CantRegistros("MOVISTO", "DOSECLAR =  '" & NFCY$ & "'", "NOMESS")
      SDPN& = CantRegistros("SADEUPEN", "CODICOM_LAR =  '" & NFCY$ & "'", "NOMESS")
      TNMSE& = CantRegistros("TANUMES", "NUMFACTURA =  '" & NFCY$ & "'", "NOMESS")
      If CJBNC& > 0 Or MVST& > 0 Or SDPN& > 0 Or TNMSE& > 0 Then
        GoTo 20
      End If
      Label30 = NFCY$
      '
   Else
      'ARMADO DE NUMERO DE FACTURA TRADICIONAL
      NFCZ$ = TRIM$(Text11): While Len(NFCZ$) < 8: NFCZ$ = "0" + NFCZ$: Wend
      NFCY$ = TICOMX$ + "-" + Label12 + "-" + NFCZ$ + "-" + Label13
   End If
   
   
Return                     ' ***********************************************************
'
VERILIMCRE:                 ' VERIFICA LIMITE DE CREDITO      '
   If VERLICRE% = 1 Then
     If TIFADOC% < 4 Then   ' EXCLUYE CREDITOS
       '
       For UKK& = 1 To ULTREG&("!CARDEFAC")
         XKK$ = REGLEIDO$("!CARDEFAC", UKK&)
         NROPEDXX& = XVALO(TRIM$(Mid$(XKK$, 371, 16)))
         If NROPEDXX& < 1 Then GoTo 22
         RFFKK$ = RECFILT$("AUTORENT", 1, MKS$(NROPEDXX&))
         If Len(RFFKK$) < 4 Then GoTo 22 ' NO AUTORIZADO
         REGAUTKK& = CVS(RFFKK$)
         XAUT$ = REGLEIDO$("AUTORENT", REGAUTKK&)
         If UCase$(Mid$(XAUT$, 31, 2)) <> "SI" Then GoTo 22
       Next UKK&
       GoTo 23
       '
22     LMCR# = XVALO(Text7)     ' LIMICRE(NCLIE%)
       SDEU# = SALDCLI#(NCLIE%)
       CHPA# = CHEPEACRE#(NCLIE%)
       TPED# = XVALO(Total(6).TEXT) * TICAMBX#
       If SDEU# + CHPA# + TPED# > LMCR# Then
         TLIMI$ = TRIM$(CSTRING$(MKD$(LMCR#), 3, 12, 2, 2))
         If TLIMI$ = "" Then TLIMI$ = "0.00"
         Call COMUNI("Imposible Emitir Este Documento.\  \Supera el Límite de Crédito Asignado\para esta Cuenta de $ " + TLIMI$ + ".\  \Informe al Supervisor de Ventas.")
         Exit Sub
       End If
     End If
   End If
   '
23 Return
   '
Return                      ' ***********************************************************
'
RECALTOTA:                  ' RE-CALCULO TOTALES DE FACTURA
   DESC$ = TRIM$(PORDESCU.TEXT)
   PODES = COEFDESCU(DESC$)
   '
   For KK% = 1 To 10: TOTAFAC#(1, KK%) = 0: Next KK%
   TOTAFAC#(1, 1) = (XVALO(BASIVA(0)) + XVALO(TLABEL28)) / (1 - PODES / 100)
   TOTAFAC#(1, 2) = XVALO(BASIVA(1)) / (1 - PODES / 100)
   TOTAFAC#(1, 3) = XVALO(Total(0).TEXT)
   TOTAFAC#(1, 4) = XVALO(Total(1).TEXT)
   TOTAFAC#(1, 5) = XVALO(BASIVA(0)) + XVALO(TLABEL28)
   TOTAFAC#(1, 6) = XVALO(Total(3).TEXT)
   TOTAFAC#(1, 7) = XVALO(Total(5).TEXT)
   TOTAFAC#(1, 8) = XVALO(Total(6).TEXT)
   TOTAFAC#(1, 9) = XVALO(BASIVA(1))
   TOTAFAC#(1, 10) = XVALO(Total(4).TEXT)
   '
   TICA = 1
   For KK% = 1 To 10
     TOTAFAC#(1, KK%) = ROUND#(TOTAFAC#(1, KK%) * TICAMBX#, 2)
     TOTAFAC#(2, KK%) = TOTAFAC#(1, KK%) / TICA
   Next KK%
   '
   TOTAGEN# = 0: TOTAGENDOL# = 0: TOTGASEXP# = 0
   NENETO# = XVALO(Total(2)) - XVALO(IEMBALA)
   NEFA7# = ROUND#(NENETO# * TICAMBX#, 2)
   NEFA7DOL# = ROUND#(NENETO#, 2)
   TOTAGEN# = TOTAGEN# + NEFA7#
   TOTAGENDOL# = TOTAGENDOL# + NEFA7DOL#
   NEMBAL7# = ROUND#(XVALO(IEMBALA) * TICAMBX#, 2)
   NEMBAL7DOL# = ROUND#(XVALO(IEMBALA), 2)
   TOTAGEN# = TOTAGEN# + NEMBAL7#
   TOTAGENDOL# = TOTAGENDOL# + NEMBAL7DOL#
   For KKI% = 0 To 2
     NIVA7#(KKI%) = ROUND#(XVALO(Total(KKI% + 3)) * TICAMBX#, 2)
     NIVA7DOL#(KKI%) = ROUND#(XVALO(Total(KKI% + 3)), 2)
     TOTAGEN# = TOTAGEN# + NIVA7#(KKI%)
     TOTAGENDOL# = TOTAGENDOL# + NIVA7DOL#(KKI%)
   Next KKI%
   PERIB7# = ROUND#(XVALO(Text15) * TICAMBX#, 2)
   PERIB7DOL# = ROUND#(XVALO(Text15), 2)
   TOTAGEN# = TOTAGEN# + PERIB7#
   TOTAGENDOL# = TOTAGENDOL# + PERIB7DOL#
   PERIVA7# = ROUND#(XVALO(Text17) * TICAMBX#, 2)
   PERIVA7DOL# = ROUND#(XVALO(Text17), 2)
   TOTAGEN# = TOTAGEN# + PERIVA7#
   TOTAGENDOL# = TOTAGENDOL# + PERIVA7DOL#
   TOTGASEXP# = ROUND#(XVALO(FORFINFAC.Text2(15)) * TICAMBX#, 2)
   TOTGASEXPDOL# = ROUND#(XVALO(FORFINFAC.Text2(15)), 2)
   TOTAGEN# = TOTAGEN# + TOTGASEXP#
   TOTAGENDOL# = TOTAGENDOL# + TOTGASEXPDOL#
   '
Return                      ' ***********************************************************
'
ACSALCUEN:                  ' ACTUALIZA SALDO DE LA CUENTA
   LU$ = LUDAT$
   ARCHIDEU$ = "SVD202"
   ARCHICLI$ = "DEUDOR"
   ARCHIFAC$ = "FACTUR"
   RD% = 1
   Tipo% = TIFADOC%
   VA% = VARIDOC%
   NRO = XVALO(Text11.TEXT)
   NC% = NCLIE%
   FFI% = FECHANUM(Text9.TEXT)
   CPAG% = XVALO(TEXT21(1).TEXT)
   '
   '
   'Call BLIMESS("Actualizando Saldo de Cuenta ...")
   SIGX% = 1: If Tipo% = 6 Then SIGX% = (-1)
   TTF# = TOTAGEN# * SIGX%
   Call GRACLI(LU$, ARCHICLI$, NC%, TTF#, DATCLI$(), 2)
   Call CIERRARCH("DEUDOR2")
   '
   FLIM% = FECHANUM("01/01/99")
   If ULTREG&("FORMFAC") > 0 Then
      FFS$ = REGLEIDO$("FORMFAC", 2&)
      FLIM% = CVI(Mid$(FFS$, 19, 2))
   End If
Return                      ' ***********************************************************
'
RENUMHOJAS:                 ' RENUMERA PARA CONSIDERAR LAS HOJAS QUE SE CONSUMIRAN
   FORIPRE$ = TRIM$(CONTROL$("", "FORFAC-" + Label13.Caption))
   CARENGLO& = ULTREG&("!CARDEFAC")
   ITPORHOJA% = XVALO(ATRIFORM$(FORIPRE$, "MAX-ITEM"))
   If ITPORHOJA% < 1 Then
     ITPORHOJA% = 80
   End If
   CAHOJNEC% = Int(0.99 + CARENGLO& / ITPORHOJA%)
   '
'   If NFCAM% < 1 Then
'     Call NUFACTURA
'   End If
   '
27 NPP& = XVALO(Text11) + CAHOJNEC% - 1
   Text11.TEXT = TRIM$(Str$(NPP&))
   Text11.Refresh
   '
   
   TICOMX$ = "FC": OPMOVIX$ = "FAC"
   If Tipo% = 3 Then TICOMX$ = "ND": OPMOVIX$ = "DEB"
   If Tipo% = 6 Then TICOMX$ = "NC": OPMOVIX$ = "CRE"
   '
   If FACELEC% < 1 Then
    NFCZ$ = TRIM$(Text11): While Len(NFCZ$) < 8: NFCZ$ = "0" + NFCZ$: Wend
    NFCX$ = TICOMX$ + "-" + Mid$(NFCZ$, 3) + "-" + Label13
    NFCY$ = TICOMX$ + "-" + Label12 + "-" + NFCZ$ + "-" + Label13
   Else
    'AGREGADA
    NFCX$ = TICOMX$ + "-" + Right(NFCZ$, 6) + "-" + Label13
   End If
   '
   OPMOVIX$ = "FC"
   If TIFADOC% = 3 Then OPMOVIX$ = "ND"
   If TIFADOC% = 6 Then OPMOVIX$ = "NC"
   VAMOVIX$ = Label13
   '
Return                      ' ***********************************************************
'
VERINUMERO:
   ' ESTO ES PARA VERIFICAR QUE EL ARCHIVO ESTÉ BLOQUEADO
   Call BLOQARCH("CUECORR")
   'CajaBanc.Close
   If ISLOCKED%("CUECORR") < 1 Then
      Call LIBARCH("CUECORR")
      Call MENSERR(24, "No fue Posible Abrir la Base en Modo Exclusivo.\  \Espere un Instante y Vuelva a Grabar.")
      Exit Sub
   End If
   '
   If COPIASFAC% > 0 Then
     If VERINUFA%(NFCY$) < 0 Then
       Call LIBARCH("CUECORR")
       Call NUFACTURA
       NROREMIT& = XVALO(Text10)
       NPVX$ = Label12
       If PIVADOC% = 5 Or PIVADOC% = 7 Then
          EPAC$ = TRIM$(UCase$(EMPREAC$))
          If InStr(EPAC$, "DOSIVA") > 0 Then
            NPVX$ = "0001"
          End If
       End If
       If CONTROFIS% = 2 Then
            Call NUREMIFA_HAS(NROREMIT&, NPVX$)
          Else
            Call NUREMIFA(NROREMIT&, NPVX$)
       End If
       Text10.TEXT = TRIM$(Str$(NROREMIT&))
       If FACELEC% > 0 Then 'si ya esta en uso genera otro
         PUNTOVENTA$ = "000" & XVALO(Label12): PUNTOVENTA$ = Right(PUNTOVENTA$, 3)
31       NFCZ$ = RANDSTRING(-6)
         NFCY$ = TICOMX$ & "-" & "9" & PUNTOVENTA$ & "-" & NFCZ$ & "-" & Label13
         If CantRegistros("CAJABANC", "CODICOM_LAR =  '" & NFCY$ & "'", "NOMESS") > 0 Then
          GoTo 31
         End If
       Else
         Call MENSERR(24, "Número de Comprobante '" + NFCY$ + "' Ocupado.\   \FLEXOFT ha Re-Numerado el Documento\como Número '" + TRIM$(Text11.TEXT) + "'.\  \Verifique y Vuelva a Grabar.")
         Exit Sub
      End If
     End If
   End If
   If TIFADOC% < 3 Then
     If COPIASREM% > 0 Then
       If VERINURE% < 0 Then
         Call LIBARCH("CUECORR")
         NREMIRE$ = TRIM$(Text10)
         NROREMIT& = 0
         NPVX$ = Label12
         If PIVADOC% = 5 Or PIVADOC% = 7 Then
           EPAC$ = TRIM$(UCase$(EMPREAC$))
           If InStr(EPAC$, "DOSIVA") > 0 Then
             NPVX$ = "0001"
           End If
         End If
         If CONTROFIS% = 2 Then
              Call NUREMIFA_HAS(NROREMIT&, NPVX$)
            Else
              Call NUREMIFA(NROREMIT&, NPVX$)
         End If
         Text10.TEXT = TRIM$(Str$(NROREMIT&))
995      If VERINURE% < 0 Then
            NROREMIT& = NROREMIT& + 1
            Text10.TEXT = TRIM$(Str$(NROREMIT&))
            GoTo 995
         End If
         Call MENSERR(24, "Número de Remito '" + NREMIRE$ + "' Ocupado.\   \FLEXOFT ha Re-Numerado el Remito\como Número '" + TRIM$(Text10.TEXT) + "'.\  \Verifique y Vuelva a Grabar.")
         Exit Sub
       End If
     End If
   End If
   '
Return                      ' ***********************************************************
'
GRATOTFAC:                  ' GRABA TOTAL DE FACTURA
30 CoComCor$ = NFCX$
   CoComLar$ = NFCY$
'   CodiEmp% = 0
'   CodiSuc% = 0
   NumeClie% = Int(GRUDIVCLI(NC))
   NumeProv% = 0
   FechEmi = CVDAT(FFI%)
   Referen$ = RASOCLI$(NC%)
   FeContab = CVDAT(FFI%)
   FeAcredi = CVDAT(FFI%)
   CuentaCo% = CUECOCLI%(NC%)
   TiMovi$ = "FVEN"
   OpMovi$ = OPMOVIX$
   VaMovi$ = Label13
       PUNTOVENTA$ = Label12.Caption
       If FACELEC% > 0 Then PUNTOVENTA$ = "XXX" & XVALO(Label12): PUNTOVENTA$ = Right(PUNTOVENTA$, 4)
   Descrip$ = Label7 + " " + Label13.Caption + " " + Label6.Caption + " " + PUNTOVENTA$ + "-" + NFCZ$
   NumeTal = XVALO(Label12)
   NumeCom = CDbl(Text11)
   If FACELEC% > 0 Then NumeCom = (-1)
   StatInit = 0
   ImpoDeb = TOTAGEN# * SIGX%
   ValoDola = TOTAGENDOL# * SIGX%
   Observa$ = RASOCLI$(NC%)
   'SE UTILIZA UN CAST POR QUE ES UN CAMPO DE TEXTO VER SI QUEDA ASI
   IdentiCae$ = TRIM$(Str$(1 + XVALO(VALOBADA("CAJABANC", "MAX(cast(substring(IDCAE,1,LEN(IDCAE))as INT))", " TIPOMOVIM = 'FVEN'", "NOMESS"))))
   ObservaFactu$ = TRIM$(FINFACTU107.Text4)
   Call GRACABAN
Return                      ' ***********************************************************
'
GRAVENETA:                  ' GRABA VENTA NETA
   'graba venta neta - Verificar que QQQ sea la IC de combinacion Cliente-Tipo de Comprobante
   TOTANETX# = NEFA7#
   If TOTANETX# >= 0.005 Then
     QQ1% = CUECOINT%(Text16)
     If QQ1% < 1 Then QQ1% = QQQ%
     CuentaCo% = QQ1%   ' TOMA LA QUE APARECE EN PANTALLA
     TiMovi$ = "VENTA"
     OpMovi$ = OPMOVIX$
     VaMovi$ = Label13
     Descrip$ = "Mercaderia Vendida " + NFCX$
     StatInit = 0
     ImpoDeb = TOTANETX# * SIGX% * (-1)
     ValoDola = NEFA7DOL# * SIGX% * (-1)
     Call GRACABAN
   End If
Return                      ' ***********************************************************
'
GRAGASEXP:                  ' GRABA GASTOS DE EXPORTACION
   TOTANETX# = TOTGASEXP#
   If TOTANETX# >= 0.005 Then
     CUEMBA$ = CONTROL$("SVD451S", "CUEVE-13")
     If CUEMBA$ = "" Or CUECOINT%(CUEMBA$) < 1 Then CUEMBA$ = Text16
     QQ1% = CUECOINT%(CUEMBA$)
     If QQ1% < 1 Then QQ1% = QQQ%
     CuentaCo% = QQ1%   ' TOMA LA QUE APARECE EN PANTALLA
     TiMovi$ = "VENTA"
     OpMovi$ = OPMOVIX$
     VaMovi$ = Label13
     Descrip$ = TRIM$(FORFINFAC.Text10) + " " + NFCX$
     StatInit = 0
     ImpoDeb = TOTANETX# * SIGX% * (-1)
     ValoDola = TOTGASEXPDOL# * SIGX% * (-1)
     Call GRACABAN
   End If
Return                      ' ***********************************************************

GRAEMBALA:                   ' GRABA IMPORTE EMBALAJES
   TOTANETX# = NEMBAL7#
   If TOTANETX# >= 0.005 Then
     CUEMBA$ = CONTROL$("SVD451S", "CUEVE-12")
     If CUEMBA$ = "" Or CUECOINT%(CUEMBA$) < 1 Then CUEMBA$ = Text16
     QQ1% = CUECOINT%(CUEMBA$)
     If QQ1% < 1 Then QQ1% = QQQ%
     CuentaCo% = QQ1%   ' TOMA LA QUE APARECE EN PANTALLA
     TiMovi$ = "EMBAL"
     OpMovi$ = OPMOVIX$
     VaMovi$ = Label13
     Descrip$ = "Gasto Embalajes " + NFCX$
     StatInit = 0
     ImpoDeb = TOTANETX# * SIGX% * (-1)
     ValoDola = NEMBAL7DOL# * SIGX% * (-1)
     Call GRACABAN
   End If
Return                      ' ***********************************************************
'
GRABAIVA:                   ' GRABA IVA INSCRIPTO Y NO INSCRIPTO
   'graba iva inscripto
   CUEIVX$(0) = CONTROL$("SVD451S", "CUEIVA")
   CUEIVX$(1) = CONTROL$("SVD451S", "CUEIVB")
   If TRIM$(CUEIVX$(1)) = "" Or CUECOINT%(CUEIVX$(1)) < 1 Then
     CUEIVX$(1) = CUEIVX$(0)
   End If
   For KKI% = 0 To 1
     TOTANETX# = NIVA7#(KKI%)
     If TOTANETX# >= 0.005 Then
       CuentaCo% = CUECOINT%(CUEIVX$(KKI%))
       TiMovi$ = "IVA"
       OpMovi$ = "DeFi"
       VaMovi$ = PORIVA(KKI%) + " %"
       Descrip$ = "Iva Debito Fiscal " + NFCX$
       StatInit = 0
       ImpoDeb = TOTANETX# * SIGX% * (-1)
       ValoDola = NIVA7DOL#(KKI%) * SIGX% * (-1)
       BasImpo# = ROUND#(XVALO(BASIVA(KKI%)) * TICAMBX#, 2)
       If KKI% = 0 Then
          BasImpo# = BasImpo# + ROUND#(XVALO(TLABEL28) * TICAMBX#, 2)
       End If
       Call GRACABAN
     End If
   Next KKI%
   BasImpo# = 0
   '
   'graba iva NO inscripto
   TOTANETX# = NIVA7#(2)
   If TOTANETX# >= 0.005 Then
     CUEIVA$ = CONTROL$("SVD451S", "CUEADU")
     CuentaCo% = CUECOINT%(CUEIVA$)
     TiMovi$ = "IVA"
     OpMovi$ = "DeFi"
     VaMovi$ = "RNI"
     Descrip$ = "IVA Resp.No Inscripto " + NFCX$
     StatInit = 0
     ImpoDeb = TOTANETX# * SIGX% * (-1)
     ValoDola = NIVA7DOL#(2) * SIGX% * (-1)
     BasImpo# = ROUND#(XVALO(BASIVA(2)) * TICAMBX#, 2)
     Call GRACABAN
   End If
   BasImpo# = 0
Return                      ' ***********************************************************
'
GRAPERCEP:                  ' PERCEPCIONES I.BRUTOS E IVA
   'graba percepcion I.Brutos
   TOTANETX# = PERIB7#
   If TOTANETX# >= 0.005 Then
     CUEIVA$ = CONTROL$("PERIBRU", "CUECONTA")
     CuentaCo% = CUECOINT%(CUEIVA$)
     TiMovi$ = "IBRU"
     OpMovi$ = "IB-B"
     VaMovi$ = "PERC"
     Descrip$ = "Perc.IB.Bs.As. " + NFCX$
     StatInit = 0
     ImpoDeb = TOTANETX# * SIGX% * (-1)
     ValoDola = PERIB7DOL# * SIGX% * (-1)
     BasImpo# = ROUND#(XVALO(Total(2)) * TICAMBX#, 2)
     Call GRACABAN
   End If
   BasImpo# = 0
   '
   'graba percepcion IVA
   TOTANETX# = PERIVA7#
   If TOTANETX# >= 0.005 Then
     CUEPER$ = CONTROL$("SVD451S", "CUERET")
     CuentaCo% = CUECOINT%(CUEPER$)
     TiMovi$ = "IVA"
     OpMovi$ = "DeFi"
     VaMovi$ = "PER-IVA"
     Descrip$ = "Percep.IVA " + NFCX$
     StatInit = 0
     ImpoDeb = TOTANETX# * SIGX% * (-1)
     ValoDola = PERIVA7DOL# * SIGX% * (-1)
     BasImpo# = ROUND#(XVALO(Total(2)) * TICAMBX#, 2)
     Call GRACABAN
   End If
   BasImpo# = 0
Return                      ' ***********************************************************
'
End Sub

'
Sub RECONCLI(NCX7%, hubore%)
   '
   hubore% = 0
   '
   Dim SALCONSI(32767), SALCONUE(32767)
   Dim DOCUSA$(2048)
   If NCX7% > 0 Then
     If RASOCLI$(NCX7%) <> "" Then
       Screen.MousePointer = 11
       FIN& = ULTREG&("CONSICLI")
       For KLK& = 1 To 32767: SALCONSI(KLK&) = 0: SALCONUE(KLK&) = 0: Next KLK&
       RFF$ = RECFILT$("CONSICLI", 1, MKI$(NCX7%))
       For U2& = 1 To Len(RFF$) Step 4   ' FIN&
         If U2& > Len(RFF$) Then Exit For
         U& = CVS(Mid$(RFF$, U2&, 4))
         Call TRACE(20, U&, FIN&)
         XX$ = REGLEIDO$("CONSICLI", U&)
         NCX8% = CVI(Left$(XX$, 2))
         If NCX8% = NCX7% Then
           CIXI% = CVI(Mid$(XX$, 35, 2))
           If CIXI% > 0 Then
             If CIXI% <= NUMAS% Then
               CACON = CVD(Mid$(XX$, 37, 8)) - CVD(Mid$(XX$, 45, 8))
               SALCONSI(CIXI%) = SALCONSI(CIXI%) + CACON
             End If
           End If
         End If
       Next U2&
       '
'NMS% = NUMAS%
'For UI% = 1 To NMS%
'If SALCONSI(UI%) > 0 Then
'AA1 = CODEXT$(UI%)
'BBB = CCC
'End If
'Next UI%

       FEXI = FETEXCOR1$(FECHANUM("22/08/08"))
       CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
       If Len(CMO$) <> 2 Then CMO$ = "RT"
       '
       SQLX$ = " SELECT * FROM MOVISTO "
       SQLX$ = SQLX$ + "WHERE Nume_Clie = " & NCX7% & " "
       SQLX$ = SQLX$ + " AND CodiMovi = '" & CMO$ & "' "
       SQLX$ = SQLX$ + " AND DoPriLar <> '' "
       SQLX$ = SQLX$ + " AND ((DoSecLar is null) OR (DoSecLar = '') OR (DoSecLar LIKE 'P/F%')) "
       SQLX$ = SQLX$ + " AND FechMov < '" & FEXI & "' "
       '
       Dim CONSITMP As ADODB.Recordset
       Set CONSITMP = New ADODB.Recordset
       CONSITMP.CursorLocation = adUseClient
25     On Error Resume Next
       CONSITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 25
       End If
       On Error GoTo 0
       '
       FIN& = 0
       With CONSITMP
         If Not (.BOF And .EOF) Then
           .MoveFirst
           Do While Not .EOF
              CIXI% = !COD_INTE
              SALCONUE(CIXI%) = SALCONUE(CIXI%) + !Cantsalid - !CantIngre
              FIN& = FIN& + 1
           .MoveNext
           Loop
         End If
       End With
       '
       CANCASOS% = 0
       For CII% = 1 To NUMAS%
          If SALCONSI(CII%) > 0 Then CANCASOS% = CANCASOS% + 1
          If SALCONSI(CII%) <> SALCONUE(CII%) Then
             hubore% = 1
          End If
       Next CII%
       If hubore% < 1 Then
          If CANCASOS% = FIN& Then
             Exit Sub
          End If
       End If
       '
       CONSITMP.Close
       Set CONSITMP = Nothing
       SQLX$ = " UPDATE MOVISTO SET DoSecCor = 'S/Cargo', DOSECLAR = 'Sin Cargo'"
       SQLX$ = SQLX$ + "WHERE Nume_Clie = " & NCX7% & " "
       SQLX$ = SQLX$ + " AND CodiMovi = '" & CMO$ & "' "
       SQLX$ = SQLX$ + " AND DoPriLar <> '' "
       SQLX$ = SQLX$ + " AND ((DoSecLar is null) OR (DoSecLar = '') OR (DoSecLar LIKE 'P/F%')) "
       SQLX$ = SQLX$ + " AND FechMov < '" & FEXI & "' "
       FLEXCONN.Execute SQLX$
       '
       For CII% = 1 To NUMAS%
          If SALCONSI(CII%) > 0 Then
             AA1 = CODEXT$(CII%)
             SQLX$ = " SELECT * FROM MOVISTO "
             SQLX$ = SQLX$ + "WHERE Nume_Clie = " & NCX7% & " "
             SQLX$ = SQLX$ + "AND Cod_Inte = " & CII% & " "
             SQLX$ = SQLX$ + " AND (CodiMovi = '" & CMO$ & "' "
             SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
             SQLX$ = SQLX$ + " AND DoPriLar <> '' "
             SQLX$ = SQLX$ + " AND ((DoSecLar is null) OR (DoSecLar = 'Sin Cargo')) "
             'SQLX$ = SQLX$ + " AND FechMov < '" & FEXI & "' "
             '
             Set CONSITMP = New ADODB.Recordset
             CONSITMP.CursorLocation = adUseClient
26              On Error Resume Next
             CONSITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockPessimistic, adCmdText
             If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                On Error GoTo 0
                Call CapturaErrorOpen
                GoTo 26
             End If
             On Error GoTo 0
             '
             JJ& = 0
             With CONSITMP
               If Not (.BOF And .EOF) Then
                 .MoveFirst
                 Do While Not .EOF
                    If !Cantsalid > 0 Then
                      JJ& = JJ& + 1
                      DOCUSA$(JJ&) = !DOPRILAR
                    End If
                 .MoveNext
                 Loop
               End If
             End With
             '
             For JJ1& = 1 To JJ&
               For JJ2& = JJ1& + 1 To JJ&
                 If DOCUSA$(JJ1&) < DOCUSA$(JJ2&) Then
                   Call SWAP(DOCUSA$(JJ1&), DOCUSA$(JJ2&))
                 End If
               Next JJ2&
             Next JJ1&
             '
             For JJ1& = 1 To JJ&
               SQLX$ = " SELECT * FROM MOVISTO "
               SQLX$ = SQLX$ + "WHERE Nume_Clie = " & NCX7% & " "
               SQLX$ = SQLX$ + "AND Cod_Inte = " & CII% & " "
               SQLX$ = SQLX$ + " AND CodiMovi = '" & CMO$ & "' "
               SQLX$ = SQLX$ + " AND DoPriLar = '" & DOCUSA$(JJ1&) & "' "
               SQLX$ = SQLX$ + " AND CANTSALID > 0 "
               SQLX$ = SQLX$ + " AND FechMov < '" & FEXI & "' "
               '
               Set CONSITMP = New ADODB.Recordset
               CONSITMP.CursorLocation = adUseClient
27                On Error Resume Next
               CONSITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockPessimistic, adCmdText
               If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                    On Error GoTo 0
                    Call CapturaErrorOpen
                    GoTo 27
               End If
               On Error GoTo 0
               '
               With CONSITMP
                 On Error Resume Next
                 .MoveFirst
                 On Error GoTo 0
                 If Not (.BOF And .EOF) Then
                   Do While Not .EOF
                     If XVALO(!Cantsalid) > 0 And XVALO(!Cantsalid) <= SALCONSI(CII%) Then
                       !DoSecCor = ""
                       !DoSecLar = ""
                       On Error Resume Next
                       .Update
                       If Err Then
                          On Error GoTo 0
                          GoTo 49
                       End If
                       On Error GoTo 0
                       SALCONSI(CII%) = SALCONSI(CII%) - XVALO(!Cantsalid)
                       If SALCONSI(CII%) < 0.5 Then GoTo 49
                     End If
                   .MoveNext
                   Loop
                 End If
               End With
               '
             Next JJ1&
          End If
          '
49     Next CII%
       '
     End If
   End If
   '
End Sub
'
Public Function VERIFSALIDA%(Optional MODO$)
   '
   'DEVUELVE(0 NINGUNO)-(1 SI ES PUNTO DE VTA NORMAL)-( 2 B.CAP.)
   '
   VRFSLIDA% = 0
   If CONTROL$("FACTURA", "SALIDA") = "FELECT" Then
      If TRIM$(Label12) = CONTROL("FACTURA", "PUVENELE") Then
           VRFSLIDA% = 1
         ElseIf TRIM$(Label12) = CONTROL("FACTURA", "PUVENCAP") Then
           VRFSLIDA% = 2
         ElseIf TRIM$(Label12) = CONTROL("FACTURA", "PUVEFAEX") Then
           VRFSLIDA% = 3
         Else
           If TRIM$(UCase$(MODO$)) <> "NOMESS" Then
              Call MENSERR(24, "Punto de Venta '" + TRIM$(Label12) + "' No Coincide\con los Configurados para Factura Electrónica")
           End If
           VRFSLIDA% = 0
      End If
   End If
   '
99 VERIFSALIDA% = VRFSLIDA%
   '
End Function

Function VALIDAPUNVEN%()
   ' VERIFICA DE ACUERDO AL PUNTO DE VENTA SI ES EL CONFIGURADO PARA FACTURAR BIENES DE CAPITALES
   ' QUE LOS ARTICULOS SELECCIONADOS ESTEN COMO TALES
   VALIPUVEN% = 1
   If CONTROL$("FACTURA", "SALIDA") <> "FELECT" Then GoTo 99
    
   VERISAL% = VERIFSALIDA%("NOMESS") 'DEVUELVE 1 SI ES PUNTO DE VTA NORMAL- 2 B.CAP. - 3 FACTURA DE EXPORTACION 0 NINGUNO DE LOS 3
   If VERISAL% < 1 Then GoTo 99
   '
   If VERISAL% = 1 Then TIPODEPROD% = 0: TIPROX$ = "(Comun)"  ' comunes
   If VERISAL% = 2 Then TIPODEPROD% = 1: TIPROX$ = "(Bien.Cap.)"  ' bienes de capital
   '
   POXIVA% = PIVACLI%(TRIM$(Text1))

   If VERISAL% = 2 Then                     ' bienes de capital
     '
     If POXIVA% < 1 Or POXIVA% = 5 Then
        Call MENSERR(24, "Imposible Facturar como Bienes de Capital\a Consumidor Final o Extranjero.")
        VALIPUVEN% = 0
        GoTo 99
     End If
   
     NUMERO_CUIT$ = CUITCLI$(XVALO(Text1))
     If VALICUIT%(NUMERO_CUIT$) <= 0 Then
       Call MENSERR(24, "Debe Tener un C.U.I.T. Válido\Para la Modalidad de Facturación\Bienes de Capitales")
       VALIPUVEN% = 0
       GoTo 99
     End If
     '
'     If IsNumeric(PORDESCU) = True Then
'        Call MENSERR(24, "No se Podran Aplicar Descuentos Generales\a Facturación de Bienes de Capital\Aplicar los Mismos Manualmente en Ventana de Carga")
'        VALIPUVEN% = 0
'        GoTo 99
'     End If
   End If
   '
   If POXIVA% > 0 And POXIVA% <> 5 Then ' si no es consumidor final
     FIN& = ULTREG&("!CARDEFAC")
     For UI& = 1 To FIN&
       XX$ = REGLEIDO$("!CARDEFAC", UI&)
       CI% = CVI(Mid$(XX$, 5, 2))
       If BIENCAP%(CI%) <> TIPODEPROD% Then
          Call MENSERR(24, "El Artículo " & CODEXT$(CI%) & TIPROX$ & " \No Puede Facturarse Por el Punto de Venta Activo.")
          VALIPUVEN% = 0
          GoTo 99
       End If
    Next UI&
   End If
   '
99 VALIDAPUNVEN% = VALIPUVEN%
  
End Function

Sub CARNUMSER(OK%)
   '
   OK% = 1
   REF2$ = ""
   FIN& = ULTREG&("!CARDEFAC")
   For U& = 1 To FIN&
      Y$ = REGLEIDO$("!CARDEFAC", U&)
      CAN# = CVS(Mid$(Y$, 1, 4))
      CI% = CVI(Mid$(Y$, 5, 2))
      If ESTRAZABLE%(CI%) = 2 Then
        If CAN# > 1 Then ' VALIDA QUE LA CANTIDAD NO SEA MAYOR A 1 PARA INGRESAR EL NUMERO DE SERIE
            Call MENSERR(24, "Cantidad No Valida - El Máximo Permitido es 1 (uno)\en los Articulos Trazables")
            OK% = 2
            Exit Sub
        End If
      Else
        REF2$ = "S/N"
        Call REPLA(Y$, REF2$, 139, 64)
        Call GRAREG("!CARDEFAC", Y$, U&)
      End If
   Next U&
   '
10 XXX% = VENTABU%("CARNUMSE")
   
   If XXX% < 0 Then
     OK% = 0
     Exit Sub
   Else
     FIN& = ULTREG&("!CARDEFAC")
     For U& = 1 To FIN&
       Y$ = REGLEIDO$("!CARDEFAC", U&)
       NUMSE$ = TRIM$(Mid$(Y$, 139, 64))
       If NUMSE$ <> "S/N" Then
         CI% = CVI(Mid$(Y$, 5, 2))
         '
         CONDI1$ = "NumeroSerie='" & NUMSE$ & "'"
         '
         REGI1$ = VALOBADA("TANUMSE", "NumeroSerie", CONDI1$)
         If REGI1$ = "" Then
           Call MENSERR(24, "Numero de Serie " & NUMSE$ & " Inexistente")
           GoTo 10
         End If
         '
         REGI2% = XVALO(VALOBADA("TANUMSE", "Cod_Inte", CONDI1$))
         If REGI2% <> CI% Then
           Call MENSERR(24, "Numero de Serie " & NUMSE$ & " Corresponde al Código: " & CODEXT$(REGI2%) & "")
           GoTo 10
         End If
         '
'         If TIFADOC% <= 2 Then           ' solo en caso de facturas.
'            REGI3$ = VALOBADA("TANUMSE", "NumRemito", CONDI1$)
'            If REGI3$ <> "" Then
'              Call MENSERR(24, "Numero de Serie " & NUMSE$ & " ya Entregado (Nro Remito: " & REGI3$ & ")")
'              GoTo 10
'            End If
'         End If
         '
         REGI3$ = TRIM$(VALOBADA("TANUMSE", "NumRemito", CONDI1$))
         If TIFADOC% <= 2 And REGI3$ <> "" Then ' solo en caso de facturas
              Call MENSERR(24, "Numero de Serie " & NUMSE$ & " ya Entregado (Nro Remito: " & REGI3$ & ")")
              GoTo 10
         ElseIf TIFADOC% = 6 And REGI3$ = "" Then ' solo en caso de notas de credito
             Call MENSERR(24, "Numero de Serie " & NUMSE$ & " No Se Encuentra Entregado")
             GoTo 10
         End If
       End If
     Next U&
   End If
   '
End Sub
Function VERIPEDMO%()
   'VALIDA QUE NO HAYAN PEDIDOS CON DIFERENTES MOTIVOS
   ' Y SI LOS PEDIDOS SELECCIONADOS ESTAN AUTORIZADOS
   
   '0 = NO
   '1 = OK
   '
   VERI% = 1
   '
   'RECORRO LOS PEDIDOS Y VALIDO SI TODOS TIENEN EL MISMO MOTIVO
   For KKU& = 1 To ULTREG&("!CARDEFAC")
        XX$ = REGLEIDO$("!CARDEFAC", KKU&)
        NPED& = XVALO(TRIM$(Mid$(XX$, 371, 16)))
        If KKU& = 1 Then
          NUMOT1% = NUMOTIVO%(NPED&)
        End If
        NUMOT2% = NUMOTIVO%(NPED&)
        If NUMOT2% <> NUMOT1% Then
           'Call COMUNI("Se Han Detectado Pedidos con Diferente Motivo Asignado")
           VERI% = 0
           GoTo 99
        End If
    Next KKU&
    '
    'RECORRO LOS PEDIDOS Y VALIDO SI TODOS ESTAN AUTORIZADOS
    For KKU& = 1 To ULTREG&("!CARDEFAC")
        XX$ = REGLEIDO$("!CARDEFAC", KKU&)
        NPED& = XVALO(TRIM$(Mid$(XX$, 371, 16)))
        If VALIDAOV%(NPED&) < 1 Then
           VERI% = 0
           GoTo 99
        End If
    Next KKU&
    '
99 VERIPEDMO% = VERI%
End Function
Function LISTADEMPAQ%(NC%)
    '
    LISEMPAQ% = 0
    CONDI$ = "CODI_CLI = " & NC% & " AND EmpaCerra > 0 and (NumFactura = '' or NumFactura is Null)"
    HAYREG& = CantRegistros("ENCEMPAQ", CONDI$)
    If HAYREG& > 0 Then
       LISEMPAQ% = 1
    End If
    '
    LISTADEMPAQ% = LISEMPAQ%
End Function


'LA RUTINA SE PASO AL MODULO MOVIST07_SQL
'Sub ASIGNALOTES()
'   'SELECCION DE LOTES A ASIGNAR AL PEDIDO
'   REGIS% = 0
'   ASIGLOT07.List1.Clear
'   ASIGLOT07.List2.Clear
'   For U& = 1 To ULTREG&("!CARDEFAC")
'       XX$ = REGLEIDO$("!CARDEFAC", U&)
'       CCII% = CVI(Mid$(XX$, 5, 2))
'       If CCII% > 0 And ESTRAZABLE%(CCII%) = 1 Then
'         TXTX$ = ""
'         Call REPLA(TXTX$, CODEXT$(CCII%), 1, 16)
'         Call REPLA(TXTX$, DESCRIT$(CCII%), 18, 33)
'         SALPEND = CVS(Mid$(XX$, 1, 4))
'         Call REPLA(TXTX$, FORMATNUM$(SALPEND, "F10.1"), 54, 12)
'         Call REPLA(TXTX$, TRIM$(FORMATNUM$(U&, "I5")), 80, 5)
'         ASIGLOT07.List1.AddItem (TXTX$)
'         REGIS% = REGIS% + 1
'       End If
'   Next U&
'   If REGIS% > 0 Then
'      ASIGLOT07.Show 1
'   End If
'   '
'End Sub
'
'LA RUTINA SE PASO AL MODULO MOVIST07_SQL
'Sub DETLOTREM(CI%, REG&)
'    '
'    If ESTRAZABLE%(CI%) = 1 Then
'       TOTALOTE$ = ""
'       For UY& = 0 To ASIGLOT07.List1.ListCount - 1
'           YY$ = ASIGLOT07.List1.List(UY&)
'           If XVALO(Mid$(YY$, 80, 5)) = REG& Then
'               TOTALOTE$ = TRIM$(Mid$(YY$, 85))
'           End If
'       Next UY&
'       '
'5      NN% = InStr(TOTALOTE$, ";")
'       If NN% > 0 Then
'           LOTTE$ = Mid$(TOTALOTE$, 1, NN% - 1)
'           '
'           CONDI$ = "IDENLOTE = '" & LOTTE$ & "'"
'           FEALT$ = (VALOBADA("LOTINGRE", "FECHALTA", CONDI$, "NOMESS"))
'           FEEXX = CVINT(FEALT$)
'           FEE$ = FECHATEX$(FEEXX)
'           ANCHOO = XVALO(VALOBADA("LOTINGRE", "ANCHOMILI", CONDI$, "NOMESS"))
'           CANTI = XVALO(VALOBADA("LOTINGRE", "CANALTA", CONDI$, "NOMESS"))
'           '
'           LOTTE$ = AJUSTI$(LOTTE$, 12) & " - " & TRIM$(FORMATNUM$(CANTI, "F10.1")) & " " & UNIMED$(CI%)
'           If ANCHOO > 0 Then
'             LOTTE$ = LOTTE$ & " - " & TRIM$(FORMATNUM$(ANCHOO, "F10.1")) & " mm"
'           End If
'           '
'           LOTTE$ = LOTTE$ & " - " & FEE$
'           Call AGREGATEXTO(LOTTE$, 2)
'           TOTALOTE$ = Mid$(TOTALOTE$, NN% + 1)
'           GoTo 5
'       End If
'    End If
'    '
'End Sub


Sub SOLICITARCAE_WSFEX1()
    FLAG% = 0
    cont% = 0
    '
    NC% = XVALO(Text1)
    'TOMO LOS DATOS DEL FRM DE INGRESO DE DATOS
    
    TIPO_EXPO$ = Left$(TRIM$(DATFAEX07.Combo1.TEXT), 1)
    DEST_COMP$ = Left$(TRIM$(DATFAEX07.Combo2.TEXT), 3)
    CUIT_PAIS$ = Left$(TRIM$(DATFAEX07.Combo4.TEXT), 11)
    INCOTS$ = TRIM$(DATFAEX07.Combo5.TEXT)
    IDIOMA$ = TRIM$(Left(DATFAEX07.Combo6.TEXT, 1))
    
    COMP_ASOCIAD% = DATFAEX07.List1.ListCount
    S_N$ = "N" 'VARIABLE QUE DEFINE SI TIENE PERMIOS O NO
    
    PERMISO% = DATFAEX07.List2.ListCount: If PERMISO% > 0 Then S_N$ = "S"
    ' EN EL CASO QUE SEA NOTA DE CREDITO O DEBITO ESTE DEBE ESTAR VACIO
    If FORFACPRINT.Label7.Caption = "N.Débito" Then S_N$ = ""
    If FORFACPRINT.Label7.Caption = "N.Crédito:" Then S_N$ = ""
    ' EN EL CASO QUE EL TIPO DE EXPORTACION NO SEA DE BIENES ESTE DEBE ESTAR VACIO
    If TIPO_EXPO$ <> "1" Then S_N$ = ""
    '
    CARPETOK$ = "C:\FLEXOFT\FACELEC\"
    If EXISTE%(CARPETOK$ + "token_WSFEX.txt") = 0 Then
      TXT$ = REPLACAR$(CARPETOK$, "\", "/")
      Call MENSERR(24, "Falta Archivo 'token_WSFEX.txt'\en Carpeta " & TXT$)
      Exit Sub
    End If
    If EXISTE%(LUDAT$ + "consulta_cae_WSFEX_original.txt") = 0 Then
      TXT$ = REPLACAR$(LUDAT$, "\", "/")
      Call MENSERR(24, "Falta Archivo 'consulta_cae_WSFEX_original.txt'\en Carpeta " & TXT$)
      Exit Sub
    End If
    NARCHI1% = FILEOPEN%(LUDAT$ + "consulta_cae_WSFEX_original.txt", 1, 128) 'ABRE EL ARCHIVO DE LECTURA
    NARCHI2% = FILEOPEN%(CARPETOK$ + "consulta_cae_WSFEX.txt", 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA
    NC% = XVALO(FORFACPRINT.Text1)
    Do While Not EOF(NARCHI1%) ' RECORRE
       Line Input #NARCHI1%, XX$ ' LEE CADA REGISTRO
       cont% = cont% + 1
       pos% = InStr(XX$, ">")
       ETIQUETACIERRE$ = TRIM$(Mid$(XX$, pos% + 1))
       TX$ = Mid$(XX$, 1, pos%)
       
          Select Case cont%
           Case 2
            FECHA$ = FECHATEXLAR$(FECHANUM(FORFACPRINT.Text9))
            DIA$ = Mid(FECHA$, 1, 2): MES$ = Mid$(FECHA$, 4, 2): ANIO$ = Mid$(FECHA$, 7, 4)
            DATO$ = ANIO$ + MES$ + DIA$
           Case 3: DATO$ = TIPOX_CBTEX$ ' TIPO DE COMPROBANTE
           Case 4: DATO$ = TRIM$(Str(XVALO(FORFACPRINT.Label12))) ' Punto de Venta
           Case 5: DATO$ = TRIM$(FORFACPRINT.Text11)            ' Numero de comprobante inicial
           Case 6: DATO$ = TIPO_EXPO$ ' TIPO DE EXPORTACION
           Case 7: DATO$ = S_N$ ' PERMISO EXISTENTE
           Case 9:
           If S_N$ = "S" Then
              For J% = 0 To DATFAEX07.List2.ListCount - 1
                PERMI$ = DATFAEX07.List2.List(J%)
                Print #NARCHI2%, "<permiso>"
                Print #NARCHI2%, "<id_permiso>" & TRIM$(PERMI$) & "</id_permiso>"
                Print #NARCHI2%, "<dst_merc>" & TRIM$(DEST_COMP$) & "</dst_merc>"
                Print #NARCHI2%, "</permiso>"
              Next J%
           End If
           Case 10: DATO$ = DEST_COMP$ ' DESTINO PAIS
           Case 11: DATO$ = TRIM$(RASOCLI$(NC%))
           Case 12: DATO$ = CUIT_PAIS$ ' CUIT DE PAIS
           Case 13: DATO$ = TRIM$(DOMICLI$(NC%)) 'DOMICILIO
           Case 14: DATO$ = TRIM$(CUITCLI$(NC%)) 'CUIT
           Case 15: DATO$ = MONEDA_CAE$ ' TIPO DE MONEDA
           Case 16: DATO$ = TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) 'COTIZACION
           Case 17: DATO$ = Left$(Text4, 2000) ' OBSERVACIONES
           Case 18: DATO$ = TRIM$(FORFACPRINT.Total(6)) ' Importe Total
           Case 19: DATO$ = Left$(Text4, 2000) ' OBSERVACIONES
           Case 21:
              If COMP_ASOCIAD% > 0 Then
                For J% = 0 To DATFAEX07.List1.ListCount - 1
                 RELEI$ = DATFAEX07.List1.List(J%)
                 CBTE$ = Mid$(TRIM$(RELEI$), 1, 2)
                 PVTA$ = Mid$(RELEI$, 20, 4)
                 NURO$ = Mid$(RELEI$, 32, 8)
                 Print #NARCHI2%, "<cmp_asoc>"
                 Print #NARCHI2%, "<cbte_tipo>" & TRIM$(CBTE$) & "</cbte_tipo>"
                 Print #NARCHI2%, "<cbte_punto_vta>" & TRIM$(PVTA$) & "</cbte_punto_vta>"
                 Print #NARCHI2%, "<cbte_nro>" & TRIM$(NURO$) & "</cbte_nro>"
                 Print #NARCHI2%, "</cmp_asoc>"
                Next J%
              End If
           Case 22: DATO$ = TRIM$(Left$(FORFACPRINT.Eco21(1), 50)) ' FORMA DE PAGO
           Case 23: DATO$ = INCOTS$ ' TIPO DE VENTA FOB,CIF,ETC
           FORFACPRINT.Text18 = INCOTS$ ' ACTUALIZO EL FRM
           Case 25: DATO$ = IDIOMA$ '1-ESPAÑOL 2-INGLES 3-PORTUGUES
           
          End Select
        TX$ = TX$ + DATO$ + ETIQUETACIERRE$ ' CONCATENA EL RENGLON A IMPRIMIR
        DATO$ = ""
        Print #NARCHI2%, TX$ ' ESCRIBE EN EL ARCHIVO EL NUEVO TEXTO
     Loop
    '
    Print #NARCHI2%, "<Items>"

    DESCRILAR$ = ""
    For I& = 1 To ULTREG&("!CARDEFAC")
      XX$ = REGLEIDO$("!CARDEFAC", I&)
      CI% = CVI(Mid$(XX$, 5, 2))
      DESCRI$ = DESCRIT$(CI%)
      If CI% = 0 Then DESCRI$ = TRIM$(Mid$(XX$, 7, 64))
      CAN = CVS(Mid$(XX$, 1, 4))
      PRECI# = CVS(Mid$(XX$, 429, 4))
      '
      PORDESCUEN = PORCEDESCU(FORFACPRINT.PORDESCU) 'MKS$(Mid$(XX$, 425, 4))
      '
      PRECI# = PRECI# * (100 - PORDESCUEN) / 100
      PRCI$ = FORMATNUM$(PRECI#, "F12.3")
      IMPORTE# = CDbl(Int(10 * CAN + 0.5)) / 10 * (Int(10000 * CDbl(PRECI#) + 0.5)) / 10000
      IMPRT$ = FORMATNUM$(IMPORTE#, "F12.2")
      CONDI$ = "Codint = " & CI%
      NOMENC$ = VALOBADA("MASTER", "CODIGO_NCM", CONDI$, "NOMESS")
      UNIMEX$ = TRIM$(VALOBADA("MASTER", "PRO_UMED", CONDI$, "NOMESS")) 'REVISA SI ESTA EN BASE DE DATOS
      If UNIMEX$ = "" Or CI% = 0 Then UNIMEX$ = Str$(UNIMEAFIP%(CI%)) ' SI NO LO TRAE DEL ARCHIVO DE UNIMED
      '
      If PRECI# > 0.005 Then
        Print #NARCHI2%, "<item>"
        Print #NARCHI2%, "<pro_codigo>" & TRIM$(NOMENC$) & "</pro_codigo>"
            If DESCRILAR$ <> "" Then DESCRI$ = DESCRILAR$ + DESCRI$
        Print #NARCHI2%, "<pro_ds>" & TRIM$(DESCRI$) & "</pro_ds>"
        Print #NARCHI2%, "<pro_qty>" & CAN & "</pro_qty>"
        Print #NARCHI2%, "<pro_umed>" & TRIM$(UNIMEX$) & "</pro_umed>"
        Print #NARCHI2%, "<pro_precio_uni>" & TRIM$(PRCI$) & "</pro_precio_uni>"
        Print #NARCHI2%, "<imp_total>" & TRIM$(IMPRT$) & "</imp_total>"
        Print #NARCHI2%, "</item>"
        DESCRILAR$ = ""
      Else
        DESCRILAR$ = DESCRILAR$ + DESCRI$
      End If
    Next I&
    Print #NARCHI2%, "</Items>"
    Print #NARCHI2%, "</Cmp>"
    
    Close #NARCHI1%
    Close #NARCHI2%
    '
99
End Sub

Sub SOLICITARCAE_WSBFE1()
    '
    CARPETOK$ = "C:\FLEXOFT\FACELEC\"
    If EXISTE%(CARPETOK$ + "token_WSBFE.txt") = 0 Then
      TXT$ = REPLACAR$(CARPETOK$, "\", "/")
      Call MENSERR(24, "Falta Archivo 'Token_WSBFE.txt'\en Carpeta " & TXT$)
      Exit Sub
    End If
    If EXISTE%(LUDAT$ + "consulta_cae_WSBFE_original.txt") = 0 Then
      TXT$ = REPLACAR$(LUDAT$, "\", "/")
      Call MENSERR(24, "Falta Archivo 'consulta_cae_WSBFE_original.txt'\en Carpeta " & TXT$)
      Exit Sub
    End If
    NARCHI1% = FILEOPEN%(LUDAT$ + "consulta_cae_WSBFE_original.txt", 1, 128) 'ABRE EL ARCHIVO DE LECTURA
    NARCHI2% = FILEOPEN%(CARPETOK$ + "consulta_cae_WSBFE.txt", 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA

    Do While Not EOF(NARCHI1%) ' RECORRE
       Line Input #NARCHI1%, XX$ ' LEE CADA REGISTRO
       cont% = cont% + 1
       pos% = InStr(XX$, ">")
       ETIQUETACIERRE$ = TRIM$(Mid$(XX$, pos% + 1))
       TX$ = Mid$(XX$, 1, pos%)
       
          Select Case cont%
           
           Case 2: DATO$ = FINFACTU107.TIPO_DOC      ' Tipo Documento (Cuit,Dni, etc) Identificador del comprador
           Case 3: DATO$ = FINFACTU107.NRO_DOC      ' Numero Dni o Cuit
           Case 4: DATO$ = "1"   ' Numero de zona
           Case 5: DATO$ = FINFACTU107.TIPO_CBTE$        ' Tipo de Comprobante(Factura A o B ,Nota de Credito,Nota de Debito)
           Case 6: DATO$ = TRIM$(Str(XVALO(FORFACPRINT.Label12))) ' Punto de Venta
           Case 7: DATO$ = TRIM$(FORFACPRINT.Text11)            ' Numero de comprobante inicial
           Case 8: DATO$ = TRIM$(FORFACPRINT.Total(6))          ' Importe Total
           Case 9: DATO$ = FINFACTU107.IMP_TOT_CONC            ' Importe no gravados = C/Iva, entre ellos neto no gravado, las percepciones de iva e IIBB
           Case 10: DATO$ = FINFACTU107.IMP_NETO                ' Importe Neto (IMPORTE NETO = IMPORTE NETO - IVA NO GRAVADO)
           Case 11: DATO$ = FINFACTU107.IMPTO_LIQ            ' Impuesto Liquidado(la suma de los importes de iva  aplicados)
           Case 12: DATO$ = "0"           ' Impto_Liq_rni - siempre es cero  o percepcion a no categorizados
           Case 13: DATO$ = "0"           ' importe operaciones externas (No dieron informacion de que se trata)
           Case 14: DATO$ = "0"
           Case 15: DATO$ = "0"
           Case 16: DATO$ = "0"
           Case 17: DATO$ = "0"
           Case 18: DATO$ = FINFACTU107.MONEDA_CAE ' QUEDA PARA AGREGAR UNA FUNCION POR EL TIPO DE CAMBIO
           Case 19: DATO$ = TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4"))
           
           Case 20:           '1)Fecha Comprobante,2)Fecha de comienzo de Servicio, 3)Fecha Fin de Servicio, Fecha de Vencimento
                    FECHA$ = FECHATEXLAR$(FECHANUM(FORFACPRINT.Text9))
                    DIA$ = Mid(FECHA$, 1, 2): MES$ = Mid$(FECHA$, 4, 2): ANIO$ = Mid$(FECHA$, 7, 4)
                    DATO$ = ANIO$ + MES$ + DIA$
          End Select
        TX$ = TX$ + DATO$ + ETIQUETACIERRE$ ' CONCATENA EL RENGLON A IMPRIMIR
        DATO$ = ""
        Print #NARCHI2%, TX$ ' ESCRIBE EN EL ARCHIVO EL NUEVO TEXTO
     Loop
    '
    Print #NARCHI2%, "<Items>"

    For I& = 1 To ULTREG&("!CARDEFAC")
      XX$ = REGLEIDO$("!CARDEFAC", I&)
      CI% = CVI(Mid$(XX$, 5, 2))
      DESCRI$ = DESCRIT$(CI%)
      CAN = CVS(Mid$(XX$, 1, 4))
      PRECI# = CVS(Mid$(XX$, 429, 4))
      '
      PORDESCUEN = PORCEDESCU(FORFACPRINT.PORDESCU) 'MKS$(Mid$(XX$, 425, 4))
      IMPDESC = (PRECI# * PORDESCUEN) / 100
      PRECI# = ROUND(PRECI# - IMPDESC)
      '
      PRCI$ = FORMATNUM$(PRECI#, "F12.2")
      IMPORTE# = CAN * PRECI#
      IMPRT$ = FORMATNUM$(IMPORTE#, "F12.2")
      CONDI$ = "Codint = " & CI%
      NOMENC$ = VALOBADA("MASTER", "CODIGO_NCM", CONDI$, "NOMESS")
      COSEC$ = VALOBADA("MASTER", "CODIGO_SEC", CONDI$, "NOMESS")
      UNIMEX$ = VALOBADA("MASTER", "PRO_UMED", CONDI$, "NOMESS")
      COTIVA$ = VALOBADA("MASTER", "IVA_ID", CONDI$, "NOMESS")
      '
      Print #NARCHI2%, "<item>"
      Print #NARCHI2%, "<pro_codigo_ncm>" & TRIM$(NOMENC$) & "</pro_codigo_ncm>"
      Print #NARCHI2%, "<pro_codigo_sec>" & TRIM$(COSEC$) & "</pro_codigo_sec>"
      Print #NARCHI2%, "<pro_ds>" & TRIM$(DESCRI$) & "</pro_ds>"
      Print #NARCHI2%, "<pro_qty>" & CAN & "</pro_qty>"
      Print #NARCHI2%, "<pro_umed>" & TRIM$(UNIMEX$) & "</pro_umed>"
      Print #NARCHI2%, "<pro_precio_uni>" & TRIM$(PRCI$) & "</pro_precio_uni>"
      Print #NARCHI2%, "<imp_bonif>0</imp_bonif>"
      Print #NARCHI2%, "<imp_total>" & TRIM$(IMPRT$) & "</imp_total>"
      Print #NARCHI2%, "<iva_id>" & TRIM$(COTIVA$) & "</iva_id>"
      Print #NARCHI2%, "</item>"
    Next I&
    Print #NARCHI2%, "</Items>"
    Print #NARCHI2%, "</Cmp>"
    
    Close #NARCHI1%
    Close #NARCHI2%
    '
99 End Sub
Function MONEDA_CAE$()
    '
    MONE$ = ""
    If MONEMI% = 1 Then
         MONE$ = "PES"
      ElseIf MONEMI% = 2 Then
         MONE$ = "DOL"
      Else
         DEMONE$ = UCase(DEMONECO$(MONEMI%))
         If Left$(DEMONE$, 4) = "EURO" Then
              MONE$ = "060"
           ElseIf Left$(DEMONE$, 4) = "REAL" Then
              MONE$ = "012"
         End If
    End If
    '
99  MONEDA_CAE$ = MONE$
    '
End Function

Function TIPOX_CBTEX$()

    TIPOCOMP$ = UCase$(FORFACPRINT.Label7 + FORFACPRINT.Label13)
    
    Select Case TIPOCOMP$
        Case "FACTURA:E": VALOR$ = "19"
        Case "N.DÉBITO:E": VALOR$ = "20"
        Case "N.CRÉDITO:E": VALOR$ = "21"
    End Select
    '
    TIPOX_CBTEX$ = VALOR$
    '
End Function

Sub SOLICITARCAE1()
    FLAG% = 0
    cont% = 0
    CARPETOK$ = "C:\FLEXOFT\FACELEC\"
    If EXISTE%(CARPETOK$ + "token.txt") = 0 Then
      TXT$ = REPLACAR$(CARPETOK$, "\", "/")
      Call MENSERR(24, "Falta Archivo 'Token.txt'\en Carpeta " & TXT$)
      GoTo 99
    End If
    If EXISTE%(LUDAT$ + "consulta_cae_original.txt") = 0 Then
      TXT$ = REPLACAR$(LUDAT$, "\", "/")
      Call MENSERR(24, "Falta Archivo 'consulta_cae_original.txt'\en Carpeta " & TXT$)
      GoTo 99
    End If
    NARCHI1% = FILEOPEN%(LUDAT$ + "consulta_cae_original.txt", 1, 128) 'ABRE EL ARCHIVO DE LECTURA
    NARCHI2% = FILEOPEN%(CARPETOK$ + "consulta_cae.txt", 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA

    Do While Not EOF(NARCHI1%) ' RECORRE
       Line Input #NARCHI1%, XX$ ' LEE CADA REGISTRO
       cont% = cont% + 1
       pos% = InStr(XX$, ">")
       ETIQUETACIERRE$ = TRIM$(Mid$(XX$, pos% + 1))
       TX$ = Mid$(XX$, 1, pos%)
       
          Select Case cont%
           'Case 4: DATO$ = TRIM$(FORFACPRINT.Label18)          ' Cantidad de registros del detalle de la factura
           Case 4: DATO$ = "1"            ' corregido porque debe ser siempre 1
           Case 5: DATO$ = "1"            ' Tipo de Presentacion dada (0 servicios , 1 Bienes de Capital)
           Case 9: DATO$ = FINFACTU107.TIPO_DOC$      ' Tipo Documento (Cuit,Dni, etc) Identificador del comprador
           Case 10: DATO$ = FINFACTU107.NRO_DOC$      ' Numero Dni o Cuit
           Case 11: DATO$ = FINFACTU107.TIPO_CBTE$        ' Tipo de Comprobante(Factura A o B ,Nota de Credito,Nota de Debito)
           Case 12: DATO$ = TRIM$(Str(XVALO(FORFACPRINT.Label12))) ' Punto de Venta
           Case 13: DATO$ = "0"                                  ' Numero de comprobante inicial
           Case 14: DATO$ = "0"                                  ' Numero de comprobante hasta
           Case 15: DATO$ = TRIM$(FORFACPRINT.Total(6))          ' Importe Total
           Case 16: DATO$ = FINFACTU107.IMP_TOT_CONC$            ' Importe no gravados = C/Iva, entre ellos neto no gravado, las percepciones de iva e IIBB
           Case 17: DATO$ = FINFACTU107.IMP_NETO$                ' Importe Neto (IMPORTE NETO = IMPORTE NETO - IVA NO GRAVADO)
           Case 18: DATO$ = FINFACTU107.IMPTO_LIQ$            ' Impuesto Liquidado(la suma de los importes de iva  aplicados)
           Case 19: DATO$ = "0"           ' Impto_Liq_rni - siempre es cero  o percepcion a no categorizados
           Case 20: DATO$ = "0"           ' importe operaciones externas (No dieron informacion de que se trata)
           Case 21, 22, 23, 24:           '1)Fecha Comprobante,2)Fecha de comienzo de Servicio, 3)Fecha Fin de Servicio, Fecha de Vencimento
                    FECHA$ = FECHATEXLAR$(FECHANUM(FORFACPRINT.Text9))
                    DIA$ = Mid(FECHA$, 1, 2): MES$ = Mid$(FECHA$, 4, 2): ANIO$ = Mid$(FECHA$, 7, 4)
                    DATO$ = ANIO$ + MES$ + DIA$
          End Select
          ' NOTA !!!!************************************************
          ' EN LA FILA 24 SE DEBE PONER LA FECHA DE VENCIMIENTO DEL COMPROBANTE
          ' PROVISORIAMENTE SE PUSO LA MISMA DEL DIA
          '*********************************************************************
        TX$ = TX$ + DATO$ + ETIQUETACIERRE$ ' CONCATENA EL RENGLON A IMPRIMIR
        DATO$ = ""
        Print #NARCHI2%, TX$ ' ESCRIBE EN EL ARCHIVO EL NUEVO TEXTO
     Loop
    '
    Close #NARCHI1%
    Close #NARCHI2%
    '
99 End Sub


