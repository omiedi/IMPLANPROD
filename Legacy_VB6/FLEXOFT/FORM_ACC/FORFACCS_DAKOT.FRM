VERSION 5.00
Begin VB.Form FORFACCS 
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
      Picture         =   "FORFACCS.frx":0000
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
            TabIndex        =   135
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
               Left            =   3360
               TabIndex        =   163
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
               TabIndex        =   148
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
               TabIndex        =   147
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
               TabIndex        =   146
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
               TabIndex        =   145
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
               TabIndex        =   144
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
               TabIndex        =   139
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
               TabIndex        =   136
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
               Left            =   165
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
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00005500&
            Height          =   330
            Left            =   105
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
            TabIndex        =   141
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
            TabIndex        =   140
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
            TabIndex        =   138
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
            TabIndex        =   137
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
         Begin VB.Label Label32 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   420
            TabIndex        =   160
            Top             =   525
            Visible         =   0   'False
            Width           =   540
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
            TabIndex        =   143
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
            TabIndex        =   142
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
      Begin VB.PictureBox FORFADAKO 
         Height          =   3200
         Left            =   210
         ScaleHeight     =   3135
         ScaleWidth      =   10455
         TabIndex        =   164
         Top             =   1785
         Visible         =   0   'False
         Width           =   10515
         Begin VB.VScrollBar VScroll1 
            Height          =   3165
            Left            =   10260
            TabIndex        =   173
            Top             =   0
            Width           =   225
         End
         Begin VB.PictureBox FORFADAKS 
            Height          =   3180
            Left            =   -30
            ScaleHeight     =   3120
            ScaleWidth      =   10245
            TabIndex        =   165
            Top             =   -30
            Width           =   10300
            Begin VB.TextBox Descom 
               Alignment       =   2  'Center
               Height          =   255
               Index           =   0
               Left            =   7170
               TabIndex        =   174
               Top             =   -20
               Width           =   960
            End
            Begin VB.TextBox Tcodigo 
               Height          =   255
               Index           =   0
               Left            =   525
               TabIndex        =   169
               Text            =   " "
               Top             =   -20
               Width           =   2010
            End
            Begin VB.TextBox Tdescri 
               Height          =   675
               Index           =   0
               Left            =   2520
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   168
               Text            =   "FORFACCS.frx":0442
               Top             =   -20
               Width           =   3270
            End
            Begin VB.TextBox Tcanti 
               Alignment       =   1  'Right Justify
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   167
               Top             =   -20
               Width           =   540
            End
            Begin VB.TextBox Tpreca 
               Alignment       =   1  'Right Justify
               Height          =   255
               Index           =   0
               Left            =   5775
               TabIndex        =   166
               Top             =   -20
               Width           =   960
            End
            Begin VB.Label Pruneto 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   0
               Left            =   8120
               TabIndex        =   172
               Top             =   -15
               Width           =   960
            End
            Begin VB.Label Ineto 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   0
               Left            =   9080
               TabIndex        =   171
               Top             =   -15
               Width           =   1170
            End
            Begin VB.Label Dlista 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   0
               Left            =   6720
               TabIndex        =   170
               Top             =   -15
               Width           =   435
            End
         End
      End
      Begin VB.Frame Frame12 
         BorderStyle     =   0  'None
         Caption         =   "Frame12"
         Height          =   3375
         Left            =   210
         TabIndex        =   149
         Top             =   8000
         Visible         =   0   'False
         Width           =   10725
         Begin VB.CommandButton Command30 
            Caption         =   "Cortar"
            Height          =   260
            Left            =   5565
            TabIndex        =   159
            Top             =   3045
            Width           =   1170
         End
         Begin VB.CommandButton Command29 
            Caption         =   "Copiar"
            Height          =   260
            Left            =   6720
            TabIndex        =   158
            Top             =   3045
            Width           =   1170
         End
         Begin VB.CommandButton Command28 
            Caption         =   "Pegar"
            Height          =   260
            Left            =   7875
            TabIndex        =   157
            Top             =   3045
            Width           =   1275
         End
         Begin VB.CommandButton Command27 
            Caption         =   "Selecc Todo"
            Height          =   260
            Left            =   9135
            TabIndex        =   156
            Top             =   3045
            Width           =   1400
         End
         Begin VB.CommandButton Command26 
            Height          =   260
            Left            =   0
            TabIndex        =   155
            Top             =   3045
            Width           =   5565
         End
         Begin VB.PictureBox Picture6 
            BackColor       =   &H00E0E0E0&
            Height          =   435
            Left            =   0
            ScaleHeight     =   375
            ScaleWidth      =   10455
            TabIndex        =   151
            Top             =   0
            Width           =   10515
            Begin VB.Label Label28 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Descripción General - Especificaciones"
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
               Height          =   225
               Left            =   0
               TabIndex        =   152
               Top             =   105
               Width           =   10515
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
            Height          =   2640
            Left            =   1365
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   150
            Top             =   430
            Width           =   7890
         End
         Begin VB.Label TEPRUE 
            AutoSize        =   -1  'True
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
            Left            =   0
            TabIndex        =   161
            Top             =   0
            Visible         =   0   'False
            Width           =   120
         End
         Begin VB.Line Line10 
            X1              =   10520
            X2              =   10520
            Y1              =   0
            Y2              =   3280
         End
         Begin VB.Label Label31 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            Height          =   2640
            Left            =   0
            TabIndex        =   154
            Top             =   435
            Width           =   1380
         End
         Begin VB.Label Label30 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            Height          =   2640
            Left            =   9255
            TabIndex        =   153
            Top             =   435
            Width           =   1270
         End
      End
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
            Picture         =   "FORFACCS.frx":0444
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
            Picture         =   "FORFACCS.frx":0AAE
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
            Picture         =   "FORFACCS.frx":0DB8
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
      Begin VB.CommandButton Command50 
         BackColor       =   &H000000FF&
         Caption         =   "O"
         Height          =   255
         Left            =   7380
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   134
         Top             =   1000
         Visible         =   0   'False
         Width           =   255
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
         Picture         =   "FORFACCS.frx":10C2
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
            Picture         =   "FORFACCS.frx":13CC
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
            Picture         =   "FORFACCS.frx":16D6
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
            Picture         =   "FORFACCS.frx":19E0
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
            Picture         =   "FORFACCS.frx":1CEA
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
            Picture         =   "FORFACCS.frx":1FF4
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
            Picture         =   "FORFACCS.frx":2436
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
            Left            =   105
            Picture         =   "FORFACCS.frx":2740
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
            Picture         =   "FORFACCS.frx":288A
            Style           =   1  'Graphical
            TabIndex        =   69
            ToolTipText     =   "Opciones Varias - Anulación / Control Diferencia de Cambio"
            Top             =   4905
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
            Picture         =   "FORFACCS.frx":29D4
            Style           =   1  'Graphical
            TabIndex        =   68
            ToolTipText     =   "Acceso al Control General de Gestión Comercial"
            Top             =   3510
            Width           =   650
         End
         Begin VB.Image Image2 
            Height          =   390
            Left            =   -315
            Picture         =   "FORFACCS.frx":2CDE
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
            Picture         =   "FORFACCS.frx":4C00
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
            Picture         =   "FORFACCS.frx":4F0A
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
         Left            =   210
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
      Begin VB.Label BARRELAB 
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
         Height          =   220
         Left            =   5355
         TabIndex        =   162
         Top             =   0
         Visible         =   0   'False
         Width           =   1035
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
         Left            =   8925
         TabIndex        =   50
         ToolTipText     =   "Doble-Click Para Cambiar Punto de Venta"
         Top             =   495
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
Attribute VB_Name = "FORFACCS"
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
Dim FVENCIM
'
Dim CONTROFIS% ' Si 0 --> No Inicializado /// 1 --> Controlador Fiscal /// -1 --> Printer Común
Dim CANCECON%  ' Si 0 --> No Inicializado /// 1 --> Permite Cancela Contado /// -1 --> No Permite Cancela Contado
'
Dim CONSICLI%, PREFACTU$
Dim FORCARGA$
'
Dim DESELEC#, HASELEC#
'
Dim NUOCOMPRA$
Dim NROPEDIDO$
Dim NROREMICO$
'
Dim DATCLI$(32)
'
Dim NFCAM%    ' flag (default=0) que se pone en 1 si se ha cambiado manualmente en nro de factura
'
Dim MONECO$, TICAMBX#
Dim APLIVA%, RNI%
'
Dim COPIASREM%
'
Dim TASACLICK


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
       If CATIVAPRO% = 2 Or CATIVAPRO% = 3 Or CATIVAPRO% = 4 Or CATIVAPRO% = 6 Then
         APLIVA% = 0
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
   If BARRELAB.Caption = "BARRE" Then Exit Sub
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
   If BARRELAB.Caption = "BARRE" Then Exit Sub
   Command11.Enabled = False
   Call CIERRARCH("*.*")
   Call CONECRUN("ARCHIG04", "Archivos Maestros")
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
   If BARRELAB.Caption = "BARRE" Then Exit Sub
   Command13.Enabled = False
   COGEVE04.Show 1
   Command13.Enabled = True
End Sub

Sub Command14_Click()
   If VALICLIE% < 1 Then Exit Sub
   If Abs(XVALO(Total(6))) < 0.005 Then Exit Sub
   '
   Command14.Enabled = False
   '
   '\\\OT-07-0023-RG-24/01/07///
   FINFACTURA.Text7 = TRIM$(Str$(DEPOEXPE%))
   If CONTROFIS% = 1 Then
       '\\\OT-08-0120-FL-03/04/08///
       'Esto es para control de las controladoras fiscales que no permiten la falta de cuit o lo validan
       If TIFADOC% = 6 Then               ' NOTA DE CREDITO
          If VARIDOC% = 2 Then            ' N/C B
               NCX% = XVALO(FORFACCS.Text1)
               CUITT$ = REPLACAR$(CUITCLI$(NCX%), "-", "")
               If VALICUIT%(CUITT$) <= 0 Then
                   Call MENSERR(24, "C.U.I.T. no Válido o Inexistente\Verifique el Número de C.U.I.T. del Cliente ")
                   Text1.TEXT = ""
                   Command14.Enabled = True
                   Exit Sub
               End If
          End If
       End If
       '\\\OT-08-0120-FL-03/04/08///
       '
       FINCONFIS.Height = 2070
       FINCONFIS.Text7 = TRIM$(Str$(DEPOEXPE%))
       FINCONFIS.Text1 = ""
       FINCONFIS.Text2 = ""
       If CANCECON% = 1 Then
          FINCONFIS.Text1 = CONTROL$("MECOBRA", "CONTADO")
          FINCONFIS.Text2 = TRIM$(Total(6))
          FINCONFIS.Height = 4050
       End If
   End If
   '
   '\\\OT-07-0023-RG-FIN///
   If XVALO(TLABEL28) < 0 Then
      Call MENSERR(24, "Imposible Registrar este Documento.\Importe No Gravado no Admisible.")
      Command14.Enabled = True
      Exit Sub
   End If
   '
   ' 5/12/07 - SILCAR - ERROR DE IMPRESION DE IMPORTE EN NOTA DE CREITO TIPO B POR CH. RECH.
   If TIFADOC% > 2 Then         ' solo débitos y creditos
     If VARIDOC% = 2 Then       ' solo tipo 'B'
       If MODODOC% = 1 Then     ' solo texto libre
         ALIAPLI = XVALO(PORIVA(0))
         If XVALO(Total(2)) > 0 Then
           ALIAPLI = 100 * (XVALO(Total(3)) + XVALO(Total(4))) / XVALO(Total(2))
         End If
         For KU& = 1 To ULTREG&("!CARDEFAC")
           XXX$ = REGLEIDO$("!CARDEFAC", KU&)
           INETO1# = CVD(Mid$(XXX$, 433, 8))
           Call REPLA(XXX$, MKS$(ALIAPLI), 441, 4)
           Call REPLA(XXX$, MKD$(INETO1# * ALIAPLI / 100), 445, 8)
           Call REPLA(XXX$, MKD$(INETO1# * (1 + ALIAPLI / 100)), 453, 8)
           Call GRAREG("!CARDEFAC", XXX$, KU&)
         Next KU&
       End If
     End If
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
22     LMCR# = LIMICRE(NCLIE%)
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
' ojo aca grabar lugares de entrega
25 Call APERBASDAT("CLIEN")
   With Deudor12
     .FindFirst "Nume_Cli = " & NCLIE%
     If .NoMatch = False Then
       .Edit
       !Lentre0_Cli = Left$(LENTREGA(0) + " ", 48)
       !Lentre1_Cli = Left$(LENTREGA(1) + " ", 48)
       !Lentre2_Cli = Left$(LENTREGA(2) + " ", 48)
       .Update
     End If
   End With
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
   FORIPRE$ = FORELFAC$
   If FORIPRE$ = Null Then Exit Sub
   '
   ' control de renglones por hoja
   If CONTROL$("FACTUREM", "FORMULAR") = "HOJASUEL" Then
       HOJASUEL% = 1
       CARENGLO& = ULTREG&("!CARDEFAC")
       ITPORHOJA% = XVALO(ATRIFORM$(FORIPRE$, "MAX-ITEM"))
       If CARENGLO& > ITPORHOJA Then
          Call MENSERR(24, "Imposible Emitir este Documento.\Supera Máximo de Renglones por Hoja.")
          Command14.Enabled = True
          Exit Sub
       End If
   End If
   '
   If MODODOC% = 0 Then    ' Material Codificado
     If PREFACTU$ = "REMITO" Then     ' QUE ORDENE SOLO EN GUIDI
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
       Call PRIFACTURA
   End If
   Call CIERRARCH("*.*")
   Call LIBARCH("*.*")
   Call BAJALDISCO
   '
   If BARRELAB.Caption = "BARRE" Then
      BARRELAB.Caption = ""
      Unload Me
      Exit Sub
   End If
   '
   If OKX% > 0 Then
     If InStr(UCase$(Eco21(1)), "CONTADO") > 0 Then
       RECIBO04.Text1 = Text1
       RECIBO04.Text2 = Text2
       RECIBO04.Text3 = Text3
       RECIBO04.Text4 = Text4
       RECIBO04.Text5 = Text5
       RECIBO04.Text9 = Text9
       If ULTREG&("COBRADOR") > 0 Then
         XXX$ = REGLEIDO$("COBRADOR", 1)
         RECIBO04.Label22.Caption = Left$(TRIM$(Asc(Left$(XXX$, 1))) + " - " + Mid$(XXX$, 2, 30), 19)
       End If
       RECIBO04.Show 1
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
   Command14.Enabled = True
End Sub

Private Sub Command23_Click()
    If BARRELAB.Caption = "BARRE" Then Exit Sub
    NDCLIEN.Show 1
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

Private Sub Command31_Click()

    OBX$ = OBJPLA$("CHEQUERECHA", VDEF$)

End Sub

Private Sub Descom_Change(Index As Integer)
    Call CALPRENETO(Index)
End Sub

Private Sub Descom_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
       Call AGREGARENGLON(Index)
       Tcanti(Index + 1).SetFocus
   End If
End Sub

Private Sub Label12_Change()
   Call NUFACTURA
End Sub

Private Sub Label12_DblClick()
   '
   If Label12 = "0001" Then
       Label12 = "0002"
     ElseIf Label12 = "0002" Then
       Label12 = "0003"
     Else
       Label12 = "0001"
   End If
   Call GRACONTROL(IDENTER$, "PREFIFA", Label12)
   '
End Sub


Private Sub Label14_DblClick()
   If BARRELAB.Caption = "BARRE" Then Exit Sub
   Call Text14_DBLCLICK
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


Private Sub Tcanti_GotFocus(Index As Integer)
     Tdescri(Index).ZOrder 0
     Tcanti(Index).TEXT = " " + Tcanti(Index).TEXT
     Tcanti(Index).SelStart = 0
     Tcanti(Index).SelLength = 1 + Len(Tcanti(Index))
     Tcanti(Index).Refresh
End Sub

Private Sub Tcanti_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = 40 Then ' FLECHA ABAJO
       Call AGREGARENGLON(Index)
       Tcanti(Index + 1).SetFocus
     ElseIf KeyCode = 13 Or KeyCode = 39 Then
       Tcodigo(Index).SetFocus
     ElseIf KeyCode = 38 Then
       If Index > 0 Then
         Tcanti(Index - 1).SetFocus
       End If
   End If
End Sub

Private Sub Tcodigo_Change(Index As Integer)
    CI% = CODINT%(Tcodigo(Index))
    If CI% > 0 Then
        Tdescri(Index) = DESCRIT0$(CI%)
        Tdescri(Index).ZOrder 0
        Tdescri(Index).SetFocus
        PRELIX = PRELISTA(ARCHIPRE$, CI%)
        If TRIM$(Tpreca(Index)) = "" Then Tpreca(Index) = FORMATNUM$(PRELIX, "F10.2")
           NULISTA% = XVALO(TEXT21(0))
           CLIEN% = XVALO(Text1)
        Dlista(Index) = FORMATNUM(DESCULISPRE(NULISTA%), "F4.1")
        Descom(Index) = DESCOMBINA$(CLIEN%, CI%)
        Call CALPRENETO(Index)
      Else
        On Error Resume Next
        Tpreca(Index) = 0
        On Error GoTo 0
    End If
End Sub

Sub CALPRENETO(Index)
     DESCULISTA# = XVALO(Dlista(Index))
     KOEFILISTA# = 1 - DESCULISTA# / 100 ' Calcula el coeficiente de Dto. de lista
         ' Trae la combinacion del descuento que tiene el articulo/Cliente
     DESCOMBIN$ = XVALO(Descom(Index))
     DESCUCOMBI# = ROUND(PORCEDESCU(DESCOMBIN$), 1)
     KOEFICOMBI# = 1 - DESCUCOMBI# / 100 'Calcula el coeficiente
         ' Determina el descuento general
     DESCUGENER# = ROUND#(100 * (1 - KOEFILISTA# * KOEFICOMBI#), 1)
     PREUNITA = XVALO(Tpreca(Index)) * (1 - DESCUGENER# / 100)
     If Index < Pruneto.Count Then Pruneto(Index) = FORMATNUM$(PREUNITA, "F10,2")
     If Index < Ineto.Count Then Ineto(Index) = FORMATNUM(PREUNITA * XVALO(Tcanti(Index)), "F12,2")
     '
     SUMATO# = 0
     For KKI% = 1 To Ineto.Count
        SUMATO# = SUMATO# + XVALO(Ineto(KKI% - 1))
     Next KKI%
     Total(0) = FORMATNUM$(SUMATO#, "F12.2")
     '
End Sub
'
Private Sub Tcodigo_DblClick(Index As Integer)
     FORSELEC.SINI = Str$(Top + (Height - ScaleHeight) + FORFADAKO.Top + Tcodigo(Index).Top + Tcodigo(Index).Height)
     FORSELEC.LINI = Str$(Left + FORFADAKO.Left + Tcodigo(Index).Left)
     FORSELEC.ABUSQUE = TRIM$(Tcodigo(Index))
     '
     DELISECO$ = ""
     If Left$(UCase$(ARCHIPRE$), 5) = "LIPRE" Then
        DELISECO$ = DENOLISTA$(XVALO(Mid$(ARCHEK$, 6, 3)))
     End If
     Call SELECHO(ARCHIPRE$ + "/" + DELISECO$)
     Tcodigo(Index) = VALACT1$(ARCHIPRE$)
End Sub

Private Sub Tcodigo_GotFocus(Index As Integer)
     Tdescri(Index).ZOrder 0
End Sub

Private Sub Tcodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
     If KeyCode = 113 Then
        Call Tcodigo_DblClick(Index)
     End If
End Sub

Private Sub Tdescri_GotFocus(Index As Integer)
    Tdescri(Index).ZOrder 0
    Tdescri(Index).SelStart = 1 + Len(Tdescri(Index).TEXT)
End Sub

Private Sub Tdescri_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
       Tpreca(Index).SetFocus
    End If
End Sub

Private Sub Text14_DBLCLICK()
    '
    If BARRELAB.Caption = "BARRE" Then Exit Sub
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
    SQLX$ = SQLX$ + " AND (TRIM(DoSECLar) = '' "
    SQLX$ = SQLX$ + " OR DoSecLar LIKE 'P/F*' ) "
    SQLX$ = SQLX$ + " ORDER BY NUPEDCLI ASC "
    On Error GoTo 0
    Set pedianuittemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    JJ& = 0
    With pedianuittemp
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        TTXX$ = REGBLAN("INDIFAPE")
        Do While Not .EOF
         If PEDANTI& <> !NUPEDCLI Then
            NRODES& = !NUPEDCLI
            CI% = !Cod_Inte
            If NRODES& <= 0 Then GoTo 15
          TTXX$ = MKS$(NRODES&)
            NOCO$ = NOCOMCLI(NRODES&)
          If NOCO$ <> "" Then
            TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
          End If
            With PedEnca
              .FindFirst "NroPed = " & NRODES& & ""
              If .NoMatch = False Then
                FEX = CVINT(PedEnca!FECHEMIS)
              End If
            End With
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
    End With
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
   If BARRELAB.Caption = "BARRE" Then
      Command14.SetFocus
      Exit Sub
   End If
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

Private Sub Command17_Click()
   If BARRELAB.Caption = "BARRE" Then Exit Sub
   Command17.Enabled = False
   EPAC$ = TRIM$(UCase$(EMPREAC$))
'   If InStr(EPAC$, "DOSIVA") > 0 Then
'       CONPEDID.Show 1
'     Else
       Call SHOWPEDIPEN
'   End If
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
    BARRELAB.Caption = "CANCEL"
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_LostFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Command20_Click()
   If BARRELAB.Caption = "BARRE" Then Exit Sub
   Command20.Enabled = False
   OPVARFA04.Show 1
   Command20.Enabled = True
End Sub

Private Sub Command22_Click()
    If BARRELAB.Caption = "BARRE" Then Exit Sub
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
    RECUEC04.Text1.TEXT = Text1.TEXT
    RECUEC04.Show 1
    Command4.Enabled = True
End Sub

Sub Command5_Click()
    Command5.Enabled = False
    Call CONECRUN("FSETUP04/OPFACTUR", "Configuración de Funcionamiento")
    Call FINAL("")
    Command5.Enabled = True
End Sub

Sub Command6_Click()
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
         Else
           Call CARDETPEDIPEN   ' CARGA DETALLE DE PEDIDOS PENDIENTES
       End If
       Call CARDETFAC       ' POR PEDIDO Y SIN CONSIGNACION PENDIENTE
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
   Call CARDETFAC
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
3  DES1# = CDbl(CVDAT(DES%))
     DESELEC# = DES1#
   HAS1# = CDbl(CVDAT(HAS%))
     HASELEC# = HAS1#
   '
   LPRE% = XVALO(TEXT21(0).TEXT)
   NCLIE% = XVALO(Text1.TEXT)
   ARCHIPRE$ = TRIM$(Str$(LPRE%))
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
   SQLX$ = SQLX$ + " AND ((ISNULL(DoSecLar) = true) or (DoSecLar = '') OR (DoSecLar LIKE 'P/F*')) "
   SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) >= " & DES1# & " "
   SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) <= " & HAS1# & " "
   '
   Set CONSITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
   With CONSITMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       RECONSI$ = REGBLAN("CONSICLI")
       J& = 0
       Do While .EOF = False
         CODIMAT% = !Cod_Inte
         CAN = !CantSalid - !CantIngre
         NPEDID& = 0: If IsNull(!NUPEDCLI) = False Then NPEDID& = !NUPEDCLI
         NUMEPED& = 0: PREPED = 0
           If NPEDID& < 1 Then
             EPAC$ = TRIM$(UCase$(EMPREAC$))
             If InStr(EPAC$, "GUIDI") > 0 Then
                FEREMIT% = CVINT%(!FechMov)
                TITUCON$ = TRIM$(!DoPriCor) + " - Cant: " + TRIM$(Str$(CAN)) + " - Cod: " + TRIM$(CODEXT$(CODIMAT%)) + " - " + DESCRIT$(CODIMAT%)
                Call RECARPED(NCLIE%, CODIMAT%, FEREMIT%, NUMEPED&, PREPED, TITUCON$)
                NPEDID& = NUMEPED&
             End If
           End If
5        z$ = REGBLAN$("!CARDEFAC")
         Call REPLA(z$, MKS$(CAN), 1, 4)      ' cantidad
           COD$ = CODEXT$(CODIMAT%)
         Call REPLA(z$, MKI$(CODIMAT%), 5, 2)
         Call REPLA(z$, DESCRIT$(CODIMAT%), 7, 64)
           NREMIT$ = !DoPriCor
         Call REPLA(z$, NREMIT$, 387, 14) ' numero de remito
           FEREMIT% = CVINT%(!FechMov)
         Call REPLA(z$, MKI$(FEREMIT%), 419, 2)
         Call REPLA(z$, NUPLANO$(CODIMAT%), 315, 24)
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
7          MONEMA% = !MoneVal
           If MONEMA% <= 1 Then
             MONEPE% = MONEMIPED%(NPEDID&)    ' RECUPERA MONEDA DE PEDIDO
             If MONEPE% > MONEMA% Then MONEMA% = MONEPE%
           End If
           If MONEMA% > ULTREG&("TAMONED") Then MONEMA% = 1 ' ATAJA FUERA DE RANGO
           '
           COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
           PRECIO = !ValoUnit * COEFCON
             '
8            If PRECIO < 0.00005 Then PRECIO = PREPED
             If PRECIO < 0.00005 Then PRECIO = PRELISTA#(ARCHIPRE$, CODIMAT%)
           PRELI = PRELISTA#(ARCHIPRE$, CODIMAT%)
           If PRELI < PRECIO Then PRELI = PRECIO
           PORDESCUEN = 0: If PRELI >= 0.0001 Then PORDESCUEN = 100 * (PRELI - PRECIO) / PRELI
         Call REPLA(z$, MKS$(PRELI), 421, 4)
         Call REPLA(z$, MKS$(PORDESCUEN), 425, 4)
         Call REPLA(z$, MKS$(PRECIO), 429, 4)
           IMPO# = CDbl(PRECIO) * CAN
         Call REPLA(z$, MKD$(IMPO#), 433, 8)
         Call REPLA(z$, MKS$(PORCEIVA(CODIMAT%)), 441, 4)
           NORIT% = !NuOrItem
         'Call REPLA(z$, MKI$(NORIT%), 503, 2)
         Call REPLA(z$, MKI$(NORIT%), 401, 2)    ' resuelve enredo entre pedido y remito
         Call REPLA(z$, MKS$(0), 505, 4)
         Call REPLA(z$, MKS$(0), 509, 4)
         Call REPLA(z$, TRIM$(Str$(NPEDID&)), 371, 16)
           NUOCOX$ = NUOCOCLI$(NPEDID&, CODIMAT%)
10       Call REPLA(z$, NUOCOX$, 355, 16)
         '
         Call REGAPP("!CARDEFAC", z$)
         CAITCAR% = CAITCAR% + 1
         '
       .MoveNext
       Loop
12   End If
   End With
   '
   JJ& = 0
   For U& = 1 To ULTREG&("!CARDEFAC")
     z$ = REGLEIDO$("!CARDEFAC", U&)
     If Abs(CVS(Mid$(z$, 1, 4))) >= 0.01 Then
       JJ& = JJ& + 1
       Call GRAREG("!CARDEFAC", z$, JJ&)
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
   ARCHIPRE$ = TRIM$(Str$(LPRE%))
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
   SQLX$ = SQLX$ + " AND (TRIM(DoSECLar) = '' "
   SQLX$ = SQLX$ + " OR DoSecLar LIKE 'P/F*' ) "
   SQLX$ = SQLX$ + " AND NuPedCli = " & NPED& & " "
   '
   On Error GoTo 0
   Set CONSITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
   With CONSITMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       RECONSI$ = REGBLAN("CONSICLI")
       J& = 0
       Do While .EOF = False
         CODIMAT% = !Cod_Inte
         CAN = !CantSalid - !CantIngre
         NPEDID& = 0: If IsNull(!NUPEDCLI) = False Then NPEDID& = !NUPEDCLI
         NUMEPED& = 0: PREPED = 0
           If NPEDID& < 1 Then
             FEREMIT% = CVINT%(!FechMov)
             TITUCON$ = TRIM$(!DoPriCor) + " - Cant: " + TRIM$(Str$(CAN)) + " - Cod: " + TRIM$(CODEXT$(CODIMAT%)) + " - " + DESCRIT$(CODIMAT%)
             Call RECARPED(NCLIE%, CODIMAT%, FEREMIT%, NUMEPED&, PREPED, TITUCON$)
             NPEDID& = NUMEPED&
           End If
5        z$ = REGBLAN$("!CARDEFAC")
         Call REPLA(z$, MKS$(CAN), 1, 4)      ' cantidad
           COD$ = CODEXT$(CODIMAT%)
         Call REPLA(z$, MKI$(CODIMAT%), 5, 2)
         Call REPLA(z$, DESCRIT$(CODIMAT%), 7, 64)
           NREMIT$ = !DoPriCor
         Call REPLA(z$, NREMIT$, 387, 14) ' numero de remito
           FEREMIT% = CVINT%(!FechMov)
         Call REPLA(z$, MKI$(FEREMIT%), 419, 2)
         Call REPLA(z$, NUPLANO$(CODIMAT%), 315, 24)
           '
           MONEMA% = !MoneVal
           If MONEMA% <= 1 Then
             MONEPE% = MONEMIPED%(NPEDID&)    ' RECUPERA MONEDA DE PEDIDO
             If MONEPE% > MONEMA% Then MONEMA% = MONEPE%
           End If
           '
           COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
           PRECIO = !ValoUnit * COEFCON
             If PRECIO < 0.00005 Then PRECIO = PREPED
             If PRECIO < 0.00005 Then PRECIO = PRELISTA#(ARCHIPRE$, CODIMAT%)
           PRELI = PRELISTA#(ARCHIPRE$, CODIMAT%)
           If PRELI < PRECIO Then PRELI = PRECIO
           PORDESCUEN = 0: If PRELI >= 0.0001 Then PORDESCUEN = 100 * (PRELI - PRECIO) / PRELI
         Call REPLA(z$, MKS$(PRELI), 421, 4)
         Call REPLA(z$, MKS$(PORDESCUEN), 425, 4)
         Call REPLA(z$, MKS$(PRECIO), 429, 4)
           IMPO# = CDbl(PRECIO) * CAN
         Call REPLA(z$, MKD$(IMPO#), 433, 8)
         Call REPLA(z$, MKS$(PORCEIVA(CODIMAT%)), 441, 4)
           NORIT% = !NuOrItem
         'Call REPLA(z$, MKI$(NORIT%), 503, 2)
         Call REPLA(z$, MKI$(NORIT%), 401, 2)   ' para resolver el enredo entre remito y pedido
         Call REPLA(z$, MKS$(0), 505, 4)
         Call REPLA(z$, MKS$(0), 509, 4)
         Call REPLA(z$, TRIM$(Str$(NPEDID&)), 371, 16)
           NUOCOX$ = NUOCOCLI$(NPEDID&, CODIMAT%)
10       Call REPLA(z$, NUOCOX$, 355, 16)
         '
         Call REGAPP("!CARDEFAC", z$)
         CAITCAR% = CAITCAR% + 1
         '
       .MoveNext
       Loop
12   End If
   End With
   '
   JJ& = 0
   For U& = 1 To ULTREG&("!CARDEFAC")
     z$ = REGLEIDO$("!CARDEFAC", U&)
     If Abs(CVS(Mid$(z$, 1, 4))) >= 0.01 Then
       JJ& = JJ& + 1
       Call GRAREG("!CARDEFAC", z$, JJ&)
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
   FEREMIX = CDbl(CVDAT(FEREMIT%))
   Call APERBASDAT("PEDCLI")
   SQLX$ = " SELECT * "
   SQLX$ = SQLX$ + " FROM PEDDETA "
   SQLX$ = SQLX$ + " WHERE NroClie = " & NCLIE% & " "
   SQLX$ = SQLX$ + " AND Status < 8 "    ' NO ANULADOS NI CUMPLIDOS
   SQLX$ = SQLX$ + " AND STATUS >= 0 "   ' APROBADOS
   SQLX$ = SQLX$ + "AND CodiInt = " & CIPD% & " "
   'SQLX$ = SQLX$ + "AND CDBL(INT(FechEmis)) <= " & FEREMIX & " "
   SQLX$ = SQLX$ + " AND (((CanPed - CantEnt) >= 0.05) OR (TipoPed = 'ABIERTO' AND CDBL(FinVigen) >= " & FEREMIX & ")) "
   SQLX$ = SQLX$ + "ORDER BY NroPed ASC"
   '
   Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
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
         NUMEPED& = !NroPed
         PREPED = !PreUnid
         ENCUPE% = 1
       ElseIf CAREGIS& > 1 Then
         .MoveFirst
         JJ& = 0
         Do While (.EOF = False)
           TTXX$ = MKS$(!NroPed)
           TTXX$ = TTXX$ + AJUSTI$(!NroOcom, 16) + "  "
              FEX = CVINT(!FECHEMIS)
           TTXX$ = TTXX$ + FECHATEX$(FEX) + "  "
              FEX = CVINT(!FESOLENT)
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
           z$ = Space$(60) + MKS$(0)
           Call REPLA(z$, DOCABX$(KKI%), 1, 18)
           Call REPLA(z$, FECHATEX$(FEMI), 19, 8)
           Call REPLA(z$, "$" + FORMATNUM$(SALPES#(KKI%), "F11.2"), 29, 12)
           Call REPLA(z$, "u$s" + FORMATNUM$(SALDOL#(KKI%), "F12.2"), 42, 16)
           Call REPLA(z$, MKS$(REDOCA&(KKI%)), 61, 4)
           Call REGAPP("SELDOCAP", z$)
         End If
       End If
     ElseIf TIFADOC% = 3 Then  ' N.DEBITO
       If SALDOL#(KKI%) > 0 Then
         If SALPES#(KKI%) <= 0.005 Then
           XXC$ = REGLEIDO$("CARBIMON", REDOCA&(KKI%))
           FEMI = CVI(Mid$(XXC$, 1, 2))
           z$ = Space$(60) + MKS$(0)
           Call REPLA(z$, DOCABX$(KKI%), 1, 18)
           Call REPLA(z$, FECHATEX$(FEMI), 19, 8)
           Call REPLA(z$, "$" + FORMATNUM$(SALPES#(KKI%), "F11.2"), 29, 12)
           Call REPLA(z$, "u$s" + FORMATNUM$(SALDOL#(KKI%), "F12.2"), 42, 16)
           Call REPLA(z$, MKS$(REDOCA&(KKI%)), 61, 4)
           Call REGAPP("SELDOCAP", z$)
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
    '
    Dim RESP%
    Dim NUMEOCXX$
    '
    If BARRELAB.Caption = "BARRE" Then Exit Sub
    Command8.Enabled = False
    modo% = COMALTER%("Opciones de Salida\\Pantalla\Impresora\Cancelar")
    If modo% < 1 Or modo% > 2 Then GoTo 99
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
    DES1# = CDbl(CVDAT(DES%))
    HAS1# = CDbl(CVDAT(HAS%))
    '
    If modo% = 1 Then GoTo 56
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
    SQLX$ = " SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE (CodiMovi = '" & CMO$ & "' "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
    SQLX$ = SQLX$ + " AND DoPriLar <> '' "
    SQLX$ = SQLX$ + " AND ((ISNULL(DoSecLar) = true) or (DoSecLar = '') OR (DoSecLar LIKE 'P/F*')) "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) >= " & DES1# & " "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) <= " & HAS1# & " "
    'PARA EVITAR COMPROBANTES ARRUINADOS
    SQLX$ = SQLX$ + " AND LEFT$(UCASE$(REFERCOR),4) <> 'ANUL' "
    'FILTRO CLIENTE
    If modo% <> 1 Then
      SQLX$ = SQLX$ + " AND Nume_Clie >= " & CLINFE & ""
      SQLX$ = SQLX$ + "AND Nume_Clie<= " & CLISUP & ""
    End If
    '
    Set CONSITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
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
         Call REPLA(RECONSI$, MKI$(CVINT(!FechMov)), 3, 2)
         Call REPLA(RECONSI$, "", 5, 18)
         Call REPLA(RECONSI$, !DoPriCor, 23, 12)
         Call REPLA(RECONSI$, MKI$(!Cod_Inte), 35, 2)
         Call REPLA(RECONSI$, MKD$(!CantSalid), 37, 8)
         Call REPLA(RECONSI$, MKD$(!CantIngre), 45, 8)
         Call REPLA(RECONSI$, MKD$(!CantSaldo), 53, 8)
         Call REPLA(RECONSI$, !ReferCor, 61, 48)
         Call REPLA(RECONSI$, String$(12, 0), 109, 12)
         Call REPLA(RECONSI$, MKS$(XVALO(Mid(!DoPriCor, 5))), 109, 4)
         Call REPLA(RECONSI$, MKI$(CVINT(!FeReMovi)), 122, 2)
          HOR = Format(!FeReMovi, "HH:MM:SS")
         Call REPLA(RECONSI$, MKI$(CVINT(HOR)), 124, 2)
         Call REPLA(RECONSI$, MKI$(!NumUsuar), 127, 2)
         '
         Call REPLA(RECONSJ$, MKI$(NCLIE%), 1, 2)
         Call REPLA(RECONSJ$, MKI$(!Cod_Inte), 3, 2)
         Call REPLA(RECONSJ$, !DoPriCor, 5, 12)
         Call REPLA(RECONSJ$, !DoPriLar, 17, 18)
         Call REPLA(RECONSJ$, MKI$(CVINT(!FechMov)), 35, 2)
         If IsNull(!NUPEDCLI) Then
            Call REPLA(RECONSJ$, MKS$(0), 37, 4)
         Else
            Call REPLA(RECONSJ$, MKS$(!NUPEDCLI), 37, 4)
         End If
         If IsNull(!NUPEDCLI) Then
             NUMEOCXX$ = ""
           Else
             NUMEOCXX$ = NUOCOCLI$(!NUPEDCLI, !Cod_Inte)
         End If
         Call REPLA(RECONSJ$, NUMEOCXX$, 41, 24)
         Call REPLA(RECONSJ$, MKD$(!CantSalid), 65, 8)
         Call REPLA(RECONSJ$, MKD$(!CantIngre), 73, 8)
         Call REPLA(RECONSJ$, MKD$(!CantSaldo), 81, 8)
         Call REPLA(RECONSJ$, MKD$(!ValoUnit), 89, 8)
           TOTI# = !ValoUnit * (!CantSalid - !CantIngre)
         Call REPLA(RECONSJ$, MKD$(TOTI#), 97, 8)
         '
         J& = J& + 1
         Call GRAREG("CONSICLI", RECONSI$, J&)
         Call GRAREG("REPENFAC", RECONSJ$, J&)
      .MoveNext
      Loop
    End With
    '
    Call SETULTREG("CONSICLI", J&)
    Call FILESORT("CONSICLI", "", 1, 6, "ASC")
    Call CIERRARCH("CONSICLI")
    '
    Call SETULTREG("REPENFAC", J&)
    Call FILESORT("REPENFAC", "", 1, 6, "ASC")
    Call CIERRARCH("REPENFAC")
    '
    If modo% = 1 Then
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
    Refresh
    Frame1.Refresh
    DoEvents
    '\\\OT-05-0419-WAR-23/06/05///
    If CONTROL$("FACTURA", "SUPCONF") <> "SI" Then
       If CONTROL$("FACTURA", "SALIDA") <> "CFISCAL" Then
          Call CARGAPUERTOS
       End If
    End If
    '
    If EXISTE%("PEDIDO04.EXE") > 0 Then
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
    If CATIVAPRO% = 2 Or CATIVAPRO% = 3 Or CATIVAPRO% = 4 Or CATIVAPRO% = 6 Then
       Label13.Caption = "C": PIVADOC% = 3
    End If
    UPREFI$ = TRIM$(CONTROL$(IDENTER$, "PREFIFA"))
    If UPREFI$ = "" Then UPREFI$ = "0001"
    Label12.Caption = UPREFI$
    '
    Printer.TrackDefault = False
    '
    VERLICRE% = 0
    If CONTROL$("FACTURA", "VERLICRE") = "SI" Then VERLICRE% = 1
    Call LEETASIVA
    '
    If FORCARGA$ = "" Then
       FORCARGA$ = "NORMAL"
       If CONTROL$("FACTUMAI", "FORMACAR") = "SINREMITO" Then
          FORCARGA$ = "SINREMITO"
       End If
    End If
    '
    MUECOMHAS% = 0
    If CONTROFIS% = 0 Then
       CONTROFIS% = (-1)
       If CONTROL$("FACTURA", "SALIDA") = "CFISCAL" Then
          CONTROFIS% = 1
          FORCARGA$ = "SINREMITO"
          Check1.Visible = True
          ' INICIALIZAR CONTROLADOR FISCAL
          Call FORHASAR.INICIA
          '
       End If
    End If
    '
    If CANCECON% = 0 Then
       CANCECON% = (-1)
       If CONTROL$("FACTURA", "CANCECON") = "SI" Then
          CANCECON% = 1
       End If
    End If
    '
    Call NUFACTURA
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
    If PREFACTU$ = "" Then
      PREFACTU$ = TRIM$(CONTROL("FACTURA", "PREFACTU"))
      If PREFACTU$ = "" Then PREFACTU$ = "REMITO"   ' si no ORDESPA / PEDIDO
    End If
    '
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

Sub Label6_Click()
   If BARRELAB.Caption = "BARRE" Then
      If Label6.Caption = "(T.Libre)" Then
        Exit Sub
      End If
   End If
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

Sub Label7_Click()
   If BARRELAB.Caption = "BARRE" Then
      If Label7.Caption = "Factura:" Then
        Exit Sub
      End If
   End If
   If ULTREG&("!CARDEFAC") > 0 Then Exit Sub
   If Label7.Caption = "Factura:" Then
       TIFADOC% = 3: Label7.Caption = "N.Débito:"
       Label6.Caption = "(T.Libre)": MODODOC% = 1
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

Private Sub List1_DBLClick()
   If Command14.Enabled = True Then
      For U& = 1 To ULTREG&("!CARDEFAC")
        X$ = REGLEIDO$("!CARDEFAC", U&)
        NPEDID& = XVALO(TRIM$(Mid$(X$, 371, 16)))
        If NPEDID& <= 1 Then       ' NO VINO DE PEDIDO
           PRELI = CVS(Mid$(X$, 421, 4))
           Call REPLA(X$, MKS$(PRELI), 429, 4)
           Call GRAREG("!CARDEFAC", X$, U&)
        End If
      Next U&
      Call CARDETFAC
   End If
End Sub

Private Sub List3_DblClick()
   Call DIFPREC2
End Sub

Private Sub MONEMIS_DblClick()
   '
   If BARRELAB.Caption = "BARRE" Then Exit Sub
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
     PRELI = ROUND(COEFCON * CVS(Mid$(X$, 421, 4)), 4)
     PRECIO = ROUND(COEFCON * CVS(Mid$(X$, 429, 4)), 4)
     IMPO# = ROUND(CDbl(CAN * PRECIO), 2)
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
End Sub

Private Sub PORIVA_Click(Index As Integer)
   PORDESCU.SetFocus
End Sub

Private Sub Text1_Change()
   Call BLANFORMU
End Sub

Private Sub Text1_DblClick()
  If BARRELAB.Caption = "BARRE" Then Exit Sub
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
   If BARRELAB.Caption = "BARRE" Then
      Command14.SetFocus
      Exit Sub
   End If
   Text1.TEXT = TRIM$(Text1.TEXT)
   Text1.SelStart = 0
   Text1.SelLength = Len(Text1.TEXT)
   MARCONOTA.Visible = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  If BARRELAB.Caption = "BARRE" Then Exit Sub
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
   If BARRELAB.Caption = "BARRE" Then Exit Sub
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
    If BARRELAB.Caption = "BARRE" Then Exit Sub
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

Private Sub Text13_GOTFOCUS()
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
   If BARRELAB.Caption = "BARRE" Then
      Command14.SetFocus
      Exit Sub
   End If
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
  End If
End Sub

Private Sub Text19_DblClick()
   '\\\OT-06-0570-RG-27/12/06///
   Call Command24_Click
   '\\\OT-06-0570-RG-FIN///
End Sub

Private Sub Text2_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text20_Change()
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
        Text19 = CVI(Mid$(XX$, 195, 2))
      End If
    End If
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
            ARCHIPRE$ = "MASTER"
          Else
            Eco21(Index).TEXT = ECOARCH$("LISTAS", MKI$(XVALO(TEXT21(Index).TEXT)))
            ARCHIPRE$ = TRIM$(TEXT21(0).TEXT)
            While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
            ARCHIPRE$ = "LIPRE" + ARCHIPRE$
        End If
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
   Text15 = ""
   Text16 = ""
   Label16 = ""
   Label26 = ""
   Text17 = "" 'LIMPIA RET-IVA
   IEMBALA = ""
   Text19 = ""
   Text20 = ""
   Label25 = ""
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
   '
   ' PONE LETRA POR OMISIÓN
   Label13.Caption = "A"
   If CATIVAPRO% = 2 Or CATIVAPRO% = 3 Or CATIVAPRO% = 4 Or CATIVAPRO% = 6 Then
      Label13.Caption = "C"
   End If
   '\\\OT-5-506-FG-02/09/05
   ' PONE PREFIJO POR OMISION AL 0001
   UPREFI$ = TRIM$(CONTROL$(IDENTER$, "PREFIFA"))
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
   For KKI% = 0 To 2
      TEXT21(KKI%).TEXT = ""
      Eco21(KKI%) = ""
      BASIVA(KKI%) = ""
      LENTREGA(KKI%) = ""
   Next KKI%
   '
   PORDESCU.TEXT = ""
   For KK% = 0 To 6
      Total(KK%).TEXT = ""
   Next KK%
   TLABEL28 = ""
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
End Sub
'
Sub CARDACLI()
    '
    Screen.MousePointer = 11
    '
    NC% = XVALO(Text1.TEXT): Text1.Refresh
    If NC% > 0 Then
      '
      '\\\OT-06-0607-OD-28/07/08///
      'EVALUA SI LA CUENTA NO ESTA SUSPENDIDA
      '
      If LIMICRE(NC%) < (-0.5) Then
          Call MENSERR(24, "Cuenta Suspendida.")
          Text1 = ""
          Call BLANFORMU
          GoTo 99
      End If
      '\\\OT-06-0607-OD-FIN///
      '
      If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
        Text2.TEXT = RASOCLI$(NC%): Text2.Refresh
        Text3.TEXT = DOMICLI$(NC%): Text3.Refresh
        Text4.TEXT = CPOSCLI$(NC%): Text4.Refresh
        Text5.TEXT = LOCACLI$(NC%): Text5.Refresh
        Label26 = CODICLI$(NC%)
        Text7.TEXT = CSTRING$(MKS$(LIMICRE(NC%)), 3, 10, 2, 2)
          Text7.Refresh
        Text12.TEXT = TRIM$(POSIVAC$(NC%)) + " - CUIT " + CUITCLI$(NC%)
          Text12.Refresh
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
        '\\\OT-5-506-FG-02/09/05
        ' PARA QUE TOME EN CUENTA EL PREFIJO DE LAS FACTURAS
        ' DE EXPORTACION
        If PPVI% = 5 Or PPVI% = 7 Then
          Label12.Caption = PREFEXPO$()
        End If
        '\\\OT-5-506-FG-FIN
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
        Label32 = "0"      ' significa que los datos son los de la base de datos de clientes
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
        If POREMBAL(NCLIE%) > 0 Then
          CUEMBA$ = CONTROL$("SVD451S", "CUEVE-12")
          If CUEMBA$ = "" Or CUECOINT%(CUEMBA$) < 1 Then
            Call MENSERR(24, "Falta Definir Cuenta Contable\Correspondiente a Embalajes.\   \FLEXOFT Utilizará la Cuenta de Ventas Netas.")
          End If
        End If
        '
      End If
    End If
99  Screen.MousePointer = 1
    '
End Sub
    
Sub CARLETDOC()
   '
   Label13.Caption = "A"
   VARIDOC% = 1
   RNIDOC% = 0
   '
   If CATIVAPRO% = 2 Or CATIVAPRO% = 3 Or CATIVAPRO% = 4 Or CATIVAPRO% = 6 Then
      Label13.Caption = "C"
      Exit Sub
   End If
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
        '
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
   Static MONUMERA$
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
   If MONUMERA$ = "" Then
      MONUMERA$ = "AUTO"
      If TRIM$(CONTROL$("", "NUMFACTU")) = "MANUAL" Then
         MONUMERA$ = "MANUAL"
      End If
   End If
   '
   OPMOVIX$ = "FC"
   If TIFADOC% = 3 Then OPMOVIX$ = "ND"
   If TIFADOC% = 6 Then OPMOVIX$ = "NC"
   VAMOVIX$ = Label13
   '
   DATEDOC% = FECHANUM(Text9.TEXT)
   If DATEDOC% < FLINUMERA% Then
      Call MENSERR(24, "Imposible Facturar en la Fecha Elegida.\  \La Fecha de Factura Debe Ser POSTERIOR\a la Fecha de Nueva Numeración.")
      Call CIERRARCH("*.*")
      Call FINAL("")
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
   
   FLIMI = FLIDIAS: MODORAPIDO% = 1
   If FLINUME > FLIMI Then
      FLIMI = FLINUME
      MODORAPIDO% = 0
   End If
   '
22 Call APERBASDAT("CABAN")
   ASELE$ = "SELECT NumeComp FROM CajaBanc "
   '
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
   'ASELE$ = ASELE$ + "AND (CDBL(FechEmisi) >= " & FLINUME & ") "
   ASELE$ = ASELE$ + "AND FechEmisi >= " & FLIMI & " "
   ASELE$ = ASELE$ & "ORDER BY NumeComp DESC;"
   '
   Set NUFAXI = CueCorri.OpenRecordset(ASELE$, dbOpenDynaset, dbReadOnly)
   '
   NUMAXI& = 0
   On Error Resume Next
   NUFAXI.MoveFirst
   If Err < 1 Then
       On Error GoTo 0
       NUMAXI& = NUFAXI!NumeComp
     Else
       On Error GoTo 0
       If MODORAPIDO% = 1 Then
          FLIMI = FLINUME
          MODORAPIDO% = 0
          GoTo 22
       End If
   End If
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
      If TRIM$(Text10.TEXT) = "" Then
        If FORCARGA$ <> "SINREMITO" Then Call NUREMIFA(NROREMIT&)
        If NROREMIT& > XVALO(Text10.TEXT) Then
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
Function VERINUFA%(NFCY$)
   '
   Static FLINUMERA%
   Static NUMECORR%
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
   Set NUFAXI = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset, dbReadOnly)
   '
   On Error Resume Next
   NUFAXI.MoveFirst
   If Err Then
       VERINUFA% = 1
     Else
       Do While Not NUFAXI.EOF
         If CDbl(NUFAXI.FechEmisi) >= FLIMI Then
            VERINUFA% = (-1)
            GoTo 90
         End If
       NUFAXI.MoveNext
       Loop
       VERINUFA% = 1
   End If
90 On Error GoTo 0
   '
End Function
'
Function VERINURE%()
   '
   Static PREFIREM$, FLIM%, CMO$
   '
   If CMO$ = "" Then
     CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
     If CMO$ = "" Then CMO$ = "RT"
   End If
   '
   If PREFIREM$ = "" Then
     PREFIREM$ = CONTROL$("", "PREFIREM")
     If PREFIREM$ = "" Then PREFIREM$ = "0001"
   End If
   '
   If FLIM% < 1 Then
     FENUE$ = CONTROL$("", "FENUREFA")
     If FENUE$ = "" Then FENUE$ = "18/04/03"
     FLIM% = FECHANUM(FENUE$)
     If FLIM% > HOY(HO$) Then FLIM% = HOY(HO$)
   End If
   FLINUME = CDbl(CVDAT(FLIM%))
   '
   REMICOR$ = "RT." + TRIM$(Text10)
   REMILAR$ = Text10: While Len(REMILAR$) < 8: REMILAR$ = "0" + REMILAR$: Wend
   REMILAR$ = CMO$ + "-X-" + PREFIREM$ + "-" + REMILAR$
   '
   Call APERBASDAT("STOCKS")
   SQLX$ = "SELECT * FROM MOVISTO "
   SQLX$ = SQLX$ + " WHERE FechMov >= " & FLINUME & " "
   SQLX$ = SQLX$ + " AND CodiMovi = '" & CMO$ & "' "
   SQLX$ = SQLX$ + " AND DoPriLar = '" & REMILAR$ & "' "
   '
   Set MOVITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
   With MOVITMP
      On Error Resume Next
      .MoveFirst
      If Err Then
          VERINURE% = 1
        Else
          VERINURE% = (-1)
      End If
   End With
   On Error GoTo 0
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

Private Sub Text6_GOTFOCUS()
   On Error Resume Next
   'Text1.SetFocus
   Command6.SetFocus
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
   If BARRELAB.Caption = "BARRE" Then Exit Sub
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
   Call CALTOFAC
   '
End Sub

Sub PRIFACTURA()
    '
    Static AGREDESP%
    '
    If CONTROFIS% = 1 Then
       'imprime por controlador fiscal
       Call PRICONFIS
       Exit Sub
    End If
    '
    TICAMBX# = XVALO(TICAMBIS)
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
    SQLX$ = SQLX$ + " AND ((ISNULL(DoSecLar) = true) or (DoSecLar = '') OR (DoSecLar LIKE 'P/F*')) "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) >= " & DESELEC# & " "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) <= " & HASELEC# & " "
    Set CONSITEMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    CAHOJAS% = 0
    FORIPRE$ = FORELFAC$            ' BUSCA POR LA FUNCION ASOCIADA AL TIPO DE COMPROBANTE
    If FORIPRE$ = "NULL" Then Exit Sub
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
'    If InStr(UCase$(TIPODOC$), "FACTURA") > 0 Then '\\\OT-06-0110-WAR-29/03/06///
'      If InStr(UCase$(TIPODOC$), "CONTADO") < 1 Then
'        If CPAG% < 1 Or InStr(UCase$(DECONPA$(CPAG%)), "CONTADO") > 0 Then
'          TIPODOC$ = TIPODOC$ + " CONTADO"
'        End If
'      End If
'    End If
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
    CODTABU1$(27) = "N-AUTENT"
    CODTABU1$(28) = "R-PLANO"
    CODTABU1$(29) = "F-PLANO"
    '
    CACOLTAB1% = 29
    '
    CAITAB1% = 0
    CAITCRU% = 0
    '
    SUMATOTA# = XVALO(Total(0).TEXT)
    IDESCUEN# = XVALO(Total(1).TEXT)
    '
    COPIASREM% = XVALO(FINFACTURA.Text2(1))
    If COPIASREM% < 1 Then Text10.TEXT = ""
    '
    NroRemi$ = TRIM$(Text10.TEXT)
    If NroRemi$ <> "" Then
      While Len(NroRemi$) < 8: NroRemi$ = "0" + NroRemi$: Wend
      PREFIREM$ = CONTROL$("", "PREFIREM")
      If PREFIREM$ = "" Then PREFIREM$ = "0001"
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
    CODPARAM$(31) = "CTAC-CLI": DOCPARAM$(31) = CONTACTCLI$(DESTIDOC%) '  CONTACTO
    '\\\OT-5-579-FG-FIN
    '\\\OT-06-0125-WAR-03/04/06///
    CODPARAM$(33) = "NUBULTOS": DOCPARAM$(33) = TRIM$(FINFACTURA.Text2(2))          ' BULTOS
    CODPARAM$(34) = "PES-CONJ": DOCPARAM$(34) = TRIM$(FINFACTURA.Text2(11)) + " Kg" ' PESO NETO
    CODPARAM$(35) = "PES-COMP": DOCPARAM$(35) = TRIM$(FINFACTURA.Text2(4)) + " Kg"  ' PESO BRUTO
    CODPARAM$(36) = "BANC-CHE": DOCPARAM$(36) = TRIM$(FINFACTURA.Text2(5))          'BANCO EXTERIOR
    CODPARAM$(37) = "DOCU-APL": DOCPARAM$(37) = TRIM$(FINFACTURA.Text2(8))          ' BANCO LOCAL
    CODPARAM$(38) = "NUME-CHE": DOCPARAM$(38) = TRIM$(FINFACTURA.Text2(10))         ' CREDITO
    CODPARAM$(39) = "DEN-CUEN": DOCPARAM$(39) = TRIM$(FINFACTURA.Text2(9))          ' GUIA DE IMPORTACION
    CODPARAM$(40) = "REF-ASIE": DOCPARAM$(40) = TRIM$(FINFACTURA.Text2(6))          ' PUERTO DE EMBARQUE
    CODPARAM$(41) = "DET-PASE": DOCPARAM$(41) = TRIM$(FINFACTURA.Text2(7))          ' DESTINO
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
    CODPARAM$(49) = "IMP-DEBE": DOCPARAM$(49) = FINFACTURA.Text2(13)  ' FLETE
    CODPARAM$(50) = "IMP-HABE": DOCPARAM$(50) = FINFACTURA.Text2(12)  ' SEGURO
    CODPARAM$(51) = "TOT-HABE": DOCPARAM$(51) = FORMATNUM(TOTATOTA#, "F13.2")
      TASAX = ALIPERBRU#(NC1%, MINIPER#, VARIDOC%)
    CODPARAM$(52) = "IMPO-CHE": DOCPARAM$(52) = FORMATNUM(TASAX, "F5.1")
    '
    CAPARDOC% = 52 '32 '\\\OT-06-0125-WAR-03/04/06///
    '
'MsgBox "PRIFACTURA-3    "
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
    POVERTI% = XVALO(TRIM$(UCase$(ATRIFORM$(CODFOR$, "OBSERVA/POSINVER"))))
    '
    PRINTERPORT$ = PORTUFAC$
    CACOPDOC% = XVALO(FINFACTURA.Text2(0).TEXT)
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
'MsgBox "PRIFACTURA-4    "
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
      TETABU1$(1, CAITAB1%) = CODDX$
      TETABU1$(2, CAITAB1%) = Mid$(XXX$, 7, 64)
      TETABU1$(3, CAITAB1%) = CAPEX$
      'TETABU1$(4, CAITAB1%) = PRELI$
      'TETABU1$(5, CAITAB1%) = DESCX$
      'TETABU1$(6, CAITAB1%) = PREUX$
      'TETABU1$(7, CAITAB1%) = ITOTX$
      'TETABU1$(8, CAITAB1%) = FENTX$
      TETABU1$(9, CAITAB1%) = Str$(U&)
      TETABU1$(10, CAITAB1%) = Mid$(XXX$, 71, 4)    ' unidad de medida
      TETABU1$(11, CAITAB1%) = Mid$(XXX$, 75, 64)   ' referencia 1
      TETABU1$(12, CAITAB1%) = Mid$(XXX$, 139, 64)  ' referencia 1
      TETABU1$(13, CAITAB1%) = Mid$(XXX$, 203, 64)  ' referencia 1
      TETABU1$(14, CAITAB1%) = Mid$(XXX$, 315, 24)  ' numero de plano
      TETABU1$(15, CAITAB1%) = Mid$(XXX$, 339, 16)  ' numero original
      TETABU1$(16, CAITAB1%) = Mid$(XXX$, 355, 16)  ' numero de o.compra
      TETABU1$(17, CAITAB1%) = Mid$(XXX$, 387, 14)  ' numero de remito
      TETABU1$(18, CAITAB1%) = Mid$(XXX$, 403, 16)  ' numero de doc-recepcion
      TETABU1$(19, CAITAB1%) = FECHATEX$(CVI(Mid$(XXX$, 419, 2))) ' fecha de recepcion
      '
      FACIVITEM = 1
      If Label13.Caption = "B" Then FACIVITEM = 1 + (CVS(Mid$(XXX$, 441, 4))) / 100
      LPRE% = XVALO(TEXT21(0).TEXT)
      If LPRE% > 0 Then
        TETABU1$(20, CAITAB1%) = Str$(FACIVITEM * CVS(Mid$(XXX$, 421, 4))) ' precio de lista
        TETABU1$(21, CAITAB1%) = Str$(CVS(Mid$(XXX$, 425, 4))) ' PORCENTAJE DE DESCUENTO
      End If
      '
      TETABU1$(22, CAITAB1%) = Str$(FACIVITEM * CVS(Mid$(XXX$, 429, 4))) ' precio unitario
      TETABU1$(23, CAITAB1%) = Str$(CVD(Mid$(XXX$, 433, 8))) ' total neto
      TETABU1$(23, CAITAB1%) = TRIM$(FORMATNUM$(CVD(Mid$(XXX$, 433, 8)), "F15.4")) ' total neto
         If TETABU1$(23, CAITAB1%) = "" Then TETABU1$(23, CAITAB1%) = "S/Cargo"
      TETABU1$(24, CAITAB1%) = Str$(CVS(Mid$(XXX$, 441, 4))) ' alicuota iva
      TETABU1$(25, CAITAB1%) = Str$(CVD(Mid$(XXX$, 445, 8))) ' importe iva del item
      TETABU1$(26, CAITAB1%) = TRIM$(FORMATNUM$(CVD(Mid$(XXX$, 453, 8)), "F15.4")) ' total bruto
         If TETABU1$(26, CAITAB1%) = "" Then TETABU1$(26, CAITAB1%) = "S/Cargo"
      TETABU1$(27, CAITAB1%) = Mid$(XXX$, 371, 16) ' numero de pedido cliente
      '
      'DESCUENTOS DE LISTA Y COMBINADO
      DESXLXTA$ = FORMATNUM$(CVS(Mid$(XXX$, 505, 4)), "F4.1")
      DESXGPO$ = FORMATNUM$(CVS(Mid$(XXX$, 509, 4)), "F4.1")
         FORDESGRU$ = TRIM$(UCase$(ATRIFORM$(FORIPRE$, "F-PLANO/FORMATO")))
         If UCase$(Left$(FORDESGRU$, 1)) = "A" Then
           DESXGPO$ = TRIM$(Mid$(XXX$, 487, 16))
         End If
         '
     TETABU1$(28, CAITAB1%) = DESXLXTA$      ' DESCUENTO DE LISTA
     TETABU1$(29, CAITAB1%) = DESXGPO$       ' DESCUENTO COMBINADO
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
      COMPDESCRI$ = TRIM$(Mid$(XXX$, 75, 64))
      If COMPDESCRI$ <> "" Then
         CAITAB1% = CAITAB1% + 1
         TETABU1$(2, CAITAB1%) = COMPDESCRI$
      End If
      '
      ' AQUI BUSCA EL NUMERO DE SERIE
      If CIPRE% > 0 Then
'        call APERbasdat("STOCKS")
'        DOPRIC$ = TRIM$(Mid$(XXX$, 387, 14))
'        NORIT% = CVI(Mid$(XXX$, 503, 2))
'        '
'        SQLX$ = "DoPriCor = '" & DOPRIC$ & "' "
'        SQLX$ = SQLX$ + "AND Cod_Inte = " & CIPRE% & " "
'        SQLX$ = SQLX$ + "AND Nume_Clie = " & DESTIDOC% & " "
'        SQLX$ = SQLX$ + "AND NuorItem = " & NORIT% & " "
'        '
'        With Movisto
'          .FindFirst SQLX$
'25        If .NoMatch = False Then
'            SERNU$ = !ReferCor
'            If Left$(SERNU$, 5) = "S.No." Then
'              CAITAB1% = CAITAB1% + 1
'              TETABU1$(2, CAITAB1%) = SERNU$
'            End If
'          End If
'        End With
        DOPRIC$ = TRIM$(Mid$(XXX$, 387, 14))
        'NORIT% = CVI(Mid$(XXX$, 503, 2))
        NORIT% = CVI(Mid$(XXX$, 401, 2))   ' resolver enredo entre pedido y remito
        '
        SQLX$ = "DoPriCor = '" & DOPRIC$ & "' "
        SQLX$ = SQLX$ + "AND Cod_Inte = " & CIPRE% & " "
        SQLX$ = SQLX$ + "AND NuorItem = " & NORIT% & " "
        '
        With CONSITEMP
          .FindFirst SQLX$
25        If .NoMatch = False Then
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
      '
      SUBTOTA# = SUBTOTA# + XVALO(TETABU1$(23, CAITAB1%))
      SUBTOTB# = SUBTOTB# + XVALO(TETABU1$(26, CAITAB1%))
      '
      If (U& Mod ITPORHOJA%) = 0 Then
        If U& < FIN& Then
          If CACOPDOC% > 0 Then
            DOCPARAM$(26) = TRIM$(CSTRING$(MKD$(SUBTOTA#), 4, 12, 2, 2))
            Call PRINTDOC("@" + FORIPRE$)
            Call CIERRARCH("*.*")
            Call BAJALDISCO
          End If
        End If
      End If

109 Next U&
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
'MsgBox "PRIFACTURA-5    "
    If POVERTI% < 1 Then
      CAITAB1% = CAITAB1% + 1
      For KKII% = 1 To CACOLTAB1%: TETABU1$(KKII%, CAITAB1%) = "": Next KKII%
      For KU& = 1 To ULTREG&("!OBSERVOF")
        TTXX$ = REGLEIDO$("!OBSERVOF", KU&)
        CAITAB1% = CAITAB1% + 1
        For KKII% = 1 To CACOLTAB1%: TETABU1$(KKII%, CAITAB1%) = "": Next KKII%
        TETABU1$(2, CAITAB1%) = RTrim$(TTXX$)
      Next KU&
      Call SETULTREG("!OBSERVOF", 0)
    End If
        
    CODPARAM$(5) = "IMP-NETO": DOCPARAM$(5) = FORMATNUM$(XVALO(Total(0).TEXT) + XVALO(IEMBALA), "F13.2")
    CODPARAM$(9) = "IMP-IVA": DOCPARAM$(9) = Total(3).TEXT
    CODPARAM$(11) = "IMP-IVA1": DOCPARAM$(11) = Total(4).TEXT
    CODPARAM$(12) = "IMP-DESC": DOCPARAM$(12) = Total(1).TEXT
    CODPARAM$(13) = "IMPNEDES": DOCPARAM$(13) = Total(2).TEXT
    CODPARAM$(14) = "IMP-TOTA": DOCPARAM$(14) = Total(6).TEXT
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
      TOTATOTA = XVALO(Total(6)) + XVALO(FINFACTURA.Text2(12)) + XVALO(FINFACTURA.Text2(13))
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
    TTTT# = XVALO(Total(6).TEXT)
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
      Call PRINTDOC("@" + FORIPRE$)
    End If
    '
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    '
'MsgBox "PRIFACTURA-6    "
    CACOPDOC% = XVALO(FINFACTURA.Text2(1).TEXT)
    If CACOPDOC% > 0 Then
      If TRIM$(Text10.TEXT) <> "" Then
        For KKI% = 0 To CACOLTAB1%
          For KKJ% = 0 To CAITAB1%
            TETABU1$(KKI%, KKJ%) = ""
          Next KKJ%
        Next KKI%
        '
        DOCUORIG$ = NUMEDOC$
        Call PRIREMIFA
      End If
    End If
    '
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    '
'MsgBox "PRIFACTURA-7    "
    If BARRELAB.Caption <> "BARRE" Then
       If NOCONFIRM% > 0 Then
          Call COMUNI("Se han Emitido los Documentos Solicitados")
       End If
    End If
    For KKI% = 0 To CACOLTAB1%
      For KKJ% = 0 To CAITAB1%
        TETABU1$(KKI%, KKJ%) = ""
      Next KKJ%
    Next KKI%
    '
End Sub

Sub PRIREMIFA()
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FOREMIFA"))
    If TRIM$(FORIPRE$) <> "" Then
       '
       '\\\OT-06-0114-WAR-31/03/06///
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
       POVERTI% = XVALO(TRIM$(UCase$(ATRIFORM$(CODFOR$, "OBSERVA/POSINVER"))))
       '\\\OT-06-0114-WAR-31/03/06///
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
       TIPODOC$ = "Remito"
       '
       If TRIM$(Text14.TEXT) <> "" Then
         NUOCOMPRA$ = TRIM$(Text14.TEXT)
       End If
       '
       CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = TRIM$(Mid$(TRXP$, 1, 32))
       CODPARAM$(2) = "NUBULTOS": DOCPARAM$(2) = TRIM$(FINFACTURA.Text2(2).TEXT)
       CODPARAM$(3) = "CUIT-TRA": DOCPARAM$(3) = TRIM$(Mid$(TRXP$, 65, 16))
       CODPARAM$(4) = "VALO-DEC": DOCPARAM$(4) = TRIM$(FINFACTURA.Text2(3).TEXT)
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
       '\\\OT-08-0866-19/09/08
       'PARAMETROS PARA IMPRESION DE DATOS DEL TRANSPORTE SEGUN LO
       'ELEGIDO EN EL FORMULARIO PRINCIPAL DE FACTURA
         NTRA% = XVALO(Text19)
       CODPARAM$(26) = "NOP-OPER": DOCPARAM$(26) = DENOTRANS(NTRA%)
       CODPARAM$(27) = "DES-OPER": DOCPARAM$(27) = DOMITRANS(NTRA%)
       CODPARAM$(28) = "OBS-OPER": DOCPARAM$(28) = LOCATRANS(NTRA%)
       '\\\OT-08-0866-FIN
       '
       CAPARDOC% = 28
       '
       CODTABU1$(1) = "COD-MAT"
       CODTABU1$(2) = "DES-MAT"
       CODTABU1$(3) = "CANTIDAD"
       CODTABU1$(4) = "ORD-ITEM"
       CODTABU1$(5) = "N-AUTENT"
       CODTABU1$(6) = "N-OCOMPR"
       CODTABU1$(7) = "N-PLANO"
       CACOLTAB1% = 7
       '
       DESTIDOC% = NC%
       CAITAB1% = 0
       '
       If FINFACTURA.Text3 <> "" Then
         Call FRATEXTO(FINFACTURA.Text3, 40)
         For KKII% = 1 To CARETEX%
           CAITAB1% = CAITAB1% + 1
           TETABU1$(1, CAITAB1%) = ""
           TETABU1$(2, CAITAB1%) = RETEX$(KKII%)
           TETABU1$(3, CAITAB1%) = ""
           TETABU1$(4, CAITAB1%) = Str$(CAITAB1%)
         Next KKII%
         FINFACTURA.Text3 = ""
         FINFACTURA.Frame5.Visible = False
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
           If TRIM$(Mid$(XXX$, 387, 14)) = "" Then
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
             TETABU1$(5, CAITAB1%) = Mid$(XXX$, 371, 16) 'numero de pedido cliente
             TETABU1$(6, CAITAB1%) = Mid$(XXX$, 355, 16)
             TETABU1$(7, CAITAB1%) = TRIM$(Mid$(XXX$, 463, 24))
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
             COMPDESCRI$ = TRIM$(Mid$(XXX$, 75, 64))
             If COMPDESCRI$ <> "" Then
               CAITAB1% = CAITAB1% + 1
               TETABU1$(2, CAITAB1%) = COMPDESCRI$
             End If
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
80     CACOPDOC% = XVALO(FINFACTURA.Text2(1).TEXT)
       PRINTERPORT$ = PORTUREM$
       Call PRINTDOC("FOREMIFA")   ' Remito de Facturacion
       '
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
           LMCR# = LIMICRE(NCLIE%)
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

Sub GRAFACTURA(OKX%)
   '
'MsgBox "Grabando en Factura_Nue"
Call GRAFACTURA_NUE(OKX%)
Exit Sub
   Dim NIVA7#(10)
   '
   OKX% = 0
   NCLIE% = XVALO(Text1.TEXT)
   If MARCONOTA.Visible = True Then Exit Sub
   If ULTREG&("!CARDEFAC") < 1 Then Exit Sub
   '
   Call APROCOPIAS(OKX%)
   COPIASFAC% = XVALO(FINFACTURA.Text2(0))
   COPIASREM% = XVALO(FINFACTURA.Text2(1))
   If OKX% < 0 Then Exit Sub
   If COPIASFAC% + COPIASREM% < 1 Then Exit Sub
   '
   If COPIASFAC% > 0 Then
     If XVALO(Total(6).TEXT) < 0.01 Then
       Call MENSERR(24, "Imposible Emitir Factura SIN CARGO !!!")
       Exit Sub
     End If
   End If
   '
   If COPIASFAC% < 1 Then
      Call GRAREMIFA
      Call PRIFACTURA
      Screen.MousePointer = 1
      Exit Sub
   End If
   '
   PPVI% = PIVACLI%(NCLIE%)
   If PPVI% < 0 Or PPVI% > 7 Then PPVI% = 1
   CUEVE$ = CONTROL$("SVD451S", "CUEVE-" + TRIM$(Str$(PPVI%)))
   QQQ% = CUECOINT%(CUEVE$)
   '
   If TIFADOC% = 3 Or TIFADOC% = 6 Then
'      OPIMPUTA.Option1.Value = True
      Call OPIMPUTA.Option1_Click
      '
      If TIFADOC% = 6 Then
        'Option1(5).Value = True
        If CONTROL$("", "IMPUNCRE") = "DIFERENCIACAMBIO" Then
'          OPIMPUTA.Option2.Value = True
'          Call OPIMPUTA.Option1_Click
        ElseIf CONTROL$("", "IMPUNCRE") = "VARIOS" Then
'           OPIMPUTA.Option3.Value = True
'           Call OPIMPUTA.Option1_Click
        End If
      End If
      
      If TRIM$(DOCUDIF) <> "" Then
'        OPIMPUTA.Option2.Value = True
      End If
      '
      OPIMPUTA.Show 1
   End If
   '
'   'NUEVO PARA INGRESAR UN NUMERO DE FACTURA MANUAL
   '
   NUPANTALLA& = XVALO(Text11.TEXT)
   If NFCAM% < 1 Then
       Call NUFACTURA
       If XVALO(Text11) <> NUPANTALLA& Then GoTo 16
     Else
       TICOMX$ = "FC"
       If TIPO% = 3 Then TICOMX$ = "ND"
       If TIPO% = 6 Then TICOMX$ = "NC"
       NFCZ$ = TRIM$(Text11): While Len(NFCZ$) < 8: NFCZ$ = "0" + NFCZ$: Wend
       NFCY$ = TICOMX$ + "-" + Label12 + "-" + NFCZ$ + "-" + Label13
       '
       If VERINUFA%(NFCY$) < 0 Then
16       Call MENSERR(24, "Número de Comprobante no Válido")
         Exit Sub
       End If
   End If
   '
21 If XVALO(Text11.TEXT) > NUPANTALLA& Then
      Call CIERRARCH("SVD202")
      Call CIERRARCH("DEUDOR2")
      Call CIERRARCH("FACTUR")
      Call CIERRARCH("CAJABANC")
      Call CIERRARCH("MOVISTO")
      Call CIERRARCH("SALSTOCK")
      Call COMUNI("FLEXOFT ha Re-Numerado el Documento\como Número '" + TRIM$(Text11.TEXT) + "'.\  \Verifique y Vuelva a Grabar.")
      Exit Sub
   End If
   Text11.TEXT = TRIM$(Str$(NUPANTALLA&))
   '
   COPIASREM% = XVALO(FINFACTURA.Text2(1))
   If COPIASREM% < 1 Then
      Text10.TEXT = ""
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
   '
   If VERLICRE% = 1 Then
     If TIFADOC% < 4 Then ' EXCLUYE CREDITOS
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
       GoTo 25
       '
22     LMCR# = LIMICRE(NCLIE%)
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
25 DESC$ = TRIM$(PORDESCU.TEXT)
   PODES = COEFDESCU(DESC$)
   '
   For KK% = 1 To 10: TOTAFAC#(1, KK%) = 0: Next KK%
   TOTAFAC#(1, 1) = XVALO(BASIVA(0)) / (1 - PODES / 100)
   TOTAFAC#(1, 2) = XVALO(BASIVA(1)) / (1 - PODES / 100)
   TOTAFAC#(1, 3) = XVALO(Total(0).TEXT)
   TOTAFAC#(1, 4) = XVALO(Total(1).TEXT)
   TOTAFAC#(1, 5) = XVALO(BASIVA(0))
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
   LU$ = LUDAT$
   ARCHIDEU$ = "SVD202"
   ARCHICLI$ = "DEUDOR"
   ARCHIFAC$ = "FACTUR"
   RD% = 1
   TIPO% = TIFADOC%
   VA% = VARIDOC%
   NRO = XVALO(Text11.TEXT)
   NC% = NCLIE%
   FFI% = FECHANUM(Text9.TEXT)
   CPAG% = XVALO(TEXT21(1).TEXT)
   '
   TOTAGEN# = 0
   NENETO# = XVALO(Total(2)) - XVALO(IEMBALA)
   NEFA7# = ROUND#(NENETO# * TICAMBX#, 2)
   TOTAGEN# = TOTAGEN# + NEFA7#
   NEMBAL7# = ROUND#(XVALO(IEMBALA) * TICAMBX#, 2)
   TOTAGEN# = TOTAGEN# + NEMBAL7#
   For KKI% = 0 To 2
     NIVA7#(KKI%) = ROUND#(XVALO(Total(KKI% + 3)) * TICAMBX#, 2)
     TOTAGEN# = TOTAGEN# + NIVA7#(KKI%)
   Next KKI%
   PERIB7# = ROUND#(XVALO(Text15) * TICAMBX#, 2)
   TOTAGEN# = TOTAGEN# + PERIB7#
   PERIVA7# = ROUND#(XVALO(Text17) * TICAMBX#, 2)
   TOTAGEN# = TOTAGEN# + PERIVA7#
   ' TOTAGEN# = XVALO(Total(6).TEXT) * TICAMBX#
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
   ' LO QUE SIGUE ES PARA CONSIDERAR LAS HOJAS QUE SE CONSUMIRAN
   'FORIPRE$ = TRIM$(CONTROL$("", "FORFAC-" + Label13.Caption))
   FORIPRE$ = FORELFAC$
   CARENGLO& = ULTREG&("!CARDEFAC")
   ITPORHOJA% = XVALO(ATRIFORM$(FORIPRE$, "MAX-ITEM"))
   If ITPORHOJA% < 1 Then
     ITPORHOJA% = 80
   End If
   CAHOJNEC% = Int(0.99 + CARENGLO& / ITPORHOJA%)
   '
   If NFCAM% < 1 Then
     Call NUFACTURA
   End If
   '
27 NPP& = XVALO(Text11) + CAHOJNEC% - 1
   Text11.TEXT = TRIM$(Str$(NPP&))
   Text11.Refresh
   '
   TICOMX$ = "FC": OPMOVIX$ = "FAC"
   If TIPO% = 3 Then TICOMX$ = "ND": OPMOVIX$ = "DEB"
   If TIPO% = 6 Then TICOMX$ = "NC": OPMOVIX$ = "CRE"
   NFCZ$ = TRIM$(Text11): While Len(NFCZ$) < 8: NFCZ$ = "0" + NFCZ$: Wend
   NFCX$ = TICOMX$ + "-" + Mid$(NFCZ$, 3) + "-" + Label13
   NFCY$ = TICOMX$ + "-" + Label12 + "-" + NFCZ$ + "-" + Label13
   '
   OPMOVIX$ = "FC"
   If TIFADOC% = 3 Then OPMOVIX$ = "ND"
   If TIFADOC% = 6 Then OPMOVIX$ = "NC"
   VAMOVIX$ = Label13
   '
   SIGX% = 1: If TIPO% = 6 Then SIGX% = (-1)
   '
   'NUEVO PARA INGRESO MANUAL DE NUM FAC
   If NFCAM% = 1 Then
      If VERINUFA%(NFCY$) < 0 Then
        OKX% = (-1)
        Exit Sub
      End If
    End If
   'HASTA ACA NUERO PARA INGRESO MANUAL DE NUM FAC
   '
' AQUI GRABAR CAJABANC PARA FACTURA
   Call APERBASDAT("CABAN")
   NOCLOSE% = 1
30 CoComCor$ = NFCX$
   CoComLar$ = NFCY$
   CodiEmp% = 0
   CodiSuc% = 0
   NumeClie% = GRUDIVCLI%(NC%)
   SucuClie% = XVALO(Text20)
   NumeProv% = 0
   FechEmi = CVDAT(FFI%)
   Referen$ = RASOCLI$(NC%)
   FeContab = CVDAT(FFI%)
   FeAcredi = CVDAT(FFI%)
   CuentaCo% = CUECOCLI%(NC%)
   TiMovi$ = "FVEN"
   OpMovi$ = OPMOVIX$
   VaMovi$ = Label13
   Descrip$ = Label7 + " " + Label13.Caption + " " + Label6.Caption + " " + Label12.Caption + "-" + Text11.TEXT
   NumeTal = XVALO(Label12)
   NumeCom = CDbl(Text11)
   StatInit = 0
   ImpoDeb = TOTAGEN# * SIGX%
   Observa$ = RASOCLI$(NC%)
   Call GRACABAN
   '
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
     Call GRACABAN
   End If
   '
   ' GRABA IMPORTE EMBALAJES
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
     Call GRACABAN
   End If
     '
   'graba iva inscripto
   For KKI% = 0 To 1
     TOTANETX# = NIVA7#(KKI%)
     If TOTANETX# >= 0.005 Then
       CUEIVA$ = CONTROL$("SVD451S", "CUEIVA")
       CuentaCo% = CUECOINT%(CUEIVA$)
       TiMovi$ = "IVA"
       OpMovi$ = "DeFi"
       VaMovi$ = PORIVA(KKI%) + " %"
       Descrip$ = "Iva Debito Fiscal " + NFCX$
       StatInit = 0
       ImpoDeb = TOTANETX# * SIGX% * (-1)
       BasImpo# = ROUND#(XVALO(BASIVA(KKI%)) * TICAMBX#, 2)
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
     BasImpo# = ROUND#(XVALO(BASIVA(2)) * TICAMBX#, 2)
     Call GRACABAN
   End If
   BasImpo# = 0
   '
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
     BasImpo# = ROUND#(XVALO(Total(2)) * TICAMBX#, 2)
     Call GRACABAN
   End If
   BasImpo# = 0
   '
   'Call BLIMESS("Calculando Vencimientos ...")
   If TIFADOC% <> 6 Or TRIM$(DOCUDIF) = "" Then
     Call GRADEUFA(TIPO%, VA%, NRO, NC%, FFI%, CPAG%, TOTAGEN#)
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
     If TIFADOC% < 3 Then   ' graba noviniento cabeza
       Call GRAMOCABE
     End If
   End If
   '
80 OKX% = 1
   Screen.MousePointer = 1
   Exit Sub

98 Screen.MousePointer = 1
   Call MENSERR(24, MERRO$)
   '
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
   TIPO% = TIFADOC%
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
     With SaDeuPen
       .FindFirst "NuClien = " & NCI% & " AND CODICOM_LAR = '" & DOCUCA$ & "' AND SALDdebe > 0"
       If .NoMatch = False Then
850      PARCAN# = ICAN#: If !SaldDebe < ICAN# Then PARCAN# = !SaldDebe
         .Edit
         !SaldDebe = !SaldDebe - PARCAN#
         .Update
         '
         DocOrCor$ = !CodiCom_Cor
         DocOrLar$ = !CodiCom_Lar
         TotOrig# = !IBruComp
         IvaOrig# = 0
         On Error Resume Next
         IvaOrig# = !IIvaComp
         On Error GoTo 0
         '
         ICAN# = ICAN# - PARCAN#
         If ICAN# >= 0.005 Then
           .FindNext "CODICOM_LAR = '" & DOCUCA$ & "' AND SALDdebe > 0"
           GoTo 850
         End If
       End If
     End With
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
Sub GRADEUFA(TIPO%, VA%, NRO, NC%, FFAC%, COPAGI%, TOTFAC#)
   '
   Dim FVI%(6), PLAII%(6), IM#(6), FEX(3)
   '
   LU$ = LUDAT$
   IORIG# = TOTFAC#
   '
   If TIPO% = 0 Then
     If IORIG# < 0 Then
       TIPO% = 7
     End If
   End If
   '
   If TIPO% > 3 Then
     NUPAG% = 1
     FVI%(1) = FFAC%
     IM#(1) = TOTFAC#     ' IORIG#
     GoTo 41
   End If
   '
   ' busca condicion de pago segun COPAGI%
20 Call BUCONPA(LU$, COPAGI%, COPAG$, FREF%, PLA$)
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
   SIGX% = 1: If TIPO% > 3 Then SIGX% = (-1)
   For I% = 1 To NUPAG%
     If Abs(IM#(I%)) >= 0.005 Then
       FechEmi = CVDAT(FFAC%)
       IBruTot# = IORIG#
       NuorVen% = NuorVen% + 1
       FeVenc = CVDAT(FVI%(I%))
       ImVenc# = (Abs(IM#(I%))) * SIGX%
       CondiPago% = COPAGI%
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
     FINFACTURA.Text2(KKI%) = ""
   Next KKI%
   '
   TICOMPROBAN$ = TCMP$
   FINFACTURA.Label2.Caption = TCMP$
   FINFACTURA.Label1(0) = TCMP$ + ":"
   FINFACTURA.Label4.Caption = CATIVAS$
   '
   CODFOR$ = TRIM$(CONTROL$("", "FORFAC-" + Label13.Caption))
   NROCON& = XVALO(Text11.TEXT)
   NCN$ = TRIM$(Str$(NROCON&))
   While Len(NCN$) < 8: NCN$ = "0" + NCN$: Wend
   PREFI$ = Label12.Caption
   NCN$ = PREFI$ + "-" + NCN$
   NUCOMPROBAN$ = NCN$
   FINFACTURA.Label3.Caption = NCN$
   '
5  COPIASFAC% = XVALO(CONTROL$("", "COPIAFAC"))
   COPIASREM% = XVALO(CONTROL$("", "COPIAREM"))
   If COPIASFAC% < 1 Then COPIASFAC% = 1
   If COPIASREM% < 1 Then COPIASREM% = 1
   '
   If TIFADOC% > 2 Then COPIASREM% = 0
   If CONSICLI% = 1 Then COPIASREM% = 0
   If CONTROL$("FACTUMAI", "FORMACAR") = "SINREMITO" Then COPIASREM% = 0
   '
   FINFACTURA.Text2(0) = TRIM$(Str$(COPIASFAC%))
   FINFACTURA.Text2(1) = TRIM$(Str$(COPIASREM%))
   FINFACTURA.Text2(2) = " "
   'FINFACTURA.Text2(3) = TRIM$(Str$(Int(TICAMBX# * XVALO(Total(2).TEXT)))) + ".="
   '\\\OT-05-0492-WAR-11/07/05///
     COVADE = XVALO(CONTROL("", "COEVADEC")) / 100 'COEFICIENTE DE VALOR AGREGADO
     If COVADE < 0.01 Then COVADE = 1
     FINFACTURA.Text2(3) = TRIM$(Str$(Int(COVADE * TICAMBX# * XVALO(Total(2).TEXT)))) + ".="
   '\\\OT-05-0492-WAR-FIN
   FINFACTURA.Label5.Caption = ""
   For U& = 0 To 8
      FINFACTURA.Text1(U&).TEXT = ""
   Next U&
   '\\\OT-06-0114-WAR-03/04/06///
   FINFACTURA.Text4.TEXT = ""
   FINFACTURA.Text5.TEXT = ""
   'LIMPIO LOS TEXT DE DATOS DE LA EXPO
   For JX% = 4 To 11
    FINFACTURA.Text2(JX%) = ""
   Next JX%
   '\\\OT-06-0114-WAR-FIN///
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   TTXX$ = ""
   '
   FINFACTURA.Label6.Enabled = False
   FINFACTURA.CLAREDOL = "0"
   If MONEMI% = 1 Then
     FINFACTURA.Label6.Enabled = True
   End If
   '
   FINFACTURA.Label6.Visible = False
   FINFACTURA.Text3 = ""
   FINFACTURA.Frame5.Visible = False
   '\\\OT-06-0125-WAR-03/04/06///
   FINFACTURA.Height = 6690
   'SI FACTURA DE EXPO => SE MUESTRA DATOS DE EXPORTACIÓN
   NC1% = XVALO(FORFACCS.Text1.TEXT)
   If PIVACLI%(NC1%) = 5 Or PIVACLI%(NC1%) = 7 Then
     FINFACTURA.Height = FINFACTURA.Height + 1935 '(8600)
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
   FINFACTURA.Text2(11) = FORMATNUM$(PT, "F8.1")
   '\\\OT-06-0125-WAR-FIN///
   '18/04/07 (OT-07-0162) 'PARA VER LA COTIZACION DEL DOLAR CUANDO EMITIMOS UNA FACTURA EN PESOS
   If MONEMI% <= 1 Then
      NLIPRE% = XVALO(FORFACCS.TEXT21(0))
      If MONELISTA%(NLIPRE%) = 2 Then
         TTXX$ = "Emitido en " + DEMONECO$(MONEMI%) + " a cotización " + DEMONECO$(2) + " de " + TRIM$(FORMATNUM$(TICAMONE(2), "F9.4")) + ". "
          FINFACTURA.Text4 = TTXX$
      End If
   End If
   FLEYE% = VARIDOC% - 1
If TRIM(FINFACTURA.Text4) <> "" Then FINFACTURA.Text4 = FINFACTURA.Text4 + Chr(10) + Chr(14)
RFF$ = RECFILT$("LEYFAC", 1, MKI$(FLEYE%))
'LEE OBSERVACIONES Y LAS CARGA EN LA FACTURA CORRESPONDIENTE
For U& = 1 To Len(RFF$) Step 4
  RELE& = CVS(Mid$(RFF$, U&, 4))
  LEYY$ = REGLEIDO$("LEYFAC", RELE&)
  FINFACTURA.Text4 = TRIM(FINFACTURA.Text4) + " " + TRIM(Mid$(LEYY$, 3))
Next U&
'
'LEE OBSERVACIONES Y LAS CARGA EN EL REMITO
RFF$ = RECFILT$("LEYFAC", 1, MKI$(3))
For U& = 1 To Len(RFF$) Step 4
  RELE& = CVS(Mid$(RFF$, U&, 4))
  LEYY$ = REGLEIDO$("LEYFAC", RELE&)
  FINFACTURA.Text5 = TRIM(FINFACTURA.Text5) + " " + TRIM(Mid$(LEYY$, 3))
Next U&

   '18/04/07 (FIN)
   If BARRELAB.Caption = "BARRE" Then
      FINFACTURA.Label5.Caption = "APRO"
      GoTo 20
   End If
   '
10 If CONTROFIS% = 1 Then
        FINCONFIS.Show 1
      Else
        FINFACTURA.Show 1
   End If
   Refresh
   '
   If FINFACTURA.Label5.Caption <> "APRO" Then
      If COMALTER%("Realmente Quiere CANCELAR este Documento\\No, Continuar\Si, Cancelar") <> 2 Then GoTo 10
      OKX% = (-1)
      Exit Sub
   End If
   '
20 COPIASFAC% = XVALO(FINFACTURA.Text2(0))
   If COPIASFAC% < 0 Or COPIASFAC% > 9 Then GoTo 5
   COPIASREM% = XVALO(FINFACTURA.Text2(1))
   If COPIASREM% > 9 Then GoTo 5
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
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "DAKOT") > 0 Then
       Call CARFADAKOT
       Screen.MousePointer = 1
       Exit Sub
   End If
   '
   HOII% = HOY(HOS$)
   '
   ATRI$ = "/NC/ARCHECO(2)=" + ARCHIPRE$
   ATRI$ = ATRI$ + "/TITUPAN=Detalle de Facturación (Mat.Codificado)"
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Frame1.Top + Picture3.Top))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture3.Left - 50))
   '
   Timer1.Enabled = True
10 PRF$ = "FACONGUI"
   NOFRESHPRE% = 1
   If CONSICLI% = 0 Then
      PRF$ = "FACDOSI"
      NOFRESHPRE% = 0
      EPAC$ = TRIM$(UCase$(EMPREAC$))
      If InStr(EPAC$, "FRIOLA") > 0 Then
        PRF$ = "FACFRIOL"
      End If
      If MODESCU% = 1 Then PRF$ = "FACPORDE"
   End If
   If MODODOC% = 1 Then
      'Call SETTELIB
      PRF$ = "FAC-LIBR"
      ATRI$ = "/NC"
      ATRI$ = ATRI$ + "/TITUPAN=Detalle de Facturación (Texto Libre)"
      ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Frame1.Top + Picture3.Top))
      ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture3.Left + 50))
   End If
   List1.Clear
   '
20 LPRE% = XVALO(TEXT21(0))
   ARCHIPRE$ = TRIM$(Str$(LPRE%))
   While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
   ARCHIPRE$ = "LIPRE" + ARCHIPRE$
   XXX% = VENTABU%(PRF$ + ATRI$)
   If XXX% < 0 Then GoTo 91 ' cancelada
   '
   HUBOCA% = 0
   CLIEN% = XVALO(Text1)
   For U& = 1 To ULTREG&("!CARDEFAC")
      X$ = REGLEIDO$("!CARDEFAC", U&)
      CIXXI% = CVI(Mid$(X$, 5, 2))
      CAN = CVS(Mid$(X$, 1, 4))
      PRELI = CVS(Mid$(X$, 421, 4))
      PREUNI = CVS(Mid$(X$, 429, 4)) '
      PORDESCUIN = CVS(Mid$(X$, 425, 4))
      If PORDESCUIN < 0 Or PORDESCUIN > 100 Then PORDESCUIN = 0
      '
      NPEDID& = XVALO(TRIM$(Mid$(X$, 371, 16)))
      If NPEDID& > 0 Then  ' SI VIENE DE PEDIDO NO SE PUEDE CAMBIAR EL PRECIO
        'AQUI VALIDAR CAMBIOS DE CODIGO
        NORITPE% = CVI(Mid$(X$, 503, 2))
        SQLX$ = "NroPed = " & NPEDID& & " "
        SQLX$ = SQLX$ + "AND NuOrItem = " & NORITPE% & " "
        SQLX$ = SQLX$ + "AND Status < 8 "
        'Set PEDDETTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
        '
        PedDeta.FindFirst SQLX$
        If PedDeta.NoMatch = True Then    ' NO LO ENCONTRO - CAMBIÓ EL CÓDIGO
            Call REPLA(X$, "", 371, 16)       ' PONE EN CERO EL PEDIDO
            Call REPLA(X$, MKI$(0), 503, 2)   ' PONE EN CERO EL NÚMERO DE ORDEN DEL PEDIDO
            Call GRAREG("!CARDEFAC", X$, U&)
          ElseIf PedDeta!CodiInt <> CIXXI% Then
            Call REPLA(X$, "", 371, 16)       ' PONE EN CERO EL PEDIDO
            Call REPLA(X$, MKI$(0), 503, 2)   ' PONE EN CERO EL NÚMERO DE ORDEN DEL PEDIDO
            Call GRAREG("!CARDEFAC", X$, U&)
        End If
        '
        NPEDID& = XVALO(TRIM$(Mid$(X$, 371, 16)))
        If NPEDID& > 0 Then  ' SI VIENE DE PEDIDO NO SE PUEDE CAMBIAR EL PRECIO
          If Abs(PREUNI - PRELI * (100 - PORDESCUIN) / 100) > 0.05 Then
            Call MENSERR(24, "No Válido Cambiar Precio\Cargado en Orden de Venta.")
            PREUNI = PRELI * (100 - PORDESCUIN) / 100
            Call REPLA(X$, MKS$(PREUNI), 429, 4)
            Call GRAREG("!CARDEFAC", X$, U&)
            GoTo 20
          End If
        End If
        '
      End If
      '
      NPEDID& = XVALO(TRIM$(Mid$(X$, 371, 16)))
      If NPEDID& <= 1 Then       ' NO VINO DE PEDIDO
         DOPRIC$ = TRIM$(Mid$(X$, 387, 14))
         If DOPRIC$ <> "" Then GoTo 29 ' VINO DE REMITO PREVIO
         '
         PRECARGADO = CVS(Mid$(X$, 429, 4))
         ' PRECARGA = PRELISTA#(ARCHIPRE$, CI%)
         Call REPLA(X$, MKS$(PRECARGADO), 421, 4)
         If TIFADOC% > 2 Or MODODOC% > 0 Then GoTo 29
         '
         ' Trae el descuento de lista
         NULISTA% = XVALO(TEXT21(0))
         DESCULISTA# = ROUND(DESCULISPRE(NULISTA%), 1)
         KOEFILISTA# = 1 - DESCULISTA# / 100 ' Calcula el coeficiente de Dto. de lista
         ' Trae la combinacion del descuento que tiene el articulo/Cliente
         DESCOMBIN$ = DESCOMBINA$(CLIEN%, CIXXI%)
         DESCUCOMBI# = ROUND(PORCEDESCU(DESCOMBIN$), 1)
         KOEFICOMBI# = 1 - DESCUCOMBI# / 100 'Calcula el coeficiente
         ' Determina el descuento general
         DESCUGENER# = ROUND#(100 * (1 - KOEFILISTA# * KOEFICOMBI#), 1)
         Call REPLA(X$, DESCOMBIN$, 487, 16)
         Call REPLA(X$, MKS$(DESCULISTA#), 505, 4)
         Call REPLA(X$, MKS$(DESCUCOMBI#), 509, 4)
         Call REPLA(X$, MKS$(DESCUGENER#), 425, 4)
         PREUNITA = PRECARGADO * (1 - DESCUGENER# / 100)
         Call REPLA(X$, MKS$(PREUNITA), 429, 4)
         Call GRAREG("!CARDEFAC", X$, U&)
         HUBOCA% = 1
      End If
29 Next U&
   If HUBOCA% > 0 Then
      Call COMUNI("ATENCIÓN: Cambios o Agregados a Ordenes de Venta.")
   End If
   '
   Call CALTOFAC
   Timer1.Enabled = False
   '
  'aqui agregar la carga de colores
   '\\\OT-08-0606-OD-28/07/08///
   If CONTROL$("PEDIDO", "VARCOLOR") = "SI" Then
      If TIFADOC% <= 2 Then
        If MODODOC% < 1 Then
          XXX% = VENTABU%("CARDECOL")
          If XXX% < 0 Then
             For U& = 1 To ULTREG&("!CARDEFAC")
               X$ = REGLEIDO$("!CARDEFAC", U&)
               NPEDID& = XVALO(TRIM$(Mid$(X$, 371, 16)))
               If NPEDID& <= 1 Then       ' NO VINO DE PEDIDO
                  PRELI = CVS(Mid$(X$, 421, 4))
                  Call REPLA(X$, MKS$(PRELI), 429, 4)
                  Call GRAREG("!CARDEFAC", X$, U&)
               End If
             Next U&
             GoTo 20
          End If
        End If
      End If
   End If
   'Graba detalles de tela y color
   '
   ' ///OT-8-812-04/09/08
   If CONTROL$("PEDIDO", "VARCOLOR") = "SI" Then
      For U& = 1 To ULTREG&("!CARDEFAC")
        XX$ = REGLEIDO$("!CARDEFAC", U&)
        TTXX$ = Mid$(XX$, 75, 64)
        COLORBA$ = TRIM$(Mid$(XX$, 283, 10))
        TELA$ = TRIM$(Mid$(XX$, 293, 10))
        COLTELA$ = TRIM$(Mid$(XX$, 315, 24))
        If TRIM$(COLOBA$) + TRIM$(TELA$) + TRIM$(COLTELA$) <> "" Then
          COMPDESCRI$ = "CB: " + COLORBA$ + " - T: " + TELA$ + " - CT: " + COLTELA$
          Call REPLA(XX$, COMPDESCRI$, 75, 64)
          Call GRAREG("!CARDEFAC", XX$, U&)
        End If
      Next U&
   End If
   ' ///OT-8-812-FIN
   '\\\OT-08-0606-OD-FIN///
   '
90 NPEDID& = 0
91 If XXX% < 0 Then
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
      If XVALO(TEXT21(2).TEXT) < 1 Then
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
   NLIS% = XVALO(Mid$(ARCHIPRE$, 6, 3)): If NLIS% > 255 Then NLIS% = 0
   For U& = 1 To FIN&
      Y$ = REGLEIDO$("!CARDEFAC", U&)
      TTXX$ = Space$(128)
      CI% = CVI(Mid$(Y$, 5, 2))
      PRECIO = CVS(Mid$(Y$, 429, 4))
      PRELI = PRELISTA#(ARCHIPRE$, CI%)
      PORDESCUEN = 0: If PRELI >= 0.0001 Then PORDESCUEN = 100 * (PRELI - PRECIO) / PRELI
      MAREPA$ = UCase$(Mid$(Y$, 461, 1)): If MAREPA$ <> "R" Then MAREPA$ = " "
      NPEDID& = XVALO(Mid$(Y$, 371, 8))
      CAN = CVS(Mid$(Y$, 1, 4))
      CANDECI% = DECILISTA%(NLIS%)
      '
      Call REPLA(TTXX$, CSTRING$(MKS$(CAN), 3, 6, 0, 2), 1, 6)
      Call REPLA(TTXX$, CODEXT$(CI%), 8, 15)
      Call REPLA(TTXX$, Mid$(Y$, 7, 64), 24, 26)
      PRELI = CVS(Mid$(Y$, 421, 4))
      Call REPLA(TTXX$, CSTRING$(MKS$(PRELI), 3, 8, CANDECI%, 2), 50, 8)
      PREUNI = CVS(Mid$(Y$, 429, 4))
      Call REPLA(TTXX$, CSTRING$(MKS$(PREUNI), 3, 8, CANDECI%, 2), 64, 8)
      PORDESCUIN = CVS(Mid$(Y$, 425, 4))
      If PRELI >= 0.1 Then
         PORDESCUIN = 100 * (PRELI - PREUNI) / PRELI
         If PORDESCUIN < 0 Then PORDESCUIN = 0
         Call REPLA(Y$, MKS$(PORDESCUIN), 425, 4)
         Call GRAREG("!CARDEFAC", Y$, U&)
      End If
      Call REPLA(TTXX$, CSTRING$(MKS$(PORDESCUIN), 3, 5, 1, 2), 59, 5)
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
     If CATIVAPRO% = 2 Or CATIVAPRO% = 3 Or CATIVAPRO% = 4 Or CATIVAPRO% = 6 Then
       APLIVA% = 0
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
       NLIS% = XVALO(Mid$(ARCHIPRE$, 6, 3)): If NLIS% > 255 Then NLIS% = 0
       CANDECI% = DECILISTA%(NLIS%)
       If CAN = 1 Then
           TOT# = ROUND(CDbl(PREUNI), 2)
         Else
           TOT# = ROUND(CDbl(ROUND(CAN, 1) * ROUND(PREUNI, CANDECI%)))
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
     If PEMBA > 0 Then
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
'      If CATIVAPRO% = 2 Or CATIVAPRO% = 3 Or CATIVAPRO% = 4 Or CATIVAPRO% = 6 Then
'        TAIVA(U& - 1) = 0
'        PORIVA(U& - 1) = "0.00"
'      End If
   Next U&
   PORIVA(2) = "50.00" ' responsable no inscripto
'   If CATIVAPRO% = 2 Or CATIVAPRO% = 3 Or CATIVAPRO% = 4 Or CATIVAPRO% = 6 Then
'      PORIVA(2) = "0.00"
'   End If
   Call CIERRARCH("GRUPIVA")
   ALIIVA = ALIVA
   '
End Sub

Sub GRACONSI()
   '
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
   SQLX$ = " SELECT * FROM MOVISTO "
   SQLX$ = SQLX$ + "WHERE Nume_Clie = " & NC1% & " "
   SQLX$ = SQLX$ + " AND (CodiMovi = '" & CMO$ & "' "
   SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
   SQLX$ = SQLX$ + " AND DoPriLar <> '' "
   SQLX$ = SQLX$ + " AND ((ISNULL(DoSecLar) = true) or (DoSecLar = '') OR (DoSecLar LIKE 'P/F*')) "
   SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) >= " & DESELEC# & " "
   SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) <= " & HASELEC# & " "
   '
   Set CONSITMP = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
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
       DOPRIC$ = TRIM$(Mid$(X$, 387, 14))
       'NORIT% = CVI(Mid$(X$, 503, 2))
       NORIT% = CVI(Mid$(X$, 401, 2))     ' resolver enredo entre pedido y remito
       PREUNI = CVS(Mid$(X$, 429, 4)) * COEFPRE
       NPEDCAR& = XVALO(Mid$(X$, 371, 16))
       '
       If CI% > 0 Then
         '
         SQLX$ = "DoPriCor = '" & DOPRIC$ & "' "
         SQLX$ = SQLX$ + "AND Cod_Inte = " & CI% & " "
         SQLX$ = SQLX$ + "AND Nume_Clie = " & NC1% & " "
         SQLX$ = SQLX$ + "AND NuorItem = " & NORIT% & " "
         '
         With CONSITMP   ' Movisto
           .FindFirst SQLX$
25         If .NoMatch = False Then
             NPEDID& = 0
             If IsNull(!NUPEDCLI) = False Then NPEDID& = !NUPEDCLI
             If NPEDID& > 0 And NPEDCAR& > 0 And NPEDID& <> NPEDCAR& Then
               .FindNext SQLX$
               GoTo 25
             End If
             '
             ' ESTO ES PARA QUE NO CONFUNDA DEVOLUCIONES CON REMITOS - 23/05/07
             If Sgn(!CantSalid - !CantIngre) <> Sgn(CANTAA) Then
               .FindNext SQLX$
               GoTo 25
             End If
             '
             ' ESTO ES PARA QUE NO CONFUNDA LOS QUE YA ESTÁN PROCESADOS
             If !DoSecLar = NUFACLAR$ Then
               .FindNext SQLX$
               GoTo 25
             End If
             
             .Edit
             !DoSecCor = NUFACORT$
             !DoSecLar = NUFACLAR$
             !ValoUnit = PREUNI
             !MoneVal = MONEMI%     ' moneda de emision
             If MONEMI% <= 1 Then
                 !TipoCam = TICAMONE(2) ' siempre cotizacion del dolar
               Else
                 !TipoCam = TICAMONE(MONEMI%)
             End If
             If CANTAA > 0 Then
                 !CantSalid = CANTAA
                 !CantIngre = 0
               Else
                 !CantSalid = 0
                 !CantIngre = Abs(CANTAA)
             End If
             If NPEDID& < 1 Then
               !NUPEDCLI = NPEDCAR&
             End If
             .Update
             GoTo 29
           End If
         End With
       End If
       '
       FECHA7% = FECHANUM(Text9)
       VUNI7# = PREUNI
       Call MOVISTOK(FECHA7%, CI%, "", CMO$, NUFACLAR$, NUFACORT$, NUFACLAR$, NUFACORT$, RASOCLI$(NC1%), NORIT%, VUNI7#, MONEMI%, NC1%, DEPOEXPE%, (-1) * CANTAA, NPEDCAR&)
       '
     End If
29 Next I&
   '
   CONSITMP.Close
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
   COPIASFAC% = XVALO(FINFACTURA.Text2(0))
   If COPIASFAC% < 1 Then
      NUFACLAR$ = ""
      NUFACORT$ = ""
   End If
   '
   PREFIREM$ = CONTROL$("", "PREFIREM")
   If PREFIREM$ = "" Then PREFIREM$ = "0001"
   REMICOR$ = "RT." + TRIM$(Text10)
   REMILAR$ = Text10: While Len(REMILAR$) < 8: REMILAR$ = "0" + REMILAR$: Wend
   REMILAR$ = CMO$ + "-X-" + PREFIREM$ + "-" + REMILAR$
   If COPIASREM% < 1 Then
     REMICOR$ = NUFACORT$
     REMILAR$ = NUFACLAR$
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
   DEPOSALI% = XVALO(FINFACTURA.Text7)
   '\\\OT-07-0023-RG-FIN///
   For I& = 1 To FIN&
     Call TRACE(20, I&, FIN&)
     X$ = REGLEIDO$("!CARDEFAC", I&)
     CANTAA = CVS(Left$(X$, 4))
     If Abs(CANTAA) > 0 Then
       CI% = CVI(Mid$(X$, 5, 2))
       If CI% >= 0 Then                 ' para incluir remito texto libre
         NORIT% = I& - 1
         PREUNI# = CDbl(CVS(Mid$(X$, 429, 4)) * COEFPRE)
         '\\\OT-05-0562-WAR-02/08/05///
         SIG% = (-1)
         If TIFADOC% = 6 Then SIG% = 1 'NC
         '\\\OT-05-0562-WAR-02/08/05///
         NUPEDIORI& = XVALO(Mid$(X$, 371, 16))
         '
         Call MOVISTOK(FF%, CI%, LOTE$, CMO$, REMILAR$, REMICOR$, NUFACLAR$, NUFACORT$, RASOCLI$(NC1%), NORIT%, PREUNI#, MONEMI%, NC1%, DEPOSALI%, SIG% * CANTAA, NUPEDIORI&) '(-1) * CANTAA)
         '
         If DESCOMPO%(CI%) > 0 Then
              Call MOVISTOK(FF%, CI%, LOTE$, CMO$, REMILAR$, REMICOR$, NUFACLAR$, NUFACORT$, RASOCLI$(NC1%), NORIT%, PREUNI#, MONEMI%, NC1%, DEPOSALI%, CANTAA, NUPEDIORI&)  '(-1) * CANTAA)            MOVIMIENTO SALIDA PARA COMPONENTES DEL KIT
              Call AREXPLO(CI%, 1, 1)
              NUORIT% = 0
              For KKJ% = 1 To CAIT%
                NUORIT% = NUORIT% + 1
                CICO% = CIJ%(KKJ%)
                CACICO = USOI(KKJ%) * CANTAA
                Call MOVISTOK(FF%, CICO%, LOTE$, CMO$, REMILAR$, REMICOR$, NUFACLAR$, NUFACORT$, RASOCLI$(NC1%), NUORIT%, PREUNI#, MONEMI%, NC1%, DEPOSALI%, SIG% * CACICO, NUPEDIORI&) '(-1) * CANTAA)
              Next KKJ%
         End If
         '
         If CANTAA > 0 Then ' PARA EVITAR CON NOTAS DE CREDITO
           CANAX# = CANTAA
           NUORITPE% = CVI(Mid$(X$, 503, 2))
           If NUPEDIORI& > 0 Then
             Call BLANPEDIPE(NC1%, FF%, REMILAR$, CI%, CANAX#, NUPEDIORI&, NUORITPE%)
           End If
         End If
         '
       End If
     End If
   Next I&
   '
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
         SQLX$ = "DoPriLar = '" & DOPRIL$ & "' "
         SQLX$ = SQLX$ + "AND Cod_Inte = " & CI% & " "
         SQLX$ = SQLX$ + "AND Nume_Clie = " & NC1% & " "
         SQLX$ = SQLX$ + "AND NuorItem = " & NORIT% & " "
         '
         With Movisto
           .FindFirst SQLX$
25         If .NoMatch = False Then
             .Edit
             !ValoUnit = PREUNI
             .Update
           End If
         End With
       End If
     End If
   Next I&
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
     '
     If YAINDI% = 0 Then
         '
         Screen.MousePointer = 11
         
         Call ABREPEDCLI
         SQLX$ = " SELECT DISTINCT PEDENCA.NroPed, PEDENCA.NroClie, PEDENCA.FechEmis, PEDENCA.NroOcom "
         SQLX$ = SQLX$ + " FROM PEDENCA "
         SQLX$ = SQLX$ + " INNER JOIN  PEDDETA "
         SQLX$ = SQLX$ + " ON (PEDENCA.NroPed = PEDDETA.NroPed) "
         SQLX$ = SQLX$ + " WHERE PEDDETA.Status < 8 "   ' para que no este anulada ni cumplida
         SQLX$ = SQLX$ + " AND PEDDETA.CantEnt < (PEDDETA.CanPed - 0.005) "
         SQLX$ = SQLX$ + " AND PEDENCA.Status < 9 "
         SQLX$ = SQLX$ + " ORDER BY PEDENCA.NroPed ASC "
         Set PediModifTemp = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
         '
         JJ& = 0
         With PediModifTemp
           On Error Resume Next
           .MoveFirst
           If Err < 1 Then
             Do While Not .EOF
                 NRODES& = !NroPed
               TTXX$ = MKS$(NRODES&)
                 NOCO$ = !NroOcom
               If NOCO$ <> "" Then
                 TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
               End If
                 FEX = CVINT(!FECHEMIS)
               TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
                 NCLIE% = !NroClie
               TTXX$ = TTXX$ + TRIM$(RASOCLI$(NCLIE%)) + Space$(52)
               JJ& = JJ& + 1
               Call GRAREG("INDIPEP", Left$(TTXX$, 52) + MKS$(REGENCA&), JJ&)
             .MoveNext
             Loop
           End If
         End With
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
5    For KKI% = 0 To 5
       EDITFORM.Picture1(KKI%).Cls
       EDITFORM.Picture1(KKI%).DrawWidth = 3
       EDITFORM.Picture1(KKI%).DrawStyle = 0
       EDITFORM.Picture1(KKI%).FillStyle = 1
     Next KKI%
     '
     REDOCU$ = REGACT$("PDOCLST")
     URE& = ULTREG&("PDOCSPL")
     LI& = 0: LS& = URE&
6    E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
     L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
     DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
     If DCC& < DOCUNU& Then LI& = L&: GoTo 6
     If DCC& > DOCUNU& Then LS& = L&: GoTo 6
     GoTo 8
     '
7    Screen.MousePointer = 1
     Call MENSERR(24, "Documento no Registrado")
     GoTo 9999
     '
8    TPSP$ = ""
     Screen.MousePointer = 11
     While L& > 1
       DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
       If DCC& < DOCUNU& Then GoTo 9
       L& = L& - 1
     Wend
     '
9    CAPAGINAS% = 0
     For KKL& = L& To URE&
       TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
       If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
       If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
         TPSP$ = TPSP$ + Mid$(TBUF$, 5)
       End If
     Next KKL&
     '
95   SPOOLSTRING$ = TPSP$
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
10   PPXX% = InStr(TPSP$, Chr$(255))
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
     'EDITFORM.Command5.Enabled = False
     EDITFORM.Show 1
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
    DES1# = CDbl(CVDAT(DES%))
    HAS1# = CDbl(CVDAT(HAS%))
    '
    Call APERBASDAT("STOCKS")
    SQLX$ = " SELECT * FROM MOVISTO  "
    SQLX$ = SQLX$ + " WHERE CDBL(FechMov)>= " & DES1# & " "
    SQLX$ = SQLX$ + "AND CDBL(FechMov)<= " & HAS1# & " "
    SQLX$ = SQLX$ + "AND CodiMovi= 'RT' "
    SQLX$ = SQLX$ + "AND LEFT(DoSecCor,2)='FC' "
    SQLX$ = SQLX$ + "AND Nume_Clie= " & NCLIE% & " "
    SQLX$ = SQLX$ + "ORDER BY FechMov ASC "
    Set DIFPRETMP = Stocks.OpenRecordset(SQLX$, 4)
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
            PRENUE# = PRELISTA(NOLIST$, !Cod_Inte)
            '
            Call REPLA(YYY$, MKI(!Cod_Inte), 1, 2)
            Call REPLA(YYY$, !DoPriCor, 3, 12)
            Call REPLA(YYY$, !DoSecCor, 15, 12)
            Call REPLA(YYY$, MKI(CVINT(!FechMov)), 27, 2)
            Call REPLA(YYY$, MKD(!CantSalid), 29, 8)
            Call REPLA(YYY$, MKD(!ValoUnit), 37, 8)
            Call REPLA(YYY$, !DoPriLar, 81, 18)
            Call REPLA(YYY$, !DoSecLar, 99, 18)
            Call REPLA(YYY$, MKI(!NuOrItem), 117, 2)
            Call REPLA(YYY$, MKD(PRENUE#), 45, 8)
            JJ& = JJ& + 1
            Call GRAREG("!DIFPRECO", YYY$, JJ&)
       .MoveNext
       Loop
    End With
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
    z$ = REGBLAN("!CARDEFAC")
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
            Call REPLA(z$, MKI(CODI1%), 5, 5)
            Call REPLA(z$, DESCRI$, 7, 64)
            Call REPLA(z$, MKI(CVINT(Text9)), 419, 2)
            Call REPLA(z$, MKS(CANTI1#), 1, 4)
            Call REPLA(z$, MKS(DIFCOT1#), 421, 4)
            Call REPLA(z$, MKS(DIFCOT1#), 429, 4)
            Call REPLA(z$, MKD(IMPORDIF#), 433, 8)
            Call REPLA(z$, RTOLAR$, 387, 14)
            'Call REPLA(z$, MKI(NUMORDIT%), 503, 2)
            Call REPLA(z$, MKI(NUMORDIT%), 401, 2) ' resolver enredo entre pedido y remito
            JJ& = JJ& + 1
            Call GRAREG("!CARDEFAC", z$, JJ&)
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
   Call CALTOFA2
   '
End Sub
'
Sub CARDETPEDIPEN()
   '
   Static AGRUPAENTREGA%
   If AGRUPAENTREGA% = 0 Then
      AGRUPAENTREGA% = (-1)
      If CONTROL$("FACTUREM", "AGRUPEDI") = "TRANSUCU" Then AGRUPAENTREGA% = 1
   End If
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
    '\\\OT-07-0371-FL-01/10/07///
   SQLX$ = " SELECT PEDDETA.codiint, PEDDETA.CanPed, PEDDETA.CantEnt, PEDDETA.NroPed,PEDDETA.FechEmis,"
   SQLX$ = SQLX$ + "PEDDETA.FeSolEnt, PEDDETA.PreUnid, PEDDETA.MoneEmis, "
   SQLX$ = SQLX$ + "PEDDETA.TipoPed, PEDDETA.NuOrItem, PEDENCA.CondPag, "
   SQLX$ = SQLX$ + "PEDDETA.DescLista, PEDDETA.DescCombi, "
   SQLX$ = SQLX$ + "PEDDETA.CBASE , PEDDETA.TTALLE , PEDDETA.TCOLOR, "
   SQLX$ = SQLX$ + "PEDENCA.DomiEnt, PEDENCA.LocaEnt, PEDENCA.TranspEnt, "
   SQLX$ = SQLX$ + "PEDENCA.NTranspo, PEDENCA.TDescApli "
   SQLX$ = SQLX$ + "FROM PEDDETA "
   SQLX$ = SQLX$ + "INNER JOIN PEDENCA ON (PEDDETA.NroPed = PEDENCA.NroPed) "
   SQLX$ = SQLX$ + "WHERE PEDDETA.NroClie = " & NC% & " "
   SQLX$ = SQLX$ + "AND PEDDETA.Status < 8 "    ' NO ANULADOS NI CUMPLIDOS
   SQLX$ = SQLX$ + "AND PEDDETA.STATUS >= 0 "   ' APROBADOS
   '\\\OT-07-371-FL-FIN///
   '
   If HAS1# > 0 Then
      SQLX$ = SQLX$ + " AND CDBL(PEDDETA.FeSolEnt) <= " & HAS1# & " "
   End If
   '
   SQLX$ = SQLX$ + " AND (((PEDDETA.CanPed - PEDDETA.CantEnt) >= 0.05) OR (PEDDETA.TipoPed = 'ABIERTO' AND CDBL(PEDDETA.FinVigen) >= " & Int(CDbl(Now)) & ")) "
   SQLX$ = SQLX$ + " ORDER BY PEDDETA.NroPed, PEDDETA.NuOrItem;" 'ESTA LINEA LA AGREGUE
   '
   Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   With PEDTEMP
     On Error Resume Next
     .MoveFirst
     '
     JJ& = 0: CAITCAR% = 0
     '\\\OT-07-0371-FL-01/10/07///
     DESCARTAPED% = 0
     '\\\OT-07-0371-FL-FIN///
     Do While .EOF = False
       CODIMAT% = !CodiInt
       If CODIMAT% > 0 And CODIMAT% <= NUMAS% Then
         CAN = !CanPed - !CantEnt
           If CAPEN < 0 Then CAPEN = 0
         If !TipoPed = "ABIERTO" Then CAPEN = 999999
         NPEDID& = !NroPed
         FEMIPED% = CVINT(!FECHEMIS)
         FENTRESOL% = CVINT(!FESOLENT)
         PRECI1# = !PreUnid
         MONEMA% = !MoneEmis
            If MONEMA% > ULTREG&("TAMONED") Then MONEMA% = 1 ' ATAJA FUERA DE RANGO
         COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
         PRECIO = PRECI1# * COEFCON
'        ESTO QUE SIGUE ES ANTERIOR A LA IMPLEMENTACION DE DAKOT
'           If PRECIO < 0.00005 Then PRECIO = PRELISTA#(ARCHIPRE$, CODIMAT%)
'         PRELI = PRELISTA#(ARCHIPRE$, CODIMAT%)
'           If PRELI < PRECIO Then PRELI = PRECIO
'         PORDESCUEN = 0: If PRELI >= 0.0001 Then PORDESCUEN = 100 * (PRELI - PRECIO) / PRELI
         '
         ' ESTO CORRESPONDE A DAKOT
         PRELI = !PreList: If PRELI < 0.00005 Then PRELI = PRELISTA#(ARCHIPRE$, CODIMAT%)
         PDLISTA = XVALO(!DescLista)
         Dcombi$ = TRIM$(SAFETEXT$(!DescCombi)) + "+" + TRIM$(SAFETEXT$(!TDescApli))
         If Left$(Dcombi$, 1) = "+" Then Dcombi$ = Mid$(Dcombi$, 2)
         If Right$(Dcombi$, 1) = "+" Then Dcombi$ = Left$(Dcombi$, Len(Dcombi$) - 1)
         If Len(Dcombi$) > 16 Then
            Call MENSERR(24, "Imposible Aplicar Descuento Combinado - Consulte")
         End If
         PDCOMBI = PORCEDESCU(Dcombi$)
         PORDESCUEN = 0: If PRELI >= 0.0001 Then PORDESCUEN = 100 * (PRELI - PRECIO) / PRELI
         '
         If AGRUPAENTREGA% > 0 Then
            DOMIENPE$ = TRIM$(UCase$(!DOMIENT))
            LOCAENPE$ = TRIM$(UCase$(!LOCAENT))
            TRANENPE$ = TRIM$(UCase$(!TRANSPENT))
            CONDPAGO% = !CONDPAG
            NTRANS% = XVALO(!NTranspo)
           '
           If XVALO(Label32) < 1 Then
              If DOMIENPE$ <> "" Then LENTREGA(0) = DOMIENPE$: Label32 = "1"
              If LOCAENPE$ <> "" Then LENTREGA(1) = LOCAENPE$: Label32 = "1"
              If NTRANS% <> 0 Then Text19.TEXT = TRIM$(Str$(NTRANS%)): Label32 = "1"
              If TRANENPE$ <> "" Then LENTREGA(2) = TRANENPE$: Label32 = "1"
              If CONDPAGO% <> 0 Then TEXT21(1).TEXT = CONDPAGO%: Label32 = "1"
              
           End If
           If DOMIENPE$ <> "" Then If DOMIENPE$ <> LENTREGA(0) Then GoTo 28
           If LOCAENPE$ <> "" Then If LOCAENPE$ <> LENTREGA(1) Then GoTo 28
           If NTRANS% <> 0 Then If NTRANS% <> XVALO(Text19.TEXT) Then GoTo 28
           If TRANENPE$ <> "" Then If TRANENPE$ <> LENTREGA(2) Then GoTo 28
           If CONDPAGO% <> 0 Then If CONDPAGO% <> XVALO(TEXT21(1).TEXT) Then GoTo 28
         End If
         '
         COMPDESCRI$ = ""
         If CONTROL$("PEDIDO", "VARCOLOR") = "SI" Then
            COLORBA$ = !CBASE
            TELA$ = !TTALLE
            COLTELA$ = !TCOLOR
            COMPDESCRI$ = "CB: " + COLORBA$ + " - T: " + TELA$ + " - CT: " + COLTELA$
         End If
         '\\\OT-07-0371-FL-FIN///
         '
5        z$ = REGBLAN$("!CARDEFAC")
         Call REPLA(z$, MKS$(CAN), 1, 4)      ' cantidad
           COD$ = CODEXT$(CODIMAT%)
         Call REPLA(z$, MKI$(CODIMAT%), 5, 2)
         Call REPLA(z$, DESCRIT$(CODIMAT%), 7, 64)
         Call REPLA(z$, COMPDESCRI$, 75, 64)
           NREMIT$ = ""
         Call REPLA(z$, NREMIT$, 387, 14) ' numero de remito
           FEREMIT% = 0
         Call REPLA(z$, MKI$(FEREMIT%), 419, 2)
         Call REPLA(z$, MKI$(0), 401, 2)
         Call REPLA(z$, NUPLANO$(CODIMAT%), 315, 24)
         '
         Call REPLA(z$, MKS$(PRELI), 421, 4)
         Call REPLA(z$, MKS$(PORDESCUEN), 425, 4)
         Call REPLA(z$, MKS$(PRECIO), 429, 4)
           IMPO# = CDbl(PRECIO) * CAN
         Call REPLA(z$, MKD$(IMPO#), 433, 8)
         Call REPLA(z$, MKS$(PORCEIVA(CODIMAT%)), 441, 4)
           NORITPE% = !NuOrItem
         Call REPLA(z$, MKI$(NORITPE%), 503, 2)
         Call REPLA(z$, Dcombi$, 487, 16)
         Call REPLA(z$, MKS$(PDLISTA), 505, 4)
         Call REPLA(z$, MKS$(PDCOMBI), 509, 4)
         Call REPLA(z$, TRIM$(Str$(NPEDID&)), 371, 16)
           NUOCOX$ = NUOCOCLI$(NPEDID&, CODIMAT%)
10       Call REPLA(z$, NUOCOX$, 355, 16)
         '
         JJ& = JJ& + 1
         Call GRAREG("!CARDEFAC", z$, JJ&)
         CAITCAR% = CAITCAR% + 1
         '
       End If
       '\\\OT-07-0371-FL-01/10/07///
       GoTo 29
28     DESCARTAPED% = 1
29   .MoveNext
       '\\\OT-07-0371-FL-FIN///
     Loop
   End With
   '
   Call SETULTREG("!CARDEFAC", JJ&)
   '
    '\\\OT-07-0371-FL-01/10/07///
   If DESCARTAPED% = 1 Then
      Call MENSERR(24, "Se han Descartado Pedidos que no Corresponden\con la Sucursal o con la Condición de Venta\del Pedido más Antiguo")
   End If
    '\\\OT-07-0371-FL-FIN///
'
   JJ& = 0
   For U& = 1 To ULTREG&("!CARDEFAC")
     z$ = REGLEIDO$("!CARDEFAC", U&)
     If Abs(CVS(Mid$(z$, 1, 4))) >= 0.01 Then
       JJ& = JJ& + 1
       Call GRAREG("!CARDEFAC", z$, JJ&)
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
   ' AGREGADO
     '\\\OT-07-0371-FL-01/10/07///
   Static AGRUPAENTREGA%
   If AGRUPAENTREGA% = 0 Then
      AGRUPAENTREGA% = (-1)
      If CONTROL$("FACTUREM", "AGRUPEDI") = "TRANSUCU" Then AGRUPAENTREGA% = 1
   End If
     '\\\OT-07-0371-FL-FIN///
   ' FIN AGREGADO
   '
   Screen.MousePointer = 11
   NC% = XVALO(Text1)
   Call ABREPEDCLI
   SQLX$ = " SELECT DISTINCT PEDENCA.NroPed, PEDENCA.NroClie, PEDENCA.FechEmis, PEDENCA.NroOcom, "
   SQLX$ = SQLX$ + " PEDENCA.FeSolent, PEDENCA.NSucEnt, PEDENCA.TDescApli "
   SQLX$ = SQLX$ + " FROM PEDENCA "
   SQLX$ = SQLX$ + " INNER JOIN  PEDDETA "
   SQLX$ = SQLX$ + " ON (PEDENCA.NroPed = PEDDETA.NroPed) "
   SQLX$ = SQLX$ + " WHERE PEDENCA.NroClie = " & NC% & " "
   SQLX$ = SQLX$ + " AND PEDDETA.STATUS >= 0 "   ' APROBADOS
   SQLX$ = SQLX$ + " AND PEDDETA.Status < 8 "    ' para que no este anulada ni cumplida
   SQLX$ = SQLX$ + " AND PEDDETA.CantEnt < (PEDDETA.CanPed - 0.005) "
   SQLX$ = SQLX$ + " AND PEDENCA.Status < 9 "    ' para que no esté anulado
   SQLX$ = SQLX$ + " ORDER BY PEDENCA.NroPed ASC "
   Set PediModifTemp = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   Dim SELPE As Form
   On Error Resume Next
   Set SELPE = SELPEDIPEN
   If InStr(UCase$(EMPREAC$), "DAKOT") > 0 Then
      Set SELPE = SELPEDIDAK
   End If
   On Error GoTo 0
   '
   SELPE.LICOPEN.Clear
   SELPE.Label5 = "0"
   SELPE.Text1 = ""
   HAYPEDIPE% = 0
   With PediModifTemp
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
         Do While Not .EOF
'           ESTO SE MODIFICO POR DAKOT EL 13 DE OCTUBRE 2008
'             NRODES& = !NroPed
'           TTXX$ = FORMATNUM$(NRODES&, "F6.0") + "  "
'             NOCO$ = TRIM$(!NroOcom)
'           If NOCO$ <> "" Then
'             TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
'           End If
'             FEX = CVINT(!FECHEMIS)
'           TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
'             NCLIE% = !NroClie
'           TTXX$ = TTXX$ + TRIM$(RASOCLI$(NCLIE%)) + Space$(52)
             NRODES& = !NroPed
           TTXX$ = FORMATNUM$(NRODES&, "F6.0") + "  "
             NSUCPE% = XVALO(!NSucEnt)
             NOCO$ = TRIM$(!NroOcom)
             FEX = CVINT(!FECHEMIS): FEMI$ = FECHATEX$(FEX)
             FEX = CVINT(!FESOLENT): FENT$ = FECHATEX$(FEX)
             TDESCU$ = AJUSTI$(SAFETEXT$(!TDescApli), 14)
           TTXX$ = TTXX$ + FEMI$ + "  " + FENT$
           TTXX$ = TTXX$ + "  " + TDESCU$
           TTXX$ = TTXX$ + "  " + FORMATNUM$(NSUCPE%, "I4")
           TTXX$ = TTXX$ + "    " + NOCO$
           SELPE.LICOPEN.AddItem TTXX$
           HAYPEDIPE% = 1
         .MoveNext
         Loop
      End If
      On Error GoTo 0
   End With
   '
   Screen.MousePointer = 1
   If HAYPEDIPE% > 0 Then
        SELPE.Show 1
      Else
        Call COMUNI("No Hay Pedidos APROBADOS Pendientes\de Entrega para el Cliente Elegido.")
   End If
   '
   CANSELE% = XVALO(SELPE.Label5)
   If CANSELE% < 1 Then
      JJ& = 0
      GoTo 20     ' PARA BLANQUEAR ARCHIVO DE CARGA
   End If
   '
   Call APERBASDAT("PEDCLI")
'     '\\\OT-07-0371-FL-01/10/07///
   SQLX$ = " SELECT PEDDETA.codiint, PEDDETA.CanPed, PEDDETA.CantEnt, PEDDETA.NroPed,PEDDETA.FechEmis, "
   SQLX$ = SQLX$ + "PEDDETA.FeSolEnt, PEDDETA.PreUnid, PEDDETA.MoneEmis,PEDDETA.TipoPed,PEDDETA.NuOrItem, PEDENCA.CondPag, "
   SQLX$ = SQLX$ + "PEDDETA.CBASE , PEDDETA.TTALLE , PEDDETA.TCOLOR, "
   SQLX$ = SQLX$ + "PEDDETA.PreList, PEDDETA.DescLista, PEDDETA.DescCombi, "
   SQLX$ = SQLX$ + "PEDENCA.DomiEnt, PEDENCA.LocaEnt, PEDENCA.TranspEnt, "
   SQLX$ = SQLX$ + "PEDENCA.NTranspo, PEDENCA.NSucEnt, PEDENCA.TDescApli "
   SQLX$ = SQLX$ + " FROM PEDDETA "
   SQLX$ = SQLX$ + "INNER JOIN PEDENCA ON (PEDDETA.NroPed = PEDENCA.NroPed) "
   SQLX$ = SQLX$ + " WHERE PEDDETA.NroClie = " & NC% & " "
   SQLX$ = SQLX$ + " AND PEDDETA.Status < 8 "    ' NO ANULADOS NI CUMPLIDOS
   SQLX$ = SQLX$ + " AND PEDDETA.STATUS >= 0 "   ' APROBADOS
   SQLX$ = SQLX$ + " AND (((PEDDETA.CanPed - PEDDETA.CantEnt) >= 0.05) OR (PEDDETA.TipoPed = 'ABIERTO' AND CDBL(PEDDETA.FinVigen) >= " & Int(CDbl(Now)) & ")) "
   '
   ATXX% = 0
   For KKU& = 1 To SELPE.LICOPEN.ListCount
      If SELPE.LICOPEN.Selected(KKU& - 1) = True Then
         NPESEL& = XVALO(Left$(SELPE.LICOPEN.List(KKU& - 1), 6))
         If ATXX% < 1 Then
             SQLX$ = SQLX$ + "AND (PEDDETA.NroPed = " & NPESEL& & " "
             ATXX% = 1
           Else
             SQLX$ = SQLX$ + "OR PEDDETA.NroPed = " & NPESEL & " "
         End If
      End If
   Next KKU&
   SQLX$ = SQLX$ + ") "
   '
   SQLX$ = SQLX$ + " ORDER BY PEDDETA.NroPed, PEDDETA.NuOrItem;" 'ESTA LINEA LA AGREGUE
   '
   Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   With PEDTEMP
     On Error Resume Next
     .MoveFirst
      If Err Then
        On Error GoTo 0
        Call COMUNI("Ningun Registro Coincide con Condición SQL.")
        GoTo 20
      End If
      On Error GoTo 0
      DESCARTAPED% = 0
      '
      JJ& = 0: CAITCAR% = 0
      Do While (.EOF = False)
        CODIMAT% = !CodiInt
        If CODIMAT% > 0 And CODIMAT% <= NUMAS% Then
          CAN = !CanPed - !CantEnt
             If CAPEN < 0 Then CAPEN = 0
          If !TipoPed = "ABIERTO" Then CAPEN = 999999
          NPEDID& = !NroPed
          FEMIPED% = CVINT(!FECHEMIS)
          FENTRESOL% = CVINT(!FESOLENT)
          PRECI1# = !PreUnid
          MONEMA% = !MoneEmis
             If MONEMA% > ULTREG&("TAMONED") Then MONEMA% = 1 ' ATAJA FUERA DE RANGO
          COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
          PRECIO = PRECI1# * COEFCON
'        ESTO QUE SIGUE ES ANTERIOR A LA IMPLEMENTACION DE DAKOT
'           If PRECIO < 0.00005 Then PRECIO = PRELISTA#(ARCHIPRE$, CODIMAT%)
'         PRELI = PRELISTA#(ARCHIPRE$, CODIMAT%)
'           If PRELI < PRECIO Then PRELI = PRECIO
'         PORDESCUEN = 0: If PRELI >= 0.0001 Then PORDESCUEN = 100 * (PRELI - PRECIO) / PRELI
         '
         ' ESTO CORRESPONDE A DAKOT
          PRELI = !PreList: If PRELI < 0.00005 Then PRELI = PRELISTA#(ARCHIPRE$, CODIMAT%)
         PDLISTA = XVALO(!DescLista)
         Dcombi$ = TRIM$(SAFETEXT$(!DescCombi))
         DPEDIDO$ = TRIM$(SAFETEXT$(!TDescApli))
         If TRIM$(DPEDIDO$) <> "" Then PORDESCU = DPEDIDO$
         'DCOMBI$ = DCOMBI0$ + "+" + DPEDID0$
         If Left$(Dcombi$, 1) = "+" Then Dcombi$ = Mid$(Dcombi$, 2)
         If Right$(Dcombi$, 1) = "+" Then Dcombi$ = Left$(Dcombi$, Len(Dcombi$) - 1)
         If Len(Dcombi$) > 16 Then
            Call MENSERR(24, "Imposible Aplicar Descuento Combinado - Consulte")
         End If
         PDCOMBI = PORCEDESCU(Dcombi$)
         'If DPEDID0$ <> "" Then PRECIO = PRECIO * (1 - PORCEDESCU(DPEDID0$) / 100)
         'PORDESCUEN = 0: If PRELI >= 0.0001 Then PORDESCUEN = 100 * (PRELI - PRECIO) / PRELI
         '
         ' AGREGADO
         If AGRUPAENTREGA% > 0 Then
           DOMIENPE$ = TRIM$(UCase$(!DOMIENT))
           LOCAENPE$ = TRIM$(UCase$(!LOCAENT))
           TRANENPE$ = TRIM$(UCase$(!TRANSPENT))
           CONDPAGO% = !CONDPAG
           SUCUEN% = !NSucEnt
           NTRANS% = XVALO(!NTranspo)
           '
           If XVALO(Label32) < 1 Then
              If DOMIENPE$ <> "" Then LENTREGA(0) = DOMIENPE$: Label32 = "1"
              If LOCAENPE$ <> "" Then LENTREGA(1) = LOCAENPE$: Label32 = "1"
              If NTRANS% <> 0 Then Text19.TEXT = TRIM$(Str$(NTRANS%)): Label32 = "1"
              If TRANENPE$ <> "" Then LENTREGA(2) = TRANENPE$: Label32 = "1"
              If CONDPAGO% <> 0 Then TEXT21(1).TEXT = CONDPAGO%: Label32 = "1"
              If SUCUEN% <> 0 Then Text20.TEXT = TRIM$(Str$(SUCUEN%)): Label32 = "1"
           End If
           If DOMIENPE$ <> "" Then If DOMIENPE$ <> LENTREGA(0) Then GoTo 28
           If LOCAENPE$ <> "" Then If LOCAENPE$ <> LENTREGA(1) Then GoTo 28
           If NTRANS% <> 0 Then If NTRANS% <> XVALO(Text19.TEXT) Then GoTo 28
           If TRANENPE$ <> "" Then If TRIM$(TRANENPE$) <> TRIM$(LENTREGA(2)) Then GoTo 28
           If CONDPAGO% <> 0 Then If CONDPAGO% <> XVALO(TEXT21(1).TEXT) Then GoTo 28
           If SUCUEN% <> 0 Then If SUCUEN% <> XVALO(Text20.TEXT) Then GoTo 28
         End If
         ' FIN AGREGADO
         '
5        z$ = REGBLAN$("!CARDEFAC")
         Call REPLA(z$, MKS$(CAN), 1, 4)      ' cantidad
           COD$ = CODEXT$(CODIMAT%)
         Call REPLA(z$, MKI$(CODIMAT%), 5, 2)
         Call REPLA(z$, DESCRIT$(CODIMAT%), 7, 64)
           NREMIT$ = ""
         Call REPLA(z$, NREMIT$, 387, 14) ' numero de remito
           FEREMIT% = 0
         Call REPLA(z$, MKI$(FEREMIT%), 419, 2)
         Call REPLA(z$, MKI$(0), 401, 2)
         Call REPLA(z$, NUPLANO$(CODIMAT%), 315, 24)
         '
         Call REPLA(z$, MKS$(PRELI), 421, 4)
         Call REPLA(z$, MKS$(PORDESCUEN), 425, 4)
         Call REPLA(z$, MKS$(PRECIO), 429, 4)
           IMPO# = CDbl(PRECIO) * CAN
         Call REPLA(z$, MKD$(IMPO#), 433, 8)
         Call REPLA(z$, MKS$(PORCEIVA(CODIMAT%)), 441, 4)
           NORITPE% = !NuOrItem
         Call REPLA(z$, MKI$(NORITPE%), 503, 2)
         Call REPLA(z$, Dcombi$, 487, 16)
         Call REPLA(z$, MKS$(PDLISTA), 505, 4)
         Call REPLA(z$, MKS$(PDCOMBI), 509, 4)
         Call REPLA(z$, TRIM$(Str$(NPEDID&)), 371, 16)
           NUOCOX$ = NUOCOCLI$(NPEDID&, CODIMAT%)
10       Call REPLA(z$, NUOCOX$, 355, 16)
         '
       
         COMPDESCRI$ = ""
         If CONTROL$("PEDIDO", "VARCOLOR") = "SI" Then
            COLORBA$ = SAFETEXT(!CBASE)
            TELA$ = SAFETEXT(!TTALLE)
            COLTELA$ = SAFETEXT(!TCOLOR)
            Call REPLA(z$, COLORBA$, 283, 10)
            Call REPLA(z$, TELA$, 293, 10)
            Call REPLA(z$, COLTELA$, 315, 24)
            If TRIM$(COLOBA$) + TRIM$(TELA$) + TRIM$(COLTELA$) <> "" Then
              COMPDESCRI$ = "CB: " + COLORBA$ + " - T: " + TELA$ + " - CT: " + COLTELA$
            End If
         End If
         Call REPLA(z$, COMPDESCRI$, 75, 64)
         '
         JJ& = JJ& + 1
         Call GRAREG("!CARDEFAC", z$, JJ&)
         CAITCAR% = CAITCAR% + 1
         '
       End If
       ' AGREGADO
         '\\\OT-07-0371-FL-01/10/07///
       GoTo 29
28     DESCARTAPED% = 1
        '\\\OT-07-0371-FL-FIN///
       ' FIN AGREGADO
29     .MoveNext
     Loop
   End With
   '
20 Call SETULTREG("!CARDEFAC", JJ&)
   '
   ' AGREGADO
     '\\\OT-07-0371-FL-01/10/07///
   If DESCARTAPED% = 1 Then
      Call MENSERR(24, "Se han Descartado Pedidos que no Corresponden\con la Sucursal o con la Condición de Venta\del Pedido más Antiguo")
   End If
     '\\\OT-07-0371-FL-FIN///
   ' FIN AGREGADO
   '
   JJ& = 0
   For U& = 1 To ULTREG&("!CARDEFAC")
     z$ = REGLEIDO$("!CARDEFAC", U&)
     If Abs(CVS(Mid$(z$, 1, 4))) >= 0.01 Then
       JJ& = JJ& + 1
       Call GRAREG("!CARDEFAC", z$, JJ&)
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
Sub GRAFACTURA_NUE(OKX%)
   '
   Dim NIVA7#(10), NIVA7DOL#(10), CUEIVX$(10)
   '
   COTIMONEM# = XVALO(TICAMBIS) ' para prever ajuste de moneda de emision
   If COTIMONEM# < 0.05 Then COTIMONEM# = 1
   '
   OKX% = 0
   GoSub APRUEBACOPIAS       ' APROBACION FINAL Y CANTIDAD DE COPIAS
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
   Call APERBASDAT("CABAN")
   NOCLOSE% = 1
   CondiPago% = XVALO(TEXT21(1))
   SucuClie% = XVALO(Text20)
   GoSub GRATOTFAC           ' GRABA CAJABANC PARA FACTURA
   GoSub GRAVENETA           ' GRABA VENTA NETA
   GoSub GRAEMBALA           ' GRABA IMPORTE EMBALAJES
   GoSub GRABAIVA            ' graba iva inscripto Y NO INSCRIPTO
   GoSub GRAPERCEP           ' graba PERCEPCIONES IVA E I.BRUTOS
   GoSub GRABACOBRA
   '
   'Call BLIMESS("Calculando Vencimientos ...")
   If TIFADOC% <> 6 Or TRIM$(DOCUDIF) = "" Then
     If TOTAGEN# > ICANCE# Then
       Call GRADEUFA(TIPO%, VA%, NRO, NC%, FFI%, CPAG%, TOTAGEN# - ICANCE#)
     End If
   End If
   CajaBanc.Close
   SaDeuPen.Close
   NOCLOSE% = 0
   Call LIBARCH("CUECORR")
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
   COPIASFAC% = XVALO(FINFACTURA.Text2(0))
   COPIASREM% = XVALO(FINFACTURA.Text2(1))
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
      OPIMPUTB.Show 1
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
   If TIPO% = 3 Then TICOMX$ = "ND"
   If TIPO% = 6 Then TICOMX$ = "NC"
   NFCZ$ = TRIM$(Text11): While Len(NFCZ$) < 8: NFCZ$ = "0" + NFCZ$: Wend
   NFCY$ = TICOMX$ + "-" + Label12 + "-" + NFCZ$ + "-" + Label13
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
22     LMCR# = LIMICRE(NCLIE%)
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
   TOTAGEN# = 0: TOTAGENDOL# = 0
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
   '
Return                      ' ***********************************************************
'
ACSALCUEN:                  ' ACTUALIZA SALDO DE LA CUENTA
   LU$ = LUDAT$
   ARCHIDEU$ = "SVD202"
   ARCHICLI$ = "DEUDOR"
   ARCHIFAC$ = "FACTUR"
   RD% = 1
   TIPO% = TIFADOC%
   VA% = VARIDOC%
   NRO = XVALO(Text11.TEXT)
   NC% = NCLIE%
   FFI% = FECHANUM(Text9.TEXT)
   CPAG% = XVALO(TEXT21(1).TEXT)
   '
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
Return                      ' ***********************************************************
'
RENUMHOJAS:                 ' RENUMERA PARA CONSIDERAR LAS HOJAS QUE SE CONSUMIRAN
   'FORIPRE$ = TRIM$(CONTROL$("", "FORFAC-" + Label13.Caption))
   FORIPRE$ = FORELFAC$
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
   If TIPO% = 3 Then TICOMX$ = "ND": OPMOVIX$ = "DEB"
   If TIPO% = 6 Then TICOMX$ = "NC": OPMOVIX$ = "CRE"
   NFCZ$ = TRIM$(Text11): While Len(NFCZ$) < 8: NFCZ$ = "0" + NFCZ$: Wend
   NFCX$ = TICOMX$ + "-" + Mid$(NFCZ$, 3) + "-" + Label13
   NFCY$ = TICOMX$ + "-" + Label12 + "-" + NFCZ$ + "-" + Label13
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
   CajaBanc.Close
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
       If FORCARGA$ <> "SINREMITO" Then Call NUREMIFA(NROREMIT&)
       Text10.TEXT = TRIM$(Str$(NROREMIT&))
       Call MENSERR(24, "Número de Comprobante '" + NFCY$ + "' Ocupado.\   \FLEXOFT ha Re-Numerado el Documento\como Número '" + TRIM$(Text11.TEXT) + "'.\  \Verifique y Vuelva a Grabar.")
       Exit Sub
     End If
   End If
   If TIFADOC% < 3 Then
     If COPIASREM% > 0 Then
       If VERINURE% < 0 Then
         Call LIBARCH("CUECORR")
         NREMIRE$ = TRIM$(Text10)
         NROREMIT& = 0
         If FORCARGA$ <> "SINREMITO" Then Call NUREMIFA(NROREMIT&)
         Text10.TEXT = TRIM$(Str$(NROREMIT&))
         If NROREMIT& > 0 Then
            Call MENSERR(24, "Número de Remito '" + NREMIRE$ + "' Ocupado.\   \FLEXOFT ha Re-Numerado el Remito\como Número '" + TRIM$(Text10.TEXT) + "'.\  \Verifique y Vuelva a Grabar.")
            Exit Sub
         End If
       End If
     End If
   End If
   '
Return                      ' ***********************************************************
'
GRATOTFAC:                  ' GRABA TOTAL DE FACTURA
30 CoComCor$ = NFCX$
   CoComLar$ = NFCY$
   CodiEmp% = 0
   CodiSuc% = 0
   NumeClie% = GRUDIVCLI%(NC%)
   NumeProv% = 0
   FechEmi = CVDAT(FFI%)
   Referen$ = RASOCLI$(NC%)
   FeContab = CVDAT(FFI%)
   FeAcredi = CVDAT(FFI%)
   CuentaCo% = CUECOCLI%(NC%)
   TiMovi$ = "FVEN"
   OpMovi$ = OPMOVIX$
   VaMovi$ = Label13
   Descrip$ = Label7 + " " + Label13.Caption + " " + Label6.Caption + " " + Label12.Caption + "-" + Text11.TEXT
   NumeTal = XVALO(Label12)
   NumeCom = CDbl(Text11)
   StatInit = 0
   ImpoDeb = TOTAGEN# * SIGX%
   ValoDola = TOTAGENDOL# * SIGX%
   Observa$ = RASOCLI$(NC%)
   CondiPago% = XVALO(TEXT21(1))
   Call GRACABAN
   '
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
GRABACOBRA:
   ICANCE# = 0
   If CONTROFIS% = 1 Then
     If CANCECON% = 1 Then
       ICANCE# = XVALO(FINCONFIS.Text2)
       MECO% = XVALO(FINCONFIS.Text1)
       CUECO% = CUEMECOB%(MECO%)
       If ICANCE# >= 0.005 Then
         TiMovi$ = "RCOB"
         OpMovi$ = OPMOVIX$
         VaMovi$ = Label13
         ImpoDeb = (-1) * ICANCE#
         Call GRACABAN
         '
         CuentaCo% = CUECO%
         TiMovi$ = "CHRECI"
         OpMovi$ = "CH"
         VaMovi$ = ""
         Descrip$ = FINCONFIS.Label2
             If FECHE% < 1 Then FECHE% = FFI%
         FechEmi = CVDAT(FFI%)
         FeAcredi = CVDAT(FFI%)
             STATUI% = 0
             If CUECO% = CUECARTE% Then STATUI% = 1
         StatInit = STATUI%
         ImpoDeb# = ICANCE#
         Observa$ = ""
         Call GRACABAN
       End If
     End If
   End If
Return

End Sub
'
Function FORELFAC$()   ' DEVUELVE EL NOMBRE DEL FORMULARIO ELECTRONICO DE FACTURACIÓN
    '
    LTFX1$ = Label13.Caption
    If LTFX1$ = "C" Then LTFX1$ = "A"
    FEFAC$ = TRIM$(CONTROL$("", "FORFAC-" + LTFX1$))
    TIPOCOM$ = TRIM$(Label6)
    If TIPOCOM$ = "(T.Libre)" Then
      If Label13.Caption <> "E" Then   ' EXCLUYE FACTURAS DE EXPORTACION
        FEFAC$ = TRIM$(CONTROL$("", "FACTEXLI"))
        If UCase$(Label13) = "B" Then
          FEFAC$ = TRIM$(CONTROL$("", "FACTELIB"))
        End If
      End If
    End If
    '\\\OT-06-0038-WAR-01/02/06///
    FORIALT$ = ""
    If TIFADOC% = 3 Or TIFADOC% = 6 Then 'SI NC O ND
      If TIFADOC% = 3 Then ' SI NOTA DE DEBITO
        If UCase$(Label13.Caption) = "A" Or UCase$(Label13.Caption) = "C" Then
          FORIALT$ = CONTROL$("", "FORNDE-A") 'ND-A
          ElseIf UCase$(Label13.Caption) = "B" Then
          FORIALT$ = CONTROL$("", "FORNDE-B") 'ND-B
        End If
      End If
      '
      If TIFADOC% = 6 Then ' SI NOTA DE CREDITO
        If UCase$(Label13.Caption) = "A" Or UCase$(Label13.Caption) = "A" Then
          FORIALT$ = CONTROL$("", "FORNCR-A") 'NC-A
          ElseIf UCase$(Label13.Caption) = "B" Then
          FORIALT$ = CONTROL$("", "FORNCR-B") 'NC-B
        End If
      End If
    End If
    If TRIM$(FORIALT$) <> "" Then
      FEFAC$ = FORIALT$
    End If
    ''\\\OT-06-0038-WAR-FIN///
    '
    ' VALIDACIONES
    If FEFAC$ = "" Then
       Call MENSERR(24, "Falta Definir Controlador de Impresión")
       FEFAC$ = "NULL"
    End If
    '
    If EXISTE%(LUDAT$ + FEFAC$ + ".FRM") < 1 Then
       Call MENSERR(24, "Imposible Imprimir Documento.\No Existe Formulario Electrónico '" + FEFAC$ + ".FRM'.")
       FEFAC$ = "NULL"
    End If
    '
    FORELFAC$ = FEFAC$
    '
End Function

Sub SETTELIB()
    FORIPRE$ = FORELFAC$
    If FORIPRE$ = "NULL" Then
       Call FINAL("")
       Exit Sub
    End If
    AA1$ = ATRIFORM$(FORIPRE$, "DES-MAT/FORMATO")
    ACOLDE = Val(Mid$(AA1$, 2))
    If ACOLDE < 1 Or ACOLDE > 64 Then ACOLDE = 64 ' PARA PODER MANDAR LO AL !CARDEFAC
    '
    TEPRUE.Font = Text28.Font
    TEPRUE.FontSize = Text28.FontSize
    TEPRUE.FontBold = Text28.FontBold
    TEPRUE = String$(ACOLDE + 3, "*")
    ATEX = TEPRUE.Width + 50
    '
          ANCHOTO = Label14.Width + Text28.Width + Label15.Width
          If ATEX > ANCHOTO - 50 Then ATEX = ANCHOTO - 50
          Label14.Width = (ANCHOTO - ATEX) / 2
          Label15.Width = (ANCHOTO - ATEX) / 2
          Text28.Left = Label15.Left + Label15.Width
          Text28.Width = ATEX
          Label14.Left = Text28.Left + Text28.Width
          '
    '
    Frame12.Top = Picture5.Top
    Frame12.Left = Picture5.Left
    Frame12.Visible = True
    '

End Sub

Sub PRICONFIS()
' DISTINGUIR:
' A) fACTURA / DEBITO / CREDITO
' B) A O B
   '
   If TIFADOC% = 2 Then                ' ticket factura
        If VARIDOC% = 1 Then           ' factura a
            Call PRITIFACA
          ElseIf VARIDOC% = 2 Then     ' factura b
            Call PRITIFACB
          Else
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

Sub PRITIFACA()
    '
    Call FORHASAR.PRITIFACA
    '
End Sub

Sub PRITIFACB()
    '
    Call FORHASAR.PRITIFACB
    '
End Sub

Sub PRINOCREA()
    '
    Call FORHASAR.PRINOCREA
    '
End Sub

Sub PRINOCREB()
   '
   Call FORHASAR.PRINOCREB
   '
End Sub

Sub PRINODEBA()
   '
   Call FORHASAR.PRINODEBA
   '
End Sub

Sub PRINODEBB()
   '
   Call FORHASAR.PRINODEBB
   '
End Sub

Sub CARFADAKOT()
   FORFADAKO.Visible = True
   Tcanti(0).SetFocus
End Sub

Sub AGREGARENGLON(Index)
    If Index > Tcanti.Count - 2 Then
       NUPIC% = Tcanti.Count
       Load Tcanti(NUPIC%): Tcanti(NUPIC%) = ""
       Load Tcodigo(NUPIC%): Tcodigo(NUPIC%) = ""
       Load Tdescri(NUPIC%): Tdescri(NUPIC%) = ""
       Load Tpreca(NUPIC%): Tpreca(NUPIC%) = ""
       Load Dlista(NUPIC%): Dlista(NUPIC%) = ""
       Load Descom(NUPIC%): Descom(NUPIC%) = ""
       Load Pruneto(NUPIC%): Pruneto(NUPIC%) = ""
       Load Ineto(NUPIC%): Ineto(NUPIC%) = ""
       Tcanti(NUPIC%).Top = Tcanti(NUPIC% - 1).Top + Tcanti(NUPIC% - 1).Height
       Tcanti(NUPIC%).Visible = True
       Tcodigo(NUPIC%).Top = Tcanti(NUPIC%).Top
       Tcodigo(NUPIC%).Visible = True
       Tdescri(NUPIC%).Top = Tcanti(NUPIC%).Top
       Tdescri(NUPIC%).Visible = True
       Tpreca(NUPIC%).Top = Tcanti(NUPIC%).Top
       Tpreca(NUPIC%).Visible = True
       Dlista(NUPIC%).Top = Tcanti(NUPIC%).Top
       Dlista(NUPIC%).Visible = True
       Descom(NUPIC%).Top = Tcanti(NUPIC%).Top
       Descom(NUPIC%).Visible = True
       Pruneto(NUPIC%).Top = Tcanti(NUPIC%).Top
       Pruneto(NUPIC%).Visible = True
       Ineto(NUPIC%).Top = Tcanti(NUPIC%).Top
       Ineto(NUPIC%).Visible = True
     End If
End Sub

Private Sub Tpreca_Change(Index As Integer)
     Call CALPRENETO(Index)
End Sub

Private Sub Tpreca_GotFocus(Index As Integer)
     Tdescri(Index).ZOrder 0
End Sub

Private Sub Tpreca_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = 40 Then ' FLECHA ABAJO
       Call AGREGARENGLON(Index)
       Tpreca(Index + 1).SetFocus
     ElseIf KeyCode = 13 Then
       Descom(Index).SetFocus
   End If
End Sub

Private Sub Tpreca_LostFocus(Index As Integer)
   Tpreca(Index) = FORMATNUM(XVALO(Tpreca(Index)), "F12.2")
End Sub
