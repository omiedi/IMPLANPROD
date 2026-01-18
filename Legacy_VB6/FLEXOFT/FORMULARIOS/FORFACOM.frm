VERSION 5.00
Begin VB.Form FORFACOM 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturación de Proveedores (Operaciones Comerciales de Compras)"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   11910
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0FFD0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8500
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   12000
      Begin VB.Frame Frame12 
         BackColor       =   &H00E0FFD0&
         Caption         =   "Resumen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   210
         TabIndex        =   128
         Top             =   3300
         Width           =   6465
         Begin VB.TextBox Text17 
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
            Left            =   1860
            TabIndex        =   129
            Top             =   480
            Width           =   1365
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Importe"
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
            Left            =   2625
            TabIndex        =   137
            Top             =   240
            Width           =   1260
         End
         Begin VB.Label Label31 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   4920
            TabIndex        =   135
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tot.Neto"
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
            Left            =   4515
            TabIndex        =   134
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label29 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   3300
            TabIndex        =   133
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label28 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   240
            TabIndex        =   132
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tot.Bruto"
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
            Left            =   -210
            TabIndex        =   131
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "% Bonif."
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
            Left            =   1260
            TabIndex        =   130
            Top             =   240
            Width           =   1260
         End
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0FFD0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   5145
         TabIndex        =   127
         Top             =   630
         Width           =   255
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0FFD0&
         Caption         =   "Vencimientos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   210
         TabIndex        =   43
         Top             =   4200
         Width           =   3315
         Begin VB.TextBox IVENCI 
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
            Index           =   5
            Left            =   5715
            TabIndex        =   62
            Text            =   " "
            Top             =   1155
            Width           =   1485
         End
         Begin VB.TextBox IVENCI 
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
            Index           =   4
            Left            =   5715
            TabIndex        =   59
            Text            =   " "
            Top             =   735
            Width           =   1485
         End
         Begin VB.TextBox IVENCI 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Index           =   3
            Left            =   1620
            TabIndex        =   56
            Text            =   " "
            Top             =   1260
            Width           =   1485
         End
         Begin VB.TextBox IVENCI 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Index           =   2
            Left            =   1620
            TabIndex        =   53
            Text            =   " "
            Top             =   945
            Width           =   1485
         End
         Begin VB.TextBox IVENCI 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Index           =   1
            Left            =   1620
            TabIndex        =   50
            Text            =   " "
            Top             =   630
            Width           =   1485
         End
         Begin VB.TextBox IVENCI 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Index           =   0
            Left            =   1620
            TabIndex        =   47
            Top             =   315
            Width           =   1485
         End
         Begin VB.CommandButton BOTFEVEN 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   5
            Left            =   5460
            TabIndex        =   61
            Top             =   1155
            Width           =   180
         End
         Begin VB.TextBox FEVEN 
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
            Index           =   5
            Left            =   4515
            TabIndex        =   60
            Text            =   " "
            Top             =   1155
            Width           =   1065
         End
         Begin VB.CommandButton BOTFEVEN 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   4
            Left            =   5460
            TabIndex        =   58
            Top             =   735
            Width           =   180
         End
         Begin VB.TextBox FEVEN 
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
            Left            =   4515
            TabIndex        =   57
            Text            =   " "
            Top             =   735
            Width           =   1065
         End
         Begin VB.CommandButton BOTFEVEN 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Index           =   3
            Left            =   1365
            TabIndex        =   55
            Top             =   1260
            Width           =   180
         End
         Begin VB.TextBox FEVEN 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Index           =   3
            Left            =   420
            TabIndex        =   54
            Text            =   " "
            Top             =   1260
            Width           =   1065
         End
         Begin VB.CommandButton BOTFEVEN 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Index           =   2
            Left            =   1365
            TabIndex        =   52
            Top             =   945
            Width           =   180
         End
         Begin VB.TextBox FEVEN 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Index           =   2
            Left            =   420
            TabIndex        =   51
            Text            =   " "
            Top             =   945
            Width           =   1065
         End
         Begin VB.CommandButton BOTFEVEN 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Index           =   1
            Left            =   1365
            TabIndex        =   49
            Top             =   630
            Width           =   180
         End
         Begin VB.TextBox FEVEN 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Index           =   1
            Left            =   420
            TabIndex        =   48
            Text            =   " "
            Top             =   630
            Width           =   1065
         End
         Begin VB.CommandButton BOTFEVEN 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Index           =   0
            Left            =   1365
            TabIndex        =   46
            Top             =   315
            Width           =   180
         End
         Begin VB.TextBox FEVEN 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Index           =   0
            Left            =   420
            TabIndex        =   45
            Top             =   315
            Width           =   960
         End
         Begin VB.Label LEFE 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "F:"
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
            Index           =   5
            Left            =   2835
            TabIndex        =   68
            Top             =   1200
            Width           =   1590
         End
         Begin VB.Label LEFE 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "F:"
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
            Index           =   4
            Left            =   2835
            TabIndex        =   67
            Top             =   780
            Width           =   1590
         End
         Begin VB.Label LEFE 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "F:"
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
            Index           =   3
            Left            =   -1260
            TabIndex        =   66
            Top             =   1260
            Width           =   1590
         End
         Begin VB.Label LEFE 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "F:"
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
            Index           =   2
            Left            =   -1260
            TabIndex        =   65
            Top             =   945
            Width           =   1590
         End
         Begin VB.Label LEFE 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "F:"
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
            Index           =   1
            Left            =   -1260
            TabIndex        =   63
            Top             =   630
            Width           =   1590
         End
         Begin VB.Label LEFE 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "F:"
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
            Index           =   0
            Left            =   -1260
            TabIndex        =   44
            Top             =   315
            Width           =   1590
         End
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00E0FFD0&
         Caption         =   "Pago Contado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   3675
         TabIndex        =   116
         Top             =   4200
         Width           =   3000
         Begin VB.CommandButton Command14 
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
            Left            =   2625
            TabIndex        =   120
            Top             =   280
            Width           =   180
         End
         Begin VB.TextBox Text14 
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
            Left            =   210
            TabIndex        =   119
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1200
            Width           =   1275
         End
         Begin VB.TextBox Text15 
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
            Left            =   1680
            TabIndex        =   118
            TabStop         =   0   'False
            Top             =   1200
            Width           =   960
         End
         Begin VB.CommandButton Command13 
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
            Left            =   2625
            TabIndex        =   117
            Top             =   1200
            Width           =   180
         End
         Begin VB.Label Label21 
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
            Height          =   280
            Left            =   210
            TabIndex        =   125
            Top             =   620
            Width           =   2610
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Medio de Pago:"
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
            Left            =   420
            TabIndex        =   124
            Top             =   345
            Width           =   1485
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Número:"
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
            Left            =   -630
            TabIndex        =   123
            Top             =   940
            Width           =   1485
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha:"
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
            Left            =   735
            TabIndex        =   122
            Top             =   940
            Width           =   1485
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   1995
            TabIndex        =   121
            Top             =   280
            Width           =   645
         End
      End
      Begin VB.CommandButton Command12 
         Caption         =   "O/C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   10920
         Picture         =   "FORFACOM.frx":0000
         TabIndex        =   112
         ToolTipText     =   "Visualizar Ordenes de Compra"
         Top             =   3780
         Width           =   855
      End
      Begin VB.TextBox Text13 
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
         Left            =   1365
         TabIndex        =   111
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1155
         Width           =   3480
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00E0FFD0&
         Caption         =   "Percepción Ing.Brutos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Left            =   6825
         TabIndex        =   79
         Top             =   4725
         Width           =   3945
         Begin VB.OptionButton Option2 
            BackColor       =   &H00E0FFD0&
            Caption         =   "Option2"
            Height          =   255
            Left            =   240
            TabIndex        =   126
            ToolTipText     =   "Doble Click para Multiples Jurisdicciones"
            Top             =   720
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox Text12 
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
            Left            =   2520
            TabIndex        =   105
            Text            =   " "
            ToolTipText     =   "Ingrese lImporte"
            Top             =   315
            Width           =   1275
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
            Left            =   630
            TabIndex        =   101
            Text            =   " "
            ToolTipText     =   "IngreselProvincia o Jurisdicción"
            Top             =   315
            Width           =   330
         End
         Begin VB.CommandButton Command10 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   945
            TabIndex        =   103
            Top             =   315
            Width           =   180
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Importe:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1530
            TabIndex        =   84
            Top             =   375
            Width           =   960
         End
         Begin VB.Label Label18 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            TabIndex        =   83
            Top             =   675
            Width           =   3150
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Jur:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   0
            TabIndex        =   82
            Top             =   375
            Width           =   540
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00E0FFD0&
         Caption         =   "Percepciones AFIP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   6825
         TabIndex        =   78
         Top             =   3925
         Width           =   3945
         Begin VB.TextBox Text8 
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
            Left            =   2520
            TabIndex        =   99
            Top             =   315
            Width           =   1275
         End
         Begin VB.TextBox Text7 
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
            Left            =   630
            TabIndex        =   97
            Top             =   315
            Width           =   1275
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Gan:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1950
            TabIndex        =   81
            Top             =   420
            Width           =   540
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "IVA:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   0
            TabIndex        =   80
            Top             =   420
            Width           =   540
         End
      End
      Begin VB.CommandButton Command8 
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
         Height          =   720
         Left            =   10920
         Picture         =   "FORFACOM.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Anular, Consultar, Imputación de Créditos, Tablas"
         Top             =   5985
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8085
         TabIndex        =   72
         Top             =   195
         Width           =   180
      End
      Begin VB.TextBox Text6 
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
         Left            =   7140
         TabIndex        =   1
         Text            =   " "
         ToolTipText     =   "Ingrese la fecha REAL del Comprobante Recibido"
         Top             =   195
         Width           =   960
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0FFD0&
         Caption         =   "Totales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2440
         Left            =   6825
         TabIndex        =   42
         Top             =   1470
         Width           =   3945
         Begin VB.TextBox IMPOIVA 
            BackColor       =   &H0080FF80&
            BorderStyle     =   0  'None
            Height          =   300
            Index           =   6
            Left            =   2520
            TabIndex        =   115
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   2100
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.TextBox IMPOGRA 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   930
            TabIndex        =   114
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   2025
            Width           =   1485
         End
         Begin VB.TextBox IMPOIVA 
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
            Index           =   5
            Left            =   2500
            TabIndex        =   95
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   1740
            Width           =   1275
         End
         Begin VB.TextBox IMPOIVA 
            BackColor       =   &H0080FF80&
            BorderStyle     =   0  'None
            Height          =   300
            Index           =   4
            Left            =   2520
            TabIndex        =   107
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   1455
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H00E0FFD0&
            BorderStyle     =   0  'None
            Caption         =   "Frame11"
            Height          =   1950
            Left            =   40
            TabIndex        =   98
            Top             =   360
            Width           =   850
            Begin VB.Label TASIVA 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Imp.Int"
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
               Left            =   0
               TabIndex        =   113
               Top             =   1680
               Width           =   765
            End
            Begin VB.Label TASIVA 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Import"
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
               Index           =   5
               Left            =   0
               TabIndex        =   109
               Top             =   1425
               Width           =   770
            End
            Begin VB.Label TASIVA 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Exento"
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
               Index           =   4
               Left            =   0
               TabIndex        =   108
               Top             =   1140
               Width           =   800
            End
            Begin VB.Label TASIVA 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "21.0%"
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
               Index           =   0
               Left            =   105
               TabIndex        =   106
               Top             =   0
               Width           =   750
            End
            Begin VB.Label TASIVA 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "27.0%"
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
               Index           =   1
               Left            =   80
               TabIndex        =   104
               Top             =   285
               Width           =   750
            End
            Begin VB.Label TASIVA 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "10.5%"
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
               Index           =   2
               Left            =   80
               TabIndex        =   102
               Top             =   570
               Width           =   750
            End
            Begin VB.Label TASIVA 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " "
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
               Index           =   3
               Left            =   80
               TabIndex        =   100
               Top             =   855
               Width           =   750
            End
         End
         Begin VB.TextBox IMPOGRA 
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
            Index           =   5
            Left            =   930
            TabIndex        =   90
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   1740
            Width           =   1485
         End
         Begin VB.TextBox IMPOGRA 
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
            Index           =   4
            Left            =   930
            TabIndex        =   89
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   1455
            Width           =   1485
         End
         Begin VB.TextBox IMPOIVA 
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
            Index           =   3
            Left            =   2500
            TabIndex        =   94
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   1170
            Width           =   1275
         End
         Begin VB.TextBox IMPOGRA 
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
            Index           =   3
            Left            =   930
            TabIndex        =   88
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   1170
            Width           =   1485
         End
         Begin VB.TextBox IMPOIVA 
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
            Left            =   2500
            TabIndex        =   93
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   885
            Width           =   1275
         End
         Begin VB.TextBox IMPOGRA 
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
            Left            =   930
            TabIndex        =   87
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   885
            Width           =   1485
         End
         Begin VB.TextBox IMPOIVA 
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
            Left            =   2500
            TabIndex        =   92
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   600
            Width           =   1275
         End
         Begin VB.TextBox IMPOGRA 
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
            Left            =   930
            TabIndex        =   86
            ToolTipText     =   "  "
            Top             =   600
            Width           =   1485
         End
         Begin VB.TextBox IMPOIVA 
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
            Left            =   2500
            TabIndex        =   91
            ToolTipText     =   "  "
            Top             =   315
            Width           =   1275
         End
         Begin VB.Label ITOGRA 
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
            Height          =   300
            Left            =   930
            TabIndex        =   96
            Top             =   315
            Width           =   1485
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0FFD0&
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   435
         Left            =   6480
         TabIndex        =   38
         Top             =   630
         Width           =   4455
         Begin VB.OptionButton Option1 
            BackColor       =   &H00E0FFD0&
            Caption         =   "CO"
            Height          =   240
            Index           =   5
            Left            =   165
            TabIndex        =   85
            ToolTipText     =   "Comprobante Genérico (Extracto Bancario - Póliza de Seguro - Liquidación)"
            Top             =   105
            Width           =   705
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00E0FFD0&
            Caption         =   "RH"
            Height          =   240
            Index           =   4
            Left            =   855
            TabIndex        =   71
            ToolTipText     =   "Recibo de Honorarios"
            Top             =   105
            Width           =   700
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00E0FFD0&
            Caption         =   "TF"
            Height          =   240
            Index           =   3
            Left            =   1560
            TabIndex        =   70
            ToolTipText     =   "Ticket - Factura"
            Top             =   105
            Width           =   700
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00E0FFD0&
            Caption         =   "NC"
            Height          =   240
            Index           =   2
            Left            =   3660
            TabIndex        =   41
            ToolTipText     =   "Nota de Crédito"
            Top             =   105
            Width           =   700
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00E0FFD0&
            Caption         =   "ND"
            Height          =   240
            Index           =   1
            Left            =   2955
            TabIndex        =   40
            ToolTipText     =   "Nota de Débito"
            Top             =   105
            Width           =   700
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00E0FFD0&
            Caption         =   "FC"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   2265
            TabIndex        =   39
            ToolTipText     =   "Factura de Compra"
            Top             =   105
            Value           =   -1  'True
            Width           =   700
         End
      End
      Begin VB.CommandButton Command6 
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
         Height          =   720
         Left            =   10920
         Picture         =   "FORFACOM.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Anular, Consultar, Imputación de Créditos, Tablas"
         Top             =   6740
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E0E0E0&
         Height          =   1695
         Left            =   210
         ScaleHeight     =   1635
         ScaleWidth      =   6405
         TabIndex        =   30
         Top             =   1585
         Width           =   6465
         Begin VB.Frame Frame4 
            BackColor       =   &H00C0C000&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   330
            Left            =   0
            TabIndex        =   31
            Top             =   0
            Width           =   10515
            Begin VB.Line Line23 
               X1              =   8505
               X2              =   8505
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "  Cod.    Referencia - Cuenta                                       Importe"
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
               Left            =   0
               TabIndex        =   32
               Top             =   50
               Width           =   10410
            End
            Begin VB.Line Line20 
               X1              =   4725
               X2              =   4725
               Y1              =   0
               Y2              =   1155
            End
            Begin VB.Line Line6 
               X1              =   630
               X2              =   630
               Y1              =   0
               Y2              =   1155
            End
            Begin VB.Line Line5 
               X1              =   6930
               X2              =   6930
               Y1              =   0
               Y2              =   1050
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0F0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   5580
            Left            =   -315
            TabIndex        =   33
            Top             =   0
            Width           =   10800
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
               Height          =   1350
               IntegralHeight  =   0   'False
               Left            =   280
               TabIndex        =   34
               Top             =   315
               Width           =   6465
            End
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00E0E0E0&
         Height          =   2230
         Left            =   210
         ScaleHeight     =   2175
         ScaleWidth      =   10470
         TabIndex        =   20
         Top             =   5985
         Width           =   10530
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0C000&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   5630
            Left            =   0
            TabIndex        =   21
            Top             =   315
            Width           =   10515
            Begin VB.ListBox List2 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1410
               Left            =   -30
               TabIndex        =   22
               Top             =   0
               Width           =   10550
            End
            Begin VB.ListBox List3 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1410
               Left            =   -30
               TabIndex        =   69
               Top             =   0
               Width           =   10450
            End
            Begin VB.Label Label10 
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
               Height          =   225
               Left            =   8295
               TabIndex        =   29
               Top             =   1535
               Width           =   1695
            End
            Begin VB.Label Label11 
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
               Height          =   225
               Left            =   0
               TabIndex        =   28
               Top             =   0
               Width           =   1695
            End
            Begin VB.Line Line4 
               X1              =   8400
               X2              =   8400
               Y1              =   0
               Y2              =   2415
            End
            Begin VB.Label Label8 
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
               Height          =   225
               Left            =   6510
               TabIndex        =   25
               Top             =   1535
               Width           =   1695
            End
            Begin VB.Label Label4 
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
               Height          =   225
               Left            =   3570
               TabIndex        =   24
               Top             =   2130
               Width           =   1695
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   " Total General"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   0
               TabIndex        =   23
               Top             =   1535
               Width           =   1695
            End
            Begin VB.Line Line14 
               BorderWidth     =   2
               X1              =   0
               X2              =   10500
               Y1              =   1420
               Y2              =   1420
            End
            Begin VB.Line Line15 
               X1              =   0
               X2              =   10500
               Y1              =   1840
               Y2              =   1840
            End
            Begin VB.Line Line16 
               X1              =   2520
               X2              =   2520
               Y1              =   1365
               Y2              =   2415
            End
            Begin VB.Line Line17 
               X1              =   3735
               X2              =   3735
               Y1              =   1365
               Y2              =   2415
            End
            Begin VB.Line Line18 
               X1              =   5460
               X2              =   5460
               Y1              =   1365
               Y2              =   2415
            End
            Begin VB.Line Line19 
               X1              =   6720
               X2              =   6720
               Y1              =   1365
               Y2              =   2415
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0C000&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   330
            Left            =   0
            TabIndex        =   26
            Top             =   0
            Width           =   10515
            Begin VB.Line Line1 
               X1              =   8400
               X2              =   8400
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line2 
               X1              =   6720
               X2              =   6720
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   " Documento  Número         Emision                Importe       Vencim.                  Saldo                Acredit.       "
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
               Left            =   0
               TabIndex        =   27
               Top             =   50
               Width           =   10410
            End
            Begin VB.Line Line9 
               X1              =   2520
               X2              =   2520
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line10 
               X1              =   3730
               X2              =   3730
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line11 
               X1              =   5460
               X2              =   5460
               Y1              =   0
               Y2              =   1050
            End
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Imprimir"
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
         Left            =   10920
         Picture         =   "FORFACOM.frx":0896
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Imprime Formulario de Carga - Vuelco de Pantalla"
         Top             =   7490
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Caption         =   "New"
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
         Left            =   10920
         Picture         =   "FORFACOM.frx":0F00
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Apertura de Nueva Cuenta de Proveedor"
         Top             =   1575
         Width           =   855
      End
      Begin VB.CommandButton Command4 
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
         Height          =   720
         Left            =   10920
         Picture         =   "FORFACOM.frx":120A
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Resumen de Cuenta Corriente Proveedores"
         Top             =   2310
         Width           =   855
      End
      Begin VB.CommandButton Command3 
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
         Height          =   720
         Left            =   10920
         Picture         =   "FORFACOM.frx":1514
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Notas de Proveedores"
         Top             =   3045
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
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
         Left            =   10920
         Picture         =   "FORFACOM.frx":181E
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Abandona la Aplicación"
         Top             =   210
         Width           =   855
      End
      Begin VB.CommandButton BOTNEXT 
         Caption         =   "Grabar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1100
         Left            =   10920
         Picture         =   "FORFACOM.frx":1B28
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Grabar - Registrar Documento en la Base de Datos"
         Top             =   4725
         Width           =   855
      End
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7980
         TabIndex        =   3
         ToolTipText     =   "Ingrese Número de Comprobante"
         Top             =   1080
         Width           =   2805
      End
      Begin VB.TextBox Text9 
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
         Left            =   9555
         TabIndex        =   2
         Text            =   " "
         ToolTipText     =   "Ingrese la Fecha Correspondiente al Mes de Contabilización del I.V.A."
         Top             =   210
         Width           =   960
      End
      Begin VB.CommandButton Command9 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10500
         TabIndex        =   5
         Top             =   210
         Width           =   180
      End
      Begin VB.TextBox Text5 
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
         Left            =   2205
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   " "
         Top             =   840
         Width           =   2640
      End
      Begin VB.TextBox Text4 
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
         Left            =   1365
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   " "
         Top             =   840
         Width           =   750
      End
      Begin VB.TextBox Text3 
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
         Left            =   1365
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   " "
         Top             =   525
         Width           =   3480
      End
      Begin VB.TextBox Text2 
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
         Left            =   1365
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   " "
         Top             =   210
         Width           =   3480
      End
      Begin VB.CommandButton Command1 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1050
         TabIndex        =   4
         Top             =   840
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
         ToolTipText     =   "Ingrese Número de Proveedor"
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Carga por Detalle"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5400
         TabIndex        =   136
         Top             =   660
         Width           =   855
      End
      Begin VB.Shape Shape1 
         Height          =   600
         Left            =   5145
         Top             =   -15
         Width           =   750
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00808080&
         X1              =   210
         X2              =   6705
         Y1              =   3280
         Y2              =   3280
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00808080&
         X1              =   6675
         X2              =   6675
         Y1              =   1575
         Y2              =   3280
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00808080&
         X1              =   10740
         X2              =   10740
         Y1              =   5985
         Y2              =   8190
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackColor       =   &H00E0FFD0&
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5250
         TabIndex        =   110
         ToolTipText     =   "Click Para Cambiar Categoria IVA"
         Top             =   0
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Emision:"
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
         Left            =   6090
         TabIndex        =   73
         Top             =   105
         Width           =   960
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3780
         TabIndex        =   36
         Top             =   6500
         Width           =   1695
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Valores:"
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
         Left            =   1470
         TabIndex        =   35
         Top             =   6500
         Width           =   2220
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Factura Nro:"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   4935
         TabIndex        =   15
         Top             =   1120
         Width           =   2955
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Contable:"
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
         Left            =   8505
         TabIndex        =   14
         Top             =   120
         Width           =   960
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
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
         Left            =   315
         TabIndex        =   8
         Top             =   525
         Width           =   960
      End
   End
   Begin VB.Label ONEGRA 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   840
      TabIndex        =   75
      Top             =   0
      Width           =   1515
   End
   Begin VB.Label IMPIVA 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   4
      Left            =   2415
      TabIndex        =   77
      Top             =   0
      Width           =   1230
   End
   Begin VB.Label IMPIVA 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   2415
      TabIndex        =   76
      Top             =   0
      Width           =   1230
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1050
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Pos.IVA:"
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
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   960
   End
End
Attribute VB_Name = "FORFACOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LIST2ANT$
Dim TOBRUFAC#        ' total bruto factura
'
Private Sub BOTFEVEN_Click(Index As Integer)
   Call SELECHO("SELFECHA")
   FEVEN(Index).TEXT = VALACT1$("SELFECHA")
   FEVEN(Index).Refresh
End Sub

Private Sub BOTNEXT_Click()
    '
    Call GRAFACOM
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    MODIFIX% = 0
    '
End Sub
'



Private Sub Check1_LOSTFOCUS()
  '\\\OT-05-0527-WAR-04/08/05///
  If Option1(1) Or Option1(2) Or Option1(4) Or Option1(5) Then
    Check1 = 0
  End If
  '\\\OT-05-0527-WAR-FIN///
End Sub
Private Sub Check1_CLICK()
  '\\\OT-05-0527-WAR-04/08/05///
'  If Val(Text1) > 0 Then
'    Check1.Value = (1 + Check1.Value) Mod 2
'    Exit Sub
'  End If
  '
If Option1(1) Or Option1(2) Or Option1(4) Or Option1(5) Then
    Check1 = 0
  Else
    'Check1 = 1
  End If
  '\\\OT-05-0527-WAR-FIN///
End Sub
Private Sub Command1_Click()
    '
    If MODIFIX% > 0 Then
       If COMALTER%("Abandona sin Guardar los Cambios ?\\No, Continuar\Si, Cancelar") <> 2 Then
          Exit Sub
       End If
    End If
    '
    Call SELECHO("ACREDOR")
    NCXI% = Val(TRIM$(VALACT1$("ACREDOR")))
    If NCXI% > 0 Then
        NC% = NCXI%
        Text1.TEXT = TRIM$(Str$(NC%))
        Text2.TEXT = RASOCLI$((-1) * NC%)
        Text3.TEXT = DOMICLI$((-1) * NC%)
        Text4.TEXT = CPOSCLI$((-1) * NC%)
        Text5.TEXT = LOCACLI$((-1) * NC%)
        Text13.TEXT = CUITCLI$((-1) * NC%)
        If PIVACLI%((-1) * NC%) = 1 Then
            Label20 = "A"
          Else
            Label20 = "C"
        End If
        If Left$(Text2.TEXT, 6) = "VARIOS" Then
            Text13.SetFocus
          Else
            Text6.SetFocus
        End If
        'Call SIGUEFACO1
    End If
    '
End Sub

Private Sub Command10_Click()
    Call SELECHO("PROVINC")
    DPX1$ = VALACT1$("PROVINC")
    If DPX1$ <> "" Then
        DP1$ = DPX1$
        Text10.TEXT = TRIM$(DP1$)
        Label18.Caption = ECOARCH$("PROVINC", DP1$)
    End If
End Sub

 Sub Command12_Click()
   Command12.Enabled = False
   NPROI% = Val(Text1)
   If NPROI% < 1 Then
       Call MENSERR(24, "Ingrese Número de Proveedor")
     Else
       RFF$ = RECFILT$("OCOMENCA", 4, MKI$(NPROI%))
       If Len(RFF$) < 4 Then
            Call MENSERR(24, "No Hay Ordenes de Compra Registradas\para este Proveedor")
         Else
            JJ& = 0
            For UK& = 1 To Len(RFF$) Step 4
              U& = CVS(Mid$(RFF$, UK&, 4))
              X$ = REGLEIDO$("OCOMENCA", U&)
              Y$ = Left$(X$, 48) + MKS$(U&)
              JJ& = JJ& + 1
              Call GRAREG("!INDIOCOM", Y$, JJ&)
            Next UK&
            Call SETULTREG("!INDIOCOM", JJ&)
            '
3           Call FRESHECHO("!INDIOCOM")
            Call SELECHO("!INDIOCOM")
            NPX& = Val(VALACT1$("!INDIOCOM"))
            If NPX& < 1 Then GoTo 99
            '
            TIDOCUM$ = "Nota de Pedido"
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
            Screen.MousePointer = 1
            Call COMUNI("Documento no Encontrado")
            GoTo 3
            '
5           Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
            GoTo 3
            '
       End If
   End If
99 Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   Call SELECHO("SELFECHA")
   Text15 = VALACT1$("SELFECHA")
   Text15.Refresh
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   Call SELECHO("MEPAGO")
   Label25 = ""
   If VALACT1$("MEPAGO") <> "" Then
     TPV% = Val(VALACT1$("MEPAGO"))
     If CUEMEPAG%(TPV%) = CUEICARCH% Then
         Call MENSERR(24, "Cartera Valores no Disponible\para Compras de Contado.\   \Debe Cancelar por Orden de Pago.")
         Label25 = ""
       Else
         Label25 = VALACT1$("MEPAGO")
     End If
   End If
   Command14.Enabled = True
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Image1_Click()

End Sub

Private Sub IMPOGRA_Change(Index As Integer)
   '
   TASIX$ = TRIM$(REPLACAR$(TASIVA(Index).Caption, "%", ""))
   TASI = Val(TASIX$)
   IGRA# = Val(IMPOGRA(Index).TEXT)
   If Index < 4 Then
     IMPOIVA(Index).TEXT = CSTRING$(MKD$(TASI * IGRA# / 100), 4, 9, 2, 2)
     IMPOIVA(Index).Refresh
   End If
   '
   TOGRAVA# = 0
   For KKI% = 1 To 6
     IGRA# = Val(IMPOGRA(KKI%).TEXT)
     TOGRAVA# = TOGRAVA# + IGRA#
   Next KKI%
   '
   ITOGRA.Caption = CSTRING$(MKD$(TOTIMPU# - TOGRAVA#), 4, 11, 2, 2)
   ITOGRA.Refresh
   Call CALTOFAC
   '
End Sub

Private Sub IMPOGRA_GotFocus(Index As Integer)
    IMPOGRA(Index).SelStart = 0
    IMPOGRA(Index).SelLength = Len(IMPOGRA(Index).TEXT)
End Sub

Private Sub IMPOGRA_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii% = 13 Then
      If Index < 5 Then
         IMPOGRA(Index + 1).SetFocus
        Else
         IMPOIVA(0).SetFocus
      End If
   End If
End Sub

Private Sub IMPOGRA_LostFocus(Index As Integer)
   IGRA# = Abs(CDbl(Val(IMPOGRA(Index).TEXT)))
   TASIX$ = TRIM$(TASIVA(Index).Caption)
   If TASIX$ = "" Then IGRA# = 0
   If LETRAFA$ <> "A" Then
      If Index <> 4 Then
        IMPOGRA(4).TEXT = CSTRING$(MKD$(Val(Label31)), 4, 11, 2, 2)
        IGRA# = 0
      End If
   End If
   IMPOGRA(Index).TEXT = CSTRING$(MKD$(IGRA#), 4, 11, 2, 2)
   IMPOGRA(Index).Refresh
End Sub

Private Sub IMPOIVA_Change(Index As Integer)
   Call CALTOFAC
End Sub

Private Sub IMPOIVA_GotFocus(Index As Integer)
    If Index = 4 Then
       IMPOIVA(5).SetFocus
       Exit Sub
    End If
    IMPOIVA(Index).SelStart = 0
    IMPOIVA(Index).SelLength = Len(IMPOIVA(Index).TEXT)
End Sub

Private Sub IMPOIVA_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii% = 13 Then
      If Index < 5 Then
         NECA% = 1: If Index = 3 Then NECA% = 2
         IMPOIVA(Index + NECA%).SetFocus
        Else
         Text7.SetFocus
      End If
   End If
End Sub

Private Sub IMPOIVA_LostFocus(Index As Integer)
   IGRA# = Abs(CDbl(Val(IMPOIVA(Index).TEXT)))
   TASIX$ = TRIM$(TASIVA(Index).Caption)
   If TASIX$ = "" Then IGRA# = 0
   If LETRAFA$ <> "A" Then
     IGRA# = 0
   End If
   IMPOIVA(Index).TEXT = CSTRING$(MKD$(IGRA#), 4, 9, 2, 2)
   IMPOIVA(Index).Refresh
End Sub

Private Sub Label20_Change()
   LETRAFA$ = Label20.Caption
   If TIDOCOM$ <> "CO" Then
      LETRAFA$ = Label20.Caption
      TTXX$ = Text11.TEXT
      If TRIM$(Text11.TEXT) <> "" Then
        Call REPLA(TTXX$, LETRAFA$, 4, 1)
        Text11.TEXT = TTXX$
        Text11.Refresh
      End If
   End If
   Call SIGUEFACO3
End Sub

Private Sub Label20_Click()
   ILETRA% = Asc(Label20.Caption + "A") + 1
   If ILETRA% > 67 Then ILETRA% = 65
   Label20.Caption = Chr$(ILETRA%)
End Sub

Private Sub Label25_Change()
   Label21 = ECOARCH$("MEPAGO", MKI$(Val(Label25)))
   If TRIM$(Label25) = "" Then
       Text14 = ""
       Text15 = ""
       'Call CALTOFAC
     Else
       HOII% = HOY(HOS$)
       Text15 = HOS$
       On Error Resume Next
       Text14.SetFocus
       On Error GoTo 0
   End If
End Sub

Private Sub Label25_DblClick()
   Call SELECHO("MEPAGO")
   If VALACT1$("MEPAGO") <> "" Then
     TPV% = Val(VALACT1$("MEPAGO"))
     If CUEMEPAG%(TPV%) = CUEICARCH% Then
         Call MENSERR(24, "Cartera Valores no Disponible\para Compras de Contado.\   \Debe Cancelar por Orden de Pago.")
         Label25 = ""
       Else
         Label25 = VALACT1$("MEPAGO")
     End If
   End If
End Sub

Private Sub Label28_CHANGE()
  '
  Call Text17_Change
  '
End Sub

Private Sub Label31_CHANGE()
  '
  IMPONET1# = Val(Label31)
  LETRAFA$ = Label20
  If LETRAFA$ = "A" Then
     ITOGRA = FORMATNUM(IMPONET1#, "F10.2")
   Else
     IMPOGRA(4) = FORMATNUM(IMPONET1#, "F10.2")
  End If
  TOTIMPU# = IMPONET1#
  '
End Sub

Private Sub Option2_DblClick()
   '
   VTB% = VENTABU%("INGREPIB")
   If VTB% >= 0 Then
     SUPIBRU# = 0
     For U& = 1 To ULTREG&("!PROVINB")
       XX$ = REGLEIDO$("!PROVINB", U&)
       SUPIBRU# = SUPIBRU# + CVD(Mid$(XX$, 57, 8))
     Next U&
     '
     If SUPIBRU# >= 0.005 Then
        Text10 = "*"
        Label18 = "Jurisdicciones Varias"
        Text12 = FORMATNUM$(SUPIBRU#, "F9.2")
        Text10.Enabled = False
        Text12.Enabled = False
        Command10.Enabled = False
      Else
        If Text10 = "*" Then
          Text10 = ""
          Label18 = ""
          Text12 = ""
        End If
        Text10.Enabled = True
        Text12.Enabled = True
        Command10.Enabled = True
     End If
   End If
   '
End Sub

Private Sub Text1_LostFocus()
   CLITE$ = TRIM$(Text1.TEXT)
   If CLITE$ <> "" Then
     If Val(CLITE$) < 1 Or Val(CLITE$) > 32767 Then
       Text1.TEXT = Space$(5)
       Exit Sub
     End If
     If ECOARCH$("ACREDOR", MKI$(Val(CLITE$))) = "" Then
       Text1.TEXT = Space$(5)
       Exit Sub
     End If
   End If
End Sub

Private Sub Text10_LostFocus()
   If TRIM$(Text10.TEXT) <> "" Then
      If ECOARCH$("PROVINC", TRIM$(Text10.TEXT)) = "" Then
         Text10.TEXT = " "
         On Error Resume Next
         Text10.SetFocus
         On Error GoTo 0
      End If
   End If
End Sub
'
Private Sub TEXT10_GotFocus()
    Text12.SelStart = 0
    Text12.SelLength = Len(Text10.TEXT)
End Sub

Private Sub Text12_Change()
   Call CALTOFAC
End Sub

'
Private Sub TEXT12_GotFocus()
    Text12.SelStart = 0
    Text12.SelLength = Len(Text12.TEXT)
End Sub

Private Sub TEXT12_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
      FEVEN(0).SetFocus
   End If
End Sub

Private Sub Text12_LostFocus()
   IGRA# = Abs(CDbl(Val(Text12.TEXT)))
   If LETRAFA$ <> "A" Then
     IGRA# = 0
   End If
   DPX1$ = TRIM$(Text10.TEXT)
   If DPX1$ = "" Or ECOARCH$("PROVINC", DPX1$) = "" Then
     Call MENSERR(24, "Falta Jurisdicción Percepción Ingresos Brutos")
     IGRA# = 0
   End If
   Text12.TEXT = CSTRING$(MKD$(IGRA#), 4, 9, 2, 2)
   Text12.Refresh
End Sub

Private Sub Text13_GOTFOCUS()
    Text13.SelStart = 0
    Text13.SelLength = Len(Text13.TEXT)
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Val(Text1.TEXT) > 0 Then
          If Val(Text1.TEXT) <= 32767 Then
             NC% = Val(Text1.TEXT)
             If ECOARCH$("ACREDOR", MKI$(NC%)) <> "" Then
                On Error Resume Next
                Text2.SetFocus
                On Error GoTo 0
                Exit Sub
             End If
          End If
       End If
       '
       Call MENSERR(24, "Número de Cuenta no Válido")
       On Error Resume Next
       Text1.SetFocus
       On Error GoTo 0
       '
    End If
    '
End Sub
Private Sub Text17_Change()
  '
  TN1# = Val(Label28)
  BN1 = COEFDESCU(TRIM$(Text17))
  IMTO1# = (Int(100 * TN1# * (1 - BN1 / 100) + 0.5)) / 100
  BN2# = TN1# - IMTO1#
  '
  Label29 = FORMATNUM(BN2#, "F11.2")
  Label31 = FORMATNUM(IMTO1#, "F11.2")
  '
End Sub

Private Sub Text17_GotFocus()
  '\\\OT-05-0527-WAR-04/08/05///
  If Check1 Then
    IVENCI(0).SetFocus
  End If
  '\\\OT-05-0527-WAR-FIN///
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Val(Text1.TEXT) > 0 Then
          If Val(Text1.TEXT) <= 32767 Then
             NC% = Val(Text1.TEXT)
             If ECOARCH$("ACREDOR", MKI$(NC%)) <> "" Then
                On Error Resume Next
                Text6.SetFocus
                On Error GoTo 0
                Exit Sub
             End If
          End If
       End If
       '
       Call MENSERR(24, "Número de Cuenta no Válido")
       On Error Resume Next
       Text1.SetFocus
       On Error GoTo 0
       '
    End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
      Text9.SetFocus
   End If
End Sub

Private Sub Text7_Change()
   Call CALTOFAC
End Sub

'
Private Sub Text7_GotFocus()
    Text7.SelStart = 0
    Text7.SelLength = Len(Text7.TEXT)
End Sub

Private Sub TEXT7_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
      Text8.SetFocus
   End If
End Sub

Private Sub TEXT7_LostFocus()
   IGRA# = Abs(CDbl(Val(Text7.TEXT)))
   If LETRAFA$ <> "A" Then
     IGRA# = 0
   End If
   Text7.TEXT = CSTRING$(MKD$(IGRA#), 4, 9, 2, 2)
   Text7.Refresh
End Sub

Private Sub Text8_Change()
   Call CALTOFAC
End Sub

'
Private Sub TEXT8_GotFocus()
    Text8.SelStart = 0
    Text8.SelLength = Len(Text8.TEXT)
End Sub

Private Sub TEXT8_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
     On Error Resume Next
     Text10.SetFocus
     On Error GoTo 0
   End If
End Sub

Private Sub TEXT8_LostFocus()
   IGRA# = Abs(CDbl(Val(Text8.TEXT)))
   If LETRAFA$ <> "A" Then
     IGRA# = 0
   End If
   Text8.TEXT = CSTRING$(MKD$(IGRA#), 4, 9, 2, 2)
   Text8.Refresh
End Sub
'
Private Sub ITOGRA_Change()
   '
   TASIX$ = TRIM$(REPLACAR$(TASIVA(0).Caption, "%", ""))
   TASI = Val(TASIX$)
   IGRA# = Val(ITOGRA.Caption)
   IMPOIVA(0).TEXT = CSTRING$(MKD$(TASI * IGRA# / 100), 4, 9, 2, 2)
   IMPOIVA(0).Refresh
   Call CALTOFAC
   '
End Sub
'
Private Sub Text10_DblClick()
    Call SELECHO("PROVINC")
    DPX1$ = VALACT1$("PROVINC")
    DP1$ = DPX1$
    Text10.TEXT = AJUSTI$(DP1$, 1)
    Label18.Caption = ECOARCH$("PROVINC", TRIM$(DP1$))
End Sub

Private Sub Text10_Change()
    DPX1$ = TRIM$(Text10.TEXT)
    DP1$ = AJUSTI$(DPX1$, 1)
    Label18.Caption = ECOARCH$("PROVINC", DP1$)
End Sub
'
Private Sub TEXT10_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
     On Error Resume Next
     Text12.SetFocus
     On Error GoTo 0
   End If
End Sub


Private Sub Command11_Click()
    Command11.Enabled = False
    Call CIERRARCH("*.*")
    Call CONECRUN("AMAPRO/NEW", "Alta de Nueva Cuenta de Proveedor")
    Call FRESHECHO("PROVED1")
    Command11.Enabled = True
End Sub

Private Sub Command2_Click()
    If MODIFIX% > 0 Then
       If COMALTER%("Abandona sin Guardar los Cambios ?\\No, Continuar\Si, Cancelar") <> 2 Then
          Exit Sub
       End If
    End If
    Screen.MousePointer = 11
    If List1.ListCount > 0 Or List2.ListCount > 0 Then
          Call CIERRARCH("*.*")
          Call BLAFACOM
          Screen.MousePointer = 1
          Text1.SetFocus
        Else
          Call FINAL("")
          Screen.MousePointer = 1
    End If
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If List1.ListCount > 0 Or List2.ListCount > 0 Then
        Command2.ToolTipText = "Limpia y Refresca Pantalla"
      Else
        Command2.ToolTipText = "Cierra y Abandona la Aplicación"
   End If
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    NCLI$ = TRIM$(Text1.TEXT)
    Call CONECRUN("*NOTAPRO/" + NCLI$, "Notas de Proveedores")
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    RESCUPRO.Text1.TEXT = Text1.TEXT
    PROLLAMA$ = Name
    'Hide
    RESCUPRO.Show 1
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    PRINTFORM
    Text1.SetFocus
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    FACOCOMP.Show 1
    Command6.Enabled = True
End Sub


Private Sub Command7_Click()
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Text6.TEXT = VDEVV$
    End If
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   OPFACOM.Show 1
   Command8.Enabled = True
End Sub

Private Sub Command9_Click()
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Text9.TEXT = VDEVV$
        FF% = FECHANUM(Text9.TEXT)
    End If
End Sub

Private Sub FEVEN_GotFocus(Index As Integer)
    FEVEN(Index).SelStart = 0
    FEVEN(Index).SelLength = Len(FEVEN(Index).TEXT)
End Sub

Private Sub FEVEN_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
      If TRIM$(FEVEN(Index).TEXT) = "" Then
         If TRIM$(IVENCI(Index).TEXT) = "" Then
            BOTNEXT.SetFocus
            Exit Sub
         End If
      End If
      IVENCI(Index).SetFocus
   End If
End Sub

Private Sub FEVEN_LostFocus(Index As Integer)
   If TRIM(FEVEN(Index).TEXT) <> "" Then
      If FECHANUM(FEVEN(Index).TEXT) < 1 Then
         Call MENSERR(24, "Fecha no Válida")
         FEVEN(Index).SetFocus
      End If
   End If
End Sub

Private Sub Form_Load()
    '
    Call CENTRAFORM(Me)
    '
    Show
    Refresh
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    If InStr(EPAC$, "SABO") > 0 Then Option2.Visible = True
    '
    For IXX% = 0 To 3
      TXXX$ = TRIM$(Str$(IXX%))
      TASA = Val(CONTROL$("", "TASIVA-" + TXXX$))
      If TASA > 0.005 Then
        TASIVA(IXX%).Caption = TRIM$(CSTRING$(MKS$(TASA), 3, 6, 2, 2) + "%")
      End If
    Next IXX%
    '
    TIDOCOM$ = "FC"
    Text1.TEXT = Space$(5)
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.TEXT)
    '
    FF% = HOY(HO$)
    Text9.TEXT = HO$
    Text6.TEXT = HO$
    Text11.TEXT = ""
    '
    Call BLAFACOM
    '
    Refresh
    Text1.SetFocus
    '
End Sub

Private Sub IVENCI_GotFocus(Index As Integer)
    If FECHANUM(FEVEN(Index).TEXT) < 1 Then
       'FEVEN(Index).SetFocus
       'Call MENSERR(24, "Fecha de Vencimiento no Válida")
       Exit Sub
    End If
    '
    IVENCI(Index).SelStart = 0
    IVENCI(Index).SelLength = Len(IVENCI(Index).TEXT)
End Sub

Private Sub IVENCI_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Index < 5 Then
         FEVEN(Index + 1).SetFocus
      End If
   End If
End Sub

Private Sub IVENCI_LOSTFOCUS(Index As Integer)
   SUMAVEN# = Val(FORFACOM.Label10.Caption)
   For KKI% = 0 To Index - 1
      SUMAVEN# = SUMAVEN# + Val(IVENCI(KKI%).TEXT)
   Next KKI%
   IPAVEN# = Val(FORFACOM.IVENCI(Index).TEXT)
   If SUMAVEN# + IPAVEN# < Val(TOBRUFAC#) Then
      If Index < 5 Then
         IVENCI(Index + 1) = CSTRING$(MKD$(TOBRUFAC# - SUMAVEN# - IPAVEN#), 3, 11, 2, 2)
      End If
   End If
   IVENCI(Index).TEXT = CSTRING$(MKD$(IPAVEN#), 4, 11, 2, 2)
End Sub

Private Sub Label10_Change()
   Call CALTOFAC
End Sub

Private Sub List1_DBLClick()
    '
    '\\\OT-05-0527-WAR-04/08/05///
    If Check1 Then
      FOFACOBR.Show 1
    Else
      Call SIGUEFACO2
    End If
    '
End Sub

Private Sub List1_GotFocus()
   MODIFIX% = 1
End Sub

Private Sub List2_Click()
    '
    If VACON% <> 0 Then
        Exit Sub
    End If
    '
    Z$ = List2.TEXT
    LIST2ANT$ = Z$
    TOLISTA& = List2.TopIndex
    ILISTA& = List2.ListIndex
    IACRE$ = TRIM$(Mid$(Z$, 70, 14))
    IPEN$ = " " + Mid$(Z$, 56, 13)
    '
    SUMACRE# = 0: MAXACRE# = TOBRUFAC#
    List2.Visible = False
    VACON% = 1
    For U& = 1 To List2.ListCount
       List2.ListIndex = U& - 1
       SUMACRE# = SUMACRE# + Val(Mid$(List2.TEXT, 70, 14))
    Next U&
    List2.TopIndex = TOLISTA&
    List2.ListIndex = ILISTA&
    List2.Visible = True
    VACON% = 0
    '
    If IACRE$ = "" Then
         ACREDIT# = Val(IPEN$)
         If ACREDIT# > MAXACRE# - SUMACRE# Then
            ACREDIT# = MAXACRE# - SUMACRE#
            If ACREDIT# < 0 Then ACREDIT# = 0
         End If
         ACRE$ = CSTRING$(MKD$(ACREDIT#), 3, 14, 2, 2)
         Call REPLA(Z$, ACRE$, 70, 14)
         SUMACRE# = SUMACRE# + CDbl(Val(ACRE$))
       Else
         Call REPLA(Z$, Space$(14), 70, 14)
         SUMACRE# = SUMACRE# - CDbl(Val(IACRE$))
    End If
    VACON% = 1
    List2.AddItem Z$, ILISTA&
    List2.RemoveItem ILISTA& + 1
    List2.Refresh
    List2.ListIndex = ILISTA&
    VACON% = 0
    '
    Label10.Caption = CSTRING$(MKD$(SUMACRE#), 3, 14, 2, 2)
    Label10.Refresh
    '
End Sub
'

Private Sub List2_DblClick()
    '
    VACON% = 1
    ILIST2& = List2.ListIndex
    TOLIST2& = List2.TopIndex
    '
    Z$ = List2.TEXT
    ILISTA& = List2.ListIndex
    IACRE$ = TRIM$(Mid$(Z$, 70, 14))
    IPEN$ = Mid$(Z$, 56, 14)
    SUMACRE# = Val(Label10.Caption)
    SUMACRE# = SUMACRE# - CDbl(Val(IACRE$))
    SUMACRE# = SUMACRE# + CDbl(Val(Mid$(LIST2ANT$, 70, 14)))
    List2.Visible = False
    List2.AddItem LIST2ANT$, ILISTA&
    List2.RemoveItem ILISTA& + 1
    '
    Call SIGUEFACO5
    '
    VACON% = 1
    List2.TopIndex = TOLIST2&
    List2.Visible = True
    If List2.ListCount > ILIST2& Then
          List2.ListIndex = ILIST2&
        Else
          Call CIERRARCH("*.*")
          List1.Clear
          List2.Clear
          Label4.Caption = ""
          Label8.Caption = ""
          Label10.Caption = ""
          Text1.SetFocus
    End If
    VACON% = 0
    '
End Sub

Private Sub Option1_Click(Index As Integer)
   For KKI% = 0 To 5
     Option1(KKI%).FontBold = False
     Option1(KKI%).FontSize = 8
     If KKI% = Index Then
       Option1(KKI%).FontBold = True
       Option1(KKI%).FontSize = 10
     End If
   Next KKI%
   '
   '\\\OT-05-0527-WAR-04/08/05///
   Check1 = 0
   '\\\OT-05-0527-WAR-FIN///
   If Option1(0) = True Then
        Label6.Caption = "Factura Nro:"
        TIDOCOM$ = "FC"
      ElseIf Option1(1) = True Then
        Label6.Caption = "N.Débito Nro:"
        TIDOCOM$ = "ND"
        '\\\OT-05-0527-WAR-04/08/05///
        Check1 = 0
        '\\\OT-05-0527-WAR-FIN///
      ElseIf Option1(2) = True Then
        Label6.Caption = "N.Crédito Nro:"
        TIDOCOM$ = "NC"
        '\\\OT-05-0527-WAR-04/08/05///
        Check1 = 0
        '\\\OT-05-0527-WAR-FIN///
      ElseIf Option1(3) = True Then
        Label6.Caption = "Ticket-Factura:"
        TIDOCOM$ = "TF"
      ElseIf Option1(4) = True Then
        Label6.Caption = "R.Honorarios:"
        TIDOCOM$ = "RH"
        '\\\OT-05-0527-WAR-04/08/05///
        Check1 = 0
        '\\\OT-05-0527-WAR-FIN///
      ElseIf Option1(5) = True Then
        Label6.Caption = "Comprobante:"
        TIDOCOM$ = "CO"
        '\\\OT-05-0527-WAR-04/08/05///
        Check1 = 0
        '\\\OT-05-0527-WAR-FIN///
   End If
   Label6.Refresh
   Call SIGUEFACO1
End Sub

Private Sub Text1_Change()
    List1.Clear
    List2.Clear
    Label4.Caption = ""
    Label8.Caption = ""
    Label10.Caption = ""
    Check1.Enabled = True
    If TRIM(Text1) <> "" Then Check1.Enabled = False
    Call BLAFACOM
    '
    If Val(Text1.TEXT) > 32767 Then
      Text1.TEXT = Left$(Text1.TEXT, Len(Text1.TEXT) - 1)
      Text1.SelStart = Len(Text1.TEXT)
    End If
    '
End Sub

Private Sub Text1_DblClick()
    '
    If MODIFIX% > 0 Then
       If COMALTER%("Abandona sin Guardar los Cambios ?\\No, Continuar\Si, Cancelar") <> 2 Then
          Exit Sub
       End If
    End If
    '
    Call SELECHO("ACREDOR")
    NCXI% = Val(TRIM$(VALACT1$("ACREDOR")))
    If NCXI% > 0 Then
        NC% = NCXI%
        Text1.TEXT = TRIM$(Str$(NC%))
        Text2.TEXT = RASOCLI$((-1) * NC%)
        Text3.TEXT = DOMICLI$((-1) * NC%)
        Text4.TEXT = CPOSCLI$((-1) * NC%)
        Text5.TEXT = LOCACLI$((-1) * NC%)
        Text13.TEXT = CUITCLI$((-1) * NC%)
        If PIVACLI%((-1) * NC%) = 1 Then
            Label20 = "A"
          Else
            Label20 = "C"
        End If
        If Left$(Text2.TEXT, 6) = "VARIOS" Then
            Text13.SetFocus
          Else
            Text6.SetFocus
        End If
        'Call SIGUEFACO1
    End If
    '
End Sub

Private Sub Text1_GotFocus()
    Static CTXX%
    If CTXX% = 0 Then
        FORFACOM.Refresh
        CTXX% = 1
    End If
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.TEXT)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Val(Text1.TEXT) > 0 Then
          If Val(Text1.TEXT) <= 32767 Then
             NC% = Val(Text1.TEXT)
             If ECOARCH$("ACREDOR", MKI$(NC%)) <> "" Then
                Text2.TEXT = RASOCLI$((-1) * NC%)
                Text3.TEXT = DOMICLI$((-1) * NC%)
                Text4.TEXT = CPOSCLI$((-1) * NC%)
                Text5.TEXT = LOCACLI$((-1) * NC%)
                Text13.TEXT = CUITCLI$((-1) * NC%)
                If PIVACLI%((-1) * NC%) = 1 Then
                    Label20 = "A"
                  Else
                    Label20 = "C"
                End If
                If Left$(Text2.TEXT, 6) = "VARIOS" Then
                    Text13.SetFocus
                  Else
                    Text6.SetFocus
                End If
                Exit Sub
             End If
          End If
       End If
       '
       Call MENSERR(24, "Número de Cuenta no Válido")
       Text1.SetFocus
       '
    End If
    '
End Sub

Private Sub Text11_GotFocus()
    If TRIM$(Text11.TEXT) = "" Then
       Text11.TEXT = Space$(18)
    End If
    Text11.SelStart = 0
    Text11.SelLength = Len(Text11.TEXT)
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      If TRIM$(Text11.TEXT) <> "" Then
        If Val(Text1.TEXT) > 0 Then
          If Val(Text1.TEXT) <= 32767 Then
            NC% = Val(Text1.TEXT)
            If ECOARCH$("ACREDOR", MKI$(NC%)) <> "" Then
              Text1.SetFocus
              Exit Sub
            End If
          End If
        End If
      End If
    End If
End Sub

Private Sub Text11_LostFocus()
    '
    If TRIM$(Text11.TEXT) <> "" Then
'      '\\\OT-05-0527-WAR-20/07/05///
'      If FORFACOM.Check1 Then
'        Call FORFACOM.SIGUEBORE(Val(FORFACOM.Text1)) 'SIGUE CON BR'S SIN FACTUAR
'        Exit Sub
'      Else
'      '\\\OT-05-0527-WAR-FIN///
        Call SIGUEFACO1
     End If
     '
    'End If
    '
End Sub

Private Sub Text2_GotFocus()
   If Left$(Text2.TEXT, 6) <> "VARIOS" Then
      Text1.SetFocus
      Exit Sub
   End If
   Text2.SelStart = 0
   Text2.SelLength = Len(Text2.TEXT)
End Sub

Private Sub Text3_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text4_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text5_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text6_Click()
    'Text1.SetFocus
End Sub

Private Sub Text6_GOTFOCUS()
    NC% = Val(Text1.TEXT)
    If NC% < 1 Or ECOARCH$("ACREDOR", MKI$(NC%)) = "" Then
       Call MENSERR(24, "Número de Cuenta no Válido")
       Exit Sub
    End If
    '
    If Left$(RASOCLI$((-1) * NC%), 6) <> "VARIOS" Then
       Text2.TEXT = RASOCLI$((-1) * NC%)
       Text3.TEXT = DOMICLI$((-1) * NC%)
       Text4.TEXT = CPOSCLI$((-1) * NC%)
       Text5.TEXT = LOCACLI$((-1) * NC%)
       Text13.TEXT = CUITCLI$((-1) * NC%)
       If PIVACLI%((-1) * NC%) = 1 Then
           Label20 = "A"
         Else
           Label20 = "C"
       End If
    End If
    '
    Text6.SelStart = 0
    Text6.SelLength = Len(Text6.TEXT)
End Sub

Private Sub Text9_DblClick()
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Text9.TEXT = VDEVV$
        FF% = FECHANUM(Text9.TEXT)
    End If
End Sub
'
Private Sub Text6_DblClick()
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Text6.TEXT = VDEVV$
    End If
End Sub

Private Sub Text9_GotFocus()
    Text9.SelStart = 0
    Text9.SelLength = Len(Text9.TEXT)
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
      Text11.SetFocus
   End If
End Sub

Private Sub Text9_LostFocus()
    VDEVV$ = TRIM$(Text9.TEXT)
    If FECHANUM(VDEVV$) > 0 Then
          FF% = FECHANUM(VDEVV$)
          FEX = FF%
          Text9.TEXT = FECHATEX(FEX)
        Else
          Call MENSERR(24, "Fecha no Válida")
          FEX = HOY(HO$)
          Text9.TEXT = HO$
          Text9.SetFocus
    End If
End Sub

Private Sub Text6_LostFocus()
    VDEVV$ = TRIM$(Text6.TEXT)
    If FECHANUM(VDEVV$) > 0 Then
          FF% = FECHANUM(VDEVV$)
          FEX = FF%
          Text6.TEXT = FECHATEX(FEX)
        Else
          Call MENSERR(24, "Fecha no Válida")
          FEX = HOY(HO$)
          Text6.TEXT = HO$
          Text6.SetFocus
    End If
End Sub

Sub CALTOFAC()
   Dim FVI%(10), IM#(10)
   '
   TOBRUFAC# = Val(ITOGRA.Caption)
   For KKI% = 0 To 6
     TOBRUFAC# = TOBRUFAC# + (Int(100 * Val(IMPOIVA(KKI%).TEXT) + 0.5)) / 100
     If KKI% > 0 Then
        TOBRUFAC# = TOBRUFAC# + (Int(100 * Val(IMPOGRA(KKI%).TEXT) + 0.5)) / 100
     End If
   Next KKI%
   TOBRUFAC# = TOBRUFAC# + (Int(100 * Val(Text7.TEXT) + 0.5)) / 100
   TOBRUFAC# = TOBRUFAC# + (Int(100 * Val(Text8.TEXT) + 0.5)) / 100
   TOBRUFAC# = TOBRUFAC# + (Int(100 * Val(Text12.TEXT) + 0.5)) / 100
   '
   NCI% = Val(Text1)
   If NCI% < 1 Then Exit Sub
   '
   COPAGI% = CPAGCLI%((-1) * NCI%)
   Call BUCONPA(LU$, COPAGI%, COPAG$, FREF%, PLA$)
   
   IORIG# = TOBRUFAC#
   FF = FECHANUM(Text6)
   FFS$ = FECHATEX$(FF)
   PLA$ = Left$(PLA$ + String$(6, 0), 6)

   If FREF% = 2 Then FREF% = 1   ' AUN NO FUNCIONA SEMANA DE FACTURA
   If FREF% = 3 Then
     If Val(Left$(FFS$, 2)) <= 15 Then FEBAS$ = "15" + Mid$(FFS$, 3)
     If Val(Left$(FFS$, 2)) > 15 Then FEBAS$ = "31" + Mid$(FFS$, 3)
   End If
   If FREF% = 1 Then FEBAS$ = FFS$
   If FREF% = 4 Then FEBAS$ = "31" + Mid$(FFS$, 3)
   '
   For I% = 1 To 6: FVI%(I%) = 0: Next I%
   A11 = FECHANUM(FEBAS$): B22 = A11
   NUPAG% = 1
   For I% = 1 To 6
     PLAI% = Asc(Mid$(PLA$, I%, 1))
     If PLAI% > 0 Then
       While DIENTRE%(A11, B22) < PLAI%
         B22 = DIPOST(B22)
       Wend
       FVI%(I%) = B22
       NUPAG% = I%
     End If
   Next I%
   '
   If NUPAG% < 1 Then NUPAG% = 1
   IM#(1) = (Int(100 * IORIG# / NUPAG% + 0.5)) / 100
   '
   For H% = 2 To NUPAG% - 1
     IM#(H%) = IM#(1)
   Next H%
   IM#(NUPAG%) = IORIG#
   '
   For H% = 1 To NUPAG% - 1
     IM#(NUPAG%) = IM#(NUPAG%) - IM#(H%)
   Next H%
   '
   ICANCE# = 0
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
   For H% = 1 To 6
     FEX = FVI(H%)
     FEVEN(H% - 1) = FECHATEX$(FEX)
     IVENCI(H% - 1) = FORMATNUM$(IM#(H%), "F11.2")
   Next H%
   '
End Sub


'
Sub BLAFACOM()
   '
   TOTIMPU# = 0
   TONEGRA# = 0
   TOEXEN# = 0
   TOTIVA# = 0
   TORETIV# = 0
   TOTFAC# = 0
   '
   Label25 = ""
   Label21 = ""
   Text14 = ""
   Text15 = ""
   '
   List1.Clear
   List2.Clear
   Label4.Caption = ""
   Label8.Caption = ""
   Label10.Caption = ""
   Text11.TEXT = ""
   '
   JJ& = 0
   For U& = 1 To ULTREG("PROVINB")
     XX$ = REGLEIDO$("PROVINB", U&)
     If CVI(Mid$(XX$, 27, 2)) > 0 Then
       If ECOARCH$("CUECON", Mid$(XX$, 27, 2)) <> "" Then
         Call REPLA(XX$, String$(8, 0), 57, 8)
         JJ& = JJ& + 1
         Call GRAREG("!PROVINB", XX$, JJ&)
       End If
     End If
   Next U&
   Call SETULTREG("!PROVINB", JJ&)
   Call CIERRARCH("PROVINB")
   '
   ITOGRA.Caption = Space$(11): ITOGRA.Refresh
   Text7.TEXT = Space$(7)
   Text8.TEXT = Space$(7)
   Text12.TEXT = Space$(7)
   Text12.Enabled = True
   Text10.TEXT = " "
   Label18.Caption = ""
   For KKI% = 1 To 6: IMPOGRA(KKI%).TEXT = Space$(11): Next KKI%
   For KKI% = 0 To 6: IMPOIVA(KKI%).TEXT = Space$(9): Next KKI%
   For KKI% = 0 To 5: FEVEN(KKI%).TEXT = Space$(8): Next KKI%
   For KKI% = 0 To 5: IVENCI(KKI%).TEXT = Space$(11): Next KKI%
   '
   For KKI% = 0 To 10
      TOFAC#(KKI%) = 0
      FVII%(KKI%) = 0
      IMV#(KKI%) = 0
      IMV1#(KKI%) = 0
      GRU%(KKI%) = 0
      IPAR#(KKI%) = 0
   Next KKI%
   '
   MODIFIX% = 0
   '\\\OT-05-0527-WAR-04/08/05///
   Label28 = ""
   Text17 = ""
   Text17.BackColor = vbWhite
   Label29 = ""
   Label31 = ""
   '\\\OT-05-0527-WAR-FIN///
End Sub
'
Sub GRAFACOM()
   '
   Dim TOFAC1#(10), GRU1%(10), IPAR1#(10)
   '
   NROX$ = TRIM$(FORFACOM.Text11.TEXT)
   If Len(NROX$) < 1 Then
      Call MENSERR(24, "Número de Comprobante no Válido")
      Text11.SetFocus
      Exit Sub
   End If
   '
   If Left$(NROX$, 3) = "CO-" Then
      Call MENSERR(24, "Número de Comprobante no Válido")
      Text11.SetFocus
      Exit Sub
   End If
   '
   If TIDOCOM$ = "CO" Then
      If Len(NROX$) > 15 Then
         Call MENSERR(24, "Número de Comprobante no Válido")
         Text11.SetFocus
         Exit Sub
      End If
      NROX$ = "CO-" + TRIM$(Label20.Caption) + "-" + AJUSTI$(NROX$, 13)
   End If
   '
   SIG% = 1: TIPO% = 2
   If Left$(TRIM$(NROX$), 2) = "TF" Then TIPO% = 22
   If Left$(TRIM$(NROX$), 2) = "RH" Then TIPO% = 32
   If Left$(TRIM$(NROX$), 2) = "CO" Then TIPO% = 42
   If Left$(TRIM$(NROX$), 2) = "ND" Then TIPO% = 3
   If Left$(TRIM$(NROX$), 2) = "NC" Then TIPO% = 6: SIG% = (-1)
   '
   If Len(NROX$) <> 18 Then
      Call MENSERR(24, "Formato Número de Comprobante no Válido")
      VACON% = 0
      Exit Sub
   End If
   '
   SUC% = 0: NRO& = 0
   If TIDOCOM$ <> "CO" Then
      VA% = Asc(Mid$(NROX$, 4, 1)) - 64
      If VA% < 1 Or VA% > 3 Then
          Call MENSERR(24, "Letra de Documento Debe ser 'A' o 'C'")
          VACON% = 0
          Exit Sub
      End If
      SUC% = Val(Mid$(NROX$, 6, 4))
      NRO& = Val(Mid$(NROX$, 11, 8))
      '
      If NRO& < 1 Then
         Call MENSERR(24, "Falta Número de Documento")
      End If
   End If
   '
   NUDOCUCOM$ = NROX$
   '
   NC% = Val(FORFACOM.Text1.TEXT)
   If NC% < 1 Or ECOARCH$("ACREDOR", MKI$(NC%)) = "" Then
      Call MENSERR(24, "Cuenta Inexistente en B.D. de Proveedores")
      VACON% = 0
      Exit Sub
   End If
   '
   If TICUEPRO%((-1) * NC%) <> 0 Then
      Call MENSERR(24, "Imposible Registrar.\Cuenta de Gastos.")
      VACON% = 0
      Exit Sub
   End If
   '
   'CUITT$ = CUITCLI$((-1) * NC%)
   CUITT$ = TRIM$(Text13.TEXT)
   If CUITT$ = "" Or VALICUIT%(CUITT$) < 1 Then
      PVCL% = PIVACLI%((-1) * NC%)
      If PVCL% <> 0 And PVCL% <> 5 Then ' EXCLUYE EXPORTACION / IMPORTACION
         Call MENSERR(24, "Número de C.U.I.T. de Proveedor no Válido")
         VACON% = 0
         Exit Sub
      End If
   End If
   '
   FF% = FECHANUM(FORFACOM.Text9.TEXT)
   HOI% = HOY(HO$)
   If FF% < 1 Or FF% > HOI% + 1 Then
       Call MENSERR(24, "Fecha no Válida")
       FORFACOM.Text9.SetFocus
       VACON% = 0
       Exit Sub
   End If
   '
   FFRE% = FECHANUM(FORFACOM.Text6.TEXT)
   HOI% = HOY(HO$)
   If FFRE% < 1 Or FFRE% > HOI% + 1 Then
       Call MENSERR(24, "Fecha de Emisión no Válida")
       FORFACOM.Text6.SetFocus
       VACON% = 0
       Exit Sub
   End If
   '
   If DIENTRE%(FF%, HOI%) > 55 Then
       Call MENSERR(24, "Fecha no Válida")
       FORFACOM.Text9.SetFocus
       VACON% = 0
       Exit Sub
   End If
   '
   FECHALIM% = Val(CONTROL$("", "CIEMECOM"))
   If FF% <= FECHALIM% Then
       Call MENSERR(24, "Fecha no Válida.\  \Ese Período Mensual Está Cerrado.")
       FORFACOM.Text9.SetFocus
       VACON% = 0
       Exit Sub
   End If
   '
   If FF% < FECHANUM(FORFACOM.Text6.TEXT) Then
       Call MENSERR(24, "No Válido Contabilizar en Fecha Anterior\a la de Emisión del Comprobante.")
       FORFACOM.Text9.SetFocus
       VACON% = 0
       Exit Sub
   End If
   '
   If Val(FORFACOM.ITOGRA.Caption) < (-0.005) Then
      Call MENSERR(24, "Importes no Válidos")
      Call SIGUEFACO3
      Exit Sub
   End If
   '
   If LETRAFA$ <> "A" Then
     If Abs(Val(FORFACOM.ITOGRA.Caption)) >= 0.005 Then
       Call MENSERR(24, "Importes no Válidos")
       Call SIGUEFACO3
       Exit Sub
     End If
   End If
   '
   SUIMPUTA# = 0
   KKI% = 0
   For U& = 1 To FORFACOM.List1.ListCount
      FORFACOM.List1.ListIndex = U& - 1
      Z$ = FORFACOM.List1.TEXT
      CUE% = Val(Left$(Z$, 5))
      IPA# = Val(Mid$(Z$, 40, 12))
      '
      If CUE% > 0 Then
         If IPA# >= 0.005 Then
            KKI% = KKI% + 1
            If KKI% > 8 Then GoTo 55
            GRU%(KKI%) = CUE%
            IPAR#(KKI%) = IPA#
            '\\\OT-05-0527-WAR-09/08/05///
            COEBON1 = COEFDESCU(TRIM$(Text17))
            If COEBON1 > 0 Then
               IPA# = (Int(100 * IPA# * (1 - COEBON1 / 100) + 0.5)) / 100
               IPAR#(KKI%) = IPA#
            End If
            '\\\OT-05-0527-WAR-FIN///
            SUIMPUTA# = SUIMPUTA# + IPA#
         End If
      End If
   Next U&
   '
55 If KKI% > 8 Then
      Call MENSERR(24, "Demasiados Conceptos - Máximo es 8")
      Exit Sub
   End If
   '
   If Val(Text12.TEXT) >= 0.005 Then
      If TRIM$(Label18.Caption) = "" Then
         Call MENSERR(24, "Imposible Registrar - Falta Jurisdiccion I.Brutos")
         Exit Sub
      End If
   End If
   '
   SUNETO# = Val(ITOGRA.Caption)
   SUBRUTO# = SUNETO#
   '
   TOFAC#(1) = SIG% * SUNETO#
   TOFAC#(2) = SIG% * Abs(Val(FORFACOM.IMPOGRA(4).TEXT))
   TOFAC#(3) = SIG% * Abs(Val(FORFACOM.IMPOGRA(6).TEXT)) ' impuestos intermos
   TOFAC#(4) = 0                    ' IVA TASAS GENERALES
   TOFAC#(5) = 0                    ' iva ADUANA
   TOFAC#(6) = SIG% * Abs(Val(FORFACOM.Text7.TEXT))
   TOFAC#(7) = SIG% * Abs(Val(FORFACOM.Text8.TEXT))
   TOFAC#(8) = SIG% * Abs(Val(FORFACOM.Text12.TEXT))
   '
   For KKI% = 0 To 6
     SUBRUTO# = SUBRUTO# + (Int(100 * Val(IMPOIVA(KKI%).TEXT) + 0.5)) / 100
     If KKI% <> 5 Then
          TOFAC#(4) = TOFAC#(4) + SIG% * (Int(100 * Val(IMPOIVA(KKI%).TEXT) + 0.5)) / 100
        Else
          TOFAC#(5) = TOFAC#(5) + SIG% * (Int(100 * Val(IMPOIVA(KKI%).TEXT) + 0.5)) / 100
     End If
     If KKI% > 0 Then
        SUNETO# = SUNETO# + (Int(100 * Val(IMPOGRA(KKI%).TEXT) + 0.5)) / 100
        SUBRUTO# = SUBRUTO# + (Int(100 * Val(IMPOGRA(KKI%).TEXT) + 0.5)) / 100
        If KKI% <> 4 And KKI% <> 6 Then
          TOFAC#(1) = TOFAC#(1) + SIG% * (Int(100 * Val(IMPOGRA(KKI%).TEXT) + 0.5)) / 100
        End If
     End If
   Next KKI%
   '
   '\\\OT-05-0527-WAR-04/08/05///
   If Abs(SUNETO - SUIMPUTA#) >= 0.005 Then
   'If Abs(SUIMPUTA# - Val(Label28)) >= 0.005 Then
      Call MENSERR(24, "Importes Incoherentes")
      Exit Sub
   End If
   '
   SUBRUTO# = SUBRUTO# + (Int(100 * Val(Text7.TEXT) + 0.5)) / 100
   SUBRUTO# = SUBRUTO# + (Int(100 * Val(Text8.TEXT) + 0.5)) / 100
   SUBRUTO# = SUBRUTO# + (Int(100 * Val(Text12.TEXT) + 0.5)) / 100
   TOFAC#(0) = (-1) * SIG% * SUBRUTO#
   '
   If SUNETO# < 0 Or SUBRUTO# < 0.005 Then
      Call MENSERR(24, "Falta Ingresar Importes")
      VACON% = 0
      Exit Sub
   End If
   '
   SUMAVEN# = 0
   For KKI% = 1 To 6
      IMV1#(KKI%) = 0
      FVII%(KKI%) = FECHANUM(FORFACOM.FEVEN(KKI% - 1).TEXT)
      If FVII%(KKI%) > 0 Then
         IMV#(KKI%) = Val(FORFACOM.IVENCI(KKI% - 1).TEXT)
         If IMV#(KKI%) < 0 Then
            Call MENSERR(24, "Importe a Vencer Negativo")
            VACON% = 0
            Exit Sub
         End If
         IMV1#(KKI%) = IMV#(KKI%) * SIG%
         SUMAVEN# = SUMAVEN# + IMV#(KKI%)
      End If
   Next KKI%
   If Abs(SUMAVEN# - SUBRUTO#) >= 0.005 Then
      Call MENSERR(24, "Suma de Vencimientos Difiere de Total")
      VACON% = 0
      Exit Sub
   End If
   '
   SUMATOT# = 0
   For KKI% = 0 To 8
      SUMATOT# = SUMATOT# + TOFAC#(KKI%)
   Next KKI%
   If Abs(SUMATOT#) >= 0.005 Then
      Call MENSERR(24, "Valores Incoherentes")
      VACON% = 0
      Exit Sub
   End If
   '
   VACON% = 1
   SUMACRE# = 0: ATX% = 0
   Call BLOQARCH("CUECORR")
   Call BLOQARCH("CUECOSS")
   Call BLOQARCH("PROVED2")
   '
   For U& = 1 To FORFACOM.List2.ListCount
      FORFACOM.List2.ListIndex = U& - 1
      Z$ = FORFACOM.List2.TEXT
      If Val(Mid$(Z$, 70, 14)) < 0 Then
         Call MENSERR(24, "Valores Aplicados Negativos")
         VACON% = 0
         Exit Sub
      End If
      '
      DOCU$ = Left$(Z$, 20)
      If DOCU$ <> "Anticipo no Aplicado" Then
         CANCE# = CDbl(Val(Mid$(Z$, 70, 14)))
         If CANCE# > 0 Then
            NREG& = Val(TRIM$(Mid$(Z$, 119, 8)))
            If NREG& > 0 Then
              If NREG& <= ULTREG&("CUECORR") Then
                ORD% = Val(Mid$(Z$, 127, 2))
                If ORD% > 0 Then
                  If ORD% <= 6 Then
                    '
                    SS$ = REGLEIDO$("CUECORR", NREG&)
                    If Mid$(SS$, 7, 18) <> Left$(DOCU$, 18) Then ATX% = 1
                    SALDO$ = Mid$(SS$, 139, 60)
                    SALDEU# = CVD(Mid$(SALDO$, 10 * ORD% - 7, 8))
                    If SALDEU# + 0.005 < CANCE# Then ATX% = 1
                    SUMACRE# = SUMACRE# + CDbl(Val(Mid$(Z$, 70, 14)))
                  End If
                End If
              End If
            End If
         End If
      End If
   Next U&
   '
   If ATX% > 0 Then
      Call MENSERR(24, "Imposible Registrar este Documento.\  \Reingrese los Datos Completos.\Si Persiste el Problema,\Consulte a su Soporte de Sistemas FLEXOFT.")
      GoTo 84
   End If
   '
   If SUMACRE# > Abs(TOFAC#(0)) + 0.005 Then
      Call MENSERR(24, "Importes Aplicados Superan Total Factura")
      VACON% = 0
      Exit Sub
   End If
   '
   If TIDOCOM$ = "CO" Or TIDOCOM$ = "NC" Then
     If Val(Label25) > 0 Then
       Call COMUNI("Imposible Cancelación Contado.\Genere Orden de Pago C.Cte.")
       NRECIB$ = TRIM$(Text11)
       Exit Sub
     End If
   End If
   '
   Call BLOQARCH("CUECORR")
   Call BLOQARCH("CUECOSS")
   Call BLOQARCH("APLIPAG")
   Call BLOQARCH("PROVED2")
   If Val(Label25) > 0 Then Call BLOQARCH("CHEQUEMI")
   Screen.MousePointer = 11
   '
   REFEDOC$ = AJUSTI$(Text2.TEXT, 24) + TRIM$(Text13.TEXT)
   RECUECO& = 1 + ULTREG&("CUECORR")
   If Val(Label25) > 0 Then
     For KKI% = 0 To 10: FVII%(KKI%) = 0: IMV1#(KKI%) = 0: Next KKI
   End If
   '
   Call CUECORR(FF%, (-1) * NC%, TIPO%, VA%, SUC%, NRO&, REFEDOC$, MO$, TOFAC#(), FVII%(), IMV1#(), GRU%(), IPAR#(), 1, FFRE%)
   Call CUECOSS(NUDOCUCOM$)
   '
   If Val(Label25) > 0 Then
     If TIDOCOM$ = "CO" Then
       Call COMUNI("Imposible Cancelación Contado.\Genere Orden de Pago C.Cte.")
       NRECIB$ = TRIM$(Text11)
       Exit Sub
     End If
     Call REPLA(NUDOCUCOM$, "FF", 1, 2)
     For KKI% = 0 To 10: TOFAC1#(KKI%) = (-1) * TOFAC#(KKI%): IMV1#(KKI%) = 0: Next KKI%
     For KKI% = 0 To 10: GRU1%(KKI%) = 0: IPAR1#(KKI%) = 0: Next KKI%
     Call CUECORR(FF%, (-1) * NC%, 14, VA%, SUC%, NRO&, REFEDOC$, MO$, TOFAC1#(), FVII%(), IMV1#(), GRU1%(), IPAR1#(), 1, FFRE%)
     Call CUECOSS(NUDOCUCOM$)
     '
     TIDOS$ = "."
     TIPO% = Val(Label25)
     BANCO$ = TRIM$(Label21)
     NUCHE$ = TRIM$(Text14)
     IMCHE# = Abs(TOFAC#(0))
     FECHE% = FECHANUM(Text9): If FECHE% < 1 Then FECHE% = HOY(HOS$)
     FVEN% = FECHANUM(Text15): If FVEN% < FECHE% Then FVEN% = FECHE%
     '
     Z$ = REGBLAN$("CHEQUEMI")
     Z$ = REPLACE$(Z$, MKI$(NC%), 1, 2)
     Z$ = REPLACE$(Z$, NUCHE$, 3, 12)
     Z$ = REPLACE$(Z$, BANCO$, 15, 16)
     Z$ = REPLACE$(Z$, MKI$(FECHE%), 31, 2)
     Z$ = REPLACE$(Z$, MKD$(IMCHE#), 33, 8)
     Z$ = REPLACE$(Z$, MKS$(NRO&), 41, 4)
     Z$ = REPLACE$(Z$, MKI$(TIPO%), 45, 2)
     Z$ = REPLACE$(Z$, MKI$(FF%), 47, 2)
     Z$ = REPLACE$(Z$, MKI$(FVEN%), 49, 2)
     Call REGAPP("CHEQUEMI", Z$)
     Call CIERRARCH("CHEQUEMI")
     '
   End If
   '
80 ' *********************************** aqui blanquear facturas canceladas
   '
   NC1% = (-1) * NC%
   VACON% = 1
   '
   XX% = 0
   For U& = 1 To FORFACOM.List2.ListCount
      FORFACOM.List2.ListIndex = U& - 1
      Z$ = FORFACOM.List2.TEXT
      DOCU$ = Left$(Z$, 20)
      '
      If DOCU$ <> "Anticipo no Aplicado" Then
         CANCE# = CDbl(Val(Mid$(Z$, 70, 14)))
         If CANCE# > 0 Then
            NREG& = Val(TRIM$(Mid$(Z$, 119, 8)))
            If NREG& > 0 Then
               If NREG& <= ULTREG&("CUECORR") Then
                  ORD% = Val(Mid$(Z$, 127, 2))
                  If ORD% > 0 Then
                    If ORD% <= 6 Then
                      '
                      SS$ = REGLEIDO$("CUECORR", NREG&)
                      TS$ = REGLEIDO$("CUECOSS", NREG&)
                      SALDO$ = Mid$(SS$, 139, 60)
                      SALDEU# = CVD(Mid$(SALDO$, 10 * ORD% - 7, 8))
                      SALDEU# = SALDEU# - CANCE#
                      Call REPLA$(SALDO$, MKD$(SALDEU#), 10 * ORD% - 7, 8)
                      '
                      If Abs(SALDEU#) < 0.005 Then
                        SALDO$ = Left$(SALDO$, 10 * (ORD% - 1)) + Mid$(SALDO$, 10 * ORD% + 1) + String$(10, 0)
                      End If
                      '
                      Call REPLA(SS$, SALDO$, 139, 60)
                      Call GRAREG("CUECORR", SS$, NREG&)
                      Call REPLA(TS$, SALDO$, 139, 60)
                      Call GRAREG("CUECOSS", TS$, NREG&)
                    End If
                  End If
               End If
            End If
         End If
      End If
   Next U&
   '
   OKX% = 3         ' depuró facturas pendientes
   '
83 ' *************************************** aqui grabar archivo APLIPAG.DAT
   '
   For U& = 1 To FORFACOM.List2.ListCount
      FORFACOM.List2.ListIndex = U& - 1
      Z$ = FORFACOM.List2.TEXT
      DOCU$ = Left$(Z$, 20)
      If DOCU$ <> "Anticipo no Aplicado" Then
         CANCE# = CDbl(Val(Mid$(Z$, 70, 14)))
         If CANCE# > 0 Then
            NREG& = Val(TRIM$(Mid$(Z$, 119, 8)))
            If NREG& > 0 Then
               If NREG& <= ULTREG&("CUECORR") Then
                  SS$ = REGLEIDO$("CUECORR", NREG&)
                  DOCORI$ = Mid$(SS$, 7, 18)
                  SALDO$ = Mid$(SS$, 139, 60)
                  SS77$ = Mid$(SS$, 51)
                  ITOTORI# = Abs(CVD(Mid$(SS$, 115, 8)) - CVD(Mid$(SS$, 123, 8)))
                  IVAORI# = Abs(CVD(Mid$(SS77$, 25, 8)) + CVD(Mid$(SS77$, 33, 8)) + CVD(Mid$(SS77$, 41, 8)))
                  '
                  Call APLIPAG(FC%, NC1%, 4, 1, SUC%, NRO&, CANCE#, RefApli$, DOCORI$, ITOTORI#, IVAORI#)
                  '
               End If
            End If
         End If
      End If
   Next U&
   '
   OKX% = 4       ' grabó aplicaciones
   'Grabo Detalle de Facturación en Archivo-Base Central
   '\\\OT-05-0527-WAR-03/08/05///
   ZZ$ = REGBLAN("DETFACOM")
   ULTREDET& = ULTREG("DETFACOM")
   CONTI1& = 0
   For J& = 1 To ULTREG("!DETFACOM")
    DETO$ = REGLEIDO("!DETFACOM", J&)
    '
    CAFA1 = CVD(Mid$(DETO$, 95, 8))
    If CAFA1 > 0.005 Then
      ZZ$ = DETO$
      '
      NUFA1$ = TRIM(Text11)
      Call REPLA(ZZ$, NUFA1$, 3, 18) 'NUMERO FACTURA
      Call REPLA(ZZ$, MKI$(FFRE%), 21, 2) 'F.EMISION
      Call REPLA(ZZ$, MKI$(FF%), 23, 2) 'FECHA CONTABLE
      '
      CONTI1& = CONTI1& + 1
      Call GRAREG("DETFACOM", ZZ$, ULTREDET& + CONTI1&)
    End If
   Next J&
   Call SETULTREG("DETFACOM", ULTREDET& + CONTI1&)
   '\\\OT-05-0527-WAR-FIN///
   '\\\OT-05-0527-WAR-04/08/05///
   XX1$ = REGBLAN("FACBONCO") 'GRABO UN REGISTRO EN EL FACBONCO
   ULTREFA& = ULTREG("FACBONCO")
   '
   TICO1$ = Left(Text11, 2)
   COMPRO1$ = Text11
   NC% = Val(Text1)
   ISINBON# = Val(Label28)
   PORBON1$ = Text17
   IBON1# = Val(Label29)
   ICONBON# = Val(Label31)
   COEFI1# = 1
   If Abs(ISINBON#) >= 0.005 Then
     COEFI1# = ICONBON# / ISINBON# * 100
   End If
   '
   Call REPLA(XX1$, MKI$(NC%), 1, 2)
   Call REPLA(XX1$, MKI$(FFRE%), 3, 2)
   Call REPLA(XX1$, MKI$(FF%), 5, 2)
   Call REPLA(XX1$, TICO1$, 7, 2)
   Call REPLA(XX1$, COMPRO1$, 9, 18)
   Call REPLA(XX1$, MKI$(NC%), 27, 2)
   Call REPLA(XX1$, MKD$(ISINBON#), 33, 8)
   Call REPLA(XX1$, PORBON1$, 41, 16)
   Call REPLA(XX1$, MKD$(IBON1#), 57, 8)
   Call REPLA(XX1$, MKD$(ICONBON#), 65, 8)
   Call REPLA(XX1$, MKD$(COEFI1#), 73, 8)
   Call GRAREG("FACBONCO", XX1$, ULTREFA& + 1)
   Call CIERRARCH("FACBONCO")
   'LIMPIO LOS SUB-TOTALES
   Unload FOFACOBR
   '
'   With FOFACOBR
'     .Text12 = ""
'     .Label7 = ""
'     .Label4 = ""
'     .LBLIMPO = ""
'     .Hide
'   End With
   '
   '\\\OT-05-0527-WAR-FIN///
84 Call CIERRARCH("APLIPAG")
   Call LIBARCH("APLIPAG")
   Call CIERRARCH("CUECOSS")
   Call LIBARCH("CUECOSS")
   Call CIERRARCH("PROVED2")
   Call LIBARCH("PROVED2")
   Call CIERRARCH("CUECORR")
   Call LIBARCH("CUECORR")
   Call CIERRARCH("CHEQUEMI")
   Call LIBARCH("CHEQUEMI")
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
   FORFACOM.List1.Clear
   FORFACOM.List2.Clear
   FORFACOM.Label4.Caption = ""
   FORFACOM.Label8.Caption = ""
   FORFACOM.Label10.Caption = ""
   FORFACOM.Text11.TEXT = ""
   FORFACOM.ITOGRA.Caption = Space$(11)
   'FORFACOM.TOEXEN.TEXT = ""
   'FORFACOM.TOTIVA.TEXT = ""
   'FORFACOM.TORETIV.TEXT = ""
   'FORFACOM.TOTFAC.Caption = ""
   For KKI% = 0 To 5
      FORFACOM.FEVEN(KKI%).TEXT = ""
      FORFACOM.IVENCI(KKI%).TEXT = ""
   Next KKI%
   Call BLAFACOM
   VACON% = 0
   Screen.MousePointer = 1
   FORFACOM.Text1.SetFocus
   '
End Sub
'
Sub CUECOSS(NUDOCUX$)
    '
    Static VALTASA(10), TETASA$(10), CTX%
    If CTX% = 0 Then
       CTX% = 1
       For IXX% = 1 To 4
          TXXX$ = TRIM$(Str$(IXX% - 1))
          TASA = Val(CONTROL$("", "TASIVA-" + TXXX$))
          If TASA > 0.005 Then
            VALTASA(IXX%) = TASA
            TETASA$(IXX%) = TRIM$(CSTRING$(MKS$(TASA), 3, 6, 2, 2) + "%")
          End If
       Next IXX%
       TETASA$(5) = "Exento": VALTASA(5) = 0
       TETASA$(6) = "Aduana": VALTASA(6) = 0
       TETASA$(7) = "Monotr": VALTASA(7) = 0
       TETASA$(8) = "I.Int.": VALTASA(8) = 0
    End If
    '
    NUDOCUY$ = AJUSTI$(NUDOCUX$, 18)
    Screen.MousePointer = 11
    URECORR& = ULTREG&("CUECORR")
    URECOSS& = ULTREG&("CUECOSS")
    For U& = URECOSS& + 1 To URECORR&
       Call TRACE(20, U&, URECORR&)
       U1& = U&
       X$ = REGLEIDO$("CUECORR", U1&)
       If U1& = U& Then
         X$ = X$ + String$(212, 0)
         Call GRAREG("CUECOSS", X$, U&)
       End If
    Next U&
    '
    For RECO& = URECORR& To 1 Step -1
       SS$ = REGLEIDO$("CUECORR", RECO&)
       If Mid$(SS$, 7, 18) = NUDOCUY$ Then GoTo 25
    Next RECO&
    Call MENSERR(24, "Inconsistencia en B.Datos de Compras- Consulte.")
    Call CIERRARCH("*.*")
    Call FINAL("")
    '
25  TS$ = REGLEIDO$("CUECOSS", RECO&)
    TT7$ = Mid$(TS$, 301, 212)
    NC1% = (-1) * Abs(CVI(Mid$(TS$, 5, 2)))
    '
    For KKI% = 1 To 8
       Call REPLA(TT7$, TETASA$(KKI%), 24 * KKI% - 7, 8)
       If KKI% < 7 Then
          If KKI% = 1 Then
              IMGRA# = Val(ITOGRA.Caption)
            Else
              IMGRA# = Val(IMPOGRA(KKI% - 1).TEXT)
          End If
          IMIVA# = Val(IMPOIVA(KKI% - 1).TEXT)
          If KKI% = 5 Then IMEXE# = IMGRA#: IMONO# = 0: IMIVA# = 0
          TASARE = 0
          If IMGRA# >= 0.005 Then
             TASARE = 100 * IMIVA# / IMGRA#
          End If
          Call REPLA(TT7$, MKD$(IMGRA#), 24 * KKI% - 23, 8)
          Call REPLA(TT7$, MKD$(IMIVA#), 24 * KKI% - 15, 8)
       End If
    Next KKI%
    '
    If PIVACLI%(NC1%) = 6 Then
       IMONO# = IMEXE#: IMEXE# = 0
    End If
    IMINT# = Val(IMPOGRA(6).TEXT)
    '
    Call REPLA(TT7$, MKD$(IMEXE#), 24 * 5 - 23, 8)
    Call REPLA(TT7$, MKD$(IMONO#), 24 * 7 - 23, 8)
    Call REPLA(TT7$, MKD$(IMINT#), 24 * 8 - 23, 8)
    '
    Call REPLA(TT7$, Text10.TEXT, 193, 1)
    Call REPLA(TT7$, Label18.Caption, 194, 19)
    '
123 Call REPLA(TS$, TT7$, 301, 212)
    Call GRAREG("CUECOSS", TS$, RECO&)
    Call CIERRARCH("CUECOSS")
    '
    If TRIM$(Text10) = "*" Then
      If Val(Text12) > 0 Then
        REBLA$ = REGBLAN$("PERCIBFA")
          FF% = FECHANUM(FORFACOM.Text9.TEXT)
        Call REPLA(REBLA$, MKI$(FF%), 1, 2)
          NPRX% = Val(Text1)
        Call REPLA(REBLA$, MKI$(NPRX%), 3, 2)
        Call REPLA(REBLA$, Text11, 5, 18)
        Call REPLA(REBLA$, RASOCLI$((-1) * NPRX%), 23, 30)
        Call REPLA(REBLA$, CUITCLI$((-1) * NPRX%), 53, 14)
        Call REPLA(REBLA$, NIBRCLI$((-1) * NPRX%), 67, 14)
          INETOX# = Val(ITOGRA): For KKI% = 1 To 5: INETOX# = INETOX# + Val(IMPOGRA(KKI%)): Next KKI%
        Call REPLA(REBLA$, MKD$(INETOX#), 81, 8)
        '
        SUPIBRU# = 0
        For U& = 1 To ULTREG&("!PROVINB")
          XX$ = REGLEIDO$("!PROVINB", U&)
          IPERPA# = CVD(Mid$(XX$, 57, 8))
          If Abs(IPERPA#) >= 0.005 Then
            REPER$ = REBLA$
              SUPIBRU# = SUPIBRU# + IPERPA#
            Call REPLA(REPER$, MKD$(IPERPA#), 89, 8)
            Call REPLA(REPER$, Mid$(XX$, 1, 20), 97, 20)
              CUCO% = CVI(Mid$(XX$, 27, 2))
            Call REPLA(REPER$, MKI$(CUCO%), 117, 2)
            Call REGAPP("PERCIBFA", REPER$)
          End If
        Next U&
        Call CIERRARCH("PERCIBFA")
      End If
    End If
    '
End Sub
'
Sub SIGUEBORE(NPROX%)
    '
    FORSELBRE.LIMATSEL.Clear
    RFF$ = RECFILT$("BOLRECEP", 10, MKI$(NPROX%))
    '
' ver mas abajo por que deshabilite esta parte
'  If Len(RFF$) < 4 Then
'    Call MENSERR(24, "No Hay Boletas de Recepción Registradas.")
'  Else
    'NUEVO FILTRO DE FECHA
    Call DESHASFECHA(DES%, HAS%, PERIO$)
    NBRVIE$ = "": NOCOVIE$ = ""
    For UK& = 1 To Len(RFF$) Step 4
      U& = CVS(Mid$(RFF$, UK&, 4))
      X$ = REGLEIDO$("BOLRECEP", U&)
      NFAC$ = Mid$(X$, 135, 16)
      CANAP = CVD(Mid$(X$, 163, 8))
      '\\\OT-05-0527-WAR-11/08/05///
      FERECEP = CVI(Mid$(X$, 81, 2))
      '\\\OT-05-0527-WAR-FIN///
      '
      If TRIM(NFAC$) = "" Then
      If FERECEP >= DES% Then
       If FERECEP <= HAS% Then
        If CANAP > 0.005 Then
          '
          NBR$ = CVS(Mid$(X$, 1, 4))  '  WALTER: Esta asignacion acá es licita pero es peligrosa. Hay que respetar el tipo de variable !!!
            While Len(NBR$) < 6: NBR$ = "0" + NBR$: Wend ' para formatear
          NRP$ = TRIM$(Mid$(X$, 119, 16))
          NOCOM$ = TRIM$(Str$(CVS(Mid$(X$, 151, 4))))
          FBORE = CVI(Mid$(X$, 81, 2))
          BORE$ = "BR." & NBR$ & " - " & FECHATEX(FBORE) & " - Rto." & NRP$
          If Val(NOCOM$) > 0 Then
            BORE$ = BORE$ + " - OC." + NOCOM$
            NOCOVIE$ = NOCOM$
          End If
          '
          With FORSELBRE
            If NBRVIE$ <> NBR$ Then
                .LIMATSEL.AddItem BORE$
                NBRVIE$ = NBR$
              ElseIf NOCOM$ <> NOCOVIE$ Then
                If Val(NOCOM$) > 0 Then
                  BORE$ = .LIMATSEL.List(.LIMATSEL.ListCount - 1)
                  .LIMATSEL.RemoveItem (.LIMATSEL.ListCount - 1)
                  BORE$ = BORE$ + ", " + NOCOM$
                  .LIMATSEL.AddItem BORE$
                  NOCOVIE$ = NOCOM$
                End If
            End If
          End With
          '
        End If
       End If
      End If
      End If
      '
    Next UK&
  'End If
  '
  ' hay que prever el caso de que no haya boletas aprobadas sin facturar !!!
    If FORSELBRE.LIMATSEL.ListCount < 1 Then
      Call MENSERR(24, "No Hay Boletas de Recepción\Pendientes de Facturación\para la Cuenta Elegida.")
      GoTo 25
    End If
  '
  FORSELBRE.Show 1
  HAYSEL% = 0
  For J& = 0 To FORSELBRE.LIMATSEL.ListCount - 1
    If FORSELBRE.LIMATSEL.Selected(J&) = True Then
      HAYSEL% = 1
    End If
  Next J&
  If HAYSEL% > 0 Then 'SI HAY SELECCION
25  With FOFACOBR
      .Text1 = Text1
      .Text2 = Text2
      .Text3 = Text3
      .Text4 = Text4
      .Text5 = Text5
      .Text6 = Text6
      .Text13 = Text13
      .Text9 = Text9
      .Text11 = Text11
      '
      With FORSELBRE
          J1& = 0
          For J& = 1 To FORSELBRE.LIMATSEL.ListCount
            If FORSELBRE.LIMATSEL.Selected(J& - 1) = True Then
              LI1$ = FORSELBRE.LIMATSEL.List(J& - 1)
              'SACO EL NUMERO DE BOLETA DE RECEPCION
              NBORE& = Val(Mid$(LI1$, 4, 6))
              On Error GoTo 0
              '
              RFF$ = RECFILT$("BOLRECEP", 1, MKS$(NBORE&))
              If Len(RFF$) >= 4 Then
                For UK& = 1 To Len(RFF$) Step 4
                  '
                  U& = CVS(Mid$(RFF$, UK&, 4))
                  X$ = REGLEIDO$("BOLRECEP", U&)
                  NFAC$ = Mid$(X$, 135, 16)
                  CANAP1# = CVD(Mid$(X$, 163, 8))
                  '
                  Y$ = REGBLAN$("!DETFACOM")
                  '
                   If TRIM(NFAC$) = "" Then
                    If CANAP1# > 0.005 Then
                      '
                      NRP$ = TRIM$(Mid$(X$, 119, 16))
                      FBORE = CVI(Mid$(X$, 81, 2))
                      CIXI% = CVI(Mid$(X$, 83, 2))
                      CARE# = CVD(Mid$(X$, 89, 8))
                      NUOC& = CVS(Mid$(X$, 151, 4))
                      LO1$ = TRIM(Mid$(X$, 69, 12))
                      LP1$ = TRIM(Mid$(X$, 107, 12))
                      'SACO EL NUMERO DE ITEM
                      ACA% = 0
                      For JJJ& = 1 To Len(LO1$)
                        JX$ = Mid$(LO1$, JJJ&, 1)
                        If JX$ = "-" Then
                          ACA% = ACA% + 1
                        End If
                        If ACA% = 2 Then Exit For
                      Next JJJ&
                      NUOIT% = Val(Mid(X$, 69 + JJJ&))
                      'PRECIO UNITARIO DE COMPRA
                      If NUOC& > 0 Then
                        REGOC$ = ULTIPRE$(NPROX%, CIXI%, NUOC&)
                        PREUN1# = CVD(Mid$(REGOC$, 67, 8))
                      Else
                        PREUN1# = 0
                      End If
                      'BUSCO SI HAY CANTIDAD YA FACTURADA
                      CANTIYAFAC# = 0
                      For JJX& = ULTREG&("DETFACOM") To 1 Step -1
                        ZZZ$ = REGLEIDO("DETFACOM", JJX&)
                        If NBORE& = CVS(Mid$(ZZZ$, 27, 4)) Then
                          If CIXI% = CVI(Mid$(ZZZ$, 25, 2)) Then
                            If NUOIT% = CVS(Mid$(ZZZ$, 31, 4)) Then
                              CANTFAC = CVD(Mid$(ZZZ$, 95, 8))
                              CANTIYAFAC# = CANTIYAFAC# + CANTFAC
                            End If
                          End If
                        End If
                      Next JJX&
                      QPENFAC# = CANAP1# - CANTIYAFAC#
                      If QPENFAC# <= 0 Then GoTo 35
                      'IMPORTE
                      IMPOR1# = PREUN1# * CANAP1#
                      'AGREGO AL ARCHIVO
                      Call REPLA(Y$, MKI$(NPROX%), 1, 2)
                      Call REPLA(Y$, MKI$(CIXI%), 25, 2)
                      Call REPLA(Y$, MKS$(NBORE&), 27, 4)
                      Call REPLA(Y$, MKS$(NUOIT%), 31, 4)
                      Call REPLA(Y$, LO1$, 35, 16)
                      Call REPLA(Y$, LP1$, 51, 16)
                      Call REPLA(Y$, MKS$(NUOC&), 67, 4)
                      Call REPLA(Y$, MKD$(CARE#), 71, 8)
                      Call REPLA(Y$, MKD$(CANAP1#), 79, 8) 'Q.aprobada
                      Call REPLA(Y$, MKD$(QPENFAC#), 87, 8) 'Q.PENDIENTE DE FAC
                      Call REPLA(Y$, MKD$(PREUN1#), 103, 8) 'P.U. O/C
                      Call REPLA(Y$, MKD$(PREUN1#), 111, 8) 'P.U. FACTURADO
                      Call REPLA(Y$, MKD$(IMPOR1#), 135, 8)
                      '
                      J1& = J1& + 1
                      Call GRAREG("!DETFACOM", Y$, J1&)
                    End If
                   End If
                  '
35              Next UK&
              End If
            End If
          Next J&
          Call SETULTREG("!DETFACOM", J1&)
      End With
      '
      If J1& < 1 Then
        Call COMUNI("No se Encontraron Recepeciones\Pendientes de Facturación\para el Set de Recepciones Elegidas.")
      End If
      'AGREGAR AL LIST DE FOFACOBR
      FOFACOBR.List1.Clear
      IMPOTOT# = 0
      For J& = 1 To ULTREG&("!DETFACOM")
        XY$ = REGLEIDO$("!DETFACOM", J&)
        NPRO1 = CVI(Mid$(XY$, 1, 2))
        CIXI1% = CVI(Mid$(XY$, 25, 2))
        '
        NBORE1 = CVS(Mid$(XY$, 27, 4))
        NUOC1& = CVS(Mid$(XY$, 67, 4))
        QFAC1 = CVD(Mid$(XY$, 95, 8))
        QPENFAC1 = CVD(Mid$(XY$, 87, 8))
        PUFAC1 = CVD(Mid$(XY$, 111, 8))
        IMPOR1# = PUFAC1 * QFAC1
        IMPOTOT# = IMPOTOT# + IMPOR1#
        '
        NPRO2$ = FORMATNUM(NPRO1, "F5")
        CIXI2$ = CODEXT(CIXI1%)
        DESCRI2$ = Left$(DESCRIT0(CIXI1%), 15)
        NBORE2$ = FORMATNUM(NBORE1, "F5.0")
        NUOC2$ = FORMATNUM(NUOC1&, "F5.0")
        QPENFAC2$ = FORMATNUM(QPENFAC1, "F10.1")
        QFAC2$ = FORMATNUM(QFAC1, "F10.1")
        PUFAC2$ = FORMATNUM(PUFAC1, "F10.4")
        IMPOR2$ = FORMATNUM(IMPOR1#, "F10.2")
        '
        LI1$ = Space(64)
        '
        Call REPLA(LI1$, CIXI2$, 1, 16)
        Call REPLA(LI1$, DESCRI2$, 17, 10)
        Call REPLA(LI1$, NBORE2$, 29, 5)
        Call REPLA(LI1$, NUOC2$, 36, 5)
        Call REPLA(LI1$, QPENFAC2$, 43, 10)
        Call REPLA(LI1$, QFAC2$, 55, 10)
        Call REPLA(LI1$, PUFAC2$, 67, 10)
        Call REPLA(LI1$, IMPOR2$, 79, 11)
        '
        .List1.AddItem LI1$
      Next J&
      '
      .LBLIMPO = FORMATNUM(IMPOTOT#, "F11.2")
      .Show
    End With
  End If
  '
End Sub
'
Function ULTIPRE$(NPRO%, CIIX%, NUOCO&)
   '
   ' trae el registro de OCOMDETA correspondiente a la
   ' compra no anulada del artículo, proveedor
   ' y que sea = al Nº O/C,
   ' con precio registrado mayor que 0
   '
   FULCO% = 0: UCO$ = REGBLAN$("OCOMDETA")
   '
   RFF$ = RECFILT$("OCOMDETA", 5, MKI$(CIIX%))
   If Len(RFF$) >= 4 Then
     For U& = 1 To Len(RFF$) Step 4
       REGO& = CVS(Mid$(RFF$, U&, 4))
       YX$ = REGLEIDO$("OCOMDETA", REGO&)
       If CVI(Mid$(YX$, 7, 2)) = NPRO% Then
         If CVI(Mid$(YX$, 9, 2)) = CIIX% Then
          'If CVI(Mid$(YX$, 1, 4)) = NOUCO& Then   ' WALTER  !!!!
          If CVS(Mid$(YX$, 1, 4)) = NUOCO& Then
            PREUNRE# = CVD(Mid$(YX$, 67, 8))
            If PREUNRE# > 0.00005 Or FULCO% = 0 Then
              If CVI(Mid$(YX$, 5, 2)) >= FULCO% Then
                UCO$ = YX$
                FULCO% = CVI(Mid$(YX$, 5, 2))
              End If
            End If
          End If
         End If
       End If
     Next U&
   End If
   '
   ULTIPRE$ = UCO$
   '
End Function
