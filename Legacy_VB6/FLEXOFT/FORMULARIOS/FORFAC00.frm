VERSION 5.00
Begin VB.Form FORFAC00 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturación de Proveedores (Operaciones Comerciales de Compras)"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
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
         TabIndex        =   113
         Top             =   4200
         Width           =   3000
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
            TabIndex        =   122
            Top             =   1200
            Width           =   180
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
            TabIndex        =   121
            TabStop         =   0   'False
            Top             =   1200
            Width           =   960
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
            TabIndex        =   120
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1200
            Width           =   1275
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
            Height          =   300
            Left            =   2625
            TabIndex        =   115
            Top             =   280
            Width           =   180
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
            TabIndex        =   119
            Top             =   280
            Width           =   645
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
            TabIndex        =   118
            Top             =   940
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
            TabIndex        =   117
            Top             =   940
            Width           =   1485
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
            TabIndex        =   116
            Top             =   345
            Width           =   1485
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
            TabIndex        =   114
            Top             =   620
            Width           =   2610
         End
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
            Index           =   2
            Left            =   1620
            TabIndex        =   53
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
            Index           =   1
            Left            =   1620
            TabIndex        =   50
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
            Height          =   315
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
            Index           =   2
            Left            =   1365
            TabIndex        =   52
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
            Index           =   2
            Left            =   420
            TabIndex        =   51
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
            Index           =   1
            Left            =   1365
            TabIndex        =   49
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
            Index           =   1
            Left            =   420
            TabIndex        =   48
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
            Height          =   300
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
            Height          =   300
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
            Index           =   2
            Left            =   -1260
            TabIndex        =   56
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
            Index           =   1
            Left            =   -1260
            TabIndex        =   54
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
            Index           =   0
            Left            =   -1260
            TabIndex        =   44
            Top             =   360
            Width           =   1590
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Frame12"
         Height          =   1485
         Left            =   315
         TabIndex        =   100
         Top             =   4305
         Visible         =   0   'False
         Width           =   3060
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
            Index           =   3
            Left            =   1680
            TabIndex        =   109
            Text            =   " "
            Top             =   0
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
            Index           =   3
            Left            =   2625
            TabIndex        =   108
            Top             =   0
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
            Left            =   1680
            TabIndex        =   107
            Text            =   " "
            Top             =   420
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
            Left            =   2625
            TabIndex        =   106
            Top             =   420
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
            Left            =   1680
            TabIndex        =   105
            Text            =   " "
            Top             =   840
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
            Index           =   5
            Left            =   2625
            TabIndex        =   104
            Top             =   840
            Width           =   180
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
            Index           =   3
            Left            =   2885
            TabIndex        =   103
            Text            =   " "
            Top             =   0
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
            Left            =   2885
            TabIndex        =   102
            Text            =   " "
            Top             =   420
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
            Index           =   5
            Left            =   2885
            TabIndex        =   101
            Text            =   " "
            Top             =   840
            Width           =   1485
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
            Left            =   525
            TabIndex        =   112
            Top             =   420
            Visible         =   0   'False
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
            Left            =   0
            TabIndex        =   111
            Top             =   465
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
            Index           =   5
            Left            =   0
            TabIndex        =   110
            Top             =   885
            Width           =   1590
         End
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
         TabIndex        =   99
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
         TabIndex        =   67
         Top             =   4725
         Width           =   3945
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
            TabIndex        =   93
            Text            =   " "
            ToolTipText     =   "Ingrese Importe Percepción I.B."
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
            TabIndex        =   89
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
            TabIndex        =   91
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
            TabIndex        =   72
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
            TabIndex        =   71
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
            TabIndex        =   70
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
         TabIndex        =   66
         Top             =   3885
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
            TabIndex        =   87
            ToolTipText     =   "Ingrese Importe Percepción Ganancias"
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
            TabIndex        =   85
            ToolTipText     =   "Ingrese Importe percepción I.V.A."
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
            TabIndex        =   69
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
            TabIndex        =   68
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
         Picture         =   "FORFAC00.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   62
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
         TabIndex        =   60
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
         Height          =   2370
         Left            =   6825
         TabIndex        =   42
         Top             =   1470
         Width           =   3945
         Begin VB.TextBox IMPOIVA 
            BackColor       =   &H00E0FFD0&
            BorderStyle     =   0  'None
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
            Left            =   2520
            TabIndex        =   95
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   1640
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H00E0FFD0&
            BorderStyle     =   0  'None
            Caption         =   "Frame11"
            Height          =   1950
            Left            =   40
            TabIndex        =   86
            Top             =   360
            Width           =   850
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
               TabIndex        =   97
               Top             =   1575
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
               TabIndex        =   96
               Top             =   1260
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
               TabIndex        =   94
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
               TabIndex        =   92
               Top             =   315
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
               TabIndex        =   90
               Top             =   630
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
               TabIndex        =   88
               Top             =   945
               Width           =   750
            End
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
            Height          =   330
            Index           =   5
            Left            =   2500
            TabIndex        =   83
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   1890
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
            Height          =   330
            Index           =   5
            Left            =   930
            TabIndex        =   78
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   1890
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
            Height          =   330
            Index           =   4
            Left            =   930
            TabIndex        =   77
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   1575
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
            Height          =   360
            Index           =   3
            Left            =   2500
            TabIndex        =   82
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   1260
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
            Height          =   360
            Index           =   3
            Left            =   930
            TabIndex        =   76
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   1260
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
            Height          =   330
            Index           =   2
            Left            =   2500
            TabIndex        =   81
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   945
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
            Height          =   330
            Index           =   2
            Left            =   930
            TabIndex        =   75
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   945
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
            Height          =   330
            Index           =   1
            Left            =   2500
            TabIndex        =   80
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   630
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
            Height          =   360
            Index           =   1
            Left            =   930
            TabIndex        =   74
            ToolTipText     =   "  "
            Top             =   630
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
            Height          =   330
            Index           =   0
            Left            =   2500
            TabIndex        =   79
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
            Height          =   330
            Left            =   930
            TabIndex        =   84
            Top             =   315
            Width           =   1485
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0FFD0&
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   435
         Left            =   5670
         TabIndex        =   38
         Top             =   630
         Width           =   5055
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFD0&
            Caption         =   "CO"
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
            Index           =   5
            Left            =   105
            TabIndex        =   73
            ToolTipText     =   "Comprobante Genérico (Extracto Bancario - Póliza de Seguro - Liquidación)"
            Top             =   105
            Width           =   705
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFD0&
            Caption         =   "RH"
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
            Index           =   4
            Left            =   945
            TabIndex        =   59
            ToolTipText     =   "Recibo de Honorarios"
            Top             =   105
            Width           =   700
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFD0&
            Caption         =   "TF"
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
            Index           =   3
            Left            =   1785
            TabIndex        =   58
            ToolTipText     =   "Ticket - Factura"
            Top             =   105
            Width           =   700
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFD0&
            Caption         =   "NC"
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
            Index           =   2
            Left            =   4305
            TabIndex        =   41
            ToolTipText     =   "Nota de Crédito"
            Top             =   105
            Width           =   700
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFD0&
            Caption         =   "ND"
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
            Index           =   1
            Left            =   3465
            TabIndex        =   40
            ToolTipText     =   "Nota de Débito"
            Top             =   105
            Width           =   700
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
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
            Left            =   2625
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
         Picture         =   "FORFAC00.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Anular, Consultar, Imputación de Créditos, Tablas"
         Top             =   6740
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E0E0E0&
         Height          =   2535
         Left            =   210
         ScaleHeight     =   2475
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
               Height          =   22335
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
               TabIndex        =   57
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
         Picture         =   "FORFAC00.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Imprime Formulario de Carga - Vuelco de Pantalla"
         Top             =   7490
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Proveed"
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
         Picture         =   "FORFAC00.frx":0BF6
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Archivo Maestro de Proveedoresta de Proveedor"
         Top             =   1580
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
         Picture         =   "FORFAC00.frx":0F00
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Resumen de Cuenta Corriente Proveedores"
         Top             =   2340
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
         Picture         =   "FORFAC00.frx":120A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Notas de Proveedores"
         Top             =   3105
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
         Height          =   1190
         Left            =   10920
         Picture         =   "FORFAC00.frx":1514
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
         Height          =   990
         Left            =   10920
         Picture         =   "FORFAC00.frx":181E
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Grabar - Registrar Documento en la Base de Datos"
         Top             =   4830
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
         Left            =   7875
         TabIndex        =   3
         ToolTipText     =   "Ingrese Número de Comprobante"
         Top             =   1050
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
      Begin VB.Line Line21 
         X1              =   10740
         X2              =   12000
         Y1              =   4120
         Y2              =   4120
      End
      Begin VB.Line Line13 
         X1              =   10710
         X2              =   12000
         Y1              =   4630
         Y2              =   4630
      End
      Begin VB.Image Image2 
         Height          =   510
         Left            =   10370
         Picture         =   "FORFAC00.frx":1B28
         Top             =   4140
         Width           =   2010
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00808080&
         X1              =   210
         X2              =   6670
         Y1              =   4120
         Y2              =   4120
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00808080&
         X1              =   6670
         X2              =   6670
         Y1              =   1575
         Y2              =   4120
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
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   5085
         TabIndex        =   98
         ToolTipText     =   "Click Para Cambiar Categoria IVA"
         Top             =   210
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
         TabIndex        =   61
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
         Left            =   4830
         TabIndex        =   15
         Top             =   1050
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
      TabIndex        =   63
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
      TabIndex        =   65
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
      TabIndex        =   64
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
Attribute VB_Name = "FORFAC00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LIST2ANT$
Dim TOBRUFAC#        ' total bruto factura
'
Dim MODORPA%
Dim MODOSORVAL%
Dim TIDOCOM$, LETRAFA$, NUDOCUCOM$
Dim TOFAC#(10), TOFAC1#(10), FVII%(10), IMV#(10), IMV1#(10), GRU%(10), GRU1%(10), IPAR#(10), IPAR1#(10)
'
Dim TOTIMPU#, TONEGRA#, TOEXEN#, TOTIVA#, TORETIV#, TOTFAC#, ALICUIVA
'
Private Sub BOTFEVEN_Click(Index As Integer)
   Call SELECHO("SELFECHA")
   FEVEN(Index).TEXT = VALACT1$("SELFECHA")
   FEVEN(Index).Refresh
End Sub

Private Sub BOTNEXT_Click()
    '
    Call GRAFACOM
    MODIFIX% = 0
    '
End Sub
'

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

Private Sub Command12_Click()
   Command12.Enabled = False
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
   Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   Call SELECHO("SELFECHA")
   Text15 = VALACT1$("SELFECHA")
   Text15.Refresh
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Hide
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
   For KKI% = 1 To 5
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
       Call CALTOFAC
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
         Text10.SetFocus
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

Private Sub TEXT12_LostFocus()
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

Private Sub Text13_GotFocus()
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
      Text10.SetFocus
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
      Text12.SetFocus
   End If
End Sub


Private Sub Command11_Click()
    Command11.Enabled = False
    Call CIERRARCH("*.*")
    Call CONECRUN("ARCH_GES/AMAPRO", "Archivos Maestros")
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
          Screen.MousePointer = 1
          Hide
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
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Call CENTRAFORM(Me)
    Refresh
    DoEvents
    '
    For IXX% = 0 To 3
      TXXX$ = TRIM$(Str$(IXX%))
      On Error Resume Next
      TASA = Val(CONTROL$("", "TASIVA-" + TXXX$))
      On Error GoTo 0
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
    On Error Resume Next
    Text1.SetFocus
    On Error GoTo 0
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
   SUMAVEN# = Val(Label10.Caption)
   For KKI% = 0 To Index - 1
      SUMAVEN# = SUMAVEN# + Val(IVENCI(KKI%).TEXT)
   Next KKI%
   IPAVEN# = Val(IVENCI(Index).TEXT)
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

Private Sub List1_DblClick()
    '
    Call SIGUEFACO2
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
   If Option1(0) = True Then
        Label6.Caption = "Factura Nro:"
        TIDOCOM$ = "FC"
      ElseIf Option1(1) = True Then
        Label6.Caption = "N.Débito Nro:"
        TIDOCOM$ = "ND"
      ElseIf Option1(2) = True Then
        Label6.Caption = "N.Crédito Nro:"
        TIDOCOM$ = "NC"
      ElseIf Option1(3) = True Then
        Label6.Caption = "Ticket-Factura:"
        TIDOCOM$ = "TF"
      ElseIf Option1(4) = True Then
        Label6.Caption = "R.Honorarios:"
        TIDOCOM$ = "RH"
      ElseIf Option1(5) = True Then
        Label6.Caption = "Comprobante:"
        TIDOCOM$ = "CO"
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
    Call BLAFACOM
    '
    If Val(Text1.TEXT) > 32767 Then
      Text1.TEXT = Left$(Text1.TEXT, Len(Text1.TEXT) - 1)
      Text1.SelStart = Len(Text1.TEXT)
    End If
    '
End Sub

Private Sub Text1_DBLCLICK()
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
        Refresh
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
       Call SIGUEFACO1
    End If
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

Private Sub Text6_GotFocus()
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
   TOBRUFAC# = Val(ITOGRA.Caption)
   For KKI% = 0 To 5
     TOBRUFAC# = TOBRUFAC# + (Int(100 * Val(IMPOIVA(KKI%).TEXT) + 0.5)) / 100
     If KKI% > 0 Then
        TOBRUFAC# = TOBRUFAC# + (Int(100 * Val(IMPOGRA(KKI%).TEXT) + 0.5)) / 100
     End If
   Next KKI%
   TOBRUFAC# = TOBRUFAC# + (Int(100 * Val(Text7.TEXT) + 0.5)) / 100
   TOBRUFAC# = TOBRUFAC# + (Int(100 * Val(Text8.TEXT) + 0.5)) / 100
   TOBRUFAC# = TOBRUFAC# + (Int(100 * Val(Text12.TEXT) + 0.5)) / 100
   IVENCI(0).TEXT = CSTRING$(MKD$(TOBRUFAC#), 4, 11, 2, 2)
   IVENCI(0).Refresh
   HOII% = HOY(HOS$)
   FEVEN(0).TEXT = HOS$
   FEVEN(0).Refresh
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
   List1.Clear
   List2.Clear
   Label4.Caption = ""
   Label8.Caption = ""
   Label10.Caption = ""
   Text11.TEXT = ""
   '
   Label25 = ""
   Label21 = ""
   Text14 = ""
   Text15 = ""
   '
   ITOGRA.Caption = Space$(11): ITOGRA.Refresh
   Text7.TEXT = Space$(7)
   Text8.TEXT = Space$(7)
   Text12.TEXT = Space$(7)
   Text10.TEXT = " "
   Label18.Caption = ""
   For KKI% = 1 To 5: IMPOGRA(KKI%).TEXT = Space$(11): Next KKI%
   For KKI% = 0 To 5: IMPOIVA(KKI%).TEXT = Space$(9): Next KKI%
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
   '
End Sub
'
Sub GRAFACOM()
   '
   NROX$ = TRIM$(Text11.TEXT)
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
   NC% = Val(Text1.TEXT)
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
   CUEMA$ = TRIM$(CUEMADRE$((-1) * NC%))
   If CUEMA$ = "" Or CUECOINT%(CUEMA$) < 1 Then
     Call MENSERR(24, "Falta Cuenta Madre de " + TRIM$(Text2))
     VACON% = 0
     Exit Sub
   End If
   '
   FF% = FECHANUM(Text9.TEXT)
   HOI% = HOY(HO$)
   If FF% < 1 Or FF% > HOI% + 1 Then
       Call MENSERR(24, "Fecha no Válida")
       Text9.SetFocus
       VACON% = 0
       Exit Sub
   End If
   '
   FFRE% = FECHANUM(Text6.TEXT)
   HOI% = HOY(HO$)
   If FFRE% < 1 Or FFRE% > HOI% + 1 Then
       Call MENSERR(24, "Fecha de Emisión no Válida")
       Text6.SetFocus
       VACON% = 0
       Exit Sub
   End If
   '
   If DIENTRE%(FF%, HOI%) > 55 Then
       Call MENSERR(24, "Fecha no Válida")
       Text9.SetFocus
       VACON% = 0
       Exit Sub
   End If
   '
   FECHALIM% = Val(CONTROL$("", "CIEMECOM"))
   If FF% <= FECHALIM% Then
       Call MENSERR(24, "Fecha no Válida.\  \Ese Período Mensual Está Cerrado.")
       Text9.SetFocus
       VACON% = 0
       Exit Sub
   End If
   '
   If FF% < FECHANUM(Text6.TEXT) Then
       Call MENSERR(24, "No Válido Contabilizar en Fecha Anterior\a la de Emisión del Comprobante.")
       Text9.SetFocus
       VACON% = 0
       Exit Sub
   End If
   '
   If Val(ITOGRA.Caption) < (-0.005) Then
      Call MENSERR(24, "Importes no Válidos")
      Call SIGUEFACO3
      Exit Sub
   End If
   '
   If LETRAFA$ <> "A" Then
     If Abs(Val(ITOGRA.Caption)) >= 0.005 Then
       Call MENSERR(24, "Importes no Válidos")
       Call SIGUEFACO3
       Exit Sub
     End If
   End If
   '
   SUIMPUTA# = 0
   KKI% = 0
   For U& = 1 To List1.ListCount
      List1.ListIndex = U& - 1
      Z$ = List1.TEXT
      CUE% = Val(Left$(Z$, 5))
      IPA# = Val(Mid$(Z$, 40, 12))
      '
      If CUE% > 0 Then
         If IPA# >= 0.005 Then
            KKI% = KKI% + 1
            GRU%(KKI%) = CUE%
            IPAR#(KKI%) = IPA#
            SUIMPUTA# = SUIMPUTA# + IPA#
         End If
      End If
   Next U&
   If KKI% > 8 Then
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
   For KKI% = 0 To 6
     If KKI% = 4 Then
         IMPOX# = Val(IMPOIVA(5))
       ElseIf KKI% = 5 Then
         IMPOX# = Val(Text7)
       ElseIf KKI% = 6 Then
         IMPOX# = Val(Text8)
       Else
         IMPOX# = Val(IMPOIVA(KKI%))
     End If
     '
     If IMPOX# >= 0.005 Then
       CUEXTE$ = CONTROL$("", "CUEIVA-" + TRIM$(Str$(KKI%)))
       CUEINTE% = CUECOINT%(CUEXTE$)
       If TRIM$(CUEXTE$) = "" Or CUEINTE% < 1 Then
         Call MENSERR(24, "Imposible Registrar.\Faltan Cuentas Contables IVA.")
         Exit Sub
       End If
     End If
   Next KKI%
   '
   If TRIM$(Text10) <> "" Then
     If Val(Text12) >= 0.005 Then
       CUEXTE$ = CUEIBRUPRO$(Text10)
       CUEINTE% = CUECOINT%(CUEXTE$)
       If TRIM$(CUEXTE$) = "" Or CUEINTE% < 1 Then
         Call MENSERR(24, "Imposible Registrar.\Falta Cuenta Contable Correspondiente\a Percepc.Ing.Brutos " + TRIM$(Label18) + ".")
         Exit Sub
       End If
     End If
   End If
   '
   SUNETO# = Val(ITOGRA.Caption)
   SUBRUTO# = SUNETO#
   '
   TOFAC#(1) = SIG% * SUNETO#
   TOFAC#(2) = SIG% * Abs(Val(IMPOGRA(4).TEXT))
   TOFAC#(3) = 0                    ' impuestos intermos
   TOFAC#(4) = 0                    ' IVA TASAS GENERALES
   TOFAC#(5) = 0                    ' iva ADUANA
   TOFAC#(6) = SIG% * Abs(Val(Text7.TEXT))
   TOFAC#(7) = SIG% * Abs(Val(Text8.TEXT))
   TOFAC#(8) = SIG% * Abs(Val(Text12.TEXT))
   '
   For KKI% = 0 To 5
     SUBRUTO# = SUBRUTO# + (Int(100 * Val(IMPOIVA(KKI%).TEXT) + 0.5)) / 100
     If KKI% < 5 Then
          TOFAC#(4) = TOFAC#(4) + SIG% * (Int(100 * Val(IMPOIVA(KKI%).TEXT) + 0.5)) / 100
        Else
          TOFAC#(5) = TOFAC#(5) + SIG% * (Int(100 * Val(IMPOIVA(KKI%).TEXT) + 0.5)) / 100
     End If
     If KKI% > 0 Then
        SUNETO# = SUNETO# + (Int(100 * Val(IMPOGRA(KKI%).TEXT) + 0.5)) / 100
        SUBRUTO# = SUBRUTO# + (Int(100 * Val(IMPOGRA(KKI%).TEXT) + 0.5)) / 100
        If KKI% <> 4 Then
          TOFAC#(1) = TOFAC#(1) + SIG% * (Int(100 * Val(IMPOGRA(KKI%).TEXT) + 0.5)) / 100
        End If
     End If
   Next KKI%
   If Abs(SUNETO - SUIMPUTA#) >= 0.005 Then
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
      FVII%(KKI%) = FECHANUM(FEVEN(KKI% - 1).TEXT)
      If FVII%(KKI%) > 0 Then
         IMV#(KKI%) = Val(IVENCI(KKI% - 1).TEXT)
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
   Call BLOQARCH("CAJABANC")
   Call BLOQARCH("PROVED2")
   '
   For U& = 1 To List2.ListCount
      List2.ListIndex = U& - 1
      Z$ = List2.TEXT
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
   '
   Call BLOQARCH("CUECORR")
   Call BLOQARCH("CUECOSS")
   Call BLOQARCH("CAJABANC")
   Call BLOQARCH("APLIPAG")
   Call BLOQARCH("PROVED2")
   Screen.MousePointer = 11
   '
   REFEDOC$ = AJUSTI$(Text2.TEXT, 24) + TRIM$(Text13.TEXT)
   RECUECO& = 1 + ULTREG&("CUECORR")
   If TIDOCOM$ = "NC" Then
     IMV1#(1) = IMV1#(1) + Val(Label10)  ' DESCONTAR IMPORTES APLICADOS
   End If
   If Val(Label25) > 0 Then
     If TIDOCOM$ <> "CO" Then
       For KKI% = 1 To 6: IMV1#(KKI%) = 0: FVII%(KKI%) = 0: Next KKI%
     End If
   End If
   If TIDOCOM$ = "CO" Then
     DOCUACTIVO$ = NUDOCUCOM$
   End If
   Call CUECORR(FF%, (-1) * NC%, TIPO%, VA%, SUC%, NRO&, REFEDOC$, MO$, TOFAC#(), FVII%(), IMV1#(), GRU%(), IPAR#(), 1, FFRE%)
   Call CUECOSS(NUDOCUCOM$)
   '
   If Val(Label25) > 0 Then
     If TIDOCOM$ = "CO" Then
       Call COMUNI("Imposible Cancelación Contado.\Genere Orden de Pago C.Cte.")
       NRECIB$ = TRIM$(Text11)
       GoTo 20
     End If
     Call REPLA(NUDOCUCOM$, "FF", 1, 2)
     For KKI% = 0 To 10: TOFAC1#(KKI%) = (-1) * TOFAC#(KKI%): Next KKI%
     For KKI% = 0 To 10: GRU1%(KKI%) = 0: IPAR1#(KKI%) = 0: Next KKI%
     Call CUECORR(FF%, (-1) * NC%, 14, VA%, SUC%, NRO&, REFEDOC$, MO$, TOFAC1#(), FVII%(), IMV1#(), GRU1%(), IPAR1#(), 1, FFRE%)
     Call CUECOSS(NUDOCUCOM$)
   End If
   '
   ' grabar factura en CAJABANC
   If Len(TRIM$(Text11)) < 18 Then
        NRECIB$ = TRIM$(Text11)
        While Len(NRECIB$) > 12: NRECIB$ = Mid$(NRECIB$, 2): Wend
      Else
        NRECIB$ = Left$(Text11, 5) + Mid$(Text11, 9, 1) + "-" + Mid$(Text11, 14, 5)
   End If
   '
20 NC1% = (-1) * NC%
   For U& = 1 To List1.ListCount
      Z$ = List1.List(U& - 1)
      CUE% = Val(Left$(Z$, 5))
      IPA# = Val(Mid$(Z$, 40, 12))
      IPA1# = IPA#: IPA2# = 0: If TIDOCOM$ = "NC" Then IPA1# = 0: IPA2# = IPA#
      DEMOVX$ = Mid$(Z$, 8, 30)
      '
      If CUE% > 0 Then
        If IPA# >= 0.005 Then
          Z$ = REGBLAN$("CAJABANC")
          Call REPLA(Z$, NRECIB$, 1, 12)
          Call REPLA(Z$, RASOCLI$(NC1%), 13, 48)
          Call REPLA(Z$, MKI$(FF%), 61, 2)
          Call REPLA(Z$, MKI$(FF%), 63, 2)
          Call REPLA(Z$, MKI$(CUE%), 65, 2)
          Call REPLA(Z$, MKI$(NC1%), 67, 2)
          Call REPLA(Z$, Left$(Text11, 2), 69, 2)
          Call REPLA(Z$, TRIM$(DEMOVX$) + " - " + RASOCLI$(NC1%), 71, 40)
          Call REPLA(Z$, MKI$(FF%), 111, 2)
          Call REPLA(Z$, MKI$(SUC%), 115, 2)
          Call REPLA(Z$, MKS$(NRO&), 117, 4)
          Call REPLA(Z$, Chr$(0), 121, 1)
          Call REPLA(Z$, MKD$(IPA#), 145, 8)
          Call REPLA(Z$, MKD$(IPA1#), 153, 8)
          Call REPLA(Z$, MKD$(IPA2#), 161, 8)
          Call REPLA(Z$, RASOCLI$(NC1%), 177, 32)
          Call REPLA(Z$, "", 245, 8)
          Call REPLA(Z$, MKS$(0), 253, 4)
          Call REPLA(Z$, IDACTU$, 257, 6)  ' FECHA - HORA - USUARIO - TERMINAL
          Call REGAPP("CAJABANC", Z$)
        End If
      End If
   Next U&
   '
   ' grabar importes de IVA
   ZZ0$ = REGBLAN$("CAJABANC")
   Call REPLA(ZZ0$, NRECIB$, 1, 12)
   Call REPLA(ZZ0$, RASOCLI$(NC1%), 13, 48)
   Call REPLA(ZZ0$, MKI$(FF%), 61, 2)
   Call REPLA(ZZ0$, MKI$(FF%), 63, 2)
   Call REPLA(ZZ0$, MKI$(NC1%), 67, 2)
   Call REPLA(ZZ0$, Left$(Text11, 2), 69, 2)
   Call REPLA(ZZ0$, TRIM$(Text11) + " - " + RASOCLI$(NC1%), 71, 40)
   Call REPLA(ZZ0$, MKI$(FF%), 111, 2)
   Call REPLA(ZZ0$, MKI$(SUC%), 115, 2)
   Call REPLA(ZZ0$, MKS$(NRO&), 117, 4)
   Call REPLA(ZZ0$, Chr$(0), 121, 1)
   Call REPLA(ZZ0$, RASOCLI$(NC1%), 177, 32)
   Call REPLA(ZZ0$, IDACTU$, 257, 6)  ' FECHA - HORA - USUARIO - TERMINAL
   '
   For KKI% = 0 To 6
     If KKI% = 4 Then
         IMPOX# = Val(IMPOIVA(5))
       ElseIf KKI% = 5 Then
         IMPOX# = Val(Text7)
       ElseIf KKI% = 6 Then
         IMPOX# = Val(Text8)
       Else
         IMPOX# = Val(IMPOIVA(KKI%))
     End If
     '
     If IMPOX# >= 0.005 Then
       CUEXTE$ = CONTROL$("", "CUEIVA-" + TRIM$(Str$(KKI%)))
       CUEINTE% = CUECOINT%(CUEXTE$)
       Z$ = ZZ0$
       IPA1# = IMPOX#: IPA2# = 0: If TIDOCOM$ = "NC" Then IPA1# = 0: IPA2# = IMPOX#
       Call REPLA(Z$, MKI$(CUEINTE%), 65, 2)
       Call REPLA(Z$, MKD$(Abs(IMPOX#)), 145, 8)
       Call REPLA(Z$, MKD$(IPA1#), 153, 8)
       Call REPLA(Z$, MKD$(IPA2#), 161, 8)
       Call REGAPP("CAJABANC", Z$)
     End If
   Next KKI%
   '
   ' grabar percepcion Ingresos Brutos
   If TRIM$(Text10) <> "" Then
     If Val(Text12) >= 0.005 Then
       CUEXTE$ = CUEIBRUPRO$(Text10)
       CUEINTE% = CUECOINT%(CUEXTE$)
       Z$ = ZZ0$
       IMPOX# = Val(Text12)
       IPA1# = IMPOX#: IPA2# = 0: If TIDOCOM$ = "NC" Then IPA1# = 0: IPA2# = IMPOX#
       Call REPLA(Z$, MKI$(CUEINTE%), 65, 2)
       Call REPLA(Z$, MKD$(Abs(IMPOX#)), 145, 8)
       Call REPLA(Z$, MKD$(IPA1#), 153, 8)
       Call REPLA(Z$, MKD$(IPA2#), 161, 8)
       Call REGAPP("CAJABANC", Z$)
     End If
   End If
   '
   ABC$ = CUEMADRE$(-1 * NC%)
   CUEIXX% = CUECOINT%(CUEMA$)
   IPA1# = 0: IPA2# = Abs(TOFAC#(0)): If TIDOCOM$ = "NC" Then IPA1# = Abs(TOFAC#(0)): IPA2# = 0
   Z$ = REGBLAN$("CAJABANC")
   Call REPLA(Z$, NRECIB$, 1, 12)
   Call REPLA(Z$, RASOCLI$(NC1%), 13, 48)
   Call REPLA(Z$, MKI$(FF%), 61, 2)
   Call REPLA(Z$, MKI$(FF%), 63, 2)
   Call REPLA(Z$, MKI$(CUEIXX%), 65, 2)
   Call REPLA(Z$, MKI$(NC1%), 67, 2)
   Call REPLA(Z$, Left$(Text11, 2), 69, 2)
   Call REPLA(Z$, TRIM$(Text11) + " - " + RASOCLI$(NC1%), 71, 40)
   Call REPLA(Z$, MKI$(FF%), 111, 2)
   Call REPLA(Z$, MKI$(SUC%), 115, 2)
   Call REPLA(Z$, MKS$(NRO&), 117, 4)
   Call REPLA(Z$, Chr$(0), 121, 1)
   Call REPLA(Z$, MKD$(Abs(TOFAC#(0))), 145, 8)
   Call REPLA(Z$, MKD$(IPA1#), 153, 8)
   Call REPLA(Z$, MKD$(IPA2#), 161, 8)
   Call REPLA(Z$, RASOCLI$(NC1%), 177, 32)
   Call REPLA(Z$, "CUECORR", 245, 8)
   Call REPLA(Z$, MKS$(URECOBRA&), 253, 4)
   Call REPLA(Z$, IDACTU$, 257, 6)  ' FECHA - HORA - USUARIO - TERMINAL
   Call REGAPP("CAJABANC", Z$)
   '
   ' grabar pago contado
   If TRIM$(Label25) <> "" Then
     If TIDOCOM$ = "CO" Then
       GoTo 85
     End If
     '
     Call REPLA(Z$, "PC", 69, 2)
     Call REPLA(Z$, MKD$(IPA2#), 153, 8)
     Call REPLA(Z$, MKD$(IPA1#), 161, 8)
     Call REGAPP("CAJABANC", Z$)
     '
     FACRE% = FECHANUM(Text15)
     If FACRE% < HOY(HOS$) Then FACRE% = HOY(HOS$)
     BANCO$ = TRIM$(Label21)
     NUCHE$ = TRIM$(Text14)
     CUEIXX% = CUEMEPAG%(Val(Label25))
     Call REPLA(Z$, MKI$(FACRE%), 63, 2) ' fecha acreditación valor
     FEX = FACRE%: FEVECHE$ = FECHATEX$(FEX)
     IDEVA$ = AJUSTI$(BANCO$, 18) + AJUSTI$(NUCHE$, 14) + FECHATEX$(FEX)
     Call REPLA(Z$, MKI$(CUEIXX%), 65, 2)
     Call REPLA(Z$, "CP", 69, 2)
     Call REPLA(Z$, IDEVA$, 71, 40)
     Call REPLA(Z$, MKD$(IPA1#), 153, 8)
     Call REPLA(Z$, MKD$(IPA2#), 161, 8)
     If TRIM$(NUCHE$) <> "" Then
       Call REPLA(Z$, TRIM$(BANCO$) + " " + TRIM$(NUCHE$) + " " + TRIM$(FEVECHE$), 177, 32)
     End If
     Call REGAPP("CAJABANC", Z$)
     '
   End If
   '
80 ' *********************************** aqui blanquear facturas canceladas
   '
85 NC1% = (-1) * NC%
   VACON% = 1
   '
   XX% = 0
   For U& = 1 To List2.ListCount
      List2.ListIndex = U& - 1
      Z$ = List2.TEXT
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
   For U& = 1 To List2.ListCount
      List2.ListIndex = U& - 1
      Z$ = List2.TEXT
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
                  Call APLIPAG(FC%, NC1%, 4, 1, SUC%, NRO&, CANCE#, REFAPLI$, DOCORI$, ITOTORI#, IVAORI#)
                  '
               End If
            End If
         End If
      End If
   Next U&
   '
   OKX% = 4       ' grabó aplicaciones
   '
84 Call CIERRARCH("APLIPAG")
   Call LIBARCH("APLIPAG")
   Call CIERRARCH("CUECOSS")
   Call LIBARCH("CUECOSS")
   Call CIERRARCH("PROVED2")
   Call LIBARCH("PROVED2")
   Call CIERRARCH("CUECORR")
   Call LIBARCH("CUECORR")
   Call CIERRARCH("CAJABANC")
   Call LIBARCH("CAJABANC")
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
   List1.Clear
   List2.Clear
   Label4.Caption = ""
   Label8.Caption = ""
   Label10.Caption = ""
   Text11.TEXT = ""
   ITOGRA.Caption = Space$(11)
   'TOEXEN.TEXT = ""
   'TOTIVA.TEXT = ""
   'TORETIV.TEXT = ""
   'TOTFAC.Caption = ""
   For KKI% = 0 To 5
      FEVEN(KKI%).TEXT = ""
      IVENCI(KKI%).TEXT = ""
   Next KKI%
   Call BLAFACOM
   VACON% = 0
   Screen.MousePointer = 1
   Text1.SetFocus
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
End Sub
'
Sub SIGUEFACO1()
    '
    Dim NRO$(10)
    Command12.Enabled = True
    Label25.Enabled = True
    If TIDOCOM$ = "NC" Then
      Command12.Enabled = False
      Label25.Enabled = False
    End If
    '
    Screen.MousePointer = 11
    '
    NCXI = Val(Text1.TEXT)
    If NCXI < 1 Or NCXI > 32767 Then
        Call MENSERR(24, "Numero de Cuenta no Válido")
        Screen.MousePointer = 1
        Text1.SetFocus
        Exit Sub
    End If
    '
    NC% = NCXI
    If ECOARCH$("ACREDOR", MKI$(NC%)) = "" Then
        Call MENSERR(24, "Cuenta no Figura\en Base de Datos de Acreedores")
        Screen.MousePointer = 1
        Text1.SetFocus
        Exit Sub
    End If
    '
    FE$ = TRIM$(Text9.TEXT)
    FF% = FECHANUM(FE$)
    HOI% = HOY(HO$)
    If FF% < 1 Or FF% > HOI% + 1 Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Fecha no Válida")
        Text9.SetFocus
        Exit Sub
    End If
    '
    If DIENTRE%(FF%, HOI%) > 55 Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Fecha no Válida")
        Text9.SetFocus
        Exit Sub
    End If
   '
    Call CIERRARCH("*.*")
    '
    Screen.MousePointer = 11
    NC% = Val(Text1.TEXT)
    If NC% > 0 Then
      If Left$(RASOCLI$((-1) * NC%), 6) <> "VARIOS" Then
        Text2.TEXT = RASOCLI$((-1) * NC%)
        Text3.TEXT = DOMICLI$((-1) * NC%)
        Text4.TEXT = CPOSCLI$((-1) * NC%)
        Text5.TEXT = LOCACLI$((-1) * NC%)
        Refresh
      End If
    End If
    '
    CUITT$ = CUITCLI$((-1) * NC%)
    If VALICUIT%(CUITT$) <> 1 Then
      If PIVACLI%((-1) * NC%) <> 5 Then
        Call MENSERR(24, "Imposible Procesar Documento.\Número de C.U.I.T. no Válido")
        Text1.SetFocus
        Screen.MousePointer = 1
        Exit Sub
      End If
    End If
    '
    If TRIM$(Text1.TEXT) = "" Then
       Screen.MousePointer = 1
       Text11.SetFocus
       Exit Sub
    End If
    '
    LETRAFA$ = Label20.Caption
    If TIDOCOM$ = "CO" Then
      If Len(TRIM$(Text11)) < 4 Then
        Screen.MousePointer = 1
        Text11.SetFocus
        Exit Sub
      End If
    End If
    '
    If TIDOCOM$ <> "CO" Then    ' evalua numero ingresado
      NROX$ = TRIM$(Text11.TEXT)
      While Len(NROX$) > 0
         If Asc(Left$(NROX$, 1)) < 48 Or Asc(Left$(NROX$, 1)) > 57 Then
              NROX$ = Mid$(NROX$, 2)
            Else
              GoTo 101
         End If
      Wend
      '
101   If NROX$ = "" Then
         Screen.MousePointer = 1
         Text11.SetFocus
         Exit Sub
      End If
      '
      SUC% = 0
      C1% = 1: For C% = 1 To 10:  NRO$(C%) = "": Next C%
      '
      For C% = 1 To Len(NROX$)
         C2% = Asc(Mid$(NROX$, C%, 1))
         If C2% = 32 Then GoTo 102
         If (C2% < 48 Or C2% > 57) Then C1% = C1% + 1: GoTo 102
         NRO$(C1%) = NRO$(C1%) + Chr$(C2%)
102   Next C%
      '
      If C1% = 1 Then
         SUC% = 0
         VD# = Val(NRO$(1))
         NRO$(2) = NRO$(1)
         NRO$(1) = "0"
      End If
      '
      If Val(NRO$(2)) < 1 Or Len(NRO$(1)) > 4 Or Len(NRO$(2)) > 8 Then
          Call MENSERR(24, "Numero de Documento no Permitido")
          Screen.MousePointer = 1
          Text11.SetFocus
          Exit Sub
      End If
      '
      While Len(NRO$(1)) < 4: NRO$(1) = "0" + NRO$(1): Wend
      While Len(NRO$(2)) < 8: NRO$(2) = "0" + NRO$(2): Wend
      '
      LETRAFA$ = "C": If PIVACLI%((-1) * NC%) = 1 Then LETRAFA$ = "A"
      NROX$ = TIDOCOM$ + "-" + LETRAFA$ + "-" + NRO$(1) + "-" + NRO$(2)
      Text11.TEXT = NROX$
      Text11.Refresh
      Label20.Caption = LETRAFA$
      Label20.Refresh
      '
    End If
    '
    Screen.MousePointer = 11
    NROY1$ = NROX$
    If TIDOCOM$ <> "CO" Then    ' evalua numero ingresado
        Call REPLA(NROY1$, "X", 4, 1)  'para busqueda de repeticion
      Else
        NROY1$ = "CO-" + NROX$
    End If
    NROY$ = AJUSTI$(NROY1$, 18)
    '
    URECU& = ULTREG&("CUECORR")
    RFF$ = RECFILT$("CUECORR", 4, MKI$(NC%))
    For KKL& = 1 To Len(RFF$) Step 4
      RECUEC& = CVS(Mid$(RFF$, KKL&, 4))
      If RECUEC& > 0 Then
        If RECUEC& <= URECU& Then
          NROZ$ = Mid$(REGLEIDO$("CUECORR", RECUEC&), 7, 18)
          If Left$(NROZ$, 2) <> "CO" Then
            Call REPLA(NROZ$, "X", 4, 1)
          End If
          If TRIM$(NROY$) = TRIM$(NROZ$) Then
            Screen.MousePointer = 1
            Call MENSERR(24, "Documento ya Registrado")
            Text1.SetFocus
            Exit Sub
          End If
        End If
      End If
    Next KKL&
    Screen.MousePointer = 1
    '
100 List1.Clear
    List1.Refresh
    List2.Clear
    List2.Refresh
    '
    MAXDEU& = ULTREG&("CUECORR")
    RFD$ = RECFILT$("CUECORR", 4, MKI$(NC%))
    PUCUEN$ = ""
    '
    For U& = 1 To Len(RFD$) Step 4
        REG& = CVS(Mid$(RFD$, U&, 4))
        X$ = REGLEIDO$("CUECORR", REG&)
        If CVI(Mid$(X$, 5, 2)) <> NC% Then
            X$ = REGLEIDO$("CUECORR", 1)
            Call GRAREG("CUECORR", X$, 1)
            Call CIERRARCH("CUECORR")
            GoTo 100
        End If
        '
        NROZ$ = Mid$(X$, 7, 18)
        If Left$(NROZ$, 3) <> "CO-" Then
           Call REPLA(NROZ$, "X", 4, 1)  'para busqueda de repeticion
        End If
        '
        If NROZ$ = NROY$ Then
           Call MENSERR(24, "Documento Pre-existente")
           Text11.TEXT = ""
           Text11.Refresh
           Text11.SetFocus
           Screen.MousePointer = 1
           Exit Sub
        End If
        '
        Y6$ = Mid$(X$, 199, 80)
        For KKI% = 1 To 80 Step 10
          CUEB$ = Mid$(Y6$, KKI%, 2)
          CUEI% = CVI(CUEB$)
          If CUEI% > 0 Then
            For KKJ% = 1 To Len(PUCUEN$) Step 2
              If Mid$(PUCUEN$, KKJ%, 2) = CUEB$ Then GoTo 110
            Next KKJ%
            PUCUEN$ = PUCUEN$ + CUEB$
          End If
110     Next KKI%
    Next U&
    '
    PUCUEX$ = ""
    For KKI% = 1 To Len(PUCUEN$) Step 2
       CUEB$ = Mid$(PUCUEN$, KKI%, 2)
       If ECOARCH$("CUECON", CUEB$) <> "" Then
          PUCUEX$ = PUCUEX$ + CUEB$
          If Len(PUCUEX$) > 16 Then
             PUCUEX$ = Mid$(PUCUEX$, 3)
          End If
       End If
    Next KKI%
    '
    RESCUPRO.LINDICE.Clear
    List1.Clear
    For KKI% = 1 To Len(PUCUEX$) Step 2
       CUEB$ = Mid$(PUCUEX$, KKI%, 2)
       Z$ = Space$(128)
       Call REPLA(Z$, CSTRING$(CUEB$, 1, 5, 0, 2), 1, 5)
       DECUE$ = TRIM$(Left$(ECOARCH$("CUECON", CUEB$), 30))
       'DECUE$ = TRIM$(TRIM$(Mid$(DECUE$, 31, 12)) + " " + TRIM$(Left$(DECUE$, 30)))
       Call REPLA(Z$, DECUE$, 8, 30)
       RESCUPRO.LINDICE.AddItem Z$
    Next KKI%
    '
    For U& = 1 To RESCUPRO.LINDICE.ListCount
       RESCUPRO.LINDICE.ListIndex = U& - 1
       Z$ = RESCUPRO.LINDICE.TEXT
       List1.AddItem Z$
    Next U&
    '
    List1.Refresh
    '
    If TIDOCOM$ = "NC" Then
      RESCUPRO.LINDICE.Clear
      List2.Clear
      For U& = 1 To Len(RFD$) Step 4
        REG& = CVS(Mid$(RFD$, U&, 4))
        X$ = REGLEIDO$("CUECORR", REG&)
        FEMI = CVI(Left$(X$, 2))
        IMPO# = CVD(Mid$(X$, 115, 8)) - CVD(Mid$(X$, 123, 8))
        SALDEU$ = Mid$(X$, 139, 60)
        For KKI% = 1 To 6
           FEVE = CVI(Mid$(SALDEU$, 10 * KKI% - 9, 2))
           SALDO# = CVD(Mid$(SALDEU$, 10 * KKI% - 7, 8))
           If Abs(SALDO#) >= 0.005 Then
              If MODORPA% = 1 Then
                  FF% = FEMI      ' orden emision
                Else
                  FF% = FEVE      ' orden vencimiento
              End If
              Z$ = CSTRING$(MKI$(FF%), 1, 6, 0, 2) + CSTRING$(MKS$(REG&), 4, 8, 0, 2) + CSTRING$(MKI$(KKI%), 1, 4, 0, 2)
              RESCUPRO.LINDICE.AddItem Z$
           End If
        Next KKI%
      Next U&
      '
      SUMADEU# = 0: SUMAFAC# = 0
      For U& = 1 To RESCUPRO.LINDICE.ListCount
        RESCUPRO.LINDICE.ListIndex = U& - 1
        REG& = Val(Mid$(RESCUPRO.LINDICE.TEXT, 7, 8))
        KKI% = Val(Mid$(RESCUPRO.LINDICE.TEXT, 15, 4))
        If REG& > 0 Then
          If REG& <= MAXDEU& Then
            If KKI% > 0 Then
              If KKI% <= 6 Then
                X$ = REGLEIDO$("CUECORR", REG&)
                TIX$ = Mid$(X$, 7, 2)
                SG% = (-1)
                TIPOS$ = "Factura": If TIX$ = "AS" Then TIPOS$ = "Ajuste de Saldo"
                                    If TIX$ = "TF" Then TIPOS$ = "Ticket-Factura"
                                    If TIX$ = "RH" Then TIPOS$ = "Rec.Honorarios"
                                    If TIX$ = "ND" Then TIPOS$ = "Nota de Debito"
                                    If TIX$ = "OP" Then TIPOS$ = "Orden de Pago": SG% = 1
                                    If TIX$ = "FF" Then TIPOS$ = "O/P Fondo Fijo": SG% = 1
                                    If TIX$ = "NC" Then TIPOS$ = "Nota de Credito": SG% = 1
                                    If TIX$ = "DP" Then TIPOS$ = "Debito Propio": SG% = 1
                                    '
                IMPO# = CVD(Mid$(X$, 115, 8)) - CVD(Mid$(X$, 123, 8))
                SALDO# = CVD(Mid$(X$, 139 + 10 * KKI% - 8, 8))
                '
                If TIX$ = "FC" Or TIX$ = "ND" Or TIX$ = "TF" Or TIX$ = "RH" Or (TIX$ = "AS" And IMPO# < 0) Then
                  If SALDO# >= 0.005 Then
                    '
                    SUMAFAC# = SUMAFAC# + IMPO#
                    IMPO# = Abs(IMPO#)
                    SUMADEU# = SUMADEU# + SALDO#
                    '
                    NROX$ = Mid$(X$, 7, 18)
                    FEMI = CVI(Mid$(X$, 1, 2))
                    FEVE = CVI(Mid$(X$, 139 + 10 * KKI% - 10, 2))
                    '
                    IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 13, 2, 2)
                    SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 13, 2, 2)
                    REGTEX$ = CSTRING$(MKS$(REG&), 3, 8, 0, 2)
                    IRETEX$ = CSTRING$(MKI$(KKI%), 1, 2, 0, 2)
                    '
                    Z$ = Space$(128)
                    Call REPLA(Z$, NROX$, 1, 18)
                    Call REPLA(Z$, FECHATEX$(FEMI), 23, 8)
                    Call REPLA(Z$, IMPTEX$, 32, 13)
                    Call REPLA(Z$, FECHATEX$(FEVE), 47, 8)
                    Call REPLA(Z$, SALTEX$, 56, 13)
                    Call REPLA(Z$, REGTEX$, 119, 8)
                    Call REPLA(Z$, IRETEX$, 127, 2)
                    '
                    List2.AddItem Z$
                    '
                  End If
                End If
              End If
            End If
          End If
        End If
      Next U&
    End If
    '
    List2.ListIndex = -1
    List2.Refresh
    Label8.Caption = CSTRING$(MKD$(SUMADEU#), 4, 14, 2, 2)
    Label8.Refresh
    Label10.Caption = ""
    Label10.Refresh
    List2.Visible = True
    If List2.ListCount > 0 Then
        VACON% = 1
        List2.ListIndex = 0
        VACON% = 0
    End If
    '
    Screen.MousePointer = 1
    Call SIGUEFACO2
    '
End Sub
'

Sub SIGUEFACO2()
    '
    Call BLANARCH("!CARICOG")
    For U& = 1 To List1.ListCount
        List1.ListIndex = U& - 1
        Z$ = List1.TEXT
        CUE% = Val(Left$(Z$, 5))
        DENOX$ = TRIM$(Mid$(Z$, 8, 30))
        IPA# = Val(Mid$(Z$, 40, 12))
        '
        Y$ = REGBLAN$("!CARICOG")
        Call REPLA(Y$, MKI$(CUE%), 1, 2)
        Call REPLA(Y$, DENOX$, 3, 38)
        Call REPLA(Y$, MKD$(IPA#), 41, 8)
        Call REGAPP("!CARICOG", Y$)
    Next U&
    Call CIERRARCH("!CARICOG")
    
    '
200 ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Imputacion Contable"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Picture1.Top))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture1.Left))
    '
    VTB% = VENTABU%("CARIFACG" + ATRI$)
    '
    If VTB% >= 0 Then
       RESCUPRO.LINDICE.Clear
       List1.Clear
       TOTIMPU# = 0
       For U& = 1 To ULTREG&("!CARICOG")
          Y$ = REGLEIDO$("!CARICOG", U&)
          CUEB$ = Left$(Y$, 2)
          If CVI(CUEB$) > 0 Then
             Z$ = Space$(128)
             Call REPLA(Z$, CSTRING$(CUEB$, 1, 5, 0, 2), 1, 5)
             DECUE$ = Mid$(Y$, 3, 38) ' ECOARCH$("CUECON", CUEB$)
             'DECUE$ = TRIM$(TRIM$(Mid$(DECUE$, 31, 12)) + " " + TRIM$(Left$(DECUE$, 30)))
             Call REPLA(Z$, DECUE$, 8, 32)
             IMBB$ = Mid$(Y$, 41, 8)
             If CVD(IMBB$) < 0 Then
                Call MENSERR(24, "No se Admiten Valores Negativos")
                GoTo 200
             End If
             Call REPLA(Z$, CSTRING$(IMBB$, 3, 13, 2, 2), 38, 13)
             TOTIMPU# = TOTIMPU# + CVD(IMBB$)
             If CVD(IMBB$) >= 0.005 Then
               RESCUPRO.LINDICE.AddItem Z$
             End If
          End If
       Next U&
       '
       For U& = 1 To RESCUPRO.LINDICE.ListCount
          RESCUPRO.LINDICE.ListIndex = U& - 1
          Z$ = RESCUPRO.LINDICE.TEXT
          List1.AddItem Z$
       Next U&
       '
       For KKI% = 1 To 5
          IMPOGRA(KKI%) = Space$(11)
          IMPOIVA(KKI%) = Space$(8)
       Next KKI%
       '
       If LETRAFA$ = "A" Then
           ITOGRA.Caption = CSTRING$(MKD$(TOTIMPU#), 4, 11, 2, 2)
           ITOGRA.Refresh
         Else
           ITOGRA.Caption = Space$(11)
           IMPOGRA(4).TEXT = CSTRING$(MKD$(TOTIMPU#), 4, 11, 2, 2)
           IMPOGRA(4).Refresh
       End If
    End If
    List1.Refresh
    IMPOGRA(1).SetFocus
    If TOTIMPU# < 0.01 Then
       Text1.SetFocus
    End If
    '
End Sub

Sub SIGUEFACO3()
   '
      For UI% = 1 To 5
        IMPOGRA(UI%).TEXT = Space$(11)
        IMPOIVA(KKI%) = Space$(8)
      Next UI%
      If LETRAFA$ = "A" Then
           ITOGRA.Caption = Space$(11)
           ITOGRA.Caption = CSTRING$(MKD$(TOTIMPU#), 4, 11, 2, 2)
           ITOGRA.Refresh
         Else
           ITOGRA.Caption = Space$(11)
           IMPOGRA(4).TEXT = CSTRING$(MKD$(TOTIMPU#), 4, 11, 2, 2)
           IMPOGRA(4).Refresh
      End If
      TOEXEN# = Val(IMPOGRA(4).TEXT)
      IMPOGRA(1).SetFocus
   TONEGRA# = TOTIMPU# - TOEXEN#
   ITOGRA.Caption = CSTRING$(MKD$(TONEGRA#), 4, 11, 2, 2)
   ITOGRA.Refresh
   '
End Sub


Sub SIGUEFACO5()
    '
    If List2.ListCount < 1 Then
        Exit Sub
    End If
    '
    Call BLANARCH("!CARGAPLI")
    For U& = 1 To List2.ListCount
        List2.ListIndex = U& - 1
        Z$ = List2.TEXT
        NROX$ = Mid$(Z$, 1, 20)
        FEMII% = FECHANUM(Mid$(Z$, 23, 8))
        IPED# = CDbl(Val(Mid$(Z$, 32, 13)))
        FEVEI% = FECHANUM(Mid$(Z$, 47, 8))
        SALD# = CDbl(Val(Mid$(Z$, 56, 13)))
        ACRE# = CDbl(Val(Mid$(Z$, 70, 14)))
        DOCU$ = Mid$(Z$, 1, 18)
        REG2& = Val(Mid$(Z$, 119, 8))
        IRE% = Val(Mid$(Z$, 127, 2))
        '
        Y$ = REGBLAN$("!CARGAPLI")
        Call REPLA(Y$, NROX$, 1, 20)
        Call REPLA(Y$, MKI$(FEMII%), 21, 2)
        Call REPLA(Y$, MKD$(IPED#), 23, 8)
        Call REPLA(Y$, MKI$(FEVEI%), 31, 2)
        Call REPLA(Y$, MKD$(SALD#), 33, 8)
        Call REPLA(Y$, MKD$(ACRE#), 41, 8)
        Call REPLA(Y$, DOCU$, 49, 18)
        Call REPLA(Y$, MKS$(REG2&), 67, 4)
        Call REPLA(Y$, MKI$(IRE%), 71, 2)
        '
        Call REGAPP("!CARGAPLI", Y$)
    Next U&
    Call CIERRARCH("!CARGAPLI")
    '
200 ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Aplicación de Pagos"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Frame8.Top))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture2.Left))
    '
    VTB% = VENTABU%("CARGAPLI" + ATRI$)
    '
    VACON% = 1
    If VTB% >= 0 Then
        List2.Clear
        For U& = 1 To ULTREG&("!CARGAPLI")
            Y$ = REGLEIDO$("!CARGAPLI", U&)
            NROX$ = Mid$(Y$, 1, 20)
            If TRIM$(NROX$) = "Anticipo no Aplicado" Then
               HOI% = HOY(HO$)
               ACRE# = CVD(Mid$(Y$, 41, 8))
               Call REPLA(Y$, MKI$(HOI%), 31, 2)
               Call REPLA(Y$, MKD$(ACRE#), 23, 8)
               Call REPLA(Y$, MKD$(ACRE#), 33, 8)
               Call GRAREG("!CARGAPLI", Y$, U&)
            End If
            FEMI = CVI(Mid$(Y$, 21, 2))
            IPED# = CVD(Mid$(Y$, 23, 8))
            IMPTEX$ = CSTRING$(MKD$(IPED#), 4, 13, 2, 2)
            FEVE = CVI(Mid$(Y$, 31, 2))
            SALD# = CVD(Mid$(Y$, 33, 8))
            SALTEX$ = CSTRING$(MKD$(SALD#), 4, 13, 2, 2)
            ACRE# = CVD(Mid$(Y$, 41, 8))
            '
            If ACRE# < 0 Then
                Call MENSERR(24, "No se Permite Aplicación Negativa")
                GoTo 200
            End If
            If U& > 1 Then
              If ACRE# > SALD# Then
                Call MENSERR(24, "El Importe Acreditado Debe Ser\Inferior al Saldo del Documento")
                 GoTo 200
              End If
            End If
            ACRTEX$ = CSTRING$(MKD$(ACRE#), 4, 14, 2, 2)
            DOCU$ = Mid$(Y$, 49, 18)
            REG2& = CVS(Mid$(Y$, 67, 4))
            IRE% = CVI(Mid$(Y$, 71, 2))
            REGTEX$ = CSTRING$(MKS$(REG2&), 3, 8, 0, 2)
            IRETEX$ = CSTRING$(MKI$(IRE%), 1, 2, 0, 2)
            '
            Z$ = Space$(128)
            Call REPLA(Z$, NROX$, 1, 20)
            Call REPLA(Z$, FECHATEX$(FEMI), 23, 8)
            Call REPLA(Z$, IMPTEX$, 32, 13)
            Call REPLA(Z$, FECHATEX$(FEVE), 47, 8)
            Call REPLA(Z$, SALTEX$, 56, 13)
            Call REPLA(Z$, ACRTEX$, 70, 14)
            Call REPLA(Z$, REGTEX$, 119, 8)
            Call REPLA(Z$, IRETEX$, 127, 2)
            List2.AddItem Z$
            '
        Next U&
    End If
    '
    SUMACRE# = 0
    For U& = 1 To List2.ListCount
        List2.ListIndex = U& - 1
        Z$ = List2.TEXT
        SUMACRE# = SUMACRE# + CDbl(Val(Mid$(Z$, 70, 14)))
    Next U&
    '
    Label10.Caption = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
    Label10.Refresh
    '
    List2.Refresh
    VACON% = 0
    '
End Sub
'


