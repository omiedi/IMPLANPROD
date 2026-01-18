VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form OPCLAV07 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración Resúmenes Mensuales de Ventas"
   ClientHeight    =   5970
   ClientLeft      =   1650
   ClientTop       =   585
   ClientWidth     =   14670
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   14670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Cuentas Contables Asiento Items Importados (Trazables = 1)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   121
      Top             =   5280
      Width           =   13780
      Begin VB.TextBox Text4 
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
         Index           =   22
         Left            =   8715
         TabIndex        =   127
         Text            =   " "
         Top             =   240
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
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
         Index           =   22
         Left            =   9870
         TabIndex        =   126
         Top             =   240
         Width           =   150
      End
      Begin VB.TextBox Text4 
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
         Index           =   21
         Left            =   1635
         TabIndex        =   123
         Text            =   " "
         Top             =   240
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
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
         Index           =   21
         Left            =   2790
         TabIndex        =   122
         Top             =   240
         Width           =   150
      End
      Begin VB.Label Label4 
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
         Index           =   22
         Left            =   10065
         TabIndex        =   129
         Top             =   240
         Width           =   3585
      End
      Begin VB.Label LABEL2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Material Importado"
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
         Index           =   22
         Left            =   6960
         TabIndex        =   128
         Top             =   300
         Width           =   1545
      End
      Begin VB.Label Label4 
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
         Index           =   21
         Left            =   3105
         TabIndex        =   125
         Top             =   240
         Width           =   3585
      End
      Begin VB.Label LABEL2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Materias Primas"
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
         Index           =   21
         Left            =   120
         TabIndex        =   124
         Top             =   300
         Width           =   1320
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   4800
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   75
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      ForeColor       =   &H80000008&
      Height          =   9720
      Left            =   13800
      ScaleHeight     =   9690
      ScaleWidth      =   1875
      TabIndex        =   52
      Top             =   0
      Width           =   1905
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   105
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   106
         Top             =   4200
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   107
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   105
         Picture         =   "OPCLAV07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   "Tabla de Cuentas Monetarias"
         Top             =   2520
         Width           =   630
      End
      Begin VB.CommandButton command1 
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
         Picture         =   "OPCLAV07.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   180
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
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
         Picture         =   "OPCLAV07.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   870
         Width           =   650
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Help"
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
         Picture         =   "OPCLAV07.frx":059E
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   1560
         Width           =   650
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Plan"
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
         Picture         =   "OPCLAV07.frx":08A8
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Archivo Maestro de Cuentas Contables"
         Top             =   3360
         Width           =   650
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E2D3C0&
      Caption         =   "OTROS MOVIMIENTOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3225
      Left            =   7080
      TabIndex        =   45
      Top             =   2010
      Width           =   6705
      Begin VB.CommandButton Command8 
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
         Index           =   20
         Left            =   2790
         TabIndex        =   117
         Top             =   2760
         Width           =   150
      End
      Begin VB.TextBox Text4 
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
         Index           =   20
         Left            =   1635
         TabIndex        =   116
         Text            =   " "
         Top             =   2760
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
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
         Index           =   19
         Left            =   2790
         TabIndex        =   113
         Top             =   2400
         Width           =   150
      End
      Begin VB.TextBox Text4 
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
         Index           =   19
         Left            =   1635
         TabIndex        =   112
         Text            =   " "
         Top             =   2400
         Width           =   1170
      End
      Begin VB.TextBox Text4 
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
         Index           =   18
         Left            =   1635
         TabIndex        =   102
         Text            =   " "
         Top             =   1640
         Width           =   1170
      End
      Begin VB.TextBox Text4 
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
         Index           =   17
         Left            =   1635
         TabIndex        =   98
         Text            =   " "
         Top             =   1280
         Width           =   1170
      End
      Begin VB.TextBox Text4 
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
         Index           =   12
         Left            =   1635
         TabIndex        =   59
         Text            =   " "
         Top             =   615
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
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
         Index           =   18
         Left            =   2790
         TabIndex        =   103
         Top             =   1640
         Width           =   150
      End
      Begin VB.CommandButton Command8 
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
         Index           =   17
         Left            =   2790
         TabIndex        =   99
         Top             =   1280
         Width           =   150
      End
      Begin VB.TextBox Text4 
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
         Index           =   13
         Left            =   1635
         TabIndex        =   68
         Text            =   " "
         Top             =   930
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
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
         Index           =   13
         Left            =   2790
         TabIndex        =   67
         Top             =   930
         Width           =   150
      End
      Begin VB.CommandButton Command8 
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
         Index           =   12
         Left            =   2790
         TabIndex        =   60
         Top             =   615
         Width           =   150
      End
      Begin VB.TextBox Text4 
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
         Index           =   10
         Left            =   1635
         TabIndex        =   47
         Text            =   " "
         Top             =   300
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
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
         Index           =   10
         Left            =   2790
         TabIndex        =   46
         Top             =   300
         Width           =   150
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Cuenas Contables de Asientos Para Items Sin Cargo"
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
         TabIndex        =   120
         Top             =   2130
         Width           =   4815
      End
      Begin VB.Label LABEL2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cta. Contrapartida"
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
         Index           =   20
         Left            =   -600
         TabIndex        =   119
         Top             =   2820
         Width           =   2145
      End
      Begin VB.Label Label4 
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
         Index           =   20
         Left            =   2985
         TabIndex        =   118
         Top             =   2760
         Width           =   3585
      End
      Begin VB.Label LABEL2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bonificaciones"
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
         Index           =   19
         Left            =   -600
         TabIndex        =   115
         Top             =   2460
         Width           =   2145
      End
      Begin VB.Label Label4 
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
         Index           =   19
         Left            =   2985
         TabIndex        =   114
         Top             =   2400
         Width           =   3585
      End
      Begin VB.Label Label4 
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
         Index           =   18
         Left            =   2985
         TabIndex        =   105
         Top             =   1635
         Width           =   3585
      End
      Begin VB.Label LABEL2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Anticipos"
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
         Index           =   18
         Left            =   -600
         TabIndex        =   104
         Top             =   1695
         Width           =   2145
      End
      Begin VB.Label Label4 
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
         Index           =   17
         Left            =   2985
         TabIndex        =   101
         Top             =   1275
         Width           =   3585
      End
      Begin VB.Label LABEL2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Anticipos A Cobrar"
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
         Index           =   17
         Left            =   -600
         TabIndex        =   100
         Top             =   1350
         Width           =   2145
      End
      Begin VB.Label LABEL2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta IVA Ventas:"
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
         Index           =   10
         Left            =   -360
         TabIndex        =   89
         Top             =   360
         Width           =   1905
      End
      Begin VB.Label LABEL2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Gastos Embalaje"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   12
         Left            =   -360
         TabIndex        =   88
         Top             =   660
         Width           =   1905
      End
      Begin VB.Label LABEL2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Gastos Exportacion"
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
         Index           =   13
         Left            =   -600
         TabIndex        =   70
         Top             =   990
         Width           =   2145
      End
      Begin VB.Label Label4 
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
         Index           =   13
         Left            =   2985
         TabIndex        =   69
         Top             =   930
         Width           =   3585
      End
      Begin VB.Label Label4 
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
         Index           =   12
         Left            =   2985
         TabIndex        =   61
         Top             =   615
         Width           =   3585
      End
      Begin VB.Label Label4 
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
         Index           =   10
         Left            =   2985
         TabIndex        =   48
         Top             =   300
         Width           =   3585
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2D3C0&
      Caption         =   "CLASIFICACIÓN DE DÉBITOS Y CRÉDITOS VARIOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   7080
      TabIndex        =   38
      Top             =   50
      Width           =   6705
      Begin VB.TextBox Text4 
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
         Index           =   8
         Left            =   1635
         TabIndex        =   39
         Text            =   " "
         Top             =   300
         Width           =   1170
      End
      Begin VB.TextBox Text4 
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
         Index           =   14
         Left            =   1635
         TabIndex        =   72
         Text            =   " "
         Top             =   1245
         Width           =   1170
      End
      Begin VB.TextBox Text4 
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
         Index           =   9
         Left            =   1635
         TabIndex        =   41
         Text            =   " "
         Top             =   615
         Width           =   1170
      End
      Begin VB.TextBox Text4 
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
         Index           =   16
         Left            =   1635
         TabIndex        =   95
         Text            =   " "
         Top             =   1560
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
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
         Index           =   16
         Left            =   2790
         TabIndex        =   94
         Top             =   1575
         Width           =   150
      End
      Begin VB.CommandButton Command8 
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
         Index           =   14
         Left            =   2790
         TabIndex        =   73
         Top             =   1260
         Width           =   150
      End
      Begin VB.TextBox Text4 
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
         Index           =   11
         Left            =   1635
         TabIndex        =   50
         Text            =   " "
         Top             =   930
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
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
         Index           =   11
         Left            =   2790
         TabIndex        =   49
         Top             =   930
         Width           =   150
      End
      Begin VB.CommandButton Command8 
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
         Index           =   9
         Left            =   2790
         TabIndex        =   42
         Top             =   615
         Width           =   150
      End
      Begin VB.CommandButton Command8 
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
         Index           =   8
         Left            =   2790
         TabIndex        =   40
         Top             =   300
         Width           =   150
      End
      Begin VB.Label LABEL2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Gastos Bancarios"
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
         Index           =   16
         Left            =   -360
         TabIndex        =   97
         Top             =   1605
         Width           =   1905
      End
      Begin VB.Label Label4 
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
         Index           =   16
         Left            =   2985
         TabIndex        =   96
         Top             =   1560
         Width           =   3585
      End
      Begin VB.Label LABEL2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta IVA Ventas:"
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
         Index           =   8
         Left            =   -360
         TabIndex        =   87
         Top             =   360
         Width           =   1905
      End
      Begin VB.Label LABEL2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta IVA Ventas:"
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
         Index           =   9
         Left            =   -360
         TabIndex        =   86
         Top             =   690
         Width           =   1905
      End
      Begin VB.Label LABEL2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta IVA Ventas:"
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
         Index           =   11
         Left            =   -360
         TabIndex        =   85
         Top             =   990
         Width           =   1905
      End
      Begin VB.Label Label4 
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
         Index           =   14
         Left            =   2985
         TabIndex        =   74
         Top             =   1245
         Width           =   3585
      End
      Begin VB.Label LABEL2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Intereses por Mora"
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
         Index           =   15
         Left            =   -360
         TabIndex        =   71
         Top             =   1305
         Width           =   1905
      End
      Begin VB.Label Label4 
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
         Index           =   11
         Left            =   2985
         TabIndex        =   51
         Top             =   930
         Width           =   3585
      End
      Begin VB.Label Label4 
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
         Index           =   9
         Left            =   2985
         TabIndex        =   44
         Top             =   615
         Width           =   3585
      End
      Begin VB.Label Label4 
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
         Index           =   8
         Left            =   2985
         TabIndex        =   43
         Top             =   300
         Width           =   3585
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      Caption         =   "CUENTAS CONTABLES I.V.A."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   0
      TabIndex        =   1
      Top             =   50
      Width           =   6900
      Begin VB.TextBox Text2 
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
         Index           =   4
         Left            =   1635
         TabIndex        =   108
         Text            =   " "
         Top             =   1575
         Width           =   1170
      End
      Begin VB.CommandButton Command4 
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
         Index           =   4
         Left            =   2790
         TabIndex        =   109
         Top             =   1575
         Width           =   150
      End
      Begin VB.TextBox Text2 
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
         Index           =   3
         Left            =   1635
         TabIndex        =   62
         Text            =   " "
         Top             =   615
         Width           =   1170
      End
      Begin VB.CommandButton Command4 
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
         Index           =   3
         Left            =   2790
         TabIndex        =   63
         Top             =   615
         Width           =   150
      End
      Begin VB.TextBox Text100 
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
         Left            =   6150
         TabIndex        =   58
         Text            =   " "
         ToolTipText     =   "Código de Régimen de Percepción"
         Top             =   1245
         Width           =   540
      End
      Begin VB.TextBox Text2 
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
         Left            =   1635
         TabIndex        =   7
         Text            =   " "
         Top             =   300
         Width           =   1170
      End
      Begin VB.CommandButton Command4 
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
         Left            =   2790
         TabIndex        =   6
         Top             =   300
         Width           =   150
      End
      Begin VB.TextBox Text2 
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
         Left            =   1635
         TabIndex        =   5
         Text            =   " "
         Top             =   930
         Width           =   1170
      End
      Begin VB.CommandButton Command4 
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
         Left            =   2790
         TabIndex        =   4
         Top             =   930
         Width           =   150
      End
      Begin VB.TextBox Text2 
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
         Index           =   2
         Left            =   1635
         TabIndex        =   3
         Text            =   " "
         Top             =   1245
         Width           =   1170
      End
      Begin VB.CommandButton Command4 
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
         Index           =   2
         Left            =   2790
         TabIndex        =   2
         Top             =   1245
         Width           =   150
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sin Uso"
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
         Left            =   -330
         TabIndex        =   111
         Top             =   1650
         Width           =   1905
      End
      Begin VB.Label Label3 
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
         Index           =   4
         Left            =   3105
         TabIndex        =   110
         Top             =   1575
         Width           =   3585
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta IVA T.Sec"
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
         Left            =   -360
         TabIndex        =   65
         Top             =   690
         Width           =   1905
      End
      Begin VB.Label Label3 
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
         Index           =   3
         Left            =   3105
         TabIndex        =   64
         Top             =   615
         Width           =   3585
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "C.R."
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
         Left            =   5595
         TabIndex        =   57
         ToolTipText     =   "Código de Régimen de Percepción"
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label3 
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
         Index           =   2
         Left            =   3105
         TabIndex        =   29
         Top             =   1245
         Width           =   2430
      End
      Begin VB.Label Label3 
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
         Index           =   1
         Left            =   3105
         TabIndex        =   28
         Top             =   930
         Width           =   3585
      End
      Begin VB.Label Label3 
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
         Index           =   0
         Left            =   3105
         TabIndex        =   27
         Top             =   300
         Width           =   3585
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta IVA Ventas"
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
         Left            =   -360
         TabIndex        =   10
         Top             =   375
         Width           =   1905
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta IVA R.N.I"
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
         Left            =   -360
         TabIndex        =   9
         Top             =   1005
         Width           =   1905
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Percepción IVA"
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
         Left            =   -360
         TabIndex        =   8
         Top             =   1320
         Width           =   1905
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E2D3C0&
      Caption         =   "CLASIFICACIÓN DE VENTAS SEGÚN POSICION DE IVA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3225
      Left            =   0
      TabIndex        =   0
      Top             =   2010
      Width           =   6900
      Begin VB.TextBox Text4 
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
         Index           =   15
         Left            =   1635
         TabIndex        =   90
         Text            =   " "
         Top             =   2820
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
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
         Index           =   15
         Left            =   2790
         TabIndex        =   91
         Top             =   2820
         Width           =   150
      End
      Begin VB.TextBox Text4 
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
         Index           =   7
         Left            =   1635
         TabIndex        =   26
         Text            =   " "
         Top             =   2505
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
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
         Index           =   7
         Left            =   2790
         TabIndex        =   25
         Top             =   2505
         Width           =   150
      End
      Begin VB.TextBox Text4 
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
         Index           =   6
         Left            =   1635
         TabIndex        =   24
         Text            =   " "
         Top             =   2190
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
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
         Index           =   6
         Left            =   2790
         TabIndex        =   23
         Top             =   2190
         Width           =   150
      End
      Begin VB.TextBox Text4 
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
         Index           =   5
         Left            =   1635
         TabIndex        =   22
         Text            =   " "
         Top             =   1875
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
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
         Index           =   5
         Left            =   2790
         TabIndex        =   21
         Top             =   1875
         Width           =   150
      End
      Begin VB.TextBox Text4 
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
         Index           =   4
         Left            =   1635
         TabIndex        =   20
         Text            =   " "
         Top             =   1560
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
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
         Index           =   4
         Left            =   2790
         TabIndex        =   19
         Top             =   1560
         Width           =   150
      End
      Begin VB.TextBox Text4 
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
         Index           =   2
         Left            =   1635
         TabIndex        =   18
         Text            =   " "
         Top             =   930
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
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
         Index           =   2
         Left            =   2790
         TabIndex        =   17
         Top             =   930
         Width           =   150
      End
      Begin VB.TextBox Text4 
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
         Left            =   1635
         TabIndex        =   16
         Text            =   " "
         Top             =   615
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
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
         Left            =   2790
         TabIndex        =   15
         Top             =   615
         Width           =   150
      End
      Begin VB.TextBox Text4 
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
         Left            =   1635
         TabIndex        =   14
         Text            =   " "
         Top             =   300
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
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
         Left            =   2790
         TabIndex        =   13
         Top             =   300
         Width           =   150
      End
      Begin VB.TextBox Text4 
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
         Index           =   3
         Left            =   1635
         TabIndex        =   12
         Text            =   " "
         Top             =   1245
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
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
         Index           =   3
         Left            =   2790
         TabIndex        =   11
         Top             =   1245
         Width           =   150
      End
      Begin VB.Label Label4 
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
         Index           =   15
         Left            =   3105
         TabIndex        =   93
         Top             =   2820
         Width           =   3585
      End
      Begin VB.Label LABEL2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Reparaciones"
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
         Index           =   14
         Left            =   -360
         TabIndex        =   92
         Top             =   2880
         Width           =   1905
      End
      Begin VB.Label LABEL2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta IVA Ventas:"
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
         Left            =   -360
         TabIndex        =   84
         Top             =   1635
         Width           =   1905
      End
      Begin VB.Label LABEL2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta IVA Ventas:"
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
         Left            =   -360
         TabIndex        =   83
         Top             =   375
         Width           =   1905
      End
      Begin VB.Label LABEL2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta IVA Ventas:"
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
         Left            =   -360
         TabIndex        =   82
         Top             =   690
         Width           =   1905
      End
      Begin VB.Label LABEL2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta IVA Ventas:"
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
         Left            =   -360
         TabIndex        =   81
         Top             =   1005
         Width           =   1905
      End
      Begin VB.Label LABEL2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta IVA Ventas:"
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
         Left            =   -360
         TabIndex        =   80
         Top             =   1320
         Width           =   1905
      End
      Begin VB.Label LABEL2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta IVA Ventas:"
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
         Left            =   -360
         TabIndex        =   79
         Top             =   1950
         Width           =   1905
      End
      Begin VB.Label LABEL2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta IVA Ventas:"
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
         Left            =   -360
         TabIndex        =   78
         Top             =   2265
         Width           =   1905
      End
      Begin VB.Label LABEL2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta IVA Ventas:"
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
         Left            =   -360
         TabIndex        =   77
         Top             =   2580
         Width           =   1905
      End
      Begin VB.Label Label4 
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
         Index           =   7
         Left            =   3105
         TabIndex        =   37
         Top             =   2505
         Width           =   3585
      End
      Begin VB.Label Label4 
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
         Index           =   6
         Left            =   3105
         TabIndex        =   36
         Top             =   2190
         Width           =   3585
      End
      Begin VB.Label Label4 
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
         Index           =   5
         Left            =   3105
         TabIndex        =   35
         Top             =   1875
         Width           =   3585
      End
      Begin VB.Label Label4 
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
         Index           =   4
         Left            =   3105
         TabIndex        =   34
         Top             =   1560
         Width           =   3585
      End
      Begin VB.Label Label4 
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
         Index           =   2
         Left            =   3105
         TabIndex        =   33
         Top             =   930
         Width           =   3585
      End
      Begin VB.Label Label4 
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
         Index           =   1
         Left            =   3105
         TabIndex        =   32
         Top             =   615
         Width           =   3585
      End
      Begin VB.Label Label4 
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
         Index           =   0
         Left            =   3105
         TabIndex        =   31
         Top             =   300
         Width           =   3585
      End
      Begin VB.Label Label4 
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
         Index           =   3
         Left            =   3105
         TabIndex        =   30
         Top             =   1245
         Width           =   3585
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "OPCLAV07.frx":09F2
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "OPCLAV07.frx":0C26
      TabIndex        =   76
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Menu MnuArchivo 
      Caption         =   "Archivo"
   End
   Begin VB.Menu mnuTAb 
      Caption         =   "Tablas"
      Begin VB.Menu mnuCtasPosFac 
         Caption         =   "Configurar Tabla de Imputaciones de Ventas"
      End
      Begin VB.Menu ImpuConPorGrupoVentas 
         Caption         =   "Imputacion de Ventas por Grupo Contable"
      End
   End
   Begin VB.Menu mnuConfiguracion 
      Caption         =   "Configuración"
      Begin VB.Menu mnuConfiCashflow 
         Caption         =   "Configuración Cashflow"
      End
   End
End
Attribute VB_Name = "OPCLAV07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Command6_Click()
   Command6.Enabled = False
   '
   Call COPYSTRU("TACUECAJ", "TACUMONE")
   '
   JJ& = 0
   NUCUEN% = ULTREG&("CUENTAS")
   Screen.MousePointer = 11
   For U& = 1 To ULTREG&("TACUMONE")
      XX$ = REGLEIDO$("TACUMONE", U&)
      CXI% = CUEINT%(TRIM$(Left$(XX$, 12)))
      If CXI% > 0 Then
         If CXI% <= NUCUEN% Then
            JJ& = JJ& + 1
            Call GRAREG("!TACUMONE", XX$, JJ&)
         End If
      End If
   Next U&
   Call SETULTREG("!TACUMONE", JJ&)
   Call CIERRARCH("!TACUMONE")
   Call CIERRARCH("TACUMONE")
   Screen.MousePointer = 1
   '
   VTB% = VENTABU%("TACUMONE")
   If VTB% >= 0 Then
      Screen.MousePointer = 11
      JJ& = 0
      For U& = 1 To ULTREG&("!TACUMONE")
         XX$ = REGLEIDO$("!TACUMONE", U&)
         CXI% = CUEINT%(TRIM$(Left$(XX$, 12)))
         If CXI% > 0 Then
            If CXI% <= NUCUEN% Then
               JJ& = JJ& + 1
               Call GRAREG("TACUMONE", XX$, JJ&)
            End If
         End If
      Next U&
      Call SETULTREG("TACUMONE", JJ&)
      Call CIERRARCH("TACUMONE")
      Call FRESHECHO("TACUMONE")
      Call CIERRARCH("!TACUMONE")
      Screen.MousePointer = 1
   End If
   '
   Command6.Enabled = True
   
End Sub

Private Sub Command8_Click(Index As Integer)
    Call SELECHO("ECUECON")
    If VALACT1$("ECUECON") <> "" Then
        Text4(Index).TEXT = VALACT1$("ECUECON")
        Label4(Index) = VALACT2$("ECUECON")
    End If
End Sub

Private Sub Form_Load()
   '
   CUEIVA$ = Control$("SVD451S", "CUEIVA")
   Text2(0) = CUEIVA$
   Label3(0) = ECOARCH$("ECUECON", CUEIVA$)
   '
   CUEIVB$ = Control$("SVD451S", "CUEIVB")
   Text2(3) = CUEIVB$
   Label3(3) = ECOARCH$("ECUECON", CUEIVB$)
   '
   CUEADU$ = Control$("SVD451S", "CUEADU")
   Text2(1) = CUEADU$
   Label3(1) = ECOARCH$("ECUECON", CUEADU$)
   '
   CUERET$ = Control$("SVD451S", "CUERET")
   Text2(2) = CUERET$
   Label3(2) = ECOARCH$("ECUECON", CUERET$)
   '
   CUEREG$ = Control$("SVD451S", "CUEREG")
   Text2(4) = CUEREG$
   Label3(4) = ECOARCH$("ECUECON", CUEREG$)
   '
   Label2(0) = "Consumidor Final"
   
   Label2(1) = "IVA Resp.Inscripto"
   Label2(2) = "Resp. No Inscripto"
   Label2(3) = "No Responsable IVA"
   Label2(4) = "IVA Exento"
   Label2(5) = "Comercio Exterior"
   Label2(6) = "Monotributo"
   Label2(7) = "Zona Promovida"
   '
   Label2(8) = "Cheques Rechazados"
   Label2(9) = "Diferencia de Cambio"
   Label2(10) = "Ajustes de Saldo"
   Label2(11) = "Otros Deb/Creditos"
   Label2(14) = "Reparaciones"
   '
   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me, Picture1)
   For KKI% = 0 To 18
     Text4(KKI%).TEXT = Control$("SVD451S", "CUEVE-" + TRIM$(Str$(KKI%)))
   Next KKI%

   Text4(19).TEXT = Control("ASVENTAS", "ASITEMSC") 'CUENTA  DE ASIENTO DE VENTAS ITEMS SIN CARGO
   Text4(20).TEXT = Control("ASVENTAS", "CONTASSC") 'CUENTA DE CONTRAPARTIDA DE ASIENTO DE VENTAS ITEMS SIN CARGO
   
   Text4(21).TEXT = Control("ASVENTIM", "ASIMAPRI") 'CUENTA  DE ASIENTO DE VENTAS MATERIAS PRIMAS
   Text4(22).TEXT = Control("ASVENTIM", "CONTIMPO") 'CUENTA DE CONTRAPARTIDA DE IMPORTADOS
   
   Screen.MousePointer = 1
   '
   '\\\OT-06-0158-WAR-30/05/06///
   Text100.TEXT = Control$("COREPER", "COREPER")
   '\\\OT-06-0158-WAR-FIN///
   '
End Sub

Private Sub Command2_Click()
    '
    CUEIVA$ = AJUSTI$(Text2(0).TEXT, 12)
    CUEIVB$ = AJUSTI$(Text2(3).TEXT, 12)
    CUEADU$ = AJUSTI$(Text2(1).TEXT, 12)
    CUERET$ = AJUSTI$(Text2(2).TEXT, 12)
    'CUEREG$ = AJUSTI$(Text2(4).TEXT, 12) 'sin uso
    '
    CUEIVA1% = 0: CUEIVB1% = 0: CUEADU1% = 0: CUERET1% = 0
    For U& = 1 To ULTREG&("ECUECON")
       x$ = REGLEIDO$("ECUECON", U&)
       If Left$(x$, 12) = CUEIVA$ Then CUEIVA1% = CVI(Mid$(x$, 41, 2)): DECIVA$ = Mid$(x$, 13, 28)
       If Left$(x$, 12) = CUEIVB$ Then CUEIVB1% = CVI(Mid$(x$, 41, 2)): DECIVA$ = Mid$(x$, 13, 28)
       If Left$(x$, 12) = CUEADU$ Then CUEADU1% = CVI(Mid$(x$, 41, 2)): DECADU$ = Mid$(x$, 13, 28)
       If Left$(x$, 12) = CUERET$ Then CUERET1% = CVI(Mid$(x$, 41, 2)): DECRET$ = Mid$(x$, 13, 28)
       'If Left$(x$, 12) = CUEREG$ Then CUEREG1% = CVI(Mid$(x$, 41, 2)): DECREG$ = Mid$(x$, 13, 28) 'sin uso
    Next U&
    '
    If CUEIVA1% < 1 Then
       Call MENSERR(24, "Falta Cuenta Contable I.V.A.")
       Text2(0).SetFocus
       Exit Sub
    End If
    '
    If CUEIVB1% < 1 Then CUEIVB$ = ""
    '
    If CUEADU1% < 1 Then
       Call MENSERR(24, "Falta Cuenta Contable I.V.A. R.N.I.")
       Text2(1).SetFocus
       Exit Sub
    End If
    '
    If CUERET1% < 1 Then
       Call MENSERR(24, "Falta Cuenta Contable Retenciones I.V.A.")
       Text2(2).SetFocus
       Exit Sub
    End If
    '
'    If CUEREG1% < 1 Then
'       Call MENSERR(24, "Falta Cuenta Contable Retenciones de Ganancias")
'       Text2(4).SetFocus
'       Exit Sub
'    End If

    'Validacion de Cuentas Contables Validas en l matriz text4
    For i& = 1 To Text4.UBound
      Codcuenta$ = TRIM$(Text4(i&).TEXT)
      If Codcuenta$ <> "" Then
        DECUE$ = ECOARCH("ECUECON", TRIM$(Text4(i&).TEXT))
        If DECUE$ = "" Then
          Call MENSERR(24, "Cuanta Contable No Válida")
          Call PINTATEXT(Text4(i&))
          Exit Sub
        End If
      End If
    Next i&
    '
    'Validacion de Cuentas Anticipos - Si se agrega una valida que este la de contrapartida.
    If TRIM$(Text4(17).TEXT) <> "" Or TRIM$(Text4(18).TEXT) <> "" Then
       If TRIM$(Text4(17).TEXT) = "" Or TRIM$(Text4(18).TEXT) = "" Then
          Call MENSERR(24, "Falta Una de Las Cuentas Contables Para Facturas Anticios\Complete o Borre Ambas")
          Call PINTATEXT(Text4(17))
          Exit Sub
       End If
    End If
    
    YA1% = 0: YA2% = 0: YA3% = 0: YA4% = 0
    For U& = 1 To ULTREG&("CUECON")
       x$ = REGLEIDO$("CUECON", U&)
       If CVI(Left$(x$, 2)) = CUEIVA1% Then YA1% = 1
       If CVI(Left$(x$, 2)) = CUEADU1% Then YA2% = 1
       If CVI(Left$(x$, 2)) = CUERET1% Then YA3% = 1
    Next U&
    '
    If YA1% < 1 Then
       x$ = REGBLAN$("CUECON")
       Call REPLA(x$, MKI$(CUEIVA1%), 1, 2)
       Call REPLA(x$, DECIVA$, 3, 29)
       Call REPLA(x$, CUEIVA$, 33, 12)
       Call REPLA(x$, CUEIVA$, 45, 12)
       Call REGAPP("CUECON", x$)
    End If
    '
    If YA2% < 1 Then
       x$ = REGBLAN$("CUECON")
       Call REPLA(x$, MKI$(CUEADU1%), 1, 2)
       Call REPLA(x$, DECADU$, 3, 29)
       Call REPLA(x$, CUEADU$, 33, 12)
       Call REPLA(x$, CUEADU$, 45, 12)
       Call REGAPP("CUECON", x$)
    End If
    '
    If YA3% < 1 Then
       x$ = REGBLAN$("CUECON")
       Call REPLA(x$, MKI$(CUERET1%), 1, 2)
       Call REPLA(x$, DECRET$, 3, 29)
       Call REPLA(x$, CUERET$, 33, 12)
       Call REPLA(x$, CUERET$, 45, 12)
       Call REGAPP("CUECON", x$)
    End If
    '
'    If YA4% < 1 Then  'sin uso
'       x$ = REGBLAN$("CUECON")
'       Call REPLA(x$, MKI$(CUEREG1%), 1, 2)
'       Call REPLA(x$, DECREG$, 3, 29)
'       Call REPLA(x$, CUEREG$, 33, 12)
'       Call REPLA(x$, CUEREG$, 45, 12)
'       Call REGAPP("CUECON", x$)
'    End If

    Call GRACONTROL("SVD451S", "CUEIVA", CUEIVA$)
    Call GRACONTROL("SVD451S", "CUEIVB", CUEIVB$)
    Call GRACONTROL("SVD451S", "CUEADU", CUEADU$)
    Call GRACONTROL("SVD451S", "CUERET", CUERET$)
   ' Call GRACONTROL("SVD451S", "CUEREG", CUEREG$)'sin uso
    '
    MOCALVE$ = "PIVACLI"
    Call GRACONTROL("REMEVENT", "MODCALVE", MOCALVE$)
    '
    If MOCALVE$ = "PIVACLI" Then
      For KKI% = 0 To 18
        CUEVENTA$(KKI%) = AJUSTI$(Text4(KKI%).TEXT, 12)
      Next KKI%
      '
      For U& = 1 To ULTREG&("ECUECON")
        x$ = REGLEIDO$("ECUECON", U&)
        For KKI% = 0 To 18
          If Left$(x$, 12) = CUEVENTA$(KKI%) Then COCUEVEN%(KKI%) = CVI(Mid$(x$, 41, 2)): DECIVA$ = Mid$(x$, 13, 28)
        Next KKI%
      Next U&
      '
      For KKI% = 0 To 18
        If COCUEVEN%(KKI%) > 0 Or KKI% = 12 Or KKI% = 13 Or KKI% = 14 Or KKI% = 15 Or KKI% = 16 Or KKI% = 17 Or KKI% = 18 Then
            Call GRACONTROL("SVD451S", "CUEVE-" + TRIM$(Str$(KKI%)), CUEVENTA$(KKI%))
          Else
            Call MENSERR(24, "Faltan Cuentas Contables de Venta.")
            Text4(KKI%).SetFocus
            Exit Sub
        End If
      Next KKI%
    End If
    
    '\\\OT-06-0158-WAR-30/05/06///
      Call GRACONTROL("COREPER", "COREPER", Text100.TEXT)
    '\\\OT-06-0158-WAR-FIN///
    
    Call GRACONTROL("ASVENTAS", "ASITEMSC", Text4(19).TEXT) 'CUENTA  DE ASIENTO DE VENTAS ITEMS SIN CARGO
    Call GRACONTROL("ASVENTAS", "CONTASSC", Text4(20).TEXT) 'CUENTA DE CONTRAPARTIDA DE ASIENTO DE VENTAS ITEMS SIN CARGO
    
    Call GRACONTROL("ASVENTIM", "ASIMAPRI", Text4(21).TEXT) 'CUENTA  DE ASIENTO DE VENTAS MATERIAS PRIMAS
    Call GRACONTROL("ASVENTIM", "CONTIMPO", Text4(22).TEXT) 'CUENTA DE CONTRAPARTIDA DE IMPORTADOS
    
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Hide
    '
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Unload Me
    On Error Resume Next
    Hide
    On Error GoTo 0
End Sub

Private Sub ImpuConPorGrupoVentas_Click()

20  Campos$ = "Grupo/I8;Descripción/A32;Cuenta contable/A20;Denominacion/A32"
    Call CARGA_ENCABEZADO(IMPUVENGRUPOCON.GRID, Campos$, IMPUVENGRUPOCON, 1, 0, ANTOT, , 1)
    Call CARGA_ENCABEZADO(IMPUVENGRUPOCON.MSF_2, Campos$, IMPUVENGRUPOCON, 1, 0, ANTOT, , 1)
    IMPUVENGRUPOCON.Caption = "Imputacion de Ventas por Grupo Contable"
    IMPUVENGRUPOCON.Command3.Caption = "Aceptar"
    IMPUVENGRUPOCON.Command3.ToolTipText = "Vincular cuenta contable con grupo de venta"
    '
    Screen.MousePointer = 11
    
    FIN& = ULTREG&("GRUCOVEN") + 1
    IMPUVENGRUPOCON.GRID.Rows = FIN&
    IMPUVENGRUPOCON.MSF_2.Rows = FIN&
    For iGrucoven& = 1 To FIN& - 1
        REGISTRO$ = REGLEIDO$("GRUCOVEN", iGrucoven&)
        codigoGrupo% = CVI(Mid$(REGISTRO$, 1, 2))
        denominacionGrupo$ = Mid$(REGISTRO$, 3, 24)
        IMPUVENGRUPOCON.GRID.TextMatrix(iGrucoven&, 1) = codigoGrupo%
        IMPUVENGRUPOCON.GRID.TextMatrix(iGrucoven&, 2) = denominacionGrupo$
        
        ctaContableInterno% = CodigoCtaContablePorGrupoVenta(codigoGrupo%)
        ctaContableExterno$ = CODCUE$(ctaContableInterno%)
        denominacionCtaContable$ = DENOCUE$(ctaContableInterno%)
        IMPUVENGRUPOCON.GRID.TextMatrix(iGrucoven&, 3) = ctaContableExterno$
        IMPUVENGRUPOCON.GRID.TextMatrix(iGrucoven&, 4) = denominacionCtaContable$
        
    Next iGrucoven&
    
    Screen.MousePointer = 1
    
    IMPUVENGRUPOCON.Show 1

    If IMPUVENGRUPOCON.HaySeleccion% > 0 Then
       CuentaContable% = SELECCUECON
       Screen.MousePointer = 11
       REFEGRUP$ = TRIM$(RESP_ZELE_VECT(2))
       If CuentaContable% < 1 Then
          OPX% = COMALTER%("Atención No se Ha Seleccionado Ningún Grupo de Actualización\Si Selecciona Grabar y Blanquear los Item Seleccionados Borrará el Dato Actual\\Cancelar\Grabar y Blanquear")
       Else
          OPX% = COMALTER%("Confirma Asociar Los Items Seleccionados al Grupo de Actualización: " & REFEGRUP$ & "\\Cancelar\Grabar")
       End If
       '
       Dim MSF As MSFlexGrid
       Set MSF = IMPUVENGRUPOCON.GRID
       If TRIM$(IMPUVENGRUPOCON.TXTBUSCAR) <> "" Then
          Set MSF = IMPUVENGRUPOCON.MSF_2
       End If
       
       If OPX% = 2 Then
          For i& = 1 To MSF.Rows - 1
            MARCA$ = TRIM$(MSF.TextMatrix(i&, 0))
            If MARCA$ = "*" Then
               grupoVenta% = XVALO(MSF.TextMatrix(i&, 1))
               
               If CantRegistros("GRUVECUECON", "GRUCOVEN=" & CStr(grupoVenta%)) > 0 Then
                    SQLX$ = "UPDATE GRUVECUECON SET CODINCUE = " & CStr(CuentaContable%)
                    SQLX$ = SQLX$ + " WHERE GRUCOVEN =  " & CStr(grupoVenta%)
               Else
                    SQLX$ = "INSERT INTO GRUVECUECON (GRUCOVEN, CODINCUE) VALUES (" & CStr(grupoVenta%) & "," & CStr(CuentaContable%) & ")"
               End If
               FLEXCONN.Execute (SQLX$)
            End If
          Next i&
          Screen.MousePointer = 1
          IMPUVENGRUPOCON.HaySeleccion% = 0
          GoTo 20
       Else
          Screen.MousePointer = 1
          IMPUVENGRUPOCON.HaySeleccion% = 0
          Exit Sub
       End If
    End If
    Screen.MousePointer = 1

End Sub

Function SELECCUECON()
   
   On Error Resume Next
   RESP_ZELE_VECT(1) = ""
   On Error GoTo 0
   Campos$ = "Cuenta/A20;Denominacion/A48"
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
   FRMZELECHO.Caption = "Lista de Cuentas Contables"
 
   Dim RS_FUN As ADODB.Recordset
   Set RS_FUN = New ADODB.Recordset
   SQLX$ = "SELECT CodExCue, Denocue FROM PLANCUEN WITH(NOLOCK) WHERE Statcue > 0 ORDER BY NUORDIT"
   Set RS_FUN = READSET(SQLX$)
   Set RS_PLANCUEN = RS_FUN
   Set RS_FUN = Nothing
  
   Dim obj As New RECORXET
   Set RS = RS_PLANCUEN
   Call Carga_MSF_Recordset(RS, Campos$, FRMZELECHO.MSF2, 1)
   Set obj = Nothing
   RS.Close
   Set RS = Nothing

   FRMZELECHO.Width = 8500
   FRMZELECHO.Show 1
   If TRIM$(RESP_ZELE_VECT(1)) <> "" Then
     codigoExterno$ = TRIM$(RESP_ZELE_VECT(1))
     CodigoInterno% = CUEINT%(codigoExterno$)
     SELECCUECON = CodigoInterno%
   End If
End Function

Private Sub mnuConfiCashflow_Click()
    FSCASHFLOW.Show 1
End Sub

Private Sub mnuCtasPosFac_Click()
    '
    Call TRALOCAL("CTASFACT", "")
    '
10  VTB% = VENTABU%("CTASFACT/NC/TITUPAN=Cuentas Posibles de Facturación")
    If VTB% < 0 Then Exit Sub
    '
    J& = 0
    For i& = 1 To ULTREG&("!CTASFACT")
        x$ = REGLEIDO$("!CTASFACT", i&)
        If CVI(Left$(x$, 2)) > 0 Then
           CUECON% = CVI(Mid$(x$, 1, 2))
           If CODCUE$(CUECON%) = "" Then
              Call MENSERR(24, "Falta Cuenta Contable en la Imputacion")
              GoTo 10
           End If
           J& = J& + 1
           Call GRAREG("CTASFACT", x$, J&)
        End If
    Next i&
    '
    Call SETULTREG("CTASFACT", J&)
    Call CIERRARCH("CTASFACT")

End Sub

Private Sub Text2_Change(Index As Integer)
    DECUE$ = ECOARCH("ECUECON", TRIM$(Text2(Index).TEXT))
    If DECUE$ <> "" Then
        Label3(Index) = DECUE$
      Else
        Label3(Index) = ""
    End If
End Sub

Private Sub Text2_DblClick(Index As Integer)
    Call SELECHO("ECUECON")
    If VALACT1$("ECUECON") <> "" Then
        Text2(Index).TEXT = VALACT1$("ECUECON")
        Label3(Index) = VALACT2$("ECUECON")
    End If
End Sub
'
Private Sub Command4_Click(Index As Integer)
    Call SELECHO("ECUECON")
    If VALACT1$("ECUECON") <> "" Then
        Text2(Index).TEXT = VALACT1$("ECUECON")
        Label3(Index) = VALACT2$("ECUECON")
    End If
End Sub
'
Private Sub Text4_Change(Index As Integer)
    DECUE$ = ECOARCH("ECUECON", TRIM$(Text4(Index).TEXT))
    If DECUE$ <> "" Then
        Label4(Index) = DECUE$
      Else
        Label4(Index) = ""
    End If
End Sub

Private Sub Text4_DblClick(Index As Integer)
    Call SELECHO("ECUECON")
    If VALACT1$("ECUECON") <> "" Then
        Text4(Index).TEXT = VALACT1$("ECUECON")
        Label4(Index) = VALACT2$("ECUECON")
    End If
End Sub

Private Sub Command11_Click()
   '
   Command11.Enabled = False
   RETEJ% = EJER%
   EJER% = 1
   Call CARCUEINGAS
   EJER% = RETEJ%
   EJU$ = EJACT$
   Command11.Enabled = True
   '
End Sub


