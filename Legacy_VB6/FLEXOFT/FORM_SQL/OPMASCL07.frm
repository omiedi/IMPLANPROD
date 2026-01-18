VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form OPMASCL07 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CLIENTES - Configuración"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   11085
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   11085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command9 
      Caption         =   "Proceso"
      Height          =   255
      Left            =   5160
      TabIndex        =   82
      ToolTipText     =   "Proceso actualiza el campo pos_iva y cuti_cli en cajabanc  Tomando del Maestro de Clientes - Seleccionando Fecha hasta."
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox MARCOBARRA 
      Appearance      =   0  'Flat
      BackColor       =   &H002B9973&
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   7550
      ScaleHeight     =   165
      ScaleWidth      =   1455
      TabIndex        =   77
      Top             =   720
      Width           =   1485
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H0048CC9D&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   -120
         TabIndex        =   78
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Dep.Entrega"
      Height          =   255
      Left            =   5160
      Picture         =   "OPMASCL07.frx":0000
      TabIndex        =   75
      ToolTipText     =   "Depuracion de domicilio de entrega."
      Top             =   4200
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Complementos"
      Height          =   255
      Left            =   7550
      Picture         =   "OPMASCL07.frx":014A
      TabIndex        =   72
      ToolTipText     =   "Actualiza Flexcrl de Configuración en Tabla Sql"
      Top             =   240
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E2D3C0&
      Caption         =   "VALORES PREDETERMINADOS"
      Height          =   4095
      Left            =   4920
      TabIndex        =   43
      Top             =   1290
      Width           =   6075
      Begin VB.TextBox COTEXT 
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
         Height          =   300
         Index           =   7
         Left            =   1485
         TabIndex        =   64
         Text            =   " "
         Top             =   450
         Width           =   570
      End
      Begin VB.TextBox DETEXT 
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
         Height          =   300
         Index           =   7
         Left            =   2325
         TabIndex        =   63
         TabStop         =   0   'False
         Text            =   "  "
         Top             =   450
         Width           =   3560
      End
      Begin VB.CommandButton BOTSEL 
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
         Index           =   7
         Left            =   2010
         TabIndex        =   62
         Top             =   450
         Width           =   175
      End
      Begin VB.TextBox COTEXT 
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
         Height          =   300
         Index           =   6
         Left            =   3480
         TabIndex        =   61
         Text            =   " "
         Top             =   3075
         Width           =   2250
      End
      Begin VB.TextBox DETEXT 
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
         Height          =   300
         Index           =   6
         Left            =   1485
         TabIndex        =   60
         TabStop         =   0   'False
         Text            =   "  "
         Top             =   3450
         Width           =   4410
      End
      Begin VB.CommandButton BOTSEL 
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
         Index           =   6
         Left            =   5715
         TabIndex        =   59
         Top             =   3075
         Width           =   175
      End
      Begin VB.CommandButton BOTSEL 
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
         Index           =   5
         Left            =   2010
         TabIndex        =   58
         Top             =   2550
         Width           =   175
      End
      Begin VB.CommandButton BOTSEL 
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
         Index           =   4
         Left            =   2010
         TabIndex        =   57
         Top             =   2130
         Width           =   175
      End
      Begin VB.CommandButton BOTSEL 
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
         Index           =   3
         Left            =   2010
         TabIndex        =   56
         Top             =   1710
         Width           =   175
      End
      Begin VB.CommandButton BOTSEL 
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
         Index           =   2
         Left            =   2010
         TabIndex        =   55
         Top             =   870
         Width           =   175
      End
      Begin VB.TextBox DETEXT 
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
         Height          =   300
         Index           =   5
         Left            =   2325
         TabIndex        =   54
         TabStop         =   0   'False
         Text            =   "  "
         Top             =   2550
         Width           =   3560
      End
      Begin VB.TextBox COTEXT 
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
         Height          =   300
         Index           =   5
         Left            =   1485
         TabIndex        =   53
         Text            =   " "
         Top             =   2550
         Width           =   570
      End
      Begin VB.TextBox COTEXT 
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
         Height          =   300
         Index           =   3
         Left            =   1485
         TabIndex        =   52
         Text            =   " "
         Top             =   1710
         Width           =   570
      End
      Begin VB.TextBox DETEXT 
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
         Height          =   300
         Index           =   2
         Left            =   2325
         TabIndex        =   51
         TabStop         =   0   'False
         Text            =   "  "
         Top             =   870
         Width           =   3560
      End
      Begin VB.TextBox COTEXT 
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
         Height          =   300
         Index           =   2
         Left            =   1485
         TabIndex        =   50
         Text            =   " "
         Top             =   870
         Width           =   570
      End
      Begin VB.TextBox COTEXT 
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
         Height          =   300
         Index           =   4
         Left            =   1485
         TabIndex        =   49
         Text            =   " "
         Top             =   2130
         Width           =   570
      End
      Begin VB.TextBox DETEXT 
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
         Height          =   300
         Index           =   3
         Left            =   2325
         TabIndex        =   48
         TabStop         =   0   'False
         Text            =   "  "
         Top             =   1710
         Width           =   3560
      End
      Begin VB.TextBox DETEXT 
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
         Height          =   300
         Index           =   4
         Left            =   2325
         TabIndex        =   47
         TabStop         =   0   'False
         Text            =   "  "
         Top             =   2130
         Width           =   3560
      End
      Begin VB.TextBox DETEXT 
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
         Height          =   300
         Index           =   0
         Left            =   2325
         TabIndex        =   46
         TabStop         =   0   'False
         Text            =   "  "
         Top             =   1290
         Width           =   3560
      End
      Begin VB.TextBox COTEXT 
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
         Height          =   300
         Index           =   0
         Left            =   1485
         TabIndex        =   45
         Text            =   "  "
         Top             =   1290
         Width           =   570
      End
      Begin VB.CommandButton BOTSEL 
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
         Index           =   0
         Left            =   2010
         TabIndex        =   44
         Top             =   1290
         Width           =   175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta Madre"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   71
         Top             =   3180
         Width           =   1905
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de Precios"
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
         Left            =   -300
         TabIndex        =   70
         Top             =   1815
         Width           =   1710
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo de Cliente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   380
         TabIndex        =   69
         Top             =   425
         Width           =   1020
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Categoría Ingresos Brutos"
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
         Left            =   135
         TabIndex        =   68
         Top             =   810
         Width           =   1260
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Condición de Pago"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   210
         TabIndex        =   67
         Top             =   2050
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor"
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
         Left            =   -510
         TabIndex        =   66
         Top             =   2655
         Width           =   1905
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Provincia"
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
         Left            =   210
         TabIndex        =   65
         Top             =   1395
         Width           =   1185
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Cnfg."
      Height          =   375
      Left            =   8520
      Picture         =   "OPMASCL07.frx":0294
      TabIndex        =   42
      ToolTipText     =   "Actualiza Flexcrl de Configuración en Tabla Sql"
      Top             =   840
      Visible         =   0   'False
      Width           =   650
   End
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   1470
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command5 
      Height          =   855
      Left            =   8040
      Picture         =   "OPMASCL07.frx":03DE
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Cálculo Automático de Límite de Crédito"
      Top             =   240
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E2D3C0&
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   7680
      TabIndex        =   35
      Top             =   480
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E2D3C0&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   7440
      TabIndex        =   34
      Top             =   315
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E2D3C0&
      Caption         =   "CONTROL"
      Height          =   950
      Left            =   9135
      TabIndex        =   31
      Top             =   120
      Width           =   1800
      Begin VB.CommandButton command1 
         Height          =   690
         Left            =   945
         Picture         =   "OPMASCL07.frx":06E8
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Cierra - Abandona la Aplicación"
         Top             =   210
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         Height          =   690
         Left            =   220
         Picture         =   "OPMASCL07.frx":0832
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Cierra - Abandona la Aplicación"
         Top             =   210
         Width           =   650
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2D3C0&
      Caption         =   "PERCEPCIONES IIBB"
      Height          =   950
      Left            =   4935
      TabIndex        =   28
      Top             =   120
      Visible         =   0   'False
      Width           =   2520
      Begin VB.CommandButton Command10 
         Height          =   690
         Left            =   1680
         Picture         =   "OPMASCL07.frx":0B3C
         Style           =   1  'Graphical
         TabIndex        =   83
         ToolTipText     =   "Configuracion Funcionamiento Percepciones IB CABA"
         Top             =   720
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command3 
         Height          =   690
         Left            =   380
         Picture         =   "OPMASCL07.frx":0F7E
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Configuracion Funcionamiento Percepciones IB Pcia de Bs.As"
         Top             =   210
         Width           =   650
      End
      Begin VB.CommandButton Command4 
         Height          =   690
         Left            =   1440
         Picture         =   "OPMASCL07.frx":13C0
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Listado Clientes para I.Brutos"
         Top             =   210
         Width           =   650
      End
   End
   Begin VB.TextBox COTEXT 
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
      Index           =   1
      Left            =   6300
      TabIndex        =   27
      Text            =   " "
      Top             =   2100
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.TextBox DETEXT 
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
      Index           =   1
      Left            =   7140
      TabIndex        =   26
      TabStop         =   0   'False
      Text            =   "  "
      Top             =   2100
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton BOTSEL 
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
      Index           =   1
      Left            =   6825
      TabIndex        =   25
      Top             =   2100
      Visible         =   0   'False
      Width           =   175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      Caption         =   "CONTROLES OPERATIVOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   90
      TabIndex        =   20
      Top             =   3015
      Width           =   4770
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Domicilio Fiscal Mandatario"
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
         Height          =   460
         Left            =   3240
         TabIndex        =   80
         Top             =   1080
         Width           =   1410
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2D3C0&
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
         Height          =   735
         Left            =   3120
         TabIndex        =   73
         ToolTipText     =   "Marcado / En Caso de no Tener Domicilio de Entrega Toma el Fiscal"
         Top             =   300
         Width           =   1575
         Begin VB.CheckBox Check6 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "No Replica Domicilio Fiscal"
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
            Height          =   460
            Left            =   120
            TabIndex        =   74
            ToolTipText     =   "Marcado / En Caso de no Tener Domicilio de Entrega Toma el Fiscal"
            Top             =   200
            Width           =   1410
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00E2D3C0&
         Caption         =   "FORMATO NÚMERO DE CUIT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   80
         TabIndex        =   36
         Top             =   1600
         Width           =   4620
         Begin VB.CheckBox Check4 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Derechos CUIT"
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
            Left            =   3120
            TabIndex        =   81
            ToolTipText     =   "Considera Derechos de Modificar Cuit"
            Top             =   300
            Width           =   1440
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Suprime Validación Formato (Arg.)"
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
            Left            =   120
            TabIndex        =   37
            Top             =   300
            Width           =   2880
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E2D3C0&
         Caption         =   "ALTA DE NUEVA CUENTA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   80
         TabIndex        =   21
         Top             =   300
         Width           =   3015
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Razón Social en Mayúscula"
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
            Left            =   120
            TabIndex        =   76
            Top             =   855
            Width           =   2760
         End
         Begin VB.TextBox Text3 
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
            Height          =   315
            Left            =   2160
            TabIndex        =   24
            Top             =   495
            Width           =   645
         End
         Begin VB.OptionButton Option3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Número Siguiente al"
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
            Height          =   270
            Left            =   120
            TabIndex        =   23
            Top             =   560
            Width           =   2100
         End
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Número Siguiente al Último"
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
            Height          =   405
            Left            =   120
            TabIndex        =   22
            Top             =   210
            Value           =   -1  'True
            Width           =   2595
         End
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E2D3C0&
      Caption         =   "TABLAS DE VALIDACION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2850
      Left            =   90
      TabIndex        =   0
      Top             =   100
      Width           =   4770
      Begin VB.CommandButton Command20 
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
         Left            =   3375
         TabIndex        =   1
         Top             =   285
         Width           =   175
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00E2D3C0&
         Caption         =   "OPCIONES DE TRABAJO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1725
         Left            =   210
         TabIndex        =   18
         Top             =   1000
         Width           =   4350
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Zonas de Reparto"
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
            Height          =   360
            Index           =   10
            Left            =   2310
            TabIndex        =   79
            Top             =   960
            Width           =   1935
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Provincias y Zonas de Venta"
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
            Height          =   525
            Index           =   5
            Left            =   120
            TabIndex        =   14
            Top             =   840
            Width           =   1800
         End
         Begin VB.OptionButton Option11 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Descuentos Combinados"
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
            Height          =   360
            Left            =   2310
            TabIndex        =   38
            Top             =   1320
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Grupos de Cliente"
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
            Height          =   450
            Index           =   9
            Left            =   2310
            TabIndex        =   17
            Top             =   570
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Cuentas Contables"
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
            Height          =   450
            Index           =   8
            Left            =   2310
            TabIndex        =   16
            Top             =   240
            Width           =   1950
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Lista de Precios"
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
            Height          =   345
            Index           =   7
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   1800
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Condiciones de Pago"
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
            Height          =   450
            Index           =   6
            Left            =   120
            TabIndex        =   13
            Top             =   500
            Width           =   2055
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Vendedores"
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
            Height          =   345
            Index           =   4
            Left            =   120
            TabIndex        =   15
            Top             =   1320
            Width           =   1590
         End
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Abrir"
         Height          =   700
         Left            =   3800
         Picture         =   "OPMASCL07.frx":1A2A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   285
         Width           =   750
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
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
         Height          =   300
         Left            =   210
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   " "
         Top             =   620
         Width           =   3345
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
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
         Height          =   300
         Left            =   1800
         TabIndex        =   9
         Text            =   " "
         Top             =   285
         Width           =   1695
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Listados de Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Left            =   2625
         TabIndex        =   2
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command21 
            Caption         =   "Análisis A-B-C"
            Height          =   435
            Left            =   2205
            TabIndex        =   8
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command22 
            Caption         =   "Estadísticas"
            Height          =   435
            Left            =   210
            TabIndex        =   7
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command23 
            Caption         =   "Críticos"
            Height          =   435
            Left            =   2205
            TabIndex        =   6
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command24 
            Caption         =   "Consignación"
            Height          =   435
            Left            =   210
            TabIndex        =   5
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command25 
            Caption         =   "Negativos"
            Height          =   435
            Left            =   2205
            TabIndex        =   4
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command26 
            Caption         =   "Inventarios"
            Height          =   435
            Left            =   210
            TabIndex        =   3
            Top             =   420
            Width           =   1590
         End
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TABLA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -615
         TabIndex        =   19
         Top             =   345
         Width           =   2325
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "OPMASCL07.frx":1D34
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "OPMASCL07.frx":1F68
      TabIndex        =   41
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Menu mnuUtilidades 
      Caption         =   "Utilidades"
      Begin VB.Menu mnuSincrprov 
         Caption         =   "Sincronización de Código Alfa de Provincias"
      End
   End
End
Attribute VB_Name = "OPMASCL07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TABLA$(16)
Dim MODUCONE$

Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Unload Me
   '
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   If DERACCE%("ACPAIBRU", "Actualizar Parametros Ingresos Brutos") = 2 Then
10   VDEF$ = String$(16, 0) + Space$(16)
     Call REPLA(VDEF$, MKS$(XVALO(Control$("PERIBCA", "TASA-A"))), 1, 4)
     Call REPLA(VDEF$, MKS$(XVALO(Control$("PERIBCA", "TASA-B"))), 5, 4)
     Call REPLA(VDEF$, MKS$(XVALO(Control$("PERIBCA", "TASA-EXP"))), 9, 4)
     Call REPLA(VDEF$, MKS$(XVALO(Control$("PERIBCA", "MINIPER"))), 13, 4)
     Call REPLA(VDEF$, Control$("PERIBCA", "CUECONTA"), 17, 12)
     '
     OBX$ = OBJPLA$("SETPERCEPIBCABA", VDEF$)
     If OBX$ <> "" Then
       Call GRACONTROL$("PERIBCA", "TASA-A", Str(CVS(Mid$(OBX$, 1, 4))))
       Call GRACONTROL$("PERIBCA", "TASA-B", Str(CVS(Mid$(OBX$, 5, 4))))
       Call GRACONTROL$("PERIBCA", "TASA-EXP", Str(CVS(Mid$(OBX$, 9, 4))))
       Call GRACONTROL$("PERIBCA", "MINIPER", Str(CVS(Mid$(OBX$, 13, 4))))
       '\\\OT-06-0087-WAR-13/03/06///
       If TRIM$(Mid$(OBX$, 1, 17)) <> "" Then
          CUECO1$ = Mid$(OBX$, 17, 12)
          If TRIM$(CUECO1$) = "" Or ECOARCH("ECUECON", CUECO1$) = "" Then
            Call MENSERR(24, "Cuenta Contable no Válida.")
            GoTo 10
          End If
       End If
       '\\\OT-06-0087-WAR-FIN///
       Call GRACONTROL$("PERIBCA", "CUECONTA", Mid$(OBX$, 17, 12))
     End If
     '
   End If
   Command10.Enabled = True
End Sub



Sub Command16_Click()
    '
    Command16.Enabled = False
    If TRIM$(MODUCONE$) <> "" Then
      If TRIM$(MODUCONE$) = "GRUCOCLI" Then
            LDT$ = LUDAT$
            Call CIERRARCH("GRUCOCLI")
            Call CIERRARCH("!GRUCOCLI")
            'Call COPYFILE(LDT$ + "GRUCOCLI.DAT", "C:\FLEXOFT\GRUCOCLI.WKF")
            Call TRALOCAL("GRUCOCLI", "")
            VTB% = VENTABU%("GRUCOCLI")
            If VTB% < 0 Then GoTo 99
            J& = 0
            For i& = 1 To ULTREG&("!GRUCOCLI")
                X$ = REGLEIDO$("!GRUCOCLI", i&)
                If CVI(Left$(X$, 2)) > 0 Then
                    J& = J& + 1
                    Call GRAREG("GRUCOCLI", X$, J&)
                End If
            Next i&
            Call SETULTREG("GRUCOCLI", J&)
            Call CIERRARCH("GRUCOCLI")
          ElseIf MODUCONE = "TADEPOSI" Then
            Call CARTADEPOSI
          ElseIf MODUCONE = "TATIMAT" Then
            Call CARTATIMA
          ElseIf MODUCONE = "CONPAGO" Then
            Call CARCONPAGOS
          ElseIf MODUCONE = "TAPROVIN" Then
            Call CARTAPROVIN
          ElseIf MODUCONE = "CUEINGAS" Then
            Call CARCUEINGAS
          ElseIf MODUCONE = "*VENDEDOR" Then
                RTA% = ALTERNA%("ABM Vendedores\Datos Complementarios")
                If RTA% = 1 Then
                    Call CONECRUN(MODUCONE$, Text2.TEXT)
                ElseIf RTA% = 2 Then
                    Call CARTAVENDE         'CARGA DATOS COMPLEMENTARIOS DEL VENDEDOR
                Else
                    GoTo 99
                End If
          ElseIf MODUCONE = "ZONAREP" Then
            Call CARZONAREP
          
          Else
            Call CONECRUN(MODUCONE$, Text2.TEXT)
      End If
    End If
99  Command16.Enabled = True
End Sub

Private Sub Command2_Click()
   OPXX$ = "-1"
   If Option3.Value = True Then OPXX$ = TRIM$(Str$(XVALO(Text3.TEXT)))
   Call GRACONTROL("", "PRONUCLI", OPXX$)
   '\\\OT-05-0198-WAR-19/04/05-///
   VALICUIT1$ = "NO"
   If Check1.Value = 1 Then VALICUIT1$ = "SI"
   Call GRACONTROL("FORMCUIT", "SUPRIVAL", VALICUIT1$)
   ''\\\OT-05-0198-WAR-FIN///
   For Index = 0 To 7
       TV11% = TIPFIELD(ARCHE$(Index), 1)
       LF11% = LENFIELD(ARCHE$(Index), 1)
       ARGU$ = CBIN$(TRIM$(COTEXT(Index).TEXT), TV11%, LF11%)
       VAPRE$ = "": ECO$ = ""
       If TRIM$(COTEXT(Index).TEXT) <> "" Then
          ECO$ = TRIM$(ECOARCH$(ARCHE$(Index), ARGU$))
          If ECO$ <> "" Then
             VAPRE$ = TRIM$(COTEXT(Index).TEXT)
          End If
       End If
       COTEXT(Index).TEXT = VAPRE$
       DETEXT(Index).TEXT = ECO$
       Call GRACONTROL("AMACLI", ARCHE$(Index), VAPRE$)
   Next Index
   '
  'SI ES NO, NO PRESENTA EL DOMICILIO FISCAL EN SOLAPA DE ENTREGA CUANDO ESTE NO TIENE INFORMACION
  DOMIFI$ = "SI"
  If Check6.Value = 1 Then DOMIFI$ = "NO"
  Call GRACONTROL("MASTER", "ENTREPFI", DOMIFI$)
  ' AUTOMATICAMENTE LLEVA LA RAZON SOCIAL A MAYÚSCULA
  RASOMAYU$ = "NO"
  If Check2.Value = 1 Then RASOMAYU$ = "SI"
  Call GRACONTROL("MASTER", "RASOMAYU", RASOMAYU$)
  
  DOMIFIMA$ = "NO"
  If Check3.Value = 1 Then DOMIFIMA$ = "SI"
  Call GRACONTROL("MASTER", "DOMIMAND", DOMIFIMA$)
  
  VALICU$ = "NO"
  If Check4.Value = 1 Then VALICU$ = "SI"
  Call GRACONTROL("AMACLI", "CONSICUI", VALICU$)
  


End Sub

 Sub Command3_Click()
 
   Command3.Enabled = False
   Opcion% = COMALTER%("Jurisdicciones Percepción Ingresos Brutos\\Bs-As\CABA\Misiones\Tucumán")
   If Opcion% < 1 Or Opcion% > 4 Then GoTo 99
   
   If Opcion% = 1 Then
        If DERACCE%("ACPAIBRU", "Actualizar Parametros Ingresos Brutos") = 2 Then
10        VDEF$ = String$(16, 0) + Space$(16)
          Call REPLA(VDEF$, MKS$(XVALO(Control$("PERIBRU", "TASA-A"))), 1, 4)
          Call REPLA(VDEF$, MKS$(XVALO(Control$("PERIBRU", "TASA-B"))), 5, 4)
          Call REPLA(VDEF$, MKS$(XVALO(Control$("PERIBRU", "TASA-EXP"))), 9, 4)
          Call REPLA(VDEF$, MKS$(XVALO(Control$("PERIBRU", "MINIPER"))), 13, 4)
          Call REPLA(VDEF$, Control$("PERIBRU", "CUECONTA"), 17, 12)
          '
          OBX$ = OBJPLA$("PARPERIBRU", VDEF$)
          If OBX$ <> "" Then
            Call GRACONTROL$("PERIBRU", "TASA-A", Str(CVS(Mid$(OBX$, 1, 4))))
            Call GRACONTROL$("PERIBRU", "TASA-B", Str(CVS(Mid$(OBX$, 5, 4))))
            Call GRACONTROL$("PERIBRU", "TASA-EXP", Str(CVS(Mid$(OBX$, 9, 4))))
            Call GRACONTROL$("PERIBRU", "MINIPER", Str(CVS(Mid$(OBX$, 13, 4))))
            '\\\OT-06-0087-WAR-13/03/06///
            If TRIM$(Mid$(OBX$, 1, 17)) <> "" Then
               CUECO1$ = Mid$(OBX$, 17, 12)
               If TRIM$(CUECO1$) = "" Or ECOARCH("ECUECON", CUECO1$) = "" Then
                 Call MENSERR(24, "Cuenta Contable no Válida.")
                 GoTo 10
               End If
            End If
            '\\\OT-06-0087-WAR-FIN///
            Call GRACONTROL$("PERIBRU", "CUECONTA", Mid$(OBX$, 17, 12))
          End If
          '
        End If
        
   ElseIf Opcion% = 2 Then
        If DERACCE%("ACPAIBRU", "Actualizar Parametros Ingresos Brutos") = 2 Then
20        VDEF$ = String$(16, 0) + Space$(16)
          Call REPLA(VDEF$, MKS$(XVALO(Control$("PERIBCA", "TASA-A"))), 1, 4)
          Call REPLA(VDEF$, MKS$(XVALO(Control$("PERIBCA", "TASA-B"))), 5, 4)
          Call REPLA(VDEF$, MKS$(XVALO(Control$("PERIBCA", "TASA-EXP"))), 9, 4)
          Call REPLA(VDEF$, MKS$(XVALO(Control$("PERIBCA", "MINIPER"))), 13, 4)
          Call REPLA(VDEF$, Control$("PERIBCA", "CUECONTA"), 17, 12)
          '
          OBX$ = OBJPLA$("SETPERCEPIBCABA/PERCEPCION IIBB C.A.B.A.", VDEF$)
          If OBX$ <> "" Then
            Call GRACONTROL$("PERIBCA", "TASA-A", Str(CVS(Mid$(OBX$, 1, 4))))
            Call GRACONTROL$("PERIBCA", "TASA-B", Str(CVS(Mid$(OBX$, 5, 4))))
            Call GRACONTROL$("PERIBCA", "TASA-EXP", Str(CVS(Mid$(OBX$, 9, 4))))
            Call GRACONTROL$("PERIBCA", "MINIPER", Str(CVS(Mid$(OBX$, 13, 4))))
            '\\\OT-06-0087-WAR-13/03/06///
            If TRIM$(Mid$(OBX$, 1, 17)) <> "" Then
               CUECO1$ = Mid$(OBX$, 17, 12)
               If TRIM$(CUECO1$) = "" Or ECOARCH("ECUECON", CUECO1$) = "" Then
                 Call MENSERR(24, "Cuenta Contable no Válida.")
                 GoTo 20
               End If
            End If
            '\\\OT-06-0087-WAR-FIN///
            Call GRACONTROL$("PERIBCA", "CUECONTA", Mid$(OBX$, 17, 12))
          End If
          '
        End If
        
   ElseIf Opcion% = 3 Then
        If DERACCE%("ACPAIBRU", "Actualizar Parametros Ingresos Brutos") = 2 Then
            OpcionMisiones% = COMALTER%("Configuraciones\\Alicuotas y otros\Código de Provincia")
            If OpcionMisiones% = 1 Then
30                VDEF$ = String$(16, 0) + Space$(16)
                  Call REPLA(VDEF$, MKS$(XVALO(Control$("PERIBMI", "TASA-A"))), 1, 4)
                  Call REPLA(VDEF$, MKS$(XVALO(Control$("PERIBMI", "TASA-B"))), 5, 4)
                  Call REPLA(VDEF$, MKS$(XVALO(Control$("PERIBMI", "TASA-EXP"))), 9, 4)
                  Call REPLA(VDEF$, MKS$(XVALO(Control$("PERIBMI", "MINIPER"))), 13, 4)
                  Call REPLA(VDEF$, Control$("PERIBMI", "CUECONTA"), 17, 12)
                  '
                  OBX$ = OBJPLA$("SETPERCEPIBCABA/PERCEPCION IIBB MISIONES", VDEF$)
                  If OBX$ <> "" Then
                    Call GRACONTROL$("PERIBMI", "TASA-A", Str(CVS(Mid$(OBX$, 1, 4))))
                    Call GRACONTROL$("PERIBMI", "TASA-B", Str(CVS(Mid$(OBX$, 5, 4))))
                    Call GRACONTROL$("PERIBMI", "TASA-EXP", Str(CVS(Mid$(OBX$, 9, 4))))
                    Call GRACONTROL$("PERIBMI", "MINIPER", Str(CVS(Mid$(OBX$, 13, 4))))
                    '\\\OT-06-0087-WAR-13/03/06///
                    If TRIM$(Mid$(OBX$, 1, 17)) <> "" Then
                       CUECO1$ = Mid$(OBX$, 17, 12)
                       If TRIM$(CUECO1$) = "" Or ECOARCH("ECUECON", CUECO1$) = "" Then
                         Call MENSERR(24, "Cuenta Contable no Válida.")
                         GoTo 30
                       End If
                    End If
                    '\\\OT-06-0087-WAR-FIN///
                    Call GRACONTROL$("PERIBMI", "CUECONTA", Mid$(OBX$, 17, 12))
                  End If
                  '
            ElseIf OpcionMisiones% = 2 Then
                Call SELECHO("PROVINC")
                CodigoProvincia$ = TRIM$(VALACT1$("PROVINC"))
                If CodigoProvincia$ <> "" Then
                    Call GRACONTROL$("PERIBMI", "CODPROV", CodigoProvincia$)
                    
                    If EXISTE%(LUDAT$ & "DAPIBMI.DAT") > 0 Then Kill LUDAT$ & "DAPIBMI.DAT"
                    Call ACLI_GES07.GENDAPIBMI
                                                           
                    ConsultaSql$ = "select nume_cli, prov_cli from deudor12 with(nolock) where stat_cli>0 and piva_cli=1"
                    Dim rst_misiones As ADODB.Recordset
                    Set rst_misiones = READSET(ConsultaSql$)
                    With rst_misiones
                        FIN& = rst_misiones.RecordCount
                        Do While Not .EOF
                            FI& = FI& + 1
                            Call TRACE(24, FI&, FIN&)
                            
                            NCLIE = XVALO(!nume_cli)
                            If SAFETEXT$(!Prov_Cli) = CodigoProvincia$ Then
                                RECORD$ = FILTERGETRECORD$("DAPIBMI", 1, MKS$(NCLIE))
                                Call REPLA(RECORD$, MKS$(NCLIE), 1, 4)
                                Call REPLA(RECORD$, Chr$(1), 5, 1)
                                Call REPLA(RECORD$, USERNAME$, 42, 20)
                                Call REPLA(RECORD$, MKI$(HOY(HOS$)), 62, 2)
                                Call REPLA(RECORD$, MKI$(Int(HORANUM(Time$))), 64, 2)
                                Call FILTERPUTRECORD("DAPIBMI", 1, MKS$(NCLIE), RECORD$)
                            End If
                            
                            .MoveNext
                        Loop
                    End With
                    Call CIERRARCH("DAPIBMI")
                    On Error Resume Next
                    rst_misiones.Close
                    Set rst_misiones = Nothing
                    On Error GoTo 0
                    
                    Call COMUNI("Código de Provincia '" & CodigoProvincia$ & "' Registrado.")
                End If
            End If
        End If
            
   Else
        If DERACCE%("ACPAIBRU", "Actualizar Parametros Ingresos Brutos") = 2 Then
40        VDEF$ = String$(16, 0) + Space$(16)
          Call REPLA(VDEF$, MKS$(XVALO(Control$("PERIBTU", "DIAVENCI"))), 1, 4)
          Call REPLA(VDEF$, MKS$(XVALO(Control$("PERIBTU", "NOINCPAD"))), 5, 4)
          'Call REPLA(VDEF$, MKS$(XVALO(Control$("PERIBTU", "TASA-EXP"))), 9, 4)
          Call REPLA(VDEF$, MKS$(XVALO(Control$("PERIBTU", "MINIPER"))), 13, 4)
          Call REPLA(VDEF$, Control$("PERIBTU", "CUECONTA"), 17, 12)
          '
          OBX$ = OBJPLA$("SETPERCEPIBTUCU/PERCEPCION IIBB TUCUMAN", VDEF$)
          If OBX$ <> "" Then
            Call GRACONTROL$("PERIBTU", "DIAVENCI", Str(CVS(Mid$(OBX$, 1, 4))))
            Call GRACONTROL$("PERIBTU", "NOINCPAD", Str(CVS(Mid$(OBX$, 5, 4))))
            'Call GRACONTROL$("PERIBTU", "TASA-EXP", Str(CVS(Mid$(OBX$, 9, 4))))
            Call GRACONTROL$("PERIBTU", "MINIPER", Str(CVS(Mid$(OBX$, 13, 4))))
            '\\\OT-06-0087-WAR-13/03/06///
            If TRIM$(Mid$(OBX$, 1, 17)) <> "" Then
               CUECO1$ = Mid$(OBX$, 17, 12)
               If TRIM$(CUECO1$) = "" Or ECOARCH("ECUECON", CUECO1$) = "" Then
                 Call MENSERR(24, "Cuenta Contable no Válida.")
                 GoTo 40
               End If
            End If
            '\\\OT-06-0087-WAR-FIN///
            Call GRACONTROL$("PERIBTU", "CUECONTA", Mid$(OBX$, 17, 12))
          End If
          '
        End If
        
   End If
   
99 Command3.Enabled = True

End Sub

Sub Command4_Click()
   Command4.Enabled = False
   
   Opcion% = COMALTER%("Listado de Ingresos Brutos\\Buenos Aires\CABA\Misiones")
   If Opcion% < 1 Or Opcion% > 3 Then GoTo 99
   
   Screen.MousePointer = 11
   
   Dim oReporte As cGenrep
   Dim Rst As ADODB.Recordset
   Set oReporte = New cGenrep
   
   If Opcion% = 1 Then
        PROVINCIA$ = "BUENOS AIRES"
   ElseIf Opcion% = 2 Then
        PROVINCIA$ = "CABA"
   ElseIf Opcion% = 3 Then
        PROVINCIA$ = "MISIONES"
   End If
   
   '"Cliente(ECO=DEUDOR1-A32)/I10;Cat(ECO=CARETIB-A32)/I3;Numero/A32;St(ECO=TACATIBC-A32)/I2;Exento(T=Motivo Exento)/A32;Jur(ECO=PROVINC-A24)/A3")
   Campos$ = "Cuenta/I10;Cliente/A48;Cat/I3;Categoria/A24;Nro IB/A24;Exento/A32;Jur/I3;Jurisdiccion/A24"
   Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
   Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
   GRILLAGRAL.Caption = "LISTADO DE INGRESOS BRUTOS (" & PROVINCIA$ & ")"
   GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
   GRILLAGRAL.mnuMenu1.Visible = True
   GRILLAGRAL.mnuMenu3.Caption = "Invertir selección"
   
   If Opcion% = 1 Then
        Call FILESORT("DAPIBCLN", "DAPIBCLJ", 1, 1, "ASC")
    
        RELISA& = 0
        FIN& = ULTREG&("DAPIBCLJ")
        For U& = 1 To FIN&
          
          Call TRACE(20, U&, FIN&)
          XX$ = REGLEIDO$("DAPIBCLJ", U&)
          NumeroCliente& = CVS(Left$(XX$, 4))
          CategoriaIngBrutos% = XVALO(CATIBRU$(NumeroCliente&))
          NumeroIngBrutos$ = NIBRCLI$(NumeroCliente&)
          Status% = Asc(Mid$(XX$, 5, 1))
          Exento$ = Mid$(XX$, 10, 32)
          Jurisdiccion$ = SAFETEXT$(VALOBADA("DEUDOR12", "JurisdiccioIIBB_CLI", "NUME_CLI=" & CStr(NumeroCliente&)))
          razonSocial$ = SAFETEXT$(VALOBADA("DEUDOR12", "Raso_Cli", "NUME_CLI=" & CStr(NumeroCliente&)))
          
          J& = J& + 1
          GRILLAGRAL.GRID.Rows = J& + 1
          GRILLAGRAL.GRID.TextMatrix(J&, 1) = NumeroCliente&
          GRILLAGRAL.GRID.TextMatrix(J&, 2) = razonSocial$
          GRILLAGRAL.GRID.TextMatrix(J&, 3) = CategoriaIngBrutos%
          GRILLAGRAL.GRID.TextMatrix(J&, 4) = ECOARCH$("CARETIB", CStr(CategoriaIngBrutos%))
          GRILLAGRAL.GRID.TextMatrix(J&, 5) = Mid$(NumeroIngBrutos$, 1, 32)
          GRILLAGRAL.GRID.TextMatrix(J&, 6) = Exento$
          GRILLAGRAL.GRID.TextMatrix(J&, 7) = Jurisdiccion$
          GRILLAGRAL.GRID.TextMatrix(J&, 8) = ECOARCH$("PROVINC", Jurisdiccion$)
          
        Next U&
       
        GRILLAGRAL.GRID.ZOrder 0
        GRILLAGRAL.TXTBUSCAR = ""
        Call POSICIONAR_ULTIMAFILA(GRILLAGRAL.GRID)
        Screen.MousePointer = 1
          
        GRILLAGRAL.Show 1
        Dim MSF0 As msFlexGrid
        If GRILLAGRAL.MENU1 > 0 Then
            Set MSF0 = GRILLAGRAL.GRID
            If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
            Call CARGA_TABLA_IMPRE(MSF0, Campos$, "LIIBBSAS", GRILLAGRAL.Caption, "TABLA_IMPRE", RET_NAMECONS$, 1)
            Call FINAL("?" & RET_NAMECONS$)
            GRILLAGRAL.MENU1 = 0
            Screen.MousePointer = 1
        End If
       
   ElseIf Opcion% = 2 Then
   
        Call FILESORT("DAPIBCA", "DAPIBCLJ", 1, 1, "ASC")
    
        RELISA& = 0
        FIN& = ULTREG&("DAPIBCLJ")
        For U& = 1 To FIN&
          
          Call TRACE(20, U&, FIN&)
          XX$ = REGLEIDO$("DAPIBCLJ", U&)
          NumeroCliente& = CVS(Left$(XX$, 4))
          CategoriaIngBrutos% = XVALO(CATIBRU$(NumeroCliente&))
          NumeroIngBrutos$ = NIBRCLI$(NumeroCliente&)
          Status% = Asc(Mid$(XX$, 5, 1))
          Exento$ = Mid$(XX$, 10, 32)
          Jurisdiccion$ = SAFETEXT$(VALOBADA("DEUDOR12", "JurisdiccioIIBB_CLI", "NUME_CLI=" & CStr(NumeroCliente&)))
          razonSocial$ = SAFETEXT$(VALOBADA("DEUDOR12", "Raso_Cli", "NUME_CLI=" & CStr(NumeroCliente&)))
          
          J& = J& + 1
          GRILLAGRAL.GRID.Rows = J& + 1
          GRILLAGRAL.GRID.TextMatrix(J&, 1) = NumeroCliente&
          GRILLAGRAL.GRID.TextMatrix(J&, 2) = razonSocial$
          GRILLAGRAL.GRID.TextMatrix(J&, 3) = CategoriaIngBrutos%
          GRILLAGRAL.GRID.TextMatrix(J&, 4) = ECOARCH$("CARETIB", CStr(CategoriaIngBrutos%))
          GRILLAGRAL.GRID.TextMatrix(J&, 5) = Mid$(NumeroIngBrutos$, 1, 32)
          GRILLAGRAL.GRID.TextMatrix(J&, 6) = Exento$
          GRILLAGRAL.GRID.TextMatrix(J&, 7) = ""
          GRILLAGRAL.GRID.TextMatrix(J&, 8) = ""
          
        Next U&
           
        GRILLAGRAL.GRID.ZOrder 0
        GRILLAGRAL.TXTBUSCAR = ""
        Call POSICIONAR_ULTIMAFILA(GRILLAGRAL.GRID)
        Screen.MousePointer = 1
          
        GRILLAGRAL.Show 1
        Dim MSF1 As msFlexGrid
        If GRILLAGRAL.MENU1 > 0 Then
            Set MSF1 = GRILLAGRAL.GRID
            If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF1 = GRILLAGRAL.MSF_2
            Call CARGA_TABLA_IMPRE(MSF1, Campos$, "LIIBCABA", GRILLAGRAL.Caption, "TABLA_IMPRE", RET_NAMECONS$, 1)
            Call FINAL("?" & RET_NAMECONS$)
            GRILLAGRAL.MENU1 = 0
            Screen.MousePointer = 1
        End If
        
   ElseIf Opcion% = 3 Then
   
        Call FILESORT("DAPIBMI", "DAPIBCLJ", 1, 1, "ASC")
    
        RELISA& = 0
        FIN& = ULTREG&("DAPIBCLJ")
        For U& = 1 To FIN&
          
          Call TRACE(20, U&, FIN&)
          XX$ = REGLEIDO$("DAPIBCLJ", U&)
          NumeroCliente& = CVS(Left$(XX$, 4))
          CategoriaIngBrutos% = XVALO(CATIBRU$(NumeroCliente&))
          NumeroIngBrutos$ = NIBRCLI$(NumeroCliente&)
          Status% = Asc(Mid$(XX$, 5, 1))
          Exento$ = Mid$(XX$, 10, 32)
          razonSocial$ = SAFETEXT$(VALOBADA("DEUDOR12", "Raso_Cli", "NUME_CLI=" & CStr(NumeroCliente&)))
          
          J& = J& + 1
          GRILLAGRAL.GRID.Rows = J& + 1
          GRILLAGRAL.GRID.TextMatrix(J&, 1) = NumeroCliente&
          GRILLAGRAL.GRID.TextMatrix(J&, 2) = razonSocial$
          GRILLAGRAL.GRID.TextMatrix(J&, 3) = CategoriaIngBrutos%
          GRILLAGRAL.GRID.TextMatrix(J&, 4) = ECOARCH$("CARETIB", CStr(CategoriaIngBrutos%))
          GRILLAGRAL.GRID.TextMatrix(J&, 5) = Mid$(NumeroIngBrutos$, 1, 32)
          GRILLAGRAL.GRID.TextMatrix(J&, 6) = Exento$
          GRILLAGRAL.GRID.TextMatrix(J&, 7) = ""
          GRILLAGRAL.GRID.TextMatrix(J&, 8) = ""
          
        Next U&
           
        GRILLAGRAL.GRID.ZOrder 0
        GRILLAGRAL.TXTBUSCAR = ""
        Call POSICIONAR_ULTIMAFILA(GRILLAGRAL.GRID)
        Screen.MousePointer = 1
          
        GRILLAGRAL.Show 1
        Dim msf2 As msFlexGrid
        If GRILLAGRAL.MENU1 > 0 Then
            Set msf2 = GRILLAGRAL.GRID
            If GRILLAGRAL.MSF_2.Rows > 1 Then Set msf2 = GRILLAGRAL.MSF_2
            Call CARGA_TABLA_IMPRE(msf2, Campos$, "LIIBMISI", GRILLAGRAL.Caption, "TABLA_IMPRE", RET_NAMECONS$, 1)
            Call FINAL("?" & RET_NAMECONS$)
            GRILLAGRAL.MENU1 = 0
            Screen.MousePointer = 1
        End If
        
   End If
   
   Screen.MousePointer = 1
            
'
'        Screen.MousePointer = 11
'
'        FIN& = ULTREG&("DAPIBCLJ")
'        For U& = 1 To FIN&
'          Call TRACE(20, U&, FIN&)
'          XX$ = REGLEIDO$("DAPIBCLJ", U&)
'          NCLIX% = CVI(Left$(XX$, 2))
'          CATIBX% = XVALO(CATIBRU$(NCLIX%))
'          NIBR$ = NIBRCLI$(NCLIX%)
'
'          Call REPLA(XX$, MKI$(CATIBX%), 65, 2)
'          Call REPLA(XX$, NIBR$, 67, 16)
'          If CATIBX% > 0 Then
'            If TRIM$(NIBR$) = "" Then
'              Call REPLA(XX$, "Falta Nro. Inscripción", 83, 24)
'            End If
'          End If
'          Call GRAREG("DAPIBCLJ", XX$, U&)
'        Next U&
'        Call CIERRARCH("DAPIBCLJ")
'        Screen.MousePointer = 1
'        '
'        Call FINAL("*LIDAPIBC")
'   End If
'   '
99 Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   'obtener fecha de 90 dias atras o ultimos 3 meses ???
   'consultar realizar un sum sobre el cajabanc donde tipomvim = FVEN AND NUCLIEN = NC% sobre IHABE
   'consultar realizar un sum sobre el cajabanc donde tipomvim = FVEN AND NUCLIEN = NC sobre IDEBE
   'VALIDAR SI EL RESULTADO ES NEGATIVO QUE EL IMPORTE SEA 0
   'EN CASO CONTRARIO CALCULAR LIMITE DE CREDITO SEGUN CASO PARA CONDICION 1,5,6,12
   'REGISTRAR EL RESULTADO COMO LIMITE DE CREDITO
   'TAMBIEN ATENDIENDO UNA SOLICITUD DE JUAN SE BORRO EL DESCUENTO STANDARD DEL CLIENTE
   'obtener fecha de 90 dias atras o ultimos 3 meses ???
   Command5.Enabled = False
   Screen.MousePointer = 11

    DIXA% = HOY(HO$)
    FECHADESDE = FETEXCOR1$(DIASANTES(DIXA%, 90))
    PAC$ = TRIM$(UCase$(EMPREAC$))
    EPAC$ = TRIM$(UCase$(EMPREAC$))

    '
    'generar recordeset de tabla deudor12
    CONDI$ = " NUME_CLI > 0"
    FIN& = CantRegistros&("DEUDOR12", CONDI$, "NOMESS")
    SQLX$ = " SELECT * FROM DEUDOR12 WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE " & CONDI$
    SQLX$ = SQLX$ + " ORDER BY NUME_CLI "
    Dim CLITMP As ADODB.Recordset
    Set CLITMP = New ADODB.Recordset
    Set CLITMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    jj& = 0
    With CLITMP
      Do While Not .EOF
        NC% = XVALO(!nume_cli)
        jj& = jj& + 1
        Call TRACE(20, jj&, FIN&)
        If NC% > 0 Then
           'consultar realizar un sum sobre el cajabanc donde tipomvim = FVEN AND NUCLIEN = NC% sobre IHABE
           'consultar realizar un sum sobre el cajabanc donde tipomvim = FVEN AND NUCLIEN = NC sobre IDEBE
           CONDI$ = "NuClien= " & NC% & " AND TIPOMOVIM = 'FVEN' AND FechEmisi >= '" & FECHADESDE & "'"
           IHAB# = XVALO(VALOBADA("CAJABANC", "SUM(IHABECOMP)", CONDI$, "NOMESS"))
           IDEB# = XVALO(VALOBADA("CAJABANC", "SUM(IDEBECOMP)", CONDI$, "NOMESS"))
           TOTAVENTA# = IDEB# - IHAB#
           If TOTAVENTA# < 0.005 Then TOTAVENTA# = 0
           CONDPAGO% = CPAGCLI%(NC%)
           If InStr(EPAC$, "AUXILIAR") > 0 Then
             If CONDPAGO% <= 1 Then
              TOTAVENTA# = (TOTAVENTA# / 3)
              Call ACTUREGISTRO("DEUDOR12", "Lcre_Cli", " NUME_CLI = " & NC%, FORMATNUM$(TOTAVENTA#, "F15.6"), REGAF&, "NOMESS")
             End If
           Else
              Select Case CONDPAGO%
                Case 0, 1: TOTAVENTA# = XVALO(!Lcre_Cli) ' PARA MANTENER EL VALOR DE AUXILIAR
                Case 5, 15: TOTAVENTA# = (TOTAVENTA# / 3) * 1.33 ' SOLICITO JUAN QUE SE CONSIDERE TAMBIEN 8 AL IGUAL QUE 5
                Case 3, 6: TOTAVENTA# = (TOTAVENTA# / 2) * 1.22
                Case 4, 12: TOTAVENTA# = (TOTAVENTA# / 1.57) * 1.17
                '
                'CASO QUE NO ENTRO EN NINGUNO QUE GRABE LIMITE DE CREDITO EN 0 (HABLADO CON JUAN PABLO)
                Case Else:
                TOTAVENTA# = 0
              End Select
              'REGISTRAR EL RESULTADO COMO LIMITE DE CREDITO
              Call ACTUREGISTRO("DEUDOR12", "Lcre_Cli", " NUME_CLI = " & NC%, FORMATNUM$(TOTAVENTA#, "F15.6"), REGAF&, "NOMESS")
           End If
        End If
        .MoveNext
        
      Loop
    End With
    '
    CLITMP.Close
    Set CLITMP = Nothing
    '
    Screen.MousePointer = 1
    '
    Call COMUNI("PROCESO FINALIZADO")
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
     'SE LEE EL ARCHIVO  FLEXCRL.DAT Y SE ACTUALIZA EN LA TABLA SQL
     'CON EL CRIETERIO QUE SI YA EXISTE EN SQL NO LO MODIFICA P QUE EL .DAT PUEDE ESTAR DESACTUALIZADO
     'CON RESPECTO AL SQL ,CONSIDERANDO QUE DESDE HACE UN TIEMPO EL GRACONTROL GRABA SOLO EN EL  SQL.
     '
     Command6.Enabled = False
     If CLACONTRA% < 1 Then GoTo 99
     '
     Screen.MousePointer = 11
    
     Call ABRECONEXION
     Call CREACAMPO("", "FLEXCRL", "PROBUS", 10, 32, 1)
     Call CREACAMPO("", "FLEXCRL", "CLABUS", 10, 32, 1)
     Call CREACAMPO("", "FLEXCRL", "VALCONTROL", 12, 0, 0)
     '
     Call TRALOCAL("FLEXCRL", "")
     '
     For i& = 1 To ULTREG&("!FLEXCRL")
          TX$ = REGLEIDO$("!FLEXCRL", i&)
          PROCE$ = TRIM$(Mid$(TX$, 1, 8))
          CLAVE$ = TRIM$(Mid$(TX$, 9, 8))
          VALCON$ = TRIM$(Mid$(TX$, 17, 40))
          CONDI$ = "PROBUS = '" & PROCE$ & "' AND CLABUS = '" & CLAVE$ & "'"
          If CantRegistros("FLEXCRL", CONDI$, "NOMESS") < 1 Then
10          RNDX$ = RANDSTRING$(-8)
            If CantRegistros("FLEXCRL", "PROBUS = '" & RNDX$ & "'", "NOMESS") > 0 Then GoTo 10
            Call CreaRegistro("FLEXCRL", "PROBUS", RNDX$, "NOMESS")
            Call ACTUREGISTRO("FLEXCRL", "CLABUS", "PROBUS = '" & RNDX$ & "'", CLAVE$, RAFE&, "NOMESS")
            Call ACTUREGISTRO("FLEXCRL", "VALCONTROL", "PROBUS = '" & RNDX$ & "'", VALCON$, RAFE&, "NOMESS")
            Call ACTUREGISTRO("FLEXCRL", "PROBUS", "PROBUS = '" & RNDX$ & "'", PROCE$, RAFE&, "NOMESS")
          End If
     Next i&
     '
     Call CIERRARCH("!FLEXCRL")
     Call CIERRARCH("FLEXCRL")
     '
99   Screen.MousePointer = 1
     Command6.Enabled = True

End Sub

Private Sub Command7_Click()
    '
    Call ABRECONEXION
    
    OPX% = COMALTER%("Opciones Complementarias\\Re-aplicacion de Comprobantes\Re-imputacion de Debitos\Re-aplicación de Comprobantes en forma Cronológica")
    If OPX% < 1 Or OPX% > 3 Then Exit Sub
    
    If OPX% = 1 Then GoTo 10
    If OPX% = 2 Then GoTo 20
    If OPX% = 3 Then GoTo 15
    Exit Sub
    
10  Dim CRED_COMP() As String, CRED_IMPO() As Single, CRED_COMC() As String
    Dim DEBI_COMP() As String, DEBI_IMPO() As Single, DEBI_COMC() As String
    
    SQLX$ = "SELECT * FROM DEUDOR12 WITH(NOLOCK) WHERE stat_cli>0 "
    SQLX$ = SQLX$ + " AND NUME_CLI<>10 AND NUME_CLI<>270 AND NUME_CLI<>457 AND NUME_CLI<>459 AND NUME_CLI<>536"
    SQLX$ = SQLX$ + " AND NUME_CLI<>576 AND NUME_CLI<>588 AND NUME_CLI<>591 AND NUME_CLI<>632 AND NUME_CLI<>670"
    SQLX$ = SQLX$ + " AND NUME_CLI<>686 AND NUME_CLI<>866 AND NUME_CLI<>939 AND NUME_CLI<>1262 AND NUME_CLI<>1271 AND NUME_CLI<>1332"
    SQLX$ = SQLX$ + " AND NUME_CLI<>851 AND NUME_CLI<>1118"
    Set rst1 = READSET(SQLX$)
    With rst1
        Do While Not .EOF
            NC% = XVALO(!nume_cli)
            GoSub 15
            .MoveNext
        Loop
    End With
    GoTo 19
    
15: Call SELECHO("DEUDOR1")
    NC% = XVALO(VALACT1$("DEUDOR1"))
    '
    ' ************* CREDITOS
    '
    CONDI$ = "((CODICOM_LAR LIKE 'AS-%' AND TIPOMOVIM='AJUD') OR (CODICOM_LAR LIKE 'AJ-%' AND TIPOMOVIM='FVEN') OR (TIPOMOVIM='AJUSTE') OR (TIPOMOVIM='FVEN' AND OPCIMOVIM='NC') OR (TIPOMOVIM='RCOB') ) AND NUCLIEN=" & CStr(NC%) & " AND IHABECOMP>0.001"
    FINCRE& = CantRegistros("CAJABANC", CONDI$)
    
    If FINCRE& < 1 Then
        Return
    End If
    Screen.MousePointer = 11
    
    ReDim CRED_COMP(FINCRE&)
    ReDim CRED_IMPO(FINCRE&)
    ReDim CRED_COMC(FINCRE&)
    i& = 0
    
    SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) WHERE ((CODICOM_LAR LIKE 'AS-%' AND TIPOMOVIM='AJUD') OR (CODICOM_LAR LIKE 'AJ-%' AND TIPOMOVIM='FVEN') OR (TIPOMOVIM='AJUSTE') OR (TIPOMOVIM='FVEN' AND OPCIMOVIM='NC') OR (TIPOMOVIM='RCOB') ) AND NUCLIEN=" & CStr(NC%) & " AND IHABECOMP>0.001 ORDER BY FECHCONTA ASC"
    Dim RSTCR As ADODB.Recordset
    Set RSTCR = READSET(SQLX$)
    
    With RSTCR
        Do While Not .EOF
            i& = i& + 1
            CRED_COMP(i&) = SAFETEXT$(!CODICOM_LAR)
            CRED_COMC(i&) = SAFETEXT$(!CODICOM_COR)
            CRED_IMPO(i&) = XVALO(!ValAbsComp)
            '
            .MoveNext
        Loop
    End With
    On Error Resume Next
    RSTCR.Close
    Set RSTCR = Nothing
    On Error GoTo 0
    '
    ' ************* DEBITOS
    '
    CONDI$ = "((CODICOM_LAR LIKE 'AS-%' AND TIPOMOVIM='AJUD') OR (CODICOM_LAR LIKE 'AJ-%' AND TIPOMOVIM='FVEN') OR (TIPOMOVIM='FVEN' AND OPCIMOVIM='FC') OR (TIPOMOVIM='FVEN' AND OPCIMOVIM='ND')) AND NUCLIEN=" & CStr(NC%) & " AND IDEBECOMP>0.001"
    FINDEB& = CantRegistros("CAJABANC", CONDI$)
    
    If FINDEB& < 1 Then
        Return
    End If
    ReDim DEBI_COMP(FINDEB&)
    ReDim DEBI_IMPO(FINDEB&)
    ReDim DEBI_COMC(FINDEB&)
    i& = 0
    
    SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) WHERE ((CODICOM_LAR LIKE 'AS-%' AND TIPOMOVIM='AJUD') OR (CODICOM_LAR LIKE 'AJ-%' AND TIPOMOVIM='FVEN') OR (TIPOMOVIM='FVEN' AND OPCIMOVIM='FC') OR (TIPOMOVIM='FVEN' AND OPCIMOVIM='ND')) AND NUCLIEN=" & CStr(NC%) & " AND IDEBECOMP>0.001 ORDER BY FECHCONTA ASC"
    Dim RSTDB As ADODB.Recordset
    Set RSTDB = READSET(SQLX$)
    
    With RSTDB
        Do While Not .EOF
            i& = i& + 1
            DEBI_COMP(i&) = SAFETEXT$(!CODICOM_LAR)
            DEBI_IMPO(i&) = XVALO(!ValAbsComp)
            DEBI_COMC(i&) = SAFETEXT$(!CODICOM_COR)
            .MoveNext
        Loop
    End With
    On Error Resume Next
    RSTDB.Close
    Set RSTDB = Nothing
    On Error GoTo 0
    '
    '
'    ' ************** REAPLICACION DE COMPROBANTES RESPETANDO LAS APLICACIONES ACTUALES
'    FF% = FECHANUM("14/01/13")
'    SQLX$ = "SELECT * FROM APLICOBRA WHERE NUCLIEN=" & NC%
'    Dim RSTAPLI As ADODB.Recordset
'    Set RSTAPLI = New ADODB.Recordset
'    RSTAPLI.Open SQLX$, FLEXCONN, adOpenKeyset, adLockPessimistic
'
'    ICRED& = 0: ENCONTRO% = 0
'    With RSTAPLI
'        Do While Not .EOF
'            DEBITO$ = SAFETEXT$(!codocor_lar)
'            CREDITO$ = SAFETEXT$(!codocap_lar)
'            IMPOAPLI = XVALO(!IMPOAPLI)
'            Fecha% = CVINT(VALOBADA("CAJABANC", "FECHCONTA", "CODICOM_LAR='" & CREDITO$ & "' AND NUCLIEN=" & NC%))
'            '
'            If Fecha% >= FF% Then
'                ' ACTUALIZAR VECTORES
'                For i& = 1 To FINCRE&
'                    If CRED_COMP(i&) = CREDITO$ Then
'                        ICRED& = i&: Exit For
'                    End If
'                Next i&
'                For i& = 1 To FINDEB&
'                    If DEBI_COMP(i&) = DEBITO$ Then
'                        CRED_IMPO(ICRED&) = CRED_IMPO(ICRED&) - IMPOAPLI
'                        DEBI_IMPO(i&) = DEBI_IMPO(i&) - IMPOAPLI
'                        ENCONTRO% = 1
'                        !CodiSucu = NC%
'                        .Update
'                        Exit For
'                    End If
'                Next i&
'                If ENCONTRO% = 0 Then
'                    ' ELIMINAR IMPUTACION
'                    !NUCLIEN = -1 * NC%
'                    .Update
'                End If
'                ENCONTRO% = 0
'            Else
'                ' ELIMINAR IMPUTACION
'                !NUCLIEN = -1 * NC%
'                .Update
'            End If
'            '
'            .MoveNext
'        Loop
'    End With
'    On Error Resume Next
'    RSTAPLI.Close
'    Set RSTAPLI = Nothing
'    On Error GoTo 0
    '
    Call BORRAREGISTROS("APLICOBRA", "NUCLIEN=" & CStr(NC%), REGAF&)
    '
    ' ************** REAPLICACION DE COMPROBANTES
    IC& = 1
    ID& = 1
    salir = False
    '
'    Call BORRAREGISTROS("APLICOBRA", "NUCLIEN=" & CStr(NC%), REGAF&)
'    If CantRegistros("APLICOBRA", "NUCLIEN=" & CStr(NC%)) > 0 Then
'        Call COMUNI("Imposible Eliminar Movimientos Aplicados")
'        Exit Sub
'    End If
    '
    MAXI& = FINCRE&
    If MAXI& > FINDEB& Then MAXI& = FINDEB&
    '
    Do While Not salir
        ii& = ii& + 1
        Call TRACE(24, ii&, MAXI&)
        '
        CRE = CRED_IMPO(IC&)
        DEB = DEBI_IMPO(ID&)
        '
        If CRE >= DEB Then
            CRED_IMPO(IC&) = CRED_IMPO(IC&) - DEB
            DEBI_IMPO(ID&) = 0
            ' APLICO
            NumeClie = NC%
            CoComLar = CRED_COMP(IC&)
            CoComCor$ = CRED_COMC(IC&)
            ImpApli# = XVALO(FORMATNUM(DEB, "F12.2"))
            RefApli$ = "Imputacion Manual"
            DocOrLar$ = DEBI_COMP(ID&)
            DocOrCor$ = DEBI_COMC(ID&)
            TotOrig# = XVALO(FORMATNUM(DEB, "F12.2"))
            IvaOrig# = XVALO(FORMATNUM(DEB - (DEB / 1.21), "F12.2"))
            Call GRAPLICOB
            ID& = ID& + 1
            If CRED_IMPO(IC&) < 0.005 Then IC& = IC& + 1
        Else
            DEBI_IMPO(ID&) = DEBI_IMPO(ID&) - CRE
            CRED_IMPO(IC&) = 0
            ' APLICO
            NumeClie = NC%
            CoComLar = CRED_COMP(IC&)
            CoComCor$ = CRED_COMC(IC&)
            ImpApli# = XVALO(FORMATNUM(CRE, "F12.2"))
            RefApli$ = "Imputacion Manual"
            DocOrLar$ = DEBI_COMP(ID&)
            DocOrCor$ = DEBI_COMC(ID&)
            TotOrig# = XVALO(FORMATNUM(CRE, "F12.2"))
            IvaOrig# = XVALO(FORMATNUM(CRE - (CRE / 1.21), "F12.2"))
            Call GRAPLICOB
            IC& = IC& + 1
            If DEBI_IMPO(ID&) < 0.005 Then ID& = ID& + 1
        End If
        If IC& > FINCRE& Then salir = True
        If ID& > FINDEB& Then salir = True
    Loop
    '
'    For i& = 1 To FINCRE&
'        If CRED_IMPO(i&) > 0.005 Then
'            Debug.Print CRED_COMP(i&) & "         " & CRED_IMPO(i&) & "        " & CStr(i&)
'        End If
'    Next i&
'
'    For i& = 1 To FINDEB&
'        If DEBI_IMPO(i&) > 0.005 Then
'            Debug.Print DEBI_COMP(i&) & "         " & DEBI_IMPO(i&) & "        " & CStr(i&)
'        End If
'    Next i&
    
    Screen.MousePointer = 1
    CONTROLGRABA% = 1
    Call RECUEC07.REAPLICRE(NC%)
    CONTROLGRABA% = 0
    GoTo 19
    
19  Call COMUNI("Re-Aplicacion Finalizada Para El Cliente '" & rasocli$(NC%) & "'")
    Exit Sub
    
20  Call COMUNI("Opción No Disponible")
    Exit Sub
    '
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=149", 225, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=151", 225, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=153", 225, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=154", 225, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=156", 225, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=158", 225, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=159", 225, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=161", 225, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=162", 225, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=163", 225, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=164", 225, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=165", 225, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=166", 225, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=167", 225, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=169", 225, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=170", 225, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=171", 225, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=172", 225, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=173", 225, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=175", 225, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=177", 225, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=179", 225, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=180", 225, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=182", 225, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=183", 225, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=184", 225, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=186", 225, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=187", 225, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=188", 225, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=191", 225, REGAF&)
    '
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=149", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=151", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=153", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=154", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=156", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=158", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=159", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=161", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=162", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=163", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=164", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=165", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=166", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=167", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=169", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=170", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=171", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=172", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=173", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=175", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=177", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=179", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=180", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=182", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=183", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=184", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=186", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=187", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=188", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=191", 214, REGAF&)
    '
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=150", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=152", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=155", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=160", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=168", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=174", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=176", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=178", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=181", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=185", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=189", 214, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='FVEN' AND OPCIMOVIM='ND' AND IDEBECOMP>0 AND NUMETALO=4 AND NUMECOMP=190", 214, REGAF&)
    '
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=150", 519, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=152", 519, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=155", 519, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=160", 519, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=168", 519, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=174", 519, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=176", 519, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=178", 519, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=181", 519, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=185", 519, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=189", 519, REGAF&)
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", "TIPOMOVIM='VENTA' AND OPCIMOVIM='ND' AND IHABECOMP>0 AND NUMETALO=4 AND NUMECOMP=190", 519, REGAF&)
    '
    Call COMUNI("Proceso Finalizado")
    '
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    '
    If CLACONTRA% > 0 Then
        'recorrer tabla deudor12 verificando los casos que coincide el domicilio del cliente con el de entrega , limpiar el de entrega
        CONDI$ = "Nume_Cli > 0 "
        REG& = CantRegistros&("DEUDOR12", CONDI$, "NOMESS")
        '
        SQLX$ = "SELECT Domi_Cli, Lentre0_Cli,Lentre1_Cli, Lentre2_Cli FROM DEUDOR12 WHERE " & CONDI$
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        rs.Open SQLX$, FLEXCONN, adOpenKeyset, adLockPessimistic
        '
        With rs
            Do While Not .EOF
              jj& = jj& + 1
              Call TRACE(20, jj&, REG&)
              DOMIFIS$ = UCase$(TRIM$(Left$(SAFETEXT$(!domi_cli), 48)))
              DOMIENTREGA$ = UCase$(TRIM$(Left$(SAFETEXT$(!Lentre0_Cli), 48)))
              If DOMIFIS$ = DOMIENTREGA$ Then
                 !Lentre0_Cli = ""
                 !Lentre1_Cli = ""
                 !Lentre2_Cli = ""
                 .Update
              End If
              .MoveNext
            Loop
        End With
        rs.Close
        Set rs = Nothing
    '
    End If
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   If CLACONTRA% > 0 Then
      OPX% = COMALTER%("Seleccione Rango de Fechas\Modificado solo para Cuit\Se Actualizara el Campo  Cuit_cli\En Cajabanc Tomando el Dato del Maestro de clientes\\Cancelar\Continuar")
      
      If OPX% < 2 Then GoTo 99
      
      Call DESHASFECHA(Des%, Has%, PERIO$)
      If Des% < 1 Then GoTo 99
    
      DES1$ = FETEXCOR1$(Des%)
      HAS1$ = FETEXCOR1$(Has%)

'      'SELECCION DE FECHA
'      HO1% = HOY(HO$)
'      Call REPLA(VDEF$, MKI$(HO1%), 1, 2)
'      OBJ1$ = OBJPLA$("ENTRAFECHA1", VDEF$)
'      If TRIM$(OBJ1$) = "" Then GoTo 99
'          '
'      FE% = CVI(Left$(OBJ1$, 2))
'      If FE% < 1 Then Call COMUNI("Fecha No Válida"): GoTo 99
'
'      FECHACORTE = FETEXCOR1$(FE%)
      '
      Call ABRECONEXION
      Call ABRECABAN

      'CARGO VECTORES CON LA POSICION DE IVA Y EL CUITCLI
      SQLX$ = "SELECT PIVA_CLI, Cuit_Cli,NUME_CLI FROM DEUDOR12 WITH(NOLOCK) ORDER BY NUME_CLI ASC"
      Dim PICLI%(), CUCLI$()
      CANTICLI = XVALO(VALOBADA("DEUDOR12", "MAX(NUME_CLI)", "NUME_CLI <> 0", "NOMESS"))
      If CANTICLI < 1 Then GoTo 99
      FIN& = CANTICLI
      ReDim Preserve PICLI%(CANTICLI), CUCLI$(CANTICLI)
      Set rs = READSET(SQLX$)
      With rs
        Do While Not .EOF
          i& = i& + 1
          Call TRACE(20, i&, FIN&)
          NC = XVALO(!nume_cli)
          'PICLI%(NC) = XVALO(!Piva_Cli)
          CUCLI$(NC) = TRIM$(SAFETEXT(!Cuit_Cli))
          .MoveNext
        Loop
      End With
      rs.Close
      Set rs = Nothing
   '
      'GRABO EN EL CAJABANC LOS DATOS DE LOS VECTORES EN LOS MOVIMIENTOS PREVIOS A LA FECHA SELECCIONADA
      CONDI$ = "NUCLIEN > 0 AND FECHEMISI >= '" & DES1$ & "' AND FECHEMISI <= '" & HAS1$ & "'"
      FIN& = CantRegistros("CAJABANC", CONDI$, "NOMESS")
      
      SQLX$ = "SELECT POS_IVA, CUIT_CLI,NUCLIEN FROM CAJABANC"
      SQLX$ = SQLX$ + " WHERE " & CONDI$
      SQLX$ = SQLX$ + " ORDER BY NUCLIEN"
      i& = 0
      Dim rs1 As ADODB.Recordset
      Set rs1 = New ADODB.Recordset
      rs1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
      With rs1
          Do While Not .EOF
            i& = i& + 1
            Call TRACE(20, i&, FIN&)
            NC = XVALO(!nuclien)
            '!POS_IVA = PICLI%(NC)
            !Cuit_Cli = CUCLI$(NC)
            .Update
            .MoveNext
          Loop
      End With
10    rs1.Close
      Set rs1 = Nothing
   End If
   
99 Command9.Enabled = True
End Sub

'
Private Sub Form_Load()
  '
  UTILIZA_SKIN% = 1
  Call APLICACIONSKINS(Me)
  '
  Screen.MousePointer = 11
  Call CARTAVALI
  Screen.MousePointer = 1
  '
  'Call APLICACIONSKINS(Me)
  TABLA$(0) = "TATIMAT"       ' tabla de tipos de material
  TABLA$(1) = "*UNIMED"
  TABLA$(2) = "TADEPOSI"
  TABLA$(3) = ""       ' grupos contables por articulo
  TABLA$(4) = "*VENDEDOR"
  TABLA$(5) = "TAPROVIN"
  TABLA$(6) = "CONPAGO"
  TABLA$(7) = "LPREAC07"
  TABLA$(8) = "CUEINGAS"
  TABLA$(9) = "GRUCOCLI"
  TABLA$(10) = "ZONAREP"
  '
  PRONUCLI$ = Control$("", "PRONUCLI")
  If XVALO(PRONUCLI$) >= 0 Then
    Option3.Value = True
    Text3 = PRONUCLI$
  End If
  '
  For Index = 0 To 7
    VAPRE$ = TRIM$(Control$("AMACLI", ARCHE$(Index)))
    COTEXT(Index) = "": DETEXT(Index) = ""
    If VAPRE$ <> "" Then
      TV11% = TIPFIELD(ARCHE$(Index), 1)
      LF11% = LENFIELD(ARCHE$(Index), 1)
      ARGU$ = CBIN$(VAPRE$, TV11%, LF11%)
      ECO$ = TRIM$(ECOARCH$(ARCHE$(Index), ARGU$))
      If ECO$ <> "" Then
          COTEXT(Index).TEXT = VAPRE$
          DETEXT(Index).TEXT = ECO$
        Else
          COTEXT(Index).TEXT = ""
          DETEXT(Index).TEXT = ""
       End If
    End If
    DETEXT(Index).Refresh
    DETEXT(Index).TabStop = False
  Next Index
  ' '\\\OT-05-0198-WAR-19/04/05-///
  If Control$("FORMCUIT", "SUPRIVAL") = "SI" Then
      Check1.Value = 1
  End If
  '''\\\OT-05-0198-WAR-FIN///
  '
  Check6.Value = 0
  DOMIFI$ = Control("MASTER", "ENTREPFI")
  If DOMIFI$ = "NO" Then Check6.Value = 1
  '
  Check2.Value = 0
  RASOMAYU$ = Control("MASTER", "RASOMAYU")
  If RASOMAYU$ = "SI" Then Check2.Value = 1

  EPAC$ = TRIM$(UCase$(EMPREAC$))
  If InStr(EPAC$, "GABAL") > 0 Then
     Command5.Visible = True 'BOTON DE CALCULO PROMEDIO DE VENTAS DE LOS  ULTIMOS 3 MESES + 33 %
  End If                     'RESULTADO ASIGNADO COMO LIMITE DE CREDITO

  EPAC$ = TRIM$(UCase$(EMPREAC$))
  If InStr(EPAC$, "FERVI") > 0 Then
     Command7.Visible = True 'BOTON DE CALCULO PROMEDIO DE VENTAS DE LOS  ULTIMOS 3 MESES + 33 %
  End If
  
  Check3.Value = 0
  If Control("MASTER", "DOMIMAND") = "SI" Then Check3.Value = 1
  '
  Check4.Value = 0
  If Control("AMACLI", "CONSICUI") = "SI" Then Check4.Value = 1
  '
End Sub

Private Sub mnuTransacciones_Click()

End Sub

Private Sub mnuSincrprov_Click()
        VENTANA.AgregaRegistro = 1
        VENTANA.UtilizaFormula = 1
        VENTANA.HabilitaBorrar = 0
        VENTANA.HabilitaInsertar = 0
        VENTANA.NoMostrarBotonTilde = 0
        VENTANA.TITULO = ""
        VENTANA.Inicializar = 1
        VENTANA.CampEditables = "4"
        VENTANA.Campox = "Cod.Oficial/A10;Provincia/A32;ID/F8;CodFlexoft/A10"

        Dim obj As New RECORXET
        Set rs1 = obj.RS_Tabla_Provincias_Config()
        Call Carga_MSF_Recordset(rs1, "CodOficial/A10;Provincia/A32;ID/F8;CodFlex/A10", VENTABNEW.MSF)
        Set obj = Nothing
        rs1.Close
        Set rs1 = Nothing
        '*****************************************************'
        VENTABNEW.Show 1
        If VENTABNEW.APROBADO > 0 Then
            Dim MSF As msFlexGrid
            Set MSF = VENTABNEW.MSF
    
    
            APARTIR% = 1
            SQLX$ = "SELECT * FROM PROVINCIA ORDER BY ID"
            Dim rs As ADODB.Recordset
            Set rs = New ADODB.Recordset
25          On Error Resume Next
            rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
            If Err Then
             On Error GoTo 0
             Call CapturaErrorOpen
             GoTo 25
            End If
              With rs
                Do While Not .EOF
                  IDX = XVALO(!ID)
                  For REG& = APARTIR% To MSF.Rows - 1
                   If XVALO(MSF.TextMatrix(REG&, 3)) = IDX Then
                      !CODFLEX = TRIM$(Left(MSF.TextMatrix(REG&, 4), 24))
                      .Update
                      APARTIR% = REG& + 1
                      GoTo 30
                    End If
                   Next REG&
30                 .MoveNext
                 Loop
              End With
            
              rs.Close
              Set rs = Nothing
        End If
        Set VENTANA = Nothing

End Sub

'
Private Sub Option1_Click(Index As Integer)
    Option1(Index).Value = True
    Text2.TEXT = Option1(Index).Caption
    MODUCONE$ = TABLA$(Index)
    TETABLA$ = MODUCONE$
    If Left$(TETABLA$, 1) = "#" Then
        TETABLA$ = Mid$(TETABLA$, 2) + "(#)"
      ElseIf Left$(TETABLA$, 1) = "*" Then
        TETABLA$ = Mid$(TETABLA$, 2) + "(*)"
    End If
    Text1.TEXT = TETABLA$
    '
    For k% = 4 To 9
       If k% <> Index Then
          Option1(k%).Value = False
       End If
    Next k%
    'Command16.SetFocus
End Sub

 Sub Option1_DblClick(Index As Integer)
    '
    If TRIM$(MODUCONE$) = "" Then
       Call Form_Load
       MODUCONE$ = TABLA$(Index)
    End If
    '
    If TRIM$(MODUCONE$) <> "" Then
        If TRIM$(MODUCONE$) = "GRUCOCLI" Then
            LDT$ = LUDAT$
            Call CIERRARCH("GRUCOCLI")
            Call CIERRARCH("!GRUCOCLI")
            'Call COPYFILE(LDT$ + "GRUCOCLI.DAT", "C:\FLEXOFT\GRUCOCLI.WKF")
            Call TRALOCAL("GRUCOCLI", "")
            VTB% = VENTABU%("GRUCOCLI")
            If VTB% < 0 Then Exit Sub
            J& = 0
            For i& = 1 To ULTREG&("!GRUCOCLI")
                X$ = REGLEIDO$("!GRUCOCLI", i&)
                If CVI(Left$(X$, 2)) > 0 Then
                    J& = J& + 1
                    Call GRAREG("GRUCOCLI", X$, J&)
                End If
            Next i&
            Call SETULTREG("GRUCOCLI", J&)
            Call CIERRARCH("GRUCOCLI")
          ElseIf MODUCONE = "TADEPOSI" Then
            Call CARTADEPOSI
          ElseIf MODUCONE = "TATIMAT" Then
            Call CARTATIMA
          ElseIf MODUCONE = "CONPAGO" Then
            Call CARCONPAGOS
          ElseIf MODUCONE = "TAPROVIN" Then
            Call CARTAPROVIN
          ElseIf MODUCONE = "CUEINGAS" Then
            Call CARCUEINGAS
          ElseIf MODUCONE = "*VENDEDOR" Then
                'RTA% = ALTERNA%("ABM Vendedores\Direccion de E-mail")
                RTA% = ALTERNA%("ABM Vendedores\Datos Complementarios")
                If RTA% = 1 Then
                    Call CONECRUN(MODUCONE$, Text2.TEXT)
                ElseIf RTA% = 2 Then
                    Call CARTAVENDE     'CARGA DATOS COMPLEMENTARIOS DEL VENDEDOR
                Else
                    Exit Sub
                End If
          ElseIf MODUCONE = "ZONAREP" Then
            Call CARZONAREP
          
          Else
            Call CONECRUN(MODUCONE$, Text2.TEXT)
        End If
    End If

End Sub


Private Sub Option11_DblClick()
    '
    Call COMUNI("Opción No Disponible")
'    If DERACCE%("DESCOMBI", "DESCUENTOS COMBINADOS") >= 2 Then
'        DEFDESC07.Show 1
'    End If
    '
End Sub

Private Sub Option2_Click()
   Option2.Value = True
   Text3.TEXT = ""
End Sub

Private Sub Option3_Click()
   Option3.Value = True
End Sub

Private Sub Option4_Click()

End Sub

Private Sub Text1_GotFocus()
   Command16.SetFocus
End Sub

Private Sub Text2_gotfocus()
   Command16.SetFocus
End Sub

Private Sub Text3_Click()
   Option3.Value = True
End Sub



Private Sub COTEXT_DblClick(Index As Integer)
    Call SELECHO(ARCHE$(Index))
    If TRIM$(VALACT1$(ARCHE$(Index))) <> "" Then
        COTEXT(Index).TEXT = VALACT1$(ARCHE$(Index))
        COTEXT(Index).Refresh
        DETEXT(Index).TEXT = VALACT2$(ARCHE$(Index))
        DETEXT(Index).Refresh
        DETEXT(Index).TabStop = False
    End If
    COTEXT(Index).SetFocus
End Sub
'
Private Sub BOTSEL_Click(Index As Integer)
    Call SELECHO(ARCHE$(Index))
    If TRIM$(VALACT1$(ARCHE$(Index))) <> "" Then
        COTEXT(Index).TEXT = VALACT1$(ARCHE$(Index))
        COTEXT(Index).Refresh
        DETEXT(Index).TEXT = VALACT2$(ARCHE$(Index))
        DETEXT(Index).Refresh
        DETEXT(Index).TabStop = False
    End If
    COTEXT(Index).SetFocus
End Sub
'
Private Sub COTEXT_LOSTFOCUS(Index As Integer)
    If TRIM$(COTEXT(Index).TEXT) <> "" Then
        TV11% = TIPFIELD(ARCHE$(Index), 1)
        LF11% = LENFIELD(ARCHE$(Index), 1)
        ARGU$ = CBIN$(TRIM$(COTEXT(Index).TEXT), TV11%, LF11%)
        ECO$ = TRIM$(ECOARCH$(ARCHE$(Index), ARGU$))
        If ECO$ <> "" Then
             COTEXT(Index).TEXT = TRIM$(COTEXT(Index).TEXT)
             DETEXT(Index).TEXT = ECO$
             DETEXT(Index).Refresh
             DETEXT(Index).TabStop = False
           Else
             DETEXT(Index).TEXT = ""
             Exit Sub
        End If
      Else
        DETEXT(Index).TEXT = ""
    End If
End Sub
'

Private Sub COTEXT_KeyPress(Index As Integer, KeyAscii As Integer)
    '
    If KeyAscii = 13 Then
       If TRIM$(COTEXT(Index).TEXT) <> "" Then
          TV11% = TIPFIELD(ARCHE$(Index), 1)
          LF11% = LENFIELD(ARCHE$(Index), 1)
          ARGU$ = CBIN$(TRIM$(COTEXT(Index).TEXT), TV11%, LF11%)
          ECO$ = TRIM$(ECOARCH$(ARCHE$(Index), ARGU$))
          If ECO$ <> "" Then
               DETEXT(Index).TEXT = ECO$
               DETEXT(Index).Refresh
               DETEXT(Index).TabStop = False
             Else
               DETEXT(Index).TEXT = ""
               COTEXT(Index).SetFocus
          End If
       End If
    End If
End Sub

Private Sub DETEXT_CLICK(Index As Integer)
    COTEXT(Index).SetFocus
End Sub


Sub CARZONAREP()
    If TRIM$(MODUCONE$) = "ZONAREP" Then
      LDT$ = LUDAT$
      Call CIERRARCH("ZONAREP")
      Call CIERRARCH("!ZONAREP")
      Call TRALOCAL("ZONAREP", "")
      VTB% = VENTABU%("ZONAREP")
      If VTB% < 0 Then GoTo 99
      J& = 0
      For i& = 1 To ULTREG&("!ZONAREP")
          X$ = REGLEIDO$("!ZONAREP", i&)
          If TRIM$(Left$(X$, 8)) <> "" Then
              J& = J& + 1
              Call GRAREG("ZONAREP", X$, J&)
          End If
      Next i&
      Call SETULTREG("ZONAREP", J&)
      Call CIERRARCH("ZONAREP")
    End If
99 End Sub

