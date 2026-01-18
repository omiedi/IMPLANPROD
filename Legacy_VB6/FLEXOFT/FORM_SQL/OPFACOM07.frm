VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form OPFACOM07 
   BackColor       =   &H00C7D9CC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración Proceso de Facturación de Compras"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   11805
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
   ScaleHeight     =   7860
   ScaleWidth      =   11805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame9 
      BackColor       =   &H00C7D9CC&
      Caption         =   "CTA. CONTABLE PARA IMPORTES  CIF (Factura Aduana)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   5520
      TabIndex        =   88
      Top             =   6600
      Width           =   5265
      Begin VB.CommandButton Command7 
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
         Left            =   2430
         TabIndex        =   89
         Top             =   360
         Width           =   175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Importes Cif"
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
         Index           =   9
         Left            =   240
         TabIndex        =   92
         Top             =   360
         Width           =   1155
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
         Left            =   1470
         TabIndex        =   91
         Top             =   360
         Width           =   975
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
         Left            =   1470
         TabIndex        =   90
         Top             =   675
         Width           =   3525
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C7D9CC&
      Caption         =   "CTA. CONTABLE RECAUDACIONES BANCARIAS (SIRCREB)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   105
      TabIndex        =   83
      Top             =   6600
      Width           =   5265
      Begin VB.CommandButton Command32 
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
         Left            =   2430
         TabIndex        =   84
         Top             =   360
         Width           =   175
      End
      Begin VB.Label DENOSIRCREB 
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
         Left            =   1470
         TabIndex        =   87
         Top             =   675
         Width           =   3525
      End
      Begin VB.Label SIRCREB 
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
         Left            =   1470
         TabIndex        =   86
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Recaudaciones Bancarias"
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
         Index           =   8
         Left            =   120
         TabIndex        =   85
         Top             =   260
         Width           =   1275
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5640
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C7D9CC&
      Caption         =   "CONFIGURACIÓN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   5520
      TabIndex        =   75
      Top             =   5520
      Width           =   5250
      Begin VB.CommandButton Command77 
         Caption         =   "Detalle de Gastos"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3680
         TabIndex        =   78
         Top             =   200
         Width           =   1500
      End
      Begin VB.CheckBox Check5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7D9CC&
         Caption         =   "No Considerar Moneda Según Tabla Proveedor/Moneda"
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
         Height          =   330
         Left            =   180
         TabIndex        =   77
         Top             =   550
         Width           =   5000
      End
      Begin VB.CheckBox Check4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7D9CC&
         Caption         =   "No Edita Ctro/Ctos en Carga de Detalle"
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
         Height          =   330
         Left            =   180
         TabIndex        =   76
         ToolTipText     =   $"OPFACOM07.frx":0000
         Top             =   250
         Width           =   3435
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C7D9CC&
      Caption         =   "CTA. CONTABLE POR DFECTO P/CARGA POR DETALLE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   5520
      TabIndex        =   69
      Top             =   4300
      Width           =   5250
      Begin VB.CommandButton CMDCTA 
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
         Left            =   4800
         TabIndex        =   70
         Top             =   360
         Width           =   180
      End
      Begin VB.Label TXTDESCRIP 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1155
         TabIndex        =   73
         Top             =   720
         Width           =   3870
      End
      Begin VB.Label TXTCTA 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3150
         TabIndex        =   72
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta"
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
         Left            =   1965
         TabIndex        =   71
         Top             =   435
         Width           =   1065
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C7D9CC&
      Caption         =   "DIFERENCIA DE CAMBIO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   5520
      TabIndex        =   64
      Top             =   3050
      Width           =   5250
      Begin VB.CommandButton Command3 
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
         Left            =   4800
         TabIndex        =   65
         Top             =   360
         Width           =   180
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta"
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
         Left            =   1965
         TabIndex        =   68
         Top             =   435
         Width           =   1065
      End
      Begin VB.Label CUEIVA 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   3150
         TabIndex        =   67
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label DECUENTA 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   1155
         TabIndex        =   66
         Top             =   720
         Width           =   3870
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C7D9CC&
      Caption         =   "LETRA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8880
      TabIndex        =   62
      Top             =   1000
      Width           =   1890
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7D9CC&
         Caption         =   "Distingue Letra"
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
         Height          =   215
         Left            =   100
         TabIndex        =   63
         Top             =   345
         Width           =   1650
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   1320
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   60
      Top             =   1200
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C7D9CC&
      Caption         =   "REDONDEO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5520
      TabIndex        =   55
      Top             =   60
      Width           =   5250
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
         Height          =   285
         Left            =   2310
         TabIndex        =   56
         Top             =   330
         Width           =   175
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta"
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
         Left            =   0
         TabIndex        =   59
         Top             =   405
         Width           =   750
      End
      Begin VB.Label Label11 
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
         Left            =   840
         TabIndex        =   58
         Top             =   330
         Width           =   1485
      End
      Begin VB.Label Label8 
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
         Left            =   2625
         TabIndex        =   57
         Top             =   330
         Width           =   2430
      End
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00C7D9CC&
      Caption         =   "MONEDA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1180
      Left            =   8865
      TabIndex        =   51
      Top             =   1830
      Width           =   1890
      Begin VB.OptionButton Option15 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7D9CC&
         Caption         =   "Bi-Monetario"
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
         Height          =   240
         Left            =   210
         TabIndex        =   54
         Top             =   615
         Width           =   1600
      End
      Begin VB.OptionButton Option16 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7D9CC&
         Caption         =   "Pesos"
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
         Height          =   240
         Left            =   210
         TabIndex        =   53
         Top             =   330
         Value           =   -1  'True
         Width           =   1600
      End
      Begin VB.OptionButton Option19 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7D9CC&
         Caption         =   "Multi-Moneda"
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
         Height          =   240
         Left            =   210
         TabIndex        =   52
         Top             =   900
         Width           =   1600
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C7D9CC&
      Caption         =   "LIBRO IVA COMPRAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5520
      TabIndex        =   49
      Top             =   2070
      Width           =   3270
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7D9CC&
         Caption         =   "Suprime Asientos de Diario"
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
         Left            =   420
         TabIndex        =   50
         Top             =   240
         Width           =   2790
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C7D9CC&
      Caption         =   "VALIDACIONES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5520
      TabIndex        =   46
      Top             =   1000
      Width           =   3270
      Begin VB.CheckBox Check6 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7D9CC&
         Caption         =   "Posición de IVA con AFIP"
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
         Height          =   210
         Left            =   420
         TabIndex        =   93
         Top             =   600
         Width           =   2685
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7D9CC&
         Caption         =   "Requiere Ingreso del C.A.I."
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
         Left            =   420
         TabIndex        =   47
         Top             =   220
         Width           =   2685
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0050765A&
      ForeColor       =   &H80000008&
      Height          =   8160
      Left            =   10920
      ScaleHeight     =   8130
      ScaleWidth      =   945
      TabIndex        =   39
      Top             =   0
      Width           =   975
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H000C93DC&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   79
         Top             =   7440
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0076C9F5&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   80
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command10 
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
         Picture         =   "OPFACOM07.frx":009B
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Clasificación de Proveedores por Moneda de Compra y Valuación"
         Top             =   3720
         Width           =   650
      End
      Begin VB.CommandButton Command9 
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
         Left            =   120
         Picture         =   "OPFACOM07.frx":01E5
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Cotización Dolar y Tipo de Cambio"
         Top             =   5160
         Width           =   650
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Tablas"
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
         Left            =   120
         Picture         =   "OPFACOM07.frx":04EF
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Tablas de Grupos, Familias e Items y Conceptos"
         Top             =   3000
         Width           =   650
      End
      Begin VB.CommandButton Command6 
         Caption         =   "C-Cos"
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
         Left            =   120
         Picture         =   "OPFACOM07.frx":0639
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Base de Datos de Centros de Costos"
         Top             =   2280
         Width           =   650
      End
      Begin VB.CommandButton Command5 
         Caption         =   "I-B"
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
         Left            =   120
         MaskColor       =   &H00800000&
         Picture         =   "OPFACOM07.frx":0943
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Percepción Ingresos Brutos - Tabla de Provincias"
         Top             =   4440
         Width           =   650
      End
      Begin VB.CommandButton Command4 
         Caption         =   "C-Con"
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
         Left            =   120
         Picture         =   "OPFACOM07.frx":0A8D
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Base de Datos de Cuentas Contables"
         Top             =   1560
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
         Left            =   120
         Picture         =   "OPFACOM07.frx":0BD7
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Guardar Configuración"
         Top             =   840
         Width           =   650
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
         Left            =   120
         Picture         =   "OPFACOM07.frx":0EE1
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Finaliza la Aplicación"
         Top             =   120
         Width           =   650
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00C7D9CC&
      Caption         =   "TASAS DE IVA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4450
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   60
      Width           =   5265
      Begin VB.CommandButton Command3 
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
         Left            =   4830
         TabIndex        =   34
         Top             =   3640
         Width           =   175
      End
      Begin VB.TextBox TASIVA 
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
         Left            =   1155
         TabIndex        =   22
         Text            =   " "
         Top             =   2800
         Width           =   855
      End
      Begin VB.CommandButton Command3 
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
         Left            =   4830
         TabIndex        =   19
         Top             =   2800
         Width           =   175
      End
      Begin VB.TextBox TASIVA 
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
         Left            =   1155
         TabIndex        =   16
         Text            =   " "
         Top             =   1960
         Width           =   855
      End
      Begin VB.CommandButton Command3 
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
         Left            =   4830
         TabIndex        =   13
         Top             =   1960
         Width           =   175
      End
      Begin VB.TextBox TASIVA 
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
         Left            =   1155
         TabIndex        =   10
         Text            =   " "
         Top             =   1120
         Width           =   855
      End
      Begin VB.CommandButton Command3 
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
         Left            =   4830
         TabIndex        =   7
         Top             =   1120
         Width           =   175
      End
      Begin VB.TextBox TASIVA 
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
         Left            =   1155
         TabIndex        =   4
         Text            =   " "
         Top             =   280
         Width           =   855
      End
      Begin VB.CommandButton Command3 
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
         Left            =   4830
         TabIndex        =   1
         Top             =   280
         Width           =   175
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta"
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
         Left            =   1995
         TabIndex        =   38
         Top             =   3715
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "I.V.A. Aduana  -"
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
         Left            =   740
         TabIndex        =   37
         Top             =   3715
         Width           =   1695
      End
      Begin VB.Label CUEIVA 
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
         Left            =   3150
         TabIndex        =   36
         Top             =   3640
         Width           =   1695
      End
      Begin VB.Label DECUENTA 
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
         Left            =   1155
         TabIndex        =   35
         Top             =   4000
         Width           =   3870
      End
      Begin VB.Label DECUENTA 
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
         Left            =   1155
         TabIndex        =   24
         Top             =   3160
         Width           =   3870
      End
      Begin VB.Label CUEIVA 
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
         Left            =   3150
         TabIndex        =   23
         Top             =   2800
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TASA-3"
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
         Left            =   210
         TabIndex        =   21
         Top             =   2850
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta"
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
         Left            =   1995
         TabIndex        =   20
         Top             =   2875
         Width           =   1065
      End
      Begin VB.Label DECUENTA 
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
         Left            =   1155
         TabIndex        =   18
         Top             =   2355
         Width           =   3870
      End
      Begin VB.Label CUEIVA 
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
         Left            =   3150
         TabIndex        =   17
         Top             =   1960
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TASA-2"
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
         Left            =   210
         TabIndex        =   15
         Top             =   2000
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta"
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
         Left            =   1995
         TabIndex        =   14
         Top             =   2035
         Width           =   1065
      End
      Begin VB.Label DECUENTA 
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
         Left            =   1155
         TabIndex        =   12
         Top             =   1515
         Width           =   3870
      End
      Begin VB.Label CUEIVA 
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
         Left            =   3150
         TabIndex        =   11
         Top             =   1120
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TASA-1"
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
         Left            =   105
         TabIndex        =   9
         Top             =   1190
         Width           =   960
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta"
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
         Left            =   1995
         TabIndex        =   8
         Top             =   1230
         Width           =   1065
      End
      Begin VB.Label DECUENTA 
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
         Left            =   1155
         TabIndex        =   6
         Top             =   640
         Width           =   3870
      End
      Begin VB.Label CUEIVA 
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
         Left            =   3150
         TabIndex        =   5
         Top             =   280
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Básica"
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
         Left            =   210
         TabIndex        =   3
         Top             =   355
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta"
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
         Left            =   1995
         TabIndex        =   2
         Top             =   355
         Width           =   1065
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00C7D9CC&
      Caption         =   "CTAS. CONTABLES PERCEPCIONES, RETENCIONES y SIAP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Index           =   1
      Left            =   105
      TabIndex        =   25
      Top             =   4560
      Width           =   5265
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
         Left            =   2430
         TabIndex        =   94
         Top             =   1500
         Width           =   175
      End
      Begin VB.TextBox Text100 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   4440
         TabIndex        =   81
         Text            =   " "
         ToolTipText     =   "Código de Régimen de Percepción"
         Top             =   360
         Width           =   540
      End
      Begin VB.CommandButton Command3 
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
         Left            =   2430
         TabIndex        =   27
         Top             =   1125
         Width           =   175
      End
      Begin VB.CommandButton Command3 
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
         Left            =   2430
         TabIndex        =   26
         Top             =   405
         Width           =   175
      End
      Begin VB.Label lblDescripcionRetIIBB 
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
         Left            =   2640
         TabIndex        =   97
         Top             =   1500
         Width           =   2325
      End
      Begin VB.Label lblRetIIBB 
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
         Left            =   1440
         TabIndex        =   96
         Top             =   1500
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Retención IIBB"
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
         Left            =   120
         TabIndex        =   95
         Top             =   1590
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código Régimen de Percepción"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   5
         Left            =   3150
         TabIndex        =   82
         ToolTipText     =   "Código de Régimen de Percepción"
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Percepcion IVA"
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
         Left            =   120
         TabIndex        =   33
         Top             =   480
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Perc. Ganancias "
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
         Left            =   120
         TabIndex        =   32
         Top             =   1200
         Width           =   1275
      End
      Begin VB.Label CUEIVA 
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
         Left            =   1440
         TabIndex        =   31
         Top             =   1125
         Width           =   975
      End
      Begin VB.Label DECUENTA 
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
         Left            =   2640
         TabIndex        =   30
         Top             =   1125
         Width           =   2325
      End
      Begin VB.Label CUEIVA 
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
         Left            =   1440
         TabIndex        =   29
         Top             =   405
         Width           =   975
      End
      Begin VB.Label DECUENTA 
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
         Left            =   1440
         TabIndex        =   28
         Top             =   720
         Width           =   3525
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "OPFACOM07.frx":102B
      TabIndex        =   61
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "OPFACOM07.frx":10A5
      Top             =   1200
   End
   Begin VB.Menu mnuConfiguracion 
      Caption         =   "Configuración"
      Begin VB.Menu mnuPercepcionesIva 
         Caption         =   "Percepciones IVA"
         Begin VB.Menu mnuPercepIvaProveedoresImportacion 
            Caption         =   "Proveedores de Importación"
         End
      End
   End
End
Attribute VB_Name = "OPFACOM07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub CargaTablasAuxiliares()

    Dim A As New DatosAuxiliares
    Dim a1() As eDatosAuxiliaresCampos
    Dim a2 As eDatosAuxiliaresCampos
    
    ' TABLA PROVEEDORES DE IMPORTACION PARA PERCEPCION DE IVA
    ' *******************************************************
    
    Set a2 = New eDatosAuxiliaresCampos
    a2.Encabezado = "Ord/I10"
    a2.Formula = ""
    a2.CampoEditable = "N"
    a2.ColumnaCampoEco = ""
    a2.Total = ""
    a2.DejarLibre = ""
    a2.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve a1(n)
    Set a1(n - 1) = a2
    
    Set a2 = New eDatosAuxiliaresCampos
    a2.Encabezado = "Cuenta/I10"
    a2.Formula = ""
    a2.CampoEditable = "S"
    a2.ColumnaCampoEco = "NUME_CLI;RASO_CLI;PROVED12;STAT_CLI>0"
    a2.Total = ""
    a2.DejarLibre = ""
    a2.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve a1(n)
    Set a1(n - 1) = a2
        
    Set a2 = New eDatosAuxiliaresCampos
    a2.Encabezado = "Razón Social/A48"
    a2.Formula = ""
    a2.CampoEditable = "N"
    a2.ColumnaCampoEco = ""
    a2.Total = ""
    a2.DejarLibre = ""
    a2.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve a1(n)
    Set a1(n - 1) = a2
           
    Call A.CrearTabla("PERCIVAPROIMPO", "PROVEEDORES DE IMPORTACION", a1)
    
End Sub

Private Sub CMDCTA_Click()
    Call SELECHO("CUECON")
    VDEVU$ = VALACT1$("CUECON")
    If VDEVU$ <> "" Then
        Me.TXTCTA.Caption = CODCUE$(XVALO(VDEVU$))
    End If

End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Sub Command10_Click()
   Command10.Enabled = False
   '
   Screen.MousePointer = 11
   '
   UREVI& = Val(CONTROL$("OCOMPRA", "RMONEPRO")) - 128
   If UREVI& < 1 Then UREVI& = 1
   '
   FIN& = ULTREG&("OCOMENCA")
   AGREPRO% = 0
   PUNPROAC$ = ""
   For U& = UREVI& To FIN&
     XX$ = REGLEIDO$("OCOMENCA", U&)
     NPRO% = CVI(Mid$(XX$, 51, 2))
     If NPRO% > 0 Then
       RFF$ = RECFILT$("CLAPROMO", 1, MKI$(NPRO%))
       If Len(RFF$) < 4 Then
         For KKI% = 1 To Len(PUNPROAC$) Step 2
           If CVI(Mid$(PUNPROAC$, KKI%, 2)) = NPRO% Then GoTo 19
         Next KKI%
         YY$ = REGBLAN$("CLAPROMO")
         Call REPLA(YY$, MKI$(NPRO%), 1, 2)
         Call REPLA(YY$, rasocli$((-1) * NPRO%), 3, 28)
         Call REPLA(YY$, Chr$(1), 31, 1)
         Call REPLA(YY$, Chr$(1), 32, 1)
         Call REGAPP("CLAPROMO", YY$)
         Call CIERRARCH("CLAPROMO")
         PUNPROAC$ = PUNPROAC$ + MKI$(NPRO%)
         AGREPRO% = 1
       End If
     End If
19 Next U&
   '
   If AGREPRO% > 0 Then
     Call FILESORT("CLAPROMO", "", 2, 2, "ASC")
   End If
   '
   Screen.MousePointer = 1
   Call TRALOCAL("CLAPROMO", "")
   VTB% = VENTABU%("CLAPROMO")
   If VTB% >= 0 Then
     Call TRACENTRAL("CLAPROMO", "")
     Call GRACONTROL("OCOMPRA", "RMONEPRO", Str$(ULTREG&("OCOMENCA")))
   End If
   '
   Command10.Enabled = True
End Sub
Private Sub Command14_Click()
   Call CommSel("ECUECON", Command14, Label11)
End Sub

Private Sub Command2_Click()
    '
    For IXX% = 0 To CUEIVA.UBound
      TXXX$ = TRIM$(Str$(IXX%))
      Call GRACONTROL$("", "CUEIVA-" + TXXX$, TRIM$(CUEIVA(IXX%).Caption))
      If IXX% < 4 Then
         TASAX$ = REPLACAR$(TASIVA(IXX%).TEXT, ",", ".")
         Call GRACONTROL$("", "TASIVA-" + TXXX$, TASAX$)
         TASA = XVALO(TASAX$)
         On Error Resume Next
         If TASA > 0.005 Then
             FORFACOM.TASIVA(IXX%).Caption = TRIM$(CSTRING$(MKS$(TASA), 3, 6, 2, 2) + "%")
             FORFAC00.TASIVA(IXX%).Caption = TRIM$(CSTRING$(MKS$(TASA), 3, 6, 2, 2) + "%")
           Else
             FORFACOM.TASIVA(IXX%).Caption = ""
             FORFAC00.TASIVA(IXX%).Caption = ""
         End If
         On Error GoTo 0
      End If
    Next IXX%
    If Check1.Value Then
      Call GRACONTROL("", "VALICAI", "SI")
    Else
      Call GRACONTROL("", "VALICAI", "NO")
    End If
    '\\\OT-06-0208-WAR-05/06/06///
    If Check2.Value Then
      Call GRACONTROL("LIVACOM", "SUPRIVA", "SI")
    Else
      Call GRACONTROL("LIVACOM", "SUPRIVA", "NO")
    End If
    '\\\OT-06-0208-WAR-FIN///
    '
    Call GRACONTROL("FORFACPR", "CTAXDETA", TRIM$(TXTCTA.Caption))
    
    MODOMONE$ = "PESOS"
    If Option15.Value = True Then MODOMONE$ = "BIMONETA"
    If Option19.Value = True Then MODOMONE$ = "MULTIMON"
    Call GRACONTROL("GESACRED", "MODOMONE", MODOMONE$)
    '
    IMPURED$ = TRIM$(Label11)
    Call GRACONTROL("FACOMPRA", "REDONDEO", IMPURED$)
    '
    COMPALET$ = "NO"
    If Check3.Value = 1 Then COMPALET$ = "SI"
    Call GRACONTROL("FACOMPRA", "VERILEFA", COMPALET$)
    '
    EditaCC$ = "SI"
    If Check4.Value = 1 Then EditaCC$ = "NO"
    Call GRACONTROL("FACOMPRA", "EDITARCC", EditaCC$)
    '
    MONESEGUNPROV$ = "SI"
    If Check5.Value = 1 Then MONESEGUNPROV$ = "NO"
    Call GRACONTROL("FACOMPRA", "MONXPROV", MONESEGUNPROV$)
    '
    Call GRACONTROL("COREPER", "COREPER", Text100.TEXT)
    '
    SIRCREB_IIBB$ = TRIM$(SIRCREB.Caption)
    Call GRACONTROL("FACOMPRA", "SIRCREB", SIRCREB_IIBB$)
    '
    MONUME$ = TRIM$(Label4.Caption)
    Call GRACONTROL("FACOMPRA", "VALORCIF", MONUME$)
    MONESEGUNPROV$ = "SI"
    '
    VALIDAPOSIVA$ = "NO"
    If Check6.Value = 1 Then VALIDAPOSIVA$ = "SI"
    Call GRACONTROL("FACOMPRA", "VAPOSIVA", VALIDAPOSIVA$)
    '
    Call GRACONTROL("FACOMPRA", "RETIIBB", lblRetIIBB)

    Hide
    '
End Sub



Private Sub Command20_Click()
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        lblRetIIBB.Caption = VDEVU$
        lblDescripcionRetIIBB.Caption = VALACT2$("ECUECON")
    End If
End Sub


Private Sub Command3_Click(Index As Integer)
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        CUEIVA(Index).Caption = VDEVU$
        DECUENTA(Index).Caption = VALACT2$("ECUECON")
    End If
    
End Sub
'

Private Sub Command32_Click()
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        SIRCREB.Caption = VDEVU$
        DENOSIRCREB.Caption = VALACT2$("ECUECON")
    End If
End Sub

Sub Command4_Click()
    Command4.Enabled = False
    If EXISTE%("CONTAGEN.EXE") > 0 Then
        Call COMUNI("Actualice la Base de Datos de Cuentas Contables\a Través del Plan de Cuentas\del Sistema de Contabilidad General.\  \Pulse 'Aceptar' para Pasar a este Sistema.")
        Call CONECRUN("CONTAGEN", "Contabilidad General - Caja y Bancos")
      Else
        Call CARCUEINGAS
    End If
    Command4.Enabled = True
End Sub

Sub Command5_Click()
   Command5.Enabled = False
   Call CARTAPROVIN
   Command5.Enabled = True
End Sub

Sub Command6_Click()
    '
    Command6.Enabled = False
    Call CARTACENCOS
99  Command6.Enabled = True
    '
End Sub



Private Sub Command7_Click()
    Command7.Enabled = False
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Label4.Caption = VDEVU$
        Label3.Caption = VALACT2$("ECUECON")
    End If
    Command7.Enabled = True
End Sub

Private Sub Command77_Click()
    Command77.Enabled = False
    '
    CANTICUENTAS% = XVALO(CantRegistros("PLANCUEN", , "NOMESS"))
    Dim FILAXX&()
    ReDim Preserve FILAXX&(CANTICUENTAS%, 3)
    If EXISTE%(LUDAT$ & "TACENCOS") < 1 Then GoTo 99
    FIN& = ULTREG&("TACENCOS")
    If FIN& < 1 Then GoTo 99
    
    MUESTRADATOS1.Height = 3840
    MUESTRADATOS1.Command1.Visible = True
10  Call CARGA_ENCABEZADO(MUESTRADATOS1.GRID, "Centro de Costos/f18;Denomincación/A24;Importe/F20", MUESTRADATOS1)
    
    ANOX_ACTUAL% = Year(Date)
    ANOSELECCIONADO% = ANOX_ACTUAL
    ANOSELECCIONADO% = XVALO(InputBox$("Selección de Año de Informe", "Ingrese Año (yyyy)", ANOSELECCIONADO%))
    If ANOSELECCIONADO% < 1980 Or ANOSELECCIONADO% > ANOX_ACTUAL% Then GoTo 99
    ANO2% = Right$(TRIM$(ANOSELECCIONADO%), 2)
    '
    For k& = 1 To FIN&
       XX$ = REGLEIDO$("TACENCOS", k&)
       CODCCOSTO$ = TRIM$(Mid$(XX$, 1, 8))
       DENOMINA$ = TRIM$(Mid$(XX$, 9, 40))
       MUESTRADATOS1.GRID.Rows = MUESTRADATOS1.GRID.Rows + 1
       MUESTRADATOS1.GRID.TextMatrix(k&, 1) = CODCCOSTO$
       MUESTRADATOS1.GRID.TextMatrix(k&, 2) = DENOMINA$
       MUESTRADATOS1.GRID.TextMatrix(k&, 3) = FORMATNUM$(IMPO_ACUMU_CENCOS#(XVALO(CODCCOSTO$), ANOSELECCIONADO%), "F14.2")
    Next k&
    '
    MUESTRADATOS1.Caption = MUESTRADATOS1.Caption & SAFETEXT$(ANOSELECCIONADO%)
    Call COLOREAR_GRILLASQL(MUESTRADATOS1.GRID, &HE0E0E0)
20  MUESTRADATOS1.Show 1
    '
    
    If MUESTRADATOS1.RESP_ZELE$ = "SELECCIONMULTIPLE" Then Call SeleccionMultipleCentroCostos: GoTo 20
    CCOSSTRING$ = MUESTRADATOS1.RESP_ZELE$
    
    CCOSTX% = XVALO(CCOSSTRING$)
    If CCOSTX% < 1 Then GoTo 10
    DENOMICENCOS$ = TRIM$(ECOARCH("TACENCOS", TRIM$(CCOSSTRING$)))
    
    Call CARGA_ENCABEZADO(MUESTRADATOS2.GRID, "Cta.Cble./A13;Denomincación/A24;Enero " & ANO2% & "/F12.2;Febrero " & ANO2% & "/F12.2;Marzo " & ANO2% & "/F12.2;Abril " & ANO2% & "/F12.2;Mayo " & ANO2% & "/F12.2;Junio " & ANO2% & "/F12.2;Julio " & ANO2% & "/F12.2;Agosto " & ANO2% & "/F12.2;Septiembre " & ANO2% & "/F12.2;Octurbre " & ANO2% & "/F12.2;Noviembre " & ANO2% & "/F12.2;Diciembre " & ANO2% & "/F12.2;Total/F14.2;CODIEMPR/F0", MUESTRADATOS2, 1, 1)
    '
    'CARGO ENCABEZADO DE GRILLA DE FILTRO
    MUESTRADATOS2.MSF_2.Rows = 3
    MUESTRADATOS2.MSF_2.FixedRows = 2
    Call CARGA_ENCABEZADO(MUESTRADATOS2.MSF_2, "Cta.Cble./A13;Denomincación/A24;Enero " & ANO2% & "/F12.2;Febrero " & ANO2% & "/F12.2;Marzo " & ANO2% & "/F12.2;Abril " & ANO2% & "/F12.2;Mayo " & ANO2% & "/F12.2;Junio " & ANO2% & "/F12.2;Julio " & ANO2% & "/F12.2;Agosto " & ANO2% & "/F12.2;Septiembre " & ANO2% & "/F12.2;Octurbre " & ANO2% & "/F12.2;Noviembre " & ANO2% & "/F12.2;Diciembre " & ANO2% & "/F12.2;Total/F14.2;CODIEMPR/F0", MUESTRADATOS2, 1, 1)
    '
    Dim Total(13)
    Erase Total
    MUESTRADATOS2.GRID.Rows = 2
    IMPOTOTAL# = 0
    VUELTA% = 0
    SQLX$ = "SELECT P.CODEXCUE ,  P.Denocue, C.CenCosto, SUM(C.IDebeComp) AS FACSUMAIDEBE, SUM(C.IHabeComp) AS NCSUMAHABER,"
    SQLX$ = SQLX$ + " MONTH(C.FECHEMISI) as mes, YEAR(C.FECHEMISI) AS ANO ,C.CodiEmpr,C.cuecontai"
    SQLX$ = SQLX$ + " FROM CAJABANC C  WITH(NOLOCK) INNER JOIN PLANCUEN P WITH(NOLOCK)"
    SQLX$ = SQLX$ + " ON C.cuecontai = P.CodInCue"
    'CONDICION
    SQLX$ = SQLX$ + " Where c.NuProve > 0 And c.CENCOSTO = " & CCOSTX% & " AND YEAR(C.FECHEMISI) = " & ANOSELECCIONADO%
    '
    SQLX$ = SQLX$ + " GROUP BY  C.CenCosto,YEAR(C.FECHEMISI), MONTH(C.FECHEMISI),C.cuecontai, P.CODEXCUE , P.Denocue,C.CodiEmpr"
    SQLX$ = SQLX$ + " ORDER BY YEAR(C.FECHEMISI), MONTH(C.FECHEMISI),C.cuecontai, P.CODEXCUE, P.Denocue, C.CodiEmpr  "
    Set RS = READSET(SQLX$)
    With RS
    i% = 1
    IMPOTOTAL# = 0
    Do While Not .EOF
      MEX = XVALO(!MES)
      SUBRAYAR% = 0
      If VUELTA% = 0 Then FILAXX&(XVALO(!CUECONTAI), XVALO(!CodiEmpr)) = i%
      If ((SAFETEXT$(!CodExCue) <> CODEXCUE_ANT) Or (XVALO(!CodiEmpr) <> CodiEmpr_ANT)) And (VUELTA% > 0) Then
           i% = i% + 1
           MUESTRADATOS2.GRID.Rows = i%
           MUESTRADATOS2.GRID.Rows = MUESTRADATOS2.GRID.Rows + 1
           IMPOTOTAL# = 0
           'USO UN VECTOR PARA LLEVAR EL CONTROL DE LA FILA YA QUE CUANDO CAMBIABA DE MES LA MISMA CUENTA LO IMPRIMIA EN OTRA FILA
           If FILAXX&(XVALO(!CUECONTAI), XVALO(!CodiEmpr)) = 0 Then
             FILAXX&(XVALO(!CUECONTAI), XVALO(!CodiEmpr)) = i%
           End If
           
      End If
      '
      If XVALO(!CodiEmpr) > 0 Then SUBRAYAR% = 1 ' PARA SUBRAYAR CASO MULTIEMPRESA  A MODO DE DIFERENCIAR UNO Y OTRO

      MUESTRADATOS2.GRID.Row = FILAXX&(XVALO(!CUECONTAI), XVALO(!CodiEmpr)): MUESTRADATOS2.GRID.COL = 1
      MUESTRADATOS2.GRID.TextMatrix(FILAXX&(XVALO(!CUECONTAI), XVALO(!CodiEmpr)), 1) = SAFETEXT$(!CodExCue): If SUBRAYAR% > 0 Then MUESTRADATOS2.GRID.CellFontUnderline = True: MUESTRADATOS2.GRID.CellForeColor = vbBlue
      MUESTRADATOS2.GRID.TextMatrix(FILAXX&(XVALO(!CUECONTAI), XVALO(!CodiEmpr)), 2) = SAFETEXT$(!DENOCUE)
        IMPO# = XVALO(!FACSUMAIDEBE) - XVALO(!NCSUMAHABER)
      MUESTRADATOS2.GRID.TextMatrix(FILAXX&(XVALO(!CUECONTAI), XVALO(!CodiEmpr)), MEX + 2) = TRIM$(FORMATNUM$(IMPO#, "F12.2"))
        IMPOTOTAL# = IMPOTOTAL# + IMPO#
      'MUESTRADATOS2.GRID.TextMatrix(FILAXX&(XVALO(!CUECONTAI), XVALO(!CodiEmpr)), 15) = FORMATNUM$(IMPOTOTAL#, "F12.2") ' LO SUMA AL FINAL
      Total(MEX) = Total(MEX) + IMPO#
      'Total(13) = Total(13) + IMPOTOTAL# 'LO SUMA AL FINAL
      MUESTRADATOS2.GRID.TextMatrix(FILAXX&(XVALO(!CUECONTAI), XVALO(!CodiEmpr)), 16) = XVALO(!CodiEmpr)
      '
      'CUANDO CAMBIA LA CUENTA CONTABLE PASA AL SIGUIENTE REGISTRO
      CODEXCUE_ANT = TRIM$(!CodExCue)
      CodiEmpr_ANT = XVALO(!CodiEmpr)
      VUELTA% = 1
      .MoveNext
    Loop
    End With

    'Call COLOREAR_GRILLASQL_VERTICAL(MUESTRADATOS2.GRID, &HE0E0E0, 0, 2)
    MUESTRADATOS2.Caption = "Informe de Gastos - Centro de Costos: " & DENOMICENCOS$ & " (Año: " & ANOSELECCIONADO% & ")"
    Screen.MousePointer = 1
    MUESTRADATOS2.GRID.FixedCols = 3 '
    '
    'ORDENO POR LA DENOMINACION DE LA CUENTA.
    MUESTRADATOS2.GRID.COL = 1
    MUESTRADATOS2.GRID.Sort = 5
    
    'IMPRIMIR TOTALES VERTICALES

    i% = MUESTRADATOS2.GRID.Rows
    'TOTALIZO AL FINAL POR COLUMNA
    MUESTRADATOS2.GRID.Rows = i% + 2
     MUESTRADATOS2.GRID.TextMatrix(i%, 2) = "TOTALES: "
    For J% = 3 To 15 ' 3 A 15 POR QUE LOS MESES ARRANCAN DE LA COLUMNA 3
       MUESTRADATOS2.GRID.TextMatrix(i%, J%) = FORMATNUM$(Total(J% - 2), "F12.2") ' MENOS 2 POR EL QUE EL VECTOR TIENE EL NUMERO REAL DEL MES
    Next J%
    '
    'TOTALIZA DE MANERA HORIZONTAL
    For L% = 1 To MUESTRADATOS2.GRID.Rows - 1
        TOTALFILA# = 0
        For r% = 3 To 15
          TOTALFILA# = TOTALFILA# + XVALO(MUESTRADATOS2.GRID.TextMatrix(L%, r%))
        Next r%
        MUESTRADATOS2.GRID.TextMatrix(L%, 15) = FORMATNUM$(TOTALFILA#, "F12.2")
    Next L%
    
    Call COLOREAR_GRILLASQL(MUESTRADATOS2.GRID, &HE0E0E0)

    'COLOREO LA FILA DE LOS TOTALES
    Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MUESTRADATOS2.GRID, &HC00000, i%)
    Call COLOREAR_GRILLA_SOLO_UNA_AFILA(MUESTRADATOS2.GRID, &H808080, i%)
    '
    'POSICIONARSE AL FINAL DE LA LISTA
     TPXI& = MUESTRADATOS2.GRID.Rows - 22: If TPXI& < 1 Then TPXI& = 1
     On Error Resume Next
     MUESTRADATOS2.GRID.TopRow = TPXI&
     On Error GoTo 0
     '
    MUESTRADATOS2.ANO = SAFETEXT$(ANOSELECCIONADO%)
    MUESTRADATOS2.CENTROCOSTO = CCOSSTRING$
    MUESTRADATOS2.Show 1
    GoTo 20
    '
    'FALTA VER EL TEMA DE LOS TOTALES
99
   Command77.Enabled = True

End Sub
Sub SeleccionMultipleCentroCostos()

  CONDI_PUBLIC$ = ""
  Dim FILAXX&()
  CANTICUENTAS% = XVALO(CantRegistros("PLANCUEN", , "NOMESS"))
  ReDim Preserve FILAXX&(CANTICUENTAS%, 3)  '
  '
  'GENERA LA CONDICION SEGUN LOS CENTRO DE COSTOS SELECCIONADOS PARA LA CONSULTA SQL
  With MUESTRADATOS1
     ATXX% = 0
     CONDI$ = ""
     For KKU& = 1 To .GRID.Rows - 1
       If .GRID.TextMatrix(KKU&, 0) = "*" Then
         CCOSTX% = XVALO(.GRID.TextMatrix(KKU&, 1))
         If ATXX% < 1 Then
              CONDI$ = CONDI$ + " (c.CENCOSTO = " & CCOSTX%
              ATXX% = 1
            Else
              CONDI$ = CONDI$ + " OR c.CENCOSTO = " & CCOSTX%
          End If
       End If
    Next KKU&
   End With
   '
   CONDI$ = CONDI$ + ") "
   If ATXX% < 1 Then Exit Sub

   'VALIDA SI SE SELECCIONARON C.COSTOS
   '+++++++++++++++++++++++++++++++++++++++++++++
   If CONDI$ = "" Then
     Call COMUNI("Falta Seleccionar Centro de Costos")
     GoTo 99
   End If
   CONDI_PUBLIC$ = CONDI$ ' LUEGO LA VOY A TENER QUE USAR EN MUESTRADATOS2
   '
   '
   '********************************************
   'CARGO EL LIST CON LOS MESES
    Call CARSELMESA
    Call SELECMESANO.LICOPEN.Clear
    For n& = 1 To ULTREG&("SELMESA")
       SELECMESANO.LICOPEN.AddItem REGLEIDO$("SELMESA", n&)
    Next n&
    SELECMESANO.Show 1
    '
    '*******************************************
    MESEX$ = "": ATXX% = 0: jH& = 0
    '´RECORRO LOS MESES SELECCIONADOS - ARMO LA CONSULTA PARA SQLX$ Y CARGO EN UNA VARIABLE LOS MESES-ENCABEZADOS
    '****************************************************************************************************************
    For KKU& = SELECMESANO.LICOPEN.ListCount To 1 Step -1
      jH& = jH& + 1
      
      If SELECMESANO.LICOPEN.Selected(KKU& - 1) = True Then
        ' CARGO EL STRING DE LOS CAMPOS PARA LOS ENCABEZADO
        TEXTOMES$ = TRIM$(SELECMESANO.LICOPEN.LIST(KKU& - 1))
        POS1% = InStr(TRIM$(TEXTOMES$), " ")
        MES$ = Left$(TEXTOMES$, POS1% - 1)
        ANO$ = TRIM$(Right(TEXTOMES$, Len(TEXTOMES$) - Len(MES$)))
        MESEX$ = MESEX$ + TEXTOMES$
        MESEX$ = MESEX$ + "/F12.2"
        MESEX$ = MESEX$ + ";"
        '*******************************************************
        'GENERO CONDICION SQL
        MESFILTRO$ = "01/" & MESNUM$(TEXTOMES$) & "/" & ANO$
        ANO1% = Right(FECHATEXLAR$(CVINT(MESFILTRO$)), 4)
        If ATXX% < 1 Then
             CONDI$ = CONDI$ + " AND (MONTH(C.FECHEMISI) = " & MESNUM$(TEXTOMES$)
             CONDI$ = CONDI$ + " AND  YEAR(C.FECHEMISI) = " & ANO1%
             ATXX% = 1
           Else
             CONDI$ = CONDI$ + " OR MONTH(C.FECHEMISI) = " & MESNUM$(TEXTOMES$)
             CONDI$ = CONDI$ + " AND  YEAR(C.FECHEMISI) = " & ANO1%
         End If
      End If
   Next KKU&
   CONDI$ = CONDI$ + ") "
   If ATXX% < 1 Then Exit Sub ' SI NO ELIGE NINGUN MES SE VA
   
    'CARGO ENCABEZADO DE GRILLA DE FILTRO
    Call CARGA_ENCABEZADO(MUESTRADATOS2.GRID, "Cta.Cble./A13;Denomincación/A24;" & MESEX$ & "Total/F14.2;CODIEMPR/F0", MUESTRADATOS2, 1, 1)
    '
    MUESTRADATOS2.MSF_2.Rows = 3
    MUESTRADATOS2.MSF_2.FixedRows = 2
    Call CARGA_ENCABEZADO(MUESTRADATOS2.MSF_2, "Cta.Cble./A13;Denomincación/A24;" & MESEX$ & "Total/F14.2;CODIEMPR/F0", MUESTRADATOS2, 1, 1)
    
    
    Dim Total()
    Erase Total
    ReDim Total(jH&)
    MUESTRADATOS2.GRID.Rows = 2
    VUELTA% = 0
    SQLX$ = "SELECT P.CODEXCUE ,  P.Denocue, C.CenCosto, SUM(C.IDebeComp) AS FACSUMAIDEBE, SUM(C.IHabeComp) AS NCSUMAHABER,"
    SQLX$ = SQLX$ + " MONTH(C.FECHEMISI) as mes, YEAR(C.FECHEMISI) AS ANO ,C.CodiEmpr,C.cuecontai"
    SQLX$ = SQLX$ + " FROM CAJABANC C  WITH(NOLOCK) INNER JOIN PLANCUEN P WITH(NOLOCK)"
    SQLX$ = SQLX$ + " ON C.cuecontai = P.CodInCue"
    'CONDICION
    SQLX$ = SQLX$ + " Where " & CONDI$
    '
    SQLX$ = SQLX$ + " GROUP BY  C.CenCosto,YEAR(C.FECHEMISI), MONTH(C.FECHEMISI),C.cuecontai, P.CODEXCUE , P.Denocue,C.CodiEmpr"
    SQLX$ = SQLX$ + " ORDER BY YEAR(C.FECHEMISI), MONTH(C.FECHEMISI),C.cuecontai, P.CODEXCUE, P.Denocue, C.CodiEmpr  "
    Set RS = READSET(SQLX$)
    With RS
    i% = 1
    Do While Not .EOF
     ' FORMATEO LA FECHA PARA QUE ME QUEDE COMO EL ENCABEZADO JULIO 15
      MEX = XVALO(!MES)  ' 7
      NUMAÑO = XVALO(!ANO) ' 2015
      MESTEX$ = TRIM$(METEXTO$(FECHANUM("01/" & Right$("00" & MEX, 2) & "/" + Right$(SAFETEXT$(NUMAÑO), 2)))) ' JULIO 2015
      MESTEX$ = Mid$(MESTEX$, 1, Len(MESTEX$) - 4) & Right$(SAFETEXT$(NUMAÑO), 2) 'JULIO 15
      '
      SUBRAYAR% = 0
      If VUELTA% = 0 Then FILAXX&(XVALO(!CUECONTAI), XVALO(!CodiEmpr)) = i%
      If ((SAFETEXT$(!CodExCue) <> CODEXCUE_ANT) Or (XVALO(!CodiEmpr) <> CodiEmpr_ANT)) And (VUELTA% > 0) Then
           i% = i% + 1
           MUESTRADATOS2.GRID.Rows = i%
           MUESTRADATOS2.GRID.Rows = MUESTRADATOS2.GRID.Rows + 1
           'USO UN VECTOR PARA LLEVAR EL CONTROL DE LA FILA YA QUE CUANDO CAMBIABA DE MES LA MISMA CUENTA LO IMPRIMIA EN OTRA FILA
           If FILAXX&(XVALO(!CUECONTAI), XVALO(!CodiEmpr)) = 0 Then
             FILAXX&(XVALO(!CUECONTAI), XVALO(!CodiEmpr)) = i%
           End If
           
      End If
      '
      If XVALO(!CodiEmpr) > 0 Then SUBRAYAR% = 1 ' PARA SUBRAYAR CASO MULTIEMPRESA  A MODO DE DIFERENCIAR UNO Y OTRO
      NUCOLUMNA% = COLUMNA%(MESTEX$, MUESTRADATOS2.GRID)
      MUESTRADATOS2.GRID.Row = FILAXX&(XVALO(!CUECONTAI), XVALO(!CodiEmpr)): MUESTRADATOS2.GRID.COL = 1
      MUESTRADATOS2.GRID.TextMatrix(FILAXX&(XVALO(!CUECONTAI), XVALO(!CodiEmpr)), 1) = SAFETEXT$(!CodExCue): If SUBRAYAR% > 0 Then MUESTRADATOS2.GRID.CellFontUnderline = True: MUESTRADATOS2.GRID.CellForeColor = vbBlue
      MUESTRADATOS2.GRID.TextMatrix(FILAXX&(XVALO(!CUECONTAI), XVALO(!CodiEmpr)), 2) = SAFETEXT$(!DENOCUE)
        IMPO# = XVALO(!FACSUMAIDEBE) - XVALO(!NCSUMAHABER)
      MUESTRADATOS2.GRID.TextMatrix(FILAXX&(XVALO(!CUECONTAI), XVALO(!CodiEmpr)), NUCOLUMNA%) = TRIM$(FORMATNUM$(IMPO#, "F12.2"))
      Total(NUCOLUMNA%) = Total(NUCOLUMNA%) + IMPO#
      MUESTRADATOS2.GRID.TextMatrix(FILAXX&(XVALO(!CUECONTAI), XVALO(!CodiEmpr)), MUESTRADATOS2.GRID.Cols - 1) = XVALO(!CodiEmpr)
      '
      'CUANDO CAMBIA LA CUENTA CONTABLE PASA AL SIGUIENTE REGISTRO
      CODEXCUE_ANT = TRIM$(!CodExCue)
      CodiEmpr_ANT = XVALO(!CodiEmpr)
      VUELTA% = 1
      .MoveNext
    Loop
    End With

    'Call COLOREAR_GRILLASQL_VERTICAL(MUESTRADATOS2.GRID, &HE0E0E0, 0, 2)
    MUESTRADATOS2.Caption = "Informe de Gastos - Centro de Costos: Selección Múltiple"
    Screen.MousePointer = 1
    MUESTRADATOS2.GRID.FixedCols = 3 '
    '
    'ORDENO POR LA DENOMINACION DE LA CUENTA.
    MUESTRADATOS2.GRID.COL = 1
    MUESTRADATOS2.GRID.Sort = 5
    
    'IMPRIMIR TOTALES VERTICALES

    i% = MUESTRADATOS2.GRID.Rows
    'TOTALIZO AL FINAL POR COLUMNA
    MUESTRADATOS2.GRID.Rows = i% + 2
    MUESTRADATOS2.GRID.TextMatrix(i%, 2) = "TOTALES: "
'    For j% = 3 To MUESTRADATOS2.GRID.Cols - 1 '
'       If j% > MUESTRADATOS2.GRID.Rows - 1 Then Exit For
'       MUESTRADATOS2.GRID.TextMatrix(I%, j%) = FORMATNUM$(Total(j%), "F12.2") ' MENOS 2 POR EL QUE EL VECTOR TIENE EL NUMERO REAL DEL MES
'    Next j%
    '
    'TOTALIZA DE MANERA VERTICAL
    Dim TOTALCOLU#()
    Erase TOTALCOLU#
    For L% = 1 To MUESTRADATOS2.GRID.Rows - 1
        For r% = 3 To MUESTRADATOS2.GRID.Cols - 1
          If L% >= MUESTRADATOS2.GRID.Rows - 2 Then Exit For
          MUESTRADATOS2.GRID.TextMatrix(MUESTRADATOS2.GRID.Rows - 2, r%) = FORMATNUM$(XVALO(MUESTRADATOS2.GRID.TextMatrix(MUESTRADATOS2.GRID.Rows - 2, r%)) + XVALO(MUESTRADATOS2.GRID.TextMatrix(L%, r%)), "F12.2")
        Next r%
    Next L%

    'TOTALIZA DE MANERA HORIZONTAL
    For L% = 1 To MUESTRADATOS2.GRID.Rows - 1
        TOTALFILA# = 0
        For r% = 3 To MUESTRADATOS2.GRID.Cols - 1
          TOTALFILA# = TOTALFILA# + XVALO(MUESTRADATOS2.GRID.TextMatrix(L%, r%))
        Next r%
        MUESTRADATOS2.GRID.TextMatrix(L%, MUESTRADATOS2.GRID.Cols - 2) = FORMATNUM$(TOTALFILA#, "F12.2")
    Next L%
    
    Call COLOREAR_GRILLASQL(MUESTRADATOS2.GRID, &HE0E0E0)

    'COLOREO LA FILA DE LOS TOTALES
    Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MUESTRADATOS2.GRID, &HC00000, i%)
    Call COLOREAR_GRILLA_SOLO_UNA_AFILA(MUESTRADATOS2.GRID, &H808080, i%)
    '
    'POSICIONARSE AL FINAL DE LA LISTA
     TPXI& = MUESTRADATOS2.GRID.Rows - 22: If TPXI& < 1 Then TPXI& = 1
     On Error Resume Next
     MUESTRADATOS2.GRID.TopRow = TPXI&
     On Error GoTo 0
     '
    MUESTRADATOS2.ANO = SAFETEXT$(ANOSELECCIONADO%)
    MUESTRADATOS2.CENTROCOSTO = "Selección Múltiple"
    MUESTRADATOS2.Show 1
'    GoTo 20
    '
    'FALTA VER EL TEMA DE LOS TOTALES
99
   Command13.Enabled = True

End Sub
Function IMPO_ACUMU_CENCOS#(CCSTO%, ANOX%)

        SQLX$ = " SELECT SUM(IHabeComp) AS HABE, SUM (IDebeComp) DEBE, CENCOSTO FROM CAJABANC"
        SQLX$ = SQLX$ + " Where CENCOSTO = " & CCSTO%
        SQLX$ = SQLX$ + " And Year(FECHEMISI) = " & ANOX%
        SQLX$ = SQLX$ + "  GROUP BY CENCOSTO"
        Dim RS As New ADODB.Recordset
        Set RS = FLEXCONN.Execute(SQLX$)
        IMPO_ACUMU_CENCOS# = XVALO(RS!DEBE) - XVALO(RS!HABE)
        RS.Close
        Set RS = Nothing
        
End Function
Private Sub Command8_Click()
    '
      Command8.Enabled = False
    '
    OPX% = COMALTER%("Seleccione la Tabla:\\Grupos\Familias\Conceptos")
    If OPX% < 1 Or OPX% > 3 Then GoTo 99
    '
    Select Case OPX%
        Case 1 'GRUPOS
10          Call TRALOCAL("TABGRUP", "")
            VTB% = VENTABU("TABGRUP")
            If VTB% >= 0 Then
              For U& = 1 To ULTREG&("!TABGRUP")
                X$ = REGLEIDO$("!TABGRUP", U&)
                CD$ = TRIM((Mid$(X$, 1, 6)))
                Call REPLA(X$, CD$, 1, 6)
                Call GRAREG("TABGRUP", X$, U&)
                '
                ' ValidA QUE EL CODIGO SEA VALIDO
                If TRIM(CD$) = "" Then
                  Call MENSERR(24, "Código de Grupo no Válido.")
                  GoTo 10
                End If
                ' Valida que la familia posea descripcion
                DESC$ = TRIM$(Mid$(X$, 7, 48))
                If Len(DESC$) < 1 Then
                  Call MENSERR(24, "Código de Grupo sin Descripción")
                  GoTo 10
                End If
                'valida que haya ingresado proratera
                PRO$ = TRIM$(Mid$(X$, 55, 1))
                If Len(PRO$) < 1 Then
                  Call MENSERR(24, "No ha Definidio el Campo Prorratea.")
                  GoTo 10
                End If
                '
              Next U&
              Call FILESORT("TABGRUP", "TABGRUP", 1, 1, "ASC")
              Call TRACENTRAL("TABGRUP", "")
            End If
        Case 2 'FAMILIAS
20          Call TRALOCAL("TABFAMI", "")
            VTB% = VENTABU("TABFAMI")
            If VTB% >= 0 Then
              For U& = 1 To ULTREG&("!TABFAMI")
                X$ = REGLEIDO$("!TABFAMI", U&)
                CD$ = TRIM$((Mid$(X$, 1, 6)))
                Call REPLA(X$, CD$, 1, 6)
                Call GRAREG("TABFAMI", X$, U&)
                '
                ' ValidA QUE EL CODIGO SEA VALIDO
                If TRIM(CD$) = "" Then
                  Call MENSERR(24, "Código de Familia no Válido.")
                  GoTo 20
                End If
                ' Valida que la familia posea descripcion
                DESC$ = TRIM$(Mid$(X$, 7, 48))
                If Len(DESC$) < 1 Then
                  Call MENSERR(24, "Código de Familia sin Descripción")
                  GoTo 20
                End If
                'valida que haya ingresado el grupo de pertenencia
                NG$ = TRIM((Mid$(X$, 55, 6)))
                If TRIM(NG$) <= "" Then
                  Call MENSERR(24, "No ha Definidio el Grupo de Pertenencia.")
                  GoTo 20
                End If
                '
                'valida que el codigo no sea repetido
                For U2& = 1 To U& - 1
                  y$ = REGLEIDO$("!TABFAMI", U2&)
                        If TRIM$(Mid$(X$, 1, 6)) = TRIM$(Mid$(y$, 1, 6)) Then
                            Call MENSERR(24, "Código de Item Repetido")
                            GoTo 20
                        End If
                Next U2&
                '
              Next U&
              Call FILESORT("TABFAMI", "TABFAMI", 1, 1, "ASC")
              Call TRACENTRAL("TABFAMI", "")
            End If
        Case 3 'ITEMS Y CONCEPTOS.
30          Call TRALOCAL("TABITCO", "")
            VTB% = VENTABU("TABITCO")
            If VTB% >= 0 Then
                For U& = 1 To ULTREG&("!TABITCO")
                X$ = REGLEIDO$("!TABITCO", U&)
                CD1$ = TRIM$(Mid$(X$, 1, 6))
                Call REPLA(X$, CD1$, 1, 6)
                Call GRAREG("TABITCO", X$, U&)
                '
                ' ValidA QUE EL CODIGO SEA VALIDO
                If Len(CD1$) <= 0 Or Len(CD1$) > 6 Then
                  Call MENSERR(24, "Código de Item no Válido.")
                  GoTo 30
                End If
                'valida que el codigo no sea repetido
                For U2& = 1 To U& - 1
                  y$ = REGLEIDO$("!TABITCO", U2&)
                        If TRIM$(Mid$(X$, 1, 6)) = TRIM$(Mid$(y$, 1, 6)) Then
                            Call MENSERR(24, "Código de Item Repetido")
                            GoTo 30
                        End If
                Next U2&
                ' Valida que EL ITEM posea descripcion
                DESC$ = TRIM$(Mid$(X$, 7, 32))
                If Len(DESC$) < 1 Then
                  Call MENSERR(24, "Código de Item sin Descripción")
                  GoTo 30
                End If
                'valida que haya familia de pertenencia
                NF$ = TRIM(Mid$(X$, 41, 6))
                If TRIM(NF$) = "" Then
                  Call MENSERR(24, "No ha Definidio la Familia de Pertenencia.")
                  GoTo 30
                End If
                'NO VALIDA LA UNIDAD DE MEDIDA XQ
                'ALGUNOS TIENEN Y OTROS NO
              Next U&
              Call FILESORT("TABITCO", "TABITCO", 1, 1, "ASC")
              Call TRACENTRAL("TABITCO", "")
            End If
    End Select
    '
99  Command8.Enabled = True
    '
    '
End Sub

Sub Command9_Click()
   '\\\OT-06-0023-FG-26/01/06
   Command9.Enabled = False
   If DERACCE%("TICAMBIO", "Ajustar Tipo de Cambio Central") < 2 Then
     GoTo 99
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
99 Command9.Enabled = True
   '\\\OT-06-0023-FG-FIN
End Sub

Private Sub CUEIVA_Change(Index As Integer)
   DECUENTA(Index).Caption = DENOCUE$(CUEINT%(CUEIVA(Index).Caption))
End Sub



Private Sub Form_Load()
    '
    For IXX% = 0 To CUEIVA.UBound
      TXXX$ = TRIM$(Str$(IXX%))
      CUEIVA(IXX%).Caption = CONTROL$("", "CUEIVA-" + TXXX$)
      If CUEINT%(CUEIVA(IXX%)) < 1 Then
        CUEIVA(IXX%) = TRIM$(CUEIVA(IXX%)) + "-inex"
      End If
      If IXX% < 4 Then
         TASA = XVALO(CONTROL$("", "TASIVA-" + TXXX$))
         TASIVA(IXX%).TEXT = CSTRING$(MKS$(TASA), 3, 6, 2, 2)
      End If
    Next IXX%
    '
    lblRetIIBB = CONTROL("FACOMPRA", "RETIIBB")
    
    'CUENTA POR DEFECTO PARA LA CARGA POR DETALLE DE FACTURA DE PROVEEDORES
    Me.TXTCTA = CONTROL("FORFACPR", "CTAXDETA")
    
    If CONTROL$("", "VALICAI") = "SI" Then
        Check1.Value = 1
    End If
    '\\\OT-06-0208-WAR-05/06/06///
    If CONTROL$("LIVACOM", "SUPRIVA") = "SI" Then
      Check2.Value = 1
    End If
    '\\\OT-06-0208-WAR-FIN///
    If CONTROL$("GESACRED", "MODOMONE") = "BIMONETA" Then Option15.Value = True
    If CONTROL$("GESACRED", "MODOMONE") = "MULTIMON" Then Option19.Value = True
    '
    Check3.Value = 0
    If CONTROL$("FACOMPRA", "VERILEFA") = "SI" Then Check3.Value = 1

    Check4.Value = 0
    If CONTROL("FACOMPRA", "EDITARCC") = "NO" Then Check4.Value = 1
    '
    If CONTROL("FACOMPRA", "MONXPROV") = "NO" Then Check5.Value = 1
    
    Label11.Caption = CONTROL$("FACOMPRA", "REDONDEO")
    
    'CODIGO DE REGIMEN DE PERCEPCION DE IVA
    Text100.TEXT = CONTROL$("COREPER", "COREPER")
    
    SIRCREB_IIBB$ = CONTROL("FACOMPRA", "SIRCREB")
    SIRCREB.Caption = SIRCREB_IIBB$
    
    MONUME$ = CONTROL("FACOMPRA", "VALORCIF")
    Label4.Caption = MONUME$
    
    Check6.Value = 0
    If CONTROL("FACOMPRA", "VAPOSIVA") = "SI" Then Check6.Value = 1
    
    Call CargaTablasAuxiliares
    
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    '
End Sub

Private Sub Label11_Change()
  Label8.Caption = DENOCUE$(CUEINT%(Label11.Caption))
End Sub

Private Sub Label4_Change()
    Label3.Caption = DENOCUE$(CUEINT%(Label4.Caption))
End Sub

Private Sub lblRetIIBB_Change()
    lblDescripcionRetIIBB.Caption = DENOCUE$(CUEINT%(lblRetIIBB.Caption))
End Sub

Private Sub mnuPercepIvaProveedoresImportacion_Click()
    
    IDLIST& = ObtenerIdTablaAuxiliar("PERCIVAPROIMPO")
    Call ABM_TABLA(IDLIST&)

End Sub

Private Sub SIRCREB_Change()
    DENOSIRCREB.Caption = DENOCUE$(CUEINT%(SIRCREB.Caption))
End Sub

Private Sub TXTCTA_Change()
  Me.TXTDESCRIP.Caption = DENOCUE$(CUEINT%(Me.TXTCTA.Caption))
End Sub

Function COLUMNA%(CAMPO_ENCA$, MSF As MSFlexGrid)

   COLUMNA% = -1
   For i% = 1 To MSF.Cols - 1
     CAMPO_LEIDO$ = TRIM$(MSF.TextMatrix(0, i%))
     If CAMPO_LEIDO$ = CAMPO_ENCA$ Then
        COLUMNA% = i%
        Exit For
     End If
   Next i%
   
End Function
