VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form OPRECIB07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración y Formulario de Recibo Standard de Cobranza"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10590
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   10590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      ForeColor       =   &H80000008&
      Height          =   8520
      Left            =   9480
      ScaleHeight     =   8490
      ScaleWidth      =   1875
      TabIndex        =   79
      Top             =   120
      Width           =   1905
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   180
         ScaleHeight     =   165
         ScaleWidth      =   735
         TabIndex        =   88
         Top             =   6960
         Width           =   765
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   89
            Top             =   0
            Width           =   12000
         End
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
         Left            =   200
         Picture         =   "OPRECIB07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   120
         Width           =   650
      End
      Begin VB.CommandButton Command2 
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
         Left            =   200
         Picture         =   "OPRECIB07.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   810
         Width           =   650
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00624E28&
         Caption         =   "Tablas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   3075
         Left            =   120
         TabIndex        =   81
         Top             =   3600
         Width           =   860
         Begin VB.CommandButton Command21 
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
            Height          =   520
            Left            =   80
            MaskColor       =   &H00800000&
            Picture         =   "OPRECIB07.frx":0454
            Style           =   1  'Graphical
            TabIndex        =   90
            ToolTipText     =   "Percepción Ingresos Brutos - Tabla de Provincias"
            Top             =   2450
            Width           =   670
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
            Height          =   520
            Left            =   80
            Picture         =   "OPRECIB07.frx":059E
            Style           =   1  'Graphical
            TabIndex        =   85
            ToolTipText     =   "Actualizar Tabla de Valores y Medios de Cobranza"
            Top             =   840
            Width           =   670
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
            Height          =   520
            Left            =   80
            Picture         =   "OPRECIB07.frx":06E8
            Style           =   1  'Graphical
            TabIndex        =   84
            ToolTipText     =   "Actualizar Tabla de Talonarios de Recibo y Cobradores"
            Top             =   1350
            Width           =   670
         End
         Begin VB.CommandButton Command11 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   520
            Left            =   80
            Picture         =   "OPRECIB07.frx":0832
            Style           =   1  'Graphical
            TabIndex        =   83
            ToolTipText     =   "Archivo Maestro de Cuentas Contables"
            Top             =   270
            Width           =   670
         End
         Begin VB.CommandButton Command7 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   520
            Left            =   80
            Picture         =   "OPRECIB07.frx":097C
            Style           =   1  'Graphical
            TabIndex        =   82
            ToolTipText     =   "Cotización Dolar y Tipo de Cambio"
            Top             =   1890
            Width           =   670
         End
      End
      Begin VB.CommandButton Command20 
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
         Left            =   200
         Picture         =   "OPRECIB07.frx":0C86
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   1510
         Width           =   650
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   57
      Top             =   360
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00E2D3C0&
      Caption         =   "CONTABILIZAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5580
      Left            =   5160
      TabIndex        =   23
      Top             =   0
      Width           =   4180
      Begin VB.TextBox Text12 
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
         Height          =   330
         Left            =   3360
         TabIndex        =   95
         Text            =   " "
         ToolTipText     =   "Código de Régimen de Retención de I.V.A."
         Top             =   4800
         Width           =   600
      End
      Begin VB.CommandButton Command22 
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
         Left            =   2520
         TabIndex        =   91
         Top             =   4820
         Width           =   175
      End
      Begin VB.CommandButton Command19 
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
         Left            =   3780
         TabIndex        =   53
         Top             =   2590
         Width           =   175
      End
      Begin VB.CommandButton Command18 
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
         Left            =   3780
         TabIndex        =   47
         Top             =   1850
         Width           =   175
      End
      Begin VB.CommandButton Command15 
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
         Left            =   3780
         TabIndex        =   36
         Top             =   1080
         Width           =   175
      End
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
         Left            =   3780
         TabIndex        =   33
         Top             =   4100
         Width           =   175
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
         Height          =   300
         Left            =   3780
         TabIndex        =   29
         Top             =   3340
         Width           =   175
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
         Left            =   3780
         TabIndex        =   25
         Top             =   310
         Width           =   175
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Cod.Reg."
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
         Left            =   2700
         TabIndex        =   96
         ToolTipText     =   "Código de Régimen de Retención de I.V.A."
         Top             =   4875
         Width           =   660
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Retención IVA"
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
         Left            =   -240
         TabIndex        =   94
         Top             =   4920
         Width           =   1500
      End
      Begin VB.Label Label24 
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
         Height          =   300
         Left            =   1305
         TabIndex        =   93
         Top             =   4815
         Width           =   1230
      End
      Begin VB.Label Label23 
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
         Height          =   300
         Left            =   255
         TabIndex        =   92
         Top             =   5175
         Width           =   3690
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ajuste Redondeo"
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
         Left            =   -120
         TabIndex        =   56
         Top             =   2700
         Width           =   2325
      End
      Begin VB.Label Label20 
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
         Height          =   300
         Left            =   2280
         TabIndex        =   55
         Top             =   1850
         Width           =   1485
      End
      Begin VB.Label Label19 
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
         Height          =   300
         Left            =   240
         TabIndex        =   54
         Top             =   2950
         Width           =   3690
      End
      Begin VB.Label Label18 
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
         Height          =   300
         Left            =   270
         TabIndex        =   50
         Top             =   2200
         Width           =   3690
      End
      Begin VB.Label Label17 
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
         Height          =   300
         Left            =   2280
         TabIndex        =   49
         Top             =   2590
         Width           =   1485
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "N/Cred.Descuentos"
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
         Left            =   -105
         TabIndex        =   48
         Top             =   1950
         Width           =   2325
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Acreedores Comerciales"
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
         Left            =   -105
         TabIndex        =   39
         Top             =   1180
         Width           =   2325
      End
      Begin VB.Label Label14 
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
         Height          =   300
         Left            =   2280
         TabIndex        =   38
         Top             =   1080
         Width           =   1485
      End
      Begin VB.Label Label13 
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
         Height          =   300
         Left            =   270
         TabIndex        =   37
         Top             =   1440
         Width           =   3690
      End
      Begin VB.Label Label12 
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
         Height          =   300
         Left            =   270
         TabIndex        =   35
         Top             =   4450
         Width           =   3690
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
         Height          =   300
         Left            =   2280
         TabIndex        =   34
         Top             =   4080
         Width           =   1485
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cheques Rechazados"
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
         Left            =   -105
         TabIndex        =   32
         Top             =   4200
         Width           =   2325
      End
      Begin VB.Label Label9 
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
         Height          =   300
         Left            =   240
         TabIndex        =   31
         Top             =   3700
         Width           =   3690
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
         Height          =   300
         Left            =   2280
         TabIndex        =   30
         Top             =   3340
         Width           =   1485
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cartera de Valores"
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
         Left            =   -105
         TabIndex        =   28
         Top             =   3450
         Width           =   2325
      End
      Begin VB.Label Label6 
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
         Height          =   300
         Left            =   270
         TabIndex        =   27
         Top             =   670
         Width           =   3690
      End
      Begin VB.Label Label5 
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
         Height          =   300
         Left            =   2310
         TabIndex        =   26
         Top             =   310
         Width           =   1485
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Deudores por Ventas"
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
         Left            =   -105
         TabIndex        =   24
         Top             =   420
         Width           =   2325
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00E2D3C0&
      Caption         =   "OTROS FORMULARIOS DE IMPRESIÓN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   5160
      TabIndex        =   22
      Top             =   5640
      Width           =   4200
      Begin VB.CommandButton Command17 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   3240
         Picture         =   "OPRECIB07.frx":0F90
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   945
         Width           =   750
      End
      Begin VB.CommandButton Command16 
         Caption         =   "."
         Height          =   280
         Left            =   3000
         TabIndex        =   65
         Top             =   945
         Width           =   175
      End
      Begin VB.TextBox Text10 
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
         Left            =   1680
         TabIndex        =   64
         Text            =   " "
         Top             =   945
         Width           =   1380
      End
      Begin VB.TextBox Text7 
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
         Left            =   240
         TabIndex        =   63
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1260
         Width           =   2940
      End
      Begin VB.CommandButton Command9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   3240
         Picture         =   "OPRECIB07.frx":129A
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   240
         Width           =   750
      End
      Begin VB.TextBox Text9 
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
         Left            =   240
         TabIndex        =   61
         TabStop         =   0   'False
         Text            =   " "
         Top             =   585
         Width           =   2940
      End
      Begin VB.CommandButton Command8 
         Caption         =   "."
         Height          =   280
         Left            =   3000
         TabIndex        =   60
         Top             =   240
         Width           =   175
      End
      Begin VB.TextBox Text8 
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
         Left            =   1680
         TabIndex        =   59
         Text            =   " "
         Top             =   240
         Width           =   1380
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Comp. Cuentas"
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
         Left            =   -1360
         TabIndex        =   68
         Top             =   1020
         Width           =   2955
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ajuste de Saldo"
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
         Left            =   -720
         TabIndex        =   67
         Top             =   320
         Width           =   2325
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E2D3C0&
      Caption         =   "RECIBO DE COBRANZA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7300
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4920
      Begin VB.Frame Frame8 
         BackColor       =   &H00E2D3C0&
         Caption         =   "VALIDACIONES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   210
         TabIndex        =   97
         Top             =   5160
         Width           =   4485
         Begin VB.CheckBox chkNoValidaFechaRecibo 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "No Valida Fecha Recibo"
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
            Height          =   300
            Left            =   120
            TabIndex        =   98
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E2D3C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   210
         TabIndex        =   76
         Top             =   6495
         Width           =   4490
         Begin VB.TextBox Text11 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3000
            TabIndex        =   78
            Text            =   " "
            Top             =   240
            Width           =   1185
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Generar Nota Débito por Diferencia de Cambio a Partir de $"
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
            Left            =   120
            TabIndex        =   77
            Top             =   200
            Width           =   3285
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E2D3C0&
         Caption         =   "PAPEL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         Left            =   210
         TabIndex        =   15
         Top             =   3045
         Width           =   2970
         Begin VB.OptionButton Option5 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "EndDoc"
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
            Height          =   225
            Left            =   200
            TabIndex        =   75
            Top             =   800
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton Option6 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "NewPage"
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
            Left            =   1700
            TabIndex        =   74
            Top             =   800
            Width           =   1035
         End
         Begin VB.OptionButton Option4 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Hoja Suelta"
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
            Height          =   465
            Left            =   1700
            TabIndex        =   73
            Top             =   240
            Width           =   1065
         End
         Begin VB.OptionButton Option3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Formulario Continuo"
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
            Height          =   465
            Left            =   200
            TabIndex        =   72
            Top             =   240
            Width           =   1275
         End
      End
      Begin VB.Frame Frame15 
         BackColor       =   &H00E2D3C0&
         Caption         =   "CÁLCULO DE DÍAS DE VENCIMIENTO PONDERADO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   210
         TabIndex        =   69
         Top             =   5880
         Width           =   4485
         Begin VB.OptionButton Option12 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Por Vencimiento Factura"
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
            Height          =   225
            Left            =   2160
            TabIndex        =   71
            Top             =   300
            Width           =   2235
         End
         Begin VB.OptionButton Option11 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Por Fecha de Recibo"
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
            Height          =   225
            Left            =   120
            TabIndex        =   70
            Top             =   300
            Value           =   -1  'True
            Width           =   1995
         End
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H00E2D3C0&
         Caption         =   "DESCUENTO"
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
         Left            =   1600
         TabIndex        =   44
         Top             =   1800
         Width           =   2175
         Begin VB.OptionButton Option10 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Cancela"
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
            Height          =   300
            Left            =   1050
            TabIndex        =   52
            Top             =   520
            Width           =   885
         End
         Begin VB.OptionButton Option9 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Valor"
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
            Height          =   300
            Left            =   250
            TabIndex        =   51
            Top             =   520
            Width           =   1275
         End
         Begin VB.CheckBox Check3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Nota de Crédito"
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
            Height          =   300
            Left            =   100
            TabIndex        =   46
            Top             =   240
            Width           =   1635
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Suprime Impresión"
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
            Height          =   300
            Left            =   100
            TabIndex        =   45
            Top             =   840
            Width           =   1665
         End
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00E2D3C0&
         Caption         =   "RENGLONES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Left            =   3440
         TabIndex        =   42
         Top             =   3645
         Width           =   1275
         Begin VB.TextBox Text6 
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
            Height          =   330
            Left            =   420
            MaxLength       =   2
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   200
            Width           =   450
         End
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1560
         TabIndex        =   41
         Text            =   " "
         Top             =   1320
         Width           =   3165
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   " Suprime Impresión al Generar el Documento"
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
         TabIndex        =   21
         Top             =   1080
         Width           =   4600
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00E2D3C0&
         Caption         =   "ORDENAMIENTO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   210
         TabIndex        =   18
         Top             =   4275
         Width           =   1530
         Begin VB.OptionButton Option8 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Vencimiento"
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
            Left            =   120
            TabIndex        =   20
            Top             =   480
            Width           =   1275
         End
         Begin VB.OptionButton Option7 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Emision"
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
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Value           =   -1  'True
            Width           =   1065
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2D3C0&
         Caption         =   "PREFIJO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   3440
         TabIndex        =   16
         Top             =   3045
         Width           =   1275
         Begin VB.TextBox Text4 
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
            Height          =   330
            Left            =   315
            TabIndex        =   17
            TabStop         =   0   'False
            Text            =   "0001"
            Top             =   200
            Width           =   660
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2D3C0&
         Caption         =   "TALONARIO EN USO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   1860
         TabIndex        =   11
         Top             =   4275
         Width           =   2850
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
            Height          =   300
            Left            =   2580
            TabIndex        =   13
            Top             =   360
            Width           =   175
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
            Height          =   300
            Left            =   1425
            TabIndex        =   12
            Text            =   " "
            Top             =   360
            Width           =   1185
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Nueva Numeración"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   165
            TabIndex        =   14
            Top             =   255
            Width           =   1170
         End
      End
      Begin VB.TextBox Text1 
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
         Left            =   2415
         TabIndex        =   10
         Text            =   " "
         Top             =   315
         Width           =   1380
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
         Height          =   300
         Left            =   3780
         TabIndex        =   9
         Top             =   315
         Width           =   175
      End
      Begin VB.TextBox Text2 
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
         Left            =   210
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " "
         Top             =   690
         Width           =   3750
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Abrir"
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
         Left            =   4095
         Picture         =   "OPRECIB07.frx":15A4
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   315
         Width           =   650
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E2D3C0&
         Caption         =   "NUMERACIÓN"
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
         Left            =   210
         TabIndex        =   3
         Top             =   1800
         Width           =   1335
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Pre Impresa"
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
            Left            =   240
            TabIndex        =   5
            Top             =   250
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Auto Numera"
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
            Left            =   240
            TabIndex        =   4
            Top             =   700
            Width           =   1065
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2D3C0&
         Caption         =   "CÓDIGO"
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
         Left            =   3840
         TabIndex        =   1
         Top             =   1800
         Width           =   855
         Begin VB.TextBox Text5 
            Alignment       =   2  'Center
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
            Height          =   360
            Left            =   190
            TabIndex        =   2
            TabStop         =   0   'False
            Text            =   "RB"
            Top             =   525
            Width           =   450
         End
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Puerto Impresion"
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
         Left            =   -240
         TabIndex        =   40
         Top             =   1500
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario"
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
         Left            =   0
         TabIndex        =   6
         Top             =   420
         Width           =   2325
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "OPRECIB07.frx":18AE
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "OPRECIB07.frx":1AE2
      TabIndex        =   58
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "OPRECIB07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command12_Click()
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Label5 = VDEVU$
        Label6 = VALACT2$("ECUECON")
    End If
End Sub

Private Sub Command13_Click()
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Label8 = VDEVU$
        Label9 = VALACT2$("ECUECON")
    End If
End Sub
'
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Hide
End Sub

 Sub Command10_Click()
    Command10.Enabled = False
    Call CARTACOBRA
    Command10.Enabled = True
End Sub

Sub Command11_Click()
   Command11.Enabled = False
   RETEJ% = EJER%
   EJER% = 1
   Call CARCUEINGAS
   EJER% = RETEJ%
   EJU$ = EJACT$
   Command11.Enabled = True
End Sub

Private Sub Command14_Click()
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Label11 = VDEVU$
        Label12 = VALACT2$("ECUECON")
    End If
End Sub

Private Sub Command15_Click()
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Label14 = VDEVU$
        Label13 = VALACT2$("ECUECON")
    End If
End Sub
'\\\OT-06-0144-WAR-17/04/06///
Private Sub Command16_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text10.TEXT = VDEVU$
        Text7.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command17_Click()
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text10.TEXT), "Formularios de Impresion")
    If EXISTE%(ACONEC$ + ".EXE") < 1 Then
        Call FINAL("")
    End If
End Sub

Private Sub Command18_Click()
    '\\\OT-07-0029-RG-26/01/07///
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Label17 = VDEVU$
        Label18 = VALACT2$("ECUECON")
    End If
    '\\\OT-07-0029-RG-FIN///
End Sub

Private Sub Command19_Click()
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Label20 = VDEVU$
        Label19 = VALACT2$("ECUECON")
    End If

End Sub

'\\\OT-06-0144-WAR-FIN///
Private Sub Command2_Click()
    '
    FENUE$ = TRIM$(Text3.TEXT)
    If FENUE$ = "" Then
        FENUE$ = "01/01/80"
    End If
    FENU% = FECHANUM(FENUE$)
    If FENU% < 1 Then
        Call MENSERR(24, "Fecha no Válida")
        Text3.SetFocus
        Exit Sub
    End If
    'CONTROL
    If EXISTE_TABLA%("ULTRECIBO") > 0 Then
        FEGRABADA% = FECHANUM(Control$("", "FENURECI"))
        'SI SE CAMBIA LA FECHA BORRA LA INFO DE LA TABLA DE CONTROL DE NUMERO DE RECIBO
        If FENU% <> FEGRABADA% Then ' SI CAMBIO LA FECHA INICIALIZA LA TABLA BORRANDOLE LOS DATOS
          Call BORRAREGISTROS("ULTRECIBO", "NUMETALO >= 0 ", REGAF&, "NOMESS")
        End If
    End If
    '
    Call GRACONTROL("", "FORECIST", Text1.TEXT)
    Call GRACONTROL("", "FORAJUSA", Text8.TEXT)
    '\\\OT-06-0144-WAR-17/04/06///
    Call GRACONTROL("", "FORCOMPE", Text10.TEXT)
    '\\\OT-06-0144-WAR-FIN///
    Call GRACONTROL("", "CODMRECI", "RB")    ' Text5.TEXT)
    '
    MONUME$ = "AUTONUME": If Option1.Value = True Then MONUME$ = "PREIMPRE"
    Call GRACONTROL("", "NURECIST", MONUME$)
    '
    TIPAPEL$ = "FORMUCON": If Option4.Value = True Then TIPAPEL$ = "HOJASUEL"
    Call GRACONTROL("", "TIPARECI", TIPAPEL$)
    '
    AVAHOJA$ = "ENDDOC": If Option6.Value = True Then AVAHOJA$ = "NEWPAGE"
    Call GRACONTROL("RECISTA", "AVAHOJA", AVAHOJA$)
    '
    MODORDE$ = "EMISION": MODORECI% = 1
    If Option8.Value = True Then
        MODORDE$ = "VENCIM"
        MODORECI% = 2
    End If
    Call GRACONTROL("RECISTA", "MODORDE", MODORDE$)
    '\\\OT-07-0029-RG-25/01/07///
    DESCUEX$ = ""
    If Check3.Value = 1 Then DESCUEX$ = "CREDESCU"
    Call GRACONTROL("RECISTA", "DESCAJA", DESCUEX$)
    '
    SUPRINT$ = "NO": CANCELPRINT% = 0
    If Check1.Value = 1 Then SUPRINT$ = "SI": CANCELPRINT% = 1
    Call GRACONTROL("CREDESCU", "SUPRINT", SUPRINT$)
    '\\\OT-07-0029-RG-FIN///
    SUPRINT$ = "NO": CANCELPRINT% = 0
    If Check2.Value = 1 Then SUPRINT$ = "SI": CANCELPRINT% = 1
    Call GRACONTROL("RECISTA", "SUPRINT", SUPRINT$)
    '
    Call GRACONTROL("", "FENURECI", FENUE$)
    '
    PorValor$ = "SI": If Option10.Value = True Then PorValor$ = "NO"
    Call GRACONTROL("", "OPCVALOR", PorValor$)

    PREFIREM$ = TRIM$(Text4.TEXT): If PREFIREM$ = "" Then PREFIREM$ = "0001"
    If PREFIREM$ <> "" Then
       If Len(PREFIREM$) <> 4 Then
          Call COMUNI("Prefijo no Válido")
          Text4.SetFocus
          Exit Sub
       End If
       For U& = 1 To 4
          UI% = Asc(Mid$(PREFIREM$, U&))
          If UI% < 48 Or UI% > 57 Then
             Call COMUNI("Prefijo no Válido")
             Text4.SetFocus
             Exit Sub
          End If
       Next U&
    End If
    Call GRACONTROL("", "PREFRECI", PREFIREM$)
    '
    Call GRACONTROL("AMACLI", "ECUECON", Label5)
    Call GRACONTROL("AMAPRO", "ECUECON", Label14)
    Call GRACONTROL("RECISTA", "CUEREDON", Label20)

    Call GRACONTROL("", "CUECHECA", Label8)
    Call GRACONTROL("", "CUECHERE", Label11)
    '\\\OT-07-0029-RG-26/01/07///
    Call GRACONTROL("", "CUEDESCU", Label17)
    Call GRACONTROL$("OPRECIBO", "CURETIVA", Label24) '''
    Call GRACONTROL$("OPRECIBO", "CREREIVA", Text12) ' CODIGO DE REGIMEN DE RETENCION DE IVA
    
    '\\\OT-07-0029-RG-FIN///
    '\\\OT-05-0513-WAR-14/07/05///
    CARENDET% = Val(Text6.TEXT)
    Call GRACONTROL("RECISTA", "CARENDET", Text6)
    '\\\OT-05-0513-WAR-FIN///
    Call GRACONTROL(IDENTER$, "PORTREC", Combo1.TEXT)
    Call GRACONTROL("*", "PORTREC", Combo1.TEXT)
    '
    valor$ = "FECHARE": If Option12.Value Then valor$ = "FECHAFA"
    Call GRACONTROL("RECISTA", "DIASVEN", valor$)
    '
    valor$ = SAFETEXT$(XVALO(Text11))
    Call GRACONTROL("RECISTA", "MINDIFCM", valor$)
    '
    valor$ = "NO": If chkNoValidaFechaRecibo.Value = 1 Then valor$ = "SI"
    Call GRACONTROL("RECISTA", "NOVAFECH", valor$)
    
    Call CIERRARCH("*.*")
    '
    On Error Resume Next
    RECIBO00.List1.Clear
    RECIBO00.List2.Clear
    RECIBO00.Label4.Caption = ""
    RECIBO00.Label8.Caption = ""
    RECIBO00.Label10.Caption = ""
    RECIBO00.Label14.Caption = ""
    RECIBO00.Label18.Caption = ""
    RECIBO00.Label20.Caption = ""
    RECIBO00.BOTNEXT.Enabled = True
    RECIBO00.Command5.Enabled = False
    On Error GoTo 0
    '
    Hide
    '
End Sub

Private Sub Command20_Click()
    '
    Command20.Enabled = False
    '
10  OPX% = ALTERNA%("Depuracion Fecha Emision de Recibos\Otros Accesos Directos")
    If OPX% < 1 Or OPX% > 1 Then GoTo 99
    '
    If OPX% = 1 Then 'Carga de Interes por Mora
        Screen.MousePointer = 11
        Call ABRECONEXION
        SQLX$ = "SELECT * FROM SADEUPEN"
        Dim RST As ADODB.Recordset
        Set RST = New ADODB.Recordset
25      On Error Resume Next
        RST.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
          On Error GoTo 0
          Call CapturaErrorOpen
          GoTo 25
        End If
        On Error GoTo 0
        '
        FIN& = CantRegistros("SADEUPEN", "FECHEMISI > '" & FETEXCOR1$(FECHANUM("01/01/80")) & "'")
        JJ& = 0
        With RST
            Do While Not .EOF
                Call TRACE(24, JJ&, FIN&)
                FF1% = CVINT%(!FECHEMISI)
                DOCU$ = SAFETEXT$(!codicom_lar)
                NCLIE% = XVALO(!NUCLIEN)
                Condi$ = "CODICOM_LAR='" & DOCU$ & "' AND NUCLIEN=" & NCLIE%
                FF2% = CVINT%(VALOBADA("CAJABANC", "FECHCONTA", Condi$))
                '
                If FF1% <> FF2% Then
                    !FECHEMISI = CVDAT(FF2%)
                    .Update
                End If
                JJ& = JJ& + 1
                .MoveNext
            Loop
        End With
        On Error Resume Next
        RST.Close
        Set RST = Nothing
        Screen.MousePointer = 1
        Call COMUNI("Proceso de Recuperacion Finalizada")
        On Error GoTo 0
        GoTo 10
    End If
    '
99  Command20.Enabled = True
    '
End Sub

 Sub Command21_Click()
   Command21.Enabled = False
   Call CARTAPROVIN
   Command21.Enabled = True
End Sub

Private Sub Command22_Click()
    Command22.Enabled = False
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Label24 = VDEVU$
        Label23 = VALACT2$("ECUECON")
    End If
    Command22.Enabled = True
End Sub

Private Sub Command3_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub
'

Private Sub Command4_Click()
    Call SELECHO("SELFECHA")
    FENUE$ = TRIM$(VALACT1$("SELFECHA"))
    If FENUE$ <> "" Then
        Text3.TEXT = FENUE$
    End If
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
    If EXISTE%(ACONEC$ + ".EXE") < 1 Then
       Call FINAL("")
    End If
    Command5.Enabled = True
End Sub

Sub Command6_Click()
    Command6.Enabled = False
    Call CARTAMECO
    Command6.Enabled = True
End Sub

Sub Command7_Click()
   Command7.Enabled = False
   If DERACCE%("TICAMBIO", "Ajustar Tipo de Cambio Central") < 2 Then
     GoTo 99
   End If
   '
   Call TRALOCAL("TAMONED", "")
   VTB% = VENTABU%("TAMONED/NC/TITUPAN=Monedas y Tipo de Cambio")
   If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!TAMONED")
         x$ = REGLEIDO$("!TAMONED", U&)
         Call REPLA(x$, Chr$(U&), 1, 1)
         Call GRAREG("TAMONED", x$, U&)
      Next U&
      FIN& = ULTREG&("!TAMONED")
      Call SETULTREG("TAMONED", FIN&)
   End If
   Call CIERRARCH("TAMONED")
99 Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text8.TEXT = VDEVU$
        Text9.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command9_Click()
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text8.TEXT), "Formularios de Impresion")
    If EXISTE%(ACONEC$ + ".EXE") < 1 Then
        Call FINAL("")
    End If
End Sub

Private Sub Form_Load()
    '
    Text1.TEXT = Control$("", "FORECIST")
    Text2.TEXT = ECOARCH$("INDIFRM", Text1.TEXT)
    Text3.TEXT = Control$("", "FENURECI")
    Text8.TEXT = Control$("", "FORAJUSA")
    Text9.TEXT = ECOARCH$("INDIFRM", Text8.TEXT)
    '\\\OT-06-0144-WAR-17/04/06///
    Text10.TEXT = Control$("", "FORCOMPE")
    Text7.TEXT = ECOARCH$("INDIFRM", Text10.TEXT)
    '\\\OT-06-0144-WAR-FIN///
    '
    '\\\OT-05-0513-WAR-14/07/05///
    Text6 = UCase$(Control$("RECISTA", "CARENDET"))
    '\\\OT-05-0513-WAR-FIN
    Label5 = Control$("AMACLI", "ECUECON")
    Label6 = ECOARCH$("ECUECON", Label5)
    Label14 = Control$("AMAPRO", "ECUECON")
    Label13 = ECOARCH$("ECUECON", Label14)
    Label20 = Control$("RECISTA", "CUEREDON")
    Label19 = ECOARCH$("ECUECON", Label20)
    Label8 = Control$("", "CUECHECA")
    Label9 = ECOARCH$("ECUECON", Label8)
    Label11 = Control$("", "CUECHERE")
    Label12 = ECOARCH$("ECUECON", Label11)
    '\\\OT-07-0029-RG-26/01/07///
    Label17 = Control$("", "CUEDESCU")
    Label18 = ECOARCH$("ECUECON", Label17)
    
    Label24 = Control$("OPRECIBO", "CURETIVA")
    Label23 = ECOARCH$("ECUECON", Label24)
    Text12 = Control$("OPRECIBO", "CREREIVA")
    
    '\\\OT-07-0029-RG-FIN///
    '
    If Control$("", "NURECIST") = "AUTONUME" Then Option2.Value = True
    If Control$("", "TIPARECI") = "HOJASUEL" Then Option4.Value = True
    If Control$("RECISTA", "AVAHOJA") = "NEWPAGE" Then Option6.Value = True
    '\\\OT-07-0029-RG-26/01/07///
    If Control$("RECISTA", "DESCAJA") = "CREDESCU" Then Check3.Value = 1
    If Control$("CREDESCU", "SUPRINT") = "SI" Then Check1.Value = 1
    '\\\OT-07-0029-RG-FIN///
    If Control$("RECISTA", "SUPRINT") = "SI" Then Check2.Value = 1
    MODORECI% = 1: If Control$("RECISTA", "MODORDE") = "VENCIM" Then Option8.Value = True: MODORECI% = 2
    '
    Option9.Value = True
    If Control$("", "OPCVALOR") = "NO" Then Option10.Value = True
    '
    PREFIREM$ = TRIM$(Control$("", "PREFRECI"))
    Text4.TEXT = PREFIREM$
    Text5.TEXT = "RB"
    '
    Combo1.Clear
    Dim PRX As Printer
    For Each PRX In Printers
      Combo1.AddItem PRX.DeviceName
    Next
    Combo1.TEXT = TRIM$(Control$(IDENTER$, "PORTREC"))
      If Combo1.TEXT = "" Then Combo1.TEXT = Control$("*", "PORTREC")
      If Combo1.TEXT = "" Then Combo1.TEXT = Printer.DeviceName
      '
    If Control("RECISTA", "DIASVEN") = "FECHAFA" Then Option12.Value = True
    Text11 = ""
    ValorMinimoNDebDifCambio# = XVALO(Control("RECISTA", "MINDIFCM"))
    If ValorMinimoNDebDifCambio# > 0 Then Text11 = FORMATNUM$(ValorMinimoNDebDifCambio#, "f12.2")
        
    valor$ = Control("RECISTA", "NOVAFECH")
    If valor$ = "SI" Then chkNoValidaFechaRecibo.Value = 1
    
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me, Picture1)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 0
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    On Error Resume Next ' SE AGREGO EL ON ERROR POR QUE ESTABA DANDO UN ERROR 402 CUANDO LLEGABA AL HIDE
    Hide
    On Error GoTo 0
End Sub

Private Sub Label11_DblClick()
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Label11 = VDEVU$
        Label12 = VALACT2$("ECUECON")
    End If
End Sub

Private Sub Label14_DblClick()
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Label14 = VDEVU$
        Label13 = VALACT2$("ECUECON")
    End If
End Sub



Private Sub Label17_DblClick()
    '\\\OT-07-0029-RG-26/01/07///
    Call Command18_Click
    '\\\OT-07-0029-RG-FIN///
End Sub

Private Sub Label5_DblClick()
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Label5 = VDEVU$
        Label6 = VALACT2$("ECUECON")
    End If
End Sub

Private Sub Label8_DblClick()
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Label8 = VDEVU$
        Label9 = VALACT2$("ECUECON")
    End If
End Sub

Private Sub Text1_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub
'\\\OT-06-0144-WAR-17/04/06///
Private Sub Text10_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text10.TEXT = VDEVU$
        Text7.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text11_LostFocus()
   Text11 = FORMATNUM$(XVALO(Text11), "f12.2")
End Sub

'\\\OT-06-0144-WAR-FIN///
Private Sub Text2_CLICK()
    Text1.SetFocus
End Sub


Private Sub Text3_DBLCLICK()
    Call SELECHO("SELFECHA")
    FENUE$ = TRIM$(VALACT1$("SELFECHA"))
    If FENUE$ <> "" Then
        Text3.TEXT = FENUE$
    End If
End Sub

Private Sub Text5_Click()
    Command2.SetFocus
End Sub
'\\\OT-06-0144-WAR-17/04/06///
Private Sub Text7_CLICK()
  Text10.SetFocus
End Sub
'\\\OT-06-0144-WAR-FIN///
Private Sub Text8_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text8.TEXT = VDEVU$
        Text9.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text9_Click()
    Text8.SetFocus
End Sub

Sub CARTAMECO()     ' carga tabla medios de cobranza
    '
    If DERACCE%("ACTAMECO", "Actualizacion Tabla de Medios de Cobranza") < 2 Then
      Exit Sub
    End If
    '
    CUECHECA$ = TRIM$(Control$("", "CUECHECA"))
    If CUECHECA$ = "" Or CUECOINT%(CUECHECA$) < 1 Then
      Call MENSERR(24, "Falta Definir Cuenta Contable\Asociada a Cartera de Valores.")
      Exit Sub
    End If
    '
    Call SETULTREG("!MECOBRA", 0)
    '
    LU$ = LUDAT$
    If EXISTE%(LU$ + "MECOBRA.DAT") = 1 Then
       If COPYFILE%(LU$ + "MECOBRA.DAT", "C:\FLEXOFT\MECOBRA.WKF") <> 1 Then
          Command1.Enabled = True
          Exit Sub
       End If
    End If
    '
    Screen.MousePointer = 11
    For U& = 1 To ULTREG&("!MECOBRA")
       x$ = REGLEIDO$("!MECOBRA", U&)
       If CVI(Left$(x$, 2)) < 900 Then
         CUECO$ = TRIM$(Mid$(x$, 45, 12))
         If CUECO$ = "" Or CUECOINT%(CUECO$) < 1 Then
           Call REPLA(x$, CUECHECA$, 45, 12)
           Call GRAREG("!MECOBRA", x$, U&)
         End If
       End If
    Next U&
    Call CIERRARCH("!MECOBRA")
    Screen.MousePointer = 1
    '
360 ATY% = VENTABU%("COMECOB/NC")
    If ATY% < 0 Then GoTo 390
    '
361 ATX% = 0: J& = 0
    Screen.MousePointer = 11
    For U& = 1 To ULTREG&("!MECOBRA")
       x$ = REGLEIDO$("!MECOBRA", U&)
       CDX% = CVI(Left$(x$, 2))
362    If CDX% > 0 Then
          If CDX% <= 999 Then
             J& = J& + 1
363          Call GRAREG("!MECOBRA", x$, J&)
             If TRIM$(Mid$(x$, 3, 12)) = "" Then ATX% = 1
             If TRIM$(Mid$(x$, 45, 12)) = "" Then ATX% = 1
364          If TRIM$(Mid$(x$, 57, 4)) = "" Then ATX% = 1
          End If
       End If
365 Next U&
    '
    Call SETULTREG("!MECOBRA", J&)
    '
366 Screen.MousePointer = 1
    If ATX% > 0 Then
        Call COMUNI("Datos Incompletos !!!")
        GoTo 360
    End If
    '
    Call TRACENTRAL("MECOBRA", "")
    Call CIERRARCH("MECOBRA")
    Call FILESORT("MECOBRA", "", 1, 1, "ASC")
    '
390 End Sub

Sub CARTACOBRA()    ' carga tabla cobradores / talonarios
    '
    Dim MCOBRA%(1000)
    '
    Call CIERRARCH("COBRADOR")
    Call CIERRARCH("!COBRADOR")
    LU$ = LUDAT$
    If EXISTE%(LU$ + "COBRADOR.DAT") = 1 Then
       If COPYFILE%(LU$ + "COBRADOR.DAT", "C:\FLEXOFT\COBRADOR.WKF") <> 1 Then
          Command6.Enabled = True
          Exit Sub
       End If
    End If
    '
360 ATY% = VENTABU%("COBRADOR/NC")
    If ATY% < 0 Then GoTo 390
    '
361 ATX% = 0: J& = 0
    Screen.MousePointer = 11
    For U& = 1 To ULTREG&("!COBRADOR")
       x$ = REGLEIDO$("!COBRADOR", U&)
       CDX% = Asc(Left$(x$, 2))
362    If CDX% >= 0 Then
          If CDX% <= 255 Then
             If MCOBRA%(CDX%) = 0 Then
                J& = J& + 1
363             Call GRAREG("!COBRADOR", x$, J&)
                MCOBRA%(CDX%) = 1
             End If
          End If
       End If
365 Next U&
    '
    Call SETULTREG("!COBRADOR", J&)
    Call CIERRARCH("!COBRADOR")
    Call CIERRARCH("COBRADOR")
    '
366 Screen.MousePointer = 1
    '
    AA% = COPYFILE%("C:\FLEXOFT\COBRADOR.WKF", LU$ + "COBRADOR.DAT")
    '
390 End Sub

