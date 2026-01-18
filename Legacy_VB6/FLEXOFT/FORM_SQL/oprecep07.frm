VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form OPRECEP07 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opciones de Configuración"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   11205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame16 
      BackColor       =   &H00CDDADA&
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
      Left            =   120
      TabIndex        =   78
      Top             =   5250
      Width           =   4950
      Begin VB.OptionButton Option10 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Bulto"
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
         Left            =   3300
         TabIndex        =   88
         ToolTipText     =   "Imprime una Etiqueta Por Bulto"
         Top             =   640
         Width           =   900
      End
      Begin VB.OptionButton Option9 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Unidad"
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
         Left            =   3300
         TabIndex        =   87
         ToolTipText     =   "Imprime una Etiqueta por cada Unidad"
         Top             =   300
         Value           =   -1  'True
         Width           =   900
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   4215
         Picture         =   "oprecep07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   240
         Width           =   645
      End
      Begin VB.TextBox Text14 
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
         Left            =   90
         TabIndex        =   81
         TabStop         =   0   'False
         Text            =   " "
         Top             =   600
         Width           =   3150
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
         Left            =   3060
         TabIndex        =   80
         Top             =   240
         Width           =   180
      End
      Begin VB.TextBox TEXT13 
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
         Left            =   1680
         TabIndex        =   79
         Text            =   " "
         Top             =   240
         Width           =   1380
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Etiqueta de Recepción"
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
         Left            =   -200
         TabIndex        =   83
         Top             =   345
         Width           =   1850
      End
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00CDDADA&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   620
      Left            =   100
      TabIndex        =   74
      Top             =   4650
      Width           =   4950
      Begin VB.CommandButton Command17 
         Caption         =   "Imprime Etiqueta de Recepción ??"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   460
         Left            =   30
         TabIndex        =   93
         Top             =   120
         Width           =   1550
      End
      Begin VB.TextBox Text15 
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
         Left            =   4320
         TabIndex        =   91
         ToolTipText     =   "Ingresar Cantidad de Columnas que posee el rollo donde se imprime la etiqueta."
         Top             =   240
         Width           =   555
      End
      Begin VB.CheckBox Check19 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Solo Items Trazables "
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
         Left            =   1680
         TabIndex        =   90
         Top             =   160
         Width           =   1095
      End
      Begin VB.CheckBox Check16 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Imprime Etiqueta de Recepción"
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
         Left            =   75
         TabIndex        =   77
         Top             =   160
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Columnas de Etiquetas en Rollo:"
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
         Index           =   1
         Left            =   2760
         TabIndex        =   92
         Top             =   120
         Width           =   1425
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   1470
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   67
      Top             =   120
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00CDDADA&
      Caption         =   "OPCIONES DE APROBACIÓN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   5250
      TabIndex        =   50
      Top             =   0
      Width           =   4950
      Begin VB.CheckBox Check15 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Utiliza Ventana Valorizada"
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
         Left            =   2520
         TabIndex        =   89
         ToolTipText     =   "Utiliza Ventana Valorizada al Recepcionar"
         Top             =   180
         Width           =   2220
      End
      Begin VB.CommandButton Command13 
         Caption         =   "..."
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
         Height          =   255
         Left            =   4320
         TabIndex        =   73
         Top             =   615
         Width           =   495
      End
      Begin VB.CheckBox Check7 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Certificado de Calidad Proveedor"
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
         Height          =   350
         Left            =   120
         TabIndex        =   51
         Top             =   600
         Width           =   2145
      End
      Begin VB.CheckBox Check14 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Control de Consignaciones"
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
         Height          =   350
         Left            =   2520
         TabIndex        =   72
         Top             =   600
         Width           =   2025
      End
      Begin VB.CheckBox Check10 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "N° de Control de Vigilancia"
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
         TabIndex        =   52
         Top             =   180
         Width           =   2295
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00CDDADA&
      Caption         =   "COMPROBANTE DE APROBACIÓN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2835
      Left            =   5260
      TabIndex        =   38
      Top             =   3390
      Width           =   4950
      Begin VB.Frame Frame17 
         BackColor       =   &H00CDDADA&
         Caption         =   "Control de Recepción Cronológica Segun O.C."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   240
         TabIndex        =   84
         Top             =   2160
         Width           =   4480
         Begin VB.CheckBox Check18 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Con Bloqueo"
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
            Left            =   2520
            TabIndex        =   86
            ToolTipText     =   "Bloquea Si No se Cumple el Orden Propuesto, Excepto Derecho Explícito."
            Top             =   240
            Width           =   1770
         End
         Begin VB.CheckBox Check17 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Recepción Cronológica "
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
            TabIndex        =   85
            ToolTipText     =   "Presenta Cantidades Ideales Según Orden de Fecha de Entrega de O.C."
            Top             =   240
            Width           =   2130
         End
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H00CDDADA&
         Caption         =   "MAIL PAÑOL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   1995
         TabIndex        =   65
         Top             =   1440
         Width           =   2715
         Begin VB.TextBox Text12 
            Appearance      =   0  'Flat
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
            Left            =   105
            TabIndex        =   66
            Top             =   210
            Width           =   2535
         End
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00CDDADA&
         Caption         =   "FORMATO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   240
         TabIndex        =   55
         Top             =   1440
         Width           =   1665
         Begin VB.OptionButton Option8 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "FRM"
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
            Left            =   105
            TabIndex        =   57
            Top             =   250
            Value           =   -1  'True
            Width           =   750
         End
         Begin VB.OptionButton Option7 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "TXT"
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
            Height          =   250
            Left            =   900
            TabIndex        =   56
            Top             =   250
            Width           =   690
         End
      End
      Begin VB.ComboBox Combo1 
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
         Height          =   345
         Left            =   945
         TabIndex        =   53
         Text            =   "Combo1"
         Top             =   975
         Width           =   3795
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
         Height          =   300
         Left            =   2415
         TabIndex        =   42
         Text            =   " "
         Top             =   240
         Width           =   1380
      End
      Begin VB.CommandButton Command9 
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
         TabIndex        =   41
         Top             =   240
         Width           =   180
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
         Left            =   210
         TabIndex        =   40
         TabStop         =   0   'False
         Text            =   " "
         Top             =   600
         Width           =   3750
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   4095
         Picture         =   "oprecep07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   240
         Width           =   645
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Puerto"
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
         Left            =   -1470
         TabIndex        =   54
         Top             =   1125
         Width           =   2325
      End
      Begin VB.Label Label4 
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
         TabIndex        =   43
         Top             =   345
         Width           =   2325
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00CDDADA&
      Caption         =   "REMITO DE DEVOLUCIÓN (RECHAZOS)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   5250
      TabIndex        =   17
      Top             =   1080
      Width           =   4950
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
         Left            =   1680
         TabIndex        =   97
         Top             =   1860
         Width           =   180
      End
      Begin VB.TextBox Text16 
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
         Left            =   1350
         TabIndex        =   96
         Text            =   " "
         ToolTipText     =   "Depósito de Ingreso para Equipos Importados"
         Top             =   1860
         Width           =   330
      End
      Begin VB.TextBox Text11 
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
         Left            =   1350
         TabIndex        =   61
         Text            =   " "
         Top             =   1500
         Width           =   330
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
         Left            =   1680
         TabIndex        =   58
         Top             =   1500
         Width           =   180
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00CDDADA&
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
         Height          =   910
         Left            =   3840
         TabIndex        =   25
         Top             =   900
         Width           =   930
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   230
            TabIndex        =   26
            Text            =   "RD"
            Top             =   330
            Width           =   450
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00CDDADA&
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
         Height          =   495
         Left            =   210
         TabIndex        =   22
         Top             =   900
         Width           =   3540
         Begin VB.OptionButton Option6 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Automática"
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
            Left            =   1995
            TabIndex        =   24
            Top             =   220
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option5 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Manual"
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
            Left            =   480
            TabIndex        =   23
            Top             =   220
            Width           =   1380
         End
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   4095
         Picture         =   "oprecep07.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   300
         Width           =   645
      End
      Begin VB.TextBox Text4 
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
         Left            =   210
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   " "
         Top             =   615
         Width           =   3750
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
         Left            =   3780
         TabIndex        =   19
         Top             =   300
         Width           =   175
      End
      Begin VB.TextBox Text3 
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
         Left            =   2415
         TabIndex        =   18
         Text            =   " "
         Top             =   300
         Width           =   1380
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ingreso Eq.Imp."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   220
         TabIndex        =   99
         ToolTipText     =   "Depósito de Ingreso para Equipos Importados"
         Top             =   1920
         Width           =   1125
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
         Left            =   1875
         TabIndex        =   98
         ToolTipText     =   "Depósito de Ingreso para Equipos Importados"
         Top             =   1860
         Width           =   1905
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
         Left            =   1875
         TabIndex        =   60
         Top             =   1500
         Width           =   1905
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Destino Devolución"
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
         Left            =   180
         TabIndex        =   59
         Top             =   1440
         Width           =   1125
      End
      Begin VB.Label Label1 
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
         TabIndex        =   27
         Top             =   270
         Width           =   2325
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Caption         =   "FORMATO LISTADOR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   2880
      TabIndex        =   14
      Top             =   3830
      Width           =   2190
      Begin VB.OptionButton Option4 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Ficha p/Item"
         Enabled         =   0   'False
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
         Left            =   315
         TabIndex        =   16
         Top             =   510
         Width           =   1800
      End
      Begin VB.OptionButton Option3 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Form. Tabulado"
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
         Left            =   315
         TabIndex        =   15
         Top             =   250
         Value           =   -1  'True
         Width           =   1800
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "OPCIONES DE RECEPCIÓN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   120
      TabIndex        =   11
      Top             =   3830
      Width           =   2655
      Begin VB.CheckBox Check21 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Mts. lineales"
         Enabled         =   0   'False
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
         Left            =   1440
         TabIndex        =   100
         Top             =   390
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Bobinas"
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
         Index           =   1
         Left            =   75
         TabIndex        =   13
         Top             =   510
         Width           =   945
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Mat. General"
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
         Index           =   0
         Left            =   75
         TabIndex        =   12
         Top             =   225
         Value           =   1  'Checked
         Width           =   1905
      End
      Begin VB.Label Label12 
         BackColor       =   &H00CDDADA&
         Caption         =   "----- >"
         Height          =   255
         Left            =   960
         TabIndex        =   103
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00CDDADA&
      Caption         =   "BOLETA DE RECEPCIÓN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3800
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4950
      Begin VB.Frame Frame10 
         BackColor       =   &H00CDDADA&
         Caption         =   "OTRAS OPCIONES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1800
         Left            =   210
         TabIndex        =   31
         Top             =   1920
         Width           =   4530
         Begin VB.CheckBox Check22 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Controles por Peso"
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
            Left            =   2730
            TabIndex        =   101
            Top             =   1200
            WhatsThisHelpID =   1000
            Width           =   1725
         End
         Begin VB.CommandButton Command18 
            Caption         =   "Informa Despacho A.Import."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   195
            TabIndex        =   95
            ToolTipText     =   "Configuración Supervisada Solicite Información a su Soporte de Sistemas  "
            Top             =   1440
            Width           =   2385
         End
         Begin VB.CheckBox Check20 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Informa Despacho Art.Imp."
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
            Height          =   230
            Left            =   210
            TabIndex        =   94
            Top             =   1440
            WhatsThisHelpID =   1000
            Width           =   2325
         End
         Begin VB.CheckBox Check12 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Cierre de O-C"
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
            Left            =   2730
            TabIndex        =   70
            Top             =   1440
            WhatsThisHelpID =   1000
            Width           =   1725
         End
         Begin VB.CheckBox Check13 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Desc. un renglón"
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
            Left            =   2730
            TabIndex        =   71
            Top             =   960
            WhatsThisHelpID =   1000
            Width           =   1725
         End
         Begin VB.TextBox Text8 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   230
            Left            =   1680
            TabIndex        =   45
            Text            =   "0.0"
            Top             =   1200
            Width           =   660
         End
         Begin VB.CheckBox Check11 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Recibe Solo O-C."
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
            Height          =   230
            Left            =   210
            TabIndex        =   69
            ToolTipText     =   "Recibe Solo O-Compras"
            Top             =   960
            Width           =   2445
         End
         Begin VB.CheckBox Check9 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Imp.Desc."
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
            Left            =   2730
            TabIndex        =   64
            Top             =   240
            Width           =   1650
         End
         Begin VB.CheckBox Check8 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Imp.Especificac."
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
            Left            =   2730
            TabIndex        =   63
            Top             =   720
            Width           =   1725
         End
         Begin VB.CheckBox Check4 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Imp.D. Comp."
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
            Left            =   2730
            TabIndex        =   62
            ToolTipText     =   "Imp.Descripción"
            Top             =   480
            Width           =   1645
         End
         Begin VB.CheckBox Check6 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Suprime Val. Rto."
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
            Height          =   230
            Left            =   210
            TabIndex        =   49
            ToolTipText     =   "Suprime la Validación del Remito"
            Top             =   480
            Width           =   2625
         End
         Begin VB.CheckBox Check5 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Sobre-Emb. %"
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
            Height          =   230
            Left            =   210
            TabIndex        =   44
            ToolTipText     =   "Bloquea Sobre-Embarque por mas del % configurado"
            Top             =   1200
            Width           =   1605
         End
         Begin VB.CheckBox Check3 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Anula Cant. Present."
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
            Height          =   230
            Left            =   210
            TabIndex        =   33
            ToolTipText     =   "Anula Cantidades Presentadas"
            Top             =   720
            Width           =   1845
         End
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "No Imp. Comprobante"
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
            Height          =   230
            Left            =   210
            TabIndex        =   32
            Top             =   240
            Width           =   2610
         End
         Begin VB.Label Label11 
            BackColor       =   &H00CDDADA&
            Caption         =   "<----"
            Height          =   375
            Left            =   2400
            TabIndex        =   102
            Top             =   1200
            Width           =   255
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00CDDADA&
         Caption         =   "ANTICIPACIÓN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   880
         Left            =   1635
         TabIndex        =   28
         Top             =   1050
         Width           =   2160
         Begin VB.TextBox Text7 
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
            Left            =   1440
            TabIndex        =   29
            Top             =   420
            Width           =   555
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Dias Fecha Solictud Entrega"
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
            Index           =   0
            Left            =   50
            TabIndex        =   30
            Top             =   315
            Width           =   1305
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
         Height          =   285
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
         Height          =   285
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
         Height          =   285
         Left            =   210
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " "
         Top             =   630
         Width           =   3750
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   4095
         Picture         =   "oprecep07.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   315
         Width           =   645
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00CDDADA&
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
         Height          =   880
         Left            =   210
         TabIndex        =   3
         Top             =   1050
         Width           =   1365
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Manual"
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
            Left            =   90
            TabIndex        =   5
            Top             =   300
            Width           =   1125
         End
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Automatica"
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
            Left            =   90
            TabIndex        =   4
            Top             =   525
            Value           =   -1  'True
            Width           =   1125
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00CDDADA&
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
         Height          =   880
         Left            =   3885
         TabIndex        =   1
         Top             =   1050
         Width           =   855
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   190
            TabIndex        =   2
            Text            =   "BR"
            Top             =   420
            Width           =   450
         End
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
         Height          =   210
         Left            =   0
         TabIndex        =   6
         Top             =   400
         Width           =   2325
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7140
      Left            =   10320
      ScaleHeight     =   7110
      ScaleWidth      =   1455
      TabIndex        =   34
      Top             =   -150
      Width           =   1485
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   75
         Top             =   4920
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   76
            Top             =   0
            Width           =   12000
         End
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
         Height          =   390
         Left            =   105
         Picture         =   "oprecep07.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Actualizar Tabla de Motivos de Rechazo."
         Top             =   1810
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
         Height          =   640
         Left            =   105
         Picture         =   "oprecep07.frx":0D72
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Salir"
         Top             =   300
         Width           =   650
      End
      Begin VB.CommandButton Command2 
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
         Left            =   120
         Picture         =   "oprecep07.frx":0EBC
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Guardar"
         Top             =   3480
         Width           =   650
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
         Height          =   630
         Left            =   105
         Picture         =   "oprecep07.frx":11C6
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Help"
         Top             =   945
         Width           =   650
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
         Height          =   390
         Left            =   105
         Picture         =   "oprecep07.frx":14D0
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Actualizar Tabla de Destinos de Rechazo."
         Top             =   2200
         Width           =   650
      End
      Begin VB.CommandButton Command14 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   105
         Picture         =   "oprecep07.frx":161A
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Autorización de Recepciones sin O/C"
         Top             =   2600
         Width           =   650
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "oprecep07.frx":1764
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "oprecep07.frx":1998
      TabIndex        =   68
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "Configuración"
      Begin VB.Menu mnuBalanza 
         Caption         =   "Balanza"
         Begin VB.Menu mnuLecturaPesoBalanza 
            Caption         =   "Configuraciones"
         End
         Begin VB.Menu mnuTablaTaras 
            Caption         =   "Tabla de Taras"
         End
         Begin VB.Menu mnuTaraPorTipoDeRecepcion 
            Caption         =   "Tara por Tipo de Recepción"
         End
      End
      Begin VB.Menu linea1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCarTabAux 
         Caption         =   "Tablas Auxiliares"
      End
   End
End
Attribute VB_Name = "OPRECEP07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check14_Click()
    
    If Check14.Value = 1 Then
        Command13.Enabled = True
    Else
        Command13.Enabled = False
    End If
    
End Sub


Private Sub Check4_Click()
    If Check4.Value = 1 Then
        Text10 = "STICKER"
        text9 = ""
      ElseIf Check4.Value = 0 Then
        If Text10 = "STICKER" Then
          Text10 = ""
          text9 = ""
        End If
    End If
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Unload Me
End Sub
'\\\OT-05-0723-WAR-22/11/05///
Private Sub Command10_Click()
  '
  'TABLA DE DESTINO DE RECHAZO
  If ULTREG&("TADESTI") < 1 Then
    Call BLANARCH("TADESTI")
    Call REGAPP("TADESTI", Chr$(0) + "Retrabajo")
    Call REGAPP("TADESTI", Chr$(1) + "Devolución")
    Call REGAPP("TADESTI", Chr$(2) + "Baja")
    Call CIERRARCH("TADESTI")
  End If
  '
  If DERACCE%("EDIMOTIV", "Edición de Tabla de Motivos de Rechazo") >= 2 Then
    Call TRALOCAL("TADESTI", "TADESTI")
    '
    ATRI$ = "BORDEIZQ=W" + TRIM$(Str$(Left - 1400))
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + 850))
    VTB% = VENTABU%("TADESTI/" + ATRI$)
    If VTB% >= 0 Then
       JJ& = 0
      For U& = 1 To ULTREG&("!TADESTI")
         X$ = REGLEIDO$("!TADESTI", U&)
         JJ& = JJ& + 1
         Call GRAREG("TADESTI", X$, JJ&)
      Next U&
      Call SETULTREG("TADESTI", JJ&)
      Call CIERRARCH("TADESTI")
    End If
  End If
  '
End Sub
'\\\OT-05-0723-WAR-FIN///
'\\\OT-05-0723-WAR-23/11/05///
Private Sub Command11_Click()
  '
  If DERACCE%("EDIMOTIV", "Edición de Tabla de Motivos de Rechazo") >= 2 Then
    Call TRALOCAL("TAMOTIV", "TAMOTIV")
    '
    ATRI$ = "BORDEIZQ=W" + TRIM$(Str$(Left - 1400))
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + 850))
    VTB% = VENTABU%("TAMOTIV/" + ATRI$)
    If VTB% >= 0 Then
       JJ& = 0
      For U& = 1 To ULTREG&("!TAMOTIV")
         X$ = REGLEIDO$("!TAMOTIV", U&)
         JJ& = JJ& + 1
         Call GRAREG("TAMOTIV", X$, JJ&)
      Next U&
      Call SETULTREG("TAMOTIV", JJ&)
      Call CIERRARCH("TAMOTIV")
    End If
  End If
  '
End Sub

Private Sub Command12_Click()
      Command12.Enabled = False
      If ULTREG&("TADESTI") < 1 Then
        Call BLANARCH("TADESTI")
        Call REGAPP("TADESTI", Chr$(0) + "Retrabajo")
        Call REGAPP("TADESTI", Chr$(1) + "Devolución")
        Call REGAPP("TADESTI", Chr$(2) + "Baja")
        Call CIERRARCH("TADESTI")
      End If
      Call SELECHO("TADESTI")
      DESTINO$ = VALACT2("TADESTI")
      If TRIM$(DESTINO$) = "" Then
          GoTo 99
      End If
    
      Text11 = XVALO(VALACT1$("TADESTI"))
    
99    Command12.Enabled = True
End Sub

Private Sub Command13_Click()
    
    Call TRALOCAL("VEBRECON", "VEBRECON")
    VTB% = VENTABU%("VEBRECON")
    Call TRACENTRAL("VEBRECON", "VEBRECON")
    
End Sub

'\\\OT-05-0723-WAR-FIN///
Private Sub Command14_Click()
  '
  '\\\OT-05-0694-WAR-25/10/05///
  If DERACCE%("STAMACO", "Autorización Recepciones sin Orden de Compra") < 2 Then
    Exit Sub
  End If
  Command14.Enabled = False
  '
  Call TRALOCAL("TAMATCON", "TAMATCON")
  ATRI$ = "BORDEIZQ=W" + TRIM$(Str$(Left - 2600))
  ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + 850))
  VTB% = VENTABU%("TAMATCON/" + ATRI$)
  If VTB% >= 0 Then
     JJ& = 0
     For U& = 1 To ULTREG&("!TAMATCON")
        X$ = REGLEIDO$("!TAMATCON", U&)
        'LE PONGO EL USUARIO QUE AUTORIZO
        US$ = Mid$(X$, 5, 24)
        If TRIM(US$) = "" Then
          Call REPLA(X$, USERNAME$, 5, 24)
        End If
        F1% = CVI(Mid$(X$, 3, 2))
        If F1% = 0 Then
          HOI% = HOY(HO$)
          Call REPLA(X$, MKI$(HOI%), 3, 2)
        End If
        JJ& = JJ& + 1
        Call GRAREG("TAMATCON", X$, JJ&)
     Next U&
     Call SETULTREG("TAMATCON", JJ&)
     Call CIERRARCH("TAMATCON")
  End If
  '
  Call FILESORT("TAMATCON", "TAMATCON", 1, 1, "ASC")
  '
  RTA% = COMALTER%("Seleccione la Siguiente Tarea:\\Emitir Listado\Terminar")
  If RTA% = 1 Then
    Call FINAL("*LIMATCON")
  End If
  Command14.Enabled = True
  '\\\OT-05-0694-WAR-FIN///
  '
End Sub

Private Sub Command15_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        TEXT13.TEXT = VDEVU$
        Text14.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command16_Click()
 Call CONECRUN("FLFORMS/" + TRIM$(TEXT13.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command17_Click()
   Command17.Enabled = False
   If CLACONTRA% > 0 Then
      Command17.Visible = False
   End If
   Command17.Enabled = True
   
End Sub

Private Sub Command18_Click()
   If CLACONTRA% = 1 Then
      Command18.Visible = False
   End If
   Command1.SetFocus
End Sub

Private Sub Command19_Click()
    Call SELECHO("TADEPOSI")
    If VALACT1$("TADEPOSI") <> "" Then
      Text16 = TRIM$(VALACT1$("TADEPOSI"))
    End If

End Sub

Private Sub Command2_Click()
    '
    Command2.Enabled = False
    Screen.MousePointer = 11
    '
    If Check18.Value > 0 And Check17.Value < 1 Then
       Call COMUNI("Atención !! Imposible Marcar Check de Bloqueo Para Recepción en Orden Cronológico\Sin Marcar Check de -Recepción Cronológica-")
       GoTo 99
    End If
    FORIPRE$ = TRIM$(UCase$(Text1.TEXT))
    Call GRACONTROL("", "FORMBREC", Text1.TEXT)
    Call GRACONTROL("", "CODMBREC", Text5.TEXT)
    '
    MONUME$ = "AUTO": If Option1.Value = True Then MONUME$ = "MANUAL"
    Call GRACONTROL("", "NUMEBREC", MONUME$)
    '
    Call GRACONTROL("", "REMIRECH", Text3.TEXT)
    Call GRACONTROL("BOLRECEP", "ETIQRECE", TEXT13.TEXT)
    Call GRACONTROL("", "CODRRECH", Text6.TEXT)
    '
    Call GRACONTROL("", "STIKAPRO", Text10.TEXT)
    '
    MONUME$ = "AUTO": If Option5.Value = True Then MONUME$ = "MANUAL"
    Call GRACONTROL("", "NUMERECH", MONUME$)
    '
    MONUME$ = "NO": If Check1(1).Value = 1 Then MONUME$ = "SI"
    Call GRACONTROL("", "RECPAPEL", MONUME$)
    '
    MONUME$ = "NO": If Check21.Value = 1 Then MONUME$ = "SI"
    Call GRACONTROL("", "REPAMTLI", MONUME$)
    '
    MONUME$ = "FORTABU": If Option4.Value = True Then MONUME$ = "FICHA"
    Call GRACONTROL("", "COMPBREC", MONUME$)
    '
    SUPRINT$ = "NO": CANCELPRINT% = 0
    If Check2.Value = 1 Then SUPRINT$ = "SI": CANCELPRINT% = 1
    Call GRACONTROL("BOLRECEP", "SUPRINT", SUPRINT$)
    '
    ANUCAN$ = "NO"
    If Check3.Value = 1 Then ANUCAN$ = "SI"
    Call GRACONTROL("BOLRECEP", "ANUCANTI", ANUCAN$)
    '
    '\\\OT-06-0515-RG-28/11/06///
    MONUME$ = "NO"
    If Check6.Value = 1 Then MONUME$ = "SI"
    Call GRACONTROL("BOLRECEP", "SUPREMI", MONUME$)
    '\\\OT-06-0515-RG-FIN///
    '
    BLOSOBRE$ = "NO"
    If Check5.Value = 1 Then BLOSOBRE$ = "SI"
    Call GRACONTROL("BOLRECEP", "BLOSOBRE", BLOSOBRE$)
    Call GRACONTROL("BOLRECEP", "PORSOBRE", Text8)
    '
    'CONTROL SI SE RECIBE SOLO POR ORDEN DE COMPRA
    MONUME$ = "NO"
    If Check11.Value = 1 Then MONUME$ = "SI"
    Call GRACONTROL("BOLRECEP", "RESOLOOC", MONUME$)
    '
    DIANREC$ = TRIM$(Text7)
    Call GRACONTROL("BOLRECEP", "DIANREC", DIANREC$)
    '
    NUCONVIG$ = "NO": If Check10.Value = 1 Then NUCONVIG$ = "SI"
    Call GRACONTROL("BOLRECEP", "NUCONVIG", NUCONVIG$)
    '
    CERCAPRO$ = "NO": If Check7.Value = 1 Then CERCAPRO$ = "SI"
    Call GRACONTROL("BOLRECEP", "CERCAPRO", CERCAPRO$)
    '
    Call GRACONTROL(IDENTER$, "PORTSTIK", Combo1.TEXT)
    Call GRACONTROL("*", "PORTSTIK", Combo1.TEXT)
    '
    MONUME$ = "FRM": If Option7.Value = True Then MONUME$ = "TXT"
    Call GRACONTROL("BOLRECEP", "FORMATO", MONUME$)
    '
    Call GRACONTROL("BOLRECEP", "DETIRECH", Text11) ' DESTINO DE DEVOLUCION
    Call GRACONTROL("BOLRECEP", "DEPEQIMP", Text16) ' DESTINO DE DEVOLUCION
    '
    X$ = REGBLAN$("TATIREC")
    Call REPLA(X$, Chr$(1), 1, 1)
    Call REPLA(X$, "Material General", 2, 31)
    Call REPLA(X$, "BOLREGEN", 49, 16)
    Call REPLA(X$, FORIPRE$, 65, 16)
    Call GRAREG("TATIREC", X$, 1)
    '
    If Check1(1).Value = 1 Then
          X$ = REGBLAN$("TATIREC")
          Call REPLA(X$, Chr$(2), 1, 1)
          LEYENDA$ = "Papel en Bobinas"
          If InStr(1, EMPREAC$, "CONTEPLAS") > 0 Then LEYENDA$ = "Bobinas de polipropileno"
          Call REPLA(X$, LEYENDA$, 2, 31)
          Call REPLA(X$, "RECEP-PAPEL", 33, 16)
          Call REPLA(X$, "BOLREGEN", 49, 16)
          Call REPLA(X$, FORIPRE$, 65, 16)
          
          Call GRAREG("TATIREC", X$, 2)
    
          If InStr(1, EMPREAC$, "CONTEPLAS") > 0 Then
                Call REPLA(X$, Chr$(3), 1, 1)
                LEYENDA$ = "Mat.General con Balanza"
                Call REPLA(X$, LEYENDA$, 2, 31)
                Call REPLA(X$, "", 33, 16)
                Call REPLA(X$, "BOLREGEN", 49, 16)
                Call REPLA(X$, FORIPRE$, 65, 16)
                
                Call GRAREG("TATIREC", X$, 3)
          End If
       Else
          X$ = String$(80, 0)
          Call GRAREG("TATIREC", X$, 2)
    End If
    Call FRESHECHO("TATIREC")
    '
    'CONTROLES DE IMPRESION DE DESCRIPCION
    IMPRDES$ = "": If Check9.Value = 1 Then IMPRDES$ = "SI"
    Call GRACONTROL("BOLRECEP", "DESCPRIM", IMPRDES$)
    '
    IMPRDES$ = "": If Check4.Value = 1 Then IMPRDES$ = "SI"
    Call GRACONTROL("BOLRECEP", "DESCOMPL", IMPRDES$)
    '
    IMPRDES$ = "": If Check8.Value = 1 Then IMPRDES$ = "SI"
    Call GRACONTROL("BOLRECEP", "ESPECIFI", IMPRDES$)
    '
    IMPRDES$ = "": If Check12.Value = 1 Then IMPRDES$ = "SI"
    Call GRACONTROL("BOLRECEP", "CIERREOC", IMPRDES$)
    
    IMPRDES$ = "": If Check22.Value = 1 Then IMPRDES$ = "SI"
    Call GRACONTROL("BOLRECEP", "CTRLPESO", IMPRDES$)
    
    MAILPA$ = TRIM$(Text12)
    Call GRACONTROL("BOLRECEP", "MAIPAÑOL", MAILPA$)
    '
    IMPRDES$ = "": If Check13.Value = 1 Then IMPRDES$ = "SI"
    Call GRACONTROL("BOLRECEP", "DESCUNRE", IMPRDES$)
    '
    MONUME$ = "NO": If Check14.Value = 1 Then MONUME$ = "SI"
    Call GRACONTROL("BOLRECEP", "VEBRECON", MONUME$)
    '
    MONUME$ = "NO" ' PRESENTA VENTANA AHP VALORIZADA
    If Check15.Value = 1 Then MONUME$ = "SI"
    Call GRACONTROL("BOLRECEP", "CONPRECI", MONUME$)
    
    MONUME$ = "NO" ' PRESENTA VENTANA AHP VALORIZADA
    If Check16.Value = 1 Then MONUME$ = "SI"
    Call GRACONTROL("BOLRECEP", "IMETIREC", MONUME$)
    '
    If Check19.Value = 1 Then MONUME$ = "1"
    Call GRACONTROL("BOLRECEP", "ETIQTRAZ", MONUME$)

    MONUME$ = "0": If Check17.Value = 1 Then MONUME$ = "1"
    Call GRACONTROL("BOLRECEP", "RECECRON", MONUME$)
    '
    MONUME$ = "0": If Check18.Value = 1 Then MONUME$ = "1"
    Call GRACONTROL("BOLRECEP", "CRONBLOQ", MONUME$)
    '
    MONUME$ = "NO": If Option9.Value = True Then MONUME$ = "SI"
    Call GRACONTROL("PRINTETI", "PORUNIDA", MONUME$)

    MONUME$ = "NO": If Option10.Value = True Then MONUME$ = "SI"
    Call GRACONTROL("PRINTETI", "PORBULTO", MONUME$)
    '
    MONUME$ = SAFETEXT$(XVALO(Text15))
    Call GRACONTROL("PRINTETI", "CANTCOLU", MONUME$)
    
    MONUME$ = "0" ' INFORMA DATOS DE IMPORTACION EN LA RECEPCION SOLO DE PROVEEDORES EXTERNOS
    If Check20.Value = 1 Then MONUME$ = "1"
    Call GRACONTROL("BOLRECEP", "RECEIMPO", MONUME$)
    
    Call CIERRARCH("*.*")
    Unload Me
    '
99  Screen.MousePointer = 1
    Command2.Enabled = True
    
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
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text3.TEXT = VDEVU$
        Text4.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command5_Click()
    Call CONECRUN("FLFORMS/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command6_Click()
    Call CONECRUN("FLFORMS/" + TRIM$(Text3.TEXT), "Formularios de Impresion")
End Sub

Private Sub Check1_CLICK(Index As Integer)
    If Index = 0 Then
        Check1(Index).Value = 1
    End If
    
    If Index = 1 And Check1(Index).Value = 1 Then
        Check21.Enabled = True
    Else
        Check21.Value = 0
        Check21.Enabled = False
    End If
End Sub

Private Sub Command8_Click()
    If Option8.Value = True Then
        Call CONECRUN("FLFORMS/" + TRIM$(Text10.TEXT), "Formularios de Impresion")
      Else
        BLOCNOTA.Show 1
    End If
End Sub

Private Sub Command9_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text10.TEXT = VDEVU$
        text9.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Form_Load()
    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    '
    Text1.TEXT = Control$("", "FORMBREC")
    Text2.TEXT = ECOARCH$("INDIFRM", Text1.TEXT)
    '
    If Control$("", "NUMEBREC") = "MANUAL" Then Option1.Value = True
    '
    CMO$ = TRIM$(Control$("", "CODMBREC")): If CMO$ <> "" Then Text5.TEXT = CMO$
    '
    Text3.TEXT = Control$("", "REMIRECH")
    Text4.TEXT = ECOARCH$("INDIFRM", Text3.TEXT)
    
    TEXT13 = Control("BOLRECEP", "ETIQRECE")
    Text14.TEXT = ECOARCH$("INDIFRM", TEXT13.TEXT)
    '
    Text10.TEXT = Control$("", "STIKAPRO")
    text9.TEXT = ECOARCH$("INDIFRM", Text10.TEXT)
    If Left$(UCase$(Text10), 3) = "STI" Then
      Check4.Value = 1
    End If
    '
    If Control$("", "NUMERECH") = "MANUAL" Then Option5.Value = True
    '
    CMO$ = TRIM$(Control$("", "CODRRECH")): If CMO$ <> "" Then Text6.TEXT = CMO$
    '
    Check1(1).Value = 0
    If TRIM$(Control$("", "RECPAPEL")) = "SI" Then
        Check1(1).Value = 1
        Check21.Enabled = True
        
        Check21.Value = 0
        If TRIM$(Control$("", "REPAMTLI")) = "SI" Then
            Check21.Value = 1
        End If
    End If
    '
    '
    If Control$("BOLRECEP", "SUPRINT") = "SI" Then Check2.Value = 1
    '
    If Control$("BOLRECEP", "RESOLOOC") = "SI" Then Check11.Value = 1
    
    Option3.Value = True
    If TRIM$(Control$("", "COMPBREC")) = "FICHA" Then
        Option4.Value = True
    End If
    '
    Text7 = TRIM$(Control$("BOLRECEP", "DIANREC"))
    If Control$("BOLRECEP", "ANUCANTI") = "SI" Then
      Check3.Value = 1
    End If
    '
    '\\\OT-06-0515-RG-28/11/06///
    If Control$("BOLRECEP", "SUPREMI") = "SI" Then Check6.Value = 1
    '\\\OT-06-0515-RG-FIN///
    '
    If Control$("BOLRECEP", "BLOSOBRE") = "SI" Then
      Check5.Value = 1
      Text8 = FORMATNUM(Val(Control("BOLRECEP", "PORSOBRE")), "F5.1")
    End If
    '
    If Control$("BOLRECEP", "NUCONVIG") = "SI" Then Check10.Value = 1
    If Control$("BOLRECEP", "CERCAPRO") = "SI" Then Check7.Value = 1
    '
    If Control$("BOLRECEP", "FORMATO") = "TXT" Then Option7.Value = True
    '
    'CONTROLES DE IMPRESION DE DESCRIPCION
    If Control("BOLRECEP", "DESCPRIM") = "SI" Then Check9.Value = 1
    If Control("BOLRECEP", "DESCOMPL") = "SI" Then Check4.Value = 1
    If Control("BOLRECEP", "ESPECIFI") = "SI" Then Check8.Value = 1
    If Control("BOLRECEP", "CIERREOC") = "SI" Then Check12.Value = 1
    If Control("BOLRECEP", "CTRLPESO") = "SI" Then
        Check22.Value = 1
        Text8 = FORMATNUM(Val(Control("BOLRECEP", "PORSOBRE")), "F5.1")
    End If
    If Control("BOLRECEP", "VEBRECON") = "SI" Then
        Check14.Value = 1
        Command13.Enabled = True
    End If
    If Control("BOLRECEP", "CONPRECI") = "SI" Then Check15.Value = 1
    If Control("BOLRECEP", "IMETIREC") = "SI" Then
       Check16.Value = 1
       Command17.Visible = False
    Else
       Command16.Visible = True
    End If

    '
    Text16 = Control("BOLRECEP", "DEPEQIMP")
    DESTI$ = Control("BOLRECEP", "DETIRECH"): If XVALO(DESTI$) > 0 Then Text11 = XVALO(DESTI$) ' DESTINO DE DEVOLUCION)
    PFAC$ = Control$(IDENTER$, "PORTSTIK")
    If PFAC$ = "" Then PFAC$ = Control$("*", "PORTSTIK")
    Combo1.Clear
    Dim PRX As Printer
    On Error Resume Next
    Combo1.TEXT = Printer.DeviceName
    On Error GoTo 0
    For Each PRX In Printers
      Combo1.AddItem PRX.DeviceName
      If PRX.DeviceName = PFAC$ Then
        Combo1.TEXT = PFAC$
      End If
    Next
    PFAC$ = TRIM$(Control$(IDENTER$, "PORTSTIK"))
    If PFAC$ <> "" Then Combo1.TEXT = PFAC$
    '
    Text12 = Control("BOLRECEP", "MAIPAÑOL")
    '
    If Control("BOLRECEP", "DESCUNRE") = "SI" Then Check13.Value = 1
    
    Check17.Value = XVALO(Control("BOLRECEP", "RECECRON"))
    '
    Check18.Value = XVALO(Control("BOLRECEP", "CRONBLOQ"))
    
    MONUME$ = Control("PRINTETI", "PORUNIDA")
    If MONUME$ = "SI" Then Option9.Value = True
    
    MONUME$ = Control("PRINTETI", "PORBULTO")
    If MONUME$ = "SI" Then Option10.Value = True
    
    Check19.Value = XVALO(Control("BOLRECEP", "ETIQTRAZ"))
    
    Text15 = XVALO(Control("PRINTETI", "CANTCOLU"))
    
    ' INFORMA DATOS DE IMPORTACION EN LA RECEPCION SOLO DE PROVEEDORES EXTERNOS
    Command18.Caption = "Informa Despacho A.Import."
    Check20.Value = XVALO(Control("BOLRECEP", "RECEIMPO"))
    If Check20.Value > 0 Then Command18.Caption = "Informa Despacho A.Import. (1)"
    
End Sub

Private Sub mnuTabAux_Click()

End Sub

Private Sub mnuCarTabAux_Click()
   Call GENERAR_ESTRUCTURAS_SQL
   Call PRESENTA_TABLAS_AUXILIARES
End Sub

Private Sub mnuLecturaPesoBalanza_Click()

    Tara# = XVALO(Control$("BOLRECEP", "TARAMET"))
    Puerto$ = TRIM$(Control$(IDENTER$, "PORTBAL"))
    Baudios# = XVALO(Control$(IDENTER$, "BAUDBAL"))
    
10  VDEF$ = MKS$(Tara#) & Left$(Puerto$ & "   ", 3) & MKS$(Baudios#)
    OBX$ = OBJPLA$("CONFIBALAN", VDEF$)
    If OBX$ = "" Then GoTo 99

    Tara# = CVS(Mid$(OBX$, 1, 4))
    Puerto$ = TRIM$(Mid$(OBX$, 5, 3))
    Baudios# = CVS(Mid$(OBX$, 8, 4))
    
    If XVALO(Puerto$) < 1 And XVALO(Puerto$) > 99 Then
        Call COMUNI("Valor de Puerto No Aceptado")
        GoTo 10
    End If
    If XVALO(Baudios#) < 1 Then
        Call COMUNI("Valor de Baudios No Aceptado")
        GoTo 10
    End If
    
    Call GRACONTROL("BOLRECEP", "TARAMET", TRIM$(FORMATNUM$(Tara#, "F10.2")))
    Call GRACONTROL(IDENTER$, "PORTBAL", TRIM$(Puerto$))
    Call GRACONTROL(IDENTER$, "BAUDBAL", TRIM$(FORMATNUM$(Baudios#, "F10.2")))
    
99:
    Exit Sub
    
End Sub

Private Sub mnuTablaTaras_Click()

    Call TablasAuxiliares
    
    IDLIST& = ObtenerIdTablaAuxiliar("TABLATARAS")
    Call ABM_TABLA(IDLIST&, "Tabla de Taras")

End Sub

Private Sub mnuTaraPorTipoDeRecepcion_Click()

    Call TablasAuxiliares
    
    On Error Resume Next
    FLEXCONN.Execute "delete #TPTiposRcepcion"
    ConsultaSql$ = "create table #TPTiposRecepcion (codigo int, denominacion varchar(64))"
    FLEXCONN.Execute ConsultaSql$
    On Error GoTo 0
    
    On Error GoTo ControlError
    Dim Rst As New ADODB.Recordset
    ConsultaSql$ = "select * from #TPTiposRecepcion"
    Rst.Open ConsultaSql$, FLEXCONN, adOpenKeyset, adLockOptimistic, adCmdText
    With Rst
        For i& = 1 To ULTREG&("TATIREC")
            regi$ = REGLEIDO$("TATIREC", i&)
            Codigo& = Asc(Mid$(regi$, 1, 1))
            Denominacion$ = Mid$(regi$, 2, 30)
                    
            .AddNew
            !Codigo = Codigo&
            !Denominacion = Denominacion$
            .Update
                
        Next i&
    End With
    Rst.Close
    Set Rst = Nothing
    On Error GoTo 0
    
    IDLIST& = ObtenerIdTablaAuxiliar("TARATIRECE")
    Call ABM_TABLA(IDLIST&, "Tara por Tipo de Recepción")
    
    On Error Resume Next
    FLEXCONN.Execute "delete #TPTiposRecepcion"
    On Error GoTo 0
    
    Exit Sub

ControlError:
    Exit Sub
    

End Sub

Private Sub Text1_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text10_Change()
'   If Left$(UCase$(Text10), 3) <> "STI" Then
'      Check4.Value = 0
'   End If
End Sub

Private Sub Text10_DblClick()
   Call Command9_Click
End Sub

Private Sub Text11_Change()
    '
    If TRIM$(Text11) <> "" And XVALO(Text11) < 3 Then
      Label8 = ECOARCH$("TADESTI", MKI$(XVALO(Text11)))
    Else
      Label8 = ""
    End If
    '
End Sub

Private Sub Text13_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        TEXT13.TEXT = VDEVU$
        Text14.TEXT = VALACT2$("INDIFRM")
    End If

End Sub

Private Sub Text15_LostFocus()
   Text15 = FORMATNUM$(XVALO(Text15), "f5.1")
End Sub

Private Sub Text16_Change()
  Label9.Caption = ECOARCH$("TADEPOSI", Chr$(XVALO(Text16)))
End Sub

Private Sub Text2_CLICK()
    Text1.SetFocus
End Sub

Private Sub Text4_CLICK()
    Text3.SetFocus
End Sub

Private Sub Text3_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text3.TEXT = VDEVU$
        Text4.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text8_LostFocus()
   Text8 = FORMATNUM$(Val(Text8), "f5.1")
End Sub

Private Sub Text9_Click()
    Text10.SetFocus
End Sub

Sub TablasAuxiliares()

    Dim TablaAuxiliar As New DatosAuxiliares
    Dim ListaCampos() As eDatosAuxiliaresCampos
    Dim CAMPO As eDatosAuxiliaresCampos
    
    ' TABLA DE TARAS
    ' **************
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Ord/I10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Descripcion/A32"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
        
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Peso/F12.2"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Call TablaAuxiliar.CrearTabla("TABLATARAS", "Tabla de Taras", ListaCampos)

    ' TABLA DE TARAS POR TIPO DE RECEPCION
    ' ************************************
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Ord/I10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "C-R/I5"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = "codigo;denominacion;#TPTiposRecepcion;codigo>0;order by codigo"
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Denominacion/A32"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
        
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "C-T/I5"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    IDLIST& = ObtenerIdTablaAuxiliar("TABLATARAS")
    CAMPO.ColumnaCampoEco = "campo1;campo2;datasql;id_lista=" & CStr(IDLIST&) & ";order by cast(campo1 as int)"
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Descripcion/A32"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Call TablaAuxiliar.CrearTabla("TARATIRECE", "Tara por Tipo Recepción", ListaCampos)

End Sub
