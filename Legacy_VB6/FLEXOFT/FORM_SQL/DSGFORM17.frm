VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form DSGFORM17 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generador de Formularios de Impresion"
   ClientHeight    =   10995
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   10995
   ScaleWidth      =   15240
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00404040&
      Height          =   75
      Left            =   11865
      TabIndex        =   9
      Top             =   4935
      Width           =   15000
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Propiedades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6100
      Left            =   11640
      TabIndex        =   57
      Top             =   5520
      Width           =   3375
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   1020
         TabIndex        =   67
         Text            =   " "
         Top             =   2430
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   810
         TabIndex        =   81
         Text            =   "0.0"
         Top             =   1470
         Width           =   750
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   810
         TabIndex        =   80
         Text            =   "0.0"
         Top             =   1770
         Width           =   750
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2160
         TabIndex        =   79
         Text            =   "0.0"
         Top             =   1470
         Width           =   750
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2160
         TabIndex        =   78
         Text            =   "0.0"
         Top             =   1770
         Width           =   750
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   810
         Sorted          =   -1  'True
         TabIndex        =   77
         Text            =   "Combo1"
         Top             =   2985
         Width           =   2115
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   820
         Left            =   810
         TabIndex        =   76
         Text            =   "8"
         Top             =   3555
         Width           =   820
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   810
         TabIndex        =   75
         Top             =   960
         Width           =   2115
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   810
         TabIndex        =   74
         Text            =   " "
         Top             =   2220
         Width           =   2115
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Del"
         Height          =   250
         Left            =   1995
         TabIndex        =   73
         ToolTipText     =   "Elimina Objeto Activo"
         Top             =   360
         Width           =   405
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   285
         Left            =   1335
         Max             =   1
         Min             =   1
         TabIndex        =   72
         Top             =   360
         Value           =   1
         Width           =   150
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Negrita"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   1755
         TabIndex        =   71
         Top             =   3345
         Width           =   1065
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Italica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   1755
         TabIndex        =   70
         Top             =   3555
         Width           =   1170
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Subrayado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   1755
         TabIndex        =   69
         Top             =   3765
         Width           =   1170
      End
      Begin VB.CommandButton Command13 
         Height          =   285
         Left            =   2370
         Picture         =   "DSGFORM17.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   2220
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Copy"
         Height          =   250
         Left            =   2415
         TabIndex        =   66
         ToolTipText     =   "Copia Objeto Activo"
         Top             =   360
         Width           =   500
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "  InVisible  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   -1  'True
         EndProperty
         Height          =   250
         Left            =   1755
         TabIndex        =   65
         Top             =   3975
         Width           =   1170
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Prio"
         Height          =   250
         Left            =   1575
         TabIndex        =   64
         ToolTipText     =   "Incrementa Prioridad"
         Top             =   360
         Width           =   405
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Encasillado"
         Height          =   250
         Left            =   1755
         TabIndex        =   63
         Top             =   4185
         Width           =   1170
      End
      Begin VB.CommandButton Command18 
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
         Height          =   285
         Left            =   2880
         TabIndex        =   62
         Top             =   2220
         Width           =   175
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Formato Multilinea "
         Height          =   255
         Left            =   810
         TabIndex        =   61
         Top             =   2595
         Width           =   2070
      End
      Begin VB.CommandButton Command20 
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
         Height          =   285
         Left            =   2880
         TabIndex        =   60
         Top             =   2235
         Width           =   175
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
         Height          =   285
         Left            =   3000
         TabIndex        =   59
         Top             =   4635
         Width           =   175
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   795
         TabIndex        =   58
         Text            =   " "
         Top             =   4635
         Width           =   2235
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Orden:"
         Height          =   225
         Left            =   -975
         TabIndex        =   94
         Top             =   390
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
         Height          =   225
         Left            =   -975
         TabIndex        =   93
         Top             =   690
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         Height          =   225
         Left            =   -975
         TabIndex        =   92
         Top             =   990
         Width           =   1695
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pos.- X1:"
         Height          =   225
         Left            =   -975
         TabIndex        =   91
         Top             =   1500
         Width           =   1695
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X2:"
         Height          =   225
         Left            =   -975
         TabIndex        =   90
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Y1:"
         Height          =   225
         Left            =   1620
         TabIndex        =   89
         Top             =   1500
         Width           =   465
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Y2:"
         Height          =   225
         Left            =   1620
         TabIndex        =   88
         Top             =   1800
         Width           =   465
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fuente:"
         Height          =   225
         Left            =   -1080
         TabIndex        =   87
         Top             =   3015
         Width           =   1800
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tamaño:"
         Height          =   225
         Left            =   -510
         TabIndex        =   86
         Top             =   3345
         Width           =   1695
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   810
         TabIndex        =   85
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label15 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   810
         TabIndex        =   84
         Top             =   660
         Width           =   2115
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Texto:"
         Height          =   225
         Left            =   -975
         TabIndex        =   83
         Top             =   2265
         Width           =   1695
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Destino D-Click:"
         Height          =   465
         Left            =   -30
         TabIndex        =   82
         Top             =   4515
         Width           =   735
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Configuracion General"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   12360
      TabIndex        =   42
      Top             =   600
      Width           =   3375
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Orientacion"
         Height          =   645
         Left            =   135
         TabIndex        =   54
         Top             =   360
         Width           =   3300
         Begin VB.OptionButton Option1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Retrato (VERT)"
            Height          =   330
            Left            =   315
            TabIndex        =   56
            Top             =   250
            Value           =   -1  'True
            Width           =   1590
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Paisaje (HOR)"
            Height          =   330
            Left            =   1785
            TabIndex        =   55
            Top             =   250
            Width           =   1380
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Margenes de Impresion"
         Height          =   1110
         Left            =   135
         TabIndex        =   48
         Top             =   1050
         Width           =   3300
         Begin VB.TextBox Text12 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   840
            TabIndex        =   51
            Text            =   "0.0"
            Top             =   315
            Width           =   750
         End
         Begin VB.TextBox Text13 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2415
            TabIndex        =   50
            Text            =   "0.0"
            Top             =   315
            Width           =   750
         End
         Begin VB.CheckBox Check7 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Impresion Automática al Guardar"
            Height          =   255
            Left            =   360
            TabIndex        =   49
            Top             =   720
            Width           =   2775
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Superior:"
            Height          =   225
            Left            =   1620
            TabIndex        =   53
            Top             =   345
            Width           =   705
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Izquierdo:"
            Height          =   225
            Left            =   -945
            TabIndex        =   52
            Top             =   345
            Width           =   1695
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Navegacion"
         Height          =   645
         Left            =   135
         TabIndex        =   45
         Top             =   2205
         Width           =   3300
         Begin VB.CheckBox Check8 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Boton P.Ant."
            Height          =   255
            Left            =   360
            TabIndex        =   47
            Top             =   300
            Width           =   1215
         End
         Begin VB.CheckBox Check9 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Boton P.Sig."
            Height          =   255
            Left            =   1920
            TabIndex        =   46
            Top             =   300
            Width           =   1215
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Formulario Operativo"
         Height          =   645
         Left            =   135
         TabIndex        =   43
         Top             =   2925
         Width           =   3300
         Begin VB.TextBox Text16 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   600
            TabIndex        =   44
            Top             =   240
            Width           =   2055
         End
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0FFFF&
      Height          =   5415
      Left            =   90
      TabIndex        =   40
      Top             =   1060
      Width           =   580
      Begin VB.PictureBox Picture9 
         BorderStyle     =   0  'None
         Height          =   5250
         Left            =   35
         ScaleHeight     =   5250
         ScaleWidth      =   510
         TabIndex        =   41
         Top             =   100
         Width           =   515
         Begin VB.Image Image1 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "DSGFORM17.frx":014A
            ToolTipText     =   "Diseño General"
            Top             =   0
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   405
            Index           =   1
            Left            =   50
            Picture         =   "DSGFORM17.frx":0A14
            ToolTipText     =   "Leyenda Fija"
            Top             =   540
            Width           =   405
         End
         Begin VB.Image Image1 
            Height          =   465
            Index           =   2
            Left            =   0
            Picture         =   "DSGFORM17.frx":1332
            ToolTipText     =   "Parametro"
            Top             =   1050
            Width           =   510
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   3
            Left            =   45
            Picture         =   "DSGFORM17.frx":200C
            ToolTipText     =   "Linea"
            Top             =   2160
            Width           =   435
         End
         Begin VB.Image Image1 
            Height          =   465
            Index           =   4
            Left            =   5
            Picture         =   "DSGFORM17.frx":2B4E
            ToolTipText     =   "Caja"
            Top             =   2700
            Width           =   525
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   5
            Left            =   10
            Picture         =   "DSGFORM17.frx":38A4
            ToolTipText     =   "Cuadro de Imagen"
            Top             =   3290
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   6
            Left            =   0
            Picture         =   "DSGFORM17.frx":3BAE
            Stretch         =   -1  'True
            ToolTipText     =   "Caja de Texto"
            Top             =   1605
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   450
            Index           =   7
            Left            =   0
            Picture         =   "DSGFORM17.frx":3CF8
            Stretch         =   -1  'True
            ToolTipText     =   "Grilla - Planilla"
            Top             =   3885
            Width           =   450
         End
         Begin VB.Image Image1 
            Height          =   450
            Index           =   8
            Left            =   30
            Picture         =   "DSGFORM17.frx":3E42
            Stretch         =   -1  'True
            ToolTipText     =   "Multiobjeto Firma Electronica"
            Top             =   4485
            Width           =   450
         End
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   4080
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   38
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "M-Sel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   0
      TabIndex        =   31
      Top             =   8880
      Width           =   735
      Begin VB.CommandButton Command25 
         Caption         =   "Del"
         Height          =   250
         Left            =   50
         TabIndex        =   35
         Top             =   1020
         Width           =   645
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Copy"
         Height          =   250
         Left            =   50
         TabIndex        =   34
         ToolTipText     =   "Hace Coincidir Borde Superior de Objetos Multiseleccionados"
         Top             =   760
         Width           =   645
      End
      Begin VB.CommandButton Command23 
         Caption         =   "V-Align"
         Height          =   250
         Left            =   50
         TabIndex        =   33
         ToolTipText     =   "Hace Coincidir Borde Superior de Objetos Multiseleccionados"
         Top             =   500
         Width           =   645
      End
      Begin VB.CommandButton Command22 
         Caption         =   "H-Align"
         Height          =   250
         Left            =   50
         TabIndex        =   32
         ToolTipText     =   "Hace Coincidir Borde Izquierdo de Objetos Multiseleccionados"
         Top             =   240
         Width           =   645
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   9840
      Top             =   360
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9270
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Test"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   11040
      TabIndex        =   28
      Top             =   600
      Width           =   720
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   10100
      ScaleHeight     =   480
      ScaleWidth      =   5175
      TabIndex        =   20
      Top             =   0
      Width           =   5180
      Begin VB.CommandButton Command26 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3360
         Picture         =   "DSGFORM17.frx":4284
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Envio por Mail "
         Top             =   0
         Width           =   600
      End
      Begin VB.CommandButton Command19 
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   15
         Picture         =   "DSGFORM17.frx":458E
         TabIndex        =   29
         ToolTipText     =   "Importa Formulario desde Formato Antiguo"
         Top             =   0
         Width           =   350
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
         Height          =   480
         Left            =   960
         Picture         =   "DSGFORM17.frx":4898
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Configuracion de Funcionamiento"
         Top             =   0
         Width           =   600
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
         Height          =   480
         Left            =   1560
         Picture         =   "DSGFORM17.frx":4CDA
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Refresca y Re-Dibuja Todo"
         Top             =   0
         Width           =   600
      End
      Begin VB.CommandButton Command3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   360
         Picture         =   "DSGFORM17.frx":4E24
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Genera Nuevo Formulario de Impresión"
         Top             =   0
         Width           =   600
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
         Height          =   480
         Left            =   4560
         Picture         =   "DSGFORM17.frx":512E
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Termina - Cierra la Aplicación"
         Top             =   0
         Width           =   600
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
         Height          =   480
         Left            =   3960
         Picture         =   "DSGFORM17.frx":5278
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   0
         Width           =   600
      End
      Begin VB.CommandButton Command5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2760
         Picture         =   "DSGFORM17.frx":5582
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Guarda Cambios en Formulario Activo"
         Top             =   0
         Width           =   600
      End
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2160
         Picture         =   "DSGFORM17.frx":588C
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Imprime Muestra del Formulario Activo"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   600
      End
   End
   Begin VB.PictureBox Picture5 
      Height          =   9780
      Left            =   735
      ScaleHeight     =   9720
      ScaleWidth      =   10980
      TabIndex        =   10
      Top             =   1155
      Width           =   11040
      Begin VB.HScrollBar HScroll1 
         Height          =   225
         LargeChange     =   100
         Left            =   0
         TabIndex        =   17
         Top             =   9450
         Width           =   10380
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H00E0E0E0&
         Height          =   6105
         Left            =   0
         ScaleHeight     =   6045
         ScaleWidth      =   480
         TabIndex        =   19
         Top             =   0
         Width           =   540
         Begin VB.Line Line2 
            X1              =   510
            X2              =   510
            Y1              =   0
            Y2              =   25000
         End
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   6960
         LargeChange     =   100
         Left            =   10710
         TabIndex        =   11
         Top             =   0
         Width           =   225
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H00E0E0E0&
         Height          =   520
         Left            =   0
         ScaleHeight     =   465
         ScaleWidth      =   11085
         TabIndex        =   18
         Top             =   0
         Width           =   11145
         Begin VB.Line Line1 
            X1              =   0
            X2              =   18000
            Y1              =   450
            Y2              =   450
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00E0E0E0&
         Height          =   16630
         Left            =   -30
         ScaleHeight     =   16575
         ScaleWidth      =   11085
         TabIndex        =   12
         Top             =   -30
         Width           =   11145
         Begin VB.PictureBox Picture1 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            Height          =   9000
            Left            =   525
            ScaleHeight     =   8940
            ScaleWidth      =   10320
            TabIndex        =   13
            Top             =   500
            Width           =   10380
            Begin VB.TextBox Text3 
               BackColor       =   &H00FFFFFF&
               Height          =   750
               Left            =   420
               MultiLine       =   -1  'True
               TabIndex        =   15
               Top             =   210
               Visible         =   0   'False
               Width           =   2220
            End
            Begin VB.ListBox List1 
               Height          =   2205
               Left            =   420
               TabIndex        =   14
               Top             =   1470
               Visible         =   0   'False
               Width           =   3165
            End
            Begin VB.Line Line4 
               BorderColor     =   &H000000C0&
               BorderStyle     =   5  'Dash-Dot-Dot
               Visible         =   0   'False
               X1              =   5145
               X2              =   5145
               Y1              =   0
               Y2              =   20000
            End
            Begin VB.Line Line3 
               BorderColor     =   &H000000FF&
               BorderStyle     =   5  'Dash-Dot-Dot
               Visible         =   0   'False
               X1              =   0
               X2              =   20000
               Y1              =   4200
               Y2              =   4200
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Label17"
               Height          =   255
               Left            =   525
               TabIndex        =   16
               Top             =   1050
               Visible         =   0   'False
               Width           =   630
            End
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00404040&
      Height          =   11040
      Left            =   11865
      TabIndex        =   8
      Top             =   -100
      Width           =   75
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Formulario Activo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   9210
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
         Height          =   360
         Left            =   3150
         TabIndex        =   3
         Top             =   370
         Width           =   170
      End
      Begin VB.CommandButton Command27 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Picture         =   "DSGFORM17.frx":5EF6
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Refresca Lista de Seleccion"
         Top             =   360
         Width           =   300
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3465
         MultiLine       =   -1  'True
         TabIndex        =   4
         Text            =   "DSGFORM17.frx":6040
         Top             =   370
         Width           =   5055
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
         Height          =   360
         Left            =   1260
         TabIndex        =   2
         Text            =   " "
         Top             =   370
         Width           =   1905
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   8600
         Picture         =   "DSGFORM17.frx":6044
         Top             =   255
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   8600
         Picture         =   "DSGFORM17.frx":634E
         Top             =   250
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
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
         Left            =   525
         TabIndex        =   1
         Top             =   525
         Width           =   700
      End
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   4725
      Pattern         =   "*.FRM"
      TabIndex        =   5
      Top             =   6930
      Visible         =   0   'False
      Width           =   2745
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "DSGFORM17.frx":6658
      Top             =   7320
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "DSGFORM17.frx":688C
      TabIndex        =   39
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   9720
      TabIndex        =   30
      Top             =   750
      Width           =   975
   End
   Begin VB.Label TEPRUORI 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   105
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label TEPRURED 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   0
      TabIndex        =   6
      Top             =   315
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "DSGFORM17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARRABOR%
Dim PORCOORDE%
'
Dim XMOUSE As Single, YMOUSE As Single
Dim ESTAPRETADO%, MUEVETODO%
Dim XORI, YORI
Dim MODO$
Dim UTAFUE, UNEGRI, UITALI, USUBRA
'
Dim XREGLA, YREGLA
Dim MULTISEL%
Dim DHOR, DVER                    ' DESPLAZAMIENTO HORIZONTAL Y VERTICAL EN MULTICOPIA
'
Dim FRESHLIST%
Dim SOLOLEC%                      ' SI YA ESTA ABIERTO PONE SOLOLEC% EN 1 Y LO ABRE EN SOLO LECTURA
Dim NKKK%                         ' NUMERO DE ARCHIVO DE BLOQUEO
'
Sub VALIDAFORM(OK%)
    '
    OK% = 0
    '
    If LOCK_DSG$ = "SI" Then
       If Image2.Visible = False Then
          Call MENSERR(25, "Diseño Bloqueado - Imposible Guardar Cambios")
          Exit Sub
       End If
    End If
    '
    If FRESHLIST% >= 0 Then
       Call FILTERSETRECORD("INDISQR", 1, CODFOR$, AJUSTI$(CODFOR$, 8) + Text2.TEXT)
       Call FRESHECHO("INDISQR")
    End If
    '
    ' VERIFICA REPETICION DE NOMBRES Y REFERENCIAS
    For KKI% = 1 To CANOBJ%
       If TRIM$(ONAME$(KKI%)) = "" Then
          Call MENSERR(24, "Falta Nombre de Objeto")
          Label14 = KKI%
          Text9.SetFocus
          Exit Sub
       End If
       '
       For KKJ% = 1 To KKI% - 1
          If ONAME$(KKI%) = ONAME$(KKJ%) Then
             Call MENSERR(24, "Nombre de Objeto Repetido")
             Label14 = KKI%
             Text9.SetFocus
             Exit Sub
          End If
       Next KKJ%
       '
       ' VERIFICA REFERENCIA A TABLAS EXISTENTES
       If TRIM$(TASEL$(KKI%)) <> "" Then
          If EXISTELATABLA%(TRIM$(TASEL$(KKI%))) < 1 Then
             Call MENSERR(24, "Referencia a Tabla SQL '" + TRIM$(TASEL$(KKI%)) + "' Inexistente.")
             Label14 = KKI%
             Text9.SetFocus
             Exit Sub
          End If
       End If
       '
       ' VERIFICA ORDENAMIENTO
       VUELTA% = 1
       TREP$ = TRIM$(FOREPLA$(KKI%))
9      If TRIM$(TREP$) <> "" Then
10        PPXX1% = InStr(TREP$, "{")
          If PPXX1% > 0 Then      ' BUSCA ENCIERRO ENTRE LLAVES
             If PPXX1% > 1 Then
               If Mid$(TREP$, PPXX1% - 1, 1) = "%" Then GoTo 35 ' REFERENCIA DE OTRA HOJA
             End If
             '
             PPXX2% = InStr(TREP$, "}"): If PPXX2% <= PPXX1% Then PPXX2% = 1 + Len(TREP$)
             IZQUI$ = Left$(TREP$, PPXX1% - 1)
             DEREX$ = Mid$(TREP$, PPXX2% + 1)
             VAREPLA$ = Mid$(TREP, PPXX1% + 1, PPXX2% - PPXX1% - 1)
               If UCase$(Left(VAREPLA$, 7)) = "COMMPAR" Then VAREPLA$ = "COMPAR" + Mid$(VAREPLA$, 8)
               If UCase$(Left$(VAREPLA$, 6)) = "PARCOM" Then VAREPLA$ = "COMPAR" + Mid$(VAREPLA$, 7)
               If VAREPLA$ = "COMPAR" Then VAREPLA$ = "COMPAR-0"
             If Left$(UCase$(VAREPLA$), 6) = "COMPAR" Then GoTo 29
             For JJ& = 1 To KKI% - 1
               If UCase$(ONAME$(JJ&)) = UCase$(VAREPLA$) Then
                  VAREPLA = ""
                  GoTo 29
               End If
             Next JJ&
             '
             POVARE% = 0: TTXX$ = "Inexistente"
             For JJ& = KKI% To CANOBJ%
               If UCase$(ONAME$(JJ&)) = UCase$(VAREPLA$) Then POVARE% = JJ&
             Next JJ&
             If POVARE% > 0 Then TTXX$ = "en Posicion " & POVARE%
             Call MENSERR(24, "Objeto '" + ONAME$(KKI%) + "' en Posición " & KKI% & "\Referencia Parametro '" & VAREPLA$ & "' " & TTXX$ & ".")
             If POVARE% > 0 Then
                 Label14 = POVARE%
               Else
                 Label14 = KKI%
             End If
             Exit Sub
             '
29           TREP$ = IZQUI$ + VAREPLA$ + DEREX$
             GoTo 10
          End If
       End If
35     If VUELTA% = 1 Then
         VUELTA% = 2
         TREP$ = CONDSEL$(KKI%)
         GoTo 9
       End If
    Next KKI%
    '
    OK% = 1
End Sub

'
Private Sub Check1_Click()
   Call PREDITEXTO
   If IOBJ% <= CANOBJ% + 1 Then
     NEGRI(IOBJ%) = Check1.Value
     UNEGRI = Check1.Value
   End If
End Sub

Private Sub Check2_Click()
   Call PREDITEXTO
   If IOBJ% <= CANOBJ% + 1 Then
     ITALI(IOBJ%) = Check2.Value
     UITALI = Check2.Value
   End If
End Sub

Private Sub Check3_Click()
   Call PREDITEXTO
   If IOBJ% <= CANOBJ% + 1 Then
     SUBRA(IOBJ%) = Check3.Value
     USUBRA = Check3.Value
   End If
End Sub

Private Sub Check4_Click()
   Call PREDITEXTO
   If IOBJ% <= CANOBJ% + 1 Then
     INVIS(IOBJ%) = Check4.Value
   End If
End Sub

Private Sub Check5_Click()
   Call PREDITEXTO
   If IOBJ% <= CANOBJ% + 1 Then
     ENCAS(IOBJ%) = Check5.Value
   End If
End Sub

Private Sub Check6_Click()
   Call PREDITEXTO
   If IOBJ% <= CANOBJ% + 1 Then
     MULIN(IOBJ%) = Check6.Value
   End If
End Sub

Private Sub Check7_Click()
    AUTO_PRN$ = ""
    If Check7.Value = 1 Then AUTO_PRN$ = "SI"
End Sub

Private Sub Check8_Click()
    PAGI_ANT$ = ""
    If Check8.Value = 1 Then PAGI_ANT$ = "SI"
End Sub

Private Sub Check9_Click()
    PAGI_SIG$ = ""
    If Check9.Value = 1 Then PAGI_SIG$ = "SI"
End Sub

Private Sub Combo1_Change()
    Call PREDITEXTO
End Sub

Private Sub Combo1_Click()
    If XVALO(Label14) > 0 Then
       If XVALO(Label14) <= CANOBJ% Then
          FUENTE$(XVALO(Label14)) = Combo1.TEXT
       End If
    End If
    Call PREDITEXTO
End Sub

Private Sub Command1_Click()
    '
    FORSELEC.ABUSQUE = Text1
    Call SELECHO("INDISQR")
    CODFOR$ = VALACT1$("INDISQR")
    If TRIM$(CODFOR$) = "" Then
        Text1.SetFocus
      Else
        Text1.TEXT = CODFOR$
    End If
    '
End Sub

Sub CARGAFORM()
    '
    CANOBJ% = 0
    NX% = FILEOPEN%(LU$ + CODFOR$ + ".FRM", 0, 128)
    Dim RF As String * 128
    '
    ' CARGA LAS CAJAS
    MAXCA% = 0
    For i% = 1 To LOF(NX%) / 128
        Get #NX%, i%, RF$
        If TRIM$(Left$(RF$, 8)) = "CAJAWIN" Then
            '
            CANOBJ% = CANOBJ% + 1
            TIOBJ$(CANOBJ%) = "BOX"
               MAXCA% = MAXCA% + 1
            ONAME$(CANOBJ%) = "BOX-" & MAXCA%
            TAFUEN(CANOBJ%) = 1
            If Asc(Mid$(RF$, 73, 1)) > 0 Then
                TAFUEN(CANOBJ%) = Asc(Mid$(RF$, 73, 1))
            End If
            '
            H1(CANOBJ%) = CVS(Mid$(RF$, 57, 4))
            V1(CANOBJ%) = CVS(Mid$(RF$, 61, 4))
            H2(CANOBJ%) = CVS(Mid$(RF$, 65, 4))
            V2(CANOBJ%) = CVS(Mid$(RF$, 69, 4))
        End If
    Next i%
    '
    ' dibuja lineas *************************************
    MAXLI% = 0
    For i% = 1 To LOF(NX%) / 128
        Get #NX%, i%, RF$
        If TRIM$(Left$(RF$, 8)) = "LINEWIN" Then
            '
            CANOBJ% = CANOBJ% + 1
            TIOBJ$(CANOBJ%) = "LINE"
               MAXLI% = MAXLI% + 1
            ONAME$(CANOBJ%) = "LINE-" & MAXLI%
            TAFUEN(CANOBJ%) = 1
            If Asc(Mid$(RF$, 73, 1)) > 0 Then
                TAFUEN(CANOBJ%) = Asc(Mid$(RF$, 73, 1))
            End If
            '
            H1(CANOBJ%) = CVS(Mid$(RF$, 57, 4))
            V1(CANOBJ%) = CVS(Mid$(RF$, 61, 4))
            H2(CANOBJ%) = CVS(Mid$(RF$, 65, 4))
            V2(CANOBJ%) = CVS(Mid$(RF$, 69, 4))
        End If
    Next i%
    '
    ' escribe leyendas fijas ***************************
    MAXLA% = 0
    For i% = 1 To LOF(NX%) / 128
        Get #NX%, i%, RF$
        If TRIM$(Left$(RF$, 8)) = "TEXTWIN" Then
           If TRIM$(Mid$(RF$, 9, 48)) <> "" Then
               FUENX$ = TRIM$(Mid$(RF$, 87, 32))
               If FUENX$ = "" Then FUENX$ = "Courier"
               TAFU = CVS(Mid$(RF$, 119, 4))
               '
               CANOBJ% = CANOBJ% + 1
               TIOBJ$(CANOBJ%) = "LABEL"
                 MAXLA% = MAXLA% + 1
               ONAME$(CANOBJ%) = "LABEL-" & MAXLA%
               FUENTE$(CANOBJ%) = FUENX$
               TAFUEN(CANOBJ%) = TAFU
               '
               Label17.Font = FUENX$
               Label17.FontSize = TAFU
               Label17.FontBold = False
               Label17.FontItalic = False
               Label17.FontUnderline = False
               '
               TEXTO(CANOBJ%) = TRIM$(Mid$(RF$, 9, 48))
               Label17.Caption = TRIM$(Mid$(RF$, 9, 48))
               '
               H1(CANOBJ%) = CVS(Mid$(RF$, 57, 4))
               H2(CANOBJ%) = H1(CANOBJ%) + Label17.Width / 56.7
               V1(CANOBJ%) = CVS(Mid$(RF$, 61, 4))
               V2(CANOBJ%) = V1(CANOBJ%) + Label17.Height / 56.7
           End If
        End If
109 Next i%
    '
End Sub
'
Private Sub Command10_Click()
   If Text3.Visible Then Exit Sub
   MODO$ = "": Label19 = ""
   Command10.Enabled = False
   Call REDIBUJA
   MULTISEL% = 0: DHOR = 0: DVER = 0
   Label19.ToolTipText = ""
   Command10.Enabled = True
End Sub

Sub REDIBUJA()
   Text3.Visible = False
   Text3 = ""
   Picture1.Cls
   '
   APRETANT% = ESTAPRETADO%
   ESTAPRETADO% = 0
   For i% = 1 To CANOBJ%
      ISEL%(i%) = 0
      Call DIBUOBJ(i%)
   Next i%
   ESTAPRETADO% = APRETANT%
End Sub
'
Sub DIBUOBJ(i%, Optional ENFOC%)
   '
   If i% > CANOBJ% Then Exit Sub
   If Text3.Visible = True Then Exit Sub
   '
   MODOX$ = TIOBJ$(i%)
   On Error Resume Next
   ' Picture1.ForeColor = RGB(0, 0, 0)
   FOCOL = RGB(0, 0, 0)
   Picture1.ForeColor = FOCOL
   If ENFOC% = 1 Or i% = XVALO(Label14) Then FOCOL = RGB(0, 150, 150): Picture1.ForeColor = FOCOL
   If ENFOC% = 2 Then FOCOL = RGB(120, 120, 0): Picture1.ForeColor = FOCOL                       ' SELECCION MULTIPLE
   '
   Select Case MODOX$
      Case "LABEL"
         Picture1.CurrentY = 56.7 * V1(i%)
         Picture1.CurrentX = 56.7 * H1(i%)
         Picture1.Font = FUENTE$(i%)
         Picture1.FontSize = TAFUEN(i%)
         Picture1.FontBold = NEGRI(i%)
         Picture1.FontItalic = ITALI(i%)
         Picture1.FontUnderline = SUBRA(i%)
         Picture1.FontStrikethru = False
            If INVIS(i%) = 1 Then Picture1.FontStrikethru = True
         If MULIN(i%) <> 1 Then
               Picture1.Print TEXTO(i%)
            Else
               TXXX$ = TEXTO(i%)
               Call FRATEXFONT(TXXX$, H2(i%) - H1(i%), FUENTE$(i%), TAFUEN(i%), NEGRI(i%))
               For KKI% = 1 To CARETEX%
                    Picture1.Print RETEX$(KKI%);
                    Label17.Font = Picture1.Font
                    Label17.FontSize = Picture1.FontSize
                    Label17.FontBold = Picture1.FontBold
                    Label17.FontItalic = Picture1.FontItalic
                    Label17 = RETEX$(KKI%) & "*M"
                    Picture1.CurrentX = 56.7 * H1(i%)
                    Picture1.CurrentY = Picture1.CurrentY + 0.85 * Label17.Height
               Next KKI%
         End If
         '
         Label17.Font = Picture1.Font
         Label17.FontSize = Picture1.FontSize
         Label17.FontBold = Picture1.FontBold
         Label17.FontItalic = Picture1.FontItalic
         Label17 = TEXTO(i%) & "*M"
      Case "PARAM"
         Picture1.DrawStyle = 0
         Picture1.DrawWidth = 1
         Picture1.Line (56.7 * (H1(i%) - 0.5), 56.7 * (V1(i%) - 0.5))-(56.7 * (H2(i%) - 1.4), 56.7 * (V1(i%) - 0.5))
         Picture1.Line (56.7 * (H1(i%) - 0.5), 56.7 * (V1(i%) - 0.5))-(56.7 * (H1(i%) - 0.5), 56.7 * (V2(i%) - 1.5))
         Picture1.DrawStyle = 3
         Picture1.Line (56.7 * (H1(i%) - 0.5), 56.7 * (V2(i%) - 1.5))-(56.7 * (H2(i%) - 1.4), 56.7 * (V2(i%) - 1.5))
         Picture1.Line (56.7 * (H2(i%) - 1.4), 56.7 * (V1(i%) - 0.5))-(56.7 * (H2(i%) - 1.4), 56.7 * (V2(i%) - 1.5))
         '
         Picture1.CurrentY = 56.7 * V1(i%)
         Picture1.CurrentX = 56.7 * H1(i%)
         Picture1.Font = FUENTE$(i%)
         Picture1.FontSize = TAFUEN(i%)
         Picture1.FontBold = NEGRI(i%)
         Picture1.FontItalic = ITALI(i%)
         Picture1.FontUnderline = SUBRA(i%)
         Picture1.FontStrikethru = False
            If INVIS(i%) = 1 Then Picture1.FontStrikethru = True
         Picture1.ForeColor = RGB(255, 0, 0)
            If ENFOC% > 0 Or i% = XVALO(Label14) Then Picture1.ForeColor = FOCOL ' RGB(0, 150, 150)
         Picture1.Print TEXTO(i%)
         Picture1.ForeColor = RGB(0, 0, 0)
         '
         Label17.Font = Picture1.Font
         Label17.FontSize = Picture1.FontSize
         Label17.FontBold = Picture1.FontBold
         Label17.FontItalic = Picture1.FontItalic
         Label17 = TEXTO(i%)
      Case "TEXT"
         Picture1.DrawStyle = 0
         Picture1.DrawWidth = 1
         Picture1.Line (56.7 * (H1(i%) - 0.5), 56.7 * (V1(i%) - 0.5))-(56.7 * (H2(i%) - 1.4), 56.7 * (V1(i%) - 0.5))
         Picture1.Line (56.7 * (H1(i%) - 0.5), 56.7 * (V1(i%) - 0.5))-(56.7 * (H1(i%) - 0.5), 56.7 * (V2(i%) - 1.5))
         Picture1.DrawStyle = 3
         Picture1.Line (56.7 * (H1(i%) - 0.5), 56.7 * (V2(i%) - 1.5))-(56.7 * (H2(i%) - 1.4), 56.7 * (V2(i%) - 1.5))
         Picture1.Line (56.7 * (H2(i%) - 1.4), 56.7 * (V1(i%) - 0.5))-(56.7 * (H2(i%) - 1.4), 56.7 * (V2(i%) - 1.5))
         '
         Picture1.CurrentY = 56.7 * V1(i%)
         Picture1.CurrentX = 56.7 * H1(i%)
         Picture1.Font = FUENTE$(i%)
         Picture1.FontSize = TAFUEN(i%)
         Picture1.FontBold = NEGRI(i%)
         Picture1.FontItalic = ITALI(i%)
         Picture1.FontUnderline = SUBRA(i%)
         Picture1.FontStrikethru = False
            If INVIS(i%) = 1 Then Picture1.FontStrikethru = True
         Label17.Font = Picture1.Font
         Label17.FontSize = Picture1.FontSize
         Label17.FontBold = Picture1.FontBold
         Label17.FontItalic = Picture1.FontItalic
         '
         TPXX$ = "": Label17 = ""
         For PINI% = 1 To Len(TEXTO(i%))
            If Label17.Width < 56.7 * (H2(i%) - H1(i%)) Then
              TPXX$ = TPXX$ + Mid$(TEXTO(i%), PINI%, 1)
              Label17 = TPXX$ + Mid$(TEXTO(i%), PINI% + 1, 1)
            End If
         Next PINI%
         Picture1.ForeColor = RGB(0, 0, 255)
            If ENFOC% > 0 Or i% = XVALO(Label14) Then Picture1.ForeColor = FOCOL ' RGB(0, 150, 150)
         Picture1.Print TPXX$    'TEXTO(I%)
         Picture1.ForeColor = RGB(0, 0, 0)
         '
         Label17 = TEXTO(i%)
      Case "LINE"
         Picture1.DrawStyle = 0
         Picture1.DrawWidth = XVALO(TAFUEN$(i%))
         Picture1.Line (56.7 * H1(i%), 56.7 * V1(i%))-(56.7 * H2(i%), 56.7 * V2(i%))
      Case "BOX"
         Picture1.DrawStyle = 0
         Picture1.DrawWidth = XVALO(TAFUEN$(i%))
         Picture1.Line (56.7 * H1(i%), 56.7 * V1(i%))-(56.7 * H2(i%), 56.7 * V2(i%)), , B
      Case "GRAPH"
         Picture1.DrawStyle = 2
         Picture1.DrawWidth = 1
         Picture1.Line (56.7 * H1(i%), 56.7 * V1(i%))-(56.7 * H2(i%), 56.7 * V2(i%)), , B
         '
         If ESTAPRETADO% = 0 Then
            If TEXTO(i%) <> "" Then
              ARIMAG$ = LUDAT$ + TEXTO(i%)
              If EXISTE%(ARIMAG$) < 1 Then ARIMAG$ = App.Path + "\" + TEXTO(i%)
              If EXISTE%(ARIMAG$) < 1 Then ARIMAG$ = TEXTO(i%)
              If EXISTE%(ARIMAG$) > 0 Then
                 Call DIBUIMAG(DSGFORM.Picture1, ARIMAG$, H1(i%), H2(i%), V1(i%), V2(i%))
              End If
            End If
         End If
      Case "GRID"
         Picture1.DrawStyle = 4
         Picture1.DrawWidth = 1
         Picture1.Line (56.7 * H1(i%), 56.7 * V1(i%))-(56.7 * H2(i%), 56.7 * V2(i%)), , B
   End Select
   Call FRESHALL
   On Error GoTo 0
   '
End Sub
'
Sub CREAOBJETO(TIPOBJ$, HX1, VX1, TTXX$, MULTLIN, FORMATEXX$, FREPLAXX$, TASELXX$, CAMPOSEXX$, CONDISEXX$, OBLIGAXX%, LOCKFIXX$)
     '
     IOBJ% = CANOBJ% + 1
     TIOBJ$(IOBJ%) = TIPOBJ$
       MAXCAI% = 1
25     ONAM$ = TIPOBJ$ + "-" + TRIM$(Str$(MAXCAI%))
       For KKI% = 1 To CANOBJ%
          If ONAME$(KKI%) = ONAM$ Then
             MAXCAI% = MAXCAI% + 1
             GoTo 25
          End If
       Next KKI%
     ONAME$(IOBJ%) = ONAM$
     '
     If XVALO(UTAFUE) < 1 Then UTAFUE = Text8
     If XVALO(UTAFUE) < 1 Then UTAFUE = 8
     Label17.Font = Combo1.TEXT
     Label17.FontSize = UTAFUE
     Label17.FontBold = UNEGRI
     Label17.FontItalic = UITALI
     Label17.FontUnderline = USUBRA
     If TTXX$ <> "" Then
         Label17 = RTrim(TTXX) + "*M"
       Else
         Label17 = TMASCARA$(FORMATEXX$)
     End If
     '
     H1(IOBJ%) = HX1
     H2(IOBJ%) = HX1 + Label17.Width / 56.7
     V1(IOBJ%) = VX1
     V2(IOBJ%) = VX1 + Label17.Height / 56.7
     CONMARCA%(IOBJ%) = 0
     TEXTO(IOBJ%) = TTXX$
     ARDOCLI$(IOBJ%) = Text15.TEXT
     FUENTE$(IOBJ%) = Combo1.TEXT
        If XVALO(UTAFUE) < 1 Then UTAFUE = Text8
     TAFUEN$(IOBJ%) = UTAFUE
     NEGRI(IOBJ%) = UNEGRI
     ITALI(IOBJ%) = UITALI
     SUBRA(IOBJ%) = USUBRA
     INVIS(IOBJ%) = 0
     ENCAS(IOBJ%) = 0
     MULIN(IOBJ%) = MULTLIN
     FORPA$(IOBJ%) = FORMATEXX$
     FOREPLA$(IOBJ%) = FREPLAXX$
     TASEL$(IOBJ%) = TASELXX$
     CASEL$(IOBJ%) = CAMPOSEXX$
     CONDSEL$(IOBJ%) = CONDISEXX$
     OBLIGA%(IOBJ%) = OBLIGAXX%
     LOCKFI$(IOBJ%) = LOCKFIXX$
        If LOCKFIXX$ = "@IDENT" Then LOCKFI$(IOBJ%) = ONAM$
     VAMINI$(IOBJ%) = ""
     VAMAXI$(IOBJ%) = ""
     TABVALI$(IOBJ%) = ""
     '
     CANOBJ% = CANOBJ% + 1
     '
End Sub
'
Private Sub Command11_Click()
   '
   Command11.Enabled = False
   OPPX% = ALTERNA%("Descarga Documentos Viejos\Control y Recuperación de Datos\Control de Documentos")
   If OPPX% = 1 Then
        Call DESCADOC
      ElseIf OPPX% = 2 Then
        Call CHECKDOC
      ElseIf OPPX% = 3 Then
        Call VERIDOC
   End If
   '
99 Screen.MousePointer = 1
   Command11.Enabled = True
   MARCOBARRA.Visible = False
   End Sub

Sub DESCADOC()
   '
   If DERACCE%("DESCADOC", "Descarga de Base de Datos de Documentos") >= 2 Then
     If COMALTER%("Realmente Quiere Descargar la Base de Datos\de Documentos Impresos ?\\No, Conservar\Si, Descargar.") = 2 Then
       MARCOBARRA.Visible = True
       Call BLOQARCH("PDOCSPL")
       On Error Resume Next
       Name LUDAT$ + "PDOCSPL.DAT" As LUDAT$ + "PDOCSPL0.DAT"
       If Err Then
         On Error GoTo 0
         Call MENSERR(24, "Proceso Imposible.\   \Cierre FLEXOFT en TODAS las\Estaciones de Trabajo, y\Repita la Petición.")
         Call CIERRARCH("MENSAJES")
         GoTo 99
       End If
       On Error GoTo 0
       '
       Screen.MousePointer = 11
       Call COPYSTRU("PDOCSPL", "PDOCSPL0")
       '
       FIN& = ULTREG&("PDOCLST")
       For U& = FIN& To 1 Step -1
         XX$ = REGLEIDO$("PDOCLST", U&)
         UDOC& = CVS(Left$(XX$, 4))
         If UDOC& > 0 Then GoTo 20
       Next U&
       UDOC& = 1
       '
20     FINDESCA& = UDOC& / 2   ' descarga la mitad del archivo
       Dim TD0 As String * 512
       NX0% = FILEOPEN%(LUDAT$ + "PDOCSPL.DT0", 0, 512)
       URE0& = LOF(NX0%) / 512
       FIN& = ULTREG&("PDOCSPL0")
       NX1% = FILEOPEN%(LUDAT$ + "PDOCSPL.DAT", 0, 512)
       URE1& = 0
       '
       For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
         XX$ = REGLEIDO$("PDOCSPL0", U&)
         TD0$ = XX$
         If CVS(Left$(XX$, 4)) <= FINDESCA& Then
              URE0& = URE0& + 1
              Put #NX0%, URE0&, TD0$
            Else
              URE1& = URE1& + 1
              Put #NX1%, URE1&, TD0$
         End If
       Next U&
       '
       Call CIERRARCH("PDOCSPL0")
       Close #NX0%
       Close #NX1%
       '
       Screen.MousePointer = 1
       On Error Resume Next
       Kill LUDAT$ + "PDOCSPL0.DAT"
       If Err Then
           On Error GoTo 0
           Call COMUNI("No se Pudo Completar el Proceso - Consulte.")
         Else
           '\\\OT-07-0111-RG-15-03-07/// 'MENSAJE QUE MUESTRA CUANDO FINALIZA EL PROCESO
           On Error GoTo 0
           Call COMUNI("Fin de Proceso de Descarga:\   \Operación Exitosa\")
           '\\\OT-07-0111-RG-FIN///
       End If
       On Error GoTo 0
       '
     End If
   End If
   '
99 End Sub

Sub CHECKDOC()
   '
   If DERACCE%("CHECKDOC", "Control y Recuperación de B.D. Documentos") >= 2 Then
     If COMALTER%("Realmente Quiere Recuperar la Base de Datos\de Documentos Impresos ?\\No, Cancelar\Si, Recuperar.") = 2 Then
       MARCOBARRA.Visible = True
       OPX% = COMALTER%("Para este Proceso debe Cerrarse FLEXOFT\en TODAS las Estaciones de Trabajo.\   \Pulse 'Continuar' para Iniciar el Proceso\\Continuar\Cancelar")
       If OPX% <> 1 Then Exit Sub
       '
       EPAC$ = UCase$(EMPREAC$)
       If InStr(EPAC$, "DOSIVAC") > 0 Then
          Call CHECKDOSI
          GoTo 99
       End If
       '
       TTXX$ = "Para el Proceso de Recuperación,\Tome Archivo de Backup 'PDOCSPL.DAT'\y Re-Nómbrelo como 'BKSPOOL.DAT'.\   \Luego Coloque este Último en la \Carpeta de Datos '" + REPLACAR$(LUDAT$, "\", "/")
       If COMALTER%(TTXX$ + "\\Cancelar\Recuperar") = 2 Then
          Screen.MousePointer = 11
          Call COPYSTRU("PDOCSPL", "BKSPOOL")
          FIN& = ULTREG&("BKSPOOL")
          For U& = 2100000 To FIN&
             Call TRACE(20, U&, FIN&)
             X$ = REGLEIDO$("PDOCSPL", U&)
             Y$ = REGLEIDO$("BKSPOOL", U&)
             If Y$ <> X$ Then
                If COMALTER%("Recupera Registro Numero " & U& & "\\No, Continuar\Si,Recuperar") = 2 Then
                   Call GRAREG("PDOCSPL", Y$, U&)
                End If
             End If
          Next U&
          Call CIERRARCH("*.*")
          Screen.MousePointer = 1
          Call COMUNI("Fin de Recuperación")
       End If
       '
     End If
   End If
   '
99 List2.Visible = False
   Frame11.Visible = False
   Call LIBARCH("*.*")
   '
End Sub

Sub VERIDOC()
    Dim ISDOC%(), DESRE&(), HASRE&()
    '
    MARCOBARRA.Visible = True
    FIN& = ULTREG&("PDOCLST")
    For U& = FIN& To 1 Step -1
       XX$ = REGLEIDO$("PDOCLST", U&)
       UDOC& = CVS(Left$(XX$, 4))
       If UDOC& > 0 Then GoTo 20
    Next U&
    UDOC& = 1
    '
20  UDOC& = UDOC& + 256
    ReDim ISDOC%(UDOC&), DESRE&(UDOC&), HASRE&(UDOC&)
    '
    PRIDOC& = 0: ULTDOC& = 0
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       XX$ = REGLEIDO$("PDOCLST", U&)
       NDOC& = CVS(Left$(XX$, 4))
' VERIFICAR AQUI ORDEN CORRELATIVO
       If NDOC& > 0 Then
         If PRIDOC& = 0 Then PRIDOC& = NDOC&
         If NDOC& <= UDOC& Then
           ISDOC%(NDOC&) = 1
           If NDOC& > ULTDOC& Then ULTDOC& = NDOC&
           If NDOC& <= DOCANT& Then
              YY$ = TRIM$(Mid$(XX$, 5, 60))
              MsgBox "Registro " & U& & " - " & YY$ & " - Desorden en Indice."
           End If
           DOCANT& = NDOC&
           GoTo 29
         End If
       End If
29   Next U&
     '
     'NX0% = FILEOPEN%(LUDAT$ + "PDOCSPL.DT0", 0, 512)
     'URE0& = LOF(NX0%) / 512
     'FIN& = ULTREG&("PDOCSPL0")
     'NX1% = FILEOPEN%(LUDAT$ + "PDOCSPL.DAT", 0, 512)
     'URE1& = 0
     '
     FIN& = ULTREG&("PDOCSPL")
     DOCANT& = 0
     For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       XX$ = REGLEIDO$("PDOCSPL", U&)
       NDOC& = CVS(Left$(XX$, 4))
       If NDOC& < DOCANT& Then
           YY$ = TRIM$(Mid$(XX$, 5, 60))
           MsgBox "Registro " & U& & " de PDOCSPL - Desorden en Detalle. (" & DOCANT& & " / " & NDOC& & ")."
       End If
       DOCANT& = NDOC&
       If NDOC& > 0 Then
         If NDOC& <= UDOC& Then
           If DESRE&(NDOC&) < 1 Then DESRE&(NDOC&) = U&
           If HASRE&(NDOC&) < U& Then HASRE&(NDOC&) = U&
         End If
       End If
     Next U&
     '
     For U& = 1 To UDOC&
        If ISDOC%(U&) = 1 Then
           If DESRE&(U&) < 1 Or HASRE&(U&) < DESRE&(U&) Then
              YY$ = TRIM$(Mid$(REGLEIDO$("PDOCLST", U&), 5, 60))
              MsgBox "Registro " & U& & " - " & YY$ & " - Falta Detalle "
           End If
           If DESRE&(U&) > 0 Then
             If DESRE&(U&) <= HASRE&(U& - 1) Then
               MsgBox "Punteros Cruzados"
             End If
           End If
        End If
     Next U&
     '
End Sub

Sub CHECKDOSI()
   '
   Call BLOQARCH("PDOCSPL")
   FIDATI# = FIDATIM#(LUDAT$ + "PDOCSPL.DAT")
   '
   Screen.MousePointer = 11
   Call COPYSTRU("PDOCSPL", "PDOCSPL0")
   On Error Resume Next
   Kill LUDAT$ + "PDOCSPL0.DAT"
   On Error GoTo 0
   '
   If EXISTE%(LUDAT$ + "PDOCSPL0.DAT") > 0 Then
      Call MENSERR(24, "Proceso Imposible. No fue Posible Borrar\Archivo de Transferencia 'PDOCSPL0.DAT'")
      GoTo 99
   End If
   '
   List2.Visible = True
   Frame11.Visible = True
   '
   YY$ = REGLEIDO$("PDOCLST", 1)
   MINI = CVS(Left$(YY$, 4))
   URELI& = ULTREG&("PDOCLST")
   YY$ = REGLEIDO$("PDOCLST", URELI&)
   MAXI = CVS(Left$(YY$, 4))
   '
   FIN& = ULTREG&("PDOCSPL")
   JJ& = 0: NDOCA& = 0: CAREDEL& = 0
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      VALOPRO& = 1 + (U& / FIN&) * (MAXI - MINI + 1)
      XX$ = REGLEIDO$("PDOCSPL", U&)
      VALEIDO = CVS(Left$(XX$, 4))
      If (Abs(VALOPRO& - VALEIDO)) < 10 Or ((Abs(VALOPRO& - VALEIDO)) / VALOPRO&) < 0.1 Then
         JJ& = JJ& + 1
         Call GRAREG("PDOCSPL0", XX$, JJ&)
         NDOC1& = VALEIDO
         If NDOC1& <> NDOCA& Then
           YY$ = FILTERGETRECORD$("PDOCLST", 1, MKS$(NDOC1&))
           ZZ$ = FORMATNUM$(NDOC1&, "F7.0") + "  " + Mid$(YY$, 5, 60)
           If List2.ListCount > 30000 Then
              List2.RemoveItem 0
           End If
           List2.AddItem ZZ$
           List2.Visible = True
           List2.ListIndex = List2.ListCount - 1
           List2.Refresh
           NDOCA& = NDOC1&
         End If
         CAREDEL& = 0
         GoTo 29
      End If
      If CAREDEL& = 0 Then
        List2.AddItem "   **********************  Documentos Perdidos  *************************"
      End If
      CAREDEL& = CAREDEL& + 1
      '
29  Next U&
   '
   Call CIERRARCH("PDOCSPL")
   Call CIERRARCH("PDOCSPL0")
   '
   If ULTREG&("PDOCSPL") <> FIN& Then
      Call MENSERR(24, "Imposible Continuar. Usuarios Conectados\Activos Durante el Proceso.")
      GoTo 99
   End If
   '
   Close
   Reset
   Call BLOQARCH("PDOCSPL")
   On Error Resume Next
   Name LUDAT$ + "PDOCSPL.DAT" As LUDAT$ + "PDOCSPL.BAD"
   If Err Then
      On Error GoTo 0
      Call MENSERR(24, "Proceso Interrumpido. No fue Posible\Re-Nombrar 'PDOCSPL.DAT' como 'PDOCSPL.BAD'.\   \Repita el Proceso.")
      GoTo 99
   End If
   On Error GoTo 0
   '
   On Error Resume Next
   Name LUDAT$ + "PDOCSPL0.DAT" As LUDAT$ + "PDOCSPL.DAT"
   If Err Then
      On Error GoTo 0
      Call MENSERR(24, "Proceso Interrumpido. No fue Posible\Re-Nombrar 'PDOCSPL0.DAT' como 'PDOCSPL.DAT'.\   \Consulte Urgente a Soporte FLEXOFT.")
      GoTo 99
   End If
   On Error GoTo 0
   '
   Call COMUNI("Fin de Proceso de Recuperación:\   \Operación Exitosa\")
99 On Error GoTo 0
   '
End Sub

Private Sub Command12_Click()
   '
   FRESHLIST% = (-1)         ' PARA MANTENER LA LISTA DE SELECCION
   Call Command5_Click       ' PARA GUARDAR ANTES DE EJECUTAR
   FRESHLIST% = 0
   '
   Call STDFORM(Text1, "")   ' "SSTK-STD", CODINT%("BI-1809A"))
   Call Text1_Change         ' PARA REFRESCAR DATOS
   Exit Sub
   
   CODIFOR$ = Text1
   CONDI$ = "HOPROFAB = '" & CODIFOR$ & "'"
   CODIPRODU% = XVALO(VALOBADA("SECOPER", "CodIConj", CONDI$))
   If CODIPRODU% > 0 Then
      NUMEOPERA% = XVALO(VALOBADA("SECOPER", "NumeOper", CONDI$))
      If NUMEOPERA% > 0 Then
         CONDI$ = "Cod_Inte = " & CODIPRODU%
         NORFA$ = TRIM$(VALOBADA("ORDEFABR", "MAX(IDSUBOR)", CONDI$))
         If NORFA$ <> "" Then
            STDSCREEN.Caption = Text2
            Timer1.Enabled = False
            Call STDFORM(CODIFOR$, NORFA$ & ";" & NUMEOPERA%)
            Timer1.Enabled = True
            GoTo 99
         End If
         Call MENSERR(24, "Testing Imposible - No hay O/F para " + CODEXT$(CODIPRODU%) + " - OP." & NUMEOPERA%)
         GoTo 99
      End If
   End If
   Call MENSERR(24, "Testing Imposible - Formulario no Vinculado a Secuencia Operación Producto")
   '
99 Call Text1_Change
End Sub

Sub CARGATEXTO(i%, TOOLTI$)
   '
   TOOLTI$ = ""
   If TIOBJ$(i%) = "PARAM" Or TIOBJ$(i%) = "TEXT" Then
      FOREP$ = LTrim$(RTrim$(FOREPLA$(i%)))
      If FOREP$ <> "" Then
         TOOLTI$ = FOREP$
         ' TRABAJA CON TEXTO REEMPLAZADO
         FOREP$ = TEXREPLA$(FOREP$)
         TEXTO(i%) = FOREP$     ' ASIGNA COMO TEXTO EL PARAMETRO
         Exit Sub
      End If
      '
      If TASEL$(i%) <> "" Then
         ' TRABAJA CON VALOBADA
         CONDIX$ = TEXREPLA$(CONDSEL$(i%), 1)
         VAREPLA = VALOBADA(TASEL$(i%), CASEL$(i%), CONDIX$)
         TOOLTI$ = "SELECT " + CASEL$(i%) + " FROM " + TASEL$(i%) + " WHERE " + CONDIX$
         If VarType(VAREPLA) = 7 Then VAREPLA = CVINT(VAREPLA)         ' CASO FECHA
         TEXTO(i%) = VAREPLA
      End If
   End If
   '
End Sub
'

Function TEXREPLA$(TEXORI$, Optional ENTRECOMI%)
     '
     TREP$ = TEXORI$
10   PPXX1% = InStr(TREP$, "{")
     If PPXX1% > 0 Then      ' BUSCA ENCIERRO ENTRE LLAVES
        PPXX2% = InStr(TREP$, "}"): If PPXX2% <= PPXX1% Then PPXX2% = 1 + Len(TREP$)
        IZQUI$ = Left$(TREP$, PPXX1% - 1)
        DEREX$ = Mid$(TREP$, PPXX2% + 1)
        VAREPLA$ = Mid$(TREP, PPXX1% + 1, PPXX2% - PPXX1% - 1)
          If UCase$(Left(VAREPLA$, 7)) = "COMMPAR" Then VAREPLA$ = "COMPAR" + Mid$(VAREPLA$, 8)
          If UCase$(Left$(VAREPLA$, 6)) = "PARCOM" Then VAREPLA$ = "COMPAR" + Mid$(VAREPLA$, 7)
          If VAREPLA$ = "COMPAR" Then VAREPLA$ = "COMPAR-0"
        For JJ& = 1 To CANOBJ%
           If UCase$(ONAME$(JJ&)) = UCase$(VAREPLA$) Then
              VAREPLA = TEXTO(JJ&)
              GoTo 29
           End If
        Next JJ&
        '
29      COMILLA$ = "": If ENTRECOMI% = 1 Then COMILLA$ = "'"
        TREP$ = IZQUI$ + COMILLA$ + VAREPLA$ + COMILLA$ + DEREX$
        GoTo 10
     End If
     '
     TEXREPLA = TREP$
     '
End Function
'
Private Sub Command13_Click()
    DEVALOBA17.Label14 = Label14
    DEVALOBA17.Height = 5100
    DEVALOBA17.Caption = "Parametro Reemplazable"
    '
    If Label19.Caption = "TEXT" Then
       DEVALOBA17.Height = DEVALOBA17.Frame3.Top + DEVALOBA17.Frame3.Height + 100 + (DEVALOBA17.Height - DEVALOBA17.ScaleHeight)
       DEVALOBA17.Caption = "Caja de Texto - Entrada de Datos"
    End If
    '
    DEVALOBA17.Show 1
    MODO$ = "": Label19 = ""
End Sub

Private Sub Command14_Click()
    '
    Command14.Enabled = False
    '
10  TEBUS$ = InputBox$("Busqueda por Texto", "Introduzca Argumento de Búsqueda", TEBUS$)
    If TEBUS$ = "" Then GoTo 99
    '
    Screen.MousePointer = 11
    MARCOBARRA.Visible = True
    Call COPYSTRU("PDOCLST", "PDOCLIS")
    URE& = ULTREG&("PDOCSPL")
    For i& = 1 To ULTREG&("PDOCSPL")
      Call TRACE(20, i&, URE&)
      RELE$ = REGLEIDO$("PDOCSPL", i&)
      X$ = Mid$(RELE$, 5)
      
      If InStr(1, X$, TEBUS$) > 0 Then
         DOCUNU& = CVS(Left$(RELE$, 4))
         '
         RECORD$ = FILTERGETRECORD$("PDOCLST", 1, MKS$(DOCUNU&))
         '
         Z& = Z& + 1
         Call GRAREG("PDOCLIS", RECORD$, Z&)
         '
         Call CIERRARCH("PDOCLIS")
      End If
    Next i&
    Call SETULTREG("PDOCLIS", Z&)
    MARCOBARRA.Visible = False
    '
    If Z& < 1 Then
       Call MENSERR(24, "No se Encontraron Documentos\para el Argumento de Búsqueda.")
       GoTo 10
    End If
    '
25  Screen.MousePointer = 1
    Call SELECHO("PDOCLIS/Búsqueda en Documentos por Texto: '" + TEBUS$ + "'")
    DOCUNU& = Val(VALACT1$("PDOCLIS"))
    If DOCUNU& > 0 Then
 '      Call MUESTRA_DOCU(DOCUNU&)
       GoTo 25
    End If
    GoTo 10
    '
99  Command14.Enabled = True
    '
End Sub

Private Sub Command15_Click()
   If Text3.Visible = True Then    ' para pasar primero a modo presentacion
      Call Picture1_GotFocus
   End If
   '
   IOBJ% = XVALO(Label14): RETIOBJ% = IOBJ
   If IOBJ% <= CANOBJ Then
      For KKI% = IOBJ% + 1 To CANOBJ%
         TIOBJ$(KKI% - 1) = TIOBJ$(KKI%)
         ONAME$(KKI% - 1) = ONAME$(KKI%)
         H1(KKI% - 1) = H1(KKI%)
         H2(KKI% - 1) = H2(KKI%)
         V1(KKI% - 1) = V1(KKI%)
         V2(KKI% - 1) = V2(KKI%)
         MULIN(KKI% - 1) = MULIN(KKI%)
         CONMARCA%(KKI% - 1) = CONMARCA%(KKI%)
         TEXTO(KKI% - 1) = TEXTO(KKI%)
         ARDOCLI$(KKI% - 1) = ARDOCLI$(KKI%)
         FUENTE$(KKI% - 1) = FUENTE$(KKI%)
         TAFUEN$(KKI% - 1) = TAFUEN$(KKI%)
         NEGRI(KKI% - 1) = NEGRI(KKI%)
         ITALI(KKI% - 1) = ITALI(KKI%)
         SUBRA(KKI% - 1) = SUBRA(KKI%)
         INVIS(KKI% - 1) = INVIS(KKI%)
         ENCAS(KKI% - 1) = ENCAS(KKI%)
         FORPA$(KKI% - 1) = FORPA$(KKI%)
         FOREPLA$(KKI% - 1) = FOREPLA$(KKI%)
         TASEL$(KKI% - 1) = TASEL$(KKI%)
         CASEL$(KKI% - 1) = CASEL$(KKI%)
         CONDSEL$(KKI% - 1) = CONDSEL$(KKI%)
         OBLIGA%(KKI% - 1) = OBLIGA%(KKI%)
         LOCKFI$(KKI% - 1) = LOCKFI$(KKI%)
         VAMINI$(KKI% - 1) = VAMINI$(KKI%)
         VAMAXI$(KKI% - 1) = VAMAXI$(KKI%)
         TABVALI$(KKI% - 1) = TABVALI$(KKI%)
      Next KKI%
      CANOBJ% = CANOBJ% - 1
      Label14 = CANOBJ%
      Call REDIBUJA
      If RETIOBJ% < CANOBJ% Then Label14 = RETIOBJ%
   End If
   '
   ESTAPRETADO% = 0
   Call Picture1_GotFocus
   MODO$ = ""
   '
End Sub

Private Sub Command16_Click()
    '
    IOBJ% = XVALO(Label14)
    If IOBJ% <= CANOBJ% Then
       If Text3.Visible = True Then
          Call Picture1_GotFocus
       End If
       '
       If MULTISEL% < 1 Or DHOR < 1 Then DHOR = H2(IOBJ%) - H1(IOBJ%) + 2
       If MULTISEL% < 1 Or DVER < 1 Then DVER = V2(IOBJ%) - V1(IOBJ%)
       CANOBJ% = CANOBJ% + 1
       TIOBJ$(CANOBJ%) = TIOBJ$(IOBJ%)
       MAXCAI% = 1
25     ONAM$ = TIOBJ$(CANOBJ%) + "-" + TRIM$(Str$(MAXCAI%))
       For KKI% = 1 To CANOBJ% - 1
          If ONAME$(KKI%) = ONAM$ Then
             MAXCAI% = MAXCAI% + 1
             GoTo 25
          End If
       Next KKI%
       ONAME$(CANOBJ%) = ONAM$
       H1(CANOBJ%) = H1(IOBJ%) + DHOR  ' H2(IOBJ%) + 2    ' 0
       H2(CANOBJ%) = H2(IOBJ%) + DHOR  ' H1(CANOBJ%) + H2(IOBJ%) - H1(IOBJ%)
       V1(CANOBJ%) = V1(IOBJ%) + DVER  ' V2(IOBJ%)        ' 0
       V2(CANOBJ%) = V2(IOBJ%) + DVER  ' V1(CANOBJ%) + V2(IOBJ%) - V1(IOBJ%)
       MULIN(CANOBJ%) = MULIN(IOBJ%)
       TEXTO(CANOBJ%) = TEXTO(IOBJ%)
       FUENTE$(CANOBJ%) = FUENTE$(IOBJ%)
       TAFUEN$(CANOBJ%) = TAFUEN$(IOBJ%)
       NEGRI(CANOBJ%) = NEGRI(IOBJ%)
       INVIS(CANOBJ%) = INVIS(IOBJ%)
       ENCAS(CANOBJ%) = ENCAS(IOBJ%)
       CONMARCA%(CANOBJ%) = 0
       ITALI(CANOBJ%) = ITALI(IOBJ%)
       SUBRA(CANOBJ%) = SUBRA(IOBJ%)
       FORPA$(CANOBJ%) = FORPA$(IOBJ%)
       FOREPLA$(CANOBJ%) = FOREPLA$(IOBJ%)
       TASEL$(CANOBJ%) = TASEL$(IOBJ%)
       CASEL$(CANOBJ%) = CASEL$(IOBJ%)
       CONDSEL$(CANOBJ%) = CONDSEL$(IOBJ%)
       OBLIGA%(CANOBJ%) = OBLIGA%(IOBJ%)
       LOCKFI$(CANOBJ%) = LOCKFI$(IOBJ%)
       VAMINI$(CANOBJ%) = VAMINI$(IOBJ%)
       VAMAXI$(CANOBJ%) = VAMAXI$(IOBJ%)
       TABVALI$(CANOBJ%) = TABVALI$(IOBJ%)
       '
       Call DIBUOBJ(CANOBJ%)
       ESTAPRETADO% = 0
       '
    End If
    MODO$ = "": Label19 = ""
End Sub

Private Sub Command17_Click()
   Command17.Enabled = False
   MODO$ = "": Label19 = ""
   If Text3.Visible = True Then
      Call Picture1_GotFocus
   End If
   '
   IOBJ% = XVALO(Label14)
   If IOBJ% <= CANOBJ Then
     If IOBJ% > 1 Then
       KKI% = IOBJ%
       SWAP TIOBJ$(KKI% - 1), TIOBJ$(KKI%)
       SWAP ONAME$(KKI% - 1), ONAME$(KKI%)
       SWAP H1(KKI% - 1), H1(KKI%)
       SWAP H2(KKI% - 1), H2(KKI%)
       SWAP V1(KKI% - 1), V1(KKI%)
       SWAP V2(KKI% - 1), V2(KKI%)
       SWAP CONMARCA%(KKI% - 1), CONMARCA%(KKI%)
       SWAP TEXTO(KKI% - 1), TEXTO(KKI%)
       SWAP FUENTE$(KKI% - 1), FUENTE$(KKI%)
       SWAP TAFUEN$(KKI% - 1), TAFUEN$(KKI%)
       SWAP NEGRI(KKI% - 1), NEGRI(KKI%)
       SWAP ITALI(KKI% - 1), ITALI(KKI%)
       SWAP SUBRA(KKI% - 1), SUBRA(KKI%)
       SWAP INVIS(KKI% - 1), INVIS(KKI%)
       SWAP ENCAS(KKI% - 1), ENCAS(KKI%)
       SWAP FORPA$(KKI% - 1), FORPA$(KKI%)
       SWAP FOREPLA$(KKI% - 1), FOREPLA$(KKI%)
       SWAP TASEL$(KKI% - 1), TASEL$(KKI%)
       SWAP CASEL$(KKI% - 1), CASEL$(KKI%)
       SWAP CONDSEL$(KKI% - 1), CONDSEL$(KKI%)
       SWAP OBLIGA%(KKI% - 1), OBLIGA%(KKI%)
       SWAP LOCKFI$(KKI% - 1), LOCKFI$(KKI%)
       SWAP VAMINI$(KKI% - 1), VAMINI$(KKI%)
       SWAP VAMAXI$(KKI% - 1), VAMAXI$(KKI%)
       SWAP TABVALI$(KKI% - 1), TABVALI$(KKI%)
       '
       Label14 = IOBJ% - 1
       Call REDIBUJA
     End If
   End If
   MODO$ = "": Label19 = ""
   Command17.Enabled = True
End Sub

Private Sub Command18_Click()   ' boton para conectar imagenes
    IOBJ% = XVALO(Label14)
    Command18.Enabled = False
10  CommonDialog1.CancelError = True
      If TRIM$(Text10) <> "" Then
        On Error Resume Next
        'CommonDialog1.InitDir = PARENTFOLDER(TRIM$(Text10))
        CommonDialog1.FileName = TRIM$(Text10)
        On Error GoTo 0
      End If
      '
    CommonDialog1.DialogTitle = "Archivo Origen de Imagen"
         ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
    CommonDialog1.Filter = "Todos los archivos (*.*)|*.*|Mapas de Bit" & _
         "(*.BMP)|*.BMP|Archivos Graficos JPEG (*.JPG)|*.JPG"
         ' Especificar el filtro predeterminado
    CommonDialog1.FilterIndex = 2
    On Error GoTo 99
    CommonDialog1.ShowOpen
    On Error GoTo 0
    '
    DoEvents
    ARIMAG$ = UCase$(CommonDialog1.FileName)
    '
    If ARIMAG$ <> "" Then
       If Left$(ARIMAG$, Len(LUDAT$)) = UCase$(LUDAT$) Then ARIMAG$ = Mid$(ARIMAG$, 1 + Len(LUDAT$))
       If Left$(ARIMAG$, Len(App.Path + "\")) = UCase$(App.Path + "\") Then ARIMAG$ = Mid$(ARIMAG$, 1 + Len(App.Path + "\"))
       Text10 = ARIMAG$
       TEXTO(IOBJ%) = ARIMAG$
       Call DIBUOBJ(IOBJ%)
    End If
99  Command18.Enabled = True
End Sub

Private Sub Command19_Click()
    '
    JJ& = 0
    For II& = 1 To ULTREG&("INDIFRM")
      XX$ = REGLEIDO$("INDIFRM", II&)
      COFO$ = UCase$(TRIM$(Left$(XX$, 8)))
      If EXISTE%(LUDAT$ + COFO$ + ".SQR") > 0 Then GoTo 19
      '
      If EXISTE%(LUDAT$ + COFO$ + ".FRM") > 0 Then
        For KK& = 1 To JJ&
          If UCase$(TRIM$(Left$(REGLEIDO$("INDIFRM", KK&), 8))) = COFO$ Then GoTo 19
        Next KK&
        JJ& = JJ& + 1
        Call GRAREG("INDIFRM", XX$, JJ&)
      End If
19  Next II&
    '
    Call SETULTREG("INDIFRM", JJ&)
    Call FILESORT("INDIFRM", "", 1, 1)
    Call FRESHECHO("INDIFRM")
    '
    Call SELECHO("INDIFRM")
    CODFOR$ = VALACT1$("INDIFRM")
    If TRIM$(CODFOR$) = "" Then
        Text1.SetFocus
      Else
        Text1.TEXT = CODFOR$
        Text2.TEXT = ECOARCH$("INDIFRM", CODFOR$)
        If EXISTE%(LUDAT$ + CODFOR$ + ".SQR") < 1 Then
           Call CARGAFORM
           Call REDIBUJA
        End If
    End If
    '
End Sub

Private Sub Command2_Click()
    Call CIERRARCH("*.*")
    Close
    Call FINAL("")
End Sub

Private Sub Command20_Click()
    IOBJ% = XVALO(Label14)
    Command20.Enabled = False
    '
    On Error Resume Next
    DEFGRID17.Label14 = Label14
    If Err Then
       On Error GoTo 0
       GoTo 99
    End If
    On Error GoTo 0
    DEFGRID17.Show 1
10  RSEL$ = TRIM$(DEFGRID17.Label4)
    If Len(RSEL$) > 4 Then
       Text10 = RSEL$
       TEXTO(IOBJ%) = Text10
    End If
    MODO$ = "": Label19 = ""
    '
99  Command20.Enabled = True
End Sub

Private Sub Command21_Click()
    IOBJ% = XVALO(Label14)
    Command21.Enabled = False
10  CommonDialog1.CancelError = True
      If TRIM$(Text10) <> "" Then
        On Error Resume Next
        CommonDialog1.FileName = TRIM$(Text15)
        On Error GoTo 0
      End If
      '
    CommonDialog1.DialogTitle = "Archivo Destino Doble-Click"
         ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
    CommonDialog1.Filter = "Todos los archivos (*.*)|*.*|Mapas de Bit" & _
         "(*.BMP)|*.BMP|Archivos Graficos JPEG (*.JPG)|*.JPG"
         ' Especificar el filtro predeterminado
    CommonDialog1.FilterIndex = 2
    On Error GoTo 99
    CommonDialog1.ShowOpen
    On Error GoTo 0
    '
    DoEvents
    ARDESTI$ = UCase$(CommonDialog1.FileName)
    '
    If ARDESTI$ <> "" Then
       If Left$(ARDESTI$, Len(LUDAT$)) = UCase$(LUDAT$) Then ARDESTI$ = Mid$(ARDESTI$, 1 + Len(LUDAT$))
       If Left$(ARDESTI$, Len(App.Path + "\")) = UCase$(App.Path + "\") Then ARDESTI$ = Mid$(ARDESTI$, 1 + Len(App.Path + "\"))
       Text15 = ARDESTI$
       ARDOCLI$(IOBJ%) = ARDESTI$
    End If
99  Command21.Enabled = True
End Sub

Private Sub Command22_Click()
  PHORI = (-1)
  For i% = 1 To CANOBJ%
     If ISEL%(i%) = 1 Then
        Call BORRAOBJ(i%)
        If PHORI < 0 Then PHORI = H1(i%)
        DIFEX = PHORI - H1(i%)
        H1(i%) = H1(i%) + DIFEX
        H2(i%) = H2(i%) + DIFEX
        Call DIBUOBJ(i%, 2)
        Picture1.SetFocus
     End If
  Next i%
End Sub

Private Sub Command23_Click()
  PVERT = (-1)
  For i% = 1 To CANOBJ%
     If ISEL%(i%) = 1 Then
        Call BORRAOBJ(i%)
        If PVERT < 0 Then PVERT = V1(i%)
        DIFEX = PVERT - V1(i%)
        V1(i%) = V1(i%) + DIFEX
        V2(i%) = V2(i%) + DIFEX
        Call DIBUOBJ(i%, 2)
        Picture1.SetFocus
     End If
  Next i%
End Sub

Private Sub Command24_Click()
  CANOBJA% = CANOBJ%
  For i% = 1 To CANOBJA%
     If ISEL%(i%) = 1 Then
        ISEL%(i%) = 0
        Call DIBUOBJ(i%)
        Label14 = TRIM$(Str$(i%))
        Call Command16_Click
        ISEL%(CANOBJ%) = 1
        Call DIBUOBJ%(CANOBJ%, 2)
        Picture1.SetFocus
     End If
  Next i%
End Sub

Private Sub Command25_Click()
    If MULTISEL% < 1 Then Exit Sub
    KKJ% = 0
    For KKI% = 1 To CANOBJ%
       If ISEL%(KKI%) < 1 Then
           KKJ% = KKJ% + 1
           TIOBJ$(KKJ%) = TIOBJ$(KKI%)
           ONAME$(KKJ%) = ONAME$(KKI%)
           H1(KKJ%) = H1(KKI%)
           H2(KKJ%) = H2(KKI%)
           V1(KKJ%) = V1(KKI%)
           V2(KKJ%) = V2(KKI%)
           MULIN(KKJ%) = MULIN(KKI%)
           CONMARCA%(KKJ%) = CONMARCA%(KKI%)
           TEXTO(KKJ%) = TEXTO(KKI%)
           ARDOCLI$(KKJ%) = ARDOCLI$(KKI%)
           FUENTE$(KKJ%) = FUENTE$(KKI%)
           TAFUEN$(KKJ%) = TAFUEN$(KKI%)
           NEGRI(KKJ%) = NEGRI(KKI%)
           ITALI(KKJ%) = ITALI(KKI%)
           SUBRA(KKJ%) = SUBRA(KKI%)
           INVIS(KKJ%) = INVIS(KKI%)
           ENCAS(KKJ%) = ENCAS(KKI%)
           FORPA$(KKJ%) = FORPA$(KKI%)
           FOREPLA$(KKJ%) = FOREPLA$(KKI%)
           TASEL$(KKJ%) = TASEL$(KKI%)
           CASEL$(KKJ%) = CASEL$(KKI%)
           CONDSEL$(KKJ%) = CONDSEL$(KKI%)
           OBLIGA%(KKJ%) = OBLIGA%(KKI%)
           LOCKFI$(KKJ%) = LOCKFI$(KKI%)
           VAMINI$(KKJ%) = VAMINI$(KKI%)
           VAMAXI$(KKJ%) = VAMAXI$(KKI%)
           TABVALI$(KKJ%) = TABVALI$(KKI)
           '
        End If
    Next KKI%
    CANOBJ% = KKJ%
    Label14 = CANOBJ%
    Call REDIBUJA
    '
    ESTAPRETADO% = 0
    Call Picture1_GotFocus
    MODO$ = ""
    '
End Sub

Private Sub Command26_Click()
     '
     Command26.Enabled = False
     '
     'SE VALIDA  EL CODIGO DEL FORMULARIO
     CODFOR$ = TRIM$(Text1)
     If TRIM$(CODFOR$) = "" Then
       Call MENSERR(24, "Falta Seleccionar Formulario")
       Text1.SetFocus
       GoTo 99
     End If

     'LLAMA A LA GRABACION DEL FRM PASANDOLE PARAMETRO DE LUGAR DE GRABACION LUCOM$
     'SIN VALIDAR LA INFORMACION
     If SOLOLEC% > 0 Then
        Call MENSERR(24, "Imposible Guardar. HECP '" + CODFOR$ + " Abierta en Modo de Solo Lectura")
        GoTo 99
     End If
     Call GRABAFRM(LUCOM$)
     '
     'VALIDA SI EXISTE EL ARCHIVO PARA EL ENVIO POR MAIL
     CODFOR$ = TRIM$(Text1)
     If EXISTE%(LUCOM$ + CODFOR$ + ".SQR") < 1 Then
        Call MENSERR(24, "No Se Encontró El Archivo: " & CODFOR$ & ".SQR" & " En la Ruta: " & REPLACAR(LUCOM$, "\", "/"))
         GoTo 99
     End If
     '
     'DATOS PARA MAIL
     ARCHIX$ = LUCOM$ + CODFOR$ + ".SQR"
     MAI_DESTI$ = "asistec@flexoft.com.ar"
     MAI_DIREL$ = MAI_DESTI$
     MAI_ASUNT$ = "Envio de Archivo " & CODFOR$
     MAI_TEXTO$ = "Envio de Archivo " & CODFOR$
     MAI_ADJUN$(0) = ARCHIX$
     '
     'ENVIO DE MAIL
     Call GENERAMAIL
     '
99   Command26.Enabled = True
End Sub

Private Sub Command27_Click()
    Command27.Enabled = False
    Screen.MousePointer = 11
    Call CARINDIFORM
    Screen.MousePointer = 1
    Command27.Enabled = True
End Sub

Private Sub Command3_Click()
    Call ALTANUEFORM
End Sub

Private Sub Command4_Click()
   '
   Printer.Orientation = 1
   If Left$(FOR_ORIE$, 3) = "HOR" Then
      On Error Resume Next
      Printer.Orientation = 2
      On Error GoTo 0
   End If
   '
   For i% = 1 To CANOBJ%
      Call PRINOBJ(i%)
   Next i%
   Printer.EndDoc
   '
End Sub

Sub PRINOBJ(i%)
   '
   MODOX$ = TIOBJ$(i%)
   On Error Resume Next
   Select Case MODOX$
      Case "LABEL"
         Printer.CurrentY = 56.7 * (V1(i%) + FOR_MGSU)
         Printer.CurrentX = 56.7 * (H1(i%) + FOR_MGIZ)
         Printer.Font = FUENTE$(i%)
         Printer.FontSize = TAFUEN(i%)
         Printer.FontBold = NEGRI(i%)
         Printer.FontItalic = ITALI(i%)
         Printer.FontUnderline = SUBRA(i%)
         If INVIS(i%) <> 1 Then Printer.Print TEXTO(i%)
      Case "PARAM"
         If ENCAS(i%) = 1 Then
            Printer.ScaleMode = 1
            Printer.ForeColor = 0
            Printer.DrawMode = 13
            Printer.DrawWidth = 1
            Printer.DrawStyle = 0
            Printer.FillStyle = 1
            If INVIS(i%) <> 1 Then Printer.Line (56.7 * (H1(i%) + FOR_MGIZ), 56.7 * (V1(i%) + FOR_MGSU))-(56.7 * (H2(i%) + FOR_MGIZ), 56.7 * (V2(i%) + FOR_MGSU)), QBColor(0), B
         End If
         Exit Sub     ' NO SE IMPRIMEN PARAMETROS
         Printer.CurrentY = 56.7 * (V1(i%) + FOR_MGSU)
         Printer.CurrentX = 56.7 * (H1(i%) + FOR_MGIZ)
         Printer.Font = FUENTE$(i%)
         Printer.FontSize = TAFUEN(i%)
         Printer.FontBold = NEGRI(i%)
         Printer.FontItalic = ITALI(i%)
         Printer.FontUnderline = SUBRA(i%)
         If INVIS(i%) <> 1 Then Printer.Print FORMATNUM$(TEXTO(i%), FORPA$(i%))
      Case "TEXT"
         If ENCAS(i%) = 1 Then
            Printer.ScaleMode = 1
            Printer.ForeColor = 0
            Printer.DrawMode = 13
            Printer.DrawWidth = 1
            Printer.DrawStyle = 0
            Printer.FillStyle = 1
            If INVIS(i%) <> 1 Then Printer.Line (56.7 * (H1(i%) + FOR_MGIZ), 56.7 * (V1(i%) + FOR_MGSU))-(56.7 * (H2(i%) + FOR_MGIZ), 56.7 * (V2(i%) + FOR_MGSU)), QBColor(0), B
         End If
         Exit Sub     ' NO SE IMPRIMEN PARAMETROS
         Printer.CurrentY = 56.7 * (V1(i%) + FOR_MGSU)
         Printer.CurrentX = 56.7 * (H1(i%) + FOR_MGIZ)
         Printer.Font = FUENTE$(i%)
         Printer.FontSize = TAFUEN(i%)
         Printer.FontBold = NEGRI(i%)
         Printer.FontItalic = ITALI(i%)
         Printer.FontUnderline = SUBRA(i%)
         If INVIS(i%) <> 1 Then Printer.Print FORMATNUM$(TEXTO(i%), FORPA$(i%))
      Case "LINE"
         Printer.ScaleMode = 1
         Printer.ForeColor = 0
         Printer.DrawMode = 13
         Printer.DrawWidth = XVALO(TAFUEN$(i%))
         Printer.DrawStyle = 0
         If INVIS(i%) <> 1 Then Printer.Line (56.7 * (H1(i%) + FOR_MGIZ), 56.7 * (V1(i%) + FOR_MGSU))-(56.7 * (H2(i%) + FOR_MGIZ), 56.7 * (V2(i%) + FOR_MGSU)), QBColor(0)
      Case "BOX"
         Printer.ScaleMode = 1
         Printer.ForeColor = 0
         Printer.DrawMode = 13
         Printer.DrawWidth = XVALO(TAFUEN$(i%))
         Printer.DrawStyle = 0
         Printer.FillStyle = 1
         If INVIS(i%) <> 1 Then Printer.Line (56.7 * (H1(i%) + FOR_MGIZ), 56.7 * (V1(i%) + FOR_MGSU))-(56.7 * (H2(i%) + FOR_MGIZ), 56.7 * (V2(i%) + FOR_MGSU)), QBColor(0), B
      Case "GRAPH"
         If INVIS(i%) <> 1 Then
            If TEXTO(i%) <> "" Then
              ARIMAG$ = LUDAT$ + TEXTO(i%)
              If EXISTE%(ARIMAG$) < 1 Then ARIMAG$ = App.Path + "\" + TEXTO(i%)
              If EXISTE%(ARIMAG$) < 1 Then ARIMAG$ = TEXTO(i%)
              If EXISTE%(ARIMAG$) > 0 Then
                 Call DIBUIMAG(Printer, ARIMAG$, H1(i%) + FOR_MGIZ, H2(i%) + FOR_MGIZ, V1(i%) + FOR_MGSU, V2(i%) + FOR_MGSU)
              End If
            End If
         End If
   End Select
   Call FRESHALL
   On Error GoTo 0
   '
End Sub

Private Sub Command5_Click()
    '
    Command5.Enabled = False
    '
    'SE VALIDA  EL CODIGO DEL FORMULARIO
    CODFOR$ = TRIM$(Text1)
    If TRIM$(CODFOR$) = "" Then
       Call MENSERR(24, "Falta Seleccionar Formulario")
       Text1.SetFocus
       GoTo 99
    End If
    '
    'VALIDA LA INFORMACION Y SI ESTA BIEN GRABA
    Call VALIDAFORM(OK%)
    If OK% > 0 Then
       Call GRABAFRM(LUDAT$)
    End If
    '
99  Command5.Enabled = True
    '
End Sub

Sub GRABAFRM(LUGRA$)
    '
    If SOLOLEC% > 0 Then
       Call MENSERR(24, "Imposible Guardar. HECP '" + CODFOR$ + "' Abierta en Modo de Solo Lectura")
       Exit Sub
    End If
    '
    IGRAF% = 0: TGRAF$ = ""

    N1% = FILEOPEN%(LUGRA$ + CODFOR$ + ".SQR", 2, 128)
    '
    ' DATOS DE ENCABEZADO
    TTXX$ = "CONFIG"
    TTXX$ = TTXX$ + Chr$(124) + "FOR_DESC = " + FOR_DESC$
    TTXX$ = TTXX$ + Chr$(124) + "FOR_ORIE = " + FOR_ORIE$
    TTXX$ = TTXX$ + Chr$(124) + "FOR_MGIZ = " + FORMATNUM$(FOR_MGIZ, "F10.1")
    TTXX$ = TTXX$ + Chr$(124) + "FOR_MGSU = " + FORMATNUM$(FOR_MGSU, "F10.1")
    TTXX$ = TTXX$ + Chr$(124) + "AUTO_PRN = " + AUTO_PRN$
    TTXX$ = TTXX$ + Chr$(124) + "PAGI_ANT = " + PAGI_ANT$
    TTXX$ = TTXX$ + Chr$(124) + "PAGI_SIG = " + PAGI_SIG$
       LOCK_DSG$ = "SI": If Image2.Visible = True Then LOCK_DSG$ = "NO"
    TTXX$ = TTXX$ + Chr$(124) + "LOCK_DSG = " + LOCK_DSG$
       FORM_OPX$ = TRIM$(Text16)
    TTXX$ = TTXX$ + Chr$(124) + "FORM_OPX = " + FORM_OPX$
       
    Print #N1%, TTXX$ + Chr$(124)
    '
    ' DATOS DE CADA OBJETO
    For i% = 1 To CANOBJ%
      TTXX$ = "TIPO = " + TIOBJ$(i%)
      TTXX$ = TTXX$ + Chr$(124) + "NAME = " + ONAME$(i%)
      TTXX$ = TTXX$ + Chr$(124) + " H1 = " + FORMATNUM$(H1(i%), "F10.1")
      TTXX$ = TTXX$ + Chr$(124) + " V1 = " + FORMATNUM$(V1(i%), "F10.1")
      TTXX$ = TTXX$ + Chr$(124) + " H2 = " + FORMATNUM$(H2(i%), "F10.1")
      TTXX$ = TTXX$ + Chr$(124) + " V2 = " + FORMATNUM$(V2(i%), "F10.1")
        '
        If TIOBJ$(i%) = "GRAPH" Then
           FNAME$ = LUDAT$ + TEXTO(i%)
           If EXISTE%(FNAME$) < 1 Then FNAME$ = App.Path + "\" + TEXTO(i%)
           If EXISTE%(FNAME$) < 1 Then FNAME$ = TEXTO(i%)
           '
           AGRAF$ = LOADFILE$(FNAME$)
           If AGRAF$ <> "" Then
              IGRAF% = IGRAF% + 1
              NGRAF$ = TRIM$(Str$(IGRAF%))
              While Len(NGRAF$) < 4: NGRAF$ = "0" + NGRAF$: Wend
              TGRAF$ = TGRAF$ + ">>>>>>>>GRAF" + NGRAF$ + AGRAF$ + "GRAF" + NGRAF$ + "<<<<<<<<"
              TEXTO(i%) = "GRAF" + NGRAF$
           End If
        End If
        '
        TX123$ = TEXTO(i%)
      If TRIM$(TX123$) <> "" Then TTXX$ = TTXX$ + Chr$(124) + " TEXT = " + TX123$
      If TRIM$(FUENTE$(i%)) <> "" Then TTXX$ = TTXX$ + Chr$(124) + " FONT = " + FUENTE$(i%)
      If TRIM$(TAFUEN$(i%)) <> "" Then TTXX$ = TTXX$ + Chr$(124) + " SIZE = " + TAFUEN$(i%)
      If TRIM$(FORPA$(i%)) <> "" Then TTXX$ = TTXX$ + Chr$(124) + " FORM = " + FORPA$(i%)
      If MULIN(i%) = 1 Then TTXX$ = TTXX$ + Chr$(124) + " MULIN = 1"
      If TRIM$(ARDOCLI$(i%)) <> "" Then TTXX$ = TTXX$ + Chr$(124) + " ADOC = " + ARDOCLI$(i%)
      '
      If NEGRI(i%) = 1 Then TTXX$ = TTXX$ + Chr$(124) + " NEGRI = 1"
      If ITALI(i%) = 1 Then TTXX$ = TTXX$ + Chr$(124) + " ITALI = 1"
      If SUBRA(i%) = 1 Then TTXX$ = TTXX$ + Chr$(124) + " SUBRA = 1"
      If INVIS(i%) = 1 Then TTXX$ = TTXX$ + Chr$(124) + " INVIS = 1"
      If ENCAS(i%) = 1 Then TTXX$ = TTXX$ + Chr$(124) + " ENCAS = 1"
      '
      If TRIM$(FOREPLA$(i%)) <> "" Then TTXX$ = TTXX$ + Chr$(124) + " FOREPLA = " + FOREPLA$(i%)
      If TRIM$(TASEL$(i%)) <> "" Then TTXX$ = TTXX$ + Chr$(124) + " TASEL = " + TASEL$(i%)
      If TRIM$(CASEL$(i%)) <> "" Then TTXX$ = TTXX$ + Chr$(124) + " CASEL = " + CASEL$(i%)
      If TRIM$(CONDSEL$(i%)) <> "" Then TTXX$ = TTXX$ + Chr$(124) + " CONDSEL = " + CONDSEL$(i%)
      If TRIM$(LOCKFI$(i%)) <> "" Then TTXX$ = TTXX$ + Chr$(124) + " LOCKFI = " + LOCKFI$(i%)
      If TRIM$(VAMINI$(i%)) <> "" Then TTXX$ = TTXX$ + Chr$(124) + " VAMINI = " + VAMINI$(i%)
      If TRIM$(VAMAXI$(i%)) <> "" Then TTXX$ = TTXX$ + Chr$(124) + " VAMAXI = " + VAMAXI$(i%)
      If TRIM$(TABVALI$(i%)) <> "" Then TTXX$ = TTXX$ + Chr$(124) + " TABVALI = " + TABVALI$(i%)
      If OBLIGA%(i%) > 0 Then TTXX$ = TTXX$ + Chr$(124) + " OBLIGA = 1"
      Print #N1%, TTXX$ + Chr$(124)
    Next i%
    '
    Close #N1%
    '
    If TGRAF$ <> "" Then
        TTZZ$ = LOADFILE$(LUGRA$ + CODFOR$ + ".SQR")
        TTZZ$ = TTZZ$ + TGRAF$
        Call SAVEFILE(LUGRA$ + CODFOR$ + ".SQR", TTZZ$)
    End If
    
    '
    Call Text1_Change                     ' PARA QUE VUELVA A CARGARLO

End Sub
Private Sub Command6_Click()
    Call HELPONLINE("PRIFORMS")
End Sub

Private Sub Command7_Click()
   '
   On Error Resume Next
   Printer.DrawWidth = 3
   Printer.DrawStyle = 0
   Printer.FillStyle = 1
   If Err Then
     Call MENSERR(24, "Imposible Imprimir.\Error de Comunicacion con la Impresora.")
     Resume
     On Error GoTo 0
     Exit Sub
   End If
   On Error GoTo 0
   '
5  Call SELECHO("PDOCLST")
   DOCUNU& = Val(VALACT1$("PDOCLST"))
   If DOCUNU& < 1 Then
      Exit Sub
   End If
   REDOCU$ = REGACT$("PDOCLST")
   '
   TPSP$ = ""
   Screen.MousePointer = 11
   '
   Dim TD0 As String * 512
   NX0% = FILEOPEN%(LUDAT$ + "PDOCSPL.DAT", 0, 512)
   Get #NX0%, 1, TD0$
   If CVS(Left$(TD0$, 4)) > DOCUNU& Then
     Close #NX0%
     NX0% = FILEOPEN%(LUDAT$ + "PDOCSPL.DT0", 0, 512)
   End If
   '
   URE& = Abs(LOF(NX0%) / 512)   ' PARA PREVENIR DESBORDAMIENTO DOSIVAC
   LI& = 0: LS& = URE&
6  E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
   L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
   Get #NX0%, L&, TD0$
   DCC& = CVS(Left$(TD0$, 4))
   If DCC& < DOCUNU& Then LI& = L&: GoTo 6
   If DCC& > DOCUNU& Then LS& = L&: GoTo 6
   GoTo 8
   '
7  Call MENSERR(24, "Documento no Registrado")
   GoTo 5
   '
8  TPSP$ = ""
   Screen.MousePointer = 11
   While L& > 1
     Get #NX0%, L&, TD0$
     DCC& = CVS(Left$(TD0$, 4))
     If DCC& < DOCUNU& Then GoTo 9
     L& = L& - 1
   Wend
   '
9  For KKL& = L& To URE&          'SE CAMBIO EL ARCHIVO DE LECTURA
     Get #NX0%, KKL&, TD0$
     TBUF$ = TD0$
     If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
         TPSP$ = TPSP$ + Mid$(TBUF$, 5)
       ElseIf CVS(Left$(TBUF$, 4)) > DOCUNU& Then
         GoTo 91
     End If
   Next KKL&
   '
91 Close #NX0%
   Screen.MousePointer = 11
   On Error Resume Next
   Printer.PrintQuality = (-1)
   MATRIXPUN% = 0
   If InStr(Printer.DeviceName, "810") > 0 Then MATRIXPUN% = 1
   If InStr(Printer.DeviceName, "286") > 0 Then MATRIXPUN% = 1
   If InStr(UCase$(Printer.DeviceName), "APEX") > 0 Then MATRIXPUN% = 1
   If InStr(UCase$(Printer.DeviceName), "FACTURA") > 0 Then MATRIXPUN% = 1
   If InStr(UCase$(Printer.DeviceName), "REMITO") > 0 Then MATRIXPUN% = 1
   'If InStr(Printer.DriverName, "EPSON9") > 0 Then MATRIXPUN% = 1
   FORMALTER.List1.Clear
   On Error GoTo 0
   '
   Printer.EndDoc        ' RESETEA IMPRESORA
   '
10 PPXX% = InStr(TPSP$, Chr$(255))
   If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
    '    Call PRINTCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
   End If
   '
   Printer.EndDoc
   Screen.MousePointer = 1
   GoTo 5
   '
End Sub
'
Private Sub Command8_Click()
   '
   For KKI% = 0 To 5
     EDIFORMU.Picture1(KKI%).Cls
     EDIFORMU.Picture1(KKI%).Top = 0
     EDIFORMU.Picture1(KKI%).Refresh
     EDIFORMU.Picture1(KKI%).Height = 7070
     EDIFORMU.Picture1(KKI%).Width = 11700
     EDIFORMU.Picture1(KKI%).DrawWidth = 3
     EDIFORMU.Picture1(KKI%).DrawStyle = 0
     EDIFORMU.Picture1(KKI%).FillStyle = 1
   Next KKI%
   ALTUPAGINA = 7070
   TOPEPAGINA = 0
   PAGINACTIVA% = 0
   CAPAGINAS% = 0
   '
5  Call SELECHO("PDOCLST")
   DOCUNU& = Val(VALACT1$("PDOCLST"))
   If DOCUNU& < 1 Then
      Exit Sub
   End If
   REDOCU$ = REGACT$("PDOCLST")
   ORIPRINTER$ = TRIM$(Mid$(REDOCU$, 129))
   '
   Dim TD0 As String * 512
   NX0% = FILEOPEN%(LUDAT$ + "PDOCSPL.DAT", 0, 512)
   Get #NX0%, 1, TD0$
   If CVS(Left$(TD0$, 4)) > DOCUNU& Then
     Close #NX0%
     NX0% = FILEOPEN%(LUDAT$ + "PDOCSPL.DT0", 0, 512)
   End If
   '
   URE& = Abs(LOF(NX0%) / 512)   ' PARA PREVENIR DESBORDAMIENTO DOSIVAC
   LI& = 0: LS& = URE&
6  E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
   L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
   '
   Get #NX0%, L&, TD0$
   DCC& = CVS(Left$(TD0$, 4))
   If DCC& < DOCUNU& Then LI& = L&: GoTo 6
   If DCC& > DOCUNU& Then LS& = L&: GoTo 6
   GoTo 8
   '
7  Call MENSERR(24, "Documento no Registrado")
   GoTo 5
   '
8  TPSP$ = ""
   Screen.MousePointer = 11
   While L& > 1
     Get #NX0%, L&, TD0$
     DCC& = CVS(Left$(TD0$, 4))
     If DCC& < DOCUNU& Then GoTo 9
     L& = L& - 1
   Wend
   '
9  CAPAGINAS% = 0
   For KKL& = L& To URE&
     Get #NX0%, KKL&, TD0$
     TBUF$ = TD0$
     If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
     If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
       TPSP$ = TPSP$ + Mid$(TBUF$, 5)
     End If
   Next KKL&
   '
95 SPOOLSTRING = TPSP$
   Close #NX0%
   '
10 PPXX% = InStr(TPSP$, Chr$(255))
   If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
   '     Call SHOWCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
   End If
   '
   Screen.MousePointer = 1
   PAGINACTIVA% = 0
   ANCHPAGINA = EDIFORMU.Picture1(PAGINACTIVA%).Width
   EDIFORMU.Picture1(PAGINACTIVA%).Refresh
   EDIFORMU.VScroll1.Min = 0
   EDIFORMU.VScroll1.Value = 0
   EDIFORMU.VScroll1.Max = 1
   If ALTUPAGINA > EDIFORMU.Frame1.Height Then
      EDIFORMU.VScroll1.Max = (ALTUPAGINA - EDIFORMU.Frame1.Height) / 100
   End If
   EDIFORMU.HScroll1.Min = 0
   EDIFORMU.HScroll1.Value = 0
   EDIFORMU.HScroll1.Max = 1
   If ANCHPAGINA > EDIFORMU.Frame1.Width Then
      EDIFORMU.HScroll1.Max = (ANCHPAGINA - EDIFORMU.Frame1.Width) / 100
   End If
   '
   EDIFORMU.Show
   '
End Sub

Private Sub Command99_Click()
   Screen.MousePointer = 11
   FIN& = ULTREG&("PDOCLST")
   For U& = 73740 To FIN&
     XX$ = REGLEIDO$("PDOCLST", U&)
     Call REPLA(XX$, MKS$(U&), 1, 4)
     If U& >= 73740 Then
       If U& <= 73742 Then
         Call REPLA(XX$, "Documento Numero " + TRIM$(Str$(U&)), 5, 124)
       End If
     End If
     Call GRAREG("PDOCLST", XX$, U&)
   Next U&
   Call CIERRARCH("PDOCLST")
   Call FRESHECHO("PDOCLST")
   '
   FIN& = ULTREG&("PDOCSPL")
   TTXX$ = ""
   For U& = 1470000 To FIN&
     XX$ = REGLEIDO$("PDOCSPL", U&)
     IU& = CVS(Left$(XX$, 4))
     If IU& < 10000 Then
       IU& = IU& + 73742
       Call REPLA(XX$, MKS$(IU&), 1, 4)
       Call GRAREG("PDOCSPL", XX$, U&)
     End If
   Next U&
   Call CIERRARCH("PDOCSPL")
   Screen.MousePointer = 1
   Call COMUNI("Recuperacion Completada")
End Sub

'
Private Sub Form_Load()
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    Call CENTRAFORM(Me)
    Call APLICACIONSKINS(Me)
    Show
    Call ACOMODATODO
    Refresh
    DoEvents
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
      Caption = Caption + "  -  " + EPAC$
    End If
    DoEvents
    '
    HOII% = HOY(HOS$)
    If HOII% = FECHANUM("21/11/05") Then
      Command99.Visible = True
    End If
    '
    Screen.MousePointer = 11
    On Error Resume Next
    Combo1.Clear
    For i% = 1 To Printer.FontCount
        Printer.FontName = Printer.Fonts(i% - 1)
        Combo1.AddItem Printer.FontName
    Next i%
    On Error GoTo 0
    Combo1.TEXT = "Courier New"
    '
    Call DIBUREGHOR
    Call DIBUREGVER
    Picture4.BackColor = Picture1.BackColor
    
    ' Call CARINDIFORM
    '
    If APLINVO$ <> "" Then
       CODFOR$ = APLINVO$
       Picture1.Cls
       Text1.TEXT = CODFOR$
       Text2.TEXT = ECOARCH$("INDIFRM", CODFOR$)
    End If
    '
    Call CENTRAFORM(Me)
    Screen.MousePointer = 1
    '
End Sub

Sub DIBUREGHOR()
    'DIBUJA REGLA HORIZONTAL
    Picture6.Line (Picture1.Left, 200)-(17000, 200)
    For IX = 0 To 305 Step 5
       INI = 150
       Picture6.CurrentX = Picture1.Left + IX * 56.7 - 90
       Picture6.CurrentY = 250
       Picture6.Font = TAHOMA
       Picture6.FontSize = 7
       IXXX$ = TRIM$(Str$(IX)): While Len(IXXX$) < 3: IXXX$ = " " + IXXX$ + " ": Wend
       If IX Mod 10 = 0 Then
          INI = 110
          If IX >= 0 Then Picture6.Print TRIM$(Str$(IX))
       End If
       If IX Mod 50 = 0 Then
          INI = 50
       End If
       Picture6.Line (Picture1.Left + IX * 56.7, INI)-(Picture1.Left + IX * 56.7, 200)
    Next IX
    '
End Sub
'
Sub DIBUREGVER()
    'DIBUJA REGLA VERTICAL
    Picture7.Line (200, Picture1.Top)-(200, 25000)
    For IX = 0 To 305 Step 5
       INI = 150
       Picture7.CurrentX = 250
       Picture7.CurrentY = Picture1.Top + IX * 56.7 - 60
       Picture7.Font = TAHOMA
       Picture7.FontSize = 7
       IXXX$ = TRIM$(Str$(IX)): While Len(IXXX$) < 3: IXXX$ = " " + IXXX$ + " ": Wend
       If IX Mod 10 = 0 Then
          INI = 110
          If IX >= 0 Then Picture7.Print TRIM$(Str$(IX))
       End If
       If IX Mod 50 = 0 Then
          INI = 50
       End If
       Picture7.Line (INI, Picture1.Top + IX * 56.7)-(200, Picture1.Top + IX * 56.7)
    Next IX
    '
End Sub

Private Sub Form_Resize()
    Call CENTRAFORM(Me)
    Picture8.Left = ScaleWidth - Picture8.Width
    ' Picture2.Left = ScaleWidth - 3500
    Frame10.Left = ScaleWidth - 3500
    Frame11.Left = ScaleWidth - 3500
    Frame3.Left = ScaleWidth - 3600
    Frame3.Height = ScaleHeight + 200
    Command12.Left = Frame3.Left - Command12.Width
    Call ACOMODATODO
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub frame3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ARRABOR% = 1
End Sub

Private Sub frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Frame3.MousePointer = 9
   If ARRABOR% = 1 Then
      If Frame3.Left + X > 20 Then
          Frame3.Left = Frame3.Left + X
          Call ACOMODATODO
          Call DIBUREGHOR
      End If
   End If
End Sub

Private Sub frame3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ARRABOR% = 0
End Sub

Sub ACOMODATODO()
    '
Printer.Orientation = 1
    '
    On Error Resume Next
    Picture1.Width = Printer.ScaleWidth
    Picture1.Height = Printer.ScaleHeight
    If FOR_ORIE$ = "HORIZONTAL" Then
       Picture1.Height = Printer.ScaleWidth
       Picture1.Width = Printer.ScaleHeight
    End If
    Picture4.Width = Picture1.Left + Picture1.Width + VScroll1.Width + 150
    Picture4.Height = Picture1.Top + Picture1.Height
    
    Frame10.Left = Frame3.Left + Frame3.Width - 30
    Frame10.Width = Width - Frame10.Left - 250
    Frame11.Left = Frame3.Left + Frame3.Width - 30
    Frame11.Width = Width - Frame11.Left - 250
    Frame5.Left = Frame3.Left + Frame3.Width - 30
    Frame5.Width = Width - Frame5.Left
    '
    Command12.Left = Frame3.Left - Command12.Width - 130
    Label19.Left = Frame4.Left + Frame4.Width + 100
    Label19.Width = Command12.Left - Label19.Left - 100
    
    Frame10.Height = Frame5.Top - Frame10.Top
    Frame11.Top = Frame5.Top + Frame5.Height - 50
    Frame11.Height = ScaleHeight - Frame11.Top - 200
    '
    Picture5.Width = Frame3.Left - Picture5.Left - 100
    Picture5.Height = ScaleHeight - Picture5.Top - 200   ' Top - 500   '100
    Frame7.Top = Picture5.Top + Picture5.Height - Frame7.Height
    VScroll1.Top = -20
    VScroll1.Left = Picture5.Width - VScroll1.Width - 50
    VScroll1.Height = Picture5.Height - HScroll1.Height - 50 '- 550
    '
    If Picture4.Width < VScroll1.Left + 20 Then Picture4.Width = VScroll1.Left + 20
    If Picture1.Width < Picture4.Width Then Picture1.Width = Picture4.Width
    
    HScroll1.Left = -20
    HScroll1.Top = Picture5.Height - HScroll1.Height - 50
    HScroll1.Width = Picture5.Width - VScroll1.Width
    '
    Picture6.Top = -30
    Picture6.Height = Picture1.Top + 60
    Picture6.Left = Picture4.Left
    Picture6.Width = Picture4.Width
    '
    Picture7.Left = -30
    Picture7.Width = Picture1.Left + 60
    Picture7.Top = Picture4.Top
    Picture7.Height = Picture4.Height
    '
    VScroll1.Value = 0
    HScroll1.Value = 0
    Picture4.Top = 0
    Picture4.Left = 0
    DIFHOR = Picture4.Width - Picture5.Width: If DIFHOR < 0 Then DIFHOR = 0
    HScroll1.Max = DIFHOR
    DIFVER = Picture4.Height - Picture5.Height: If DIFVER < 0 Then DIFVER = 0
    VScroll1.Max = DIFVER
    Call VScroll1_Scroll
    Call HScroll1_Scroll
    On Error GoTo 0
    '
End Sub
'
Private Sub frame5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ARRABOR% = 1
End Sub

Private Sub frame5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Frame5.MousePointer = 7
   If ARRABOR% = 1 Then
      Frame5.Top = Frame5.Top + Y
      Call ACOMODATODO
   End If
End Sub

Private Sub frame5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ARRABOR% = 0
End Sub

Private Sub HScroll1_Scroll()
   Picture4.Left = -30 - HScroll1.Value
   Picture6.Left = Picture4.Left
   Call DIBUREGHOR
End Sub

Private Sub Image1_Click(Index As Integer)
    '
    Call Picture1_GotFocus
    '
    If Index >= 0 Then
       IOBJ% = CANOBJ% + 1
       Label14 = IOBJ%
       Label16 = "Texto:"
         If XVALO(UTAFUE) < 1 Then UTAFUE = Text8
       Select Case Index
          Case 1
            MODO$ = "LABEL": Label15 = "LABEL (Leyenda Fija)"
            If Combo1.TEXT = "" Then Combo1.TEXT = "Courier New"
            If UTAFUE = "" Then UTAFUE = "8"
            Text11.Visible = False
            Label16 = "Texto:"
            Command13.Visible = False: Command18.Visible = False: Command20.Visible = False
            Text8 = UTAFUE
            Check1.Value = UNEGRI: NEGRI(IOBJ%) = UNEGRI
            Check2.Value = UITALI: ITALI(IOBJ%) = UITALI
            Check3.Value = USUBRA: SUBRA(IOBJ%) = USUBRA
            Check4.Value = 0: INVIS(IOBJ%) = 0
            Check5.Value = 0: ENCAS(IOBJ%) = 0
            Check6.Value = 0: MULIN(IOBJ%) = 0
          Case 2
            MODO$ = "PARAM": Label15 = "PARAM (Parametro)"
            If Combo1.TEXT = "" Then Combo1.TEXT = "Courier New"
            If UTAFUE = "" Then UTAFUE = "8"
            Text8 = UTAFUE
            Check1.Value = UNEGRI: NEGRI(IOBJ%) = UNEGRI
            Check2.Value = UITALI: ITALI(IOBJ%) = UITALI
            Check3.Value = USUBRA: SUBRA(IOBJ%) = USUBRA
            Check4.Value = 0: INVIS(IOBJ%) = 0
            Check5.Value = 0: ENCAS(IOBJ%) = 0
            Check6.Value = 0: MULIN(IOBJ%) = 0
            Text11.Top = Text10.Top
            Text11.Left = Text10.Left
            Text11.Visible = True
            Text11 = "F10.2": Call Text11_Change
            Label16 = "Formato:"
            Command13.Visible = True: Command18.Visible = False: Command20.Visible = False
            If FORPA$(IOBJ%) <> "" Then Text11.TEXT = FORPA$(IOBJ%)
            Text10 = TMASCARA$(Text11)
          Case 3
            MODO$ = "LINE": Label15 = "LINE (Linea)"
            ALIN = XVALO(Text8)
            If ALIN < 1 Or ALIN > 3 Then ALIN = 1
            Text8 = CStr(ALIN)                    ' espesor
          Case 4
            MODO$ = "BOX": Label15 = "BOX (Rectangulo)"
            ALIN = XVALO(Text8)
            If ALIN < 1 Or ALIN > 3 Then ALIN = 1
            Text8 = CStr(ALIN)                    ' espesor
          Case 5
            MODO$ = "GRAPH": Label15 = "GRAPH (Imagen)"
            Command13.Visible = False
            Text11.Visible = False
            Command18.FontSize = 1
            Command18.Left = Text10.Left + Text10.Width
            Command18.Top = Text10.Top
            Command18.FontSize = 14
            Command18.Visible = True: Command20.Visible = False
            Text8 = ""
            Label16 = "Archivo:"
          Case 6
            MODO$ = "TEXT": Label15 = "TEXT (Caja de Texto)"
            If Combo1.TEXT = "" Then Combo1.TEXT = "Courier New"
            If UTAFUE = "" Then UTAFUE = "8"
            Text8 = UTAFUE
            Check1.Value = UNEGRI: NEGRI(IOBJ%) = UNEGRI
            Check2.Value = UITALI: ITALI(IOBJ%) = UITALI
            Check3.Value = USUBRA: SUBRA(IOBJ%) = USUBRA
            Check4.Value = 0: INVIS(IOBJ%) = 0
            Check5.Value = 0: ENCAS(IOBJ%) = 1
            Check6.Value = UMULIN: MULIN(IOBJ%) = UMULIN
            Text11.Top = Text10.Top
            Text11.Left = Text10.Left
            Text11.Visible = True
            Text11 = "F10.2": Call Text11_Change
            Label16 = "Formato:"
            Command13.Visible = True: Command18.Visible = False: Command20.Visible = False
            If FORPA$(IOBJ%) <> "" Then Text11.TEXT = FORPA$(IOBJ%)
            Text10 = TMASCARA$(Text11)
          Case 7
            MODO$ = "GRID": Label15 = "GRID (Grilla)"
            Combo1.TEXT = "Lucida Console"
            Text8 = "9"                         ' tamaño de fuente
            Command13.Visible = False
            Text11.Visible = False
            Command20.FontSize = 1
            Command20.Left = Text10.Left + Text10.Width
            Command20.Top = Text10.Top
            Command20.FontSize = 14
            Command20.Visible = True: Command18.Visible = False
            Text8 = ""
            Label16 = "Consulta:"
          Case 8
            MULTISEL% = 1
            MODO$ = "MULTISEL": Label19 = "MULTISEL"
            CANOBJA% = CANOBJ%
            OFHOR = (-1) * Picture4.Left / 56.7
            OFVER = (-1) * Picture4.Top / 56.7
            Combo1.TEXT = "Arial": Text8 = "8"
            Call CREAOBJETO("LABEL", 0 + OFHOR, 0 + OFVER, "User - Password:", 0, "A16", "", "", "", "", 0, "")
            Call CREAOBJETO("TEXT", H2(CANOBJA% + 1), 0 + OFVER, "", 0, "A12", "@PASSWORD", "", "", "", 1, "@IDENT")
            Call CREAOBJETO("TEXT", 0 + OFHOR, 2 + V2(CANOBJA% + 1), "", 0, "A26", "APEYNOM({" + ONAME$(CANOBJA% + 2) + "})", "", "", "", 0, ONAME$(CANOBJA% + 2))
            '
            PHORI = H2(CANOBJA% + 2): If H2(CANOBJA% + 3) > PHORI Then PHORI = H2(CANOBJA% + 3)
            Call CREAOBJETO("LABEL", 5 + PHORI, 0 + OFVER, "Fecha:", 0, "A6", "", "", "", "", 0, "")
            Call CREAOBJETO("TEXT", H2(CANOBJA% + 4), 0 + OFVER, "", 0, "D8", "@HOY", "", "", "", 1, ONAME$(CANOBJA% + 2))
            Call CREAOBJETO("LABEL", 5 + PHORI, 2 + V2(CANOBJA% + 1), " Hora:", 0, "A6", "", "", "", "", 0, "")
            Call CREAOBJETO("TEXT", H2(CANOBJA% + 4), 2 + V2(CANOBJA% + 1), "", 0, "H8", "@AHORA", "", "", "", 1, ONAME$(CANOBJA% + 2))
            '
            For KKI% = CANOBJA% + 1 To CANOBJ%
               ISEL%(KKI%) = 1
               Call DIBUOBJ%(KKI%, 2)
            Next KKI%
            Call Picture1.SetFocus
            Exit Sub
       End Select
       '
       MAXCAI% = 1
25     ONAM$ = MODO$ + "-" + TRIM$(Str$(MAXCAI%))
       For KKI% = 1 To CANOBJ%
          If ONAME$(KKI%) = ONAM$ Then
             MAXCAI% = MAXCAI% + 1
             GoTo 25
          End If
       Next KKI%
       Text9 = ONAM$
       Text4 = "0.0"
       Text5 = "0.0"
       Text6 = "0.0"
       Text7 = "0.0"
    End If
End Sub

Private Sub Image2_DBLClick()
   TXT$ = TRIM$(Text1.TEXT)
   If TXT$ <> "" Then
      If DERACCE%("ULOCKFRM", "Bloqueo y Liberacion de Diseño") >= 2 Then
         If COMALTER%("Confirme Bloqueo de Diseño del Formulario\\Cancelar\Bloquear") = 2 Then
            Image2.Visible = False
            Image3.Visible = True
         End If
      End If
   End If
End Sub

Private Sub Image3_DblClick()
   TXT$ = TRIM$(Text1.TEXT)
   If TXT$ <> "" Then
      If DERACCE%("ULOCKFRM", "Bloqueo y Liberacion de Diseño") >= 2 Then
         If COMALTER%("Confirme Re-Apertura de Diseño del Formulario\\Cancelar\Re-Abrir") = 2 Then
            Image3.Visible = False
            Image2.Visible = True
         End If
      End If
   End If
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Left$(Label15, 4) = "LINE" Then
      XORI = 56.7 * XVALO(Text4)
      YORI = 56.7 * XVALO(Text6)
      XMOUSE = 56.7 * XVALO(Text5)
      YMOUSE = 56.7 * XVALO(Text7)
      ESTAPRETADO% = 1
      Call Picture1_MouseMove(1, 0, XMOUSE, YMOUSE)
      Picture1.SetFocus
'      Call Picture1_MouseMove(1, 0, XMOUSE, YMOUSE)
   End If
      
End Sub

Private Sub Label14_Change()
    Static IOBJA%
    If VScroll2.Max < CANOBJ% Then
       If VScroll2.Value < CANOBJ% Then
          VScroll2.Max = CANOBJ%
       End If
    End If
    On Error Resume Next
    VScroll2.Value = CANOBJ% + 1 - XVALO(Label14)
    On Error GoTo 0
    '
    Call DIBUOBJ(IOBJA%)
    IOBJ% = Label14
    IOBJA% = IOBJ%
    Call DIBUOBJ%(IOBJ%, 1)
    '
    Label15 = TIOBJ$(IOBJ%)
    MODO$ = TIOBJ$(IOBJ%)
    '
    Text4 = FORMATNUM$(H1(IOBJ%), "F7.1")
    Text5 = FORMATNUM$(H2(IOBJ%), "F7.1")
    Text6 = FORMATNUM$(V1(IOBJ%), "F7.1")
    Text7 = FORMATNUM$(V2(IOBJ%), "F7.1")
    '
    If IOBJ% <= CANOBJ% Then
        If FUENTE$(IOBJ%) <> "" Then Combo1 = FUENTE$(IOBJ%)
        If XVALO(TAFUEN$(IOBJ)) >= 1 Then Text8 = TAFUEN$(IOBJ%)
        Check1.Value = NEGRI(IOBJ%)
        Check2.Value = ITALI(IOBJ%)
        Check3.Value = SUBRA(IOBJ%)
        Check4.Value = INVIS(IOBJ%)
        Check5.Value = ENCAS(IOBJ%)
        Check6.Value = MULIN(IOBJ%)
      Else
        Check1.Value = UNEGRI
        Check2.Value = UITALI
        Check3.Value = USUBRA
        Check4.Value = 0
        Check5.Value = 0
        Check6.Value = UMULIN
    End If
    '
    Text9 = ONAME$(IOBJ%)
    Text10 = TEXTO(IOBJ%)
    Text15 = ARDOCLI$(IOBJ%)
    '
    Command13.Visible = False: Command18.Visible = False: Command20.Visible = False
    Text11.Visible = False
    Label16 = "Texto:"
    If MODO$ = "GRAPH" Then
       Label16 = "Archivo:"
       Command18.Visible = True: Command20.Visible = False
    End If
    If MODO$ = "GRID" Then
       Label16 = "Consulta:"
       Command18.Visible = False: Command20.Visible = True
       Command20.FontSize = 1
       Command20.Left = Text10.Left + Text10.Width
       Command20.FontSize = 14
    End If
    If MODO$ = "PARAM" Or MODO$ = "TEXT" Then
       Label16 = "Formato:"
       Command13.Visible = True
       Text11.Top = Text10.Top
       Text11.Left = Text10.Left
       Text11.Visible = True
       Text11.TEXT = FORPA$(IOBJ%)
    End If
    '
End Sub

Private Sub Label19_dblClick()
    '
    If MULTISEL% > 0 Or Label19 = "MULTISEL" Then
       For i% = 1 To CANOBJ%
          If ISEL%(i%) > 0 Then
             ISEL%(i%) = 0
             Call DIBUOBJ%(i%)
          End If
       Next i%
       MULTISEL% = 0: DHOR = 0: DVER = 0
       Label19.ToolTipText = ""
       Call Command10_Click
       If MODO$ = "MULTISEL" Then
          MODO$ = "": Label19 = MODO$
       End If
    End If
    '
End Sub

Private Sub Option1_Click()
   If Option1.Value = True Then FOR_ORIE$ = "VERTICAL"
   Call ACOMODATODO
End Sub

Private Sub Option2_Click()
   If Option2.Value = True Then FOR_ORIE$ = "HORIZONTAL"
   Call ACOMODATODO
End Sub

Private Sub Picture1_DBLClick()
  '
  If MULTISEL% > 0 Then Exit Sub
  '
  Call Label19_dblClick
  OBUS$ = "LABEL"
  '
2 For i% = CANOBJ% To 1 Step -1
    If TIOBJ$(i%) = OBUS$ Then
      HA = H1(i%): HB = H2(i%)
      If HA > HB Then
         HA = H2(i%): HB = H1(i%)
      End If
      VXA = V1(i%): VXB = V2(i%)
      If VXA > VXB Then
         VXA = V2(i%): VXB = V1(i%)
      End If
      If XMOUSE >= 56.7 * (HA - 0.5) Then
        If XMOUSE <= 56.7 * (HB + 0.5) Then
          If (YMOUSE >= 56.7 * (VXA - 0.5) And YMOUSE <= 56.7 * (VXB + 0.5)) Or (YMOUSE >= 56.7 * (VXB - 0.5) And YMOUSE <= 56.7 * (VXA + 0.5)) Then
               GoTo 6
          End If
        End If
      End If
    End If
  Next i%
  '
  If OBUS$ = "LABEL" Then OBUS$ = "PARAM": GoTo 2
  If OBUS$ = "PARAM" Then OBUS$ = "TEXT": GoTo 2
  If OBUS$ = "TEXT" Then OBUS$ = "LINE": GoTo 2
  If OBUS$ = "LINE" Then OBUS$ = "GRAPH": GoTo 2
  If OBUS$ = "GRAPH" Then OBUS$ = "GRID": GoTo 2
  If OBUS$ = "GRID" Then OBUS$ = "BOX": GoTo 2
  Exit Sub
  '
6 ' PRESENTA DATOS DEL OBJETO
  IOBJ% = i%
  Label14 = IOBJ%
  Label15 = TIOBJ$(i%)
  MODO$ = TIOBJ$(i%)
  '
  Text4 = FORMATNUM$(H1(i%), "F7.1")
  Text5 = FORMATNUM$(H2(i%), "F7.1")
  Text6 = FORMATNUM$(V1(i%), "F7.1")
  Text7 = FORMATNUM$(V2(i%), "F7.1")
  Call DIBUOBJ(i%)
  '
  If TIOBJ$(i%) = "LABEL" Or TIOBJ$(i%) = "PARAM" Or TIOBJ$(i%) = "TEXT" Then
        '
        Combo1.TEXT = FUENTE$(i%)
        Text8.TEXT = TAFUEN(i%)
        Check1.Value = NEGRI(i%):     Text3.FontBold = Check1.Value
        Check2.Value = ITALI(i%):     Text3.FontItalic = Check2.Value
        Check3.Value = SUBRA(i%):     Text3.FontUnderline = Check3.Value
        Check4.Value = INVIS(i%):     Text3.FontStrikethru = Check4.Value
        Check5.Value = ENCAS(i%)
        Check6.Value = MULIN(i%)
        '
        TFS = Text3.FontSize
        Text3.FontSize = 1
        On Error Resume Next
        Text3.Top = 56.7 * V1(i%) - 50: Text3.Height = 56.7 * (V2(i%) - V1(i%))
        Text3.Left = 56.7 * H1(i%) - 50: Text3.Width = 56.7 * (H2(i%) - H1(i%))
        Text3.FontSize = TFS
        On Error GoTo 0
        '
        Text3.ForeColor = RGB(0, 0, 0)
        If TIOBJ$(i%) = "PARAM" Then Text3.ForeColor = RGB(255, 0, 0)
        If TIOBJ$(i%) = "TEXT" Then Text3.ForeColor = RGB(0, 0, 255)
        Text3.Visible = True
        If TIOBJ$(i%) = "LABEL" Then
             Text3 = TEXTO(i%)
           Else
             Text3 = TMASCARA$(FORPA$(i%))
        End If
        Text3.Refresh
        Text3.SetFocus
        Exit Sub
        '
     ElseIf TIOBJ$(i%) = "BOX" Then
        MODO$ = "BOX"
        ESTAPRETADO% = 2
        '
     ElseIf TIOBJ$(i%) = "GRAPH" Then
        MODO$ = "GRAPH"
        ESTAPRETADO% = 2
        '
     ElseIf TIOBJ$(i%) = "GRID" Then
        MODO$ = "GRID"
        Combo1.TEXT = FUENTE$(i%)
        Text8.TEXT = TAFUEN(i%)
        ESTAPRETADO% = 2
        '
     ElseIf TIOBJ$(i%) = "LINE" Then
        MODO$ = "LINE"
        ESTAPRETADO% = 2
  End If
  '
End Sub

Private Sub Picture1_GotFocus()
   Call BORRATRASTEX
   Text3.Visible = False
   Text3 = ""
   Text3.Width = 0
   Text3.Height = 0
   Command13.Visible = False
   Text11.Visible = False: Label16 = "Texto:"
   IOBJ% = XVALO(Label14)
   If TIOBJ$(IOBJ%) = "PARAM" Or MODO$ = "PARAM" Then
      Call DIBUOBJ(IOBJ%)
      Command13.Visible = True: Command18.Visible = False: Command20.Visible = False
      Text11.Visible = True
      Label16 = "Formato:"
   ElseIf TIOBJ$(IOBJ%) = "TEXT" Or MODO$ = "TEXT" Then
      Call DIBUOBJ(IOBJ%)
      Command13.Visible = True: Command18.Visible = False: Command20.Visible = False
      Text11.Visible = True
      Label16 = "Formato:"
   ElseIf TIOBJ$(IOBJ%) = "LABEL" Or MODO$ = "LABEL" Then
      Call DIBUOBJ(IOBJ%)
   ElseIf TIOBJ$(IOBJ%) = "GRAPH" Then
      Command18.Visible = True
   ElseIf TIOBJ$(IOBJ%) = "GRID" Then
      Command20.Visible = True
   End If
End Sub

Private Sub Picture1_KeyDOWN(KeyCode As Integer, Shift As Integer)
   Static TANTE
   If KeyCode = 46 Then          ' TECLA SUPRIMIR
      Call Command15_Click
      Exit Sub
   End If
   '
   If Timer < TANTE + 0.1 Then Exit Sub
   TANTE = Timer
   If KeyCode = 38 Then ' FLECHA ARRIBA
       If MULTISEL% = 1 Then
           Call MUEVESELEC("ARRIBA", Shift)
           Exit Sub
       End If
       '
       PORCOORDE% = 1
       DIFTE = XVALO(Text7) - XVALO(Text6)
       T6 = XVALO(Text6) - 0.1
          If Shift > 0 Then T6 = Int(T6)
          If Shift = 2 Then T6 = T6 - 9
          If T6 < 0 Then T6 = 0
       Text6 = FORMATNUM(T6, "F7.1")
       Text7 = FORMATNUM(T6 + DIFTE, "F7.1")
       PORCOORDE% = 0
       Exit Sub
     ElseIf KeyCode = 40 Then ' FLECHA ABAJO
       If MULTISEL% = 1 Then
           Call MUEVESELEC("ABAJO", Shift)
           Exit Sub
       End If
       '
       PORCOORDE% = 1
       DIFTE = XVALO(Text7) - XVALO(Text6)
       T6 = XVALO(Text6) + 0.1
          If Shift > 0 Then T6 = 1 + Int(T6)
          If Shift = 2 Then T6 = T6 + 9
       Text6 = FORMATNUM(T6, "F7.1")
       Text7 = FORMATNUM(T6 + DIFTE, "F7.1")
       PORCOORDE% = 0
       Exit Sub
     ElseIf KeyCode = 37 Then ' FLECHA IZQUIERDA
       If MULTISEL% = 1 Then
           Call MUEVESELEC("IZQUIER", Shift)
           Exit Sub
       End If
       '
       PORCOORDE% = 1
       DIFTE = XVALO(Text5) - XVALO(Text4)
       T4 = XVALO(Text4) - 0.1
          If Shift > 0 Then T4 = Int(T4)
          If Shift = 2 Then T4 = T4 - 9
       Text4 = FORMATNUM(T4, "F7.1")
       Text5 = FORMATNUM(T4 + DIFTE, "F7.1")
       PORCOORDE% = 0
       Exit Sub
     ElseIf KeyCode = 39 Then ' FLECHA DERECHA
       If MULTISEL% = 1 Then
           Call MUEVESELEC("DERECH", Shift)
           Exit Sub
       End If
       '
       PORCOORDE% = 1
       DIFTE = XVALO(Text5) - XVALO(Text4)
       T4 = XVALO(Text4) + 0.1
          If Shift > 0 Then T4 = 1 + Int(T4)
          If Shift = 2 Then T4 = T4 + 9
          If T4 < 0 Then T4 = 0
       Text4 = FORMATNUM(T4, "F7.1")
       Text5 = FORMATNUM(T4 + DIFTE, "F7.1")
       PORCOORDE% = 0
       Exit Sub
  End If

End Sub

Sub MUEVESELEC(ADON$, SH)
  For i% = 1 To CANOBJ%
     If ISEL%(i%) = 1 Then
        Call BORRAOBJ(i%)
        DIFEX = 0.1: If SH = 1 Then DIFEX = 1 Else If SH = 2 Then DIFEX = 10
        If ADON$ = "ARRIBA" Then
             V1(i%) = V1(i%) - DIFEX
             V2(i%) = V2(i%) - DIFEX
          ElseIf ADON$ = "ABAJO" Then
             V1(i%) = V1(i%) + DIFEX
             V2(i%) = V2(i%) + DIFEX
          ElseIf ADON$ = "DERECH" Then
             H1(i%) = H1(i%) + DIFEX
             H2(i%) = H2(i%) + DIFEX
          ElseIf ADON$ = "IZQUIER" Then
             H1(i%) = H1(i%) - DIFEX
             H2(i%) = H2(i%) - DIFEX
        End If
        Call DIBUOBJ(i%, 2)
        Picture1.SetFocus
     End If
  Next i%
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '
    If Shift > 0 Then
       MULTISEL% = 1
       MODO$ = "MULTISEL": Label19 = "MULTISEL"
       Label19.ToolTipText = "Doble Click para Desactivar Seleccion Multiple"
       Call SELMULTI(X, Y)
       Exit Sub
    End If
    '
    If MULTISEL% > 0 Then
       Call Label19_dblClick
       Exit Sub
    End If
    '
    PORCOORDE% = 0
    XMOUSE = X: YMOUSE = Y
    If MODO$ = "" Then Exit Sub
    If Text3.Visible = True Then Exit Sub
    '
    If ESTAPRETADO = 0 Then
       ESTAPRETADO% = 1
       IOBJ% = CANOBJ% + 1
       Label14 = IOBJ%
       '
       MAXCAI% = 1
25     ONAM$ = MODO$ + "-" + TRIM$(Str$(MAXCAI%))
       For KKI% = 1 To CANOBJ%
          If ONAME$(KKI%) = ONAM$ Then
             MAXCAI% = MAXCAI% + 1
             GoTo 25
          End If
       Next KKI%
       Text9 = ONAM$
       '
       XORI = X: XMOUSE = X
         Text4 = FORMATNUM$(X / 56.7, "F7.1")
         Text5 = FORMATNUM$(X / 56.7, "F7.1")
       YORI = Y: YMOUSE = Y
         Text6 = FORMATNUM$(Y / 56.7, "F7.1")
         Text7 = FORMATNUM$(Y / 56.7, "F7.1")
    End If
    '
    If ESTAPRETADO = 2 Then
       i% = IOBJ%
       HA = H1(i%): HB = H2(i%): If HA > HB Then HA = H2(i%): HB = H1(i%)
       VXA = V1(i%): VXB = V2(i%): If VXA > VXB Then VXA = V2(i%): VXB = V1(i%)
       If XMOUSE >= 56.7 * (HA - 0.5) Then
         If XMOUSE <= 56.7 * (HB + 0.5) Then
           If (YMOUSE >= 56.7 * (VXA - 0.5) And YMOUSE <= 56.7 * (VXB + 0.5)) Or (YMOUSE >= 56.7 * (VXB - 0.5) And YMOUSE <= 56.7 * (VXA + 0.5)) Then
'       If XMOUSE >= 56.7 * (H1(IOBJ%) - 0.5) Then
'         If XMOUSE <= 56.7 * (H2(IOBJ%) + 0.5) Then
'           If (YMOUSE >= 56.7 * (V1(IOBJ%) - 0.5) And YMOUSE <= 56.7 * (V2(IOBJ%) + 0.5)) Or (YMOUSE >= 56.7 * (V2(I%) - 0.5) And YMOUSE <= 56.7 * (V1(I%) + 0.5)) Then
             ESTAPRETADO% = 1
             GoTo 35
           End If
         End If
       End If
       Call Picture1_MouseUP(1, 0, XMOUSE, YMOUSE)
       ESTAPRETADO% = 0
       Exit Sub
       '
35     MUEVETODO% = 1
       If Abs(XMOUSE - 56.7 * XVALO(Text5)) < 100 Then
          If Abs(YMOUSE - 56.7 * XVALO(Text7)) < 100 Then
             MUEVETODO% = 0
          End If
       End If
       Screen.MousePointer = 1
       If MUEVETODO% = 1 Then
          Screen.MousePointer = 5
          Exit Sub
       End If
       Screen.MousePointer = 8
       '
       IOBJ% = XVALO(Label14)
       XORI = 56.7 * XVALO(Text4): XMOUSE = X
         Text5 = FORMATNUM$(X / 56.7, "F7.1")
       YORI = 56.7 * XVALO(Text6): YMOUSE = Y
         Text7 = FORMATNUM$(Y / 56.7, "F7.1")
    End If
End Sub
'
Sub SELMULTI(XX1, YY1)
   '
   OBUS$ = "LABEL"
2  For i% = CANOBJ% To 1 Step -1
     If TIOBJ$(i%) = OBUS$ Then
       HA = H1(i%): HB = H2(i%)
       If HA > HB Then
          HA = H2(i%): HB = H1(i%)
       End If
       VXA = V1(i%): VXB = V2(i%)
       If VXA > VXB Then
          VXA = V2(i%): VXB = V1(i%)
       End If
       If XX1 >= 56.7 * (HA - 0.5) Then
         If XX1 <= 56.7 * (HB + 0.5) Then
           If (YY1 >= 56.7 * (VXA - 0.5) And YY1 <= 56.7 * (VXB + 0.5)) Or (YY1 >= 56.7 * (VXB - 0.5) And YY1 <= 56.7 * (VXA + 0.5)) Then
                GoTo 6      ' SELECCIONADO
           End If
         End If
       End If
     End If
   Next i%
   '
   If OBUS$ = "LABEL" Then OBUS$ = "PARAM": GoTo 2
   If OBUS$ = "PARAM" Then OBUS$ = "TEXT": GoTo 2
   If OBUS$ = "TEXT" Then OBUS$ = "LINE": GoTo 2
   If OBUS$ = "LINE" Then OBUS$ = "GRAPH": GoTo 2
   If OBUS$ = "GRAPH" Then OBUS$ = "GRID": GoTo 2
   If OBUS$ = "GRID" Then OBUS$ = "BOX": GoTo 2
   Exit Sub
   '
6  If ISEL%(i%) = 1 Then
        ISEL%(i%) = 0
        Call DIBUOBJ(i%)
      Else
        ISEL%(i%) = 1
        Call DIBUOBJ(i%, 2)
   End If
   '
End Sub
'
Private Sub Picture1_MouseUP(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '
    If ESTAPRETADO% <> 1 Then Exit Sub
    '
    If MUEVETODO% = 1 Then
       MUEVETODO% = 0
       XORI = H1(IOBJ%) * 56.7
       YORI = V1(IOBJ%) * 56.7
       XMOUSE = H2(IOBJ%) * 56.7
       YMOUSE = V2(IOBJ%) * 56.7
    End If
    Screen.MousePointer = 1
    '
    If MODO$ <> "LINE" Then
       If XORI > XMOUSE Then SWAP XORI, XMOUSE
          Text4 = FORMATNUM$(XORI / 56.7, "F7.1")
          Text5 = FORMATNUM$(XMOUSE / 56.7, "F7.1")
       If YORI > YMOUSE Then SWAP YORI, YMOUSE
          Text6 = FORMATNUM$(YORI / 56.7, "F7.1")
          Text7 = FORMATNUM$(YMOUSE / 56.7, "F7.1")
    End If
    '
    Call GUARDAELOBJETO
    '
    ESTAPRETADO% = 0
    If MODO$ = "LABEL" Or MODO$ = "PARAM" Or MODO$ = "TEXT" Then
       Picture1.Line (XORI, YORI)-(XMOUSE, YMOUSE), RGB(255, 255, 255), B ' BORRA RECUADRO
       If XMOUSE <> XORI Or YMOUSE <> YORI Then
          Text3.Visible = True
          Text3.Left = XORI + 10
          Text3.Top = YORI + 10
          On Error Resume Next
            Text3.Width = X - XORI - 20
            If MODO$ = "LABEL" Then
                 Text3.ForeColor = RGB(0, 0, 0)
                 Call Text3_Change
                 Text3.SetFocus
               ElseIf MODO$ = "PARAM" Then
                 Text3.ForeColor = RGB(255, 0, 0)
                 Text3 = Text10
                 Call Text3_Change
                 Text11.SetFocus
               ElseIf MODO$ = "TEXT" Then
                 Text3.ForeColor = RGB(0, 0, 255)
                 Text3 = Text10
                 Call Text3_Change
                 Text11.SetFocus
           End If
            Call PREDITEXTO
          On Error GoTo 0
       End If
    End If
    '
    If MODO$ = "GRAPH" Or MODO$ = "GRID" Then
       IOBJ% = XVALO(Label14)
       Call DIBUOBJ(IOBJ%)
    End If
    '
    MODO$ = ""
End Sub

Sub GUARDAELOBJETO()
    ' GUARDA EL OBJETO
    If Text4 <> Text5 Or Text6 <> Text7 Then
       IOBJ% = XVALO(Label14)
       If CANOBJ% < IOBJ% Then CANOBJ% = IOBJ%
       '
       PPXX% = InStr(Label15, "("): If PPXX% < 1 Then PPXX% = 1 + Len(Label15)
       MODO$ = TRIM$(Left$(Label15, PPXX% - 1))
       TIOBJ$(IOBJ%) = MODO$
       ONAME$(IOBJ%) = Text9.TEXT
       H1(IOBJ%) = XVALO(Text4)
       H2(IOBJ%) = XVALO(Text5)
       V1(IOBJ%) = XVALO(Text6)
       V2(IOBJ%) = XVALO(Text7)
       '
       TEXTO(IOBJ%) = Text10
       FUENTE$(IOBJ%) = Combo1.TEXT
       TAFUEN(IOBJ%) = Text8
       '
    End If
End Sub
'
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If ESTAPRETADO = 1 Then
      Call BORRAELOBJETOACTIVO
      If MUEVETODO = 1 Then
         DIFEX = X - XMOUSE: DIFEY = Y - YMOUSE
         XMOUSE = X: YMOUSE = Y
         H1(IOBJ%) = H1(IOBJ%) + DIFEX / 56.7
         H2(IOBJ%) = H2(IOBJ%) + DIFEX / 56.7
         V1(IOBJ%) = V1(IOBJ%) + DIFEY / 56.7
         V2(IOBJ%) = V2(IOBJ%) + DIFEY / 56.7
         Text4 = FORMATNUM$(H1(IOBJ%), "F7.1")
         Text5 = FORMATNUM$(H2(IOBJ%), "F7.1")
         Text6 = FORMATNUM$(V1(IOBJ%), "F7.1")
         Text7 = FORMATNUM$(V2(IOBJ%), "F7.1")
         Call DIBUOBJ(IOBJ%)
         Exit Sub
      End If
      '
      If MODO$ = "LINE" Then
           Picture1.Line (XORI, YORI)-(XMOUSE, YMOUSE), RGB(255, 255, 255)
         ElseIf MODO$ = "BOX" Then
           Picture1.Line (XORI, YORI)-(XMOUSE, YMOUSE), RGB(255, 255, 255), B
         ElseIf MODO$ = "GRAPH" Or MODO$ = "GRID" Then
           Picture1.Line (XORI, YORI)-(XMOUSE, YMOUSE), RGB(255, 255, 255), BF
      End If
      XMOUSE = X: YMOUSE = Y
        Text5 = FORMATNUM$(X / 56.7, "F7.1")
        Text7 = FORMATNUM$(Y / 56.7, "F7.1")
      Picture1.DrawStyle = 2
      Picture1.DrawWidth = 1
      If MODO$ = "LINE" Or MODO$ = "BOX" Then
         Picture1.DrawStyle = 0
         Picture1.DrawWidth = 1
           ESPESOR = XVALO(Text8): If ESPESOR < 1 Then ESPESOR = 1
           Picture1.DrawWidth = ESPESOR
      End If
      If MODO$ = "GRAPH" Then
         Picture1.DrawStyle = 2
         Picture1.DrawWidth = 1
      End If
      If MODO$ = "GRID" Then
         Picture1.DrawStyle = 4
         Picture1.DrawWidth = 1
      End If
      If MODO$ = "LINE" Then
           If Shift = 1 Then
              If Abs(XMOUSE - XORI) > 10 Then
                  YMOUSE = YORI
                ElseIf Abs(YMOUSE - YORI) > 10 Then
                  XMOUSE = XORI
              End If
           End If
           Picture1.Line (XORI, YORI)-(XMOUSE, YMOUSE)
        Else
           Picture1.Line (XORI, YORI)-(XMOUSE, YMOUSE), , B
      End If
   End If
End Sub

Sub BORRAELOBJETOACTIVO()
   IOBJ% = XVALO(Label14)
   OBXX$ = TIOBJ$(IOBJ%): If OBXX$ = "" Then OBXX$ = MODO$
   If OBXX$ = "LINE" Then
         Picture1.Line (56.7 * H1(IOBJ%), 56.7 * V1(IOBJ%))-(56.7 * H2(IOBJ%), 56.7 * V2(IOBJ%)), RGB(255, 255, 255)
      ElseIf TIOBJ$(IOBJ%) = "GRAPH" Or MODO$ = "GRID" Then
         Picture1.Line (56.7 * H1(IOBJ%), 56.7 * V1(IOBJ%))-(56.7 * H2(IOBJ%), 56.7 * V2(IOBJ%)), RGB(255, 255, 255), BF
      Else
         Picture1.Line (56.7 * H1(IOBJ%), 56.7 * V1(IOBJ%))-(56.7 * H2(IOBJ%), 56.7 * V2(IOBJ%)), RGB(255, 255, 255), B
   End If
End Sub

Private Sub Picture2_DblClick()
   '
   Picture2.Visible = False
   Frame1.Visible = True
   VScroll1.Visible = True
   HScroll1.Visible = True
   Command1.Enabled = True
   Command2.Enabled = True
   Command3.Enabled = True
   Command4.Enabled = True
   Command5.Enabled = True
   Command6.Enabled = True
   Command7.Enabled = True
   Command8.Enabled = True
   Line2.Visible = True
   '
   Picture1.Visible = True
   Picture1.Refresh
   '
End Sub

Private Sub Picture6_Click()
    Line4.Visible = False
End Sub

Private Sub Picture6_DblClick()
    Line4.X1 = XREGLA
    Line4.X2 = XREGLA
    Line4.Visible = True
End Sub

Private Sub Picture6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    XREGLA = X - Picture1.Left
End Sub

Private Sub Picture7_Click()
    Line3.Visible = False
End Sub

Private Sub Picture7_DBLClick()
    Line3.Y1 = YREGLA
    Line3.Y2 = YREGLA
    Line3.Visible = True
End Sub

Private Sub Picture7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    YREGLA = Y - Picture1.Top
    'MsgBox YREGLA
End Sub

Private Sub Text1_Change()
    '
    If NKKK% > 0 Then
       Close #NKKK%: NKKK% = 0    ' PARA LIBERAR BLOQUEO
    End If
    '
    CODFOR$ = TRIM$(Text1)
    Text2 = ECOARCH$("INDISQR", CODFOR$)
    If TRIM$(CODFOR$) <> "" Then
       CANOBJ% = 0
       Call CARGADATFOR(CODFOR$)
       Call BLOJA(CODFOR$, CANCELOJ%)
       If CANCELOJ% > 0 Then Exit Sub
       Call ACOMODATODO
       Call DATGENFOR
       Call REDIBUJA
       VScroll2.Max = CANOBJ%
       Label14 = CStr(CANOBJ%)
    End If
    '
    ESTAPRETADO% = 0
    Call Picture1_GotFocus
    MODO$ = ""
    '
End Sub

Sub DATGENFOR()
    '
    If FOR_DESC$ <> "" Then Text2 = FOR_DESC$
    If Left$(FOR_ORIE$, 3) = "HOR" Then
         Option2.Value = True
       Else
         Option1.Value = True
    End If
    Text12 = FORMATNUM$(FOR_MGIZ, "F7.1")
    Text13 = FORMATNUM$(FOR_MGSU, "F7.1")
    '
    If AUTO_PRN$ = "SI" Then Check7.Value = 1 Else Check7.Value = 0
    If PAGI_ANT$ = "SI" Then Check8.Value = 1 Else Check8.Value = 0
    If PAGI_SIG$ = "SI" Then Check9.Value = 1 Else Check9.Value = 0
       Image2.Visible = False: Image3.Visible = False
    If LOCK_DSG$ = "SI" Then Image3.Visible = True Else Image2.Visible = True
    '
    Text16 = FORM_OPX$
    '
End Sub
'
Private Sub Text1_DblClick()
    '
    Call Command1_Click
    '
End Sub

Private Sub Text10_Change()
    IKX% = XVALO(Label14)
    TEXTO(IKX%) = Text10
    ESTAPRETADO = 0
    Call DIBUOBJ(IKX%)
End Sub

Private Sub Text11_Change()
    IKX% = XVALO(Label14)
    FORPA$(IKX%) = UCase$(Text11)
    VATEX$ = TMASCARA$(FORPA$(IKX%))
    If Text3.Visible = True Then
      If VATEX$ <> "" Then
        Text3 = VATEX$
        Text10 = VATEX$
      End If
    End If
'    If VATEX$ <> "" Then Text10 = VATEX$
'    Call Text8_LostFocus
End Sub

Private Sub Text12_Change()
   FOR_MGIZ = XVALO(Text12)
End Sub

Private Sub Text12_LostFocus()
   Text12 = FORMATNUM(FOR_MGIZ, "F7.1")
End Sub

Private Sub Text13_Change()
   FOR_MGSU = XVALO(Text13)
End Sub

Private Sub Text13_LostFocus()
   Text13 = FORMATNUM(FOR_MGSU, "F7.1")
End Sub


Private Sub Text2_Change()
    FOR_DESC$ = Text2
End Sub

Private Sub Text2_LOSTFOCUS()
    Call GRATITUFORM
End Sub

Sub CARLISFUEN()
    '
    NOMIMPRE$ = TRIM$(Printer.DeviceName)
    If NOMIMPRE$ = "" Then
         Call MENSERR(24, "Falta Seleccionar Impresora\en la Configuración de Windows")
         Call FINAL("")
    End If
    '
    'Call BLANARCH("FUENPRIN")
    Call SETULTREG("FUENPRIN", 0)
    On Error Resume Next
    For i% = 1 To Printer.FontCount
        Printer.FontName = Printer.Fonts(i% - 1)
        List1.AddItem Printer.FontName
    Next i%
    On Error GoTo 0
    '
    JJ& = 0
    For i% = 1 To List1.ListCount
      YY$ = TRIM$(List1.List(i% - 1))
      If YY$ <> "" Then
        If InStr(UCase$(YY$), "BENGUIAT") < 1 Then
          XX$ = REGBLAN$("FUENPRIN")
          Call REPLA(XX$, YY$, 1, 42)
          Printer.FontName = YY$
          A1 = 0: A2 = 0
          On Error Resume Next
          A1 = Printer.TextWidth(String$(80, "l"))
          A2 = Printer.TextWidth(String$(80, "M"))
          On Error GoTo 0
          If A1 > 0 And A2 > 0 Then
            If A1 = A2 Then
              Call REPLA(XX$, "(fija)", 43, 6)
            End If
            'Call REGAPP("FUENPRIN", XX$)
            JJ& = JJ& + 1
            Call GRAREG("FUENPRIN", XX$, JJ&)
          End If
        End If
      End If
    Next i%
    '
    Call SETULTREG("FUENPRIN", JJ&)
    Call CIERRARCH("FUENPRIN")
    '
End Sub

Sub CARINDIFORM()
    Call COPYSTRU("INDIFRM", "INDISQR")
    JJ& = 0
    For II& = 1 To ULTREG&("INDISQR")
      XX$ = REGLEIDO$("INDISQR", II&)
      COFO$ = UCase$(TRIM$(Left$(XX$, 8)))
      If EXISTE%(LUDAT$ + COFO$ + ".SQR") > 0 Then
        For KK& = 1 To JJ&
          If UCase$(TRIM$(Left$(REGLEIDO$("INDISQR", KK&), 8))) = COFO$ Then GoTo 9
        Next KK&
        JJ& = JJ& + 1
        Call GRAREG("INDISQR", XX$, JJ&)
      End If
9   Next II&
    '
    For II& = 1 To ULTREG&("INDIFRM")
      XX$ = REGLEIDO$("INDIFRM", II&)
      COFO$ = UCase$(TRIM$(Left$(XX$, 8)))
      If EXISTE%(LUDAT$ + COFO$ + ".SQR") > 0 Then
        For KK& = 1 To JJ&
          If UCase$(TRIM$(Left$(REGLEIDO$("INDISQR", KK&), 8))) = COFO$ Then GoTo 19
        Next KK&
        JJ& = JJ& + 1
        Call GRAREG("INDISQR", XX$, JJ&)
      End If
19  Next II&
    '
    Call ARMALISTARCH(LUDAT$, "*.SQR")
    For II& = 1 To ULTREG&("DIRFS")
      XX$ = AJUSTI$(REGLEIDO$("DIRFS", II&), 8)
      COFO$ = UCase$(TRIM$(Left$(XX$, 8)))
      If EXISTE%(LUDAT$ + COFO$ + ".SQR") > 0 Then
        For KK& = 1 To JJ&
          If UCase$(TRIM$(Left$(REGLEIDO$("INDISQR", KK&), 8))) = COFO$ Then GoTo 29
        Next KK&
        FOR_DESC = ""
        Call CARGADATFOR(COFO$)
        If FOR_DESC = "" Then FOR_DESC = "Formulario " + COFO$
        JJ& = JJ& + 1
        Call GRAREG("INDISQR", XX$ + FOR_DESC, JJ&)
      End If
29  Next II&
    Call SETULTREG("INDISQR", JJ&)
    Call FILESORT("INDISQR", "", 1, 1)
    Call FRESHECHO("INDISQR")
    '
End Sub
'
Sub SHOWFORM()
   '
   Static CACOMM%, PRICOMM$(128)
   '
   FORMALTER.PSPOOL.Clear
 '  MARIZW = Val(ATRIFORM$(CODFOR$, "MARIZQMM"))
 '  MARSUW = Val(ATRIFORM$(CODFOR$, "MARSUPMM"))
   '
   MUESTRAFORBLA% = 1
'   Call CARFORWIN
   MUESTRAFORBLA% = 0
'   Call SETSPOOL("ENDDOC", "")
   '
   If CACOMM% < 1 Then
      PRICOMM$(1) = "ORIENTATION"
      PRICOMM$(2) = "NEWPAGE"
      PRICOMM$(3) = "ENDDOC"
      PRICOMM$(4) = "HEIGHT"
      PRICOMM$(5) = "WIDTH"
      PRICOMM$(6) = "SCALEMODE"
      PRICOMM$(7) = "FORECOLOR"
      PRICOMM$(8) = "DRAWMODE"
      PRICOMM$(9) = "DRAWWIDTH"
      PRICOMM$(10) = "DRAWSTYLE"
      PRICOMM$(11) = "FILLSTYLE"
      PRICOMM$(12) = "LINE"
      PRICOMM$(13) = "FONTNAME"
      PRICOMM$(14) = "FONTITALIC"
      PRICOMM$(15) = "FONTSIZE"
      PRICOMM$(16) = "FONTBOLD"
      PRICOMM$(17) = "CURRENTX"
      PRICOMM$(18) = "CURRENTY"
      PRICOMM$(19) = "PRINT"
      PRICOMM$(20) = "FONTUNDERLINE"
      PRICOMM$(21) = "PAINTPICTURE"
      CACOMM% = 21
   End If
   '
10 TPSP$ = ""
   For KKII% = 1 To FORMALTER.PSPOOL.ListCount
      TTTT$ = TRIM$(FORMALTER.PSPOOL.List(KKII% - 1))
      PCOMM$ = TRIM$(Left$(TTTT$, 16))
      PATRI$ = Mid$(TTTT$, 17)
      For KKJJ% = 1 To CACOMM%
        If PRICOMM$(KKJJ%) = PCOMM$ Then
           TPSP$ = TPSP$ + Chr$(KKJJ%) + PATRI$ + Chr$(255)
           GoTo 19
        End If
      Next KKJJ%
19 Next KKII%
   '
   Picture1.Cls
   Picture1.Top = 0
   Picture1.Refresh
   Picture1.Height = 7070
   Picture1.Width = 11700
   Picture1.DrawWidth = 3
   Picture1.DrawStyle = 0
   Picture1.FillStyle = 1
   '
   ALTUPAGINA = 7070
   TOPEPAGINA = 0
   '
   SPOOLSTRING = TPSP$
   '
20 PPXX% = InStr(TPSP$, Chr$(255))
   If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        If COPRI% = 2 Or COPRI% = 3 Then GoTo 30
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call SHOWCOMM(COPRI%, ATRIP$)
      End If
      GoTo 20
      '
   End If
   '
30 Screen.MousePointer = 1
   ANCHPAGINA = Picture1.Width
   Picture1.Refresh
   VScroll1.Min = 0
   VScroll1.Value = 0
   VScroll1.Max = 1
   If ALTUPAGINA > Frame1.Height Then
      VScroll1.Max = (ALTUPAGINA - Frame1.Height) / 100
   End If
   HScroll1.Min = 0
   HScroll1.Value = 0
   HScroll1.Max = 1
   If ANCHPAGINA > Frame1.Width Then
      HScroll1.Max = (ANCHPAGINA - Frame1.Width) / 100
   End If
   '
   Picture1.Refresh
   '
'   MARIZW = Val(ATRIFORM$(CODFOR$, "MARIZQMM"))
'   MARSUW = Val(ATRIFORM$(CODFOR$, "MARSUPMM"))
'   DETDESMM = Val(ATRIFORM$(CODFOR$, "DETDESMM"))
'   DETHASMM = Val(ATRIFORM$(CODFOR$, "DETHASMM"))
   If DETHASMM < DETDESMM + 4 + INTERLI Then
      DETHASMM = DETDESMM + 4 * INTERLI
   End If
   Picture1.ForeColor = RGB(0, 100, 0)
   Picture1.Line (56.7 * (MARIZW + 1), 56.7 * (DETDESMM + MARSUW))-(EXTREDET - 56.7, 56.7 * (DETHASMM + MARSUW)), RGB(200, 255, 255), BF
'   INTERX$ = TRIM$(ATRIFORM$(CODFOR$, "INTERLIN"))
   INTERLI = Val(INTERX$)
   If INTERLI < 1 Then INTERLI = 4
   '
   AFORMX$ = LOADFILE$(LUDAT$ + CODFOR$ + ".FRM")
   LNXP& = 1 + Len(AFORMX$) / 128
   For KU& = 1 To LNXP&
     RFP$ = Mid$(AFORMX$, 128 * KU& - 127, 128)
200  FUENX$ = TRIM$(Mid$(RFP$, 87, 32))
     FORMA1$ = TRIM$(Mid$(RFP$, 75, 6))
     FORMA2$ = TRIM$(Mid$(RFP$, 81, 6))
     If TRIM$(FORMA1$ + FORMA2$) <> "" Then
       On Error Resume Next
       Picture1.FontName = FUENX$
       If Err Then
         Picture1.FontName = "Courier"
       End If
       On Error GoTo 0
       On Error Resume Next
       Picture1.FontSize = Str$(CVS(Mid$(RFP$, 119, 4)))
       On Error GoTo 0
       Picture1.FontBold = False
       If Asc(Mid$(RFP$, 123, 1)) > 32 Then
         Picture1.FontBold = True
       End If
       Picture1.FontItalic = False
       If Asc(Mid$(RFP$, 124, 1)) > 32 Then
         Picture1.FontItalic = True
       End If
       Picture1.FontUnderline = False
       If Asc(Mid$(RFP$, 125, 1)) > 32 Then
         Picture1.FontUnderline = True
       End If
       '
       PH1 = 56.7 * (CVS(Mid$(RFP$, 57, 4)) + MARIZW + X0)
       PV1 = 56.7 * (CVS(Mid$(RFP$, 61, 4)) + MARSUW + Y0)
       FORMA1$ = TRIM$(Mid$(RFP$, 75, 6))
       '
       PH2 = 56.7 * (CVS(Mid$(RFP$, 65, 4)) + MARIZW + X0)
       PV2 = 56.7 * (CVS(Mid$(RFP$, 69, 4)) + MARSUW + Y0)
       FORMA2$ = TRIM$(Mid$(RFP$, 81, 6))
       '
       If FORMA1$ <> "" Then
         TTXX$ = TRIM$(Left$(RFP$, 8))
         AMAXI% = Val(Mid$(FORMA1$, 2))
         TTXX$ = Left$(TTXX$, AMAXI%)
         If Left$(FORMA1$, 1) = "A" Or Left$(FORMA1$, 1) = "C" Or Left$(FORMA1$, 1) = "D" Then
             TTXX$ = Left$(TTXX$ + String$(256, "_"), AMAXI%)
           ElseIf Left$(FORMA1$, 1) <> "X" Then
             While Len(TTXX$) < AMAXI%
               TTXX$ = "_" + TTXX$
             Wend
         End If
         If UCase$(Left$(TRIM$(FORMA1$), 1) = "X") Then
           TTXX$ = TTXX$ + CRLF$ + TRIM$(Mid$(RFP$, 9, 40))
           PPXXZZ% = InStr(FORMA1$, "/"): If PPXXZZ% < 1 Then PPXXZZ% = 1 + Len(FORMA1$)
           CAREMA% = Val(Mid$(FORMA1$, 1 + PPXXZZ%)): If CAREMA% < 1 Then CAREMA% = 1
           INTERLIX = INTERLI: If INTERLIX / CAREMA% < 3 Then INTERLIX = INTERLI * (1 + CAREMA%)
           Call FRATEXTO(TTXX$, AMAXI%)
           For CAREI% = 1 To CARETEX%
             If CAREI% <= CAREMA% Then
               Picture1.ForeColor = RGB(255, 0, 0)
               Picture1.CurrentX = PH1
               Picture1.CurrentY = PV1
               If Picture1.Height < Picture1.CurrentY + 1000 Then
                  Picture1.Height = Picture1.CurrentY + 1000
               End If
               If CAREI% = 1 Then
                 If PV1 <= (MARSUW + 4) * 56.7 Then
                   PV1 = PV1 + 56.7 * DETDESMM
                 End If
               End If
               TTYY$ = Left$(RETEX$(CAREI%) + String$(256, "_"), AMAXI%)
               Picture1.Print TTYY$
               PV1 = PV1 + INTERLIX / (CAREMA% + 1) * 56.7
             End If
           Next CAREI%
           GoTo 880
         End If
         If OFSECOPIA% > 0 Then
           TTXY$ = TTXX$
           While Len(TTXX$) < OFSECOPIA%
             TTXX$ = TTXX$ + " "
           Wend
           TTXX$ = RTrim$(TTXX$ + TTXY$)
         End If
         If TRIM$(TTXX$) <> "" Then
           If PV1 <= (MARSUW + 4) * 56.7 Then
             PV1 = PV1 + 56.7 * DETDESMM
           End If
           Picture1.CurrentX = PH1
           Picture1.CurrentY = PV1
           If Picture1.Height < Picture1.CurrentY + 1000 Then
              Picture1.Height = Picture1.CurrentY + 1000
           End If
           If TRIM$(TTXX$) <> "" Then
              Picture1.ForeColor = RGB(255, 0, 0)
              Picture1.Print TTXX$
           End If
         End If
       End If
       '
880    If FORMA2$ <> "" Then
         TTXX$ = TRIM$(Left$(RFP$, 8))
         AMAXI% = Val(Mid$(FORMA2$, 2))
         TTXX$ = Left$(TTXX$, AMAXI%)
         If Left$(FORMA2$, 1) = "A" Or Left$(FORMA2$, 1) = "C" Or Left$(FORMA2$, 1) = "D" Then
             TTXX$ = Left$(TTXX$ + String$(256, "_"), AMAXI%)
           ElseIf Left$(FORMA2$, 1) <> "X" Then
             While Len(TTXX$) < AMAXI%
               TTXX$ = "_" + TTXX$
             Wend
         End If
         If UCase$(Left$(TRIM$(FORMA2$), 1) = "X") Then
           TTXX$ = TTXX$ + CRLF$ + TRIM$(Mid$(RFP$, 9, 40))
           PPXXZZ% = InStr(FORMA2$, "/"): If PPXXZZ% < 1 Then PPXXZZ% = 1 + Len(FORMA2$)
           CAREMA% = Val(Mid$(FORMA2$, 1 + PPXXZZ%)): If CAREMA% < 1 Then CAREMA% = 1
           INTERLIX = INTERLI: If INTERLIX / CAREMA% < 3 Then INTERLIX = INTERLI * (1 + CAREMA%)
           Call FRATEXTO(TTXX$, AMAXI%)
           For CAREI% = 1 To CARETEX%
             If CAREI% <= CAREMA% Then
               Picture1.ForeColor = RGB(255, 0, 0)
               Picture1.CurrentX = PH2
               Picture1.CurrentY = PV2
               If Picture1.Height < Picture1.CurrentY + 1000 Then
                  Picture1.Height = Picture1.CurrentY + 1000
               End If
               If CAREI% = 1 Then
                 If PV2 <= (MARSUW + 4) * 56.7 Then
                   PV2 = PV2 + 56.7 * DETDESMM
                 End If
               End If
               TTYY$ = Left$(RETEX$(CAREI%) + String$(256, "_"), AMAXI%)
               Picture1.Print TTYY$
               PV2 = PV2 + INTERLIX / (CAREMA% + 1) * 56.7
             End If
           Next CAREI%
           GoTo 890
         End If
         If OFSECOPIA% > 0 Then
           TTXY$ = TTXX$
           While Len(TTXX$) < OFSECOPIA%
             TTXX$ = TTXX$ + " "
           Wend
           TTXX$ = RTrim$(TTXX$ + TTXY$)
         End If
         If TRIM$(TTXX$) <> "" Then
           If PV2 <= (MARSUW + 4) * 56.7 Then
             PV2 = PV2 + 56.7 * DETDESMM
           End If
           Picture1.CurrentX = PH2
           Picture1.CurrentY = PV2
           If Picture1.Height < Picture1.CurrentY + 1000 Then
              Picture1.Height = Picture1.CurrentY + 1000
           End If
           If TRIM$(TTXX$) <> "" Then
              Picture1.ForeColor = RGB(255, 0, 0)
              Picture1.Print TTXX$
           End If
         End If
       End If
       '
890 End If
     '
   Next KU&
   '
   If ALTUPAGINA < Picture1.Height Then
     ALTUPAGINA = Picture1.Height
   End If
   If ALTUPAGINA > Frame1.Height Then
      VScroll1.Max = (ALTUPAGINA - Frame1.Height) / 100
   End If
End Sub

Sub SHOWCOMM(COPRI%, ATRIP$)
   '
   If PAGINACTIVA% > CAPAGINAS% Then
      CAPAGINAS% = PAGINACTIVA%
   End If
10 On Error Resume Next
   Select Case COPRI%
   Case 1
     '
   Case 2
     PAGINACTIVA% = PAGINACTIVA% + 1
   Case 3
     PAGINACTIVA% = PAGINACTIVA% + 1
   Case 4
     '
   Case 5
     Picture1.Width = Val(TRIM$(ATRIP$))
   Case 6
     Picture1.ScaleMode = Val(TRIM$(ATRIP$))
   Case 7
     Picture1.ForeColor = Val(TRIM$(ATRIP$))
   Case 8
     Picture1.DrawMode = Val(TRIM$(ATRIP$))
   Case 9
     Picture1.DrawWidth = Val(TRIM$(ATRIP$))
   Case 10
     Picture1.DrawStyle = Val(TRIM$(ATRIP$))
   Case 11
     Picture1.FillStyle = Val(TRIM$(ATRIP$))
   Case 12
     AA11 = Val(Mid$(ATRIP$, 1, 16))
     BB11 = Val(Mid$(ATRIP$, 17, 16))
     AA22 = Val(Mid$(ATRIP$, 33, 16))
     BB22 = Val(Mid$(ATRIP$, 49, 16))
     CCOLO = Val(Mid$(ATRIP$, 65, 16))
     If BB11 + 200 > ALTUPAGINA Then ALTUPAGINA = BB11 + 200
     If BB22 + 200 > ALTUPAGINA Then ALTUPAGINA = BB22 + 200
     If Picture1.Height < ALTUPAGINA - 1 Then
       Picture1.Height = ALTUPAGINA
     End If
     If AA11 > Picture1.Width - 300 Then
        Picture1.Width = AA11 + 300
     End If
     If AA22 > Picture1.Width - 300 Then
        Picture1.Width = AA22 + 300
     End If
     BBLO$ = UCase$(TRIM$(Mid$(ATRIP$, 81, 16)))
     If BBLO$ = "B" Then
        Picture1.Line (AA11, BB11 + TOPEPAGINA)-(AA22, BB22 + TOPEPAGINA), CCOLO, B
       Else
        Picture1.Line (AA11, BB11 + TOPEPAGINA)-(AA22, BB22 + TOPEPAGINA), CCOLO
     End If
   Case 13
     Picture1.FontName = TRIM$(ATRIP$)
   Case 14
     Picture1.FontItalic = Val(TRIM$(ATRIP$))
   Case 15
     Picture1.FontSize = Val(TRIM$(ATRIP$))
   Case 16
     Picture1.FontBold = Val(TRIM$(ATRIP$))
   Case 17
     If Val(TRIM$(ATRIP$)) > Picture1.Width - 300 Then
        Picture1.Width = Val(TRIM$(ATRIP$)) + 300
     End If
     Picture1.CurrentX = Val(TRIM$(ATRIP$))
   Case 18
     If Val(TRIM$(ATRIP$)) + 200 > ALTUPAGINA Then ALTUPAGINA = Val(TRIM$(ATRIP$)) + 1000
       If Picture1.Height < ALTUPAGINA - 1 Then
         Picture1.Height = ALTUPAGINA
       End If
     Picture1.CurrentY = Val(TRIM$(ATRIP$)) + TOPEPAGINA
   Case 19
     ATRIP$ = RTrim$(ATRIP$)
     If Right$(ATRIP$, 1) = ";" Then
         ATRIP$ = Left$(ATRIP$, Len(ATRIP$) - 1)
         Picture1.Print ATRIP$;
        Else
         Picture1.Print ATRIP$
     End If
   Case 20
     Picture1.FontUnderline = Val(TRIM$(ATRIP$))
   Case 21
     REDUFORMU = 1
     Call INCLUDEIMG(ATRIP$, "SHOW")
   End Select
   On Error GoTo 0
   '
End Sub

Private Sub Text3_Change()
   '
   If Text3.Visible = False Then Exit Sub
   '
   Text3.Font = Combo1.TEXT
   On Error Resume Next
   Text3.FontSize = XVALO(Text8)
   On Error GoTo 0
   If Text3.Visible = True Then
       Text10 = Text3
       Call BORRATRASTEX
   End If
   '
   Picture1.CurrentY = Text3.Top + 50
   Picture1.CurrentX = Text3.Left + 50
   Picture1.Font = Text3.Font
   Picture1.FontSize = Text3.FontSize
   Picture1.FontBold = Text3.FontBold
   Picture1.FontItalic = Text3.FontItalic
   Picture1.FontUnderline = Text3.FontUnderline
      FCOL = Picture1.ForeColor
   Picture1.ForeColor = Text3.ForeColor
   '
   TPXX$ = "": Label17 = ""
   For PINI% = 1 To Len(Text3.TEXT)
     If Label17.Width < Text3.Width Then
       TPXX$ = TPXX$ + Mid$(Text3.TEXT, PINI%, 1)
       Label17 = TPXX$ + Mid$(Text3.TEXT, PINI% + 1, 1)
     End If
   Next PINI%
   '
   IOBJ% = XVALO(Label14)
     If MULIN(IOBJ%) <> 1 Then
       Picture1.Print TPXX$          ' Text3.TEXT
     End If
   Picture1.ForeColor = FCOL     ' PARA RECUPERAR EL COLOR
   '
   Label17.Font = Text3.Font
   Label17.FontSize = Text3.FontSize
   Label17.FontBold = Text3.FontBold
   Label17.FontItalic = Text3.FontItalic
   Label17.FontUnderline = Text3.FontUnderline
   '
   IOBJ% = XVALO(Label14)
   Label17 = RTrim(Text3.TEXT) + "*M"
   If TIOBJ$(IOBJ%) = "PARAM" Or TIOBJ$(IOBJ%) = "TEXT" Then
      Label17 = RTrim(Text3.TEXT)
      GoTo 15
   End If
   '
   If Label17.Width > Text3.Width Then
15   If Text3 <> "" Then
       If MULIN%(IOBJ%) <> 1 Then
         Text3.Width = Label17.Width + 250
         Text5 = FORMATNUM$((Text3.Left + Text3.Width) / 56.7, "F7.1")
       End If
     End If
   End If
   If MULIN(IOBJ%) <> 1 Then
        TFS = Text3.FontSize
        Text3.FontSize = 1
        Text3.Height = Label17.Height ' - 51
        Text3.FontSize = TFS
      'Else
      '  Text3.Height = (V2(IOBJ%) - V1(IOBJ%)) / 56.7
   End If
   V2(IOBJ%) = V1(IOBJ%) + Text3.Height / 56.7
   Text7 = FORMATNUM$(V2(IOBJ%), "F7.1")
   '
   Call GUARDAELOBJETO
   '
End Sub

Sub BORRATRASTEX()
   If Text3.Visible = True Then
      Picture1.Line (Text3.Left, Text3.Top)-(Text3.Left + Text3.Width, Text3.Top + Text3.Height), RGB(255, 255, 255), BF
   End If
End Sub

Sub BORRAOBJ(NOBJ%)
   OFXX = 1
   Picture1.Line (56.7 * (H1(NOBJ%) - OFXX), 56.7 * (V1(NOBJ%) - OFXX))-(56.7 * (H2(NOBJ%) + 0), 56.7 * (V2(NOBJ%) + 0)), RGB(255, 255, 255), BF
End Sub


Private Sub Text3_DblClick()
   If Command13.Visible = True Then
      Call Command13_Click
   End If
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
   '
   If KeyCode >= 37 Then
     If KeyCode <= 40 Then
       Picture1.SetFocus
       Exit Sub
     End If
   End If
   '
   If KeyCode = 46 Then     ' TECLA SURIMIR
     If TIOBJ$(IOBJ%) <> "LABEL" Then
       Call Command15_Click
       Exit Sub
     End If
   End If
   '
   If TIOBJ$(IOBJ%) = "PARAM" Or TIOBJ$(IOBJ%) = "TEXT" Then
      KeyCode = 0
   End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   If TIOBJ$(IOBJ%) = "PARAM" Or TIOBJ$(IOBJ%) = "TEXT" Then
      KeyAscii = 0
   End If
End Sub
'
Private Sub Text3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Picture1.Line (XORI, YORI)-(XMOUSE, YMOUSE), RGB(255, 255, 255), B
   Picture1.Line (Text3.Left, Text3.Top)-(Text3.Left + Text3.Width, Text3.Top + Text3.Height), RGB(255, 255, 255), BF
   '
   ESTAPRETADO% = 2
   XORI = X: YORI = Y
   If XORI >= Text3.Width - 100 Then
      ESTAPRETADO% = 3
   End If
End Sub

Private Sub Text3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If ESTAPRETADO% = 2 Then
        Text3.Top = Text3.Top + Y - YORI
        Text3.Left = Text3.Left + X - XORI
      ElseIf ESTAPRETADO% = 3 Then
        Text3.Width = Text3.Width + X - XORI
        XORI = X
        IOBJ% = Label14
        If MULIN%(IOBJ%) = 1 Then
           TFS = Text3.FontSize
           Text3.FontSize = 1
           Text3.Height = Text3.Height + Y - YORI
           Text3.FontSize = TFS
           YORI = Y
        End If
   End If
   '
   If ESTAPRETADO% > 1 Then
      Text4 = FORMATNUM$((Text3.Left + 50) / 56.7, "F7.1")
      Text5 = FORMATNUM$((Text3.Left + 50 + Text3.Width) / 56.7, "F7.1")
      Text6 = FORMATNUM$((Text3.Top + 50) / 56.7, "F7.1")
      Text7 = FORMATNUM$((Text3.Top + 50 + Text3.Height) / 56.7, "F7.1")
   End If
   '
End Sub

Private Sub Text3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Text3_Change
   ESTAPRETADO = 0
End Sub

Private Sub Text4_Change()
   If PORCOORDE% = 1 Then
      IOBJ% = XVALO(Label14)
      Call BORRAOBJ%(IOBJ%)
      H1(IOBJ%) = XVALO(Text4)
      If Text3.Visible = True Then
          Call BORRATRASTEX
          Text3.Left = 56.7 * XVALO(Text4)
          Call Text3_Change
        Else
          Call DIBUOBJ(IOBJ%)
      End If
   End If
   H1(IOBJ%) = XVALO(Text4)
End Sub

Private Sub Text5_Change()
   If PORCOORDE% = 1 Then
      IOBJ% = XVALO(Label14)
      Call BORRAOBJ%(IOBJ%)
      H2(IOBJ%) = XVALO(Text5)
      If Text3.Visible = True Then
          Call BORRATRASTEX
          Text3.Left = 56.7 * XVALO(Text5) - Text3.Width
          Call Text3_Change
        Else
          Call DIBUOBJ(IOBJ%)
      End If
   End If
   H2(IOBJ%) = XVALO(Text5)
End Sub

Private Sub Text4_GotFocus()
   PORCOORDE% = 1
End Sub

Private Sub Text5_GotFocus()
   PORCOORDE% = 1
End Sub

Private Sub Text6_Change()
   IOBJ% = XVALO(Label14)
   If PORCOORDE% = 1 Then
      Call BORRAOBJ%(IOBJ%)
      V1(IOBJ%) = XVALO(Text6)
      If Text3.Visible = True Then
          Call BORRATRASTEX
          Text3.Top = 56.7 * XVALO(Text6)
          Call Text3_Change
        Else
          Call DIBUOBJ(IOBJ%)
      End If
   End If
   V1(IOBJ%) = XVALO(Text6)
   Text14 = V2(IOBJ%) - V1(IOBJ%) & "    " & Text3.Height / 56.7
End Sub

Private Sub Text6_GotFocus()
   PORCOORDE% = 1
End Sub

Private Sub Text7_Change()
   IOBJ% = XVALO(Label14)
   If PORCOORDE% = 1 Then
      Call BORRAOBJ%(IOBJ%)
      V2(IOBJ%) = XVALO(Text7)
      If Text3.Visible = True Then
          Call BORRATRASTEX
          Text3.Top = 56.7 * XVALO(Text7) - Text3.Height
          Call Text3_Change
        Else
          Call DIBUOBJ(IOBJ%)
      End If
   End If
   V2(IOBJ%) = XVALO(Text7)
   Text14 = V2(IOBJ%) - V1(IOBJ%) & "    " & Text3.Height / 56.7
End Sub

Private Sub Text7_GotFocus()
   PORCOORDE% = 1
End Sub

Private Sub Text4_lostfocus()
   PORCOORDE% = 0
End Sub

Private Sub Text5_lostfocus()
   PORCOORDE% = 0
End Sub
'
Private Sub Text6_lostfocus()
   PORCOORDE% = 0
End Sub

Private Sub Text7_lostfocus()
   PORCOORDE% = 0
End Sub

Private Sub Text8_Change()
   Call PREDITEXTO
   If XVALO(Text8) >= 1 Then
     IOBJ% = XVALO(Label14)
     If IOBJ% <= CANOBJ% + 1 Then
       TAFUEN$(IOBJ%) = TRIM$(Text8)
       MODOI$ = UCase$(TIOBJ$(IOBJ%))   ' (UCase$(Label15)
       If MODOI$ = "LABEL" Or MODOI$ = "PARAM" Or MODOI$ = "TEXT" Or MODOI$ = "GRID" Then
           UTAFUE = TRIM$(Text8)
         ElseIf MODOI$ = "BOX" Or MODOI$ = "LINE" Or MODOI$ = "GRAPH" Or MODOI$ = "GRID" Then
           Call DIBUOBJ(IOBJ%)
       End If
     End If
   End If
End Sub

Sub PREDITEXTO()
   If Text3.Visible = True Then
     Call BORRATRASTEX
     If Combo1.TEXT <> "" Then
         Text3.Font = Combo1.TEXT
       Else
         Combo1.TEXT = Text3.Font
     End If
     On Error Resume Next
     Text3.FontSize = Text8.TEXT
     Text3.FontBold = Check1.Value
     Text3.FontItalic = Check2.Value
     Text3.FontUnderline = Check3.Value
     On Error GoTo 0
     Call Text3_Change
   End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
      Call Text8_LostFocus
   End If
End Sub

Private Sub Text8_LostFocus()
   If Text3.Visible = False Then
      Call REDIBUJA
   End If
End Sub

Private Sub Text9_Change()
   IOBJ% = XVALO(Label14)
   If IOBJ% <= CANOBJ% + 1 Then
      ONAME$(IOBJ%) = UCase$(Text9)
   End If
End Sub

Private Sub Timer1_Timer()
   Static VECES%
   Label19 = MODO$
   Label19.Refresh
   VECES% = (VECES% + 1) Mod 10
   If VECES% = 0 Then
      Call DIBUREGHOR
      Call DIBUREGVER
   End If
End Sub
'
Private Sub VScroll1_Scroll()
   Picture4.Top = -30 - VScroll1.Value
   Picture7.Top = Picture4.Top
   Call DIBUREGVER
End Sub

Private Sub Command9_Click()
      '
      Command9.Enabled = False
      '
      On Error Resume Next
      FORCONSHK.Label7 = ""
      If Err Then
         On Error GoTo 0
         Call MENSERR(24, "Falta Formulario 'FORCONSHK.FRM' en Proyecto '" + App.EXEName + "'.")
         GoTo 99
      End If
      On Error GoTo 0
      '
      FORCONSHK.Label6 = "0"
      ADIR$ = IDENTER$
      If ADIR$ = "" Then
         Call MENSERR(24, "Falta Identificador de Estacion de Trabajo IDENTER.")
         GoTo 99
      End If
      '
      With FORCONSHK
         .Caption = "Configuracion Captura de Imagenes"
         .Frame1.Caption = "Acceso al Modulo Ejecutable"
         .Label3.Caption = "Linea de Comando (Opcional):"
         '
         .Frame2.Caption = "Destino Imagen Capturada"
         .Label5.Caption = "Ruta - Carpeta"
         '
         .Text1 = VALCONTROL(0, ADIR$, "APLICFOT")
         .Text2 = VALCONTROL(0, ADIR$, "LINCOFOT")
         .Text3 = VALCONTROL(0, ADIR$, "ARCHIFOT")
         .Text4 = VALCONTROL(0, ADIR$, "TOOTIFOT")
            If .Text4 = "" Then .Text4 = "Doble Click para Capturar Imagen"
         '
         .Show 1
         If .Label7 = "OK" Then
            DESTAPLI$ = .Text1
            LINCOMMA$ = .Text2
            ARCHICON$ = .Text3
            TOOTIPTE$ = .Text4
            '
            Call GRAVALCON(0, ADIR$, "APLICFOT", DESTAPLI$)
            Call GRAVALCON(0, ADIR$, "LINCOFOT", LINCOMMA$)
            Call GRAVALCON(0, ADIR$, "ARCHIFOT", ARCHICON$)
            Call GRAVALCON(0, ADIR$, "TOOTIFOT", TOOTIPTE$)
         End If
         '
      End With
      '
99    Command9.Enabled = True
      Unload FORCONSHK
      '
End Sub

Sub SHOWCOMMZOOM(COPRI%, ATRIP$)
   '
   If PAGINACTIVA% > CAPAGINAS% Then
      CAPAGINAS% = PAGINACTIVA%
   End If
10 On Error Resume Next
   Select Case COPRI%
   Case 1
     'Printer.Orientation = Val(TRIM$(ATRIP$))
   Case 2
     PAGINACTIVA% = PAGINACTIVA% + 1
   Case 3
     PAGINACTIVA% = PAGINACTIVA% + 1
   Case 4
     'ALTUPAGINA = Val(TRIM$(ATRIP$))
   Case 5
     'Picture1.Width = Val(TRIM$(ATRIP$))
   Case 6
     Picture2.ScaleMode = Val(TRIM$(ATRIP$))
   Case 7
     Picture2.ForeColor = Val(TRIM$(ATRIP$))
   Case 8
     Picture2.DrawMode = Val(TRIM$(ATRIP$))
   Case 9
     Picture2.DrawWidth = 0.66 * Val(TRIM$(ATRIP$))
   Case 10
     Picture2.DrawStyle = Val(TRIM$(ATRIP$))
   Case 11
     Picture2.FillStyle = Val(TRIM$(ATRIP$))
   Case 12
     AA11 = Val(Mid$(ATRIP$, 1, 16)) * REDUFORMU
     BB11 = Val(Mid$(ATRIP$, 17, 16)) * REDUFORMU
     AA22 = Val(Mid$(ATRIP$, 33, 16)) * REDUFORMU
     BB22 = Val(Mid$(ATRIP$, 49, 16)) * REDUFORMU
     CCOLO = Val(Mid$(ATRIP$, 65, 16))
     'If BB11 + 1000 > ALTUPAGINA Then ALTUPAGINA = BB11 + 1000
     'If BB22 + 1000 > ALTUPAGINA Then ALTUPAGINA = BB22 + 1000
     'For KKI% = 0 To CAPAGINAS%
     '  If Picture1(KKI%).Height < ALTUPAGINA Then
     '    Picture1(KKI%).Height = ALTUPAGINA
     '  End If
     'Next KKI%
     'If AA11 > Picture1.Width Then
     '   Picture1.Width = AA11
     'End If
     'If AA22 > Picture1.Width Then
     '   Picture1.Width = AA22
     'End If
     BBLO$ = UCase$(TRIM$(Mid$(ATRIP$, 81, 16)))
     If BBLO$ = "B" Then
        Picture2.Line (AA11, BB11 + TOPEPAGINA)-(AA22, BB22 + TOPEPAGINA), CCOLO, B
       Else
        Picture2.Line (AA11, BB11 + TOPEPAGINA)-(AA22, BB22 + TOPEPAGINA), CCOLO
     End If
   Case 13
     TEPRUORI.FontName = TRIM$(ATRIP$)
   Case 14
     TEPRUORI.FontItalic = Val(TRIM$(ATRIP$))
   Case 15
     TEPRUORI.FontSize = Val(TRIM$(ATRIP$))
   Case 16
     TEPRUORI.FontBold = Val(TRIM$(ATRIP$))
   Case 17
     Picture2.CurrentX = Val(TRIM$(ATRIP$)) * REDUFORMU
   Case 18
     Picture2.CurrentY = (Val(TRIM$(ATRIP$)) + TOPEPAGINA) * REDUFORMU
   Case 19
     TEPRUORI.Caption = ATRIP$
     AMATEX = TEPRUORI.Width * REDUFORMU
     TEPRURED.Font = "Arial"
     If InStr(UCase$(TEPRUORI.FontName), "COURIER") > 0 Then
        TEPRURED.FontName = "Courier New"
     End If
     TEPRURED.FontBold = False
     TEPRURED.Caption = ATRIP$
     For KKKK = 24 To 1 Step -1
       TEPRURED.FontSize = KKKK / 2
       If TEPRURED.Width <= AMATEX Then GoTo 17
     Next KKKK
     '
17   Picture2.FontName = TEPRURED.FontName '"Arial"
     Picture2.FontSize = KKKK / 2
     Picture2.FontBold = False
     ATRIP$ = RTrim$(ATRIP$)
     If Right$(ATRIP$, 1) = ";" Then
         ATRIP$ = Left$(ATRIP$, Len(ATRIP$) - 1)
         Picture2.Print ATRIP$;
        Else
         Picture2.Print ATRIP$
     End If
   Case 20
     Picture2.FontUnderline = Val(TRIM$(ATRIP$))
   Case 21
     Call INCLUDEIMG(ATRIP$, "SHOW")
   End Select
   On Error GoTo 0
   '
End Sub
'
Sub INCLUDEIMG(ATRIP$, MODOPRE$)
   '
   BBUS% = Val(Mid$(ATRIP$, 1, 16))
   IMAGENUMBER& = Val(Mid$(ATRIP$, 17, 16))
   PHX1 = Val(Mid$(ATRIP$, 33, 16)) * REDUFORMU
   PVX1 = Val(Mid$(ATRIP$, 49, 16)) * REDUFORMU
   ACHX = Val(Mid$(ATRIP$, 65, 16)) * REDUFORMU
   ALTX = Val(Mid$(ATRIP$, 81, 16)) * REDUFORMU
   '
   RFF$ = RECFILT$("PIMGLST", 1, MKI$(BBUS%))
   For KKI% = 1 To Len(RFF$) Step 4
     RELI& = CVS(Mid$(RFF$, KKI%, 4))
     If RELI& > 0 Then
       ZZ$ = REGLEIDO$("PIMGLST", RELI&)
       If CVS(Mid$(ZZ$, 53, 4)) = IMAGENUMBER& Then GoTo 10
     End If
   Next KKI%
   If Len(RFF$) < 4 Then
      Call MENSERR(24, "Imagen no Encontrada - Imposible Incluir.")
      Exit Sub
   End If
   '
10 NOMBIMAG$ = TRIM$(Mid$(ZZ$, 3, 38))
   LARGIMAG& = CVS(Mid$(ZZ$, 41, 4))
   FECHIMAG# = CVD(Mid$(ZZ$, 45, 8))
   FirstRecord& = CVS(Mid$(ZZ$, 57, 4))
   LastRecord& = CVS(Mid$(ZZ$, 61, 4))
   '
   If FirstRecord& < 1 Or LastRecord& < FirstRecord& Or LastRecord& > ULTREG&("PIMGSPL") Then
      Call MENSERR(24, "Imagen no Encontrada - Imposible Incluir.")
      Exit Sub
   End If
   '
   If FirstRecord& > 1 Then
      XX$ = REGLEIDO$("PIMGSPL", FirstRecord& - 1)
      If CVS(Left$(XX$, 4)) = IMAGENUMBER& Then
        Call MENSERR(24, "Imagen no Encontrada - Imposible Incluir.")
        Exit Sub
      End If
   End If
   '
   If LastRecord& < ULTREG&("PIMGSPL") Then
      XX$ = REGLEIDO$("PIMGSPL", LastRecord& + 1)
      If CVS(Left$(XX$, 4)) = IMAGENUMBER& Then
        Call MENSERR(24, "Imagen no Encontrada - Imposible Incluir.")
        Exit Sub
      End If
   End If
   '
   IMAGSTRI$ = String$(508 * (1 + LastRecord& - FirstRecord&), 0)
   PINSERT& = 1
   For U& = FirstRecord& To LastRecord&
      XX$ = REGLEIDO$("PIMGSPL", U&)
      If CVS(Left$(XX$, 4)) <> IMAGENUMBER& Then
        Call MENSERR(24, "Imagen no Encontrada - Imposible Incluir.")
        Exit Sub
      End If
      Mid$(IMAGSTRI$, PINSERT&, 508) = Mid$(XX$, 5, 508)
      PINSERT& = PINSERT& + 508
   Next U&
   '
   If Len(IMAGSTRI$) < LARGIMAG& Then
      Call MENSERR(24, "Imagen no Encontrada - Imposible Incluir.")
      Exit Sub
   End If
   IMAGSTRI$ = Left$(IMAGSTRI$, LARGIMAG&)
   '
   NX% = FreeFile
   Open LUCOM$ + NOMBIMAG$ For Binary As #NX%
   Put #NX%, , IMAGSTRI$
   Close #NX%
   '
   FORMALTER.Image3.Picture = LoadPicture("")
   FORMALTER.Image3.Picture = LoadPicture(LUCOM$ + NOMBIMAG$)
   '
   HIMAGEN = FORMALTER.Image3.Picture.Width
   VIMAGEN = FORMALTER.Image3.Picture.Height
   '
   AA11 = PHX1
   BB11 = PVX1 + TOPEPAGINA * REDUFORMU
   AA22 = AA11 + ACHX
   BB22 = BB11 + ALTX
   If BB11 + 200 > ALTUPAGINA Then ALTUPAGINA = BB11 + 200
   If BB22 + 200 > ALTUPAGINA Then ALTUPAGINA = BB22 + 200
   '
   If Picture1.Height < ALTUPAGINA - 1 Then
     Picture1.Height = ALTUPAGINA
   End If
   '
   If AA11 > Picture1.Width Then
     Picture1.Width = AA11
   End If
   '
   If AA22 > Picture1.Width Then
     Picture1.Width = AA22
   End If
   '
   If REDUFORMU = 1 Then
       Picture1.PaintPicture FORMALTER.Image3.Picture, PHX1, PVX1 + TOPEPAGINA, ACHX, ALTX
     Else
       Picture2.PaintPicture FORMALTER.Image3.Picture, PHX1, PVX1 + TOPEPAGINA * REDUFORMU, ACHX, ALTX
   End If
   '
End Sub

Sub GRATITUFORM()
   '
   Screen.MousePointer = 11
   If TRIM$(Text1.TEXT) = TRIM$(CODFOR$) Then
     If TRIM$(CODFOR$) <> "" Then
       ARXYZ$ = TRIM$(CODFOR$)
       TITUFOR$ = Text2.TEXT
       NX% = FILEOPEN%(LUDAT$ + ARXYZ$ + ".FRM", 0, 128)
       Dim RF As String * 128
       '
       For i% = 1 To LOF(NX%) / 128
         Get #NX%, i%, RF$
         If TRIM$(Left$(RF$, 8)) = "TITUFORM" Then
           RFY$ = RF$
           Call REPLA(RFY$, TITUFOR$, 9, 48)
           LSet RF$ = RFY$
           Put #NX%, i%, RF$
           GoTo 10
         End If
       Next i%
       '
       i% = 1 + LOF(NX%) / 128
       RFY$ = Space$(56) + String$(72, 0)
       Call REPLA(RFY$, "TITUFORM", 1, 8)
       Call REPLA(RFY$, TITUFOR$, 9, 48)
       LSet RF$ = RFY$
       Put #NX%, i%, RF$
       '
10     Close #NX%: NX% = 0
       '
       For V& = 1 To ULTREG&("INDIFRM")
         INF$ = REGLEIDO$("INDIFRM", V&)
         If Left$(INF$, 8) = AJUSTI$(ARXYZ$, 8) Then
           If TITUFOR$ <> "" Then
             Call REPLA(INF$, TITUFOR$, 9, 48)
             Call GRAREG("INDIFRM", INF$, V&)
           End If
           GoTo 19
         End If
       Next V&
       '
       INF$ = REGBLAN$("INDIFRM")
       Call REPLA(INF$, ARXYZ$, 1, 8)
       Call REPLA(INF$, TITUFOR$, 9, 48)
       Call REGAPP("INDIFRM", INF$)
       '
     End If
   End If
   '
19 Call CIERRARCH("INDIFRM")
   Call FRESHECHO("INDIFRM")
   Screen.MousePointer = 1
   '
End Sub
'

Private Sub VScroll2_Change()
    Text3.Visible = False
    Label14 = CANOBJ% + 1 - VScroll2.Value
    MODO$ = "": Label19 = ""
End Sub

Sub ALTANUEFORM()
        '
1000    CCC$ = OBJPLA$("CODEFORMU", CCC$)
        CODFOR$ = UCase(AJUSTI$(Left$(CCC$, 8), 8))
        Call REPLA(CCC$, CODFOR$, 1, 8)
        
        If TRIM$(CODFOR$) = "" Then Exit Sub
        For UU% = 1 To Len(TRIM$(CODFOR$))
           VV% = Asc(Mid$(CODFOR$, UU%))
           If VV% < 48 Or VV% > 90 Or (VV% > 57 And VV% < 65) Then
              If VV% <> 45 Then
                 Call COMUNI("Código no Válido")
                 GoTo 1000
              End If
           End If
        Next UU%
        '
        DESFOR$ = Mid$(CCC$, 9, 48)
        If TRIM$(DESFOR$) = "" Then
            Beep
            GoTo 1000
        End If
        '
        If EXISTE%(LU$ + TRIM$(CODFOR$) + ".FRM") > 0 Or EXISTE%(LU$ + TRIM(CODFOR$) + ".SQR") Then
           COMUNI ("Imposible dar de Alta, Formulario Pre-Existente")
           GoTo 1000
         Else
           Call COPYSTRU("INDIFRM", "INDISQR")
           Call REGAPP("INDISQR", CCC$)
           Call CIERRARCH("INDISQR")
           Call FILESORT("INDISQR", "", 1, 1)
           '
           JJ& = 0: CODFORA$ = ""
           For U& = 1 To ULTREG("INDISQR")
             XX$ = REGLEIDO$("INDISQR", U&)
             If TRIM$(Left$(XX$, 8)) <> CODFORA$ Then
               JJ& = JJ& + 1
               Call GRAREG("INDISQR", XX$, JJ&)
               CODFORA$ = TRIM$(Left$(XX$, 8))
             End If
           Next U&
           Call SETULTREG("INDISQR", JJ&)
           Call FRESHECHO("INDISQR")
        End If
        '
        If ULTREG&("INDISQR") > 1 Then
           If ALTERNA%("Copia Formato Existente") = 1 Then
1060          Call SELECHO("INDISQR")
              CCC$ = VALACT1$("INDISQR")
              If Len(CCC$) < 1 Then Exit Sub
              '
              CODFOR1$ = AJUSTI$(Left$(CCC$, 8), 8)
              If EXISTE%(LUDAT$ + TRIM(CODFOR1$) + ".SQR") > 0 Then
                 If CODFOR1$ = CODFOR$ Then
                  Beep
                  GoTo 1060
                 End If
                 FileCopy LUDAT$ + TRIM$(CODFOR1$) + ".SQR ", LUDAT$ + TRIM$(CODFOR$) + ".SQR"
                Else
                 Call COMUNI("Imposible Continuar, Formulario " & CODFOR1$ & " Inexistente")
                 GoTo 1060
              End If
              
              '
           End If
        End If
        Call COMUNI("Archivo Generado Correctamente")
        Text1 = CODFOR$
        '
End Sub

Sub BLOJA(NOMHOJA$, CANCELOJ%)
   '
   CANCELOJ% = 0
   TEBLO$ = Space$(128)
   PRXX$ = TRIM$(NOMHOJA$)
   If PRXX$ = "" Then Exit Sub
   '
   On Error Resume Next
   Kill LUDAT$ + "*.SQK"
   On Error GoTo 0
   '
   USX$ = USERNAME$
   ITX$ = IDENTER$
   TRX$ = NOMTER$
   HOII = HOY(HO$)
   '
   Call REPLA(TEBLO$, PRXX$, 1, 16)
   Call REPLA(TEBLO$, USX$, 17, 24)
   Call REPLA(TEBLO$, TRX$, 45, 24)
   Call REPLA(TEBLO$, HO$, 69, 8)
   Call REPLA(TEBLO$, Time$, 79, 8)
   Call REPLA(TEBLO$, ITX$, 119, 8)
   '
   SCMA = Screen.MousePointer
11 Screen.MousePointer = 11
12 On Error GoTo 15
   Dim AA As String * 128
   If NKKK% < 1 Then
      NKKK% = FreeFile
      Open LUDAT$ + PRXX$ + ".SQK" For Random Access Read Write Lock Write As #NKKK% Len = 128
   End If
   LSet AA$ = TEBLO$
   Put #NKKK%, 1, AA$
   SOLOLEC% = 0
   On Error GoTo 0
   '
99 Screen.MousePointer = SCMA
   Exit Sub
   '
'  **************************************************
   '
15 Resume 16
16 On Error GoTo 0
   Close #NKKK%: NKKK% = 0
   '
   If NKKK% < 1 Then NKKK% = FreeFile
   Open LUDAT$ + PRXX$ + ".SQK" For Random Access Read As #NKKK% Len = 128
   Get #NKKK%, 1, AA$
   B2$ = Mid$(AA$, 1, 16)
   B4$ = Mid$(AA$, 69, 8)
   B5$ = Mid$(AA$, 79, 8)
   B6$ = Mid$(AA$, 17, 24)
   Close #NKKK%: NKKK% = 0
   '
   LOMAX% = 0
   USY$ = "La Hoja de Proceso '" + TRIM$(B2$) + "'"
   TRY$ = "Está Abierta en Modo de Edición"
   FFY$ = "Desde el " + TRIM$(B4$) + " a las " + TRIM$(B5$) + " horas"
   If TRIM$(B6$) <> "" Then FFY$ = FFY$ + "\por el Usuario " + TRIM$(B6$)
   '
   COMU$ = USY$ + "\" + TRY$ + "\" + FFY$ + "\  \Puede Abrir en Modo de Solo Lectura o Cancelar.\  \En Solo Lectura no Podrá Guardar Modificaciones."
   Screen.MousePointer = SCMA
   If COMALTER%(COMU$ + "\\Solo Lectura\Cancelar") = 1 Then
        SOLOLEC% = 1
      Else
        Text1 = ""
        Call Text1_Change
        Text3.Visible = False
        Text3 = ""
        Picture1.Cls
        CANCELOJ% = 1
   End If
   '
End Sub
