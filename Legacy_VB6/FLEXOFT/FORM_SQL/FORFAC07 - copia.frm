VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FORFAC07 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturación de Proveedores (Operaciones Comerciales de Compras)"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   142
      Top             =   1320
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0FFD0&
      BorderStyle     =   0  'None
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
      TabIndex        =   6
      Top             =   0
      Width           =   12000
      Begin VB.Frame Frame17 
         BackColor       =   &H00E0FFD0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   210
         TabIndex        =   135
         Top             =   5760
         Width           =   3315
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   1550
            TabIndex        =   137
            Top             =   150
            Width           =   1485
         End
         Begin VB.Label Label42 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Imp. Total:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   -360
            TabIndex        =   136
            Top             =   150
            Width           =   1680
         End
      End
      Begin VB.Frame Frame16 
         BackColor       =   &H00E0FFD0&
         Caption         =   "Redondeo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Left            =   9345
         TabIndex        =   132
         Top             =   5145
         Width           =   1425
         Begin VB.TextBox Text22 
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
            Left            =   210
            TabIndex        =   133
            Text            =   " "
            ToolTipText     =   "Importe de Redondeo"
            Top             =   630
            Width           =   1065
         End
         Begin VB.Label Label41 
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
            Left            =   0
            TabIndex        =   134
            Top             =   380
            Width           =   960
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00E0FFD0&
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   5040
         ScaleHeight     =   540
         ScaleWidth      =   960
         TabIndex        =   130
         Top             =   1050
         Visible         =   0   'False
         Width           =   960
         Begin VB.Label Label40 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   420
            TabIndex        =   131
            Top             =   105
            Width           =   450
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   0
            Picture         =   "FORFAC07.frx":0000
            Top             =   0
            Width           =   480
         End
      End
      Begin VB.CheckBox Check2 
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
         Left            =   5040
         TabIndex        =   120
         Top             =   630
         Width           =   200
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00E0FFD0&
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
         Height          =   1335
         Left            =   210
         TabIndex        =   118
         Top             =   3255
         Width           =   6480
         Begin VB.TextBox Text19 
            Alignment       =   1  'Right Justify
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
            Left            =   1860
            TabIndex        =   121
            Top             =   580
            Width           =   1365
         End
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
            Height          =   285
            Left            =   615
            TabIndex        =   70
            ToolTipText     =   "Ingrese Observaciones"
            Top             =   945
            Width           =   5760
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Obs:"
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
            TabIndex        =   129
            Top             =   970
            Width           =   1140
         End
         Begin VB.Label Label38 
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
            TabIndex        =   128
            Top             =   315
            Width           =   1260
         End
         Begin VB.Label Label37 
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
            TabIndex        =   127
            Top             =   315
            Width           =   1140
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   315
            TabIndex        =   126
            Top             =   580
            Width           =   1455
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   3360
            TabIndex        =   125
            Top             =   580
            Width           =   1455
         End
         Begin VB.Label Label34 
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
            TabIndex        =   124
            Top             =   315
            Width           =   1140
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   4920
            TabIndex        =   123
            Top             =   580
            Width           =   1455
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
            TabIndex        =   122
            Top             =   315
            Width           =   1260
         End
      End
      Begin VB.TextBox Text17 
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
         Left            =   3885
         TabIndex        =   109
         Top             =   1230
         Width           =   960
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
         Height          =   280
         Left            =   1365
         TabIndex        =   107
         Top             =   1230
         Width           =   2220
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00004000&
         Height          =   8200
         Left            =   10920
         ScaleHeight     =   8145
         ScaleWidth      =   1440
         TabIndex        =   92
         Top             =   0
         Width           =   1500
         Begin VB.CommandButton Command11 
            Caption         =   "A-B-M"
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
            Left            =   160
            Picture         =   "FORFAC07.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   96
            ToolTipText     =   "Archivo Maestro de Proveedores"
            Top             =   2270
            Width           =   650
         End
         Begin VB.PictureBox MARCOBARRA 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   160
            ScaleHeight     =   75
            ScaleWidth      =   585
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   3780
            Width           =   640
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   -420
               TabIndex        =   103
               Top             =   0
               Width           =   12000
            End
         End
         Begin VB.CommandButton Command20 
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
            Left            =   160
            Picture         =   "FORFAC07.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   101
            ToolTipText     =   "Terminar - Salir de la Aplicación"
            Top             =   110
            Width           =   650
         End
         Begin VB.CommandButton Command29 
            Caption         =   "Help"
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
            Left            =   160
            Picture         =   "FORFAC07.frx":075E
            Style           =   1  'Graphical
            TabIndex        =   100
            ToolTipText     =   "Ayuda Flexoft en Línea"
            Top             =   800
            Width           =   650
         End
         Begin VB.CommandButton BOTNEXT 
            Caption         =   "Save"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   160
            Picture         =   "FORFAC07.frx":0A68
            Style           =   1  'Graphical
            TabIndex        =   99
            ToolTipText     =   "Grabar - Registrar Documento en la Base de Datos"
            Top             =   4980
            Width           =   650
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Note"
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
            Left            =   160
            Picture         =   "FORFAC07.frx":0D72
            Style           =   1  'Graphical
            TabIndex        =   98
            ToolTipText     =   "Notas de Proveedores"
            Top             =   2940
            Width           =   650
         End
         Begin VB.CommandButton Command4 
            Caption         =   "R-Cue"
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
            Left            =   160
            Picture         =   "FORFAC07.frx":107C
            Style           =   1  'Graphical
            TabIndex        =   97
            ToolTipText     =   "Resumen de Cuenta Corriente Proveedores"
            Top             =   1580
            Width           =   650
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Print"
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
            Left            =   160
            Picture         =   "FORFAC07.frx":1386
            Style           =   1  'Graphical
            TabIndex        =   95
            ToolTipText     =   "Imprime Formulario de Carga - Vuelco de Pantalla"
            Top             =   7270
            Width           =   650
         End
         Begin VB.CommandButton Command6 
            Caption         =   "More"
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
            Left            =   160
            Picture         =   "FORFAC07.frx":19F0
            Style           =   1  'Graphical
            TabIndex        =   94
            ToolTipText     =   "Anular, Consultar, Imputación de Créditos, Tablas"
            Top             =   6615
            Width           =   650
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
            Height          =   690
            Left            =   160
            Picture         =   "FORFAC07.frx":1B3A
            Style           =   1  'Graphical
            TabIndex        =   93
            ToolTipText     =   "Opciones de Configuración"
            Top             =   5840
            Width           =   650
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00004000&
            Caption         =   "Libro I.V.A"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   160
            TabIndex        =   112
            ToolTipText     =   "Incluir / Excluir de Listado IVA Compras"
            Top             =   4480
            Value           =   1  'Checked
            Width           =   915
         End
         Begin VB.Image Image1 
            Height          =   390
            Left            =   -255
            Picture         =   "FORFAC07.frx":1F7C
            Top             =   4095
            Width           =   1515
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
         TabIndex        =   82
         Top             =   4620
         Width           =   3000
         Begin VB.CommandButton Command21 
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   210
            TabIndex        =   111
            ToolTipText     =   "Click Aqui para Cancelar Pago Contado "
            Top             =   280
            Width           =   280
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
            TabIndex        =   91
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
            TabIndex        =   90
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
            TabIndex        =   89
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
            TabIndex        =   84
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
            TabIndex        =   88
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
            TabIndex        =   87
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
            TabIndex        =   86
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
            TabIndex        =   85
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
            TabIndex        =   83
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
         Height          =   1155
         Left            =   210
         TabIndex        =   37
         Top             =   4620
         Width           =   3315
         Begin VB.PictureBox Picture6 
            BackColor       =   &H00E0FFD0&
            BorderStyle     =   0  'None
            Height          =   720
            Left            =   80
            ScaleHeight     =   720
            ScaleWidth      =   3135
            TabIndex        =   144
            Top             =   320
            Width           =   3130
            Begin VB.PictureBox Picture5 
               BackColor       =   &H00E0FFD0&
               BorderStyle     =   0  'None
               Height          =   3705
               Left            =   0
               ScaleHeight     =   3705
               ScaleWidth      =   3045
               TabIndex        =   146
               Top             =   0
               Width           =   3050
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
                  Index           =   8
                  Left            =   1200
                  TabIndex        =   147
                  Top             =   3360
                  Width           =   180
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
                  Index           =   7
                  Left            =   1200
                  TabIndex        =   148
                  Top             =   2940
                  Width           =   180
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
                  Left            =   1200
                  TabIndex        =   149
                  Top             =   2100
                  Width           =   180
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
                  Left            =   1200
                  TabIndex        =   150
                  Top             =   1680
                  Width           =   180
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
                  Left            =   1200
                  TabIndex        =   151
                  Top             =   840
                  Width           =   180
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
                  Left            =   1200
                  TabIndex        =   152
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
                  Index           =   6
                  Left            =   255
                  TabIndex        =   173
                  Top             =   2520
                  Width           =   960
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
                  Index           =   6
                  Left            =   1200
                  TabIndex        =   172
                  Top             =   2520
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
                  Index           =   7
                  Left            =   255
                  TabIndex        =   171
                  Text            =   " "
                  Top             =   2940
                  Width           =   1065
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
                  Index           =   8
                  Left            =   255
                  TabIndex        =   170
                  Text            =   " "
                  Top             =   3360
                  Width           =   1065
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
                  Index           =   6
                  Left            =   1515
                  TabIndex        =   169
                  Top             =   2520
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
                  Index           =   7
                  Left            =   1515
                  TabIndex        =   168
                  Text            =   " "
                  Top             =   2940
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
                  Index           =   8
                  Left            =   1515
                  TabIndex        =   167
                  Text            =   " "
                  Top             =   3360
                  Width           =   1485
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
                  Index           =   3
                  Left            =   255
                  TabIndex        =   166
                  Top             =   1260
                  Width           =   960
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
                  Left            =   1200
                  TabIndex        =   165
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
                  Height          =   300
                  Index           =   4
                  Left            =   255
                  TabIndex        =   164
                  Text            =   " "
                  Top             =   1680
                  Width           =   1065
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
                  Left            =   255
                  TabIndex        =   163
                  Text            =   " "
                  Top             =   2100
                  Width           =   1065
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
                  Left            =   1515
                  TabIndex        =   162
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
                  Height          =   315
                  Index           =   4
                  Left            =   1515
                  TabIndex        =   161
                  Text            =   " "
                  Top             =   1680
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
                  Left            =   1515
                  TabIndex        =   160
                  Text            =   " "
                  Top             =   2100
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
                  Index           =   0
                  Left            =   1200
                  TabIndex        =   159
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
                  Index           =   0
                  Left            =   255
                  TabIndex        =   158
                  Top             =   0
                  Width           =   960
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
                  Left            =   255
                  TabIndex        =   157
                  Text            =   " "
                  Top             =   420
                  Width           =   1065
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
                  Left            =   255
                  TabIndex        =   156
                  Text            =   " "
                  Top             =   840
                  Width           =   1065
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
                  Left            =   1515
                  TabIndex        =   155
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
                  Height          =   330
                  Index           =   1
                  Left            =   1515
                  TabIndex        =   154
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
                  Index           =   2
                  Left            =   1515
                  TabIndex        =   153
                  Text            =   " "
                  Top             =   840
                  Width           =   1485
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "7)"
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
                  Index           =   11
                  Left            =   -1320
                  TabIndex        =   182
                  Top             =   2565
                  Width           =   1590
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "8)"
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
                  Index           =   10
                  Left            =   -1320
                  TabIndex        =   181
                  Top             =   2985
                  Width           =   1590
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "9)"
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
                  Index           =   9
                  Left            =   -1320
                  TabIndex        =   180
                  Top             =   3405
                  Width           =   1590
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "4)"
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
                  Index           =   8
                  Left            =   -1320
                  TabIndex        =   179
                  Top             =   1320
                  Width           =   1590
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "5)"
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
                  Index           =   7
                  Left            =   -1320
                  TabIndex        =   178
                  Top             =   1725
                  Width           =   1590
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "6)"
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
                  Index           =   6
                  Left            =   -1320
                  TabIndex        =   177
                  Top             =   2145
                  Width           =   1590
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "1)"
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
                  Left            =   -1320
                  TabIndex        =   176
                  Top             =   45
                  Width           =   1590
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "2)"
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
                  Left            =   -1320
                  TabIndex        =   175
                  Top             =   465
                  Width           =   1590
               End
               Begin VB.Label LEFE 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "3)"
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
                  Left            =   -1320
                  TabIndex        =   174
                  Top             =   885
                  Width           =   1590
               End
            End
            Begin VB.VScrollBar VScroll1 
               Height          =   720
               LargeChange     =   7
               Left            =   3020
               TabIndex        =   145
               Top             =   0
               Width           =   120
            End
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
         Height          =   280
         Left            =   1365
         TabIndex        =   81
         TabStop         =   0   'False
         Text            =   " "
         Top             =   960
         Width           =   3480
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00E0FFD0&
         Caption         =   "Percep. I.Brutos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Left            =   6825
         TabIndex        =   47
         Top             =   5145
         Width           =   2505
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
            Height          =   330
            Left            =   1050
            TabIndex        =   69
            Text            =   " "
            ToolTipText     =   "Ingrese Importe Percepción I.B."
            Top             =   720
            Width           =   1380
         End
         Begin VB.TextBox TEXT10 
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
            Left            =   315
            TabIndex        =   67
            Text            =   " "
            ToolTipText     =   "Ingrese Provincia o Jurisdicción"
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
            Left            =   630
            TabIndex        =   68
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
            Left            =   0
            TabIndex        =   52
            Top             =   820
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
            Left            =   840
            TabIndex        =   51
            Top             =   315
            Width           =   1575
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "J."
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
            Left            =   -240
            TabIndex        =   50
            ToolTipText     =   "Haga Doble Click para Varias Jurisdicciones"
            Top             =   360
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
         TabIndex        =   46
         Top             =   4305
         Width           =   3945
         Begin VB.TextBox Text8 
            Alignment       =   1  'Right Justify
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
            TabIndex        =   66
            ToolTipText     =   "Ingrese Importe Percepción Ganancias"
            Top             =   315
            Width           =   1275
         End
         Begin VB.TextBox Text7 
            Alignment       =   1  'Right Justify
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
            TabIndex        =   65
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
            TabIndex        =   49
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
            TabIndex        =   48
            Top             =   420
            Width           =   540
         End
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
         TabIndex        =   41
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
         Height          =   2790
         Left            =   6825
         TabIndex        =   36
         Top             =   1470
         Width           =   3945
         Begin VB.TextBox Text18 
            Alignment       =   1  'Right Justify
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
            TabIndex        =   117
            ToolTipText     =   "Ingrese Cotización"
            Top             =   2380
            Width           =   1485
         End
         Begin VB.TextBox IMPOIVA 
            Alignment       =   1  'Right Justify
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
            TabIndex        =   63
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   1110
            Width           =   1275
         End
         Begin VB.TextBox IMPOIVA 
            Alignment       =   1  'Right Justify
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
            Index           =   6
            Left            =   2500
            TabIndex        =   64
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   1920
            Width           =   1275
         End
         Begin VB.TextBox IMPOGRA 
            Alignment       =   1  'Right Justify
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
            Index           =   6
            Left            =   930
            TabIndex        =   59
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   1920
            Width           =   1485
         End
         Begin VB.TextBox IMPOGRA 
            Alignment       =   1  'Right Justify
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
            TabIndex        =   58
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   1650
            Width           =   1485
         End
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
            Index           =   5
            Left            =   2520
            TabIndex        =   113
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   1840
            Visible         =   0   'False
            Width           =   1170
         End
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
            TabIndex        =   77
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   1260
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H00E0FFD0&
            BorderStyle     =   0  'None
            Height          =   1950
            Left            =   40
            TabIndex        =   72
            Top             =   360
            Width           =   850
            Begin VB.Label TASIVA 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "N/Grav"
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
               TabIndex        =   114
               Top             =   1350
               Width           =   795
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
               Index           =   6
               Left            =   0
               TabIndex        =   79
               Top             =   1620
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
               TabIndex        =   78
               Top             =   1080
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
               TabIndex        =   76
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
               TabIndex        =   75
               Top             =   270
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
               TabIndex        =   74
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
               TabIndex        =   73
               Top             =   810
               Width           =   750
            End
         End
         Begin VB.TextBox IMPOGRA 
            Alignment       =   1  'Right Justify
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
            TabIndex        =   57
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   1380
            Width           =   1485
         End
         Begin VB.TextBox IMPOGRA 
            Alignment       =   1  'Right Justify
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
            TabIndex        =   56
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   1110
            Width           =   1485
         End
         Begin VB.TextBox IMPOIVA 
            Alignment       =   1  'Right Justify
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
            TabIndex        =   62
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   860
            Width           =   1275
         End
         Begin VB.TextBox IMPOGRA 
            Alignment       =   1  'Right Justify
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
            TabIndex        =   55
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   860
            Width           =   1485
         End
         Begin VB.TextBox IMPOIVA 
            Alignment       =   1  'Right Justify
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
            TabIndex        =   61
            Text            =   " "
            ToolTipText     =   "  "
            Top             =   570
            Width           =   1275
         End
         Begin VB.TextBox IMPOGRA 
            Alignment       =   1  'Right Justify
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
            TabIndex        =   54
            ToolTipText     =   "  "
            Top             =   570
            Width           =   1485
         End
         Begin VB.TextBox IMPOIVA 
            Alignment       =   1  'Right Justify
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
            TabIndex        =   60
            ToolTipText     =   "  "
            Top             =   300
            Width           =   1275
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cotiz.:"
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
            Left            =   210
            TabIndex        =   116
            Top             =   2415
            Width           =   645
         End
         Begin VB.Line Line13 
            BorderColor     =   &H00C0C0C0&
            X1              =   0
            X2              =   3885
            Y1              =   2310
            Y2              =   2310
         End
         Begin VB.Label Label29 
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
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   2520
            TabIndex        =   115
            ToolTipText     =   "T.C.: 1.0000 - Doble-Click para Cambiar Moneda"
            Top             =   2360
            Width           =   1275
         End
         Begin VB.Label ITOGRA 
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
            Height          =   300
            Left            =   930
            TabIndex        =   71
            Top             =   315
            Width           =   1485
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E0E0E0&
         Height          =   1575
         Left            =   210
         ScaleHeight     =   1515
         ScaleWidth      =   6405
         TabIndex        =   25
         Top             =   1585
         Width           =   6465
         Begin VB.Frame Frame4 
            BackColor       =   &H00C0C000&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   0
            TabIndex        =   26
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
               TabIndex        =   27
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
            Height          =   2150
            Left            =   -315
            TabIndex        =   28
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
               Height          =   22140
               IntegralHeight  =   0   'False
               Left            =   285
               TabIndex        =   29
               Top             =   315
               Width           =   6465
            End
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00E0E0E0&
         Height          =   1580
         Left            =   210
         ScaleHeight     =   1515
         ScaleWidth      =   10470
         TabIndex        =   15
         Top             =   6400
         Width           =   10530
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0C000&
            BorderStyle     =   0  'None
            Height          =   5630
            Left            =   0
            TabIndex        =   16
            Top             =   260
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
               Height          =   935
               IntegralHeight  =   0   'False
               Left            =   -30
               TabIndex        =   17
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
               Height          =   935
               IntegralHeight  =   0   'False
               Left            =   -30
               TabIndex        =   38
               Top             =   0
               Width           =   10455
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
               Left            =   8400
               TabIndex        =   24
               Top             =   980
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
               TabIndex        =   23
               Top             =   0
               Visible         =   0   'False
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
               Left            =   6615
               TabIndex        =   20
               Top             =   980
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
               TabIndex        =   19
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
               Left            =   105
               TabIndex        =   18
               Top             =   980
               Width           =   1695
            End
            Begin VB.Line Line14 
               X1              =   0
               X2              =   10500
               Y1              =   935
               Y2              =   935
            End
            Begin VB.Line Line15 
               X1              =   0
               X2              =   10500
               Y1              =   1250
               Y2              =   1250
            End
            Begin VB.Line Line16 
               X1              =   2520
               X2              =   2520
               Y1              =   735
               Y2              =   2415
            End
            Begin VB.Line Line17 
               X1              =   3735
               X2              =   3735
               Y1              =   735
               Y2              =   2415
            End
            Begin VB.Line Line18 
               X1              =   5460
               X2              =   5460
               Y1              =   630
               Y2              =   2415
            End
            Begin VB.Line Line19 
               X1              =   6720
               X2              =   6720
               Y1              =   735
               Y2              =   2415
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0C000&
            BorderStyle     =   0  'None
            Height          =   250
            Left            =   0
            TabIndex        =   21
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
               Height          =   250
               Left            =   0
               TabIndex        =   22
               Top             =   0
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
         Top             =   1100
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
         Height          =   280
         Left            =   2205
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   " "
         Top             =   690
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
         Height          =   280
         Left            =   1365
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   " "
         Top             =   690
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
         Height          =   280
         Left            =   1365
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   " "
         Top             =   420
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
         Height          =   285
         Left            =   1365
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " "
         Top             =   150
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
         Left            =   1035
         TabIndex        =   4
         Top             =   540
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
         Top             =   540
         Width           =   855
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0FFD0&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   5950
         TabIndex        =   32
         Top             =   580
         Width           =   4935
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFD0&
            Caption         =   "DP"
            Height          =   240
            Index           =   7
            Left            =   4020
            TabIndex        =   106
            ToolTipText     =   "Débito Propio"
            Top             =   230
            Width           =   700
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFD0&
            Caption         =   "CP"
            Height          =   240
            Index           =   6
            Left            =   2820
            TabIndex        =   105
            ToolTipText     =   "Crédito Propio"
            Top             =   230
            Width           =   700
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFD0&
            Caption         =   "CO"
            Height          =   240
            Index           =   5
            Left            =   105
            TabIndex        =   53
            ToolTipText     =   "Comprobante Genérico (Extracto Bancario - Póliza de Seguro - Liquidación)"
            Top             =   30
            Width           =   705
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFD0&
            Caption         =   "RH"
            Height          =   240
            Index           =   4
            Left            =   420
            TabIndex        =   40
            ToolTipText     =   "Recibo de Honorarios"
            Top             =   230
            Width           =   700
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFD0&
            Caption         =   "TF"
            Height          =   240
            Index           =   3
            Left            =   1305
            TabIndex        =   39
            ToolTipText     =   "Ticket - Factura"
            Top             =   30
            Width           =   700
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFD0&
            Caption         =   "NC"
            Height          =   240
            Index           =   2
            Left            =   3705
            TabIndex        =   35
            ToolTipText     =   "Nota de Crédito del Proveedor"
            Top             =   30
            Width           =   700
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFD0&
            Caption         =   "ND"
            Height          =   240
            Index           =   1
            Left            =   2505
            TabIndex        =   34
            ToolTipText     =   "Nota de Débito del Proveedor"
            Top             =   30
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
            Left            =   1620
            TabIndex        =   33
            ToolTipText     =   "Factura de Compra"
            Top             =   230
            Value           =   -1  'True
            Width           =   700
         End
      End
      Begin VB.Label Label46 
         Height          =   225
         Left            =   795
         TabIndex        =   141
         Top             =   270
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Label44 
         Height          =   225
         Left            =   750
         TabIndex        =   139
         Top             =   30
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "C. Imp:"
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
         Index           =   7
         Left            =   0
         TabIndex        =   138
         Top             =   30
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Carga por Detalle"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5295
         TabIndex        =   119
         Top             =   630
         Width           =   855
      End
      Begin VB.Line Line21 
         BorderColor     =   &H00808080&
         X1              =   210
         X2              =   6680
         Y1              =   3160
         Y2              =   3160
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5145
         TabIndex        =   80
         ToolTipText     =   "Click Para Cambiar Categoria IVA"
         Top             =   150
         Width           =   660
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   3690
         TabIndex        =   110
         Top             =   1260
         Width           =   330
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "C.A.I.:"
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
         Left            =   690
         TabIndex        =   108
         Top             =   1290
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
         Height          =   285
         Left            =   210
         TabIndex        =   104
         Top             =   900
         Width           =   1035
      End
      Begin VB.Line Line22 
         X1              =   0
         X2              =   12000
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00808080&
         X1              =   210
         X2              =   6650
         Y1              =   3850
         Y2              =   3850
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00808080&
         X1              =   6675
         X2              =   6675
         Y1              =   1575
         Y2              =   3160
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00808080&
         X1              =   10740
         X2              =   10740
         Y1              =   6405
         Y2              =   7960
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
         TabIndex        =   42
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
         TabIndex        =   31
         Top             =   6500
         Visible         =   0   'False
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
         TabIndex        =   30
         Top             =   6500
         Visible         =   0   'False
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
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   4830
         TabIndex        =   14
         Top             =   1100
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
         TabIndex        =   13
         Top             =   120
         Width           =   960
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
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
         TabIndex        =   7
         Top             =   270
         Width           =   960
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "FORFAC07.frx":3E9E
      TabIndex        =   143
      Top             =   1200
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "FORFAC07.frx":3F18
      Top             =   1200
   End
   Begin VB.Label Label45 
      Height          =   225
      Left            =   0
      TabIndex        =   140
      Top             =   0
      Visible         =   0   'False
      Width           =   585
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
      TabIndex        =   43
      Top             =   0
      Visible         =   0   'False
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
      TabIndex        =   45
      Top             =   500
      Visible         =   0   'False
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
      TabIndex        =   44
      Top             =   0
      Visible         =   0   'False
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
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Menu Archivos 
      Caption         =   "Archivos"
      Begin VB.Menu nva_fac 
         Caption         =   "Nueva Factura"
      End
      Begin VB.Menu anul_comp 
         Caption         =   "Anulación de Comprobantes"
      End
      Begin VB.Menu imp_cre 
         Caption         =   "Imputación de Créditos"
      End
      Begin VB.Menu ajuste 
         Caption         =   "Ajustes de Saldos Acreedores"
      End
      Begin VB.Menu linea1 
         Caption         =   "-"
      End
      Begin VB.Menu maes_pro 
         Caption         =   "Maestro de Proveedores"
      End
      Begin VB.Menu linea2 
         Caption         =   "-"
      End
      Begin VB.Menu print 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu exit 
         Caption         =   "Terminar"
      End
   End
   Begin VB.Menu edicion 
      Caption         =   "Edición"
      Begin VB.Menu anu_comp1 
         Caption         =   "Anulación de Proveedores"
      End
      Begin VB.Menu ajuste1 
         Caption         =   "Ajustes de Saldos de Acreedores"
      End
   End
   Begin VB.Menu cons_list 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu cta_cte 
         Caption         =   "Cuenta Corriente"
      End
      Begin VB.Menu notas_pro 
         Caption         =   "Notas de Proveedores"
      End
   End
   Begin VB.Menu setup 
      Caption         =   "Configuración"
      Begin VB.Menu tasa_iva 
         Caption         =   "Tasas IVA y Percepciones "
      End
      Begin VB.Menu ingbru 
         Caption         =   "Ingresos Brutos"
      End
      Begin VB.Menu cuencon 
         Caption         =   "Cuentas Contables"
      End
   End
End
Attribute VB_Name = "FORFAC07"
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
Dim CAMON%, DMONE$(256), TCMONE#(256)
'

Private Sub ajuste_Click()
   Call FACOCOMP07.Command4_Click
End Sub

Private Sub ajuste1_Click()
   Call FACOCOMP07.Command4_Click
End Sub

Private Sub anu_comp1_Click()
   Call FACOCOMP07.Command25_Click
End Sub

Private Sub anul_comp_Click()
   Call FACOCOMP07.Command25_Click
End Sub

'
Private Sub BOTFEVEN_Click(Index As Integer)
   Call SELECHO("SELFECHA")
   VDEVU$ = VALACT1$("SELFECHA")
   If VDEVU$ <> "" Then
     FEVEN(Index).TEXT = VDEVU$
     FEVEN(Index).Refresh
   End If
End Sub

Private Sub BOTNEXT_Click()
    '
    COTIMONEM# = TCMONE#(MONEMI%)
    NiveCla% = 0
    If Picture4.Visible = True Then NiveCla% = XVALO(Mid$(Label40, 2))
    If CONTROL$("", "VALICAI") = "SI" Then
        If VALIDACAI = 1 Then
          Call GRAFACOM
          MODIFIX% = 0
        End If
     Else
        Call GRAFACOM
        MODIFIX% = 0
    End If
    COTIMONEM# = 0
    Screen.MousePointer = 1
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
    NCXI% = XVALO(TRIM$(VALACT1$("ACREDOR")))
    If NCXI% > 0 Then
        NC% = NCXI%
        Text1.TEXT = TRIM$(Str$(NC%))
        Text2.TEXT = rasocli$((-1) * NC%)
        Text3.TEXT = DOMICLI$((-1) * NC%)
        Text4.TEXT = CPOSCLI$((-1) * NC%)
        Text5.TEXT = LOCACLI$((-1) * NC%)
        Text13.TEXT = CUITCLI$((-1) * NC%)
        Label26 = CODICLI((-1) * NC%)
          If LETRAFA$ <> "M" Then LETRAFA$ = "A"
          PVCL% = PIVACLI%(-1 * NC%)
          If PVCL% = 0 Or PVCL% = 2 Or PVCL% = 6 Or PVCL% = 4 Then LETRAFA$ = "C"
          If PVCL% = 5 Then LETRAFA$ = "E"
        Label20 = LETRAFA$
        If Left$(Text2.TEXT, 6) = "VARIOS" Then
            Text13.SetFocus
          Else
            Call BUSCACAI(NC%)
            If PVCL% = 0 Or PVCL% = 5 Or PVCL% = 6 Then
                Text6.SetFocus     ' FECHA DE EMISION
              Else
                Text16.SetFocus    ' NUMERO DE CAI
            End If
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
        TEXT10.TEXT = TRIM$(DP1$)
        Label18.Caption = ECOARCH$("PROVINC", DP1$)
    End If
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   Call SELECHO("MEPAGO")
   If VALACT1$("MEPAGO") <> "" Then
     TPV% = XVALO(VALACT1$("MEPAGO"))
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

Private Sub Command20_Click()
   Unload Me
End Sub

Private Sub ctacte_imp_Click()
   Call RECUEP07.Command3_Click
End Sub

Private Sub ctacte_pan_Click()
   Call Command4_Click
End Sub

Private Sub Command21_Click()
   Label25 = ""
   Text14 = ""
   Text15 = ""
End Sub

Private Sub Command29_Click()
   Call HELPONLINE("ACRE_GES")
End Sub

Private Sub cta_cte_Click()
   Call Command4_Click
End Sub

Private Sub cuencon_Click()
   Call OPFACOM07.Command4_Click
End Sub

Private Sub exit_Click()
   Call Command20_Click
End Sub

Private Sub FEVEN_DblClick(Index As Integer)
   Call BOTFEVEN_Click(Index)
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


Private Sub imp_cre_Click()
   Call FACOCOMP07.Command3_Click
End Sub

Private Sub IMPOGRA_Change(Index As Integer)
   '
   TASIX$ = TRIM$(REPLACAR$(TASIVA(Index).Caption, "%", ""))
   TASI = XVALO(TASIX$)
   IGRA# = XVALO(IMPOGRA(Index).TEXT)
   If Index < 4 Then
     IMPOIVA(Index).TEXT = CSTRING$(MKD$(TASI * IGRA# / 100), 4, 9, 2, 2)
     IMPOIVA(Index).Refresh
   End If
   '
   TOGRAVA# = 0
   For KKI% = 1 To 6
     IGRA# = XVALO(IMPOGRA(KKI%).TEXT)
     TOGRAVA# = TOGRAVA# + IGRA#
   Next KKI%
   '
   ITOGRA.Caption = CSTRING$(MKD$(TOTIMPU# - TOGRAVA#), 4, 11, 2, 2)
   ITOGRA.Refresh
   Call CALTOFAC
   '
End Sub

Private Sub IMPOGRA_DblClick(Index As Integer)
     If XVALO(ITOGRA) > 0 And XVALO(IMPOGRA(Index)) < 0.005 Then
        IMPOGRA(Index) = CSTRING$(MKD$(XVALO(ITOGRA)), 4, 11, 2, 2)
        IMPOGRA(Index).Refresh
     End If
End Sub


Private Sub IMPOGRA_GotFocus(Index As Integer)
    IMPOGRA(Index).SelStart = 0
    IMPOGRA(Index).SelLength = Len(IMPOGRA(Index).TEXT)
End Sub

Private Sub IMPOGRA_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii% = 13 Then
      If Index < 6 Then
         IMPOGRA(Index + 1).SetFocus
        Else
         IMPOIVA(0).SetFocus
      End If
   End If
End Sub

Private Sub IMPOGRA_LostFocus(Index As Integer)
   IGRA# = Abs(CDbl(XVALO(IMPOGRA(Index).TEXT)))
   TASIX$ = TRIM$(TASIVA(Index).Caption)
   If TASIX$ = "" Then IGRA# = 0
   If LETRAFA$ <> "A" And LETRAFA$ <> "M" Then
      If Index <> 4 And Index <> 5 Then
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
      If Index < 6 Then
         NECA% = 1: If Index = 3 Then NECA% = 3
         IMPOIVA(Index + NECA%).SetFocus
        Else
         Text7.SetFocus
      End If
   End If
End Sub

Private Sub IMPOIVA_LostFocus(Index As Integer)
   IGRA# = Abs(CDbl(XVALO(IMPOIVA(Index).TEXT)))
   TASIX$ = TRIM$(TASIVA(Index).Caption)
   If TASIX$ = "" Then IGRA# = 0
   If LETRAFA$ <> "A" And LETRAFA$ <> "M" Then
     IGRA# = 0
   End If
   IMPOIVA(Index).TEXT = CSTRING$(MKD$(IGRA#), 4, 9, 2, 2)
   IMPOIVA(Index).Refresh
End Sub

Private Sub ingbru_Click()
   Call OPFACOM07.Command5_Click
End Sub

Private Sub IVENCI_Change(Index As Integer)
    IMPTOTT = 0
    For i& = 0 To 9
        IMPTOTT = IMPTOTT + XVALO(IVENCI(i&))
    Next i&
    Label43 = TRIM$(FORMATNUM(IMPTOTT, "F12.2"))
End Sub

Private Sub Label17_DblClick()
   '
   If ULTREG("!PROVINB") < 1 Then
      Call COMUNI("Falta Configurar las Percepciones de Ingresos Brutos")
      Exit Sub
   End If
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
        TEXT10 = "*"
        Label18 = "Jurisdicciones Varias"
        Text12 = FORMATNUM$(SUPIBRU#, "F9.2")
        TEXT10.Enabled = False
        Text12.Enabled = False
        Command10.Enabled = False
      Else
        If TEXT10 = "*" Then
          TEXT10 = ""
          Label18 = ""
          Text12 = ""
        End If
        TEXT10.Enabled = True
        Text12.Enabled = True
        Command10.Enabled = True
     End If
   End If
   '
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
   LETRAX$ = Label20
   If LETRAX$ = "A" Then Label20 = "B"
   If LETRAX$ = "B" Then Label20 = "C"
   If LETRAX$ = "C" Then Label20 = "E"
   If LETRAX$ = "E" Then Label20 = "M"
   If LETRAX$ = "M" Then Label20 = "A"
   'ILETRA% = Asc(Label20.Caption + "A") + 1
   'If ILETRA% > 67 Then ILETRA% = 65
   'Label20.Caption = Chr$(ILETRA%)
End Sub

Private Sub Label25_Change()
   Label21 = ECOARCH$("MEPAGO", MKI$(XVALO(Label25)))
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
     TPV% = XVALO(VALACT1$("MEPAGO"))
     If CUEMEPAG%(TPV%) = CUEICARCH% Then
         Call MENSERR(24, "Cartera Valores no Disponible\para Compras de Contado.\   \Debe Cancelar por Orden de Pago.")
         Label25 = ""
       Else
         Label25 = VALACT1$("MEPAGO")
     End If
   End If
End Sub

Private Sub Label29_DblClick()
   '
   If Check2.Value > 0 Then
      Call MENSERR(24, "Carga por Detalle Solo en Moneda Básica.")
      Exit Sub
   End If
   '
   MONEMIA% = MONEMI%
15 MONEMI% = MONEMI% + 1
   If MONEMI% > CAMON% Then MONEMI% = 1
   If DMONE$(MONEMI%) = "" Then GoTo 15
   Label29 = DMONE$(MONEMI%)
   Label29.ForeColor = RGB(0, 0, 255)
   If MONEMI% > 1 Then
      Label29.ForeColor = RGB(255, 0, 0)
   End If
   '
   Call RECALTO(TCMONE#(MONEMI%), TCMONE#(MONEMIA%))
   Text18 = FORMATNUM(TCMONE#(MONEMI%), "F11.4")
'   Label29.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TCMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
'   '
'   COEFCON = TICAMONE(MONEMIA%) / TICAMONE(MONEMI%)
'   TICAMBX# = TICAMONE(MONEMI%)
'   '
'   TOTIMPU# = 0
'   For U& = 1 To List1.ListCount
'     Z$ = List1.List(U& - 1)
'     IPA# = XVALO(Mid$(Z$, 40, 12))
'     PREDOL# = IPA# * COEFCON
'     TOTIMPU# = TOTIMPU# + PREDOL#
'     Call REPLA(Z$, FORMATNUM$(PREDOL#, "F12.2"), 40, 12)
'     '
'     List1.RemoveItem U& - 1
'     List1.AddItem Z$, U& - 1
'   Next U&
'   '
'   For KKI% = 1 To 6
'     IMPOGRA(KKI%) = Space$(11)
'     IMPOIVA(KKI%) = Space$(8)
'   Next KKI%
'   '
'   NC1% = (-1) * XVALO(Text1)
'   PVCL% = PIVACLI%(NC1%)
'   If (LETRAFA$ = "A" Or LETRAFA$ = "M") And PVCL% <> 4 Then  ' CASO DE PROVEEDORES EXENTOS
'        ITOGRA.Caption = CSTRING$(MKD$(TOTIMPU#), 4, 11, 2, 2)
'        ITOGRA.Refresh
'      Else
'        ITOGRA.Caption = Space$(11)
'        IMPOGRA(4).TEXT = CSTRING$(MKD$(TOTIMPU#), 4, 11, 2, 2)
'        IMPOGRA(4).Refresh
'   End If
'   List1.Refresh
'   '
'   IMPOGRA(1).SetFocus
'   If TOTIMPU# < 0.01 Then
'      Text1.SetFocus
'   End If
'   '
End Sub

Sub RECALTO(TCNUE, TCVIE)
   '
   Label29.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TCNUE, "F10.4")) + " - Doble-Click para Cambiar Moneda"
   '
   COEFCON = TCVIE / TCNUE
   TICAMBX# = TCNUE
   '
   TOTIMPU# = 0
   For U& = 1 To List1.ListCount
     Z$ = List1.List(U& - 1)
     IPA# = XVALO(Mid$(Z$, 40, 12))
     PREDOL# = IPA# * COEFCON
     TOTIMPU# = TOTIMPU# + PREDOL#
     Call REPLA(Z$, FORMATNUM$(PREDOL#, "F12.2"), 40, 12)
     '
     List1.RemoveItem U& - 1
     List1.AddItem Z$, U& - 1
   Next U&
   '
   For KKI% = 1 To 6
     IMPOGRA(KKI%) = Space$(11)
     IMPOIVA(KKI%) = Space$(8)
   Next KKI%
   '
   Label36 = FORMATNUM$(TOTIMPU#, "F11.2")
'   NC1% = (-1) * XVALO(Text1)
'   PVCL% = PIVACLI%(NC1%)
'   If (LETRAFA$ = "A" Or LETRAFA$ = "M") And PVCL% <> 4 Then  ' CASO DE PROVEEDORES EXENTOS
'        ITOGRA.Caption = CSTRING$(MKD$(TOTIMPU#), 4, 11, 2, 2)
'        ITOGRA.Refresh
'      Else
'        ITOGRA.Caption = Space$(11)
'        IMPOGRA(4).TEXT = CSTRING$(MKD$(TOTIMPU#), 4, 11, 2, 2)
'        IMPOGRA(4).Refresh
'   End If
   List1.Refresh
   '
   IMPOGRA(1).SetFocus
   If TOTIMPU# < 0.01 Then
      Text1.SetFocus
   End If
   '
End Sub

Private Sub Label31_Change()
  '
  For KKII% = 1 To 6
     IMPOGRA(KKII%) = ""
  Next KKII%
  '
  IMPONET1# = XVALO(Label31)
  LETRAFA$ = Label20
  If LETRAFA$ = "A" Then
     ITOGRA = FORMATNUM(IMPONET1#, "F10.2")
    Else
     IMPOGRA(4) = FORMATNUM(IMPONET1#, "F10.2")
  End If
  TOTIMPU# = IMPONET1#
  '
End Sub

Private Sub Label36_Change()
  '
  Call Text19_Change
  '
End Sub

Private Sub Label40_DblClick()
   If DERACCE%("BAJNICLA/NOMESS", "Reducir Nivel de Clasificación") >= 2 Then
      NICLAC% = XVALO(Mid$(Label40, 2))
      If NICLAC% > 0 Then
         Label40 = "<" + TRIM$(Str$(NICLAC% - 1))
         Label40.Refresh
      End If
   End If
End Sub

Private Sub Label6_DblClick()
   If List1.ListCount > 0 Then
      NICON% = NIVCONFI%
      If NICON% > 0 Then
         Label40 = "<" + TRIM$(Str$(NICON%))
         Picture4.Visible = True
      End If
   End If
End Sub

'
Private Sub maes_pro_Click()
   Call Command11_Click
End Sub

Private Sub notas_pro_Click()
   Call Command3_Click
End Sub

Private Sub nva_fac_Click()
   Text1.SetFocus
End Sub

Private Sub print_Click()
   Call Command5_Click
End Sub

Private Sub tasa_iva_Click()
   Call Command8_Click
End Sub

Private Sub Text1_LostFocus()
   CLITE$ = TRIM$(Text1.TEXT)
   If CLITE$ <> "" Then
     If XVALO(CLITE$) < 1 Or XVALO(CLITE$) > 32767 Then
       Text1.TEXT = Space$(5)
       Exit Sub
     End If
     If ECOARCH$("ACREDOR", MKI$(XVALO(CLITE$))) = "" Then
       Text1.TEXT = Space$(5)
       Exit Sub
     End If
   End If
End Sub

Private Sub Text10_LostFocus()
   If TRIM$(TEXT10.TEXT) <> "" Then
     If TRIM$(TEXT10.TEXT) <> "*" Then
       If ECOARCH$("PROVINC", TRIM$(TEXT10.TEXT)) = "" Then
         TEXT10.TEXT = " "
         On Error Resume Next
         TEXT10.SetFocus
         On Error GoTo 0
       End If
     End If
   End If
End Sub
'
Private Sub Text10_GotFocus()
    Text12.SelStart = 0
    Text12.SelLength = Len(TEXT10.TEXT)
End Sub

Private Sub Text12_Change()
   Call CALTOFAC
End Sub

'
Private Sub Text12_GotFocus()
    Text12.SelStart = 0
    Text12.SelLength = Len(Text12.TEXT)
End Sub

Private Sub TEXT12_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
      Text20.SetFocus
   End If
End Sub

Private Sub TEXT12_LostFocus()
   IGRA# = Abs(CDbl(XVALO(Text12.TEXT)))
   If LETRAFA$ <> "A" And LETRAFA$ <> "M" Then
     IGRA# = 0
   End If
   DPX1$ = TRIM$(TEXT10.TEXT)
   If IGRA# >= 0.005 Then
     If DPX1$ = "" Or ECOARCH$("PROVINC", DPX1$) = "" Then
       Call MENSERR(24, "Falta Jurisdicción Percepción Ingresos Brutos")
       IGRA# = 0
     End If
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
       If XVALO(Text1.TEXT) > 0 Then
          If XVALO(Text1.TEXT) <= 32767 Then
             NC% = XVALO(Text1.TEXT)
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

Private Sub Text15_DblClick()
    '\\\OT-06-0570-RG-20/12/06///
    Call Command13_Click
    '\\\OT-06-0570-RG-FIN///
End Sub

Private Sub Text16_GotFocus()
    Text16.SelStart = 0
    Text16.SelLength = Len(Text16)
End Sub

Private Sub Text17_GotFocus()
    Text17.SelStart = 0
    Text17.SelLength = Len(Text17)
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      Text17.SetFocus
    End If
End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      Text6.SetFocus
    End If
End Sub

Private Sub Text17_LostFocus()
    If TRIM$(Text17) <> "" Then
      Text17 = FECHATEX$(FECHANUM(Text17))
      Text17.Refresh
    End If
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Call Text18_lostfocus
   End If
End Sub

Private Sub Text18_lostfocus()
  '
  TC1# = XVALO(Text18)
  If MONEMI% <= 1 Then TC1# = 1
  If TC1# < 0.1 Then
    TC1# = TICAMONE(MONEMI%)
  End If
  Text18.TEXT = FORMATNUM$(TC1#, "F11.4")
  '
  TCMONE#(MONEMI%) = XVALO(Text18)
  Label29.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TCMONE#(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
  '
End Sub

Private Sub Text19_Change()
  '
  TN1# = Val(Label36)
  BN1 = COEFDESCU(TRIM$(Text19))
  IMTO1# = (Int(100 * TN1# * (1 - BN1 / 100) + 0.5)) / 100
  BN2# = TN1# - IMTO1#
  '
  Label35 = FORMATNUM(BN2#, "F11.2")
  Label31 = FORMATNUM(IMTO1#, "F11.2")
  '
End Sub

Private Sub Text19_GotFocus()
   If Check2.Value > 0 Then
      Text20.SetFocus
   End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If XVALO(Text1.TEXT) > 0 Then
          If XVALO(Text1.TEXT) <= 32767 Then
             NC% = XVALO(Text1.TEXT)
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

Private Sub Text20_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
      FEVEN(0).SetFocus
   End If
End Sub

Private Sub Text22_Change()
   Call CALTOFAC
End Sub

Private Sub Text22_GotFocus()
    Text22.SelStart = 0
    Text22.SelLength = Len(Text22.TEXT)

End Sub

Private Sub Text22_LostFocus()
   REDON# = CDbl(XVALO(Text22.TEXT))
   Text22.TEXT = CSTRING$(MKD$(REDON#), 4, 7, 2, 2)
   Text22.Refresh

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
   IGRA# = Abs(CDbl(XVALO(Text7.TEXT)))
   If LETRAFA$ <> "A" And LETRAFA$ <> "M" Then
     IGRA# = 0
   End If
   Text7.TEXT = CSTRING$(MKD$(IGRA#), 4, 9, 2, 2)
   Text7.Refresh
End Sub

Private Sub Text8_Change()
   Call CALTOFAC
End Sub

'
Private Sub Text8_GotFocus()
    Text8.SelStart = 0
    Text8.SelLength = Len(Text8.TEXT)
End Sub

Private Sub TEXT8_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
      TEXT10.SetFocus
   End If
End Sub

Private Sub TEXT8_LostFocus()
   IGRA# = Abs(CDbl(XVALO(Text8.TEXT)))
   If LETRAFA$ <> "A" And LETRAFA$ <> "M" Then
     IGRA# = 0
   End If
   Text8.TEXT = CSTRING$(MKD$(IGRA#), 4, 9, 2, 2)
   Text8.Refresh
End Sub
'
Private Sub ITOGRA_Change()
   '
   TASIX$ = TRIM$(REPLACAR$(TASIVA(0).Caption, "%", ""))
   TASI = XVALO(TASIX$)
   IGRA# = XVALO(ITOGRA.Caption)
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
    TEXT10.TEXT = AJUSTI$(DP1$, 1)
    Label18.Caption = ECOARCH$("PROVINC", TRIM$(DP1$))
End Sub

Private Sub Text10_Change()
    DPX1$ = TRIM$(TEXT10.TEXT)
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
    Call CONECRUN("ARCHIG07/AMAPRO", "Archivos Maestros")
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
    RECUEP07.Text1 = Text1
    PROLLAMA$ = Name
    'Hide
    RECUEP07.Show 1
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    PRINTFORM
    Text1.SetFocus
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    FACOCOMP07.Show 1
    Command6.Enabled = True
End Sub


Private Sub Command7_Click()
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Text6.TEXT = VDEVV$
    End If
End Sub

Sub Command8_Click()
   Command8.Enabled = False
   If DERACCE%("STPFACP", "Configuración de Funcionamiento de Facturas de Proveedores") > 1 Then
     OPFACOM07.Show 1
   End If
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
    VScroll1.Max = (Picture5.Height - VScroll1.Height) / 100
    VScroll1.Value = 0
    '
    For IXX% = 0 To 6
      TXXX$ = TRIM$(Str$(IXX%))
      If IXX% <= 3 Then
        TASA = XVALO(CONTROL$("", "TASIVA-" + TXXX$))
        TASIVA(IXX%).Caption = "0.00%"
        If TASA > 0.005 Then
          TASIVA(IXX%).Caption = TRIM$(CSTRING$(MKS$(TASA), 3, 6, 2, 2) + "%")
        End If
      End If
      '
      CUEM$ = CONTROL$("", "CUEIVA-" + TXXX$)
      If CUEM$ = "" Or CUEINT%(CUEM$) < 1 Then
        Call MENSERR(24, "Falta Definir Cuentas Contables\Correspondientes a IVA Crédito Fiscal.")
        OPFACOM07.Show 1
        Hide
      End If
    Next IXX%
    Call APLICACIONSKINS(Me)
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
    CAMON% = ULTREG&("TAMONED")
    If CAMON% < 1 Then
        CAMON% = 1
        DMONE$(1) = "Pesos"
        TCMONE#(1) = 1
      Else
        For U& = 1 To CAMON%
          XX$ = REGLEIDO$("TAMONED", U&)
          DMX$ = TRIM$(Left$(XX$, 8)) + " "
          PPXX% = InStr(DMX$, " ")
          DMX$ = Left$(DMX$, PPXX% - 1)
          '
          VHX = CVS(Mid$(XX$, 29, 4)): If VHX < 0.005 Then VHX = 1
          DMONE$(U&) = TRIM$(DMX$)
          TCMONE#(U&) = VHX
        Next U&
    End If
    '
    MONEMI% = 1
    Label29 = DMONE$(MONEMI%)
    Text18 = FORMATNUM(TCMONE#(MONEMI%), "F11.4")
    '
    If NIVCONFI% > 0 Then
       Label6.ToolTipText = "Doble Click para Clasificar"
    End If
    '
    Refresh
    On Error Resume Next
    Text1.SetFocus
    On Error GoTo 0
    Screen.MousePointer = 1
    '
    CODCARPE$ = BUSCAVALOR(PARAPLI$, "CODCARPE =") ' codigo de carpeta de import / export
    If CODCARPE$ <> "" Then
        CONDI$ = "CodcarpImp ='" & CODCARPE$ & "' AND STATUS < 9"
        If CantRegistros("CARPEIMP", CONDI$) > 0 Then
'          CCOS$ = VALOBADA("CARPEIMP", "CodCenCos", CONDI$, "NOMESS")
           CCOS$ = COCENCOS$(CODCARPE$)
           Label44 = CODCARPE$
           Label46 = CCOS$
        End If
    End If

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
      If Index < 8 Then
         FEVEN(Index + 1).SetFocus
      End If
   End If
End Sub

Private Sub IVENCI_LOSTFOCUS(Index As Integer)
   SUMAVEN# = XVALO(Label10.Caption)
   For KKI% = 0 To Index - 1
      SUMAVEN# = SUMAVEN# + XVALO(IVENCI(KKI%).TEXT)
   Next KKI%
   IPAVEN# = XVALO(IVENCI(Index).TEXT)
   If SUMAVEN# + IPAVEN# < XVALO(TOBRUFAC#) Then
      If Index < 8 Then
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
    If Check2 Then
        FOFACOBR07.Show 1
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
       SUMACRE# = SUMACRE# + XVALO(Mid$(List2.TEXT, 70, 14))
    Next U&
    List2.TopIndex = TOLISTA&
    List2.ListIndex = ILISTA&
    List2.Visible = True
    VACON% = 0
    '
    If IACRE$ = "" Then
         ACREDIT# = XVALO(IPEN$)
         If ACREDIT# > MAXACRE# - SUMACRE# Then
            ACREDIT# = MAXACRE# - SUMACRE#
            If ACREDIT# < 0 Then ACREDIT# = 0
         End If
         ACRE$ = CSTRING$(MKD$(ACREDIT#), 3, 14, 2, 2)
         Call REPLA(Z$, ACRE$, 70, 14)
         SUMACRE# = SUMACRE# + CDbl(XVALO(ACRE$))
       Else
         Call REPLA(Z$, Space$(14), 70, 14)
         SUMACRE# = SUMACRE# - CDbl(XVALO(IACRE$))
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
    SUMACRE# = XVALO(Label10.Caption)
    SUMACRE# = SUMACRE# - CDbl(XVALO(IACRE$))
    SUMACRE# = SUMACRE# + CDbl(XVALO(Mid$(LIST2ANT$, 70, 14)))
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
   For KKI% = 0 To 7
     Option1(KKI%).FontSize = 8
     Option1(KKI%).FontBold = False
     If Option1(KKI%).Value = True Then
       Option1(KKI%).FontSize = 10
       Option1(KKI%).FontBold = True
     End If
   Next KKI%
   '
   Check2 = 0
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
      ElseIf Option1(6) = True Then
        Label6.Caption = "Crédito Propio:"
        TIDOCOM$ = "CP"
      ElseIf Option1(7) = True Then
        Label6.Caption = "Débito Propio:"
        TIDOCOM$ = "DP"
   End If
   Label6.Refresh
   Call Check2_LOSTFOCUS
   Call SIGUEFACO1
End Sub

Private Sub Text1_Change()
    '
    List1.Clear
    List2.Clear
    Label4.Caption = ""
    Label8.Caption = ""
    Label10.Caption = ""
    Check2.Enabled = True
    If TRIM(Text1) <> "" Then Check2.Enabled = False
    Call BLAFACOM
    '
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Text13 = ""
    Text16 = ""
    Text17 = ""
    '
    If XVALO(Text1.TEXT) > 32767 Then
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
    NCXI% = XVALO(TRIM$(VALACT1$("ACREDOR")))
    If NCXI% > 0 Then
        NC% = NCXI%
        Text1.TEXT = TRIM$(Str$(NC%))
        Text2.TEXT = rasocli$((-1) * NC%)
        Text3.TEXT = DOMICLI$((-1) * NC%)
        Text4.TEXT = CPOSCLI$((-1) * NC%)
        Text5.TEXT = LOCACLI$((-1) * NC%)
        Text13.TEXT = CUITCLI$((-1) * NC%)
          If LETRAFA$ <> "M" Then LETRAFA$ = "A"
          PVCL% = PIVACLI%(-1 * NC%)
          If PVCL% = 0 Or PVCL% = 2 Or PVCL% = 6 Or PVCL% = 4 Then LETRAFA$ = "C"
          If PVCL% = 5 Then LETRAFA$ = "E"
        Label20 = LETRAFA$
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
    Screen.MousePointer = 1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If XVALO(Text1.TEXT) > 0 Then
          If XVALO(Text1.TEXT) <= 32767 Then
             NC% = XVALO(Text1.TEXT)
             If ECOARCH$("ACREDOR", MKI$(NC%)) <> "" Then
                Text2.TEXT = rasocli$((-1) * NC%)
                Text3.TEXT = DOMICLI$((-1) * NC%)
                Text4.TEXT = CPOSCLI$((-1) * NC%)
                Text5.TEXT = LOCACLI$((-1) * NC%)
                Text13.TEXT = CUITCLI$((-1) * NC%)
                Label26 = CODICLI((-1) * NC%)
                If LETRAFA$ <> "M" Then LETRAFA$ = "A"
                PVCL% = PIVACLI%(-1 * NC%)
                If PVCL% = 0 Or PVCL% = 2 Or PVCL% = 6 Or PVCL% = 4 Then LETRAFA$ = "C"
                If PVCL% = 5 Then LETRAFA$ = "E"
                Label20 = LETRAFA$
                If Left$(Text2.TEXT, 6) = "VARIOS" Then
                    Text13.SetFocus
                  Else
                    Call BUSCACAI(NC%)
                    If PVCL% = 0 Or PVCL% = 5 Or PVCL% = 6 Then
                        Text6.SetFocus     ' FECHA DE EMISION
                      Else
                        Text16.SetFocus    ' NUMERO DE CAI
                    End If
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
        If XVALO(Text1.TEXT) > 0 Then
          If XVALO(Text1.TEXT) <= 32767 Then
            NC% = XVALO(Text1.TEXT)
            If REGICLI&((-1) * NC%) > 0 Then
              If TICUEPRO%((-1) * NC%) = 0 Then
                List1.SetFocus
                Exit Sub
              End If
            End If
          End If
        End If
      End If
      Text1.SetFocus
    End If
End Sub

Private Sub Text11_LostFocus()
    '
    If TRIM$(Text11.TEXT) <> "" Then
       Call SIGUEFACO1
    End If
    '
End Sub

Private Sub Text2_gotfocus()
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
    NC% = XVALO(Text1.TEXT)
    If NC% < 1 Or ECOARCH$("ACREDOR", MKI$(NC%)) = "" Then
       Call MENSERR(24, "Número de Cuenta no Válido")
       Exit Sub
    End If
    '
    If CONTROL$("", "VALICAI") = "SI" Then
      If VALIDACAI = 0 Then Exit Sub
    End If
    '
    If Left$(rasocli$((-1) * NC%), 6) <> "VARIOS" Then
       Text2.TEXT = rasocli$((-1) * NC%)
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
   Dim FVI%(10), IM#(10)
   '
   TOBRUFAC# = XVALO(ITOGRA.Caption)
   For KKI% = 0 To 6
     TOBRUFAC# = TOBRUFAC# + (Int(100 * XVALO(IMPOIVA(KKI%).TEXT) + 0.5)) / 100
     If KKI% > 0 Then
        TOBRUFAC# = TOBRUFAC# + (Int(100 * XVALO(IMPOGRA(KKI%).TEXT) + 0.5)) / 100
     End If
   Next KKI%
   TOBRUFAC# = TOBRUFAC# + (Int(100 * XVALO(Text7.TEXT) + 0.5)) / 100
   TOBRUFAC# = TOBRUFAC# + (Int(100 * XVALO(Text8.TEXT) + 0.5)) / 100
   TOBRUFAC# = TOBRUFAC# + (Int(100 * XVALO(Text12.TEXT) + 0.5)) / 100
   TOBRUFAC# = TOBRUFAC# + (Int(100 * XVALO(Text22.TEXT) + 0.5)) / 100
   '
   IVENCI(0).TEXT = CSTRING$(MKD$(TOBRUFAC# - XVALO(Label10)), 4, 11, 2, 2)
   '
   If XVALO(Text1) > 32767 Then
     Call MENSERR(24, "Número de Cliente Inválido ")
     Exit Sub
   End If
   NCI% = XVALO(Text1)
   If NCI% < 1 Then Exit Sub
   '
   COPAGI% = CPAGCLI%((-1) * NCI%)
   Call BUCONPA(LU$, COPAGI%, CoPag$, FREF%, PLA$)
   
   IORIG# = TOBRUFAC#
   FF = FECHANUM(Text6)
   FFS$ = FECHATEX$(FF)
   PLA$ = Left$(PLA$ + String$(6, 0), 6)

   If FREF% = 2 Then FREF% = 1   ' AUN NO FUNCIONA SEMANA DE FACTURA
   If FREF% = 3 Then
     If XVALO(Left$(FFS$, 2)) <= 15 Then FEBAS$ = "15" + Mid$(FFS$, 3)
     If XVALO(Left$(FFS$, 2)) > 15 Then FEBAS$ = "31" + Mid$(FFS$, 3)
   End If
   If FREF% = 1 Then FEBAS$ = FFS$
   If FREF% = 4 Then FEBAS$ = "31" + Mid$(FFS$, 3)
   '
   For i% = 1 To 6: FVI%(i%) = 0: Next i%
   A11 = FECHANUM(FEBAS$): B22 = A11
   NUPAG% = 1
   For i% = 1 To 6
     PLAI% = Asc(Mid$(PLA$, i%, 1))
     If PLAI% > 0 Or i% = 1 Then
       While DIENTRE%(A11, B22) < PLAI%
         B22 = DIPOST(B22)
       Wend
       FVI%(i%) = B22
       NUPAG% = i%
     End If
   Next i%
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
   For H% = 1 To 9
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
   List1.Clear
   List2.Clear
   Label4.Caption = ""
   Label8.Caption = ""
   Label10.Caption = ""
   Text11.TEXT = ""
   Check1.Value = 1
   '
   FF% = HOY(HO$)
   Text9.TEXT = HO$
   Text6.TEXT = HO$
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
   Label25 = ""
   Label21 = ""
   Text14 = ""
   Text15 = ""
   Text16 = ""
   Text17 = ""
   '
   ITOGRA.Caption = Space$(11): ITOGRA.Refresh
   Text7.TEXT = Space$(7)
   Text8.TEXT = Space$(7)
   Text12.TEXT = Space$(7)
   TEXT10.TEXT = " "
   Text20 = ""
   Text22 = "" ' TEXT REDONDEO
   '
   Label18.Caption = ""
   For KKI% = 1 To 6: IMPOGRA(KKI%).TEXT = Space$(11): Next KKI%
   For KKI% = 0 To 6: IMPOIVA(KKI%).TEXT = Space$(9): Next KKI%
   For KKI% = 0 To 8: FEVEN(KKI%).TEXT = Space$(8): Next KKI%
   For KKI% = 0 To 8: IVENCI(KKI%).TEXT = Space$(11): Next KKI%
   VScroll1.Value = 0
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
   Label36 = ""
   Text19 = ""
   Text19.BackColor = vbWhite
   Label35 = ""
   Label31 = ""
   '
   Label40 = "0"
   Picture4.Visible = False
   '
End Sub
'
Sub GRAFACOM()
   '
   Dim VENCIPESOS#(16)
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
      NROX$ = "-" + NROX$
      While Len(NROX$) < 13: NROX$ = "0" + NROX$: Wend
      NROX$ = "CO-" + TRIM$(Label20.Caption) + "-" + AJUSTI$(NROX$, 13)
   End If
   '
   SIG% = 1: Tipo% = 2
   If Left$(TRIM$(NROX$), 2) = "TF" Then Tipo% = 22
   If Left$(TRIM$(NROX$), 2) = "RH" Then Tipo% = 32
   If Left$(TRIM$(NROX$), 2) = "CO" Then Tipo% = 42
   If Left$(TRIM$(NROX$), 2) = "ND" Then Tipo% = 3
   If Left$(TRIM$(NROX$), 2) = "CP" Then Tipo% = 3
   If Left$(TRIM$(NROX$), 2) = "NC" Then Tipo% = 6: SIG% = (-1)
   If Left$(TRIM$(NROX$), 2) = "DP" Then Tipo% = 6: SIG% = (-1)
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
      If VA% = 5 Then VA% = 0 ' EXPORTACION
      If LETRAFA$ = "M" Then VA% = 1
      If VA% < 0 Or VA% > 3 Then
          Call MENSERR(24, "Letra de Documento Debe ser 'A', 'C', 'M' o 'E'")
          VACON% = 0
          Exit Sub
      End If
      SUC% = XVALO(Mid$(NROX$, 6, 4))
      NRO& = XVALO(Mid$(NROX$, 11, 8))
      '
      If NRO& < 1 Then
         Call MENSERR(24, "Falta Número de Documento")
      End If
   End If
   '
   NUDOCUCOM$ = NROX$
   '
   NC% = XVALO(Text1.TEXT)
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
   FECHALIM% = XVALO(CONTROL$("", "CIEMECOM"))
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
   If XVALO(ITOGRA.Caption) < (-0.005) Then
      Call MENSERR(24, "Importes no Válidos")
      Call SIGUEFACO3
      Exit Sub
   End If
   '
   If LETRAFA$ <> "A" And LETRAFA$ <> "M" Then
     If Abs(XVALO(ITOGRA.Caption)) >= 0.005 Then
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
      CUE% = XVALO(Left$(Z$, 5))
      IPA# = XVALO(Mid$(Z$, 40, 12))
      '
      If CUE% > 0 Then
         If IPA# >= 0.005 Then
            KKI% = KKI% + 1
            If KKI% > 8 Then GoTo 55
            GRU%(KKI%) = CUE%
            IPAR#(KKI%) = IPA#
            COEBON1 = COEFDESCU(TRIM$(Text19))
            If COEBON1 > 0 Then
               IPA# = ROUND#(IPA# * (1 - COEBON1 / 100))
               IPAR#(KKI%) = IPA#
            End If
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
   If XVALO(Text12.TEXT) >= 0.005 Then
      If TRIM$(Label18.Caption) = "" Then
         Call MENSERR(24, "Imposible Registrar - Falta Jurisdiccion I.Brutos")
         Exit Sub
      End If
   End If
   '
   For KKI% = 0 To 6
     If KKI% = 4 Then
         IMPOX# = XVALO(IMPOIVA(6))     ' iva aduana
       ElseIf KKI% = 5 Then
         IMPOX# = XVALO(Text7)
       ElseIf KKI% = 6 Then
         IMPOX# = XVALO(Text8)
       Else
         IMPOX# = XVALO(IMPOIVA(KKI%))
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
   If TRIM$(TEXT10) <> "" And TRIM$(TEXT10) <> "*" Then
     If XVALO(Text12) >= 0.005 Then
       CUEXTE$ = CUEIBRUPRO$(TEXT10)
       CUEINTE% = CUECOINT%(CUEXTE$)
       If TRIM$(CUEXTE$) = "" Or CUEINTE% < 1 Then
         Call MENSERR(24, "Imposible Registrar.\Falta Cuenta Contable Correspondiente\a Percepc.Ing.Brutos " + TRIM$(Label18) + ".\ \Causas posibles:\- Falta Configurar Cuenta Contable\- Cuenta Contable No Esta Marcada como Gasto")
         Exit Sub
       End If
     End If
   End If
   '
   If Abs(XVALO(Text22)) > 0 Then
     If CUECOINT%(CONTROL$("FACOMPRA", "REDONDEO")) < 1 Then
       Call MENSERR(24, "Falta Cuenta Contable de Redondeo")
       Exit Sub
     End If
   End If
   '
   SUNETO# = XVALO(ITOGRA.Caption)
   SUBRUTO# = SUNETO#
   '
   TOFAC#(1) = SIG% * SUNETO#
   TOFAC#(2) = SIG% * (Abs(XVALO(IMPOGRA(4).TEXT)) + Abs(XVALO(IMPOGRA(5).TEXT)))
   TOFAC#(3) = 0                    ' impuestos intermos
   TOFAC#(4) = 0                    ' IVA TASAS GENERALES
   TOFAC#(5) = 0                    ' iva ADUANA
   TOFAC#(6) = SIG% * Abs(XVALO(Text7.TEXT))
   TOFAC#(7) = SIG% * Abs(XVALO(Text8.TEXT))
   TOFAC#(8) = SIG% * Abs(XVALO(Text12.TEXT))
   '
   REDON# = XVALO(Text22.TEXT)
   TOFAC#(9) = SIG% * Abs(REDON#)
   If REDON# < 0 Then TOFAC#(9) = (-1) * TOFAC#(9) ' EN EL CASO DEL QUE REDONDEO SEA NEGATIVO

   '
   For KKI% = 0 To 6
     SUBRUTO# = SUBRUTO# + (Int(100 * XVALO(IMPOIVA(KKI%).TEXT) + 0.5)) / 100
     If KKI% < 6 Then
          TOFAC#(4) = TOFAC#(4) + SIG% * (Int(100 * XVALO(IMPOIVA(KKI%).TEXT) + 0.5)) / 100
        Else
          TOFAC#(5) = TOFAC#(5) + SIG% * (Int(100 * XVALO(IMPOIVA(KKI%).TEXT) + 0.5)) / 100
     End If
     If KKI% > 0 Then
        SUNETO# = SUNETO# + (Int(100 * XVALO(IMPOGRA(KKI%).TEXT) + 0.5)) / 100
        SUBRUTO# = SUBRUTO# + (Int(100 * XVALO(IMPOGRA(KKI%).TEXT) + 0.5)) / 100
        If KKI% <> 4 And KKI% <> 5 Then
          TOFAC#(1) = TOFAC#(1) + SIG% * (Int(100 * XVALO(IMPOGRA(KKI%).TEXT) + 0.5)) / 100
        End If
     End If
   Next KKI%
   If Abs(SUNETO - SUIMPUTA#) >= 0.005 Then
      Call MENSERR(24, "Importes Incoherentes")
      Exit Sub
   End If
   '
   SUBRUTO# = SUBRUTO# + (Int(100 * XVALO(Text7.TEXT) + 0.5)) / 100
   SUBRUTO# = SUBRUTO# + (Int(100 * XVALO(Text8.TEXT) + 0.5)) / 100
   SUBRUTO# = SUBRUTO# + (Int(100 * XVALO(Text12.TEXT) + 0.5)) / 100
   SUBRUTO# = SUBRUTO# + (Int(100 * XVALO(Text22.TEXT) + 0.5)) / 100 'REDONDEO

   TOFAC#(0) = (-1) * SIG% * SUBRUTO#
   '
   If SUNETO# < 0 Or SUBRUTO# < 0.005 Then
      Call MENSERR(24, "Falta Ingresar Importes")
      VACON% = 0
      Exit Sub
   End If
   '
   SUMAVEN# = 0
   For KKI% = 1 To 9
      IMV1#(KKI%) = 0
      FVII%(KKI%) = FECHANUM(FEVEN(KKI% - 1).TEXT)
      If FVII%(KKI%) > 0 Then
         IMV#(KKI%) = XVALO(IVENCI(KKI% - 1).TEXT)
         If IMV#(KKI%) < 0 Then
            Call MENSERR(24, "Importe a Vencer Negativo")
            VACON% = 0
            Exit Sub
         End If
         IMV1#(KKI%) = IMV#(KKI%) * SIG%
         SUMAVEN# = SUMAVEN# + IMV#(KKI%)
      End If
   Next KKI%
   If Abs(SUMAVEN# + XVALO(Label10) - SUBRUTO#) >= 0.005 Then
      Call MENSERR(24, "Suma de Vencimientos Difiere de Total")
      VACON% = 0
      Exit Sub
   End If
   '
   SUMATOT# = 0
   For KKI% = 0 To 9
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
      If XVALO(Mid$(Z$, 70, 14)) < 0 Then
         Call MENSERR(24, "Valores Aplicados Negativos")
         VACON% = 0
         Exit Sub
      End If
      '
      DOCU$ = Left$(Z$, 20)
      If DOCU$ <> "Anticipo no Aplicado" Then
         CANCE# = CDbl(XVALO(Mid$(Z$, 70, 14)))
         If CANCE# > 0 Then
            NREG& = XVALO(TRIM$(Mid$(Z$, 119, 8)))
            If NREG& > 0 Then
              If NREG& <= ULTREG&("CUECORR") Then
                ORD% = XVALO(Mid$(Z$, 127, 2))
                If ORD% > 0 Then
                  If ORD% <= 6 Then
                    '
                    ss$ = REGLEIDO$("CUECORR", NREG&)
                    If Mid$(ss$, 7, 18) <> Left$(DOCU$, 18) Then ATX% = 1
                    SALDO$ = Mid$(ss$, 139, 60)
                    SALDEU# = CVD(Mid$(SALDO$, 10 * ORD% - 7, 8))
                    If SALDEU# + 0.005 < CANCE# Then ATX% = 1
                    SUMACRE# = SUMACRE# + CDbl(XVALO(Mid$(Z$, 70, 14)))
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
   If TIDOCOM$ = "NC" Or TIDOCOM$ = "DP" Then
     IMV1#(1) = IMV1#(1) + XVALO(Label10)  ' DESCONTAR IMPORTES APLICADOS
   End If
   If XVALO(Label25) > 0 Then
     If TIDOCOM$ <> "CO" Then
       For KKI% = 1 To 6: IMV1#(KKI%) = 0: FVII%(KKI%) = 0: Next KKI%
     End If
   End If
   If TIDOCOM$ = "CO" Then
     DOCUACTIVO$ = NUDOCUCOM$
   End If
'   Call CUECORR(FF%, (-1) * NC%, TIPO%, VA%, SUC%, NRO&, REFEDOC$, MO$, TOFAC#(), FVII%(), IMV1#(), GRU%(), IPAR#(), 1, FFRE%)
'   Call CUECOSS(NUDOCUCOM$)
   '
   If XVALO(Label25) > 0 Then
     If TIDOCOM$ = "CO" Then
       Call COMUNI("Imposible Cancelación Contado.\Genere Orden de Pago C.Cte.")
       NRECIB$ = TRIM$(Text11)
       Exit Sub
     End If
     Call REPLA(NUDOCUCOM$, "FF", 1, 2)
     For KKI% = 0 To 10: TOFAC1#(KKI%) = (-1) * TOFAC#(KKI%): Next KKI%
     For KKI% = 0 To 10: GRU1%(KKI%) = 0: IPAR1#(KKI%) = 0: Next KKI%
'     Call CUECORR(FF%, (-1) * NC%, 14, VA%, SUC%, NRO&, REFEDOC$, MO$, TOFAC1#(), FVII%(), IMV1#(), GRU1%(), IPAR1#(), 1, FFRE%)
'     Call CUECOSS(NUDOCUCOM$)
   End If
   '
   ' grabar factura en CAJABANC
   TOFA0# = 0
   TICA0# = TCMONE#(MONEMI%)
   If Len(TRIM$(Text11)) < 18 Then
        NRECIB$ = TRIM$(Text11)
        While Len(NRECIB$) > 9: NRECIB$ = Mid$(NRECIB$, 2): Wend
        NRECIB$ = Left$(NROX$, 3) + NRECIB$
      Else
        NRECIB$ = Left$(Text11, 5) + Mid$(Text11, 9, 1) + "-" + Mid$(Text11, 14, 5)
   End If
   '
20 NC1% = (-1) * NC%
   Call APERBASDAT("CABAN")
   '
   CoComCor$ = NRECIB$
   CoComLar$ = NROX$
'   CodiEmp% = 0
'   CodiSuc% = SUC%
   NumeClie = 0
   NumeProv% = NC%
   FechEmi = CVDAT(FF%)
   '
   Referen$ = rasocli$(NC1%)
     If Left$(UCase$(rasocli$(NC1%)), 6) = "VARIOS" Then
       Referen$ = Left$(Left$(Text2, 32) + "//" + Left$(Text13, 16), 48)
     End If
   FeContab = CVDAT(FF%)
   FeAcredi = CVDAT(FF%)
   NumeTal = 0
   NumeCom = NRO&
   Observa$ = TRIM$(Text20)
      If Observa$ = "" Then Observa$ = rasocli$(NC1%)
   NumCai$ = TRIM$(Mid$(Text16, 1, 14))
   FechCai = CVDAT(FECHANUM(Text17))
   InclIva$ = "": If Check1.Value = 0 Then InclIva$ = "NO"
   On Error GoTo 0
   '
   ' graba detalle por imputación de compras
   For U& = 1 To List1.ListCount
     Z$ = List1.List(U& - 1)
     CuentaCo% = XVALO(Left$(Z$, 5))
     CentroCos$ = Mid$(Z$, 65, 8)
     ItConcep$ = Mid$(Z$, 75, 8) 'P/ conceptos
       COEBON1 = COEFDESCU(TRIM$(Text19))
     ImpoDeb# = REDONVAL#((1 - COEBON1 / 100) * XVALO(Mid$(Z$, 40, 12)) * TICA0#, 2)
       If TIDOCOM$ = "NC" Or TIDOCOM$ = "DP" Then ImpoDeb# = (-1) * ImpoDeb#
     DEMOVX$ = Mid$(Z$, 8, 30)
     '
     If CuentaCo% > 0 Then
       If Abs(ImpoDeb#) >= 0.005 Then
         TiMovi$ = "COMPRA"
         OpMovi$ = Left$(Text11, 2)
         VaMovi$ = Label20
         Descrip$ = TRIM$(DEMOVX$) + " - " + rasocli$(NC1%)
         TOFA0# = TOFA0# + ImpoDeb#
         Call GRACABAN
       End If
     End If
     CentroCos$ = ""
   Next U&
   ' Anotador de Carpetas de Importacion
   If TRIM$(CODCARPE$) <> "" Then
          Call REPLA(TTXX$, HOS$, 1, 8)
          Call REPLA(TTXX$, USERNAME$, 69, 24)
          NR$ = Left$(CoComCor$, 12)
          Call REPLA(TTXX$, "Ingreso de Factura Nº: " + NR$, 11, 56)
          Call AGREGA_ANOTACARPI(CODCARPE$, TTXX$)
          TTXX$ = Space(76)
          Call REPLA(TTXX$, "Proveedor: " + Left$(rasocli$((-1) * NumeProv%), 64), 11, 56)
          Call AGREGA_ANOTACARPI(CODCARPE$, TTXX$)
          Call AGREGA_ANOTACARPI(CODCARPE$, String$(88, "="))
   End If
   '
   ' grabar importes de IVA
   SUMIVA# = 0
   If XVALO(IMPOGRA(4)) >= 0.005 Then   ' GRABAR IMPORTE EXENTO
     TiMovi$ = "IVA"
     OpMovi$ = "CrFi"
     VaMovi$ = "Exento"
     ImpoDeb# = 0
     BasImpo# = REDONVAL#(XVALO(IMPOGRA(4)) * TICA0#, 2)
     Descrip$ = "IVA B.Imp. Exento " + NRECIB$
     CuentaCo% = CUECOINT%(CONTROL$("", "CUEIVA-0"))
     TOFA0# = TOFA0# + ImpoDeb#
     Call GRACABAN
   End If
   '
   If XVALO(IMPOGRA(5)) >= 0.005 Then   ' GRABAR IMPORTE NO GRAVADO
     TiMovi$ = "IVA"
     OpMovi$ = "CrFi"
     VaMovi$ = "No Grav"
     ImpoDeb# = 0
     BasImpo# = REDONVAL#(XVALO(IMPOGRA(5)) * TICA0#, 2)
     Descrip$ = "IVA B.Imp. No Grav." + NRECIB$
     CuentaCo% = CUECOINT%(CONTROL$("", "CUEIVA-0"))
     TOFA0# = TOFA0# + ImpoDeb#
     Call GRACABAN
   End If
   '
   For KKI% = 0 To 6
     TiMovi$ = "IVA"
     OpMovi$ = "CrFi"
     VaMovi$ = ""
     On Error Resume Next
     VaMovi$ = TASIVA(KKI%)
     On Error GoTo 0
     Descrip$ = "IVA Credito Fiscal " + NRECIB$
     '
     BasImpo# = 0
     If KKI% = 4 Then
         ImpoDeb# = XVALO(IMPOIVA(6))
         BasImpo# = XVALO(IMPOGRA(6))
         VaMovi$ = "Aduana"
       ElseIf KKI% = 5 Then
         ImpoDeb# = XVALO(Text7)
         VaMovi$ = "Per-IVA"
         Descrip$ = "Percepcion I.V.A. " + NRECIB$
       ElseIf KKI% = 6 Then
         ImpoDeb# = XVALO(Text8)
         TiMovi$ = "GAN"
         VaMovi$ = "Per-Gan"
         Descrip$ = "Percepcion Ganancias " + NRECIB$
       Else
         ImpoDeb# = XVALO(IMPOIVA(KKI%))
         SUMIVA# = SUMIVA# + ImpoDeb#
         If KKI% > 0 Then
             BasImpo# = XVALO(IMPOGRA(KKI%))
           Else
             BasImpo# = XVALO(ITOGRA)
         End If
     End If
     '
     If TIDOCOM$ = "NC" Or TIDOCOM$ = "DP" Then ImpoDeb# = (-1) * ImpoDeb#
     If Abs(ImpoDeb#) >= 0.005 Then
       CuentaCo% = CUECOINT%(CONTROL$("", "CUEIVA-" + TRIM$(Str$(KKI%))))
       ImpoDeb# = REDONVAL#(ImpoDeb# * TICA0#, 2)
       BasImpo# = REDONVAL#(BasImpo# * TICA0#, 2)
       TOFA0# = TOFA0# + ImpoDeb#
       Call GRACABAN
     End If
     BasImpo# = 0
   Next KKI%
   '
   ' grabar percepcion Ingresos Brutos - CON OPCION MULTIPLES JURISDICCIONES
   If TRIM$(TEXT10) <> "" Then
     If XVALO(Text12) >= 0.005 Then
       If TRIM$(TEXT10) = "*" Then
            For U& = 1 To ULTREG&("!PROVINB")
              XXZ$ = REGLEIDO$("!PROVINB", U&)
              IPERPA# = CVD(Mid$(XXZ$, 57, 8))
              If Abs(IPERPA#) >= 0.005 Then
                CuentaCo% = CVI(Mid$(XXZ$, 27, 2))
                ImpoDeb# = IPERPA#
                  If TIDOCOM$ = "NC" Or TIDOCOM$ = "DP" Then ImpoDeb# = (-1) * ImpoDeb#
                TiMovi$ = "IBRU"
                OpMovi = "IB-" + Left$(XXZ$, 1)
                VaMovi = "PERC"
                Descrip$ = "Percep.I.Brutos " + NRECIB$
                ImpoDeb# = REDONVAL#(ImpoDeb# * TICA0#, 2)
                TOFA0# = TOFA0# + ImpoDeb#
                Call GRACABAN
              End If
            Next U&
          Else
            CuentaCo% = CUECOINT%(CUEIBRUPRO$(TEXT10))
            ImpoDeb# = XVALO(Text12)
              If TIDOCOM$ = "NC" Or TIDOCOM$ = "DP" Then ImpoDeb# = (-1) * ImpoDeb#
            TiMovi$ = "IBRU"
            OpMovi = "IB-" + TEXT10
            VaMovi = "PERC"
            Descrip$ = "Percep.I.Brutos " + NRECIB$
            ImpoDeb# = REDONVAL#(ImpoDeb# * TICA0#, 2)
            TOFA0# = TOFA0# + ImpoDeb#
            Call GRACABAN
       End If
     End If
   End If
   '
   'SE VALIDA QUE TEX22 DE REDONDEO TENGA IMPORTE SI NO SIGUE
   If Abs(XVALO(Text22)) > 0 Then
        CuentaCo% = CUECOINT(CONTROL("FACOMPRA", "REDONDEO"))
        ImpoDeb# = XVALO(Text22)
        If TIDOCOM$ = "NC" Or TIDOCOM$ = "DP" Then ImpoDeb# = (-1) * ImpoDeb#
        TiMovi$ = "REDCOM"
        OpMovi = Left$(Text11, 2)
        VaMovi = Label20
        Descrip$ = "Redondeo " + NRECIB$
        ImpoDeb# = REDONVAL#(ImpoDeb# * TICA0#, 2)
        TOFA0# = TOFA0# + ImpoDeb#
        Call GRACABAN
   End If
   '
   ' graba factura en c.corriente proveedor
   CuentaCo% = CUECOINT%(CUEMADRE$(-1 * NC%))
   'ImpoDeb# = (-1) * Abs(TOFAC#(0))
   ImpoDeb# = (-1) * Abs(TOFA0#)
     If TIDOCOM$ = "NC" Or TIDOCOM$ = "DP" Then ImpoDeb# = (-1) * ImpoDeb#
   TiMovi = "FCOM"
   OpMovi = Left$(Text11, 2)
   VaMovi = Label20
   Descrip$ = CoComLar$ + " - " + rasocli$(NC1%)
   FeAcredi = CVDAT(FFRE%)
   Call GRACABAN
   FeAcredi = CVDAT(FF%)
   '
   ' grabar pago contado
   If TRIM$(Label25) <> "" Then
     If TIDOCOM$ = "CO" Then
       GoTo 80
     End If
     '
     CuentaCo% = CUECOINT%(CUEMADRE$(-1 * NC%))
     TiMovi = "OPAG"          ' ORDEN DE PAGO
     OpMovi = "PC"            ' pago contado
     VaMovi = "CTDO"          ' CONTADO
     Descrip = TRIM$(DEMOVX$) + " - " + rasocli$(NC1%)
     'ImpoDeb# = Abs(TOFAC#(0))
     ImpoDeb# = Abs(TOFA0#)
       If TIDOCOM$ = "NC" Or TIDOCOM$ = "DP" Then ImpoDeb# = (-1) * ImpoDeb#
     Call GRACABAN
     '
     FeContab = CVDAT(FF%)
     FeAcredi = CVDAT(FECHANUM(Text15))
       If FeAcredi < Now Then FeAcredi = Now
     CuentaCo% = CUEMEPAG%(XVALO(Label25))
     TiMovi = "CHEMI"    ' VALOR EMITIDO / ENTREGADO
     OpMovi = TRIM$(Label25)
     VaMovi = Label20
       BANCO$ = TRIM$(Label21)
       NUCHE$ = TRIM$(Text14)
       FEVECHE$ = Text15
     Descrip$ = AJUSTI$(BANCO$, 18) + AJUSTI$(NUCHE$, 14) + FECHATEX$(FEX)
     'ImpoDeb# = (-1) * Abs(TOFAC#(0))
     ImpoDeb# = (-1) * Abs(TOFA0#)
       If TIDOCOM$ = "NC" Or TIDOCOM$ = "DP" Then ImpoDeb# = (-1) * ImpoDeb#
     Observa$ = TRIM$(BANCO$) + " " + TRIM$(NUCHE$) + " " + TRIM$(FEVECHE$)
     Call GRACABAN
     '
   End If
   '
   ' Grabar Vencimientos de Facturas ******************************************
   SIGNO% = 1
   If TIDOCOM$ = "NC" Or TIDOCOM$ = "DP" Then SIGNO% = (-1)
   If XVALO(Label25) < 1 Then ' no hay cancelacion de contado
     'IBruTot# = Abs(TOFAC#(0))
     IBruTot# = Abs(TOFA0#)
     IIvaTot# = SUMIVA# * TICA0#
     '
     PRICONVEN% = -1
     SALBRUTO# = Abs(TOFA0#) * SIGNO%
     For KKI% = 0 To 8
       VENCIPESOS#(KKI%) = 0
       If FEVEN(KKI%) <> "" Then
         If XVALO(IVENCI(KKI%)) >= 0.005 Then
           If PRICONVEN% < 0 Then PRICONVEN% = KKI%
           VENCIPESOS#(KKI%) = REDONVAL#(XVALO(IVENCI(KKI%)) * SIGNO% * TICA0#, 2)
           SALBRUTO# = SALBRUTO# - VENCIPESOS#(KKI%)
         End If
       End If
     Next KKI%
     If PRICONVEN% >= 0 Then
       VENCIPESOS#(PRICONVEN%) = VENCIPESOS#(PRICONVEN%) + SALBRUTO#
     End If
     '
     For KKI% = 0 To 8
       If FEVEN(KKI%) <> "" Then
         If Abs(VENCIPESOS#(KKI%)) >= 0.005 Then
           NuorVen% = KKI% + 1
           FeVenc = CVDAT(FECHANUM(FEVEN(KKI%)))
           ImVenc# = VENCIPESOS#(KKI%)
           Call GRACREPEN
         End If
       End If
     Next KKI%
     '
   End If
   '
   ' **************************************************************************
   
80 ' *********************************** aqui blanquear facturas canceladas
   '
   FechEmi = CVDAT(FF%)
   RefApli$ = CoComLar$
   NumeProv% = XVALO(Text1)
   '
   XX% = 0
   For U& = 1 To List2.ListCount
     List2.ListIndex = U& - 1
     Z$ = List2.List(U& - 1)
     DOCU$ = Left$(Z$, 20)
     CANCE# = CDbl(XVALO(Mid$(Z$, 70, 14)))
     If CANCE# > 0 Then
       DOCUCA$ = TRIM$(Left$(Z$, 18))
       ICAN# = CANCE#
       '
       DocOrCor$ = ""
       TotOrig# = 0
       IvaOrig# = 0
       '
       Call APERBASDAT("CABAN")
       SQLX$ = "SELECT * FROM SACREPEN"
       SQLX$ = SQLX$ + " WHERE NuProve = " & NumeProv% & " "
       SQLX$ = SQLX$ + " AND CODICOM_LAR = '" & DOCUCA$ & "' "
       SQLX$ = SQLX$ + " AND SALDACRE > 0 "
       Dim SACCCREPEN As ADODB.Recordset
       Set SACCCREPENP = New ADODB.Recordset
25     On Error Resume Next
       SACCCREPENP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
          On Error GoTo 0
          Call CapturaErrorOpen
          GoTo 25
       End If
       On Error GoTo 0

       With SACCCREPENP
         '.FindFirst "NuProve = " & NumeProv% & " AND CODICOM_LAR = '" & DOCUCA$ & "' AND SALDACRE > 0"
         'If .NoMatch = False Then
         On Error Resume Next
         If Not (.EOF And .BOF) Then
           On Error GoTo 0
           Do While Not (.EOF)
            PARCAN# = ICAN#: If !SALDAcre < ICAN# Then PARCAN# = !SALDAcre
            '.Edit
            !SALDAcre = !SALDAcre - PARCAN#
            .Update
            '
            DocOrCor$ = !CodiCom_Cor
            TotOrig# = !IBruComp
            IIVAORIG# = 0
            On Error Resume Next
            IvaOrig# = !IIvaComp
            On Error GoTo 0
            '
            ICAN# = ICAN# - PARCAN#
            'If ICAN# >= 0.005 Then
            ' .FindNext "NuProve = " & NumeProv% & " AND CODICOM_LAR = '" & DOCUCA$ & "' AND SALDACRE > 0"
'              GoTo 850
            If ICAN# < 0.005 Then
              Exit Do
            End If
            .MoveNext
           Loop
         End If
         On Error GoTo 0
       End With
       SACCCREPENP.Close
       Set SACCCREPENP = Nothing
       '
       DocOrLar$ = TRIM$(Left$(Z$, 18))
       ImpApli# = CANCE#
       Call GRAPLIPAG
     End If
   Next U&
   '
   OKX% = 4       ' grabó aplicaciones
   '
   'Grabo Detalle de Facturación en Archivo-Base Central
   If Check2.Value > 0 Then
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
   End If
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
   For KKI% = 0 To 8
      FEVEN(KKI%).TEXT = ""
      IVENCI(KKI%).TEXT = ""
   Next KKI%
   T16$ = Text16: T17$ = Text17
   Call BLAFACOM
   Text16 = T16$: Text17 = T17$  ' PARA QUE MANTENGA EN NUMERO DE CAI
   VACON% = 0
   '
   ' PEDIDO POR TELEBIT QUE VUELVA A MONEDA PESOS - 24/04/06
   MONEMI% = 1
   Label29 = DMONE$(MONEMI%)
   Text18 = FORMATNUM(TCMONE#(MONEMI%), "F11.4")
   Label29.ForeColor = RGB(0, 0, 255)
   '
   Screen.MousePointer = 1
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
   '
End Sub
'
Sub SIGUEFACO1()
    '
    Dim NRO$(10)
    Command12.Enabled = True
    Label25.Enabled = True
    If TIDOCOM$ = "NC" Or TIDOCOM$ = "DP" Then
      Command12.Enabled = False
      Label25.Enabled = False
    End If
    '
    Screen.MousePointer = 11
    '
    NCXI = XVALO(Text1.TEXT)
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
    NC% = XVALO(Text1.TEXT)
    If NC% > 0 Then
      If Left$(rasocli$((-1) * NC%), 6) <> "VARIOS" Then
        Text2.TEXT = rasocli$((-1) * NC%)
        Text3.TEXT = DOMICLI$((-1) * NC%)
        Text4.TEXT = CPOSCLI$((-1) * NC%)
        Text5.TEXT = LOCACLI$((-1) * NC%)
        Refresh
      End If
    End If
    '
    CUITT$ = CUITCLI$((-1) * NC%)
    If VALICUIT%(CUITT$) <> 1 Then
      PVCL% = PIVACLI%((-1) * NC%)
      If PVCL% <> 5 And PVCL% <> 7 Then
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
      C1% = 1: For c% = 1 To 10:  NRO$(c%) = "": Next c%
      '
      For c% = 1 To Len(NROX$)
         C2% = Asc(Mid$(NROX$, c%, 1))
         If C2% = 32 Then GoTo 102
         If (C2% < 48 Or C2% > 57) Then C1% = C1% + 1: GoTo 102
         NRO$(C1%) = NRO$(C1%) + Chr$(C2%)
102   Next c%
      '
      If C1% = 1 Then
         SUC% = 0
         VD# = XVALO(NRO$(1))
         NRO$(2) = NRO$(1)
         NRO$(1) = "0"
      End If
      '
      If XVALO(NRO$(2)) < 1 Or Len(NRO$(1)) > 4 Or Len(NRO$(2)) > 8 Then
          Call MENSERR(24, "Numero de Documento no Permitido")
          Screen.MousePointer = 1
          Text11.SetFocus
          Exit Sub
      End If
      '
      While Len(NRO$(1)) < 4: NRO$(1) = "0" + NRO$(1): Wend
      While Len(NRO$(2)) < 8: NRO$(2) = "0" + NRO$(2): Wend
      '
      If LETRAFA$ <> "M" Then LETRAFA$ = "A"
      PVCL% = PIVACLI%(-1 * NC%)
      If PVCL% = 0 Or PVCL% = 2 Or PVCL% = 6 Or PVCL% = 4 Then LETRAFA$ = "C"
      If PVCL% = 5 Then LETRAFA$ = "E"
      NROX$ = TIDOCOM$ + "-" + LETRAFA$ + "-" + NRO$(1) + "-" + NRO$(2)
      Text11.TEXT = NROX$
      Text11.Refresh
      Label20.Caption = LETRAFA$
      Label20.Refresh
      '
    End If
    '
    ' verifica repeticion del numero ******************************************
    Screen.MousePointer = 11
    NROY1$ = NROX$
    If TIDOCOM$ = "CO" Then           ' evalua numero ingresado
      NROY1$ = "-" + TRIM$(Text11)
      While Len(NROY1$) < 13: NROY1$ = "0" + NROY1$: Wend
      NROY1$ = "CO-" + TRIM$(Label20.Caption) + "-" + AJUSTI$(NROY1$, 13)
    End If
    'Call REPLA(NROY1$, "?", 4, 1)  ' para busqueda de repeticion
    Call REPLA(NROY1$, "%", 4, 1) '(en sql like %) para q tome por iguales facturas q difieren en la letra
    '
    Call APERBASDAT("CABAN")
    'jandy sql
    
    Dim MIRS As ADODB.Recordset
    Set MIRS = New ADODB.Recordset
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE NuProve = " & NC% & " "
    SQLX$ = SQLX$ + "AND TipoMovim = 'FCOM' "
    SQLX$ = SQLX$ + "AND CodiCom_Lar LIKE '" & NROY1$ & "' "
    
    'Set CuecoTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    Set MIRS = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With MIRS
    If Not (.EOF And .BOF) Then
     ' .FindFirst "CodiCom_Lar Like '" & NROY1$ & "' "
     ' If .NoMatch = False Then
         Call MENSERR(24, "Documento ya Registrado")
         Text11 = ""
         Text1.SetFocus
         Exit Sub
      End If
    End With
    MIRS.Close
    Set MIRS = Nothing

95  On Error GoTo 0
    Screen.MousePointer = 1
    ' *************************************************************************
    '
100 List1.Clear: List1.Refresh
    List2.Clear: List2.Refresh
    '
    ' selecciona ultimas imputaciones contables *******************************
    Call APERBASDAT("CABAN")
    Dim CuecoTemp As ADODB.Recordset
    Set CuecoTemp = New ADODB.Recordset
    CuecoTemp.CursorLocation = adUseClient
    
    SQLX$ = "SELECT distinct CUECONTAI  FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE NuProve = " & NC% & " "
    SQLX$ = SQLX$ + "AND TipoMovim = 'COMPRA' "
    'Set CuecoTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
24  On Error Resume Next
    CuecoTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 24
    End If
    On Error GoTo 0
    '
    With CuecoTemp
      'On Error Resume Next
      '.MoveLast
      'If Err < 1 Then
      '  On Error GoTo 0
      If Not (.EOF And .BOF) Then
        .MoveFirst
        Do While Not .EOF
          Z$ = Space$(128)
          Call REPLA(Z$, FORMATNUM$(!CueContaI, "I5"), 1, 5)
          DECUE$ = TRIM$(Left$(ECOARCH$("CUECON", MKI$(!CueContaI)), 30))
          Call REPLA(Z$, DECUE$, 8, 30)
          List1.AddItem Z$
        .MoveNext
        Loop
      End If
      'On Error GoTo 0
    End With
    CuecoTemp.Close
    Set CuecoTemp = Nothing
    
    List1.Refresh
    ' *************************************************************************
    '
    If TIDOCOM$ = "NC" Or TIDOCOM$ = "DP" Then
      Call APERBASDAT("CABAN")
      Dim FacPenTemp As ADODB.Recordset
      Set FacPenTemp = New ADODB.Recordset
      FacPenTemp.CursorLocation = adUseClient
      SQLX$ = "SELECT * FROM SACREPEN "
      SQLX$ = SQLX$ + "WHERE NuProve = " & NC% & " "
      SQLX$ = SQLX$ + "ORDER BY FeVencim ASC;"
      'Set FacPenTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
      '
26    On Error Resume Next
      FacPenTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText

      If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 26
      End If
      On Error GoTo 0
      With FacPenTemp
        'On Error Resume Next
        '.MoveLast
        'If Err < 1 Then
        '  On Error GoTo 0
        If Not (.EOF And .BOF) Then
          .MoveFirst
          Do While Not .EOF
            TIX$ = UCase$(Left$(!CodiCom_Lar, 2))
            X$ = !CodiCom_Lar
            SG% = (-1)
            TIPOS$ = "Factura": If TIX$ = "AS" Then TIPOS$ = "Ajuste de Saldo"
                                If TIX$ = "TF" Then TIPOS$ = "Ticket-Factura"
                                If TIX$ = "RH" Then TIPOS$ = "Rec.Honorarios"
                                If TIX$ = "ND" Then TIPOS$ = "Nota de Debito"
                                If TIX$ = "CP" Then TIPOS$ = "Crédito Propio"
                                If TIX$ = "CO" Then TIPOS$ = "Comprobante"
                                If TIX$ = "OP" Then TIPOS$ = "Orden de Pago": SG% = 1
                                If TIX$ = "FF" Then TIPOS$ = "O/P Fondo Fijo": SG% = 1
                                If TIX$ = "NC" Then TIPOS$ = "Nota de Credito": SG% = 1
                                If TIX$ = "DP" Then TIPOS$ = "Debito Propio": SG% = 1
                                If TIX$ = "CC" Then TIPOS$ = "Comp.Cuenta": SG% = 1
                                '
            If TIX$ = "FC" Or TIX$ = "ND" Or TIX$ = "CP" Or TIX$ = "TF" Or TIX$ = "RH" Or TIX$ = "CO" Or TIX$ = "AS" Or (TIX$ = "OP" And Mid$(X$, 6, 4) = "anul") Then
              SALDO# = !SALDAcre - !SaldDebe
              SG% = s
              If SALDO# >= 0.005 Then
                IMPO# = (-1) * !IBruComp * SG%
                SUMAFAC# = SUMAFAC# + IMPO#
                IMPO# = Abs(IMPO#)
                '
                SUMADEU# = SUMADEU# + SALDO#
                '
                NROX$ = !CodiCom_Lar
                FEMI = CVINT(!FechEmisi)
                FEVE = CVINT(!FEVENCIM)
                '
                IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 13, 2, 2)
                IVATEX$ = CSTRING$(MKD$(IVAO#), 3, 13, 2, 2)
                SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 13, 2, 2)
                '
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
          .MoveNext
          Loop
        End If
        'On Error GoTo 0
      End With
      FacPenTemp.Close
      Set FacPenTemp = Nothing
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
    If Check2 Then
          List1.Clear 'Para que Borre las Imputaciones Contables por Default.
          Call SIGUEBORE(XVALO(Text1)) 'SIGUE CON BR'S SIN FACTUAR
       Else
          Call SIGUEFACO2
    End If
    '
End Sub
'

Sub SIGUEFACO2()
    '
    '!*!=NUEVO PARA GRABACION DE CONCEPTO EN FACT. PROV.
    '
    Call BLANARCH("!CARICOG")
    For U& = 1 To List1.ListCount
        List1.ListIndex = U& - 1
        Z$ = List1.TEXT
        CUE% = XVALO(Left$(Z$, 5))
        '\\\OT-08-989-29/10/08-FL
        'DENOX$ = Mid$(Z$, 8,30)+ CODCUE$(CUE%)
        DENOX$ = Mid$(Z$, 8, 29) + " " + CODCUE$(CUE%)
        '\\\OT-08-989-FIN-FL
        IPA# = XVALO(Mid$(Z$, 40, 12))
        CENCOS$ = TRIM$(Mid$(Z$, 65, 8))
        CONCEP$ = TRIM$(Mid$(Z$, 75, 8)) '!*!
        '
        Y$ = REGBLAN$("!CARICOG")
        Call REPLA(Y$, MKI$(CUE%), 1, 2)
        Call REPLA(Y$, DENOX$, 3, 48)
        Call REPLA(Y$, MKD$(IPA#), 51, 8)
        Call REPLA(Y$, CENCOS$, 59, 8)
        Call REPLA(Y$, CONCEP$, 67, 8) '!*!
        Call REGAPP("!CARICOG", Y$)
    Next U&
    Call CIERRARCH("!CARICOG")
    
    '
200 ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Imputacion Contable"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(250 + Top + Picture1.Top))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture1.Left))
    '
    FOVENTA$ = "CARIFACG"
    If EXISTE%(LUDAT$ + "TABITCO.DAT") > 0 Then
       FOVENTA$ = "CARIFCON"
    End If
    '
    VTB% = VENTABU%(FOVENTA$ + ATRI$) '!*!
    If VTB% >= 0 Then
       RECUEP07.LINDICE.Clear
       List1.Clear
       TOTIMPU# = 0
       For U& = 1 To ULTREG&("!CARICOG")
          Y$ = REGLEIDO$("!CARICOG", U&)
          CUEB$ = Left$(Y$, 2)
          If CVI(CUEB$) > 0 Then
             Z$ = Space$(128)
             Call REPLA(Z$, CSTRING$(CUEB$, 1, 5, 0, 2), 1, 5)
             DECUE$ = Mid$(Y$, 3, 48) ' ECOARCH$("CUECON", CUEB$)
             'DECUE$ = TRIM$(TRIM$(Mid$(DECUE$, 31, 12)) + " " + TRIM$(Left$(DECUE$, 30)))
             Call REPLA(Z$, DECUE$, 8, 32)
             IMBB$ = Mid$(Y$, 51, 8)
             If CVD(IMBB$) < 0 Then
                Call MENSERR(24, "No se Admiten Valores Negativos")
                GoTo 200
             End If
             '
             Call REPLA(Z$, CSTRING$(IMBB$, 3, 13, 2, 2), 38, 13)
             CENCOS$ = Mid$(Y$, 59, 8)
             '\\\OT-06-0244-WAR-28/06/06///
             FCON1% = FECHANUM(Text9.TEXT)
             If TRIM$(CENCOS$) <> "" Then
              If VALIPROY%(CENCOS$, FCON1%) = (-1) Then
                 Call MENSERR(24, "Fecha Contable se Encuentra fuera del Límite\ de Vigencia del Centro de Costos: " + CENCOS$)
                 GoTo 200
              End If
             End If
             ' Validacion de Carpeta de Importacion
             If TRIM$(Label44) <> "" Then
                 CCOS$ = Label46
                 If TRIM$(CENCOS$) <> CCOS$ Then
                   Call REPLA(Y$, CCOS$, 59, 8)
                   Call GRAREG("!CARICOG", Y$, U&)
                   GoTo 200
                 End If
             End If
             '
             '\\\OT-06-0244-WAR-FIN///
             Call REPLA(Z$, CENCOS$, 65, 8)
             '
             CONCEP$ = Mid$(Y$, 67, 8) '!*!
             Call REPLA(Z$, CONCEP$, 75, 8) '!*!
             '
             TOTIMPU# = TOTIMPU# + CVD(IMBB$)
             If CVD(IMBB$) >= 0.005 Then
               RECUEP07.LINDICE.AddItem Z$
             End If
          End If
       Next U&
       '
       For U& = 1 To RECUEP07.LINDICE.ListCount
          RECUEP07.LINDICE.ListIndex = U& - 1
          Z$ = RECUEP07.LINDICE.TEXT
          List1.AddItem Z$
       Next U&
       '
       For KKI% = 1 To 6
          IMPOGRA(KKI%) = Space$(11)
          IMPOIVA(KKI%) = Space$(8)
       Next KKI%
       '
       NC1% = (-1) * XVALO(Text1)
       PVCL% = PIVACLI%(NC1%)
       Label36 = FORMATNUM$(TOTIMPU#, "F11.2") ' TOTIMPU# 27/04/07
'       If (LETRAFA$ = "A" Or LETRAFA$ = "M") And PVCL% <> 4 Then  ' CASO DE PROVEEDORES EXENTOS
'           ITOGRA.Caption = CSTRING$(MKD$(TOTIMPU#), 4, 11, 2, 2)
'           ITOGRA.Refresh
'         Else
'           ITOGRA.Caption = Space$(11)
'           IMPOGRA(4).TEXT = CSTRING$(MKD$(TOTIMPU#), 4, 11, 2, 2)
'           IMPOGRA(4).Refresh
'       End If
    End If
    List1.Refresh
    IMPOGRA(1).SetFocus
    Label36 = FORMATNUM$(TOTIMPU#, "F11.2")
    If TOTIMPU# < 0.01 Then
       Text1.SetFocus
    End If
    '
End Sub

Sub SIGUEFACO3()
   '
   For UI% = 1 To 6
     IMPOGRA(UI%).TEXT = Space$(11)
     IMPOIVA(KKI%) = Space$(8)
   Next UI%
   If LETRAFA$ = "A" Or LETRAFA$ = "M" Then
        ITOGRA.Caption = Space$(11)
        ITOGRA.Caption = CSTRING$(MKD$(TOTIMPU#), 4, 11, 2, 2)
        ITOGRA.Refresh
      Else
        ITOGRA.Caption = Space$(11)
        IMPOGRA(4).TEXT = CSTRING$(MKD$(TOTIMPU#), 4, 11, 2, 2)
        IMPOGRA(4).Refresh
   End If
   TOEXEN# = XVALO(IMPOGRA(4).TEXT) + XVALO(IMPOGRA(5).TEXT)
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
        IPED# = CDbl(XVALO(Mid$(Z$, 32, 13)))
        FEVEI% = FECHANUM(Mid$(Z$, 47, 8))
        SALD# = CDbl(XVALO(Mid$(Z$, 56, 13)))
        ACRE# = CDbl(XVALO(Mid$(Z$, 70, 14)))
        DOCU$ = Mid$(Z$, 1, 18)
        REG2& = XVALO(Mid$(Z$, 119, 8))
        IRE% = XVALO(Mid$(Z$, 127, 2))
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
        SUMACRE# = SUMACRE# + CDbl(XVALO(Mid$(Z$, 70, 14)))
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
Sub SIGUEBORE(NPROX%)
    '
    FORSELBRE.LIMATSEL.Clear
    '
    'NUEVO FILTRO DE FECHA
    Call DESHASFECHA(DES%, HAS%, PERIO$)
    If PERIO$ = "" Then Exit Sub
    '
    DESDEFE = CDbl(CVDAT(DES%))
    HASTAFE = CDbl(CVDAT(HAS%))
    NBRVIE$ = "": NOCOVIE$ = ""
    DESDEF1 = CVDAT(DES%)
    HASTAF1 = CVDAT(HAS%)
    '
    Call APERBASDAT("STOCKS")
    SQLX$ = "SELECT * FROM BOLRECEP "
    SQLX$ = SQLX$ + "WHERE NumProv = " & NPROX% & " "
    SQLX$ = SQLX$ + "AND FechRecep >= '" & DESDEF1 & "' "
    SQLX$ = SQLX$ + "AND FechRecep <= '" & HASTAF1 & "' "
    SQLX$ = SQLX$ + "ORDER BY NumeBoRe ASC "
    'Set BOLDET = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim BOLDET As ADODB.Recordset
    Set BOLDET = New ADODB.Recordset
24  On Error Resume Next
    BOLDET.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 24
    End If
    On Error GoTo 0

    With BOLDET
       On Error Resume Next
       .MoveFirst
       If Err > 1 Then
          On Error GoTo 0
          GoTo 126
       End If
       On Error GoTo 0
       Do While Not .EOF
           NFAC$ = !NumFac_Prov
           CANAP = !CanToApro
           FERECEP = CVINT(!FechRecep)
           '
           If TRIM(NFAC$) = "" Then
              If CANAP > 0.005 Then
                 '
                 NBR$ = TRIM$(Str$(!NUMEBORE))
                   While Len(NBR$) < 6: NBR$ = "0" + NBR$: Wend ' para formatear
                 NRP$ = TRIM$(!NumRem_Prov)
                 NOCOM$ = TRIM$(Str$(!NUMOC))
                 FBORE = FERECEP
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
           '
       .MoveNext
       Loop
    End With
    Set BOLDET = Nothing
    '
  ' hay que prever el caso de que no haya boletas aprobadas sin facturar !!!
126 Screen.MousePointer = 1
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
  '
  If HAYSEL% > 0 Then 'SI HAY SELECCION
25  With FOFACOBR07
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
              SQLX$ = "SELECT * FROM BOLRECEP "
              SQLX$ = SQLX$ + "WHERE NumProv = " & NPROX% & " "
              SQLX$ = SQLX$ + "AND NumeBoRe = " & NBORE& & " "
              SQLX$ = SQLX$ + "AND FechRecep >= '" & DESDEF1 & "' "
              SQLX$ = SQLX$ + "AND FechRecep <= '" & HASTAF1 & "' "
              'Set BOLDET = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
              Dim BOLDET1 As ADODB.Recordset
              Set BOLDET1 = New ADODB.Recordset
26            On Error Resume Next
              BOLDET1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
 
              If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                 On Error GoTo 0
                 Call CapturaErrorOpen
                 GoTo 26
              End If
              On Error GoTo 0
              '
              With BOLDET1
                On Error Resume Next
                .MoveFirst
                If Err Then
                   On Error GoTo 0
                   GoTo 39
                End If
                On Error GoTo 0
                Do While Not .EOF
                   NFAC$ = !NumFac_Prov
                   CANAP1# = !CanToApro
                   FERECEP = CVINT(!FechRecep)
                   '
                   Y$ = REGBLAN$("!DETFACOM")
                   '
                   If TRIM(NFAC$) = "" Then
                      If CANAP1# > 0.005 Then
                         '
                         NRP$ = TRIM$(!NumRem_Prov)
                         FBORE = FERECEP
                         CIXI% = !Cod_Inte
                         CARE# = !CantNom
                         NUOC& = !NUMOC
                         LO1$ = TRIM(!IdenLote)
                         LP1$ = TRIM(!LoteCol_Prov)
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
                             COEFU = CVS(Mid$(REGOC$, 63, 4))
                             MONEX% = Asc(Mid$(REGOC$, 122, 1))
                             If MONEX% < 1 Then MONEX% = 1
                             '
                             PREUN1# = CVD(Mid$(REGOC$, 67, 8)) * TICAMONE(MONEX%) / TICAMONE(MONEMI%)
                             If COEFU > 1 Then PREUN1# = PREUN1# / COEFU
                             'FIN DE TIPO DE MONEDA
                         Else
                             PREUN1# = 0
                         End If
                         'BUSCO SI HAY CANTIDAD YA FACTURADA
                         CANTIYAFAC# = 0
                         For JJX& = ULTREG&("DETFACOM") To 1 Step -1
                           ZZZ$ = REGLEIDO("DETFACOM", JJX&)
                           If CVS(Mid$(ZZZ$, 27, 4)) = NBORE& Then
                             If CVI(Mid$(ZZZ$, 25, 2)) = CIXI% Then
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
35              .MoveNext
                Loop
              End With
              Set BOLDET1 = Nothing

            End If
39        Next J&
          Call SETULTREG("!DETFACOM", J1&)
      End With

      If J1& < 1 Then
        Call COMUNI("No se Encontraron Recepciones\Pendientes de Facturación\para el Set de Recepciones Elegidas.")
      End If
      'AGREGAR AL LIST DE FOFACOBR
      FOFACOBR07.List1.Clear
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
      .Text12 = ""
      .LBLIMPO = FORMATNUM(IMPOTOT#, "F11.2")
      .Show 1
    End With
    Set BOLDET1 = Nothing

  End If
  '
End Sub
'
Function ULTIPRE$(NPROX%, CIXI%, NUOC&)
   '
   ' trae el registro de OCOMDETA correspondiente a la
   ' última compra no anulada del artículo y proveedor
   ' con precio registrado mayor que cero
   '
   Call APERBASDAT("COMPRAS")
   VUELTA% = 0
   '
20 UCO$ = REGBLAN$("OCOMDETA")
   Call REPLA(UCO$, MKI$(NPROX%), 7, 2)
   Call REPLA(UCO$, MKI$(CIXI%), 9, 2)
   '
   SQLX$ = "SELECT * FROM OCOMDETA "
   SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIXI% & " "
   SQLX$ = SQLX$ + "AND Npro_Ocom = " & NPROX% & " "
   SQLX$ = SQLX$ + "AND Nume_Ocom = " & NUOC& & " "
   If VUELTA% = 0 Then SQLX$ = SQLX$ + "AND Pre_Unit >= 0.00005 "
   SQLX$ = SQLX$ + "AND Stat_Ocom < 9 "
   'Set OcomdetaTemp = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim OcomdetaTemp As ADODB.Recordset
   Set OcomdetaTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))

   On Error Resume Next
   With OcomdetaTemp
     .MoveFirst
     
     If Err < 1 Then
         On Error GoTo 0
         Call REPLA(UCO$, MKS$(!Nume_Ocom), 1, 4)
         Call REPLA(UCO$, MKI$(FECHANUM(Format(!Femi_Ocom, "dd/mm/yy"))), 5, 2)
         Call REPLA(UCO$, !Uni_Mas, 11, 4)
         Call REPLA(UCO$, !Uni_Com, 59, 4)
         Call REPLA(UCO$, MKS$(!Coef_Unids), 63, 4)
         Call REPLA(UCO$, Chr$(!Mone_Emis), 122, 1)
         Call REPLA(UCO$, MKD$(!Pre_Unit), 67, 8)
         Call REPLA(UCO$, MKS$(!Desc_Item), 75, 4)
       Else
         On Error GoTo 0
         If VUELTA% < 1 Then
           VUELTA% = VUELTA% + 1
           GoTo 20
         End If
     End If
   End With
   '
   ULTIPRE$ = UCO$
   '
End Function

















' AQUI GRABAR CAJABANC PARA FACTURA
'   CALL Aperbasdat("CABAN")
'   With CajaBanc
'     .AddNew
'     !CodiCom_Cor = NFCX$
'     !CodiCom_Lar = NFCY$
'     !RefeCom = RASOCLI$(NC%)
'     !FechConta = CVDAT(FFI%)   ' fecha de emision
'     !FechAcred = CVDAT(FFI%) ' fecha de acreditacion
'     !CueContaI = CUECOCLI%(NC%)
'     !NuClien = NC%
'     !NuProve = 0
'     !TipoMovim = "FCC"
'     !OpciMovim = OPMOVIX$
'     !VariMovim = Label13
'     !DescriMov = Label7 + " " + Label13.Caption + " " + Label6.Caption + " " + Label12.Caption + "-" + Text11.TEXT
'     !FechEmisi = CVDAT(FFI%)
'     !CodiEmpr = 0
'     !CodiSucu = 0
'     !NumeTalo = XVALO(Label12)
'     !Numecomp = CDbl(Text11)
'     !STATUS = 0
'     !ValAbsComp = TOTAGEN#
'     If TIPO% <> 6 Then
'         !IdebeComp = TOTAGEN#
'         !IHabeComp = 0
'       Else
'         !IdebeComp = 0
'         !IHabeComp = TOTAGEN#
'     End If
'     !ObserGen = RASOCLI$(NC%)
'     !IdenAlta = IDACTU$           ' FECHA - HORA - USUARIO - TERMINAL
'     !IdenBaja = String$(6, 0)
'     '
'     .Update
'     '
'   'graba venta neta
'     TOTANETX# = XVALO(Total(2))
'     If TOTANETX# >= 0.005 Then
'       .AddNew
'       !CodiCom_Cor = NFCX$
'       !CodiCom_Lar = NFCY$
'       !RefeCom = RASOCLI$(NC%)
'       !FechConta = CVDAT(FFI%)   ' fecha de emision
'       !FechAcred = CVDAT(FFI%) ' fecha de acreditacion
'       !CueContaI = QQQ%
'       !NuClien = NC%
'       !NuProve = 0
'       !TipoMovim = "MERC"
'       !OpciMovim = "VEN": If TIPO% = 6 Then !OpciMovim = "DEV"
'       !VariMovim = "COD": If InStr(UCase$(Label6), "LIB") > 0 Then OpciMovim = "LIB"
'       !DescriMov = "Mercaderia Vendida " + NFCX$
'       !FechEmisi = CVDAT(FFI%)
'       !CodiEmpr = 0
'       !CodiSucu = 0
'       !NumeTalo = XVALO(Label12)
'       !Numecomp = CDbl(Text11)
'       !STATUS = 0
'       !ValAbsComp = TOTANETX#
'        If TIPO% <> 6 Then
'            !IdebeComp = 0
'            !IHabeComp = TOTANETX#
'          Else
'            !IdebeComp = TOTANETX#
'            !IHabeComp = 0
'        End If
'        !ObserGen = RASOCLI$(NC%)
'        !IdenAlta = IDACTU$           ' FECHA - HORA - USUARIO - TERMINAL
'        !IdenBaja = String$(6, 0)
'        '
'        .Update
'        '
'     End If
'     '
'   'graba iva inscripto
'     TOTANETX# = XVALO(Total(3)) + XVALO(Total(4))
'     If TOTANETX# >= 0.005 Then
'       .AddNew
'       !CodiCom_Cor = NFCX$
'       !CodiCom_Lar = NFCY$
'       !RefeCom = RASOCLI$(NC%)
'       !FechConta = CVDAT(FFI%)   ' fecha de emision
'       !FechAcred = CVDAT(FFI%) ' fecha de acreditacion
'       '
'       CUEIVA$ = CONTROL$("SVD451S", "CUEIVA")
'       QQX% = CUECOINT%(CUEIVA$)
'       !CueContaI = QQX%
'       '
'       !NuClien = NC%
'       !NuProve = 0
'       !TipoMovim = "IVA"
'       !OpciMovim = "DEFI"
'       !VariMovim = "TBAS"
'       !DescriMov = "Iva Debito Fiscal " + NFCX$
'       !FechEmisi = CVDAT(FFI%)
'       !CodiEmpr = 0
'       !CodiSucu = 0
'       !NumeTalo = XVALO(Label12)
'       !Numecomp = CDbl(Text11)
'       !STATUS = 0
'       !ValAbsComp = TOTANETX#
'        If TIPO% <> 6 Then
'            !IdebeComp = 0
'            !IHabeComp = TOTANETX#
'          Else
'            !IdebeComp = TOTANETX#
'            !IHabeComp = 0
'        End If
'        !ObserGen = RASOCLI$(NC%)
'        !IdenAlta = IDACTU$           ' FECHA - HORA - USUARIO - TERMINAL
'        !IdenBaja = String$(6, 0)
'        '
'        .Update
'        '
'     End If
'     '
'   'graba iva NO inscripto
'     TOTANETX# = XVALO(Total(5))
'     If TOTANETX# >= 0.005 Then
'       .AddNew
'       !CodiCom_Cor = NFCX$
'       !CodiCom_Lar = NFCY$
'       !RefeCom = RASOCLI$(NC%)
'       !FechConta = CVDAT(FFI%)   ' fecha de emision
'       !FechAcred = CVDAT(FFI%) ' fecha de acreditacion
'       '
'       CUEIVA$ = CONTROL$("SVD451S", "CUEADU")
'       QQX% = CUECOINT%(CUEIVA$)
'       !CueContaI = QQX%
'       '
'       !NuClien = NC%
'       !NuProve = 0
'       !TipoMovim = "IVA"
'       !OpciMovim = "DEFI"
'       !VariMovim = "TRNI"
'       !DescriMov = "IVA Resp.No Inscripto " + NFCX$
'       !FechEmisi = CVDAT(FFI%)
'       !CodiEmpr = 0
'       !CodiSucu = 0
'       !NumeTalo = XVALO(Label12)
'       !Numecomp = CDbl(Text11)
'       !STATUS = 0
'       !ValAbsComp = TOTANETX#
'        If TIPO% <> 6 Then
'            !IdebeComp = 0
'            !IHabeComp = TOTANETX#
'          Else
'            !IdebeComp = TOTANETX#
'            !IHabeComp = 0
'        End If
'        !ObserGen = RASOCLI$(NC%)
'        !IdenAlta = IDACTU$           ' FECHA - HORA - USUARIO - TERMINAL
'        !IdenBaja = String$(6, 0)
'        '
'        .Update
'        '
'     End If
'     '
'   End With
'   ' TERMINA DE GRABAR CAJABANC
'   '
'
'
Private Sub Label26_Click()
   Call SELECHO("INDIACRE")
   CPROX$ = TRIM$(VALACT1$("INDIACRE"))
   If CPROX$ <> "" Then
     NCXI% = NUMEPRO%(CPROX$)
     If NCXI% > 0 Then
       If REGICLI&((-1) * NCXI%) > 0 Then
         Label26 = CPROX$
         NC% = NCXI%
         Text1.TEXT = TRIM$(Str$(NC%))
         Text2.TEXT = rasocli$((-1) * NC%)
         Text3.TEXT = DOMICLI$((-1) * NC%)
         Text4.TEXT = CPOSCLI$((-1) * NC%)
         Text5.TEXT = LOCACLI$((-1) * NC%)
         Text13.TEXT = CUITCLI$((-1) * NC%)
           If LETRAFA$ <> "M" Then LETRAFA$ = "A"
           PVCL% = PIVACLI%(-1 * NC%)
           If PVCL% = 0 Or PVCL% = 2 Or PVCL% = 6 Or PVCL% = 4 Then LETRAFA$ = "C"
           If PVCL% = 5 Then LETRAFA$ = "E"
         Label20 = LETRAFA$
         If Left$(Text2.TEXT, 6) = "VARIOS" Then
             Text13.SetFocus
           Else
             Text6.SetFocus
         End If
       End If
     End If
   End If
End Sub

Private Sub Label26_DblClick()
   Call Label26_Click
End Sub

Private Sub BUSCACAI(NC%)
    '
    Call APERBASDAT("CABAN")
    '
    'jandy sql
    ' obtengo el ultimo nro cai (no necesito el movelast)
    'recupero un unico registro aumento la performance
    
    'SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = "SELECT Top 1 * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE NuProve = " & NC% & " "
    SQLX$ = SQLX$ + "AND TipoMovim = '" & "FCOM" & "' "
    SQLX$ = SQLX$ + "AND NroCai <> '' "
'    SQLX$ = SQLX$ + "AND ISNULL(NroCai) = False "
'    SQLX$ = SQLX$ + "AND FeVtoCai >= Now "
    SQLX$ = SQLX$ + "AND NroCai is not null "
    SQLX$ = SQLX$ + "AND FeVtoCai >= '" & FechaCentral & "' "
    SQLX$ = SQLX$ + "ORDER BY FechEmisi desc"
    'SQLX$ = SQLX$ + "ORDER BY FechEmisi ASC"
    'Set FacCompras = CueCorri.OpenRecordset(SQLX$)
    '
    
    Dim FacCompras As ADODB.Recordset
    Set FacCompras = FLEXCONN.Execute(FILTSQL$(SQLX$))
    
    With FacCompras
      'On Error Resume Next
      '.MoveLast
      'If Err < 1 Then
      If Not (.EOF And .BOF) Then
        NUCAI$ = !NROCAI
        FECHACAI = CVINT(!FEVTOCAI)
        If FECHACAI > HOY(HOS$) Then
          Text16 = NUCAI$
          Text17 = FECHATEX$(FECHACAI)
        End If
      End If
      'On Error GoTo 0
    End With
    '
End Sub

Private Function VALIDACAI() As Integer
    '
    VALIDACAI = 1
    NC% = XVALO(Text1)
    PVCL% = PIVACLI%(-1 * NC%)
    If PVCL% = 0 Or PVCL% = 5 Or PVCL% = 6 Then
      Exit Function
    End If
    '
    Text16 = TRIM$(Text16)
    If Text16 = "" Then ' Valida la existencia del CAI
      Call MENSERR(24, "Falta Ingresar Número CAI")
      Text16.SetFocus: GoTo 99
    Else
      If Len(Text16) > 24 Then ' Valida la longitud del CAI
        Call MENSERR(24, "Número de CAI Supera Máximo de 14 Dígitos")
        Text16.SetFocus: GoTo 99
      End If
    End If
    '
    Text17 = TRIM$(Text17)
    If Text17 = "" Then ' Valida la existencia de la fecha del CAI
        Call MENSERR(24, "Falta Fecha de Vencimiento del CAI.")
        Text17.SetFocus: GoTo 99
      Else
        FECHACAI = FECHANUM(Text17)
        If FECHACAI < FECHANUM(Text6) Then
          Call MENSERR(24, "Vencimiento del CAI Anterior\a la Fecha de Emisión de Factura")
          Text17.SetFocus: GoTo 99
        End If
    End If
    '
    Exit Function
    '
99  VALIDACAI = 0
    '
End Function

Private Sub VScroll1_Change()
   Picture5.Top = (-1) * VScroll1.Value * 100
End Sub

Private Sub VScroll1_Scroll()
   Picture5.Top = (-1) * VScroll1.Value * 100
End Sub

Private Sub Check2_LOSTFOCUS()
  For KKI% = 1 To 7
    If Option1(KKI%) Then
      Check2 = 0
    End If
  Next KKI%
End Sub
'
Private Sub Check2_CLICK()
  '
  Call Check2_LOSTFOCUS
  '
  If Check2.Value > 0 Then
     MONEMIA% = MONEMI%
15   MONEMI% = 1
     Label29 = DMONE$(MONEMI%)
     Label29.ForeColor = RGB(0, 0, 255)
     '
     Call RECALTO(TCMONE#(MONEMI%), TCMONE#(MONEMIA%))
     Text18 = FORMATNUM(TCMONE#(MONEMI%), "F11.4")
  End If
  '
End Sub

