VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form TRAZORF14 
   BackColor       =   &H00B0D0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Trazabilidad"
   ClientHeight    =   10185
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10185
   ScaleWidth      =   10605
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "TRAZORF14.frx":0000
      Top             =   120
   End
   Begin VB.CommandButton Command23 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      Picture         =   "TRAZORF14.frx":0234
      Style           =   1  'Graphical
      TabIndex        =   113
      ToolTipText     =   "Recarga Orden de Fabricación"
      Top             =   240
      Width           =   315
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
      Height          =   735
      Left            =   3000
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   106
      Top             =   9360
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00B0D0E0&
      Caption         =   "Trazabilidad vía código de barras"
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
      Height          =   195
      Left            =   240
      TabIndex        =   105
      Top             =   7320
      Width           =   4575
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Limpiar Lista AUTO FIFO"
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
      Left            =   5880
      TabIndex        =   102
      Top             =   9600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command19 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   245
      Left            =   1920
      TabIndex        =   101
      Top             =   8640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Trazar Lotes"
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
      Left            =   8520
      TabIndex        =   100
      Top             =   9600
      Width           =   1095
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Calidad"
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
      Left            =   7320
      TabIndex        =   99
      Top             =   9600
      Width           =   1095
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
      Height          =   375
      Left            =   240
      TabIndex        =   98
      Top             =   9720
      Width           =   2655
   End
   Begin MSFlexGridLib.MSFlexGrid mfgLotes 
      Height          =   1935
      Left            =   240
      TabIndex        =   97
      Top             =   7560
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   3413
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   87
      Top             =   720
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00B0D0E0&
      Caption         =   "CANTIDAD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   6075
      TabIndex        =   77
      Top             =   210
      Width           =   2130
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Left            =   1080
         TabIndex        =   83
         Top             =   1050
         Width           =   915
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Left            =   1080
         TabIndex        =   82
         Top             =   735
         Width           =   915
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Left            =   1080
         TabIndex        =   81
         Top             =   420
         Width           =   915
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo"
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
         Left            =   -360
         TabIndex        =   80
         Top             =   1120
         Width           =   1365
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Proyectado"
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
         Left            =   -360
         TabIndex        =   79
         Top             =   510
         Width           =   1365
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Realizado"
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
         Left            =   -360
         TabIndex        =   78
         Top             =   820
         Width           =   1365
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frame9"
      Height          =   540
      Left            =   2520
      TabIndex        =   71
      Top             =   4560
      Visible         =   0   'False
      Width           =   6630
      Begin VB.Label Label20 
         Height          =   225
         Left            =   4725
         TabIndex        =   76
         Top             =   210
         Width           =   960
      End
      Begin VB.Label Label17 
         Height          =   225
         Left            =   3570
         TabIndex        =   75
         Top             =   210
         Width           =   1065
      End
      Begin VB.Label Label21 
         Height          =   225
         Left            =   2415
         TabIndex        =   74
         Top             =   210
         Width           =   1065
      End
      Begin VB.Label Label33 
         Height          =   225
         Left            =   735
         TabIndex        =   73
         Top             =   210
         Width           =   1590
      End
      Begin VB.Label Label8 
         Height          =   225
         Left            =   105
         TabIndex        =   72
         Top             =   210
         Width           =   540
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00B0D0E0&
      Caption         =   "# "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   8295
      TabIndex        =   68
      Top             =   210
      Width           =   1275
      Begin VB.TextBox Text6 
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
         Height          =   285
         Left            =   315
         TabIndex        =   4
         Top             =   1050
         Width           =   700
      End
      Begin VB.TextBox Text5 
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
         Height          =   285
         Left            =   315
         TabIndex        =   3
         Top             =   525
         Width           =   700
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PLC"
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
         Left            =   -960
         TabIndex        =   70
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PNL"
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
         Left            =   -840
         TabIndex        =   69
         Top             =   315
         Width           =   1335
      End
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
      Left            =   5775
      TabIndex        =   1
      Top             =   315
      Width           =   175
   End
   Begin VB.TextBox Text8 
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
      Left            =   4305
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   735
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
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
      Left            =   3150
      TabIndex        =   0
      Top             =   315
      Width           =   175
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00B0D0E0&
      Caption         =   "DETALLE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   240
      TabIndex        =   37
      Top             =   1750
      Width           =   9360
      Begin VB.Frame Frame5 
         Caption         =   "Frame3"
         Height          =   1760
         Left            =   8900
         TabIndex        =   67
         Top             =   380
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   1770
         Left            =   7570
         TabIndex        =   66
         Top             =   380
         Width           =   15
      End
      Begin VB.Frame Frame6 
         Caption         =   "Frame3"
         Height          =   1770
         Left            =   6250
         TabIndex        =   62
         Top             =   380
         Width           =   15
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame3"
         Height          =   1760
         Left            =   1850
         TabIndex        =   61
         Top             =   380
         Width           =   15
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Roboto Mono Medium"
            Size            =   8.25
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1605
         Left            =   210
         MousePointer    =   1  'Arrow
         TabIndex        =   5
         Top             =   520
         Width           =   9000
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo          Descripción                                   Reserva      Trazado"
         BeginProperty Font 
            Name            =   "Roboto Mono Medium"
            Size            =   8.25
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   210
         TabIndex        =   42
         Top             =   250
         Width           =   9000
      End
   End
   Begin VB.PictureBox Picture14 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   10905
      Left            =   9765
      ScaleHeight     =   10875
      ScaleWidth      =   1350
      TabIndex        =   32
      Top             =   0
      Width           =   1380
      Begin VB.CommandButton AGenRep 
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
         Left            =   105
         Picture         =   "TRAZORF14.frx":037E
         Style           =   1  'Graphical
         TabIndex        =   114
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   5130
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   105
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   103
         Top             =   8000
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   104
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command1 
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
         Picture         =   "TRAZORF14.frx":0688
         Style           =   1  'Graphical
         TabIndex        =   96
         ToolTipText     =   "Salir"
         Top             =   240
         Width           =   650
      End
      Begin VB.CommandButton Command16 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "TRAZORF14.frx":07D2
         Style           =   1  'Graphical
         TabIndex        =   95
         ToolTipText     =   "Impresión de Registro de Trazabilidad"
         Top             =   1890
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command3 
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
         Left            =   105
         Picture         =   "TRAZORF14.frx":0ADC
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   930
         Width           =   650
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Kdx"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   105
         Picture         =   "TRAZORF14.frx":0DE6
         Style           =   1  'Graphical
         TabIndex        =   93
         ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
         Top             =   3180
         Width           =   650
      End
      Begin VB.CommandButton Command13 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   120
         Picture         =   "TRAZORF14.frx":10F0
         Style           =   1  'Graphical
         TabIndex        =   92
         ToolTipText     =   "Guardar Cierre de Orden de Fabricacion"
         Top             =   6200
         Width           =   650
      End
      Begin VB.CommandButton Command14 
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
         Height          =   650
         Left            =   105
         Picture         =   "TRAZORF14.frx":13FA
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   3840
         Width           =   650
      End
      Begin VB.CommandButton Command8 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "TRAZORF14.frx":183C
         Style           =   1  'Graphical
         TabIndex        =   90
         ToolTipText     =   "Realiza Control de Movimientos de Stock Repetidos o Duplicados"
         Top             =   4485
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command7 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "TRAZORF14.frx":1B46
         Style           =   1  'Graphical
         TabIndex        =   89
         ToolTipText     =   "Listado Materiales Reservados por Lote"
         Top             =   2535
         Width           =   650
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00B0D0E0&
      Caption         =   "TRAZABILIDAD DE COMPONENTES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   240
      TabIndex        =   33
      Top             =   4050
      Width           =   9360
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
         Height          =   250
         Index           =   5
         Left            =   3360
         TabIndex        =   112
         Top             =   2700
         Width           =   175
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
         Height          =   250
         Index           =   4
         Left            =   3360
         TabIndex        =   111
         Top             =   2430
         Width           =   175
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
         Height          =   250
         Index           =   3
         Left            =   3360
         TabIndex        =   110
         Top             =   2160
         Width           =   175
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
         Height          =   250
         Index           =   2
         Left            =   3360
         TabIndex        =   109
         Top             =   1890
         Width           =   175
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
         Height          =   250
         Index           =   1
         Left            =   3360
         TabIndex        =   108
         Top             =   1620
         Width           =   175
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
         Height          =   250
         Index           =   0
         Left            =   3360
         TabIndex        =   107
         Top             =   1350
         Width           =   175
      End
      Begin VB.TextBox Text16 
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
         Height          =   250
         Index           =   0
         Left            =   5040
         TabIndex        =   8
         Top             =   1350
         Width           =   1380
      End
      Begin VB.TextBox Text11 
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
         Height          =   285
         Index           =   0
         Left            =   3675
         TabIndex        =   7
         ToolTipText     =   "Doble-Click para Cancelar Reserva Trazada"
         Top             =   1350
         Width           =   1065
      End
      Begin VB.TextBox Text11 
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
         Height          =   250
         Index           =   1
         Left            =   3675
         TabIndex        =   11
         ToolTipText     =   "Doble-Click para Cancelar Reserva Trazada"
         Top             =   1620
         Width           =   1065
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
         Height          =   750
         Left            =   8280
         Picture         =   "TRAZORF14.frx":1E50
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "Graba - Guardar Modificaciones"
         Top             =   2190
         Width           =   855
      End
      Begin VB.CommandButton Command15 
         Caption         =   "AUTO FIFO"
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
         Left            =   210
         TabIndex        =   86
         Top             =   2280
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "FIFO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   85
         Top             =   1950
         Width           =   855
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Ins"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   8280
         TabIndex        =   31
         Top             =   1665
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   8280
         TabIndex        =   30
         Top             =   1320
         Width           =   840
      End
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
         Height          =   285
         Left            =   5460
         TabIndex        =   63
         Top             =   300
         Width           =   960
      End
      Begin VB.CommandButton Command6 
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
         Height          =   250
         Index           =   5
         Left            =   3150
         TabIndex        =   26
         Top             =   2700
         Width           =   175
      End
      Begin VB.CommandButton Command6 
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
         Height          =   250
         Index           =   4
         Left            =   3150
         TabIndex        =   22
         Top             =   2430
         Width           =   175
      End
      Begin VB.CommandButton Command6 
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
         Height          =   250
         Index           =   3
         Left            =   3150
         TabIndex        =   18
         Top             =   2160
         Width           =   175
      End
      Begin VB.CommandButton Command6 
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
         Height          =   250
         Index           =   2
         Left            =   3150
         TabIndex        =   14
         Top             =   1890
         Width           =   175
      End
      Begin VB.CommandButton Command6 
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
         Height          =   250
         Index           =   1
         Left            =   3150
         TabIndex        =   10
         Top             =   1620
         Width           =   175
      End
      Begin VB.CommandButton Command6 
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
         Height          =   250
         Index           =   0
         Left            =   3150
         TabIndex        =   6
         Top             =   1350
         Width           =   175
      End
      Begin VB.TextBox Text11 
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
         Height          =   250
         Index           =   5
         Left            =   3675
         TabIndex        =   27
         Top             =   2700
         Width           =   1065
      End
      Begin VB.TextBox Text16 
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
         Height          =   250
         Index           =   5
         Left            =   5040
         TabIndex        =   28
         Top             =   2700
         Width           =   1380
      End
      Begin VB.TextBox Text17 
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
         Height          =   250
         IMEMode         =   3  'DISABLE
         Index           =   5
         Left            =   6720
         TabIndex        =   29
         Top             =   2700
         Width           =   1380
      End
      Begin VB.TextBox Text11 
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
         Height          =   250
         Index           =   4
         Left            =   3675
         TabIndex        =   23
         Top             =   2430
         Width           =   1065
      End
      Begin VB.TextBox Text16 
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
         Height          =   250
         Index           =   4
         Left            =   5040
         TabIndex        =   24
         Top             =   2430
         Width           =   1380
      End
      Begin VB.TextBox Text17 
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
         Height          =   250
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   6720
         TabIndex        =   25
         Top             =   2430
         Width           =   1380
      End
      Begin VB.TextBox Text11 
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
         Height          =   250
         Index           =   3
         Left            =   3675
         TabIndex        =   19
         Top             =   2160
         Width           =   1065
      End
      Begin VB.TextBox Text16 
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
         Height          =   250
         Index           =   3
         Left            =   5040
         TabIndex        =   20
         Top             =   2160
         Width           =   1380
      End
      Begin VB.TextBox Text17 
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
         Height          =   250
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   6720
         TabIndex        =   21
         Top             =   2160
         Width           =   1380
      End
      Begin VB.TextBox Text11 
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
         Height          =   250
         Index           =   2
         Left            =   3675
         TabIndex        =   15
         Top             =   1890
         Width           =   1065
      End
      Begin VB.TextBox Text16 
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
         Height          =   250
         Index           =   2
         Left            =   5040
         TabIndex        =   16
         Top             =   1890
         Width           =   1380
      End
      Begin VB.TextBox Text17 
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
         Height          =   250
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   6720
         TabIndex        =   17
         Top             =   1890
         Width           =   1380
      End
      Begin VB.TextBox Text16 
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
         Height          =   250
         Index           =   1
         Left            =   5040
         TabIndex        =   12
         Top             =   1620
         Width           =   1380
      End
      Begin VB.TextBox Text17 
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
         Height          =   250
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   6720
         TabIndex        =   13
         Top             =   1620
         Width           =   1380
      End
      Begin VB.TextBox Text17 
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
         Height          =   250
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   6720
         TabIndex        =   9
         Top             =   1350
         Width           =   1380
      End
      Begin VB.Image Image2 
         Height          =   2055
         Left            =   120
         Top             =   1080
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
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
         Left            =   8200
         TabIndex        =   65
         Top             =   300
         Width           =   945
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Trazado"
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
         Left            =   6790
         TabIndex        =   64
         Top             =   375
         Width           =   1335
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
         Height          =   255
         Index           =   5
         Left            =   1470
         TabIndex        =   59
         Top             =   2700
         Width           =   1755
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
         Height          =   255
         Index           =   4
         Left            =   1470
         TabIndex        =   58
         Top             =   2430
         Width           =   1755
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
         Height          =   255
         Index           =   3
         Left            =   1470
         TabIndex        =   57
         Top             =   2160
         Width           =   1755
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
         Height          =   255
         Index           =   2
         Left            =   1470
         TabIndex        =   56
         Top             =   1890
         Width           =   1755
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
         Height          =   255
         Index           =   1
         Left            =   1470
         TabIndex        =   55
         Top             =   1620
         Width           =   1755
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PNL / PLC Fin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   5550
         TabIndex        =   54
         Top             =   1140
         Width           =   1965
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PNL / PLC Com."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   4005
         TabIndex        =   53
         Top             =   1140
         Width           =   1965
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   2100
         TabIndex        =   52
         Top             =   1140
         Width           =   1965
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Requerido"
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
         Left            =   3990
         TabIndex        =   48
         Top             =   375
         Width           =   1395
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
         Height          =   285
         Left            =   1470
         TabIndex        =   45
         ToolTipText     =   "Descripción"
         Top             =   615
         Width           =   7680
      End
      Begin VB.Label Label27 
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
         TabIndex        =   44
         ToolTipText     =   "Código"
         Top             =   300
         Width           =   2175
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
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
         Left            =   60
         TabIndex        =   43
         Top             =   375
         Width           =   1335
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
         Height          =   255
         Index           =   0
         Left            =   1470
         TabIndex        =   38
         Top             =   1350
         Width           =   1755
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "N.C.Recepción"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   0
         TabIndex        =   34
         Top             =   1140
         Width           =   1965
      End
   End
   Begin VB.Label Label7 
      Height          =   225
      Left            =   210
      TabIndex        =   84
      Top             =   1470
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Lote"
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
      Left            =   3600
      TabIndex        =   60
      Top             =   820
      Width           =   615
   End
   Begin VB.Label Label32 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Left            =   1365
      TabIndex        =   51
      ToolTipText     =   "Doble-Click para Blanquear"
      Top             =   315
      Width           =   1800
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción"
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
      Left            =   120
      TabIndex        =   50
      Top             =   1120
      Width           =   1125
   End
   Begin VB.Label Label30 
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
      Height          =   495
      Left            =   1365
      TabIndex        =   49
      Top             =   1155
      Width           =   4635
   End
   Begin VB.Label Label28 
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
      Left            =   1365
      TabIndex        =   47
      Top             =   735
      Width           =   1965
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Producto:"
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
      Left            =   -315
      TabIndex        =   46
      Top             =   800
      Width           =   1545
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Lanzamiento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   3400
      TabIndex        =   41
      Top             =   165
      Width           =   1335
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
      Height          =   285
      Left            =   4830
      TabIndex        =   40
      Top             =   315
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "O/Fabricación"
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
      Left            =   105
      TabIndex        =   39
      Top             =   400
      Width           =   1125
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      Caption         =   "Fecha:"
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
      TabIndex        =   36
      Top             =   4830
      Width           =   1335
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
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
      TabIndex        =   35
      Top             =   4830
      Width           =   1335
   End
   Begin VB.Menu OpcionesGrillaLotes 
      Caption         =   "OpcionesGrillaLotes"
      Visible         =   0   'False
      Begin VB.Menu GrillaLotesEliminar 
         Caption         =   "Eliminar Lote"
      End
      Begin VB.Menu GrillaQuitarMarcaLectura 
         Caption         =   "Quitar Marcar de Lectura"
      End
   End
   Begin VB.Menu ConsultasListados 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu InformeTecnico 
         Caption         =   "Informe Tecnico en Word"
      End
   End
End
Attribute VB_Name = "TRAZORF14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BRTEMP As Recordset
Dim BAUTOTRA%
Dim ModoTrazabilidad$
Dim ExigeCalidadCompleto As Boolean
Dim FilaActiva As Long

Sub InformeTecnicaOFEnWord(OrdenFabricacion$)

    Dim oWord As cWord
    Dim Encabezado As Object
    Dim TABLA As Object
    Dim RstComponentes As ADODB.Recordset
    Dim hNew2 As Long
    
    ' INSTANCIA EL OBJETO DE LA CLASE cWord
    Set oWord = New cWord
       
    Call oWord.AgregaTabla(Encabezado, ";", 3, "Arial", 12, True)
    Encabezado.Borders.Enable = False
    Encabezado.Cell(1, 1).Range.InlineShapes.AddPicture (LUDAT$ & "LOGOASB.BMP")
    Encabezado.Cell(1, 2).Range.InsertAfter "REGISTRO DE TRAZABILIDAD  (F-7.1.5-00 Rev. 00)"
    CodigoConjunto$ = VALOBADA("ORDEFABR", "COD_EXTE", "IDSUBOR='" & OrdenFabricacion$ & "'")
    DescripcionConjunto$ = VALOBADA("ORDEFABR", "DESCRIP", "IDSUBOR='" & OrdenFabricacion$ & "'")
    Encabezado.Cell(3, 1).Range.InsertAfter "O-Fabricación / Lote: " & OrdenFabricacion$ & vbCrLf
    Encabezado.Cell(3, 1).Range.InsertAfter "Codigo: " & CodigoConjunto$ & vbCrLf
    Encabezado.Cell(3, 1).Range.InsertAfter "Descripción: " & DescripcionConjunto$
    
    CantidadFilas& = CantRegistros("MOVISTO", "CODIMOVI='CF' AND DOPRILAR='" & OrdenFabricacion$ & "'", , , "IDENLOTE")
    Call oWord.AgregaTabla(TABLA, "Codigo;Descripcion;Parte;Lote;PNL-PLC;Imagen", CantidadFilas&, "Arial", 10, True)
    
    ConsultaSql$ = "SELECT COD_EXTE, DESCRIP, IDENLOTE FROM MOVISTO WITH(NOLOCK) WHERE MOVISTO.CODIMOVI='CF' AND MOVISTO.DOPRILAR='" & OrdenFabricacion$ & "' GROUP BY COD_EXTE, DESCRIP, IDENLOTE"
    Set RstComponentes = READSET(ConsultaSql$)
    f& = 1
    With RstComponentes
        Do While Not .EOF
            f& = f& + 1
            LoteRecepcion$ = SAFETEXT$(!IDENLOTE)
            Parte$ = SAFETEXT$(VALOBADA("BOLRECEP", "NUMEPARTE", "IDENLOTE='" & LoteRecepcion$ & "'"))
            RutaFoto$ = SAFETEXT$(VALOBADA("INFOCALIDAD", "PATHFOTOCALIDAD", "DOCUMENTOORIGEN='" & LoteRecepcion$ & "'"))
            Clipboard.Clear
            If TRIM$(RutaFoto$) <> "" And Mid$(RutaFoto$, 1, 8) <> "@CAPTURA" Then
                Image2.Picture = LoadPicture(RutaFoto$)
                Image2.Refresh
                
                hNew2 = CopyImage(Image2.Picture, IMAGE_BITMAP, Val(320), Val(240), LR_COPYRETURNORG)
                OpenClipboard Me.hwnd
                EmptyClipboard
                SetClipboardData CF_BITMAP, hNew2
                CloseClipboard
                Image2.Picture = Clipboard.GetData(2)
    
                On Error Resume Next
                VUELTA% = 0
5               Clipboard.SetData Image2.Picture
                If Err > 0 And VUELTA% < 4 Then
                    VUELTA% = VUELTA% + 1
                    GoTo 5
                End If
                On Error GoTo 0
                
                If Abs(Image2.Picture) > 0 Then TABLA.Cell(f&, 6).Range.Paste
            End If
            TABLA.Cell(f&, 1).Range.InsertAfter SAFETEXT$(!COD_EXTE)
            TABLA.Cell(f&, 2).Range.InsertAfter SAFETEXT$(!DESCRIP)
            TABLA.Cell(f&, 3).Range.InsertAfter Parte$
            TABLA.Cell(f&, 4).Range.InsertAfter LoteRecepcion$
            TABLA.Cell(f&, 5).Range.InsertAfter " "
            
            ConsultaSql$ = "SELECT PNL_DES, PLC_HAS FROM MOVISTO WITH(NOLOCK) WHERE MOVISTO.CODIMOVI='CF' AND MOVISTO.DOPRILAR='" & OrdenFabricacion$ & "' AND IDENLOTE='" & LoteRecepcion$ & "'"
            Set RstPnlPlc = READSET(ConsultaSql$)
            PNL$ = "": PLC$ = "": TotalConsumos& = 0
            With RstPnlPlc
                Do While Not .EOF
                    PNL$ = PNL$ & AJUSTD$(!PNL_DES, 16)
                    PLC$ = PLC$ & AJUSTD$(!PLC_HAS, 16)
                    TotalConsumos& = TotalConsumos& + 1
                    .MoveNext
                Loop
                fil% = 1
                Set TablaAnidada = TABLA.Tables.Add(TABLA.Cell(f&, 5).Range, TotalConsumos& + 1, 2)
                TablaAnidada.Cell(1, 1).Range.InsertAfter "PNL"
                TablaAnidada.Cell(1, 2).Range.InsertAfter "PLC"
                For i& = 1 To Len(PNL$) Step 16
                    fil% = fil% + 1
                    TablaAnidada.Cell(fil%, 1).Range.InsertAfter TRIM$(Mid$(PNL$, i&, 16))
                    TablaAnidada.Cell(fil%, 2).Range.InsertAfter TRIM$(Mid$(PLC$, i&, 16))
                Next i&
            End With
            On Error Resume Next
            RstPnlPlc.Close
            Set RstPnlPlc = nohing
            On Error GoTo 0
            
            .MoveNext
        Loop
    End With
    
    RstComponentes.Close
    Set RstComponentes = Nothing
    
    Call oWord.GuardarDocumento(LUCOM$ & OrdenFabricacion$)
    Call oWord.CloseI
    
    Call COMUNI("Informe Técnico Generado en:\ \" & REPLACAR$(LUCOM$, "\", "/") & OrdenFabricacion$ & ".doc")

End Sub

Private Sub AGenRep_Click()
    '
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub

Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub

Private Sub Command1_Click()
   '
   If UCase$(TRIM$(Label7)) = "AUTO" Then
      If COMALTER%("¿ Realmente Cancela la Asignación\Automática de Lotes Propuesta ?\\No, Conservarla\Si, Cancelarla") = 2 Then GoTo 10
      Exit Sub
   End If
   '
10 Call CIERRARCH("*.*")
   Hide
   '
End Sub

Sub Command10_Click()

    Command10.Enabled = False
    If List1.ListIndex >= 0 Then
       FIKARDEX09.Text1 = Left$(List1.TEXT, 16)
    End If
    FIKARDEX09.Show 1
    Command10.Enabled = True
    
End Sub

Private Sub Command11_Click()
    Command11.Enabled = False
    Call SELECHO("MASTER")
    VDEVU$ = VALACT1$("MASTER")
    If VDEVU$ <> "" Then
       CICOMPO% = CODINT%(VDEVU$)
       CACONSU = 0
       TTYY$ = Space$(128)
       Call REPLA(TTYY$, VDEVU$, 1, 16)
       Call REPLA(TTYY$, DESCRIT$(CICOMPO%), 17, 40)
       Call REPLA(TTYY$, FORMATNUM$(CACONSU, "F13.1"), 57, 13)
       On Error Resume Next
       USOUNI = CACONSU / XVALO(Text1)
       On Error GoTo 0
       Call REPLA(TTYY$, FORMATNUM$(USOUNI, "F16.6"), 97, 16)
       List1.AddItem TTYY$
       List1.ListIndex = List1.ListCount - 1
       Text4 = ""
       Text4.SetFocus
    End If
    '
    Command11.Enabled = True
End Sub

Private Sub Command12_Click()
    Command12.Enabled = False
    If List1.ListIndex >= 0 Then
       ILISTA& = List1.ListIndex
       Call SELECHO("MASTER")
       VDEVU$ = VALACT1$("MASTER")
       If VDEVU$ <> "" Then
          CICOMPO% = CODINT%(VDEVU$)
          CACONSU = 0
          TTYY$ = Space$(128)
          Call REPLA(TTYY$, VDEVU$, 1, 16)
          Call REPLA(TTYY$, DESCRIT$(CICOMPO%), 17, 40)
          Call REPLA(TTYY$, FORMATNUM$(CACONSU, "F13.1"), 57, 13)
          On Error Resume Next
          USOUNI = CACONSU / XVALO(Text1)
          On Error GoTo 0
          Call REPLA(TTYY$, FORMATNUM$(USOUNI, "F16.6"), 97, 16)
          List1.AddItem TTYY$, ILISTA&
          List1.ListIndex = ILISTA&
          Text4 = ""
          Text4.SetFocus
       End If
    End If
    Command12.Enabled = True
End Sub

Private Sub Command13_Click()
    Command13.Enabled = False
    '
    ' VALIDAR FECHA
    HOII% = HOY(HO$)
    FECHA% = FECHANUM(Label24)
    If FECHA% < 1 Then
       Call MENSERR(24, "Fecha de Lanzamiento No Válida")
       GoTo 99
    End If
    '
    'VALIDAR CANTIDAD
    VALONU = XVALO(Label38)
    If VALONU < 1 Then
       Call MENSERR(24, "Cantidad No Válida")
       GoTo 99
    End If
    '
    CODICON$ = Label28
    IDMO$ = TRIM$(Label33)
    NORFA$ = Label32
    Call APERBASDAT("STOCKS")
    Call APERBASDAT("RESERVA")
    '
'    If IDMO$ = "" Then
'25     IDMO$ = NORFA$ + "/" + RANDSTRING$(-3)
'       Movisto.FindFirst "IdenMovi = '" & IDMO$ & "'"
'       If Movisto.NoMatch = False Then GoTo 25
'    End If
'    '
    ' GRABAR RESERVA
' CREAR UN RECORDSET MARESER PARA GRABAR LA RESERVA

' MODIFIQUE EL RECORDSET POR UN UPDATE MASIVO (OMAR)
    Call APERBASDAT("RESERVA")
    SQLX$ = "UPDATE RESERVA SET Status = 9 "
    SQLX$ = SQLX$ + " WHERE IdSubOr = '" & Label32 & "'"
    FLEXCONN.Execute (FILTSQL$(SQLX$))
    
'    SQLX$ = " SELECT * FROM RESERVA "
'    SQLX$ = SQLX$ + "WHERE IdSubOr = '" & Label32 & "' "
    'Set MARESER = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
    '
'    With MARESER
'       On Error Resume Next
'       .MoveFirst
'       If Err < 1 Then
'          On Error GoTo 0
'          Do While Not .EOF
'             .Edit
'             !Status = 9
'             .Update
'          .MoveNext
'          Loop
'       End If
'       On Error GoTo 0
'    End With
    '
    For KKI% = 0 To List1.ListCount
       TTXX$ = List1.LIST(KKI%)
       code$ = TRIM$(Left$(TTXX$, 16))
       CIXI% = CODINT%(code$)
       CACONSU = XVALO(Mid$(TTXX$, 57, 13))
       USOUN = CACONSU / VALONU
       '
       SUMATRA = 0
       For KKL& = 129 To Len(TTXX$) Step 64
          TTYY$ = Mid$(TTXX$, KKL&, 64)
          LOTE$ = TRIM$(Left$(TTYY$, 16))
          If LOTE$ <> "" Then
             CANTIX = XVALO(TRIM$(Mid$(TTYY$, 17, 16)))
             If CANTIX >= 0.05 Then
                PCOM$ = TRIM$(Mid$(TTYY$, 33, 16))
                PFIN$ = TRIM$(Mid$(TTYY$, 49, 16))
                SUMATRA = SUMATRA + CANTIX
                NORIT% = NORIT% + 1
                GoSub GRARESERVA
             End If
          End If
       Next KKL&
       '
       CANTIX = CACONSU - SUMATRA
       If CANTIX >= 0.05 Then
          LOTE$ = ""
          PCOM$ = ""
          PFIN$ = ""
          NORIT% = NORIT% + 1
          GoSub GRARESERVA
       End If
       '
    Next KKI%
    
    Dim MARESANT As ADODB.Recordset
    SQLX$ = "SELECT * FROM RESERVA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE IdSubOr = '" & Label32 & "' AND Status = 9 "
    Set MARESANT = READSET(SQLX$)
    '
    With MARESANT
       Do While Not .EOF
          CIXI% = !Cod_Inte
          LTXX$ = !IDENLOTE
          Call ACSALOTE(CIXI%, LTXX$, "NOMESS")
          .MoveNext
       Loop
    End With
    MARESANT.Close
    Set MARESANT = Nothing
    '
    SQLX$ = "DELETE FROM RESERVA "
    SQLX$ = SQLX$ + "WHERE IdSubOr = '" & Label32 & "' AND Status = 9 "
    FLEXCONN.Execute (FILTSQL$(SQLX$))
    '
    GoTo 99
    '
    '
    '
GRARESERVA:
     Call APERBASDAT("RESERVA")
     SQLX$ = " SELECT * FROM RESERVA "
     SQLX$ = SQLX$ + " WHERE IdSubOr <> '' and IDSUBOR = ''"
    Dim MARESER As ADODB.Recordset
    Set MARESER = New ADODB.Recordset
25   On Error Resume Next

     MARESER.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0
    With MARESER
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CodiEmpr = CodiEmp%
      On Error GoTo 0
      !Cod_Inte = CIXI%
      !COD_EXTE = CODEXT$(CIXI%)
      !DESCRIP = DESCRIT0$(CIXI%)
      !IDSUBOR = NORFA$
      !Status = 0
      !NuOrItem = NORIT%
      !FechRes = CVDAT(FECHANUM(Label24))
      !CANTRES = CANTIX
      !CantCons = 0
      !Saldo_Entre = CANTIX
        LOTEA$ = ""
        On Error Resume Next
        LOTEA$ = !IDENLOTE
        On Error GoTo 0
      !IDENLOTE = LOTE$
      !Deposito = DEPOPRE(CIXI%)
      !UsoUnit = USOUN
      !OBSERVA = ""
      !PNL_DES = PCOM$
      !PLC_HAS = PFIN$
      .Update
      '
      Call ACSALOTE(CIXI%, LOTE$, "NOMESS")
      '
    End With
    MARESER.Close
    Set MARESER = Nothing
Return
'
99  Screen.MousePointer = 1
    Command13.Enabled = True
    '
    If UCase$(TRIM$(Label7)) = "AUTO" Then
       Call CIERRARCH("*.*")
       Hide
    End If
    '
End Sub

Sub Command14_Click()
    '\\\OT-06-0476-RG-07/12/06///
    '
    If UCase$(TRIM$(Label7)) = "AUTO" Then Exit Sub
    '
    OPTRAZ07.Show 1
    '\\\OT-06-0476-RG-FIN///
End Sub

Private Sub Command15_Click()
    Command15.Enabled = False
    If Check1.Value = 1 Then
        Call TrazabilidadAutoFifoConPackingList
        If mfgLotes.Rows > 1 Then mfgLotes.Row = 1
        Text1.SetFocus
    Else
        BAUTOTRA% = 1
        For KKU& = 1 To List1.ListCount
          List1.Selected(KKU& - 1) = True
          Call AUTOTRA
        Next KKU&
        BAUTOTRA% = 0
    End If
    Command15.Enabled = True
End Sub

Sub TrazabilidadAutoFifoConPackingList(Optional modoSoloItem As Boolean)

    Dim CodigoInterno%
    Call APERBASDAT("STOCKS")
    mfgLotes.Rows = 1
    
    iniComponentes& = 1: finComponentes& = List1.ListCount
    If modoSoloItem = True Then
        If List1.ListIndex < 0 Then Exit Sub
        iniComponentes& = List1.ListIndex + 1: finComponentes& = iniComponentes&
    End If
    
    For KKU& = iniComponentes& To finComponentes&
        
        ' ARTICULO A TRAZAR
        List1.Selected(KKU& - 1) = True
        CantidadTrazada# = XVALO(Label16)
        cantidadRequerida# = XVALO(Text4)

        ' OBTIENE LOS LOTES DISPONIBLES ORDENAS SEGUN FIFO
        filaList1$ = List1.LIST(KKU& - 1)
        CodigoInterno% = CODINT%(Mid$(filaList1$, 1, 16))
        ConsultaSql$ = " SELECT * FROM LOTINGRE WITH(NOLOCK)"
        ConsultaSql$ = ConsultaSql$ + " WHERE Cod_Inte = " & CStr(CodigoInterno%) & " "
        ConsultaSql$ = ConsultaSql$ + " ORDER BY FechAlta ASC, IdenLote ASC "
        Dim LOTINGTEMP As ADODB.Recordset
        Set LOTINGTEMP = READSET(ConsultaSql$)
        '
        With LOTINGTEMP
25         On Error Resume Next
           .MoveFirst
           If Err < 1 Then
             On Error GoTo 0
             Do While Not .EOF
                ' CALCULA SALDO LOGICO (INCLUYE LO YA RESERVADO DE ESTE ITEM PARA ESTA OF)
                SaldoLogico# = XVALO(!CanAlta) - XVALO(!CanConsu) - XVALO(!CanReser): If SaldoLogico# < 0 Then SaldoLogico# = 0
               
                aTrazar# = cantidadRequerida# - CantidadTrazada#
                If SaldoLogico# < aTrazar# Then aTrazar# = SaldoLogico#
                '
                If aTrazar# > 0.05 Then
                    filaMfgLotes& = mfgLotes.Rows
                    mfgLotes.Rows = mfgLotes.Rows + 1
                    mfgLotes.Row = filaMfgLotes&
                    '
                    codigoInternoArticulo% = XVALO(!Cod_Inte)
                    LoteRecepcion$ = SAFETEXT(!IDENLOTE)
                    mfgLotes.TextMatrix(filaMfgLotes&, 1) = CODEXT$(codigoInternoArticulo%)
                    mfgLotes.TextMatrix(filaMfgLotes&, 2) = DESCRIT$(codigoInternoArticulo%)
                    mfgLotes.TextMatrix(filaMfgLotes&, 3) = LoteRecepcion$
                    mfgLotes.TextMatrix(filaMfgLotes&, 4) = FORMATNUM$(aTrazar#, "F9.1")
                    LoteRecepcionFraccionado$ = TRIM$(SAFETEXT$(!LOTECOL_PROV))
                    ' EXCLUYE DE LA TRAZABILIDAD VIA CODIGO DE BARRAS, A LOS AJUSTES Y A LAS O-FABRICACION
                    If LoteRecepcionFraccionado$ = "" Then LoteRecepcionFraccionado$ = LoteRecepcion$
                    If CantRegistros("BOLRECEP", "IDENLOTE='" & LoteRecepcion$ & "'") = 0 And CantRegistros("BOLRECEP", "IDENLOTE='" & LoteRecepcionFraccionado$ & "'") = 0 Then
                        mfgLotes.TextMatrix(filaMfgLotes&, 0) = "*"
                        mfgLotes.TextMatrix(filaMfgLotes&, 5) = FORMATNUM$(aTrazar#, "F9.1")
                    End If
                    
                    If TieneRegistroCalidad(LoteRecepcion$, observaciones$) = False And ExigeCalidadCompleto = True Then
                        Call COLOREAR_GRILLA_SOLO_UNA_AFILA(mfgLotes, &HC0E0FF, filaMfgLotes&)
                        mfgLotes.TextMatrix(filaMfgLotes&, 0) = "(E)"
                        mfgLotes.TextMatrix(filaMfgLotes&, 6) = observaciones$
                    End If
                    
                    CantidadTrazada# = CantidadTrazada# + aTrazar#
                End If
               
             .MoveNext
             Loop
            End If
        End With
        
        On Error Resume Next
        LOTINGTEMP.Close
        Set LOTINGTEMP = Nothing
        On Error GoTo 0
    Next KKU&
        
End Sub

Function TieneRegistroCalidad(LoteRecepcion As String, observaciones As String) As Boolean

    TieneRegistroCalidad = True
    Embalaje$ = ""
    Encapsulado$ = ""
    Certificado$ = ""
    pathFoto$ = ""
    estadoCalidad$ = ""
    numeroParte$ = ""
    
    Call ABRECONEXION
    
    ' SE VALIDAN SOLO AQUELLOS LOTES QUE CORRESPONDEN A RECEPCIONES
    If CantRegistros("BOLRECEP", "IDENLOTE='" & LoteRecepcion & "'") = 0 Then GoTo 99
    
    ConsultaSql$ = "SELECT EMBALAJE, ENCAPSULADO, CERTIFICADO, STATUS, PATHFOTOCALIDAD FROM INFOCALIDAD WITH(NOLOCK) WHERE DOCUMENTOORIGEN='" & LoteRecepcion & "'"
    Set RST = READSET(ConsultaSql$)
    
    With RST
        Do While Not .EOF
            Embalaje$ = TRIM$(!Embalaje)
            Encapsulado$ = TRIM$(!Encapsulado)
            Certificado$ = TRIM$(!Certificado)
            pathFoto$ = TRIM$(!PATHFOTOCALIDAD)
            estadoCalidad$ = TRIM$(!Status)
            .MoveNext
        Loop
    End With
    
    On Error Resume Next
    RST.Close
    Set RST = Nothing
    On Error GoTo 0
    
    numeroParte$ = TRIM$(VALOBADA("BOLRECEP", "NUMEPARTE", "IDENLOTE='" & LoteRecepcion & "'"))
    
    HUBOERROR% = 0
    observaciones = "REGISTRO DE CALIDAD INCOMPLETO. Falta "
    If Embalaje$ = "" Then
        observaciones = observaciones & "embalaje, ": HUBOERROR% = 1
    End If
    If Encapsulado$ = "" Then
        observaciones = observaciones & "encapsulado, ": HUBOERROR% = 1
    End If
    If Certificado$ = "" Then
        observaciones = observaciones & "certificado, ": HUBOERROR% = 1
    End If
    If pathFoto$ = "" Or Mid$(pathFoto$, 1, 1) = "@" Then
        observaciones = observaciones & "foto, ": HUBOERROR% = 1
    End If
    If estadoCalidad$ = "" Then
        observaciones = observaciones & "estado de calidad, ": HUBOERROR% = 1
    End If
    If numeroParte$ = "" Then
        observaciones = observaciones & "numero de parte, ": HUBOERROR% = 1
    End If
    observaciones = TRIM$(Mid$(observaciones, 1, Len(observaciones) - 1))
    If HUBOERROR% = 1 Then TieneRegistroCalidad = False

99  Exit Function

End Function

Private Sub Command16_Click()
   
   If UCase$(TRIM$(Label7)) = "AUTO" Then Exit Sub
   '
   Command16.Enabled = False
   NORFA$ = TRIM$(Label32)
   If NORFA$ = "" Then
      Call ABREORFAB
      Call SELECHO("ECORDEF/Indice General de Ordenes de Fabricación (Pendientes-Cumplidas-Anuladas)")
      NORFA$ = TRIM$(VALACT1$("ECORDEF"))
   End If
   If NORFA$ = "" Then GoTo 99
   '
   FORIPRE$ = TRIM$(Control$("TRAZABI", "FORTRAZA"))
   If FORIPRE$ <> "" Then
     If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") > 0 Then
       GoTo 10
     End If
   End If
   Call MENSERR(24, "Formulario de Impresión Inexistente\o no Definido.")
   GoTo 99
   '
   Dim CODEX$(4096), CACONSU(4096), IDLOT$(4096), PNLD$(4096), PLCH$(4096)
   '
10 Screen.MousePointer = 11
   Call APERBASDAT("STOCKS")
   FEX = HOY(HOS$)
   FECHDOC$ = HOS$
   NUMEDOC$ = NORFA$
   TIPODOC$ = "Registro de Trazabilidad"
   '
   CODPARAM$(1) = "REF-MAT1"
   CODPARAM$(2) = "MEDIDA-1"
   CODPARAM$(3) = "MEDIDA-2"
   CODPARAM$(4) = "MEDIDA-3"
   CAPARDOC% = 8
   '
   DOCPARAM$(1) = Text8
      If TRIM$(Text8) = "" Then
         DOCPARAM$(1) = NORFA$
      End If
   DOCPARAM$(2) = TRIM$(Label30)
   DOCPARAM$(3) = TRIM$(Text5)
   DOCPARAM$(4) = TRIM$(Text6)
   '
   CODTABU1$(1) = "COD-MAT"
   CODTABU1$(2) = "DES-MAT"
   CODTABU1$(3) = "N-PLANO"
   CODTABU1$(4) = "R-PLANO"
   CODTABU1$(5) = "F-PLANO"
   CODTABU1$(6) = "RASO-CLI"
   CODTABU1$(7) = "Z-PLANO"
   CODTABU1$(8) = "FECH-REC"
   CODTABU1$(9) = "FECH-VEN"
   CODTABU1$(10) = "NRO-LOTE"
   CODTABU1$(11) = "N-AUTENT"
   CODTABU1$(12) = "REF-MAT2"
   CODTABU1$(13) = "N-HPROF"
   CODTABU1$(14) = "N-HINSP"
   CACOLTAB1% = 14
   '
   Call APERBASDAT("STOCKS")
   SQLT$ = "SELECT CANTSALID, COD_INTE, CANTSALID, IdenLote,  PNL_DES, PLC_HAS  FROM MOVISTO WITH(NOLOCK)"
   SQLT$ = SQLT$ + " WHERE CodiMovi = 'CF' "
   SQLT$ = SQLT$ + " AND DoPriCor = '" & NORFA$ & "' "
   SQLT$ = SQLT$ + "ORDER BY NuOrItem ASC "
   'Set MACONSUM = Stocks.OpenRecordset(SQLT$, dbOpenSnapshot)
'   Dim MACONSUM As ADODB.Recordset
'   Set MACONSUM = New ADODB.Recordset
'   MACONSUM.Open FILTSQL$(SQLT$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   Set MACONSUM = READSET(SQLT$)
   '
15 JJ& = 0
   With MACONSUM
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
          If !Cantsalid >= 0.05 Then
            JJ& = JJ& + 1
            CICOMPO% = !Cod_Inte
            CODEX$(JJ&) = CODEXT$(CICOMPO%)
            CACONSU(JJ&) = !Cantsalid
            IDLOT$(JJ&) = !IDENLOTE
              PNLD(JJ&) = "": PLCH(JJ&) = ""
              On Error Resume Next
            PNLD(JJ&) = !PNL_DES
            PLCH(JJ&) = !PLC_HAS
              On Error GoTo 0
          End If
        .MoveNext
        Loop
     End If
   End With
   MACONSUM.Close
   Set MACONSUM = Nothing
   '
   ' AQUI AGREGAR LOS DE RESERVAS
   Call APERBASDAT("RESERVA")
   SQLX$ = " SELECT CantRes, CANTCONS, IdenLote, COD_INTE, IdenLote, PNL_DES, PLC_HAS FROM RESERVA WITH(NOLOCK) "
   SQLX$ = SQLX$ + "WHERE IdSubOr = '" & NORFA$ & "' "
   'Set RESERTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
'   Dim RESERTMP As ADODB.Recordset
'   Set RESERTMP = New ADODB.Recordset
'   RESERTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   Set RESERTMP = READSET(SQLX$)
   With RESERTMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
          On Error GoTo 0
          Do While (.EOF = False)
            CACOLO = !CANTRES - !CantCons: If CACOLO < 0 Then CACOLO = 0
            If CACOLO >= 0.05 Then
               If TRIM$(!IDENLOTE) <> "" Then    ' van solo las reservas trazadas
                  JJ& = JJ& + 1
                  CICOMPO% = !Cod_Inte
                  CODEX$(JJ&) = CODEXT$(CICOMPO%)
                  CACONSU(JJ&) = CACOLO
                  IDLOT$(JJ&) = TRIM$(!IDENLOTE)
                    PNLD(JJ&) = "": PLCH(JJ&) = ""
                    On Error Resume Next
                  PNLD(JJ&) = !PNL_DES
                  PLCH(JJ&) = !PLC_HAS
                    On Error GoTo 0
               End If
            End If
19        .MoveNext
          Loop
       End If
       On Error GoTo 0
   End With
   RESERTMP.Close
   Set RESERTMP = Nothing
   '
   ' AQUI ORDENAR ASCENDENTE POR CODEX$
   For U1& = 1 To JJ&
     For U2& = U1& + 1 To JJ&
       If CODEX$(U1&) > CODEX$(U2&) Then
         SW = SWAP(CODEX$(U1&), CODEX$(U2&))
         SW = SWAP(CACONSU(U1&), CACONSU(U2&))
         SW = SWAP(IDLOT$(U1&), IDLOT$(U2&))
         SW = SWAP(PNLD$(U1&), PNLD$(U2&))
         SW = SWAP(PLCH$(U1&), PLCH$(U2&))
         KLJ = KKK
       End If
     Next U2&
   Next U1
   '
   CAITAB1% = 0
   For U& = 1 To JJ&
      code$ = CODEX$(U&)
      CIX0% = CODINT%(code$)
      '
      IDLOTE$ = IDLOT$(U&)
      CACOLO = CACONSU(U&)
      If CACOLO >= 0.05 Then
         PDES$ = PNLD(U&)
         PHAS$ = PLCH(U&)
         '
         TIPOX$ = ""
         MARCAX$ = ""
         PARTNUX$ = ""
         PROVX$ = ""
         FINGRE = 0
         FEFAB = 0
         TRACLIX$ = ""
         VERIX$ = ""
         If IDLOTE$ <> "" Then
            SQLX$ = "SELECT Tipo, MARCA, NUMEPARTE, Raso_Prov, FechRecep, FechFab, COD_CLI, TRAZ_CLI, USUAPRO FROM BOLRECEP WITH(NOLOCK) "
            SQLX$ = SQLX$ + " WHERE IdenLote = '" & IDLOTE$ & "'"
            'Set BRSELE = Stocks.OpenRecordset(SQLX, dbOpenSnapshot)
'            Dim BRSELE As ADODB.Recordset
'            Set BRSELE = New ADODB.Recordset
'            BRSELE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
            Set BRSELE = READSET(SQLX$)
            With BRSELE
              On Error Resume Next
              .MoveFirst
              If Err Then
                 On Error GoTo 0
                 ' Suprime mensaje de error porque puede provenir de una O/F
                  ' Call MENSERR(24, "No Encontrado Lote '" + IDLOTE$ + "'")
                 GoTo 30
              End If
              On Error GoTo 0
              '
              TIPOX$ = "": MARCAX$ = "": PARTNUX$ = "": PROVX$ = "": TRACLIX$ = "": VERIX$ = "": COLOTX$ = ""
              On Error Resume Next
              TIPOX$ = TRIM$(!Tipo)
              MARCAX$ = !MARCA
              PARTNUX$ = !NUMEPARTE
              PROVX$ = !Raso_Prov
              FINGRE = CVINT(!FechRecep)
              FEFAB = CVINT(!FechFab)
              COLOTX$ = !COD_CLI
              TRACLIX$ = !TRAZ_CLI
              VERIX$ = !USUAPRO
              On Error GoTo 0
              '
            End With
            BRSELE.Close
            Set BRSELE = Nothing
         End If
         '
30       CAITAB1% = CAITAB1% + 1
         TETABU1$(1, CAITAB1%) = code$
         TETABU1$(2, CAITAB1%) = DESCRIT0$(CIX0%)
           If TIPOX$ = "" Then TIPOX$ = ATRIB1$(CIX0%)
         TETABU1$(3, CAITAB1%) = TIPOX$
         TETABU1$(4, CAITAB1%) = MARCAX$
         TETABU1$(5, CAITAB1%) = PARTNUX$
         TETABU1$(6, CAITAB1%) = PROVX$
         If Left$(IDLOTE$, 2) = "BR" Then
            TETABU1$(7, CAITAB1%) = IDLOTE$
         End If
         TETABU1$(8, CAITAB1%) = FECHATEX$(FINGRE)
         TETABU1$(9, CAITAB1%) = FECHATEX$(FEFAB)
         If Left$(IDLOTE$, 2) <> "BR" Then
              TETABU1$(10, CAITAB1%) = IDLOTE$              ' NUMERO DE LOTE
            Else
              TETABU1$(10, CAITAB1%) = COLOTX$
         End If
         TETABU1$(11, CAITAB1%) = TRACLIX$
         TETABU1$(12, CAITAB1%) = VERIX$
         TETABU1$(13, CAITAB1%) = PDES$
         TETABU1$(14, CAITAB1%) = PHAS$
      End If
   Next U&
   '
   Call PRINTDOC("@" + FORIPRE$)
   '
99 Screen.MousePointer = 1
   Command16.Enabled = True
   '
End Sub

Private Sub Command17_Click()
    
    Command17.Enabled = False
    
    If mfgLotes.Row > 0 Then
        LoteRecepcion$ = TRIM$(mfgLotes.TextMatrix(mfgLotes.Row, 3))
        LoteRecepcionFraccion$ = TRIM$(SAFETEXT$(VALOBADA("LOTINGRE", "LOTECOL_PROV", "IDENLOTE='" & LoteRecepcion$ & "'")))
        If CantRegistros("BOLRECEP", "IDENLOTE='" & LoteRecepcion$ & "'") = 1 Or CantRegistros("BOLRECEP", "IDENLOTE='" & LoteRecepcionFraccion$ & "'") = 1 Then
        
            opcionCalidad% = COMALTER%("Opciones de Calidad:\\Informar Registro\Imprimir Etiqueta")
            If opcionCalidad% < 1 Or opcionCalidad% > 2 Then GoTo 99
            
            LoteRecepcion$ = TRIM$(mfgLotes.TextMatrix(mfgLotes.Row, 3))
            codigoArticulo$ = TRIM$(mfgLotes.TextMatrix(mfgLotes.Row, 1))
            CodigoInterno% = CODINT%(codigoArticulo$)
                
            If opcionCalidad% = 1 Then
                descripcionArticulo$ = TRIM$(mfgLotes.TextMatrix(mfgLotes.Row, 2))
                Call FILTERFILE("TAMARCAS", "!TAMARCAS", 4, MKI$(CodigoInterno%))
                If LoteRecepcionFraccion$ <> "" Then LoteRecepcion$ = LoteRecepcionFraccion$
                
                FILTX$ = LoteRecepcion$ & ";" & codigoArticulo$ & ";" & descripcionArticulo$
                formularioCalidad$ = TRIM$(Control("TRAZABI", "REGICAL"))
                Call STDFORM(formularioCalidad$, FILTX$)
            Else
                CantidadLote# = XVALO(VALOBADA("LOTINGRE", "CANALTA", "IDENLOTE='" & LoteRecepcion$ & "'"))
                
                Call PRINSTIKAPRO(CodigoInterno%, LoteRecepcion$, 1, CantidadLote#)
            End If
        Else
            Call COMUNI("El Lote No Requiere Registro de Calidad")
        End If
    Else
        Call COMUNI("No Existen Lotes de Recepción")
    End If
   
99  Command17.Enabled = True
    
End Sub

Private Sub Command18_Click()

    ' LOTES QUE PROVIENEN DEL AUTOFIFO
    For FilaGrilla% = 1 To mfgLotes.Rows - 1
    
        aTrazar# = XVALO(mfgLotes.TextMatrix(FilaGrilla%, 5))
        If aTrazar# > 0.05 And TRIM$(mfgLotes.TextMatrix(FilaGrilla%, 0)) = "*" Then
            codigoArticulo$ = TRIM$(mfgLotes.TextMatrix(FilaGrilla%, 1))
        
            ' UBICA EL COMPONENTE EN LA LISTA DE MATERIALES RESERVADOS
            articuloLista% = 0
            For ind% = 0 To List1.ListCount - 1
                CodigoComponente$ = TRIM$(Mid$(List1.LIST(ind%), 1, 16))
                If CodigoComponente$ = codigoArticulo$ Then
                    articuloLista% = ind%: Exit For
                End If
            Next ind%
            
            renglonComponente$ = List1.LIST(articuloLista%)
            lotesTrazados$ = Mid$(renglonComponente$, 129)
            LoteRecepcion$ = TRIM$(mfgLotes.TextMatrix(FilaGrilla%, 3))

            ' ACTUALIZA LOTES TRAZADOS
            If TRIM$(Mid$(lotesTrazados$, 1, 16)) = "" Then renglonComponente$ = Mid$(renglonComponente$, 1, 128)
            renglonComponente$ = renglonComponente$ & AJUSTI$(LoteRecepcion$, 16) & FORMATNUM$(aTrazar#, "F16.1") & Space(32)
            List1.LIST(articuloLista%) = renglonComponente$
            
            ' CALCULA EL TOTAL DE LO TRAZADO PARA LUEGO ACTUALIZAR ESTE DATO EN EL LIST
            renglonComponente$ = List1.LIST(articuloLista%)
            lotesTrazados$ = Mid$(renglonComponente$, 129)
            totalTrazado# = 0
            For indLotes% = 1 To Len(lotesTrazados$) Step 64
                loteTrazado$ = Mid$(lotesTrazados, indLotes%, 64)
                LOTE$ = TRIM$(Mid$(loteTrazado$, 1, 16))
                CantidadTrazada# = XVALO(Mid$(loteTrazado$, 17, 16))
                
                If LOTE$ <> "" Then totalTrazado# = totalTrazado# + CantidadTrazada#
            Next indLotes%
            
            ' ACTUALIZA TOTAL TRAZADO QUE SE OBSERVA EN EL LIST
            Call REPLA(renglonComponente$, FORMATNUM$(totalTrazado#, "F12.1"), 71, 12)
            List1.LIST(articuloLista%) = renglonComponente$
        End If
        
    Next FilaGrilla%
    
    ' LOTES QUE NO PROVIENEN DEL AUTOFIFO
    For FilaGrilla% = 1 To mfgLotes.Rows - 1
    
        If TRIM$(mfgLotes.TextMatrix(FilaGrilla%, 0)) = "¿?" Then
            codigoArticulo$ = TRIM$(mfgLotes.TextMatrix(FilaGrilla%, 1))
            LoteRecepcion$ = TRIM$(mfgLotes.TextMatrix(FilaGrilla%, 3))
            
            ' UBICA EL COMPONENTE EN LA LISTA DE MATERIALES RESERVADOS
            articuloLista% = 0
            For ind% = 0 To List1.ListCount - 1
                CodigoComponente$ = TRIM$(Mid$(List1.LIST(ind%), 1, 16))
                If CodigoComponente$ = codigoArticulo$ Then
                    articuloLista% = ind%: Exit For
                End If
            Next ind%
            
            ' CALCULA EL SALDO PENDIENTE A TRAZAR
            renglonComponente$ = List1.LIST(articuloLista%)
            lotesTrazados$ = Mid$(renglonComponente$, 129)
            totalTrazado# = 0
            For indLotes% = 1 To Len(lotesTrazados$) Step 64
                loteTrazado$ = Mid$(lotesTrazados, indLotes%, 64)
                LOTE$ = TRIM$(Mid$(loteTrazado$, 1, 16))
                CantidadTrazada# = XVALO(Mid$(loteTrazado$, 17, 16))
                
                If LOTE$ <> "" Then totalTrazado# = totalTrazado# + CantidadTrazada#
            Next indLotes%
            CantidadReservada# = XVALO(Mid$(renglonComponente$, 57, 13))
            SaldoATrazar# = CantidadReservada# - totalTrazado#
            If SaldoATrazar# < 0.05 Then
                Call COMUNI("Reserva Completa para el Artículo '" & CodigoProducto$ & "'.\No se Trazará el Lote '" & LoteRecepcion & "'")
            End If
            
            ' OBTENER SALDO DEL LOTE
            SaldoDelLote# = SaldoLote(CODINT%(codigoArticulo$), LoteRecepcion$)
            
            ' CALCULAR CANTIDAD A TRAZAR
            aTrazar# = SaldoATrazar#
            If SaldoATrazar# > SaldoDelLote# Then
                aTrazar# = SaldoDelLote#
            End If
            
            ' ACTUALIZA LOTES TRAZADOS
            lotesTrazados$ = Mid$(renglonComponente$, 129)
            If TRIM$(Mid$(lotesTrazados$, 1, 16)) = "" Then renglonComponente$ = Mid$(renglonComponente$, 1, 128)
            renglonComponente$ = renglonComponente$ & AJUSTI$(LoteRecepcion$, 16) & FORMATNUM$(aTrazar#, "F16.1") & Space(32)
            List1.LIST(articuloLista%) = renglonComponente$
            
            ' CALCULA EL TOTAL DE LO TRAZADO PARA LUEGO ACTUALIZAR ESTE DATO EN EL LIST
            renglonComponente$ = List1.LIST(articuloLista%)
            lotesTrazados$ = Mid$(renglonComponente$, 129)
            totalTrazado# = 0
            For indLotes% = 1 To Len(lotesTrazados$) Step 64
                loteTrazado$ = Mid$(lotesTrazados, indLotes%, 64)
                LOTE$ = TRIM$(Mid$(loteTrazado$, 1, 16))
                CantidadTrazada# = XVALO(Mid$(loteTrazado$, 17, 16))
                
                If LOTE$ <> "" Then totalTrazado# = totalTrazado# + CantidadTrazada#
            Next indLotes%
            
            ' ACTUALIZA TOTAL TRAZADO QUE SE OBSERVA EN EL LIST
            Call REPLA(renglonComponente$, FORMATNUM$(totalTrazado#, "F12.1"), 71, 12)
            List1.LIST(articuloLista%) = renglonComponente$
        End If
        
40  Next FilaGrilla%
    
    
    Command19.Visible = False
    mfgLotes.Rows = 1
    Call Command13_Click
    List1.ListIndex = -1
    
End Sub

Private Sub Command19_Click()

   PopupMenu OpcionesGrillaLotes
    
End Sub

Private Sub Command2_Click()

   Command2.Enabled = False
   '
   CodigoConjunto$ = Label28
   CodigoIntConjunto% = CODINT%(CodigoConjunto$)
   CodigoComponente$ = Label27
   CodigoIntComponente% = CODINT%(CodigoComponente$)
   UsoUnitarioComponente# = XVALO(VALOBADA("FORMULAS", "USOBRUTO", "CODICONJ=" & CStr(CodigoIntConjunto%) & " AND CODIELEM=" & CStr(CodigoIntComponente%)))
   SUMATRA = 0
   For KKU& = 0 To 5
      TTYY$ = Space$(64)
      If TRIM$(Label23(KKU&)) <> "" Then
        If XVALO(Text11(KKU&)) >= 0.05 Then
            SUMATRA = SUMATRA + XVALO(Text11(KKU&))
'            If TieneRegistroCalidad(TRIM$(Label23(KKU&)), observaciones$) = True Or ExigeCalidadCompleto = False Then
'                SUMATRA = SUMATRA + XVALO(Text11(KKU&))
'            End If
            ' CONTROLA EL TOTAL DE COMPONENTES EN CONTRASTE CON LA CANT. DE PLACAS. EN PLC Y PNL INGRESAN CANTIDAD DE PLACAS,
            ' MIENTRAS QUE EN LA CANTIDAD RESERVADA INFORMAN CANTIDAD DE COMPONENTES. LA CANTIDAD DE COMPONENTES
            ' POR PLACA ESTA DADO EN EL USO UNITARIO DE LA FORMULA.
'            CantidadPlacasDesde# = XVALO(Text16(KKU&))
'            CantidadPlacasHasta# = XVALO(Text17(KKU&))
'            TotalPlacas# = (CantidadPlacasHasta# - CantidadPlacasDesde# + 1) * UsoUnitarioComponente#
'            SaldoDelLote# = XVALO(Text11(KKU&))
'            If TotalPlacas# - SaldoDelLote# > 0 And (CantidadPlacasHasta# - CantidadPlacasDesde#) > 0 Then
'                Call MENSERR(24, "Imposible Registrar.\ \La Cantidad de Componentes Resultante del Número de Placas\es Mayor a la Cantidad Reservada del Lote '" & TRIM$(Label23(KKU&)) & "'\ \" _
'                & "Componentes Según Placas: " & TRIM$(FORMATNUM$(TotalPlacas#, "F12.1")))
'                GoTo 99
'            End If
        End If
      End If
   Next KKU&
   If SUMATRA > XVALO(Text4) + 0.005 Then
      Call MENSERR(24, "Imposible Registrar.\Los Lotes Trazados Superan el Total Consumido.")
      GoTo 99
   End If
   '
   TTXX$ = Left$(List1.TEXT + Space$(128), 128)
   CACONSU = XVALO(Text4)
   Call REPLA(TTXX$, FORMATNUM$(CACONSU, "F13.1"), 57, 13)
   On Error Resume Next
   USOUNI = CACONSU / XVALO(Text1)
   On Error GoTo 0
   Call REPLA(TTXX$, FORMATNUM$(USOUNI, "F16.6"), 97, 16)
   Call REPLA(TTXX$, FORMATNUM$(SUMATRA, "F12.1"), 71, 12)
   '
   For KKU& = 0 To 5
      TTYY$ = Space$(64)
      If TRIM$(Label23(KKU&)) <> "" Then
        If XVALO(Text11(KKU&)) >= 0.05 Then
            If TieneRegistroCalidad(TRIM$(Label23(KKU&)), observaciones$) = True Or ExigeCalidadCompleto = False Then
                Call REPLA(TTYY$, TRIM$(Label23(KKU&)), 1, 16)
                Call REPLA(TTYY$, Text11(KKU&), 17, 16)
                Call REPLA(TTYY$, Text16(KKU&), 33, 16)
                Call REPLA(TTYY$, Text17(KKU&), 49, 16)
                TTXX$ = TTXX$ + TTYY$
            Else
                Call COMUNI("Imposible Asignar Lote.\ \Lote '" & TRIM$(Label23(KKU&)) & "'\con Registro de Calidad Incompleto.")
                GoTo 99
            End If
        End If
      End If
   Next KKU&
   '
   ILISTA& = XVALO(Label8)
   If List1.ListCount > 0 Then
      If ILISTA& >= 0 Then
        List1.RemoveItem (ILISTA&)
        List1.AddItem TTXX$, ILISTA&
        List1.ListIndex = ILISTA&
      End If
   End If
   '
99 Command2.Enabled = True

End Sub

Private Sub Command20_Click()
    
    mfgLotes.Rows = 1
    Command19.Visible = False
    
End Sub



Private Sub Command22_Click(Index As Integer)

    LoteRecepcion$ = TRIM$(Label23(Index))
    
    If LoteRecepcion$ <> "" Then
        If List1.ListIndex > -1 Then
            codigoArticulo$ = TRIM$(Mid$(List1.LIST(List1.ListIndex), 1, 16))
            CodigoInterno% = CODINT%(codigoArticulo$)
            
            opcionCalidad% = COMALTER%("Opciones de Calidad:\\Informar Registro\Imprimir Etiqueta")
            If opcionCalidad% < 1 Or opcionCalidad% > 2 Then Exit Sub
            
            If opcionCalidad% = 1 Then
                descripcionArticulo$ = TRIM$(DESCRIT0$(CodigoInterno%))
                Call FILTERFILE("TAMARCAS", "!TAMARCAS", 4, MKI$(CodigoInterno%))
                
                FILTX$ = LoteRecepcion$ & ";" & codigoArticulo$ & ";" & descripcionArticulo$
                formularioCalidad$ = TRIM$(Control("TRAZABI", "REGICAL"))
                Call STDFORM(formularioCalidad$, FILTX$)
            Else
                CantidadLote# = XVALO(VALOBADA("LOTINGRE", "CANALTA", "IDENLOTE='" & LoteRecepcion$ & "'"))
                
                Call PRINSTIKAPRO(CodigoInterno%, LoteRecepcion$, 1, CantidadLote#)
            End If
        End If
    Else
        Call COMUNI("Falta Elegir un Lote")
    End If

End Sub

Private Sub Command23_Click()

    OrdenFabricacion$ = Label32.Caption
    Label32.Caption = ""
    Label32.Caption = OrdenFabricacion$
    
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   CUDI$ = UCase$(CurDir)
   PPXX% = InStr(CUDI$, "\WINDOWS")
   If PPXX% > 0 Then
      RUTAX$ = Left$(CUDI$, PPXX%) + "MANUALES\"
'      If EXISTE%(RUTAX$ + "TRAZAB04.PDF") > 0 Then
'         FORHLPDF.Pdf1.setShowToolbar (True)
'         On Error Resume Next
'         FORHLPDF.Pdf1.LOADFILE RUTAX$ + "TRAZAB04.pdf"
'         If Err Then GoTo 20
'         On Error GoTo 0
'         FORHLPDF.Show 1
'         GoTo 99
'      End If
   End If
   '
20 Call MENSERR(24, "Imposible Abrir Archivo de Ayuda")
   '
99 Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Call CommSel("SELFECHA", Command4, Label24)
End Sub

Private Sub Command5_Click()
    '
    If UCase$(TRIM$(Label7)) = "AUTO" Then Exit Sub
    '
    Call ACTECOR(2)     ' para seleccionar solo las activas lanzadas
    Call CommSel("ECORDEP/Indice de Ordenes de Fabricación Pendientes", Command5, Label32)
    If Label32 <> "" Then
       On Error Resume Next
       Text8.SetFocus
       On Error GoTo 0
    End If
    '
End Sub

Sub TAMARCAS()
    '
    Ci% = CODINT%(Label27)
    If Ci% > 0 Then GoTo 20
    '
10  Call SELECHO("MASTER")
    COD1$ = TRIM$(VALACT1$("MASTER"))
    If COD1$ <> "" Then
      Ci% = CODINT(COD1$)
20    If Ci% > 0 Then
         Call FILTERFILE("TAMARCAS", "!TAMARCAS", 4, MKI$(Ci%))
         VTB% = VENTABU%("TAMARCAS")
         If VTB% >= 0 Then
            Call FILTERUPDATE("TAMARCAS", "!TAMARCAS", 4, MKI$(Ci%))
            If CODINT%(Label27) < 1 Then GoTo 10
         End If
      End If
    End If
    '
End Sub
'
Private Sub Command6_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = 113 Then ' CORRESPONDE TECLA <F2>
     Call Command6_Click(Index)
   End If
End Sub

Private Sub Command7_Click()
    '
    If UCase$(TRIM$(Label7)) = "AUTO" Then Exit Sub
    '
    Command7.Enabled = False
    '
    Call APERBASDAT("RESERVA")
    SQLX$ = " SELECT * FROM RESERVA "
    SQLX$ = SQLX$ + " WHERE CantRes - CantCons > 0 "
    SQLX$ = SQLX$ + " ORDER BY IdSubOr ASC, Idenlote ASC"
    '
    'Set RESERTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim RESERTMP As ADODB.Recordset
    Set RESERTMP = New ADODB.Recordset
    RESERTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

    With RESERTMP
      On Error Resume Next
      .MoveLast
      If Err Then
         Call MENSERR(24, " No Existen Reservas efectuadas. ")
         GoTo 99
      End If
      FIN& = .RecordCount
      .MoveFirst
      On Error GoTo 0
      JJ& = 0
      Do While (.EOF = False)
        k& = .AbsolutePosition
        Call TRACE(20, k&, FIN&)
        NUMORFA$ = !IDSUBOR
        If STATORFA%(NUMORFA$) < 3 Then
           '
           y$ = REGBLAN("RESERVA")
           Call REPLA(y$, MKI$(!Cod_Inte), 1, 2)
           Call REPLA(y$, !IDSUBOR, 3, 12)
           Call REPLA(y$, MKI$(!Cod_Inte), 15, 2)
           Call REPLA(y$, MKI$(CVINT(!FechRes)), 17, 2)
           Call REPLA(y$, MKD$(!CANTRES), 19, 8)
           Call REPLA(y$, MKD$(!CantCons), 27, 8)
           Call REPLA(y$, MKD$(!CANTRES - !CantCons), 35, 8)
           Call REPLA(y$, !IDENLOTE, 43, 16)
           Call REPLA(y$, Chr$(!Deposito), 60, 1)
           Call REPLA(y$, MKS$(!UsoUnit), 75, 4)
           '
           JJ& = JJ& + 1
           Call GRAREG("RESERVA", y$, JJ&)
           '
         End If
      .MoveNext
      Loop
    End With
    RESERTMP.Close
    Set RESERTMP = Nothing
    '
    Call SETULTREG("RESERVA", JJ&)
    Call FILESORT("RESERVA", "", 3, 4)
    '
    Call FINAL("*LIRESLOT")
    '
99  Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   '
   If UCase$(TRIM$(Label7)) = "AUTO" Then Exit Sub
   '
   Command8.Enabled = False
   Call TRAZAUTO07.Command15_Click
   Command8.Enabled = True
End Sub

Function CARPRODOF$(NROOF$)
  '
  CARPRODOF$ = ""
  Call ABREORFAB
  '
  SQLX$ = " SELECT * FROM OPERFAB "
  SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "' "
  'Set OPFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
  Dim OPFABTMP As ADODB.Recordset
  Set OPFABTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))

  '
  With OPFABTMP
    '
    On Error Resume Next
    .MoveFirst
    If Err Then
      On Error GoTo 0
      Exit Function
    End If
    On Error GoTo 0
    '
    Ci% = !Cod_Inte
    CARPRODOF$ = CODEXT$(Ci%)
  End With
  OPFABTMP.Close
  Set OPFABTMP = Nothing
  '
End Function

Private Sub Command6_Click(Index As Integer)
    If TRIM$(Label27) = "" Or CODINT%(Label27) < 0 Then
       Exit Sub
    End If
    '
    For KKU& = 0 To Index - 1
       If Label23(KKU&) = "" Then
          Call MENSERR(24, "No Dejar Campos no Informados")
          Exit Sub
       End If
    Next KKU&
    '
    CIXI% = CODINT%(Label27)
    Call ARMALILOT1(CIXI%, Label23(Index))
    Call FRESHECHO("!LILOTSE")
    '
    Call SELECHO("!LILOTSE")
    VDEVU$ = VALACT1$("!LILOTSE")
    VDEV2$ = VALACT2$("!LILOTSE")
    If VDEVU$ <> "" Then
      For KKU& = 0 To 5
        If KKU& <> Index Then
          If Label23(KKU&) = VDEVU$ Then
             Call MENSERR(24, "No Válido Repetir Lotes")
             Exit Sub
          End If
        End If
      Next KKU&
      Label23(Index) = VDEVU$
      '
      CAMAXI = XVALO(Text4)
      For KKU& = 0 To 5
         If KKU& <> Index Then
            CAMAXI = CAMAXI - XVALO(Text11(KKU&))
         End If
      Next KKU&
      If CAMAXI < 0 Then CAMAXI = 0
      '
      If XVALO(Mid$(VDEV2$, 21, 12)) < CAMAXI Then
         CAMAXI = XVALO(Mid$(VDEV2$, 33, 12))
      End If
      Text11(Index) = FORMATNUM$(CAMAXI, "F10.1")
    End If
    '
    Text11(Index).SetFocus
End Sub

Private Sub Command9_Click()
    Command9.Enabled = False
    If Check1.Value = 1 Then
        Call TrazabilidadAutoFifoConPackingList(True)
    Else
        BAUTOTRA% = 1
        Call AUTOTRA
        BAUTOTRA% = 0
    End If
    Command9.Enabled = True
End Sub

Private Sub Form_Load()
    '
    Call APLICACIONSKINS(Me, Picture14)
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    If INTERPRE% = 1 Then
       Frame9.Visible = True
    End If
    '
    Call FRESHALL
    '
    If Left$(APLINVO$, 7) = "OFABRIC" Then
       Label32 = TRIM$(Mid$(APLINVO$, 8))
       If Label32 <> "" Then
          If ECOARCH$("ECORDEP", Label32) <> "" Then
            HOII% = HOY(HOS$)
            Label24 = HOS$
          End If
       End If
    End If
    '
    Call CARGA_ENCABEZADO(mfgLotes, "Codigo/A16;Descripcion/A28;Lote/A10;Cantidad/F9.1;A Trazar/F9.1;Observaciones/A0", FRMZELECHO)
    '
    ' *** ACTIVA LA MODALIDAD DE FIFO CON LECTURA DE CODIGO DE BARRAS
    ModoTrazabilidad$ = TRIM$(Control("TRAZABI", "MODOTRAZ"))
    If ModoTrazabilidad$ = "FIFOCOBA" Then
        Check1.Value = 1
    End If
    
    ExigeCalidadCompleto = False
    If Control("TRAZABI", "EXICOCAL") = "SI" Then ExigeCalidadCompleto = True
    

End Sub

Private Sub GrillaLotesEliminar_Click()
            
   If FilaActiva = 1 And mfgLotes.Rows = 2 Then
       mfgLotes.Rows = 1
   Else
       mfgLotes.RemoveItem (FilaActiva)
   End If
   Command19.Visible = False
   
End Sub

Private Sub GrillaQuitarMarcaLectura_Click()

    If TRIM$(mfgLotes.TextMatrix(FilaActiva, 0)) = "*" Then
        mfgLotes.TextMatrix(FilaActiva, 0) = ""
        mfgLotes.TextMatrix(FilaActiva, 5) = ""
    ElseIf TRIM$(mfgLotes.TextMatrix(FilaActiva, 0)) <> "" Then
        Call COMUNI("Imposible Quitar Marca\a Lotes con Falta de Datos de Calidad")
    End If
    
End Sub


Private Sub InformeTecnico_Click()

    TituloLista$ = "Ordenes de Fabricación Cerradas (Parciales / Totales)"
    If TRIM$(FRMZELECHO.Caption) = TituloLista$ Then GoTo 50
    
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "O-Fabricación/A18;Referencia/A48", FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "O-Fabricación/A18;Referencia/A48", FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$
    
    Dim RST As ADODB.Recordset
    Set RST = READSET("SELECT IDSUBOR, REFEREN FROM ORDEFABR WITH(NOLOCK) WHERE CANAPRO>0 ORDER BY NUMEORFA DESC, IDSUBOR DESC")
    Call Carga_MSF_Recordset(RST, "O-Fabricación/A18;Referencia/A48", FRMZELECHO.msf2, 1)
    On Error Resume Next
    RST.Close
    Set RST = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 10000
50  FRMZELECHO.Show 1
    
    OrdenFabricacion$ = TRIM(RESP_ZELE_VECT(1))

    If OrdenFabricacion$ = "" Then Exit Sub
    
    InformeTecnicaOFEnWord (OrdenFabricacion$)
    
End Sub

Private Sub Label27_Change()
   CIXI% = CODINT%(Label27)
   If CIXI% > 0 Then
      Call ARMALILOT(CIXI%)
      Screen.MousePointer = 1
   End If
End Sub

Sub ARMALILOT1(CIXI%, LOTEAC$)
    '
    Static VERIRECLE%
    '
    If VERIRECLE% < 1 Then
       VERIRECLE% = 1
       If RECLEN%("LILOTSE") < 60 Then
          Call MENSERR(24, "Longitud de Registro Incorrecta\para Controlador 'LILOTSE'.\   \Reinstale el Controlador .RFS")
          Call FINAL("")
       End If
    End If
    '
    NORFA$ = Label32
    '
    Call APERBASDAT("STOCKS")
    SQLX$ = " SELECT * FROM LOTINGRE WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE Cod_Inte = " & CIXI% & " "
    SQLX$ = SQLX$ + " AND ((CanAlta - CanConsu) > 0 "    ' no trae los consumidos
    SQLX$ = SQLX$ + " OR CanConsu IS NULL or CanReser IS NULL OR CANCONSU = 0 OR CANRESER = 0)"
    SQLX$ = SQLX$ + " ORDER BY FechAlta ASC, IdenLote ASC "
    'Set LotIngTemp = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
'    Dim LOTINGTEMP As ADODB.Recordset
'    Set LOTINGTEMP = New ADODB.Recordset
'    LOTINGTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
     Set LOTINGTEMP = READSET(SQLX$)
    '
    REBLA$ = REGBLAN$("LILOTSE")
    JJ& = 0
    With LOTINGTEMP
         Do While Not .EOF
           IDLOT$ = TRIM$(SAFETEXT$(!IDENLOTE))
           '
           'Call ACSALOT(CIXI%, IDLOT$, CANTIN#, CACONRE#, CARESRE#, CSALDO#, NPROVX%)
           Call ACSALOTE(CIXI%, IDLOT$, "NOMESS")
           '
           TTXX$ = REBLA$
             FEX = CVINT(!FechAlta)
           Call REPLA(TTXX$, IDLOT$, 1, 16)
           Call REPLA(TTXX$, FECHATEX$(FEX), 17, 8)
           Call REPLA(TTXX$, FORMATNUM$(XVALO(!CanAlta), "F12,1"), 25, 12)
           '
           CANCON# = 0
           On Error Resume Next
           CANCON# = !CanConsu
           On Error GoTo 0
           CSALDO# = XVALO(!CanAlta) - CANCON#
           CSALFI# = CSALDO#
           SQLZ$ = " SELECT * FROM RESERVA WITH(NOLOCK) "
           SQLZ$ = SQLZ$ + " WHERE Cod_Inte = " & CIXI% & " "
           SQLZ$ = SQLZ$ + " AND IdenLote = '" & IDLOT$ & "' "
           SQLZ$ = SQLZ$ + " AND IdSubOr <> '" & NORFA$ & "' "
           SQLZ$ = SQLZ$ + " AND Status = 0 "     ' para no considerar los marcados para anular
           SQLZ$ = SQLZ$ + " ORDER BY FechRes ASC"
           'Set ReserTemp = Stocks.OpenRecordset(SQLZ$, dbOpenSnapshot)
'           Dim ReserTemp As ADODB.Recordset
'           Set ReserTemp = New ADODB.Recordset
'           ReserTemp.Open FILTSQL$(SQLZ$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
           Set ReserTemp = READSET(SQLZ$)
           With ReserTemp
              On Error Resume Next
              .MoveFirst
              If Err < 1 Then
                 On Error GoTo 0
                 Do While Not .EOF
                   CSALI# = XVALO(!CANTRES) - XVALO(!CantCons)
                   If CSALI# >= 0.05 Then
                     CSALDO# = CSALDO# - CSALI#
                   End If
                 .MoveNext
                  Loop
              End If
           End With
           ReserTemp.Close
           Set ReserTemp = Nothing
           '
           For KKI% = 0 To 5
              If Label23(KKI%) = IDLOT$ Then
                 If Label23(KKI%) <> LOTEAC$ Then
                   CSALDO# = CSALDO# - XVALO(Text11(KKI%))
                 End If
              End If
           Next KKI%
           '
           If CSALDO# >= 0.05 Or CSALFI# > 0 Then
              Call REPLA(TTXX$, FORMATNUM$(CSALFI#, "F12,1"), 37, 12)
              Call REPLA(TTXX$, FORMATNUM$(CSALDO#, "F12,1"), 49, 12)
              JJ& = JJ& + 1
              Call GRAREG("!LILOTSE", TTXX$, JJ&)
           End If
         .MoveNext
         Loop
    End With
    LOTINGTEMP.Close
    Set LOTINGTEMP = Nothing
    
    Call SETULTREG("!LILOTSE", JJ&)
    Call FRESHECHO("!LILOTSE")
    '
End Sub
'

Private Sub Label32_Change()
    ' CASO 1 - PONE EL CAMPO EN BLANCO
    NORFA$ = Label32
    Label17 = ""
    Label20 = ""
    Label21 = ""
    Label33 = ""
    HOII% = HOY(HOS$)
    Label24 = HOS$
    Text8.Enabled = True
    ' PONER EN BLANCO TODOS LOS RESTANTES CAMPOS Y LIST'
    List1.Clear
    Label28 = ""
    Label30 = ""
    Text8 = NORFA$
    Text1 = ""
    Text5 = ""
    Text6 = ""
    Text2 = ""
    Text3 = ""
    Label34 = ""
    Label27 = ""
    Label6 = ""
    Label16 = ""
    Text4 = ""
    '
    For KKU& = 0 To 5
       Label23(KKU&) = ""
       Text11(KKU&) = ""
       Text16(KKU&) = ""
       Text17(KKU&) = ""
    Next KKU&
    
    mfgLotes.Rows = 1
    Command19.Visible = False
    '
'    If UCase$(TRIM$(Label7)) = "AUTO" Then
'       Call CARLICONSU(1)
'       GoTo 99
'    End If
'    '
    If NORFA$ <> "" Then
         ' presenta datos de la orden de fabricación
         Call ABREORFAB
         SQLX$ = " SELECT * FROM ORDEFABR "
         SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "'"
         'Set ORDEFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
         Dim ORDEFABTMP As ADODB.Recordset
         Set ORDEFABTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))

         With ORDEFABTMP
           On Error Resume Next
           .MoveFirst
           If Err Then
              On Error GoTo 0
              Call MENSERR(24, "Orden de Fabricación no Válida. Consulte.")
              Call FINAL("")
           End If
           On Error GoTo 0
           CIXI% = !Cod_Inte
           Label28 = CODEXT$(CIXI%)
           Label30 = DESCRIT0$(CIXI%)
              CAPENDE# = !CanProy - !CanApro: If CAPENDE < 0 Then CAPENDE = 0
              If !Statuorf >= 3 Then CAPENDE# = 0
           Label34 = FORMATNUM$(!CanProy, "F8.1")
           Label37 = FORMATNUM$(!CanApro, "F8.1")
           Label38 = FORMATNUM$(CAPENDE#, "F8.1")
           '
         End With
         ORDEFABTMP.Close
         Set ORDEFABTMP = Nothing
         
         Call CARLICONSU(1)
         '
    End If
99  Screen.MousePointer = 1
End Sub

Sub CARLICONSU(MODO%)
    ' MODO%=1 DESDE LAS RESERVAS
    List1.Clear
    Dim LOTRESE$(2048), CANRESE(2048), CANTRAZA(2048)
    Label8 = (-1)
    CACONJU# = XVALO(Label38)
    If CACONJU# < 0.1 Then Exit Sub
    '
    AUTOTRAZA% = 0
    If UCase(TRIM$(Label7)) = "AUTO" Then
       If Control$("TRAZABI", "MODOTRAZ") = "AUTOMA" Then
          AUTOTRAZA% = 1
       End If
    End If
    '
    Call APERBASDAT("RESERVA")
    SQLX$ = " SELECT * FROM RESERVA "
    SQLX$ = SQLX$ + "WHERE IdSubOr = '" & Label32 & "' "
    SQLX$ = SQLX$ + "ORDER BY NuOrItem ASC "
    '
    On Error Resume Next    ' POR SI NO ESTÁ SETEADO NuOrItem
    'Set RESERTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim RESERTMP As ADODB.Recordset
    Set RESERTMP = New ADODB.Recordset
    RESERTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

    If Err Then
       On Error GoTo 0 'SEGUI EL CRITERIO PERO ESTA REPETIDO ARRIBA (OMAR)
       SQLX$ = " SELECT * FROM RESERVA "
       SQLX$ = SQLX$ + "WHERE IdSubOr = '" & Label32 & "' "
'       Set RESERTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
       Set RESERTMP = New ADODB.Recordset
       RESERTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

    End If
    On Error GoTo 0
    '
    CAIT% = 0
    With RESERTMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
          On Error GoTo 0
          Do While (.EOF = False)
            CIXI% = !Cod_Inte
            TTZZ$ = Space$(64)
            IDLOCO$ = TRIM$(!IDENLOTE)
            CACOLO = !CANTRES - !CantCons: If CACOLO < 0 Then CACOLO = 0
            CANTRA = 0
            If IDLOCO$ <> "" Then CANTRA = CACOLO
            Call REPLA(TTZZ$, IDLOCO$, 1, 16)
            Call REPLA(TTZZ$, FORMATNUM$(CACOLO, "F12.1"), 17, 16)
            On Error Resume Next
            Call REPLA(TTZZ$, !PNL_DES, 33, 16)
            Call REPLA(TTZZ$, !PLC_HAS, 49, 16)
            On Error GoTo 0
            '
            For JJ& = 1 To CAIT%
               If CIJ%(JJ&) = CIXI% Then
                  CANRESE(JJ&) = CANRESE(JJ&) + CACOLO
                  CANTRAZA(JJ&) = CANTRAZA(JJ&) + CANTRA
                  If IDLOCO$ <> "" Then
                     LOTRESE$(JJ&) = LOTRESE$(JJ&) + TTZZ$
                  End If
                  GoTo 19
               End If
            Next JJ&
            CAIT% = CAIT% + 1
            CIJ%(CAIT%) = !Cod_Inte
            USOI(CAIT%) = !UsoUnit
            CANRESE(CAIT%) = CACOLO
            CANTRAZA(CAIT%) = CANTRA
            LOTRESE$(CAIT%) = TTZZ$
19        .MoveNext
          Loop
       End If
       On Error GoTo 0
    End With
    RESERTMP.Close
    Set RESERTMP = Nothing
    '
    If CAIT% < 1 Then    ' No Existen Reservas efectuadas.
       ' arma composicion segun formula standard
       CODPRO$ = Label28
       Call LEEEXPLOMUL(CODPRO$, 1, 1)
       For JJ& = 1 To CAIT%
          LOTRESE$(JJ&) = ""
          CANRESE(JJ&) = CACONJU# * USOI(JJ&)
          CANTRAZA(JJ&) = 0
       Next JJ&
    End If
    '
    If CAIT% < 1 Then             ' producto sin estructura
      Call MENSERR(24, "No hay Fórmula de Consumo Standard\para este Producto (" + TRIM$(CODPRO$) + ").")
    End If
    '
    FIN& = CAIT%
    For I1% = 1 To CAIT%
       UI1& = I1%
       Call TRACE(20, UI1&, FIN&)
       If CIJ%(I1%) > 0 Then
          zz$ = Space$(128)
          Call REPLA(zz$, CODEXT$(CIJ%(I1%)), 1, 16)
          Call REPLA(zz$, DESCRIT$(CIJ%(I1%)), 17, 40)
            NEBRU = CANRESE(I1%)
          Call REPLA(zz$, FORMATNUM$(NEBRU, "F13.1"), 57, 13)
            SUMATRA = CANTRAZA(I1%)
          Call REPLA(zz$, FORMATNUM$(SUMATRA, "F12.1"), 71, 12)
          Call REPLA(zz$, FORMATNUM$(USOI(I1%), "F16.6"), 97, 16)
          '
          zz$ = zz$ + LOTRESE$(I1%)
          List1.AddItem zz$
          '
          ' TRAZADO AUTOMATICO POR F.I.F.O. PARA DETERPLUS
          If AUTOTRAZA% > 0 Then
             If SUMATRA < NEBRU Then
                LINDEX = List1.ListCount
                List1.ListIndex = LINDEX - 1
                Call List1_Click
                Call AUTOTRA
             End If   ' FIN DE AUTOTRAZA
          End If
          '
       End If
    Next I1%
    Screen.MousePointer = 1
    '
End Sub

Sub AUTOTRA()
    '
    NEBRU = XVALO(Text4)
    SUMATRA = XVALO(Label16)
    CIIX% = CODINT%(Label27)
    '
    If CIIX% < 1 Then Exit Sub
    If List1.ListIndex < 0 Then Exit Sub
    If SUMATRA >= NEBRU Then Exit Sub
    '
    Call ARMALILOT2(CIIX%)
    For KKUU& = 1 To ULTREG&("!LILOTSE")
       XX3$ = REGLEIDO$("!LILOTSE", KKUU&)
       VDEVU$ = TRIM$(Left$(XX3$, 16))
       VDEV2$ = Mid$(XX3$, 17)
       '
       If VDEVU$ <> "" Then
          CAMAXI = XVALO(Text4)
          For KKU& = 0 To 5
             CAMAXI = CAMAXI - XVALO(Text11(KKU&))
          Next KKU&
          If CAMAXI < 0 Then CAMAXI = 0
          '
          If XVALO(Mid$(VDEV2$, 33, 12)) < CAMAXI Then  ' CONSIDERA EL SALDO LOGICO (CANTIDAD LOTE - CONSUMO - RESERVA)
             CAMAXI = XVALO(Mid$(VDEV2$, 33, 12))
          End If
          '
          If CAMAXI >= 0.05 Then
            For KKU& = 0 To 5
              If TRIM$(Label23(KKU&)) = VDEVU$ Then
                 CANLO = XVALO(Text11(KKU&)) + CAMAXI
                 CAMAXI = 0
                 Text11(KKU&) = FORMATNUM$(CANLO, "F10.1")
                 SUMATRA = SUMATRA + CAMAXI
                 GoTo 29
              End If
            Next KKU&
            '
            ' BUSCA EL PRIMERO VACÍO
            For KKU& = 0 To 5
              If TRIM$(Label23(KKU&)) = "" Then
                 Label23(KKU&) = VDEVU$
                 Text11(KKU&) = FORMATNUM$(CAMAXI, "F10.1")
                 SUMATRA = SUMATRA + CAMAXI
                 GoTo 29
              End If
            Next KKU&
          End If
       End If
29     If SUMATRA >= NEBRU Then GoTo 39
    Next KKUU&
    '
39  Call Command2_Click    ' PARA SUBIRLO
    '
End Sub
'
Private Sub Label32_DblClick()
    ' USAR PARA LIMPIAR TODO
    Label32 = ""
End Sub

'
Private Sub List1_Click()
    
    List1.Enabled = False
    
    TTXX$ = List1.TEXT
    COD$ = Mid(TTXX$, 1, 16)
    Ci% = CODINT%(COD$)
    DESC$ = DESCRIT$(Ci%)
    CantNom$ = Mid(TTXX$, 57, 13)
    Label27.Caption = COD$
    Label6.Caption = DESC$
    Text4 = TRIM$(CantNom$)
    Label8 = List1.ListIndex
    '
    For KKU& = 0 To 5
       Label23(KKU&) = ""
       Text11(KKU&) = ""
       Text16(KKU&) = ""
       Text17(KKU&) = ""
    Next KKU&
    '
    KKU& = 0
    For KKL& = 129 To Len(TTXX$) Step 64
       TTYY$ = Mid$(TTXX$, KKL&, 64)
       If KKU& < 6 Then
         Label23(KKU&) = RTrim(Left$(TTYY$, 16))
         Text11(KKU&) = TRIM$(Mid$(TTYY$, 17, 16))
         Text16(KKU&) = TRIM$(Mid$(TTYY$, 33, 16))
         Text17(KKU&) = TRIM$(Mid$(TTYY$, 49, 16))
       End If
       KKU& = KKU& + 1
    Next KKL&
    '
    List1.Enabled = True
    '
End Sub

Sub LIMPLAB()

    Label27.Caption = ""
    Label6.Caption = ""
    Text4.TEXT = ""
  
End Sub

Sub LIMPTEXT()

    Text5 = ""
    Text3 = ""
    Text6 = ""
    Text8 = ""
    Text9 = ""
    Text14 = ""
    Text12 = ""
  
End Sub



Private Sub mfgLotes_Click()

    With mfgLotes
        If mfgLotes.Rows > 1 Then
            .COL = 1
            Command19.Move .Left + .CellLeft + (.CellWidth - Command19.Width), .Top + .CellTop, 255, 255
            Command19.Visible = True
            Command19.ZOrder 0
            FilaActiva = mfgLotes.Row
            Command19.Visible = True
            
            Text2.TEXT = mfgLotes.TextMatrix(mfgLotes.Row, 6)
        End If
    End With
    
End Sub



Private Sub Text1_KeyPress(KeyAscii As Integer)

    Text2 = Text2 + CStr(KeyAscii)
    If KeyAscii <> 13 And KeyAscii <> 124 Then Exit Sub
    If KeyAscii = 124 Then KeyAscii = 0
    If TRIM$(Text1) = "" Then Exit Sub
    
    LoteRecepcion$ = Text1
    'sLecturaCodigoBarras$ = Text1
    'sLecturaCodigoBarras$ = REPLACAR(sLecturaCodigoBarras$, "-", "/")
    'sLecturaCodigoBarras$ = REPLACAR(sLecturaCodigoBarras$, "'", "-")
    'sLecturaCodigoBarras$ = REPLACAR(sLecturaCodigoBarras$, "ç", "|")
    
'    Call TEXTOLOTES(sLecturaCodigoBarras$, "|")
'    On Error Resume Next
'    For i& = 1 To UBound(CADENATEX$) 'RECORRO EL VECTOR QUE TIENE CUALES SON LAS CELDAS EDITABLES
'        Select Case i&
'        Case 1
'         LoteRecepcion$ = UCase(CADENATEX$(i&))
'        Case 2
'         codigoExterno$ = TRIM$(CADENATEX$(i&))
'        Case 3
'         cantidadAprobada# = XVALO(CADENATEX$(i&))
'        End Select
'    Next i&
'    On Error GoTo 0
    '
    EncontroLote% = 0
    For indGrilla% = 1 To mfgLotes.Rows - 1
        loteCargado$ = TRIM$(mfgLotes.TextMatrix(indGrilla%, 3))
        ' ENCONTRO EL LOTE LEIDO EN LA GRILLA DE LOTES AUTO FIFO
        If LoteRecepcion$ = loteCargado$ Then
            If TRIM$(mfgLotes.TextMatrix(indGrilla%, 0)) <> "(E)" Then
                mfgLotes.TextMatrix(indGrilla%, 0) = "*"
                mfgLotes.TextMatrix(indGrilla%, 5) = mfgLotes.TextMatrix(indGrilla%, 4)
                EncontroLote% = 1
            End If
        End If
    Next indGrilla%
    
    If EncontroLote% = 0 Then
        CodigoProducto$ = VALOBADA("BOLRECEP", "COD_EXTE", "IDENLOTE='" & LoteRecepcion$ & "'")
        DescripcionProducto$ = VALOBADA("BOLRECEP", "DESCRIPITEM", "IDENLOTE='" & LoteRecepcion$ & "'")
        
        mfgLotes.Rows = mfgLotes.Rows + 1
        mfgLotes.TextMatrix(mfgLotes.Rows - 1, 0) = "¿?"
        mfgLotes.TextMatrix(mfgLotes.Rows - 1, 1) = CodigoProducto$
        mfgLotes.TextMatrix(mfgLotes.Rows - 1, 2) = DescripcionProducto$
        mfgLotes.TextMatrix(mfgLotes.Rows - 1, 3) = LoteRecepcion$
    End If
    
99  Text1 = "": Text2 = ""

End Sub

Private Sub Text11_Change(Index As Integer)
    SUMATRA = 0
    For KKI% = 0 To 5
       If Label23(KKI%) <> "" Then
           SUMATRA = SUMATRA + XVALO(Text11(KKI%))
         Else
           Text11(KKI%) = ""
       End If
    Next KKI%
    Label16 = TRIM$(FORMATNUM$(SUMATRA, "F10.1"))
    
    CIXI% = CODINT%(Label27)
    If CIXI% > 0 Then
       LTXX$ = TRIM$(Label23(Index))
       If LTXX$ <> "" Then
          Call ACSALOTE(CIXI%, LTXX$, "NOMESS")
          If BAUTOTRA% < 1 And UCase$(TRIM$(Label7)) <> "AUTO" Then
             Call ARMALILOT1(CIXI%, Label23(Index))
          End If
          Screen.MousePointer = 1
       End If
    End If
End Sub

Private Sub Text11_DblClick(Index As Integer)
    Text11(Index) = ""
End Sub

Private Sub Text16_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
       Text17(Index).SetFocus
    End If
End Sub

Private Sub Text17_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Index < 5 Then
          Command6(Index + 1).SetFocus
       End If
    End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)

    If KeyAscii% = 13 Then
       Call Text4_LostFocus
    End If
    
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Text6.SetFocus
   End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       Text2.SetFocus
    End If

End Sub

Private Sub Text11_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call Text11_LostFocus(Index)
       Text16(Index).SetFocus
    End If
End Sub

Private Sub Text11_LostFocus(Index As Integer)
    Text11(Index) = TRIM$(FORMATNUM$(Text11(Index), "F10.1"))
End Sub

Private Sub Text4_LostFocus()
   '
   Text4 = TRIM$(FORMATNUM$(Text4, "F12.1"))
   '
   If TRIM$(Label27) <> "" Then
      If XVALO(Text4) < 0.05 Then
         For JJ& = 0 To List1.ListCount - 1
            TTYY$ = List1.LIST(JJ&)
            If TRIM$(Left$(TTYY$, 16)) = TRIM$(Label27) Then
               List1.RemoveItem JJ&
               GoTo 80
            End If
         Next JJ&
         '
80       Label27 = ""
         Label6 = ""
         Text4 = ""
         Label8 = (-1)
      End If
   End If
   '
   List1.Refresh
   '
End Sub

Private Sub Text8_GotFocus()
    If UCase$(TRIM$(Label7)) = "AUTO" Then Command13.SetFocus
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Text3.SetFocus
   End If
End Sub
