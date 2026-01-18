VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form INFOCORTE 
   BackColor       =   &H00E2D3C0&
   Caption         =   "Informe de Operaciones "
   ClientHeight    =   8790
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   15015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   15015
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command15 
      Caption         =   "!"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   7.5
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   143
      Top             =   1200
      Width           =   300
   End
   Begin MSFlexGridLib.MSFlexGrid MSF3 
      Height          =   795
      Left            =   120
      TabIndex        =   142
      Top             =   1440
      Visible         =   0   'False
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   1402
      _Version        =   393216
      Rows            =   26
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Informar Fusión de Operaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   122
      Top             =   6720
      Visible         =   0   'False
      Width           =   14835
      Begin VB.CommandButton CmdCerrarFucion 
         Appearance      =   0  'Flat
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
         Height          =   300
         Left            =   14400
         TabIndex        =   136
         Top             =   120
         Width           =   315
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         Caption         =   "."
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
         Left            =   10590
         TabIndex        =   133
         Top             =   840
         Width           =   200
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   9960
         TabIndex        =   132
         Top             =   840
         Width           =   600
      End
      Begin VB.CommandButton CmdInfoFusion 
         Caption         =   "Seleccione Operación de Fusión"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   126
         ToolTipText     =   "Seleccione Operación de Fusión"
         Top             =   360
         Width           =   3045
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Registrar Peso"
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
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   125
         ToolTipText     =   "Lee el Peso de La Balanza"
         Top             =   360
         Width           =   1485
      End
      Begin VB.TextBox Text16 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   12240
         TabIndex        =   124
         Text            =   " "
         Top             =   360
         Width           =   936
      End
      Begin VB.CommandButton Command13 
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
         Height          =   315
         Left            =   13320
         Style           =   1  'Graphical
         TabIndex        =   123
         ToolTipText     =   "O.K."
         Top             =   840
         Width           =   1365
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Operarios:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   9120
         TabIndex        =   135
         Top             =   840
         Width           =   765
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   5
         Left            =   10800
         TabIndex        =   134
         ToolTipText     =   "Haga Doble-Click para Informe por Célula"
         Top             =   840
         Width           =   2385
      End
      Begin VB.Label lblCodFusion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3240
         TabIndex        =   131
         Top             =   360
         Width           =   810
      End
      Begin VB.Label lblDescFusion 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   4080
         TabIndex        =   130
         Top             =   360
         Width           =   3570
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   9960
         TabIndex        =   129
         ToolTipText     =   "Kg. Con Tara Descontada"
         Top             =   360
         Width           =   1290
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Kilos:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   15
         Left            =   9480
         TabIndex        =   128
         Top             =   435
         Width           =   615
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   11280
         TabIndex        =   127
         Top             =   435
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Informes de Sub-Operaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Index           =   1
      Left            =   120
      TabIndex        =   76
      Top             =   6720
      Width           =   14835
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   11760
         TabIndex        =   113
         Top             =   600
         Width           =   600
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         Caption         =   "."
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
         Left            =   12390
         TabIndex        =   112
         Top             =   600
         Width           =   200
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Registrar Peso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   13200
         Style           =   1  'Graphical
         TabIndex        =   104
         ToolTipText     =   "Lee el Peso de La Balanza"
         Top             =   1080
         Width           =   1485
      End
      Begin VB.TextBox TxtCantimpresa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   11760
         Locked          =   -1  'True
         TabIndex        =   98
         Top             =   1560
         Width           =   1275
      End
      Begin VB.CommandButton Command10 
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
         Height          =   315
         Left            =   13560
         Style           =   1  'Graphical
         TabIndex        =   97
         ToolTipText     =   "O.K."
         Top             =   1560
         Width           =   1125
      End
      Begin VB.TextBox TXTLECTURA 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   3720
         TabIndex        =   78
         Top             =   360
         Width           =   7035
      End
      Begin VB.Label lblKgRegSelec 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   12960
         TabIndex        =   140
         Top             =   120
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "KilosRegSelec"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   17
         Left            =   11760
         TabIndex        =   139
         Top             =   120
         Visible         =   0   'False
         Width           =   1650
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "CantDelRegistroSelec"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   16
         Left            =   8880
         TabIndex        =   138
         Top             =   120
         Visible         =   0   'False
         Width           =   1650
      End
      Begin VB.Label lblCantidadRegSelec 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   10560
         TabIndex        =   137
         Top             =   120
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Cortes C/Proceso Informados: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   14
         Left            =   7440
         TabIndex        =   121
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   8520
         TabIndex        =   120
         Top             =   1560
         Width           =   1050
      End
      Begin VB.Label LblId 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   7560
         TabIndex        =   119
         Top             =   120
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "IdRegistro:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   13
         Left            =   6720
         TabIndex        =   118
         Top             =   120
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   6
         Left            =   4920
         TabIndex        =   117
         Top             =   120
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "O:Fabricación.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   6
         Left            =   3720
         TabIndex        =   116
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   4
         Left            =   12600
         TabIndex        =   115
         ToolTipText     =   "Haga Doble-Click para Informe por Célula"
         Top             =   600
         Width           =   2100
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Operarios:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   10920
         TabIndex        =   114
         Top             =   600
         Width           =   765
      End
      Begin VB.Label KgImpresos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   11760
         TabIndex        =   103
         ToolTipText     =   "Kg. Con Tara Descontada"
         Top             =   1080
         Width           =   1290
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Kilos Informados:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   12
         Left            =   10440
         TabIndex        =   102
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label LblKgSinImpresion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   8520
         TabIndex        =   101
         ToolTipText     =   "Kg. Con Tara Descontada"
         Top             =   840
         Width           =   1050
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Kilos Sin Sub-Proceso:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   11
         Left            =   7440
         TabIndex        =   100
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad de         Cortes (Sub-Proceso)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   10
         Left            =   10440
         TabIndex        =   99
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "ID Ubicacion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   9
         Left            =   4680
         TabIndex        =   96
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Ubicación:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   8
         Left            =   4680
         TabIndex        =   95
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "N° Orden Inf.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   7
         Left            =   4680
         TabIndex        =   94
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "N° Pedido:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   2520
         TabIndex        =   93
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "N° Cliente:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   2520
         TabIndex        =   92
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "N°Operación:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   2520
         TabIndex        =   91
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Cant.Cortes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   240
         TabIndex        =   90
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Lote:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   240
         TabIndex        =   89
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   240
         TabIndex        =   88
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   9
         Left            =   5880
         TabIndex        =   87
         Top             =   840
         Visible         =   0   'False
         Width           =   1400
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   8
         Left            =   5880
         TabIndex        =   86
         Top             =   1560
         Width           =   1400
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   7
         Left            =   5880
         TabIndex        =   85
         Top             =   1200
         Width           =   1400
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   5
         Left            =   3600
         TabIndex        =   84
         Top             =   1560
         Width           =   1005
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   4
         Left            =   3600
         TabIndex        =   83
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   3
         Left            =   3600
         TabIndex        =   82
         Top             =   840
         Width           =   1005
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   1440
         TabIndex        =   81
         Top             =   1560
         Width           =   930
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   840
         TabIndex        =   80
         Top             =   1200
         Width           =   1530
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lectura de Código QR  >>>>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   79
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   0
         Left            =   840
         TabIndex        =   77
         Top             =   840
         Width           =   1530
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Informes de Retiros de Corte de Orden de Fabricación Activa  (No se Visualizarán Los Informes que ya hayan pasado a la Linea)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   45
      Top             =   4200
      Width           =   14835
      Begin VB.TextBox Text222 
         Alignment       =   2  'Center
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
         Height          =   1890
         Index           =   1
         Left            =   5880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   75
         TabStop         =   0   'False
         Text            =   "INFOCORTE.frx":0000
         Top             =   360
         Visible         =   0   'False
         Width           =   5055
      End
      Begin MSFlexGridLib.MSFlexGrid MSF2 
         Height          =   2130
         Left            =   120
         TabIndex        =   46
         ToolTipText     =   "Doble Clic Para Guardar o Editar Ubicación - Las Celdas Pintadas Corresponden a Sub-Procesos"
         Top             =   240
         Width           =   14640
         _ExtentX        =   25823
         _ExtentY        =   3757
         _Version        =   393216
         Rows            =   26
         Cols            =   4
         BackColor       =   16777215
         BackColorFixed  =   16501405
         BackColorBkg    =   16777215
         FillStyle       =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   9120
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   40
      Top             =   3240
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Parte de O/F a Informar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   0
      Left            =   120
      TabIndex        =   34
      Top             =   0
      Width           =   14835
      Begin VB.CommandButton Command14 
         Caption         =   "!"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   7.5
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3000
         TabIndex        =   141
         Top             =   360
         Width           =   175
      End
      Begin VB.TextBox Text14 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   864
         Locked          =   -1  'True
         TabIndex        =   65
         Text            =   " "
         Top             =   720
         Width           =   1920
      End
      Begin VB.TextBox Text13 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   12870
         Locked          =   -1  'True
         TabIndex        =   63
         Text            =   " "
         Top             =   600
         Width           =   720
      End
      Begin VB.TextBox Text12 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10680
         Locked          =   -1  'True
         TabIndex        =   57
         Text            =   " "
         Top             =   600
         Width           =   720
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10680
         Locked          =   -1  'True
         TabIndex        =   56
         Text            =   " "
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton Command8 
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
         Left            =   14090
         Picture         =   "INFOCORTE.frx":007D
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Salir"
         Top             =   660
         Width           =   600
      End
      Begin VB.CommandButton Command3 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   7.5
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2760
         TabIndex        =   2
         Top             =   360
         Width           =   175
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   864
         Locked          =   -1  'True
         TabIndex        =   0
         ToolTipText     =   "Ingresar Código de Producto Terminado"
         Top             =   960
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.CommandButton Command1 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   7.5
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2760
         TabIndex        =   1
         Top             =   960
         Visible         =   0   'False
         Width           =   175
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cant.Proy.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -45
         TabIndex        =   66
         Top             =   720
         Width           =   885
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cant.del Pedido:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   11535
         TabIndex        =   64
         Top             =   600
         Width           =   1245
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   11520
         TabIndex        =   60
         Top             =   240
         Width           =   3210
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pedido N°"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9720
         TabIndex        =   59
         Top             =   600
         Width           =   885
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   9480
         TabIndex        =   58
         Top             =   300
         Width           =   1125
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3000
         TabIndex        =   39
         Top             =   960
         Visible         =   0   'False
         Width           =   6810
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   870
         TabIndex        =   38
         ToolTipText     =   "Seleccione Orden de Fabricación Asociada"
         Top             =   360
         Width           =   1890
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -510
         TabIndex        =   37
         Top             =   1035
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de O/F: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -480
         TabIndex        =   36
         Top             =   435
         Width           =   1335
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3240
         TabIndex        =   35
         Top             =   360
         Width           =   6570
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Informar Operación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2724
      Left            =   120
      TabIndex        =   19
      Top             =   1560
      Width           =   14835
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   7680
         TabIndex        =   109
         Top             =   600
         Width           =   600
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         Caption         =   "."
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
         Left            =   8310
         TabIndex        =   108
         Top             =   600
         Width           =   200
      End
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   106
         Text            =   " "
         ToolTipText     =   "Es la medida del corte en mts."
         Top             =   3120
         Width           =   1260
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Registrar Peso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   11640
         Style           =   1  'Graphical
         TabIndex        =   105
         ToolTipText     =   "Lee el Peso de La Balanza"
         Top             =   2280
         Width           =   1485
      End
      Begin VB.TextBox TXTRESTAN 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   13245
         Locked          =   -1  'True
         TabIndex        =   73
         Text            =   " "
         Top             =   600
         Width           =   1260
      End
      Begin VB.TextBox TXTNECESIDADCORTES 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   13245
         Locked          =   -1  'True
         TabIndex        =   71
         Text            =   " "
         Top             =   240
         Width           =   1260
      End
      Begin VB.TextBox TXTCORTESOPERACION 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   13245
         Locked          =   -1  'True
         TabIndex        =   69
         Text            =   " "
         Top             =   1680
         Width           =   1260
      End
      Begin VB.TextBox TXTCORTESOPYLOTE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   13245
         Locked          =   -1  'True
         TabIndex        =   67
         Text            =   " "
         Top             =   1200
         Width           =   1260
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2295
         Left            =   14640
         ScaleHeight     =   2265
         ScaleWidth      =   2385
         TabIndex        =   61
         Top             =   360
         Visible         =   0   'False
         Width           =   2415
         Begin VB.Image Image1 
            Appearance      =   0  'Flat
            Height          =   2280
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Visible         =   0   'False
            Width           =   2400
         End
      End
      Begin VB.TextBox TXTCANTICORTES 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   53
         Text            =   " "
         ToolTipText     =   "La cantidad de cortes que tiene por unidad "
         Top             =   1320
         Width           =   696
      End
      Begin VB.TextBox TXTCORTE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   52
         Text            =   " "
         ToolTipText     =   "Es la medida del corte en mts."
         Top             =   1320
         Width           =   1260
      End
      Begin VB.TextBox TXTCODELEM 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   50
         Text            =   " "
         Top             =   960
         Width           =   1656
      End
      Begin VB.TextBox TXTLOTE 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9840
         Locked          =   -1  'True
         TabIndex        =   48
         Text            =   " "
         Top             =   960
         Width           =   1170
      End
      Begin VB.CommandButton Command2 
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
         Height          =   315
         Left            =   9885
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "O.K."
         Top             =   2280
         Width           =   1125
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   240
         Width           =   600
      End
      Begin VB.CommandButton Command6 
         Appearance      =   0  'Flat
         Caption         =   "."
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
         Left            =   8304
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.CommandButton Command7 
         Caption         =   "."
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
         Left            =   1935
         TabIndex        =   5
         Top             =   4080
         Width           =   175
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
         Height          =   285
         Left            =   990
         TabIndex        =   4
         Top             =   4080
         Width           =   960
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
         Left            =   1560
         TabIndex        =   7
         Top             =   4560
         Width           =   175
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
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   6
         Top             =   4560
         Width           =   480
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
         Left            =   1560
         TabIndex        =   9
         Top             =   5040
         Width           =   175
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
         Height          =   285
         Index           =   2
         Left            =   1080
         TabIndex        =   8
         Top             =   5040
         Width           =   480
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
         Left            =   1560
         TabIndex        =   11
         Top             =   6840
         Width           =   175
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
         Height          =   285
         Index           =   3
         Left            =   1080
         TabIndex        =   10
         Top             =   6840
         Width           =   480
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7680
         TabIndex        =   12
         Text            =   " "
         Top             =   1800
         Width           =   940
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         Caption         =   "."
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
         Left            =   8664
         TabIndex        =   13
         Top             =   1800
         Width           =   200
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10290
         TabIndex        =   14
         Text            =   " "
         Top             =   1800
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7680
         TabIndex        =   15
         Text            =   " "
         Top             =   2280
         Width           =   936
      End
      Begin VB.TextBox Text8 
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
         Height          =   285
         Left            =   4800
         TabIndex        =   18
         Text            =   " "
         Top             =   3168
         Width           =   480
      End
      Begin VB.TextBox Text9 
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
         Height          =   285
         Left            =   3720
         TabIndex        =   17
         Text            =   " "
         Top             =   3168
         Width           =   480
      End
      Begin VB.TextBox Text11 
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
         Height          =   285
         Left            =   2760
         TabIndex        =   16
         Text            =   " "
         Top             =   3168
         Width           =   480
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   2364
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   6600
         _ExtentX        =   11642
         _ExtentY        =   4180
         _Version        =   393216
         Rows            =   26
         Cols            =   4
         BackColor       =   16777215
         BackColorFixed  =   16501405
         BackColorBkg    =   16777215
         FillStyle       =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tolerancia 5%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   150
         Left            =   8560
         TabIndex        =   144
         Top             =   2350
         Width           =   1215
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   0
         Left            =   8520
         TabIndex        =   111
         ToolTipText     =   "Haga Doble-Click para Informe por Célula"
         Top             =   600
         Width           =   2490
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Operario:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   6720
         TabIndex        =   110
         Top             =   600
         Width           =   885
      End
      Begin VB.Label LblKgImpresos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   13245
         TabIndex        =   107
         ToolTipText     =   "Kg. Con Tara Descontada"
         Top             =   2280
         Width           =   1260
      End
      Begin VB.Label LABEL12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Restan:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   12120
         TabIndex        =   74
         Top             =   600
         Width           =   885
      End
      Begin VB.Label LABEL12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Necesidad de Cortes Para la Operación Activa:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   5
         Left            =   11190
         TabIndex        =   72
         Top             =   240
         Width           =   1965
      End
      Begin VB.Label LABEL12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cortes Informados de Operación Activa:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   6
         Left            =   11400
         TabIndex        =   70
         Top             =   1620
         Width           =   1725
      End
      Begin VB.Label LABEL12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cortes Informados de operación y lote activo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   4
         Left            =   11400
         TabIndex        =   68
         Top             =   1155
         Width           =   1725
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Imagen QR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   14520
         TabIndex        =   62
         Top             =   120
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label LABEL12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cant. Cortes:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   9120
         TabIndex        =   55
         Top             =   1395
         Width           =   1215
      End
      Begin VB.Label LABEL12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Corte:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   6790
         TabIndex        =   54
         Top             =   1392
         Width           =   768
      End
      Begin VB.Label LABEL12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   6840
         TabIndex        =   51
         Top             =   1032
         Width           =   768
      End
      Begin VB.Label LABEL12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lote:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   9240
         TabIndex        =   49
         Top             =   1032
         Width           =   528
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   8520
         TabIndex        =   33
         Top             =   240
         Width           =   2490
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Operación:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   228
         Left            =   6480
         TabIndex        =   32
         Top             =   300
         Width           =   1128
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Equipo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   228
         Left            =   0
         TabIndex        =   31
         Top             =   4152
         Width           =   852
      End
      Begin VB.Label Label10 
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
         Height          =   288
         Left            =   2256
         TabIndex        =   30
         Top             =   4080
         Width           =   3000
      End
      Begin VB.Label Label5 
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
         Height          =   312
         Index           =   1
         Left            =   1920
         TabIndex        =   29
         ToolTipText     =   "Haga Doble-Click para Informe por Célula"
         Top             =   4560
         Width           =   3348
      End
      Begin VB.Label Label5 
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
         Height          =   312
         Index           =   2
         Left            =   1920
         TabIndex        =   28
         ToolTipText     =   "Haga Doble-Click para Informe por Célula"
         Top             =   5040
         Width           =   3348
      End
      Begin VB.Label Label5 
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
         Height          =   312
         Index           =   3
         Left            =   1920
         TabIndex        =   27
         ToolTipText     =   "Haga Doble-Click para Informe por Célula"
         Top             =   6840
         Width           =   3348
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6480
         TabIndex        =   26
         Top             =   1872
         Width           =   1128
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hora:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9480
         TabIndex        =   25
         Top             =   1800
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad de Cortes:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6720
         TabIndex        =   24
         Top             =   2240
         Width           =   852
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Seg"
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
         Left            =   3960
         TabIndex        =   23
         Top             =   3192
         Width           =   708
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Min"
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
         Left            =   3000
         TabIndex        =   22
         Top             =   3192
         Width           =   708
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Horas"
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
         Left            =   2040
         TabIndex        =   21
         Top             =   3192
         Width           =   708
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "INFOCORTE.frx":01C7
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "INFOCORTE.frx":03FB
      TabIndex        =   41
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   47
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
   End
   Begin VB.Menu mnuCtaListado 
      Caption         =   "Consultas y Listados"
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Ver "
   End
   Begin VB.Menu mnuUtilidades 
      Caption         =   "Utilidades"
      Begin VB.Menu mnuOrigScrap 
         Caption         =   "Informar Origen de Scrap"
      End
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "Configuración"
      Begin VB.Menu mnuTabAux 
         Caption         =   "Tablas Auxiliares"
      End
      Begin VB.Menu mnuTablaConfScrap 
         Caption         =   "Configuración Tabla de Scrap"
      End
      Begin VB.Menu mnuDepoCinESl 
         Caption         =   "Configuración Depósitos (Producción)"
      End
   End
   Begin VB.Menu mnuAccesDirect 
      Caption         =   "Accesos Directos"
      Begin VB.Menu mnuInfOpEslinlgas 
         Caption         =   "Informe de Operaciones de Eslingas"
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIngCintillas 
         Caption         =   "Informe de Ingreso de Materiales a Depósito de Cintillas"
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUbic 
         Caption         =   "Ubicaciones"
      End
   End
End
Attribute VB_Name = "INFOCORTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub BLANFORMU_FUCION()
   lblCodFusion = ""
    lblDescFusion = ""
    Label26 = ""
    Text16 = ""
    Text1(5) = ""
    Label5(5) = ""
End Sub

Sub BLANFORMUIMPRESION()
    TXTLECTURA = ""
    For i% = 0 To Label31.UBound
       Label31(i%) = ""
    Next i%
    LblId = ""
    lblKgRegSelec = ""
    lblCantidadRegSelec = ""
    Text1(4) = ""
    Label5(4) = ""
    LblKgSinImpresion = ""
    KgImpresos = ""
    TxtCantimpresa = ""
    Label25 = ""
    
End Sub

Sub GenerarConsumoEslingasCintillas(NORFA$, Codigo$, LOTE$, CantCortes, MedidaDelCorte, NOPE%)
        '1.-CONTROLO QUE LA CANTIDAD TRANSFERIDA P LA O.F. MENOS LO YA CONSUMIDO P LA O.F. TENGA SALDO PARA GENERAR EL CONSUMO
        '2.-SI NO TIENE SALDO ES POR QUE SE DEVOLVIO, O YA SE REALIZO EL CONSUMO, SE VERA EN CONSUMLOT
        '3.-ESTO ES PARA LAS CINTILLAS Y ESLINGSA QUE EN EL CAMBIO DE GENERAR EL CONSUMO AL INFORMAR EL CORTE _
           QUEDARON AFUERA ES DECIR SE INFORMO EL CORTE ANTES DEL CAMBIO Y AHORA CUANDO SE CIERRA NO REALIZA EL CONSUMO.
        '
        DEPOCORTE% = XVALO(Control("", "DEPODEST"))
        CI1% = CODINT%(Codigo$)
        CantiEnMetros# = XVALO(CantCortes) * XVALO(MedidaDelCorte)
        '
        CODFAMI$ = TRIM$(CODFAMIL$(CI1%))
        PorcentajeTolerancia% = 2: EsEslinga% = 0
        If InStr(CODFAMI$, "EE") > 0 Then
            PorcentajeTolerancia% = 20 'si son eslingas PorcentajeTolerancia% = 10
            EsEslinga% = 1
        End If
         '
         'VALIDA SI TIENE SALDO DEL LOTE PARA LA OPERACION ACTUAL OPERACION
        SQLX$ = "SELECT  SUM(CASE WHEN CANTINGRE > 0 THEN PESOKG  ELSE -PESOKG END) AS PESOKG"
        SQLX$ = SQLX$ + " ,SUM(CANTINGRE - CANTSALID) As SaldoMetros"
            'ACA TRAE LO CONSUMIDO DE LA MISMA ORDEN Y OPERACION PARA DESCONTAR, POR QUE EL SALDO SI NO NO LO CONSIDERA
            SQLX$ = SQLX$ + " ,(SELECT SUM(CantSalid) FROM MOVISTO WHERE IDENLOTE = '" & TRIM$(LOTE$) & "'"
            SQLX$ = SQLX$ + " AND  DEPOMOVI = " & DEPOCORTE%
            'SQLX$ = SQLX$ + " AND  ID_OPERACION = " & NOPE%
            SQLX$ = SQLX$ + " AND  CODIMOVI = 'CF'"
            SQLX$ = SQLX$ + " AND DOPRILAR = '" & NORFA$ & "') AS YA_CONSUMIDO"
         SQLX$ = SQLX$ + " From Movisto WITH(NOLOCK)"
         SQLX$ = SQLX$ + " Where   IDENLOTE = '" & TRIM$(LOTE$) & "'"
         SQLX$ = SQLX$ + " AND  DEPOMOVI = " & DEPOCORTE%
         'SQLX$ = SQLX$ + " AND  ID_OPERACION = " & NOPE%
         SQLX$ = SQLX$ + "  AND DOSECLAR = '" & NORFA$ & "'"
         Set rs = READSET(SQLX$)
         If Not rs.EOF Then
            SALDO = XVALO(rs!SaldoMetros): MAXIMOTOLERANCIA# = (SALDO * PorcentajeTolerancia%) / 100

            TRANSFERIDO = SALDO 'LO GUARDO EN OTRA VARIABLE MOSTRARLO EN EL MENSAJE
            YaConsumido# = XVALO(rs!YA_CONSUMIDO)
            SALDO = SALDO - YaConsumido#
            
            If (SALDO + MAXIMOTOLERANCIA# < (CantiEnMetros# - YaConsumido#)) Or SALDO <= 0 Then
                Call COMUNI("El Material Seleccionado No Posee Saldo Suficiente en el Depósito de Corte\Para esta Operación \(Transferido Mts. " & ROUND(TRANSFERIDO, 2) & " \Consumido: " & ROUND(XVALO(rs!YA_CONSUMIDO), 2) & ")")
                GoTo 99
            End If
         End If
         '
         TXT$ = "Se va a Realizar un Consumo por: " & ROUND(CantiEnMetros# - YaConsumido#, 2)
         TXT$ = TXT$ + "\Transferido Para Esta Orden: " & ROUND(TRANSFERIDO, 2) & "\Yá Consumido: " & ROUND(XVALO(rs!YA_CONSUMIDO), 2) & "\Saldo Disponible:" & ROUND(XVALO(SALDO), 2)
         
         OPX% = COMALTER%(TXT$ & "\\Cancelar\Realizar Consumo")

         If OPX% < 2 Then GoTo 99

         'CONSUMO
        CMOV$ = "CF"
        REFE$ = "CONSUMO " & Text10
        DOPRI$ = NORFA$
        DOSEC$ = "PF-" + TRIM$(Mid$(NORFA$, 4))
        NC = XVALO(Text6)
        PesoPorMetro# = ROUND(CoeficientePesoPorMetroSegunLote#(TRIM$(TXTLOTE)), 4)
        PesoTeoricoSegunCantiEnMetros# = CantiEnMetros# * PesoPorMetro#

        Call MOVISTOK(HOY(HO$), CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, 0, 1, NC, DEPOCORTE%, (-1) * CantiEnMetros#, , , , , , , , , , NOPE%, CantiEnMetros#, PesoTeoricoSegunCantiEnMetros#)    'GUARDO EL NUMERO DE OPERACION DE CADA ITEMS TRANSFERIDO
        SQLX$ = "SELECT TOP 1 MOVISTO_ID FROM MOVISTO WITH(NOLOCK) ORDER BY MOVISTO_ID DESC"
        Set rs = READSET(SQLX$)
        If Not rs.EOF Then
          IdMovistoGenerado = XVALO(rs!MOVISTO_ID)
        End If
        rs.Close: Set rs = Nothing

         
            
99     Screen.MousePointer = 1
End Sub

Sub GRABARIMPRESION(IDSECTOR, SECTOR$)
     NROOF$ = TRIM$(Label4)
     Dim CTP As New Conteplast
     CI1% = CODINT(Text10)
     If CI1% < 1 Then
        Call MENSERR%(24, "Codigo Articulo Inválido")
        GoTo 99
     End If
     '
     CIELEM% = CODINT(Label31(0))
     If CIELEM% < 1 Then
        Call MENSERR%(24, "Codigo Componente Inválido")
        GoTo 99
     End If

     If NROOF$ = "" Then
        Call MENSERR%(24, "Debe Seleccionar la Orden de Fabricación.")
        GoTo 99
     End If
     '
     Condi1$ = "CodIConj = " & CI1% & " and  NumeOper = " & XVALO(Label31(3))
     If VALOBADA("SECOPER", "DescOper", Condi1$) = "" Then
         Call MENSERR(24, "Operación Incorrecta para el Articulo Seleccionado.")
         GoTo 99
     End If
     '
     If TRIM$(Label31(1)) = "" Then
         Call MENSERR(24, "Falta Lote Para el Registro Seleccionado.")
         GoTo 99
     End If
     
'     If XVALO(Text1(0)) < 1 Then
'         OPX% = COMALTER%("Falta Seleccionar Operario\\Cancelar\Contiunar sin Informar")
'         If OPX% < 2 Then GoTo 99
'     End If
'
'     If XVALO(TXTCORTE) <= 0.005 Then
'         Call MENSERR(24, "Falta Valor de Corte")
'         GoTo 99
'     End If
     '
     If XVALO(Text1(4)) < 1 Then
         OPX% = COMALTER%("Falta Seleccionar Operario\\Cancelar\Contiunar sin Informar")
         If OPX% < 2 Then GoTo 99
     End If
     NOPE% = IDSECTOR + 1000 ' si es un subproceso como informe de accesororios
     If SECTOR$ = "IMPRESION" Then
       NOPE% = CTP.NumeOperXNombre%(CI1%, "IMPRESION")
       If NOPE% < 1 Then
         Call MENSERR(24, "La Operación IMPRESION, NO Figura en Las Generadas Para Esta Operación")
         GoTo 99
       End If
     End If
     HOII% = HOY(HOS$)
     FEOP% = HOII%
     HORAX$ = Time
     HORAX$ = Left$(HORANUM(HORAX$), 4)
     HORACENTECIIMAL = CHoraSentecimal(HORAX$)
     
     If FEOP = 0 Then
        Call MENSERR(24, "Fecha Incorrecta.")
        GoTo 99
     End If
     If FEOP% > HOII% Then
           Call MENSERR(24, "Fecha No Puede Ser Mayor a la Actual.")
           GoTo 99
         ElseIf FEOP% = HOII% Then
           If HOOP% > HORANUM(Time$) Then
             Call MENSERR(24, "Hora Incorrecta")
             GoTo 99
           End If
     End If
     FEX = FEOP%
     FEHOLI = CDate(FECHATEX$(FEX) + " " + HORATEX$(HOOP%))
     '
     CONDI$ = " IdSubOr = '" & NROOF$ & "'"
     CAPROG = XVALO(VALOBADA("ORDEFABR", "CanProy", CONDI$))
     '
     CAINF# = XVALO(TxtCantimpresa)
     If CAINF# <= 0 Then
        Call MENSERR(24, "Cantidad Inválida")
        GoTo 99
     End If
     '
     CAINF# = XVALO(TxtCantimpresa)
     CANTORIG# = XVALO(Label31(2))
     If CAINF# > CANTORIG# Then
        Call MENSERR(24, "Cantidad Informada Supera La Cantidad Informada Originalmente")
        Call PINTATEXT(TxtCantimpresa)
        GoTo 99
     End If

     TIEINF = 60 * XVALO(1) + XVALO(50) + XVALO(10) / 60
     '
     If TIEINF <= 0 Then 'PARA evitar OVERFLOW
        Call COMUNI("Debe Ingresar un Tiempo Expresado en: \ Horas, Minutos y Segundos")
        GoTo 99
     End If
     '
     If XVALO(KgImpresos) <= 0.0005 Then
        Call MENSERR(24, "Peso Inválido")
        GoTo 99
     End If
     '
     If (XVALO(TxtCantimpresa) + XVALO(Label25)) > XVALO(Label31(2)) Then
        Call MENSERR(24, "Cantidad Ingresada Supera la Cantidad Orignal de los Cortes.")
        GoTo 99
     End If

     On Error GoTo ERROR_GUARDAR
     FLEXCONN.BeginTrans
     Call ABREORFAB
     
     'VALIDO SI LA CANTIDAD QUE SE VA A INFORMAR ES < A LA CANTIDAD DEL ITEM SELECCIONADO
     'DADO QUE SI ES ASI SE DEBERA GENERAR UN NUEVO REGISTRO POR LA DIFERENCIA
     
     LblId1& = XVALO(LblId)
     lblCantidadRegSelec1& = XVALO(lblCantidadRegSelec)
     KILOCALCU# = XVALO(Me.KgImpresos) 'si entra en el if lo modifica
     If TxtCantimpresa < lblCantidadRegSelec1& Then
        CantidadIngresada = XVALO(TxtCantimpresa)
        KilosRegSelec = XVALO(lblKgRegSelec)
        Call CTP.DuplicaRegistrosDeSaldoXDiferencia(lblCantidadRegSelec1&, LblId1&, CantidadIngresada, KilosRegSelec, NOPE%, XVALO(KgImpresos))
        '
'        Diferencia& = XVALO(lblCantidadRegSelec1&) - XVALO(CantidadIngresada)
'        KilosRegistrosSeleccionado = XVALO(KilosRegSelec)
'        CoefKilos# = KilosRegSelec / XVALO(lblCantidadRegSelec1&)
'        KILOCALCU# = CoefKilos# * Diferencia&
'        CAINF# = Diferencia&
'        LblId1& = 0
     End If
     '
     SQLX$ = "UPDATE INFOFAB SET CANTREAL = 0 WHERE INFOFAB_ID =" & LblId1& 'ACTUALIZO EL CAMPO ORIGINAL LA CANTIDAD A 0
     FLEXCONN.Execute (SQLX$) 'POR QUE SE CREA POR EJ DE UNO DE 100  SI SE INFORMAN 30 SE GENERAN 2
                            'UNO DE 30 Y OTRO DE 70 (EL DE 30 ES HIJO DEL DE 100
     '
     
     SQLX$ = " SELECT * FROM INFOFAB "
     SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "' "
     SQLX$ = SQLX$ + " AND NumeOper = " & NOPE% & " "
     '
     CAINFANT% = 0
     Dim INFOFABTMP As ADODB.Recordset
     Set INFOFABTMP = New ADODB.Recordset
     INFOFABTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

     With INFOFABTMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
           ' CAREAL# = CAREAL# + XVALO(!CANTREAL)
           ' TIEINF = TIEINF + XVALO(!TiemEjec)   ' deshabilitado porque si no acumulaba en el nuevo parte
           CAINFANT% = CAINFANT% + 1
         .MoveNext
         Loop
       End If
     End With
'   AL GUARDAR FACTOR DE CONVERSION Y LA CANTIDAD REAL DEL BOLSON
     
     CONDI$ = " INFOFAB_ID = " & XVALO(LblId)
     FAXTORCORTEX = XVALO(VALOBADA("InfoFab", "FACTORCORTE", CONDI$, "NOMESS"))
     If FAXTORCORTEX < 1 Then FAXTORCORTEX = 1
     CORTEX = XVALO(VALOBADA("InfoFab", "CORTE", CONDI$, "NOMESS"))
     
     Call ABREORFAB
     SQLX$ = " SELECT * FROM InfoFab "
     Dim INFOFFFAB As ADODB.Recordset
     Set INFOFFFAB = New ADODB.Recordset
     INFOFFFAB.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
     With INFOFFFAB
        .AddNew   ' YA CONVERTIDO
        On Error Resume Next
        !CODIEMPR = CodiEmp%
        On Error GoTo 0
        !idsubor = NROOF$
        !DescriOF = Left$(ECOARCH$("ECORDEF", NROOF$), 64)
        !NuOrInf = CAINFANT%
        !cod_inte = CI1%
        !Cod_Exte = CODEXT$(CI1%)
        !Descrip = Left$(DESCRIT$(CI1%), 64)
        !NUMEOPER = NOPE%
        !AltOper = 0
        !ReTraba = 0
        !descoper = SECTOR$
        !CanOrPro = CAPROG
        !MAQUINA = ""
        !DescMaq = ""
        !LegaOps = ""
        !FeHorLib = FEHOLI
        !FEHORINF = Now
        !Hora_Des = HORACENTECIIMAL
        !Hora_Has = HORACENTECIIMAL
        !MaqLiber = ""  ' LA MAQUINA ESTA LIBERADA PORQUE LA O/F ESTA CERRADA
        !TiemEjec = HORACENTECIIMAL
           MODO% = CAINFANT%
        !TiemStan = TIESTAN(CI1%, NOPE%, CAINF#, MODO%)
        RENDI = 999.9: If !TiemStan > 0 Then RENDI = 100 * XVALO(!TiemEjec) / XVALO(!TiemStan)
        If RENDI > 999.9 Then RENDI = 999.9
        !RENDIM = RENDI
        !Observa = ""
        !Referen = " "
        !MotiRecha = 0
        !idenlote = Label31(1)
        !codielem = CODINT%(Label31(0))
        !CodXElem = TRIM$(Label31(0))
        'se divide por la cantidad p.ej. un bolson usa 4 cortes de 80 cm
        'en la formula para ese componente es 3.20 m.
        'de este modo se guarda la cantidad 1 a 1 del producto fabricado 4 cortes equivalen a 1 bolson
        '!CANTREAL = Int(CAINF# / FactCte) 'lo cambie a solo la cantidad informada 04/05/20
        !CANTREAL = Int(CAINF#) 'cuando se consuma debera considerar el calculo para que coincida con formula
        !CantOrigCte = 0
        !FACTORCORTE = FAXTORCORTEX 'REFIERE A LA CANTIDAD DE CORTES
        !NUMOPERARIO = XVALO(Text1(4))
        !NameOperario = Left$(Label5(4), 64)
        !corte = XVALO(CORTEX)
        !Peso = KILOCALCU# 'XVALO(Me.KgImpresos)
        !IDREGPADRE = LblId1&
       .Update
     End With
     Set INFOFFFAB = Nothing
     '
     SQLX$ = "SELECT @@IDENTITY AS NUEVOID FROM INFOFAB WITH(NOLOCK)"
     Set rs = READSET(SQLX$)
     NuevoIDGenerado& = rs!NUEVOID

     SQLX$ = " SELECT * FROM INFOFAB "
     SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "' "
     SQLX$ = SQLX$ + " AND NumeOper = " & NOPE% & " "
     '
     CAREAL# = 0: TIEINF = 0
     Dim INFOFABTMP1 As ADODB.Recordset
     Set INFOFABTMP1 = New ADODB.Recordset
     INFOFABTMP1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    
     With INFOFABTMP1
         On Error Resume Next
         .MoveFirst
         If Err < 1 Then
           On Error GoTo 0
           Do While Not .EOF
             'If IsNull(!ReTraba) Or !ReTraba < 1 Then
             If XVALO(!ReTraba) < 1 Then
               CAREAL# = CAREAL# + XVALO(!CANTREAL)
             End If
             TIEINF = TIEINF + XVALO(!TiemEjec)
           .MoveNext
           Loop
         End If
         On Error GoTo 0
     End With
    '
    SQLZ$ = " SELECT * FROM OPERFAB "
    SQLZ$ = SQLZ$ + " WHERE IdSubOr = '" & NROOF$ & "' AND NumeOper = " & NOPE%
    Dim OPERBU As ADODB.Recordset
    Set OPERBU = New ADODB.Recordset
    OPERBU.Open FILTSQL$(SQLZ$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

    With OPERBU
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
        On Error GoTo 0
       !CANTREAL = CAREAL#
       !CANTPEND = XVALO(!CANTPROG) - CAREAL#
         If XVALO(!CANTPEND) < 0 Or STATUOP% = 3 Or STATUOP% = 8 Then !CANTPEND = 0
         CANTPE# = XVALO(!CANTPEND)
         MODO% = CAINFANT%
       !TiemEjec = TIEINF
       !TiemPend = TIEMAQUI(CI1%, NOPE%, CANTPE#, MODO%)  ' 1 es para que no considere preparacion
          If !TiemPend < 0 Or STATUOP% = 3 Or STATUOP% = 0 Then !TiemPend = 0
       !StatOper = STATUOP%
       .Update
     End If
     On Error GoTo 0
   End With
   '

   'BORRAR LA UBICACION
   SQLX$ = "UPDATE INFOFAB SET COORDUBIC = '',IDCOORDUBIC=0"
   SQLX$ = SQLX$ + "  WHERE INFOFAB_ID = " & XVALO(Me.LblId)
   FLEXCONN.Execute SQLX$

ERROR_GUARDAR:
    If Err <> 0 Then
       MsgBox Err.Description
       On Error Resume Next
       On Error GoTo 0
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION")
       FLEXCONN.RollbackTrans
       Exit Sub
    Else
      
      FLEXCONN.CommitTrans
      
      Call COMUNI("Proceso Realizado Exitosamente")
      
       
      INFOFABTMP.Close
      Set INFOFABTMP = Nothing

      INFOFABTMP1.Close
      Set INFOFABTMP1 = Nothing

      OPERBU.Close
      Set OPERBU = Nothing

    End If
    

     Call CargaMatTranfXOrdenSoloIngresos
     
     CIFAB% = CODINT%(Text10)
     Call CTP.CargaOperacionesInformadasMasSubProcesos(NROOF$, CIFAB%, MSF2, 0, "", 99)
     Set CTP = Nothing

     '++++++++++++++++++++++++++++++++++++++++++++++++++++
     'BUSCO EL REGISTRO QUE SE CREO PARA QUE QUEDE SELECCIONADO
     
     For i& = 1 To MSF2.Rows - 1
        If XVALO(MSF2.TextMatrix(i&, 16)) = NuevoIDGenerado& Then
               Call SELECCIONARFILA(MSF2, i&)
               Exit For
        End If
     Next i&
'
'
'
'
'
'
'     CIELEM% = CODINT%(Label31(0)) 'material
'     LOTE$ = TRIM$(Label31(1))
'     CANTI = XVALO(TxtCantimpresa) 'cantidad de cortes)
'
'
'     NC = XVALO(Text6)
'     NROPED& = XVALO(Text12)
'     NROOF$ = Label4
'     ORDENINFO = XVALO(MSF2.TextMatrix(REG&, 11))
'     UBICACION$ = CAINFANT%
'     IDUBICACION = XVALO(MSF2.TextMatrix(REG&, 12))
'     DESCRIOPERACION$ = MSF2.TextMatrix(REG&, 7)
'     CODIGOINTFAB = CODINT(Text10)
'     Call PRINTETIQPALLET(CODEXT$(CIELEM%), LOTE$, CANTI, NOPE%, NC, NROPED&, NROOF$, ORDENINFO, UBICACION$, IDUBICACION, DESCRIOPERACION$, CODIGOINTFAB)
     '+++++++++++++++++++++++++++++++++++++++++++++++


     Call BLANFORMU
     Call BLANFORMUIMPRESION
99
End Sub

Private Sub CmdCerrarFucion_Click()
    NROOF$ = TRIM$(Label4)
    If NROOF$ = "" Then Exit Sub
    
    CIFAB% = CODIFAB%(NORFA$)
    
    Call BLANFORMU_FUCION
    Frame4.Visible = False
    Frame2(1).Visible = True
    Dim CTP As New Conteplast
    Call CTP.CargaOperacionesInformadasMasSubProcesos(NROOF$, CODIGOINTFAB%, MSF2, 0, "", 99)
    Set CTP = Nothing
    Call VALIDARBOTONFUSION



End Sub

Private Sub CmdInfoFusion_Click()
    For i& = 1 To MSF2.Rows - 1
       If MSF2.TextMatrix(i&, 17) > 0 And MSF2.TextMatrix(i&, 0) = "*" Then
          Call COMUNI("No Válido Informar Sobre Sub-Procesos")
          Exit Sub
       End If
    Next i&
    '
    
    AB = LISTASELECGENERAL("OPERFUSI", "", "", 0)
    Me.lblCodFusion = XVALO(RESP_ZELE_VECT(1))
    Me.lblDescFusion = TRIM$(RESP_ZELE_VECT(2))

End Sub

Private Sub Command1_Click()
   command1.Enabled = False
   Dim ActualizaListaSeleccion As Boolean
   ActualizaListaSeleccion = True
   COD$ = ListaSeleccionArticulos("", ActualizaListaSeleccion)
   ActualizaListaSeleccion = False

   If COD$ <> "" Then
     Text10.TEXT = COD$
   End If
   command1.Enabled = True
End Sub



Private Sub Command10_Click()
    If TRIM$(Label31(7)) = "" Then
      Call COMUNI("Lectura Incorrecta,Vuelva a Leer Los Datos o Seleccione El Corte de la Grilla")
      Exit Sub
    End If
    AB = LISTASELECGENERAL("SECTPLANTA", "", "", 0, "CONVERT(int, campo1) < 100")
    IDSECTOR = XVALO(RESP_ZELE_VECT(1))
    SECTOR$ = TRIM$(RESP_ZELE_VECT(2))
    If SECTOR$ = "" Then Exit Sub

    Call GRABARIMPRESION(IDSECTOR, SECTOR$)
    TxtCantimpresa.Locked = True 'vuelve a poner a locked = true
End Sub

Private Sub Command11_Click()


    Command11.Enabled = False
    Codigo$ = TRIM$(TXTCODELEM)
    CI1% = CODINT%(Codigo$)
    If CI1% < 1 Then
      Call COMUNI("Código No Válido")
      GoTo 99
    End If
    
'    Dim ctp As New Conteplast
'    PESOX# = ctp.PesoLeidoBalanza(CI1%, 2)
    
    ''''''
    LblKgImpresos = 0
    PREBALANZA.FORMULARIO_NAME = Me.Name
    PREBALANZA.LblKgImpresos = 0
    PREBALANZA.Show 1
    PESOLEIDO# = XVALO(PREBALANZA.LblKgImpresos)
    If PESOLEIDO# <= 0 Then
       Call COMUNI("Peso Ingresado No Válido")
       GoTo 99
    End If
    LblKgImpresos = PESOLEIDO#
    If PESOLEIDO# < 0 Then PESOLEIDO# = 0
    
    LblKgImpresos = FORMATNUM$(PESOLEIDO#, "f12.2")

    '''''
99  Command11.Enabled = True

End Sub



Private Sub Command12_Click()
    Command12.Enabled = False
'    Codigo$ = Label31(0)
'    CI1% = CODINT%(Codigo$)
'    If CI1% < 1 Then
'      Call COMUNI("Código No Válido")
'      GoTo 99
'    End If
    
    Dim CTP As New Conteplast
    PESOX# = CTP.PesoLeidoBalanza(CI1%, 2) '
    If PESOX# < 0 Then PESOX# = 0
    Label26 = TRIM$(FORMATNUM$(PESOX#, "f12.2"))

'    KgImpresos = FORMATNUM$(CTP.PesoLeidoBalanza(CI1%), "f12.2")
99  Command12.Enabled = True

End Sub

Private Sub Command13_Click()
        'GENERAR UN REGISTRO CON NUMEOPER DE LA FUSION + 1200
        'GRABAR DESPUES QUE SE GENERO EL REGISTRO SOBRE LOS DOS CODIGOS EN EL OperFucionada EL ID del registro creado
        '
'       ver que lote (juntar los 2 lotes)separados x ;
'       que nuopper  + 1200
'       que codigo interno de componentes =0
'       que corte en metros = 0
        Dim CTP As New Conteplast
        Dim IDSELEC&(2)
        Dim CantEnRegSelex&(2)
        Dim NumOperacionRegSelec%(2)
        Dim KGEnRegSelec(2)
        
        Command13.Enabled = False
        CantdeCortes = XVALO(Text16)
        PESOKG = XVALO(Label26)
        
        If XVALO(lblCodFusion) < 1 Then
           Call COMUNI("Atención! Debe Ingresar Una Operación de Fusión")
           GoTo 99
        End If
        
        If XVALO(PESOKG) <= o Then
           Call COMUNI("Atención! Peso No Válido")
           GoTo 99
        End If
        '
        AB = LISTASELECGENERAL("SECTPLANTA", "", "", 0, "CONVERT(int, campo1) < 100")
        IDSECTOR = XVALO(RESP_ZELE_VECT(1))
        SECTOR$ = TRIM$(RESP_ZELE_VECT(2))
        If SECTOR$ = "" Then
           Call COMUNI("Atención! Debe Ingresar Un Sector de Ubicación\Podrá Cambiarlo Luego de Generar la Transacción")
           GoTo 99
        End If
        NROOF$ = TRIM$(Label4)

        If NROOF$ = "" Then
          Call MENSERR%(24, "Debe Seleccionar la Orden de Fabricación.")
          GoTo 99
        End If
        '
        If XVALO(Text1(5)) < 1 Then
           OPX% = COMALTER%("Falta Seleccionar Operario\\Cancelar\Contiunar sin Informar")
           If OPX% < 2 Then GoTo 99
        End If


        NUMEOPERARIO% = XVALO(Text1(5))
        If NUMEOPERARIO% < 1 Then
           Call COMUNI("Atención! Falta Ingresar Operario")
           GoTo 99
        End If
        
        If ECOARCH$("PADRON", MKI$(NUMEOPERARIO%)) = "" Then
           Call MENSERR(24, "Legajo Operario no Válido")
           GoTo 99
        End If
        
        
        If CantdeCortes < 1 Then
           Call COMUNI("Atención !! Imposible Continuar, Cantidad No Válida")
           GoTo 99
        End If
        
        '
        'GRABA LA FUSION DE 2 OPERACIONES
        NUMEOPERACION% = XVALO(lblCodFusion)
        CIELEM% = 0
        DESCRIOPERACION$ = TRIM$(lblDescFusion)
        LOTE$ = ""
        NC = XVALO(Text6)
        'CARGO EL LOTE
        
        
        JX& = 0
        LblId = XVALO(MSF2.TextMatrix(Reg&, 16))
        For i& = 1 To MSF2.Rows - 1
           If TRIM$(MSF2.TextMatrix(i&, 0)) = "*" Then
             If LOTE$ = "" Then
               LOTE$ = TRIM$(MSF2.TextMatrix(i&, 3))
               
             Else
               LOTE$ = LOTE$ & ";" & TRIM$(MSF2.TextMatrix(i&, 3))
             End If
             JX& = JX& + 1
             CantEnRegSelex&(JX&) = XVALO(MSF2.TextMatrix(i&, 4))
             NumOperacionRegSelec%(JX&) = XVALO(MSF2.TextMatrix(i&, 6))
             KGEnRegSelec(JX&) = XVALO(MSF2.TextMatrix(i&, 15))
             IDSELEC&(JX&) = XVALO(MSF2.TextMatrix(i&, 16))
           End If
        Next i&
        '
        If CantdeCortes > CantEnRegSelex&(1) Or CantdeCortes > CantEnRegSelex&(2) Then
           Call COMUNI("Atención !! Imposible Continuar, Cantidad No Válida (No puede ser Superior a Ninguno de los Materiales Seleccionados")
           GoTo 99
        End If
       
        IDCOORDUBIC& = XVALO(IDSECTOR)
        UBICACION$ = TRIM$(SECTOR$)
        NROPED& = XVALO(Text12)
        CODIGOINTFAB% = CODINT(Text10)
        On Error GoTo ERROR_GUARDAR
        FLEXCONN.BeginTrans
        
        'SI LA CANTIDAD DE CORTES INGRSADOS SON MENORES A LA CANTIDAD DEL 1ER REGISTRO SELECCIONADO
        If CantdeCortes < CantEnRegSelex&(1) Then
          Call CTP.DuplicaRegistrosDeSaldoXDiferencia(CantEnRegSelex&(1), IDSELEC&(1), CantdeCortes, KGEnRegSelec(1), NumOperacionRegSelec%(1), XVALO(PESOKG), 1)
        End If
        
        'SI LA CANTIDAD DE CORTES INGRSADOS SON MENORES A LA CANTIDAD DEL 2DO REGISTRO SELECCIONADO
        If CantdeCortes < CantEnRegSelex&(2) Then
          Call CTP.DuplicaRegistrosDeSaldoXDiferencia(CantEnRegSelex&(2), IDSELEC&(2), CantdeCortes, KGEnRegSelec(2), NumOperacionRegSelec%(2), XVALO(PESOKG), 1)
        End If
        '
        
        Call CTP.GrabarOperacion(NROOF$, CODIEMPR%, CODIGOINTFAB%, NUMEOPERACION%, CantdeCortes, LOTE$, CIELEM%, 0, NUMEOPERARIO%, 0, PESOKG, RET_NEWORDEN%, Ret_NuevoIDGenerado&, IDCOORDUBIC&, UBICACION$, IDSELEC&(1), DESCRIOPERACION$)
       
       'BORRAR LA UBICACION, GRABAR EN OPERFUSIONADA
        SQLX$ = "UPDATE INFOFAB SET COORDUBIC = '',IDCOORDUBIC=0,OPERFUCIONADA=  " & Ret_NuevoIDGenerado&
        SQLX$ = SQLX$ + "  WHERE INFOFAB_ID = " & IDSELEC&(1)
        FLEXCONN.Execute SQLX$
        '
        SQLX$ = "UPDATE INFOFAB SET COORDUBIC = '',IDCOORDUBIC=0, OPERFUCIONADA= " & Ret_NuevoIDGenerado&
        SQLX$ = SQLX$ + "  WHERE INFOFAB_ID = " & IDSELEC&(2)
        FLEXCONN.Execute SQLX$


ERROR_GUARDAR:
    If Err <> 0 Then
       On Error Resume Next
       On Error GoTo 0
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION")
       FLEXCONN.RollbackTrans
       GoTo 99
    Else
      FLEXCONN.CommitTrans
      Call COMUNI("Proceso Realizado Exitosamente")
      Call BLANFORMU_FUCION
    End If
    Call CTP.CargaOperacionesInformadasMasSubProcesos(NROOF$, CODIGOINTFAB%, MSF2, 0, "", 99)
    Call VALIDARBOTONFUSION

    Set CTP = Nothing

99  Command13.Enabled = True

End Sub

Private Sub Command14_Click()
    Label4 = OrdenFabricacionAbiertas$(1)

End Sub

Private Sub Command15_Click()
   If Command15.Caption = "!" Then
      If TRIM$(Label4) = "" Then Exit Sub
      Call CargarGrillaInfofabCompleta(Label4, MSF3)
      MSF3.Visible = True
      MSF3.Height = 7290
      MSF3.Width = 14865
      Command15.Caption = "X"
   Else
      MSF3.Visible = False
      MSF3.Height = 800
      MSF3.Width = 800
      Command15.Caption = "!"
   End If
      
End Sub
Sub CargarGrillaInfofabCompleta(NORFA$, MSF As msFlexGrid)
'SELECT NuOrInf,NumeOper,DescOper,CANTREAL,FeHorInf,CODXELEM,FACTORCORTE,COORDUBIC,IDCOORDUBIC,CORTE,NAMEOPERARIO,PESO,INFOFAB_ID,IdRegPadre,OperFucionada
'FROM INFOFAB WHERE IDSUBOR = 'OF-000103-1'  and IdRegPadre > 0
'ORDER BY IDENLOTE,CodXElem, INFOFAB_ID,IdRegPadre
'
'SELECT IDSUBOR, NuOrInf,NumeOper,DescOper,CANTREAL,FeHorInf,CODXELEM,FACTORCORTE,COORDUBIC,IDCOORDUBIC,CORTE,NAMEOPERARIO,PESO,INFOFAB_ID,IdRegPadre,OperFucionada
'FROM INFOFAB WHERE (OperFucionada < = 0 or OperFucionada is null) and (IdRegPadre < = 0 or IdRegPadre is null) and IDSUBOR = 'OF-000103-1'
'ORDER BY IDSUBOR ,IDENLOTE,CodXElem, INFOFAB_ID,IdRegPadre
'
'SELECT NuOrInf,NumeOper,DescOper,CANTREAL,FeHorInf,CODXELEM,FACTORCORTE,COORDUBIC,IDCOORDUBIC,CORTE,NAMEOPERARIO,PESO,INFOFAB_ID,IdRegPadre,OperFucionada
'FROM INFOFAB WHERE IDSUBOR = 'OF-000103-1'  and OperFucionada > 0
'ORDER BY IDENLOTE,CodXElem, INFOFAB_ID,IdRegPadre
'
'SELECT NuOrInf,NumeOper,DescOper,CANTREAL,IDENLOTE,FeHorInf,CODXELEM,FACTORCORTE,COORDUBIC,IDCOORDUBIC,CORTE,NAMEOPERARIO,PESO,INFOFAB_ID,IdRegPadre,OperFucionada
'FROM INFOFAB WHERE IDSUBOR = 'OF-000103-1'  and (INFOFAB_ID = 7709  or OperFucionada = 7709)
'ORDER BY OperFucionada desc
'
'SELECT NuOrInf,NumeOper,DescOper,CANTREAL,FeHorInf,IDENLOTE,CODXELEM,FACTORCORTE,COORDUBIC,IDCOORDUBIC,CORTE,NAMEOPERARIO,PESO,INFOFAB_ID,IdRegPadre,OperFucionada
'FROM INFOFAB I WHERE IDSUBOR = 'OF-000103-1'  AND
'EXISTS (SELECT * FROM INFOFAB AS I2 WHERE I2.OperFucionada = I.INFOFAB_ID)

        Call CARGA_ENCABEZADO(MSF3, "CI/A0;Código/A26;NUMEOPER/A0;Operación/A24;Lote/A16;Cantreal/F9.1;Corte/F9.1;Peso/f9.2;Infofab_id/A0;idcoordubic/A0;Coordubic/A24;Consumo/A0", Me)
        'CONSULTA TRAE TODAS LOS INFORMES QUE NO TIENEN IDREGPADRE NI OPERFUCIONADA
        J& = 0
        SQLX$ = "SELECT  NuOrInf,NumeOper,DescOper,CANTREAL,FeHorInf,CODXELEM,FACTORCORTE,COORDUBIC,CODIELEM"
        SQLX$ = SQLX$ + " ,IDCOORDUBIC,CORTE,NAMEOPERARIO,PESO,INFOFAB_ID,IdRegPadre,OperFucionada,IDENLOTE,CONSUMO"
        SQLX$ = SQLX$ + " FROM INFOFAB WITH(NOLOCK) WHERE (OperFucionada < = 0 or OperFucionada is null)"
        SQLX$ = SQLX$ + " and (IdRegPadre < = 0 or IdRegPadre is null) and IDSUBOR = '" & NORFA$ & "'"
        SQLX$ = SQLX$ + " ORDER BY DescOper,IDENLOTE,CodXElem, INFOFAB_ID,IdRegPadre"
        Set rs = READSET(SQLX$)
        With rs
            Do While Not .EOF
              J& = J& + 1
              MSF.Rows = J& + 1
              MSF.TextMatrix(J&, 1) = XVALO(!codielem)
              MSF.TextMatrix(J&, 2) = SAFETEXT$(!CodXElem)
              MSF.TextMatrix(J&, 3) = XVALO(!NUMEOPER)
              MSF.TextMatrix(J&, 4) = SAFETEXT$(!descoper)
              LOTE$ = SAFETEXT$(!idenlote)
              MSF.TextMatrix(J&, 5) = LOTE$
              MSF.TextMatrix(J&, 6) = FORMATNUM$(XVALO(!CANTREAL), "F9.1")
              MSF.TextMatrix(J&, 7) = FORMATNUM$(XVALO(!corte), "F9.1")
              MSF.TextMatrix(J&, 8) = FORMATNUM$(XVALO(!Peso), "F9.2")
              ID1& = XVALO(!INFOFAB_ID)
              MSF.TextMatrix(J&, 9) = XVALO(ID1&)
              MSF.TextMatrix(J&, 10) = SAFETEXT$(!IDCOORDUBIC)
              MSF.TextMatrix(J&, 11) = SAFETEXT$(!COORDUBIC)
              MSF.TextMatrix(J&, 12) = XVALO(!CONSUMO)
              INGRESO% = 0 'INCALIZO CONTADOR DE LA SIGUIENTE CONSULTA
              'ESTA CONSULTA TRAE LOS SUB-PROCESOS PARA EL LOTE POR EL REGISTRO PADRE
              SQLX$ = "SELECT CODIELEM,CODXELEM,NUMEOPER,DescOper,IDENLOTE,CANTREAL,CORTE,PESO,INFOFAB_ID"
              SQLX$ = SQLX$ + ",IDCOORDUBIC,COORDUBIC,CONSUMO FROM INFOFAB WITH(NOLOCK) WHERE IDSUBOR = '" & NORFA$ & "'"
              SQLX$ = SQLX$ + " AND IDENLOTE = '" & LOTE$ & "'"
              SQLX$ = SQLX$ + " AND IdRegPadre = " & ID1&
              Set rs1 = READSET(SQLX$)
              With rs1
                Do While Not .EOF
                    If INGRESO% < 1 Then
                       Call COLOREAR_GRILLA_SOLO_UNA_AFILA(MSF, &HE2D3C0, J&)
                       Call NEGRITA_SOLO_UNA_FILA(MSF, J&, 10)
                       INGRESO% = 1
                    End If
                    
                    J& = J& + 1
                    MSF.Rows = J& + 1
                    MSF.TextMatrix(J&, 1) = XVALO(!codielem)
                    MSF.TextMatrix(J&, 2) = "**********" & SAFETEXT$(!CodXElem)
                    MSF.TextMatrix(J&, 3) = XVALO(!NUMEOPER)
                    MSF.TextMatrix(J&, 4) = SAFETEXT$(!descoper)
                    LOTE$ = SAFETEXT$(!idenlote)
                    MSF.TextMatrix(J&, 5) = LOTE$
                    MSF.TextMatrix(J&, 6) = FORMATNUM$(XVALO(!CANTREAL), "F9.1")
                    MSF.TextMatrix(J&, 7) = FORMATNUM$(XVALO(!corte), "F9.1")
                    MSF.TextMatrix(J&, 8) = FORMATNUM$(XVALO(!Peso), "F9.2")
                    ID1& = XVALO(!INFOFAB_ID)
                    MSF.TextMatrix(J&, 9) = XVALO(ID1&)
                    MSF.TextMatrix(J&, 10) = SAFETEXT$(!IDCOORDUBIC)
                    MSF.TextMatrix(J&, 11) = SAFETEXT$(!COORDUBIC)
                    MSF.TextMatrix(J&, 12) = XVALO(!CONSUMO)
                    .MoveNext
                Loop
              End With
              rs1.Close
              Set rs1 = Nothing

              .MoveNext
            Loop
        End With
        rs.Close
        Set rs = Nothing
        
        
        'AHORA PRESENTO LAS FUSIONES
        SQLX$ = "SELECT CODIELEM,CODXELEM,NUMEOPER,DescOper,IDENLOTE,CANTREAL,CORTE,PESO,INFOFAB_ID"
        SQLX$ = SQLX$ + ",IDCOORDUBIC,COORDUBIC,CONSUMO FROM INFOFAB I WITH(NOLOCK) WHERE IDSUBOR = '" & NORFA$ & "'"
        SQLX$ = SQLX$ + " AND EXISTS (SELECT * FROM INFOFAB AS I2 WHERE I2.OperFucionada = I.INFOFAB_ID)"
        'ESTA CONSULTA ME TRAE TODAS LAS TAPAS COMPLETAS
        Set RS2 = READSET(SQLX$)
        With RS2
              With RS2
                Do While Not .EOF
                    ID1& = XVALO(!INFOFAB_ID)
                    'AHORA BUSCO LOS REGISTROS QUE ARMARON LA TAPA COMPLETA
                    SQLX$ = "SELECT CODIELEM,CODXELEM,NUMEOPER,DescOper,IDENLOTE,CANTREAL,CORTE,PESO,INFOFAB_ID"
                    SQLX$ = SQLX$ + ",IDCOORDUBIC,COORDUBIC,CONSUMO FROM INFOFAB I WITH(NOLOCK) WHERE IDSUBOR = '" & NORFA$ & "'"
                    SQLX$ = SQLX$ + " AND OPERFUCIONADA = " & ID1&
                    'ESTA CONSULTA ME TRAE LAS VALVULAS Y LAS TAPAS QUE AARMARON LA TAPA COMPLETA
                    Set rs3 = READSET(SQLX$)
                    With rs3
                            Do While Not .EOF
                                J& = J& + 1
                                MSF.Rows = J& + 1
                                MSF.TextMatrix(J&, 1) = XVALO(!codielem)
                                MSF.TextMatrix(J&, 2) = SAFETEXT$(!CodXElem)
                                MSF.TextMatrix(J&, 3) = XVALO(!NUMEOPER)
                                MSF.TextMatrix(J&, 4) = SAFETEXT$(!descoper)
                                LOTE$ = SAFETEXT$(!idenlote)
                                MSF.TextMatrix(J&, 5) = LOTE$
                                MSF.TextMatrix(J&, 6) = FORMATNUM$(XVALO(!CANTREAL), "F9.1")
                                MSF.TextMatrix(J&, 7) = FORMATNUM$(XVALO(!corte), "F9.1")
                                MSF.TextMatrix(J&, 8) = FORMATNUM$(XVALO(!Peso), "F9.2")
                                ID1& = XVALO(!INFOFAB_ID)
                                MSF.TextMatrix(J&, 9) = XVALO(ID1&)
                                MSF.TextMatrix(J&, 10) = SAFETEXT$(!IDCOORDUBIC)
                                MSF.TextMatrix(J&, 11) = SAFETEXT$(!COORDUBIC)
                                MSF.TextMatrix(J&, 12) = XVALO(!CONSUMO)
                                Call COLOREAR_GRILLA_SOLO_UNA_AFILA(MSF, 13400576, J&)
                                Call NEGRITA_SOLO_UNA_FILA(MSF, J&, 10)

                                .MoveNext
                            Loop
                    End With
                    rs3.Close
                    Set rs3 = Nothing
                    '
                    J& = J& + 1
                    MSF.Rows = J& + 1
                    MSF.TextMatrix(J&, 1) = XVALO(!codielem)
                    MSF.TextMatrix(J&, 2) = ">>>>>>>>>" & SAFETEXT$(!CodXElem)
                    MSF.TextMatrix(J&, 3) = XVALO(!NUMEOPER)
                    MSF.TextMatrix(J&, 4) = SAFETEXT$(!descoper)
                    LOTE$ = SAFETEXT$(!idenlote)
                    MSF.TextMatrix(J&, 5) = LOTE$
                    MSF.TextMatrix(J&, 6) = FORMATNUM$(XVALO(!CANTREAL), "F9.1")
                    MSF.TextMatrix(J&, 7) = FORMATNUM$(XVALO(!corte), "F9.1")
                    MSF.TextMatrix(J&, 8) = FORMATNUM$(XVALO(!Peso), "F9.2")
                    ID1& = XVALO(!INFOFAB_ID)
                    MSF.TextMatrix(J&, 9) = XVALO(ID1&)
                    MSF.TextMatrix(J&, 10) = SAFETEXT$(!IDCOORDUBIC)
                    MSF.TextMatrix(J&, 11) = SAFETEXT$(!COORDUBIC)
                    MSF.TextMatrix(J&, 12) = XVALO(!CONSUMO)
                    .MoveNext
                Loop
              End With

            RS2.Close
            Set RS2 = Nothing

        End With

        
End Sub
Private Sub Command2_Click()
     Command2.Enabled = False
    '
     CantiEnMetros# = ROUND(XVALO(Text7) * XVALO(TXTCORTE), 4)
     EPAC$ = UCase$(TRIM$(EMPREAC$))

     NROOF$ = TRIM$(Label4)

     CI% = CODINT(Text10)
     If CI% < 1 Then
        Call MENSERR%(24, "Codigo Articulo Inválido")
        GoTo 99
     End If
     '
     CIELEM% = CODINT(TXTCODELEM)
     If CIELEM% < 1 Then
        Call MENSERR%(24, "Codigo Componente Inválido")
        GoTo 99
     End If

     If NROOF$ = "" Then
        Call MENSERR%(24, "Debe Seleccionar la Orden de Fabricación.")
        GoTo 99
     End If
     '
     Condi1$ = "CodIConj = " & CI% & " and  NumeOper = " & XVALO(Text3)
     If VALOBADA("SECOPER", "DescOper", Condi1$) = "" Then
         Call MENSERR(24, "Operación Incorrecta para el Articulo Seleccionado.")
         GoTo 99
     End If
     '
     If TRIM$(TXTLOTE) = "" Then
         Call MENSERR(24, "Falta Lote Para el Registro Seleccionado.")
         GoTo 99
     End If
     '
     NOPE% = XVALO(Text3)
     If NOPE% < 1 Then
        Call MENSERR%(24, "Falta Identificar Tipo de Operación.")
        GoTo 99
     End If
'     If (XVALO(Text7)) > XVALO(TXTRESTAN) * 1.025 Then  'SOLO TOLERA TRANSFERIR HASTA UN 2,5
      MaximoATolerar = ROUND(XVALO(TXTNECESIDADCORTES) * 1.05, 0) 'se aumento al 5 % '
      CortesInformadosMasCortesAInformar = XVALO(TXTCORTESOPERACION) + XVALO(Text7)
      If CortesInformadosMasCortesAInformar > MaximoATolerar Then
         MaxiATolerar$ = FORMATNUM$(MaximoATolerar, "f10.0")
         MaximoAInformar = TRIM$(FORMATNUM$(XVALO(TXTNECESIDADCORTES * 1.05) - XVALO(TXTCORTESOPERACION), "f10.0"))
         Call MENSERR(24, "Imposible Informar Cantidad de Cortes Superior a la Necesidad\Tolerancia 5%\El Màximo a Informar es: " & MaximoAInformar)
         GoTo 99
     End If
     
     If XVALO(TXTCORTE) <= 0.005 Then
         Call MENSERR(24, "Falta Valor de Corte")
         GoTo 99
     End If
     '
     
'     If TRIM$(ECOARCH$("PADMAQ", TRIM$(Text5))) = "" Then
'         Call MENSERR(24, "Equipo Incorrecto.")
'         GoTo 99
'     End If
     '
'     LEGA$ = ""
'     OPERS% = 0
'     For i% = 0 To 3
'        If Text1(i%) <> "" Then
'            OPERS% = 1
'            LEGAJO% = XVALO(Text1(i%))
'            If LEGAJO% > 0 Then
'                If ECOARCH$("PADRON", MKI$(LEGAJO%)) = "" Then
'                  Call MENSERR(24, "Legajo Operario no Válido")
'                  GoTo 99
'                Else
'                  LEGA$ = LEGA$ & MKI$(LEGAJO%)
'                End If
'            End If
'         End If
'     Next i%
'     If OPERS% = 0 Then
'         Call MENSERR(24, "Falta Definir Operarios")
'         GoTo 99
'     End If
     '
     FEOP% = FECHANUM(Text4)
     HORAX$ = Time
     HORAX$ = Left$(HORANUM(HORAX$), 4)
     HORACENTECIIMAL = CHoraSentecimal(HORAX$)
     
     HOII% = HOY(HOS$)
     If FEOP = 0 Then
        Call MENSERR(24, "Fecha Incorrecta.")
        GoTo 99
     End If
     If FEOP% > HOII% Then
           Call MENSERR(24, "Fecha No Puede Ser Mayor a la Actual.")
           GoTo 99
         ElseIf FEOP% = HOII% Then
           If HOOP% > HORANUM(Time$) Then
             Call MENSERR(24, "Hora Incorrecta")
             GoTo 99
           End If
     End If
     FEX = FEOP%
     FEHOLI = CDate(FECHATEX$(FEX) + " " + HORATEX$(HOOP%))
     '
     CONDI$ = " IdSubOr = '" & NROOF$ & "'"
     CAPROG = XVALO(VALOBADA("ORDEFABR", "CanProy", CONDI$))
     '
     CAINF# = XVALO(Text7)
     If CAINF# = 0 Then
        Call MENSERR(24, "Cantidad Invalida")
        GoTo 99
     End If
     '
     TIEINF = 60 * XVALO(1) + XVALO(50) + XVALO(10) / 60
     '
     If TIEINF <= 0 Then 'PARA evitar OVERFLOW
        Call COMUNI("Debe Ingresar un Tiempo Expresado en: \ Horas, Minutos y Segundos")
        GoTo 99
     End If
     '
     If XVALO(LblKgImpresos) <= 0.0005 Then
        Call MENSERR(24, "Peso Inválido o Falta Informar")
        GoTo 99
     End If
     '
     'VALIDAR EL PESO SEGUN LOS METROS INFORMADOS EN CORTES
     PesoPorMetro# = ROUND(CoeficientePesoPorMetroSegunLote#(TRIM$(TXTLOTE)), 4) '0.612
     PesoTeoricoSegunCantiEnMetros# = CantiEnMetros# * PesoPorMetro# '42.28
     If PesoPorMetro# > 0 Then
        diferenciaPorcentual = Abs(((XVALO(LblKgImpresos) - PesoTeoricoSegunCantiEnMetros#) / PesoTeoricoSegunCantiEnMetros#) * 100)  'Porcentaje de diferencia = ((Nuevo valor - Valor de origen) / Valor de origen) * 100
        If diferenciaPorcentual > 10 Then
           Call MENSERR(24, "Peso Informado No Tiene Relación con el informado en la Recepción")
           GoTo 99
        End If
     End If
     
     'VERIFICAR SI SIGUE TENIENDO SALDO LA BOBINA, EVITANDO UN ERROR SI LA MISMA FUE DEVUELTA.
     DEPOCORTE% = XVALO(Control("", "DEPODEST"))
     '
     PorcentajeTolerancia% = 2: EsEslinga% = 0
     If InStr(CODFAMI$, "EE") > 0 Then
         PorcentajeTolerancia% = 10 'si son eslingas PorcentajeTolerancia% = 10
         EsEslinga% = 1
     End If
     'VALIDA SI TIENE SALDO DEL LOTE PARA LA OPERACION ACTUAL OPERACION
     'LE AGREGO UNA TOLERANCIA DEL 2 % o del 10 si es eslinga SI VA A CONSUMIR 100 Y EL SALDO ES 98.5 PASA IGUAL SI ES 97 NO PASA.
     EntroALeerYaConsumido% = 0
     SQLX$ = "SELECT  SUM(CASE WHEN CANTINGRE > 0 THEN PESOKG  ELSE -PESOKG END) AS PESOKG"
     SQLX$ = SQLX$ + " ,SUM(CANTINGRE - CANTSALID) As SaldoMetros"
     If (XVALO(Mid$(NROOF$, 4, 6)) >= 2341 Or XVALO(Mid$(NROOF$, 4, 6)) = 2325) Then
        'ACA TRAE LO CONSUMIDO DE LA MISMA ORDEN Y OPERACION PARA DESCONTAR, POR QUE EL SALDO SI NO NO LO CONSIDERA
        SQLX$ = SQLX$ + " ,(SELECT SUM(CantSalid) FROM MOVISTO WHERE IDENLOTE = '" & TRIM$(TXTLOTE) & "'"
        SQLX$ = SQLX$ + " AND  DEPOMOVI = " & DEPOCORTE%
        If EsEslinga% < 1 Then SQLX$ = SQLX$ + " AND  ID_OPERACION = " & NOPE% 'si no es eslinga filtra operacion
        SQLX$ = SQLX$ + " AND  CODIMOVI = 'CF'"
        SQLX$ = SQLX$ + " AND DOPRILAR = '" & NROOF$ & "') AS YA_CONSUMIDO"
        EntroALeerYaConsumido% = 1
     End If
     SQLX$ = SQLX$ + " From Movisto WITH(NOLOCK)"
     SQLX$ = SQLX$ + " Where   IDENLOTE = '" & TRIM$(TXTLOTE) & "'"
     SQLX$ = SQLX$ + " AND  DEPOMOVI = " & DEPOCORTE%
     If EsEslinga% < 1 Then SQLX$ = SQLX$ + " AND  ID_OPERACION = " & NOPE%  'si no es eslinga filtra operacion
     SQLX$ = SQLX$ + "  AND DOSECLAR = '" & NROOF$ & "'"
     Set rs = READSET(SQLX$)
     If Not rs.EOF Then
        SALDO = XVALO(rs!SaldoMetros): MAXIMOTOLERANCIA# = (SALDO * PorcentajeTolerancia%) / 100
        TRANSFERIDO = SALDO 'LO GUARDO EN OTRA VARIABLE MOSTRARLO EN EL MENSAJE
        If EntroALeerYaConsumido% = 1 Then SALDO = SALDO - XVALO(rs!YA_CONSUMIDO)
        
        If SALDO + MAXIMOTOLERANCIA# < CantiEnMetros# Then
           If EntroALeerYaConsumido% > 0 Then
            Call COMUNI("La Bobina Seleccionada No Posee Saldo Suficiente en el Depósito de Corte\Para esta Operación \(Transferido Mts. " & ROUND(TRANSFERIDO, 2) & " \Consumido: " & ROUND(XVALO(rs!YA_CONSUMIDO), 2) & ")")
           Else
            Call COMUNI("La Bobina Seleccionada No Posee Saldo Suficiente en el Depósito de Corte\Para esta Operación \(Transferido Mts. " & ROUND(TRANSFERIDO, 2) & ")")
           End If
            GoTo 99
        End If
     
     End If
     On Error GoTo ERROR_GUARDAR
     FLEXCONN.BeginTrans
     Call ABREORFAB
     
     SQLX$ = " SELECT * FROM INFOFAB "
     SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "' "
     SQLX$ = SQLX$ + " AND NumeOper = " & NOPE% & " "
     '
     CAINFANT% = 0
     Dim INFOFABTMP As ADODB.Recordset
     Set INFOFABTMP = New ADODB.Recordset
     INFOFABTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
     With INFOFABTMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
           ' CAREAL# = CAREAL# + XVALO(!CANTREAL)
           ' TIEINF = TIEINF + XVALO(!TiemEjec)   ' deshabilitado porque si no acumulaba en el nuevo parte
           CAINFANT% = CAINFANT% + 1
         .MoveNext
         Loop
       End If
     End With
     CIFAB% = CODINT%(Text10)
     
    CODFAMI$ = TRIM$(CODFAMIL$(CIFAB%))
     
     
    IdMovistoGenerado = 0
    If (XVALO(Mid$(NROOF$, 4, 6)) >= 2341 Or XVALO(Mid$(NROOF$, 4, 6)) = 2325) And (InStr(CODFAMI$, "BB") > 0 Or InStr(CODFAMI$, "EE") Or InStr(CODFAMI$, "CC")) And InStr(EPAC$, "CONTEPLAST") > 0 Then  'ESTA ES A PARTIR DE LA CUAL SE VA CONSIDRERAR DIFERENTE LA DEVOLUCION DE LA BOBINA
        CMOV$ = "CF"
        REFE$ = "CONSUMO " & CODEXT$(CI%)
        DOPRI$ = NROOF$
        DOSEC$ = "PF-" + TRIM$(Mid$(NROOF$, 4))
        NC = XVALO(Text6)
        'CantiEnMetros# = ROUND(XVALO(Text7) * XVALO(TXTCORTE), 4) SE ESTA LEYENDO ARRIBA
        Call MOVISTOK(HOY(HO$), CIELEM%, TXTLOTE, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, 0, 1, NC, DEPOCORTE%, (-1) * CantiEnMetros#, , , , , , , , , , NOPE%, CantiEnMetros#, XVALO(LblKgImpresos))    'GUARDO EL NUMERO DE OPERACION DE CADA ITEMS TRANSFERIDO
        SQLX$ = "SELECT TOP 1 MOVISTO_ID FROM MOVISTO WITH(NOLOCK) ORDER BY MOVISTO_ID DESC"
        Set rs = READSET(SQLX$)
        If Not rs.EOF Then
          IdMovistoGenerado = XVALO(rs!MOVISTO_ID)
        End If
        rs.Close: Set rs = Nothing
    End If
'   AL GUARDAR FACTOR DE CONVERSION Y LA CANTIDAD REAL DEL BOLSON
     
     Call ABREORFAB
     SQLX$ = " SELECT * FROM InfoFab "
     Dim INFOFFFAB As ADODB.Recordset
     Set INFOFFFAB = New ADODB.Recordset
     INFOFFFAB.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
     With INFOFFFAB
        .AddNew   ' YA CONVERTIDO
        On Error Resume Next
        !CODIEMPR = CodiEmp%
        On Error GoTo 0
        !idsubor = NROOF$
        !DescriOF = Left$(ECOARCH$("ECORDEF", NROOF$), 64)
        !NuOrInf = CAINFANT%
        !cod_inte = CI%
        !Cod_Exte = CODEXT$(CI%)
        !Descrip = Left$(DESCRIT$(CI%), 64)
        !NUMEOPER = NOPE%
        !AltOper = 0
        !ReTraba = 0
        !descoper = TRIM$(Label7)
        !CanOrPro = CAPROG
        !MAQUINA = ""
        !DescMaq = ""
        !LegaOps = ""
        !FeHorLib = FEHOLI
        !FEHORINF = Now
        !Hora_Des = HORACENTECIIMAL
        !Hora_Has = HORACENTECIIMAL
        !MaqLiber = ""  ' LA MAQUINA ESTA LIBERADA PORQUE LA O/F ESTA CERRADA
        !TiemEjec = HORACENTECIIMAL
           MODO% = CAINFANT%
        !TiemStan = TIESTAN(CI%, NOPE%, CAINF#, MODO%)
        RENDI = 999.9: If !TiemStan > 0 Then RENDI = 100 * XVALO(!TiemEjec) / XVALO(!TiemStan)
        If RENDI > 999.9 Then RENDI = 999.9
        !RENDIM = RENDI
        !Observa = ""
        !Referen = " "
        !MotiRecha = 0
        !idenlote = TXTLOTE
        !codielem = CIELEM%
        !CodXElem = TRIM$(TXTCODELEM)
        FactCte = XVALO(TXTCANTICORTES)
        If FactCte < 1 Then FactCte = 1
        'se divide por la cantidad p.ej. un bolson usa 4 cortes de 80 cm
        'en la formula para ese componente es 3.20 m.
        'de este modo se guarda la cantidad 1 a 1 del producto fabricado 4 cortes equivalen a 1 bolson
        '!CANTREAL = Int(CAINF# / FactCte) 'lo cambie a solo la cantidad informada 04/05/20
        !CANTREAL = Int(CAINF#) 'cuando se consuma debera considerar el calculo para que coincida con formula
        !CantOrigCte = Int(CAINF#) 'SE GUARDA LA CANTIDAD ORIGINAL CUANDO SE GENERA EL INFORME
        !FACTORCORTE = FactCte 'REFIERE A LA CANTIDAD DE CORTES
        !NUMOPERARIO = XVALO(Text1(0))
        !NameOperario = Left$(Label5(0), 64)
        !corte = XVALO(TXTCORTE)
        !Peso = XVALO(LblKgImpresos)
        '03-08-21 SOLICITARON RAUL Y LESTER QUE POR DEFECTO LOS CORTES INGRESEN A SECTOR SEMITERMINADOS
        !COORDUBIC = "SEMITERMINADOS"  'SECTOR$
        !IDCOORDUBIC = 3 'IDSECTOR
        !IDMOVISTO = XVALO(IdMovistoGenerado)
       .Update
     End With
     Set INFOFFFAB = Nothing
     '
     
     SQLX$ = " SELECT * FROM INFOFAB "
     SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "' "
     SQLX$ = SQLX$ + " AND NumeOper = " & NOPE% & " "
     '
     CAREAL# = 0: TIEINF = 0
     Dim INFOFABTMP1 As ADODB.Recordset
     Set INFOFABTMP1 = New ADODB.Recordset
     INFOFABTMP1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    
     With INFOFABTMP1
         On Error Resume Next
         .MoveFirst
         If Err < 1 Then
           On Error GoTo 0
           Do While Not .EOF
             'If IsNull(!ReTraba) Or !ReTraba < 1 Then
             If XVALO(!ReTraba) < 1 Then
               CAREAL# = CAREAL# + XVALO(!CANTREAL)
             End If
             TIEINF = TIEINF + XVALO(!TiemEjec)
           .MoveNext
           Loop
         End If
         On Error GoTo 0
     End With
    '
    SQLZ$ = " SELECT * FROM OPERFAB "
    SQLZ$ = SQLZ$ + " WHERE IdSubOr = '" & NROOF$ & "' AND NumeOper = " & NOPE%
    Dim OPERBU As ADODB.Recordset
    Set OPERBU = New ADODB.Recordset
    OPERBU.Open FILTSQL$(SQLZ$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

    With OPERBU
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
        On Error GoTo 0
       !CANTREAL = CAREAL#
       !CANTPEND = XVALO(!CANTPROG) - CAREAL#
         If XVALO(!CANTPEND) < 0 Or STATUOP% = 3 Or STATUOP% = 8 Then !CANTPEND = 0
         CANTPE# = XVALO(!CANTPEND)
         MODO% = CAINFANT%
       !TiemEjec = TIEINF
       !TiemPend = TIEMAQUI(CI%, NOPE%, CANTPE#, MODO%)  ' 1 es para que no considere preparacion
          If !TiemPend < 0 Or STATUOP% = 3 Or STATUOP% = 0 Then !TiemPend = 0
       !StatOper = STATUOP%
       .Update
     End If
     On Error GoTo 0
   End With

'   If TRIM$(NROOF$) <> "" Then
'       If CAINF# >= 0.05 Then
'         If NOPE% = ULTIMOP%(CI%) Then
'           TTXX$ = "La Operación Informada es la Última\de la Secuencia Standard de '" + TRIM$(CODEXT$(CI%)) + "'."
'           TTXX$ = TTXX$ + "\   \¿ Desea Informar el Cierre Parcial Correspondiente ?"
'           If COMALTER%(TTXX$ + "\\Si, Cierre Parcial\No, Continuar") = 1 Then
'             NROOF7$ = NROOF$
'             CAINF7 = CAINF#
'             FECHINFO$ = Text4
'             Call ORFCIE(NROOF7$, CAINF7, FECHINFO$)
'           End If
'         End If
'       End If
'   End If
   '
   
ERROR_GUARDAR:
    If Err <> 0 Then
       On Error Resume Next
       On Error GoTo 0
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION")
       FLEXCONN.RollbackTrans
       Exit Sub
    Else
      FLEXCONN.CommitTrans
      Call COMUNI("Proceso Realizado Exitosamente")
       
      INFOFABTMP.Close
      Set INFOFABTMP = Nothing

      INFOFABTMP1.Close
      Set INFOFABTMP1 = Nothing

      OPERBU.Close
      Set OPERBU = Nothing

    End If
    
'    OPX% = COMALTER%("Desea Imprmir Etiqueta Para Pallets de Corte\\No, No Imprimir\Sí, Imprimir")
'    If OPX% < 2 Then GoTo 99

'     IMAG$ = "QR_ETIQPROCESO.JPG"
'     IDENTIFICADOR$ = "QR_ETIQPROCESO"
'        '
'     On Error Resume Next
'     Kill "C:\FLEXOFT\" & IMAG$
'     NC = XVALO(Text6)
'     CANTI = XVALO(Text7)
'     NOP% = XVALO(Text3)
'     NROPED& = XVALO(Text12)
'     CORTE = XVALO(TXTCORTE)
'     Codigo$ = TXTCODELEM
'     CIELEM% = CODINT(TXTCODELEM)
'     LOTE$ = TRIM$(TXTLOTE)
'
'     TEXTOQR$ = CODEXT$(CIELEM%) & "|" & LOTE$ & "|" & CANTI & "|" & NOP% & "|" & NC & "|" & NROPED& & "|" & NROOF$ & "|" & ORDENINFO & "|" & UBICACION$ & "|" & IDUBICACION
'
'     On Error GoTo 0
'     'RUTINA QUE LLAMA A APLICACION DE .NET PARA GENERAR LA IMAGEN DEL CODIGO QR
'     QRX$ = TRIM$(IMAGENCODBARRA$("C:\FLEXOFT\", IMAG$, TEXTOQR$, "QR_CODE", ANCHOX%, ALTOX%))
'     Call T_Espera(3)
'     Image1.Picture = LoadPicture("C:\FLEXOFT\" & IMAG$)
'     Call GRABAR_QR("C:\FLEXOFT\" & IMAG$, IDENTIFICADOR$)
''     FILTX$ = TRIM$(COD$) & ";" & "Rev." & REVI$ & ";" & TRIM$(LOTE$) & ";" & IDENTIFICADOR$ & ";" & TRIM$(str$(User%))
''     Call STDFORM(FORIMPRE$, FILTX$, "PRINT", "", 0, CANTI&)
'
'     FILTX$ = TRIM$(str$(XVALO(Text6))) & ";" & NROOF$ & ";" & TRIM$(str$(XVALO(Text7))) & ";" & TRIM$(Label7) & ";" & TRIM$(str$(XVALO(TXTCORTE)))
'     Call STDFORM("PRPR-CTP", FILTX$)


     Call CargaMatTranfXOrdenSoloIngresos
     
     Dim CTP As New Conteplast
     CIFAB% = CODINT%(Text10)
     Call CTP.CargaOperacionesInformadasMasSubProcesos(NROOF$, CIFAB%, MSF2, 0, "", 99)
     Set CTP = Nothing

     '++++++++++++++++++++++++++++++++++++++++++++++++++++
     'BUSCO EL REGISTRO QUE SE CREO segun CAINFANT%
     
     For i& = 1 To MSF2.Rows - 1
        If XVALO(MSF2.TextMatrix(i&, 11)) = CAINFANT% Then
          If MSF2.TextMatrix(i&, 9) = TRIM$(TXTCODELEM) Then
            If TRIM$(MSF2.TextMatrix(i&, 7)) = TRIM$(Label7) Then
               Reg& = i&
               Exit For
            End If
          End If
        End If
     Next i&
     
     UBICACION$ = TRIM$(MSF2.TextMatrix(Reg&, 10))
     OPX% = COMALTER%("Desea Imprimir Etiqueta de Pallet (Ubicación: " & UBICACION$ & ")\\Cancelar\Imprimir")
     If OPX% < 2 Then GoTo 98

     CIELEM% = CODINT(MSF2.TextMatrix(Reg&, 9)) 'material
     LOTE$ = TRIM$(MSF2.TextMatrix(Reg&, 3))
     CANTI = XVALO(MSF2.TextMatrix(Reg&, 4)) 'cantidad de cortes)
     NOP% = XVALO(MSF2.TextMatrix(Reg&, 6))
    
     NC = XVALO(Text6)
     NROPED& = XVALO(Text12)
     NROOF$ = Label4
     ORDENINFO = XVALO(MSF2.TextMatrix(Reg&, 11))
     IDUBICACION = XVALO(MSF2.TextMatrix(Reg&, 12))
     DESCRIOPERACION$ = MSF2.TextMatrix(Reg&, 7)
     CODIGOINTFAB = CODINT(Text10)
     IDINFOFAB& = XVALO(MSF2.TextMatrix(Reg&, 16))
     Call PRINTETIQPALLET(CODEXT$(CIELEM%), LOTE$, CANTI, NOP%, NC, NROPED&, NROOF$, ORDENINFO, UBICACION$, IDUBICACION, DESCRIOPERACION$, CODIGOINTFAB, IDINFOFAB&)
     '+++++++++++++++++++++++++++++++++++++++++++++++


98   Call BLANFORMU
   

99 Command2.Enabled = True
End Sub

Private Sub Command3_Click()
    '
    Label4 = OrdenFabricacionAbiertas$
    
    Exit Sub
    CIXI% = CODINT%(TRIM$(Text10))
    If CIXI% <= 0 Then
        Call MENSERR(24, "Primero debe Seleccionar un Articulo")
        Exit Sub
    End If
    '
    CIRANG = CODIGORANGO&(CIXI%)
    CONDI$ = "COD_INTE = " & CIXI%
    CONDI$ = CONDI$ + " AND STATUORF < 3"
    Reg& = CantRegistros&("ORDEFABR", CONDI$)
    If Reg& < 1 Then
        Call MENSERR(24, "No Existen Ordenes de Fabricación\Para este Articulo")
        Exit Sub
    End If
    '
    Call SELECHO("!ECORDEF1/Ordenes de Fabricación")
    If VALACT1$("!ECORDEF1") <> "" Then
     Label4 = VALACT1$("!ECORDEF1")
     Label2 = VALACT2$("!ECORDEF1")
    End If
    '
End Sub

Private Sub Command4_Click(Index As Integer)
  '
  Call SELECHO("PADRON")
  VDEV$ = VALACT1$("PADRON")
  If VDEV$ <> "" Then
    Text1(Index) = TRIM$(VDEV$)
    Label5(Index) = TRIM$(VALACT2$("PADRON"))
  End If
  
End Sub '

Private Sub Command5_Click()
  Call SELECHO("SELFECHA")
  VDEV$ = VALACT1$("SELFECHA")
  If VDEV$ <> "" Then
    Text4.TEXT = VDEV$
  End If
End Sub

Private Sub Command6_Click()
    '
    Call COMUNI("Atención!!, La opción de Seleccionar la Operación en esta Instancia\Está Deshabilitada\Si la Informada al Ingreso de la Bobina No es la correcta\Retornela y vuelva a Ingresar Correctamente.")
    Exit Sub
    
    CIXI% = CODINT%(TRIM$(Text10))
    If CIXI% <= 0 Then
        Call MENSERR(24, "Primero debe Seleccionar un Articulo")
        Exit Sub
        If TRIM$(Label4) = "" Then
            Call MENSERR(24, "Primero debe Seleccionar la Orden de Fabricacion")
            Exit Sub
        End If
    End If
    '
    Call COPYSTRU("OPERINF", "OPERINF1")
    Condi1$ = "CodIConj = " & CIXI%
    Condi1$ = Condi1$ + " ORDER BY NUMEOPER"
    Call CARGALISEL("!OPERINF1", "SECOPER", "NumeOper/F2;DescOper/A32", Condi1$)
    
    CONDI$ = "CodIConj = " & CIXI%
    Reg& = CantRegistros&("SECOPER", CONDI$)
    If Reg& < 1 Then
        Call MENSERR(24, "No Existen Operaciones de Fabricación\Para este Articulo")
        Exit Sub
    End If
    '
    
    Call SELECHO("!OPERINF1/Operaciones de Ordenes de Fabricacion")
    If TRIM$(VALACT1$("!OPERINF1")) <> "" Then
      Text3.TEXT = TRIM$(VALACT1$("!OPERINF1"))
'      Label7.Caption = TRIM$(VALACT2$("!OPERINF1"))
    End If
    '
End Sub

Private Sub Command7_Click()
   '
   Call SELECHO("PADMAQ")
   VDEVU$ = VALACT1$("PADMAQ")
   If VDEVU$ <> "" Then
      Text5.TEXT = TRIM$(VDEVU$)
      Label10.Caption = TRIM$(VALACT2$("PADMAQ"))
   End If
   '
End Sub

Private Sub Command8_Click()
    Call CIERRARCH("*.*")
    Unload Me
End Sub



Private Sub Command9_Click()
    Command9.Enabled = False
    Codigo$ = Label31(0)
    CI1% = CODINT%(Codigo$)
    If CI1% < 1 Then
      Call COMUNI("Código No Válido")
      GoTo 99
    End If
    
    Dim CTP As New Conteplast
    PESOX# = CTP.PesoLeidoBalanza(CI1%, 2) '
    If PESOX# < 0 Then PESOX# = 0
    KgImpresos = TRIM$(FORMATNUM$(PESOX#, "f12.2"))

'    KgImpresos = FORMATNUM$(CTP.PesoLeidoBalanza(CI1%), "f12.2")
99  Command9.Enabled = True
End Sub

Private Sub Form_Load()
    '
'    Me.Height = 7380
    Call CARGA_ENCABEZADO(MSF1, "Código/A16;Descripción/A28;Lote/A16;Cons.Mts./a0;CI/a0;ID_OPERACION/A7", Me)
'    Call CARGA_ENCABEZADO(MSF2, "Código/A16;Descripción/A28;Lote/A8;Cant.Cortes/F10.1;CI/a0;N°Op./F5;Descrip/A12;Hora/A0;Material/A16;Ubicación/A12;NroInf/a0;IDUbic/a0;N°Operario/F9;Nombre/A16;Peso/f8.2;IdReg/a0", Me)
    Call CARGA_ENCABEZADO(MSF2, "Código/A0;Descripción/A0;Lote/A8;Cant.Cortes/F8.1;CI/a0;N°Op./F5;Descrip/A12;Hora/A0;Material/A16;Ubicación/A12;NroInf/a0;IDUbic/a0;N°Operario/F9;Nombre/A10;Peso/f7.2;IdReg/a0;IDREGPADRE/A0;OPERFUCIONADA/A0;IDMOVISTO/A0", Me)
    Text4 = FECHATEX$(HOY(HO$))
    Call CREATABLA_UBILOTE_ENVAP
End Sub
Public Sub Centrar(ByVal poForm As Form)
    With poForm
        .Move (Screen.Width - .Width) \ 2, (Screen.Height - .Height) \ 2
     End With
End Sub
Sub CargaMatTranfXOrdenSoloIngresos()
  MSF1.Rows = 1: Reg& = 0
  DEPOCORTE% = XVALO(Control("", "DEPODEST"))
  'aca deberia ser el resultado de lo transferido codimovi = 'tr' ingreso - salida entre los depositos 1 y deposito de corte.
  SQLX$ = "SELECT sum(CANTINGRE) AS RESULTADO ,COD_INTE,COD_EXTE,DESCRIP, IDENLOTE,ID_OPERACION "
  SQLX$ = SQLX$ + " FROM MOVISTO WITH(NOLOCK)"
  SQLX$ = SQLX$ + " where doseclar = '" & TRIM(Label4) & "'  AND DEPOMOVI = " & DEPOCORTE%
  SQLX$ = SQLX$ + " GROUP BY COD_INTE,COD_EXTE,DESCRIP, IDENLOTE,ID_OPERACION"
  SQLX$ = SQLX$ + " ORDER BY IDENLOTE "
  Set rs = READSET(SQLX$)
  With rs
    Do While Not .EOF
       Reg& = Reg& + 1
       MSF1.Rows = Reg& + 1
       MSF1.TextMatrix(Reg&, 1) = SAFETEXT$(!Cod_Exte)
       MSF1.TextMatrix(Reg&, 2) = SAFETEXT$(!Descrip)
       MSF1.TextMatrix(Reg&, 3) = SAFETEXT$(!idenlote)
       MSF1.TextMatrix(Reg&, 4) = FORMATNUM$(XVALO(!Resultado), "F10.1")
       MSF1.TextMatrix(Reg&, 5) = XVALO(!cod_inte)
       MSF1.TextMatrix(Reg&, 6) = XVALO(!ID_OPERACION)
       .MoveNext
    Loop
 End With
 rs.Close
 Set rs = Nothing
 
 '********************************************************************
 'VER ACA SI ESTA BIEN QUE MUESTRE SOLO CANTINGRE EN VEZ DE CANTINGRE - CANTSALID
  '********************************************************************

End Sub

Private Sub Label4_Change()
   MSF1.Rows = 1
   MSF2.Rows = 2
'   Call BLANFORMUGENERAL

   'CIERRA GRILLA DE DETEALLE DE INFOFAB
   MSF3.Visible = False 'CON ESTO OBLIGA A QUE SI ABRE SE ACTUALIZA LA INFORMACION
   MSF3.Height = 800
   MSF3.Width = 800
   Command15.Caption = "!"
   'FIN CIERRA GRILLA DE DETEALLE DE INFOFAB
   
   Call BLANFORMU
   Call BLANFORMUIMPRESION
   Call BLANFORMU_FUCION
   Frame4.Visible = False
   Frame2(1).Visible = True
   
'   Me.Height = 7380
   NORFA$ = TRIM$(Label4)
   If NORFA$ = "" Then MSF1.Rows = 1: Exit Sub
   Label2 = RESP_ZELE_VECT(2)
   CI1% = CODIFAB%(NORFA$)
   TXTCODINT = CI1%
   Text10 = CODEXT$(CI1%)
   Label3 = DESCRIT0$(CI1%)
   Call CargaMatTranfXOrdenSoloIngresos
        
        
     
   CIFAB% = CODINT%(Text10)
   Dim CTP As New Conteplast
   Call CTP.CargaOperacionesInformadasMasSubProcesos(NORFA$, CIFAB%, MSF2, 0, "", 99)
   Set CTP = Nothing
   
'   Call CargaOperacionesInformadasMasSubProcesos(NORFA$)
   NROPED& = NroPedidoSegunOf&(NORFA$)
   NC = cliente_SEGUN_PEDIDO(NROPED&)
   Text6 = NC
   Text12 = NROPED&
   Text14 = CantiProyectadaOF(NORFA$)
   CONDI$ = "NROPED = " & NROPED& & " AND  CODIINT = " & CIFAB%
   TEXT13 = XVALO(VALOBADA("PEDDETA", "SUM(CanPed)", CONDI$, "NOMESS"))
   
End Sub



Private Sub mnuDepoCinESl_Click()
    ConfigDepo.Show 1
End Sub

Private Sub mnuInfOpEslinlgas_Click()
    INGESLINGA.Show 1
    Call Label4_Change
End Sub

Private Sub mnuIngCintillas_Click()
    INGCINTILLA.Show 1
    Call Label4_Change
End Sub

Private Sub mnuOrigScrap_Click()
    FRMSCRAP.Show 1
    
End Sub

Private Sub mnuTabAux_Click()

' If DERACCE%("TABAUXGE", "ACCESO A TABLAS AUXILIARES GENERALES") < 2 Then Exit Sub
' Call MOSTRARTABLASCONFIGURADAS(Me.Name)
  Call GENERAR_ESTRUCTURAS_SQL
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  FRMZELECHO.Caption = "LISTA DE ESTRUCTURAS PRF Y RFS"

  Dim obj As New RECORXET
  Set rs1 = obj.RS_ENCAPRFLISTAS()
  Call Carga_MSF_Recordset(rs1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO.MSF2)
  Set obj = Nothing
  rs1.Close
  Set rs1 = Nothing
  EPAC$ = UCase$(TRIM$(EMPREAC$))
  '

  FRMZELECHO.Width = 8500
  FRMZELECHO.Show 1
  IDLIST& = XVALO(RESP_ZELE_VECT(1))
  If InStr(EPAC$, "ENVAPLAST") > 0 Then
    LIXTA$ = TRIM$(RESP_ZELE_VECT(2))
    If LIXTA$ = "MATERIALES ENVAPLAST" Then
       Call COMUNI("Atención !!! Editar Tabla de Familias Desde Maestro de Artículos")
       Exit Sub
    End If
  End If

  Call ABM_TABLA(IDLIST&)

End Sub

Private Sub mnuTablaConfScrap_Click()
        Call mostrarTablaScrap
End Sub


Private Sub mnuUbic_Click()
   INFOUBIC.NOUBICAR = 1
   INFOUBIC.Show 1
   '
   
   CIFAB% = CODINT%(Text10)
   If CIFAB% < 1 Then Exit Sub
   Dim CTP As New Conteplast
   Call CTP.CargaOperacionesInformadasMasSubProcesos(Label4, CIFAB%, MSF2, 0, "", 99)
   Set CTP = Nothing

End Sub

Private Sub MSF1_Click()
   Reg& = MSF1.MouseRow
   J& = MSF1.MouseCol
   
   If Reg& = 0 Then
     Static MODODEORDEN$
     If MODODEORDEN$ = "DESC" Then
        MODODEORDEN$ = ""
     Else
        MODOORD% = 0
        MODODEORDEN$ = "DESC"
     End If
     Call ORDENAR_MSF(INFOCORTE.MSF1, J&, MODODEORDEN$)
     Exit Sub
   End If
   
   If Reg& >= MSF1.Rows Then Exit Sub
   
   Call SELECCIONARFILA(MSF1, Reg&)
   Me.LblKgImpresos = ""
   Me.Text7 = ""
   NOP% = XVALO(MSF1.TextMatrix(Reg&, 6))
   Text3 = NOP%
   TXTLOTE = MSF1.TextMatrix(Reg&, 3)
   
   TXTCODELEM = MSF1.TextMatrix(Reg&, 1)
   CIELEM% = MSF1.TextMatrix(Reg&, 5)
   CODIGOFAB = CODINT%(Text10)
   
   Condi1$ = "CodIConj = " & CODIGOFAB & " and  NumeOper = " & NOP%
   Label7 = VALOBADA("SECOPER", "DescOper", Condi1$)
   
   CONDI$ = "CODICONJ = " & CODIGOFAB
   CONDI$ = CONDI$ + " AND (CODIELEM = " & CIELEM% & "  or CODIELEM = " & CODIGORANGO&(CIELEM%) & ")"
   DESOP$ = TRIM$(Label7)
   
   If DESOP$ <> "" Then
    CONDI$ = CONDI$ + " AND COMPONENTE = '" & DESOP$ & "'"
   End If
   'LO PASO A MTS 'EL CORTE EN LA TABLA FORMULAS ESTA EN CM POR QUE ASI LO CARGAN EN LA COTIZACION
   TXTCORTE = XVALO(VALOBADA("FORMULAS", "CORTE", CONDI$, "NOMESS")) / 100
   Me.TXTCANTICORTES = XVALO(VALOBADA("FORMULAS", "MODUXBB", CONDI$, "NOMESS"))

   
   NORFA$ = TRIM$(Label4)
   NOPER% = XVALO(Text3)
   LOTE$ = TRIM$(TXTLOTE)
   'CANTIDAD DE CORTES DE LA OF X LOTE Y OPERACION
   CONDI$ = "IDSUBOR = '" & NORFA$ & "' AND IDENLOTE = '" & LOTE$ & "' AND NUMEOPER = " & NOPER%
   IDREGSELEC& = XVALO(VALOBADA("INFOFAB", "INFOFAB_ID", CONDI$, "NOMESS"))
   CONDI2$ = "IdRegPadre = " & IDREGSELEC&
   TieneHerencia& = XVALO(VALOBADA("InfoFab", "IdRegPadre", CONDI2$, "NOMESS"))
   If TieneHerencia& > 0 Then '14/07/21
      'cuando se realiza subprocesos se pone cantreal en 0 y se aplica el valor al nuevo registro generado
      Me.TXTCORTESOPYLOTE = XVALO(VALOBADA("InfoFab", "sum(CantOrigCte)", CONDI$, "NOMESS"))
      CONDI$ = "IDSUBOR = '" & NORFA$ & "' AND NUMEOPER = " & NOPER%
      Me.TXTCORTESOPERACION = XVALO(VALOBADA("INFOFAB", "Sum(CantOrigCte)", CONDI$, "NOMESS"))
   ElseIf TRIM$(UCase$(CODFAMIL$(CIELEM%))) = "EE" Then ' ESLINGAS
      Me.TXTCORTESOPYLOTE = XVALO(VALOBADA("InfoFab", "sum(CANTREAL)", CONDI$, "NOMESS"))
      CONDI$ = "IDSUBOR = '" & NORFA$ & "' AND NUMEOPER = " & NOPER%
      Me.TXTCORTESOPERACION = XVALO(VALOBADA("INFOFAB", "Sum(CANTREAL)", CONDI$, "NOMESS"))
   Else
      Me.TXTCORTESOPYLOTE = XVALO(VALOBADA("INFOFAB", "Sum(CANTREAL)", CONDI$, "NOMESS"))
      CONDI$ = "IDSUBOR = '" & NORFA$ & "' AND NUMEOPER = " & NOPER%
      Me.TXTCORTESOPERACION = XVALO(VALOBADA("INFOFAB", "Sum(CANTREAL)", CONDI$, "NOMESS"))
'21/07/21 se corrigio a esta linea queda igual que if tieneherencia pero lo dejo separado para ver si va todo bien
      
   End If
   
   'CANTIDAD DE CORTES DE LA OF X  OPERACION
   Me.TXTNECESIDADCORTES = FORMATNUM$(XVALO(TXTCANTICORTES) * XVALO(Text14), "F8")
   
   TXTRESTAN = FORMATNUM$(XVALO(TXTNECESIDADCORTES) - XVALO(TXTCORTESOPERACION), "F8")
End Sub

Private Sub MSF1_DblClick()
   EPAC$ = TRIM$(UCase$(EMPREAC$))
    
   If InStr(EPAC$, "CONTEPLAST") < 1 Then Exit Sub

    Reg& = MSF1.MouseRow
    COL& = MSF1.MouseCol
    If Reg& = 0 Or Reg& >= MSF1.Rows Then Exit Sub
    If COL& < 1 Then
        IDLOT$ = TRIM$(MSF1.TextMatrix(Reg&, 3))
        If IDLOT$ <> "" Then
          On Error Resume Next
          CONSUMLOT07.Label4 = TRIM$(MSF1.TextMatrix(Reg&, 1))
          If Err Then Exit Sub
          On Error GoTo 0
          CONSUMLOT07.Label1 = IDLOT$
          CONSUMLOT07.Show 1
        End If
   Else
         'por ahora solo eslingas , puede servir para otros pero para manejarlo bien.
         Codigo$ = TXTCODELEM
         CI1% = CODINT(Codigo$)
         CODFAMI$ = TRIM$(CODFAMIL$(CI1%))
         If InStr(CODFAMI$, "EE") < 1 Then Exit Sub

         USER1% = USNBR% Mod 100
         If USER1% <> 1 Then
            Call COMUNI("Solo Puede Ingresar el Supervisor del Sistema")
            Exit Sub
         End If

         NORFA$ = TRIM$(Label4)
         LOTE$ = TXTLOTE
         CantCortes = XVALO(TXTCORTESOPYLOTE)
         MedidaDelCorte = XVALO(TXTCORTE)
         NOPE% = XVALO(Text3)
         
         Call GenerarConsumoEslingasCintillas(NORFA$, Codigo$, LOTE$, CantCortes, MedidaDelCorte, NOPE%)
   End If
End Sub


Private Sub MSF2_Click()
    Reg& = MSF2.MouseRow
    COL% = MSF2.MouseCol
    
    If Reg& = 0 Then
     Static MODODEORDEN$
     If MODODEORDEN$ = "DESC" Then
        MODODEORDEN$ = ""
     Else
        MODOORD% = 0
        MODODEORDEN$ = "DESC"
     End If
     Call ORDENAR_MSF(INFOCORTE.MSF2, COL%, MODODEORDEN$)
     Exit Sub
    End If
    
    
    If Reg& = 0 Or Reg& >= MSF2.Rows Then Exit Sub
    If COL% > 0 Then Exit Sub
    Call Asterizco_En_Celda(MSF2, Reg&, 0, vbYellow)
    Call VALIDARBOTONFUSION
    If Me.CmdInfoFusion.Visible = False Then Exit Sub
    
    
'    REG& = MSF2.MouseRow
'    COL& = MSF2.MouseCol
'    If REG& = 0 Or REG& >= MSF2.Rows Then Exit Sub
''    Call CARGA_ENCABEZADO(MSF2, "1 Código/A16;22Descripción/A28;33LOTE/A8;44Cant.Ctes./F10.1
''    ;55 CI/a0;N°Op./F5;66 Descrip/A12;77 Hora/A12;88 Material/A16;99 Ubicación/A10", Me)
'    OPX% = COMALTER%("Puede Informar Una Ubicación en Sector o En Rack\\Cancelar\Ubicación en Sector\Ubicación en Rack\Imprmir Etiqueta Para Pallet\Informar Impresión")
'    If OPX% < 2 Then Exit Sub
'    If OPX% = 2 Then
'        NORFA$ = Label4 'o.f
'        NOP% = XVALO(MSF2.TextMatrix(REG&, 6)) 'numeoper
'        CODIEXLM$ = MSF2.TextMatrix(REG&, 9) 'material
'        LOTE$ = MSF2.TextMatrix(REG&, 3) 'lote
'        CANTI = MSF2.TextMatrix(REG&, 4) 'cantidad de cortes
'        NUORD% = XVALO(MSF2.TextMatrix(REG&, 11))
'       '**** CAMBIOS EN UBICACIONES 21/04/20
''        If XVALO(MSF2.TextMatrix(REG&, 10)) > 0 Then 'SI TIENE UBICACION EN RACK COMIENZA CON NUMERO
''            OPX% = COMALTER%("Atención !! Deberá Limpiar la Ubicación en Rack Antes de Asignar una Nueva Ubicación en Sector\\Cancelar\Ver Lista Para Borrar")
''            If OPX% < 2 Then Exit Sub
''            On Error Resume Next
''            TEXTOX$ = MSF2.TextMatrix(REG&, 10)
''            strArray = Split(TEXTOX$, ";")
''            RK = XVALO(strArray(0))
''            FL = XVALO(strArray(1))
''            KL = XVALO(strArray(2))
''            On Error GoTo 0
'''            NUID& = INFOUBIC.BORRARUBICACIONCONZELECHO(RK, FL, KL)
''            If NUID& > 0 Then
''              Call BORRAREGISTROS("UBILOTE_ENVAP", "UBILOTE_ENVAP_ID = " & NUID&, REGAF&, "NOMESS")
''              Condi$ = "IDSUBOR = '" & NORFA$ & "'"
''              Condi$ = Condi$ + " AND NUMEOPER = " & NOP%
''              Condi$ = Condi$ + " AND  NUORINF = " & NUORD%
''              Call ACTUREGISTRO("INFOFAB", "COORDUBIC", Condi$, "", REGAF&, "NOMESS")
''              Call CargaOperacionesInformadasMasSubProcesos(NORFA$)
''            End If
''            Exit Sub
''        End If
'
'        'SE LLAMA CON CONDICION PARA QUE MUESTRE LOS < 100 DE LOS SECTORES PARA ENCUADRAR DEPENDIENDO DE DONDE SE LLAME.
'        x = LISTASELECGENERAL("SECTPLANTA", "", "", 0, "CONVERT(int, campo1) < 100")
'        IDSECTOR = XVALO(RESP_ZELE_VECT(1))
'        SECTOR$ = TRIM$(RESP_ZELE_VECT(2))
'        If SECTOR$ = "" Then Exit Sub
'        Condi$ = "IDSUBOR = '" & NORFA$ & "'"
'        Condi$ = Condi$ + " AND NUMEOPER = " & NOP%
'        Condi$ = Condi$ + " AND CANTREAL = " & CANTI
'        Condi$ = Condi$ + " AND IDENLOTE = '" & LOTE$ & "'"
'        Condi$ = Condi$ + " AND  CODIELEM = " & CODINT%(CODIEXLM$)
'        Condi$ = Condi$ + " AND  NUORINF = " & NUORD%
'        Call ACTUREGISTRO("INFOFAB", "COORDUBIC", Condi$, SECTOR$, REGAF&, "NOMESS")
'        Call ACTUREGISTRO("INFOFAB", "IDCOORDUBIC", Condi$, IDSECTOR, REGAF&, "NOMESS")
'    End If
'
'    If OPX% = 3 Then
'        INFOUBIC.Label4 = Label4 'o.f
'        INFOUBIC.Text3 = XVALO(MSF2.TextMatrix(REG&, 6)) 'numeoper
'        INFOUBIC.TXTCODELEM = MSF2.TextMatrix(REG&, 9) 'material
'        INFOUBIC.TXTLOTE = MSF2.TextMatrix(REG&, 3) 'lote
'        INFOUBIC.TXTCANTICORTES = MSF2.TextMatrix(REG&, 4) 'cantidad de cortes
'        INFOUBIC.LBLNUMEORDENINFO = XVALO(MSF2.TextMatrix(REG&, 11))
'        INFOUBIC.Show 1
'    End If
'
'    CIFAB% = CODINT%(Text10)
'    Dim CTP As New Conteplast
'    Call CTP.CargaOperacionesInformadasMasSubProcesos(Label4, CIFAB%, MSF2, 0, "", 99)
'    Set CTP = Nothing
'
'    'Impresion
'    If OPX% = 4 Or OPX% = 5 Then
'        CIELEM% = CODINT(MSF2.TextMatrix(REG&, 9)) 'material
'        LOTE$ = TRIM$(MSF2.TextMatrix(REG&, 3))
'        CANTI = XVALO(MSF2.TextMatrix(REG&, 4)) 'cantidad de cortes)
'        NOP% = XVALO(MSF2.TextMatrix(REG&, 6))
'        NC = XVALO(Text6)
'        NROPED& = XVALO(Text12)
'        NROOF$ = Label4
'        ORDENINFO = XVALO(MSF2.TextMatrix(REG&, 11))
'        UBICACION$ = TRIM$(MSF2.TextMatrix(REG&, 10))
'        IDUBICACION = XVALO(MSF2.TextMatrix(REG&, 12))
'        DESCRIOPERACION$ = MSF2.TextMatrix(REG&, 7)
'        CODIGOINTFAB = CODINT(Text10)
'       If OPX% = 4 Then
'        Call PRINTETIQPALLET(CODEXT$(CIELEM%), LOTE$, CANTI, NOP%, NC, NROPED&, NROOF$, ORDENINFO, UBICACION$, IDUBICACION, DESCRIOPERACION$, CODIGOINTFAB)
'       End If
'       If OPX% = 5 Then
'          Me.Height = 9510
'          Call Centrar(Me)
'          'SI TIENE UN VALOR EN IdRegPadre ES POR QUE ES UN INFORME QUE TIENE HERENCIA DE UN CORTE
'          'ES DECIR YA SE GENERO UN INFORM DE ESTE, POR ENDE ES UN INFORME.
'          LblId = XVALO(MSF2.TextMatrix(REG&, 16))
'          Condi$ = " INFOFAB_ID = " & XVALO(LblId)
'          TieneHerencia& = XVALO(VALOBADA("InfoFab", "IdRegPadre", Condi$, "NOMESS"))
'          If TieneHerencia& > 0 Then
'             Call COMUNI("Atención !! El Registro Seleccionado Es Informe de Otro\Solo Puede Realizar Informes de Sub-Proceso de Cortes Primarios")
'             Exit Sub
'          End If
'
'          Label25 = CTP.CantAcumuladaSubProceso&(LblId)
'
'          TXTLECTURA = CODEXT$(CIELEM%) & "|" & LOTE$ & "|" & CANTI & "|" & NOP% & "|" & NC & "|" & NROPED& & "|" & NROOF$ & "|" & ORDENINFO & "|" & UBICACION$ & "|" & IDUBICACION
'       End If
'     End If
End Sub
Sub PRINTETIQPALLET(codielem$, LOTE$, CANTI, NOP%, NC, NROPED&, NROOF$, ORDENINFO, UBICACION$, IDUBICACION, DESCRIOPERACION$, CODIGOINTFAB, IDINFOFAB&)
     If IDINFOFAB& < 1 Then
        CIELEM = CODINT(codielem$)
        CONDI$ = "CODICONJ = " & CODIGOINTFAB
        CONDI$ = CONDI$ + " AND (CODIELEM = " & CIELEM & "  or CODIELEM = " & CODIGORANGO&(CIELEM) & ")"
        corte = XVALO(VALOBADA("FORMULAS", "CORTE", CONDI$, "NOMESS"))
     Else 'SI TIENE ID
         CONDI$ = "INFOFAB_ID = " & IDINFOFAB&
         corte = XVALO(VALOBADA("INFOFAB", "CORTE", CONDI$, "NOMESS"))
         PESOX$ = FORMATNUM$(XVALO(VALOBADA("INFOFAB", "PESO", CONDI$, "NOMESS")), "F12.2")
     End If
     IMAG$ = "QR_ETIQPROCESO.JPG"
     IDENTIFICADOR$ = "QR_ETIQPROCESO"
     Dim CTP As New Conteplast
     DESCRIUBICACION$ = CTP.DescripcionyCoordenadasUbicacion$(UBICACION$)
     DESCRIUBICACION$ = REPLACAR(DESCRIUBICACION$, ";", "-")
     On Error Resume Next
     Kill "C:\FLEXOFT\" & IMAG$
     UBICACION$ = REPLACAR(UBICACION$, ";", "-")
     TEXTOQR$ = codielem$ & "|" & LOTE$ & "|" & CANTI & "|" & NOP% & "|" & NC & "|" & NROPED& & "|" & NROOF$ & "|" & ORDENINFO & "|" & UBICACION$ & "|" & IDUBICACION

     On Error GoTo 0
     'RUTINA QUE LLAMA A APLICACION DE .NET PARA GENERAR LA IMAGEN DEL CODIGO QR
     QRX$ = TRIM$(IMAGENCODBARRA$("C:\FLEXOFT\", IMAG$, TEXTOQR$, "QR_CODE", ANCHOX%, ALTOX%))
     Call T_Espera(3)
'     Image1.Picture = LoadPicture("C:\FLEXOFT\" & IMAG$)
     Call GRABAR_QR("C:\FLEXOFT\" & IMAG$, IDENTIFICADOR$)

     FILTX$ = TRIM$(str$(NC)) & ";" & NROOF$ & ";" & TRIM$(str$(CANTI)) & ";" & TRIM$(DESCRIOPERACION$) & ";" & TRIM$(str$(corte)) & ";" & TRIM$(DESCRIUBICACION$) & ";" & TRIM$(PESOX$)
     Call STDFORM("PRPR-CTP", FILTX$)
     Set CTP = Nothing
    
End Sub

Private Sub MSF2_DblClick()
    Reg& = MSF2.MouseRow
    COL& = MSF2.MouseCol
    If Reg& = 0 Or Reg& >= MSF2.Rows Then Exit Sub
    If COL& = 0 Then Exit Sub 'NO ENTRA SI EL CLIC ES COLUMNA 0
    '
    If DERACCE%("EDITINFO", "Editar Informe de Operaciones") < 2 Then Exit Sub
    
    Call SELECCIONARFILA(MSF2, Reg&)
    NUID& = XVALO(MSF2.TextMatrix(Reg&, 16))
  
    CantOriginalEnGrilla = XVALO(MSF2.TextMatrix(Reg&, 4))
    Unload VENTABNEW
    Campos$ = "Código/A16;Descripción/A28;Lote/A8;Cant.Cortes/F10.1;CI/a0;N°Op./F5;Descrip/A12;Hora/A0;Material/A16;Ubicación/A12;NroInf/a0;IDUbic/a0;N°Operario/F9;Nombre/A16;Peso/f8.2;IdReg/a0"
    VENTANA.Inicializar = 0
    VENTANA.Campox = Campos$
    VENTANA.UtilizaFormula = 1
    VENTANA.HabilitaBorrar = 1
    VENTANA.HabilitaInsertar = 0
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = "Modificaciones de Informes"
    VENTANA.CampEditables = "4;13;15"
    VENTANA.Inicializar = 1
    '
    Call EDITARFILA(MSF2, VENTABNEW.MSF, Reg&)
    '****************************************************************
    Call T_Espera(1 / 10) ' PARA QUE EL DOBLE CLICK NO APLIQUE SOBRE LA VENTANA QUE SE CARGA
30  VENTABNEW.Show 1
    If VENTABNEW.APROBADO% > 0 Then
       CantEnGrillaEditada = XVALO(VENTABNEW.MSF.TextMatrix(1, 4))
       
       Call CopiarMsfAOtroMsf(VENTABNEW.MSF, MSF2, 1)

       ID = XVALO(MSF2.TextMatrix(1, 16))
       idDelMovisto = XVALO(MSF2.TextMatrix(1, 19))
       If ID > 0 Then
          CONDI$ = "INFOFAB_ID = " & ID
          CantiOriginal = XVALO(MSF2.TextMatrix(1, 4))
          CANTI = XVALO(MSF2.TextMatrix(1, 4))
          NOPER% = XVALO(MSF2.TextMatrix(1, 13))
          If CANTI < 1 Then 'SI PONE CANTIDAD 0 SOLICITA CONFIRMACION DE BORRADO DE REGISTRO
             OPX% = COMALTER%("Confirmar ELIMINACIÓN  de Registro de Modo Permanente\\Cancelar\Eliminar Registro Seleccionado")
              
             If OPX% < 2 Then GoTo 30
             If OPX% = 2 Then
                If DERACCE%("ELIMINFO", "Eliminar Informe de Operaciones") < 2 Then Exit Sub

                'SOLO SE BORRA SI ES UN INFORME DE SUB-PROCESO EJ IMPRESION
                'O SI ES UN INFORME DE CORTE AL CUAL NO SE LE HAYA INFORMADO NINGUN PROCESO.
                SIBORRA% = 0
                If CantRegistros("INFOFAB", "IDREGPADRE > 0 and INFOFAB_ID = " & NUID&, "NOMESS") > 0 Then
                    IDREGISTROPADRE = XVALO(VALOBADA("INFOFAB", "IDREGPADRE", "INFOFAB_ID = " & NUID&, "NOMESS"))
                    SIBORRA% = 1
                ElseIf CantRegistros("INFOFAB", "OPERFUCIONADA  = " & NUID&, "NOMESS") < 1 Then
                      SIBORRA% = 1
                Else
                   If CantRegistros("INFOFAB", "IDREGPADRE  = " & NUID&, "NOMESS") < 1 Then
                      SIBORRA% = 1
                   End If
                End If
                
                If SIBORRA% = 1 Then
                   Call BORRAREGISTROS("INFOFAB", "INFOFAB_ID = " & NUID&, REGAF&, "NOMESS")
                   Call BORRAREGISTROS("MOVISTO", "MOVISTO_ID = " & idDelMovisto, REGAF&, "NOMESS")
                   Call ACTUREGISTRO("INFOFAB", "OPERFUCIONADA", "OPERFUCIONADA = " & NUID&, 0, REGAF&, "NOMESS") ' BORRO SI TIENE OPERFUCIONADA
                   If IDREGISTROPADRE > 0 Then 'SI TIENE PADRE ACTUALIZA LA CANTIDAD BORRADA
                     CANTI_bd = XVALO(VALOBADA("INFOFAB", "CANTREAL", "INFOFAB_ID = " & IDREGISTROPADRE, "NOMESS"))

                     ValorActualizado = CANTI_bd + CantOriginalEnGrilla
                     Call ACTUREGISTRO("INFOFAB", "CANTREAL", "INFOFAB_ID = " & IDREGISTROPADRE, ValorActualizado, REGAF&, "NOMESS")  ' BORRO SI TIENE OPERFUCIONADA
                   End If
                   GoTo 70
                Else
                   Call COMUNI("Imposible Eliminar Registro El Mismo Posee Informes de Sub-Proceso\Deberá Borrar Estos Primero")
                   GoTo 30
                End If
                
                   
             End If
          Else
             'si cantidad > 0 presenta mensaje y sale
             Call COMUNI("Opciòn de Modfificar Cantidad Deshabilitada\Ingrese Cantidad = '0' Para Elimnar el Registro y Vuelva a Cargarlo")
             Exit Sub
          End If
          
          nombre$ = TRIM$(ECOARCH$("PADRON", MKI$(NOPER%)))
          If nombre$ = "" Then
             Call COMUNI("Operario Inválido")
             GoTo 30
          End If
          Peso = XVALO(MSF2.TextMatrix(1, 15))
          Call ACTUREGISTRO("INFOFAB", "CANTREAL", CONDI$, CANTI, REGAF&, "NOMESS")
          Call ACTUREGISTRO("INFOFAB", "NUMOPERARIO", CONDI$, NOPER%, REGAF&, "NOMESS")
          Call ACTUREGISTRO("INFOFAB", "NAMEOPERARIO", CONDI$, nombre$, REGAF&, "NOMESS")
          Call ACTUREGISTRO("INFOFAB", "PESO", CONDI$, Peso, REGAF&, "NOMESS")
          If CantRegistros("INFOFAB", "IDREGPADRE > 0 and INFOFAB_ID = " & NUID&, "NOMESS") > 0 Then
                IDREGISTROPADRE = XVALO(VALOBADA("INFOFAB", "IDREGPADRE", "INFOFAB_ID = " & NUID&, "NOMESS"))
                If IDREGISTROPADRE > 0 Then 'SI TIENE PADRE ACTUALIZA LA CANTIDAD BORRADA
                   CANTI_bd = XVALO(VALOBADA("INFOFAB", "CANTREAL", "INFOFAB_ID = " & IDREGISTROPADRE, "NOMESS"))
                   ValorActualizado = CANTI_bd + (CantOriginalEnGrilla - CantEnGrillaEditada)
                   If ValorActualizado < 1 Then ValorActualizado = 0
                   Call ACTUREGISTRO("INFOFAB", "CANTREAL", "INFOFAB_ID = " & IDREGISTROPADRE, ValorActualizado, REGAF&, "NOMESS")  ' BORRO SI TIENE OPERFUCIONADA
                End If
          End If
       End If

    End If
70  Set VENTANA = Nothing
    Dim CTP As New Conteplast
    CIFAB% = CODINT%(Text10)
    NROOF$ = Label4
    Call CTP.CargaOperacionesInformadasMasSubProcesos(NROOF$, CIFAB%, MSF2, 0, "", 99)
    Set CTP = Nothing

99  Unload VENTABNEW

End Sub



Private Sub MSF2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Reg& = MSF2.MouseRow
    COL& = MSF2.MouseCol
    If Reg& = 0 Or Reg& >= MSF2.Rows Then Exit Sub
    If Button = 2 Then
     Call SELECCIONARFILA(MSF2, Reg&)
    EPAC$ = TRIM$(UCase$(EMPREAC$))

    

'    Call CARGA_ENCABEZADO(MSF2, "1 Código/A16;22Descripción/A28;33LOTE/A8;44Cant.Ctes./F10.1
'    ;55 CI/a0;N°Op./F5;66 Descrip/A12;77 Hora/A12;88 Material/A16;99 Ubicación/A10", Me)
    OPX% = COMALTER%("Puede Informar Una Ubicación en Sector o En Rack\\Cancelar\Ubicación en Sector\Ubicación en Rack\Imprmir Etiqueta Para Pallet\Informar Sub-Proceso")
    If OPX% < 2 Then Exit Sub
    If OPX% = 2 Then
        NORFA$ = Label4 'o.f
        NOP% = XVALO(MSF2.TextMatrix(Reg&, 6)) 'numeoper
        CODIEXLM$ = MSF2.TextMatrix(Reg&, 9) 'material
        LOTE$ = MSF2.TextMatrix(Reg&, 3) 'lote
        CANTI = MSF2.TextMatrix(Reg&, 4) 'cantidad de cortes
        NUORD% = XVALO(MSF2.TextMatrix(Reg&, 11))
       '**** CAMBIOS EN UBICACIONES 21/04/20
'        If XVALO(MSF2.TextMatrix(REG&, 10)) > 0 Then 'SI TIENE UBICACION EN RACK COMIENZA CON NUMERO
'            OPX% = COMALTER%("Atención !! Deberá Limpiar la Ubicación en Rack Antes de Asignar una Nueva Ubicación en Sector\\Cancelar\Ver Lista Para Borrar")
'            If OPX% < 2 Then Exit Sub
'            On Error Resume Next
'            TEXTOX$ = MSF2.TextMatrix(REG&, 10)
'            strArray = Split(TEXTOX$, ";")
'            RK = XVALO(strArray(0))
'            FL = XVALO(strArray(1))
'            KL = XVALO(strArray(2))
'            On Error GoTo 0
''            NUID& = INFOUBIC.BORRARUBICACIONCONZELECHO(RK, FL, KL)
'            If NUID& > 0 Then
'              Call BORRAREGISTROS("UBILOTE_ENVAP", "UBILOTE_ENVAP_ID = " & NUID&, REGAF&, "NOMESS")
'              Condi$ = "IDSUBOR = '" & NORFA$ & "'"
'              Condi$ = Condi$ + " AND NUMEOPER = " & NOP%
'              Condi$ = Condi$ + " AND  NUORINF = " & NUORD%
'              Call ACTUREGISTRO("INFOFAB", "COORDUBIC", Condi$, "", REGAF&, "NOMESS")
'              Call CargaOperacionesInformadasMasSubProcesos(NORFA$)
'            End If
'            Exit Sub
'        End If

        'SE LLAMA CON CONDICION PARA QUE MUESTRE LOS < 100 DE LOS SECTORES PARA ENCUADRAR DEPENDIENDO DE DONDE SE LLAME.
        AB = LISTASELECGENERAL("SECTPLANTA", "", "", 0, "CONVERT(int, campo1) < 100")
        IDSECTOR = XVALO(RESP_ZELE_VECT(1))
        SECTOR$ = TRIM$(RESP_ZELE_VECT(2))
        If SECTOR$ = "" Then Exit Sub
        CONDI$ = "IDSUBOR = '" & NORFA$ & "'"
        CONDI$ = CONDI$ + " AND NUMEOPER = " & NOP%
        CONDI$ = CONDI$ + " AND CANTREAL = " & CANTI
        CONDI$ = CONDI$ + " AND IDENLOTE = '" & LOTE$ & "'"
        CONDI$ = CONDI$ + " AND  CODIELEM = " & CODINT%(CODIEXLM$)
        CONDI$ = CONDI$ + " AND  NUORINF = " & NUORD%
        Call ACTUREGISTRO("INFOFAB", "COORDUBIC", CONDI$, SECTOR$, REGAF&, "NOMESS")
        Call ACTUREGISTRO("INFOFAB", "IDCOORDUBIC", CONDI$, IDSECTOR, REGAF&, "NOMESS")
    End If
    
    If OPX% = 3 Then
        INFOUBIC.Label4 = Label4 'o.f
        INFOUBIC.Text3 = XVALO(MSF2.TextMatrix(Reg&, 6)) 'numeoper
        INFOUBIC.TXTCODELEM = MSF2.TextMatrix(Reg&, 9) 'material
        INFOUBIC.TXTLOTE = MSF2.TextMatrix(Reg&, 3) 'lote
        INFOUBIC.TXTCANTICORTES = MSF2.TextMatrix(Reg&, 4) 'cantidad de cortes
        INFOUBIC.LBLNUMEORDENINFO = XVALO(MSF2.TextMatrix(Reg&, 11))
        INFOUBIC.lblidregistro = XVALO(MSF2.TextMatrix(Reg&, 16))
        INFOUBIC.Show 1
    End If
    
    CIFAB% = CODINT%(Text10)
    Dim CTP As New Conteplast
    Call CTP.CargaOperacionesInformadasMasSubProcesos(Label4, CIFAB%, MSF2, 0, "", 99)
    Set CTP = Nothing
    
    'Impresion
    If OPX% = 4 Or OPX% = 5 Then
        CIELEM% = CODINT(MSF2.TextMatrix(Reg&, 9)) 'material
        LOTE$ = TRIM$(MSF2.TextMatrix(Reg&, 3))
        CANTI = XVALO(MSF2.TextMatrix(Reg&, 4)) 'cantidad de cortes)
        NOP% = XVALO(MSF2.TextMatrix(Reg&, 6))
        NC = XVALO(Text6)
        NROPED& = XVALO(Text12)
        NROOF$ = Label4
        ORDENINFO = XVALO(MSF2.TextMatrix(Reg&, 11))
        UBICACION$ = TRIM$(MSF2.TextMatrix(Reg&, 10))
        IDUBICACION = XVALO(MSF2.TextMatrix(Reg&, 12))
        DESCRIOPERACION$ = MSF2.TextMatrix(Reg&, 7)
        CODIGOINTFAB = CODINT(Text10)
        IDINFOFAB& = XVALO(MSF2.TextMatrix(Reg&, 16))
       If OPX% = 4 Then
        Call PRINTETIQPALLET(CODEXT$(CIELEM%), LOTE$, CANTI, NOP%, NC, NROPED&, NROOF$, ORDENINFO, UBICACION$, IDUBICACION, DESCRIOPERACION$, CODIGOINTFAB, IDINFOFAB&)
       End If
       If OPX% = 5 Then
          Me.Height = 9510
          Call Centrar(Me)
          'SI TIENE UN VALOR EN IdRegPadre ES POR QUE ES UN INFORME QUE TIENE HERENCIA DE UN CORTE
          'ES DECIR YA SE GENERO UN INFORM DE ESTE, POR ENDE ES UN INFORME.
          LblId = XVALO(MSF2.TextMatrix(Reg&, 16))
          lblCantidadRegSelec = XVALO(MSF2.TextMatrix(Reg&, 4))
          lblKgRegSelec = XVALO(MSF2.TextMatrix(Reg&, 15))
          Me.TxtCantimpresa = XVALO(MSF2.TextMatrix(Reg&, 4)) 'cantidad de cortes)
          CONDI$ = " INFOFAB_ID = " & XVALO(LblId)
          TieneHerencia& = XVALO(VALOBADA("InfoFab", "IdRegPadre", CONDI$, "NOMESS"))
          If TieneHerencia& > 0 Then
             Call COMUNI("Atención !! El Registro Seleccionado Es Informe de Otro\Solo Puede Realizar Informes de Sub-Proceso de Cortes Primarios")
             Exit Sub
          End If
          
          Label25 = CTP.CantAcumuladaSubProceso&(LblId)
          
          TXTLECTURA = CODEXT$(CIELEM%) & "|" & LOTE$ & "|" & CANTI & "|" & NOP% & "|" & NC & "|" & NROPED& & "|" & NROOF$ & "|" & ORDENINFO & "|" & UBICACION$ & "|" & IDUBICACION
       End If
     End If
  End If

End Sub

Private Sub Text1_Change(Index As Integer)
    Label5(Index) = ECOARCH$("PADRON", MKI$(XVALO(Text1(Index))))
End Sub

Private Sub Text1_DblClick(Index As Integer)
    Call Command4_Click(Index)
End Sub

Private Sub Text10_Change()
   '
''   Call BLANFORMUGENERAL
'   CIXXI% = CODINT%(Text10)
'   Label3 = DESCRIT0$(CIXXI%)
'   '
'   '
'   Call COPYSTRU("OPERINF", "OPERINF1")
'   CONDI1$ = "CodIConj = " & CIXXI%
'   CONDI1$ = CONDI1$ + " ORDER BY NUMEOPER"
'   Call CARGALISEL("!OPERINF1", "SECOPER", "NumeOper/F2;DescOper/A32", CONDI1$)
'   Call COPYSTRU("ECORDEF", "ECORDEF1")
'   CONDI2$ = "COD_INTE = " & CIXXI% & " AND StatuOrf < 3"
'   Call CARGALISEL("!ECORDEF1", "ORDEFABR", "IdSubOr/A12;Cod_Exte/a16;Descrip/a32", CONDI2$)
   '
End Sub

Private Sub Text10_DblClick()
    Call Command1_Click
End Sub

Private Sub Text19_Change()
    On Error Resume Next
    TXT$ = TRIM$(ECOARCH$("TAMOTIV", Chr$(Text19)))
    On Error GoTo 0
    If TXT$ <> "" Then
       Label55 = TXT$
    Else
       Label55 = ""
    End If
End Sub

Private Sub Text2_LostFocus()
  Text2 = HORATEX$(HORANUM(Text2.TEXT))
End Sub

Private Sub Text3_DblClick()
    Call Command6_Click
End Sub

Private Sub Text4_DblClick()
    Call Command5_Click
End Sub

Private Sub Text5_DblClick()
    Call Command7_Click
End Sub

Private Sub Text6_Change()
   Label21 = rasocli$(XVALO(Text6))
End Sub

Private Sub Text7_LostFocus()
    Text7 = FORMATNUM$(XVALO(Text7), "F8.1")
End Sub


Sub BLANFORMU()
    '
    Text3 = ""
    Text1(0) = ""
    Label7 = ""
    Text1(0) = ""
    Label5(0) = ""
    TXTCODELEM = ""
    TXTLOTE = ""
    TXTCORTE = ""
    TXTCANTICORTES = ""
    Text4 = FECHATEX$(HOY(HO$))
    Text2 = ""
    Text7 = ""
    TXTNECESIDADCORTES = ""
    TXTRESTAN = ""
    TXTCORTESOPYLOTE = ""
    TXTCORTESOPERACION = ""

    '
End Sub

Sub BLANFORMUGENERAL()
    MSF1.Rows = 1
    MSF2.Rows = 1
    
    Label4 = ""
    Text6 = ""
    Text12 = ""
    Call BLANFORMU
End Sub

Private Sub TxtCantimpresa_DblClick()
   If Me.TxtCantimpresa.Locked = True Then
      Me.TxtCantimpresa.Locked = False
   Else
      Me.TxtCantimpresa.Locked = True
   End If
   'si se cambia ESTO VER TAMBIEN EN EL BOTON GRABAR QUE LO VUELVE A LLEVAR A TRUE
End Sub


Private Sub TXTLECTURA_Change()
   If Len(TXTLECTURA) > 10 Then
    LEIDO$ = TRIM$(TXTLECTURA)
     Dim strArray() As String
     On Error Resume Next
     strArray = Split(LEIDO$, "|")
     CANTVECTOR = UBound(strArray)
     For i% = 0 To CANTVECTOR
       Dato$ = (strArray(i%))
       Label31(i%) = Dato$
     Next i%
     On Error GoTo 0
   End If
   
   
End Sub


Private Sub TXTLECTURA_LostFocus()
   If TRIM(TXTLECTURA) <> "" Then
        If XVALO(LblId) < 1 Then
           CONDI$ = "numeoper= " & XVALO(Label31(3))
           CONDI$ = CONDI$ + "  AND codxelem = '" & Label31(0) & "'"
           CONDI$ = CONDI$ + "  AND NUORINF = " & XVALO(Label31(7))
           CONDI$ = CONDI$ + "  AND IDENLOTE = '" & Label31(1) & "'"
           CONDI$ = CONDI$ + "  AND IDSUBOR = '" & Label31(6) & "'"
           LblId = XVALO(VALOBADA("INFOFAB", "INFOFAB_ID", CONDI$, "NOMESS"))
           If XVALO(LblId) < 1 Then
             Call COMUNI("Atención La Lectura No se Corresponde Con la Orden de Fabricación Activa\Orden Activa: " & Label4 & "\ Orden Leída: " & Label31(6))
           End If
        End If
   End If
End Sub
Sub VALIDARBOTONFUSION()
    Me.Frame4.Visible = False
    Me.Frame2(1).Visible = True
    CANT_SELECT% = 0
    For i& = 1 To MSF2.Rows - 1
       If MSF2.TextMatrix(i&, 0) = "*" Then
          CANT_SELECT% = CANT_SELECT% + 1
       End If
    Next i&
    '
    If CANT_SELECT% > 2 Then
       Call COMUNI("Atención Se Han Seleccionado Más de 2 Registros")
    ElseIf CANT_SELECT% = 2 Then
       Me.Frame2(1).Visible = False
       Me.Frame4.Visible = True
    End If
    '
End Sub

