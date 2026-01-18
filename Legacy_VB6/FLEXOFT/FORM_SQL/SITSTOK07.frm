VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form SITSTOK07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00CDDADA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stocks - Existencias y Disponibilidad"
   ClientHeight    =   8040
   ClientLeft      =   405
   ClientTop       =   1050
   ClientWidth     =   11910
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SITSTOK07.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11910
   Begin MSFlexGridLib.MSFlexGrid mfgStockProyectado 
      Height          =   495
      Left            =   240
      TabIndex        =   76
      Top             =   8280
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   873
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CMD 
      Left            =   5760
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid gridStockProyectado 
      Height          =   3375
      Left            =   120
      TabIndex        =   75
      Top             =   4560
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   5953
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      BackColorFixed  =   8433840
      BackColorBkg    =   16777215
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.OptionButton Option3 
      Appearance      =   0  'Flat
      BackColor       =   &H00CDDADA&
      Caption         =   "Pedidos"
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
      Left            =   4215
      TabIndex        =   72
      ToolTipText     =   "Doble click Para Visualizar Pedidios Pendientes"
      Top             =   510
      Width           =   1240
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H00CDDADA&
      Caption         =   "L.Precios"
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
      Left            =   4215
      TabIndex        =   71
      Top             =   40
      Value           =   -1  'True
      Width           =   1275
   End
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H00CDDADA&
      Caption         =   "Maestro"
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
      Left            =   4215
      TabIndex        =   70
      Top             =   270
      Width           =   1240
   End
   Begin VB.Timer Timer5 
      Interval        =   15000
      Left            =   7320
      Top             =   1440
   End
   Begin VB.CommandButton Command20 
      Height          =   345
      Left            =   5600
      Picture         =   "SITSTOK07.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   69
      ToolTipText     =   "Analisis Estadisticos - Consumos de Stocks"
      Top             =   720
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Est."
      Height          =   255
      Left            =   5670
      TabIndex        =   68
      ToolTipText     =   "Analisis Estadisticos - Consumos de Stocks"
      Top             =   720
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   55000
      Left            =   5715
      Top             =   1620
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   55000
      Left            =   5610
      Top             =   1305
   End
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   105
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   67
      Top             =   735
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command13 
      Height          =   480
      Left            =   6160
      Picture         =   "SITSTOK07.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   61
      ToolTipText     =   "Muestra Datos de Item de Stock"
      Top             =   210
      Width           =   550
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Conver-Repa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8640
      TabIndex        =   60
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command11 
      Caption         =   "MIGRA-REPA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7200
      TabIndex        =   59
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   0
      Top             =   480
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8100
      Left            =   11040
      ScaleHeight     =   8070
      ScaleWidth      =   1170
      TabIndex        =   50
      Top             =   0
      Width           =   1200
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   105
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   73
         Top             =   7440
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   74
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command17 
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
         Picture         =   "SITSTOK07.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Stocks de Componentes de Item Activo"
         Top             =   4080
         Width           =   650
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Recup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   105
         TabIndex        =   64
         Top             =   4850
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.CommandButton Command15 
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
         Picture         =   "SITSTOK07.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Acceso Directo a Deposito Caotico"
         Top             =   3450
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Impo."
         Height          =   300
         Left            =   105
         TabIndex        =   62
         ToolTipText     =   "Importacion de Tabla Stockvie"
         Top             =   7770
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Print"
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
         Picture         =   "SITSTOK07.frx":0D72
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Imprime Situación de Stocks"
         Top             =   6480
         Width           =   650
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Move"
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
         Picture         =   "SITSTOK07.frx":13DC
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Realizar Movimientos de Stock"
         Top             =   5805
         Width           =   650
      End
      Begin VB.CommandButton COMMAND3 
         Caption         =   "Kdx"
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
         Picture         =   "SITSTOK07.frx":1526
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
         Top             =   2160
         Width           =   650
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Price"
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
         Picture         =   "SITSTOK07.frx":1830
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Consulta de Precios de Venta"
         Top             =   2805
         Width           =   650
      End
      Begin VB.CommandButton Command9 
         Caption         =   "List"
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
         Picture         =   "SITSTOK07.frx":1B3A
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Acceso a Consultas y Listados"
         Top             =   5160
         Width           =   650
      End
      Begin VB.CommandButton Command1 
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
         Picture         =   "SITSTOK07.frx":1C84
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Terminar - Salir de Consulta de Stocks"
         Top             =   60
         Width           =   650
      End
      Begin VB.CommandButton Command5 
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
         Picture         =   "SITSTOK07.frx":1DCE
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   750
         Width           =   650
      End
      Begin VB.CommandButton Command27 
         Caption         =   "Setup"
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
         Picture         =   "SITSTOK07.frx":20D8
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   1440
         Width           =   650
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2685
      Left            =   8580
      ScaleHeight     =   2655
      ScaleWidth      =   2295
      TabIndex        =   48
      Top             =   1560
      Visible         =   0   'False
      Width           =   2325
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Histórico"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   0
         TabIndex        =   49
         Top             =   0
         Width           =   2325
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2685
      Left            =   8580
      ScaleHeight     =   2625
      ScaleWidth      =   2265
      TabIndex        =   45
      Top             =   1560
      Visible         =   0   'False
      Width           =   2325
      Begin VB.Line Line5 
         X1              =   0
         X2              =   2310
         Y1              =   270
         Y2              =   270
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Cobertura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   0
         TabIndex        =   47
         Top             =   0
         Width           =   2325
      End
      Begin VB.Shape Shape6 
         FillColor       =   &H0080C0FF&
         FillStyle       =   0  'Solid
         Height          =   645
         Left            =   1220
         Top             =   525
         Width           =   950
      End
      Begin VB.Shape Shape5 
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   645
         Left            =   1220
         Top             =   1155
         Width           =   950
      End
      Begin VB.Shape Shape4 
         FillColor       =   &H000000C0&
         FillStyle       =   0  'Solid
         Height          =   645
         Left            =   1220
         Top             =   1785
         Width           =   950
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Solid
         Height          =   750
         Left            =   105
         Top             =   525
         Width           =   950
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   540
         Left            =   105
         Top             =   1260
         Width           =   950
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00C00000&
         FillStyle       =   0  'Solid
         Height          =   645
         Left            =   105
         Top             =   1785
         Width           =   950
      End
   End
   Begin VB.CommandButton Command10 
      Height          =   480
      Left            =   5600
      Picture         =   "SITSTOK07.frx":251A
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Visualizar Ordenes de Venta Pendientes de Entrega"
      Top             =   210
      Width           =   550
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00CDDADA&
      Caption         =   "SITUACIÓN DE STOCK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1120
      Left            =   105
      TabIndex        =   31
      Top             =   3120
      Width           =   8310
      Begin VB.CommandButton Command4 
         Height          =   345
         Left            =   105
         Picture         =   "SITSTOK07.frx":2824
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Muestra / Oculta Situación de Cobertura"
         Top             =   630
         Visible         =   0   'False
         Width           =   560
      End
      Begin VB.Image Image1 
         Height          =   250
         Left            =   105
         Picture         =   "SITSTOK07.frx":2B2E
         Stretch         =   -1  'True
         ToolTipText     =   "Histórico de Existencias"
         Top             =   315
         Visible         =   0   'False
         Width           =   560
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No Reservado"
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
         Left            =   2290
         TabIndex        =   44
         Top             =   420
         Width           =   1380
      End
      Begin VB.Label Label404 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   2640
         TabIndex        =   43
         Top             =   675
         Width           =   1170
      End
      Begin VB.Label Label401 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   960
         TabIndex        =   37
         Top             =   675
         Width           =   1170
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sock Neto"
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
         Left            =   920
         TabIndex        =   36
         Top             =   420
         Width           =   780
      End
      Begin VB.Label Label402 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   4560
         TabIndex        =   35
         Top             =   675
         Width           =   1170
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Disponible"
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
         Left            =   4020
         TabIndex        =   34
         Top             =   420
         Width           =   1275
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Faltantes"
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
         Left            =   5970
         TabIndex        =   33
         Top             =   420
         Width           =   1170
      End
      Begin VB.Label Label403 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   6480
         TabIndex        =   32
         Top             =   675
         Width           =   1170
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00CDDADA&
      Caption         =   "A INGRESAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   5670
      TabIndex        =   15
      Top             =   1470
      Width           =   2745
      Begin VB.Label Label303 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Height          =   330
         Left            =   1365
         TabIndex        =   30
         Top             =   1155
         Width           =   1170
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
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
         Left            =   420
         TabIndex        =   29
         Top             =   1260
         Width           =   855
      End
      Begin VB.Label Label301 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Height          =   330
         Left            =   1365
         TabIndex        =   19
         Top             =   315
         Width           =   1170
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Compras"
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
         Left            =   420
         TabIndex        =   18
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label302 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Height          =   330
         Left            =   1365
         TabIndex        =   17
         Top             =   735
         Width           =   1170
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fabricación"
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
         TabIndex        =   16
         Top             =   840
         Width           =   1065
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Caption         =   "COMPROMETIDO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   2835
      TabIndex        =   10
      Top             =   1470
      Width           =   2745
      Begin VB.Label Label203 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   1365
         TabIndex        =   28
         Top             =   1155
         Width           =   1170
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
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
         Left            =   420
         TabIndex        =   27
         Top             =   1260
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Asignados y Reservas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   210
         TabIndex        =   14
         Top             =   660
         Width           =   1065
      End
      Begin VB.Label Label202 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   1365
         TabIndex        =   13
         Top             =   735
         Width           =   1170
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pedidos"
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
         Left            =   420
         TabIndex        =   12
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label201 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   1365
         TabIndex        =   11
         Top             =   315
         Width           =   1170
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "STOCK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   105
      TabIndex        =   8
      Top             =   1470
      Width           =   2640
      Begin VB.Label Label103 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   1260
         TabIndex        =   26
         Top             =   1155
         Width           =   1170
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
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
         Left            =   315
         TabIndex        =   25
         Top             =   1260
         Width           =   855
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Remoto"
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
         Left            =   315
         TabIndex        =   24
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label102 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   1260
         TabIndex        =   23
         Top             =   735
         Width           =   1170
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Propio"
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
         Left            =   315
         TabIndex        =   22
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label101 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   1260
         TabIndex        =   9
         Top             =   315
         Width           =   1170
      End
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2004
      Left            =   8580
      TabIndex        =   5
      Top             =   480
      Width           =   2325
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
      Left            =   1020
      TabIndex        =   0
      Top             =   210
      Width           =   2850
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   7770
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   210
      Width           =   645
   End
   Begin VB.TextBox Text3 
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
      Height          =   600
      Left            =   105
      MultiLine       =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   735
      Width           =   5370
   End
   Begin VB.CommandButton Command2 
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
      Left            =   3885
      TabIndex        =   1
      Top             =   210
      Width           =   150
   End
   Begin VB.ListBox LIUBIC 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   8580
      TabIndex        =   2
      Top             =   3070
      Width           =   2325
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   960
      OleObjectBlob   =   "SITSTOK07.frx":2E38
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   1680
      OleObjectBlob   =   "SITSTOK07.frx":306C
      TabIndex        =   66
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Line Line7 
      X1              =   0
      X2              =   12000
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Economico:"
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
      Left            =   5985
      TabIndex        =   41
      Top             =   1150
      Width           =   1275
   End
   Begin VB.Label LECONOMI 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Left            =   7350
      TabIndex        =   40
      Top             =   1050
      Width           =   1065
   End
   Begin VB.Label STOMI 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Left            =   7350
      TabIndex        =   39
      Top             =   720
      Width           =   1065
   End
   Begin VB.Label Label30 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mínimo:"
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
      Left            =   5985
      TabIndex        =   38
      Top             =   815
      Width           =   1275
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ubicaciones"
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
      Left            =   8580
      TabIndex        =   21
      ToolTipText     =   "Doble-Click para Actualizar Ubicaciones"
      Top             =   2805
      Width           =   2325
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Existencias"
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
      Left            =   8580
      TabIndex        =   20
      Top             =   210
      Width           =   2325
   End
   Begin VB.Label Label3 
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
      Height          =   330
      Left            =   -15
      TabIndex        =   7
      Top             =   315
      Width           =   960
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "U-M"
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
      Left            =   7305
      TabIndex        =   6
      Top             =   315
      Width           =   375
   End
   Begin VB.Menu Archiv 
      Caption         =   "Transacciones"
      Begin VB.Menu Abrir 
         Caption         =   "Abrir"
      End
      Begin VB.Menu php0 
         Caption         =   "-"
      End
      Begin VB.Menu FichKar 
         Caption         =   "Ficha Tipo Kardex"
      End
      Begin VB.Menu SitCob 
         Caption         =   "Mostrar / Ocultar Situación de Cobertura"
      End
      Begin VB.Menu CoPreVe 
         Caption         =   "Consulta de Precios de Venta"
      End
      Begin VB.Menu mnuTrazabilidadXOF 
         Caption         =   "Trazabilidad x O.F."
      End
      Begin VB.Menu php3 
         Caption         =   "-"
      End
      Begin VB.Menu menuMigrarUbicacionesASql 
         Caption         =   "Migrar Ubicaciones a SQL"
      End
      Begin VB.Menu php4 
         Caption         =   "-"
      End
      Begin VB.Menu Salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu Edicion 
      Caption         =   "Edición"
      Begin VB.Menu ubicacionxprov 
         Caption         =   "Ubicaciones por Proveedor"
         Begin VB.Menu stockdetallexproveedor 
            Caption         =   "Stock Detalle x Proveedor"
         End
         Begin VB.Menu soloxproveedor 
            Caption         =   "Por proveedor"
         End
      End
      Begin VB.Menu asignalotesitemfacturad 
         Caption         =   "Consulta y Asignación de Lotes a items Facturados"
         Begin VB.Menu VerSoloSinLote 
            Caption         =   "Ver Solo Sin Lote"
         End
         Begin VB.Menu VerTodasLasFacturas 
            Caption         =   "Ver Todas Las Facturas"
         End
      End
   End
   Begin VB.Menu ConsyList 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu Gen 
         Caption         =   "Generales"
         Begin VB.Menu mnuStockConMetrosyKilos 
            Caption         =   "Stock con Metros y Kilos"
         End
         Begin VB.Menu mnuStockConMetrosyKilosEnvaplast 
            Caption         =   "Stock de Bobinas Agrupadas x código"
         End
         Begin VB.Menu ExistyUbi 
            Caption         =   "Existencias y Ubicaciones"
         End
         Begin VB.Menu FichaKardex 
            Caption         =   "Ficha Kardex -Historico de Movimientos-"
         End
         Begin VB.Menu GesRese 
            Caption         =   "Gestión de Reservas"
         End
         Begin VB.Menu mnuStockCodigoRango 
            Caption         =   "Stock Códigos Rango"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStockRangoAgrupado 
            Caption         =   "Stock por Rango (AGRUPADO)"
         End
         Begin VB.Menu mnuStockProveedorRango 
            Caption         =   "Stock por Proveedor y Rango"
         End
         Begin VB.Menu mnuStockRangoTodos 
            Caption         =   "Stock por Rango (TODOS)"
         End
         Begin VB.Menu mnuStFeMatFam 
            Caption         =   "Stock x Fechas, T.Mat. y Familias (Stock de Bobinas Totales todos los  Deposito)"
         End
         Begin VB.Menu mnuScrapMatPrim 
            Caption         =   "Stock x Fechas, T.Mat . y Familias (Scrap o Materias Primas)"
         End
         Begin VB.Menu mnuReportePlantaLaminado 
            Caption         =   "Reporte Planta Laminado"
         End
         Begin VB.Menu SEP 
            Caption         =   "-"
         End
         Begin VB.Menu mnuStTotLuzCamb 
            Caption         =   "STOCK TOTAL BOBINAS (Detalle Luzuriaga y Cambaceres)"
         End
         Begin VB.Menu mnuMovStLuzuriaga 
            Caption         =   "MOVIMIENTOS DE STOCKS LUZURIAGA"
         End
         Begin VB.Menu mnuStockMetrosvsFtecnica 
            Caption         =   "Stock Bobinas (Solo Difieren Con Metros de Ficha Técnica)"
         End
      End
      Begin VB.Menu DeCtrl 
         Caption         =   "De Control"
         Begin VB.Menu ExFiAct 
            Caption         =   "Existencias Físicas Actuales"
            Begin VB.Menu ExFiPan 
               Caption         =   "Por Pantalla"
            End
            Begin VB.Menu ExFiImp 
               Caption         =   "Por Impresión"
            End
         End
         Begin VB.Menu ExiNeg 
            Caption         =   "Existencias Negativas"
            Begin VB.Menu ExNegPan 
               Caption         =   "Por Pantalla"
            End
            Begin VB.Menu ExNegImpr 
               Caption         =   "Por Impresión"
            End
         End
         Begin VB.Menu ItStSubMi 
            Caption         =   "Items con Stock Sub-Mínimo"
            Begin VB.Menu ItStPant 
               Caption         =   "Por Pantalla"
            End
            Begin VB.Menu ItStImp 
               Caption         =   "Por Impresión"
            End
         End
         Begin VB.Menu DiGenSt 
            Caption         =   "Disponibilidad General de Stock"
            Begin VB.Menu DiStPan 
               Caption         =   "Por Pantalla"
            End
            Begin VB.Menu DiStImp 
               Caption         =   "Por Impresión"
            End
         End
         Begin VB.Menu mnuDisponibilidadStockRango 
            Caption         =   "Disponibilidad de Stock Por Rango"
            Begin VB.Menu mnuListadoDisponibilidadPorRango 
               Caption         =   "Listado de Disponibilidad"
            End
            Begin VB.Menu mnuTablaStockSeguridad 
               Caption         =   "Tabla de Stock de Seguridad"
            End
            Begin VB.Menu mnuCtrlResvsTransf 
               Caption         =   "Control Reservas vs. Transferidos"
            End
         End
      End
      Begin VB.Menu mnuAnalisisFabricacionSegunStockRecepcion 
         Caption         =   "Análisis de Fabricación Según Stock y Recepciones"
         Begin VB.Menu mnuFabricacionSegunStockRecepcion 
            Caption         =   "Fabricacion Según Stock y Recepción"
         End
         Begin VB.Menu mnuConjuntosAnalizarFabricacion 
            Caption         =   "Tabla de Conjuntos a Analizar"
         End
         Begin VB.Menu mnuImportarConjuntosDesdeExcel 
            Caption         =   "Importar Conjuntos desde Excel"
         End
      End
      Begin VB.Menu MaDep 
         Caption         =   "Material en Depósitos"
      End
      Begin VB.Menu GesReser 
         Caption         =   "Gestión de Reservas"
      End
      Begin VB.Menu mnuRepfabVsVtasXFamilia 
         Caption         =   "Reporte de Fabricación Vs. Ventas x Familia"
      End
      Begin VB.Menu mnuQuiebreDeStock 
         Caption         =   "Rotura de Stock"
         Visible         =   0   'False
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStockMulDep 
         Caption         =   "Stock Multiples Depósitos"
      End
      Begin VB.Menu MnuComparativoVR 
         Caption         =   "Reporte Comparativo Ventas vs. Recepciones"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuConfig_RtaMDB 
         Caption         =   "Configurar Ruta B.D. MDB"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu AcceDire 
      Caption         =   "Accesos Directos"
      Begin VB.Menu MovStock 
         Caption         =   "Movimientos de Stock"
      End
      Begin VB.Menu php25 
         Caption         =   "-"
      End
      Begin VB.Menu ParFabr 
         Caption         =   "Partes de Fabricación"
      End
      Begin VB.Menu php7 
         Caption         =   "-"
      End
      Begin VB.Menu OtrAccDirec 
         Caption         =   "Otros Accesos Directos"
      End
   End
   Begin VB.Menu Confi 
      Caption         =   "Configuración"
      Begin VB.Menu mnuConfiguracionGeneral 
         Caption         =   "Configuración General"
      End
   End
   Begin VB.Menu mnuInfConteplast 
      Caption         =   "Informes Conteplast"
      Begin VB.Menu mnuProductosTerminados 
         Caption         =   "Stock Productos Terminados"
      End
      Begin VB.Menu mnuReporteEStadistico 
         Caption         =   "Reporte Estadistico"
      End
      Begin VB.Menu mnuKgIngVsSalKg 
         Caption         =   "Reporte Kg ingresados vs. Kg. producidos"
      End
   End
   Begin VB.Menu Ayud 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "SITSTOK07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CONTROCLIKK%

   Dim objexcel As Excel.Application
   Dim Libro As Excel.Workbook
   Dim Hoja  As Excel.Worksheet
   Dim FACTURADO()
   Public BaDaMoviExt As Database

Sub ArtFaltantes()
   
   
    ' ARMO EL ENCABEZADO
    TX$ = ""
    Call REPLA(TX$, AJUSTI$("CÓDIGO", 20), 1, 20)
    Call REPLA(TX$, AJUSTI$("DESCRIPCIÓN", 36), 22, 36)
    Call REPLA(TX$, AJUSTD$("FALTANTE", 9), 60, 9)
   
    MAI_TEXTO$ = TX$ + vbCrLf
    MAI_TEXTO$ = MAI_TEXTO$ + String$(69, "=") + vbCrLf
    '
    J& = 0
    ' RECORRO EL ARCHIVO Y GUARDO LOS QUE TIENEN FALTANTES > 0
    ' Y QUE ESTEN EN LISTA DE PRECIOS
    FIN& = ULTREG&("STODISP")
   
    For i& = 1 To FIN&
      XX$ = REGLEIDO$("STODISP", i&)
      CI% = CVI(Mid$(XX$, 1, 2))
      CANTFALTA = CVS(Mid$(XX$, 33, 4))
      If CANTFALTA > 0 Then
         CANTLIPRE& = CantRegistros&("PREVENTA", "COD_PIEZA = " & CI%)
         If CANTLIPRE& > 0 Then
            J& = J& + 1
            TX$ = ""
            Call REPLA(TX$, CODEXT$(CI%), 1, 20)
            Call REPLA(TX$, DESCRIT$(CI%), 22, 36)
            Call REPLA(TX$, AJUSTD$(FORMATNUM$(CANTFALTA, "f9.1"), 9), 60, 9)
            MAI_TEXTO$ = MAI_TEXTO$ + TX$ + vbCrLf
            
         End If
      End If
    Next i&
    MAI_TEXTO$ = MAI_TEXTO$ + String$(69, "=") + vbCrLf + vbCrLf + vbCrLf
    '
    MAI_TEXTO$ = MAI_TEXTO$ + " Artículos en O.C. con Fecha de Entrega Solicitada Anterior al Día de Hoy" + vbCrLf
    MAI_TEXTO$ = MAI_TEXTO$ + String$(69, "=") + vbCrLf + vbCrLf + vbCrLf
    Call ABRECOMPRAS
    '
    'FECHA DE  HOY
    HOII% = HOY(HOS$)
    DIAX = FETEXCOR1$(HOII%)
    '
    TX$ = ""
    '
    ' ARMO EL ENCABEZADO
    Call REPLA(TX$, AJUSTD$("O.COMPRA", 6), 1, 6)
    Call REPLA(TX$, AJUSTD$("N.PR.", 6), 8, 6)
    Call REPLA(TX$, AJUSTI$("R.SOCIAL", 24), 15, 24)
    Call REPLA(TX$, AJUSTD$("CANT.COMP.", 10), 40, 10)
    Call REPLA(TX$, AJUSTD$("CANT.REC.", 10), 51, 10)
    Call REPLA(TX$, AJUSTD$("FALT.ENT", 10), 62, 10)
    Call REPLA(TX$, AJUSTD$("F.ENT.SL.", 10), 73, 8)
    MAI_TEXTO$ = MAI_TEXTO$ + TX$ + vbCrLf
    MAI_TEXTO$ = MAI_TEXTO$ + String$(69, "=") + vbCrLf
    '
    ' CONSULTA AL OCOMDETA DE LAS O.C. PENDIENTES FECHA DE ENTREGA VENCIDA
    SQLX$ = "SELECT COD_INTE, Nume_Ocom, Npro_Ocom, cant_ocom, Cant_Rec, Fentre_Sol FROM OCOMDETA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE Stat_Ocom < 8 "
    SQLX$ = SQLX$ + " AND FENTRE_SOL < '" & DIAX & "'"
    SQLX$ = SQLX$ + " AND CANT_OCOM > CANT_REC "
    SQLX$ = SQLX$ + " AND  Stat_Ocom <> -1 " 'Para que no  hayan O/C's sin aprobar
    SQLX$ = SQLX$ + " order by NUME_OCOM "
'    Dim COMTMP As ADODB.Recordset
'    Set COMTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    Set COMTMP = READSET(SQLX$)
    With COMTMP
      Do While Not .EOF
         'ASIGNO
         CI% = XVALO(!cod_inte)
         CANTLIPRE& = CantRegistros&("PREVENTA", "COD_PIEZA = " & CI%)
         If CANTLIPRE& > 0 Then
            J& = J& + 1
            TX$ = ""
            OCOMP% = XVALO(!nume_ocom)
            PROV% = XVALO(!Npro_Ocom)
            CANCOMPRA = XVALO(!cant_ocom)
            CANTRECIB = XVALO(!Cant_Rec)
            FECHENTRE% = CVINT(!Fentre_Sol)
            CANTFALT = CANCOMPRA - CANTRECIB
            '
            'ARMO LA LINEA DE TEXTO
            Call REPLA(TX$, FORMATNUM$(OCOMP%, "F6.0"), 1, 6)
            Call REPLA(TX$, FORMATNUM$(PROV%, "F6.0"), 8, 6)
            Call REPLA(TX$, rasocli$(-1 * PROV%), 15, 24)
            Call REPLA(TX$, FORMATNUM$(CANCOMPRA, "F10.1"), 40, 10)
            Call REPLA(TX$, FORMATNUM$(CANTRECIB, "F10.1"), 51, 10)
            Call REPLA(TX$, FORMATNUM$(CANTFALT, "F10.1"), 62, 10)
            Call REPLA(TX$, FECHATEX$(FECHENTRE%), 73, 8)
            MAI_TEXTO$ = MAI_TEXTO$ + TX$ + vbCrLf
         End If
      .MoveNext
      Loop
    End With
    COMTMP.Close
    Set COMTMP = Nothing

   
    ' SI NO SE REGISTRARON FALTANTES
    If J& < 1 Then
      Exit Sub
    End If
    '
    'GENERACION DE ARCHIVO TXT CON INFORM
    Call SAVEFILE(LUCOM$ + "ART_FALTANTES.txt", MAI_TEXTO$)
    '
    'ENVIO DE MAIL A VENTAS
    'MAIL SEGUN CONFIGURACION
    MAILADMIN$ = TRIM$(LCase(Control("DIREMAI", "MAILVNTA")))
    '
    'Validacion de mail
    If MAILADMIN$ = "" Or InStr(MAILADMIN$, "@") < 1 Then
      Call COMUNI("Dirección de Mail (" & MAILADMIN$ & ") No Válida")
      Exit Sub
    End If
   
    'ARMO TEXTO DE ASUNTO
    HOII% = HOY(HO$)
    ASUNTO$ = "Informe de Faltantes de Stocks y Compras Con Fecha de Entrega Vencida al Día: " & HO$
    '
    'ASIGNO DATOS A LOS PARAMETROS
    MAI_DESTI$ = TRIM$(MAILADMIN$)
    MAI_DIREL$ = MAI_DESTI$
    MAI_ASUNT$ = ASUNTO$
   '
    'ENVIO MAIL
    Call GENERAMAIL("/AUTO")

End Sub

Function CantingreRangoFecha(CI, Des%, Has%)
   CantingreRangoFecha = 0
   SQLX$ = "SELECT SUM(CANTINGRE) AS ING FROM MOVISTO WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE COD_INTE = " & CI
   SQLX$ = SQLX$ + " AND Fechmov >= '" & FETEXCOR1$(Des%) & "'"
   SQLX$ = SQLX$ + " AND Fechmov <= '" & FETEXCOR1$(Has%) & "'"
   SQLX$ = SQLX$ + " AND CODIMOVI = 'BR'"
   SQLX$ = SQLX$ + " GROUP BY COD_INTE"
   Set rs = READSET(SQLX$)
   With rs
     CantingreRangoFecha = XVALO(!ING)
   End With
   rs.Close
   Set rs = Nothing
End Function



Sub CargarDatosSysflom(Des%, Has%)
    '
   Desde = CDbl(CVDAT(Des%))
   Hasta = CDbl(CVDAT(Has%))
   Screen.MousePointer = 11
   Dim FACTURADO#()
   Erase FACTURADO#()
   
   ReDim Preserve FACTURADO#(10)
   SQLX$ = " SELECT SUM(UNIDADES) AS CANTING, ARTICULO FROM ITITEMS "
   SQLX$ = SQLX$ + " WHERE (letra = 'A'"
   SQLX$ = SQLX$ + " OR letra = 'B'"
   SQLX$ = SQLX$ + " OR letra = 'E')"
   SQLX$ = SQLX$ + " AND INT(Cdbl(FECEMISION)) >= " & Desde & " "
   SQLX$ = SQLX$ + " AND INT(Cdbl(FECEMISION))  <= " & Hasta & " "
   SQLX$ = SQLX$ + " AND ES  = 'E' "
   SQLX$ = SQLX$ + " GROUP BY ARTICULO"
   SQLX$ = SQLX$ + " ORDER BY ARTICULO"
   '
   Set rs = BaDaMoviExt.OpenRecordset(SQLX$, dbOpenSnapshot)
   With rs
     Do While Not .EOF
            J& = J& + 1
            CI = CODINT%(!ARTICULO)
            If CI > 0 Then
              If UBound(FACTURADO#) < CI Then
                ReDim Preserve FACTURADO#(CI)
              End If
              FACTURADO#(CI) = FACTURADO#(CI) + XVALO(!CANTING)
            End If
            .MoveNext
     Loop
   End With
   rs.Close
   Set rs = Nothing
   
   SQLX$ = " SELECT SUM(UNIDADES) AS CANTSAL , ARTICULO FROM ITITEMS "
   SQLX$ = SQLX$ + " WHERE (letra = 'A'"
   SQLX$ = SQLX$ + " OR letra = 'B'"
   SQLX$ = SQLX$ + " OR letra = 'E')"
   SQLX$ = SQLX$ + " AND INT(Cdbl(FECEMISION)) >= " & Desde & " "
   SQLX$ = SQLX$ + " AND INT(Cdbl(FECEMISION))  <= " & Hasta & " "
   SQLX$ = SQLX$ + " AND ES  = 'S'"
   SQLX$ = SQLX$ + " GROUP BY ARTICULO"
   SQLX$ = SQLX$ + " ORDER BY ARTICULO"
   '
   Set RS1 = BaDaMoviExt.OpenRecordset(SQLX$, dbOpenSnapshot)
   With RS1
     Do While Not .EOF
            J& = J& + 1
            CI = CODINT%(!ARTICULO)
            If CI > 0 Then
              If UBound(FACTURADO#) < CI Then ReDim Preserve FACTURADO#(CI)
              FACTURADO#(CI) = FACTURADO#(CI) + XVALO(!CANTSAL)
            End If
            .MoveNext
     Loop
   End With
   RS1.Close
   Set RS1 = Nothing
   

   SQLX$ = " SELECT ARTICULO FROM ITITEMS "
   SQLX$ = SQLX$ + " WHERE (letra = 'A'"
   SQLX$ = SQLX$ + " OR letra = 'B'"
   SQLX$ = SQLX$ + " OR letra = 'E')"
   SQLX$ = SQLX$ + " AND INT(Cdbl(FECEMISION)) >= " & Desde & " "
   SQLX$ = SQLX$ + " AND INT(Cdbl(FECEMISION))  <= " & Hasta & " "
   SQLX$ = SQLX$ + " GROUP BY ARTICULO"
   SQLX$ = SQLX$ + " ORDER BY ARTICULO"
   '
   MUESTRASTOCKS% = 0
   ddias = Control("COMPARAT", "RECEFACT")
   If CompRecFa.Check1.Value = 1 Then
       MUESTRASTOCKS% = 1
   End If
   If MUESTRASTOCKS% > 0 Then
      DEXDE% = DIANTE(Des%, 1)
      HAXTA% = DIASPOST(Has%, 1)
   Else
      DEXDE% = DIANTE(Des%, XVALO(ddias))
      HAXTA% = DIANTE(Has%, XVALO(ddias))
      d = FECHATEX$(DEXDE%)
      HA = FECHATEX$(HAXTA%)
   End If
   '
   ABC = UBound(FACTURADO#)
   Set RS2 = BaDaMoviExt.OpenRecordset(SQLX$, dbOpenSnapshot)
   With RS2
     Do While Not .EOF
         CI = CODINT%(!ARTICULO)
         If CI > 0 Then
            REG& = REG& + 1
            CompRecFa.MSFDET.Rows = REG& + 1

            CompRecFa.MSFDET.TextMatrix(REG&, 1) = !ARTICULO
            CompRecFa.MSFDET.TextMatrix(REG&, 2) = DESCRIT0$(CI)
            CompRecFa.MSFDET.TextMatrix(REG&, 3) = FORMATNUM$(SALDOAFECHA(CI, DEXDE%, 1), "F12.1")
            CompRecFa.MSFDET.TextMatrix(REG&, 4) = FORMATNUM$(FACTURADO#(CI), "F12.1")
               INGRE = CantingreRangoFecha(CI, DEXDE%, HAXTA%)
            CompRecFa.MSFDET.TextMatrix(REG&, 5) = FORMATNUM$(INGRE, "F12.1")
            CompRecFa.MSFDET.TextMatrix(REG&, 6) = FORMATNUM$(FACTURADO#(CI) - INGRE, "F12.1")
            CompRecFa.MSFDET.TextMatrix(REG&, 7) = FORMATNUM$(SALDOAFECHA(CI, HAXTA%, 1), "F12.1") ' saldofin
'            CompRecFa.MSFDET.TextMatrix(REG&, 8) = CI
          End If
         .MoveNext
     Loop
   End With
   RS2.Close
   Set RS2 = Nothing
   
   CompRecFa.MSFDET.Rows = REG& + 3
   Call COLOREAR_GRILLA_SOLAFILA(CompRecFa.MSFDET, vbBlack, REG&)
   Call COLOREAR_GRILLA_SOLAFILA(CompRecFa.MSFDET, &HCDDADA, REG& + 2)
   
   CompRecFa.MSFDET.TextMatrix(REG& + 2, 1) = "Consumo de Componentes"
   
   REG& = REG& + 3
   CompRecFa.MSFDET.Rows = REG&
   Dim CONSUMO#()
   ReDim Preserve CONSUMO#(10)
'   Dim USODECOMP(32767, 32767) ' CODIGO DE COMPONENTE, CODIGO PTERMINADO
'   Dim CANTIUSADA(32767, 32767) ' CODIGO DE COMPONENTE , CODIGO PTERMINADO
   'ACUMULO EL CONSUMO DE CADA COMPONENTE EN UN VECTOR CONSUMO#(CI)
   Campos$ = "CI/F0;Código Comp./A16;Descripción/A24;Cipt/f0;Código P.Term./A16;Descripción/A24;Uso Un./F12.1;Qt.Fact./F12.1;Total/F12.1"
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
   FRMZELECHO.Width = 15000
   FRMZELECHO.MSF2.Rows = 1
   J& = 0
   For i& = 1 To CompRecFa.MSFDET.Rows - 1
      CODI$ = CompRecFa.MSFDET.TextMatrix(i&, 1)
      CI1% = CODINT%(CODI$)
      If CI1% > 0 Then
        CANTI = FACTURADO#(CI1%)
'       Call LEEEXPLO(CODI$, 1)
        Call LEEEXPLOMUL(CODI$, 1, 3)
         
        For h% = 1 To CAIT%
             CIY% = CIJ%(h%): CI = CIY%
             CAN = USOI(h%) * CANTI
'             USODECOMP(CI1, CI) = CAN
             If UBound(CONSUMO#) < CI Then
               ReDim Preserve CONSUMO#(CI)
             End If
             CONSUMO#(CI) = CONSUMO#(CI) + CAN
             CODIGOEXT$ = CODEXT$(CIY%)
             DESCRIPCIOX$ = DESCRIT$(CIY%)
             'APROVECHO Y GUARDO LA RELACION DE CONSUMO CON CADA PRODUCTO TERMINADO EN EL ZELECHO PARA PODER MOSTRAR ON DEMAND
             J& = J& + 1
             FRMZELECHO.MSF2.Rows = J& + 1
             FRMZELECHO.MSF2.TextMatrix(J&, 1) = CIY% ' CODIGO COMPONENTE
             FRMZELECHO.MSF2.TextMatrix(J&, 2) = CODIGOEXT$
             FRMZELECHO.MSF2.TextMatrix(J&, 3) = DESCRIPCIOX$
             FRMZELECHO.MSF2.TextMatrix(J&, 4) = CI1 'CODIGO PRODUCTO TERMINADO
             FRMZELECHO.MSF2.TextMatrix(J&, 5) = CODI$
             FRMZELECHO.MSF2.TextMatrix(J&, 6) = DESCRIT0$(CI1%)
             FRMZELECHO.MSF2.TextMatrix(J&, 7) = FORMATNUM$(USOI(h%), "F12.5")
             FRMZELECHO.MSF2.TextMatrix(J&, 8) = FORMATNUM$(CANTI, "F12.5")
             FRMZELECHO.MSF2.TextMatrix(J&, 9) = FORMATNUM$(USOI(h%) * CANTI, "F12.5")
             
         Next h%
         
'         CANTIUSADA(CI1, CI) = CANTIUSADA(CI1, CI) + CONSUMO#(CI)
      End If
   Next i&
   '
   Screen.MousePointer = 1

   'RECORRO EL VECTOR DE CONSUMO Y AGREGO AL FINAL
   For I1% = 1 To UBound(CONSUMO#)
      If Abs(CONSUMO#(I1%)) > 0 Then
         CI = I1%
         REG& = REG& + 1
         CompRecFa.MSFDET.Rows = REG& + 1
         CompRecFa.MSFDET.TextMatrix(REG&, 1) = CODEXT$(CI)
         CompRecFa.MSFDET.TextMatrix(REG&, 2) = DESCRIT$(CI)
         CompRecFa.MSFDET.TextMatrix(REG&, 3) = FORMATNUM$(SALDOAFECHA(CI, DEXDE%, 1), "F12.1")
         CompRecFa.MSFDET.TextMatrix(REG&, 4) = FORMATNUM$(CONSUMO#(CI), "F12.1")
            INGRE = CantingreRangoFecha(CI, DEXDE%, HAXTA%)
         CompRecFa.MSFDET.TextMatrix(REG&, 5) = FORMATNUM$(INGRE, "F12.1")
         CompRecFa.MSFDET.TextMatrix(REG&, 6) = FORMATNUM$(INGRE - CONSUMO#(CI), "F12.1")
         CompRecFa.MSFDET.TextMatrix(REG&, 7) = FORMATNUM$(SALDOAFECHA(CI, HAXTA%, 1), "F12.1") ' saldofin
         Call COLOREAR_GRILLA_SOLAFILA(CompRecFa.MSFDET, &HC0FFFF, REG&)
      End If
      If I1% >= UBound(CONSUMO#) Then Exit For
   Next I1%
   
   For i& = 1 To CompRecFa.MSFDET.Rows - 1
       For J1% = 1 To CompRecFa.MSFDET.Cols - 1
          If XVALO(CompRecFa.MSFDET.TextMatrix(i&, J1%)) < 0 Then
             Call COLOREAR_GRILLA_SOLACELDA(CompRecFa.MSFDET, vbRed, i&, J1%, 1)
          End If
       Next J1%
   Next i&
    '
    
   CompRecFa.GRID.COL = 0
   CompRecFa.GRID.Sort = 3
   Screen.MousePointer = 1
End Sub


Sub MuestraLotesExpress(CI1%, NDEP%)
  If CI1% < 1 Then Exit Sub
  '
  SQLX$ = "SELECT COD_EXTE,IDENLOTE,SUM(CantIngre - CANTSALID) AS CANTI"
  SQLX$ = SQLX$ + ",(SELECT TOP 1 METROS from MOVISTO m2 WHERE m1.IDENLOTE= m2.IDENLOTE   and (codimovi = 'of' or codimovi = 'br') and METROS > 0 AND CANTINGRE > 0)AS METROS"
  SQLX$ = SQLX$ + ",(SELECT TOP 1 PESOKG from MOVISTO m2 WHeRE m1.idEnLoTe= m2.IDEnloTe and (codimovi = 'of' or codimovi = 'br')   and PESOKG > 0 AND CANTINGRE > 0)AS PESOKG"
  SQLX$ = SQLX$ + ",(SELECT TOP 1 FECHMOV from MOVISTO m2 WHeRE m1.idEnLoTe= m2.IDEnloTe and (codimovi = 'of'or codimovi = 'br')) AS FECHMOV"
  SQLX$ = SQLX$ + " FROM  MOVISTO m1 wiTh(NOLOCK) WHERE COD_INTE = " & CI1%
  SQLX$ = SQLX$ + " AND DEPOMOVI = " & NDEP%
  SQLX$ = SQLX$ + " GROUP BY IDENLOTE,Cod_Exte "
  SQLX$ = SQLX$ + " Having Sum(CanTiNGre - CANTSALID) > 0"
  SQLX$ = SQLX$ + " ORDER BY FECHMOV, IDENLOTE "

'SQLX$ = " SELECT COD_EXTE,IDENLOTE,MOVISTO_ID,SUM(CantIngre - CANTSALID) AS CANTI"
'SQLX$ = SQLX$ + " ,METROS,PESOKG,CODIMOVI,FechMov"
'SQLX$ = SQLX$ + " From Movisto WITH(NOLOCK)"
'SQLX$ = SQLX$ + " Where COD_INTE = " & CI1%
'SQLX$ = SQLX$ + " AND DEPOMOVI = " & NDEP%
'SQLX$ = SQLX$ + " GROUP BY IDENLOTE,Cod_Exte,MOVISTO_ID ,METROS,PESOKG ,CODIMOVI,FECHMOV"
'SQLX$ = SQLX$ + " Having Sum(CanTiNGre - CANTSALID) > 0.05"
'SQLX$ = SQLX$ + " ORDER BY FECHMOV"
  


  J& = 0
  ACUMUCANTI = 0
  Set rs = READSET(SQLX$)
  With rs
    Do While Not .EOF
      J& = J& + 1
      SALOTES07.GRID.Rows = J& + 1
      SALOTES07.GRID.TextMatrix(J&, 1) = FECHATEX$(CVINT(!FechMov))
      SALOTES07.GRID.TextMatrix(J&, 2) = SAFETEXT$(!idenlote)
      SALOTES07.GRID.TextMatrix(J&, 3) = ""
      SALOTES07.GRID.TextMatrix(J&, 4) = FORMATNUM$(XVALO(!CANTI), "F12.1")
      ACUMUCANTI = ACUMUCANTI + XVALO(!CANTI)
      SALOTES07.GRID.TextMatrix(J&, 5) = FORMATNUM$(XVALO(!PESOKG), "F10.1")
      SALOTES07.GRID.TextMatrix(J&, 6) = FORMATNUM$(XVALO(!METROS), "F10.1")
      SALOTES07.GRID.TextMatrix(J&, 7) = ""
      .MoveNext
    Loop
  End With
  rs.Close
  Set rs = Nothing
  SALOTES07.Label7.Caption = FORMATNUM$(XVALO(ACUMUCANTI), "F10.1")
  
End Sub

Function PedidosPendientesPorItem(i%, FechaCorte%, PedidosDesde%) As Single

    ' PEDIDOS QUE FUERON PARCIAL O TOTALMENTE ENTREGADOS
    ' LA CONSULTA ES UN LEFT JOIN ENTRE UN RECORDSET DE PEDIDOS Y LA TABLA MOVISTO. LOS PEDIDOS SON LOS EMITIDOS EN EL RANGO DE FECHA ELEGIDO
    ' Y SOLO SE CONSIDERAN LOS ITEMS DE PEDIDOS QUE TIENEN FECHA DE ENTREGA MENOR A LA DEL DIA.
    Dim RstPedidosConEntrega As ADODB.Recordset
    ConsultaSql$ = "SELECT PDD.NumeroPedido, PDD.CodigoInterno, PDD.CantidadPedida, PDD.FechaEntrega, SUM(MOVISTO.CantSalid) AS CantidadSalida, SUM(MOVISTO.CANTINGRE) AS CantidadIngresada FROM" _
                    & " (SELECT PE.NroPed AS NumeroPedido, PD.CodiInt as CodigoInterno, PD.FeSolEnt as FechaEntrega, PD.CanPed as CantidadPedida FROM PEDENCA PE WITH(NOLOCK) INNER JOIN PEDDETA PD WITH (NOLOCK) ON PE.NROPED = PD.NROPED WHERE PD.CODIINT=" & CStr(i%) & " AND PE.FechEmis > '" & FETEXCOR1$(PedidosDesde%) & "' AND PE.FechEmis < '" & FETEXCOR1$(FechaCorte%) & "' AND PD.FeSolEnt <' " & FETEXCOR1$(FechaCorte%) & "' AND PD.STATUS < 8 AND PE.STATUS < 9 GROUP BY PE.NroPed, PD.CodiInt, PD.FeSolEnt, PD.CanPed) PDD" _
                    & " LEFT JOIN MOVISTO WITH(NOLOCK) ON (MOVISTO.NuPedCli = PDD.NumeroPedido AND MOVISTO.Cod_Inte = PDD.CodigoInterno)" _
                    & " WHERE MOVISTO.FECHMOV < '" & FETEXCOR1$(FechaCorte%) & "' AND CODIMOVI='RT' GROUP BY PDD.NumeroPedido, PDD.CodigoInterno, PDD.CantidadPedida, PDD.FechaEntrega"
    Set RstPedidosConEntrega = READSET(ConsultaSql$)
    With RstPedidosConEntrega
        Do While Not .EOF
            CantidadEntregada# = XVALO(!CantidadSalida) - XVALO(!CantidadIngresada)
            If CantidadEntregada# < 0 Then CantidadEntregada# = 0
            
            CantidadPedida# = !CantidadPedida
            PendienteEntregar# = CantidadPedida# - CantidadEntregada#
                            
            CantidadPendiente# = CantidadPendiente# + PendienteEntregar#
            
            .MoveNext
        Loop
    End With
    On Error Resume Next
    RstPedidosConEntrega.Close
    Set RstPedidosConEntrega = Nothing
    On Error GoTo 0

    Call TRACE(24, 2, FIN&)
    ' PEDIDOS QUE NO FUERON ENTREGADOS
    ' LA CONSULTA ES UN LEFT JOIN UN RECORDSET DE PEDIDOS Y LA TABLA MOVISTO: LOS PEDIDOS SON LOS EMITIDOS EN EL RANGO DE FECHA ELEGIDO
    ' Y SE BUSCAN AQUELLOS QUE NO TUVIERON ENTREGAS
    Dim RstPedidosSinEntrega As ADODB.Recordset
    ConsultaSql$ = "SELECT PDD.NumeroPedido, PDD.CodigoInterno, PDD.CantidadPedida, PDD.FechaEntrega, MOVISTO.FECHMOV AS FechaMovimiento, MOVISTO.CantSalid AS CantidadSalida, MOVISTO.CANTINGRE AS CantidadIngresada FROM" _
                    & " (SELECT PE.NroPed AS NumeroPedido, PD.CodiInt as CodigoInterno, PD.FeSolEnt as FechaEntrega, PD.CanPed as CantidadPedida FROM PEDENCA PE WITH(NOLOCK) INNER JOIN PEDDETA PD WITH(NOLOCK) ON PE.NROPED = PD.NROPED WHERE PD.CODiINT=" & CStr(i%) & " AND PE.FechEmis > '" & FETEXCOR1$(PedidosDesde%) & "' AND PE.FechEmis < '" & FETEXCOR1$(FechaCorte%) & "' AND PD.FeSolEnt <' " _
                    & FETEXCOR1$(FechaCorte%) & "' AND PD.STATUS < 8 AND PE.STATUS < 9 GROUP BY PE.NroPed, PD.CodiInt, PD.FeSolEnt, PD.CanPed) PDD LEFT JOIN MOVISTO WITH(NOLOCK) ON (MOVISTO.NuPedCli = PDD.NumeroPedido AND MOVISTO.Cod_Inte = PDD.CodigoInterno)" _
                    '& " GROUP BY PDD.NumeroPedido, PDD.CodigoInterno, PDD.CantidadPedida, PDD.FechaEntrega"
    Set RstPedidosSinEntrega = READSET(ConsultaSql$)

    With RstPedidosSinEntrega
        Do While Not .EOF
            If (IsNull(!CantidadSalida) And IsNull(!CantidadIngresada)) Or CVINT(!FechaMovimiento) > FechaCorte% Then    ' AMBOS CAMPOS EN NULL, NOS DICE QUE SE TRATA DE UN PEDIDO SIN MATERIAL ENTREGADO
                SinEntregar# = !CantidadPedida
                CantidadPendiente# = CantidadPendiente# + SinEntregar#
            End If

            .MoveNext
        Loop
    End With
    On Error Resume Next
    RstPedidosSinEntrega.Close
    Set RstPedidosSinEntrega = Nothing
    On Error GoTo 0

    PedidosPendientesPorItem = CantidadPendiente#
    
End Function

Sub EstadoPedidosAnterioresADesde(EstadoItemDePedido() As Single, FechaCorte%, PedidosDesde%)

    ' PEDIDOS QUE FUERON PARCIAL O TOTALMENTE ENTREGADOS
    ' LA CONSULTA ES UN LEFT JOIN ENTRE UN RECORDSET DE PEDIDOS Y LA TABLA MOVISTO. LOS PEDIDOS SON LOS EMITIDOS EN EL RANGO DE FECHA ELEGIDO
    ' Y SOLO SE CONSIDERAN LOS ITEMS DE PEDIDOS QUE TIENEN FECHA DE ENTREGA MENOR A LA DEL DIA.
    Dim RstPedidosConEntrega As ADODB.Recordset
    ConsultaSql$ = "SELECT PDD.NumeroPedido, PDD.CodigoInterno, PDD.CantidadPedida, PDD.FechaEntrega, SUM(MOVISTO.CantSalid) AS CantidadSalida, SUM(MOVISTO.CANTINGRE) AS CantidadIngresada FROM" _
                    & " (SELECT PE.NroPed AS NumeroPedido, PD.CodiInt as CodigoInterno, PD.FeSolEnt as FechaEntrega, PD.CanPed as CantidadPedida FROM PEDENCA PE WITH(NOLOCK) INNER JOIN PEDDETA PD WITH (NOLOCK) ON PE.NROPED = PD.NROPED WHERE PE.FechEmis > '" & FETEXCOR1$(PedidosDesde%) & "' AND PE.FechEmis < '" & FETEXCOR1$(FechaCorte%) & "' AND PE.STATUS < 3 AND PD.STATUS < 3 AND PD.FeSolEnt <' " & FETEXCOR1$(FechaCorte%) & "' GROUP BY PE.NroPed, PD.CodiInt, PD.FeSolEnt, PD.CanPed) PDD" _
                    & " LEFT JOIN MOVISTO WITH(NOLOCK) ON (MOVISTO.NuPedCli = PDD.NumeroPedido AND MOVISTO.Cod_Inte = PDD.CodigoInterno)" _
                    & " WHERE MOVISTO.FECHMOV < '" & FETEXCOR1$(FechaCorte%) & "' AND CODIMOVI='RT' GROUP BY PDD.NumeroPedido, PDD.CodigoInterno, PDD.CantidadPedida, PDD.FechaEntrega"
    Set RstPedidosConEntrega = READSET(ConsultaSql$)
    
    With RstPedidosConEntrega
        Do While Not .EOF
            CantidadEntregada# = XVALO(!CantidadSalida) - XVALO(!CantidadIngresada)
            If CantidadEntregada# < 0 Then CantidadEntregada# = 0
            
            CantidadPedida# = !CantidadPedida
            CantidadPendiente# = CantidadPedida# - CantidadEntregada#
            
            EstadoItemDePedido(!CodigoInterno) = EstadoItemDePedido(!CodigoInterno) + CantidadPendiente#
            
            .MoveNext
        Loop
    End With
    On Error Resume Next
    RstPedidosConEntrega.Close
    Set RstPedidosConEntrega = Nothing
    On Error GoTo 0

    Call TRACE(24, 2, FIN&)
    ' PEDIDOS QUE NO FUERON ENTREGADOS
    ' LA CONSULTA ES UN LEFT JOIN UN RECORDSET DE PEDIDOS Y LA TABLA MOVISTO: LOS PEDIDOS SON LOS EMITIDOS EN EL RANGO DE FECHA ELEGIDO
    ' Y SE BUSCAN AQUELLOS QUE NO TUVIERON ENTREGAS
    Dim RstPedidosSinEntrega As ADODB.Recordset
    ConsultaSql$ = "SELECT PDD.NumeroPedido, PDD.CodigoInterno, PDD.CantidadPedida, PDD.FechaEntrega, SUM(MOVISTO.CantSalid) AS CantidadSalida, SUM(MOVISTO.CANTINGRE) AS CantidadIngresada FROM" _
                    & " (SELECT PE.NroPed AS NumeroPedido, PD.CodiInt as CodigoInterno, PD.FeSolEnt as FechaEntrega, PD.CanPed as CantidadPedida FROM PEDENCA PE WITH(NOLOCK) INNER JOIN PEDDETA PD WITH(NOLOCK) ON PE.NROPED = PD.NROPED WHERE PE.FechEmis > '" & FETEXCOR1$(PedidosDesde%) & "' AND PE.FechEmis < '" & FETEXCOR1$(FechaCorte%) & "' AND PE.STATUS < 3 AND PD.STATUS < 3 AND PD.FeSolEnt <' " _
                    & FETEXCOR1$(FechaCorte%) & "' GROUP BY PE.NroPed, PD.CodiInt, PD.FeSolEnt, PD.CanPed) PDD LEFT JOIN MOVISTO WITH(NOLOCK) ON (MOVISTO.NuPedCli = PDD.NumeroPedido AND MOVISTO.Cod_Inte = PDD.CodigoInterno)" _
                    & " GROUP BY PDD.NumeroPedido, PDD.CodigoInterno, PDD.CantidadPedida, PDD.FechaEntrega"
    Set RstPedidosSinEntrega = READSET(ConsultaSql$)

    With RstPedidosSinEntrega
        Do While Not .EOF
            If IsNull(!CantidadSalida) And IsNull(!CantidadIngresada) Then    ' AMBOS CAMPOS EN NULL, NOS DICE QUE SE TRATA DE UN PEDIDO SIN MATERIAL ENTREGADO
                CantidadPendiente# = !CantidadPedida
                EstadoItemDePedido(!CodigoInterno) = EstadoItemDePedido(!CodigoInterno) + CantidadPendiente#
                
                If !CodigoInterno = 514 Then
                A = B
                End If
            End If

            .MoveNext
        Loop
    End With
    On Error Resume Next
    RstPedidosSinEntrega.Close
    Set RstPedidosSinEntrega = Nothing
    On Error GoTo 0
    
    
'    Dim RstPedidosSinEntrega0 As ADODB.Recordset
'    ConsultaSql$ = "SELECT PDD.NumeroPedido, PDD.CodigoInterno, PDD.CantidadPedida, PDD.FechaEntrega, SUM(MOVISTO.CantSalid) AS CantidadSalida, SUM(MOVISTO.CANTINGRE) AS CantidadIngresada FROM" _
'                    & " (SELECT PE.NroPed AS NumeroPedido, PD.CodiInt as CodigoInterno, PD.FeSolEnt as FechaEntrega, PD.CanPed as CantidadPedida FROM PEDENCA PE WITH(NOLOCK) INNER JOIN PEDDETA PD WITH(NOLOCK) ON PE.NROPED = PD.NROPED WHERE PE.FechEmis > '" & FETEXCOR1$(PedidosDesde%) & "' AND PE.FechEmis < '" & FETEXCOR1$(FechaCorte%) & "' AND PE.STATUS < 3 AND PD.STATUS < 3 AND PD.FeSolEnt <' " _
'                    & FETEXCOR1$(FechaCorte%) & "' GROUP BY PE.NroPed, PD.CodiInt, PD.FeSolEnt, PD.CanPed) PDD LEFT JOIN MOVISTO WITH(NOLOCK) ON (MOVISTO.NuPedCli = PDD.NumeroPedido AND MOVISTO.Cod_Inte = PDD.CodigoInterno)" _
'                    & " WHERE MOVISTO.FECHMOV > '" & FETEXCOR1$(FechaCorte%) & "' AND CODIMOVI='RT'  GROUP BY PDD.NumeroPedido, PDD.CodigoInterno, PDD.CantidadPedida, PDD.FechaEntrega"
'    Set RstPedidosSinEntrega0 = READSET(ConsultaSql$)
'
'    With RstPedidosSinEntrega0
'        Do While Not .EOF
'                CantidadPendiente# = !CantidadPedida
'                EstadoItemDePedido(!CodigoInterno) = EstadoItemDePedido(!CodigoInterno) + CantidadPendiente#
'
'                If !CodigoInterno = 514 Then
'                A = B
'                End If
'
'            .MoveNext
'        Loop
'    End With
'    On Error Resume Next
'    RstPedidosSinEntrega0.Close
'    Set RstPedidosSinEntrega0 = Nothing
'    On Error GoTo 0

    
End Sub

Function EstadoPedidosDesdeAFechaEntrega(CodigoInterno%, Desde%, Hasta%, PedidosDesde%) As Single

    CantidadPendienteItem# = 0
    
    ' PEDIDOS QUE FUERON PARCIAL O TOTALMENTE ENTREGADOS
    ' LA CONSULTA ES UN LEFT JOIN ENTRE UN RECORDSET DE PEDIDOS Y LA TABLA MOVISTO. LOS PEDIDOS SON LOS EMITIDOS EN EL RANGO DE FECHA ELEGIDO
    ' Y SOLO SE CONSIDERAN LOS ITEMS DE PEDIDOS QUE TIENEN FECHA DE ENTREGA MENOR A LA DEL DIA.
    Dim RstPedidosConEntrega As ADODB.Recordset
    ConsultaSql$ = "SELECT PDD.NumeroPedido, PDD.CodigoInterno, PDD.CantidadPedida, PDD.FechaEntrega, SUM(MOVISTO.CantSalid) AS CantidadSalida, SUM(MOVISTO.CANTINGRE) AS CantidadIngresada FROM" _
                    & " (SELECT PE.NroPed AS NumeroPedido, PD.CodiInt as CodigoInterno, PD.FeSolEnt as FechaEntrega, PD.CanPed as CantidadPedida FROM PEDENCA PE WITH(NOLOCK) INNER JOIN PEDDETA PD WITH (NOLOCK) ON PE.NROPED = PD.NROPED WHERE PE.FechEmis > '" & FETEXCOR1$(PedidosDesde%) & "' AND PE.FechEmis < '" & FETEXCOR1$(Desde%) & "' AND PE.STATUS < 3 AND PD.STATUS < 3 AND PD.FeSolEnt >=' " & FETEXCOR1$(Desde%) & "' AND PD.FeSolEnt <=' " & FETEXCOR1$(Hasta%) & "' AND PD.CODIINT=" & CStr(CodigoInterno%) & " GROUP BY PE.NroPed, PD.CodiInt, PD.FeSolEnt, PD.CanPed) PDD" _
                    & " LEFT JOIN MOVISTO WITH(NOLOCK) ON (MOVISTO.NuPedCli = PDD.NumeroPedido AND MOVISTO.Cod_Inte = PDD.CodigoInterno)" _
                    & " WHERE MOVISTO.FECHMOV < '" & FETEXCOR1$(Hasta%) & "' AND CODIMOVI='RT' GROUP BY PDD.NumeroPedido, PDD.CodigoInterno, PDD.CantidadPedida, PDD.FechaEntrega"
    Set RstPedidosConEntrega = READSET(ConsultaSql$)
    
    With RstPedidosConEntrega
        Do While Not .EOF
            CantidadEntregada# = XVALO(!CantidadSalida) - XVALO(!CantidadIngresada)
            If CantidadEntregada# < 0 Then CantidadEntregada# = 0
            CantidadPedida# = !CantidadPedida
            CantidadPendiente# = CantidadPedida# - CantidadEntregada#
           
            CantidadPendienteItem# = CantidadPendienteItem# + CantidadPendiente#
            
            .MoveNext
        Loop
    End With
    On Error Resume Next
    RstPedidosConEntrega.Close
    Set RstPedidosConEntrega = Nothing
    On Error GoTo 0

    
    ' PEDIDOS QUE NO FUERON ENTREGADOS
    ' LA CONSULTA ES UN LEFT JOIN UN RECORDSET DE PEDIDOS Y LA TABLA MOVISTO: LOS PEDIDOS SON LOS EMITIDOS EN EL RANGO DE FECHA ELEGIDO
    ' Y SE BUSCAN AQUELLOS QUE NO TUVIERON ENTREGAS
    Dim RstPedidosSinEntrega As ADODB.Recordset
    ConsultaSql$ = "SELECT PDD.NumeroPedido, PDD.CodigoInterno, PDD.CantidadPedida, PDD.FechaEntrega, SUM(MOVISTO.CantSalid) AS CantidadSalida, SUM(MOVISTO.CANTINGRE) AS CantidadIngresada FROM" _
                    & " (SELECT PE.NroPed AS NumeroPedido, PD.CodiInt as CodigoInterno, PD.FeSolEnt as FechaEntrega, PD.CanPed as CantidadPedida FROM PEDENCA PE WITH(NOLOCK) INNER JOIN PEDDETA PD WITH(NOLOCK) ON PE.NROPED = PD.NROPED WHERE PE.FechEmis > '" & FETEXCOR1$(PedidosDesde%) & "' AND PE.FechEmis < '" & FETEXCOR1$(Desde%) & "' AND PE.STATUS < 3 AND PD.STATUS < 3 AND PD.FeSolEnt >=' " _
                    & FETEXCOR1$(Desde%) & "' AND PD.FeSolEnt <='" & FETEXCOR1$(Hasta%) & "' AND PD.CODIINT=" & CStr(CodigoInterno%) & " GROUP BY PE.NroPed, PD.CodiInt, PD.FeSolEnt, PD.CanPed) PDD LEFT JOIN MOVISTO WITH(NOLOCK) ON (MOVISTO.NuPedCli = PDD.NumeroPedido AND MOVISTO.Cod_Inte = PDD.CodigoInterno)" _
                    & " GROUP BY PDD.NumeroPedido, PDD.CodigoInterno, PDD.CantidadPedida, PDD.FechaEntrega "
    Set RstPedidosSinEntrega = READSET(ConsultaSql$)

    With RstPedidosSinEntrega
        Do While Not .EOF
            If IsNull(!CantidadSalida) And IsNull(!CantidadIngresada) Then    ' AMBOS CAMPOS EN NULL, NOS DICE QUE SE TRATA DE UN PEDIDO SIN MATERIAL ENTREGADO
                CantidadPendiente# = !CantidadPedida
                
                CantidadPendienteItem# = CantidadPendienteItem# + CantidadPendiente#
            End If

            .MoveNext
        Loop
    End With
    On Error Resume Next
    RstPedidosSinEntrega.Close
    Set RstPedidosSinEntrega = Nothing
    On Error GoTo 0
    
    EstadoPedidosDesdeAFechaEntrega = CantidadPendienteItem#
    
End Function

Function PesoXMetroPromedio#(CI1%)
   PesoXMetroPromedio# = 0
   CONDI$ = "Cod_Inte = " & CI1%
   PesoXMetroPromedio# = XVALO(VALOBADA("BOLRECEP", "AVG(PESO/CANTNOM)", CONDI$, "NOMESS"))

End Function

Sub ReporteCambaLuzuriaga()
100 'SI ES CAMBACERES V
'
'   'Realizado para Cambaceres muestra el stock a una fecha de cambaceres luzuriaga y totaliza
110 OPX% = COMALTER%("Stocks de Bobinas\\Cancelar\Generar Listado\Correr Proceso de Depuraciòn Correctivo")
   If OPX% = 1 Then Exit Sub
   If OPX% = 3 Then
'     Call PROCESO_RECUPERO_DE_KG_Y_METROS_ENCONSUMOS
'     Call PROCESO_RECUPERO_DE_KG_Y_METROS_EN_TRANSFERENCIAS
      Call PROCESO_RECUPERO_DE_KG_Y_METROS
     Call COMUNI("Proceso Terminado")
     GoTo 110
   End If

'   Call PROCESO_RECUPERO_DE_KG_Y_METROS  'SE AGREGO PARA QU LO GENERE CADA VEZ QUE SACA EL REPORTE
   


120  OBX$ = OBJPLA$("ENTRAFECHA1", VDEF$)
    If OBX$ = "" Then
        Des% = 0: Has% = 0: PERIO$ = ""
        Exit Sub
    End If
    VDEF$ = OBX$
    '
    FE% = CVI(Left$(OBX$, 2))
    If FE% <= 30 Then
        GoTo 120
    End If

    FEX = FE%: PERIO$ = FECHATEX$(FEX)


    FECHAX = FETEXCOR1$(FE%)
    '
    Campos$ = "PRODUCTO/A16" '1
    Campos$ = Campos$ + ";Descsripciòn/A24" '2
    Campos$ = Campos$ + ";Camb.U./F11.0" '3
    Campos$ = Campos$ + ";Camb.Mts./F11.0" '4
    Campos$ = Campos$ + ";Camb.Kg./F11.0" '5

    Campos$ = Campos$ + ";Luzur.U./F11.0" '6
    Campos$ = Campos$ + ";Luzur.Mts./F11.0" '7
    Campos$ = Campos$ + ";Luzur.Kg../F11.0" '8

    Campos$ = Campos$ + ";St.Fin.U./F11.0" '9
    Campos$ = Campos$ + ";St.Fin.Mts./F11.0" '10
    Campos$ = Campos$ + ";St.Fin.Kg../F11.0" '11

    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
    J& = 0
    TOTINICIAL = 0
    TOTINGRESO = 0
    TOTEGRESO = 0
    TOTFINAL = 0
    Screen.MousePointer = 11
    TM% = 41
    J& = 0
    TOTUnidadesCambaceres = 0
    TOTMetrosCambaceres = 0
    TOTKgCamaceres = 0

    TOTUnidadesLuzur = 0
    TOTMetrosLuzur = 0
    TOTKgLuzur = 0

    TOTFinalU = 0
    TOTFinalMetros = 0
    TOTFinalKg = 0

    SQLX$ = "select CODINT,CODIGO,Descripcion from master WITH(NOLOCK)"
    SQLX$ = SQLX$ + " where CODINT > 0 AND STATUS >=0"
    If TM% > 0 Then SQLX$ = SQLX$ + " AND TIMATE = " & TM%

    SQLX$ = SQLX$ + " ORDER BY CODIGO"
    Set rs = READSET(SQLX$)
    With rs
    Do While Not .EOF
        CI1% = XVALO(!CODINT)
        Codigo$ = SAFETEXT$(!Codigo)
        DESXRI$ = SAFETEXT$(!Descripcion)
'        OpcionDepoNoIncluir$ = "3,9"
        UnidadesTotales = STODEPFECH(CI1%, 0, FE%, , "")
        MetrosTotales = STODEPFECHENMETROS#(CI1%, 0, FE%, , "")
        KgTotales = STODEPFECHENKG#(CI1%, 0, FE%, , "")
        
        UnidadesLuzur = STODEPFECH(CI1%, 3, FE%) 'STOCK DE LUZURIAGA
        If XVALO(UnidadesLuzur) > 0 Then
            MetrosLuzur = STODEPFECHENMETROS#(CI1%, 3, FE%) 'metros depo luzuriaga
            KgLuzur = STODEPFECHENKG#(CI1%, 3, FE%) 'kg depo luzuriaga
        Else
            MetrosLuzur = 0
            KgLuzur = 0

        End If
            
        UnidadesCambaceres = UnidadesTotales  'TOTAL STOCK MENOS STOCK DE CAMBACERES
        If XVALO(UnidadesCambaceres) > 0 Then
            UnidadesCambaceres = UnidadesCambaceres - XVALO(UnidadesLuzur)
            MetrosCambaceres = MetrosTotales - XVALO(MetrosLuzur) 'TOTAL metros MENOS metros DE CAMBACERES
            KgCamaceres = KgTotales - XVALO(KgLuzur) 'TOTAL kg MENOS kg DE CAMBACERES
        Else
            MetrosCambaceres = 0
            KgCamaceres = 0
        End If


        If XVALO(UnidadesLuzur) > 0 Then
                MetrosLuzur = STODEPFECHENMETROS#(CI1%, 3, FE%) 'metros depo luzuriaga
                KgLuzur = STODEPFECHENKG#(CI1%, 3, FE%) 'kg depo luzuriaga
        Else
                MetrosLuzur = 0
                KgLuzur = 0
    
        End If
        

        FinalU = UnidadesTotales
        FinalMetros = MetrosTotales
        FinalKg = KgTotales
        J& = J& + 1

        GRILLAGRAL.GRID.Rows = J& + 1
        GRILLAGRAL.GRID.TextMatrix(J&, 1) = Codigo$
        GRILLAGRAL.GRID.TextMatrix(J&, 2) = DESXRI$

        GRILLAGRAL.GRID.TextMatrix(J&, 3) = FORMATNUM$(UnidadesCambaceres, "F10.0")
        GRILLAGRAL.GRID.TextMatrix(J&, 4) = FORMATNUM$(MetrosCambaceres, "F10.0")
        GRILLAGRAL.GRID.TextMatrix(J&, 5) = FORMATNUM$(KgCamaceres, "F10.0")

        GRILLAGRAL.GRID.TextMatrix(J&, 6) = FORMATNUM$(UnidadesLuzur, "F10.0")
        GRILLAGRAL.GRID.TextMatrix(J&, 7) = FORMATNUM$(MetrosLuzur, "F10.0")
        GRILLAGRAL.GRID.TextMatrix(J&, 8) = FORMATNUM$(KgLuzur, "F10.0")

        GRILLAGRAL.GRID.TextMatrix(J&, 9) = FORMATNUM$(UnidadesTotales, "F10.0")
        GRILLAGRAL.GRID.TextMatrix(J&, 10) = FORMATNUM$(FinalMetros, "F10.0")
        GRILLAGRAL.GRID.TextMatrix(J&, 11) = FORMATNUM$(FinalKg, "F10.0")

        TOTUnidadesCambaceres = TOTUnidadesCambaceres + XVALO(UnidadesCambaceres)
        TOTMetrosCambaceres = TOTMetrosCambaceres + XVALO(MetrosCambaceres)
        TOTKgCamaceres = TOTKgCamaceres + XVALO(KgCamaceres)

        TOTUnidadesLuzur = TOTUnidadesLuzur + XVALO(UnidadesLuzur)
        TOTMetrosLuzur = TOTMetrosLuzur + XVALO(MetrosLuzur)
        TOTKgLuzur = TOTKgLuzur + XVALO(KgLuzur)

        TOTFinalU = TOTFinalU + XVALO(FinalU)
        TOTFinalMetros = TOTFinalMetros + XVALO(FinalMetros)
        TOTFinalKg = TOTFinalKg + XVALO(FinalKg)

       .MoveNext
    Loop
   End With
   rs.Close
   Set rs = Nothing
   Screen.MousePointer = 1
   'TOTALES
    GRILLAGRAL.GRID.Rows = J& + 3
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 2) = "TOTALES:"
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 3) = FORMATNUM$(XVALO(TOTUnidadesCambaceres), "F10.0")
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 4) = FORMATNUM$(XVALO(TOTMetrosCambaceres), "F10.0")
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 5) = FORMATNUM$(XVALO(TOTKgCamaceres), "F10.0")

    GRILLAGRAL.GRID.TextMatrix(J& + 2, 6) = FORMATNUM$(XVALO(TOTUnidadesLuzur), "F10.0")
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 7) = FORMATNUM$(XVALO(TOTMetrosLuzur), "F10.0")
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 8) = FORMATNUM$(XVALO(TOTKgLuzur), "F10.0")

    GRILLAGRAL.GRID.TextMatrix(J& + 2, 9) = FORMATNUM$(XVALO(TOTFinalU), "F10.0")
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 10) = FORMATNUM$(XVALO(TOTFinalMetros), "F10.0")
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 11) = FORMATNUM$(XVALO(TOTFinalKg), "F10.0")

    '
    GRILLAGRAL.GRID.FixedCols = 3
    Call COLOREAR_GRILLASQL_CadenaDeColumnas(GRILLAGRAL.GRID, &HCDDADA, "6;7;8")
    Call COLOREAR_GRILLA_SOLAFILA(GRILLAGRAL.GRID, &H808080, J& + 2, 0, 1)
   GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
   GRILLAGRAL.mnuMenu1.Visible = True
   GRILLAGRAL.SinEfectoClickEnGrilla% = 1
   GRILLAGRAL.GRID.ZOrder 0
   GRILLAGRAL.TXTBUSCAR = ""
   TX$ = ""
   If TM% > 0 Then
     TX$ = TX$ + " Material: " & TIMAT$
   End If
   If FMLIA$ <> "" Then
     TX$ = TX$ + " Familia: " & FAMI$
   End If

   GRILLAGRAL.Caption = "Stock Cambaceres-Luzuriaga  (" & PERIO$ & ") " & TITULO1$

   GRILLAGRAL.Show 1
   Dim MSF As msFlexGrid

   If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
'        Call CARGA_TABLA_IMPRE(GRILLAGRAL.MSF_2, Campos$, "STOCKRGO", "STOCKS DE CÓDIGOS RANGO", "TABLA_IMPRE", RET_NAMECONS$, 1)

        Call CARGA_TABLA_IMPRE(MSF, Campos$, "STCAMLUZ", GRILLAGRAL.Caption, "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
        Exit Sub
   End If
   GoTo 120
End Sub





Sub TablasAuxiliares()

    Dim TablaAuxiliar As New DatosAuxiliares
    Dim ListaCampos() As eDatosAuxiliaresCampos
    Dim CAMPO As eDatosAuxiliaresCampos
    
    ' ESTADISTICAS - TABLA PRODUCTOS
    ' ******************************
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
    CAMPO.Encabezado = "Codigo/A24"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = "codigo;descripcion;master;status>0;order by codigo"
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
        
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Descripción/A48"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Call TablaAuxiliar.CrearTabla("ANALISTO", "", ListaCampos)
    
End Sub

Sub Trazador(ORDEN&, FIN&)

    Call TRACE(24, ORDEN&, FIN&)
    
End Sub

Function UBISTOSQL$(CIII%, NPRO)
    CONDI$ = "COD_INTE = " & CIII%
    CONDI$ = CONDI$ + " AND Nume_Pro = " & NPRO
    '
    UBICACION$ = ""
    SQLX$ = " SELECT CODIUBICA FROM UBICASQL WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE " & CONDI$
    Set rs = READSET(SQLX$)
    With rs
      Do While Not .EOF
         UBIC$ = SAFETEXT$(!CODIUBICA)
         If TRIM$(UBIC$) <> "" Then
            UBIC$ = UBIC$ + ", "
         End If
         UBICACION$ = UBICACION$ + UBIC$
        .MoveNext
      Loop
    End With
    '
    If Right(UBICACION$, 2) = ", " Then UBICACION$ = Left$(UBICACION$, Len(UBICACION$) - 2)
    UBISTOSQL$ = UBICACION$
   '
End Function

Sub VERUBICA(CIII%, PROVED)
    Call CARUBICSQL(CIII%, PROVED)
    '
End Sub

'Sub CARUBICSQL(CIII%, NPRO)
'    '
'    If CIII% <= 0 Then
'        Call MENSERR(24, "Código de Artículo no Válido")
'        Exit Sub
'    End If
'    '
'    If NPRO < 1 Then
'      Call Label13_DblClick
'      Exit Sub
'    End If
'    '
''    If NPRO <= 0 Then
''        Call MENSERR(24, "Proveedor no Válido")
''        Exit Sub
''    End If
'    '
'    Screen.MousePointer = 11
'    Call BLANARCH("!CARUBIC")
'    Call SETULTREG("!CARUBIC", 0)
'    '
'    K1& = 0
'    sqlx$ = "SELECT * FROM UBICASQL  WITH(NOLOCK) "
'    sqlx$ = sqlx$ + " WHERE COD_INTE = " & CIII%
'    sqlx$ = sqlx$ + " AND Nume_Pro = " & NPRO
'    sqlx$ = sqlx$ + " ORDER BY CODIUBICA"
'    Set RS = READSET(sqlx$)
'    With RS
'      Do While Not .EOF
'         CODUB$ = TRIM$(SAFETEXT$(!CODIUBICA))
'         Observa$ = SAFETEXT$(!Observacion)
'         If CODUB$ <> "" Then
'            Call REPLA(TX$, MKI$(CIII%), 1, 2)
'            Call REPLA(TX$, CODUB$, 3, 24)
'            Call REPLA(TX$, MKS$(0), 27, 4)
'            Call REPLA(TX$, Observa$, 31, 30)
'            Call REPLA(TX$, MKS$(NPRO), 61, 4)
'            Call REGAPP("!CARUBIC", TX$)
'         End If
'         .MoveNext
'      Loop
'    End With
'    RS.Close
'    Set RS = Nothing
'
'    Screen.MousePointer = 1
'    '
'    VTB% = VENTABU%("CARUBIC/BORDEIZQ=20/BORDESUP=12/NC/TITUPAN=" + TRIM$(CODEXT$(CIII%)) + " - " + DESCRIT$(CIII%))
'    '
'    If VTB% >= 0 Then
'        CONDI$ = "COD_INTE = " & CIII%
'        CONDI$ = CONDI$ + " AND Nume_Pro = " & NPRO
'        '
'        Call ACTUREGISTRO("UBICASQL", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")
'
'        Screen.MousePointer = 11
'        K& = ULTREG&("!CARUBIC")
'        K2& = 0
'        For I& = 1 To K&
'            TX$ = REGLEIDO$("!CARUBIC", I&)
'            If TRIM$(Mid$(TX$, 3, 24)) <> "" Then
'                '
'                CODUB$ = TRIM$(Mid$(TX$, 3, 24))
'                Observa$ = TRIM$(Mid$(TX$, 25, 36))
'                'LLAMO A STORED PROCEDURE PARA DAR DE ALTA UBICACION.
'                sqlx$ = "SPX_ADDUBICACION  " & CIII%
'                sqlx$ = sqlx$ + ", " & CodiEmp%
'                sqlx$ = sqlx$ + ", '" & CODUB$ & "'"
'                sqlx$ = sqlx$ + ", 0"
'                sqlx$ = sqlx$ + ", '" & Observa$ & "'"
'                sqlx$ = sqlx$ + ", " & NPRO
'                sqlx$ = sqlx$ + ", 0"
'                FLEXCONN.Execute sqlx$
'            End If
'            '
'        Next I&
'        '
'        CONDI$ = CONDI$ + " AND MARBORRA = 1"
'        Call BORRAREGISTROS("UBICASQL", CONDI$, REGAF&, "NOMESS")
'    End If
'    '
'    Call CIERRARCH("UBISTOCK")
'    Screen.MousePointer = 1
'    '
'End Sub
''

Private Sub Abrir_Click()
    Call Command2_Click
End Sub

Sub asignalotesitemfacturados(SinLote%)
    
10  Campos$ = "Fecha/D8;Documento/A22;Cta./F9;R.Social/A32"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO, , , ANTOT)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO, , , ANTOT)
    If ANTOT < 5500 Then ANTOT = 5500
    FRMZELECHO.Width = ANTOT + 1500
    '
    Screen.MousePointer = 11
    'EN ESTA CONSULTA TRAIGO TODO LO QUE SE INGRESO POR LOTE Y SALIO POR LOTE
    FRMZELECHO.MSF1.Rows = 1
    J& = 0
    ACUMLA_SALDO = 0
    SQLX$ = "SELECT  M.FECHMOV, M.DOSECLAR,  M.NUME_CLIE, M.COD_INTE,   M.VALOUNITIVA FROM MOVISTO M WITH(NOLOCK)"
    SQLX$ = SQLX$ + " INNER JOIN MASTER A ON M.COD_INTE = A.CODINT"
    SQLX$ = SQLX$ + " WHERE  M.CODIMOVI = 'RT'"
    If SinLote% > 0 Then SQLX$ = SQLX$ + " AND ( M.IDENLOTE IS NULL OR  M.IDENLOTE = '') "
    SQLX$ = SQLX$ + " AND A.TRAZABLE = 1"
    
'    SQLX$ = SQLX$ + " GROUP M.COD_INTE M.FECHMOV,M.COD_INTE DOSECLAR, COD_INTENUME_CLIE "
    SQLX$ = SQLX$ + " ORDER BY M.FECHMOV, M.DOSECLAR, M.COD_INTE ASC"
    Set RS1 = READSET(SQLX$)
    With RS1
        Do While Not .EOF
           DOCU$ = SAFETEXT$(!DOSECLAR)
           If DOCUORIG_ANT$ <> DOCU$ Then
             CI1% = XVALO(!cod_inte)
             FE% = CVINT(!FechMov)
             NCL = XVALO(!NUME_CLIE)
             RASOC$ = rasocli$(NCL)
             '
             J& = J& + 1
             FRMZELECHO.MSF2.Rows = J& + 1
             FRMZELECHO.MSF2.TextMatrix(J&, 1) = FECHATEX$(FE%)
             FRMZELECHO.MSF2.TextMatrix(J&, 2) = DOCU$
             FRMZELECHO.MSF2.TextMatrix(J&, 3) = NCL
             FRMZELECHO.MSF2.TextMatrix(J&, 4) = RASOC$
             CI_ANT% = XVALO(CI1%)
             DOCUORIG_ANT$ = DOCU$
           End If
          .MoveNext
        Loop
    End With
    RS1.Close
    Set RS1 = Nothing
    Screen.MousePointer = 1
    FRMZELECHO.Show 1
    
    FACTURAX$ = RESP_ZELE_VECT(2)
    TRAZCOAHP.lblfactura = ""
    If TRIM$(FACTURAX$) <> "" Then
      TRAZCOAHP.LBLCLIENTE = RESP_ZELE_VECT(4)
      TRAZCOAHP.lblfactura = FACTURAX$
      TRAZCOAHP.Show 1
      GoTo 10
    End If
End Sub

Private Sub Ayud_Click()
    Call Command5_Click
End Sub

Private Sub Command1_Click()
    IDTX$ = IDENTER$
    AVALSTOK$ = "LISTAPRE"
    If Option2.Value = True Then
       AVALSTOK$ = "MAESTRO"
    End If
    Call GRACONTROL(IDTX$, "AVALSTOK", AVALSTOK$)
    Call CIERRARCH("*.*")
    If Left$(App.EXEName, 6) = "ESTOCK" Then
        Hide
      Else
        Call FINAL("")
    End If
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   CONPED07.Show 1
   Command10.Enabled = True
End Sub



Private Sub Command11_Click()
'Sub MIGRASTOCKS()
  'ESTE BOTON SE AGREGO PARA  MIGRAR LA TABLA MOVISTO HACIA MIGRAREPA PARA DOSIVAC
  '\\\OT-08-0819-OD-08/09/08///
  Exit Sub
  '
  Command11.Enabled = False
  Call ABRECONEXION
  'Picture1.ToolTipText = "Del - Movimientos de Stock"
  On Error Resume Next
   FLEXCONN.Execute (FILTSQL$("DROP TABLE MOVIREPA"))
  On Error GoTo 0
  Call GENERBASE("STOCKS")
  Call APERBASDAT("STOCKS")
  '
  'Picture1.ToolTipText = "Conversión de Movimientos de Stocks"
  '
  Screen.MousePointer = 11

  FIN& = ULTREG&("MOVIREPA")
  DOCANT$ = ""
  PREFIREM$ = TRIM$(Control$("", "PREFIREM")): If PREFIREM$ = "" Then PREFIREM$ = "0001"
  '
  SQLX$ = "SELECT * FROM MOVIREPA "
  Dim MOVIREPATMP As ADODB.Recordset
  Set MOVIREPATMP = New ADODB.Recordset
  MOVIREPATMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
  '
  For U& = 1 To FIN&
    X$ = REGLEIDO$("MOVIREPA", U&)
    Call TRACE(20, U&, FIN&)
    'Label1 = U&
    'LECTURA
      FEC% = CVI(Mid$(X$, 1, 2)) 'FECHA DE MOVIMIENTO
      CI% = CVI(Mid$(X$, 3, 2)) 'CODIGO DE PIEZA
      LOT$ = Mid$(X$, 5, 16) 'LOTE / PARTIDA
      COMO$ = Mid$(X$, 21, 2) 'CODIGO DE MOVIMIENTO
      DOC1$ = Mid$(X$, 23, 10) 'DOCUMENTO PRIMARIO ASOCIADO
      DOC1LAR$ = DOC1$
      If COMO$ = "RP" Then
         NUMEDO$ = TRIM$(Mid$(DOC1$, 4))
         While Len(NUMEDO$) < 8: NUMEDO$ = "0" + NUMEDO$: Wend
         DOC1LAR$ = Left$(DOC1$, 2) + "-X-" + PREFIREM$ + "-" + NUMEDO$
      End If
      DOC2$ = Mid$(X$, 33, 10) 'DOCUMENTO SECUNDARIO
      DOC2LAR$ = DOC2$
      If Left$(DOC2$, 2) = "FC" Then
         NUMEDO$ = TRIM$(Mid$(DOC2$, 4))
         While Len(NUMEDO$) < 8: NUMEDO$ = "0" + NUMEDO$: Wend
         DOC2LAR$ = Left$(DOC2$, 2) + "-0001-" + NUMEDO$ + "-A"
      End If
      REFE$ = Mid$(X$, 43, 30) 'REFERENCIA
      VAUN# = CVD(Mid$(X$, 73, 8)) 'VALOR UNITARIO
      MONE1$ = Mid$(X$, 81, 2) 'MONEDA
      DEP% = Asc(Mid$(X$, 83, 1)) 'DEPOSITO
      NUCLI% = CVI(Mid$(X$, 84, 2)) 'NUEMRO DE CLIENTE
      NUPRO% = CVI(Mid$(X$, 86, 2)) 'NUMERO DE PROVEEDOR
      CAEN = CVD(Mid$(X$, 88, 8)) 'CANTIDAD ENTRADA
      CASA = CVD(Mid$(X$, 96, 8)) 'CANTIDAD SALIDA
      CASALD = CVD(Mid$(X$, 104, 8)) 'CANTIDAD SALDO
      Fill$ = Mid$(X$, 112, 10) 'FILLER NO IDENTIFICADO
      FEMO% = CVI(Mid$(X$, 122, 2)) 'FECHA DE MOVIMIENTO
      USU% = Asc(Mid$(X$, 126, 1)) 'USUARIO
    'GRABACIÓN
    With MOVIREPATMP
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CODIEMPR = CodiEmp%
      On Error GoTo 0
        !FechMov = CVDAT(FEC%)
        !cod_inte = CI%
        !Cod_Exte = " " 'CODEXT$(CI%) - eliminado para ganar velocidad
        !Descrip = " " 'DESCRIT0(CI%) - eliminado para ganar velocidad
        !idenlote = LOT$
        !CODIMOVI = COMO$
        !dopricor = DOC1$
        !DOPRILAR = DOC1LAR$
           If DOC1$ <> DOCANT$ Then
             NORI% = 0
             DOCANT$ = DOC1$
           End If
        !DoSecCor = DOC2$
        !DOSECLAR = DOC2LAR$
        !NumRem_Cli = "" 'POR CONVERSIÓN
           NORI% = NORI% + 1
        !NuOrItem = NORI%
        !NUPEDCLI = CVS(Left$(Fill$, 4))      'POR CONVERSIÓN
          NOPXX% = CVI(Mid$(Fill$, 5, 2))
          If NOPXX% = 8224 Then NOPXX% = 0
        !NUMEORDPED = NOPXX%
        !ReferCor = REFE$
        !VALOUNIT = VAUN
        !MoneVal = XVALO(MONE1$)
        !depomovi = DEP%
        !NUME_CLIE = NUCLI%
        !Nume_Prov = NUPRO%
        !CANTINGRE = CAEN
        !CANTSALID = CASA
        !CantSaldo = CASALD
        !FeReMovi = CVDAT(FEMO%)
      .Update
    End With
    '
    Call TRACE(20, U&, FIN&)
  Next U&
  MOVIREPATMP.Close
  Set MOVIREPATMP = Nothing
  '
  Screen.MousePointer = 1
  Command11.Enabled = True

  '
  '\\\OT-08-0819-OD-FIN///
  '
End Sub
Private Sub Command12_Click()
    '\\\OT-08-0839-OD-12/08/08///
    '
    'ESTE BOTON SE AGREGO PAARA PASAR LOS DATOS DE REPARACIONES QUE ESTABAN EN EL MOVISTO AL MOVIREPA
    'PARA DOSIVAC
    
    Exit Sub
    Command12.Enabled = False
    Screen.MousePointer = 11
    FLIM$ = FETEXCOR1$(FECHANUM("22/08/08"))
    '
    Call ABRESTOCKS
    'CONSULTA AL MOVISTO CON MOVIMIENTOS SUPERIORES O IGUALES A 22/08/08
    '
    SQLX$ = "SELECT * FROM MOVISTO"
    SQLX$ = SQLX$ + " WHERE FechMov >= '" & FLIM$ & "'"
    SQLX$ = SQLX$ + " AND NuPedCli > 0"
    '
    Dim MOVITMP As ADODB.Recordset
    Set MOVITMP = New ADODB.Recordset
    MOVITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With MOVITMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
          NumPedido& = XVALO(!NUPEDCLI)
          ResObtenido = VALOBADA("PEDENCA", "VENREP", "NroPed = " & NumPedido&)
          If ResObtenido = "R" Then 'SOLO LOS QUE SON REPARACIONES
            Dim VALORCA(128), NOMBREDECAMPO(128) 'DIMENSIONE VECTORES
            CANCAMPOS = .Fields.Count 'CUENTO LA CANTIDAD DE CAMPOS
            For i% = 1 To CANCAMPOS 'LOS RECORRO
              NOMBREDECAMPO(i%) = .Fields(i% - 1).Name 'VOY ASIGNANDO EN EL VECTOR LOS NOMBRES DE LOS CAMPOS
              VALORCA(i%) = .Fields(i% - 1).Value 'ASIGNO EL CONTENIDO DE CADA CAMPO EN EL OTRO VECTOR
            Next i%
            'ABRO EL RECORDSET DE MOVIREPA
            SQLX$ = "SELECT * FROM MOVIREPA" 'HABRO EL MOVIREPA PARA REALIZAR EL ADDNEW
            Dim RS1 As ADODB.Recordset
            Set RS1 = New ADODB.Recordset
            RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
            With RS1
              .AddNew   ' YA CONVERTIDO
              On Error Resume Next
              !CODIEMPR = CodiEmp%
              On Error GoTo 0
                For i% = 1 To CANCAMPOS 'DE 1 A LA CANTIDAD DE CAMPOS
                  NOCAM$ = NOMBREDECAMPO(i%) 'ASIGNO EL NOMBRE DEL CAMPO
                  For J% = 1 To .Fields.Count
                    If .Fields(J% - 1).Name = NOCAM$ Then ' SI COINCIDE EL NOMBRE CAMPO CON NOCAM$
                      .Fields(J% - 1).Value = VALORCA(i%) 'A CADA CAMPO LE ASIGNO EL VALOR
                    End If
                  Next J%
                Next i%
              .Update
              MOVITMP!dopricor = "MIGRADO" 'MARCA PARA DESPUES BORRAR DEL MOVISTO LOS MIGRADOS
              MOVITMP.Update
            End With
          End If
          .MoveNext
        Loop
      End If
      On Error GoTo 0
    End With
    'ELIMINO LOS MOVIMIENTOS  DEL MOVISTO PASADOS AL MOVIREPA
    '
    SQLX$ = "DELETE  FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE DOPRICOR = 'MIGRADO'"
     FLEXCONN.Execute (FILTSQL$(SQLX$))
    '
    Command12.Enabled = True
    Screen.MousePointer = 1

    Call COMUNI("Proceso Terminado")
    
    '\\\OT-08-0839-OD-FIN///
End Sub




Private Sub Command13_Click()
 '\\\OT-1069-FL-24/11/08/// MUESTRA FORMULARIO FORDACOD07 CON LOS ATRIBUTOS, LA IMAGEN Y
 'LAS ESPECIFICACIONES DEL CODIGO DEL ARTICULO SELECCIOANDO EN LA SITUACION DE STOCK
 CODXX$ = Text1.TEXT
 CODII% = CODINT%(CODXX$)
 If CODII > 0 Then
    Call MUESTRADATCOD(CODXX$)
 End If
 '\\\OT-1069-FL-FIN///
End Sub

Private Sub Command14_Click()
    Command14.Enabled = False
    Mensaje$ = REPLACAR(LUDAT$ + "BDACCESS\STOCKVIE.mdb", "\", "/") 'esta variable(mensaje$) tambien se utiliza al final del codigo
    Screen.MousePointer = 11
    If EXISTE%(LUDAT$ + "BDACCESS\STOCKVIE.mdb") <> 1 Then
      Call COMUNI("Archivo " & Mensaje$ & "\Inexistente")
      GoTo 99
    End If
    JJ& = 0
    Dim BASE As DAO.Database
    Dim rs As DAO.Recordset
    
    Set BASE = OpenDatabase(LUDAT$ + "BDACCESS\STOCKVIE.mdb")
    '
    Dim RECSET As DAO.Recordset
    Set RECSET = BASE.OpenRecordset("SELECT COUNT(*)AS CANTREG FROM MOVISTO ", dbOpenSnapshot)
    If XVALO(RECSET!CANTREG) < 1 Then GoTo 99
    '
    On Error Resume Next
     FLEXCONN.Execute (FILTSQL$("DROP TABLE MOVIVIE"))
    On Error GoTo 0
    '
    'Call GENERBASE("STOCKVIE")
    Call GENERBASE("STOCKS")
    Call ABRESTOCKS
    '
    Set rs = BASE.OpenRecordset("SELECT * FROM MOVISTO ", dbOpenSnapshot)
    With rs
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
           FIN& = .RecordCount
           JJ& = JJ& + 1
           Call TRACE(20, JJ&, FIN&)
            Dim STONUE As ADODB.Recordset
            Set STONUE = New ADODB.Recordset
            SQLX$ = "SELECT * FROM MOVIVIE WHERE Cod_Exte = 'A'" ' PARA QUE NO TRAIGA DATOS
            STONUE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
            With STONUE
             .AddNew   ' YA CONVERTIDO
             On Error Resume Next
             !CODIEMPR = CodiEmp%
             On Error GoTo 0
             !FechMov = rs!FechMov
             !cod_inte = XVALO(rs!cod_inte)
             !Cod_Exte = SAFETEXT$(rs!Cod_Exte)
             !Descrip = SAFETEXT$(rs!Descrip)
             !idenlote = SAFETEXT$(rs!idenlote)
             !CODIMOVI = SAFETEXT$(rs!CODIMOVI)
             !dopricor = SAFETEXT$(rs!dopricor)
             !DOPRILAR = SAFETEXT$(rs!DOPRILAR)
             !DoSecCor = SAFETEXT$(rs!DoSecCor)
             !DOSECLAR = SAFETEXT$(rs!DOSECLAR)
             !NuOrItem = XVALO(rs!NuOrItem)
             !NUPEDCLI = XVALO(rs!NUPEDCLI)
             !ReferCor = SAFETEXT$(rs!ReferCor)
             !VALOUNIT = XVALO(rs!VALOUNIT)
             !MoneVal = XVALO(rs!MoneVal)
             !depomovi = XVALO(rs!depomovi)
             !NUME_CLIE = XVALO(rs!NUME_CLIE)
             !Nume_Prov = XVALO(rs!Nume_Prov)
             !CANTINGRE = XVALO(rs!CANTINGRE)
             !CANTSALID = XVALO(rs!CANTSALID)
             !CantSaldo = XVALO(rs!CantSaldo)
             !FeReMovi = rs!FeReMovi
             !NumUsuar = XVALO(rs!NumUsuar)
             .Update
            End With
            STONUE.Close
            Set STONUE = Nothing
           'End If
         .MoveNext
         Loop
       End If
       On Error GoTo 0
   End With
   rs.Close
   Set rs = Nothing
   '
   Call COMUNI("El Proceso de Migración ha Finalizado con Exito\Se Deberá  Borrar o Eliminar el Archivo STOCKVIE.mdb\Ubicado en la Siguiente Ruta :\" & Mensaje$ & "\Para Evitar Sobreescribir la Migración Realizada")
99 Command14.Enabled = Enabled
   Screen.MousePointer = 1

End Sub


Private Sub Command15_Click()
    Command15.Visible = False
    ACONEC$ = "DECAOT09"
    Call CONECRUN(ACONEC$, "Depósito Caotico de Materiales")
    Command15.Visible = True
End Sub

Private Sub Command16_Click()

    If DERACCE%("RECUNUSE", "Recuperacion de Numeros de Serie") < 1 Then
       Exit Sub
    End If
    '
    SQLX$ = "select * from movisto inner join master "
    SQLX$ = SQLX$ + " on movisto.cod_inte = master.codint"
    SQLX$ = SQLX$ + " where codimovi = 'RT'"
    SQLX$ = SQLX$ + " and master.trazable=2"
    SQLX$ = SQLX$ + " and movisto.idenlote = ''"
    '
    Dim MOVITMP As ADODB.Recordset
    Set MOVITMP = New ADODB.Recordset
    MOVITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With MOVITMP
      Do While Not .EOF
         JJ& = JJ& + 1
         Call TRACE(20, JJ&, QTRECU&)
         CIXI% = !cod_inte
         NPED = !NUPEDCLI
         NUREMI$ = !DOPRILAR
         NUFAC$ = !DOSECLAR
         NUPED = !NUPEDCLI
         NUSERI$ = VALOBADA("TANUMSE", "NUMEROSERIE", "NUMREMITO = '" & NUREMI$ & "' AND COD_INTE = " & CIXI%)
         If TRIM$(NUSERI$) = "" Then
            NUSERI$ = VALOBADA("TANUMSE", "NUMEROSERIE", "NUMFACTURA = '" & NUFAC$ & "' AND COD_INTE = " & CIXI%)
         End If
         If TRIM$(NUSERI$) = "" Then
            NUSERI$ = VALOBADA("PEDDETA", "NUMEROSERIE", "NROPED = " & NUPED & " AND CODIINT = " & CIXI%)
         End If
         If TRIM$(NUSERI$) <> "" Then
            !idenlote = NUSERI$
            .Update
            CONTA2% = CONTA2% + 1
         End If
      .MoveNext
      Loop
    End With
    '
    Call COMUNI("Proceso de Recuperacion Finalizado")
    '
End Sub

Private Sub Command17_Click()
   '
   Command17.Enabled = False
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   'If VERILICEN%("STOCOMPO", "Consulta Stock de Componentes por Fórmula") > 0 Then
   'If InStr(EMPREAC$, "GABAL") > 0 Then
      'VALIDACION QUE TENGA SELECCIONADO UN ARTICULO
      If TRIM$(Text1) = "" Then
         Call MENSERR(24, "Falta Seleccionar Código"): GoTo 99
      End If
      '
      ' VALIDACION DEL CODIGO
      CIXI% = CODINT%(Text1)
      If CIXI% < 1 Then
         Call MENSERR(24, "Codigo Inexistente o no Valido"): GoTo 99
      End If
      '
      'VALIDACION QUE TENGA ESTRUCTURA
      Call LEEEXPLO(TRIM$(Text1), 1)
      If CAIT% < 1 Then
         Call MENSERR(24, "Código Seleccionado No Posee Estructura"): GoTo 99
      End If
      '
      If InStr(1, EMPREAC$, "ARCON") > 0 Then
        'PRESENTA FRM CON INFORMACION DE STOCKS DE LOS COMPONENTES
        
        STOCOMXDE.Label4 = Text1
        STOCOMXDE.Show 1
      Else
        STOCOMPO07.Label4 = Text1
        STOCOMPO07.Show 1
      End If
   'End If
   '
99 Command17.Enabled = True

End Sub


Private Sub Command19_Click()
   Command19.Enabled = False
   If DERACCE%("ESTACONS", "Analisis Estadisticos - Consumos de Stocks") < 2 Then GoTo 99
   
   If EXISTE%("ESTOCK07.EXE") Then '
     CI1% = CODINT%(Text1)
     Call CONECRUN("ESTOCK07/CONEST" & CI1%, "MODULO DE ESTADISTICAS DE STOCKS")
   End If
99 Command19.Enabled = True
End Sub

Private Sub Command2_Click()
    '
    ASEL$ = "MASTER"
    If Option1.Value = True Then
       ASEL$ = "LIPRE001"
       Call SELECHO(ASEL$)
       CODD$ = VALACT1$(ASEL$)
       GoTo 20
    End If
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "AHP") > 0 Then
     Call ZELECHOAHP(1)
     CODD$ = RESP_ZELE_VECT(2)
    Else
     Call SELECHO("MASTER")
     CODD$ = TRIM$(VALACT1$("MASTER"))
    End If
    
    If CODD$ = "" Then
        Exit Sub
    End If
    '

20  COD$ = CODD$
    Text1.TEXT = COD$
    CI% = CODINT%(COD$)
    Text3.TEXT = DESCRIT$(CI%)
    Text2.TEXT = Unimed$(CI%)
    '
End Sub

Private Sub Command20_Click()
   Command20.Enabled = False
   If DERACCE%("ESTACONS", "Analisis Estadisticos - Consumos de Stocks") < 2 Then GoTo 99
   
   If EXISTE%("ESTOCK07.EXE") Then '
     CI1% = CODINT%(Text1)
     Call CONECRUN("ESTOCK07/CONEST" & CI1%, "MODULO DE ESTADISTICAS DE STOCKS")
   End If
99 Command20.Enabled = True
End Sub




Private Sub Command27_Click()
    Command27.Enabled = False
'    If (USNBR% Mod 100) = 1 Then
    If DERACCE%("STPSTOCK", "Configuración de Funcionamiento de Stock") >= 2 Then
         OPDISTK07.Show 1
       Else
         Call MENSERR(24, "La Opción Elegida Requiere Privilegios de Supervisor")
    End If
    Command27.Enabled = True
End Sub

Sub Command3_Click()
   Command3.Enabled = False
   FIKARDEX09.Text1 = Text1
   FIKARDEX09.Show 1
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    Call GRAPHDIS
    SITCOB09.Show 1
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Call MENSERR(24, "No se Encuentra Ayuda Específica")
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   MOVISTK07.Show 1
   If CODINT%(Text1.TEXT) > 0 Then
      COD$ = TRIM$(Text1.TEXT)
      CI% = CODINT%(COD$)
      Text2.TEXT = Unimed$(CI%)
      Text3.TEXT = DESCRIT$(CI%)
      Call MUEDIS(CI%)                ' recalcular existencias
   End If
   Command6.Enabled = True
End Sub

Private Sub Command7_Click()
   Static CTXX%, MODOCORE%, FECACOS$
   '
   If DERACCE%("VERCOSTO", "Visualizar Costos y Precios") < 1 Then
      Exit Sub
   End If
   '
   CODD$ = TRIM$(Text1.TEXT)
   CIXX% = CODINT%(CODD$)
   If CIXX% <= 0 Then
      Exit Sub
   End If
   '
   If MODOCORE% < 1 Then
      MODOCORE% = 1
      MCDR$ = TRIM$(UCase$(Control$("SITUASTO", "MODOCORE")))
      If MCDR$ = "PREBASE" Then
           MODOCORE% = 2
         ElseIf MCDR$ = "COSCALCU" Then
           MODOCORE% = 3
           If FECACOS$ = "" Then
             FECACOS$ = TRIM$(Control("ANACOSTO", "FECALCOS"))
           End If
      End If
   End If
   '
   MUEPRECIO.List1.Clear
   ZZ$ = "   C.Reposición" + Space$(30): coefi = 0
   If MODOCORE% = 2 Then
        ZZ$ = "Precio Base" + Space$(30)
        coefi = 15
     ElseIf MODOCORE% = 3 Then
        ZZ$ = "Cos.Calcul.(" + FECACOS$ + ")" + Space$(30)
   End If
   '
   COPRES = COSUNI(CIXX%) * (1 + coefi / 100)
   If MODOCORE% = 3 Then COPRES = COSCALCU(CIXX%, MODOCO%)
   Call REPLA(ZZ$, CSTRING$(MKS$(COPRES), 4, 11, 4, 2), 25, 11)
   MUEPRECIO.List1.AddItem ZZ$
   MUEPRECIO.List1.AddItem Space$(80)    'String$(80, 45)
   '
   For U& = 1 To ULTREG&("LISTAS")
      ZZ$ = Mid$(REGLEIDO$("LISTAS", U&), 2, 32)
      If TRIM$(ZZ$) <> "" Then
         ZZ$ = AJUSTD$(TRIM$(Str$(U&)), 2) + "-" + ZZ$
         CLIS$ = TRIM$(Str$(U&))
         While Len(CLIS$) < 3: CLIS$ = "0" + CLIS$: Wend
         APREC$ = "LIPRE" + CLIS$
         PRLST# = PRELISTA#(APREC$, CIXX%)
         If PRLST# >= 0.00005 Then
           Call REPLA(ZZ$, CSTRING$(MKD$(PRLST#), 4, 11, 4, 2), 25, 11)
           MUEPRECIO.List1.AddItem ZZ$
         End If
      End If
   Next U&
   '
   MUEPRECIO.Show 1
   '
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   ESENVAPLAST% = 0
   If InStr(1, EPAC$, "ENVAPLAST") > 0 Then ESENVAPLAST% = 1
   FORIPRE$ = TRIM$(Control$("", "FORDISTO"))
   If FORIPRE$ = "" Then
        'PRINTFORM
        Campos$ = "Comprobante/A18;Referencia/A50;Entra/F10.1;Sale/F10.1;Proyectado/F10.1;Fecha/D8;Observaciones/A40"
        Call CARGA_ENCABEZADO(mfgStockProyectado, Campos$, FRMZELECHO)
        For fi = 1 To gridStockProyectado.Rows - 1
            mfgStockProyectado.Rows = mfgStockProyectado.Rows + 1
            For co = 0 To gridStockProyectado.Cols - 2
                mfgStockProyectado.TextMatrix(fi, co + 1) = gridStockProyectado.TextMatrix(fi, co)
            Next co
        Next fi
        Call CARGA_TABLA_IMPRE(mfgStockProyectado, Campos$, "STOPROYE", "STOCK PROYECTADO PARA (" & Text1 & " - " & Text3 & ")", "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        mfgStockProyectado.Rows = 1
      Else
        Screen.MousePointer = 11
        FEX = HOY(HOS$)
        FECHDOC$ = HOS$
        NUMEDOC$ = ""
        TIPODOC$ = "Situacion de Stocks"
        '
        CODPARAM$(1) = "COD-MAT"
        CODPARAM$(2) = "DES-MAT"
        CODPARAM$(3) = "CANTIDAD"
        CODPARAM$(4) = "CANT-MIN"
        CODPARAM$(5) = "CANT-RES"
        CODPARAM$(6) = "CANT-ING"
        CODPARAM$(7) = "CANT-DIS"
        CODPARAM$(8) = "CANT-FAL"
        CAPARDOC% = 8
        '
        DOCPARAM$(1) = TRIM$(Text1.TEXT)
        DOCPARAM$(2) = TRIM$(Text3.TEXT)
        DOCPARAM$(3) = TRIM$(Label103.Caption)
        DOCPARAM$(4) = TRIM$(STOMI.Caption)
        DOCPARAM$(5) = TRIM$(Label203.Caption)
        DOCPARAM$(6) = TRIM$(Label303.Caption)
        DOCPARAM$(7) = TRIM$(Label402.Caption)
        DOCPARAM$(8) = TRIM$(Label403.Caption)
        '
        CODTABU1$(1) = "REF-MAT1"
        CODTABU1$(2) = "HORA-EMI"
        CACOLTAB1% = 2
        '
        
        CAITAB1% = 0
        With gridStockProyectado
            For FilaGrid& = 1 To .Rows - 1
               Comprobante$ = AJUSTI$(TRIM$(.TextMatrix(FilaGrid&, 0)), 9)
               REFERENCIA$ = AJUSTI$(TRIM$(.TextMatrix(FilaGrid&, 1)), 29)
               EntradaProyectada$ = AJUSTD$(TRIM$(.TextMatrix(FilaGrid&, 2)), 9)
               SalidaProyectada$ = AJUSTD$(TRIM$(.TextMatrix(FilaGrid&, 3)), 9)
               SaldoProyectado$ = AJUSTD$(TRIM$(.TextMatrix(FilaGrid&, 4)), 12)
               FechaProyectada$ = AJUSTI$(TRIM$(.TextMatrix(FilaGrid&, 5)), 9)
               OBSERVACIONES$ = AJUSTI$(TRIM$(.TextMatrix(FilaGrid&, 6)), 22)
               If ESENVAPLAST% > 0 Then OBSERVACIONES$ = Mid$(OBSERVACIONES$, 6) 'PARA QUITARLE EL 'CTE. '
               CAITAB1% = CAITAB1% + 1
               TETABU1$(1, CAITAB1%) = "." & Comprobante$ & Space(5) & REFERENCIA$ & EntradaProyectada$ & SalidaProyectada$ & Space(3) & FechaProyectada$
               TETABU1$(2, CAITAB1%) = OBSERVACIONES$
            Next FilaGrid&
        End With
        '
        Call PRINTDOC("FORDISTO")
        Screen.MousePointer = 1
        '
      End If
      '
   Command8.Enabled = True
End Sub

Private Sub Command9_Click()
    Command9.Enabled = False
    QUERSTK07.Show 1
    Command9.Enabled = True
End Sub



Private Sub CoPreVe_Click()
    Call Command7_Click
End Sub

Private Sub DiStImp_Click()
    QUERSTK07.Option4 = True
    Call QUERSTK07.Command5_Click
End Sub

Private Sub DiStPan_Click()
    QUERSTK07.Option5 = True
    Call QUERSTK07.Command5_Click
End Sub

Private Sub ExFiImp_Click()
    QUERSTK07.Option4 = True
Call QUERSTK07.Command19_Click
End Sub

Private Sub ExFiPan_Click()
    QUERSTK07.Option5 = True
    Call QUERSTK07.Command19_Click
End Sub

Private Sub ExistyUbi_Click()
    Call QUERSTK07.Command2_Click
End Sub

Private Sub ExNegImpr_Click()
    QUERSTK07.Option4 = True
    Call QUERSTK07.Command30_Click
End Sub

Private Sub ExNegPan_Click()
    QUERSTK07.Option5 = True
    Call QUERSTK07.Command30_Click
End Sub

Private Sub FichaKardex_Click()
    Call QUERSTK07.Command9_Click
End Sub

Private Sub FichKar_Click()
    Call Command3_Click
End Sub

Private Sub Form_Activate()
   Call GRATITFOR(Caption)
   WindowState = 0
   '
   Openforms = DoEvents
   Call FRESHALL
   '
   
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     Openforms = DoEvents
   Wend
End Sub

Private Sub Form_Load()

    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    '
    If Screen.Height > 9000 Then
       Top = (Screen.Height - Height) / 2
       Left = (Screen.Width - Width) / 2
    End If
    
    EPAC$ = TRIM$(EMPREAC$)
    VERANTER$ = "CONSTOCK"
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    Call InstalaFuenteAnchoFijo
    Call ABRESTOCKS
    Call ABREPEDCLI
    Call ABRECOMPRAS
    Call CONAPLI(APLINVO$, TERMINA%) ' Busca la aplicacion a mostrar
    If TERMINA% > 0 Then Call FINAL("")
    '
    ' *** DISEÑO DEL MSFLEXGRID DE STOCK PROYECTADO
    With gridStockProyectado
        .TextMatrix(0, 0) = "Comprobante": .ColWidth(0) = .ColWidth(0) + 100: .ColAlignment(0) = flexAlignLeftBottom
        .TextMatrix(0, 1) = "Referencia": .ColWidth(1) = .ColWidth(1) + 2200: .ColAlignment(1) = flexAlignLeftBottom
        .TextMatrix(0, 2) = "Entra": .ColAlignment(2) = flexAlignRightBottom
        .TextMatrix(0, 3) = "Sale": .ColAlignment(3) = flexAlignRightBottom
        .TextMatrix(0, 4) = "Proyectado": .ColWidth(4) = .ColWidth(4) + 100: .ColAlignment(4) = flexAlignRightBottom
        .TextMatrix(0, 5) = "Fecha": .ColAlignment(5) = flexAlignCenterBottom
        .TextMatrix(0, 6) = "Observaciones": .ColWidth(6) = .ColWidth(6) + 1300: .ColAlignment(6) = flexAlignLeftBottom
        .ColWidth(7) = 0
    End With
    
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
    End If
    
    If InStr(1, EPAC$, "CONTEPLAST") > 0 Then
      If DerechoSupervisor% < 1 Then
        If DERACCE%("SOYVENDE", "Derecho Bloqueado") < 2 Then
          Call FINAL("")
        End If
      End If
    End If
    '
    If InStr(1, EPAC$, "ARCON") > 0 Then
        mnuTrazabilidadXOF.Visible = True
    End If
    '
    Call GRATITFOR(Caption)
    '
    If EXISTE%("PEDIDO02.EXE") > 0 Then
      Command10.Visible = True
    End If
    '
    If EXISTE%("DECAOT09.EXE") > 0 Then
       Command15.Visible = True
    End If
    '
    If InStr(EPAC$, "DOSIVA") > 0 Then
      Command10.ToolTipText = "Anotador de Pedidos - Despachar Materiales Vendidos"
      Command20.Visible = True 'ACCESO DIRECTO A ESTADISTICAS DE CONSUMO
      HOII% = HOY(HOS$)
      If HOII% <= FECHANUM("12/05/03") Then
        Command11.Visible = True   ' recupera packing list
        Command20.Visible = True
      End If
      '\\\OT-08-0643-OD-20/08/08///
      'La linea de abajo se comento segun O/T. 08-0643 para que la informacion _
       aparezcan ordenada con respecto a los encabezados (Queda para ver si _
       queda finallmente así)
      '
      'Label15 = " Comprob.           Referencia                                                Entra          Sale         Fecha      Observ."
      '\\\OT-08-0643-OD-Fin///

    End If
    '
    '\\\OT-08-0990-OD-30/10/08///
    LOMA% = XVALO(Control$("MASTER", "LOMAXCOD"))
    If LOMA% > 15 Then Text1.MaxLength = LOMA%
    '\\\OT-08-0990-OD-FIN///
    '
    LU$ = LUDAT$
    On Error Resume Next
    Show
    On Error GoTo 0
    Openforms = DoEvents
    '
    If EXISTE%(LUDAT$ + "LIPRE001.DAT") <> 1 Then
       Option2.Value = True
    End If
    '
    If EXISTE%("DECAOT09.EXE") > 0 Then
       Command15.Visible = True
    End If
    '
    IDTX$ = IDENTER$
    If Control$(IDTX$, "AVALSTOK") = "MAESTRO" Then
       Option2.Value = True
    End If
    '
    If DERACCE%("VERCOSTO/NOMESS", "Visualizar Costos y Precios") < 1 Then
       Option2.Value = True
    End If
    '
    HOII% = HOY(HO$)
    FEVE% = FECHANUM("04/06/09") 'BOTON CON VENCIMIENTO PARA GUIDI
    If HOII% < FEVE% Then 'VALIDA QUE LA FECHA SEA MENOR AL
                          'VENCIMIENTO PARA QUE EL BOTON SEA VISIBLE
      Command14.Visible = True ' BOTON DE IMPORTACION DE MOVIMIENTOS VIEJOS
    End If
    
    Call ABREORFAB
    Call ABRESERVA
    Call ABRECABAN
    Call ABRECLIEN
    '
    'QTRECU& = CantRegistros&("MOVISTO", "DEPOMOVI = 99")
    'If QTRECU& > 0 Then
    '   Call COMUNI("Movimientos a Recuperar\Sobre Depósito 99.\   \Click sobre 'Aceptar' para Iniciar.")
    '   Call RECUMO99(QTRECU&)
    'End If
    '
    'If InStr(EPAC$, "MEDITEA") > 0 Then
    '   Command16.Visible = True
    'End If
    '
    ' PROCESO PARA GABAL  QUE VERIFICA LOS FALTANTES DE STOCK Y LOS ARTICULOS
    ' QUE ESTEN EN ORDEN DE COMPRAS CON FECHA DE ENTREGA SOLICITADA ANTERIOR A HOY
    ' EN LOS DOS CASOS LOS ITEMS DEBEN FIGURAR EN LISTA DE PRECIOS
    If InStr(EPAC$, "GABAL") > 0 Then
'       If XVALO(Control("CONSTO07", "VERISTCK")) <> HOY(HO$) Then
'         Screen.MousePointer = 11
'         'LLAMO A LA CONSULTA DE STOCK  DEL FORMULARIO DISPSTK07
'         'DONDE GENERA LA GRABACION DEL ARCHIVO STODISP
'         Call DISPSTK07.Command12_Click
'         '
'         ' llamada para cargar articulos faltantes y en O.C. Vencida y envio de mail
'         Call ArtFaltantes
'         '
'         Screen.MousePointer = 1
'         MAILADVTA$ = TRIM$(LCase(Control("DIREMAI", "MAILVNTA")))
'         FE$ = TRIM$(Str$(HOY(HO$)))
'         Call GRACONTROL("CONSTO07", "VERISTCK", FE$)
'         TX$ = "Fin de Proceso de Informe de Faltantes de Stocks y Artículos en O.C.\ con Fecha de Entrega Solicitada Menor a Hoy"
'         TX$ = TX$ + "\ \ Se ha Generado Archivo Art_Faltantes.TXT en la Ruta C:/Flexoft/\ Y Se Ha Enviado Mail Informativo a la\ Direccion: " & MAILADVTA$
'         Call COMUNI(TX$)
'       End If
    End If
    
    'DISPARA GENERACION DE PROCESO A HORARIO
    If InStr(EPAC$, "ALEJPA") > 0 Then
      MnuComparativoVR.Visible = True
      mnuConfig_RtaMDB.Visible = True
      Timer3.Enabled = True
    End If
    
    '
    If InStr(1, EMPREAC$, "GABAL") > 0 Then
        '
        Call CREACAMPO("", "UBICAVOL", "SECTOR", 10, 2)
        Call CREACAMPO("", "UBICAVOL", "ESTANTERIA", 10, 2)
        Call CREACAMPO("", "UBICAVOL", "FAMILIA", 10, 2)
        Call CREACAMPO("", "UBICAVOL", "FILA", 10, 3)
        Call CREACAMPO("", "UBICAVOL", "COLUMNA", 10, 3)
        Call CREACAMPO("", "UBICAVOL", "VOLUMEN", 6, 0)
        Call CREACAMPO("", "UBICAVOL", "MARBORRA", 3, 0)
        '
        Call CREACAMPO("", "UBICACOD", "SECTOR", 10, 2)
        Call CREACAMPO("", "UBICACOD", "ESTANTERIA", 10, 2)
        Call CREACAMPO("", "UBICACOD", "FAMILIA", 10, 2)
        Call CREACAMPO("", "UBICACOD", "FILA", 10, 3)
        Call CREACAMPO("", "UBICACOD", "COLUMNA", 10, 3)
        Call CREACAMPO("", "UBICACOD", "CODIGO", 3, 0)
        Call CREACAMPO("", "UBICACOD", "MARBORRA", 3, 0)
        '
        Call CREACAMPO("", "UBICALOT", "SECTOR", 10, 2)
        Call CREACAMPO("", "UBICALOT", "ESTANTERIA", 10, 2)
        Call CREACAMPO("", "UBICALOT", "FAMILIA", 10, 2)
        Call CREACAMPO("", "UBICALOT", "FILA", 10, 3)
        Call CREACAMPO("", "UBICALOT", "COLUMNA", 10, 3)
        Call CREACAMPO("", "UBICALOT", "CODIGO", 3, 0)
        Call CREACAMPO("", "UBICALOT", "CANTIDAD", 6, 0)
        Call CREACAMPO("", "UBICALOT", "LOTE", 10, 12)
        '
    End If

    
    Call CREAR_STORED_PROCEDURE

   If InStr(1, EMPREAC$, "CONTEPLAST") < 1 Then
        mnuStockRangoAgrupado.Visible = False
        mnuStockProveedorRango.Visible = False
        mnuStockRangoTodos.Visible = False
        mnuInfConteplast.Visible = False
        mnuStockConMetrosyKilosEnvaplast.Visible = False
    End If
    
    If InStr(1, EMPREAC$, "CAMBACERES") < 1 Then
       mnuStFeMatFam.Visible = False
       mnuReportePlantaLaminado.Visible = False
       mnuStTotLuzCamb.Visible = False
       mnuMovStLuzuriaga.Visible = False
       mnuStockMetrosvsFtecnica.Visible = False
    End If
    
    If InStr(1, EMPREAC$, "QUILMES") < 1 Then
       mnuStockConMetrosyKilos.Visible = False
    End If


    
End Sub
'
Sub RECUMO99(QTRECU&)
    '
    Screen.MousePointer = 11
    JJ& = 0: KK& = 0
    SQLX$ = "SELECT * FROM MOVISTO"
    SQLX$ = SQLX$ + " WHERE DepoMovi = 99"
    '
    Dim MOVITMP As ADODB.Recordset
    Set MOVITMP = New ADODB.Recordset
    MOVITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With MOVITMP
      Do While Not .EOF
         JJ& = JJ& + 1
         Call TRACE(20, JJ&, QTRECU&)
         CIXI% = !cod_inte
         DPX% = DEPOPRE%(CIXI%)
         If DPX% > 0 Then
              !depomovi = DPX%
              .Update
              KK& = KK& + 1
            Else
              OPP% = COMALTER%("Imposible Determinar Depósito Predeterminado\para el Código '" + CODEXT$(CIXI%) + "'.\\Continuar\Cancelar Proceso")
              If OPP% = 2 Then GoTo 99
         End If
      .MoveNext
      Loop
    End With
    '
    Screen.MousePointer = 1
    Call COMUNI("Proceso Completo - " & KK& & " Registros Recuperados.")
    '
99  Screen.MousePointer = 1
    '
End Sub

'
Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub GesRese_Click()
    Call QUERSTK07.Command14_Click
End Sub

Private Sub GesReser_Click()
    Call QUERSTK07.Command7_Click
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   CIII% = CODINT%(Text1)
   Picture2.Cls
   If CIII% > 0 Then
     Screen.MousePointer = 11
     Fact% = HOY(HOS$)
     FINI% = Fact% - 32 * 13 + 1
     RFF$ = RECFILT$("MOVISTO", 2, MKI$(CIII%))
     STK = 0: STMAX = 0
     For KKL& = 1 To Len(RFF$) Step 4
       REMO& = CVS(Mid$(RFF$, KKL&, 4))
       XX$ = REGLEIDO$("MOVISTO", REMO&)
       FEMO% = CVI(Left$(XX$, 2))
       STK = STK + CVD(Mid$(XX$, 88, 8)) - CVD(Mid$(XX$, 96, 8))
       If FEMO% >= FINI% Then
         If STK > STMAX Then STMAX = STK
         If STMAXAN > STMAX Then STMAX = STMAXAN
       End If
       If STMAX < 0.1 Then STMAXAN = STK
     Next KKL&
     If STMAX < 0.1 Then STMAX = STMAXAN
     Picture2.ForeColor = RGB(0, 0, 0)
     Picture2.Line (0.05 * Picture2.Width, 0.94 * Picture2.Height)-(0.97 * Picture2.Width, 0.94 * Picture2.Height)
     Picture2.Line (0.05 * Picture2.Width, 0.94 * Picture2.Height)-(0.05 * Picture2.Width, 300)
     If STMAX > 0.1 Then
       COEFV = 0.8 * (Picture2.Height - 300) / STMAX
       COEFH = 0.88 * Picture2.Width / (Fact% - FINI%)
       XINIA = 0.05 * Picture1.Width
       YINIA = 0.94 * Picture1.Height - COEFV * STMAXAN
       Picture2.ForeColor = RGB(255, 0, 0)
       STK = 0
       For KKL& = 1 To Len(RFF$) Step 4
         REMO& = CVS(Mid$(RFF$, KKL&, 4))
         XX$ = REGLEIDO$("MOVISTO", REMO&)
         FEMO% = CVI(Left$(XX$, 2))
         STK = STK + CVD(Mid$(XX$, 88, 8)) - CVD(Mid$(XX$, 96, 8))
         STX = STK: If STX < 0 Then STX = 0
         YINI = 0.94 * Picture2.Height - COEFV * STX
         XINI = 0.05 * Picture2.Width + COEFH * (FEMO% - FINI%)
         If XINI >= 0.05 * Picture1.Width Then
           Picture2.Line (XINIA, 0.94 * Picture2.Height)-(XINI, YINIA), RGB(255, 80, 0), BF
         End If
         If XINI >= XINIA Then
           XINIA = XINI
           YINIA = YINI
         End If
       Next KKL&
       XINI = 0.05 * Picture2.Width + COEFH * (Fact% - FINI%)
       If XINI >= 0.05 * Picture1.Width Then
          Picture2.Line (XINIA, 0.94 * Picture2.Height)-(XINI, YINI), RGB(255, 80, 0), BF
       End If
       Picture2.ForeColor = RGB(0, 150, 160)
       Picture2.Line (0.05 * Picture2.Width, 0.94 * Picture2.Height - COEFV * XVALO(Label103))-(0.95 * Picture2.Width, 0.94 * Picture2.Height - COEFV * XVALO(Label103))
       Picture2.ForeColor = RGB(200, 150, 0)
       Picture2.Line (0.05 * Picture2.Width, 0.94 * Picture2.Height - COEFV * XVALO(STOMI))-(0.95 * Picture2.Width, 0.94 * Picture2.Height - COEFV * XVALO(STOMI))
       Picture2.Visible = True
     End If
     Screen.MousePointer = 1
   End If
       
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Picture2.Visible = False
   Screen.MousePointer = 1
End Sub

Private Sub Image2_DblClick()
    Call ACCDIR("CONSTOCK")
End Sub

Private Sub ItStImp_Click()
    QUERSTK07.Option4 = True
    Call QUERSTK07.Command20_Click
End Sub

Private Sub ItStPant_Click()
    QUERSTK07.Option5 = True
    Call QUERSTK07.Command20_Click
End Sub

Private Sub Label13_DblClick()
    Call LIUBIC_DblClick
    Exit Sub
    
    If DERACCE%("MODIUBIC", "ALTA O MODIFICACION DE UBICACIONES") < 2 Then
       Exit Sub
    End If
   
    CIII% = CODINT%(Text1.TEXT)
    Call CARUBIC(CIII%)
    Screen.MousePointer = 1
    Text1.TEXT = ""
    Text1.TEXT = CODEXT$(CIII%)
    '\\\OT-08-1021-OD-07/11/08///
20  LIUBIC.Clear
    RGXX$ = RECFILT$("UBISTOCK", 1, MKI$(CIII%)) ' se descomento por que si no no refrescaba
    For U& = 1 To Len(RGXX$) Step 4
        REG& = CVS(Mid$(RGXX$, U&, 4))
        X$ = REGLEIDO$("UBISTOCK", REG&)
        If CVI(Left$(X$, 2)) <> CIII% Then
           X$ = REGLEIDO$("UBISTOCK", 1)
           Call GRAREG("UBISTOCK", X$, 1)
           Call CIERRARCH("UBISTOCK")
           GoTo 20
        End If
        '
        If TRIM$(Mid$(X$, 3, 24)) <> "" Then
           Y$ = Mid$(X$, 3, 24)
           LIUBIC.AddItem Y$
        End If
    Next U&
'    '\\\OT-08-1021-OD-FIN///
'    Screen.MousePointer = 1
End Sub


Private Sub Label202_DblClick()
'   CONTROCLIKK% = 1
'   TOPIN = List2.TopIndex
'   For U& = 1 To List2.ListCount
'     List2.Selected(U& - 1) = False
'     If Left$(List2.LIST(U& - 1), 2) = "OD" Then
'       List2.Selected(U& - 1) = True
'     End If
'   Next U&
'   On Error Resume Next
'   List2.TopIndex = TOPIN
'   On Error GoTo 0
'   CONTROCLIKK% = 0
End Sub

Private Sub List1_DblClick()
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   Static MODOTRAZA$
   If MODOTRAZA$ = "" Then
       MODOTRAZA$ = Control$("", "RETRAZA")
       If MODOTRAZA$ = "" Then MODOTRAZA$ = "FIFO"
   End If
   '
   'PRESENTA CARGALISEL CON LA CANTIDAD DE MATERIAL EN TERCEROS QUE TIENE ESTE ARTICULO
   If InStr(List1.TEXT, "TERCEROS") > 0 Then
     CONDI$ = "COD_INTE=" & CODINT(Text1) & " AND (CANT_ENTREG-CANT_RENDID > 0 OR CANT_DEVU <> 0) AND NPRO_OCOM > 0"
     Call CARGALISEL("TERCEROS", "MACONSIG", "NPRO_OCOM/I5;RASO_PROV/A24;NREMI_CONSIG/A8;CANT_ENTREG-CANT_RENDID-CANT_DEVU/F12.2", CONDI$, "GROUP BY NPRO_OCOM, CAMPO2")
     Call SELECHO("TERCEROS")
     Exit Sub
   End If
   
   CIXI% = CODINT%(Text1)
   NDEP% = XVALO(Mid$(List1.TEXT, 61, 4))
   If ESTRAZABLE(CIXI%) = 1 Then
        If MODOTRAZA = "FIFO" Then
            If InStr(EMPREAC$, "MEDITEA") = 0 And InStr(EMPREAC$, "CAMBACERES") = 0 Then
                Call RECUSALOT(CIXI%, NDEP%)
            End If
        End If
           
        If InStr(EPAC$, "CAMBACERES") > 0 Or InStr(EPAC$, "QUILMES") > 0 Then
          SALOTES07.Label9 = TRIM$(Str$(NDEP%))
          SALOTES07.Label4 = Text1
          Call MuestraLotesExpress(CODINT%(Text1), NDEP%)
          SALOTES07.Show 1
          Exit Sub
        End If
        Call ARMALILOT(CIXI%, NDEP%, 0, 1, 1)
        If EXISTE%("DECAOT09.EXE") > 0 Then
          SALOTES09.Label9 = TRIM$(Str$(NDEP%))
          SALOTES09.Label4 = Text1
          Call SALOTES09.MUESTRASALOT
          SALOTES09.Show 1
          Exit Sub
        Else
          SALOTES07.Label9 = TRIM$(Str$(NDEP%))
          SALOTES07.Label4 = Text1
          Call SALOTES07.MUESTRASALOT
          SALOTES07.Show 1
          Exit Sub
        End If
      Else
        Call COMUNI("Imposible Consulta por Lotes en Material no Trazable.")
   End If
   '
99 Screen.MousePointer = 1
   '
End Sub

Sub RECUSALOT(CIXI%, NDEP%)
   '
   Static MODOTRAZA$
   If MODOTRAZA$ = "" Then
      MODOTRAZA$ = Control$("", "RETRAZA")
      If MODOTRAZA$ = "" Then MODOTRAZA$ = "FIFO"
   End If
   '
   If MODOTRAZA = "FIFO" Then
     If ESTRAZABLE(CIXI%) = 1 Then
       Screen.MousePointer = 11
       Call APERBASDAT("STOCKS")
       Call ARMALILOT(CIXI%, NDEP%, 0, 1, 1)
       For U& = 1 To ULTREG&("!LILOTSE")
         TTYY$ = REGLEIDO$("!LILOTSE", U&)
         IDLO$ = TRIM$(Mid$(TTYY$, 1, 16))
         SALOT = SaldoLote(CIXI%, IDLO$)
         SQLZ$ = " SELECT CantSalid,CANTINGRE,IdenLote FROM MOVISTO "
         SQLZ$ = SQLZ$ + "WHERE Cod_Inte = " & CIXI% & " "
         SQLZ$ = SQLZ$ + "AND (IdenLote = '' OR IdenLote IS NULL) "
         '
         Dim MOVCONTEMP As ADODB.Recordset
         Set MOVCONTEMP = New ADODB.Recordset
25       On Error Resume Next
         MOVCONTEMP.Open FILTSQL$(SQLZ$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
         If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
            On Error GoTo 0
            Call CapturaErrorOpen
            GoTo 25
         End If
         On Error GoTo 0
         '
         With MOVCONTEMP
            Do While Not .EOF
               CANCON = XVALO(!CANTSALID) - XVALO(CANTINGRE)
               If CANCON <= SALOT + FRACMIN(CIXI%) Then
                  !idenlote = IDLO$
                  .Update
                  SALOT = SALOT - CANCON
               End If
            .MoveNext
            Loop
         End With
         MOVCONTEMP.Close
         Set MOVCONTEMP = Nothing
         Call ACSALOTE(CIXI%, IDLO$, "NOMESS")
       Next U&
     End If
   End If
   Screen.MousePointer = 1
   '
End Sub
'


Private Sub LIUBIC_DblClick()
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   

    If DERACCE%("MODIUBIC", "ALTA O MODIFICACION DE UBICACIONES") < 2 Then
       Exit Sub
    End If
    CIII% = CODINT%(Text1.TEXT)
    CNTREG& = 0
    
    If InStr(EPAC$, "AHP") < 1 Then
      CNTREG& = CantRegistros&("UBICASQL", "", "NOMESS")
    End If
    
    If CNTREG& > 0 Then  'si ya migro las ubicaciones a la tabla sql las carga de aqui
            Call CARUBICSQL(CIII%, 0)
             LIUBIC.Clear
             SQLX$ = "SELECT CODIUBICA FROM UBICASQL  WITH(NOLOCK) "
             SQLX$ = SQLX$ + " WHERE COD_INTE = " & CI%
             SQLX$ = SQLX$ + " ORDER BY CODIUBICA"
             Set rs = READSET(SQLX$)
             With rs
               Do While Not .EOF
                  CODUB$ = TRIM$(SAFETEXT$(!CODIUBICA))
                  If CODUB$ <> "" Then
                     LIUBIC.AddItem CODUB$
                  End If
                  .MoveNext
               Loop
             End With
             rs.Close
             Exit Sub
    Else
        Call CARUBIC(CIII%)
    End If
    
    Screen.MousePointer = 1
    Text1.TEXT = ""
    Text1.TEXT = CODEXT$(CIII%)
    '\\\OT-08-1021-OD-07/11/08///
20  LIUBIC.Clear ' se descomento por que si no no refrescaba
    RGXX$ = RECFILT$("UBISTOCK", 1, MKI$(CIII%))
    For U& = 1 To Len(RGXX$) Step 4
        REG& = CVS(Mid$(RGXX$, U&, 4))
        X$ = REGLEIDO$("UBISTOCK", REG&)
        If CVI(Left$(X$, 2)) <> CIII% Then
           X$ = REGLEIDO$("UBISTOCK", 1)
           Call GRAREG("UBISTOCK", X$, 1)
           Call CIERRARCH("UBISTOCK")
           GoTo 20
        End If
        '
        If TRIM$(Mid$(X$, 3, 24)) <> "" Then
           Y$ = Mid$(X$, 3, 24)
           LIUBIC.AddItem Y$
        End If
     Next U&
'    '\\\OT-08-1021-OD-FIN///
    Screen.MousePointer = 1
End Sub

Private Sub MaDep_Click()
    Call QUERSTK07.Command6_Click
End Sub


Private Sub gridStockProyectado_DblClick()

   'MUESTRA POR  PANTALLA LA IMAGEN DEL DOCUMENTO ELEGIDO
   '****************************************************
   With gridStockProyectado
        FilaGrid& = .RowSel
        
        TIP$ = UCase$(Mid$(.TextMatrix(FilaGrid&, 0), 1, 2))
        If TRIM$(TIP$) = "" Then Exit Sub
        
        NUDOCU$ = TRIM$(Mid$(.TextMatrix(FilaGrid&, 0), 4, 6))
   End With
   
   Select Case TIP$
     Case "PC"
       'PEDIDO ' TENER EN CUENTA DERECHOS PARA VISUALIZAR PRECIOS SI NO PRESENTAR ORDEN DE DESPACHO
10      OPX% = ALTERNA%("Visualizar Pedidos - Orden de Venta\Visualizar Orden de Despacho")
      If OPX% < 1 Or OPX% > 2 Then Exit Sub
      '
      If OPX% <> 2 Then
        If DERACCE%("VERCOSTO/NOMESS", "Visualizar Costos y Precios") < 1 Then
           Call MENSERR(24, "Usuario No Autorizado para Visualizar Pedidos.\Seleccione Orden de Despacho")
           GoTo 10
        End If
      End If
      '
      If OPX% = 1 Then
         TIDOCUM$ = "Pedido de Cliente"
         DOCUMENTO_ELEGIDO$ = "Orden de Venta"
      ElseIf OPX% = 2 Then
         TIDOCUM$ = "Orden de Despacho"
         DOCUMENTO_ELEGIDO$ = TIDOCUM$
      End If
      While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
      NUDOCU$ = NUDOCU$ + " "
      'FIN PEDIDO / ORDEN DE DESPACHO
      '
      Case "OF"
        'ORDEN DE FABRICACION
        While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
        NUDOCU$ = "OF-" + NUDOCU$ + "-" + "1"
        TIDOCUM$ = "Orden de Fabricación " + NUDOCU$
     Case "OC"
        'ORDEN DE COMPRA
        TIDOCUM$ = "Nota de Pedido"
        While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
        NUDOCU1$ = " " + NUDOCU$ + " "
        NUDOCU2$ = " " + NUDOCU$ + "-"
     Case "BR"
         'BOLETA DE RECEPCION
         TIDOCUM$ = "Boleta de Recepción"
         NUDOCU$ = " " + NUDOCU$ + " -"
     Case Else
         Exit Sub
   End Select
   '
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
        XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
        If InStr(XX$, TIDOCUM$) > 4 Then
          PPXX% = InStr(XX$, "Nro.")
          '
          If PPXX% > 0 Then
            XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
          End If
          '
          If TIP$ = "OC" Then
            'CASO DE ORDEN DE COMPRA BUSCA POR UN OR CON DOS VARIABLES
            If InStr(XX$, NUDOCU1$) > 4 Or InStr(XX$, NUDOCU2$) > 4 Then
              DOCUNU& = CVS(Left$(XX$, 4))
              GoTo 5
            End If
          ElseIf InStr(XX$, NUDOCU$) > 4 Then
             'CASOS RESTANTES
             DOCUNU& = CVS(Left$(XX$, 4))
             GoTo 5
          End If
        End If
    Next UI&
    Screen.MousePointer = 1
    Call COMUNI("Documento no Encontrado")
    Exit Sub
    '
    'MUESTRA DOCUMENTO
    Screen.MousePointer = 1
5   Call MUESTRADOC(DOCUNU&)


End Sub




Private Sub menuMigrarUbicacionesASql_Click()
    '
    If CLACONTRA% = 0 Then GoTo 99
    Screen.MousePointer = 11
    CNTREG& = CantRegistros&("UBICASQL", "", "NOMESS")
    If CNTREG& > 0 Then
       Call COMUNI("La Tabla UbicaSQl, ya posee información")
       Exit Sub
    End If
    
    '
    FIN& = NUMAS%
    For CI1% = 1 To FIN&
        RGXX$ = RECFILT$("UBISTOCK", 1, MKI$(CI1%))
        For U& = 1 To Len(RGXX$) Step 4
           REG& = CVS(Mid$(RGXX$, U&, 4))
           X$ = REGLEIDO$("UBISTOCK", REG&)
'           If CVI(Left$(X$, 2)) <> CI1% Then
'              X$ = REGLEIDO$("UBISTOCK", 1)
'              Call GRAREG("UBISTOCK", X$, 1)
'              Call CIERRARCH("UBISTOCK")
'              GoTo 20
'           End If
           '
           If TRIM$(Mid$(X$, 3, 24)) <> "" Then
              CODIUBICA$ = TRIM$(Mid$(X$, 3, 24))
              OBSERVACION$ = TRIM$(Mid$(X$, 31, 30))
              Call INSERTUBICASQL(CI1%, CODIUBICA$, OBSERVACION$, 0)
           End If
        Next U&
   Next CI1%


    Call COMUNI("Proceso Terminado")
    Call CIERRARCH("UBISTOCK")
99  Screen.MousePointer = 1
    '
End Sub


Private Sub MnuComparativoVR_Click()

10 Call DESHASFECHA(Des%, Has%, PERIO$)
   If PERIO$ = "" Then
    Exit Sub
   End If
   Screen.MousePointer = 11
   Desde = CDbl(CVDAT(Des%))
   Hasta = CDbl(CVDAT(Has%))

   CANDIA% = DIENTRE%(Desde, Hasta)
   If CANDIA% < 0 Then GoTo 10
  
   DES1 = FETEXCOR1$(Des%)
   HAS1 = FETEXCOR1$(Has%)
   '
   Campos$ = "Fecha Fact/D12;Código/A20;Descripción/A32;Documento/A24;Cant.Fact/F10.1;FechaNum/D0;CI/F0"
   Call CARGA_ENCABEZADO(CompRecFa.GRID, Campos$, CompRecFa, 1)
   Call CARGA_ENCABEZADO(CompRecFa.MSF_2, Campos$, CompRecFa)
   '
   Campos$ = "Código/A20;Descripción/A32;Stock Prev./F0;Cant.Fact/F10.1;Cant.Recibida/F10.1;Diferencia/F10.1;Sock Fin/F0"
   Call CARGA_ENCABEZADO(CompRecFa.MSFDET, Campos$, CompRecFa)
   Call CARGA_ENCABEZADO(CompRecFa.MSF3, Campos$, CompRecFa)
   '
   Call CargarComboSegunGrid(CompRecFa.MSF3, CompRecFa.CMBFILTRO1, 1, Campos$)

   ARRCH$ = Control("CONFIRUT", "RUTAXMDB")
   If EXISTE%(ARRCH$) < 1 Then
    Call COMUNI("Atención No se Encontró La Ruta de Base Datos de WALDBOTT ")
    Exit Sub
   End If
    
   Set BaDaMoviExt = OpenDatabase(ARRCH$)
   
   SQLX$ = " SELECT * FROM ITITEMS "
   SQLX$ = SQLX$ + " WHERE (letra = 'A'"
   SQLX$ = SQLX$ + " OR letra = 'B'"
   SQLX$ = SQLX$ + " OR letra = 'E')"
   SQLX$ = SQLX$ + " AND Cdbl(FECEMISION) >= " & Desde & " "
   SQLX$ = SQLX$ + " AND Cdbl(FECEMISION)  <= " & Hasta & " "
   SQLX$ = SQLX$ + " ORDER BY FECEMISION,LETRA"
   '
   Dim MaterialNoCodif$()
   ReDim Preserve MaterialNoCodif$(2)
   J& = 0: n& = 0
   Set rs = BaDaMoviExt.OpenRecordset(SQLX$, dbOpenSnapshot)
   With rs
     Do While Not .EOF
            J& = J& + 1
            CI% = CODINT%(SAFETEXT$(!ARTICULO))
            If CI% < 1 Then
              MATENOCODIF% = 1
              n& = n& + 1
              If UBound(MaterialNoCodif$) < n& Then ReDim Preserve MaterialNoCodif$(n&)
              MaterialNoCodif$(n&) = SAFETEXT$(!ARTICULO)
            End If
             
            
            CompRecFa.GRID.Rows = J& + 1
            CompRecFa.GRID.TextMatrix(J&, 0) = J&
            CompRecFa.GRID.TextMatrix(J&, 1) = FECHATEX$(CVINT((!FECEMISION)))
            CompRecFa.GRID.TextMatrix(J&, 2) = SAFETEXT$(!ARTICULO)
            CompRecFa.GRID.TextMatrix(J&, 3) = DESCRIT0$(CODINT%(SAFETEXT$(!ARTICULO)))
            CompRecFa.GRID.TextMatrix(J&, 4) = SAFETEXT$(!NROCOMP)
            TiMovi = !es
            If TiMovi = "E" Then
               SIG% = -1
            Else
               SIG% = 1
            End If
            CompRecFa.GRID.TextMatrix(J&, 5) = FORMATNUM$(XVALO(!UNIDADES) * SIG%, "F12.1")
            CompRecFa.GRID.TextMatrix(J&, 6) = CVINT(!FECEMISION)
            CompRecFa.GRID.TextMatrix(J&, 7) = CODINT%(!ARTICULO)
            .MoveNext
     Loop
   End With
   '
   Screen.MousePointer = 1

   If J& < 1 Then
      Call COMUNI("Sin Información Segun Sistema Externo - Para el Rango Seleccionado")
      Screen.MousePointer = 1
      Exit Sub
   End If
   '
   'ACA ENTRA SI HAY CODIGOS DE WALDBOT QUE NO EXISTEN EN FLEXOFT
   If MATENOCODIF% = 1 Then
50      OPX% = COMALTER%("Se Han Detectado Códigos No Existentes en Base de Datos Flexoft \\Cancelar\Ver Códigos\Continuar")
        If OPX% < 2 Then Exit Sub
        If OPX% = 2 Then
           J& = 0
           Campos$ = "Código/A24;Descripción/A48"
           Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
           Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
           For J& = 1 To UBound(MaterialNoCodif$)
             FRMZELECHO.MSF2.Rows = J& + 1
             FRMZELECHO.MSF2.TextMatrix(J&, 1) = MaterialNoCodif$(J&)
             CONDI$ = "NUMERO = '" & MaterialNoCodif$(J&) & "'"
             SQLX$ = "select descrip from articulo  where " & CONDI$
             Set rs = BaDaMoviExt.OpenRecordset(SQLX$, dbOpenSnapshot)
             FRMZELECHO.MSF2.TextMatrix(J&, 2) = rs!Descrip
             
             If J& = UBound(MaterialNoCodif$) Then Exit For
           Next J&
           Screen.MousePointer = 1
           rs.Close
           Set rs = Nothing
           '
           'ELIMINA ITEM REPETIDO
           FRMZELECHO.MSF2.COL = 1
           FRMZELECHO.MSF2.Sort = 1
           i& = 0
           'elimino los repetidos
           While i& < FRMZELECHO.MSF2.Rows - 1
              i& = i& + 1
              CODIX$ = FRMZELECHO.MSF2.TextMatrix(i&, 1)
'              Debug.Print codix$
              If CODIX$ = CODIX_ANT$ Then
                 Call cmdEliminaItem(FRMZELECHO.MSF2, i&)
                 i& = i& - 1
              End If
              CODIX_ANT$ = CODIX$
           Wend
           '
           FRMZELECHO.Width = 9500
55         FRMZELECHO.Show 1
           GoTo 50
         End If
   End If
   '
   CompRecFa.LDEs = SAFETEXT$(Des%)
   CompRecFa.LHAS = SAFETEXT$(Has%)
   Call CargarDatosSysflom(Des%, Has%)
   Screen.MousePointer = 1
   CompRecFa.Show 1
End Sub
Function UltimoRegistro(COLUMNA) As Double
   Dim valor As Variant
   Dim CantFilasVacias&
   CantFilasVacias& = 0
   UltFila& = 0
   For Fila& = 2 To Hoja.Rows.Count
      If CantFilasVacias >= 100 Then
         GoTo 10
      End If
         valor = TRIM$(Hoja.Cells(Fila&, COLUMNA))
         If valor = "" Then 'Si no hay valor
            If UltFila& = 0 Then 'Si no hay registro de ultima fila se registra, si no se mantiene la almacenada
               UltFila& = Fila& - 1
            End If
            CantFilasVacias& = CantFilasVacias& + 1
           Else
            CantFilasVacias& = 0
            UltFila& = 0
          End If
   Next Fila&
10 UltimoRegistro = UltFila&


End Function

Private Sub mnuConfig_RtaMDB_Click()
   SelecRuta.Show 1
End Sub

Private Sub mnuConfiguracionGeneral_Click()
    Call Command27_Click
End Sub

Private Sub mnuConjuntosAnalizarFabricacion_Click()

    Call TablasAuxiliares
    
    IDLIST& = ObtenerIdTablaAuxiliar("ANALISTO")
    Call ABM_TABLA(IDLIST&, , , "cast(campo1 as int)")
    
End Sub

Private Sub mnuCtrlResvsTransf_Click()
    Erase RESP_ZELE_VECT
    Campos$ = "Cód.Rango/A10;" '1
    Campos$ = Campos$ + "Descripción/A64;"      '2
    
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = ""
    Dim obj As New RECORXET
    CAMPOSTABLA$ = "C.CODRANGO ,C.Descripcion "
    TABLA$ = "CODIGORANGO C  "
    CONDI$ = "C.CODIGORANGO_ID > 0 group by C.CODRANGO,C.Descripcion Order by C.CODRANGO"
    '
    Set RS1 = obj.RS_GENERAL(TABLA$, CAMPOSTABLA$, CONDI$)
    Call Carga_MSF_Recordset(RS1, Campos$, FRMZELECHO.MSF2, 1)
    On Error Resume Next
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    On Error GoTo 0
    FRMZELECHO.Caption = "Seleccione Còdigo Rango"
    FRMZELECHO.Width = 13000
50  FRMZELECHO.Show 1
'__________________________________________________________________________

    CODIRANGO$ = TRIM(RESP_ZELE_VECT(1))
    If CODIRANGO$ = "" Then Exit Sub
    
    
   
    Dim CTP As New Conteplast
    Campos$ = "O.Fabric./A16;" '1
    Campos$ = Campos$ + "Còdigo/A16;"      '2
    Campos$ = Campos$ + "CantiReser/G12.0;"      '3
    Campos$ = Campos$ + "CantiTransf./G12.0;"      '4
    Campos$ = Campos$ + "Diferencia./G12.0;"      '4
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = ""

   J& = 0
   CONDI$ = "CODRANGO = '" & CODIRANGO$ & "'"
   CONDIOR$ = CTP.CONSULTA_OR("CODIGORANGO", "CODMATPRIMA", CONDI$, "COD_EXTE")
   CONDINORFA$ = ""
   TotCantReservada_Transf# = 0
   'PRIMERO CONSULTO LAS RESERVAS CORRESPONDIENTES AL CODIGO RANGO(EN LAS RESERVAS FIGURAN CON EL CODIGO RANGO)
   'INNER JOIN CON ORDEFABR CON STATUS < 3
   SQLX$ = " select r.Cod_Exte,r.idsubor, sum(r.CantRes) as Canti "
   SQLX$ = SQLX$ + " from reserva r inner join ordefabr o WITH(NOLOCK)"
   SQLX$ = SQLX$ + " on o.idsubor= r.IdSubOr"
   SQLX$ = SQLX$ + " where r.COD_EXTE = '" & CODIRANGO$ & "'"
   SQLX$ = SQLX$ + " and  o.StatuOrf < 3"
   SQLX$ = SQLX$ + " group by r.Cod_Exte,r.idsubor"
'Call mostrarDatosx("1 " & SQLX$)
   Set rs = READSET(SQLX$)
   With rs
     Do While Not .EOF
        NORFA$ = SAFETEXT$(!idsubor)
        Codigo$ = SAFETEXT$(!Cod_Exte)
        'CantiTransferida_X_OF_Y_CodigosAsociadosalRango# devuelve la canti transferida para el param NORFA$ Y LOS CODIGOS REALES DEL RANGO
        
        CANTITRANSFERIDA# = CTP.CantiTransferida_X_OF_Y_CodigosAsociadosalRango#(NORFA$, CONDIOR$)
        parcialCantReservada# = XVALO(!CANTI) 'CANTIDAD RESERVADA DE LA OF DE LOS CODIGOS RANGO DEL RESERVA
        '
        'ParcialResultado# = parcialCantReservada# - CANTITRANSFERIDA#
'        If ParcialResultado# <= 0 Then ParcialResultado# = 0 'SI SE TRANSFIRIO MAS DE LO RESERVADO ASIGNA 0
        TotCantReservada_Transf# = TotCantReservada_Transf# + ParcialResultado#
             J& = J& + 1
             FRMZELECHO.MSF2.Rows = J& + 1
             FRMZELECHO.MSF2.TextMatrix(J&, 1) = NORFA$
             FRMZELECHO.MSF2.TextMatrix(J&, 2) = Codigo$
             FRMZELECHO.MSF2.TextMatrix(J&, 3) = FORMATNUM$(parcialCantReservada#, "G10.0")
             FRMZELECHO.MSF2.TextMatrix(J&, 4) = FORMATNUM$(CANTITRANSFERIDA#, "G10.0")
             Diferencia# = FORMATNUM$(parcialCantReservada# - CANTITRANSFERIDA#, "G10.0")
             FRMZELECHO.MSF2.TextMatrix(J&, 5) = Diferencia#
            .MoveNext
        Loop
    End With
    rs.Close
    Set rs = Nothing
        
    FRMZELECHO.Caption = "Seleccione Orden de Fabricaciòn"
    FRMZELECHO.Width = 13000
90  FRMZELECHO.Show 1
'___________________________________________________________________________________________

    NORFASEL$ = TRIM(RESP_ZELE_VECT(1))
    If NORFASEL$ = "" Then Exit Sub
    
    Campos$ = "Còdigo/A16;"       '1
    Campos$ = Campos$ + "CantiTransf./G12.0;"      '3

    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = ""
    J& = 0
    SQLX$ = " SELECT  sum(CANTSALID-CANTINGRE) AS CANTI,COD_EXTE  FROM MOVISTO WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE CODIMOVI = 'TR'   and depomovi = 1"
    If CONDIOR$ <> "" Then SQLX$ = SQLX$ + " AND " & CONDIOR$
    SQLX$ = SQLX$ + " AND  DOSECLAR = '" & NORFASEL$ & "'"
    SQLX$ = SQLX$ + "  group by cod_exte"
    Set rs = READSET(SQLX$)
    With rs
     Do While Not .EOF
        Codigo$ = SAFETEXT$(!Cod_Exte)
        
        parcialCantReservada# = XVALO(!CANTI) 'CANTIDAD RESERVADA DE LA OF DE LOS CODIGOS RANGO DEL RESERVA
        CANTITRANSFERIDA# = XVALO(!CANTI)
        '
        'ParcialResultado# = parcialCantReservada# - CANTITRANSFERIDA#
'        If ParcialResultado# <= 0 Then ParcialResultado# = 0 'SI SE TRANSFIRIO MAS DE LO RESERVADO ASIGNA 0
        TotCantReservada_Transf# = TotCantReservada_Transf# + ParcialResultado#
             J& = J& + 1
             FRMZELECHO.MSF2.Rows = J& + 1
             FRMZELECHO.MSF2.TextMatrix(J&, 1) = Codigo$
             FRMZELECHO.MSF2.TextMatrix(J&, 2) = FORMATNUM$(CANTITRANSFERIDA#, "G10.0")
            .MoveNext
        Loop
    End With
    rs.Close
    Set rs = Nothing
    
    FRMZELECHO.Caption = "Informe por Codigo de Orden de Fabricaciòn: " & NORFASEL$
    FRMZELECHO.Width = 13000
    FRMZELECHO.Show 1

End Sub


Private Sub mnuFabricacionSegunStockRecepcion_Click()
    
    ' *** OBTENGO TOTAL COMPONENTES
    Call SELECHO("TADEPOSI")
    DEPOSITO% = XVALO(VALACT1$("TADEPOSI"))
    If DEPOSITO% = 0 Then GoTo 99
    
    IncluyeCompras% = COMALTER%("¿Incluye Compras Pendientes?\\No, Excluir\Si, Incluir")
    IncluyeCompras% = IncluyeCompras%
    
    ConsultaSql$ = "select max(cod_inte) maxitem from" _
    & " (select mo.cod_inte, sum(mo.cantingre) ingreso, sum(mo.cantsalid) salida from movisto mo with(nolock) inner join master ma with(nolock) on (mo.cod_inte=ma.codint)" _
    & " Where mo.cod_inte > 0 And mo.depomovi = " & CStr(DEPOSITO%) & " And ma.Status > 0" _
    & " group by mo.cod_inte) stock50"
    Set rst_TotalComponentes = READSET(ConsultaSql$)
    With rst_TotalComponentes
        If Not (.EOF And .BOF) Then
            TotalComponentes& = !maxitem
        End If
    End With
    On Error Resume Next
    rst_TotalComponentes.Close
    Set rst_TotalComponentes = Nothing
    On Error GoTo 0
    
    ' *** MATRIZ COMPONENTES
    Dim StockComponentes() As Double
    ReDim StockComponentes(TotalComponentes&)

    ' *** OBTENGO STOCK DE COMPONENTES
    ConsultaSql$ = "select mo.cod_inte, sum(mo.cantingre) ingreso, sum(mo.cantsalid) salida from movisto mo with(nolock) inner join master ma with(nolock) on (mo.cod_inte=ma.codint)" _
    & " Where mo.cod_inte > 0 And mo.depomovi = 50 And ma.Status > 0" _
    & " group by mo.cod_inte"
    Set rst_StockComponentes = READSET(ConsultaSql$)
    With rst_StockComponentes
        Do While Not .EOF
            CodigoInterno% = XVALO(!cod_inte)
            StockComponentes(CodigoInterno%) = XVALO(!ingreso) - XVALO(!SALIDA)
            
            .MoveNext
        Loop
    End With
    On Error Resume Next
    rst_StockComponentes.Close
    Set rst_StockComponentes = Nothing
    On Error GoTo 0
    
    ' *** OBTENGO PENDIENTE DE RECEPCION
    ConsultaSql$ = "select cod_inte, cantnom from bolrecep with(nolock) where status<1"
    Set rst_RecepcionesPendientes = READSET(ConsultaSql$)
    With rst_RecepcionesPendientes
        Do While Not .EOF
            CodigoInterno% = XVALO(!cod_inte)
            If CodigoInterno% > TotalComponentes& Then
                ReDim Preserve StockComponentes(CodigoInterno%)
            End If
            StockComponentes(CodigoInterno%) = StockComponentes(CodigoInterno%) + XVALO(!CantNom)
            
            .MoveNext
        Loop
    End With
    On Error Resume Next
    rst_RecepcionesPendientes.Close
    Set rst_RecepcionesPendientes = Nothing
    On Error GoTo 0
    
    If IncluyeCompras% = 2 Then
        ' *** OBTENGO PENDIENTE DE COMPRAS
        ConsultaSql$ = "SELECT od.cod_inte, sum(od.cant_ocom) compra, sum(od.cant_rec) recibido FROM ocomdeta od with(nolock) inner join ocomenca oc with(nolock) on (od.nume_ocom=oc.nume_ocom) where oc.stat_ocom<9 and od.stat_ocom<3 and od.doc_orig='' group by cod_inte"
        Set rst_ComprasPendientes = READSET(ConsultaSql$)
        With rst_ComprasPendientes
            Do While Not .EOF
                CodigoInterno% = XVALO(!cod_inte)
                If CodigoInterno% > TotalComponentes& Then
                    ReDim Preserve StockComponentes(CodigoInterno%)
                End If
                SALDO# = XVALO(!compra) - XVALO(!recibido)
                If SALDO# < 0.05 Then SALDO# = 0
                StockComponentes(CodigoInterno%) = StockComponentes(CodigoInterno%) + SALDO#
                
                .MoveNext
            Loop
        End With
        On Error Resume Next
        rst_ComprasPendientes.Close
        Set rst_ComprasPendientes = Nothing
        On Error GoTo 0
    End If
    
    ' *** OBTENGO LISTA DE CONJUNTOS A ANALIZAR
    Dim Conjuntos() As Integer: ReDim Conjuntos(2048)
    Dim ConjuntosMensaje() As String: ReDim ConjuntosMensaje(2048)
    Dim TotalDeConjuntos As Integer
    Dim IConjunto As Integer
    Dim SubConjuntos(1024, 2) As Double                         ' (ci, uso)
    Dim ISubConjunto As Integer
    Dim Componentes() As Double: ReDim Componentes(32767)       ' (uso)
    
    IDLIST& = ObtenerIdTablaAuxiliar("ANALISTO")
    If CantRegistros("datasql", "id_lista=" & CStr(IDLIST&)) > 0 Then
        ConsultaSql$ = "select campo2 from datasql with(nolock) where id_lista=" & CStr(IDLIST&) & " order by cast(campo1 as int)"
        Dim RstCodigos As ADODB.Recordset
        Set RstCodigos = READSET(ConsultaSql$)
        
        With RstCodigos
            Do While Not .EOF
                CodigoProducto$ = SAFETEXT$(!campo2)
                TotalDeConjuntos = TotalDeConjuntos + 1
                Conjuntos(TotalDeConjuntos) = CODINT%(CodigoProducto$)
                
                .MoveNext
            Loop
        End With
        On Error Resume Next
        RstCodigos.Close
        Set RstCodigos = Nothing
        On Error GoTo 0
    End If
    
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, "Codigo/A24;Descripcion/A48;A Fabricar/F14.2;Faltantes/A48", FRMZELECHO)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, "Codigo/A24;Descripcion/A48;A Fabricar/F14.2;Faltantes/A48", FRMZELECHO)
    GRILLAGRAL.GRID.Rows = TotalDeConjuntos + 1
    
    ' *** RECORRE CONJUNTOS
    For IConjunto = 1 To TotalDeConjuntos
        For IX = 1 To 1023
            SubConjuntos(IX, 1) = 0
        Next IX
        Erase Componentes
        ReDim Componentes(32767)
        
        SubConjuntos(1, 1) = Conjuntos(IConjunto)
        SubConjuntos(1, 2) = 1
        ISubConjunto = 1
        
        ' *** RECORRE SUBCONJUNTOS
        For i = 1 To 1023
            ' *** OBTENGO ESTRUCTURA MULTINIVEL
            CodigoConjunto% = SubConjuntos(i, 1)
            If CodigoConjunto% > 0 Then
                ConsultaSql$ = "select codielem, usobruto from formulas fo with(nolock) inner join master ma with(nolock) on (fo.codielem=ma.codint) " _
                & " where fo.codiconj=" & CStr(CodigoConjunto%) & " and ma.status>0"
                Set rst_Componentes = READSET(ConsultaSql$)
                With rst_Componentes
                    Do While Not .EOF
                        CodigoElemento% = XVALO(!codielem)
                        USO# = XVALO(!usobruto)
                        If CantRegistros("formulas", "codiconj=" & CStr(CodigoElemento%)) > 0 Then
                            ' ES SUBCONJUNTO
                            ISubConjunto = ISubConjunto + 1
                            SubConjuntos(ISubConjunto, 1) = CodigoElemento%
                            SubConjuntos(ISubConjunto, 2) = USO# * SubConjuntos(i, 2)
                        Else
                            ' ES COMPONENTE
                            Componentes(CodigoElemento%) = Componentes(CodigoElemento%) + (USO# * SubConjuntos(i, 2))
                        End If
                        .MoveNext
                    Loop
                End With
            Else
                Exit For
            End If
        Next i
        
        ' *** ANALIZO CUANTO PUEDO FABRICAR
        MaximoAFabricar& = 500000
        Entro% = 0
        For i = 1 To 32766
            If Componentes(i) > 0.05 Then
                Entro% = 1
                If StockComponentes(i) > 0.05 Then
                    AFabricarSegunComponente& = Int(StockComponentes(i) / Componentes(i))
                    
                    If AFabricarSegunComponente& < MaximoAFabricar& Then
                        MaximoAFabricar& = AFabricarSegunComponente&
                        
                        If MaximoAFabricar& = 0 Then
                            If ConjuntosMensaje$(IConjunto) = "" Then
                                ConjuntosMensaje$(IConjunto) = CODEXT$(CInt(i)) '& " - (" & CStr(Componentes(i)) & ")"
                            Else
                                ConjuntosMensaje$(IConjunto) = ConjuntosMensaje$(IConjunto) & "; " & CODEXT$(CInt(i)) '& " - (" & CStr(Componentes(i)) & ")"
                            End If
                        End If
                    End If
                Else
                    MaximoAFabricar& = 0
                    If ConjuntosMensaje$(IConjunto) = "" Then
                        ConjuntosMensaje$(IConjunto) = CODEXT$(CInt(i)) '& " - (" & CStr(Componentes(i)) & ")"
                    Else
                        ConjuntosMensaje$(IConjunto) = ConjuntosMensaje$(IConjunto) & "; " & CODEXT$(CInt(i)) '& " - (" & CStr(Componentes(i)) & ")"
                    End If
                End If
            End If
        Next i
        
        ' ***
        If Entro% = 0 Then MaximoAFabricar& = 0
        ' *** DESCUENTA DEL STOCK DE LOS COMPONENTES EL UTILIZADO PARA FABRICAR EL CONJUNTO
        If MaximoAFabricar& > 0 Then
            For i = 1 To 32766
                If Componentes(i) > 0.05 Then
                    StockComponentes(i) = StockComponentes(i) - Componentes(i) * MaximoAFabricar&
                End If
            Next i
        End If
        
        GRILLAGRAL.GRID.TextMatrix(IConjunto, 1) = CODEXT$(Conjuntos(IConjunto))
        GRILLAGRAL.GRID.TextMatrix(IConjunto, 2) = DESCRIT0$(Conjuntos(IConjunto))
        GRILLAGRAL.GRID.TextMatrix(IConjunto, 3) = CStr(MaximoAFabricar&)
        GRILLAGRAL.GRID.TextMatrix(IConjunto, 4) = ConjuntosMensaje$(IConjunto)
        
    Next IConjunto
    
    GRILLAGRAL.Caption = "Análisis de Fabricación Según Stock y Recepciones"
    GRILLAGRAL.mnuMenu1.Caption = "Exportar a Excel"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    
    GRILLAGRAL.Show 1
    
    Dim MSF As msFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
        CommonDialog1.CancelError = True
        CommonDialog1.DialogTitle = "Exportar a Excel"
            ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
        CommonDialog1.FILTER = "Excel Files(*.XLS)|*.XLS|"
        CommonDialog1.FilterIndex = 1
        On Error GoTo 99
        CommonDialog1.ShowOpen
        On Error GoTo 0
        '
        DoEvents
        Screen.MousePointer = 11
        AREXPO$ = UCase$(CommonDialog1.FileName)
        If TRIM$(AREXPO$) <> "" Then
           Call GeneraExcelDeGrilla(AREXPO$, GRILLAGRAL.GRID, GRILLAGRAL.Caption, False)
        End If
   End If
   Set obj = Nothing

   Exit Sub
   
99:
    Exit Sub
    
End Sub

Private Sub mnuImportarConjuntosDesdeExcel_Click()

    Dim ta As New DatosAuxiliares
    
    On Error GoTo NoExisteExcel
    Set objexcel = New Excel.Application
    On Error GoTo 0
    
    CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
    CommonDialog1.DialogTitle = "Archivo Origen de Datos"
    CommonDialog1.FILTER = "Hojas de Cálculo Excel(*.XLS)|*.XLS|Hojas de Cálculo Excel(*.XLSX)|*.XLSX"
    CommonDialog1.FilterIndex = 1
    On Error GoTo 99
    CommonDialog1.ShowOpen
    On Error GoTo 0
    '
    DoEvents
    Screen.MousePointer = 11
    '
    Ruta$ = UCase$(CommonDialog1.FileName)
    Set Libro = objexcel.Workbooks.Open(Ruta$)
    Set Hoja = Libro.Sheets(1)
    '
    URECORD& = UltimoRegistro(1)
    If URECORD& > 0 Then
        idTabla& = ta.ObtenerIdTablaAuxiliar("ANALISTO")
        Call BORRAREGISTROS("datasql", "lista='ANALISTO'", REGAF&)
    End If
    '
    Dim Rst As New ADODB.Recordset
    CosnultaSql$ = "select * from datasql where id_lista=0"
    Rst.Open FILTSQL$(CosnultaSql$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    For Fila& = 1 To URECORD&
        '
        On Error Resume Next
        Codigo$ = TRIM$(Hoja.Cells(Fila&, 1))
        CI% = CODINT%(Codigo$)
        On Error GoTo 0
        If CI% < 1 Then
            Call MENSERR(24, "Código '" & Codigo$ & "' Inexistente.")
            GoTo 20
        End If

        If Codigo$ <> "" Then
            U& = U& + 1
            Rst.AddNew
            Rst!ID_Lista = idTabla&
            Rst!Lista = "ANALISTO"
            Rst!CAMPO1 = CStr(U&)
            Rst!campo2 = Codigo$
            Rst!CAMPO3 = DESCRIT0$(CI%)
            Rst.Update
        End If
        '
20 Next Fila&
   '
   Libro.Close
   objexcel.Quit

   Screen.MousePointer = 1
   Call MENSERR(24, "Proceso de Importación Finalizada")
   
99 Exit Sub

NoExisteExcel:
    Call MENSERR(24, "No Se Encuentra Excel Instalado.")
    GoTo 99
    
End Sub

Private Sub mnuKgIngVsSalKg_Click()
'       Call DESHASFECHA(Desde%, Hasta%, PERIODO$)
'        If PERIODO$ = "" Then Exit Sub
'        DesdeFecha$ = FETEXCOR1$(Desde%)
'        HastaFecha$ = FETEXCOR1$(Hasta%)
'        FECHAHOY$ = FETEXCOR1$(HOY(HO$))
        
        Call SELECHO("SELFECHA")
        Desde% = FECHANUM(VALACT1$("SELFECHA"))
        DesdeFecha$ = FETEXCOR1$(Desde%)
        
        Campos$ = "O.Fabricación/A12"  '1
        Campos$ = Campos$ + ";Pedido/F10" '2
        Campos$ = Campos$ + ";Código/A16" '3
        Campos$ = Campos$ + ";Descripción/A32" '4
        Campos$ = Campos$ + ";KG.Transf./F10" '5
        Campos$ = Campos$ + ";KG.Inf.Corte/F10" '6
        Campos$ = Campos$ + ";KG.Inf.Corte/F10" '7
        Campos$ = Campos$ + ";CI/A0" '8


    
        J& = 0
        Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
        Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)

        
        'ESTA CONSULTA TRAE TODAS LAS OF, QUE SU PRIMER CIERRE EN EL MOVISTO SE PRODUJO DENTRO DE LAS FECHA
        'EJ, RANGO SELECCIONADO 01/01/24 AL 31/01/24 SI LA OF SE GENERO ANTES DEL DESDE PERO SU PRIMER CIERRE FUE
        'DENTRO DEL RANGO LA TRAE,SI SU PRIMER CIERRE FUE FUERA DEL RANGO NO LA TRAE
        '
        SQLX$ = "SELECT DOPRILAR,COD_INTE,COD_EXTE,DESCRIP FROM MOVISTO WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE DOPRILAR IN (  SELECT DOPRILAR FROM MOVISTO   "
        SQLX$ = SQLX$ + " GROUP BY DOPRILAR HAVING MIN(FECHMOV) >= '" & DesdeFecha$ & "'"
        SQLX$ = SQLX$ + " ) AND CodiMovi = 'OF'"
        SQLX$ = SQLX$ + " GROUP BY DOPRILAR,COD_INTE,COD_EXTE,DESCRIP"
        SQLX$ = SQLX$ + " ORDER BY DOPRILAR "
        Set rs = READSET(SQLX$)
        With rs
          Do While Not .EOF
             NORFA$ = SAFETEXT$(!DOPRILAR)
             If STATORFA%(NORFA$) <> 3 Then GoTo 60
             CI1% = XVALO(!cod_inte)
             CONDI$ = "depomovi = 2   AND CODIMOVI = 'TR' and DOSECLAR = '" & NORFA$ & "'"
              KilosGramosTransferidosAlDeposito2XOF = XVALO(VALOBADA("MOVISTO", "SUM(CASE WHEN CODIMOVI = 'TR' AND Cantingre > 0 THEN pesokg ELSE -pesokg END)", CONDI$, "NOMESS"))
             CONDI$ = "IDSUBOR = '" & NORFA$ & "'"
              KilosGramosConsumidosXOF = XVALO(VALOBADA("INFOFAB", "SUM(PESO)", CONDI$, "NOMESS"))
             CONDI$ = "CODIMOVI = 'OF' and DOPRILAR = '" & NORFA$ & "'"
             KilosGramosProdTerminadoXOF = XVALO(VALOBADA("MOVISTO", "SUM(CASE WHEN CODIMOVI = 'OF' AND Cantingre > 0 THEN pesokg ELSE -pesokg END)", CONDI$, "NOMESS"))
             
             J& = J& + 1
             GRILLAGRAL.GRID.Rows = J& + 1
             GRILLAGRAL.GRID.TextMatrix(J&, 1) = NORFA$
             GRILLAGRAL.GRID.TextMatrix(J&, 2) = NroPedidoSegunOf&(NORFA$)
             GRILLAGRAL.GRID.TextMatrix(J&, 3) = SAFETEXT$(!Cod_Exte)
             GRILLAGRAL.GRID.TextMatrix(J&, 4) = SAFETEXT$(!Descrip)
             GRILLAGRAL.GRID.TextMatrix(J&, 5) = FORMATNUM$(KilosGramosTransferidosAlDeposito2XOF, "F10.1")
             GRILLAGRAL.GRID.TextMatrix(J&, 6) = FORMATNUM$(KilosGramosConsumidosXOF, "F10.1")
             GRILLAGRAL.GRID.TextMatrix(J&, 7) = FORMATNUM$(KilosGramosProdTerminadoXOF, "F10.1")
             GRILLAGRAL.GRID.TextMatrix(J&, 8) = CI1%

60          .MoveNext
          Loop
        End With
        rs.Close
    Set rs = Nothing
   

   
   GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
   GRILLAGRAL.mnuMenu1.Visible = True
   GRILLAGRAL.GRID.ZOrder 0
   GRILLAGRAL.TXTBUSCAR = ""
   TX$ = "Reporte De Ingresos Vs.Salidas en KG.(Solo Considera O.Fab. Cerradas) Rango Desde: " & FECHATEX$(Desde%) & " Hasta: " & FECHATEX$(Hasta%)
   
   GRILLAGRAL.Caption = TX$
   Screen.MousePointer = 1
   GRILLAGRAL.Show 1
   
   NORFA$ = RESP_ZELE_VECT(1)
   
   Dim MSF As msFlexGrid
   If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(MSF, Campos$, "INGVSSAL", " Reporte De Ingresos Vs.Salidas", "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
        Exit Sub
   End If
99 Exit Sub

        
End Sub

Private Sub mnuListadoDisponibilidadPorRango_Click()

    Campos$ = "Rango/A16;Descripcion/A48;Stock/G14.0;Compras/G14.0;O-Fabricación/G14.0;Notas de Ventas/G14.0;Stock Seguridad/G14;Positivo/G14.0;Negativo/G14.0"
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)

    ConsultaAgrupaPorRangoIni$ = "select stcomres.rango, sum(stcomres.stock) stock, sum(stcomres.compras) compras, sum(stcomres.reservasof) reservasof, sum(stcomres.reservaspedidos) reservaspedidos from" & vbCrLf
    ConsultaAgrupaPorRangoIni$ = ConsultaAgrupaPorRangoIni$ & " (" & vbCrLf

    ConsultaSqlStock$ = "select" & vbCrLf
    ConsultaSqlStock$ = ConsultaSqlStock$ & " case when cr.codrango is null then st.cod_exte" & vbCrLf
    ConsultaSqlStock$ = ConsultaSqlStock$ & " else cr.codrango" & vbCrLf
    ConsultaSqlStock$ = ConsultaSqlStock$ & " end as rango, stock Stock, (0) Compras, (0) ReservasOF, (0) ReservasPedidos from" & vbCrLf
    ConsultaSqlStock$ = ConsultaSqlStock$ & " codigorango cr inner join" & vbCrLf
    ConsultaSqlStock$ = ConsultaSqlStock$ & " (select cod_inte, cod_exte, sum(cantingre)-sum(cantsalid) stock from movisto where depomovi=1 group by cod_inte, cod_exte) st" & vbCrLf
    ConsultaSqlStock$ = ConsultaSqlStock$ & " on (st.cod_inte=cr.cimatprima)" & vbCrLf


    ConsultaSqlCompras$ = "select" & vbCrLf
    ConsultaSqlCompras$ = ConsultaSqlCompras$ & " case when cr.codrango is null then compendneta.cod_exte" & vbCrLf
    ConsultaSqlCompras$ = ConsultaSqlCompras$ & " else cr.codrango" & vbCrLf
    ConsultaSqlCompras$ = ConsultaSqlCompras$ & " end as rango, (0) Stock, Compras, (0) ReservasOF, (0) ReservasPedidos from" & vbCrLf
    ConsultaSqlCompras$ = ConsultaSqlCompras$ & " (select compend.cod_inte, compend.cod_exte," & vbCrLf
    ConsultaSqlCompras$ = ConsultaSqlCompras$ & " case when compend.Compras > 0 then compend.Compras" & vbCrLf
    ConsultaSqlCompras$ = ConsultaSqlCompras$ & " else 0" & vbCrLf
    ConsultaSqlCompras$ = ConsultaSqlCompras$ & " end as Compras from" & vbCrLf
    ConsultaSqlCompras$ = ConsultaSqlCompras$ & " (select cod_inte, cod_exte, sum(cant_ocom)-sum(cant_rec) compras from ocomdeta where stat_ocom<3 group by cod_inte, cod_exte) compend) compendneta" & vbCrLf
    ConsultaSqlCompras$ = ConsultaSqlCompras$ & " left join codigorango cr on (compendneta.cod_inte=cr.cimatprima)" & vbCrLf


    ConsultaSqlReservasOF$ = "select reservaofnetatran.rango, (0) Stock, (0) Compras," & vbCrLf
    ConsultaSqlReservasOF$ = ConsultaSqlReservasOF$ & " case when reservaofnetatran.reservado>reservaofnetatran.reserva then 0" & vbCrLf
    ConsultaSqlReservasOF$ = ConsultaSqlReservasOF$ & " else reservaofnetatran.Reserva -reservaofnetatran.reservado" & vbCrLf
    ConsultaSqlReservasOF$ = ConsultaSqlReservasOF$ & " end as ReservasOF, (0) ReservasPedidos from" & vbCrLf
    ConsultaSqlReservasOF$ = ConsultaSqlReservasOF$ & " (select reservaneta.rango, sum(reservaneta.reserva) Reserva, sum(reservaneta.reservado) Reservado from" & vbCrLf
    ConsultaSqlReservasOF$ = ConsultaSqlReservasOF$ & " (select case when resra.codrango is null then resra.cod_exte" & vbCrLf
    ConsultaSqlReservasOF$ = ConsultaSqlReservasOF$ & " else resra.codrango" & vbCrLf
    ConsultaSqlReservasOF$ = ConsultaSqlReservasOF$ & " end as rango, resra.reserva, (0) reservado from" & vbCrLf
    ConsultaSqlReservasOF$ = ConsultaSqlReservasOF$ & " (select r.cod_exte, cr.codrango, r.reserva from" & vbCrLf
    ConsultaSqlReservasOF$ = ConsultaSqlReservasOF$ & " (select cod_exte, cod_inte, sum(cantres) reserva from reserva group by cod_exte, cod_inte) r" & vbCrLf
    ConsultaSqlReservasOF$ = ConsultaSqlReservasOF$ & " left join codigorango cr on (r.Cod_Inte=cr.CIMATPRIMA)) resra" & vbCrLf
    ConsultaSqlReservasOF$ = ConsultaSqlReservasOF$ & " Union All" & vbCrLf
    
    ConsultaSqlReservasOF$ = ConsultaSqlReservasOF$ & " select resofra.rango, (0) reserva, sum(resofra.reservado) from " & vbCrLf
    ConsultaSqlReservasOF$ = ConsultaSqlReservasOF$ & " (select case when cr.codrango is null then resof.cod_exte else cr.codrango end as rango, reservado from" & vbCrLf
    ConsultaSqlReservasOF$ = ConsultaSqlReservasOF$ & " (select m.Cod_Inte, m.cod_exte, sum(m.cantsalid)-sum(m.cantingre) reservado from" & vbCrLf
    ConsultaSqlReservasOF$ = ConsultaSqlReservasOF$ & " (select distinct idsubor from reserva) r" & vbCrLf
    ConsultaSqlReservasOF$ = ConsultaSqlReservasOF$ & " inner join movisto m on (r.IdSubOr=m.DoSecLar)" & vbCrLf
    ConsultaSqlReservasOF$ = ConsultaSqlReservasOF$ & " where m.depomovi=1 and codimovi='TR' group by m.Cod_Inte, cod_exte) resof" & vbCrLf
    ConsultaSqlReservasOF$ = ConsultaSqlReservasOF$ & " inner join codigorango cr on (resof.Cod_Inte=cr.CIMATPRIMA)) resofra" & vbCrLf
    ConsultaSqlReservasOF$ = ConsultaSqlReservasOF$ & " group by resofra.rango) reservaneta" & vbCrLf
    ConsultaSqlReservasOF$ = ConsultaSqlReservasOF$ & " group by reservaneta.rango) reservaofnetatran" & vbCrLf

    ConsultaSqlReservasPedidos$ = "select reservapednetadeof.rango, (0) Stock, (0) Compras, (0) ReservaOF, sum(reservapednetadeof.reservapedido) ReservaPedidos from" & vbCrLf
    ConsultaSqlReservasPedidos$ = ConsultaSqlReservasPedidos$ & " (" & vbCrLf
    ConsultaSqlReservasPedidos$ = ConsultaSqlReservasPedidos$ & " select forran.rango, pednetofab.CantPedidaNetoFabrica*forran.usobruto reservapedido from" & vbCrLf
    ConsultaSqlReservasPedidos$ = ConsultaSqlReservasPedidos$ & " (select pedfab.Pedido, pedfab.CodIConj," & vbCrLf
    ConsultaSqlReservasPedidos$ = ConsultaSqlReservasPedidos$ & " case when pedfab.CantFabricada is null then pedfab.CantPedida" & vbCrLf
    ConsultaSqlReservasPedidos$ = ConsultaSqlReservasPedidos$ & " when pedfab.CantFabricada>pedfab.CantPedida then 0" & vbCrLf
    ConsultaSqlReservasPedidos$ = ConsultaSqlReservasPedidos$ & " else pedfab.CANTPEDIDA -pedfab.CantFabricada" & vbCrLf
    ConsultaSqlReservasPedidos$ = ConsultaSqlReservasPedidos$ & " end as CantPedidaNetoFabrica from" & vbCrLf
    ConsultaSqlReservasPedidos$ = ConsultaSqlReservasPedidos$ & " (select pedpen.nroped Pedido, pedpen.codiint CodIConj, pedpen.canped CantPedida, sum(o.CanProy) CantFabricada from" & vbCrLf
    ConsultaSqlReservasPedidos$ = ConsultaSqlReservasPedidos$ & " (select pt.nroped, pt.codiint, pt.canped from pedenca pd inner join peddeta pt on (pd.nroped=pt.nroped) where pd.status<3 and pt.Status<3) pedpen" & vbCrLf
    ConsultaSqlReservasPedidos$ = ConsultaSqlReservasPedidos$ & " left join ordefabr o on (pedpen.nroped=o.nroped and pedpen.codiint=o.Cod_Inte)" & vbCrLf
    ConsultaSqlReservasPedidos$ = ConsultaSqlReservasPedidos$ & " group by pedpen.nroped, pedpen.codiint, pedpen.canped) pedfab) pednetofab" & vbCrLf
    ConsultaSqlReservasPedidos$ = ConsultaSqlReservasPedidos$ & " Inner Join" & vbCrLf
    ConsultaSqlReservasPedidos$ = ConsultaSqlReservasPedidos$ & " (select fr.codiconj," & vbCrLf
    ConsultaSqlReservasPedidos$ = ConsultaSqlReservasPedidos$ & " case when fr.codrango is null then fr.codxelem" & vbCrLf
    ConsultaSqlReservasPedidos$ = ConsultaSqlReservasPedidos$ & " else fr.codrango" & vbCrLf
    ConsultaSqlReservasPedidos$ = ConsultaSqlReservasPedidos$ & " end as rango, fr.usobruto from" & vbCrLf
    ConsultaSqlReservasPedidos$ = ConsultaSqlReservasPedidos$ & " (select f.codiconj, f.codxelem, cr.codrango, f.usobruto from" & vbCrLf
    ConsultaSqlReservasPedidos$ = ConsultaSqlReservasPedidos$ & " (select codiconj, codielem, codxelem, usobruto from formulas) f" & vbCrLf
    ConsultaSqlReservasPedidos$ = ConsultaSqlReservasPedidos$ & " left join codigorango cr on (f.codielem=cr.cimatprima)) fr" & vbCrLf
    ConsultaSqlReservasPedidos$ = ConsultaSqlReservasPedidos$ & " ) forran" & vbCrLf
    ConsultaSqlReservasPedidos$ = ConsultaSqlReservasPedidos$ & " on (pednetofab.CodIConj=forran.CodIConj)" & vbCrLf
    ConsultaSqlReservasPedidos$ = ConsultaSqlReservasPedidos$ & " ) reservapednetadeof" & vbCrLf
    ConsultaSqlReservasPedidos$ = ConsultaSqlReservasPedidos$ & " group by reservapednetadeof.rango" & vbCrLf

    ConsultaAgrupaPorRangoFin$ = ") stcomres" & vbCrLf
    ConsultaAgrupaPorRangoFin$ = ConsultaAgrupaPorRangoFin$ & " group by stcomres.rango" & vbCrLf
    ConsultaAgrupaPorRangoFin$ = ConsultaAgrupaPorRangoFin$ & " order by stcomres.rango"

    ConsultaSql$ = ""
    ConsultaSql$ = ConsultaSql$ & ConsultaAgrupaPorRangoIni$
    ConsultaSql$ = ConsultaSql$ & ConsultaSqlStock$
    ConsultaSql$ = ConsultaSql$ & " union all "
    ConsultaSql$ = ConsultaSql$ & ConsultaSqlCompras$
    ConsultaSql$ = ConsultaSql$ & " union all "
    ConsultaSql$ = ConsultaSql$ & ConsultaSqlReservasOF$
    ConsultaSql$ = ConsultaSql$ & " union all "
    ConsultaSql$ = ConsultaSql$ & ConsultaSqlReservasPedidos$
    ConsultaSql$ = ConsultaSql$ & ConsultaAgrupaPorRangoFin$

    Dim Rst As ADODB.Recordset
    Set Rst = FLEXCONN.Execute(ConsultaSql$)
    
    With Rst
        Do While Not Rst.EOF
            Rango$ = TRIM$(!Rango)
            STOCK# = XVALO(!STOCK)
            compra# = XVALO(!compras)
            'reservaof# = XVALO(!reservasof)
            reservaof# = XVALO(ReservaMenosTransfCodRango#(Rango$))
            reservapedidos# = XVALO(!reservaspedidos)
                    
            
            With GRILLAGRAL.GRID
                CI1% = CODINT(Rango$)
                If ESCODIGORANGO&(CI1%) > 0 Then 'PIDIO RAUL QUE SI NO SON CODIGOS RANGOS NO LOS PRESENTE
                  If XVALO(Rango$) <= 700 Then   'TAMBIEN PEDIDO POR RAUL
                        Fila& = Fila& + 1
                        .Rows = .Rows + 1
                        .TextMatrix(Fila&, 1) = Rango$
                        .TextMatrix(Fila&, 2) = VALOBADA("codigorango", "descripcion", "codrango='" & Rango$ & "'")
                        .TextMatrix(Fila&, 3) = FORMATNUM$(STOCK#, "G14.0")
                        .TextMatrix(Fila&, 4) = FORMATNUM$(compra#, "G14.0")
                        .TextMatrix(Fila&, 5) = FORMATNUM$(reservaof#, "G14.0")
                        .TextMatrix(Fila&, 6) = FORMATNUM$(reservapedidos#, "G14.0")
                            stockseguridad# = XVALO(VALOBADA("datasql", "campo3", "campo1='" & Rango$ & "'"))
                        .TextMatrix(Fila&, 7) = FORMATNUM$(stockseguridad#, "G14.0")
                        
                        NETO# = STOCK# + compra# - reservaof# - reservapedidos# - stockseguridad#
                        If NETO# > 0 Then
                            .TextMatrix(Fila&, 8) = FORMATNUM$(NETO#, "G14.0")
                        Else
                            .TextMatrix(Fila&, 9) = FORMATNUM$(Abs(NETO#), "G14.0")
                        End If
                  End If
                End If
            End With
            
            .MoveNext
        Loop
    End With
    
    On Error Resume Next
    Rst.Close
    Set Rst = Nothing
    On Error GoTo 0
    
    ' ********************
    GRILLAGRAL.Caption = "Disponibilidad de Stock por Rango"
    
    GRILLAGRAL.mnuMenu1.Caption = "Exportar a Excel"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.MnuEdición.Visible = False
    GRILLAGRAL.SinEfectoClickEnGrilla = 1
    GRILLAGRAL.GRID.ZOrder 0
    
    Screen.MousePointer = 1
    GRILLAGRAL.Show 1
    Dim MSF As msFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
        CommonDialog1.CancelError = True
        CommonDialog1.DialogTitle = "Exportar a Excel"
            ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
        CommonDialog1.FILTER = "Excel Files(*.XLS)|*.XLS|"
        CommonDialog1.FilterIndex = 1
        On Error GoTo 99
        CommonDialog1.ShowOpen
        On Error GoTo 0
        '
        DoEvents
        Screen.MousePointer = 11
        AREXPO$ = UCase$(CommonDialog1.FileName)
        If TRIM$(AREXPO$) <> "" Then
           Call GeneraExcelDeGrilla(AREXPO$, GRILLAGRAL.GRID, GRILLAGRAL.Caption, False, Campos$)
        End If
   End If
   Set obj = Nothing

99 Exit Sub

End Sub

Function ReservaMenosTransfCodRango#(CODIRANGO$)
   'ESTA FUNCION  CTP.CONSULTA_OR DEVUELVE UNA CONDICION CON LOS VALORES POSIBLES DEL CODIGO RANGO
   Dim CTP As New Conteplast
   
   CONDI$ = "CODRANGO = '" & CODIRANGO$ & "'"
   CONDIOR$ = CTP.CONSULTA_OR("CODIGORANGO", "CODMATPRIMA", CONDI$, "COD_EXTE")
   CONDINORFA$ = ""
   TotCantReservada_Transf# = 0
   'PRIMERO CONSULTO LAS RESERVAS CORRESPONDIENTES AL CODIGO RANGO(EN LAS RESERVAS FIGURAN CON EL CODIGO RANGO)
   'INNER JOIN CON ORDEFABR CON STATUS < 3
   SQLX$ = " select r.Cod_Exte,r.Cod_Inte,r.idsubor, sum(r.CantRes) as Canti "
   SQLX$ = SQLX$ + " from reserva r inner join ordefabr o WITH(NOLOCK)"
   SQLX$ = SQLX$ + " on o.idsubor= r.IdSubOr"
   SQLX$ = SQLX$ + " where r.COD_EXTE = '" & CODIRANGO$ & "'"
   SQLX$ = SQLX$ + " and  o.StatuOrf < 3"
   SQLX$ = SQLX$ + " group by r.Cod_Exte,r.Cod_Inte,r.idsubor"
'Call mostrarDatosx("1 " & SQLX$)
   Set rs = READSET(SQLX$)
   With rs
     Do While Not .EOF
        NORFA$ = SAFETEXT$(!idsubor)
        'CantiTransferida_X_OF_Y_CodigosAsociadosalRango# devuelve la canti transferida para el param NORFA$ Y LOS CODIGOS REALES DEL RANGO
        CANTITRANSFERIDA# = CTP.CantiTransferida_X_OF_Y_CodigosAsociadosalRango#(NORFA$, CONDIOR$)
        parcialCantReservada# = XVALO(!CANTI) 'CANTIDAD RESERVADA DE LA OF DE LOS CODIGOS RANGO DEL RESERVA
        '
        ParcialResultado# = parcialCantReservada# - CANTITRANSFERIDA#
        If ParcialResultado# <= 0 Then ParcialResultado# = 0 'SI SE TRANSFIRIO MAS DE LO RESERVADO ASIGNA 0
        TotCantReservada_Transf# = TotCantReservada_Transf# + ParcialResultado#
        .MoveNext
     Loop
   End With
   
'   If CONDINORFA$ <> "" Then CONDINORFA$ = CONDINORFA$ & " )"
   rs.Close
   Set rs = Nothing
   
   ReservaMenosTransfCodRango# = TotCantReservada_Transf#
'   If TotCantReservada_Transf# < 1 Then 'SI NO POSEE RESERVAS DEVUELVE 0 Y SE VAA
'      ReservaMenosTransfCodRango# = 0
'      Exit Function
'   End If
'
'   'AHORA VERIFICO LA CANTIDAD TRANSFERIDA DEL DEPOSITO 1 A OTROS (UTILIZANDO LAS CONDICIONES CARGADAS ARRIBA.
'   TotCantReservada_Transf# = 0
'   SQLX$ = "SELECT  sum(CANTSALID-CANTINGRE) AS CANTI  FROM MOVISTO WITH(NOLOCK) "
'   SQLX$ = SQLX$ + " WHERE CODIMOVI = 'TR'  AND CANTSALID > 0 and depomovi = 1  "
'   If CONDIOR$ <> "" Then SQLX$ = SQLX$ + " AND " & CONDIOR$
'   If CONDINORFA$ <> "" Then SQLX$ = SQLX$ + " AND " & CONDINORFA$
'Call mostrarDatosx("2 " & SQLX$)
'   Set rs = READSET(SQLX$)
'   With rs
'     Do While Not .EOF
'       TotCantTransferida# = TotCantTransferida# + XVALO(!CANTI)
'       .MoveNext
'     Loop
'   End With
'   rs.Close
'   Set rs = Nothing
'
'  ReservaMenosTransfCodRango# = XVALO(TotCantReservada_Transf#) - XVALO(TotCantTransferida#)
'  Set CTP = Nothing
'
'Call mostrarDatosx("3 " & SAFETEXT$(TotCantReservada_Transf#) & "  " & SAFETEXT$(TotCantTransferida#))
'  'EJEMPLO COMO QUEDA LA CONSULTA AL MOVISTO
''  SELECT  sum(CANTSALID-CANTINGRE) AS CANTI  FROM MOVISTO
''WHERE CODIMOVI = 'TR'  AND CANTSALID > 0 and depomovi = 1
''AND( DOSECLAR ='OF-000062-1'
''OR DOSECLAR ='OF-000134-1'
''OR DOSECLAR ='OF-000151-1'
''OR DOSECLAR ='OF-000332-1'
''OR DOSECLAR ='OF-000340-1'
''OR DOSECLAR ='OF-000358-1'
''OR DOSECLAR ='OF-000449-1'
''OR DOSECLAR ='OF-000491-1'
''OR DOSECLAR ='OF-000492-1'
'')
''AND (COD_EXTE = 'CA TP 101 190 BL'
''OR COD_EXTE ='CA TPL 101 74 BL' )

End Function

Private Sub mnuMovStLuzuriaga_Click()
 
 
20  FILTROFECHA% = 0
    FILTROCODIGO% = 0
    OBJX$ = OBJPLA$("CODIFECH", VDEF$)
    If OBJX$ <> "" Then
        DESA% = CVI(Mid$(OBJX$, 1, 2))
        HASA% = CVI(Mid$(OBJX$, 3, 2))
        CI1% = CVI(Mid$(OBJX$, 5, 2))
        CI2% = CVI(Mid$(OBJX$, 7, 2))
    Else
        Exit Sub
    End If
    '
    If DESA% > 0 And HASA% > 0 Then
      FILTROFECHA% = 1
    End If
    If CI1% > 0 And CI2% > 0 Then
      FILTROCODIGO% = 1
    End If
    '
    TOTCOLUMNA7 = 0
    TOTCOLUMNA8 = 0
    TOTCOLUMNA9 = 0
    TOTCOLUMNA10 = 0
    TOTCOLUMNA11 = 0
    TOTCOLUMNA12 = 0


    
    Campos$ = "FECHA/D8" '1
    Campos$ = Campos$ + ";REMITO/A18" '2
    Campos$ = Campos$ + ";TRANSF./A12" '3
    Campos$ = Campos$ + ";Cta/F7.0" '4
    Campos$ = Campos$ + ";R.SOCIAL/A24" '5
    Campos$ = Campos$ + ";CODIGO/A16" '6
    Campos$ = Campos$ + ";INGRESO/F7.0" '7
    Campos$ = Campos$ + ";METROS/F7" '8
    Campos$ = Campos$ + ";KILOS/F7" '9
    Campos$ = Campos$ + ";EGRESO/F7.0" '10
    Campos$ = Campos$ + ";METROS/F12.0" '11
    Campos$ = Campos$ + ";KILOS/F9.0" '12

    
    J& = 0: T& = 0
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)

    SQLX$ = "select FECHMOV,COD_EXTE,DESCRIP, Nume_Clie,D.Raso_Cli,DOPRILAR, SUM(METROS) METROS, SUM(PESOKG) AS PESOKG"
    SQLX$ = SQLX$ + " ,CODIMOVI,SUM(CANTINGRE)AS INGRE,SUM(CANTSALID) AS SALIDA"
    SQLX$ = SQLX$ + " from movisto WITH(NOLOCK) LEFT JOIN DEUDOR12 D ON D.Nume_Cli = Nume_Clie"
    SQLX$ = SQLX$ + "  WHERE DepoMovi = 3 AND (CODIMOVI = 'TR' OR CODIMOVI = 'RT')"
    If FILTROFECHA% > 0 Then
      Desde = FETEXCOR1$(DESA%)
      Hasta = FETEXCOR1$(HASA%)
      SQLX$ = SQLX$ + " AND FECHMOV >= '" & Desde & "'"
      SQLX$ = SQLX$ + " AND FECHMOV <= '" & Hasta & "'"
    End If
    If FILTROCODIGO% > 0 Then
      CODIX1$ = CODEXT$(CI1%)
      CODIX2$ = CODEXT$(CI2%)
      SQLX$ = SQLX$ + " AND COD_EXTE >= '" & CODIX1$ & "'"
      SQLX$ = SQLX$ + " AND COD_EXTE <= '" & CODIX2$ & "'"
    End If
    
    SQLX$ = SQLX$ + " GROUP BY FECHMOV,COD_EXTE,DESCRIP,DOPRILAR, CODIMOVI,Nume_Clie,D.Raso_Cli"
    SQLX$ = SQLX$ + " ORDER BY FECHMOV"
    '
    Screen.MousePointer = 11
    Set rs = READSET(SQLX$)
    With rs
      Do While Not .EOF
         J& = J& + 1
         GRILLAGRAL.GRID.Rows = J& + 1
         FECHAX$ = FECHATEX$(CVINT(!FechMov))
         Codigo$ = SAFETEXT$(!Cod_Exte)
         DESCRIPX$ = SAFETEXT$(!Descrip)
         NC = XVALO(!NUME_CLIE)
         RSOCIAL$ = SAFETEXT$(!raso_cli)
         DOCUMENTO$ = SAFETEXT$(!DOPRILAR)
         METROS# = XVALO(!METROS)
         KG# = XVALO(!PESOKG)
         CODIMOVI$ = SAFETEXT$(!CODIMOVI)
         INGESO# = XVALO(!INGRE)
         SALIDA# = XVALO(!SALIDA)
         GRILLAGRAL.GRID.TextMatrix(J&, 1) = FECHAX$
         If CODIMOVI$ = "TR" Then
            GRILLAGRAL.GRID.TextMatrix(J&, 3) = DOCUMENTO$
         ElseIf CODIMOVI$ = "RT" Then
            GRILLAGRAL.GRID.TextMatrix(J&, 2) = DOCUMENTO$
         End If

         GRILLAGRAL.GRID.TextMatrix(J&, 4) = NC
         GRILLAGRAL.GRID.TextMatrix(J&, 5) = RSOCIAL$
         GRILLAGRAL.GRID.TextMatrix(J&, 6) = Codigo$
         If INGESO# > 0 Then
            GRILLAGRAL.GRID.TextMatrix(J&, 7) = INGESO#
            GRILLAGRAL.GRID.TextMatrix(J&, 8) = METROS#
            GRILLAGRAL.GRID.TextMatrix(J&, 9) = KG#
            TOTCOLUMNA7 = TOTCOLUMNA7 + INGESO#
            TOTCOLUMNA8 = TOTCOLUMNA8 + METROS#
            TOTCOLUMNA9 = TOTCOLUMNA9 + KG#

         Else
            GRILLAGRAL.GRID.TextMatrix(J&, 10) = INGESO#
            GRILLAGRAL.GRID.TextMatrix(J&, 11) = METROS#
            GRILLAGRAL.GRID.TextMatrix(J&, 12) = KG#
            TOTCOLUMNA10 = TOTCOLUMNA10 + INGESO#
            TOTCOLUMNA11 = TOTCOLUMNA11 + METROS#
            TOTCOLUMNA12 = TOTCOLUMNA12 + KG#

         End If
         

        .MoveNext
      Loop
    End With
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    On Error GoTo 0
    'CALCULO LOS PROMEDIOS
    
    '
   'TOTALES
    GRILLAGRAL.GRID.Rows = J& + 3
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 2) = "Totales:"
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 7) = FORMATNUM$(XVALO(TOTCOLUMNA7), "F10.0")
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 8) = FORMATNUM$(XVALO(TOTCOLUMNA8), "F10.0")
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 9) = FORMATNUM$(XVALO(TOTCOLUMNA9), "F10.0")
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 10) = FORMATNUM$(XVALO(TOTCOLUMNA10), "F10.0")
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 11) = FORMATNUM$(XVALO(TOTCOLUMNA11), "F10.0")
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 12) = FORMATNUM$(XVALO(TOTCOLUMNA12), "F10.0")

    'COLOREO COLUMNAS
'    Call COLOREAR_GRILLASQL_CadenaDeColumnas(GRILLAGRAL.GRID, &HCDDADA, "2;3;4;5;9;10;11;12;")
    
    'COLOREO FILA TOTALES
    Call COLOREAR_GRILLA_SOLAFILA(GRILLAGRAL.GRID, &H808080, J& + 2, 0, 1)

    
   GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
   GRILLAGRAL.mnuMenu1.Visible = True
   GRILLAGRAL.GRID.ZOrder 0
   GRILLAGRAL.TXTBUSCAR = ""
   Screen.MousePointer = 1
   GRILLAGRAL.Show 1
   
   Dim MSF As msFlexGrid
   If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(MSF, Campos$, "MOVDEPLU", "MOVIMIENTOS STOCK LUZURIAGA", "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
        Exit Sub
   End If
   
   GoTo 20
End Sub



Private Sub mnuProductosTerminados_Click()
''''''

   Campos$ = "Cta/F7" '1
   Campos$ = Campos$ + ";RAZÓN SOCIAL/a32" '2
   Campos$ = Campos$ + ";N VENTA/F7" '3
   Campos$ = Campos$ + ";O FAB/A12" '4
   Campos$ = Campos$ + ";F.CIERRE/D8" '5
   Campos$ = Campos$ + ";CÓDIGO/A12" '6
   Campos$ = Campos$ + ";DESCRIPCIÓN/A48" '7
   Campos$ = Campos$ + ";STOCK/F8" '8
   Campos$ = Campos$ + ";KG/F8" '9
   
   Dim CTP As New Conteplast

   Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
   Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
   J& = 0
    SQLX$ = "SELECT P.NroClie,P.RasoClie"
    SQLX$ = SQLX$ + " ,O.NroPed"
    SQLX$ = SQLX$ + " ,M.IDENLOTE"
    SQLX$ = SQLX$ + " ,O.FechCierr"
    'SQLX$ = SQLX$ + " ,CODESPTECTE"
    SQLX$ = SQLX$ + " ,SUM(CANTINGRE - CANTSALID) AS STOCK"
    SQLX$ = SQLX$ + " ,M.Cod_Inte"
    SQLX$ = SQLX$ + " FROM MOVISTO M WITH(NOLOCK)"
    SQLX$ = SQLX$ + " LEFT JOIN ORDEFABR O ON O.IdSubOr = M.IdenLote"
    SQLX$ = SQLX$ + " LEFT JOIN PEDENCA P ON P.NroPed = O.NroPed"
    'SQLX$ = SQLX$ + " LEFT JOIN RECETACTP R ON R.CODINT =  M.Cod_Inte"
    SQLX$ = SQLX$ + " Where DepoMovi = 9"
    SQLX$ = SQLX$ + " GROUP BY M.Cod_Inte,M.IDENLOTE,O.NroPed,O.FechCierr,P.NroClie,P.RasoClie"
    'SQLX$ = SQLX$ + " ,CODESPTECTE"
    SQLX$ = SQLX$ + " Having Sum(CANTINGRE - CANTSALID) > 0"
    SQLX$ = SQLX$ + " ORDER BY NroClie,NroPed,M.Cod_Inte,M.IdenLote"

   Set rs = READSET(SQLX$)
   With rs
      Do While Not .EOF
           J& = J& + 1

           GRILLAGRAL.GRID.Rows = J& + 1
           NORFA$ = SAFETEXT(!idenlote)
           GRILLAGRAL.GRID.TextMatrix(J&, 1) = XVALO(!NROCLIE)
           GRILLAGRAL.GRID.TextMatrix(J&, 2) = SAFETEXT(!RasoClie)
           GRILLAGRAL.GRID.TextMatrix(J&, 3) = XVALO(!NROPED)
    
           GRILLAGRAL.GRID.TextMatrix(J&, 4) = NORFA$
           GRILLAGRAL.GRID.TextMatrix(J&, 5) = FECHATEX$(CVINT(!Fechcierr))
           GRILLAGRAL.GRID.TextMatrix(J&, 6) = CODEXT$(XVALO(!cod_inte))  'lo traigo asi por que cambiaron codigos
           GRILLAGRAL.GRID.TextMatrix(J&, 7) = DESCRIT0$(XVALO(!cod_inte))
           GRILLAGRAL.GRID.TextMatrix(J&, 8) = XVALO(!STOCK)
            'EL PESO ES PROMEDIO DE LOS CIERRES EN FORMA UNITARIA POR LA CANTIDAD DE STOCK
           GRILLAGRAL.GRID.TextMatrix(J&, 9) = TRIM$(FORMATNUM$(CTP.PesoPromedioDeCierresFabricSegunOF(NORFA$) * XVALO(!STOCK), "F12"))
           .MoveNext
     Loop
    End With
    rs.Close
    Set rs = Nothing
    '
    Screen.MousePointer = SCMA
    
     TX$ = "Stocks Productos Terminados Emitido el: " & FECHATEX$(HOY(HO$))
     GRILLAGRAL.Caption = TX$
     '*******************************************
     GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
     GRILLAGRAL.mnuMenu1.Visible = True
     GRILLAGRAL.GRID.ZOrder 0
     GRILLAGRAL.TXTBUSCAR = ""
     Call POSICIONAR_ULTIMAFILA(GRILLAGRAL.GRID)
     
     GRILLAGRAL.SinEfectoClickEnGrilla% = 1
     
     GRILLAGRAL.Show 1
     Dim MSF As msFlexGrid
     If GRILLAGRAL.MENU1 > 0 Then
         Set MSF = GRILLAGRAL.GRID
         If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
         Call CARGA_TABLA_IMPRE(MSF, Campos$, "STOPROTE", TX$, "TABLA_IMPRE", , 1)
         TOTMETOS& = FORMATNUM$(XVALO(SUMA_COLUMNA_SELECCION(MSF, 8)), "F10")
         TOTKILOS& = FORMATNUM$(XVALO(SUMA_COLUMNA_SELECCION(MSF, 9)), "F10")
         FILTX$ = TRIM$(FECHATEX$(HOY(HO$))) & ";" & TRIM$(TOTMETOS&) & ";" & TRIM$(TOTKILOS&)
         Call STDFORM("STOC-CTP", FILTX$)
     
         GRILLAGRAL.MENU1 = 0
         Screen.MousePointer = 1
     End If

End Sub

Private Sub mnuQuiebreDeStock_Click()
   
    Dim QuiebreEntregaItemFecha() As Single
    Dim CantidadEntregadaItemFecha() As Single
    Dim PrecioEntregadoItemFecha() As Single

    Call DESHASFECHA(Desde%, Hasta%, PERIODO$)
    If PERIODO$ = "" Then Exit Sub
    DesdeFecha$ = FETEXCOR1$(Desde%)
    HastaFecha$ = FETEXCOR1$(Hasta%)
    FECHAHOY$ = FETEXCOR1$(HOY(HO$))
    
    Screen.MousePointer = 11
    
    FIN& = 6
    Call TRACE(24, 1, FIN&)
    CantidadDias% = DiferenciaDiasEntreFechas(Desde%, Hasta%)
    
    ReDim Preserve QuiebreEntregaItemFecha(32767, CantidadDias%)
    ReDim Preserve CantidadEntregadaItemFecha(32767, CantidadDias%)
    ReDim Preserve PrecioEntregadoItemFecha(32767, CantidadDias%)
        
    ' QUIEBRE DE ENTREGA - PEDIDOS CON ENTREGAS PARCIALES O TOTALES
    Dim RstPedidosConEntrega As ADODB.Recordset
    ConsultaSql$ = "SELECT PDD.NumeroPedido, PDD.CodigoInterno, PDD.CantidadPedida, PDD.FechaEntrega, SUM(MOVISTO.CantSalid) AS CantidadSalida, SUM(MOVISTO.CANTINGRE) AS CantidadIngresada, AVG(ValoUnit) AS PrecioPromedio FROM" _
                    & " (SELECT PE.NroPed AS NumeroPedido, PD.CodiInt as CodigoInterno, PD.NuOrItem AS NumeroOrden, PD.FeSolEnt as FechaEntrega, PD.CanPed as CantidadPedida FROM PEDENCA PE WITH(NOLOCK) INNER JOIN PEDDETA PD WITH (NOLOCK) ON PE.NROPED = PD.NROPED WHERE PE.FechEmis >= '" & DesdeFecha$ & "' AND PE.FechEmis <= '" & HastaFecha$ & "' AND PE.STATUS < 9 AND PD.STATUS < 8 AND PD.FeSolEnt <' " & HastaFecha$ & "' GROUP BY PE.NroPed, PD.CodiInt, PD.NuOrItem, PD.FeSolEnt, PD.CanPed) PDD" _
                    & " LEFT JOIN MOVISTO WITH(NOLOCK) ON (MOVISTO.NuPedCli = PDD.NumeroPedido AND MOVISTO.Cod_Inte = PDD.CodigoInterno AND MOVISTO.NumeOrdPed = PDD.NumeroOrden)" _
                    & " WHERE MOVISTO.FECHMOV <= PDD.FechaEntrega AND CODIMOVI='RT' GROUP BY PDD.NumeroPedido, PDD.CodigoInterno, PDD.CantidadPedida, PDD.FechaEntrega"
    Set RstPedidosConEntrega = READSET(ConsultaSql$)
    
    With RstPedidosConEntrega
        Do While Not .EOF
            CantidadEntregada# = XVALO(!CantidadSalida) - XVALO(!CantidadIngresada)
            If CantidadEntregada# < 0 Then CantidadEntregada# = 0
            
            CantidadPedida# = !CantidadPedida
            If CantidadEntregada# < CantidadPedida# Then
                QuiebreDeEntrega# = CantidadPedida# - CantidadEntregada#
                CodigoInterno% = !CodigoInterno
                PrecioPromedio# = XVALO(!PrecioPromedio) ' SE TOMA EL PRECIO PROMEDIO DE LAS DISTINTAS ENTREGAS EN PESOS DE LA COMBINACION PEDIDO-CODIGO-ORDEN
                FechaEntrega% = CVINT(!FechaEntrega%)
                DistanciaEnDias% = DiferenciaDiasEntreFechas(Desde%, FechaEntrega%)
                                
                QuiebreEntregaItemFecha(CodigoInterno%, DistanciaEnDias%) = QuiebreEntregaItemFecha(CodigoInterno%, DistanciaEnDias%) + QuiebreDeEntrega#
                CantidadEntregadaItemFecha(CodigoInterno%, DistanciaEnDias%) = CantidadEntregadaItemFecha(CodigoInterno%, DistanciaEnDias%) + 1
                PrecioEntregadoItemFecha(CodigoInterno%, DistanciaEnDias%) = PrecioEntregadoItemFecha(CodigoInterno%, DistanciaEnDias%) + PrecioPromedio#
            End If
            
            .MoveNext
        Loop
    End With
    On Error Resume Next
    RstPedidosConEntrega.Close
    Set RstPedidosConEntrega = Nothing
    On Error GoTo 0

    Call TRACE(24, 2, FIN&)
    
    ' QUIEBRE DE ENTREGA - SIN ENTREGAS
    Dim RstPedidosSinEntrega As ADODB.Recordset
    ConsultaSql$ = "SELECT PDD.NumeroPedido, PDD.CodigoInterno, PDD.CantidadPedida, PDD.FechaEntrega, PDD.PrecioUnitario, PDD.TipoCambio, SUM(MOVISTO.CantSalid) AS CantidadSalida, SUM(MOVISTO.CANTINGRE) AS CantidadIngresada FROM" _
                    & " (SELECT PE.NroPed AS NumeroPedido, PD.CodiInt as CodigoInterno, PD.NuOrItem AS NumeroOrden, PD.FeSolEnt as FechaEntrega, PD.CanPed as CantidadPedida, AVG(PreUnid) as PrecioUnitario, AVG(TipoCam) TipoCambio FROM PEDENCA PE WITH(NOLOCK) INNER JOIN PEDDETA PD WITH(NOLOCK) ON PE.NROPED = PD.NROPED WHERE PE.FechEmis >= '" _
                    & DesdeFecha$ & "' AND PE.FechEmis <= '" & HastaFecha$ & "' AND PE.STATUS < 9 AND PD.STATUS < 8 AND PD.FeSolEnt <' " _
                    & HastaFecha$ & "' GROUP BY PE.NroPed, PD.CodiInt, PD.NuOrItem, PD.FeSolEnt, PD.CanPed) PDD LEFT JOIN MOVISTO WITH(NOLOCK) ON (MOVISTO.NuPedCli = PDD.NumeroPedido AND MOVISTO.Cod_Inte = PDD.CodigoInterno AND MOVISTO.NumeOrdPed = PDD.NumeroOrden)" _
                    & " GROUP BY PDD.NumeroPedido, PDD.CodigoInterno, PDD.CantidadPedida, PDD.FechaEntrega, PDD.PrecioUnitario, PDD.TipoCambio "
    Set RstPedidosSinEntrega = READSET(ConsultaSql$)

    With RstPedidosSinEntrega
        Do While Not .EOF
            If IsNull(!CantidadSalida) And IsNull(!CantidadIngresada) Then    ' AMBOS CAMPOS EN NULL, NOS DICE QUE SE TRATA DE UN PEDIDO SIN MATERIAL ENTREGADO
                QuiebreDeEntrega# = !CantidadPedida
                CodigoInterno% = !CodigoInterno
                PrecioPromedio# = XVALO(!PrecioUnitario) * XVALO(!TIPOCAMBIO) ' SE TOMA EL PRECIO UNITARIO PECIFICADO DEL PEDIDO
                FechaEntrega% = CVINT(!FechaEntrega%)
                DistanciaEnDias% = DiferenciaDiasEntreFechas(Desde%, FechaEntrega%)
                                
                QuiebreEntregaItemFecha(CodigoInterno%, DistanciaEnDias%) = QuiebreEntregaItemFecha(CodigoInterno%, DistanciaEnDias%) + QuiebreDeEntrega#
                CantidadEntregadaItemFecha(CodigoInterno%, DistanciaEnDias%) = CantidadEntregadaItemFecha(CodigoInterno%, DistanciaEnDias%) + 1
                PrecioEntregadoItemFecha(CodigoInterno%, DistanciaEnDias%) = PrecioEntregadoItemFecha(CodigoInterno%, DistanciaEnDias%) + PrecioPromedio#
            End If

            .MoveNext
        Loop
    End With
    On Error Resume Next
    RstPedidosSinEntrega.Close
    Set RstPedidosSinEntrega = Nothing
    On Error GoTo 0

    Call TRACE(24, 3, FIN&)
    
    ' QUIEBRE DE STOCK POR ITEM-FECHA
           
    ' PEDIDOS PENDIENTES
    ' {  (1) } {            (2)                  }
    ' --------|-----------------------------------|-----
    '      DESDE                                HASTA
    
    ' (1) PEDIDOS PENDIENTES CON FECHA EMISION < Desde% Y FECHA DE ENTREGA < Desde% (TODOS LOS ARTICULOS)
'    Dim EstadoItemPedidoAnteriorADesde() As Single
'    ReDim EstadoItemPedidoAnteriorADesde(32767) As Single
    PedidosDesde% = DIANTE(HOY(HO$), 255)
    PedidosDesde% = DIANTE(PedidosDesde%, 110)
'    Call EstadoPedidosAnterioresADesde(EstadoItemPedidoAnteriorADesde, Desde%, PedidosDesde%)
    
    Dim PedidosPendientesItemFecha() As Single
    ReDim PedidosPendientesItemFecha(32767, CantidadDias%)
    
    Dim StockItemFecha() As Single
    ReDim StockItemFecha(32767, CantidadDias%)
    
    Dim QuiebreStockItemFecha() As Single
    ReDim Preserve QuiebreStockItemFecha(32767, CantidadDias%)
    
    For i% = 1 To 32766
        Call TRACE(24, CLng(i%), 32766)
        For J% = 1 To CantidadDias%
            QuiebreStockItemFecha(i%, J%) = QuiebreEntregaItemFecha(i%, J%)
            If QuiebreStockItemFecha(i%, J%) > 0.05 Then
                FechaEntrega% = DIASPOST(Desde%, J%)
                
                StockAFecha# = SALDOAFECHA(i%, FechaEntrega% - 1, 1)
                If StockAFecha# > 0.05 Then
                    ' (2) PEDIDOS PENDIENTES DEL ARTICULO i%, CON FECHA EMISION < Desde% Y FECHA DE ENTREGA EN EL RANGO (Desde% y FechaEntrega%)
                    'PedidosPendientesEntreDesdeFechaEntrega# = EstadoPedidosDesdeAFechaEntrega(i%, Desde%, FechaEntrega%, PedidosDesde%)
                    'PedidosPendientesEntregaAnterioresDesde# = EstadoItemPedidoAnteriorADesde(i%) ' (1)
                    PedidosPendientesEntreDesdeFechaEntrega# = PedidosPendientesPorItem(i%, FechaEntrega%, PedidosDesde%)
                    PedidosPendientesItemFecha(i%, J%) = PedidosPendientesEntregaAnterioresDesde# + PedidosPendientesEntreDesdeFechaEntrega#
                    StockItemFecha(i%, J%) = StockAFecha#
                    
                    StockAFecha# = StockAFecha# - PedidosPendientesEntregaAnterioresDesde# - PedidosPendientesEntreDesdeFechaEntrega#
                End If
                If StockAFecha# < 0 Then StockAFecha# = 0
                
                If StockAFecha# > 0.05 Then
                    If QuiebreEntregaItemFecha(i%, J%) > 0.05 Then
                        If StockAFecha# > QuiebreEntregaItemFecha(i%, J%) Then
                            QuiebreStockItemFecha(i%, J%) = 0
                        Else
                            QuiebreStockItemFecha(i%, J%) = QuiebreEntregaItemFecha(i%, J%) - StockAFecha#
                        End If
                    End If
                End If
            End If
        Next J%
    Next i%
    
    ReDim EstadoItemPedidoAnteriorADesde(1)
    
    Call TRACE(24, 4, FIN&)
      
    ' QUIEBRE DE STOCK POR ITEM DEL PERIODO
    
    Dim QuiebreStockPorItem() As Single
    Dim QuiebreEntregaPorItem() As Single
    Dim QuiebreStockPrecioUnitarioPorItem() As Single
    ReDim QuiebreStockPorItem(32767)
    ReDim QuiebreEntregaPorItem(32767)
    ReDim QuiebreStockPrecioUnitarioPorItem(32767)
    For i% = 1 To 32766
        CantidadEntregada# = 0
        For J% = 1 To CantidadDias%
            QuiebreStockPorItem(i%) = QuiebreStockPorItem(i%) + QuiebreStockItemFecha(i%, J%)
            QuiebreEntregaPorItem(i%) = QuiebreEntregaPorItem(i%) + QuiebreEntregaItemFecha(i%, J%)
            QuiebreStockPrecioUnitarioPorItem(i%) = QuiebreStockPrecioUnitarioPorItem(i%) + PrecioEntregadoItemFecha(i%, J%) ' SUMA DE PRECIOS PROMEDIO POR DIA
            CantidadEntregada# = CantidadEntregada# + CantidadEntregadaItemFecha(i%, J%) ' SUMA DE CANTIDADES ENTREGADAS POR DIA
        Next J%
        If QuiebreStockPrecioUnitarioPorItem(i%) > 0.05 Then
            If CantidadEntregada# > 0.05 Then QuiebreStockPrecioUnitarioPorItem(i%) = QuiebreStockPrecioUnitarioPorItem(i%) / CantidadEntregada# ' PRECIO PROMEDIO
        End If
    Next i%


    Call TRACE(24, 5, FIN&)
20  Campos$ = "Código/A24;Descripción/A48;Quiebre Entrega/F0;Rotura Stock/F14.1;Ind. Pérdida Item/F20.1; Ind. Pérdida Total/F20.1; Valorizacion/F0; Facturacion/F0"
    
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)

    GRILLAGRAL.Caption = "LISTADO DE ROTURA DE STOCK - PERIODO: " & PERIODO$
    GRILLAGRAL.mnuMenu1.Caption = "IMPRIMIR"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.TXTBUSCAR = ""
    GRILLAGRAL.NOPINTAR = 1
    
    ' VENTA NETA: SOLO FACTURAS Y NOTAS DE CREDITO
    FacturacionNeta# = XVALO(VALOBADA("CAJABANC", "SUM(IHABECOMP-IDEBECOMP)", "TIPOMOVIM='VENTA' AND (OPCIMOVIM='FC' OR OPCIMOVIM='NC') AND NUCLIEN > 0 AND FECHCONTA >='" & DesdeFecha$ & "' AND FECHCONTA <= '" & HastaFecha$ & "'"))
    FilaGrilla% = 0
    TotalCasoConQuiebreDeStock% = 0
    AcumuladoCantidadDeQuiebre# = 0
    For i% = 1 To 32766
        If QuiebreStockPorItem(i%) > 0.05 Then
                GRILLAGRAL.GRID.Rows = GRILLAGRAL.GRID.Rows + 1
                FilaGrilla% = FilaGrilla% + 1
                GRILLAGRAL.GRID.TextMatrix(FilaGrilla%, 1) = CODEXT$(i%)
                GRILLAGRAL.GRID.TextMatrix(FilaGrilla%, 2) = DESCRIT0$(i%)
                GRILLAGRAL.GRID.TextMatrix(FilaGrilla%, 3) = FORMATNUM$(QuiebreEntregaPorItem(i%), "F10.1")
                GRILLAGRAL.GRID.TextMatrix(FilaGrilla%, 4) = FORMATNUM$(QuiebreStockPorItem(i%), "F10.1")
                    IndicePorItem# = 0
                    ValorizacionDeQuiebre# = QuiebreStockPrecioUnitarioPorItem(i%) * QuiebreStockPorItem(i%)
                    If FacturacionNeta# > 0.05 Then IndicePorItem# = ValorizacionDeQuiebre# / FacturacionNeta#
                GRILLAGRAL.GRID.TextMatrix(FilaGrilla%, 5) = FORMATNUM$(IndicePorItem# * 100, "F10.1")
                    IndiceAcumulado# = IndiceAcumulado# + IndicePorItem#
                GRILLAGRAL.GRID.TextMatrix(FilaGrilla%, 6) = FORMATNUM$(IndiceAcumulado# * 100, "F10.1")
                
                GRILLAGRAL.GRID.TextMatrix(FilaGrilla%, 7) = FORMATNUM$(ValorizacionDeQuiebre#, "F10.1")
                GRILLAGRAL.GRID.TextMatrix(FilaGrilla%, 8) = FORMATNUM$(FacturacionNeta#, "F10.1")
                
                TotalCasoConQuiebreDeStock% = TotalCasoConQuiebreDeStock% + 1
                AcumuladoCantidadDeQuiebre# = AcumuladoCantidadDeQuiebre# + QuiebreStockPorItem(i%)
        End If
    Next i%
    
    If TotalCasoConQuiebreDeStock% > 0 Then
        GRILLAGRAL.GRID.Rows = GRILLAGRAL.GRID.Rows + 2
        FilaGrilla% = FilaGrilla% + 2
        GRILLAGRAL.GRID.TextMatrix(FilaGrilla%, 2) = "Promedio General de Quiebre (Cantidad / Indice)"
        IndicePromedioGeneralQuiebre# = IndiceAcumulado# / TotalCasoConQuiebreDeStock%
        PromedioGeneralQuiebre# = AcumuladoCantidadDeQuiebre# / TotalCasoConQuiebreDeStock%
        GRILLAGRAL.GRID.TextMatrix(FilaGrilla%, 4) = FORMATNUM$(PromedioGeneralQuiebre#, "F10.1")
        GRILLAGRAL.GRID.TextMatrix(FilaGrilla%, 5) = FORMATNUM$(IndicePromedioGeneralQuiebre#, "F10.1")
    End If
    
    Call TRACE(24, 6, FIN&)
    
    ReDim CantidadEntregadaItemFecha(1, 1)
    ReDim PrecioEntregadoItemFecha(1, 1)
    ReDim QuiebreStockPorItem(1)
    ReDim QuiebreEntregaPorItem(1)
    ReDim QuiebreStockPrecioUnitarioPorItem(1)
    
    Screen.MousePointer = 1
30  GRILLAGRAL.Show 1
    Dim MSF As msFlexGrid
    
    If GRILLAGRAL.SeleccionaFilaDobleClick = 1 Then
        codigoExterno$ = TRIM(RESP_ZELE_VECT(1))
        CodigoInterno% = CODINT%(codigoExterno$)
        
        TituloLista$ = "DETALLE DE ROTURA DE STOCK PARA EL CODIGO: " & codigoExterno$
        Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, "Fecha/D8;Rotura Stock/F18.1;Quiebre Entrega/F18.1;Ped. Pend./F15.1; Stock/F15.1", FRMZELECHO)
        Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Fecha/D8;Rotura Stock/F18.1;Quiebre Entrega/F18.1;Ped. Pend./F15.1; Stock/F15.1", FRMZELECHO)
        FRMZELECHO.Caption = TituloLista$
        
        For J% = 1 To CantidadDias%
            If QuiebreEntregaItemFecha(CodigoInterno%, J%) > 0.05 Then
                FRMZELECHO.MSF2.Rows = FRMZELECHO.MSF2.Rows + 1
                FilaGrilla% = FRMZELECHO.MSF2.Rows - 1
                FRMZELECHO.MSF2.TextMatrix(FilaGrilla%, 1) = FECHATEX$(DIASPOST(Desde%, J%))
                FRMZELECHO.MSF2.TextMatrix(FilaGrilla%, 2) = TRIM$(FORMATNUM$(QuiebreStockItemFecha(CodigoInterno%, J%), "F10.1"))
                FRMZELECHO.MSF2.TextMatrix(FilaGrilla%, 3) = TRIM$(FORMATNUM$(QuiebreEntregaItemFecha(CodigoInterno%, J%), "F10.1"))
                FRMZELECHO.MSF2.TextMatrix(FilaGrilla%, 4) = TRIM$(FORMATNUM$(PedidosPendientesItemFecha(CodigoInterno%, J%), "F10.1"))
                FRMZELECHO.MSF2.TextMatrix(FilaGrilla%, 5) = TRIM$(FORMATNUM$(StockItemFecha(CodigoInterno%, J%), "F10.1"))
            End If
        Next J%
        
        FRMZELECHO.Width = 10000
        FRMZELECHO.Show 1
        
        GoTo 30
    End If
    
    ReDim QuiebreEntregaItemFecha(1, 1)
    ReDim QuiebreStockItemFecha(1, 1)
    
    If GRILLAGRAL.MENU1 > 0 Then
            Set MSF = GRILLAGRAL.GRID
            If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
            Call CARGA_TABLA_IMPRE(MSF, Campos$, "ROTUSTO", "LISTADO DE INDICADORES DE ROTURA DE STOCK - " & PERIODO$, "TABLA_IMPRE", RET_NAMECONS$, 1)
            Call FINAL("?" & RET_NAMECONS$)
            GRILLAGRAL.MENU1 = 0
    End If
     
99  Exit Sub
        

End Sub


Private Sub mnuRepfabVsVtasXFamilia_Click()

    If DERACCE%("FABVSVTA", "REPORTE DE FABRICACION VS. VENTAS") < 2 Then Exit Sub
20  Campos$ = "Interno/F9;Código/A18;Descripción/A32;Fabricado/F10.1;O.Vta.Rango/f12.1;Stock/F12.1;O.Vta.Pendiente/F15.1;Pendiente/F12.1"
    
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
    
    Call SELECHO("TAFAMIL")
    VALONUE$ = VALACT1$("TAFAMIL")
    If VALONUE$ = "" Then GoTo 99
    
    Call REPLA(VDEF$, MKI$(HOY(HO$)), 1, 2)
10  OBX$ = OBJPLA$("ENTRAFECHA1", VDEF$)
    If OBX$ = "" Then GoTo 99
    Des% = CVI(Mid$(OBX$, 1, 2))
    Has% = HOY(HO$)
    FECDESDE$ = FETEXCOR1(Des%)
    FECHASTA$ = FETEXCOR1(Has%)
   
    GRILLAGRAL.Caption = "Reporte de Artículos Clasificados (Familia: " & VALACT2$("TAFAMIL") & ")"
    GRILLAGRAL.mnuMenu1.Caption = "IMPRIMIR"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.TXTBUSCAR = ""
    
    'DECLARO TOOLTIPTEXT A VISUALIZAR SEGUN COLUMNA EN LA GRILLA DE FRM GRILLAGRAL (VectorToolTipText ES UNA VARIABLE PUBLICA)
    Erase VectorToolTipText 'EL INDICE ES LA COLUMNA DONDE SE APLICA EL TOOLTIPTEXT
    VectorToolTipText(4) = "Cantidad Fabricada En las Fechas Seleccionada"
    VectorToolTipText(5) = "Cantidades Pedidas En las Fechas Seleccionada Sin Considerar Entregas"
    VectorToolTipText(6) = "Cantidades Remitidas En las Fechas Seleccionada Facturadas o No"
    VectorToolTipText(7) = "Cantidades Pedidas - Cantidades Remitidas"
    
    '
    Screen.MousePointer = 11
    'CARGO EL RECORDSET
    Dim obj As New RECORXET
    Set RS1 = obj.RS_ArtClasificado(VALONUE$)
    Call Carga_MSF_Recordset(RS1, Campos$, GRILLAGRAL.GRID, 1, 1)
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    '*******************************************
    Screen.MousePointer = 1
    If GRILLAGRAL.GRID.Rows < 2 Then Exit Sub
    '
'    CAMPOS$ = "Codint/F0;Código/A18;Descripción/A32;Fabricado/F10.1;Ped.O.Vta./f12.1;Cant.Fact/F12.1;Q.Pendiente/F12.1"
    
    FIN& = GRILLAGRAL.GRID.Rows - 1
    For i& = 1 To GRILLAGRAL.GRID.Rows - 1
      Call TRACE(24, i&, FIN&)
      CI1% = XVALO(GRILLAGRAL.GRID.TextMatrix(i&, 1))
      If CI1% > 0 Then
        CONDI$ = "COD_INTE =  " & CI1%
        CONDI$ = CONDI$ + " AND CODIMOVI  = 'PF' "
        CONDI$ = CONDI$ + " AND FECHMOV >= '" & FECDESDE$ & "'"
        CONDI$ = CONDI$ + " AND FECHMOV <= '" & FECHASTA$ & "'"
        
        FABRICADO = XVALO(VALOBADA("MOVISTO", "SUM(CANTINGRE)", CONDI$, "NOMESS"))
        PedidosRango# = CANTIDAD_PEDIDA(CI1%, Des%, Has%)
        'CONDI$ = "CODIMOVI  = 'RT' AND COD_INTE = " & CI1% & " AND FECHMOV >= '" & FECDESDE$ & "' AND FECHMOV <= '" & FECHASTA$ & "'"
        'CANTIFACT = XVALO(VALOBADA("MOVISTO", "SUM(CANTSALID - CANTINGRE)", CONDI$, "NOMESS"))
        PedidosPendientes# = 0
        GoSub CantidadPedidaPendiente
        
        STOCKAHORA = STOCKACT(CI1%)
        PENDIENTEX = PedidosPendientes# - STOCKAHORA
        If PENDIENTEX < 0 Then PENDIENTEX = 0
        GRILLAGRAL.GRID.TextMatrix(i&, 4) = FORMATNUM$(FABRICADO, "F10.1")
        GRILLAGRAL.GRID.TextMatrix(i&, 5) = FORMATNUM$(PedidosRango#, "F10.1")
        GRILLAGRAL.GRID.TextMatrix(i&, 6) = FORMATNUM$(STOCKAHORA, "F10.1")
        GRILLAGRAL.GRID.TextMatrix(i&, 7) = FORMATNUM$(PedidosPendientes#, "F10.1")
        GRILLAGRAL.GRID.TextMatrix(i&, 8) = FORMATNUM$(PENDIENTEX, "F10.1")
      End If
    Next i&
    
    Screen.MousePointer = 1
30  GRILLAGRAL.Show 1
    Dim MSF As msFlexGrid

    If GRILLAGRAL.MENU1 > 0 Then
            Set MSF = GRILLAGRAL.GRID
            If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
            Call CARGA_TABLA_IMPRE(MSF, Campos$, "ARTCONFI", "Reporte de Artículos Clasificados (Familia: " & VALACT2$("TAFAMIL") & ")", "TABLA_IMPRE", RET_NAMECONS$, 1)
            Call FINAL("?" & RET_NAMECONS$)
            GRILLAGRAL.MENU1 = 0
     End If
     
99   Exit Sub

CantidadPedidaPendiente:

     SQLX$ = " SELECT CanPed,  CantEnt, AcuDespa, AcuPrepa FROM PEDDETA WITH(NOLOCK) "
     SQLX$ = SQLX$ + " WHERE CodiInt = " & CI1%
     SQLX$ = SQLX$ + " AND VenRep = 'V'"
     SQLX$ = SQLX$ + " AND Status < 8"   ' para que no este anulada ni cumplida
     SQLX$ = SQLX$ + " AND CantEnt < (CanPed - 0.005)"
     SQLX$ = SQLX$ + " ORDER BY NroPed ASC"
     '
     Set PEDTEMP = READSET(SQLX$)
     TOPEDID = 0
     With PEDTEMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           CAPEN = XVALO(!CanPed) - XVALO(!CANTENT)
           TOPEDID = TOPEDID + CAPEN
         .MoveNext
         Loop
       End If
     End With
     PedidosPendientes# = TOPEDID
        
     Return
    
End Sub



Private Sub mnuReporteEStadistico_Click()
'VER EL RANGO DE FECHAS HACERLO CON LO ULTIMOS
     Call CARSELMESA
     
     '
10   Screen.MousePointer = 1
     Call SELECHO("SELMESA")
     MESASEL$ = VALACT1$("SELMESA")
     If TRIM$(MESASEL$) = "" Then
        GoTo 99
     End If
     CORRESQ$ = MESANO$(MESASEL$, Des%, Has%)
     '
     Desde = FETEXCOR1$(Des%)
     Hasta = FETEXCOR1$(Has%)
     
    Campos$ = "Ord.Fab./A12;" '1
    Campos$ = Campos$ + "Código/a16;" '2
    Campos$ = Campos$ + "Descripción/A32;" '3
    Campos$ = Campos$ + "Proveedor/A24;"      '4
    Campos$ = Campos$ + "Transf.Mt./f9.2;"       '5
    Campos$ = Campos$ + "Cortes Mt./F9.2;"   '6
    Campos$ = Campos$ + "Dif.Metros/F12.2;"      '7
    Campos$ = Campos$ + "Kg.Transf/F9.2;"       '8
    Campos$ = Campos$ + "Kg.Corte/F9.2;"        '9
    Campos$ = Campos$ + "%Rend.Kgs./F9.2;"     '10
    Campos$ = Campos$ + "%Rend.Mts/F10.2;" '11
    Campos$ = Campos$ + "CI/A0" '12
    
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
       
    Erase VectorToolTipText 'EL INDICE ES LA COLUMNA DONDE SE APLICA EL TOOLTIPTEXT
    VectorToolTipText(4) = "Proveedor Habitual"
    VectorToolTipText(5) = "Metros Transferidos a la orden de fabricación "
    VectorToolTipText(6) = "Suma de Metros según los cortes informados (largo corte * cantidad de cortes)"
    VectorToolTipText(7) = "Diferencia en mts entre la columna Transf.MT. y Cortes Mt."
    VectorToolTipText(8) = "Kgs. Transferidos a la orden de fabricación"
    VectorToolTipText(9) = "Kgs. Informados cuando se realizan los cortes"
    VectorToolTipText(10) = "Porcentaje de rendimiento calculados según Kg.Transf. y Kg.Corte (100% valor óptimo)"
    VectorToolTipText(11) = "Porcentaje de rendimiento calculados según Mt.Transf. y Mt.Corte (100% valor óptimo)"
    SQLX$ = " SELECT M.CODINT, C.CODRANGO ,C.Descripcion ,C.CODMATPRIMA ,C.DESCRIMATPRIMA,M.COSUNI,C.CIMATPRIMA ,M.FAMILIA"
    SQLX$ = SQLX$ + " FROM CODIGORANGO C WITH(NOLOCK) INNER JOIN MASTER M ON C.CIMATPRIMA = M.CODINT"
    SQLX$ = SQLX$ + " WHERE C.CODIGORANGO_ID > 0 AND M.FAMILIA ='TE' Order by CODRANGO"
    Screen.MousePointer = 11
    Set rs = READSET(SQLX$)
    With rs
       Do While Not .EOF
            CI1% = XVALO(!CODINT)
            SQLX$ = " SELECT m.cod_inte,m.cod_exte,m.descrip,DOSECLAR,"
            SQLX$ = SQLX$ + " SUM( m.cantingre - m.cantsalid ) AS Material_Transferido"
            SQLX$ = SQLX$ + " ,SUM(CASE WHEN m.CODIMOVI = 'TR' AND Cantingre > 0 THEN pesokg ELSE -pesokg END) AS SumaPesokg"
            SQLX$ = SQLX$ + " FROM movisto m WITH(NOLOCK)  WHERE depomovi = 2   AND CODIMOVI = 'TR' and Cod_Inte = " & CI1%
            SQLX$ = SQLX$ + " AND FECHMOV >=  '" & Desde & "'"
            'SQLX$ = SQLX$ + " AND FECHMOV <=  '" & Hasta & "'"
            SQLX$ = SQLX$ + " Group By DOSECLAR,m.cod_inte, m.cod_exte,m.descrip,DOSECLAR"
            'SQLX$ = SQLX$ + " HAVING SUM(CASE WHEN m.CODIMOVI = 'TR' THEN m.cantingre - m.cantsalid ELSE 0 END) > 0"
            Set RS1 = READSET(SQLX$)
            With RS1
                Do While Not .EOF
                    NORFA$ = SAFETEXT$(!DOSECLAR)
                    If STATORFA%(NORFA$) <> 3 Then GoTo 60 'SOLO ORDENES DE FABRICACION CERRADAS
                    CI1% = XVALO(!cod_inte)
                    Codigo$ = SAFETEXT$(!Cod_Exte)
                    DESXRI$ = SAFETEXT$(!Descrip)
                    LOTE$ = ""
                    CantTransferidaAlDepo2 = XVALO(!Material_Transferido)
                    NOPER% = 0 'XVALO(!ID_OPERACION)
                    CONDI$ = "codielem = " & CI1% & " AND  FeHorInf >= '" & Desde & "' AND CANTREAL > 0 AND IDSUBOR = '" & NORFA$ & "'"
'                    Descrioper$ = VALOBADA("INFOFAB", "DESCOPER", CONDI$, "NOMESS")
                    CantiMetros = XVALO(VALOBADA("INFOFAB", "SUM(CORTE*Cantreal)", CONDI$, "NOMESS"))
                    PesoInfofab = XVALO(VALOBADA("INFOFAB", "SUM(PESO)", CONDI$, "NOMESS"))
                    SumaPesokg = XVALO(!SumaPesokg) 'del movisto
                    J& = J& + 1
                    GRILLAGRAL.GRID.Rows = J& + 1
                    GRILLAGRAL.GRID.TextMatrix(J&, 1) = NORFA$
                    GRILLAGRAL.GRID.TextMatrix(J&, 2) = Codigo$
                    GRILLAGRAL.GRID.TextMatrix(J&, 3) = DESXRI$
                      NP = PROVHABI(CI1%)
                    GRILLAGRAL.GRID.TextMatrix(J&, 4) = rasocli$(-1 * NP)
                    GRILLAGRAL.GRID.TextMatrix(J&, 5) = FORMATNUM$(CantTransferidaAlDepo2, "F12.2")
                    GRILLAGRAL.GRID.TextMatrix(J&, 6) = FORMATNUM$(CantiMetros, "F12.2")
                    GRILLAGRAL.GRID.TextMatrix(J&, 7) = FORMATNUM$(XVALO(GRILLAGRAL.GRID.TextMatrix(J&, 5)) - XVALO(GRILLAGRAL.GRID.TextMatrix(J&, 6)), "F12.2")
                    
                    GRILLAGRAL.GRID.TextMatrix(J&, 8) = FORMATNUM$(SumaPesokg, "F12.2")
                    GRILLAGRAL.GRID.TextMatrix(J&, 9) = FORMATNUM$(PesoInfofab, "F12.2")
                    If XVALO(GRILLAGRAL.GRID.TextMatrix(J&, 8)) <= 0 Then GRILLAGRAL.GRID.TextMatrix(J&, 8) = 1
                    GRILLAGRAL.GRID.TextMatrix(J&, 10) = FORMATNUM$((XVALO(GRILLAGRAL.GRID.TextMatrix(J&, 9)) / XVALO(GRILLAGRAL.GRID.TextMatrix(J&, 8)) * 100), "F12.2")
                    GRILLAGRAL.GRID.TextMatrix(J&, 11) = FORMATNUM$((XVALO(GRILLAGRAL.GRID.TextMatrix(J&, 6)) / XVALO(GRILLAGRAL.GRID.TextMatrix(J&, 5))) * 100, "F12.2") ' (Metros informados como cortes / Metros Ingresados) * 100
                    GRILLAGRAL.GRID.TextMatrix(J&, 12) = CI1%

60                 .MoveNext
                Loop
            End With
            RS1.Close
            Set RS1 = Nothing
            
            .MoveNext
       Loop
    End With
    rs.Close
    Set rs = Nothing
   

   
   GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
   GRILLAGRAL.mnuMenu1.Visible = True
   GRILLAGRAL.GRID.ZOrder 0
   GRILLAGRAL.TXTBUSCAR = ""
   TX$ = "Reporte Estadístico (Material Transferido vs Informado en Cortes (Solo Considera O.Fab. Cerradas) A Partir de:  " & FECHATEX$(Des%)
   
   GRILLAGRAL.Caption = TX$
   Screen.MousePointer = 1
   GRILLAGRAL.Show 1
   
   Dim MSF As msFlexGrid
   If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
'        Call CARGA_TABLA_IMPRE(GRILLAGRAL.MSF_2, Campos$, "STOCKRGO", "STOCKS DE CÓDIGOS RANGO", "TABLA_IMPRE", RET_NAMECONS$, 1)

        Call CARGA_TABLA_IMPRE(MSF, Campos$, "REPESTAD", "REPORTE ESTADISTICO", "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
        Exit Sub
   End If
99 Exit Sub
End Sub

Private Sub mnuReportePlantaLaminado_Click()
20  FILTROFECHA% = 0
    FILTROCODIGO% = 0
    OBJX$ = OBJPLA$("CODIFECH", VDEF$)
    If OBJX$ <> "" Then
        DESA% = CVI(Mid$(OBJX$, 1, 2))
        HASA% = CVI(Mid$(OBJX$, 3, 2))
        CI1% = CVI(Mid$(OBJX$, 5, 2))
        CI2% = CVI(Mid$(OBJX$, 7, 2))
    Else
        Exit Sub
    End If
    '
    If DESA% > 0 And HASA% > 0 Then
      FILTROFECHA% = 1
    End If
    If CI1% > 0 And CI2% > 0 Then
      FILTROCODIGO% = 1
    End If
    '
    TOTCOLUMNA2 = 0
    TOTCOLUMNA5 = 0
    TOTCOLUMNA8 = 0
    TOTCOLUMNA11 = 0
    TOTCOLUMNA12 = 0
    PROMEDIOCOLUMNA13 = 0
    PROMEDIOCOLUMNA14 = 0
    PROMEDIOCOLUMNA15 = 0


    
    Campos$ = "FECHA/D8" '1
    Campos$ = Campos$ + ";METROS/F8.0" '2
    Campos$ = Campos$ + ";TELA LAMINADA/A16" '3
    Campos$ = Campos$ + ";LOTE LAM./A12" '4
    Campos$ = Campos$ + ";PESO(kg)/F8.0" '5
    Campos$ = Campos$ + ";TELA BASE/A16" '6
    Campos$ = Campos$ + ";LOTE BASE/A12" '7
    Campos$ = Campos$ + ";Kg BASE/F7.0" '8
    Campos$ = Campos$ + ";Ancho(cm)/A0" '9
    Campos$ = Campos$ + ";Gramaje/A0" '10
    Campos$ = Campos$ + ";APORTE/F7.0" '11
    Campos$ = Campos$ + ";AP sin REF/F12.0" '12
    Campos$ = Campos$ + ";GRAM BASE/F9.0" '13
    Campos$ = Campos$ + ";GRAM LAM/F8.0" '14
    Campos$ = Campos$ + ";GRAM APORTE/F12.0" '15
    
    J& = 0: T& = 0
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)

    SQLX$ = "SELECT M.COD_INTE,M.COD_EXTE,M.FECHMOV,M.IDENLOTE,M.PESOKG,M.METROS,M.DOPRILAR"
    SQLX$ = SQLX$ + " FROM MOVISTO M WITH(NOLOCK) INNER JOIN MASTER MAS "
    SQLX$ = SQLX$ + " ON M.COD_INTE = MAS.CODINT"
    SQLX$ = SQLX$ + " WHERE (MAS.FAMILIA = 'TL' OR MAS.FAMILIA = 'PL') "
    SQLX$ = SQLX$ + " AND CODIMOVI = 'OF' AND CANTINGRE > 0"
    If FILTROFECHA% > 0 Then
      Desde = FETEXCOR1$(DESA%)
      Hasta = FETEXCOR1$(HASA%)
      SQLX$ = SQLX$ + " AND M.FECHMOV >= '" & Desde & "'"
      SQLX$ = SQLX$ + " AND M.FECHMOV <= '" & Hasta & "'"
    End If
    If FILTROCODIGO% > 0 Then
      CODIX1$ = CODEXT$(CI1%)
      CODIX2$ = CODEXT$(CI2%)
      SQLX$ = SQLX$ + " AND M.COD_EXTE >= '" & CODIX1$ & "'"
      SQLX$ = SQLX$ + " AND M.COD_EXTE <= '" & CODIX2$ & "'"
    End If
    
    SQLX$ = SQLX$ + " ORDER BY FECHMOV,COD_EXTE,IDENLOTE"
    '
    Screen.MousePointer = 11
    Set rs = READSET(SQLX$)
    With rs
      Do While Not .EOF
         J& = J& + 1
         CI1LAMINADA% = XVALO(!cod_inte)
         LOTELAMINADO$ = SAFETEXT$(!idenlote)
         NORFALAMINADA$ = SAFETEXT$(!DOPRILAR)
         PESOBOBLAMINADA# = XVALO(!PESOKG)
         METROSBOBLAMINADA& = XVALO(!METROS): If METROSBOBLAMINADA& <= 0 Then METROSBOBLAMINADA& = 1
         FECHAX$ = FECHATEX$(CVINT(!FechMov))
         CODIGOLAMINADO$ = SAFETEXT$(!Cod_Exte)
         
         GRILLAGRAL.GRID.Rows = J& + 1
         GRILLAGRAL.GRID.TextMatrix(J&, 1) = FECHAX$
         GRILLAGRAL.GRID.TextMatrix(J&, 2) = FORMATNUM$(METROSBOBLAMINADA&, "F10.0")
         GRILLAGRAL.GRID.TextMatrix(J&, 3) = CODIGOLAMINADO$
         GRILLAGRAL.GRID.TextMatrix(J&, 4) = LOTELAMINADO$
         GRILLAGRAL.GRID.TextMatrix(J&, 5) = FORMATNUM$(PESOBOBLAMINADA#, "F10.0")
'OJO QUE TOMA VALORES DE LA TABLA DATSQL
'******    COMPLETAR DATOS DE BOBINA LAMINDA TIPOTELACMB Y SI EL ITEM NO TIENE LOS DATOS EN LA FICHA TECNICA DIVIDE X 0   ****
   ' ***** CAPTURAR CUAL ES EL CODIGO DONDE SE GENERA LA DIVISION X 0 Y VERIFICAR LA F.TECNICA

         'DATOS BOBINA BASE
         sqlx2$ = "SELECT MOV.PESOKG,MOV.METROS,MOV.COD_INTE,MOV.COD_EXTE,MOV.DESCRIP,MOV.IDENLOTE "
         sqlx2$ = sqlx2$ + " ,FTE.TramaCob_GRAMAJE,FTE.ANCHO_CM,FTE.TramaCob_TIPOTELA"
         sqlx2$ = sqlx2$ + " from MOVISTO MOV WITH(NOLOCK) "
         sqlx2$ = sqlx2$ + "  INNER JOIN FTECNICACAMB FTE ON FTE.CODINT = MOV.COD_INTE "
         sqlx2$ = sqlx2$ + " WHERE DOCUORIG = '" & LOTELAMINADO$ & "' AND CODIMOVI = 'CF'"
         sqlx2$ = sqlx2$ + " AND DoSecCor LIKE 'PF-%' AND IDENLOTE <> '' ORDER BY FEREMOVI DESC" 'PARA QUE TRAIGA LA ULTIMA
         Set RS2 = READSET(sqlx2$)
         If RS2.EOF <> True Then
           T& = T& + 1
           CODIGOBASE$ = SAFETEXT$(RS2!Cod_Exte)
           LOTEBASE$ = TRIM$(SAFETEXT$(RS2!idenlote))
           PESOBASE# = XVALO(RS2!PESOKG)
           If Left$(LOTEBASE$, 3) = "BR-" Then
            CONDI$ = "IDENLOTE = '" & LOTEBASE$ & "'"
            PESOBASE# = XVALO(VALOBADA("BOLRECEP", "PESO", CONDI$, "NOMESS"))
           End If
           ANCHOBASE# = XVALO(RS2!ANCHO_CM)
           GRAMAJEBASE# = XVALO(RS2!TramaCob_GRAMAJE)
           TIPOTELA$ = TRIM$(UCase$(SAFETEXT$(RS2!TramaCob_TIPOTELA)))
           CONDI$ = "CAMPO2 = '" & TIPOTELA$ & "' AND LISTA = 'TIPOTELACMB'"
           REFILEVALOR = XVALO(VALOBADA("DATASQL", "CAMPO4", CONDI$, "NOMESS"))
           REFILEFACTOR = XVALO(VALOBADA("DATASQL", "CAMPO5", CONDI$, "NOMESS")): If REFILEFACTOR = 0 Then REFILEFACTOR = 1
           '
           KgLaminados = PESOBOBLAMINADA# - PESOBASE#
           KgLaminadosSinRefiles = KgLaminados - (REFILEVALOR / 1000 * METROSBOBLAMINADA&)
           GramajeTelaBase = PESOBASE# / (METROSBOBLAMINADA& * ANCHOBASE#) * 100000 / REFILEFACTOR
'excel     =(5-11+12)/(2*9)*100000/1(factor)
'           GramajeTelaLaminada = (PESOBOBLAMINADA# - KgLaminados + ANCHOBASE#) / (METROSBOBLAMINADA&) * 100000 / REFILEFACTOR
           GramajeTelaLaminada = (PESOBOBLAMINADA# - KgLaminados + KgLaminadosSinRefiles) / (METROSBOBLAMINADA& * ANCHOBASE#) * 100000 / REFILEFACTOR
           LaminadosSinRefiles = GramajeTelaLaminada - GramajeTelaBase
           
           GRILLAGRAL.GRID.TextMatrix(J&, 6) = CODIGOBASE$
           GRILLAGRAL.GRID.TextMatrix(J&, 7) = LOTEBASE$
           GRILLAGRAL.GRID.TextMatrix(J&, 8) = FORMATNUM$(PESOBASE#, "F10.0")
           GRILLAGRAL.GRID.TextMatrix(J&, 9) = FORMATNUM$(ANCHOBASE#, "F10.0")
           GRILLAGRAL.GRID.TextMatrix(J&, 10) = FORMATNUM$(GRAMAJEBASE#, "F10.0")
           'DATOS POR FORMULAS
           GRILLAGRAL.GRID.TextMatrix(J&, 11) = FORMATNUM$(KgLaminados, "F10.0")
           GRILLAGRAL.GRID.TextMatrix(J&, 12) = FORMATNUM$(KgLaminadosSinRefiles, "F10.0")
           GRILLAGRAL.GRID.TextMatrix(J&, 13) = FORMATNUM$(GramajeTelaBase, "F10.0")
           GRILLAGRAL.GRID.TextMatrix(J&, 14) = FORMATNUM$(GramajeTelaLaminada, "F10.0")
           GRILLAGRAL.GRID.TextMatrix(J&, 15) = FORMATNUM$(LaminadosSinRefiles, "F10.0")
           
           TOTCOLUMNA2 = TOTCOLUMNA2 + METROSBOBLAMINADA&
           TOTCOLUMNA5 = TOTCOLUMNA5 + PESOBOBLAMINADA#
           TOTCOLUMNA8 = TOTCOLUMNA8 + PESOBASE#
           TOTCOLUMNA11 = TOTCOLUMNA11 + KgLaminados
           TOTCOLUMNA12 = TOTCOLUMNA12 + KgLaminadosSinRefiles
           TOTCOLUMNA13 = TOTCOLUMNA13 + GramajeTelaBase
           TOTCOLUMNA14 = TOTCOLUMNA14 + GramajeTelaLaminada
           TOTCOLUMNA15 = TOTCOLUMNA15 + LaminadosSinRefiles
           
           
           RS2.Close
           Set RS2 = Nothing
         End If
        .MoveNext
      Loop
    End With
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    On Error GoTo 0
    'CALCULO LOS PROMEDIOS
    
    If T& < 1 Then T& = 1
    PROMEDIOCOLUMNA13 = TOTCOLUMNA13 / T& 'GramajeTelaBase
    PROMEDIOCOLUMNA14 = TOTCOLUMNA14 / T& 'GramajeTelaLaminada
    PROMEDIOCOLUMNA15 = TOTCOLUMNA15 / T& 'LaminadosSinRefiles
    '
   'TOTALES
    GRILLAGRAL.GRID.Rows = J& + 3
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 1) = "Tot:"
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 2) = FORMATNUM$(XVALO(TOTCOLUMNA2), "F10.0")
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 5) = FORMATNUM$(XVALO(TOTCOLUMNA5), "F10.0")
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 8) = FORMATNUM$(XVALO(TOTCOLUMNA8), "F10.0")
    
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 11) = FORMATNUM$(XVALO(TOTCOLUMNA11), "F10.0")
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 12) = FORMATNUM$(XVALO(TOTCOLUMNA12), "F10.0")
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 13) = FORMATNUM$(XVALO(PROMEDIOCOLUMNA13), "F10.0")
    
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 14) = FORMATNUM$(XVALO(PROMEDIOCOLUMNA14), "F10.0")
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 15) = FORMATNUM$(XVALO(PROMEDIOCOLUMNA15), "F10.0")
    'COLOREO COLUMNAS
    Call COLOREAR_GRILLASQL_CadenaDeColumnas(GRILLAGRAL.GRID, &HCDDADA, "2;3;4;5;9;10;11;12;")
    
    'COLOREO FILA TOTALES
    Call COLOREAR_GRILLA_SOLAFILA(GRILLAGRAL.GRID, &H808080, J& + 2, 0, 1)

    
   GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
   GRILLAGRAL.mnuMenu1.Visible = True
   GRILLAGRAL.GRID.ZOrder 0
   GRILLAGRAL.TXTBUSCAR = ""
   Screen.MousePointer = 1
   GRILLAGRAL.Show 1
   
   Dim MSF As msFlexGrid
   If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(MSF, Campos$, "REPOLAMI", "SREPORTE PLANTA DE LAMINADO", "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
        Exit Sub
   End If
   
   GoTo 20
   

End Sub

Private Sub mnuScrapMatPrim_Click()
   'Realizado para Cambaceres
    VDEF$ = ""
20  OBX$ = OBJPLA$("FECHASYFAMI", VDEF$)
    If OBX$ = "" Then
        Des% = 0: Has% = 0: PERIO$ = ""
        Exit Sub
    End If
    VDEF$ = OBX$
    '
    Des% = CVI(Left$(OBX$, 2))
    Has% = CVI(Mid$(OBX$, 3, 2))
    If Des% > Has% Then
        GoTo 20
    End If
    FEX = Des%: PERIO$ = FECHATEX$(FEX)
    FEX = Has%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    
    TM% = Asc(Mid$(OBX$, 5, 1))
    FMLIA$ = TRIM$(Mid$(OBX$, 6, 2))

    If TM% > 0 Then
     TIMAT$ = TRIM$(ECOARCH$("TATIMAT", Chr$(TM%)))
     If TIMAT$ = "" Then
        Call COMUNI("Tipo de Material No Valido")
        GoTo 20
     End If
    End If
    If FMLIA$ <> "" Then
     FAMI$ = UCase$(TRIM$(ECOARCH$("TAFAMIL", FMLIA$)))
     If FAMI$ = "" Then
        Call COMUNI("Familia No Valida")
        GoTo 20
     End If
    End If

    DESFE = FETEXCOR1$(Des%)
    HASFE = FETEXCOR1$(Has%, "HASTA")
    '
    Campos$ = "Código/A16;Descripción/A36;Stock In/F10;Ingresos/F10;Egresos/F10;Stock Final/F12"
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
    J& = 0
    TOTINICIAL = 0
    TOTINGRESO = 0
    TOTEGRESO = 0
    TOTFINAL = 0

    SQLX$ = "select CODINT,CODIGO,Descripcion from master WITH(NOLOCK)"
    SQLX$ = SQLX$ + " where CODINT > 0 AND STATUS >=0"
    If TM% > 0 Then SQLX$ = SQLX$ + " AND TIMATE = " & TM%
    If FMLIA$ <> "" Then SQLX$ = SQLX$ + " AND FAMILIA = '" & FMLIA$ & "'"
    
    SQLX$ = SQLX$ + " ORDER BY CODIGO"
    Set rs = READSET(SQLX$)
    
    With rs
    Do While Not .EOF
        CI1% = XVALO(!CODINT)
        Codigo$ = SAFETEXT$(!Codigo)
        DESXRI$ = SAFETEXT$(!Descripcion)
        
        CONDI$ = "COD_INTE = " & CI1% & " AND FECHMOV > '" & DESFE & "'" 'SOLO > Y < POR QUE SI ES
        CONDI$ = CONDI$ + " AND FECHMOV < '" & HASFE & "'"                '>= SE SUMA AL STOCK ING A FECHA
        J& = J& + 1
        STINICIAL = STODEPFECH(CI1%, 0, Des%)
        INGRESOS = XVALO(VALOBADA("MOVISTO", "SUM(CANTINGRE)", CONDI$, "NOMESS"))
        EGRESOS = XVALO(VALOBADA("MOVISTO", "SUM(CANTSALID)", CONDI$, "NOMESS"))
        STFINAL = STODEPFECH(CI1%, 0, Has%)
        
        GRILLAGRAL.GRID.Rows = J& + 1
        GRILLAGRAL.GRID.TextMatrix(J&, 1) = Codigo$
        GRILLAGRAL.GRID.TextMatrix(J&, 2) = DESXRI$
        GRILLAGRAL.GRID.TextMatrix(J&, 3) = FORMATNUM$(STINICIAL, "F12")
        GRILLAGRAL.GRID.TextMatrix(J&, 4) = FORMATNUM$(INGRESOS, "F12")
        GRILLAGRAL.GRID.TextMatrix(J&, 5) = FORMATNUM$(EGRESOS, "F12")
        GRILLAGRAL.GRID.TextMatrix(J&, 6) = FORMATNUM$(STFINAL, "F12")
        
        TOTINICIAL = TOTINICIAL + XVALO(STINICIAL)
        TOTINGRESO = TOTINGRESO + XVALO(INGRESOS)
        TOTEGRESO = TOTEGRESO + XVALO(EGRESOS)
        TOTFINAL = TOTFINAL + XVALO(STFINAL)
       .MoveNext
    Loop
   End With
   rs.Close
   Set rs = Nothing
   'TOTALES
    GRILLAGRAL.GRID.Rows = J& + 3
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 2) = "TOTALES:"
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 3) = FORMATNUM$(XVALO(TOTINICIAL), "F12")
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 4) = FORMATNUM$(XVALO(TOTINGRESO), "F12")
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 5) = FORMATNUM$(XVALO(TOTEGRESO), "F12")
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 6) = FORMATNUM$(XVALO(TOTFINAL), "F12")

   
   GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
   GRILLAGRAL.mnuMenu1.Visible = True
   GRILLAGRAL.GRID.ZOrder 0
   GRILLAGRAL.TXTBUSCAR = ""
   TX$ = ""
   If TM% > 0 Then
     TX$ = TX$ + " Material: " & TIMAT$
   End If
   If FMLIA$ <> "" Then
     TX$ = TX$ + " Familia: " & FAMI$
   End If
   
   GRILLAGRAL.Caption = "STOCK MATERIA PRIMA (" & PERIO$ & ") " & TX$
   
   GRILLAGRAL.Show 1
   
   Dim MSF As msFlexGrid
   If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(GRILLAGRAL.MSF_2, Campos$, "STOCKRGO", "STOCKS DE CÓDIGOS RANGO", "TABLA_IMPRE", RET_NAMECONS$, 1)

        Call CARGA_TABLA_IMPRE(MSF, Campos$, "STMATPR", "STOCK MATERIA PRIMA", "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
        Exit Sub
   End If
   GoTo 20
End Sub



Private Sub mnuStFeMatFam_Click()
   'Realizado para Cambaceres
    VDEF$ = ""
20  OBX$ = OBJPLA$("FECHASYFAMI", VDEF$)
    If OBX$ = "" Then
        Des% = 0: Has% = 0: PERIO$ = ""
        Exit Sub
    End If
    VDEF$ = OBX$
    '
    Des% = CVI(Left$(OBX$, 2))
    Has% = CVI(Mid$(OBX$, 3, 2))
    If Des% > Has% Then
        GoTo 20
    End If
    FEX = Des%: PERIO$ = FECHATEX$(FEX)
    FEX = Has%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    
    TM% = Asc(Mid$(OBX$, 5, 1))
    FMLIA$ = TRIM$(Mid$(OBX$, 6, 2))

    If TM% > 0 Then
     TIMAT$ = TRIM$(ECOARCH$("TATIMAT", Chr$(TM%)))
     If TIMAT$ = "" Then
        Call COMUNI("Tipo de Material No Valido")
        GoTo 20
     End If
    End If
    If FMLIA$ <> "" Then
     FAMI$ = UCase$(TRIM$(ECOARCH$("TAFAMIL", FMLIA$)))
     If FAMI$ = "" Then
        Call COMUNI("Familia No Valida")
        GoTo 20
     End If
    End If

    DESFE = FETEXCOR1$(Des%)
    HASFE = FETEXCOR1$(Has%, "HASTA")
    '
    Campos$ = "Código/A16" '1
    Campos$ = Campos$ + ";Descripción/A32" '2
    Campos$ = Campos$ + ";INI/F6.0" '3
    Campos$ = Campos$ + ";MTS INI/F7.0" '4
    Campos$ = Campos$ + ";KG INI/F7.0" '5
    '
    Campos$ = Campos$ + ";PROD/F6.0" '6
    Campos$ = Campos$ + ";MTS PR/F7.0" '7
    Campos$ = Campos$ + ";KG PR/F7.0" '8
    
    Campos$ = Campos$ + ";A LAM/F6.0" '9
    Campos$ = Campos$ + ";MTS LAM/F7.0" '10
    Campos$ = Campos$ + ";KG LAM/F7.0" '11

    Campos$ = Campos$ + ";DESP./F6.0" '12
    Campos$ = Campos$ + ";MTS DES/F7.0" '13
    Campos$ = Campos$ + ";KG DES/F7.0" '14
    
    Campos$ = Campos$ + ";FINAL/F6.0" '15
    Campos$ = Campos$ + ";MTS FIN/F7.0" '16
    Campos$ = Campos$ + ";KG FIN/F7.0" '17

    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
    J& = 0
    TOTINICIAL = 0
    TOTINGRESO = 0
    TOTEGRESO = 0
    TOTFINAL = 0
    Screen.MousePointer = 11
    
    SQLX$ = "select CODINT,CODIGO,Descripcion from master WITH(NOLOCK)"
    SQLX$ = SQLX$ + " where CODINT > 0 AND STATUS >=0"
    If TM% > 0 Then SQLX$ = SQLX$ + " AND TIMATE = " & TM%
    If FMLIA$ <> "" Then SQLX$ = SQLX$ + " AND FAMILIA = '" & FMLIA$ & "'"
    
    SQLX$ = SQLX$ + " ORDER BY CODIGO"
    Set rs = READSET(SQLX$)
    
    With rs
    Do While Not .EOF
        CI1% = XVALO(!CODINT)
        Codigo$ = SAFETEXT$(!Codigo)
        DESXRI$ = SAFETEXT$(!Descripcion)
        
        CONDI$ = "COD_INTE = " & CI1% & " AND FECHMOV >= '" & DESFE & "'" 'SOLO > Y < POR QUE SI ES
        CONDI$ = CONDI$ + " AND FECHMOV <= '" & HASFE & "'"                '>= SE SUMA AL STOCK ING A FECHA
        J& = J& + 1
        Des1Ante% = DIANTE(Des%)
        STINICIAL = STODEPFECH(CI1%, 0, Des1Ante%)
        STINICIALKG = STODEPFECHENKG#(CI1%, 0, Des1Ante%)
        STINIMETROS = STODEPFECHENMETROS#(CI1%, 0, Des1Ante%)
        If STINICIALKG > 0 Then
          X = 1
        End If
        
'        Ingresos = XVALO(VALOBADA("MOVISTO", "SUM(CANTINGRE)", CONDI$, "NOMESS"))
'        EGRESOS = XVALO(VALOBADA("MOVISTO", "SUM(CANTSALID)", CONDI$, "NOMESS"))
        '
        'LO FABRICADO EN EL PERIODO (TODO BOBINAS BASE )
        CONDI2$ = CONDI$ + " AND CODIMOVI = 'OF' AND CANTINGRE > 0 "
        PRODUCIDOUN = XVALO(VALOBADA("MOVISTO", "SUM(CANTINGRE)", CONDI2$, "NOMESS"))
        PRODUCIDOKG = XVALO(VALOBADA("MOVISTO", "SUM(PESOKG)", CONDI2$, "NOMESS"))
        PRODUCIDOMETROS = XVALO(VALOBADA("MOVISTO", "SUM(METROS)", CONDI2$, "NOMESS"))
        '
        'LO CONSUMIDO EN EL PERIODO (LAS BOBINAS BASE CONSUMIDAS PARA LAMINAR)
        CONDI3$ = CONDI$ + " AND CODIMOVI = 'CF' AND CANTSALID > 0"
        CONSUMIDOUN = XVALO(VALOBADA("MOVISTO", "SUM(CANTSALID)", CONDI3$, "NOMESS"))
        CONSUMIDOKG = XVALO(VALOBADA("MOVISTO", "SUM(PESOKG)", CONDI3$, "NOMESS"))
        CONSUMIDOMETROS = XVALO(VALOBADA("MOVISTO", "SUM(METROS)", CONDI3$, "NOMESS"))
        
         'LO REMITIDO EN EL PERIODO ()
        CONDI4$ = CONDI$ + " AND CODIMOVI = 'RT'  "
        REMITIDOUN = XVALO(VALOBADA("MOVISTO", "SUM(CANTSALID - CANTINGRE) ", CONDI4$, "NOMESS"))
        REMITIDOKG = XVALO(VALOBADA("MOVISTO", "SUM(PESOKG)", CONDI4$, "NOMESS"))
        REMITIDOMETROS = XVALO(VALOBADA("MOVISTO", "SUM(METROS)", CONDI4$, "NOMESS"))
        
        HAS1MAS% = DIASPOST(Has%, 1) 'LE AGREGO 1 DIA POR QUE LA CONDI ES < DE ESTE MODO QUEDARIA <= (ASI NO CAMBIO LA CONDI)
        STFINALUN = STODEPFECH(CI1%, 0, HAS1MAS%)
        STFINALKG = STODEPFECHENKG#(CI1%, 0, HAS1MAS%)
        STFINALMETROS = STODEPFECHENMETROS#(CI1%, 0, HAS1MAS%)
        
        GRILLAGRAL.GRID.Rows = J& + 1
        GRILLAGRAL.GRID.TextMatrix(J&, 1) = Codigo$
        GRILLAGRAL.GRID.TextMatrix(J&, 2) = DESXRI$
        GRILLAGRAL.GRID.TextMatrix(J&, 3) = FORMATNUM$(STINICIAL, "F10.0")
        GRILLAGRAL.GRID.TextMatrix(J&, 4) = FORMATNUM$(STINIMETROS, "F10.0")
        GRILLAGRAL.GRID.TextMatrix(J&, 5) = FORMATNUM$(STINICIALKG, "F10.0")
        
        GRILLAGRAL.GRID.TextMatrix(J&, 6) = FORMATNUM$(PRODUCIDOUN, "F10.0")
        GRILLAGRAL.GRID.TextMatrix(J&, 7) = FORMATNUM$(PRODUCIDOMETROS, "F10.0")
        GRILLAGRAL.GRID.TextMatrix(J&, 8) = FORMATNUM$(PRODUCIDOKG, "F10.0")
        
        GRILLAGRAL.GRID.TextMatrix(J&, 9) = FORMATNUM$(CONSUMIDOUN, "F10.0")
        GRILLAGRAL.GRID.TextMatrix(J&, 10) = FORMATNUM$(CONSUMIDOMETROS, "F10.0")
        GRILLAGRAL.GRID.TextMatrix(J&, 11) = FORMATNUM$(CONSUMIDOKG, "F10.0")
       
        GRILLAGRAL.GRID.TextMatrix(J&, 12) = FORMATNUM$(REMITIDOUN, "F10.0")
        GRILLAGRAL.GRID.TextMatrix(J&, 13) = FORMATNUM$(REMITIDOMETROS, "F10.0")
        GRILLAGRAL.GRID.TextMatrix(J&, 14) = FORMATNUM$(REMITIDOKG, "F10.0")
       
        GRILLAGRAL.GRID.TextMatrix(J&, 15) = FORMATNUM$(STFINALUN, "F10.0")
        GRILLAGRAL.GRID.TextMatrix(J&, 16) = FORMATNUM$(STFINALMETROS, "F10.0")
        GRILLAGRAL.GRID.TextMatrix(J&, 17) = FORMATNUM$(STFINALKG, "F10.0")
       
        TOTINICIAL = TOTINICIAL + XVALO(STINICIAL)
        TOTINICIALMETROS = TOTINICIALMETROS + XVALO(STINIMETROS)
        TOTINICIALKG = TOTINICIALKG + XVALO(STINICIALKG)
        
        TOTPRODUCIDOUN = TOTPRODUCIDOUN + XVALO(PRODUCIDOUN)
        TOTPRODUCIDOMETROS = TOTPRODUCIDOMETROS + XVALO(PRODUCIDOMETROS)
        TOTPRODUCIDOKG = TOTPRODUCIDOKG + XVALO(PRODUCIDOKG)
        
        TOTCONSUMIDOUN = TOTCONSUMIDOUN + XVALO(CONSUMIDOUN)
        CONSUMIDOMETROS = TOTCONSUMIDOMETROS + XVALO(CONSUMIDOMETROS)
        TOTCONSUMIDOKG = TOTCONSUMIDOKG + XVALO(CONSUMIDOKG)
        
        TOTREMITIDOUN = TOTREMITIDOUN + XVALO(REMITIDOUN)
        TOTREMITIDOMETROS = TOTREMITIDOMETROS + XVALO(REMITIDOMETROS)
        TOTREMITIDOKG = TOTREMITIDOKG + XVALO(REMITIDOKG)
        
        TOTSTFINALUN = TOTSTFINALUN + XVALO(STFINALUN)
        TOTSTFINALMETROS = TOTSTFINALMETROS + XVALO(STFINALMETROS)
        TOTSTFINALKG = TOTSTFINALKG + XVALO(STFINALKG)
       .MoveNext
    Loop
   End With
   rs.Close
   Set rs = Nothing
   Screen.MousePointer = 1
   'TOTALES
    GRILLAGRAL.GRID.Rows = J& + 3
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 2) = "TOTALES:"
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 3) = FORMATNUM$(XVALO(TOTINICIAL), "F10.0")
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 4) = FORMATNUM$(XVALO(TOTINICIALMETROS), "F10.0")
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 5) = FORMATNUM$(XVALO(TOTINICIALKG), "F10.0")
    
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 6) = FORMATNUM$(XVALO(TOTPRODUCIDOUN), "F10.0")
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 7) = FORMATNUM$(XVALO(TOTPRODUCIDOMETROS), "F10.0")
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 8) = FORMATNUM$(XVALO(TOTPRODUCIDOKG), "F10.0")
    
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 9) = FORMATNUM$(XVALO(TOTCONSUMIDOUN), "F10.0")
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 10) = FORMATNUM$(XVALO(CONSUMIDOMETROS), "F10.0")
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 11) = FORMATNUM$(XVALO(TOTCONSUMIDOKG), "F10.0")
    
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 12) = FORMATNUM$(XVALO(TOTREMITIDOUN), "F10.0")
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 13) = FORMATNUM$(XVALO(TOTREMITIDOMETROS), "F10.0")
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 14) = FORMATNUM$(XVALO(TOTREMITIDOKG), "F10.0")
    
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 15) = FORMATNUM$(XVALO(TOTSTFINALUN), "F10.0")
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 16) = FORMATNUM$(XVALO(TOTSTFINALMETROS), "F10.0")
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 17) = FORMATNUM$(XVALO(TOTSTFINALKG), "F10.0")
    '
    GRILLAGRAL.GRID.FixedCols = 3
    Call COLOREAR_GRILLASQL_CadenaDeColumnas(GRILLAGRAL.GRID, &HCDDADA, "6;7;8;12;13;14")
    Call COLOREAR_GRILLA_SOLAFILA(GRILLAGRAL.GRID, &H808080, J& + 2, 0, 1)
   GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
   GRILLAGRAL.mnuMenu1.Visible = True
   GRILLAGRAL.GRID.ZOrder 0
   GRILLAGRAL.TXTBUSCAR = ""
   TX$ = ""
   If TM% > 0 Then
     TX$ = TX$ + " Material: " & TIMAT$
   End If
   If FMLIA$ <> "" Then
     TX$ = TX$ + " Familia: " & FAMI$
   End If
   
   GRILLAGRAL.Caption = "STOCK  (" & PERIO$ & ") " & TX$
   
   GRILLAGRAL.Show 1
   
   Dim MSF As msFlexGrid
   If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
'        Call CARGA_TABLA_IMPRE(GRILLAGRAL.MSF_2, Campos$, "STOCKRGO", "STOCKS DE CÓDIGOS RANGO", "TABLA_IMPRE", RET_NAMECONS$, 1)

        Call CARGA_TABLA_IMPRE(MSF, Campos$, "STMATPR", "STOCK X T.MAT. Y FAMILIAS", "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
        Exit Sub
   End If
   GoTo 20
End Sub

Private Sub mnuStockCodigoRango_Click()
    DEPOPAÑOL1% = XVALO(Control$("", "DEPOCEN"))
    
20  Campos$ = "Ci/A0;" '1
    Campos$ = Campos$ + "Código Rango/A16;" '2
    Campos$ = Campos$ + "Descripción/A32;" '3
    Campos$ = Campos$ + "CIMATERIAPRIMA/A0;" '4
    Campos$ = Campos$ + "Código M.Prima/A16;" '5
    Campos$ = Campos$ + "Descripción/A32;" '6
    Campos$ = Campos$ + "Stock/F12.1;" '7
    Campos$ = Campos$ + "Total Rango/F12.1" '8
    
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
    
    
    GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
    GRILLAGRAL.mnuMenu1.Visible = True

    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.TXTBUSCAR = ""
    'CARGO EL RECORDSET
    SQLX$ = "SELECT  CIRANGO,CODRANGO,C.Descripcion,C.CIMATPRIMA,C.CODMATPRIMA,C.DESCRIMATPRIMA,T1.STOCK "
    SQLX$ = SQLX$ + " FROM CODIGORANGO C WITH(NOLOCK)  INNER JOIN (select CODMATPRIMA,CIMATPRIMA, SUM(M.Cantingre - M.CANTSALID) AS STOCK "
    SQLX$ = SQLX$ + " from CODIGORANGO c"
    SQLX$ = SQLX$ + " LEFT join MOVISTO M ON M.Cod_Inte = C.CIMATPRIMA "
    SQLX$ = SQLX$ + " GROUP BY CODMATPRIMA,CIMATPRIMA ) T1 ON C.CIMATPRIMA = T1.CIMATPRIMA"
    SQLX$ = SQLX$ + " ORDER BY CIRANGO "

    Dim obj As New RECORXET
    Set RS1 = obj.RS_CONSULTA_COMPLETA(SQLX$)

    Call Carga_MSF_Recordset(RS1, Campos$, GRILLAGRAL.GRID, 1)

    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    '*******************************************
    Screen.MousePointer = 1
    If GRILLAGRAL.GRID.Rows < 2 Then Exit Sub
    GRILLAGRAL.GRID.Gridlines = flexGridInset
    TX$ = "Reporte de Stocks Código Rango + Detalle : "
    GRILLAGRAL.Caption = TX$
    CANTSTOCK = 0: filas% = 0
    For i& = 1 To GRILLAGRAL.GRID.Rows - 1
       CI1% = XVALO(GRILLAGRAL.GRID.TextMatrix(i&, 1))
       filas% = filas% + 1
       If CI1% <> CI1_ANT% And CI1_ANT% > 0 Then
28        GRILLAGRAL.GRID.TextMatrix(i& - 1, 8) = FORMATNUM$(CANTSTOCK, "f12.1")
          Call NEGRITA_SOLO_UNA_CELDA(GRILLAGRAL.GRID, i& - 1, 8)
          filas% = 0
          CANTSTOCK = 0
       End If
       If i& < GRILLAGRAL.GRID.Rows Then 'para evitar error cuando vuelve if filas%
         CANTSTOCK = CANTSTOCK + XVALO(GRILLAGRAL.GRID.TextMatrix(i&, 7))
       End If
       CI1_ANT% = CI1%
    Next i&
    If filas% > 0 Then GoTo 28
    '*******************************************
'    CANTSTOCK = 0: FILAS% = 0
'    For i& = 1 To GRILLAGRAL.GRID.Rows - 1
'       CI1% = XVALO(GRILLAGRAL.GRID.TextMatrix(i&, 1))
'       FILAS% = FILAS% + 1
'       CANTSTOCK = CANTSTOCK + XVALO(GRILLAGRAL.GRID.TextMatrix(i&, 7))
'       If CI1% <> CI1_ANT% And CI1_ANT% > 0 Then
'28        For J& = FILAS% To i& - FILAS% Step -1
'             GRILLAGRAL.GRID.TextMatrix(J&, 8) = CANTSTOCK
'          Next J&
'          FILAS% = 0
'          CANTSTOCK = 0
'       End If
'       CI1_ANT% = CI1%
'    Next i&
'    If FILAS% > 0 Then GoTo 28 'si filas es > 0 es pq' al llegar al final no entro
    '*******************************************
    
30  GRILLAGRAL.Show 1

     If GRILLAGRAL.MENU1 > 0 Then
            If GRILLAGRAL.MSF_2.Rows > 1 Then
               Call CARGA_TABLA_IMPRE(GRILLAGRAL.MSF_2, Campos$, "STOCKRGO", "STOCKS DE CÓDIGOS RANGO", "TABLA_IMPRE", RET_NAMECONS$, 1)
            Else
               Call CARGA_TABLA_IMPRE(GRILLAGRAL.GRID, Campos$, "STOCKRGO", "STOCKS DE CÓDIGOS RANGO", "TABLA_IMPRE", RET_NAMECONS$, 1)
            End If
            Call FINAL("?" & RET_NAMECONS$)
            GRILLAGRAL.MENU1 = 0
            Screen.MousePointer = 1
     End If
End Sub

Private Sub mnuStockConMetrosyKilos_Click()

    EPAC$ = TRIM$(UCase$(EMPREAC$))
    SQLX$ = "SELECT COD_EXTE,IDENLOTE, COD_INTE,DESCRIP,SUM(CANTINGRE - CANTSALID) AS STOCK,MAS.Umedida"
    SQLX$ = SQLX$ + ", SUM(CASE WHEN Cantingre > 0 THEN pesokg ELSE -pesokg END) AS PESO_TOTAL"
    SQLX$ = SQLX$ + " ,SUM(CASE WHEN Cantingre > 0 THEN METROS ELSE -METROS END) AS METROS_TOTAL ,MAS.Trazable "
    SQLX$ = SQLX$ + " FROM MOVISTO MOV WITH(NOLOCK)  INNER JOIN MASTER MAS ON MAS.CODINT = MOV.Cod_Inte"
    SQLX$ = SQLX$ + " WHERE COD_INTE > 0 and TRAZABLE = 1 AND MAS.Umedida <> 'KG'"
    If InStr(EPAC$, "QUILMES") > 0 Then SQLX$ = SQLX$ + " AND DEPOMOVI NOT IN(9) "
    SQLX$ = SQLX$ + " GROUP BY  IDENLOTE, COD_INTE,DESCRIP,COD_EXTE,MAS.Umedida ,MAS.Trazable"
    SQLX$ = SQLX$ + " Having Sum(CANTINGRE - CANTSALID) > 0.005"
    SQLX$ = SQLX$ + " ORDER BY COD_INTE,COD_EXTE,IDENLOTE"

    Campos$ = "Codigo/A18" '1
    Campos$ = Campos$ + ";Descripcion/A32" '2
    Campos$ = Campos$ + ";Lote/A18" '3
    Campos$ = Campos$ + ";Stock/F9.2" '4
    Campos$ = Campos$ + ";SubTotCant./F11.2" '5
    Campos$ = Campos$ + ";Unid.Medida/A12" '6
    Campos$ = Campos$ + ";Metros/F9.2" '7
    Campos$ = Campos$ + ";SubTotMts/F11.2" '8
    Campos$ = Campos$ + ";PesoKg./F9.2" '9
    Campos$ = Campos$ + ";SubTotKgs/F11.2" '10
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)

    Screen.MousePointer = 11
    Set rs = READSET(SQLX$)

    With rs
        Do While Not .EOF
            J& = J& + 1
            CI1% = XVALO(!cod_inte)
            GRILLAGRAL.GRID.Rows = J& + 1
            GRILLAGRAL.GRID.TextMatrix(J&, 1) = SAFETEXT$(!Cod_Exte)
            GRILLAGRAL.GRID.TextMatrix(J&, 2) = SAFETEXT$(!Descrip)
            GRILLAGRAL.GRID.TextMatrix(J&, 3) = SAFETEXT$(!idenlote)
            GRILLAGRAL.GRID.TextMatrix(J&, 4) = FORMATNUM(XVALO(!STOCK), "F12.1")
            SUBSTOCK# = SUBSTOCK# + XVALO(GRILLAGRAL.GRID.TextMatrix(J&, 4))
'            GRILLAGRAL.GRID.TextMatrix(J&, 5) = SUBSTOCK#
            GRILLAGRAL.GRID.TextMatrix(J&, 6) = SAFETEXT$(!UMEDIDA)
            If SAFETEXT$(!UMEDIDA) = "M." Then  'SI LA U.MEDIDA ES METROS LA COLUMNA METROS = STOCK
              GRILLAGRAL.GRID.TextMatrix(J&, 7) = FORMATNUM(XVALO(!STOCK), "F12.1")
            Else
              GRILLAGRAL.GRID.TextMatrix(J&, 7) = FORMATNUM(XVALO(!METROS_TOTAL), "F12.1")
            End If

            SUBMETROS_TOTAL# = SUBMETROS_TOTAL# + XVALO(GRILLAGRAL.GRID.TextMatrix(J&, 7))
'            GRILLAGRAL.GRID.TextMatrix(J&, 8) = SUBMETROS_TOTAL#
            If SAFETEXT$(!UMEDIDA) = "KG" Then 'SI LA U.MEDIDA ES KG LA COLUMNA PESO_TOTAL = STOCK
               GRILLAGRAL.GRID.TextMatrix(J&, 9) = FORMATNUM(XVALO(!STOCK), "F12.1")
            Else
               GRILLAGRAL.GRID.TextMatrix(J&, 9) = FORMATNUM(XVALO(!PESO_TOTAL), "F12.1")
            End If
            SUBPESO_TOTAL# = SUBPESO_TOTAL# + XVALO(GRILLAGRAL.GRID.TextMatrix(J&, 9))
'            GRILLAGRAL.GRID.TextMatrix(J&, 10) = FORMATNUM(XVALO(SUBPESO_TOTAL#), "F12.1")
            CI1_ANT% = CI1%
            COD$ = SAFETEXT$(!Cod_Exte)
            DESC$ = SAFETEXT$(!Descrip)

            .MoveNext 'AVANZA
            CI1_SIGUIENTE% = XVALO(!cod_inte) 'AQUI YA LEE EL SIGUIENTE Y VALIDA SI CAMBIO
            COD_SIG$ = SAFETEXT$(!Cod_Exte)
            DESC_SIG$ = SAFETEXT$(!Descrip)

            If CI1_ANT% <> CI1_SIGUIENTE% Then
             'actualiza todavia con el mismo indice del ultimo registro de la grilla (por que j& no se incremento aun)
             GRILLAGRAL.GRID.TextMatrix(J&, 5) = SUBSTOCK#: Call NEGRITA_SOLO_UNA_CELDA(GRILLAGRAL.GRID, J&, 5)
             GRILLAGRAL.GRID.TextMatrix(J&, 8) = SUBMETROS_TOTAL#: Call NEGRITA_SOLO_UNA_CELDA(GRILLAGRAL.GRID, J&, 8)
             GRILLAGRAL.GRID.TextMatrix(J&, 10) = SUBPESO_TOTAL#: Call NEGRITA_SOLO_UNA_CELDA(GRILLAGRAL.GRID, J&, 10)
             SUBSTOCK# = 0
             SUBMETROS_TOTAL# = 0
             SUBPESO_TOTAL# = 0
            End If

        Loop
    End With
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    On Error GoTo 0

    '_____________________________________________________
    J& = J& + 1
    SQLX$ = "SELECT COD_EXTE, COD_INTE,DESCRIP,SUM(CANTINGRE - CANTSALID) AS STOCK"
    SQLX$ = SQLX$ + ",MAS.Umedida, MAS.Trazable"
    SQLX$ = SQLX$ + " FROM MOVISTO MOV WITH(NOLOCK)  INNER JOIN MASTER MAS ON MAS.CODINT = MOV.Cod_Inte"
    SQLX$ = SQLX$ + " WHERE COD_INTE > 0  AND (Trazable = 0 OR Trazable IS NULL) AND MAS.Umedida = 'KG'"
    If InStr(EPAC$, "QUILMES") > 0 Then SQLX$ = SQLX$ + " AND DEPOMOVI NOT IN(9) "
    SQLX$ = SQLX$ + " GROUP BY  COD_INTE,DESCRIP,COD_EXTE,MAS.Umedida,MAS.Trazable"
    SQLX$ = SQLX$ + " Having Sum(CANTINGRE - CANTSALID) > 0.005 ORDER BY COD_INTE,COD_EXTE"
    Set rs = READSET(SQLX$)

    With rs
        Do While Not .EOF
            J& = J& + 1
            CI1% = XVALO(!cod_inte)
            GRILLAGRAL.GRID.Rows = J& + 1
            GRILLAGRAL.GRID.TextMatrix(J&, 1) = SAFETEXT$(!Cod_Exte)
            GRILLAGRAL.GRID.TextMatrix(J&, 2) = SAFETEXT$(!Descrip)
            GRILLAGRAL.GRID.TextMatrix(J&, 3) = "-"
            GRILLAGRAL.GRID.TextMatrix(J&, 4) = FORMATNUM(XVALO(!STOCK), "F12.1")
            SUBSTOCK# = SUBSTOCK# + XVALO(GRILLAGRAL.GRID.TextMatrix(J&, 4))
'            GRILLAGRAL.GRID.TextMatrix(J&, 5) = SUBSTOCK#
            GRILLAGRAL.GRID.TextMatrix(J&, 6) = SAFETEXT$(!UMEDIDA)
            .MoveNext
        Loop
    End With
    
    '______________________________________________________
    Screen.MousePointer = 1
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    On Error GoTo 0
    
    ' ********************
    GRILLAGRAL.Caption = "Stock No Considera Depósito 9"
    GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.TXTBUSCAR = ""
    GRILLAGRAL.SinEfectoClickEnGrilla = 1
    
    Screen.MousePointer = 1
    GRILLAGRAL.Show 1
    Dim MSF As msFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(MSF, Campos$, "STOCKMEKI", "Stock No Considera Depósito 9", "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
        Exit Sub
   End If
   Set obj = Nothing
       
99  Exit Sub

End Sub


Private Sub mnuStockConMetrosyKilosEnvaplast_Click()

     Call SELECHO("SELFECHA")
     VDEVU$ = VALACT1$("SELFECHA")
     If VDEVU$ <> "" Then
        fechaHasta% = FECHANUM(VDEVU$)
     Else
        fechaHasta% = HOY(HO$)
     End If
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    
     
     
'    SQLX$ = "SELECT COD_EXTE, COD_INTE,DESCRIP,FORMAT(SUM(CANTINGRE - CANTSALID), 'N2') AS STOCK,MAS.Umedida"
'    SQLX$ = SQLX$ + ", FORMAT(SUM(CASE WHEN Cantingre > 0 THEN pesokg ELSE -pesokg END), 'N2') AS PESO_TOTAL"
'    SQLX$ = SQLX$ + " ,FORMAT(SUM(CASE WHEN Cantingre > 0 THEN METROS ELSE -METROS END), 'N2') AS METROS_TOTAL"
'    SQLX$ = SQLX$ + " FROM MOVISTO MOV WITH(NOLOCK)  INNER JOIN MASTER MAS ON MAS.CODINT = MOV.Cod_Inte"
'    SQLX$ = SQLX$ + " WHERE FECHMOV <= '" & FETEXCOR1$(fechaHasta%) & "'"
'    SQLX$ = SQLX$ + " AND MOV.DEPOMOVI = 1 AND MAS.FAMILIA = 'TE' AND MAS.TIMATE = 41" 'BOBINAS Y FAMILIA TELAS
'    SQLX$ = SQLX$ + " GROUP BY   COD_INTE,DESCRIP,COD_EXTE,MAS.Umedida"
'    SQLX$ = SQLX$ + " Having Sum(CANTINGRE - CANTSALID) > 0.005"
'    SQLX$ = SQLX$ + " ORDER BY COD_INTE,COD_EXTE"

    SQLX$ = "SELECT  COD_INTE,FORMAT(SUM(CANTINGRE - CANTSALID), 'N2') AS STOCK"
    SQLX$ = SQLX$ + ", FORMAT(SUM(CASE WHEN Cantingre > 0 THEN pesokg ELSE -pesokg END), 'N2') AS PESO_TOTAL"
    SQLX$ = SQLX$ + " ,FORMAT(SUM(CASE WHEN Cantingre > 0 THEN METROS ELSE -METROS END), 'N2') AS METROS_TOTAL"
    SQLX$ = SQLX$ + " FROM MOVISTO MOV WITH(NOLOCK)  INNER JOIN MASTER MAS ON MAS.CODINT = MOV.Cod_Inte"
    SQLX$ = SQLX$ + " WHERE FECHMOV <= '" & FETEXCOR1$(fechaHasta%) & "'"
    SQLX$ = SQLX$ + " AND MOV.DEPOMOVI = 1 AND MAS.FAMILIA = 'TE' AND MAS.TIMATE = 41 and MAS.UMEDIDA = 'M.'" 'BOBINAS Y FAMILIA TELAS
    SQLX$ = SQLX$ + " GROUP BY   COD_INTE"
    SQLX$ = SQLX$ + " Having Sum(CANTINGRE - CANTSALID) > 0.005"
    SQLX$ = SQLX$ + " ORDER BY COD_INTE"

    Campos$ = "Codigo/A18" '1
    Campos$ = Campos$ + ";Descripcion/A32" '2
    Campos$ = Campos$ + ";Stock Mts./F9.2" '3
    Campos$ = Campos$ + ";PesoKg./F9.2" '4
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)

    Screen.MousePointer = 11
    Set rs = READSET(SQLX$)

    With rs
        Do While Not .EOF
            J& = J& + 1
            CI1% = XVALO(!cod_inte)
            GRILLAGRAL.GRID.Rows = J& + 1
            GRILLAGRAL.GRID.TextMatrix(J&, 1) = CODEXT$(CI1%)
            GRILLAGRAL.GRID.TextMatrix(J&, 2) = DESCRIT0(CI1%)
            GRILLAGRAL.GRID.TextMatrix(J&, 3) = XVALO(!STOCK)
            GRILLAGRAL.GRID.TextMatrix(J&, 4) = FORMATNUM$(PesoXMetroPromedio#(CI1%) * XVALO(!STOCK), "F10.2")
            
'            GRILLAGRAL.GRID.TextMatrix(J&, 4) = XVALO(!PESO_TOTAL) LOS KG, NO ESTA BIEN LA OPCION SERIA TRAER LOS KG SOLO DE LOS QUE PSEEN SALDO > 0
 '     SI NO VER SI TRAER LOS KG EN BASE A LO DEL MAESTRO.
'            CI1_ANT% = CI1%
            .MoveNext 'AVANZA
'            CI1_SIGUIENTE% = XVALO(!cod_inte) 'AQUI YA LEE EL SIGUIENTE Y VALIDA SI CAMBIO
'            COD_SIG$ = SAFETEXT$(!Cod_Exte)
'            DESC_SIG$ = SAFETEXT$(!Descrip)
'
'            If CI1_ANT% <> CI1_SIGUIENTE% Then
'             'actualiza todavia con el mismo indice del ultimo registro de la grilla (por que j& no se incremento aun)
'             GRILLAGRAL.GRID.TextMatrix(J&, 5) = SUBSTOCK#: Call NEGRITA_SOLO_UNA_CELDA(GRILLAGRAL.GRID, J&, 5)
'             GRILLAGRAL.GRID.TextMatrix(J&, 8) = SUBMETROS_TOTAL#: Call NEGRITA_SOLO_UNA_CELDA(GRILLAGRAL.GRID, J&, 8)
'             GRILLAGRAL.GRID.TextMatrix(J&, 10) = SUBPESO_TOTAL#: Call NEGRITA_SOLO_UNA_CELDA(GRILLAGRAL.GRID, J&, 10)
'             SUBSTOCK# = 0
'             SUBMETROS_TOTAL# = 0
'             SUBPESO_TOTAL# = 0
'            End If

        Loop
    End With
    Screen.MousePointer = 1
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    On Error GoTo 0
    
    ' ********************
    GRILLAGRAL.Caption = "Stock "
    GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.TXTBUSCAR = ""
    GRILLAGRAL.SinEfectoClickEnGrilla = 1
    
    Screen.MousePointer = 1
    GRILLAGRAL.Show 1
    Dim MSF As msFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(MSF, Campos$, "STOCKMEKI", "Stock", "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
        Exit Sub
   End If
   Set obj = Nothing
       
99  Exit Sub

End Sub

Private Sub mnuStockMetrosvsFtecnica_Click()


     
     SQLX$ = "SELECT Cod_Exte,M.DESCRIP,IDENLOTE,SUM(CANTINGRE - CANTSALID) AS RESUL, F.BOBINASMT,M.METROS,M.PESOKG"
     SQLX$ = SQLX$ + " FROM MOVISTO M WITH(NOLOCK)"
     SQLX$ = SQLX$ + " INNER JOIN MASTER MAS ON M.Cod_Inte = MAS.Codint"
     SQLX$ = SQLX$ + " inner join FTECNICACAMB F ON M.COD_INTE = F.CODINT"
     SQLX$ = SQLX$ + " WHERE MAS.Umedida = 'U.'"
     SQLX$ = SQLX$ + " GROUP BY COD_EXTE,M.DESCRIP, IDENLOTE, F.BOBINASMT ,M.METROS,M.PESOKG"
     SQLX$ = SQLX$ + " Having Sum(CANTINGRE - CANTSALID) = 1 And BOBINASMT <> METROS"
     SQLX$ = SQLX$ + " ORDER BY BOBINASMT, METROS"
     
    'HAVING SUM(CANTINGRE - CANTSALID) = 1 ESTO POR QUE EN LOS CASOS QUE SON MAS DE 1 SE HACE DESPELOTE EJEMPLO
    'SELECT * FROM MOVISTO WHERE Cod_Exte = 'CA TP 101 52 BL' AND IDENLOTE = 'OF-000132076' QUE TIENE UN AJUSTE DE 5

    Campos$ = "Codigo/A18;Descripcion/A48;Lote/A18;Cantidad/F6;Mts.Teoricos/f12;Mts.Reales/f10;Kg./F10.1"
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)


    Set rs = READSET(SQLX$)

    With rs
        Do While Not .EOF

            J& = J& + 1
        
            GRILLAGRAL.GRID.Rows = J& + 1
            
            GRILLAGRAL.GRID.TextMatrix(J&, 1) = SAFETEXT$(!Cod_Exte)
            GRILLAGRAL.GRID.TextMatrix(J&, 2) = SAFETEXT$(!Descrip)
            GRILLAGRAL.GRID.TextMatrix(J&, 3) = SAFETEXT$(!idenlote)
            GRILLAGRAL.GRID.TextMatrix(J&, 4) = XVALO(!RESUL)
            GRILLAGRAL.GRID.TextMatrix(J&, 5) = XVALO(!BOBINASMT)
            GRILLAGRAL.GRID.TextMatrix(J&, 6) = XVALO(!METROS)
            GRILLAGRAL.GRID.TextMatrix(J&, 7) = XVALO(!PESOKG)
            
            .MoveNext
        Loop
    End With
    
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    On Error GoTo 0
    
    ' ********************
    GRILLAGRAL.Caption = "Stock de Bobinas con Diferencia con los Metros Teóricos"
    GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.TXTBUSCAR = ""
    GRILLAGRAL.SinEfectoClickEnGrilla = 1
    
    Screen.MousePointer = 1
    GRILLAGRAL.Show 1
    Dim MSF As msFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(MSF, Campos$, "STDIFTEO", "Stock de Bobinas con Diferencia con los Metros Teóricos", "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
        Exit Sub
   End If
   Set obj = Nothing
       
99  Exit Sub

End Sub




Private Sub mnuStockMulDep_Click()
    DepoMult.Show 1
End Sub

Private Sub mnuStockProveedorRango_Click()

    Campos$ = "Cta/I8;Proveedor/A24;Rango/I6;Denominacion/A48;Codigo/A18;Descripcion/A32;Bobinas/I10;Metros/G14.0;Subtotal Mts/G14.0;Kilos/G14.0;Subtotal Kg/G14.0"
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
    
    On Error Resume Next
    FLEXCONN.Execute "drop view StockRangoProveedor"
    On Error GoTo 0
    
    ConsultaSql$ = "create view StockRangoProveedor as " _
    & " (select stloteran.rango, stloteran.denominacion, stloteran.codinte, m.codigo, m.descripcion, stloteran.lote, stloteran.metros, stloteran.prov,  m.pesmetro pesometromas, stloteran.pesometrorece from" _
    & " (select c.CODRANGO rango, c.Descripcion denominacion, stlote.codinte, stlote.lote, stlote.prov, (stlote.ingreso-stlote.salida) metros, (stlote.peso/stlote.cantnom) pesometrorece from" _
    & " (select * from" _
    & " (select m.Cod_Inte codinte, m.idenlote lote, b.NumProv prov, sum(cantsalid) salida, sum(cantingre) ingreso, sum(peso) peso, sum(cantnom) cantnom from bolrecep b inner join movisto m on (b.IdenLote=m.IdenLote)" _
    & " where m.depomovi=1 group by m.cod_inte, m.idenlote, b.NumProv) st" _
    & " where (st.ingreso-st.salida) > 0.05) stlote" _
    & " left join codigorango c on (stlote.codinte=c.CIMATPRIMA)) stloteran" _
    & " inner join master m on (stloteran.codinte=m.Codint))"
    FLEXCONN.Execute ConsultaSql$

    Dim RstStockRangoProveedor As ADODB.Recordset
    ConsultaSql$ = "select prov, rango, denominacion, codigo, descripcion, count(lote) bobinas, sum(metros) metros, sum(kilos) kilos from" _
    & " (select prov, rango, denominacion, codigo, descripcion, lote, metros, (metros*pesometrorece) kilos from stockrangoproveedor with(nolock)) st" _
    & " group by prov, rango, denominacion, codigo, descripcion order by prov, rango, codigo"
    Set RstStockRangoProveedor = READSET(ConsultaSql$)
    
    RangoActual& = 0
    ColorFila1 = RGB(162, 210, 196)
    ColorFila2 = RGB(201, 184, 143)
    ColorFila = ColorFila2
    CuentaActiva& = 0
    With RstStockRangoProveedor
        Do While Not .EOF
            Cuenta& = XVALO(!PROV)
            razonSocial$ = rasocli$((-1) * Cuenta&)
            If CuentaActiva& = 0 Then CuentaActiva& = Cuenta&
            Rango& = XVALO(!Rango)
            denominacion$ = SAFETEXT$(!denominacion)
            Codigo$ = !Codigo
            Descripcion$ = !Descripcion
            bobinas% = !bobinas
            METROS# = XVALO(!METROS)
            KILOS# = XVALO(!KILOS)
            
            With GRILLAGRAL
                .GRID.Rows = .GRID.Rows + 1
                
                .GRID.TextMatrix(.GRID.Rows - 1, 1) = CStr(Cuenta&)
                .GRID.TextMatrix(.GRID.Rows - 1, 2) = razonSocial$
                .GRID.TextMatrix(.GRID.Rows - 1, 3) = CStr(Rango&)
                .GRID.TextMatrix(.GRID.Rows - 1, 4) = denominacion$
                .GRID.TextMatrix(.GRID.Rows - 1, 5) = Codigo$
                .GRID.TextMatrix(.GRID.Rows - 1, 6) = Descripcion$
                .GRID.TextMatrix(.GRID.Rows - 1, 7) = CStr(bobinas%)
                .GRID.TextMatrix(.GRID.Rows - 1, 8) = TRIM$(FORMATNUM$(METROS#, "F14.0"))
                .GRID.TextMatrix(.GRID.Rows - 1, 10) = TRIM$(FORMATNUM$(KILOS#, "F14.0"))
                
                If CuentaActiva& <> Cuenta& Then
                    CuentaActiva& = Cuenta&
                    If ColorFila = ColorFila1 Then
                        ColorFila = ColorFila2
                    Else
                        ColorFila = ColorFila1
                    End If
                    .GRID.TextMatrix(.GRID.Rows - 2, 9) = TRIM$(FORMATNUM$(TOTALMETROS#, "F14.0"))
                    .GRID.TextMatrix(.GRID.Rows - 2, 11) = TRIM$(FORMATNUM$(TotalKilos#, "F14.0"))
                    RangoActual& = Rango&
                    TOTALMETROS# = METROS#
                    TotalKilos# = KILOS#
                    Call COLOREAR_GRILLA_SOLAFILA(.GRID, ColorFila, .GRID.Rows - 1, 0, 1)
                Else
                    If RangoActual& <> Rango& And .GRID.Rows > 2 Then
                        .GRID.TextMatrix(.GRID.Rows - 2, 9) = TRIM$(FORMATNUM$(TOTALMETROS#, "F14.0"))
                        .GRID.TextMatrix(.GRID.Rows - 2, 11) = TRIM$(FORMATNUM$(TotalKilos#, "F14.0"))
                        RangoActual& = Rango&
                        TOTALMETROS# = METROS#
                        TotalKilos# = KILOS#
                    Else
                        RangoActual& = Rango&
                        TOTALMETROS# = TOTALMETROS# + METROS#
                        TotalKilos# = TotalKilos# + KILOS#
                    End If
                    
                    Call COLOREAR_GRILLA_SOLAFILA(.GRID, ColorFila, .GRID.Rows - 1, 0, 1)
                End If
            End With
            
            .MoveNext
        Loop
        
        GRILLAGRAL.GRID.TextMatrix(GRILLAGRAL.GRID.Rows - 1, 9) = TRIM$(FORMATNUM$(TOTALMETROS#, "F14.0"))
        GRILLAGRAL.GRID.TextMatrix(GRILLAGRAL.GRID.Rows - 1, 11) = TRIM$(FORMATNUM$(TotalKilos#, "F14.0"))
    End With
    
    On Error Resume Next
    RstStockRangoProveedor.Close
    Set RstStockRangoProveedor = Nothing
    
    FLEXCONN.Execute "drop view StockRangoProveedor"
    On Error GoTo 0
    
    ' ********************
    GRILLAGRAL.Caption = "Stock por Proveedor y Rango"
    
    GRILLAGRAL.mnuMenu1.Caption = "Exportar a Excel"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.MnuEdición.Visible = False
    GRILLAGRAL.SinEfectoClickEnGrilla = 1
    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.GRID.FixedCols = 3
    GRILLAGRAL.GRID.MergeCells = flexMergeRestrictColumns
    GRILLAGRAL.GRID.MergeCol(1) = True
    GRILLAGRAL.GRID.MergeCol(2) = True
    GRILLAGRAL.GRID.ColAlignment(1) = 7
    GRILLAGRAL.MSF_2.FixedCols = 3
    GRILLAGRAL.MSF_2.MergeCells = flexMergeRestrictColumns
    GRILLAGRAL.MSF_2.MergeCol(1) = True
    GRILLAGRAL.MSF_2.MergeCol(2) = True
    GRILLAGRAL.MSF_2.ColAlignment(1) = 7
    
    Screen.MousePointer = 1
    GRILLAGRAL.Show 1
    Dim MSF As msFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
        CommonDialog1.CancelError = True
        CommonDialog1.DialogTitle = "Exportar a Excel"
            ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
        CommonDialog1.FILTER = "Excel Files(*.XLS)|*.XLS|"
        CommonDialog1.FilterIndex = 1
        On Error GoTo 99
        CommonDialog1.ShowOpen
        On Error GoTo 0
        '
        DoEvents
        Screen.MousePointer = 11
        AREXPO$ = UCase$(CommonDialog1.FileName)
        If TRIM$(AREXPO$) <> "" Then
           Call GeneraExcelDeGrilla(AREXPO$, GRILLAGRAL.GRID, GRILLAGRAL.Caption, False, Campos$)
        End If
   End If
   Set obj = Nothing
       
99  Exit Sub

End Sub

Private Sub mnuStockRangoAgrupado_Click()

    Campos$ = "Rango/I6;Descripcion/A48;Codigo/A18;Descripcion/A32;Bobinas/I10;Metros/G14.0;Subtotal Mts/G14.0;Kilos/G14.0;Subtotal Kg/G14.0"
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
    
    On Error Resume Next
    FLEXCONN.Execute "drop view StockRangoProveedor"
    On Error GoTo 0
    
    ConsultaSql$ = "create view StockRangoProveedor as " _
    & " (select stloteran.rango, stloteran.denominacion, stloteran.codinte, m.codigo, m.descripcion, stloteran.lote, stloteran.metros, stloteran.prov,  m.pesmetro pesometromas, stloteran.pesometrorece from" _
    & " (select c.CODRANGO rango, c.Descripcion denominacion, stlote.codinte, stlote.lote, stlote.prov, (stlote.ingreso-stlote.salida) metros, (stlote.peso/stlote.cantnom) pesometrorece from" _
    & " (select * from" _
    & " (select m.Cod_Inte codinte, m.idenlote lote, b.NumProv prov, sum(cantsalid) salida, sum(cantingre) ingreso, sum(peso) peso, sum(cantnom) cantnom from bolrecep b inner join movisto m on (b.IdenLote=m.IdenLote)" _
    & " where m.depomovi=1 group by m.cod_inte, m.idenlote, b.NumProv) st" _
    & " where (st.ingreso-st.salida) > 0.05) stlote" _
    & " left join codigorango c on (stlote.codinte=c.CIMATPRIMA)) stloteran" _
    & " inner join master m on (stloteran.codinte=m.Codint))"
    FLEXCONN.Execute ConsultaSql$

    Dim RstStockRango As ADODB.Recordset
    ConsultaSql$ = "select rango, denominacion, codigo, descripcion, count(lote) bobinas, sum(metros) metros, sum(kilos) kilos from" _
    & " (select rango, denominacion, codigo, descripcion, lote, metros, (metros*pesometrorece) kilos from stockrangoproveedor with(nolock)) st" _
    & " group by rango, denominacion, codigo, descripcion order by rango, codigo"
    Set RstStockRango = READSET(ConsultaSql$)
    
    RangoActual& = 0
    ColorFila1 = RGB(162, 210, 196)
    ColorFila2 = RGB(201, 184, 143)
    ColorFila = ColorFila2
    With RstStockRango
        Do While Not .EOF
            Rango& = XVALO(!Rango)
            denominacion$ = SAFETEXT$(!denominacion)
            Codigo$ = !Codigo
            Descripcion$ = !Descripcion
            bobinas% = !bobinas
            METROS# = XVALO(!METROS)
            KILOS# = XVALO(!KILOS)
            
            With GRILLAGRAL
                .GRID.Rows = .GRID.Rows + 1
                
                .GRID.TextMatrix(.GRID.Rows - 1, 1) = CStr(Rango&)
                .GRID.TextMatrix(.GRID.Rows - 1, 2) = denominacion$
                .GRID.TextMatrix(.GRID.Rows - 1, 3) = Codigo$
                .GRID.TextMatrix(.GRID.Rows - 1, 4) = Descripcion$
                .GRID.TextMatrix(.GRID.Rows - 1, 5) = CStr(bobinas%)
                .GRID.TextMatrix(.GRID.Rows - 1, 6) = TRIM$(FORMATNUM$(METROS#, "F14.0"))
                .GRID.TextMatrix(.GRID.Rows - 1, 8) = TRIM$(FORMATNUM$(KILOS#, "F14.0"))
                
                If RangoActual& <> Rango& And Rango& > 0 Then
                    .GRID.TextMatrix(.GRID.Rows - 2, 7) = TRIM$(FORMATNUM$(TOTALMETROS#, "F14.0"))
                    .GRID.TextMatrix(.GRID.Rows - 2, 9) = TRIM$(FORMATNUM$(TotalKilos#, "F14.0"))
                    If ColorFila = ColorFila1 Then
                        ColorFila = ColorFila2
                    Else
                        ColorFila = ColorFila1
                    End If
                    RangoActual& = Rango&
                    TOTALMETROS# = METROS#
                    TotalKilos# = KILOS#
                Else
                    TOTALMETROS# = TOTALMETROS# + METROS#
                    TotalKilos# = TotalKilos# + KILOS#
                End If
                
                Call COLOREAR_GRILLA_SOLAFILA(.GRID, ColorFila, .GRID.Rows - 1, 0, 1)
            End With
            
            .MoveNext
        Loop
        
        GRILLAGRAL.GRID.TextMatrix(GRILLAGRAL.GRID.Rows - 1, 7) = TRIM$(FORMATNUM$(TOTALMETROS#, "F14.0"))
        GRILLAGRAL.GRID.TextMatrix(GRILLAGRAL.GRID.Rows - 1, 9) = TRIM$(FORMATNUM$(TotalKilos#, "F14.0"))
    End With
    
    On Error Resume Next
    RstStockRango.Close
    Set RstStockRango = Nothing
    
    FLEXCONN.Execute "drop view StockRangoProveedor"
    On Error GoTo 0
    
    ' ********************
    GRILLAGRAL.Caption = "Stock por Rango (AGRUPADO)"
    
    GRILLAGRAL.mnuMenu1.Caption = "Exportar a Excel"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.MnuEdición.Visible = False
    GRILLAGRAL.SinEfectoClickEnGrilla = 1
    GRILLAGRAL.GRID.ZOrder 0
    
    Screen.MousePointer = 1
    GRILLAGRAL.Show 1
    Dim MSF As msFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
        CommonDialog1.CancelError = True
        CommonDialog1.DialogTitle = "Exportar a Excel"
            ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
        CommonDialog1.FILTER = "Excel Files(*.XLS)|*.XLS|"
        CommonDialog1.FilterIndex = 1
        On Error GoTo 99
        CommonDialog1.ShowOpen
        On Error GoTo 0
        '
        DoEvents
        Screen.MousePointer = 11
        AREXPO$ = UCase$(CommonDialog1.FileName)
        If TRIM$(AREXPO$) <> "" Then
           Call GeneraExcelDeGrilla(AREXPO$, GRILLAGRAL.GRID, GRILLAGRAL.Caption, False, Campos$)
        End If
   End If
   Set obj = Nothing
       
99  Exit Sub

End Sub


Private Sub mnuStockRangoTodos_Click()

    Campos$ = "Rango/I6;Descripcion/A48;Codigo/A18;Descripcion/A32;Lote/A16;Cta/I10;Proveedor/A24;Metros/G14.0;Subtotal Mts/G14.0;Kilos/G14.0;Subtotal Kg/G14.0"
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
    
    On Error Resume Next
    FLEXCONN.Execute "drop view StockRangoProveedor"
    On Error GoTo 0
    
    ConsultaSql$ = "create view StockRangoProveedor as " _
    & " (select stloteran.rango, stloteran.denominacion, stloteran.codinte, m.codigo, m.descripcion, stloteran.lote, stloteran.metros, stloteran.prov,  m.pesmetro pesometromas, stloteran.pesometrorece from" _
    & " (select c.CODRANGO rango, c.Descripcion denominacion, stlote.codinte, stlote.lote, stlote.prov, (stlote.ingreso-stlote.salida) metros, (stlote.peso/stlote.cantnom) pesometrorece from" _
    & " (select * from" _
    & " (select m.Cod_Inte codinte, m.idenlote lote, b.NumProv prov, sum(cantsalid) salida, sum(cantingre) ingreso, sum(peso) peso, sum(cantnom) cantnom from bolrecep b inner join movisto m on (b.IdenLote=m.IdenLote)" _
    & " where m.depomovi=1 group by m.cod_inte, m.idenlote, b.NumProv) st" _
    & " where (st.ingreso-st.salida) > 0.05) stlote" _
    & " left join codigorango c on (stlote.codinte=c.CIMATPRIMA)) stloteran" _
    & " inner join master m on (stloteran.codinte=m.Codint))"
    FLEXCONN.Execute ConsultaSql$

    Dim RstStockRango As ADODB.Recordset
    ConsultaSql$ = "select * from StockRangoProveedor with(nolock) order by rango, codigo"
    Set RstStockRango = READSET(ConsultaSql$)
    
    RangoActual& = 0
    ColorFila1 = RGB(162, 210, 196)
    ColorFila2 = RGB(201, 184, 143)
    ColorFila = ColorFila2
    With RstStockRango
        Do While Not .EOF
            CI1% = XVALO(!codinte)
            Rango& = XVALO(!Rango)
            denominacion$ = SAFETEXT$(!denominacion)
            Codigo$ = !Codigo
            Descripcion$ = !Descripcion
            LOTE$ = !LOTE
            Cuenta& = XVALO(!PROV)
            razonSocial$ = rasocli$((-1) * Cuenta&)
            METROS# = XVALO(!METROS)
            PesoMetro# = XVALO(!pesometrorece)
            
            With GRILLAGRAL
                .GRID.Rows = .GRID.Rows + 1
                
                .GRID.TextMatrix(.GRID.Rows - 1, 1) = CStr(Rango&)
                .GRID.TextMatrix(.GRID.Rows - 1, 2) = denominacion$
                .GRID.TextMatrix(.GRID.Rows - 1, 3) = Codigo$
                .GRID.TextMatrix(.GRID.Rows - 1, 4) = Descripcion$
                .GRID.TextMatrix(.GRID.Rows - 1, 5) = LOTE$
                .GRID.TextMatrix(.GRID.Rows - 1, 6) = CStr(Cuenta&)
                .GRID.TextMatrix(.GRID.Rows - 1, 7) = razonSocial$
                .GRID.TextMatrix(.GRID.Rows - 1, 8) = TRIM$(FORMATNUM$(METROS#, "F14.0"))
                    'KILOS# = METROS# * PesoMetro#  SE CAMBIA PARA QUE TRAIGA EL ULTIMO PESO INFORMADO TIPO TR CON INGRESO AL DEPOSITO 1
                    KILOS# = KilosSegunUltimoMovimientoDeIngresoAlDepo1#(CI1%, LOTE$, 1, "TR")
                    If KILOS# <= 0.05 Then KILOS# = METROS# * PesoMetro# 'POR DEFECTO SI NO DEVUELVE INFO MANTENGO LO ANTERIOR
                    
                .GRID.TextMatrix(.GRID.Rows - 1, 10) = TRIM$(FORMATNUM$(KILOS#, "F14.0"))
                
                If RangoActual& <> Rango& And Rango& > 0 Then
                    .GRID.TextMatrix(.GRID.Rows - 2, 9) = TRIM$(FORMATNUM$(TOTALMETROS#, "F14.0"))
                    .GRID.TextMatrix(.GRID.Rows - 2, 11) = TRIM$(FORMATNUM$(TotalKilos#, "F14.0"))
                    If ColorFila = ColorFila1 Then
                        ColorFila = ColorFila2
                    Else
                        ColorFila = ColorFila1
                    End If
                    RangoActual& = Rango&
                    TOTALMETROS# = METROS#
                    TotalKilos# = KILOS#
                Else
                    TOTALMETROS# = TOTALMETROS# + METROS#
                    TotalKilos# = TotalKilos# + KILOS#
                End If
                
                Call COLOREAR_GRILLA_SOLAFILA(.GRID, ColorFila, .GRID.Rows - 1, 0, 1)
            End With
            
            .MoveNext
        Loop
        
        GRILLAGRAL.GRID.TextMatrix(GRILLAGRAL.GRID.Rows - 1, 9) = TRIM$(FORMATNUM$(TOTALMETROS#, "F14.0"))
        GRILLAGRAL.GRID.TextMatrix(GRILLAGRAL.GRID.Rows - 1, 11) = TRIM$(FORMATNUM$(TotalKilos#, "F14.0"))
    End With
    
    On Error Resume Next
    RstStockRango.Close
    Set RstStockRango = Nothing
    
    FLEXCONN.Execute "drop view StockRangoProveedor"
    On Error GoTo 0
    
    ' ********************
    GRILLAGRAL.Caption = "Stock por Rango (TODOS)"
    
    GRILLAGRAL.mnuMenu1.Caption = "Exportar a Excel"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.MnuEdición.Visible = False
    GRILLAGRAL.SinEfectoClickEnGrilla = 1
    GRILLAGRAL.GRID.ZOrder 0
    
    Screen.MousePointer = 1
    GRILLAGRAL.Show 1
    Dim MSF As msFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
        CommonDialog1.CancelError = True
        CommonDialog1.DialogTitle = "Exportar a Excel"
            ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
        CommonDialog1.FILTER = "Excel Files(*.XLS)|*.XLS|"
        CommonDialog1.FilterIndex = 1
        On Error GoTo 99
        CommonDialog1.ShowOpen
        On Error GoTo 0
        '
        DoEvents
        Screen.MousePointer = 11
        AREXPO$ = UCase$(CommonDialog1.FileName)
        If TRIM$(AREXPO$) <> "" Then
           Call GeneraExcelDeGrillaTabulada(AREXPO$, GRILLAGRAL.GRID, GRILLAGRAL.Caption, False, Me)
        End If
   End If
   Set obj = Nothing
       
99  Exit Sub

End Sub

Sub PROCESO_RECUPERO_DE_KG_Y_METROS()
    'USADO PARA CAMBACERES PARA RECUPERAR EL PESO Y LOS METROS
    'CONSIDERANDO LOS LOTES GENERADOS EN EL CIERRE DE OF, DONDE AHI SE INFORMA EL PESO Y LOS METROS
    SQLX$ = "SELECT CODIMOVI,COD_INTE,IDENLOTE,METROS,PESOKG FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE (IdenLote <> '' ) AND ((METROS <= 0 OR METROS IS NULL)"
    SQLX$ = SQLX$ + " OR  (PESOKG <= 0 OR PESOKG IS NULL) )"
    SQLX$ = SQLX$ + " AND (CODIMOVI = 'RT' OR CODIMOVI = 'AJ' "
    SQLX$ = SQLX$ + " OR CODIMOVI = 'CF' OR CODIMOVI = 'OF' or CODIMOVI = 'TR' OR CODIMOVI = 'AJ') AND REFERCOR <> 'Anul.Fabricación Propia'"
    SQLX$ = SQLX$ + "ORDER BY CODIMOVI"
    
    REG& = 0
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
25  On Error Resume Next
    rs.Open (SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0
    '
    With rs
      Do While Not .EOF
          CMOV$ = SAFETEXT$(!CODIMOVI)
          CI1% = XVALO(!cod_inte)
          LOTE$ = SAFETEXT$(!idenlote)

          CONDI$ = "IDENLOTE = '" & LOTE$ & "' AND (CODIMOVI = 'OF' OR CODIMOVI = 'BR') AND (METROS > 0 OR PESOKG > 0)"
          METROS# = XVALO(VALOBADA("MOVISTO", "METROS", CONDI$, "NOMESS"))
          KG1# = XVALO(VALOBADA("MOVISTO", "PESOKG", CONDI$, "NOMESS"))
          !METROS = METROS#
          !PESOKG = KG1#
          .MoveNext
      Loop
    End With
    rs.Close
    Set rs = Nothing

End Sub
Private Sub mnuStTotLuzCamb_Click()

    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "QUILMES") > 0 Then GoTo 100 ' VA AL REPORTE VIEJO
    If InStr(EPAC$, "CAMBACERES") > 0 Then Call ReporteCambaLuzuriaga: Exit Sub
10 OPX% = COMALTER%("Stocks de Bobinas\\Cancelar\Generar Listado\Correr Proceso de Depuraciòn Correctivo")
   If OPX% = 1 Then Exit Sub
   If OPX% = 3 Then
'     Call PROCESO_RECUPERO_DE_KG_Y_METROS_ENCONSUMOS
'     Call PROCESO_RECUPERO_DE_KG_Y_METROS_EN_TRANSFERENCIAS
      Call PROCESO_RECUPERO_DE_KG_Y_METROS
     Call COMUNI("Proceso Terminado")
     GoTo 10
   End If
       '
    Dim MSF As msFlexGrid

    Campos$ = "PRODUCTO/A16" '1
    Campos$ = Campos$ + ";Descsripciòn/A24" '2
    Campos$ = Campos$ + ";Lote/A12" '3
    Campos$ = Campos$ + ";Dep./F5.0" '4
    Campos$ = Campos$ + ";Stock U./F11.0" '5
    
    Campos$ = Campos$ + ";Peso Kg./F11.1" '6
    Campos$ = Campos$ + ";Metros./F11.1" '7
    Campos$ = Campos$ + ";CI/A0" '8

    
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
    J& = 0
    SQLX$ = "SELECT CASE WHEN COD_EXTE IS NULL THEN 'Total' ELSE COD_EXTE END AS COD_EXTE,"
    SQLX$ = SQLX$ + "  IDENLOTE, depomovi,COD_INTE,DESCRIP,"
    SQLX$ = SQLX$ + " FORMAT(SUM(CANTINGRE - CANTSALID), 'N2') AS STOCK,"
    SQLX$ = SQLX$ + " FORMAT(SUM(CASE WHEN Cantingre > 0 THEN pesokg ELSE -pesokg END), 'N2') AS PESO_TOTAL,"
    SQLX$ = SQLX$ + " FORMAT(SUM(CASE WHEN Cantingre > 0 THEN METROS ELSE -METROS END), 'N2') AS METROS_TOTAL"
    SQLX$ = SQLX$ + " FROM MOVISTO WITH(NOLOCK) INNER JOIN MASTER M ON M.CODINT = COD_INTE"
    SQLX$ = SQLX$ + " WHERE M.TIMATE =41 AND M.FAMILIA IN( 'TP','TT','TL','PL' )"
    SQLX$ = SQLX$ + " GROUP BY GROUPING SETS ((COD_INTE,COD_EXTE, IDENLOTE, depomovi,DESCRIP), (COD_INTE), ())"
    SQLX$ = SQLX$ + " HAVING SUM(CANTINGRE - CANTSALID) > 0"
    SQLX$ = SQLX$ + " ORDER BY COD_INTE,COD_EXTE,IDENLOTE"
    'En esta consulta, utilizamos la cláusula GROUPING SETSpara generar un registro adicional con el subtotal
    'para cada cambio en el código ( COD_EXTE). La cláusula GROUPING SETSnos permite especificar múltiples niveles
    'de agrupamiento en una sola consulta.

'se mantuvo la consulta como estaba, pero en vez de tomar los subtotales de la consultas los recalculo cada vez que
'va agregar el registro de Total con lo que voy acumulando
'y al final cargo en el primer registro el total total

    Set rs = READSET(SQLX$)
    
    With rs
    Do While Not .EOF
        CI1% = XVALO(!cod_inte)
        Codigo$ = SAFETEXT$(!Cod_Exte)
        DESXRI$ = SAFETEXT$(!Descrip)
'        '
        J& = J& + 1
        GRILLAGRAL.GRID.Rows = J& + 1
        GRILLAGRAL.GRID.TextMatrix(J&, 1) = Codigo$
        GRILLAGRAL.GRID.TextMatrix(J&, 2) = DESXRI$
        
        GRILLAGRAL.GRID.TextMatrix(J&, 3) = SAFETEXT$(!idenlote)
        GRILLAGRAL.GRID.TextMatrix(J&, 4) = XVALO(!depomovi)
        GRILLAGRAL.GRID.TextMatrix(J&, 5) = FORMATNUM$(XVALO(!STOCK), "F10.0")
        GRILLAGRAL.GRID.TextMatrix(J&, 6) = FORMATNUM$(XVALO(!PESO_TOTAL), "F10.1")
        GRILLAGRAL.GRID.TextMatrix(J&, 7) = FORMATNUM$(XVALO(!METROS_TOTAL), "F10.1")
        GRILLAGRAL.GRID.TextMatrix(J&, 8) = CI1%


        If J& > 1 And TRIM$(UCase$(GRILLAGRAL.GRID.TextMatrix(J&, 1))) <> "TOTAL" Then
           AcumuPesoTotal = AcumuPesoTotal + XVALO(!PESO_TOTAL): AcumuPesoTotalTotal = AcumuPesoTotalTotal + AcumuPesoTotal
           AcumuMetrosTotal = AcumuMetrosTotal + XVALO(!METROS_TOTAL): AcumuMetrosTotalTotal = AcumuMetrosTotalTotal + AcumuMetrosTotal
        End If
        
        If TRIM$(UCase$(GRILLAGRAL.GRID.TextMatrix(J&, 1))) = "TOTAL" And J& > 1 Then
        'si es la de total es por que cambio y le pongo el acumulado que venia sumando y luego incializo la variable
           GRILLAGRAL.GRID.TextMatrix(J&, 6) = FORMATNUM$(XVALO(AcumuPesoTotal), "F10.1")
           GRILLAGRAL.GRID.TextMatrix(J&, 7) = FORMATNUM$(XVALO(AcumuMetrosTotal), "F10.1")
           AcumuPesoTotalTotal = AcumuPesoTotalTotal + AcumuPesoTotal
           AcumuMetrosTotalTotal = AcumuMetrosTotalTotal + AcumuMetrosTotal
           AcumuPesoTotal = 0
           AcumuMetrosTotal = 0
        End If
       .MoveNext
    Loop
   End With
   rs.Close
   Set rs = Nothing
   Screen.MousePointer = 1
   If J& < 1 Then
     Call COMUNI("No hay Registros Según la condición seleccionada")
     Exit Sub
   End If
   'pongo los acumulados totales en la primer fila
   GRILLAGRAL.GRID.TextMatrix(1, 6) = FORMATNUM$(XVALO(AcumuPesoTotalTotal), "F10.1")
   GRILLAGRAL.GRID.TextMatrix(1, 7) = FORMATNUM$(XVALO(AcumuMetrosTotalTotal), "F10.1")

    GRILLAGRAL.GRID.FixedCols = 3

   GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
   GRILLAGRAL.mnuMenu1.Visible = True
   GRILLAGRAL.SinEfectoClickEnGrilla% = 1
   GRILLAGRAL.GRID.ZOrder 0
   GRILLAGRAL.TXTBUSCAR = ""
   TX$ = ""
   If TM% > 0 Then
     TX$ = TX$ + " Material: " & TIMAT$
   End If
   If FMLIA$ <> "" Then
     TX$ = TX$ + " Familia: " & FAMI$
   End If
   
   GRILLAGRAL.Caption = "Stock "
   
   GRILLAGRAL.Show 1
   
   If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
'        Call CARGA_TABLA_IMPRE(GRILLAGRAL.MSF_2, Campos$, "STOCKRGO", "STOCKS DE CÓDIGOS RANGO", "TABLA_IMPRE", RET_NAMECONS$, 1)

        Call CARGA_TABLA_IMPRE(MSF, Campos$, "STCAMLUZ", "Stock Discriminado Cambaceres y Luzuriaga", "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
        Exit Sub
   End If
   Exit Sub
'_________________________________________________________________________________________________________________

100 'SI ES CAMBACERES VUELVE AL REPORTE ANTERIOR
   ''''CODIGO DE RESPALDO ANTERIOR
'
'   'Realizado para Cambaceres muestra el stock a una fecha de cambaceres luzuriaga y totaliza
110 OPX% = COMALTER%("Stocks de Bobinas\\Cancelar\Generar Listado\Correr Proceso de Depuraciòn Correctivo")
   If OPX% = 1 Then Exit Sub
   If OPX% = 3 Then
'     Call PROCESO_RECUPERO_DE_KG_Y_METROS_ENCONSUMOS
'     Call PROCESO_RECUPERO_DE_KG_Y_METROS_EN_TRANSFERENCIAS
      Call PROCESO_RECUPERO_DE_KG_Y_METROS
     Call COMUNI("Proceso Terminado")
     GoTo 110
   End If

'   Call PROCESO_RECUPERO_DE_KG_Y_METROS  'SE AGREGO PARA QU LO GENERE CADA VEZ QUE SACA EL REPORTE
    VDEF$ = ""                           'EL MISMO SE MODIFICO PARA QUE SE GENERE RAPIDAMENTE
    OpcionDepoNoIncluir$ = "": OpcionDepo$ = "0"
    If InStr(EPAC$, "QUILMES") > 0 Then ' SI ES QUILMES
       OPX% = COMALTER%("Stocks de Bobinas\\Cancelar\No Considera Dep. 9 \Considera Solo Dep. 9")
       If OPX% = 2 Then OpcionDepoNoIncluir$ = "3,9": TITULO1$ = "No Considera Dep. 9" ' no incluye dep 3,9
       If OPX% = 3 Then OpcionDepo$ = "9": TITULO1$ = "Considera Solo Dep. 9"                  'el 3 no incluye por defecto por que es el
'       If OPX% = 4 Then OpcionDepo$ = "11": TITULO1$ = "Considera Solo Dep. 11"            'de luzuriaga
       
   Else
       IncluyeONo$ = ""
   End If
   

120  OBX$ = OBJPLA$("ENTRAFECHA1", VDEF$)
    If OBX$ = "" Then
        Des% = 0: Has% = 0: PERIO$ = ""
        Exit Sub
    End If
    VDEF$ = OBX$
    '
    FE% = CVI(Left$(OBX$, 2))
    If FE% <= 30 Then
        GoTo 120
    End If

    FEX = FE%: PERIO$ = FECHATEX$(FEX)


    FECHAX = FETEXCOR1$(FE%)
    '
    Campos$ = "PRODUCTO/A16" '1
    Campos$ = Campos$ + ";Descsripciòn/A24" '2
    Campos$ = Campos$ + ";Camb.U./F11.0" '3
    Campos$ = Campos$ + ";Camb.Mts./F11.0" '4
    Campos$ = Campos$ + ";Camb.Kg./F11.0" '5

    Campos$ = Campos$ + ";Luzur.U./F11.0" '6
    Campos$ = Campos$ + ";Luzur.Mts./F11.0" '7
    Campos$ = Campos$ + ";Luzur.Kg../F11.0" '8

    Campos$ = Campos$ + ";St.Fin.U./F11.0" '9
    Campos$ = Campos$ + ";St.Fin.Mts./F11.0" '10
    Campos$ = Campos$ + ";St.Fin.Kg../F11.0" '11

    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
    J& = 0
    TOTINICIAL = 0
    TOTINGRESO = 0
    TOTEGRESO = 0
    TOTFINAL = 0
    Screen.MousePointer = 11
    TM% = 41
    J& = 0
    TOTUnidadesCambaceres = 0
    TOTMetrosCambaceres = 0
    TOTKgCamaceres = 0

    TOTUnidadesLuzur = 0
    TOTMetrosLuzur = 0
    TOTKgLuzur = 0

    TOTFinalU = 0
    TOTFinalMetros = 0
    TOTFinalKg = 0

    SQLX$ = "select CODINT,CODIGO,Descripcion from master WITH(NOLOCK)"
    SQLX$ = SQLX$ + " where CODINT > 0 AND STATUS >=0"
    If TM% > 0 Then SQLX$ = SQLX$ + " AND TIMATE = " & TM%

    SQLX$ = SQLX$ + " ORDER BY CODIGO"
    Set rs = READSET(SQLX$)
'cambios el 060325 el stock cambaceres va a ser todo con las excepciones previstass
'se cambia en vez de hacer la resta, trae todo el stock menos el del 9,  y 3 como deposito de cambaceres
    With rs
    Do While Not .EOF
        If InStr(EPAC$, "CAMBACERES") > 0 Then OpcionDepoNoIncluir$ = "" 'primero vacio para que calcule aqui en las lineas siguientes los totales (22_10_25)
        CI1% = XVALO(!CODINT)
        Codigo$ = SAFETEXT$(!Codigo)
        DESXRI$ = SAFETEXT$(!Descripcion)
        If InStr(EPAC$, "CAMBACERES") > 0 Then OpcionDepoNoIncluir$ = "3,9" ' LUEGO QUE CALCULO LOS TOTALES LE APLICO LA OpcionDepoNoIncluir$ (22_10_25)

        UnidadesTotales = STODEPFECH(CI1%, XVALO(OpcionDepo$), FE%, , OpcionDepoNoIncluir$)
        MetrosTotales = STODEPFECHENMETROS#(CI1%, XVALO(OpcionDepo$), FE%, , OpcionDepoNoIncluir$)
        KgTotales = STODEPFECHENKG#(CI1%, XVALO(OpcionDepo$), FE%, , OpcionDepoNoIncluir$)
        

        'UnidadesCambaceres = UnidadesTotales - STODEPFECH(CI1%, 3, FE%)  'TOTAL STOCK MENOS STOCK DE CAMBACERES
        UnidadesCambaceres = UnidadesTotales  'TOTAL STOCK MENOS STOCK DE CAMBACERES
        If XVALO(UnidadesCambaceres) > 0 Then
'            MetrosCambaceres = MetrosTotales - STODEPFECHENMETROS#(CI1%, 3, FE%) 'TOTAL metros MENOS metros DE CAMBACERES
            MetrosCambaceres = MetrosTotales
'            KgCamaceres = KgTotales - STODEPFECHENKG#(CI1%, 3, FE%) 'TOTAL kg MENOS kg DE CAMBACERES
            KgCamaceres = KgTotales
        Else
            MetrosCambaceres = 0
            KgCamaceres = 0
        End If

        If OpcionDepo$ > 0 Then  'es que va a considerar depositos 9 u 11 entonces luzuriaga no debe contarlo, carga el deposito seleccionado en columnas cambaceres
           UnidadesLuzur = 0
           MetrosLuzur = 0
           KgLuzur = 0
        Else
            UnidadesLuzur = STODEPFECH(CI1%, 3, FE%) 'STOCK DE LUZURIAGA
            If XVALO(UnidadesLuzur) > 0 Then
                MetrosLuzur = STODEPFECHENMETROS#(CI1%, 3, FE%) 'metros depo luzuriaga
                KgLuzur = STODEPFECHENKG#(CI1%, 3, FE%) 'kg depo luzuriaga
            Else
                MetrosLuzur = 0
                KgLuzur = 0
    
            End If
        End If
        
'        '
'        UnidadesCambaceres = STODEPFECH(CI1%, 2, FE%) 'TOTAL STOCK MENOS STOCK DE CAMBACERES
'        If XVALO(UnidadesCambaceres) > 0 Then
'            MetrosCambaceres = STODEPFECHENMETROS#(CI1%, 2, FE%)  'TOTAL metros MENOS metros DE CAMBACERES
'            KgCamaceres = STODEPFECHENKG#(CI1%, 2, FE%)  'TOTAL kg MENOS kg DE CAMBACERES
'        End If
'
'        UnidadesLuzur = STODEPFECH(CI1%, 3, FE%) 'STOCK DE LUZURIAGA
'        If XVALO(UnidadesLuzur) > 0 Then
'            MetrosLuzur = STODEPFECHENMETROS#(CI1%, 3, FE%) 'metros depo luzuriaga
'            KgLuzur = STODEPFECHENKG#(CI1%, 3, FE%) 'kg depo luzuriaga
'        End If
'
        FinalU = UnidadesTotales: If InStr(EPAC$, "CAMBACERES") > 0 Then FinalU = UnidadesTotales + UnidadesLuzur
        FinalMetros = MetrosTotales: If InStr(EPAC$, "CAMBACERES") > 0 Then FinalMetros = MetrosTotales + MetrosLuzur
        FinalKg = KgTotales: If InStr(EPAC$, "CAMBACERES") > 0 Then FinalKg = KgTotales + KgLuzur
        J& = J& + 1

        GRILLAGRAL.GRID.Rows = J& + 1
        GRILLAGRAL.GRID.TextMatrix(J&, 1) = Codigo$
        GRILLAGRAL.GRID.TextMatrix(J&, 2) = DESXRI$

        GRILLAGRAL.GRID.TextMatrix(J&, 3) = FORMATNUM$(UnidadesCambaceres, "F10.0")
        GRILLAGRAL.GRID.TextMatrix(J&, 4) = FORMATNUM$(MetrosCambaceres, "F10.0")
        GRILLAGRAL.GRID.TextMatrix(J&, 5) = FORMATNUM$(KgCamaceres, "F10.0")

        GRILLAGRAL.GRID.TextMatrix(J&, 6) = FORMATNUM$(UnidadesLuzur, "F10.0")
        GRILLAGRAL.GRID.TextMatrix(J&, 7) = FORMATNUM$(MetrosLuzur, "F10.0")
        GRILLAGRAL.GRID.TextMatrix(J&, 8) = FORMATNUM$(KgLuzur, "F10.0")

        GRILLAGRAL.GRID.TextMatrix(J&, 9) = FORMATNUM$(UnidadesTotales, "F10.0")
        GRILLAGRAL.GRID.TextMatrix(J&, 10) = FORMATNUM$(FinalMetros, "F10.0")
        GRILLAGRAL.GRID.TextMatrix(J&, 11) = FORMATNUM$(FinalKg, "F10.0")

        TOTUnidadesCambaceres = TOTUnidadesCambaceres + XVALO(UnidadesCambaceres)
        TOTMetrosCambaceres = TOTMetrosCambaceres + XVALO(MetrosCambaceres)
        TOTKgCamaceres = TOTKgCamaceres + XVALO(KgCamaceres)

        TOTUnidadesLuzur = TOTUnidadesLuzur + XVALO(UnidadesLuzur)
        TOTMetrosLuzur = TOTMetrosLuzur + XVALO(MetrosLuzur)
        TOTKgLuzur = TOTKgLuzur + XVALO(KgLuzur)

        TOTFinalU = TOTFinalU + XVALO(FinalU)
        TOTFinalMetros = TOTFinalMetros + XVALO(FinalMetros)
        TOTFinalKg = TOTFinalKg + XVALO(FinalKg)

       .MoveNext
    Loop
   End With
   rs.Close
   Set rs = Nothing
   Screen.MousePointer = 1
   'TOTALES
    GRILLAGRAL.GRID.Rows = J& + 3
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 2) = "TOTALES:"
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 3) = FORMATNUM$(XVALO(TOTUnidadesCambaceres), "F10.0")
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 4) = FORMATNUM$(XVALO(TOTMetrosCambaceres), "F10.0")
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 5) = FORMATNUM$(XVALO(TOTKgCamaceres), "F10.0")

    GRILLAGRAL.GRID.TextMatrix(J& + 2, 6) = FORMATNUM$(XVALO(TOTUnidadesLuzur), "F10.0")
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 7) = FORMATNUM$(XVALO(TOTMetrosLuzur), "F10.0")
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 8) = FORMATNUM$(XVALO(TOTKgLuzur), "F10.0")

    GRILLAGRAL.GRID.TextMatrix(J& + 2, 9) = FORMATNUM$(XVALO(TOTFinalU), "F10.0")
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 10) = FORMATNUM$(XVALO(TOTFinalMetros), "F10.0")
    GRILLAGRAL.GRID.TextMatrix(J& + 2, 11) = FORMATNUM$(XVALO(TOTFinalKg), "F10.0")

    '
    GRILLAGRAL.GRID.FixedCols = 3
    Call COLOREAR_GRILLASQL_CadenaDeColumnas(GRILLAGRAL.GRID, &HCDDADA, "6;7;8")
    Call COLOREAR_GRILLA_SOLAFILA(GRILLAGRAL.GRID, &H808080, J& + 2, 0, 1)
   GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
   GRILLAGRAL.mnuMenu1.Visible = True
   GRILLAGRAL.SinEfectoClickEnGrilla% = 1
   GRILLAGRAL.GRID.ZOrder 0
   GRILLAGRAL.TXTBUSCAR = ""
   TX$ = ""
   If TM% > 0 Then
     TX$ = TX$ + " Material: " & TIMAT$
   End If
   If FMLIA$ <> "" Then
     TX$ = TX$ + " Familia: " & FAMI$
   End If

   GRILLAGRAL.Caption = "Stock Cambaceres-Luzuriaga  (" & PERIO$ & ") " & TITULO1$

   GRILLAGRAL.Show 1

   If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
'        Call CARGA_TABLA_IMPRE(GRILLAGRAL.MSF_2, Campos$, "STOCKRGO", "STOCKS DE CÓDIGOS RANGO", "TABLA_IMPRE", RET_NAMECONS$, 1)

        Call CARGA_TABLA_IMPRE(MSF, Campos$, "STCAMLUZ", GRILLAGRAL.Caption, "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
        Exit Sub
   End If
   GoTo 120
End Sub


Private Sub mnuTablaStockSeguridad_Click()

    Dim ta As New DatosAuxiliares
    
    Call ta.CrearTablasAuxiliares
    IDLIST& = ta.ObtenerIdTablaAuxiliar("STOSEGRANGO")
    Call ABM_TABLA(IDLIST&, "", , "campo1")
    
End Sub

Private Sub mnuTrazabilidadXOF_Click()
   CONTRCOM.Show 1
End Sub

Private Sub MovStock_Click()
    Call Command6_Click
End Sub

Private Sub Option1_Click()
   If DERACCE%("VERCOSTO/NOMESS", "Visualizar Costos y Precios") < 1 Then
      Option2.Value = True
   End If
End Sub

 Sub Option2_Click()

End Sub

Private Sub Option3_DblClick()
    Call SELECHO("!INDIPEP/Pedidos Pendientes de Entrega")
    NPED& = XVALO(VALACT1$("!INDIPEP"))
    If NPED& < 1 Then
       Call Option2_Click
       Exit Sub
    End If
    '
    Campos$ = "Código/A20;C.Pedida/F10.1;Entrega/F10.1;Saldo/F10.1"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO, , , ANTOT)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO, , , ANTOT)
    If ANTOT < 5500 Then ANTOT = 5500
    FRMZELECHO.Width = ANTOT + 1500
    '
    'EN ESTA CONSULTA TRAIGO TODO LO QUE SE INGRESO POR LOTE Y SALIO POR LOTE
    FRMZELECHO.MSF1.Rows = 1
    J& = 0
    ACUMLA_SALDO = 0
    SQLX$ = " SELECT CodiExt, CANPED,CantEnt FROM PEDDETA with(nolock)"
    SQLX$ = SQLX$ + "  where NroPed=" & NPED&
    SQLX$ = SQLX$ + "  AND STATUS < 8 "
    SQLX$ = SQLX$ + "  ORDER BY CODIINT"
    Set RS1 = READSET(SQLX$)
    With RS1
        Do While Not .EOF
             Codigo$ = SAFETEXT$(!CODIEXT)
             CANPEDIDA$ = FORMATNUM$(XVALO(!CanPed), "F110.1")
             CANENTREG$ = FORMATNUM$(XVALO(!CANTENT), "F110.1")
             CANSALDOX$ = FORMATNUM$(XVALO(!CanPed) - XVALO(!CANTENT), "F110.1")
             If XVALO(CANSALDOX$) < 1 Then CANSALDOX$ = ""
             J& = J& + 1
             FRMZELECHO.MSF2.Rows = J& + 1
             FRMZELECHO.MSF2.TextMatrix(J&, 1) = Codigo$
             FRMZELECHO.MSF2.TextMatrix(J&, 2) = CANPEDIDA
             FRMZELECHO.MSF2.TextMatrix(J&, 3) = CANENTREG$
             FRMZELECHO.MSF2.TextMatrix(J&, 4) = CANSALDOX$
             
            .MoveNext
        Loop
    End With
    RS1.Close
    Set RS1 = Nothing
    '
    FRMZELECHO.Caption = "Items Pendientes - Pedido N°: " & NPED&
    FRMZELECHO.Show 1
    
    Codigo$ = TRIM$((RESP_ZELE_VECT(1)))
    If Codigo$ <> "" Then
       Text1 = Codigo$
    End If

End Sub

Private Sub OtrAccDirec_Click()
5   RTA% = ALTERNA%("\Alta Numeros de Serie\Listado Numeros de Serie\Salida Numeros de Serie")
    If RTA% = 1 Then
      'ACA LLAMA AL OBJETO NUSERI
      VDEF$ = ""
      VDEF$ = Space(128)
10    HOII% = HOY(HO$)
      Call REPLA(VDEF$, MKI$(HOII%), 3, 2)
      obj$ = OBJPLA$("NUSERIE", VDEF$)
      OBJANT$ = obj$
      If TRIM$(obj$) <> "" Then
        CODI1% = CVI(Mid$(obj$, 1, 2))
        FA1% = CVI(Mid$(obj$, 3, 2))
        ORI1$ = Mid$(obj$, 5, 12)
        NUSERI1$ = Mid$(obj$, 17, 24)
        OBS1$ = Mid$(obj$, 41, 64)
        'VALIDACIONES
        If CODI1% > NUMAS% Or CODI1% <= 0 Then
          Call MENSERR(24, "Código de Articulo no Válido")
          VDEF = OBJANT$
          GoTo 10
        End If
        If TRIM$(ORI1$) = "" Then
          Call MENSERR(24, "Origen no Válido")
          VDEF = OBJANT$
          GoTo 10
        End If
        If TRIM$(NUSERI1$) = "" Then
          Call MENSERR(24, "Número de Serie no Válido")
          VDEF = OBJANT$
          GoTo 10
        End If
        'CODIGO + Nº DE SERIE NO SE PUEDEN REPETIR
        Call FILTERFILE("SERIEQUI", "!SERIEQUI", 1, MKI$(CODI1%))
        For J& = 1 To ULTREG&("!SERIEQUI")
          X$ = REGLEIDO$("!SERIEQUI", J&)
          NUMESERI$ = Mid$(X$, 17, 24)
          If UCase(NUMESERI$) = UCase(NUSERI1$) Then
            Call MENSERR(24, "Código + Número de Serie no Válido o Pre-Existente")
            VDEF = OBJANT$
            GoTo 10
          End If
        Next J&
        'GRABAR EN BASE DE DATOS
        Y$ = REGBLAN$("SERIEQUI")
        Call REPLA(Y$, MKI$(CODI1%), 1, 2)
        Call REPLA(Y$, MKI$(FA1%), 3, 2)
        Call REPLA(Y$, ORI1$, 5, 12)
        Call REPLA(Y$, NUSERI1$, 17, 24)
        Call REPLA(Y$, OBS1$, 123, 64)
        Call REGAPP("SERIEQUI", Y$)
        '
        Call CIERRARCH("SERIEQUI")
        'PREPARO EL OBJETO PARA LA PROXIMA ALTA
        PRONUSERI$ = ""
        NUSE& = XVALO(Right$(TRIM$(NUSERI1$), 1))
        If NUSE& > 0 Then
          PRONUSE& = NUSE& + 1
          If PRONUSE& = 10 Then PRONUSE& = 1
          PRONUSE1$ = FORMATNUM$(PRONUSE&, "F1.0")
          PRONUSERI$ = Mid$(NUSERI1$, 1, Len(TRIM$(NUSERI1$)) - 1) + PRONUSE1$
        End If
        Call REPLA(OBJANT$, "", 5, 105)
        Call REPLA(OBJANT$, PRONUSERI$, 17, 24)
        VDEF$ = OBJANT$
        GoTo 10
        '
      End If
      GoTo 5
    ElseIf RTA% = 2 Then
      Call LINUSE 'ACA LISTADO
      GoTo 5
    ElseIf RTA% = 3 Then
      Call PROCENUSE 'PROCESO DE SALIDA DE NU. SERIE
      GoTo 5
    End If
'\\\OT-05-0711-WAR-FIN///
End Sub
 
Sub PROCENUSE()
      '1) SELECCION DEL CODIGO
30    OBJ3$ = OBJPLA$("PRESCODIGO", VDEF$)
      If OBJ3$ <> "" Then
        Call FRESHECHO("!INDINUSE")
        CI3% = CVI(Mid$(OBJ3$, 1, 2))
        'VALIDACIONES
        If CI3% <= 0 Then
          Call COMUNI("Codigo no Válido o Inexistente.")
          GoTo 30
        End If
        'AVERIGUO SI EXISTEN Nº DE SERIE PARA ESE ARTICULO
        'PENDIENTES DE SALIDA
        Call FILTERFILE("SERIEQUI", "!SERIEQUI", 1, MKI$(CI3%))
        If ULTREG&("!SERIEQUI") <= 0 Then
          Call COMUNI("No existen Numeros de Serie para \ " & CODEXT$(CI3%) & " - " & DESCRIT0(CI3%) & "")
          GoTo 30
        End If
        JJ& = 0
        Call BLANARCH("!INDINUSE")
        For J& = 1 To ULTREG&("SERIEQUI")
          X$ = REGLEIDO$("SERIEQUI", J&)
          CODI1% = CVI(Mid$(X$, 1, 2))
          If CODI1% = CI3% Then
            REMI3$ = Mid$(X$, 75, 24)
            FACT3$ = Mid$(X$, 99, 24)
            If TRIM$(REMI3$) = "" And TRIM$(FACT3$) = "" Then
              FA1% = CVI(Mid$(X$, 3, 2))
              ORI1$ = Mid$(X$, 5, 12)
              NUSE1$ = Mid$(X$, 17, 24)
              REFE1$ = NUSE1$ + "   " + ORI1$ + "   " + FECHATEX$(CVINT(CVDAT(FA1%)))
              'ACA AGREGO EL NUMERO DE SERIE PENDIENTE AL INDICE
              Y$ = REGBLAN$("INDINUSE")
              JJ& = JJ& + 1
              NO% = XVALO(J&) 'NUMERO DE REGISTRO
              Call REPLA(Y$, MKI$(NO%), 1, 2)
              Call REPLA(Y$, REFE1$, 3, 62)
              Call GRAREG("!INDINUSE", Y$, JJ&)
            End If
          End If
        Next J&
        '
        If JJ& <= 0 Then
          Call COMUNI("No existen Numeros de Serie Pendientes para \ " & CODEXT$(CI3%) & " - " & DESCRIT0(CI3%) & "")
          GoTo 30
        End If
        '
        Call SETULTREG("!INDINUSE", JJ&)
        '
32      Call SELECHO("!INDINUSE")
        NREG1& = XVALO(VALACT1$("!INDINUSE"))
        If NREG1& <= 0 Then
          GoTo 30
        End If
        '
35      X$ = REGLEIDO$("SERIEQUI", NREG1&)
        CI4% = CVI(Mid$(X$, 1, 2))
        FA4% = CVI(Mid$(X$, 3, 2))
        ORI4$ = Mid$(X$, 5, 12)
        NUSE4$ = Mid$(X$, 17, 24)
        OBS4A$ = Mid$(X$, 123, 32)
        OBS4B$ = Mid$(X$, 155, 64)
        '
        Call REPLA(VDEF$, MKI$(CI4%), 1, 2)
        Call REPLA(VDEF$, MKI$(FA4%), 3, 2)
        Call REPLA(VDEF$, ORI4$, 5, 12)
        Call REPLA(VDEF$, NUSE4$, 17, 24)
        Call REPLA(VDEF$, OBS4A$, 41, 32)
        Call REPLA(VDEF$, OBS4B$, 73, 32)
        '
        HOII% = HOY(HO$)
        Call REPLA(VDEF$, MKI$(HOII%), 105, 2)
        '
37      OBJ4$ = OBJPLA$("SANUSE", VDEF$)
        OBJANT4$ = OBJ4$
        If OBJ4$ <> "" Then
          'VALIDACIONES
          REMI4$ = Mid$(OBJ4$, 109, 24)
          FACT4$ = Mid$(OBJ4$, 133, 24)
          If TRIM$(REMI4$) = "" Or TRIM$(FACT4$) = "" Then
            Call COMUNI("Número de Remito o Factura no Valido.")
            VDEF$ = OBJANT4$
            GoTo 37
          End If
          NC4% = CVI(Mid$(OBJ4$, 107, 2))
          If NC4% <= 0 Then
            Call COMUNI("Cliente / Destino no Válido.")
            VDEF$ = OBJANT4$
            GoTo 37
          End If
          'EDITA EL REGISTRO
          FS4% = CVI(Mid$(OBJ4$, 105, 2))
          '
          X$ = REGLEIDO$("SERIEQUI", NREG1&)
          Call REPLA(X$, MKI$(FS4%), 41, 2)
          Call REPLA(X$, MKI$(NC4%), 43, 2)
          Call REPLA(X$, rasocli$(NC4%), 45, 30)
          Call REPLA(X$, REMI4$, 75, 24)
          Call REPLA(X$, FACT4$, 99, 24)
          Call GRAREG("SERIEQUI", X$, NREG1&)
          '
        Else
          VDEF$ = ""
          GoTo 32
        End If
        '
      End If
End Sub

Sub LINUSE()
    'ACA LLAMA AL OBJETO DESHASCOD2
      VDEF$ = ""
20    HOII% = HOY(HO$)
      ULTCO% = NUMAS%
      Call REPLA(VDEF$, MKI$(1), 1, 2)
      Call REPLA(VDEF$, MKI$(ULTCO%), 3, 2)
      Call REPLA(VDEF$, Chr$(0), 5, 1)
      Call REPLA(VDEF$, MKI$(COMESACT%), 6, 2)
      Call REPLA(VDEF$, MKI$(HOII%), 8, 2)
      OBJ1$ = OBJPLA$("DESHASCOD2", VDEF$)
      If TRIM$(OBJ1$) <> "" Then
        DESA% = CVI(Mid$(OBJ1$, 1, 2))
        HASA% = CVI(Mid$(OBJ1$, 3, 2))
        DESF% = CVI(Mid$(OBJ1$, 6, 2))
        HASF% = CVI(Mid$(OBJ1$, 8, 2))
        TIP% = Asc%(Mid$(OBJ1$, 5, 1))
        '
        Call COPYSTRU("SERIEQUI", "LINUSERI")
        Z$ = REGBLAN$("LINUSERI")
        '
        JJ& = 0
        For J& = 1 To ULTREG&("SERIEQUI")
          X$ = REGLEIDO$("SERIEQUI", J&)
          CI1% = CVI(Mid$(X$, 1, 2))
          FA1% = CVI(Mid$(X$, 3, 2))
          If TIP% > 0 Then
            If TIP% <> TIMATE%(CI1%) Then
              GoTo 25
            End If
          End If
          If CI1% >= DESA% Then
            If CI1% <= HASA% Then
              If FA1% >= DESF% Then
                If FA1% <= HASF% Then
                  ORI1$ = Mid$(X$, 5, 12)
                  NUSE1$ = Mid$(X$, 17, 24)
                  FESA1% = CVI(Mid$(X$, 41, 2)) 'FECHA DE SALIDA
                  CTE1% = CVI(Mid$(X$, 43, 2)) 'CLIENTE
                  DESTI1$ = Mid$(X$, 45, 30) 'DESTINO
                  REMI1$ = Mid$(X$, 75, 24) 'REMITO
                  FACT1$ = Mid$(X$, 99, 24) 'FACTURA
                  OBS1$ = Mid$(X$, 123, 64) 'OBSERVACIONES
                  Call REPLA(Y$, MKI$(CI1%), 1, 2)
                  Call REPLA(Y$, MKI$(FA1%), 3, 2)
                  Call REPLA(Y$, ORI1$, 5, 12)
                  Call REPLA(Y$, NUSE1$, 17, 24)
                  Call REPLA(Y$, MKI$(FESA1%), 41, 2)
                  Call REPLA(Y$, MKI$(CTE1%), 43, 2)
                  Call REPLA(Y$, DESTI1$, 45, 30)
                  Call REPLA(Y$, REMI1$, 75, 24)
                  Call REPLA(Y$, FACT1$, 99, 24)
                  Call REPLA(Y$, OBS1$, 123, 64)
                  JJ& = JJ& + 1
                  Call GRAREG("LINUSERI", Y$, JJ&)
                End If
              End If
            End If
          End If
25      Next J&
        Call SETULTREG("LINUSERI", JJ&)
        Call CIERRARCH("LINUSERI")
        If JJ& > 0 Then
          Call FILESORT("LINUSERI", "LINUSERI", 3, 3, "ASC")
          Call FILESORT("LINUSERI", "LINUSERI", 1, 1, "ASC")
          Call FINAL("*LINUSERI")
        Else
          Call COMUNI("No existen Numeros de Serie Asignados")
        End If
      End If
End Sub
'\\\OT-05-0711-WAR-FIN///
Private Sub ParFabr_Click()
    Call QUERSTK07.Command3_Click
End Sub

Private Sub salir_Click()
    Call Command1_Click
End Sub

Private Sub SitCob_Click()
    Call Command4_Click
End Sub


Private Sub soloxproveedor_Click()

    CODI$ = TRIM$(Text1)
    CIII% = CODINT%(CODI$)
    '
    If CIII% < 1 Then
       Call COMUNI("Código Inexistentes o No Válido")
       Exit Sub
    End If
    '
    Call SELECHO("PROVED1")
    PROVED = XVALO(VALACT1$("PROVED1"))
    If PROVED > 0 Then
       Call CARUBICSQL(CIII%, PROVED)
    End If

End Sub

Private Sub stockdetallexproveedor_Click()
    CODXX$ = Text1.TEXT
    CODII% = CODINT%(CODXX$)
    If CODII < 1 Then
      Call COMUNI("Código Inexistente o No Válido")
      Exit Sub
    End If
    '
    Campos$ = "Cta.Prov/F9;R.Social/A40;O.Stock/F10.1;Ubicación/A48"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO, , , ANTOT)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO, , , ANTOT)
    If ANTOT < 5500 Then ANTOT = 5500
    FRMZELECHO.Width = ANTOT + 1500
    '
    'EN ESTA CONSULTA TRAIGO TODO LO QUE SE INGRESO POR LOTE Y SALIO POR LOTE
    FRMZELECHO.MSF1.Rows = 1
    J& = 0
    ACUMLA_SALDO = 0
    SQLX$ = " SELECT B.COD_EXTE, B.NUMPROV, M.COD_INTE, SUM(CANTSALID) AS SALEN, SUM(CANTINGRE) AS INGRESO FROM MOVISTO M with(nolock)"
    SQLX$ = SQLX$ + " INNER JOIN BOLRECEP B ON M.IDENLOTE = B.IDENLOTE"
    SQLX$ = SQLX$ + " WHERE B.COD_INTE = " & CODII%
    SQLX$ = SQLX$ + " AND M.COD_INTE = " & CODII%
    SQLX$ = SQLX$ + " GROUP BY M.Cod_Inte, B.NumProv,B.COD_EXTE"
    'SQLX$ = SQLX$ + " ORDER BY M.FECHMOV"
    Set RS1 = READSET(SQLX$)
    With RS1
        Do While Not .EOF
             NPROV = XVALO(!NumProv)
             RAZONPROVEEDOR$ = rasocli$(-1 * NPROV)
             SALDOX = XVALO(!ingreso) - XVALO(!SALEN)
             '
             J& = J& + 1
             FRMZELECHO.MSF2.Rows = J& + 1
             FRMZELECHO.MSF2.TextMatrix(J&, 1) = NPROV
             FRMZELECHO.MSF2.TextMatrix(J&, 2) = RAZONPROVEEDOR$
             FRMZELECHO.MSF2.TextMatrix(J&, 3) = FORMATNUM$((XVALO(SALDOX)), "F12.1")
             If NPROV > 0 Then
                FRMZELECHO.MSF2.TextMatrix(J&, 4) = UBISTOSQL$(CODII%, NPROV)
             Else
                FRMZELECHO.MSF2.TextMatrix(J&, 4) = UBISTO$(CODII%)
             End If
             
             ACUMLA_SALDO = ACUMLA_SALDO + SALDOX
            .MoveNext
        Loop
    End With
    RS1.Close
    Set RS1 = Nothing
    '
    'AGREGO AL FINAL EL ACUMULADO DE LOS MOVIMIENTOS DEL ITEM QUE SE GENERARON SIN LOTE POR EJEMPLO FACTURAS GENERADAS SIN LOTE.
    SQLX$ = " SELECT COD_INTE, SUM(CANTSALID) AS SALEN, SUM(CANTINGRE) AS INGRESO FROM MOVISTO With(nolock)"
    SQLX$ = SQLX$ + " WHERE (IDENLOTE IS NULL OR IDENLOTE = '') "
    SQLX$ = SQLX$ + " AND COD_INTE = " & CODII%
    SQLX$ = SQLX$ + " GROUP BY Cod_Inte"
    Set RS1 = READSET(SQLX$)
    With RS1
        Do While Not .EOF
             SALDOX = XVALO(!ingreso) - XVALO(!SALEN)
             '
             J& = J& + 1
             FRMZELECHO.MSF2.Rows = J& + 1
             FRMZELECHO.MSF2.TextMatrix(J&, 1) = "0"
             FRMZELECHO.MSF2.TextMatrix(J&, 2) = "No Identificado"
             FRMZELECHO.MSF2.TextMatrix(J&, 3) = FORMATNUM$((XVALO(SALDOX)), "F12.1")
             '
             ACUMLA_SALDO = ACUMLA_SALDO + SALDOX
            .MoveNext
        Loop
    End With
    RS1.Close
    Set RS1 = Nothing
    FRMZELECHO.Caption = "Código: " & CODXX$ & "     -     Saldo: " & ACUMLA_SALDO
    FRMZELECHO.Show 1
    
    PROVED = XVALO(RESP_ZELE_VECT(1))
    If PROVED > 0 Then Call VERUBICA(CODII%, PROVED)
    

End Sub

Private Sub Timer2_Timer()
    Static TextA$, FUEAMOSTRAR%
    '
    If Text1 = TextA$ Then Exit Sub     ' NO CAMBIO EL CODIGO DE ARTICULO
    If FUEAMOSTRAR% > 0 Then Exit Sub   ' ESTÁ PRESENTANDO EL ANTERIOR
    '
    TextA$ = Text1
    Text2.TEXT = ""
    Text3.TEXT = ""
    Label101.Caption = ""
    Label102.Caption = ""
    Label103.Caption = ""
    Label201.Caption = ""
    Label202.Caption = ""
    Label203.Caption = ""
    Label301.Caption = ""
    Label302.Caption = ""
    Label303.Caption = ""
    Label401.Caption = ""
    Label404.Caption = ""
    Label402.Caption = ""
    Label403.Caption = ""
    List1.Clear
    gridStockProyectado.Rows = 1
    LIUBIC.Clear
    COD$ = TRIM$(Text1.TEXT)
    CIX1% = CODINT%(COD$)
    If CIX1 > 0 Then
        Text2.TEXT = Unimed$(CIX1%)
        Text3.TEXT = DESCRIT$(CIX1%)
        Openforms = DoEvents
        FUEAMOSTRAR% = 1
        Call MUEDIS(CIX1%)                ' recalcular existencias
        FUEAMOSTRAR% = 0
    End If
    '
End Sub

Private Sub Text1_Change()
    '
    ' REEMPLAZADO EL 17/10/2008
    ' EL SUCESO CHANGE DEL TEXT1 DISPATA EL SUCESO TIMER CON INTERVAL 0 0,5 SEG.
    ' Y ESTE A SU VEZ LA PRESENTACION DE LOS DATOS.
    '
    ' SE DETECTÓ EN DOSIVAC QUE SI SE INGRESABA EL CÓDIGO POR TECLADO, AL
    ' LLEGAR AL PRIMER CODIGO QUE MACHEABA COMENZABA A PRESENTAR. sI LUEGO
    ' SE SEGÍAN INGRESANDO CARACTERES HASTA LLEGAR AL SEGUNDO CÓDIGO, PRO-
    ' DUCIA UNA MEZCLA DE DATOS ENTRE LOS DOS, CON DATOS FALTANTES Y REPETICIONES.
    '
    ' LLEVADO AL SUCESO TIMER, ESPERA HASTA QUE HAYA COMPLETADO LA PRESEN-
    ' TACION DEL PRIMER ARTICULO (BANDERA FUEAMOSTRAR%) Y RECÉN DESPUES
    ' DISPARA LA PRESENTACION DEL SEGUNDO ARTICULO.
    '
    ' CASO DE ESTUDIO LOS CODIGOS 19032 y 19032/1
    '
End Sub
'

Private Sub Text1_DblClick()
    '
    ASEL$ = "MASTER"
    If Option1.Value = True Then
       ASEL$ = "LIPRE001"
    End If
    '
    Text1.SelLength = 0
    Call SELECHO(ASEL$)
    Text1.SelLength = 0
    CODD$ = VALACT1$(ASEL$)
    If CODD$ = "" Then
        Exit Sub
    End If
    COD$ = CODD$
    Text1.TEXT = COD$
    CI% = CODINT%(COD$)
    Text3.TEXT = DESCRIT$(CI%)
    Text2.TEXT = Unimed$(CI%)
    '
End Sub
'

Private Sub Text1_LostFocus()
    COD$ = TRIM$(Text1.TEXT)
    CI% = CODINT%(COD$)
    Text3.TEXT = DESCRIT$(CI%)
    Text2.TEXT = Unimed$(CI%)
End Sub

Private Sub Text2_gotfocus()
    Text1.SetFocus
End Sub

Private Sub Text3_Click()
    Text1.SetFocus
End Sub

Sub SHOWPEDIPEN()
     '
     Screen.MousePointer = 11
     Call ABREPEDCLI
     SQLX$ = " SELECT DISTINCT PEDENCA.NroPed, PEDENCA.NroClie, PEDENCA.FechEmis, PEDENCA.NroOcom "
     SQLX$ = SQLX$ + " FROM PEDENCA "
     SQLX$ = SQLX$ + " INNER JOIN  PEDDETA "
     SQLX$ = SQLX$ + " ON (PEDENCA.NroPed = PEDDETA.NroPed) "
     SQLX$ = SQLX$ + " WHERE PEDDETA.Status < 8 "   ' para que no este anulada ni cumplida
     SQLX$ = SQLX$ + " AND PEDDETA.CantEnt < (PEDDETA.CanPed - 0.005) "
     SQLX$ = SQLX$ + " AND PEDENCA.Status < 9 "
     SQLX$ = SQLX$ + " ORDER BY PEDENCA.NroPed ASC "
     'Set PediModifTemp = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     Dim PediModifTemp As ADODB.Recordset
     Set PediModifTemp = New ADODB.Recordset
     PediModifTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockReadOnly, adCmdText
     JJ& = 0
     With PediModifTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
             NRODES& = !NROPED
             TTXX$ = MKS$(NRODES&)
             NOCO$ = !NROOCOM
           If NOCO$ <> "" Then
             TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
           End If
             FEX = CVINT(!FECHEMIS)
             TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
             NCLIE% = !NROCLIE
             TTXX$ = TTXX$ + TRIM$(rasocli$(NCLIE%)) + Space$(52)
             JJ& = JJ& + 1
           Call GRAREG("INDIPEP", Left$(TTXX$, 52) + MKS$(REGENCA&), JJ&)
         .MoveNext
         Loop
       End If
     End With
     PediModifTemp.Close
     Set PediModifTemp = Nothing
     
     Call SETULTREG("INDIPEP", JJ&)
     Call CIERRARCH("INDIPEP")
     '
     Screen.MousePointer = 1
     YAINDI% = 1
     '
2112 If ULTREG&("INDIPEP") < 1 Then
        Call COMUNI("No hay O/V's Pendientes\que Puedan Consultarse")
        GoTo 9999
     End If
     '
2110 X$ = REGSEL$("INDIPEP/Ordenes de Venta Pendientes de Entrega")
     If Len(X$) <= 4 Then
        GoTo 9999
     End If
     '
     NPX& = CVS(Left$(X$, 4))
     If NPX& < 1 Then GoTo 9999
     '
     TIDOCUM$ = "Orden de Despacho"
     NUDOCU$ = TRIM$(Str$(NPX&))
     While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
     NUDOCU$ = NUDOCU$ + " "
     Screen.MousePointer = 11
     '
     For UI& = ULTREG&("PDOCLST") To 1 Step -1
       XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
       If InStr(XX$, TIDOCUM$) > 4 Then
         PPXX% = InStr(XX$, "Nro.")
         If PPXX% > 0 Then
           XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
         End If
         If InStr(XX$, NUDOCU$) > 4 Then
           DOCUNU& = CVS(Left$(XX$, 4))
           GoTo 5
         End If
       End If
     Next UI&
     Screen.MousePointer = 1
     Call COMUNI("Documento no Encontrado")
     GoTo 2110
     '
5    For KKI% = 0 To 5
       EDITFORM.Picture1(KKI%).Cls
       EDITFORM.Picture1(KKI%).DrawWidth = 3
       EDITFORM.Picture1(KKI%).DrawStyle = 0
       EDITFORM.Picture1(KKI%).FillStyle = 1
     Next KKI%
     '
     REDOCU$ = REGACT$("PDOCLST")
     URE& = ULTREG&("PDOCSPL")
     LI& = 0: LS& = URE&
6    E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
     L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
     DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
     If DCC& < DOCUNU& Then LI& = L&: GoTo 6
     If DCC& > DOCUNU& Then LS& = L&: GoTo 6
     GoTo 8
     '
7    Screen.MousePointer = 1
     Call MENSERR(24, "Documento no Registrado")
     GoTo 9999
     '
8    TPSP$ = ""
     Screen.MousePointer = 11
     While L& > 1
       DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
       If DCC& < DOCUNU& Then GoTo 9
       L& = L& - 1
     Wend
     '
9    CAPAGINAS% = 0
     For KKL& = L& To URE&
       TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
       If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
       If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
         TPSP$ = TPSP$ + Mid$(TBUF$, 5)
       End If
     Next KKL&
     '
95   SPOOLSTRING$ = TPSP$
     For KKI% = 0 To 5
       EDITFORM.Picture1(KKI%).Cls
       EDITFORM.Picture1(KKI%).Top = 0
       EDITFORM.Picture1(KKI%).Refresh
       EDITFORM.Picture1(KKI%).Height = 7170
       EDITFORM.Picture1(KKI%).Width = 11700
       EDITFORM.Picture1(KKI%).DrawWidth = 3
       EDITFORM.Picture1(KKI%).DrawStyle = 0
       EDITFORM.Picture1(KKI%).FillStyle = 1
     Next KKI%
     ALTUPAGINA = 7170
     TOPEPAGINA = 0
     PAGINACTIVA% = 0
     CAPAGINAS% = 0
     '
10   PPXX% = InStr(TPSP$, Chr$(255))
     If PPXX% > 0 Then
       COPRIAT$ = Left$(TPSP$, PPXX% - 1)
       TPSP$ = Mid$(TPSP$, PPXX% + 1)
       Call TRACE(20, FIN& - Len(TPSP$), FIN&)
       '
       If Len(COPRIAT$) > 0 Then
         COPRI% = Asc(Left$(COPRIAT$, 1))
         ATRIP$ = Mid$(COPRIAT$, 2)
         Call SHOWCOMMAND(COPRI%, ATRIP$)
       End If
       GoTo 10
       '
     End If
     '
     Screen.MousePointer = 1
     PAGINACTIVA% = 0
     For KKI% = 0 To 5
       EDITFORM.Picture1(KKI%).Visible = False
     Next KKI%
     EDITFORM.Picture1(PAGINACTIVA%).Visible = True
     EDITFORM.Picture1(PAGINACTIVA%).Refresh
     EDITFORM.VScroll1.min = 0
     EDITFORM.VScroll1.Value = 0
     EDITFORM.VScroll1.max = 1
     If ALTUPAGINA > EDITFORM.Frame1.Height Then
       EDITFORM.VScroll1.max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
     End If
     EDITFORM.Command3.Enabled = False
     'EDITFORM.Command5.Enabled = False
     EDITFORM.Show 1
     '
     GoTo 2110
     '
9999 End Sub

Sub MUEDIS(CI%)
     '
     Static CTX%, VECOPEN$, CMPED$, HOYY, MODODOSI%
     '
     EPAC$ = TRIM$(UCase$(EMPREAC$))
     CIXI% = CODINT(Text1)
     CADECI% = XVALO(CNTDC$(CIXI%))

     Screen.MousePointer = 11
     If CTX% <> 1 Then
        IDENSEDE$ = Control$("", "IDENSEDE")
        VECOPEN$ = PUNOCPEN$
        CMPED$ = "PC"
        CMO$ = TRIM$(Control$("", "CODPECLI"))
        If CMO$ <> "" Then CMPED$ = CMO$
        HOYY = HOY(HOS$)
        If InStr(EPAC$, "DOSIVA") > 0 Then MODODOSI% = 1
        CTX% = 1
     End If
     '
     List1.Clear
     STOPROP = 0
     STOREMO = 0
     TOPEDID = 0
     TOSUMIN = 0
     TOFABRI = 0
     TORESER = 0
     STOMINI = STOMIN(CI%)
     LECONOM = LOTESTAN(CI%)
     GoSub PRESEDIS
     '
     Call CARDEPOS
     '
     Screen.MousePointer = 11
     Call VESARRAS(CI%)
     '
     Screen.MousePointer = 11
     STOPROP = 0
     If CI% > 0 Then
       For U& = 1 To NUSU%
         CAPAR = STODEPOS(CI%, COSU%(U&)) '        XXZ$ = Mid$(SALCA$, 4 * U& - 3, 4)
         If Abs(CAPAR) >= 0.05 Then
           XXX$ = ADEPO$(U&) + CSTRING$(MKS$(CAPAR), 4, 11, CADECI%, 2)
           If DISTO(U&) < 0 Then XXX$ = XXX$ + "(x)"
           XXX$ = AJUSTI$(XXX$, 60) + FORMATNUM$(COSU%(U&), "F4.0")
           List1.AddItem XXX$
           If DISTO%(U&) >= 0 Then
             STOPROP = STOPROP + CAPAR
           End If
         End If
       Next U&
       GoSub PRESEDIS
     End If
     '
10   'List2.Clear
     ' *** MSFLEXGRID
     gridStockProyectado.Rows = 1
     
     HOII% = HOY(HO$)
'     LI2TEX$ = Space$(96)
'     Call REPLA(LI2TEX$, "Stock Propio al " + HO$, 15, 29)
'     Call REPLA(LI2TEX$, CSTRING$(MKS$(STOPROP), 4, 9, CADECI%, 2), 44, 9)
'     Call REPLA(LI2TEX$, HO$, 74, 9)
'     OBSER$ = ""
'     If STOPROP < 0 Then
'          OBSER$ = "!!! Negativo !!!"
'        ElseIf STOPROP < STOMINI Then
'          OBSER$ = "Sub-Mínimo"
'     End If
'     Call REPLA(LI2TEX$, OBSER$, 84, 24)
'     List2.AddItem LI2TEX$
     
     ' *** COMPLETA MSFLEXGRID STOCK PROYECTADO
     FilaGrid& = 1
     With gridStockProyectado
        .Rows = .Rows + 1
        .TextMatrix(FilaGrid&, 1) = "Stock Propio al " + HO$
        .TextMatrix(FilaGrid&, 2) = TRIM$(FORMATNUM$(STOPROP, "F12." & CStr(CADECI%)))
        .TextMatrix(FilaGrid&, 5) = HO$
        OBSER$ = ""
        If STOPROP < 0 Then
             OBSER$ = "!!! Negativo !!!"
           ElseIf STOPROP < STOMINI Then
             OBSER$ = "Sub-Mínimo"
        End If
        .TextMatrix(FilaGrid&, 6) = OBSER$
        .TextMatrix(FilaGrid&, 7) = Format("01/01/1980", "yyyymmdd")
     End With
     
     SALSTO = STOTOT
     '
12   TORECEP = 0
     '
     Call ABRESTOCKS
     SQLX$ = " SELECT CANTNOM, NumProv, NUMEBORE , FECHRECEP, NumRem_Prov, CanToApro  FROM BOLRECEP WITH(NOLOCK) "
     SQLX$ = SQLX$ + " WHERE Status = 0 "
     SQLX$ = SQLX$ + " AND Cod_Inte = " & CI% & " "
     SQLX$ = SQLX$ + " ORDER BY NumeBoRe "
     '
     'Set COBRECEPTMP = Stocks.OpenRecordset(sqlx$, dbOpenSnapshot)
     '
'     Dim COBRECEPTMP As ADODB.Recordset
'     Set COBRECEPTMP = New ADODB.Recordset
'     COBRECEPTMP.CursorType = adOpenDynamic
'     COBRECEPTMP.LockType = adLockReadOnly
'     COBRECEPTMP.Open FILTSQL$(SQLX$), FLEXCONN, , , adcmtext
     Set COBRECEPTMP = READSET(SQLX$)
     With COBRECEPTMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         JJ& = 0
         k& = 0
         Do While (.EOF = False)
            CARECEP = XVALO(!CantNom) - XVALO(!CanToApro)
            If CARECEP > 0 Then
              NCI% = XVALO(!NumProv)
              REMIPRO$ = SAFETEXT$(!NumRem_Prov)
              NROBOL& = XVALO(!NUMEBORE)
              FEX = CVINT(!FechRecep)
'              LI2TEX$ = Space$(96)
'              Call REPLA(LI2TEX$, "BR." + TRIM$(CSTRING$(MKS$(NROBOL&), 3, 6, 0, 2)), 1, 9)
'              Call REPLA(LI2TEX$, RASOCLI$((-1) * NCI%), 15, 29)
'              Call REPLA(LI2TEX$, CSTRING$(MKS$(CARECEP), 4, 9, CADECI%, 2), 44, 9)
'              Call REPLA(LI2TEX$, FECHATEX$(FEX), 74, 9)
'              Call REPLA(LI2TEX$, "RT." + REMIPRO$, 84, 24)
'              List2.AddItem LI2TEX$
              
              ' *** COMPLETA MSFLEXGRID STOCK PROYECTADO
              FilaGrid& = FilaGrid& + 1
              With gridStockProyectado
                 .Rows = .Rows + 1
                 .TextMatrix(FilaGrid&, 0) = "BR." + TRIM$(CSTRING$(MKS$(NROBOL&), 3, 6, 0, 2))
                 .TextMatrix(FilaGrid&, 1) = rasocli$((-1) * NCI%)
                 .TextMatrix(FilaGrid&, 2) = CSTRING$(MKS$(CARECEP), 4, 9, CADECI%, 2)
                 .TextMatrix(FilaGrid&, 5) = FECHATEX$(FEX)
                 .TextMatrix(FilaGrid&, 6) = "RT." + REMIPRO$
                 .TextMatrix(FilaGrid&, 7) = Format(FECHATEXLAR$(FEX), "yyyymmdd")
              End With
              
              TOSUMIN = TOSUMIN + CARECEP
              TORECEP = TORECEP + CARECEP
            End If
         .MoveNext
         Loop
       End If
       On Error GoTo 0
     End With
     '
     ' *** RECEPCIONES PENDIENTES DE APROBACION (NO ES LA INSPECCION Y APROBACION COMUN)
     ' *** LAS BR NO APROBADAS NO EXISTEN MAS COMO BR PENDIENTES Y TAMPOCO INGRESARON AL STOCK
     ' *** ES POR ESO QUE SE INCLUYE EL CODIGO SIGUIENTE QUE LAS AGREGA A LA DISPONIBILIDAD
     If InStr(1, EMPREAC$, "ARCON") > 0 Then
        Call COPYSTRU("BRECEPEN", "BRECEPEL")
        Call FILESORT("BRECEPEN", "BRECEPEL", 1, 1)
        FIN& = ULTREG&("BRECEPEL")
        For U& = 1 To FIN&
          TXTX$ = REGLEIDO$("BRECEPEL", U&)
          CIII% = CVI(Mid$(TXTX$, 3, 2))
          If CIII% = CI% Then
               DOPRI$ = TRIM$(Mid$(TXTX$, 5, 16))
               CANTINGRE = CVD(Mid$(TXTX$, 88, 8))
               CANTSALID = CVD(Mid$(TXTX$, 96, 8))
               CANTI = CANTINGRE - CANTSALID
               If CANTI > 0.05 Then
                    FEX = CVI(Mid$(TXTX$, 1, 2))
                    REFE$ = Mid$(TXTX$, 43, 30)
                    NCI% = CVI(Mid$(TXTX$, 86, 2))
                    DOPRI$ = Mid$(DOPRI$, 1, InStr(4, DOPRI$, "-") - 1)
                    DOPRI$ = REPLACAR$(DOPRI$, "-", ".")
                    ' *** COMPLETA MSFLEXGRID STOCK PROYECTADO
                    FilaGrid& = FilaGrid& + 1
                    
                    With gridStockProyectado
                       .Rows = .Rows + 1
                       .TextMatrix(FilaGrid&, 0) = DOPRI$
                       .TextMatrix(FilaGrid&, 1) = rasocli$((-1) * NCI%)
                       .TextMatrix(FilaGrid&, 2) = CSTRING$(MKS$(CANTI), 4, 9, CADECI%, 2)
                       .TextMatrix(FilaGrid&, 5) = FECHATEX$(FEX)
                       .TextMatrix(FilaGrid&, 6) = "Pend. de Aprobación"
                       .TextMatrix(FilaGrid&, 7) = "Pend. de Aprobación"
                       
                    End With
                    TOSUMIN = TOSUMIN + CANTI
                    TORECEP = TORECEP + CANTI
               End If
          End If
        Next U&
     End If
     
     If TORECEP > 0.05 Then
        XXX$ = "RECEPC. " + CSTRING$(MKS$(TORECEP), 4, 11, CADECI%, 2)
        List1.AddItem XXX$
     End If
     GoSub PRESEDIS
     '
     Call ABREPEDCLI
     SQLX$ = " SELECT NROPED, NroClie, CanPed,  CantEnt, FechEmis, AcuDespa, AcuPrepa, FeSolEnt FROM PEDDETA WITH(NOLOCK) "
     SQLX$ = SQLX$ + "WHERE CodiInt = " & CI% & " "
     SQLX$ = SQLX$ + "AND VenRep = 'V' "
     SQLX$ = SQLX$ + "AND Status < 8 "   ' para que no este anulada ni cumplida
     If MODODOSI% = 1 Then
          SQLX$ = SQLX$ + " AND StaDespa < 2 "
          SQLX$ = SQLX$ + "AND AcuDespa < (CanPed - 0.005) "
        Else
          SQLX$ = SQLX$ + "AND CantEnt < (CanPed - 0.005) "
     End If
     SQLX$ = SQLX$ + "ORDER BY NroPed ASC "
     '
'     Dim PEDTEMP As ADODB.Recordset
'     Set PEDTEMP = New ADODB.Recordset
'     PEDTEMP.CursorType = adOpenDynamic
'     PEDTEMP.LockType = adLockReadOnly
'     PEDTEMP.Open FILTSQL$(SQLX$), FLEXCONN, , , adcmtext
     Set PEDTEMP = READSET(SQLX$)
     With PEDTEMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           NROX& = !NROPED
           NCI% = !NROCLIE
           CAPEN = !CanPed - !CANTENT
           
           SALPREP& = 0
           If MODODOSI% = 1 Then
              FEMIPED% = CVINT(!FECHEMIS)
              If FEMIPED% >= FECHANUM("07/11/05") Then
                 CAPEN = XVALO(!CanPed) - XVALO(!AcuDespa)
                 SALPREP& = XVALO(!AcuPrepa) - XVALO(!AcuDespa)
              End If
           End If
           FEX = CVINT(!FeSolEnt)
           If CAPEN < 0 Then CAPEN = 0
           '
           If CAPEN > 0 Then
             TOPEDID = TOPEDID + CAPEN
'             LI2TEX$ = Space$(96)
'             Call REPLA(LI2TEX$, CMPED$ + "." + TRIM$(CSTRING$(MKS$(NROX&), 3, 6, 0, 2)), 1, 9)
'             Call REPLA(LI2TEX$, RASOCLI$(NCI%), 15, 29)
'             Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN), 4, 9, CADECI%, 2), 53, 9)
'             Call REPLA(LI2TEX$, FECHATEX$(FEX), 74, 9)
'               OBSER$ = ""
'               If FEX < HOYY Then
'                 ATRAS% = DIENTRE%(FEX, HOYY)
'                 OBSER$ = "Atr." + TRIM$(Str$(ATRAS%)) + " Ds."
'               End If
'               If SALPREP > 0 Then
'                 If OBSER$ <> "" Then OBSER$ = OBSER$ + "/ "
'                 OBSER$ = OBSER$ + "Pr." + TRIM$(Str$(SALPREP&))
'               End If
'             Call REPLA(LI2TEX$, OBSER$, 84, 24)
'             List2.AddItem LI2TEX$
             
             ' *** COMPLETA MSFLEXGRID STOCK PROYECTADO
             FilaGrid& = FilaGrid& + 1
             With gridStockProyectado
                .Rows = .Rows + 1
                .TextMatrix(FilaGrid&, 0) = CMPED$ + "." + TRIM$(CSTRING$(MKS$(NROX&), 3, 6, 0, 2))
                .TextMatrix(FilaGrid&, 1) = rasocli$(NCI%)
                .TextMatrix(FilaGrid&, 3) = CSTRING$(MKS$(CAPEN), 4, 9, CADECI%, 2)
                .TextMatrix(FilaGrid&, 5) = FECHATEX$(FEX)
                    OBSER$ = ""
                    If FEX < HOYY Then
                      ATRAS% = DIENTRE%(FEX, HOYY)
                      OBSER$ = "Atr." + TRIM$(Str$(ATRAS%)) + " Ds."
                    End If
                    If SALPREP > 0 Then
                      If OBSER$ <> "" Then OBSER$ = OBSER$ + "/ "
                      OBSER$ = OBSER$ + "Pr." + TRIM$(Str$(SALPREP&))
                    End If
                .TextMatrix(FilaGrid&, 6) = OBSER$
                .TextMatrix(FilaGrid&, 7) = Format(FECHATEXLAR$(FEX), "yyyymmdd")
             End With
              
           End If
         .MoveNext
         Loop
       End If
     End With
     GoSub PRESEDIS
     GoTo 15
     '
     '
15   Call ABRECOMPRAS
     SQLX$ = " SELECT N_ORSERV, Coef_Unids, cant_ocom, Cant_Rec, Npro_Ocom, Nume_Ocom, Fentre_Sol FROM OCOMDETA WITH(NOLOCK) "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
     SQLX$ = SQLX$ + "AND Stat_Ocom < 8 "   ' para que no este anulada ni cumplida
     SQLX$ = SQLX$ + "AND  Stat_Ocom <> -1 " 'Para que no  hayan O/C's sin aprobar
     SQLX$ = SQLX$ + "AND Cant_Rec < (Cant_Ocom - 0.005) "
     SQLX$ = SQLX$ + "ORDER BY Nume_Ocom ASC "
     '
     'Set OcoTemp = Compras.OpenRecordset(sqlx$, dbOpenSnapshot)
     
'     Dim OcoTemp As ADODB.Recordset
'     Set OcoTemp = New ADODB.Recordset
'     OcoTemp.CursorType = adOpenDynamic
'     OcoTemp.LockType = adLockReadOnly
'     OcoTemp.Open FILTSQL$(SQLX$), FLEXCONN, , , adcmtext
     Set OcoTemp = READSET(SQLX$)
     With OcoTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           NUORSE& = !N_ORSERV
           If NUORSE& < 1 Then
             COEFCAN = !Coef_Unids
             If COEFCAN < 0.005 Then COEFCAN = 1
             CACOM = !cant_ocom * COEFCAN
             CAREC = !Cant_Rec * COEFCAN
             CAPEN = CACOM - CAREC
             If CAPEN > 0 Then
               NCI% = !Npro_Ocom
               NROX& = !nume_ocom
               FESOL = CVINT(!Fentre_Sol)
               FEX = FESOL: OBSER$ = ""
                 If FEX < HOYY Then
                   ATRAS% = DIENTRE%(FESOL, HOYY)
                   'OBSER$ = Left$(FECHATEX$(FESOL), 5) + " - Atr." + TRIM$(Str$(ATRAS%)) + " Ds."
                   OBSER$ = "Atr." + TRIM$(Str$(ATRAS%)) + " Ds."
                   'FEX = HOYY
                 End If
'               LI2TEX$ = Space$(96)
'               Call REPLA(LI2TEX$, "OC." + TRIM$(CSTRING$(MKS$(NROX&), 3, 6, 0, 2)), 1, 9)
'               Call REPLA(LI2TEX$, RASOCLI$((-1) * NCI%), 15, 29)
'               Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN), 4, 9, 1, 2), 44, 9)
'               Call REPLA(LI2TEX$, FECHATEX$(FEX), 74, 9)
'               Call REPLA(LI2TEX$, OBSER$, 84, 24)
'               List2.AddItem LI2TEX$
               
               ' *** COMPLETA MSFLEXGRID STOCK PROYECTADO
               FilaGrid& = FilaGrid& + 1
               With gridStockProyectado
                  .Rows = .Rows + 1
                  .TextMatrix(FilaGrid&, 0) = "OC." + TRIM$(CSTRING$(MKS$(NROX&), 3, 6, 0, 2))
                  .TextMatrix(FilaGrid&, 1) = rasocli$((-1) * NCI%)
                  .TextMatrix(FilaGrid&, 2) = CSTRING$(MKS$(CAPEN), 4, 9, 1, 2)
                  .TextMatrix(FilaGrid&, 5) = FECHATEX$(FEX)
                  .TextMatrix(FilaGrid&, 6) = OBSER$
                  .TextMatrix(FilaGrid&, 7) = Format(FECHATEXLAR$(FEX), "yyyymmdd")
               End With
             
               TOSUMIN = TOSUMIN + CAPEN
             End If
           End If
         .MoveNext
         Loop
       End If
     End With
     GoSub PRESEDIS
     '

     '\\\OT-06-0149-WAR-25/04/06///
22      Call ABRECOMPRAS
     SQLX$ = " SELECT Cant_Aten, Acuen_Aten, Npro_Aten, Nume_Aten, Fentre_Sol FROM AUTOENT WITH(NOLOCK) "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
     SQLX$ = SQLX$ + "AND Stat_Aten < 8 "   ' para que no este anulada ni cumplida
     SQLX$ = SQLX$ + "AND  Stat_Aten <> -1 " 'Para que no  hayan O/C's sin aprobar
     SQLX$ = SQLX$ + "AND Acuen_Aten < (Cant_Aten - 0.005) "
     SQLX$ = SQLX$ + "ORDER BY Nume_Aten ASC "
     '
     'Set AtenTemp = Compras.OpenRecordset(sqlx$, dbOpenSnapshot)
     
'     Dim AtenTemp As ADODB.Recordset
'     Set AtenTemp = New ADODB.Recordset
'     AtenTemp.CursorType = adOpenDynamic
'     AtenTemp.LockType = adLockReadOnly
'     AtenTemp.Open FILTSQL$(SQLX$), FLEXCONN, , , adcmtext
     '
     Set AtenTemp = READSET(SQLX$)
     With AtenTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           NUORSE& = 0 '!N_OrServ
           If NUORSE& < 1 Then
             COEFCAN = 1 '!Coef_Unids
             If COEFCAN < 0.005 Then COEFCAN = 1
             CACOM = !Cant_Aten * COEFCAN
             CAREC = !Acuen_Aten * COEFCAN
             CAPEN = CACOM - CAREC
             If CAPEN > 0 Then
               NCI% = !Npro_Aten
               NROX& = !Nume_Aten
               FEX = CVINT(!Fentre_Sol)
               FESOL = CVINT(!Fentre_Sol)
               FEX = FESOL: OBSER$ = ""
                 If FEX < HOYY Then
                   ATRAS% = DIENTRE%(FESOL, HOYY)
                   OBSER$ = Left$(FECHATEX$(FESOL), 5) + " - Atr." + TRIM$(Str$(ATRAS%)) + " Ds."
                   FEX = HOYY
                 End If
'               LI2TEX$ = Space$(96)
'               Call REPLA(LI2TEX$, "AE." + TRIM$(CSTRING$(MKS$(NROX&), 3, 6, 0, 2)), 1, 9)
'               Call REPLA(LI2TEX$, RASOCLI$((-1) * NCI%), 15, 29)
'               Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN), 4, 9, CADECI%, 2), 44, 9)
'               Call REPLA(LI2TEX$, FECHATEX$(FEX), 74, 9)
'               Call REPLA(LI2TEX$, OBSER$, 84, 24)
'               List2.AddItem LI2TEX$
               
               ' *** COMPLETA MSFLEXGRID STOCK PROYECTADO
               FilaGrid& = FilaGrid& + 1
               With gridStockProyectado
                  .Rows = .Rows + 1
                  .TextMatrix(FilaGrid&, 0) = "AE." + TRIM$(CSTRING$(MKS$(NROX&), 3, 6, 0, 2))
                  .TextMatrix(FilaGrid&, 1) = rasocli$((-1) * NCI%)
                  .TextMatrix(FilaGrid&, 2) = CSTRING$(MKS$(CAPEN), 4, 9, CADECI%, 2)
                  .TextMatrix(FilaGrid&, 5) = FECHATEX$(FEX)
                  .TextMatrix(FilaGrid&, 6) = OBSER$
                  .TextMatrix(FilaGrid&, 7) = Format(FECHATEXLAR$(FEX), "yyyymmdd")
               End With
               
               TOSUMIN = TOSUMIN + CAPEN
             End If
           End If
         .MoveNext
         Loop
       End If
     End With
     GoSub PRESEDIS
     '\\\OT-06-0149-WAR-FIN///
     '
25   Call ABREORFAB
     SQLX$ = " SELECT CanProy,  CanApro,  NUMEORFA, FechEnSol, Referen, FechProyEntre, FechRecal FROM ORDEFABR WITH(NOLOCK) "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
     SQLX$ = SQLX$ + "AND StatuOrf < 3 "   ' para que no este anulada/cumplida/cerrada
     SQLX$ = SQLX$ + "AND CanApro < (CanProy - 0.005) "
     SQLX$ = SQLX$ + "ORDER BY NumeOrFa ASC "
     '
     'Set OfaTemp = Ofabric.OpenRecordset(sqlx$, dbOpenSnapshot)
'     Dim OfaTemp As ADODB.Recordset
'     Set OfaTemp = New ADODB.Recordset
'     OfaTemp.CursorType = adOpenDynamic
'     OfaTemp.LockType = adLockReadOnly
'     OfaTemp.Open FILTSQL$(SQLX$), FLEXCONN, , , adcmtext
     Set OfaTemp = READSET(SQLX$)
     With OfaTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           CACOM = !CANPROY
           CAREC = !CanApro
           CAPEN = CACOM - CAREC
           If CAPEN > 0 Then
             NROPED& = !NUMEORFA
             FEX = CVINT(!FechEnSol)
             REFE$ = !Referen
             LI2TEX$ = Space$(96)
             FPE% = CVINT(!FechProyEntre)
'             Call REPLA(LI2TEX$, "OF." + TRIM$(CSTRING$(MKS$(NROPED&), 3, 6, 0, 2)), 1, 9)
'             Call REPLA(LI2TEX$, REFE$, 15, 29)
'             Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN), 4, 9, CADECI%, 2), 44, 9)
'             Call REPLA(LI2TEX$, FECHATEX$(FEX), 74, 9)
'             'Call REPLA(LI2TEX$, IDENSEDE$, 84, 8)
'             '
'             If FPE% > 0 Then
'               FEX = FPE%
'               Call REPLA(LI2TEX$, FECHATEX$(FEX), 74, 9)
'               FEX = CVINT(!FechRecal)
'               Call REPLA(LI2TEX$, "Calc." + FECHATEX$(FEX), 84, 24)
'             End If
'             '
'             List2.AddItem LI2TEX$
             
             ' *** COMPLETA MSFLEXGRID STOCK PROYECTADO
             FilaGrid& = FilaGrid& + 1
             With gridStockProyectado
                .Rows = .Rows + 1
                .TextMatrix(FilaGrid&, 0) = "OF." + TRIM$(CSTRING$(MKS$(NROPED&), 3, 6, 0, 2))
                .TextMatrix(FilaGrid&, 1) = TRIM$(REFE$)
                .TextMatrix(FilaGrid&, 2) = CSTRING$(MKS$(CAPEN), 4, 9, CADECI%, 2)
                .TextMatrix(FilaGrid&, 5) = FECHATEX$(FEX)
                .TextMatrix(FilaGrid&, 7) = Format(FECHATEXLAR$(FEX), "yyyymmdd")
                If FPE% > 0 Then
                  FEX = FPE%
                  .TextMatrix(FilaGrid&, 5) = FECHATEX$(FEX)
                  .TextMatrix(FilaGrid&, 7) = Format(FECHATEXLAR$(FEX), "yyyymmdd")
                  FEX = CVINT(OfaTemp!FechRecal)
                  .TextMatrix(FilaGrid&, 6) = "Calc." + FECHATEX$(FEX)
                Else
                  If FEX < HOYY Then
                    ATRAS% = DIENTRE%(FEX, HOYY)
                    OBSER$ = "Atr." + TRIM$(Str$(ATRAS%)) + " Ds."
                    .TextMatrix(FilaGrid&, 6) = OBSER$
                  End If
                End If
             End With
               
             TOFABRI = TOFABRI + CAPEN
           End If
         .MoveNext
         Loop
       End If
     End With
     GoSub PRESEDIS
     'HASTA ACA ACTUALIZADO.
35   Call ABRESERVA
     If VersionBD% <= 2005 Then 'valida version por que cast(FechRes as date) da error en version anterior 2005 o menor
        SQLX$ = " SELECT CantRes, CANTCONS, Idsubor, IdenLote,  FechRes FROM RESERVA WITH(NOLOCK) "
        SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
        SQLX$ = SQLX$ + "AND CantCons < (CantRes - 0.005) "
        SQLX$ = SQLX$ + "ORDER BY IdSubOr ASC "
     Else
        'NOTA SE AGREGO EL SUM DE LAS CANTIDADES RESERV Y CONSUM PARA QUE EXPRESE
        'SI ES DE UNA MISMA OF Y MISMO LOTE LA RESERVA DE EL ITEM ACTIVO LA ACUMULE EN UNA SOLA FILA (ENVAPLAST)
        'EL cast(FechRes as date) ES PARA QUE SE AGRUPE BIEN P Q EL CAMPO ES DE FECHA Y HORA Y SI TIENE DIF DE SEGUNDOS NO AGRUPA
        SQLX$ = " SELECT SUM(CantRes)AS CANTRES, SUM(CANTCONS) AS CANTCONS, Idsubor, IdenLote,  cast(FechRes as date) FROM RESERVA WITH(NOLOCK) "
        SQLX$ = SQLX$ + " WHERE Cod_Inte = " & CI% & " "
        SQLX$ = SQLX$ + " AND CantCons < (CantRes - 0.005) "
        SQLX$ = SQLX$ + " GROUP BY  Idsubor, IdenLote,  cast(FechRes as date) "
        SQLX$ = SQLX$ + " ORDER BY IdSubOr ASC "
     End If
     '
     'Set ReserTemp = Stocks.OpenRecordset(sqlx$, dbOpenSnapshot)
     
'     Dim ReserTemp As ADODB.Recordset
'     Set ReserTemp = New ADODB.Recordset
'     ReserTemp.CursorType = adOpenDynamic
'     ReserTemp.LockType = adLockReadOnly
'     ReserTemp.Open FILTSQL$(SQLX$), FLEXCONN, , , adcmtext
     '
     Set ReserTemp = READSET(SQLX$)
     '
     With ReserTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           TORE# = XVALO(!CANTRES)
           CACO# = XVALO(!CANTCONS)
           '
           If TRIM$(UCase$(Unimed$(CI%))) = "GK" Then
              TORE# = TORE# / 1000
              CACO# = CACO# / 1000
           End If
           '
           SIGNO% = Sgn(CVS(Mid$(X$, 75, 4)))
           SARE = TORE# - CACO#
           If SARE * SIGNO% < 0 Then
             SARE = 0
           End If
           CAPEN = SARE
           If Abs(CAPEN) > 0 Then
             NORFA$ = TRIM$(SAFETEXT$(!idsubor))
             IDLOTXX$ = TRIM$(SAFETEXT$(!idenlote))
             OBSER$ = ""
             If InStr(EPAC$, "ENVAPLAST") > 0 Then
               NC = NroClienteSegunOf&(NORFA$)
               RASOC$ = rasocli$(NC)
               OBSER$ = "Cte: " + RASOC$
             Else
               If IDLOTXX$ <> "" Then OBSER$ = "Lote " + IDLOTXX$
            End If

             'FEX = CVINT(!FechRes) 'Este parrafo se ve que se agrego y pisa lo anterior de poner el lote
             If InStr(EPAC$, "ENVAPLAST") < 1 Then 'agregue el if para que no pise lo de envaplast de arriba
               FEX = CVINT(VALOBADA("ordefabr", "fechensol", "idsubor='" & NORFA$ & "'"))
               If FEX < HOYY Then
                 ATRAS% = DIENTRE%(FEX, HOYY)
                 OBSER$ = "Atr." + TRIM$(Str$(ATRAS%)) + " Ds."
               End If
             End If
             If Left$(NORFA$, 2) = "OR" Then
               REFE$ = "Reserva p/Reparación"
               GoTo 38
             End If
             '
             DESTIN$ = NORFA$
             CICON% = CODIFAB%(NORFA$)
             If CICON% > 0 Then
               DESTIN$ = TRIM$(CODEXT$(CICON%)) + " - " + TRIM$(DESCRIT0$(CICON%))
             End If
             '
             REFE$ = "Reserva p/" + DESTIN$
             '
             ' *** COMPLETA MSFLEXGRID STOCK PROYECTADO
38           FilaGrid& = FilaGrid& + 1
             With gridStockProyectado
                .Rows = .Rows + 1
                .TextMatrix(FilaGrid&, 0) = Left$(NORFA$, 12)
                .TextMatrix(FilaGrid&, 1) = REFE$
                .TextMatrix(FilaGrid&, 3) = CSTRING$(MKS$(CAPEN), 4, 9, CADECI%, 2)
                .TextMatrix(FilaGrid&, 5) = FECHATEX$(FEX)
                .TextMatrix(FilaGrid&, 6) = OBSER$
                .TextMatrix(FilaGrid&, 7) = Format(FECHATEXLAR$(FEX), "yyyymmdd")
             End With
                            
             TORESER = TORESER + CAPEN
           End If
         .MoveNext
         Loop
       End If
     End With
     '
     ' FABRICACIONES PENDIENTES DE APROBACION
     ' **************************************
     '
     SQLX$ = "SELECT COD_INTE, COD_EXTE, SUM(CANTINGRE) AS CANIN, SUM(CANTSALID) AS CANAPRO, DoPriCor, FECHMOV, IDTEXT"
     SQLX$ = SQLX$ + " FROM ORFAPEN WITH(NOLOCK)"
     SQLX$ = SQLX$ + " WHERE CodiMovi = 'OF' AND Cod_Inte = " & CStr(CI%)
     SQLX$ = SQLX$ + " GROUP BY Cod_Exte, DoPriCor, FechMov, COD_INTE, IDTEXT"
     SQLX$ = SQLX$ + " Having Sum(CANTINGRE) > Sum(CANTSALID)"
     SQLX$ = SQLX$ + " ORDER BY DOPRICOR"
     Set rs = READSET(SQLX$)
     With rs
         Do While Not .EOF
             CIII% = XVALO(!cod_inte)
             DOPRI$ = SAFETEXT$(!dopricor)
             REFE$ = TRIM$(VALOBADA("ORFAPEN", "ReferCor", "DoPriCor = '" & DOPRI$ & "' AND CODIMOVI= 'OF'"))
             CAPEN = XVALO(!CANIN) - XVALO(!CanApro)
    
             ' *** COMPLETA MSFLEXGRID STOCK PROYECTADO
             FilaGrid& = FilaGrid& + 1
             With gridStockProyectado
               .Rows = .Rows + 1
               .TextMatrix(FilaGrid&, 0) = Left$(DOPRI$, 12)
               .TextMatrix(FilaGrid&, 1) = REFE$
               .TextMatrix(FilaGrid&, 3) = CSTRING$(MKS$(CAPEN), 4, 9, CADECI%, 2)
               .TextMatrix(FilaGrid&, 5) = FECHATEX$(HOY(HO$))
               .TextMatrix(FilaGrid&, 6) = "PEND. APROBACION"
               .TextMatrix(FilaGrid&, 7) = Format(FECHATEXLAR$(HOY(HO$)), "yyyymmdd")
             End With
             
             TOFABRI = TOFABRI + CAPEN
             
             .MoveNext
         Loop
     End With
     rs.Close
     Set rs = Nothing
     '
     ' ********************************************************************************
     GoSub PRESEDIS
     '
     STOTER = 0
     Call ABRESTOCKS
     SQLX$ = "SELECT SUM(CANT_ENTREG) AS ENTREGADO, SUM(CANT_RENDID) AS RENDIDO, SUM(Cant_Devu) AS DEVU, SUM(CANT_ASIG) AS ASIGNA, SUBSTRING(Nro_OCOM,1,9) AS OCOM, Npro_Ocom FROM MACONSIG WITH(NOLOCK) "
     SQLX$ = SQLX$ + " WHERE Cod_Inte = " & CI% & " "
     SQLX$ = SQLX$ + " GROUP BY SUBSTRING(Nro_OCOM,1,9), NPRO_OCOM, COD_INTE "
     ' SUBSTRING(Nro_OCOM,1,9) TIENE EL FORMATO OC-999999
     
     Set LIMACONTMP = READSET(SQLX$)
     With LIMACONTMP
       On Error Resume Next
       .MoveFirst
       'Valida que la tabla no este vacia
       If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
           TORE# = XVALO(!Entregado)
           CACO# = XVALO(!RENDIDO)
           DEVU# = XVALO(!DEVU)
           
           CAPEN = TORE# - CACO# - DEVU#
           CAPENENT# = XVALO(!ASIGNA) - XVALO(!Entregado)
           NUOC& = XVALO(Mid$(!OCOM, 4, 6))
           '
           If CAPEN > 0 Then
              NUREX$ = "RC"
              NORFA$ = SAFETEXT$(!OCOM)
              NPRO% = XVALO(!Npro_Ocom)
              REFE$ = "Pendiente de rendir por " & rasocli$((-1) * NPRO%)
              ORDENCOMPRA& = XVALO(Mid$(NORFA$, 4, 6))
              FEX = HOY(HO$)
              SDOR$ = ""
              
              ' *** COMPLETA MSFLEXGRID STOCK PROYECTADO
              FilaGrid& = FilaGrid& + 1
              With gridStockProyectado
                 .Rows = .Rows + 1
                 .TextMatrix(FilaGrid&, 0) = Left$(NUREX$, 12)
                 .TextMatrix(FilaGrid&, 1) = REFE$
                 If CAPEN > 0 Then
                   .TextMatrix(FilaGrid&, 2) = CSTRING$(MKS$(CAPEN), 4, 9, CADECI%, 2)
                 Else
                   .TextMatrix(FilaGrid&, 3) = CSTRING$(MKS$(Abs(CAPEN)), 4, 9, CADECI%, 2)
                 End If
                 .TextMatrix(FilaGrid&, 5) = FECHATEX$(FEX)
                 .TextMatrix(FilaGrid&, 6) = NORFA$
                 .TextMatrix(FilaGrid&, 7) = Format(FECHATEXLAR$(FEX), "yyyymmdd")
              End With
             
              STOREMO = STOREMO + CAPEN
              STOTER = STOTER + CAPEN
              'GoTo 47
           End If
           '
47         If CAPENENT# >= 0.05 Then
             If (STATOCOM%(NUOC&) < 8 And STATOCOM%(NUOC&) >= 0) Then
                NORFA$ = SAFETEXT$(!OCOM)
                NPRO% = XVALO(!Npro_Ocom)
                REFE$ = rasocli$((-1) * NPRO%)
                FEX = CVINT(VALOBADA("OCOMENCA", "FEMI_OCOM", "NUME_OCOM = " & XVALO(Mid$(NORFA$, 4))))
                SDOR$ = ""

                ' *** COMPLETA MSFLEXGRID STOCK PROYECTADO
                FilaGrid& = FilaGrid& + 1
                With gridStockProyectado
                   .Rows = .Rows + 1
                   .TextMatrix(FilaGrid&, 0) = Left$(NORFA$, 12)
                   .TextMatrix(FilaGrid&, 1) = REFE$
                   If CAPENENT# > 0 Then
                     .TextMatrix(FilaGrid&, 3) = CSTRING$(MKS$(CAPENENT#), 4, 9, CADECI%, 2)
                   Else
                     .TextMatrix(FilaGrid&, 2) = CSTRING$(MKS$(Abs(CAPENENT#)), 4, 9, CADECI%, 2)
                   End If
                   .TextMatrix(FilaGrid&, 5) = FECHATEX$(FEX)
                   .TextMatrix(FilaGrid&, 7) = Format(FECHATEXLAR$(FEX), "yyyymmdd")
                End With

                TORESER = TORESER + CAPENENT
             End If
           End If
           '
         .MoveNext
         Loop
       End If
     End With
     '
     If Abs(STOTER) >= 0.005 Then
        XXX$ = "TERCEROS" + CSTRING$(MKS$(STOTER), 4, 11, CADECI%, 2)
        List1.AddItem XXX$
     End If
     '
    GoSub PRESEDIS
     '
'     ' ORDENAMIENTO POR FECHA
'     For U& = 1 To List2.ListCount
'45     TTXX$ = List2.LIST(U& - 1)
'       FE1% = FECHANUM(Mid$(TTXX$, 74, 8))
'       If Mid$(UCase$(TTXX$), 15, 12) = "STOCK PROPIO" Then FE1% = 0
'       For U2& = U& + 1 To List2.ListCount
'         TTYY$ = List2.LIST(U2& - 1)
'         FE2% = FECHANUM(Mid$(TTYY$, 74, 8))
'         If FE2% < FE1% Then
'           List2.RemoveItem U2& - 1
'           List2.AddItem TTYY$, U& - 1
'           GoTo 45
'         End If
'       Next U2&
'     Next U&
'     List2.Refresh
          
     ' *** ORDENAMIENTO DE LA GRILLA POR FECHA. SOLO LA COLUMNA OCULTA
     ' *** CON LA FECHA EN FORMATO INGLES
     gridStockProyectado.COL = 7
     gridStockProyectado.ColSel = 7
     gridStockProyectado.Sort = 1

'     ' STOCK PROYECTADO
'     STOPRO = 0
'     For U& = 1 To List2.ListCount
'       TTXX$ = List2.LIST(U& - 1)
'       STOPRO = STOPRO + XVALO(Mid$(TTXX$, 44, 9)) - XVALO(Mid$(TTXX$, 53, 9))
'       'STOPROX$ = FORMATNUM$(STOPRO, "F10.1")
'       STOPROX$ = FORMATNUM$(STOPRO, "F10." & CNTDC$(CIXI%))
'       Call REPLA(TTXX$, STOPROX$, 62, 10)
'       List2.RemoveItem U& - 1
'       List2.AddItem TTXX$, U& - 1
'     Next U&
'     List2.Refresh
     
     ' *** CALCULA STOCK PROYECTADO POR FECHA EN EL MSFLEXGRID
     STOPRO = 0
     With gridStockProyectado
        For FilaGrid& = 1 To .Rows - 1
          STOPRO = STOPRO + XVALO(.TextMatrix(FilaGrid&, 2)) - XVALO(.TextMatrix(FilaGrid&, 3))
          .TextMatrix(FilaGrid&, 4) = TRIM$(FORMATNUM$(STOPRO, "F12.1"))
        Next FilaGrid&
     End With
     
     '
20   LIUBIC.Clear
     CNTREG& = 0
     If InStr(EPAC$, "AHP") < 1 Then
       CNTREG& = CantRegistros&("UBICASQL", "", "NOMESS")
     End If
     If CNTREG& > 0 Then  'si ya migro las ubicaciones a la tabla sql las carga de aqui
            SQLX$ = "SELECT CODIUBICA FROM UBICASQL  WITH(NOLOCK) "
            SQLX$ = SQLX$ + " WHERE COD_INTE = " & CI%
            SQLX$ = SQLX$ + " ORDER BY CODIUBICA"
            Set rs = READSET(SQLX$)
            With rs
              Do While Not .EOF
                 CODUB$ = TRIM$(SAFETEXT$(!CODIUBICA))
                 If CODUB$ <> "" Then
                    LIUBIC.AddItem CODUB$
                 End If
                 .MoveNext
              Loop
            End With
            rs.Close
     Else
        RGXX$ = RECFILT$("UBISTOCK", 1, MKI$(CI%))
        For U& = 1 To Len(RGXX$) Step 4
           REG& = CVS(Mid$(RGXX$, U&, 4))
           X$ = REGLEIDO$("UBISTOCK", REG&)
           If CVI(Left$(X$, 2)) <> CI% Then
              X$ = REGLEIDO$("UBISTOCK", 1)
              Call GRAREG("UBISTOCK", X$, 1)
              Call CIERRARCH("UBISTOCK")
              GoTo 20
           End If
           '
           If TRIM$(Mid$(X$, 3, 24)) <> "" Then
              Y$ = Mid$(X$, 3, 24)
              LIUBIC.AddItem Y$
           End If
        Next U&
     End If
     '
     Screen.MousePointer = 1
'     Call GRAPHDIS
     '
     Exit Sub
     '
PRESEDIS:
     '
     
     STOMI.Caption = TRIM$(CSTRING$(MKS$(STOMINI), 4, 12, CADECI%, 2))
     LECONOMI.Caption = TRIM$(CSTRING$(MKS$(LECONOM), 4, 12, CADECI%, 2))
     '
     Label101.Caption = TRIM$(CSTRING$(MKS$(STOPROP), 4, 12, CADECI%, 2))
     Label102.Caption = TRIM$(CSTRING$(MKS$(STOREMO), 4, 12, CADECI%, 2))
     STOTOT = STOPROP + STOREMO
     Label103.Caption = TRIM$(CSTRING$(MKS$(STOTOT), 4, 12, CADECI%, 2))
     '
     Label201.Caption = TRIM$(CSTRING$(MKS$(TOPEDID), 4, 12, CADECI%, 2))
     Label202.Caption = TRIM$(CSTRING$(MKS$(TORESER), 4, 12, CADECI%, 2))
     TOCOMPR = TOPEDID + TORESER
     Label203.Caption = TRIM$(CSTRING$(MKS$(TOCOMPR), 4, 12, CADECI%, 2))
     '
     Label301.Caption = TRIM$(CSTRING$(MKS$(TOSUMIN), 4, 12, CADECI%, 2))
     Label302.Caption = TRIM$(CSTRING$(MKS$(TOFABRI), 4, 12, CADECI%, 2))
     TOINPRE = TOSUMIN + TOFABRI
     Label303.Caption = TRIM$(CSTRING$(MKS$(TOINPRE), 4, 12, CADECI%, 2))
     '
     STONET = STOTOT - TOCOMPR
     STODIS = STONET - STOMINI + TOINPRE
     STOFAL = 0
     If STODIS < 0 Then
        STOFAL = Abs(STODIS)
        STODIS = 0
     End If
     '
     Label401.Caption = TRIM$(CSTRING$(MKS$(STONET), 4, 12, CADECI%, 2))
     Label402.Caption = TRIM$(CSTRING$(MKS$(STODIS), 4, 12, CADECI%, 2))
     Label403.Caption = TRIM$(CSTRING$(MKS$(STOFAL), 4, 12, CADECI%, 2))
     Label404.Caption = TRIM$(CSTRING$(MKS$(XVALO(Label103) - XVALO(Label202)), 4, 12, CADECI%, 2))
     '
     DoEvents
     '
Return
'
COBRECEPTMP.Close
Set COBRECEPTMP = Nothing
PEDTEMP.Close
Set PEDTEMP = Nothing
OcoTemp.Close
Set OcoTemp = Nothing
AtenTemp.Close
Set AtenTemp = Nothing
OfaTemp.Close
Set OfaTemp = Nothing
ReserTemp.Close
Set ReserTemp = Nothing
LIMACONTMP.Close
Set LIMACONTMP = Nothing
'
End Sub
'
Sub GRAPHDIS()
   '
   '
   ALTUMA = 0.78 * SITCOB09.Picture1.Height
   STOK = XVALO(Label103): If STOK < 0 Then STOK = 0
   COMP = XVALO(Label301): If COMP < 0 Then COMP = 0
   FABR = XVALO(Label302): If FABR < 0 Then FABR = 0
   '
   MINI = XVALO(STOMI): If MINI < 0 Then MINI = 0
   RESE = XVALO(Label203): If RESE < 0 Then RESE = 0
   '
   Screen.MousePointer = 11
   CIIX% = CODINT%(Text1)
   TENT = 0
   RFF$ = RECFILT$("COBERFA", 2, MKI$(CIIX%))
   For KU& = 1 To Len(RFF$) Step 4
     KL& = CVS(Mid$(RFF$, KU&, 4))
     XZ$ = REGLEIDO$("COBERFA", KL&)
     If CVI(Mid$(XZ$, 51, 2)) = 22 Then
       TENT = TENT + CVS(Mid$(XZ$, 57, 4))
     End If
   Next KU&
   If TENT < 0 Then TENT = 0
   '
   MAXI = STOK + COMP + FABR
   If MINI + RESE + TENT > MAXI Then
     MAXI = MINI + RESE + TENT
   End If
   Screen.MousePointer = 1
   '
   If MAXI < 1 Then
     SITCOB09.Shape1.Visible = False
     SITCOB09.Shape2.Visible = False
     SITCOB09.Shape3.Visible = False
     SITCOB09.Shape4.Visible = False
     SITCOB09.Shape5.Visible = False
     SITCOB09.Shape6.Visible = False
     Exit Sub
   End If
   '
   COEF = ALTUMA / MAXI
   
   SITCOB09.Shape1.Height = STOK * COEF
   SITCOB09.Shape1.Top = 0.94 * SITCOB09.Picture1.Height - SITCOB09.Shape1.Height
   SITCOB09.Shape2.Height = COMP * COEF
   SITCOB09.Shape2.Top = SITCOB09.Shape1.Top - SITCOB09.Shape2.Height
   SITCOB09.Shape3.Height = FABR * COEF
   SITCOB09.Shape3.Top = SITCOB09.Shape2.Top - SITCOB09.Shape3.Height
   '
   SITCOB09.Shape4.Height = MINI * COEF
   SITCOB09.Shape4.Top = 0.94 * SITCOB09.Picture1.Height - SITCOB09.Shape4.Height
   SITCOB09.Shape5.Height = RESE * COEF
   SITCOB09.Shape5.Top = SITCOB09.Shape4.Top - SITCOB09.Shape5.Height
   SITCOB09.Shape6.Height = TENT * COEF
   SITCOB09.Shape6.Top = SITCOB09.Shape5.Top - SITCOB09.Shape6.Height
   '
   If STOK > 0 Then SITCOB09.Shape1.Visible = True
   If COMP > 0 Then SITCOB09.Shape2.Visible = True
   If FABR > 0 Then SITCOB09.Shape3.Visible = True
   If MINI > 0 Then SITCOB09.Shape4.Visible = True
   If RESE > 0 Then SITCOB09.Shape5.Visible = True
   If TENT > 0 Then SITCOB09.Shape6.Visible = True

End Sub

Private Sub Timer1_Timer()
    If Me.Enabled = False Then
      On Error Resume Next
      For U& = 0 To Controls.Count - 1
        If Left$(UCase$(Me.Controls(U&).Name), 5) = "COMMA" Then
           Me.Controls(U&).Refresh
        End If
      Next U&
      On Error GoTo 0
    End If
End Sub

'
Sub CONAPLI(APLINVO$, TERMINA%)
   TERMINA% = 1
   Select Case APLINVO$
       '
       'TRANSACCIONES
       '
       Case "AJUSEXIS"
           Call MOVISTK07.Command25_Click                 ' AJUSTES DE EXISTENCIAS
       Case "TRANSTOCK"
           Call MOVISTK07.Command24_Click                 ' TRANSFERENCIAS DE STOCK
       Case "VALCONMAT"
           Call MOVISTK07.Command21_Click                 ' VALE DE CONSUMO DE MATERIALES
       Case "DESCRAP"
           Call MOVISTK07.Command17_Click                 ' DESCARTE DE SCRAP
       Case "ENTPROPROD"
           Call MOVISTK07.Command18_Click                 ' PROPIA PRODUCCION
       
       'CONSULTAS Y LISTADOS
       '
       Case "FICKARDEX"
           Call QUERSTK07.Command9_Click                 ' FICHA KARDEX
       Case "LISEXIUBI"
           Call QUERSTK07.Command2_Click                 ' LISTADOS EXISTENCIAS Y UBICACIONES
       Case "GESRESTOCK"
           Call QUERSTK07.Command14_Click                 ' GESTION DE RESERVAS
       Case "EXISACT"
       '
       'LISTADO DE CONTROL
       '
           Call QUERSTK07.Command19_Click                ' EXISTENCIAS ACTUALES
       Case "EXISNEG"
           Call QUERSTK07.Command30_Click                ' EXISTENCIAS NEGATIVAS
       Case "EXISUBMIN"
           Call QUERSTK07.Command20_Click                ' EXISTENCIAS SUB MINIMAS
           
       Case "LISTGENCON"
           Call Command9_Click                 ' LISTADOS GENERALES
       Case "LISCONCON"
           Call Command9_Click                 ' LISTADO DE CONTROL
       Case "MATENDEPO"
           Call QUERSTK07.Command6_Click                ' MATERIAL EN DEPOSITO
       Case "PRODPROCE"
           Call QUERSTK07.Command3_Click                ' PRODUCTOS EN PROCESO
      Case "DISGENSTOCK"
           Call QUERSTK07.Command5_Click                ' DISPONIBILIDAD GENERAL DE STOCK
      Case "CONGENSTOCK"
           Call MOVISTK07.Command26_Click                ' CONFIGURACION GENERAL DE STOCK
      Case "ESTASTOCK"
           Call QUERSTK07.Command27_Click                ' ESTADISTICAS DE STOCK
           
      Case "STOCKRANGOAGRUPADO"
         Call mnuStockRangoAgrupado_Click
      Case "STOCKPROVRANGO"
      Call mnuStockProveedorRango_Click
      Case "STOCKRANGOTOD"
      Call mnuStockRangoTodos_Click
    
      Case Else
           TERMINA% = 0
   End Select
End Sub

Private Sub Timer3_Timer()
    '
    ACTUAL = Mid$(AHORA, 13, 5)
    HORAPROGRAMADA = Control("SITUASTO", "HORAPROC")
    HORAPROGRAMADA = Mid$(HORAPROGRAMADA, 1, 5)
    If HORAPROGRAMADA = ACTUAL Then
       Timer3.Enabled = False
       Timer4.Enabled = True

       Call CFGSALSTO.RECALCUVALO("AUTO")
    End If
    '
End Sub

Private Sub Timer4_Timer()
     Static ENTRADAS%
     If ENTRADAS% > 20 Then
        Timer4.Enabled = False
        ENTRADAS% = 0
        Call CFGSALSTO.STOCKS_TALLERES("AUTO")
     Else
       ENTRADAS% = ENTRADAS% + 1
     End If
     
End Sub

Private Sub Timer5_Timer()

    If InStr(1, EMPREAC$, "ACHERNAR") > 0 Then
        '
        USUAR$ = Control("CONSTO07", "USLIAUUB")
        If USUAR$ = IDENTER$ Then
            HORA$ = TRIM$(Control("CONSTO07", "HOAUTUBI"))
            HORAVB$ = Format(Time, "hh:mm")
            If InStr(1, HORAVB$, HORA$, vbTextCompare) > 0 Then
                ULTIGENEAUTO% = FECHANUM(Control("CONSTO07", "ULTLIUBI"))
                If ULTIGENEAUTO% <= HOY(HO$) Then
                    Archivo$ = TRIM$(Control("CONSTO07", "LIAUTUBI"))
                    If Archivo$ <> "" Then
                        Call GRACONTROL("CONSTO07", "ULTLIUBI", FECHATEX$(HOY(HO$)))
                        Call APARTIMA
                        Call QUERSTK07.GENERAXCEL(Archivo$)
                    End If
                End If
            End If
        Else
            Timer5.Enabled = False
        End If
        '
    Else
        Timer5.Enabled = False
    End If

End Sub


Private Sub VerSoloSinLote_Click()
   Call asignalotesitemfacturados(1)
End Sub


Private Sub VerTodasLasFacturas_Click()
   Call asignalotesitemfacturados(0)
End Sub


