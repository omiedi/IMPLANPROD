VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form ACLI_GES07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MAESTRO DE CLIENTES"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11790
   DrawStyle       =   1  'Dash
   FillColor       =   &H00C0C0FF&
   FillStyle       =   3  'Vertical Line
   Icon            =   "ACLI_GES07.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   11790
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      ForeColor       =   &H80000008&
      Height          =   8100
      Index           =   0
      Left            =   10920
      ScaleHeight     =   8070
      ScaleWidth      =   1245
      TabIndex        =   205
      Top             =   0
      Width           =   1275
      Begin VB.CommandButton Command4 
         Caption         =   "Ref."
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
         Left            =   120
         TabIndex        =   219
         ToolTipText     =   "RefrescarDeudor12"
         Top             =   1440
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.CommandButton Command21 
         Caption         =   "New"
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
         Picture         =   "ACLI_GES07.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   218
         ToolTipText     =   "Alta de Nuevo Cliente"
         Top             =   1890
         Width           =   650
      End
      Begin VB.CommandButton Command3 
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
         Left            =   105
         Picture         =   "ACLI_GES07.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   217
         ToolTipText     =   "Exit"
         Top             =   270
         Width           =   650
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Sort"
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
         Picture         =   "ACLI_GES07.frx":06D6
         Style           =   1  'Graphical
         TabIndex        =   216
         ToolTipText     =   "Ordena, Depura Duplicados y Corrige Errores en  Base de Datos de Clientes"
         Top             =   3960
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Baja"
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
         Picture         =   "ACLI_GES07.frx":09E0
         Style           =   1  'Graphical
         TabIndex        =   215
         ToolTipText     =   "Baja Lógica de la Cuenta"
         Top             =   3270
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command99 
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
         Left            =   105
         Picture         =   "ACLI_GES07.frx":0CEA
         Style           =   1  'Graphical
         TabIndex        =   214
         ToolTipText     =   "Control de Integridad de la Base de Datos"
         Top             =   8100
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
         Height          =   690
         Left            =   105
         Picture         =   "ACLI_GES07.frx":0FF4
         Style           =   1  'Graphical
         TabIndex        =   213
         ToolTipText     =   "Configuración - Tablas de Validación"
         Top             =   6320
         Width           =   650
      End
      Begin VB.CommandButton Command13 
         Caption         =   "List"
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
         Picture         =   "ACLI_GES07.frx":1436
         Style           =   1  'Graphical
         TabIndex        =   212
         ToolTipText     =   "Emisión de Listados"
         Top             =   4935
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command25 
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
         Height          =   690
         Left            =   105
         Picture         =   "ACLI_GES07.frx":1740
         Style           =   1  'Graphical
         TabIndex        =   211
         ToolTipText     =   "Imprime Ficha del Cliente"
         Top             =   5625
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   100
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   209
         Top             =   7560
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   210
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Parar"
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
         Picture         =   "ACLI_GES07.frx":1DAA
         Style           =   1  'Graphical
         TabIndex        =   208
         ToolTipText     =   "Suspensión de Operaciones"
         Top             =   2580
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command28 
         Caption         =   "Clte"
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
         Left            =   105
         TabIndex        =   207
         Top             =   4670
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   206
         ToolTipText     =   "DOBLE CLICK PARA DETENER PROCESO"
         Top             =   7750
         Visible         =   0   'False
         Width           =   850
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10200
      Top             =   1200
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "ACLI_GES07.frx":20B4
      Top             =   720
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   4620
      OleObjectBlob   =   "ACLI_GES07.frx":22E8
      TabIndex        =   143
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   6090
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   142
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E2D3C0&
      Caption         =   "Destino Fábrica"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   5595
      TabIndex        =   104
      Top             =   180
      Width           =   1170
   End
   Begin VB.TextBox Text6 
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
      Left            =   3240
      TabIndex        =   97
      Top             =   315
      Width           =   1965
   End
   Begin VB.CommandButton Command10 
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
      Left            =   5200
      Picture         =   "ACLI_GES07.frx":2346
      TabIndex        =   96
      Top             =   315
      Width           =   175
   End
   Begin VB.Frame Frame2 
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
      Height          =   880
      Left            =   8720
      TabIndex        =   91
      Top             =   210
      Width           =   1990
      Begin VB.CommandButton Command20 
         Height          =   500
         Left            =   1280
         Picture         =   "ACLI_GES07.frx":2650
         Style           =   1  'Graphical
         TabIndex        =   94
         ToolTipText     =   "Búsqueda por CUIT, Domicilio, Nombre de Fantasía, etc"
         Top             =   250
         Width           =   550
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
         Height          =   500
         Left            =   155
         Picture         =   "ACLI_GES07.frx":39B6
         Style           =   1  'Graphical
         TabIndex        =   93
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   250
         Width           =   550
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
         Height          =   500
         Left            =   725
         Picture         =   "ACLI_GES07.frx":3CC0
         Style           =   1  'Graphical
         TabIndex        =   92
         ToolTipText     =   "Restablece los Valores Anteriores - Descarta Cambios Efectuados"
         Top             =   250
         Width           =   550
      End
   End
   Begin VB.TextBox Text23 
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
      Left            =   6915
      TabIndex        =   88
      Top             =   735
      Width           =   1425
   End
   Begin VB.ListBox LINDICE 
      Height          =   1035
      Left            =   10920
      Sorted          =   -1  'True
      TabIndex        =   77
      Top             =   3780
      Visible         =   0   'False
      Width           =   1170
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
      Height          =   300
      Left            =   1035
      TabIndex        =   2
      Top             =   720
      Width           =   5550
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
      Left            =   1035
      MaxLength       =   6
      TabIndex        =   0
      Top             =   315
      Width           =   780
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
      Left            =   1815
      Picture         =   "ACLI_GES07.frx":3FCA
      TabIndex        =   1
      Top             =   315
      Width           =   175
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6900
      Left            =   120
      TabIndex        =   36
      Top             =   1320
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   12171
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      BackColor       =   14865344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "ACLI_GES07.frx":42D4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label17"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Anotaciones"
      TabPicture(1)   =   "ACLI_GES07.frx":42F0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Line3"
      Tab(1).Control(1)=   "Frame3(0)"
      Tab(1).Control(2)=   "Picture4"
      Tab(1).Control(3)=   "Command5"
      Tab(1).Control(4)=   "Command6"
      Tab(1).Control(5)=   "Command8"
      Tab(1).Control(6)=   "Command17"
      Tab(1).Control(7)=   "Text28"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Cuenta Corriente"
      TabPicture(2)   =   "ACLI_GES07.frx":430C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame18"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Entrega"
      TabPicture(3)   =   "ACLI_GES07.frx":4328
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4(0)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame22"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame20"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Frame1(0)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Cobranza"
      TabPicture(4)   =   "ACLI_GES07.frx":4344
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame7(0)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Pedidos"
      TabPicture(5)   =   "ACLI_GES07.frx":4360
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame13"
      Tab(5).Control(1)=   "Frame8"
      Tab(5).Control(2)=   "Label82"
      Tab(5).Control(3)=   "Label74"
      Tab(5).Control(4)=   "Label73"
      Tab(5).ControlCount=   5
      Begin VB.Frame Frame13 
         Caption         =   "DETALLE DE PEDIDOS PENDIENTES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   -74760
         TabIndex        =   169
         Top             =   3600
         Width           =   10095
         Begin VB.CommandButton Command29 
            Caption         =   "Envio de Pendientes por Email"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   177
            Top             =   2450
            Width           =   2775
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
            Height          =   2055
            Left            =   120
            TabIndex        =   170
            Top             =   350
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   3625
            _Version        =   393216
            Cols            =   8
            ForeColor       =   0
            WordWrap        =   -1  'True
            AllowUserResizing=   1
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
         Begin VB.Label Label75 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1530
            TabIndex        =   178
            Top             =   0
            Visible         =   0   'False
            Width           =   2025
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label51 
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
            Height          =   300
            Left            =   8340
            TabIndex        =   174
            Top             =   2475
            Width           =   1575
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total en Pesos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   7395
            TabIndex        =   173
            Top             =   2400
            Width           =   870
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Moneda Original"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   4560
            TabIndex        =   172
            Top             =   2400
            Width           =   1110
         End
         Begin VB.Label Label11 
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
            Height          =   300
            Left            =   5760
            TabIndex        =   171
            Top             =   2475
            Width           =   1575
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "PEDIDOS PENDIENTES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   -74760
         TabIndex        =   167
         Top             =   840
         Width           =   10095
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
            Height          =   2175
            Left            =   120
            TabIndex        =   168
            Top             =   360
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   3836
            _Version        =   393216
            Cols            =   5
            ForeColor       =   0
            AllowUserResizing=   1
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
      End
      Begin VB.Frame Frame7 
         Height          =   6255
         Index           =   0
         Left            =   -74970
         TabIndex        =   145
         Top             =   360
         Width           =   10520
         Begin VB.Frame Frame7 
            Caption         =   "DATOS DE COBRANZA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1590
            Index           =   3
            Left            =   130
            TabIndex        =   150
            Top             =   120
            Width           =   10260
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
               Index           =   2
               Left            =   930
               MaxLength       =   48
               TabIndex        =   156
               Text            =   " "
               Top             =   1100
               Width           =   4320
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
               Index           =   0
               Left            =   930
               MaxLength       =   48
               TabIndex        =   155
               Text            =   " "
               Top             =   315
               Width           =   4320
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
               Index           =   1
               Left            =   930
               MaxLength       =   48
               TabIndex        =   154
               Text            =   " "
               Top             =   700
               Width           =   4320
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
               Index           =   3
               Left            =   7110
               MaxLength       =   48
               TabIndex        =   153
               Text            =   " "
               Top             =   315
               Width           =   2955
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
               Index           =   4
               Left            =   7110
               TabIndex        =   152
               Top             =   700
               Width           =   2955
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
               Index           =   5
               Left            =   6480
               MaxLength       =   128
               TabIndex        =   151
               Text            =   " "
               Top             =   1100
               Width           =   3585
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Mail"
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
               Left            =   -15
               TabIndex        =   162
               Top             =   1200
               Width           =   855
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Teléfono"
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
               Left            =   -15
               TabIndex        =   161
               Top             =   840
               Width           =   855
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Contacto"
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
               Left            =   -15
               TabIndex        =   160
               Top             =   420
               Width           =   855
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Días de Reclamo"
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
               Index           =   4
               Left            =   5445
               TabIndex        =   159
               Top             =   420
               Width           =   1590
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Días de Pago"
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
               Index           =   5
               Left            =   5535
               TabIndex        =   158
               Top             =   795
               Width           =   1485
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Mail Factura"
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
               Left            =   5220
               TabIndex        =   157
               Top             =   1200
               Width           =   1170
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "NOVEDADES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1950
            Index           =   1
            Left            =   50
            TabIndex        =   148
            Top             =   4215
            Width           =   10400
            Begin VB.Frame Frame16 
               Height          =   495
               Left            =   120
               TabIndex        =   185
               Top             =   180
               Width           =   10170
               Begin VB.Label Label44 
                  BackStyle       =   0  'Transparent
                  Caption         =   $"ACLI_GES07.frx":437C
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
                  TabIndex        =   186
                  Top             =   210
                  Width           =   9735
               End
            End
            Begin VB.CommandButton Command30 
               Caption         =   "O.K"
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
               Left            =   8240
               Picture         =   "ACLI_GES07.frx":442D
               Style           =   1  'Graphical
               TabIndex        =   163
               ToolTipText     =   "Agrega el Texo al Anotador"
               Top             =   1120
               Width           =   750
            End
            Begin VB.TextBox Text14 
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
               Height          =   1050
               Left            =   1320
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   149
               Top             =   735
               Width           =   6840
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
               Height          =   300
               Index           =   0
               Left            =   8235
               OleObjectBlob   =   "ACLI_GES07.frx":4737
               TabIndex        =   165
               Top             =   720
               Width           =   2100
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
               Height          =   300
               Index           =   2
               Left            =   120
               OleObjectBlob   =   "ACLI_GES07.frx":478D
               TabIndex        =   166
               Top             =   960
               Width           =   1155
            End
            Begin VB.Line Line23 
               X1              =   10580
               X2              =   10580
               Y1              =   315
               Y2              =   1900
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "ANOTACIONES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2340
            Index           =   1
            Left            =   50
            TabIndex        =   146
            Top             =   1800
            Width           =   10400
            Begin VB.Frame Frame15 
               Height          =   495
               Left            =   110
               TabIndex        =   181
               Top             =   200
               Width           =   10170
               Begin VB.Label Label91 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Usuario"
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
                  Left            =   8160
                  TabIndex        =   184
                  Top             =   210
                  Width           =   975
               End
               Begin VB.Label Label90 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Detalle"
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
                  Left            =   1320
                  TabIndex        =   183
                  Top             =   210
                  Width           =   975
               End
               Begin VB.Label Label10 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Fecha"
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
                  TabIndex        =   182
                  Top             =   210
                  Width           =   975
               End
            End
            Begin VB.ListBox List4 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1380
               Left            =   100
               TabIndex        =   147
               Top             =   690
               Width           =   10155
            End
            Begin VB.Line Line4 
               Index           =   1
               X1              =   0
               X2              =   10605
               Y1              =   4410
               Y2              =   4410
            End
            Begin VB.Line Line1 
               Index           =   1
               X1              =   10580
               X2              =   10580
               Y1              =   320
               Y2              =   6225
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "SUCURSALES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1665
         Index           =   0
         Left            =   -68740
         TabIndex        =   133
         Top             =   600
         Width           =   4005
         Begin VB.CommandButton CmdBorrar 
            Caption         =   "Borrar"
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
            Left            =   600
            Style           =   1  'Graphical
            TabIndex        =   244
            ToolTipText     =   "Borrar"
            Top             =   1200
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton CMDINSTRUCCION 
            Caption         =   "Instrucción"
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
            Left            =   1400
            Style           =   1  'Graphical
            TabIndex        =   237
            ToolTipText     =   "Agrega Nueva Sucursal"
            Top             =   1200
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.TextBox Text13 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Height          =   360
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   227
            Top             =   1200
            Width           =   420
         End
         Begin VB.CommandButton Command26 
            Caption         =   "Editar"
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
            Left            =   3180
            Style           =   1  'Graphical
            TabIndex        =   136
            ToolTipText     =   "Modificar Datos de la Sucursal  Activa"
            Top             =   1200
            Width           =   600
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Nuevo"
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
            Left            =   2500
            Style           =   1  'Graphical
            TabIndex        =   135
            ToolTipText     =   "Agrega Nueva Sucursal"
            Top             =   1200
            Width           =   600
         End
         Begin VB.ListBox List1 
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
            Height          =   870
            ItemData        =   "ACLI_GES07.frx":47E3
            Left            =   120
            List            =   "ACLI_GES07.frx":47E5
            TabIndex        =   134
            Top             =   240
            Width           =   3705
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "TRANSPORTISTA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3800
         Left            =   -74760
         TabIndex        =   116
         Top             =   3000
         Width           =   10050
         Begin VB.CommandButton Command32 
            Caption         =   "Nuevo"
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
            Left            =   300
            Style           =   1  'Graphical
            TabIndex        =   231
            ToolTipText     =   "Modificar Datos del Transporte Activo"
            Top             =   3240
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.CommandButton Command31 
            Caption         =   "Editar"
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
            Left            =   300
            Style           =   1  'Graphical
            TabIndex        =   226
            ToolTipText     =   "Modificar Datos del Transporte Activo"
            Top             =   2880
            Visible         =   0   'False
            Width           =   1065
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
            Height          =   330
            Left            =   6795
            TabIndex        =   131
            Top             =   525
            Width           =   175
         End
         Begin VB.TextBox Text45 
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
            Left            =   6000
            TabIndex        =   130
            Top             =   540
            Width           =   750
         End
         Begin VB.CommandButton Command16 
            Caption         =   "Nuevo"
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
            Left            =   7815
            Style           =   1  'Graphical
            TabIndex        =   118
            ToolTipText     =   "Agrega Nuevo Transporte a la Base de Datos"
            Top             =   525
            Width           =   1050
         End
         Begin VB.CommandButton Command18 
            Caption         =   "Editar"
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
            Left            =   8820
            Style           =   1  'Graphical
            TabIndex        =   117
            ToolTipText     =   "Modificar Datos del Transporte Activo"
            Top             =   525
            Width           =   1065
         End
         Begin MSFlexGridLib.MSFlexGrid MSFINSTRUC 
            Height          =   795
            Left            =   1440
            TabIndex        =   229
            Top             =   2880
            Visible         =   0   'False
            Width           =   8475
            _ExtentX        =   14949
            _ExtentY        =   1402
            _Version        =   393216
            Cols            =   5
            ForeColor       =   0
            WordWrap        =   -1  'True
            AllowBigSelection=   -1  'True
            FocusRect       =   2
            ScrollBars      =   2
            AllowUserResizing=   2
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
         Begin VB.TextBox COTEXT3 
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
            MaxLength       =   1
            TabIndex        =   238
            Text            =   " "
            Top             =   1680
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.TextBox Text15 
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
            Left            =   6850
            TabIndex        =   242
            Top             =   1680
            Width           =   3030
         End
         Begin VB.CommandButton BOTSEL3 
            Caption         =   "."
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
            Left            =   1905
            TabIndex        =   239
            Top             =   1680
            Visible         =   0   'False
            Width           =   175
         End
         Begin VB.Label label53 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Código Postal"
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
            Index           =   19
            Left            =   5550
            TabIndex        =   243
            Top             =   1785
            Width           =   1185
         End
         Begin VB.Label label53 
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
            Index           =   18
            Left            =   285
            TabIndex        =   240
            Top             =   1725
            Width           =   1065
         End
         Begin VB.Label label53 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Atte. Sr."
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
            Left            =   5770
            TabIndex        =   194
            Top             =   2565
            Width           =   645
         End
         Begin VB.Label Label69 
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
            Left            =   6510
            TabIndex        =   193
            Top             =   2460
            Width           =   3375
         End
         Begin VB.Label Label67 
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
            Left            =   1430
            TabIndex        =   192
            Top             =   2460
            Width           =   4215
         End
         Begin VB.Label Label66 
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
            Left            =   6855
            TabIndex        =   191
            Top             =   2060
            Width           =   3015
         End
         Begin VB.Label Label65 
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
            Left            =   1430
            TabIndex        =   190
            Top             =   2060
            Width           =   4215
         End
         Begin VB.Label label53 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono(s)"
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
            Index           =   3
            Left            =   285
            TabIndex        =   189
            Top             =   2115
            Width           =   1065
         End
         Begin VB.Label label53 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Fax/Horario"
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
            Left            =   5760
            TabIndex        =   188
            Top             =   2145
            Width           =   1005
         End
         Begin VB.Label label53 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Email"
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
            Left            =   645
            TabIndex        =   187
            Top             =   2525
            Width           =   645
         End
         Begin VB.Label label53 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "C.U.I.T:"
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
            Left            =   6405
            TabIndex        =   129
            Top             =   920
            Width           =   1275
         End
         Begin VB.Label label53 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Número"
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
            Left            =   5310
            TabIndex        =   128
            Top             =   600
            Width           =   585
         End
         Begin VB.Label label53 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre"
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
            Index           =   0
            Left            =   285
            TabIndex        =   127
            Top             =   600
            Width           =   1065
         End
         Begin VB.Label label53 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Domicilio"
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
            Left            =   45
            TabIndex        =   126
            Top             =   1350
            Width           =   1275
         End
         Begin VB.Label label53 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Localidad"
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
            Left            =   5145
            TabIndex        =   125
            Top             =   1430
            Width           =   1275
         End
         Begin VB.Label label53 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Razón Social"
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
            Left            =   -165
            TabIndex        =   124
            Top             =   970
            Width           =   1485
         End
         Begin VB.Label Label61 
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
            Left            =   1430
            TabIndex        =   123
            Top             =   525
            Width           =   3780
         End
         Begin VB.Label Label62 
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
            Left            =   1430
            TabIndex        =   122
            Top             =   920
            Width           =   5370
         End
         Begin VB.Label Label63 
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
            Left            =   1430
            TabIndex        =   121
            Top             =   1325
            Width           =   3900
         End
         Begin VB.Label Label60 
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
            Left            =   7770
            TabIndex        =   120
            Top             =   920
            Width           =   2115
         End
         Begin VB.Label Label64 
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
            Left            =   6510
            TabIndex        =   119
            Top             =   1325
            Width           =   3375
         End
         Begin VB.Label DETEXT3 
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
            Left            =   2100
            TabIndex        =   241
            Top             =   1680
            Visible         =   0   'False
            Width           =   2565
         End
      End
      Begin VB.Frame Frame22 
         Caption         =   "DATOS DE ENTREGA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2340
         Left            =   -74720
         TabIndex        =   105
         Top             =   600
         Width           =   5865
         Begin VB.CommandButton Command33 
            Caption         =   "Guardar Prov."
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
            Left            =   4320
            Style           =   1  'Graphical
            TabIndex        =   236
            ToolTipText     =   "Guardar Provincia y Horario a Sucursal Activa"
            Top             =   1920
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.TextBox COTEXT2 
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
            Left            =   1005
            MaxLength       =   1
            TabIndex        =   233
            Text            =   " "
            Top             =   1920
            Width           =   465
         End
         Begin VB.CommandButton BOTSEL2 
            Caption         =   "."
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
            Left            =   1485
            TabIndex        =   232
            Top             =   1950
            Width           =   175
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
            Left            =   990
            MaxLength       =   12
            TabIndex        =   106
            Top             =   800
            Width           =   1065
         End
         Begin VB.TextBox Text5 
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
            Left            =   990
            MaxLength       =   48
            TabIndex        =   108
            Text            =   " "
            Top             =   1180
            Width           =   4665
         End
         Begin VB.TextBox Text43 
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
            Left            =   990
            MaxLength       =   48
            TabIndex        =   109
            Text            =   " "
            Top             =   420
            Width           =   4665
         End
         Begin VB.TextBox Text44 
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
            Left            =   3030
            MaxLength       =   48
            TabIndex        =   107
            Text            =   " "
            Top             =   800
            Width           =   2625
         End
         Begin VB.TextBox Text35 
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
            Left            =   990
            MaxLength       =   48
            TabIndex        =   110
            Top             =   1580
            Width           =   2475
         End
         Begin VB.TextBox Text37 
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
            Left            =   4350
            TabIndex        =   112
            Top             =   1580
            Width           =   1305
         End
         Begin VB.Label label53 
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
            Index           =   14
            Left            =   -180
            TabIndex        =   235
            Top             =   2000
            Width           =   1065
         End
         Begin VB.Label DETEXT2 
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
            Left            =   1680
            TabIndex        =   234
            Top             =   1950
            Width           =   2565
         End
         Begin VB.Label label53 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cod. Postal"
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
            Left            =   -195
            TabIndex        =   179
            Top             =   885
            Width           =   1065
         End
         Begin VB.Label label53 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Transporte"
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
            Index           =   12
            Left            =   -165
            TabIndex        =   132
            Top             =   1275
            Width           =   1065
         End
         Begin VB.Label label53 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Localidad"
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
            Left            =   1860
            TabIndex        =   115
            Top             =   885
            Width           =   1065
         End
         Begin VB.Label label53 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Calle y Número"
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
            Index           =   10
            Left            =   -165
            TabIndex        =   114
            Top             =   315
            Width           =   1065
         End
         Begin VB.Label label53 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Horario"
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
            Left            =   -285
            TabIndex        =   113
            Top             =   1650
            Width           =   1065
         End
         Begin VB.Label label53 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Zona Reparto"
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
            Index           =   16
            Left            =   3480
            TabIndex        =   111
            Top             =   1485
            Width           =   705
         End
      End
      Begin VB.Frame Frame18 
         BorderStyle     =   0  'None
         Height          =   6250
         Left            =   -74970
         TabIndex        =   78
         Top             =   360
         Width           =   10520
         Begin MSFlexGridLib.MSFlexGrid mfgComposicionDeuda 
            Height          =   1935
            Left            =   120
            TabIndex        =   225
            Top             =   3720
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   3413
            _Version        =   393216
            BackColorFixed  =   16501405
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
         Begin VB.Frame Frame14 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   2460
            Left            =   120
            TabIndex        =   221
            Top             =   3720
            Width           =   10335
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
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
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   5520
               TabIndex        =   223
               Top             =   2040
               Width           =   1695
            End
            Begin VB.Label Label28 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
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
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   8280
               TabIndex        =   222
               Top             =   2040
               Width           =   1695
            End
            Begin VB.Label Label94 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   540
               Left            =   0
               TabIndex        =   224
               Top             =   1920
               Width           =   10335
            End
         End
         Begin MSFlexGridLib.MSFlexGrid mfgCuentaCorriente 
            Height          =   2775
            Left            =   120
            TabIndex        =   220
            Top             =   720
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   4895
            _Version        =   393216
            BackColorFixed  =   16501405
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
         Begin VB.Label label 
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
            Index           =   1
            Left            =   7755
            TabIndex        =   86
            Top             =   350
            Width           =   765
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cheques en Cartera"
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
            Index           =   0
            Left            =   5040
            TabIndex        =   85
            Top             =   160
            Width           =   885
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Límite Crédito"
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
            Index           =   2
            Left            =   120
            TabIndex        =   84
            Top             =   160
            Width           =   765
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Deuda en Cta.Cte."
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
            Index           =   3
            Left            =   2655
            TabIndex        =   83
            Top             =   160
            Width           =   765
         End
         Begin VB.Label LTEXT36 
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
            Left            =   960
            TabIndex        =   82
            Top             =   275
            Width           =   1500
         End
         Begin VB.Label LTEXT37 
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
            Left            =   3450
            TabIndex        =   81
            Top             =   275
            Width           =   1500
         End
         Begin VB.Label LTEXT35 
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
            Left            =   5985
            TabIndex        =   80
            Top             =   275
            Width           =   1500
         End
         Begin VB.Label LTEXT23 
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
            Left            =   8595
            TabIndex        =   79
            Top             =   275
            Width           =   1500
         End
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   6255
         Left            =   30
         TabIndex        =   43
         Top             =   360
         Width           =   10520
         Begin VB.Frame Frame12 
            Caption         =   "DATOS CONTABLES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1650
            Left            =   5260
            TabIndex        =   66
            Top             =   4470
            Width           =   5055
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
               Height          =   300
               Left            =   1995
               TabIndex        =   103
               Top             =   1155
               Width           =   2880
            End
            Begin VB.CommandButton Command11 
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
               Left            =   1725
               Picture         =   "ACLI_GES07.frx":47E7
               TabIndex        =   101
               Top             =   1155
               Width           =   175
            End
            Begin VB.TextBox Text2 
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
               Left            =   945
               MaxLength       =   6
               TabIndex        =   100
               Top             =   1155
               Width           =   780
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
               Left            =   4680
               TabIndex        =   31
               Top             =   800
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
               Index           =   7
               Left            =   1725
               TabIndex        =   29
               Top             =   315
               Width           =   175
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
               Index           =   7
               Left            =   1155
               TabIndex        =   28
               Text            =   " "
               Top             =   315
               Width           =   570
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
               Left            =   2400
               TabIndex        =   30
               Text            =   " "
               Top             =   800
               Width           =   2270
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Grupo Divisional"
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
               Left            =   0
               TabIndex        =   102
               Top             =   1080
               Width           =   855
            End
            Begin VB.Label DETEXT 
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
               Height          =   300
               Index           =   7
               Left            =   1995
               TabIndex        =   76
               Top             =   315
               Width           =   2880
            End
            Begin VB.Label DETEXT 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   6
               Left            =   420
               TabIndex        =   75
               Top             =   1680
               Visible         =   0   'False
               Width           =   4455
            End
            Begin VB.Label Label43 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Grupo"
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
               TabIndex        =   68
               Top             =   420
               Width           =   765
            End
            Begin VB.Label Label42 
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
               Height          =   330
               Left            =   420
               TabIndex        =   67
               Top             =   890
               Width           =   1905
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "DATOS COMERCIALES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2540
            Left            =   5260
            TabIndex        =   60
            Top             =   1720
            Width           =   5055
            Begin VB.CheckBox Check2 
               Appearance      =   0  'Flat
               Caption         =   "Mail Deuda"
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
               Left            =   920
               TabIndex        =   140
               Top             =   1750
               Width           =   1395
            End
            Begin VB.TextBox Text8 
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
               Left            =   3150
               TabIndex        =   138
               Text            =   " "
               Top             =   2080
               Width           =   1710
            End
            Begin VB.TextBox Text20 
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
               Left            =   3150
               TabIndex        =   21
               Text            =   " "
               Top             =   1640
               Width           =   1710
            End
            Begin VB.TextBox Text21 
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
               Left            =   900
               TabIndex        =   20
               Text            =   " "
               Top             =   2065
               Width           =   1485
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
               Left            =   1730
               TabIndex        =   19
               Top             =   1180
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
               Left            =   1730
               TabIndex        =   17
               Top             =   800
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
               Left            =   1730
               TabIndex        =   15
               Top             =   420
               Width           =   175
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
               Left            =   1155
               TabIndex        =   14
               Text            =   " "
               Top             =   420
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
               Left            =   1155
               TabIndex        =   16
               Text            =   " "
               Top             =   800
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
               Index           =   5
               Left            =   1155
               TabIndex        =   18
               Text            =   " "
               Top             =   1180
               Width           =   570
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Lim.Cred. cta.cte"
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
               Left            =   2400
               TabIndex        =   139
               Top             =   2040
               Width           =   705
            End
            Begin VB.Label DETEXT 
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
               Height          =   300
               Index           =   5
               Left            =   1995
               TabIndex        =   74
               Top             =   1180
               Width           =   2880
            End
            Begin VB.Label DETEXT 
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
               Height          =   300
               Index           =   4
               Left            =   1995
               TabIndex        =   73
               Top             =   800
               Width           =   2880
            End
            Begin VB.Label DETEXT 
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
               ForeColor       =   &H00000000&
               Height          =   300
               Index           =   3
               Left            =   1995
               TabIndex        =   72
               Top             =   420
               Width           =   2880
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Límite de Crédito"
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
               Left            =   2040
               TabIndex        =   65
               Top             =   1575
               Width           =   1065
            End
            Begin VB.Label Label18 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Descuento"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   50
               TabIndex        =   64
               Top             =   2180
               Width           =   855
            End
            Begin VB.Label Label19 
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
               Height          =   435
               Left            =   120
               TabIndex        =   63
               Top             =   315
               Width           =   945
            End
            Begin VB.Label Label40 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Condición Pago"
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
               Left            =   0
               TabIndex        =   62
               Top             =   750
               Width           =   1080
            End
            Begin VB.Label Label41 
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
               Height          =   225
               Left            =   -630
               TabIndex        =   61
               Top             =   1280
               Width           =   1710
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "IMPUESTOS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2170
            Left            =   220
            TabIndex        =   55
            Top             =   3950
            Width           =   4845
            Begin VB.CheckBox chkRegimenFacturaCredito 
               Appearance      =   0  'Flat
               Caption         =   "Incluido en el Régimen de Factura de Crédito"
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
               Left            =   960
               TabIndex        =   203
               Top             =   1080
               Width           =   3615
            End
            Begin VB.TextBox Text12 
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
               Height          =   280
               Left            =   2865
               MaxLength       =   1
               TabIndex        =   197
               Text            =   " "
               Top             =   1800
               Width           =   465
            End
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
               Height          =   300
               Left            =   3315
               TabIndex        =   196
               Top             =   1800
               Width           =   175
            End
            Begin VB.CommandButton Command22 
               Caption         =   "Configurar"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   3520
               TabIndex        =   99
               ToolTipText     =   "Datos Percepción I.Brutos Cliente Activo"
               Top             =   1485
               Width           =   1140
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
               Left            =   1260
               TabIndex        =   26
               Top             =   1440
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
               Index           =   1
               Left            =   1260
               TabIndex        =   23
               Top             =   315
               Width           =   175
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
               Height          =   280
               Index           =   1
               Left            =   945
               MaxLength       =   2
               TabIndex        =   22
               Text            =   " "
               Top             =   315
               Width           =   360
            End
            Begin VB.TextBox Text33 
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
               Height          =   280
               Left            =   2310
               MaxLength       =   24
               TabIndex        =   24
               Top             =   315
               Width           =   2355
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
               Height          =   280
               Index           =   2
               Left            =   945
               MaxLength       =   2
               TabIndex        =   25
               Text            =   " "
               Top             =   1440
               Width           =   360
            End
            Begin VB.TextBox Text34 
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
               Height          =   280
               Left            =   2280
               MaxLength       =   15
               TabIndex        =   27
               Text            =   " "
               Top             =   1440
               Width           =   1185
            End
            Begin VB.Label Label23 
               BackStyle       =   0  'Transparent
               Caption         =   "Jurisd."
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
               Left            =   2350
               TabIndex        =   204
               Top             =   1890
               Width           =   615
            End
            Begin VB.Label Label85 
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
               Left            =   3600
               TabIndex        =   199
               Top             =   1800
               Width           =   1050
            End
            Begin VB.Label Label84 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Jurisd."
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
               Left            =   1680
               TabIndex        =   198
               Top             =   1515
               Width           =   1185
            End
            Begin VB.Label DETEXT 
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
               Left            =   945
               TabIndex        =   71
               Top             =   1800
               Width           =   1320
            End
            Begin VB.Label DETEXT 
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
               Height          =   280
               Index           =   1
               Left            =   945
               TabIndex        =   70
               Top             =   665
               Width           =   3720
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Posición IVA"
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
               Left            =   120
               TabIndex        =   59
               Top             =   240
               Width           =   780
            End
            Begin VB.Label Label34 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "CUIT"
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
               Left            =   1500
               TabIndex        =   58
               Top             =   400
               Width           =   750
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Categoría Ing. Brutos"
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
               Left            =   120
               TabIndex        =   57
               Top             =   1320
               Width           =   780
            End
            Begin VB.Label Label39 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Nro Ing. Brutos"
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
               Left            =   1560
               TabIndex        =   56
               Top             =   1320
               Width           =   630
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "AGENDA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2180
            Left            =   220
            TabIndex        =   50
            Top             =   1720
            Width           =   4845
            Begin VB.CommandButton Command98 
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
               Left            =   4440
               TabIndex        =   141
               ToolTipText     =   "Cargar Mail de Contactos"
               Top             =   1035
               Width           =   175
            End
            Begin VB.TextBox Text36 
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
               Left            =   1365
               TabIndex        =   89
               Text            =   " "
               Top             =   1736
               Width           =   3300
            End
            Begin VB.CommandButton Command27 
               Caption         =   "Users"
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
               Left            =   3780
               TabIndex        =   137
               Top             =   1755
               Width           =   855
            End
            Begin VB.TextBox Text19 
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
               Left            =   1365
               TabIndex        =   13
               Text            =   " "
               Top             =   1395
               Width           =   3300
            End
            Begin VB.TextBox Text30 
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
               Left            =   1365
               TabIndex        =   10
               Text            =   " "
               Top             =   315
               Width           =   3300
            End
            Begin VB.TextBox Text31 
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
               Left            =   1365
               TabIndex        =   11
               Text            =   " "
               Top             =   675
               Width           =   3300
            End
            Begin VB.TextBox Text32 
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
               Left            =   1365
               TabIndex        =   12
               Text            =   " "
               Top             =   1035
               Width           =   3090
            End
            Begin VB.Label Label71 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Horario"
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
               Left            =   -630
               TabIndex        =   90
               Top             =   1880
               Width           =   1905
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Contacto"
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
               Left            =   -630
               TabIndex        =   54
               Top             =   1480
               Width           =   1905
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Teléfono"
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
               Left            =   -630
               TabIndex        =   53
               Top             =   400
               Width           =   1905
            End
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Fax"
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
               Left            =   -630
               TabIndex        =   52
               Top             =   760
               Width           =   1905
            End
            Begin VB.Label Label33 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "E-Mail"
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
               Left            =   -630
               TabIndex        =   51
               Top             =   1120
               Width           =   1905
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "DOMICILIO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1580
            Left            =   220
            TabIndex        =   44
            Top             =   100
            Width           =   10095
            Begin VB.TextBox Text24 
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
               Left            =   7980
               TabIndex        =   4
               Text            =   " "
               Top             =   315
               Width           =   1935
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
               Left            =   1800
               TabIndex        =   8
               Top             =   1050
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
               Index           =   0
               Left            =   1365
               MaxLength       =   1
               TabIndex        =   7
               Text            =   " "
               Top             =   1050
               Width           =   465
            End
            Begin VB.TextBox Text25 
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
               Height          =   300
               Left            =   1365
               TabIndex        =   3
               Top             =   315
               Width           =   5295
            End
            Begin VB.TextBox Text26 
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
               Left            =   1365
               TabIndex        =   5
               Text            =   " "
               Top             =   680
               Width           =   1305
            End
            Begin VB.TextBox Text27 
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
               Left            =   3885
               TabIndex        =   6
               Text            =   " "
               Top             =   680
               Width           =   6030
            End
            Begin VB.TextBox Text29 
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
               Left            =   7035
               TabIndex        =   9
               Text            =   " "
               Top             =   1050
               Width           =   2880
            End
            Begin VB.Label Label38 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFF80&
               BackStyle       =   0  'Transparent
               Caption         =   "Nombre de Fantasía"
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
               Left            =   6770
               TabIndex        =   95
               Top             =   225
               Width           =   1110
            End
            Begin VB.Label DETEXT 
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
               Index           =   0
               Left            =   2100
               TabIndex        =   69
               Top             =   1050
               Width           =   4005
            End
            Begin VB.Label Label29 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Localidad"
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
               Left            =   2850
               TabIndex        =   49
               Top             =   770
               Width           =   945
            End
            Begin VB.Label Label21 
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
               Left            =   105
               TabIndex        =   48
               Top             =   1150
               Width           =   1185
            End
            Begin VB.Label Label26 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Calle y Número"
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
               Left            =   120
               TabIndex        =   47
               Top             =   400
               Width           =   1185
            End
            Begin VB.Label Label27 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Código Postal"
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
               Left            =   -630
               TabIndex        =   46
               Top             =   770
               Width           =   1905
            End
            Begin VB.Label Label30 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "País"
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
               Left            =   6405
               TabIndex        =   45
               Top             =   1150
               Width           =   540
            End
         End
      End
      Begin VB.TextBox Text28 
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
         Height          =   4650
         Left            =   -74580
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   1450
         Width           =   9780
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Cortar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   -70130
         TabIndex        =   42
         Top             =   6100
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Copiar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   -68945
         TabIndex        =   33
         Top             =   6100
         Width           =   1170
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Pegar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   -67760
         TabIndex        =   34
         Top             =   6100
         Width           =   1275
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Seleccionar Todo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   -66480
         TabIndex        =   35
         Top             =   6100
         Width           =   1650
      End
      Begin VB.PictureBox Picture4 
         AutoRedraw      =   -1  'True
         Height          =   1590
         Left            =   -72690
         ScaleHeight     =   1530
         ScaleWidth      =   2055
         TabIndex        =   41
         Top             =   2175
         Width           =   2115
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   6250
         Index           =   0
         Left            =   -74970
         TabIndex        =   144
         Top             =   360
         Width           =   10520
         Begin VB.TextBox Text7 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   9075
            TabIndex        =   200
            Top             =   2040
            Width           =   750
         End
         Begin VB.Label label53 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Recargo por Embalaje           %"
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
            Left            =   7320
            TabIndex        =   202
            Top             =   2160
            Width           =   1770
         End
         Begin VB.Label Label52 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Index           =   1
            Left            =   9960
            TabIndex        =   201
            Top             =   2130
            Width           =   180
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   6255
         Index           =   0
         Left            =   -74970
         TabIndex        =   180
         Top             =   360
         Width           =   10520
      End
      Begin VB.Label Label82 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Pedidos Pendientes en Pesos"
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
         Left            =   -74640
         TabIndex        =   195
         Top             =   360
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.Label Label74 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Pedidos Pendientes en Pesos"
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
         Left            =   -68160
         TabIndex        =   176
         Top             =   465
         Width           =   1590
      End
      Begin VB.Label Label73 
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
         Height          =   300
         Left            =   -66360
         TabIndex        =   175
         Top             =   540
         Width           =   1575
      End
      Begin VB.Line Line3 
         X1              =   -64770
         X2              =   -64770
         Y1              =   1455
         Y2              =   6300
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   120
         Left            =   3795
         TabIndex        =   40
         Top             =   4905
         Width           =   15
      End
      Begin VB.Label Label17 
         Caption         =   "Label8"
         Height          =   120
         Left            =   4215
         TabIndex        =   39
         Top             =   3960
         Width           =   15
      End
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
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
      Height          =   285
      Left            =   3360
      TabIndex        =   230
      Top             =   0
      Visible         =   0   'False
      Width           =   7215
      WordWrap        =   -1  'True
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
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
      TabIndex        =   228
      Top             =   0
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Label46 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor Número"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   7125
      TabIndex        =   87
      Tag             =   " "
      Top             =   285
      Width           =   1005
   End
   Begin VB.Label Label45 
      Alignment       =   2  'Center
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
      Height          =   225
      Left            =   0
      TabIndex        =   164
      Top             =   0
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Código de Cuenta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2085
      TabIndex        =   98
      Top             =   240
      Width           =   1110
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   0
      X2              =   12000
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Razón Social"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   200
      TabIndex        =   38
      Top             =   650
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Número de Cuenta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   50
      TabIndex        =   37
      Top             =   200
      Width           =   900
   End
   Begin VB.Menu mnuTransacciones 
      Caption         =   "Transacciones"
      Begin VB.Menu NueCli 
         Caption         =   "Nuevo Cliente"
      End
      Begin VB.Menu GuaCli 
         Caption         =   "Guardar Cambios"
      End
      Begin VB.Menu DeHaCaCli 
         Caption         =   "Deshacer Cambios"
      End
      Begin VB.Menu plh1 
         Caption         =   "-"
      End
      Begin VB.Menu EditCli 
         Caption         =   "Abrir Cliente Existente"
      End
      Begin VB.Menu Baj_Cli 
         Caption         =   "Baja Cuenta Activa"
      End
      Begin VB.Menu Susp_Cli 
         Caption         =   "Suspender Crédito"
      End
      Begin VB.Menu SuspOper 
         Caption         =   "Suspender Operación"
      End
      Begin VB.Menu mnuEliminarCtesSinMovimientos 
         Caption         =   "Suspensión de Clientes Sin Movimientos"
      End
      Begin VB.Menu OrDBCli 
         Caption         =   "Ordenar Clientes por Razón Social"
      End
      Begin VB.Menu plh4 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu ConLiCli 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu BusCli 
         Caption         =   "Búsqueda de Cliente"
      End
      Begin VB.Menu separa4 
         Caption         =   "-"
      End
      Begin VB.Menu LiGeClaCli 
         Caption         =   "Listado General y Clasificado"
      End
      Begin VB.Menu LiCuInCli 
         Caption         =   "Listado de Cuentas Inactivas"
      End
      Begin VB.Menu mnuListadoCteParaIB 
         Caption         =   "Listado de Clientes para Ingresos Brutos"
      End
      Begin VB.Menu Print 
         Caption         =   "Imprimir Ficha"
      End
      Begin VB.Menu php6b 
         Caption         =   "-"
      End
      Begin VB.Menu ObtenerCuitDeDNI 
         Caption         =   "Obtener CUIT/CUIL a partir de DNI"
      End
      Begin VB.Menu ConstanciaAfip 
         Caption         =   "Descarga Constancia de Inscripcion Afip"
      End
      Begin VB.Menu separa1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReportesVarios 
         Caption         =   "Reportes Varios"
      End
   End
   Begin VB.Menu mnuUtilidades 
      Caption         =   "Utilidades"
      Begin VB.Menu DeLiCreCli 
         Caption         =   "Deudas y Límites de Crédito"
      End
      Begin VB.Menu FacPenCob 
         Caption         =   "Facturas Pendientes de Cobranza"
      End
      Begin VB.Menu TabCotizacion 
         Caption         =   "Cotización de Moneda Extranjera"
      End
   End
   Begin VB.Menu ConfCli 
      Caption         =   "Configuración"
      Begin VB.Menu mnuAsociarReporte 
         Caption         =   "Vincular Reporte con Formulario"
      End
      Begin VB.Menu php5 
         Caption         =   "-"
      End
      Begin VB.Menu ConPagCli 
         Caption         =   "Condiciones de Pago"
      End
      Begin VB.Menu VenCli 
         Caption         =   "Vendedores"
      End
      Begin VB.Menu PrZoVeCli 
         Caption         =   "Provincias y Zonas de Venta"
      End
      Begin VB.Menu GruCli 
         Caption         =   "Grupo de Clientes"
      End
      Begin VB.Menu mnuIngresosBrutos 
         Caption         =   "Ingresos Brutos"
         Begin VB.Menu mnuPercepcionIB 
            Caption         =   "Percepción Ingresos Brutos"
         End
         Begin VB.Menu mnuInicializaJursdiccion 
            Caption         =   "Inicializar Jurisdicción desde Provincia"
         End
      End
      Begin VB.Menu mnuTokenAltaClientes 
         Caption         =   "Token Alta de Clientes"
      End
      Begin VB.Menu php6 
         Caption         =   "-"
      End
      Begin VB.Menu ConGenCli 
         Caption         =   "Configuración General"
      End
   End
   Begin VB.Menu menuAccesosDirectos 
      Caption         =   "Accesos Directos"
      Begin VB.Menu LiPreCli 
         Caption         =   "Listas de Precios"
      End
      Begin VB.Menu CuCoCli 
         Caption         =   "Plan de Cuentas Contables"
      End
   End
End
Attribute VB_Name = "ACLI_GES07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHEX$(16)
Dim YACAR%(8), MODIFIC%(8)
Dim GRABATODO%
Dim AGRECLI%
Dim RECONTANT$
Dim FOC6%
Public CambioInstruccionEntrega%
Sub CARDACOBRANZA()
   '
   NCLIE = XVALO(TRIM$(Text1.TEXT))
   If NCLIE < 1 Then Exit Sub
   '
   If REGICLI(NCLIE) < 1 Then Exit Sub
   CONDI$ = "Nume_Cli =" & NCLIE
   Text10(0) = VALOBADA("DEUDOR12", "Ctact_Cob", CONDI$, "NOMESS")
   Text10(1) = VALOBADA("DEUDOR12", "Tele_Cob", CONDI$, "NOMESS")
   Text10(2) = VALOBADA("DEUDOR12", "Mail_Cob", CONDI$, "NOMESS")
   Text10(3) = VALOBADA("DEUDOR12", "DiasRec_Cob", CONDI$, "NOMESS")
   Text10(4) = VALOBADA("DEUDOR12", "DiasPago_Cob", CONDI$, "NOMESS")
   Text10(5) = VALOBADA("DEUDOR12", "MailFactura_Cob", CONDI$, "NOMESS")
   '
   Call CARDA_ANOTA_COBRANZA(NCLIE)
   '
   YACAR%(4) = 1: MODIFIC%(4) = 0
   '
   
End Sub


Sub CargarInstrucciones()

    Me.MSFINSTRUC.Rows = 1
    
    NC = XVALO(Text1)
    NTRANSPORTE% = XVALO(Text45) 'NRO TRANSPORTE
    NSUCURSAL% = XVALO(Text13) 'NRO SUCURSAL
    
    If NC < 1 Then Exit Sub
    
    If NSUCURSAL% < 1 Then Exit Sub
    
    If NTRANSPORTE% < 1 Then Exit Sub

    CONDI$ = "Nume_Cli = " & NC
    CONDI$ = CONDI$ + " AND Nume_Suc = " & NSUCURSAL%
    CONDI$ = CONDI$ + " AND Codi_Tra = " & NTRANSPORTE%
    CONDI$ = CONDI$ + " AND MARBORRA < 1"
    
    Dim MSF As msFlexGrid
    Set MSF = MSFINSTRUC
    MSF.Rows = 1
    If MSF.Rows < 2 Then MSF.Rows = 2
    '
    SQLX$ = "SELECT * FROM INSTRUC_ENTRE WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE " & CONDI$
    SQLX$ = SQLX$ + " ORDER BY INSTRUC_ENTRE_ID  "
    Set rs = READSET(SQLX$)
    With rs
      Do While Not .EOF
        REG& = REG& + 1
        MSF.Rows = REG& + 1
        MSF.TextMatrix(REG&, 1) = XVALO(!INSTRUC_ENTRE_ID)
            INSTRU$ = SAFETEXT$(!INSTRUCCION)
            Label22 = INSTRU$
            MSF.RowHeight(REG&) = Label22.Height + 100
            MSF.TextMatrix(REG&, 2) = SAFETEXT$(!INSTRUCCION)
            MSF.TextMatrix(REG&, 3) = XVALO(!nume_cli)
            MSF.TextMatrix(REG&, 4) = XVALO(!Nume_Suc)
            MSF.TextMatrix(REG&, 5) = XVALO(!CODI_TRA)
        
       .MoveNext
      Loop
    End With

    rs.Close
    Set rs = Nothing

End Sub

Function existeValorEnString%(valor, Cadena$)
  existeValorEnString% = 0
  VALORX = Split(Cadena$, ",")
  For i = 0 To UBound(VALORX)
      If VALORX(i) = SAFETEXT$(valor) Then
          existeValorEnString% = 1
          Exit For
      End If
  Next
End Function

Sub GRABAR_GRILLA_INSTRUCCIONES()
   'grabar las instrucciones
   NTRANSPORTE% = XVALO(Text45) 'NRO TRANSPORTE
   NSUCURSAL% = XVALO(Text13) 'NRO SUCURSAL
   NCLIE = XVALO(Text1)
   On Error GoTo ERROR_GUARDAR
   FLEXCONN.BeginTrans
   'MARCAR MARBORRA -1
   CONDI$ = " Nume_Cli = " & NCLIE
   CONDI$ = CONDI$ + " AND Nume_Suc= " & NSUCURSAL%
   CONDI$ = CONDI$ + " AND Codi_Tra =" & NTRANSPORTE%
   SQLX$ = "Update INSTRUC_ENTRE SET Marborra = 1 WHERE " & CONDI$
   FLEXCONN.Execute (FILTSQL$(SQLX$))
   '
   'UPDATE DE LOS ITEMS EXISTENTES QUE SON LOS QUE TIENEN ID
   For i& = 1 To MSFINSTRUC.Rows - 1
     ID& = XVALO(MSFINSTRUC.TextMatrix(i&, 1))
     If ID& > 0 Then
        INSTRUCCION$ = MSFINSTRUC.TextMatrix(i&, 2)
        SQLX$ = "UPDATE INSTRUC_ENTRE"
        SQLX$ = SQLX$ + " SET Nume_Cli = " & NCLIE
        SQLX$ = SQLX$ + " , Nume_Suc = " & NSUCURSAL%
        SQLX$ = SQLX$ + " , Codi_Tra = " & NTRANSPORTE%
        SQLX$ = SQLX$ + " , Codiempr = " & CodiEmp%
        SQLX$ = SQLX$ + " , Instruccion = '" & INSTRUCCION & "'"
        SQLX$ = SQLX$ + " , MARBORRA = '0'"
        SQLX$ = SQLX$ + " WHERE INSTRUC_ENTRE_ID = " & ID&
        FLEXCONN.Execute (SQLX$)
     End If
  Next i&
   'INSERT DE LOS DATOS QUE ESTAN EN LA GRILLA QUE NO TIENEN ID
   For i& = 1 To MSFINSTRUC.Rows - 1
     ID& = XVALO(MSFINSTRUC.TextMatrix(i&, 1))
     If ID& < 1 Then
        INSTRUCCION$ = MSFINSTRUC.TextMatrix(i&, 2)
        SQLX$ = "INSERT INTO INSTRUC_ENTRE"
        SQLX$ = SQLX$ + " (Nume_Cli,Nume_Suc,Codi_Tra,Codiempr,Instruccion,MARBORRA)"
        SQLX$ = SQLX$ + " VALUES(" & NCLIE
        SQLX$ = SQLX$ + " ," & NSUCURSAL%
        SQLX$ = SQLX$ + " ," & NTRANSPORTE%
        SQLX$ = SQLX$ + " ," & CodiEmp%
        SQLX$ = SQLX$ + " ,'" & INSTRUCCION & "'"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + ")"
        FLEXCONN.Execute (SQLX$)
     End If
  Next i&
  '
  'BORRAR LOS MARBORRA = 1 QUE VAN AA SER SOLO SI ALGUNO SE QUITO.
  CONDI$ = CONDI$ + " AND Marborra = 1"
  Call BORRAREGISTROS("INSTRUC_ENTRE", CONDI$, REGAF&, "NOMESS")
  SQLX$ = "DELETE  FROM INSTRUC_ENTRE WHERE " & CONDI$
  FLEXCONN.Execute (SQLX$)

ERROR_GUARDAR:
    If Err.Number <> 0 Then
       FLEXCONN.RollbackTrans
       Call COMUNI("ERROR AL GUARDAR INSTRUCCIONES DE ENTREGAS" & vbCrLf & Err.Number & " " & Err.Description)
       On Error GoTo 0
    Else
      FLEXCONN.CommitTrans
    End If

End Sub

Sub REMP_COMISIMP()
    ' RUTINA PARA REEMPLAZAR COMILLAS SIMPLES POR UN GUION
     Screen.MousePointer = 11
     Call APERBASDAT("CLIEN")
     '
     FIN& = CantRegistros("DEUDOR12", "NUME_CLI > 0 ")
     SQLX$ = "Select * from Deudor12 "
     Dim CLIENTMP As ADODB.Recordset
     Set CLIENTMP = New ADODB.Recordset
25   On Error Resume Next
     CLIENTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
     End If
     On Error GoTo 0
     '
     J& = 0
     CAMBIO% = 0
     With CLIENTMP
         Do While Not .EOF
           J& = J& + 1
           Call TRACE(20, J&, FIN&)
           For i& = 1 To .Fields.Count
             VALO$ = ""
             VALO$ = SAFETEXT$(.Fields(i& - 1).Value)
             If InStr(VALO$, "'") > 0 Then
                 .Fields(i& - 1) = REPLACAR$(VALO$, "'", "-")
                 CAMBIO% = 1
             End If
           Next i&
           If CAMBIO% > 0 Then
             .Update
             CAMBIO% = 0
           End If
         .MoveNext
         Loop
     End With
     CLIENTMP.Close
     Set CLIENTMP = Nothing
     Screen.MousePointer = 1

End Sub

Sub DEPURADUP()
    ' RUTINA PARA REEMPLAZAR COMILLAS SIMPLES POR UN GUION
     Screen.MousePointer = 11
     Call APERBASDAT("CLIEN")
     '
     FIN& = CantRegistros("DEUDOR12", "NUME_CLI > 0 ")
     SQLX$ = "Select * from Deudor12 "
     Dim CLIENTMP As ADODB.Recordset
     Set CLIENTMP = New ADODB.Recordset
25   On Error Resume Next
     CLIENTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
        GoTo 25
     End If
     On Error GoTo 0
     '
     J& = 0
     CAMBIO% = 0
     With CLIENTMP
        Do While Not .EOF
           J& = J& + 1
           Call TRACE(20, J&, FIN&)
           NCX& = XVALO(!nume_cli)
           QTR& = CantRegistros("DEUDOR12", "NUME_CLI = " & NCX&)
           If QTR& > 1 Then
              .Delete
              CLIENTMP.Close
              GoTo 25
           End If
        .MoveNext
        Loop
     End With
     CLIENTMP.Close
     Set CLIENTMP = Nothing
     Screen.MousePointer = 1
     '
End Sub


Sub MIGRATRANS()
     '
     If ULTREG&("TRANSPOR") > 0 Then
        CONDI$ = "Codi_Tra > 0"
        Call ABRECLIEN
        If CantRegistros&("TRANSPORTES", CONDI$) < 1 Then
           Call COMUNI("FLEXOFT Debe Ahora Actualizar la Tabla de TRANSPORTES")
           Dim TRANSPOR As ADODB.Recordset
           Set TRANSPOR = New ADODB.Recordset
           SQLX$ = "Select * from TRANSPORTES"
25         On Error Resume Next
           TRANSPOR.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
           If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
              On Error GoTo 0
              Call CapturaErrorOpen
              GoTo 25
           End If
           '
           With TRANSPOR
             For J& = 1 To ULTREG&("TRANSPOR")
               XX$ = REGLEIDO("TRANSPOR", J&)
               COD1% = CVI(Mid$(XX$, 1, 2))
               DENO$ = TRIM$(Mid$(XX$, 3, 32))
               RASO$ = TRIM$(Mid$(XX$, 35, 48))
               DOMI$ = TRIM$(Mid$(XX$, 83, 48))
               LOCA$ = TRIM$(Mid$(XX$, 131, 48))
               Cuit$ = TRIM$(Mid$(XX$, 179, 15))
               PIVA% = Asc(Mid$(XX$, 194, 1))
               TELE$ = TRIM$(Mid$(XX$, 195, 48))
               FAXI$ = TRIM$(Mid$(XX$, 243, 48))
               mail$ = TRIM$(Mid$(XX$, 291, 48))
               CTAC$ = TRIM$(Mid$(XX$, 339, 48))
               If COD1% > 0 Then
                 .AddNew
                 !CODI_TRA = COD1%
                 !DENO_TRA = DENO$
                 !RASO_TRA = RASO$
                 !DOMI_TRA = DOMI$
                 !LOCA_TRA = LOCA$
                 !CUIT_TRA = Cuit$
                 !PIVA_TRA = PIVA%
                 !TELE_TRA = TELE$
                 !FAXI_TRA = FAXI$
                 !MAIL_TRA = mail$
                 !CTAC_TRA = CTAC$
                 .Update
               End If
             Next J&
           End With
           TRANSPOR.Close
        End If
     End If
     '
End Sub
Sub ACTUTRANSPOR()
   Exit Sub
'   CONDI$ = "Codi_Tra > 0"
'   Call ABRECLIEN
'   If CantRegistros&("TRANSPORTES", CONDI$) < 1 Then
'      Call SETULTREG&("TRANSPOR", 0)
'           Dim TRANSPOR As ADODB.Recordset
'           Set TRANSPOR = New ADODB.Recordset
'           sqlx$ = "Select * from TRANSPORTES WHERE " & CONDI$
'                 COD1% = XVALO(!CODI_TRA)
'                 DENO$ = SAFETEXT$(!DENO_TRA)
'                 RASO$ = SAFETEXT$(!RASO_TRA)
'                 DOMI$ = SAFETEXT$(!DOMI_TRA)
'                 LOCA$ = SAFETEXT$(!LOCA_TRA)
'                 CUIT$ = SAFETEXT$(!CUIT_TRA)
'                 PIVA% = SAFETEXT$(!PIVA_TRA)
'                 TELE$ = SAFETEXT$(!TELE_TRA)
'                 FAXI$ = SAFETEXT$(!FAXI_TRA)
'                 mail$ = SAFETEXT$(!MAIL_TRA)
'                 CTAC$ = SAFETEXT$(!CTAC_TRA)
'                 Call REPLA(XX$, COD1%, 1, 2)
'                 'HASTA AQUI LLEGUE
'                 'FALTA TERMINAR
'               COD1% = CVI(Mid$(XX$, 1, 2))
'               DENO$ = TRIM$(Mid$(XX$, 3, 32))
'               RASO$ = TRIM$(Mid$(XX$, 35, 48))
'               DOMI$ = TRIM$(Mid$(XX$, 83, 48))
'               LOCA$ = TRIM$(Mid$(XX$, 131, 48))
'               CUIT$ = TRIM$(Mid$(XX$, 179, 15))
'               PIVA% = Asc(Mid$(XX$, 194, 1))
'               TELE$ = TRIM$(Mid$(XX$, 195, 48))
'               FAXI$ = TRIM$(Mid$(XX$, 243, 48))
'               mail$ = TRIM$(Mid$(XX$, 291, 48))
'               CTAC$ = TRIM$(Mid$(XX$, 339, 48))
'               If COD1% > 0 Then
'

End Sub
Sub MIGRASUCU()
     '
     If ULTREG&("SUCENTRE") > 0 Then
        CONDI$ = "Nume_Suc > 0"
        Call ABRECLIEN
        If CantRegistros&("SUCENTRE", CONDI$) < 1 Then
           Call COMUNI("FLEXOFT Debe Ahora Actualizar la Tabla de SUCURSALES")
           Dim SUCUR As ADODB.Recordset
           Set SUCUR = New ADODB.Recordset
           SQLX$ = "Select * from SUCENTRE"
25         On Error Resume Next
           SUCUR.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
           If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
              On Error GoTo 0
              Call CapturaErrorOpen
              GoTo 25
           End If
           '
           With SUCUR
             For J& = 1 To ULTREG&("SUCENTRE")
               XX$ = REGLEIDO("SUCENTRE", J&)
               NUSU% = CVI(Mid$(XX$, 1, 2))
               DENO$ = TRIM$(Mid$(XX$, 3, 62))
               DOMI$ = TRIM$(Mid$(XX$, 65, 48))
               LOCA$ = TRIM$(Mid$(XX$, 113, 48))
               NUCLI% = Asc(Mid$(XX$, 161, 2))
               TELE$ = TRIM$(Mid$(XX$, 163, 32))
               If NUSU% > 0 Then
                 .AddNew
                  !Nume_Suc = NUSU%
                  !deno_suc = DENO$
                  !Domi_Suc = DOMI$
                  !Loca_Suc = LOCA$
                  !nume_cli = NUCLI%
                  !Tele_Suc = TELE$
                  !CODI_TRA = CODT%
                  .Update
               End If
             Next J&
           End With
           SUCUR.Close
        End If
     End If
     '
End Sub


Private Sub Baj_Cli_Click()
   Call Command24_Click
End Sub

Private Sub BOTSEL2_Click()
'        If CantRegistros("PROVINCIA", "ID > 0 ", "NOMESS") > 0 Then
'           If CantRegistros("LOCALIDAD", "ID > 0 ", "NOMESS") > 0 Then
'                CONDI$ = "CODFLEX = '" & COTEXT(0) & "'"
'                PROVLOCA.TXTPROVNUM = VALOBADA("PROVINCIA", "ID", CONDI$, "NOMESS")
'                PROVLOCA.TXTLOCALNUM = ""
'                PROVLOCA.TXTNAMEPROVINCIA = DETEXT(0)
'                PROVLOCA.TXTPROVLET = COTEXT(0)
'                PROVLOCA.TXTLOCALIDAD = Text27
'                PROVLOCA.TXTCODPOSTAL = Text26
'
'                PROVLOCA.Show 1
'                If PROVLOCA.APROBO > 0 Then
'                    DETEXT2 = PROVLOCA.TXTNAMEPROVINCIA
'                    COTEXT2 = PROVLOCA.TXTPROVLET
'                    Text27 = PROVLOCA.TXTLOCALIDAD
'                    Text26 = PROVLOCA.TXTCODPOSTAL
'                End If
'                Exit Sub
'           End If
'        End If
    
        Call SELECHO(ARCHEX$(0))
        VDEVU$ = VALACT1$(ARCHEX$(0))
        If TRIM$(VDEVU$) <> "" Then
           COTEXT2.TEXT = VDEVU$
        End If
End Sub

Private Sub BOTSEL3_Click()
        Call SELECHO(ARCHEX$(0))
        VDEVU$ = VALACT1$(ARCHEX$(0))
        If TRIM$(VDEVU$) <> "" Then
           COTEXT3.TEXT = VDEVU$
        End If
End Sub


Private Sub BusCli_Click()
Call Command20_Click
End Sub

Private Sub CoInBDCli_Click()
Call Command99_Click
End Sub

Private Sub Check1_Click()
   If Check1.Value = 1 Then
     If TRIM$(Text24) = "" Then
       Text24 = TRIM$(Left$(Text3, 8))
     End If
   End If
   MODIFIC%(0) = 1
End Sub

Private Sub Check2_Click()
   MODIFIC%(0) = 1
End Sub

Private Sub CmdBorrar_Click()
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   CmdBorrar.Enabled = False
   
   NSuc% = XVALO(Text13)
   If NSuc% < 1 Then
      Call COMUNI("Debe Seleccionar una Sucursal")
      GoTo 99
   End If
   CONDI$ = "NUME_CLI=" & XVALO(Text1) & " and Nume_Suc=" & NSuc%
   SUCURSAL$ = SAFETEXT$(VALOBADA("SUCENTRE", "Deno_Suc", CONDI$))

   OPX% = COMALTER%("Opciones de Borrado:\\Cancelar\Borrar Sucursal N°: " & NSuc% & " " & SUCURSAL$)
   If OPX% < 2 Then GoTo 99
   Call ACTUREGISTRO("SUCENTRE", "STATUS", CONDI$, -1, REGAF&, "NOMESS")
   'Call CARGALISTA(List1, "SUCENTRE", "DENO_SUC/A62;DOMI_SUC/A48;LOCA_SUC/A48;TELE_SUC/A32;NUME_SUC/I4;CODI_TRA/I4;CODPOSTAL_SUC/A10", "Nume_Cli =" & XVALO(Text1) & " AND STATUS >=0 ")
   
   Text13 = ""
   Call CARGALISTA(List1, "SUCENTRE", "DENO_SUC/A62;DOMI_SUC/A48;LOCA_SUC/A48;TELE_SUC/A32;NUME_SUC/I4;CODI_TRA/I4;CODPOSTAL_SUC/A10", "Nume_Cli =" & XVALO(Text1) & " AND STATUS >=0 ")
   If InStr(EPAC$, "AHP") > 0 Then Call OrdenarListBox(List1)
99 CmdBorrar.Enabled = True
End Sub

Private Sub CMDINSTRUCCION_Click()
   InstrucSuc.ltext1 = XVALO(Me.Text1)
   InstrucSuc.Show 1
End Sub

Private Sub Command11_Click()
   '
   If XVALO(TRIM$(Text1)) > 0 Then
     Call SELECHO("DEUDOR1")
     NCLIE = XVALO(TRIM$(VALACT1$("DEUDOR1")))
     If NCLIE < 1 Then Exit Sub
     '
     NC% = NCLIE
     Text2 = TRIM$(Str$(NC%))
     Text4 = rasocli$(NC%)
   End If
   '
End Sub

Private Sub Command12_Click()
    '
    Command12.Enabled = False
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    NCLIEN = XVALO(Text1)
    If NCLIEN > 0 Then
       'Call FILTERFILE("SUCENTRE", "", 5, MKI$(NC%))
       'URE% = 1 + ULTREG&("!SUCENTRE")
       URE% = 1 + CantRegistros("sucentre", "nume_cli=" & NCLIEN)
       VDEF$ = MKI$(URE%) + Space$(190)
10     X$ = OBJPLA("SUCENTRE", VDEF$)
       If X$ <> "" Then
          XX1$ = TRIM$(Mid$(X$, 3, 62))
          XX2$ = TRIM$(Mid$(X$, 65, 48))
          XX3$ = TRIM$(Mid$(X$, 113, 48))
          XX4$ = TRIM$(Mid$(X$, 161, 32))
          XX5% = CVI(Mid$(X$, 193, 2))          ' TOMA NUMERO DE TRANSPORTE DEL OBJETO
          XX6$ = TRIM$(Mid$(X$, 203, 10))          ' TOMA CODIGO POSTAL
         '
          If XX1$ = "" Or XX2$ = "" Or XX3$ = "" Or XX4$ = "" Then
             Call MENSERR(24, "Información Incompleta - Faltan Datos")
             VDEF$ = X$
             GoTo 10
          End If
          '
'          Y$ = REGBLAN$("SUCENTRE")
'          Call REPLA(Y$, X$, 1, 160)
'          Call REPLA(Y$, MKI$(NC%), 161, 2)
'          Call REPLA(Y$, Mid$(X$, 161, 32), 163, 32)
'          Call REPLA(Y$, MKI$(XX5%), 195, 2)            'CARGA NUMERO DE SUCURSAL
'          Call REGAPP("!SUCENTRE", Y$)
'          Call CIERRARCH("!SUCENTRE")
'          Call FILTERUPDATE("SUCENTRE", "", 5, MKI$(NC%))
          '
          ' Grabacion de SQL
          NUMESUC% = URE%
          Dim SUCUR As ADODB.Recordset
          Set SUCUR = New ADODB.Recordset
          SQLX$ = "Select * from SUCENTRE where Nume_Cli =" & NCLIEN
25        On Error Resume Next
          SUCUR.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
          If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
            On Error GoTo 0
            Call CapturaErrorOpen
            GoTo 25
          End If
          With SUCUR
             If NUMESUC% > 0 Then
                .AddNew
                !Nume_Suc = NUMESUC%
                !nume_cli = NCLIEN
                !deno_suc = XX1$
                !Domi_Suc = XX2$
                !CodPostal_Suc = XX6$
                !Loca_Suc = XX3$
                !Tele_Suc = XX4$
                !CODI_TRA = XX5%
                !Status = 0
                .Update
             End If
          End With
          SUCUR.Close
          
          ' Arma la lista con las sucursales
          'Call CARGALISTA(List1, "SUCENTRE", "DENO_SUC/A62;DOMI_SUC/A48;LOCA_SUC/A48;TELE_SUC/A32;NUME_SUC/I4;CODI_TRA/I4;CODPOSTAL_SUC/A10", "Nume_Cli =" & NCLIEN) 'MAL LA POSICION
          Call CARGALISTA(List1, "SUCENTRE", "DENO_SUC/A62;DOMI_SUC/A48;LOCA_SUC/A48;TELE_SUC/A32;NUME_SUC/I4;CODI_TRA/I4;CODPOSTAL_SUC/A10", "Nume_Cli =" & NCLIEN & " AND STATUS >=0 ")
          If InStr(EPAC$, "AHP") > 0 Then Call OrdenarListBox(List1)
       End If
    End If
    '
    Command12.Enabled = True
    '
End Sub

Private Sub Command16_Click()
   '
   Command16.Enabled = False
   '
   If DERACCE%("ATRANSPO", "Actualizar Base de Datos de Transportes") < 2 Then
      GoTo 99
   End If
   '
   VDEF$ = REGBLAN$("TRANSPOR")
   CONDI$ = "CODI_TRA > 0"
10 PROXI% = 1 + XVALO(VALOBADA("TRANSPORTES", "MAX(CODI_TRA)", CONDI$, "NOMESS"))
   'PROXI% = 1 + ULTREG&("TRANSPOR")
'11 RFF$ = RECFILT$("TRANSPOR", 1, MKI$(PROXI%))
'   If RFF$ <> "" Then
'     PROXI% = PROXI% + 1
'     GoTo 11
'   End If
   Call REPLA(VDEF$, MKI$(PROXI%), 1, 2)
   '
   OBX$ = OBJPLA$("CARGA-TRANSPOR", VDEF$)
   If OBX$ <> "" Then
     If TRIM$(Mid$(OBX$, 3, 32)) = "" Then
       Call MENSERR(24, "Falta Nombre del Transporte")
       VDEF$ = OBX$
       GoTo 10
     End If
     '
     Call BLOQARCH("TRANSPOR")
     PROXI% = 1 + XVALO(VALOBADA("TRANSPORTES", "MAX(CODI_TRA)", CONDI$, "NOMESS"))
'21   RFF$ = RECFILT$("TRANSPOR", 1, MKI$(PROXI%))
'     If RFF$ <> "" Then
'       PROXI% = PROXI% + 1
'       GoTo 21
'     End If
     Call REPLA(OBX$, MKI$(PROXI%), 1, 2)
     Call REGAPP("TRANSPOR", OBX$)
     Call CIERRARCH("TRANSPOR")
     ' Graba en Tabla SQL Transportes.
     COD1% = PROXI%
     DENO$ = TRIM$(Mid$(OBX$, 3, 32))
     RASO$ = TRIM$(Mid$(OBX$, 35, 48))
     DOMI$ = TRIM$(Mid$(OBX$, 83, 48))
     LOCA$ = TRIM$(Mid$(OBX$, 131, 48))
     Cuit$ = TRIM$(Mid$(OBX$, 179, 15))
     PIVA% = Asc(Mid$(OBX$, 194, 1))
     TELE$ = TRIM$(Mid$(OBX$, 195, 48))
     FAXI$ = TRIM$(Mid$(OBX$, 243, 48))
     mail$ = TRIM$(Mid$(OBX$, 291, 48))
     CTAC$ = TRIM$(Mid$(OBX$, 339, 48))
     Dim TRANSPOR As ADODB.Recordset
     Set TRANSPOR = New ADODB.Recordset
     SQLX$ = "Select * from TRANSPORTES"
25   On Error Resume Next
     TRANSPOR.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 25
     End If
     With TRANSPOR
          If COD1% > 0 Then
                 .AddNew
                 !CODI_TRA = COD1%
                 !DENO_TRA = DENO$
                 !RASO_TRA = RASO$
                 !DOMI_TRA = DOMI$
                 !LOCA_TRA = LOCA$
                 !CUIT_TRA = Cuit$
                 !PIVA_TRA = PIVA%
                 !TELE_TRA = TELE$
                 !FAXI_TRA = FAXI$
                 !MAIL_TRA = mail$
                 !CTAC_TRA = CTAC$
                 .Update
          End If
     End With
     '
   End If
99 Command16.Enabled = True
   '
End Sub

Private Sub Command18_Click()
   Command18.Enabled = False
   '
   If DERACCE%("ATRANSPO", "Actualizar Base de Datos de Transportes") < 2 Then
      GoTo 99
   End If
   '
   NTRA = XVALO(Text45)
   If NTRA > 0 Then
     If NTRA < 32767 Then
       NTRAI% = NTRA
       'If ECOARCH$("TRANSPOR", MKI$(NTRAI%)) <> "" Then
       
       'MODIFICADO POR QUE EN GABAL SI GRABA EN UNA EMPRESA
       'NO LE DEJA EDITAR POR QUE NO ENCUENTRA EL ECOARCH
       CONDI$ = "CODI_TRA = " & NTRAI%
       If CantRegistros("TRANSPORTES", CONDI$) > 0 Then
         'ASIGNO CON LA FUNICION SI NO ENCUENTRA EN SQL TRAE DE DAT
         '*********************************************************
         CUITRA$ = CUITTRANS$(NTRAI%)
         DENOTRA$ = DENOTRANS$(NTRAI%)
         RASOTRA$ = RASOTRANS$(NTRAI%)
         DOMITRA$ = DOMITRANS$(NTRAI%)
         LOCATRA$ = LOCATRANS$(NTRAI%)
         TELETRA$ = TELETRANS$(NTRAI%)
         FAXITRA$ = FAXITRANS$(NTRAI%)
         MAILTRA$ = MAILTRANS$(NTRAI%)
         CTACTRA$ = CTACTRANS$(NTRAI%)
         PIVATRA% = PIVATRANS%(NTRAI%)
         '
         'ARMO EL STRING
         '**************
         Call REPLA(VDEF$, MKI$(NTRAI%), 1, 2)
         Call REPLA(VDEF$, DENOTRA$, 3, 32)
         Call REPLA(VDEF$, RASOTRA$, 35, 48)
         Call REPLA(VDEF$, DOMITRA$, 83, 48)
         Call REPLA(VDEF$, LOCATRA$, 131, 48)
         Call REPLA(VDEF$, CUITRA$, 179, 15)
         Call REPLA(VDEF$, Chr$(PIVATRA%), 194, 1)
         Call REPLA(VDEF$, TELETRA$, 195, 48)
         Call REPLA(VDEF$, FAXITRA$, 243, 48)
         Call REPLA(VDEF$, MAILTRA$, 291, 48)
         Call REPLA(VDEF$, CTACTRA$, 339, 48)

         'VDEF$ = FILTERGETRECORD$("TRANSPOR", 1, MKI$(NTRAI%))
10       OBX$ = OBJPLA$("CARGA-TRANSPOR", VDEF$)
         If OBX$ <> "" Then
           If TRIM$(Mid$(OBX$, 3, 32)) = "" Then
             Call MENSERR(24, "Falta Nombre del Transporte")
             VDEF$ = OBX$
             GoTo 10
           End If
           '
           Call BLOQARCH("TRANSPOR")
           Call FILTERPUTRECORD("TRANSPOR", 1, MKI$(NTRAI%), OBX$)
           Call CIERRARCH("TRANSPOR")
           ' Edita y Graba en Tabla SQL Transportes.
           COD1% = NTRAI%
           DENO$ = TRIM$(Mid$(OBX$, 3, 32))
           RASO$ = TRIM$(Mid$(OBX$, 35, 48))
           DOMI$ = TRIM$(Mid$(OBX$, 83, 48))
           LOCA$ = TRIM$(Mid$(OBX$, 131, 48))
           Cuit$ = TRIM$(Mid$(OBX$, 179, 15))
           PIVA% = Asc(Mid$(OBX$, 194, 1))
           TELE$ = TRIM$(Mid$(OBX$, 195, 48))
           FAXI$ = TRIM$(Mid$(OBX$, 243, 48))
           mail$ = TRIM$(Mid$(OBX$, 291, 48))
           CTAC$ = TRIM$(Mid$(OBX$, 339, 48))
'           SQLX$ = "Codi_Tra = " & COD1% & ""
           
          Dim rs1 As ADODB.Recordset
          Set rs1 = New ADODB.Recordset
          SQLX$ = "Select * from TRANSPORTES"
          SQLX$ = SQLX$ + " WHERE Codi_Tra = " & COD1%
30        On Error Resume Next
          rs1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

          If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
            On Error GoTo 0
            Call CapturaErrorOpen
            GoTo 30
          End If
          On Error GoTo 0
        
          '
          With rs1
           Do While Not .EOF
              !DENO_TRA = Left$(TRIM$(DENO$), 128)
              !RASO_TRA = Left$(TRIM$(RASO$), 128)
              !DOMI_TRA = Left$(TRIM$(DOMI$), 128)
              !LOCA_TRA = Left$(TRIM$(LOCA$), 128)
              !CUIT_TRA = Left$(TRIM$(Cuit$), 32)
              !PIVA_TRA = PIVA%
              
              
              !TELE_TRA = Left$(TRIM$(TELE$), 64)
              !FAXI_TRA = Left$(TRIM$(FAXI$), 64)
              !MAIL_TRA = Left$(TRIM$(mail$), 64)
              !MAIL_TRA = Left$(TRIM$(Text10(3)), 128)
              !CTAC_TRA = Left$(TRIM$(CTAC$), 128)
              .Update
              .MoveNext
           Loop
          End With
          rs1.Close
          Set rs1 = Nothing
          
'           Call ACTUREGISTRO("TRANSPORTES", "DENO_TRA", SQLX$, DENO$, REG&)
'           Call ACTUREGISTRO("TRANSPORTES", "RASO_TRA", SQLX$, RASO$, REG&)
'           Call ACTUREGISTRO("TRANSPORTES", "DOMI_TRA", SQLX$, DOMI$, REG&)
'           Call ACTUREGISTRO("TRANSPORTES", "LOCA_TRA", SQLX$, LOCA$, REG&)
'           Call ACTUREGISTRO("TRANSPORTES", "CUIT_TRA", SQLX$, CUIT$, REG&)
'           Call ACTUREGISTRO("TRANSPORTES", "PIVA_TRA", SQLX$, PIVA%, REG&)
'           Call ACTUREGISTRO("TRANSPORTES", "TELE_TRA", SQLX$, TELE$, REG&)
'           Call ACTUREGISTRO("TRANSPORTES", "FAXI_TRA", SQLX$, FAXI$, REG&)
'           Call ACTUREGISTRO("TRANSPORTES", "MAIL_TRA", SQLX$, MAIL$, REG&)
'           Call ACTUREGISTRO("TRANSPORTES", "CTAC_TRA", SQLX$, CTAC$, REG&)
           '
           Call Text45_Change
         End If
         GoTo 99
       End If
     End If
   End If
   '
   Call MENSERR(24, "Imposible Editar Datos Transportista")
   '
99 Command18.Enabled = True
   '
End Sub

Private Sub Command19_Click()
   '
   If List1.ListCount > 0 Then
        Call Command26_Click
        Exit Sub
   End If
   '
   Call ABRECLIEN
   Call COPYSTRU("TRANSPOR", "TRANSPO1")
   Call CARGALISEL("TRANSPO1", "TRANSPORTES", "CODI_TRA/F5;DENO_TRA/A32", "CODI_TRA > 0 ORDER BY CODI_TRA", "NOMESS")
   Call SELECHO("TRANSPO1")
   NTRAN% = XVALO(VALACT1$("TRANSPO1"))
   If NTRAN% > 0 Then
        Text45 = TRIM$(Str$(NTRAN%))
        Label51 = Text45
   End If
   '
End Sub

Private Sub Command22_Click()
   Static CTXX%
   Command22.Enabled = False
   '
   NCLIE = XVALO(TRIM$(Text1.TEXT))
   OpcionesIB% = COMALTER%("Jurisdicción Ingresos Brutos\\Buenos Aires\C.A.B.A.\Misiones\Tucumán")
   If OpcionesIB% < 1 Or OpcionesIB% > 4 Then GoTo 99
   If OpcionesIB% = 1 Then

            If ULTREG&("DAPIBCLN") < 1 Then
              Call GENDAPIBCLI
            End If
            '
            If NCLIE > 0 Then
                '
                If CTXX% < 1 Then
                  Call GRAREG("TACATIBC", Chr$(0) + "Exento Percepcion IB", 1)
                  Call GRAREG("TACATIBC", Chr$(1) + "Sujeto Percepcion IB", 2)
                  Call SETULTREG("TACATIBC", 2)
                  CTXX% = 1
                End If
                '
                RECORD$ = FILTERGETRECORD$("DAPIBCLN", 1, MKS$(NCLIE))
                Call REPLA(RECORD$, MKS$(NCLIE), 1, 4)
                CATIB1% = Asc(Mid$(RECORD$, 5, 1))
                If CATIB1% < 0 Or CATIB1% > 1 Then
                  Call REPLA(RECORD$, Chr$(1), 5, 1)
                  Call REPLA(RECORD$, Space$(32), 10, 32)
                End If
                '
10              OBX$ = OBJPLA$("CARDAPERIB", RECORD$)
                If OBX$ <> "" Then
                  RECORD$ = OBX$
                  CATIB1% = Asc(Mid$(RECORD$, 5, 1))
                  If CATIB1% < 0 Or CATIB1% > 1 Then
                    Call MENSERR(24, "Mal Categoría I.Brutos")
                    GoTo 10
                  End If
                  If CATIB1% = 0 Then
                    If TRIM$(Mid$(RECORD$, 10, 32)) = "" Then
                      Call MENSERR(24, "Mandatorio Indicar Motivo de Exención")
                      GoTo 10
                    End If
                  End If
                  If CATIB1% = 1 Then
                    Call REPLA(RECORD$, Space$(32), 10, 32)
                  End If
                  PORCONVE = CVS(Mid$(RECORD$, 6, 4))
                  If PORCONVE < 0 Or PORCONVE > 100 Then
                    Call MENSERR(24, "% Convenio Multilateral\Fuera de Rango")
                    GoTo 10
                  End If
                  '
                  Call REPLA(RECORD$, MKS$(NCLIE), 1, 4)
                  Call REPLA(RECORD$, USERNAME$, 42, 20)
                  Call REPLA(RECORD$, MKI$(HOY(HOS$)), 62, 2)
                  Call REPLA(RECORD$, MKI$(Int(HORANUM(Time$))), 64, 2)
                  Call FILTERPUTRECORD("DAPIBCLN", 1, MKS$(NCLIE), RECORD$)
                  Call CIERRARCH("DAPIBCLN")
                  '
                End If
            End If
            
    ElseIf OpcionesIB% = 2 Then
        
            Call COPYSTRU("DAPIBCLN", "DAPIBCA")
            
            If ULTREG&("DAPIBCA") < 1 Then
              Call GENDAPIBCA
            End If
            '
            If NCLIE > 0 Then
                '
                If CTXX% < 1 Then
                  Call GRAREG("TACATIBC", Chr$(0) + "Exento Percepcion IB", 1)
                  Call GRAREG("TACATIBC", Chr$(1) + "Sujeto Percepcion IB", 2)
                  Call SETULTREG("TACATIBC", 2)
                  CTXX% = 1
                End If
                '
                RECORD$ = FILTERGETRECORD$("DAPIBCA", 1, MKS$(NCLIE))
                Call REPLA(RECORD$, MKS$(NCLIE), 1, 4)
                CATIB1% = Asc(Mid$(RECORD$, 5, 1))
                If CATIB1% < 0 Or CATIB1% > 1 Then
                  Call REPLA(RECORD$, Chr$(1), 5, 1)
                  Call REPLA(RECORD$, Space$(32), 10, 32)
                End If
                '
20              OBX$ = OBJPLA$("CARDAPERIB/PERCEPCION DE IIBB C.A.B.A.", RECORD$)
                If OBX$ <> "" Then
                  RECORD$ = OBX$
                  CATIB1% = Asc(Mid$(RECORD$, 5, 1))
                  If CATIB1% < 0 Or CATIB1% > 1 Then
                    Call MENSERR(24, "Mal Categoría I.Brutos")
                    GoTo 20
                  End If
                  If CATIB1% = 0 Then
                    If TRIM$(Mid$(RECORD$, 10, 32)) = "" Then
                      Call MENSERR(24, "Mandatorio Indicar Motivo de Exención")
                      GoTo 20
                    End If
                  End If
                  If CATIB1% = 1 Then
                    Call REPLA(RECORD$, Space$(32), 10, 32)
                  End If
                  PORCONVE = CVS(Mid$(RECORD$, 6, 4))
                  If PORCONVE < 0 Or PORCONVE > 100 Then
                    Call MENSERR(24, "% Convenio Multilateral\Fuera de Rango")
                    GoTo 20
                  End If
                  '
                  Call REPLA(RECORD$, MKS$(NCLIE), 1, 4)
                  Call REPLA(RECORD$, USERNAME$, 42, 20)
                  Call REPLA(RECORD$, MKI$(HOY(HOS$)), 62, 2)
                  Call REPLA(RECORD$, MKI$(Int(HORANUM(Time$))), 64, 2)
                  Call FILTERPUTRECORD("DAPIBCA", 1, MKS$(NCLIE), RECORD$)
                  Call CIERRARCH("DAPIBCA")
                  '
                End If
            End If
            
   ElseIf OpcionesIB% = 3 Then
   
            Call COPYSTRU("DAPIBCLN", "DAPIBMI")
            
            If ULTREG&("DAPIBMI") < 1 Then
              Call GENDAPIBMI
            End If
            '
            If NCLIE > 0 Then
                '
                If CTXX% < 1 Then
                  Call GRAREG("TACATIBC", Chr$(0) + "Exento Percepcion IB", 1)
                  Call GRAREG("TACATIBC", Chr$(1) + "Sujeto Percepcion IB", 2)
                  Call SETULTREG("TACATIBC", 2)
                  CTXX% = 1
                End If
                '
                RECORD$ = FILTERGETRECORD$("DAPIBMI", 1, MKS$(NCLIE))
                Call REPLA(RECORD$, MKS$(NCLIE), 1, 4)
                CATIB1% = Asc(Mid$(RECORD$, 5, 1))
                If CATIB1% < 0 Or CATIB1% > 1 Then
                  Call REPLA(RECORD$, Chr$(1), 5, 1)
                  Call REPLA(RECORD$, Space$(32), 10, 32)
                End If
                '
30              OBX$ = OBJPLA$("CARDAPERIB/PERCEPCION IIBB MISIONES", RECORD$)
                If OBX$ <> "" Then
                  RECORD$ = OBX$
                  CATIB1% = Asc(Mid$(RECORD$, 5, 1))
                  If CATIB1% < 0 Or CATIB1% > 1 Then
                    Call MENSERR(24, "Mal Categoría I.Brutos")
                    GoTo 30
                  End If
                  If CATIB1% = 0 Then
                    If TRIM$(Mid$(RECORD$, 10, 32)) = "" Then
                      Call MENSERR(24, "Mandatorio Indicar Motivo de Exención")
                      GoTo 30
                    End If
                  End If
                  If CATIB1% = 1 Then
                    Call REPLA(RECORD$, Space$(32), 10, 32)
                  End If
                  PORCONVE = CVS(Mid$(RECORD$, 6, 4))
                  If PORCONVE < 0 Or PORCONVE > 100 Then
                    Call MENSERR(24, "% Convenio Multilateral\Fuera de Rango")
                    GoTo 30
                  End If
                  '
                  Call REPLA(RECORD$, MKS$(NCLIE), 1, 4)
                  Call REPLA(RECORD$, USERNAME$, 42, 20)
                  Call REPLA(RECORD$, MKI$(HOY(HOS$)), 62, 2)
                  Call REPLA(RECORD$, MKI$(Int(HORANUM(Time$))), 64, 2)
                  Call FILTERPUTRECORD("DAPIBMI", 1, MKS$(NCLIE), RECORD$)
                  Call CIERRARCH("DAPIBMI")
                  '
                End If
            End If
    
    Else
            Call COPYSTRU("DAPIBCLN", "DAPIBTU")
            
            If ULTREG&("DAPIBTU") < 1 Then
              Call GENDAPIBTU
            End If
            '
            If NCLIE > 0 Then
                '
                If CTXX% < 1 Then
                  Call GRAREG("TACATIBC", Chr$(0) + "Exento Percepcion IB", 1)
                  Call GRAREG("TACATIBC", Chr$(1) + "Sujeto Percepcion IB", 2)
                  Call SETULTREG("TACATIBC", 2)
                  CTXX% = 1
                End If
                '
                RECORD$ = FILTERGETRECORD$("DAPIBTU", 1, MKS$(NCLIE))
                Call REPLA(RECORD$, MKS$(NCLIE), 1, 4)
                CATIB1% = Asc(Mid$(RECORD$, 5, 1))
                If CATIB1% < 0 Or CATIB1% > 1 Then
                  Call REPLA(RECORD$, Chr$(1), 5, 1)
                  Call REPLA(RECORD$, Space$(32), 10, 32)
                End If
                '
40              OBX$ = OBJPLA$("CARDAPERIB/PERCEPCION IIBB TUCUMAN", RECORD$)
                If OBX$ <> "" Then
                  RECORD$ = OBX$
                  CATIB1% = Asc(Mid$(RECORD$, 5, 1))
                  If CATIB1% < 0 Or CATIB1% > 1 Then
                    Call MENSERR(24, "Mal Categoría I.Brutos")
                    GoTo 40
                  End If
                  If CATIB1% = 0 Then
                    If TRIM$(Mid$(RECORD$, 10, 32)) = "" Then
                      Call MENSERR(24, "Mandatorio Indicar Motivo de Exención")
                      GoTo 40
                    End If
                  End If
                  If CATIB1% = 1 Then
                    Call REPLA(RECORD$, Space$(32), 10, 32)
                  End If
                  PORCONVE = CVS(Mid$(RECORD$, 6, 4))
                  If PORCONVE < 0 Or PORCONVE > 100 Then
                    Call MENSERR(24, "% Convenio Multilateral\Fuera de Rango")
                    GoTo 40
                  End If
                  '
                  Call REPLA(RECORD$, MKS$(NCLIE), 1, 4)
                  Call REPLA(RECORD$, USERNAME$, 42, 20)
                  Call REPLA(RECORD$, MKI$(HOY(HOS$)), 62, 2)
                  Call REPLA(RECORD$, MKI$(Int(HORANUM(Time$))), 64, 2)
                  Call FILTERPUTRECORD("DAPIBTU", 1, MKS$(NCLIE), RECORD$)
                  Call CIERRARCH("DAPIBTU")
                  '
                End If
            End If
   
   End If
99 Command22.Enabled = True
End Sub

Private Sub Command26_Click()
  '
  EPAC$ = TRIM$(UCase$(EMPREAC$))
  Command26.Enabled = False
  NCLIEN = XVALO(Text1)
  If List1.ListIndex >= 0 Then
     VDEF$ = Space$(224)
     SUC$ = List1.List(List1.ListIndex)
     NUMSUC% = XVALO(Mid$(SUC$, 195, 4)) 'MAL LA POSICION
     If NUMSUC% = 0 Then NUMSUC% = List1.ListIndex + 1
     NUTRAI% = XVALO(Mid$(SUC$, 200, 4)) 'MAL LA POSICION
     RODR$ = Mid$(SUC$, 190)
     CodiPostal$ = Mid$(SUC$, 205, 10)
     VDEF$ = MKI$(NUMSUC%) + Mid$(SUC$, 1, 190) + MKI$(NUTRAI%) + Space(8) + CodiPostal$
     VALANT = MODIFIC%(3): MODIFIC%(3) = 0
10   X$ = OBJPLA$("SUCENTRE", VDEF$)
     MODIFIC%(3) = VALANT
     '
     If X$ <> "" Then
        XX1$ = TRIM$(Mid$(X$, 3, 62))
        XX2$ = TRIM$(Mid$(X$, 65, 48))
        XX3$ = TRIM$(Mid$(X$, 113, 48))
        XX4$ = TRIM$(Mid$(X$, 161, 32))
        XX5% = CVI(Mid$(X$, 193, 2))
        XX6$ = TRIM$(Mid$(X$, 203, 10))
        '
        If XX1$ = "" Or XX2$ = "" Or XX3$ = "" Or XX4$ = "" Then
           Call MENSERR(24, "Información Incompleta - Faltan Datos")
           VDEF$ = X$
           GoTo 10
        End If
        '
'        A% = CVI(Mid$(X$, 1, 2))
'        URE& = ULTREG&("!SUCENTRE")
'        For KKU& = 1 To URE&
'            XX$ = REGLEIDO$("!SUCENTRE", KKU&)
'            B% = CVI(Mid$(XX$, 1, 2))
'            If A% = B% Then
'              Call REPLA(XX$, X$, 1, 158)
'              Call REPLA(XX$, Mid(X$, 161, 32), 163, 32)
'              Call REPLA(XX$, MKI$(XX5%), 195, 2)
'              Call GRAREG("!SUCENTRE", XX$, KKU&)
'              Exit For
'            End If
'        Next KKU&
'        '
'        Call CIERRARCH("!SUCENTRE")
'        Call FILTERUPDATE("SUCENTRE", "", 5, MKI$(NC%))
        '
        ' Grabacion de SQL
        SQLX$ = "Nume_Cli = " & NCLIEN & ""
        SQLX$ = SQLX$ + " And Nume_Suc = " & NUMSUC% & ""
        Call ACTUREGISTRO("SUCENTRE", "Deno_Suc", SQLX$, XX1$, REG&)
        Call ACTUREGISTRO("SUCENTRE", "Domi_Suc", SQLX$, XX2$, REG&)
        Call ACTUREGISTRO("SUCENTRE", "CodPostal_Suc", SQLX$, XX6$, REG&)
        Call ACTUREGISTRO("SUCENTRE", "Loca_Suc", SQLX$, XX3$, REG&)
        Call ACTUREGISTRO("SUCENTRE", "Tele_Suc", SQLX$, XX4$, REG&)
        Call ACTUREGISTRO("SUCENTRE", "Codi_Tra", SQLX$, XX5%, REG&)
        '
        ' Arma la lista con las sucursales
        'Call CARGALISTA(List1, "SUCENTRE", "DENO_SUC/A62;DOMI_SUC/A48;LOCA_SUC/A48;TELE_SUC/A32;NUME_SUC/I4;CODI_TRA/I4;CODPOSTAL_SUC/A10", "Nume_Cli =" & NCLIEN) 'MAL LA POSICION
        Text45 = TRIM$(Str$(XX5%))
        Label51 = Text45
        Call CARGALISTA(List1, "SUCENTRE", "DENO_SUC/A62;DOMI_SUC/A48;LOCA_SUC/A48;TELE_SUC/A32;NUME_SUC/I4;CODI_TRA/I4;CODPOSTAL_SUC/A10", "Nume_Cli =" & NCLIEN & " AND STATUS >=0 ")
        If InStr(EPAC$, "AHP") > 0 Then Call OrdenarListBox(List1)
        '
     End If
  End If
  Command26.Enabled = True
  '
  '\\\OT-07-0010-RG-FIN///
End Sub

Private Sub Command27_Click()
    '
    Command27.Enabled = False
    NC% = XVALO(Text1)
    If NC% < 1 Then GoTo 99
    If rasocli$(NC%) = "" Then GoTo 99
    '
    LANID$ = TRIM$(HORARCLI$(NC%))
    If LANID$ = "" Then
       Call MENSERR(24, "Imposible Cargar usuarios de WEB Services.\Debe Primero Identificar la LAN del Cliente\Mediante el Numero de Serie del Volumen.")
       GoTo 99
    End If
    '
    Call COPYSTRU("FLEXUSR", "WEBUSER")
    Call FILTERFILE("WEBUSER", "!FLEXUSR", 4, LANID$)
    '
320 XUX% = VENTABU%("FLEXUSR")
    If XUX% < 0 Then GoTo 99
    '
    Call FILTERUPDATE("WEBUSER", "!FLEXUSR", 4, LANID$)
    '
99  Call CIERRARCH("*.*")
    Command27.Enabled = True
    '
End Sub

Private Sub Command28_Click()
   '
   Command28.Enabled = False
   '
   Call ABRECLIEN
   CTOTREG& = CantRegistros("DEUDOR12", , "NOMESS")
   MsgBox "Cantidad de Registros en Base de Datos de Clientes: " & CTOTREG&
   '
   Command28.Enabled = True
End Sub

Private Sub Command29_Click()

    NC% = XVALO(Text1)
    If NC% = 0 Then Exit Sub
    '
    If TRIM$(EMAILCLI(NC%)) = "" Or InStr(EMAILCLI(NC%), "@") < 1 Then 'TIENE DIRECCION VALIDA
        Call MENSERR(24, "Imposible Enviar por Mail.\  \Falta Direccion de Email o la Existente es Invalida.")
        Exit Sub
    End If
    '
    If MSFlexGrid1.Rows = 1 Then
        Call MENSERR(24, "Imposible Enviar por Mail.\  \No Existen Pedidos Pendientes.")
        Exit Sub
    End If
    'TITULOS DE DETALLE
    ENCA$ = ""
    Call REPLA(ENCA$, AJUSTD$("Pedido", 10), 1, 10)
    Call REPLA(ENCA$, "Código", 13, 20)
    Call REPLA(ENCA$, "Descripción", 33, 28)
    Call REPLA(ENCA$, AJUSTD$("Saldo", 10), 61, 10)
    ENCA$ = ENCA$ + vbCrLf + vbCrLf + String$(70, "=") + vbCrLf
    '
    'CARGA DE DETALLE
    SQLX$ = " SELECT * FROM PEDDETA "
    TTXX$ = ""
    '
    For ii = 1 To MSFlexGrid1.Rows - 1
         MSFlexGrid1.Row = ii
         NPESEL& = XVALO(MSFlexGrid1.TextMatrix(ii, 0))
         TTXX$ = TTXX$ + " NroPed = " & CStr(NPESEL&) & " OR "
    Next ii
    TTXX$ = Mid$(TTXX$, 1, Len(TTXX$) - 4)
    SQLX$ = SQLX$ & " WHERE (" & TTXX$ & ") "
    SQLX$ = SQLX$ & " AND STATUS<8 AND (CANPED-CANTENT)>0.05"
    SQLX$ = SQLX$ + " ORDER BY NROPED ASC"
    '
    Dim PEDTEMP As ADODB.Recordset
    Set PEDTEMP = New ADODB.Recordset
    PEDTEMP.CursorLocation = adUseClient
26         On Error Resume Next
    PEDTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 26
    End If
    On Error GoTo 0
    '
    TXTX$ = "": TX$ = ""
    With PEDTEMP
      On Error Resume Next
      .MoveFirst
      Do While .EOF = False
         '
         NUPED& = XVALO(!NROPED)
         CODIMAT% = XVALO(!CodiInt)
         CANP# = XVALO(!CanPed)
         CANTEN# = XVALO(!CantEnt)
         SALDO# = CANP - CANTEN
         OBSERVACIONITEM$ = TRIM$(SAFETEXT$(!ObservaItem))
         TX$ = Space(72)
         
         Call REPLA(TX$, FORMATNUM$(NUPED&, "F10.0"), 1, 10)
         Call REPLA(TX$, CODEXT$(CODIMAT%), 13, 20)
         Call REPLA(TX$, DESCRIT0$(CODIMAT%), 33, 28)
         Call REPLA(TX$, FORMATNUM$(SALDO#, "F10.1"), 61, 10)
         TXTX$ = TXTX$ + TX$

         ACUTTYY$ = ""
         If TRIM$(OBSERVACIONITEM$) <> "" Then
            Call FRATEXTO(OBSERVACIONITEM$, 56) ' FRACCIONO EL TEXTO
             For KKI% = 1 To CARETEX% ' RECORRO DESDE 1 A LA CANTIDAD DE FRACCIONES
                TX$ = Space(100)
                If KKI% = 1 Then
                  Call REPLA(TX$, "Nota: ", 5, 6)
                End If
                Call REPLA(TX$, RETEX$(KKI%), 13, 56)
                TXTX$ = TXTX$ + vbCrLf + TX$
            Next KKI%
         End If
         TXTX$ = TXTX$ + vbCrLf
         '
         .MoveNext
      Loop
    End With
    PEDTEMP.Close
    Set PEDTEMP = Nothing
    '
    'CARGA DE ENCABEZADO
    TXTX$ = TXTX$ + vbCrLf + String$(70, "=")
    ENCA2$ = "Sres." + vbCrLf + rasocli(NC%) + vbCrLf + vbCrLf
    ENCA2$ = ENCA2$ + "Nos dirigimos a Uds a fin de hacerles llegar " + vbCrLf
    ENCA2$ = ENCA2$ + "la lista de las productos pendientes de entrega:" + vbCrLf
    textox$ = ENCA2$ + vbCrLf + ENCA$ + vbCrLf + TXTX$ + vbCrLf
    '
    textox$ = textox$ + VALORES$ + vbCrLf + String$(70, "=")
    '
    'ENVIO DE MAIL
    FMAIL09.Text3 = Text1
    FMAIL09.Text6 = "Detalle de Pedidos Pendientes de Entrega"
    FMAIL09.Text5 = textox$
    FMAIL09.Show 1
    'MAILCONFIRMADO$ = InputBox("Email de Envio", "Confirmación de Email", EMAILCLI(NC%))
    'If InStr(MAILCONFIRMADO$, "@") < 1 Then 'TIENE DIRECCION VALIDA
    '    Call MENSERR(24, "Imposible Enviar por Email.\  \Falta Direccion de Email o la Existente es Invalida.")
    '    Exit Sub
    'End If
    'MAI_DESTI$ = rasocli$(NC%)
    'MAI_DIREL$ = MAILCONFIRMADO$
    'MAI_TEXTO$ = TEXTOX$
    'MAI_ASUNT$ = "Detalle de Pedidos Pendientes de Entrega"
    'Call GENERAMAIL
    
End Sub

Private Sub Command30_Click()
   Command30.Enabled = False
   NC = XVALO(Text1)
   If NC > 0 Then
     Call ANOTADATO_COBRANZA(NC, Text14)
     Call CARDA_ANOTA_COBRANZA(NC)
     Text14 = ""
   End If
   
   Command30.Enabled = True
   
End Sub
Sub CARDA_ANOTA_COBRANZA(NCLIE)
   '
   NCLIE = XVALO(Text1.TEXT)
   If NCLIE < 1 Then Exit Sub
   '
   CONDI$ = "Nume_Cli = " & NCLIE
   XX$ = VALOBADA("DEUDOR12", "SEGUI_COB", CONDI$, "NOMESS")
   LONGI& = Len(XX$)
   '
   List4.Clear
   INICIO& = 1
   Do While INICIO& < LONGI&
        FIN& = InStr(INICIO&, XX$, Chr$(13))
        If FIN& > 0 Then
           TTXX$ = Mid$(XX$, INICIO&, FIN& - INICIO&)
           List4.AddItem TTXX$
           INICIO& = FIN& + 2
        Else
           TTXX$ = Mid$(XX$, INICIO&)
           List4.AddItem TTXX$
           Exit Do
        End If
   Loop
   '
   If List4.ListCount > 0 Then
        List4.ListIndex = List4.ListCount - 1
   End If
   '
End Sub
Sub ANOTADATO_COBRANZA(NCLIE, TXTX$, Optional FF%, Optional MODO$)
    '
    
   If NCLIE < 1 Then Exit Sub
   If TRIM$(TXTX$) = "" Then Exit Sub
   '
   HOII% = HOY(HOS$)
   fech$ = HOS$
   If FF% > 0 Then
      fech$ = FECHATEX(FF%) ' SI TRAE FECHA EL PARAMETRO TOMO ESTA
   End If
   
   Call FRATEXTO(TXTX$, 56) ' FRACCIONO EL TEXTO
   '
   For KKI% = 1 To CARETEX% ' RECORRO DESDE 1 A LA CANTIDAD DE FRACCIONES
    TTYY$ = Space$(90)
    If KKI% = 1 Then
       Call REPLA(TTYY$, fech$, 1, 8) ' SI ES LA PRIMERA VEZ GRABA LA FECHA Y USUARIO
       Call REPLA(TTYY$, USERTER$, 69, 15)
    End If
    Call REPLA(TTYY$, RETEX$(KKI%), 11, 56) ' SIEMPRE QUE PASA GUARDA EL CENTRO DEL TEXTO
    ANOTANTE$ = ANOTANTE$ + TTYY$ + vbCrLf
   Next KKI%
   '
   ANOTANTE$ = ANOTANTE$ & String$(85, "=") & vbCrLf
   '
   CONDI$ = "Nume_Cli = " & NCLIE
   '
   ANOTRIOR$ = VALOBADA("DEUDOR12", "SEGUI_COB", CONDI$, "NOMESS") 'LA NOTA GUARDADA EN LA TABLA
   '
   ANOTA$ = ANOTRIOR$ & ANOTANTE$   ' CONCATENO LA ANTERIOR CON LA NUEVA
   Call ACTUREGISTRO("DEUDOR12", "SEGUI_COB", CONDI$, ANOTA$, REGAF&, "NOMESS") 'GRABO
   '
End Sub

Private Sub Command32_Click()
   If XVALO(Text13) < 1 Then Exit Sub 'si no se selecciono ninguna sucursal
   
   Label22.Caption = "" 'este es el label del alto de la fila de grilla
   REG& = Me.MSFINSTRUC.Rows
   'crea un registro nuevo al final
   Me.MSFINSTRUC.Rows = Me.MSFINSTRUC.Rows + 1
   Me.MSFINSTRUC.Row = REG&
   Call POSICIONAR_ULTIMAFILA(Me.MSFINSTRUC)
   EDITINSTRUC.Show 1
   
   If CambioInstruccionEntrega% = 1 Then
      Call GRABAR_GRILLA_INSTRUCCIONES
      CambioInstruccionEntrega% = 0
   Else
      Call cmdEliminaItem(Me.MSFINSTRUC, REG&) 'si cancelo elimina el registro creado
   End If

End Sub

Sub Command33_Click()
    Command33.Enabled = False
    NCLIEN = XVALO(Text1)
    NUMSUC% = XVALO(Text13)
    If NCLIEN < 1 Then
       'Call COMUNI("Falta Nro de Cliente")
       GoTo 99
    End If
    If NUMSUC% < 1 Then
'       Call COMUNI("Debe Seleccionar Una Sucursal")
       GoTo 99
    End If
    
    SQLX$ = "Nume_Cli = " & NCLIEN
    SQLX$ = SQLX$ + " And Nume_Suc = " & NUMSUC% & ""
    Call ACTUREGISTRO("SUCENTRE", "Prov_Cli", SQLX$, COTEXT2.TEXT, REG&)
    Call ACTUREGISTRO("SUCENTRE", "Hora_Ent", SQLX$, Left(Text35.TEXT, 48), REG&)

    
99  Command33.Enabled = True
End Sub

Private Sub Command4_Click()
        Command4.Enabled = False
        Dim Rst As ADODB.Recordset
        SQLX$ = "select nume_cli, raso_cli from deudor12 with(nolock) where stat_cli>0"
        Set Rst = READSET(SQLX$)
        '
        With Rst
          jj& = 0
          Do While Not .EOF
              NCLIEN = XVALO(!nume_cli)
              RASOC$ = SAFETEXT$(!raso_cli)
              '
              XX$ = REGBLAN$("DEUDOR12")
              Call REPLA(XX$, MKS$(NCLIEN), 1, 4)
              Call REPLA(XX$, RASOC$, 5, 28)
              jj& = jj& + 1
              Call GRAREG("DEUDOR12", XX$, jj&)
              .MoveNext
          Loop
        End With
        On Error Resume Next
        Rst.Close
        Set Rst = Nothing
        On Error GoTo 0
        Call SETULTREG("DEUDOR12", jj&)
        Call FRESHECHO("DEUDOR12")

End Sub

Private Sub Command7_Click()
   Call SELECHO("PROVINC")
   VDEVU$ = VALACT1$("PROVINC")
   If TRIM$(VDEVU$) <> "" Then
      Text12 = VDEVU$
   End If
End Sub

Private Sub Command98_Click()
    'ACCESO A LOS MAILS DE CONTACTOS
    Command98.Enabled = False
    COCLI$ = TRIM$(Text6)
    If COCLI$ <> "" Then
      FRMMAILS.Label26 = TRIM$(Text6)
      FRMMAILS.Show 1
    Else
      Call COMUNI("Código de Cliente No Válido o Inexistente")
      GoTo 99
    End If
99  Command98.Enabled = True
End Sub

Private Sub Command99_Click()
   Command99.Enabled = False
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("DEUDOR1")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     R0$ = REGLEIDO$("DEUDOR1", U&)
     NCLIE = CVI(Left$(R0$, 2))
     R1$ = Mid$(R0$, 3, 30)
     If Left$(R1$, 1) = Chr$(64) Then
       R1$ = AJUSTI$(TRIM$(Mid$(R1$, 2)), 30)
     End If
     R1$ = REPLACAR$(R1$, "@", " ")
     R2$ = AJUSTI$(rasocli$(NCLIE), 30)
     If R1$ <> R2$ Then
       HAYINCO% = HAYINCO% + 1
       Call MENSERR(24, "Revisar Cuenta " + TRIM$(Str$(NCLIE)))
     End If
   Next U&
   '
   If HAYINCO% < 1 Then
     Call COMUNI("Control Completo.\No se Detectaron Errores.")
   End If
   '
   Screen.MousePointer = 1
   Command99.Enabled = True
End Sub

Private Sub ConGenCli_Click()
Call Command14_Click
End Sub

Private Sub ConPagCli_Click()
With OPMASCL07
  .Option1(6).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub ConstanciaAfip_Click()

   cuitIngresado$ = TRIM$(Text33)
   If cuitIngresado$ <> "" And VALICUIT%(cuitIngresado$) > 0 Then
        Dim oConstanciaAfip As New DatosEmpresaAfip
        
        oConstanciaAfip.DescargarConstanciaInscripcion (cuitIngresado$)
        Set oConstanciaAfip = Nothing
   Else
        Call MENSERR(24, "Falta Cuit o el Ingresado no es Válido")
   End If
    
End Sub

Private Sub COTEXT2_Change()

    DETEXT2 = TRIM$(ECOARCH$("PROVINC", TRIM$(COTEXT2)))
    MODIFIC%(3) = 1
End Sub

Private Sub COTEXT3_Change()
    DETEXT3 = TRIM$(ECOARCH$("PROVINC", TRIM$(COTEXT3)))
    MODIFIC%(3) = 1
    If XVALO(Text45) > 0 Then
      CONDI$ = " Codi_Tra = " & XVALO(Text45)
      Call ACTUREGISTRO("TRANSPORTES", "Prov_Tra", CONDI$, TRIM$(Left$(COTEXT3, 4)), REGAF&, "NOMESS")
    End If
End Sub


Private Sub CuCoCli_Click()
With OPMASCL07
  .Option1(8).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub DeHaCaCli_Click()
Call Command2_Click
End Sub

Private Sub DeLiCreCli_Click()
    Call AMACLIS07.Command3_Click
End Sub

Private Sub EditCli_Click()
   Call Command9_Click
End Sub

Private Sub exit_Click()
   Call Command3_Click
End Sub

Private Sub FacPenCob_Click()
    FAPECOB07.Show 1
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    SSTab1.Tab = 0
    Call CENTRAFORM(Me)
    Call TRANSLABELS(Name)
    '
    'Picture1(0).BackColor = RGB(40, 78, 98)
    'Me.BackColor = RGB(192, 211, 226)
    'MARCOBARRA.BackColor = RGB(115, 153, 43)
    'BARRATRAZA.BackColor = RGB(157, 204, 72)
    'fecha y usuario solapa de cobranza
    SkinLabel1(0).Caption = Left$(USERNAME, 17)
    HOII% = HOY(HOS$)
    SkinLabel1(2).Caption = " " + HOS$
    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    Call ABRECONEXION

    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    
    If InStr(EPAC$, "MINEPARTS") > 0 Then
       Command4.Visible = True
    End If
    Call GRATITFOR(Caption)
    '
    If InStr(UCase$(App.Path), "COTELEF") > 0 Or InStr(UCase$(App.Path), "ADMIFLEX") > 0 Then
       Label71 = "LAN - ID:"
       VIFO = Text36.FontSize
       Text36.FontSize = 4
       Text36.Width = 2250
       Text36.FontSize = VIFO
    End If
    '
    ARCHEX$(0) = "PROVINC"
    ARCHEX$(1) = "CAPOSIV"
    ARCHEX$(2) = "CARETIB"
    ARCHEX$(3) = "LISTAS"
    ARCHEX$(4) = "CONPAG"
    ARCHEX$(5) = "VENDEDOR"
    ARCHEX$(6) = "ECUECON"
    ARCHEX$(7) = "GRUCOCLI"
    '

    Call CARTAVALI

    '
    If ULTREG&("BKDATCLI") < 1 Then
      LU$ = LUDAT$
      If EXISTE%(LU$ + "DEUDOR1.BAK") > 0 Then
        If EXISTE%(LU$ + "DEUDOR2.BAK") > 0 Then
          Screen.MousePointer = 11
          jj& = 0
          Dim D1 As String * 32
          Dim D2 As String * 128
          N1% = FreeFile
          Open LU$ + "DEUDOR1.BAK" For Random Access Read Write Shared As #N1% Len = 32
          N2% = FreeFile
          Open LU$ + "DEUDOR2.BAK" For Random Access Read Write Shared As #N2% Len = 128
          For k& = 1 To LOF(N1%) / 32
            Get #N1%, k&, D1$
            NC% = CVI(Left$(D1$, 2))
            If NC% > 0 Then
              Get #N2%, k&, D2$
              D3$ = D1$ + D2$
              jj& = jj& + 1
              Call GRAREG("BKDATCLI", D3$, jj&)
            End If
          Next k&
          Call SETULTREG("BKDATCLI", jj&)
          Call CIERRARCH("BKDATCLI")
          Screen.MousePointer = 1
        End If
      End If
    End If
    '

    If Control("AMACLI", "CONSICUI") = "SI" Then 'derecho de modificar CUIT
      If DERACCE%("MODICUIT/NOMESS", "Modificar CUIT") < 2 Then
         Text33.Locked = True
      End If
    End If
    '
    With MSFlexGrid1
        .TextMatrix(0, 0) = "Pedido": .ColWidth(0) = 900:
        .TextMatrix(0, 1) = "O/Compra": .ColWidth(1) = 2300: .ColAlignment(1) = flexAlignLeftTop
        .TextMatrix(0, 2) = "Emisión": .ColWidth(2) = 1000:  .ColAlignment(2) = flexAlignCenterTop
        .TextMatrix(0, 3) = "Cond.Pago": .ColWidth(3) = 2000: .ColAlignment(3) = flexAlignLeftTop
        .TextMatrix(0, 4) = "Observaciones": .ColWidth(4) = 3300
    End With
    With MSFlexGrid2
        .Cols = 9
        .TextMatrix(0, 0) = "Código": .ColWidth(0) = 1400: .ColAlignment(0) = flexAlignLeftTop
        .TextMatrix(0, 1) = "Descripción": .ColWidth(1) = 2000: .ColAlignment(1) = flexAlignLeftTop
        .TextMatrix(0, 2) = "Entrega": .ColWidth(2) = 700: .ColAlignment(2) = flexAlignCenterTop
        .TextMatrix(0, 3) = "Pedido": .ColWidth(3) = 900: .ColAlignment(3) = flexAlignRightTop
        .TextMatrix(0, 4) = "Entregado": .ColWidth(4) = 850: .ColAlignment(4) = flexAlignRightTop
        .TextMatrix(0, 5) = "Saldo": .ColWidth(5) = 900: .ColAlignment(5) = flexAlignRightTop
        .TextMatrix(0, 6) = "Importe Original": .ColWidth(6) = 1275: .ColAlignment(6) = flexAlignRightTop
        .TextMatrix(0, 7) = "Importe Pesos": .ColWidth(7) = 1275: .ColAlignment(7) = flexAlignRightTop
        .TextMatrix(0, 8) = "Observación": .ColWidth(8) = 2000: .ColAlignment(1) = flexAlignLeftTop
    End With

    Call MIGRATRANS
    
    Call MIGRASUCU
    
    ' *** DEFINE ENCABEZADO DEL GRID DE CUENTA CORRIENTE
    Call CARGA_ENCABEZADO(mfgCuentaCorriente, "Fecha/D8;Documento/A18;Número/A18;Debe/F12.2;Haber/F12.2;Saldo/F12.2", GRILLAGRAL)
    For i = 1 To 6
         mfgCuentaCorriente.COL = i: mfgCuentaCorriente.CellFontBold = False: mfgCuentaCorriente.CellFontSize = 9
    Next i
    mfgCuentaCorriente.Font.Size = 9
    '
    ' *** DEFINE ENCABEZADO DEL GRID DE COMPOSICION DE DEUDA
    Call CARGA_ENCABEZADO(mfgComposicionDeuda, "Documento/A19;Número/A18;Emisión/D9;Importe/F12.2;Vencimiento/D10;Saldo ($)/F12.2", GRILLAGRAL)
    For i = 1 To 6
         mfgComposicionDeuda.COL = i: mfgComposicionDeuda.CellFontBold = False: mfgComposicionDeuda.CellFontSize = 9
    Next i
    mfgComposicionDeuda.Font.Size = 9
    
    MODOMONE$ = Left$(TRIM$(UCase$(Control$("GESVENTA", "MODOMONE"))), 8)
    If MODOMONE$ = "MULTIMON" Or MODOMONE$ = "BIMONETA" Then
       '
       ' *** DEFINE ENCABEZADO DEL GRID DE CUENTA CORRIENTE
       Call CARGA_ENCABEZADO(mfgCuentaCorriente, "Fecha/D8;Documento/A14;Número/A18;Debe/F10.2;Haber/F10.2;Saldo/F12.2;Mon/A3;Cotiz/F7.2", GRILLAGRAL)
       For i = 1 To 8
            mfgCuentaCorriente.COL = i: mfgCuentaCorriente.CellFontBold = False: mfgCuentaCorriente.CellFontSize = 9
       Next i
       mfgCuentaCorriente.Font.Size = 9
       '
       ' *** DEFINE ENCABEZADO DEL GRID DE COMPOSICION DE DEUDA EN MODALIDAD BIMONETARIA
       mfgComposicionDeuda.Cols = 1
       mfgComposicionDeuda.Rows = 1
       Call CARGA_ENCABEZADO(mfgComposicionDeuda, "Documento/A14;Número/A18;Emisión/D8;Importe/F12.2;Vencimiento/D8;Saldo ($)/F12.2; Saldo (U$)/F12.2", GRILLAGRAL)
       For i = 1 To 7
            mfgComposicionDeuda.COL = i: mfgComposicionDeuda.CellFontBold = False: mfgComposicionDeuda.CellFontSize = 9
       Next i
       mfgComposicionDeuda.Font.Size = 9
    End If
    
    If Control("FACTURA", "VAPOSIVA") = "SI" Then
        Timer1.Enabled = True 'comentado por que da error en Fervi cuando intenta abrir tabla cuecorri
    End If
    
    Call ABRECABAN
    If InStr(EPAC$, "AHP") > 0 Then
        MSFINSTRUC.Visible = True
        Command31.Visible = True
        Command32.Visible = True
        CMDINSTRUCCION.Visible = True
        COTEXT3.Visible = True
        DETEXT3.Visible = True
        BOTSEL3.Visible = True
        CmdBorrar.Visible = True
        Call PoneEnCeroNull("SUCENTRE", "STATUS", REGAF&, "NOMESS")
    End If
End Sub



'
Private Sub Form_Activate()
   WindowState = 0
   Openforms = DoEvents
End Sub

'
Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub
'
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     Openforms = DoEvents
   Wend
End Sub

Private Sub BOTSEL_Click(Index As Integer)
    If Index = 0 Then
        If CantRegistros("PROVINCIA", "ID > 0 ", "NOMESS") > 0 Then
           If CantRegistros("LOCALIDAD", "ID > 0 ", "NOMESS") > 0 Then
                CONDI$ = "CODFLEX = '" & COTEXT(0) & "'"
                PROVLOCA.TXTPROVNUM = VALOBADA("PROVINCIA", "ID", CONDI$, "NOMESS")
                PROVLOCA.TXTLOCALNUM = ""
                PROVLOCA.TXTNAMEPROVINCIA = DETEXT(0)
                PROVLOCA.TXTPROVLET = COTEXT(0)
                PROVLOCA.TXTLOCALIDAD = Text27
                PROVLOCA.TXTCODPOSTAL = Text26
                
                PROVLOCA.Show 1
                If PROVLOCA.APROBO > 0 Then
                    DETEXT(0) = PROVLOCA.TXTNAMEPROVINCIA
                    COTEXT(0) = PROVLOCA.TXTPROVLET
                    Text27 = PROVLOCA.TXTLOCALIDAD
                    Text26 = PROVLOCA.TXTCODPOSTAL
                End If
                Exit Sub
           End If
        End If
    End If
    
   Call SELECHO(ARCHEX$(Index))
   VDEVU$ = VALACT1$(ARCHEX$(Index))
   If TRIM$(VDEVU$) <> "" Then
      COTEXT(Index).TEXT = VDEVU$
   End If

End Sub

Private Sub Command1_Click()
   GRABATODO% = 1
   Call GRAGENCLI
   Call GRANOTACLI
   Call GRADAENTREGA
   Call GRADACOBRANZA
   GRABATODO% = 0
End Sub

 Sub Command13_Click()
   AMACLIS07.Show 1
End Sub

Sub Command14_Click()
   Command14.Enabled = False
   If DERACCE%("STPCLI", "Configuracion de Setup de Cientes") = 2 Then
   OPMASCL07.Show 1
   End If
   Command14.Enabled = True
End Sub

Private Sub Command15_Click()
   Command15.Enabled = False
   If USNBR% < 100 Then
      Call MENSERR(24, "Estas Opciones Requieren Privilegios de Supervisor")
      GoTo 99
   End If
   '
   OPX% = COMALTER%("Opciones de Trabajo:\\Depurar Duplicados\Registros en Tabla\Vaciar No Utilizados")
   If OPX% = 1 Then
       If COMALTER%("¿ Está Seguro de Querer\Depurar Eventuales Duplicados ?\   \Esta Transaccion NO PODRÁ DESHACERSE !!!\\No, Cancelar\Si, Continuar") <> 2 Then GoTo 99
       UPDATCLI% = 1
       Call DEPURADUP
       Call GENECOCLI     ' DEPURA DUPLICADOS
       Call REMP_COMISIMP ' RUTINA PARA REEMPLAZAR COMILLAS SIMPLES POR UN GUION
       Call COMUNI("Proceso Terminado")
     ElseIf OPX% = 2 Then
       Call ABRECLIEN
       CTOTREG& = CantRegistros("DEUDOR12", , "NOMESS")
       MsgBox "Cantidad de Registros en Tabla de Clientes: " & CTOTREG&
     ElseIf OPX% = 3 Then
       Call LIMPIADEUDOR
   End If
99 Command15.Enabled = True
   '
End Sub

Sub LIMPIADEUDOR()
   '
   Call ABRECONEXION
   Screen.MousePointer = 11
   Dim CLACT%(32767)
   ta$ = "CAJABANC": CA$ = "NUCliEN": GoSub 100
   ta$ = "APLICOBRA": CA$ = "NuClien": GoSub 100
   ta$ = "SACREPEN": CA$ = "NuClien": GoSub 100
   ta$ = "SADEUPEN": CA$ = "NuClien": GoSub 100
   ta$ = "MOVIREPA": CA$ = "Nume_CliE": GoSub 100
   ta$ = "MOVISTO": CA$ = "Nume_CliE": GoSub 100
   ta$ = "PEDENCA": CA$ = "NroCliE": GoSub 100
   ta$ = "PEDDETA": CA$ = "NroCliE": GoSub 100
   ta$ = "PRESUVEN": CA$ = "NroCliE": GoSub 100
   '
   jj& = 0
   For i& = 1 To 32767
     If CLACT%(i&) < 1 Then
       jj& = jj& + 1
     End If
   Next i&
    
   OPX% = COMALTER%("Podrán Recuperarse " & jj& & "Registros\\Cancelar\Continuar")
   If OPX% <> 2 Then Exit Sub
   '
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Piva_Cli", 3, 0)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Cuit_Cli", 10, 24, 1)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Nibr_Cli", 10, 24)
   '
   For i& = 1 To 32767
     Call TRACE(20, i&, 32767)
     If CLACT%(i&) < 1 Then
        NCLIE = i&
        QTR& = CantRegistros("DEUDOR12", "Nume_Cli = " & NCLIE)
        If QTR& < 1 Then GoTo 29
        '
        COCLI$ = CODICLI$(NCLIE)
        If COCLI$ = "" Then GoTo 29       ' NO SE PUEDE RESPALDAR
        '
        CONDI$ = "Codi_Cli = '" & COCLI$ & "'"
        QTRX& = CantRegistros("BSEGCOM", CONDI$, "NOMESS")
        If QTRX& < 1 Then
           Call CreaRegistro("BSEGCOM", "Codi_Cli", COCLI$, "NOMESS")
        End If
        '
        Call ACTUREGISTRO("BSEGCOM", "Nume_Cli", CONDI$, 0, RAFE&)
        '
        If TRIM$(VALOBADA("BSEGCOM", "Raso_Cli", CONDI$)) = "" Then
           Call ACTUREGISTRO("BSEGCOM", "Raso_Cli", CONDI$, rasocli$(NCLIE), RAFE)
        End If
        '
        If TRIM$(VALOBADA("BSEGCOM", "Mail_Cli", CONDI$)) = "" Then
           Call ACTUREGISTRO("BSEGCOM", "Mail_Cli", CONDI$, EMAILCLI$(NCLIE), RAFE)
        End If
        '
        If TRIM$(VALOBADA("BSEGCOM", "Domi_Cli", CONDI$)) = "" Then
           Call ACTUREGISTRO("BSEGCOM", "Domi_Cli", CONDI$, DOMICLI$(NCLIE), RAFE)
        End If
        '
        If TRIM$(VALOBADA("BSEGCOM", "Cpos_Cli", CONDI$)) = "" Then
           Call ACTUREGISTRO("BSEGCOM", "Cpos_Cli", CONDI$, CPOSCLI$(NCLIE), RAFE)
        End If
        '
        If TRIM$(VALOBADA("BSEGCOM", "Loca_Cli", CONDI$)) = "" Then
           Call ACTUREGISTRO("BSEGCOM", "Loca_Cli", CONDI$, LOCACLI$(NCLIE), RAFE)
        End If
        '
        If TRIM$(VALOBADA("BSEGCOM", "Prov_Cli", CONDI$)) = "" Then
           Call ACTUREGISTRO("BSEGCOM", "Prov_Cli", CONDI$, PROVCLI$(NCLIE), RAFE)
        End If
        '
        If TRIM$(VALOBADA("BSEGCOM", "Pais_Cli", CONDI$)) = "" Then
           Call ACTUREGISTRO("BSEGCOM", "Pais_Cli", CONDI$, PAISCLI$(NCLIE), RAFE)
        End If
        '
        If TRIM$(VALOBADA("BSEGCOM", "Tele_cli", CONDI$)) = "" Then
           Call ACTUREGISTRO("BSEGCOM", "Tele_Cli", CONDI$, TEL1CLI$(NCLIE), RAFE)
        End If
        '
        If TRIM$(VALOBADA("BSEGCOM", "Tele_cli2", CONDI$)) = "" Then
           Call ACTUREGISTRO("BSEGCOM", "Tele_Cli2", CONDI$, TEL2CLI$(NCLIE), RAFE)
        End If
        '
        If TRIM$(VALOBADA("BSEGCOM", "Fax_Cli", CONDI$)) = "" Then
           Call ACTUREGISTRO("BSEGCOM", "Fax_Cli", CONDI$, FAXCLI$(NCLIE), RAFE)
        End If
        '
        Call ACTUREGISTRO("BSEGCOM", "Piva_Cli", CONDI$, PIVACLI%(NCLIE), RAFE)
        Call ACTUREGISTRO("BSEGCOM", "Cuit_Cli", CONDI$, CuitCli$(NCLIE), RAFE)
        Call ACTUREGISTRO("BSEGCOM", "Nibr_Cli", CONDI$, NIBRCLI$(NCLIE), RAFE)
        '
        Call BORRAREGISTROS("DEUDOR12", "Nume_Cli = " & NCLIE, RAFE&)
        '
     End If
29 Next i&
   '
   Screen.MousePointer = 1
   UPDATCLI% = 1
   Call GENECOCLI                   ' genera lista de seleccion
   REGLI& = 32767 - CantRegistros("DEUDOR12")
   Call COMUNI("Proceso Completo. Se Liberaron " & REGLI& & " Registros.")
   '
Exit Sub



100 Dim CLIMO As ADODB.Recordset
    Set CLIMO = New ADODB.Recordset
    SQLX$ = "SELECT " + CA$ + " AS NCX FROM " + ta$ + " WHERE " + CA$ + " > 0"
    Set CLIMO = FLEXCONN.Execute(SQLX$)
    With CLIMO
       Do While Not .EOF
         NCY& = XVALO(!NCX)
         If NCY& > 0 Then
           If NCY& <= 32767 Then
             CLACT%(NCY&) = 1
           End If
         End If
       .MoveNext
       Loop
    End With
    CLIMO.Close
    Set CLIMO = Nothing
Return
'
End Sub

Private Sub Command17_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Clipboard.SetText Mid$(Text28.TEXT, Text28.SelStart + 1, Text28.SelLength)
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub

Private Sub Command2_Click()
   Call GRAGENCLI
   Call GRANOTACLI
   Call GRADAENTREGA
   For Index = 0 To 8: YACAR%(Index) = 0: Next Index
   Call CARDATCLIE
End Sub

Private Sub Command20_Click()
   Static ARGU$
   Command20.Enabled = False
10 ARGU$ = UCase$(TRIM$(InputBox$("Buscar", "Búsqueda Cliente por Texto", ARGU$)))
   If ARGU$ <> "" Then
     Screen.MousePointer = 11
     jj& = 0
'     Dim RS As ADODB.Recordset
     SQLX$ = "SELECT * FROM DEUDOR12 WITH(NOLOCK) WHERE NUME_CLI > 0 "
     Set rs = READSET(SQLX$)
     '
     
     If rs.EOF Then
        On Error Resume Next
        rs.Close
        Set rs = Nothing
        On Error GoTo 0
        GoTo 99
     End If
     '
     CONDI$ = ""
     With rs
            For ii% = 1 To .Fields.Count
                If CONDI$ = "" Then
                    CONDI$ = .Fields(ii% - 1).Name & " LIKE '%" & ARGU$ & "%' OR "
                Else
                    CONDI$ = CONDI$ & .Fields(ii% - 1).Name & " LIKE '%" & ARGU$ & "%' OR "
                End If
            Next ii%
     End With
     '
     On Error Resume Next
     rs.Close
     Set rs = Nothing
     On Error GoTo 0
     '
     CONDI$ = Mid$(CONDI$, 1, Len(CONDI$) - 3)
     Call CARGALISEL("SELECLI", "Deudor12", "Nume_Cli/F6.0>;Raso_Cli/A48", CONDI$)
     Call SELECHO("SELECLI/CLIENTES QUE COINCIDENCIAS CON EL ARGUMENTO")
     NCLIX$ = VALACT1$("SELECLI")
     If NCLIX$ <> "" Then
       Text1 = NCLIX$
     End If
   End If
'     Call COPYSTRU("DEUDOR1", "SELECLI")
'     '
'     Call APERBASDAT("CLIEN")
'     Dim Deudor12 As ADODB.Recordset
'     Set Deudor12 = New ADODB.Recordset
'     Deudor12.CursorLocation = adUseClient
'     strsql = "Select * from Deudor12 "
'     Deudor12.Open FILTSQL$(strsql), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
'
'     With Deudor12
'      On Error Resume Next
'      .MoveFirst
'       If Err < 1 Then
'         On Error GoTo 0
'         Do While Not .EOF
'           For KKI% = 1 To .Fields.Count
'             On Error Resume Next
'             VATEX$ = ""
'             VATEX$ = .Fields(KKI% - 1).Value
'             On Error GoTo 0
'             If InStr(VATEX$, ARGU$) > 0 Then
'               NCLIE = !NUME_CLI
'               RACLI$ = !Raso_Cli
'               XX$ = MKI$(Int(NCLIE)) + RACLI$
'               JJ& = JJ& + 1
'               Call GRAREG("!SELECLI", XX$, JJ&)
'               Exit For
'             End If
'           Next KKI%
'         .MoveNext
'         Loop
'       End If
'     End With
'     Deudor12.Close
'     Set Deudor12 = Nothing
'
'     Call SETULTREG("!SELECLI", JJ&)
'     '
'     Screen.MousePointer = 1
'     If JJ& < 1 Then
'         Call MENSERR(24, "No se Encontraron Coincidencias")
'         GoTo 10
'       ElseIf JJ& = 1 Then
'         Text1 = TRIM$(Str$(CVI(Left$(REGLEIDO$("!SELECLI", 1), 2))))
'       Else
'         Call FRESHECHO("!SELECLI")
'         Call SELECHO("!SELECLI")
'         NCLIX$ = VALACT1$("!SELECLI")
'         If NCLIX$ <> "" Then
'           Text1 = NCLIX$
'         End If
'     End If
'   End If
99  Command20.Enabled = True
End Sub

Private Sub Command21_Click()
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If DERACCE%("ALTACLI", "Alta de Nuevo Cliente") < 2 Then
      Exit Sub
   End If
   '
   If EmpresaNoRealizaCambios% > 0 Then Exit Sub

   Call APERBASDAT("CLIEN")
          '(llama a abreconexion y setea la conexionstring al sql)
   'defino el recordset

   Dim Tranclie As ADODB.Recordset
   '
   Command21.Enabled = False
   Call GRAGENCLI
   Call GRANOTACLI
   Call GRADAENTREGA
   '
   Screen.MousePointer = 11
   PRONUCLI& = XVALO(Control$("", "PRONUCLI"))
   If PRONUCLI& < 0 Then
        NROPRO& = XVALO(VALOBADA("deudor12", "max(nume_cli)", "nume_cli>0"))
        GoTo 3
   Else
        If PRONUCLI& > 999999 Then PRONUCLI& = 0
        If InStr(1, EMPREAC$, "MEDITEA") > 0 Then
           If CantRegistros("DEUDOR12", "Nume_Cli = 40001") > 0 Then
                GoTo 2
           Else
                PRONUCLI& = 40000
           End If
        End If
   End If
2  NROPRO& = PRONUCLI&
   If NROPRO& < 0 Then NROPRO& = 0
3  NROPRO& = NROPRO& + 1
   '
   Set Tranclie = FLEXCONN.Execute(FILTSQL$("SELECT Nume_Cli FROM Deudor12 WHERE Nume_Cli = " & NROPRO&))
   If Not (Tranclie.EOF And Tranclie.BOF) Then 'si la encuentra (si el record no esta vacío)
      GoTo 3
   End If
   '
   NCLIE = NROPRO&
   LOCA$ = ""
   CPOSTAL$ = ""
   PROVINCIA$ = ""
   
   ModoAltaCliente% = COMALTER%("Modalidades de Alta de Cliente.\ \Ingreso Manual\u Obteniendo Datos desde Afip a Partir del Cuit (UNICAMENTE)\(Requiere Conexión a Internet)\\Alta Manual\Obtener Datos desde Afip")
   If ModoAltaCliente% < 1 Or ModoAltaCliente% > 2 Then ModoAltaCliente% = 1
   
   If ModoAltaCliente% = 1 Then GoTo 8
   
   ' *** ALTA MEDIANTE CONEXION CON AFIP
   CarpetaSistema$ = App.Path + "\"
   If EXISTE%(CarpetaSistema$ & "toolsafip.EXE") = 0 Then
        Call MENSERR(24, "Imposible Utilizar esta Modalidad.\Faltan Utilitarios de Conexión con Afip.")
        GoTo 8
   End If
   
   cuitIngresado$ = InputBox$("Ingrese el Cuit del Cliente", "Alta de Cliente", "")
   If cuitIngresado$ <> "" Then
   
        If Len(cuitIngresado$) < 9 Then
             Opcion% = COMALTER("Se Ha Detectado el Ingreso de un DNI. ¿Desea Obtener el Cuit Asociado?\\Si, Obtener\No, Continuar")
             If Opcion% = 1 Then
                cuitIngresado$ = ObtieneCuitDesdeDni(cuitIngresado$)
             Else
                GoTo 8
             End If
        End If
   
        Dim oAltaCliente As New DatosEmpresaAfip
        Dim intentosCredencial As Integer
        
        If oAltaCliente.ObtenerDatosEmpresaAfip(cuitIngresado$, response$) Then
           NCLINUE = NCLIE
           DCLINUE$ = oAltaCliente.razonSocial
           COCLINU$ = "CLI-" + TRIM$(Str$(NCLINUE))
           CUTT$ = oAltaCliente.Cuit
           PIVCLI% = oAltaCliente.condicionIva
           DOMI$ = oAltaCliente.domicilioFiscal
           LOCA$ = oAltaCliente.localidadDomicilioFiscal
           PROVINCIA$ = oAltaCliente.provinciaDomicilioFiscal
           CPOSTAL$ = oAltaCliente.codigoPostalDomicilioFiscal
           Set oAltaCliente = Nothing
           
           GoTo 15
        Else
           MensajeAfip$ = "Fallo al Dar de Alta desde Afip\ \"
           Call FRATEXTO(response, 64)
           For KKI% = 1 To CARETEX%: MensajeAfip$ = MensajeAfip$ & RETEX$(KKI%) & "\": Next KKI%
           Screen.MousePointer = 1
           Call COMUNI(MensajeAfip$)
        End If
   Else
        Call MENSERR(24, "Falta Ingresar Cuit")
        Screen.MousePointer = 1
        GoTo 99
   End If
   
   
   ' *** ALTA MEDIANTE INGRESO MANUAL
8  VDEF$ = String$(128, 0)
   Call REPLA(VDEF$, MKS$(NCLIE), 13, 4)
   Call REPLA(VDEF$, Space$(62), 17, 62)
   Call REPLA(VDEF$, Chr$(1), 79, 1)
   Call REPLA(VDEF$, Space$(15), 80, 15)
   Call REPLA(VDEF$, DOMI$, 95, 49)
   '
10 Screen.MousePointer = 1
   OBX$ = OBJPLA$("ALTACLIENACCN", VDEF$)
   If OBX$ = "" Then GoTo 99
   VDEF$ = OBX$
   '
   NCLINUE = CVS(Mid$(OBX$, 13, 4))
   If NCLINUE < 1 Or NCLINUE > 999999 Then
      Call MENSERR(24, "Número de Cuenta no Válido.")
      GoTo 10
   End If
   '
   COCLINU$ = TRIM$(Left$(OBX$, 12))
   If COCLINU$ = "" Then
     COCLINU$ = "CLI-" + TRIM$(Str$(NCLINUE))
     Call REPLA(VDEF$, COCLINU$, 1, 12)
   End If
   '
   DOMI$ = TRIM$(Mid$(OBX$, 95, 49))
   If DOMI$ = "" Then
     If Control("MASTER", "DOMIMAND") = "SI" Then
        Call REPLA(VDEF$, DOMI$, 95, 49)
        Call MENSERR(24, "Debe Ingesar un Domicilio")
        GoTo 10
     End If
   End If
   
   Set Tranclie = FLEXCONN.Execute(FILTSQL$("SELECT Codi_Cli FROM Deudor12 where Codi_cli='" & COCLINU$ & "'"))
   If Not (Tranclie.EOF And Tranclie.BOF) Then
      Call MENSERR(24, "Ingreso no Válido.\Código de Cuenta Pre-existente.")
      GoTo 10
   End If
   
   Set Tranclie = FLEXCONN.Execute(FILTSQL$("SELECT Nume_Cli FROM Deudor12 where Nume_cli=" & NCLINUE))
   If Not (Tranclie.EOF And Tranclie.BOF) Then
      Call MENSERR(24, "Ingreso no Válido.\Número de Cuenta Pre-existente.")
      GoTo 10
   End If
   '
   DCLINUE$ = TRIM$(Mid$(OBX$, 17, 62))
   If Len(DCLINUE$) < 1 Then
      Call MENSERR(24, "Ingreso no Válido.\Falta Razón Social del Cliente")
      GoTo 10
   End If
   If Control$("MASTER", "RASOMAYU") = "SI" Then
      DCLINUE$ = UCase(DCLINUE$)
   End If
   '
   PIVCLI% = Asc(Mid$(OBX$, 79, 1))
   If PIVCLI% < 0 Or PIVCLI% > 7 Then
      Call MENSERR(24, "Ingreso no Válido.\Posición IVA no Figura en Tablas.")
      GoTo 10
   End If
   '
   CUTT$ = TRIM$(Mid$(OBX$, 80, 15))
   If VALICUIT%(CUTT$) <> 1 Then
     If PIVCLI% <> 0 Then
       If PIVCLI% <> 5 Then
         Call MENSERR(24, "Ingreso no Válido.\Número de C.U.I.T. Incorrecto.")
         GoTo 10
       End If
     End If
   End If
   '
   If PIVCLI% = 0 And Len(CUTT$) > 12 Then
        Call MENSERR(24, "El Límite Máximo Para Número de Documento es de 12 Dígitos")
        GoTo 10
   End If
   '
15 If CUTT$ <> "" Then
        If VALICUIT%(CUTT$) = 1 Then CUTT$ = FormatearCuitValido$(CUTT$)
        CONDI$ = "Cuit_Cli = '" & CUTT$ & "' AND nume_cli <> " & NCLIE
        razonSocial$ = TRIM$(VALOBADA("DEUDOR12", "RASO_CLI", CONDI$, "NOMESS"))
        If razonSocial$ <> "" Then
            Screen.MousePointer = 1
            OPX% = COMALTER%("Ciente: " & razonSocial$ & " Posee el Mismo Número de C.U.I.T.\\Cancelar\Grabar Igual")
            If OPX% < 2 Then
              Call PINTATEXT(Text33)
              GoTo 99
            End If
            '
            If DERACCE%("REPICLIE", "Repetir C.U.I.T. en Diferentes Clientes") < 2 Then
              Call PINTATEXT(Text33)
              Exit Sub
            End If
               
        End If
   End If
   
   Screen.MousePointer = 11
   
20 If Err < 1 Then
     On Error GoTo 0
     ABUS$ = "Nume_Cli = " & NCLINUE
   End If
   '
   Dim Rdeudor12 As ADODB.Recordset
   Set Rdeudor12 = New ADODB.Recordset
   Rdeudor12.CursorType = adOpenKeyset
   Rdeudor12.LockType = adLockPessimistic
   Rdeudor12.Open FILTSQL$("Select * from Deudor12 where nume_cli=" & NCLINUE), FLEXCONN, , , adCmdText
   If Not (Rdeudor12.EOF And Rdeudor12.BOF) Then
      GoTo 20
   End If
   With Rdeudor12
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CODIEMPR = CodiEmp%
      On Error GoTo 0
      !nume_cli = TRIM$(Str$(NCLINUE))
      !raso_cli = DCLINUE$
      !Codi_Cli = COCLINU$
      !Cuit_Cli = CUTT$
      !Piva_Cli = PIVCLI%
      !domi_cli = DOMI$
      !loca_Cli = LOCA$
      !Cpos_Cli = CPOSTAL$
      !Prov_Cli = PROVINCIA$
      !Stat_Cli = 1
      '
      On Error Resume Next
      .Update
      If Err Then
         Call MENSERR(24, "Alta no Procesada Correctamente")
      End If
      On Error GoTo 0
   End With
   '
   Rdeudor12.Close
   Set Rdeudor12 = Nothing
   If InStr(EPAC$, "MINEPARTS") < 1 And InStr(EPAC$, "AHP") > 0 Then
     CONDI$ = "nume_cli=" & NCLINUE
     Call DuplicarRegistroDeUnaBaseAOtra("AHP", "Deudor12", "MINEPARTS", "Deudor12", CONDI$)
   End If
   'UPDATCLI% = 1
   'Call GENECOCLI
   
    Call CARGALISEL("DEUDOR1", "Deudor12", "Nume_Cli/F6.0>;Raso_Cli/A48", "Stat_Cli >= 0 and Nume_Cli>0 ORDER BY Raso_Cli")
    Call CARGALISEL("INDICLIE", "Deudor12", "Codi_Cli/A12<;Raso_Cli/A50", "Stat_Cli >= 0 AND Codi_Cli <> '' ORDER BY Raso_Cli")
        CLINUE$ = REGBLAN$("DEUDOR12")
    Call REPLA(CLINUE$, MKS$(NCLINUE), 1, 4)
    Call REPLA(CLINUE$, DCLINUE$, 5, 28)
    Call REGAPP("DEUDOR12", CLINUE$)
    Call FILESORT("DEUDOR12", "", 1, 1)
   
   
'   If LONGINT% < 1 Then If EXISTE%(LUDAT$ + "DEUDORL1.RFS") > 0 Then LONGINT% = 1
'   ADEU$ = "DEUDOR1": If LONGINT% > 0 Then ADEU$ = "DEUDORL1"
'   REBLA$ = REGBLAN$(ADEU$)
'   REBLB$ = REGBLAN$("INDICLIE")
'   NCLIE1& = NCLINUE%
'   If NCLIE1& >= 0.5 Then
'      EL1$ = REBLA$
'      If LONGINT% > 0 Then
'           Call REPLA(EL1$, MKS$(NCLIE1&), 1, 4): PINI% = 5
'         Else
'           NCLI1% = NCLIE1&
'           Call REPLA(EL1$, MKI$(NCLI1%), 1, 2): PINI% = 3
'      End If
'      Call REPLA(EL1$, DCLINUE$, PINI%, 64)
'      Call FILTERPUTRECORD(ADEU$, 1, Left$(EL1$, PINI% - 1), EL1$)
'      '
'      EL2$ = REBLB$
'      Call REPLA(EL2$, COCLINU$, 1, 12)
'      Call REPLA(EL2$, DCLINUE$, 13, 50)
'      If LONGINT% > 0 Then
'           Call REPLA(EL2$, MKS$(NCLIE1&), 61, 4)
'         Else
'           NCLI1% = NCLIE1&
'           Call REPLA(EL2$, MKI$(NCLI1%), 63, 2)
'      End If
'      Call FILTERPUTRECORD("INDICLIE", 1, TRIM$(Left$(EL2$, 12)), EL2$)
'      Call FRESHECHO(ADEU$)
'      Call FRESHECHO("INDICLIE")
'    End If
'    '
   '
   AGRECLI% = 1
   Text1.TEXT = TRIM$(Str$(NCLINUE))
   SSTab1.Tab = 0
   Call CARDATCLIE
   Call CARTAVALI
   '
   For Index = 0 To 7
     If Index <> 1 Then
       VAPRE$ = TRIM$(Control$("AMACLI", ARCHE$(Index)))
       If ModoAltaCliente% = 2 And Index = 0 Then GoTo 70 '  OBVIA LA PROVINCIA CUANDO OBTIENE LOS DATOS DE AFIP
       COTEXT(Index) = "": DETEXT(Index) = ""
       If VAPRE$ <> "" Then
         TV11% = TIPFIELD(ARCHE$(Index), 1)
         LF11% = LENFIELD(ARCHE$(Index), 1)
         ARGU$ = CBIN$(VAPRE$, TV11%, LF11%)
         ECO$ = TRIM$(ECOARCH$(ARCHE$(Index), ARGU$))
         If ECO$ <> "" Then
             COTEXT(Index).TEXT = VAPRE$
             DETEXT(Index).Caption = ECO$
           Else
             COTEXT(Index).TEXT = ""
             DETEXT(Index).Caption = ""
         End If
       End If
     End If
70 Next Index
   '
   Text3 = DCLINUE$
   Text25 = DOMI$
   If PROVINCIA$ <> "" Then COTEXT(0) = PROVINCIA$
   COTEXT(1) = TRIM$(Str$(PIVCLI%))
   Text33 = TRIM$(CUTT$)
   '
   Call Command1_Click
   '
'   If XVALO(COTEXT(1)) > 0 And PIVCLI% > 0 Then '  OBIVA A LOS CONSUMIDORES FINALES
'        OpcionDescargarConstancia% = COMALTER%("¿Desea Descargar Constancia de Inscripción Afip?\\No, Continuar\Si, Descargar")
'        If OpcionDescargarConstancia% = 2 Then
'             Dim oConstanciaAfip As New DatosEmpresaAfip
'
'             cuitIngresado$ = TRIM$(Text33)
'             oConstanciaAfip.DescargarConstanciaInscripcion (cuitIngresado$)
'
'        End If
'   End If
   
   Screen.MousePointer = 1
   '
99 Command21.Enabled = True
   '
End Sub

Private Sub Command23_Click()
   SSTab1.Tab = 0
   Text20.TEXT = " Suspendido"
   Text20.Refresh
End Sub

Private Sub Command24_Click()
   '
   Command24.Enabled = False
   Screen.MousePointer = 11
   
   If EmpresaNoRealizaCambios% > 0 Then GoTo 99

   '
   Call GRAGENCLI
   Call GRANOTACLI
   Call GRADAENTREGA
   '
   If DERACCE%("BAJACLI", "Baja de Cuenta de Cliente") < 2 Then
     GoTo 99
   End If
   '
   NCLIE = Abs(XVALO(TRIM$(Text1.TEXT))) ' SE QUITA EL (-1) PORQUE ESO ES PARA PROVEEDORES
   If Abs(NCLIE) < 1 Or REGICLI(NCLIE) < 1 Then
      RESP1% = 2
      GoTo 90
   End If
   '
   If RESP1% = 0 Then
    RESP1% = ALTERNA%("Dar de Baja Cuenta de Cliente.\Restauración de Cuenta de Cliente.")
    If RESP1% < 1 Or RESP1% > 2 Then
        GoTo 99
    End If
   End If
   '
   If RESP1% = 1 Then 'SI QUIERE ANULAR
     If XVALO(Text1.TEXT) > 0 Then
       If XVALO(Text1.TEXT) <= 999999 Then
         NCLIE = XVALO(Text1.TEXT)
         DCLI$ = TRIM$(Text6)
         TTXX$ = "La Baja de un Cliente\Puede Producir Inconsistencias en la Base de Datos\en Caso de que la Cuenta Registre\Facturación, Cobranza o Pedidos."
         TTXX$ = TTXX$ + "\  \Puede Ahora Cancelar la Operación\o Confirmar la Baja de la Cuenta:\'" + TRIM$(Text3.TEXT) + "'"
         If COMALTER%(TTXX$ + "\\Cancelar Operación\Confirmar Baja") = 2 Then
            strsql = ("Update Deudor12 set stat_cli= -1 where nume_cli= ") & NCLIE
             FLEXCONN.Execute (FILTSQL$(strsql))
            Call GRANOTACLI
            '
            UPDATCLI% = 1
            Call GENECOCLI
'            If LONGINT% < 1 Then If EXISTE%(LUDAT$ + "DEUDORL1.RFS") > 0 Then LONGINT% = 1
'            ADEU$ = "DEUDOR1": If LONGINT% > 0 Then ADEU$ = "DEUDORL1"
'            If LONGINT% > 0 Then
'                 ABUS$ = MKS$(NCLIE&)
'               Else
'                 NCLI1% = NCLIE&
'                 ABUS$ = MKI$(NCLI1%)
'            End If
'            Call FILTERDELETE(ADEU$, 1, ABUS$)
'            '
'            Call FILTERDELETE("INDICLIE", 1, DCLI$)
'            Call FRESHECHO(ADEU$)
'            Call FRESHECHO("INDICLIE")
'            '
            Text1.TEXT = ""
            Call BLANFORMU
            For KKI% = 0 To 4: MODIFIC%(KKI%) = 0: YACAR%(KKI%) = 0: Next KKI%
            GRABATODO% = 0
            Call BAJALDISCO
         End If
       End If
     End If
   End If
   '
90 If RESP1% = 2 Then 'SI QUIERE RESTAURAR
'     Call GENCLIANU
'     Call BAJALDISCO
'     RPA$ = ""
'     RPA$ = REGSEL$("INDICLIA/INDICE DE CLIENTES DE BAJA") 'PROV ANULADOS
     Call CARGALISEL("INDICLIBA", "DEUDOR12", "CODI_CLI/A12;RASO_CLI/A50;NUME_CLI/F6.0", "Stat_Cli=-1")
     Call SELECHO("INDICLIBA/INDICE DE CLIENTES DE BAJA")
     RPA$ = VALACT2$("INDICLIBA")
     If Len(RPA$) > 4 Then
       'NCLIE1 = CVI(Mid$(RPA$, 63, 2))
       NCLIE1 = XVALO(Mid$(RPA$, 51))
       If NCLIE1 > 0 Then
         NC = NCLIE1
         strsql = ("Update Deudor12 set stat_cli= 1 where nume_cli= ") & NC
         FLEXCONN.Execute (FILTSQL$(strsql))
         UPDATCLI% = 1
         Call GENECOCLI
         Text1.TEXT = TRIM$(Str$(NC))
       End If
     End If
   End If
   '
99 Screen.MousePointer = 1
   Command24.Enabled = True
   '
End Sub

Private Sub Command25_Click()
    PRINTFORM
End Sub

Private Sub Command3_Click()
   '
   If SSTab1.Tab = 5 Then SSTab1.Tab = 0
   Call GRAGENCLI
   Call GRANOTACLI
   Call GRADAENTREGA
   Call CARDATCLIE
   '
   Call CIERRARCH("*.*")
   If AGRECLI% > 0 Then
      'Call SORTCLI(1)
      AGRECLI% = 0
   End If
   Call BAJALDISCO
   '
   'VALIDA DE DONDE VIENE PARA QUE SI ES AMACLI O AMAPRO HAGA UN CALL FINAL
   If UCase$(App.EXEName) = "ARCHIG07" Then
      On Error Resume Next
      Unload Me
      MARCH_GES07.Enabled = True
      MARCH_GES07.SetFocus
      On Error GoTo 0
   Else
      Call SETULTREG("!SELPRBX", 0) 'de busqueda por CUIT, Tel, Direccio, ETC
      Call CIERRARCH("*.*")
      Call FINAL("")
   End If
   '
'   If APLINVO$ <> "" Then
'    Hide
'   Else
'    On Error Resume Next
'    Unload Me
'    MARCH_GES07.Enabled = True
'    MARCH_GES07.SetFocus
'    On Error GoTo 0
'   End If
   '
End Sub

Private Sub Command5_Click()
   Text28.SelStart = 0
   Text28.SelLength = Len(Text28.TEXT)
   Text28.SetFocus
End Sub

Private Sub Command6_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Clipboard.GetText + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub

Private Sub Command8_Click()
   Clipboard.SetText Mid$(Text28.TEXT, Text28.SelStart + 1, Text28.SelLength)
End Sub

Private Sub Command9_Click()
   '
   Call GRAGENCLI
   Call GRANOTACLI
   Call GRADAENTREGA
   Call GRADACOBRANZA
   '
   'Call SELECHO("DEUDOR12")
   'NCLIE = XVALO(TRIM$(VALACT1$("DEUDOR12")))
   NCLIE = ListaSeleccionClientes("", True)
   If NCLIE < 1 Then Exit Sub
   '
   NC = NCLIE
   Text1.TEXT = TRIM$(Str$(NC))
   'Call CARDATCLIE
   '
End Sub


Private Sub Label53_Click(Index As Integer)
'    Unload VENTABNEW
'    Campos$ = "SUCURSAL/f8;Desc.Suc/A18;CodTRansp/F8;TRANSPORTE/A30;ID_iNSTR/f8;INSTRUCCION/A64;ULT/A0"
'
'    VENTANA.Inicializar = 0
'    VENTANA.Campox = Campos$
'    VENTANA.AgregaRegistro = 0
'    VENTANA.UtilizaFormula = 1
'    VENTANA.HabilitaBorrar = 1
'    VENTANA.HabilitaInsertar = 0
'    VENTANA.NoMostrarBotonTilde = 0
'    VENTANA.TITULO = "PARTE DIARIO DE PRODUCCIÓN"
'    VENTANA.ColEcoCampo(1) = "NUME_SUC;DENO_SUC;SUCENTRE" ';LISTA = 'MATERIALES ENVAPLAST'"
'    VENTANA.ColEcoCampo(3) = "CODI_TRA;DENO_TRA;SUCUTRANSP;NUME_CLI =" & XVALO(Text1) 'OJO ES UNA VISTA
'    VENTANA.ColEcoCampo(5) = "INSTRUC_ENTRE_ID;INSTRUCCION;INSTRUC_ENTRE;NUME_CLI =" & XVALO(Text1)
'    VENTANA.CampEditables = "1;3;5"
'    VENTANA.Condicion2(1) = " NUME_CLI = " & XVALO(Text1)
'    VENTANA.Condicion2(3) = " AND NUME_SUC = COL@1  "
'    VENTANA.Condicion2(5) = " AND NUME_SUC = COL@1 "
'    VENTANA.Condicion2(5) = " AND CODI_TRA = COL@3"
'    VENTANA.Inicializar = 1
'
'    VENTABNEW.Show 1

End Sub

Private Sub LiCuInCli_Click()
    Call AMACLIS07.Command6_Click
End Sub

Private Sub LiGeClaCli_Click()
 Call AMACLIS07.Command8_Click
End Sub

Private Sub LIST1_Click()
   TTXX$ = List1.TEXT
   Nusuc% = XVALO(Mid$(TTXX$, 195, 4))
   Text43 = TRIM$(Mid$(TTXX$, 63, 48))
   Text44 = TRIM$(Mid$(TTXX$, 111, 48))
   COTEXT2 = SAFETEXT$(VALOBADA("SUCENTRE", "Prov_Cli", "NUME_CLI=" & XVALO(Text1) & " and Nume_Suc=" & Nusuc%))
   Text35 = SAFETEXT$(VALOBADA("SUCENTRE", "Hora_Ent", "NUME_CLI=" & XVALO(Text1) & " and Nume_Suc=" & Nusuc%))
   ' SE CORRIGE LA POSICION DE LECTURA DEL TRANSPORTE. LA RUTINA CARGALISTA
   ' UTILIZADA PARA ARMAR EL LIST1, DEJA AUTOMATICAMENTE UN ESPACIO EN BLANCO
   ' ENTRE CADA DATO QUE NO FUE TOMADO AL MOMENTO DE LEER EL TRANSPORTE.
   ' ADEMAS SE CORRIGE LA LONGITUD DE LECTURA QUE SEGUN EL CARGALISTA ES DE 2
   Text5 = RASOTRANS$(XVALO(Mid$(TTXX$, 200, 4))) 'MAL LA POSICION
   Text45 = XVALO(Mid$(TTXX$, 200, 4))            'MAL LA POSICION
   Text11 = TRIM$(Mid$(TTXX$, 205, 10))
   NUMSUC% = XVALO(Mid$(TTXX$, 195, 4))
   Text13 = SAFETEXT$(NUMSUC%)
   
   MODIFIC(3) = 1
End Sub

Private Sub List1_DblClick()
   '\\\OT-07-0010-RG-12/01/07///
   '
   Call Command26_Click
  '
  '\\\OT-07-0010-RG-FIN///
'***********************************************************************************
  
'   NC% = XVALO(Text1)
'   If NC% > 0 Then
'     If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
'10     VTB% = VENTABU%("SUCENTRE")
'       If VTB% >= 0 Then
'
'         URE& = ULTREG&("!SUCENTRE")
'         For KKU& = 1 To URE&
'           XX$ = REGLEIDO$("!SUCENTRE", KKU&)
'           XX1$ = TRIM$(Mid$(XX$, 3, 62))
'           XX2$ = TRIM$(Mid$(XX$, 65, 48))
'           XX3$ = TRIM$(Mid$(XX$, 113, 48))
'           If XX1$ = "" Or XX2$ = "" Or XX3$ = "" Then
'             Call MENSERR(24, "Información Incompleta - Faltan Datos")
'             GoTo 10
'           End If
'           KKI% = KKU&
'           Call REPLA(XX$, MKI$(KKI%), 1, 2)
'           Call GRAREG("!SUCENTRE", XX$, KKU&)
'         Next KKU&
'         Call CIERRARCH("!SUCENTRE")
'
'         Call FILTERUPDATE("SUCENTRE", "", 5, MKI$(NC%))
'
'         List1.Clear
'         If URE& > 0 Then
'           For KKU& = 1 To URE&
'
'             TTXX$ = Mid$(REGLEIDO$("!SUCENTRE", KKU&), 3, 160)
'             List1.AddItem TTXX$
'           Next KKU&
'           List1.ListIndex = 0
'           Call List1_Click
'         End If
'
'       End If
'     End If
'   End If
End Sub

Private Sub List2_DblClick()
    '
    ' CARGO EL LIST DEL FORMULARIO DEUDACLI07 CON LA MISMA INFORMACION DE ESTE LIST
      For U& = 1 To List2.ListCount
        DEUDACLI07.LICHESEL.AddItem List2.List(U& - 1)
        VALOR2# = VALOR2# + XVALO(Mid$(List2.List(U& - 1), 63, 14))
      Next U&
      '
      DEUDACLI07.Label15 = FORMATNUM$(VALOR2#, "F14.2")
      DEUDACLI07.Show 1
      '
   
End Sub

 Private Sub mfgComposicionDeuda_DblClick()

      ' CARGO EL LIST DEL FORMULARIO DEUDACLI07 CON LA MISMA INFORMACION DE ESTE LIST
      DEUDACLI07.LICHESEL.Clear
      For FilaCompDeuda& = 1 To mfgComposicionDeuda.Rows - 1
        TIPOS$ = Left(TRIM$(mfgComposicionDeuda.TextMatrix(FilaCompDeuda&, 1)) & Space(16), 16)
            Select Case Left(mfgComposicionDeuda.TextMatrix(FilaCompDeuda&, 1), 2)
                Case "FC", "ND", "NC", "VC"
                    NumeroCbte& = XVALO(Mid$(mfgComposicionDeuda.TextMatrix(FilaCompDeuda&, 2), 9, 8))
                Case Else
                    NumeroCbte& = XVALO(Mid$(mfgComposicionDeuda.TextMatrix(FilaCompDeuda&, 2), 11, 8))
            End Select
        COCOLAR$ = Right(Space(8) & CStr(NumeroCbte&), 8)
        FECHAEMI$ = Left(TRIM$(mfgComposicionDeuda.TextMatrix(FilaCompDeuda&, 3)) & Space(8), 8)
        IMPO1$ = Right(Space(14) & TRIM$(mfgComposicionDeuda.TextMatrix(FilaCompDeuda&, 4)), 14)
        FECHAVEN$ = Left(TRIM$(mfgComposicionDeuda.TextMatrix(FilaCompDeuda&, 5)) & Space(8), 8)
        SALDOPES$ = Right(Space(14) & TRIM$(mfgComposicionDeuda.TextMatrix(FilaCompDeuda&, 6)), 14)
        If mfgComposicionDeuda.Cols = 8 Then
            SALDODOL$ = Right(Space(13) & TRIM$(mfgComposicionDeuda.TextMatrix(FilaCompDeuda&, 7)), 13)
        End If
        A$ = TIPOS$ & COCOLAR$ & "  " & FECHAEMI$ & "  " & IMPO1$ & "  " & FECHAVEN$ & " " & SALDOPES$
        
        DEUDACLI07.LICHESEL.AddItem A$
        VALOR2# = VALOR2# + XVALO(SALDOPES$)
      Next FilaCompDeuda&
      '
      DEUDACLI07.Label15 = FORMATNUM$(VALOR2#, "F14.2")
      DEUDACLI07.Show 1

End Sub


Private Sub mfgCuentaCorriente_DblClick()
    '
    FilaSeleccionada& = mfgCuentaCorriente.Row
    FE% = FECHANUM(mfgCuentaCorriente.TextMatrix(FilaSeleccionada&, 1))
    TIDOCUM$ = TRIM$(mfgCuentaCorriente.TextMatrix(FilaSeleccionada&, 2))
    If TIDOCUM$ = "Recibo Cobranza" Then TIDOCUM$ = "Recibo"
    If TIDOCUM$ = "Ajuste de Saldo" Then TIDOCUM$ = "Ajuste Saldo Cliente"
    If TIDOCUM$ = "Comp.Ctas." Then TIDOCUM$ = "Compensación de Cuentas"
    NUDOCU$ = " " + CStr(XVALO(Mid$(mfgCuentaCorriente.TextMatrix(FilaSeleccionada&, 3), 11, 8))) + " "
    NUDOCU1$ = Str$(XVALO(NUDOCU$))
    If InStr(UCase$(TIDOCUM$), "FACTURA") > 0 Or InStr(UCase$(TIDOCUM$), "NOTA DE") > 0 Or InStr(UCase$(TIDOCUM$), "N. DE") > 0 Then
        NUDOCU$ = TRIM$(mfgCuentaCorriente.TextMatrix(FilaSeleccionada&, 3))
        NUDOCU1$ = TRIM$(Mid$(mfgCuentaCorriente.TextMatrix(FilaSeleccionada&, 3), 9, 8))
        NUDOCU2$ = Mid$(mfgCuentaCorriente.TextMatrix(FilaSeleccionada&, 3), 4, 13)
        TIPOCOM% = 1
    ElseIf InStr(UCase$(TIDOCUM$), "AJUSTE SALDO") > 0 Then
      If TRIM$(mfgCuentaCorriente.TextMatrix(FilaSeleccionada&, 3)) <> "" Then
         NUDOCU$ = TRIM$(mfgCuentaCorriente.TextMatrix(FilaSeleccionada&, 3))
         NUDOCU1$ = ""
         NUDOCU2$ = NUDOCU$
      End If
      TIPOCOM% = 2
    ElseIf InStr(UCase$(TIDOCUM$), "RECIBO") > 0 Then
      NUDOCU1$ = NUDOCU$
      NUDOCU2$ = NUDOCU$
      TIPOCOM% = 3
      'EN EL CASO DEL RECIBO SE BUSCA LA FECHA REAL DE GRABACION POR QUE EN LOS CASOS QUE SE CAMBIO
      'NO MACHEA CON EL PDOC_TBL P Q ALLI GUARDA LA FECHA REAL
      CONDI$ = "NUMECOMP= " & XVALO(NUDOCU$) & " AND fechemisi = '" & FETEXCOR1$(FE%) & "' and TIPOMOVIM = 'RCOB'"
      FE% = CVINT(VALOBADA("CAJABANC", "FeHoReal", CONDI$, "NOMESS"))
    End If
    
    Screen.MousePointer = 11
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "FERVI") > 0 Then
        If OPX% > 0 And TIPOCOM% = 1 Then
           If FE% < FECHANUM("14/01/13") Then
              NUDOQUI& = FACTUFERVI&(TIDOCUM$, TRIM$(NUDOCU1$), FE%)
              If NUDOQUI& > 0 Then Call MUESTRADOC(NUDOQUI&, "", TIDOCUM$, NUDOCU1$, NUDOCU2$): Exit Sub
           End If
        End If
    End If
    
    If TIPOCOM% = 1 Then GoTo 50
    
    VARCONTROL$ = FECHATEX$(FE%)
    If TIDOCUM$ = "Recibo" Then VARCONTROL$ = TRIM$(Text2)
    DOCUNU& = DOCID&(TIDOCUM$, NUDOCU1$, NUDOCU2$, VARCONTROL$)
50  If DOCUNU& < 1 Then
        If TIPOCOM% = 1 Then ' SI OPX% ES MAYOR A 0 POR QUE ES FACTURA O NOTA DE ..
            CONSIDERAR_MOVIVIE% = 1
            Call MUESTRASTSCREEN(XVALO(Text1.TEXT), NUDOCU$)
            CONSIDERAR_MOVIVIE% = 0
        Else
            Call COMUNI("Documento No Encontrado")
        End If
    Else
        OPPXX$ = ""
        If NUDOCU$ = "00000032" Then OPPXX$ = "RECUP"
        Call MUESTRADOC(DOCUNU&, OPPXX$, TIDOCUM$, NUDOCU1$, NUDOCU2$)
    End If
    Screen.MousePointer = 1

End Sub


Private Sub mnuAsociarReporte_Click()
    Call CONECTA_CONSUL_PRE("AMACLIS07")
End Sub

Private Sub mnuEliminarCtesSinMovimientos_Click()
    'SITESIS
    'SELECCIONA TODOS LOS CLIENTES QUE EXISTEN Y NO TUVERON PEDIDOS Y LOS CARGA EN UN VECTOR CON VALOR = 1
    'SELECCIONA TODOS LOS CLIENTES QUE EXISTEN Y SI  TUVIERON TRANSACCIONES EN EL CAJABANC Y LES RESTA 1 AL MISMO VECTOR DE ESTE
    'MODO QUEDARIAN ANULADOS ESTOS ULTIMOS POR QUE QUEDARIA SU VALOR EN 0
    'RECORRE CON UN FOR EL VECTOR Y LOS CLIENTES QUE TENGAN VALOR > 0 LOS SUSPENDE.
    If EmpresaNoRealizaCambios% > 0 Then Exit Sub

    If DERACCE%("SUCLPEFA", "Suspender Clientes Sin Pedidos y Sin Facturas") < 2 Then Exit Sub
    
    Screen.MousePointer = 11
    Dim NCLIX()
    ReDim Preserve NCLIX(1)
    NC_ANT = 0
    Text9.Visible = True
    'ESTA CONSULTA DEVUELVE LOS CLIENTES QUE NO EXISTEN EN LA TABLA PEDENCA
    SQLX$ = "Select Nume_Cli  from DEUDOR12 WITH(NOLOCK)"
    SQLX$ = SQLX$ + " where Lcre_Cli >= 0 AND  not exists (select 1 from PEDENCA where PEDENCA.NroClie  = DEUDOR12.Nume_Cli) ORDER BY NUME_CLI"
    Set rs = READSET(SQLX$)
    With rs
      Do While Not .EOF
        'GUARDO EN UN VECTOR LOS CLIENTES QUE NO ESTAN EN PEDENCA
          If Text9 = "DETENER" Then GoTo 30
          Text9 = "A" & XVALO(!nume_cli) ' PROGRESO
          Text9.Refresh
          If NC_ANT <> XVALO(!nume_cli) Then
            If UBound(NCLIX) < XVALO(!nume_cli) Then ReDim Preserve NCLIX(XVALO(!nume_cli))
             NCLIX(XVALO(!nume_cli)) = 1
            NC_ANT = XVALO(!nume_cli)
         End If
         .MoveNext
      Loop
    End With
30  rs.Close
    Set rs = Nothing
    '
    If Text9 = "DETENER" Then GoTo 99
    NC_ANT = 0
    'VALIDAR SI EXISTEN EN CAJABANC. SI EXISTE LE RESTO 1 AL VECTOR PARA QUE LUEGO ABAJO NO LO CONSIDERE
    'EJEMPLO SI NO LO ENCONTRO EN PEDENCA NCLIX(NC) = 1 , SI LUEGO LO ENCUENTRA EN EL CAJABANC   NCLIX(NC) =  NCLIX(NC) - 1  QUEDANDO EN 0
    SQLX$ = "Select Nume_Cli  from DEUDOR12 WITH(NOLOCK)"
    SQLX$ = SQLX$ + " where  exists (select 1 from CAJABANC where CAJABANC.NuClien  = DEUDOR12.Nume_Cli) ORDER BY NUME_CLI"
    Set rs = READSET(SQLX$)
    With rs
      Do While Not .EOF
        'GUARDO EN UN VECTOR LOS CLIENTES QUE NO ESTAN EN PEDENCA
        If Text9 = "DETENER" Then GoTo 60
        If NC_ANT <> XVALO(!nume_cli) Then
            Text9 = "B" & XVALO(!nume_cli) ' PROGRESO
            Text9.Refresh

            If UBound(NCLIX) < XVALO(!nume_cli) Then ReDim Preserve NCLIX(XVALO(!nume_cli))
            NCLIX(XVALO(!nume_cli)) = NCLIX(XVALO(!nume_cli)) - 1
            NC_ANT = XVALO(!nume_cli)
        End If
        .MoveNext
      Loop
    End With
60  rs.Close
    Set rs = Nothing
    '
    If Text9 = "DETENER" Then GoTo 99
    CLIENTESSUSPENDIDOS& = 0
    For i& = 1 To UBound(NCLIX)
       If Text9 = "DETENER" Then GoTo 99
       Text9 = "C" & i& ' PROGRESO
       Text9.Refresh

       If NCLIX(i&) > 0 Then
          CLIENTESSUSPENDIDOS& = CLIENTESSUSPENDIDOS& + 1
          CONDI$ = "Nume_Cli = " & i&
          Call ACTUREGISTRO("DEUDOR12", "Lcre_Cli", CONDI$, -1, REGAF&)
       End If
    Next i&
99  Screen.MousePointer = 1
    Call COMUNI("Proceso Terminado\Se han Suspendido: " & CLIENTESSUSPENDIDOS& & " Clientes")
    Text9.Visible = False
'____________________________________________________________________________________
''MODO PARA CONTROLAR LA PRIMER CONSULTA
'CREE UNA VISTA PARA PODER VERIIFICAR SI NO HABIA REPETICIONES
'create view clentesnoped as  Select Nume_Cli as canti from DEUDOR12
'where not exists (select 10 from PEDENCA where PEDENCA.NroClie  = DEUDOR12.Nume_Cli)
'CONSULTA PARA VALIDAR QUE NO HAYA REPETICIONES
'select  count(*), canti  from clentesnoped  group by canti having count(*) > 1
'___________________________________________________________________________________
'HICE LO MISMO PARA LA SEGUNDA
'create view clentesnocac as
'Select Nume_Cli  from DEUDOR12 WITH(NOLOCK)
'where  exists (select 1 from CAJABANC where CAJABANC.NuClien  = DEUDOR12.Nume_Cli)

'CONSULTA PARA VALIDAR QUE NO HAYA REPETICIONES
'select  count(*), Nume_Cli  from clentesnocac  group by Nume_Cli having count(*) > 1
'__________________________________________________________________________________

End Sub

Private Sub mnuInicializaJursdiccion_Click()
    
    Call COMUNI("IMPORTANTE !!!\ \Solo se Inicializarán las Cuentas de Clientes\que NO Posean Jurisdicción.")
    
    Dim Rst As ADODB.Recordset
    
    Set Rst = New ADODB.Recordset
    Call ABRECONEXION
    ConsultaSql$ = "SELECT * FROM DEUDOR12"
    Rst.Open FILTSQL$(ConsultaSql$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With Rst
      Do While Not .EOF
        If SAFETEXT$(!JurisdiccioIIBB_CLI) = "" Then
            !JurisdiccioIIBB_CLI = !Prov_Cli
            .Update
        End If
        .MoveNext
      Loop
    End With
    Rst.Close
    Set Rst = Nothing
    
    Call COMUNI("Proceso de Inicialización Finalizado Correctamente")
    
End Sub

Private Sub mnuJurisdiccionesIIBB_Click()
    
End Sub

Private Sub mnuListadoCteParaIB_Click()
    Call OPMASCL07.Command4_Click
End Sub

Private Sub mnuPercepcionIB_Click()
    Call OPMASCL07.Command3_Click
End Sub

Private Sub mnuReportesVarios_Click()
    Call AMACLIS07.AGenRep_Click
End Sub

Private Sub mnuTokenAltaClientes_Click()

     Token$ = Control("ALTAAFIP", "CRE_TKN")
     Sign$ = Control("ALTAAFIP", "CRE_SGN")
     Expira$ = Control("ALTAAFIP", "CRE_EXP")
     
     MsgBox "Token: " & Token$ & vbCrLf & "Sign: " & Sign$ & vbCrLf & "Venc: " & Expira$
                    
End Sub

Private Sub MSFINSTRUC_DblClick()

    NC = XVALO(Text1)
    NTRANSPORTE% = XVALO(Text45) 'NRO TRANSPORTE
    NSUCURSAL% = XVALO(Text13) 'NRO SUCURSAL
    If NC < 1 Then Exit Sub
    
    If NSUCURSAL% < 1 Then
       Call COMUNI("Falta Seleccionar Sucursal")
       Exit Sub
    End If
    
    If NTRANSPORTE% < 1 Then
       Call COMUNI("Falta Seleccionar Transporte")
       Exit Sub
    End If
   
    REG& = MSFINSTRUC.MouseRow
    COL& = MSFINSTRUC.MouseCol
    If REG& < 1 Then Exit Sub
    '
    If MSFINSTRUC.Row >= 1 Then
     With EDITINSTRUC
          .TXTDESCRIPCION.TEXT = MSFINSTRUC.TextMatrix(MSFINSTRUC.Row, 2)
          .Show 1
      End With
    End If

    If CambioInstruccionEntrega% = 1 Then
      Call GRABAR_GRILLA_INSTRUCCIONES
      CambioInstruccionEntrega% = 0
    End If

'    Unload VENTABNEW
'
'
'    Campos$ = "id/A0;Instrucción/A90;NC/A0;NSUCURSAL/A0;CODITRANSP/A0"
'    VENTANA.Inicializar = 0
'    VENTANA.AgregaRegistro = 1
'    VENTANA.UtilizaFormula = 1
'    VENTANA.HabilitaBorrar = 1
'    VENTANA.HabilitaInsertar = 0
'    VENTANA.NoMostrarBotonTilde = 0
'    VENTANA.TITULO = "Carga de Instrucciones para Despachos"
'    VENTANA.CampEditables = "2"
'    VENTANA.Campox = Campos$
'    VENTANA.Inicializar = 1
'
'
'    Call CopiarMsfAOtroMsf(MSFINSTRUC, VENTABNEW.MSF, 1)
'
'    Call T_Espera(1 / 10) ' PARA QUE EL DOBLE CLICK NO APLIQUE SOBRE LA VENTANA QUE SE CARGA
'30  VENTABNEW.Show 1
'
'    '****************************************************************
    
    '
'    Set VENTANA = Nothing
'    Unload VENTABNEW
    

End Sub
Sub ACOMODARMSF()
     MSFINSTRUC.RowHeight(MSFINSTRUC.Row) = Label22.Height + 100
     MSFINSTRUC.TextMatrix(MSFINSTRUC.Row, 2) = Label22.Caption
End Sub
Private Sub MSFlexGrid1_Click()

   With MSFlexGrid1
        If .Row > 0 Then
            '
            If .TextMatrix(.Row, 0) = "" Then Exit Sub
            '
            Fila = .Row
            For i = 0 To MSFlexGrid1.Rows - 1
                .Row = i: .COL = 1
                If .CellBackColor = &HD8D8C0 Then
                    For J = 1 To MSFlexGrid1.Cols - 1
                    .COL = J
                    .CellBackColor = &HFFFFFF
                    Next J
                    Exit For
                 End If
            Next i
            .Row = Fila
            For i = 1 To MSFlexGrid1.Cols - 1
                .COL = i
                .CellBackColor = &HD8D8C0
            Next i
            '
            MSFlexGrid2.Rows = 1
            '
            NPESEL& = CLng(.TextMatrix(.Row, 0))
            NCLI0 = XVALO(Text1)
            POSIVA% = PIVACLI%(NCLI0)
            '
            SQLX$ = " SELECT * FROM PEDDETA "
            SQLX$ = SQLX$ + " WHERE NroPed = " & NPESEL&
            SQLX$ = SQLX$ + " ORDER BY NuOrItem;" 'ESTA LINEA LA AGREGUE
            '
            Dim PEDTEMP As ADODB.Recordset
            Set PEDTEMP = New ADODB.Recordset
            PEDTEMP.CursorLocation = adUseClient
26          On Error Resume Next
            PEDTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
            If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
               On Error GoTo 0
               Call CapturaErrorOpen
               GoTo 26
            End If
            On Error GoTo 0
            '
            Label82.Caption = NPESEL& ' guardo el nro de pedido
            With PEDTEMP
              On Error Resume Next
              .MoveFirst
              '
              jj& = 1: TOTALORIG# = 0: TOTALPESOS# = 0
              Do While (.EOF = False)
                CODIMAT% = !CodiInt
                If CODIMAT% > 0 And CODIMAT% <= NUMAS% Then
                    jj& = jj& + 1
                    MSFlexGrid2.Rows = jj&
                    '
                    CANP# = XVALO(!CanPed)
                    CANTEN# = XVALO(!CantEnt)
                    SALDO# = CANP - CANTEN
                    FENTRESOL% = CVINT(!FeSolEnt)
                    PRECI1# = XVALO(FORMATNUM(!PreUnid, "F10.2"))
                    LOTEX1$ = SAFETEXT$(!IDENLOTE)
                    St% = XVALO(!Status)
                    MONEMI% = XVALO(!MoneEmis)
                    TICA# = TICAMONE(MONEMI%)
                    GIVA% = GRUPIVA%(CODIMAT%)
                    POIVA# = PORCEIVA(GIVA%)
                    '
                    If (St% = 8 Or St% = 9) Then
                        Call COLOREAR_GRILLA_SOLAFILA(MSFlexGrid2, &H8080FF, jj& - 1)
                    End If

                    OBSERVACIONITEM$ = ""
                    MSFlexGrid2.TextMatrix(jj& - 1, 0) = CODEXT$(CODIMAT%)
                      OBSERVACIONITEM$ = SAFETEXT$(!ObservaItem)
                      If TRIM$(OBSERVACIONITEM$) <> "" Then
                       Call COLOREAR_GRILLA_SOLACELDA(MSFlexGrid2, &H8000000F, jj& - 1, 8)
                       Label75 = OBSERVACIONITEM$
                       MSFlexGrid2.RowHeight(jj& - 1) = Label75.Height + 100
                      End If
                    MSFlexGrid2.TextMatrix(jj& - 1, 1) = DESCRIT0$(CODIMAT%)
                    MSFlexGrid2.TextMatrix(jj& - 1, 2) = FECHATEX$(FENTRESOL%)
                    MSFlexGrid2.TextMatrix(jj& - 1, 3) = FORMATNUM$(CANP#, "F10.2")
                    MSFlexGrid2.TextMatrix(jj& - 1, 4) = FORMATNUM$(CANTEN#, "F10.2")
                    MSFlexGrid2.TextMatrix(jj& - 1, 8) = OBSERVACIONITEM$
                        If (St% = 8 Or St% = 9) Then SALDO# = 0
                        If SALDO# < 0 Then SALDO# = 0
                    MSFlexGrid2.TextMatrix(jj& - 1, 5) = FORMATNUM$(SALDO#, "F10.2")
                        IMPOR# = SALDO# * PRECI1#
                        If POSIVA% <> 5 And POSIVA% <> 7 Then IMPOR# = IMPOR# * (1 + POIVA# / 100)
                    If IMPOR# > 0.005 Then MSFlexGrid2.TextMatrix(jj& - 1, 6) = SIMBOLPES$(MONEMI%) & " " & TRIM$(FORMATNUM$(IMPOR#, "F10.2"))
                        If TICA# < 0.005 Then TICA# = 1
                        IMPORPESOS# = SALDO# * PRECI1# * TICA#
                        If POSIVA% <> 5 And POSIVA% <> 7 Then IMPORPESOS# = IMPORPESOS# * (1 + POIVA# / 100)
                    If IMPORPESOS# > 0.005 Then MSFlexGrid2.TextMatrix(jj& - 1, 7) = "$ " & TRIM$(FORMATNUM$(IMPORPESOS#, "F10.2"))
                    
                    TOTALORIG# = TOTALORIG# + IMPOR#
                    TOTALPESOS# = TOTALPESOS# + IMPORPESOS#
                    '
                End If
              .MoveNext
              Loop
            End With
            PEDTEMP.Close
            Set PEDTEMP = Nothing
            Label11.Caption = SIMBOLPES$(MONEMI%) & " " & TRIM$(FORMATNUM$(TOTALORIG#, "F10.2"))
            Label51.Caption = "$ " & TRIM$(FORMATNUM$(TOTALPESOS#, "F10.2"))
         End If
    End With

End Sub

Private Sub MSFlexGrid1_DblClick()
      '
   If DERACCE("MODOBSPE", "Modificar Observación de Item de Pedido") < 2 Then Exit Sub
   
   With MSFlexGrid1
        If .Row > 0 Then
            '
            If .TextMatrix(.Row, 0) = "" Then Exit Sub
            '
            NPX& = .TextMatrix(.Row, 0)
            '
            NUDOCU$ = TRIM$(Str$(NPX&))
            '
10          OPX% = ALTERNA%("Visualizar Pedidos - Orden de Venta\Visualizar Orden de Despacho")
            If OPX% < 1 Or OPX% > 2 Then GoTo 99
            '
            If OPX% = 1 Then
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
            
            Screen.MousePointer = 11
            'BUSCA EL DOCUMENTO EN LA TABLA SQL CON LA NUEVA FUNCION
            DOCUNU& = DOCID&(TIDOCUM$, Str$(NPX&), NUDOCU$)
            If DOCUNU& < 1 Then
                 Screen.MousePointer = 1
                 Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a la " & DOCUMENTO_ELEGIDO$ & " .\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
                 GoTo 10
            End If
            Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, Str$(NPX&), NUDOCU$)
            Screen.MousePointer = 1
            GoTo 10
        End If
    End With

99: Exit Sub
    
End Sub


Private Sub MSFlexGrid2_Click()
    Static MODO  As Boolean
    If (MSFlexGrid2.MouseRow = 0) Then
          
        ' Ordena en forma ascendente
        If MODO Then
            MSFlexGrid2.COL = MSFlexGrid2.MouseCol
            MSFlexGrid2.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            MSFlexGrid2.COL = MSFlexGrid2.MouseCol
            MSFlexGrid2.Sort = 1
            MODO = True
        End If
    End If
End Sub

Private Sub MSFlexGrid2_DblClick()
   i& = MSFlexGrid2.MouseRow
   J& = MSFlexGrid2.MouseCol
   If i& = 0 Or i& > MSFlexGrid2.Rows Then Exit Sub
   '
   MSFlexGrid2.Row = i&
   MSFlexGrid2.COL = J&
   '
   OBSERVACION$ = MSFlexGrid2.TextMatrix(i&, 8)
   Codigo$ = MSFlexGrid2.TextMatrix(i&, 0)
   CI1% = CODINT%(Codigo$)
   DESCRIX$ = DESCRIT0$(CI1%)
   NPED& = XVALO(Label82.Caption)
   '
   OBSERVAC.Label24.Caption = Codigo$ & " - - " & DESCRIX$
   OBSERVAC.Label2 = 256
   OBSERVAC.Text28.MaxLength = 256
   OBSERVAC.Text28 = ""
   OBSERVAC.Caption = "Observación de Item"
   OBS$ = MSFlexGrid2.TextMatrix(i&, 8)
   'QUITA LOS ENTER AL FINAL O AL PRINCIPIO
10 If InStr(OBS$, Chr(13)) = 1 Or InStr(OBS$, Chr(10)) = 1 Then
      OBS$ = Mid$(OBS$, 2)
      GoTo 10
   End If
20 If InStrRev(OBS$, Chr(13)) = 1 Or InStrRev(OBS$, Chr(10)) = 1 Then
      OBS$ = Mid$(OBS$, 2)
      GoTo 20
   End If
   
   OBSERVAC.Text28 = OBS$
   Call PINTATEXT(OBSERVAC.Text28)
   '
   OBSERVAC.Show 1
   '
   OBS$ = OBSERVAC.Text28
50   If InStr(OBS$, Chr(13)) = 1 Or InStr(OBS$, Chr(10)) = 1 Then
      OBS$ = Mid$(OBS$, 2)
      GoTo 10
   End If
60 If InStrRev(OBS$, Chr(13)) = 1 Or InStrRev(OBS$, Chr(10)) = 1 Then
      OBS$ = Mid$(OBS$, 2)
      GoTo 20
   End If

   If OBSERVAC.Label1 = "OK" Then
     OBSERVACIONITEM$ = OBS$
     Label75 = LTrim$(RTrim$(OBSERVACIONITEM$))
     MSFlexGrid2.RowHeight(i&) = Label75.Height + 100
     MSFlexGrid2.TextMatrix(i&, 8) = OBSERVACIONITEM$
     
     CONDI$ = " Nroped = " & NPED&
     CONDI$ = CONDI$ + " AND CODIINT = " & CI1%

     Call ACTUREGISTRO("PEDDETA", "ObservaItem", CONDI$, OBS$, REGAF&, "NOMESS")
     If TRIM$(OBSERVACIONITEM$) <> "" Then
       Call COLOREAR_GRILLA_SOLACELDA(MSFlexGrid2, &H8000000F, jj& - 1, 8)
     Else
       Call COLOREAR_GRILLA_SOLACELDA(MSFlexGrid2, vbWhite, jj& - 1, 8)
     End If
   End If
   '
End Sub

Private Sub ObtenerCuitDeDNI_Click()

    Dni$ = TRIM$(InputBox$("Ingresar el Numero de DNI", "DNI"))
    If Len(Dni$) < 9 Then
        Cuit$ = ObtieneCuitDesdeDni(Dni$)
        Call COMUNI("El CUIT Obtenido es: " & Cuit$)
    End If
    
End Sub


Private Sub SuspOper_Click()
    Call Command3_Click
End Sub

Private Sub TabCotizacion_Click()
   If DERACCE%("TICAMBIO", "Ajustar Tipo de Cambio Central") = 2 Then
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
   End If
End Sub

Private Sub Text10_Change(Index As Integer)
   MODIFIC%(4) = 1
End Sub

Private Sub Text11_Change()
   MODIFIC%(3) = 1
End Sub

Private Sub Text12_Change()
    Label85.Caption = ECOARCH$("PROVINC", TRIM$(Text12))
End Sub

Private Sub Text12_DblClick()
    Call Command7_Click
End Sub

Private Sub Text13_Change()
   Call CargarInstrucciones

End Sub

Private Sub Text15_Change()
    MODIFIC%(3) = 1
    If XVALO(Text45) > 0 Then
      CONDI$ = " Codi_Tra = " & XVALO(Text45)
      Call ACTUREGISTRO("TRANSPORTES", "CodPostal_Tra", CONDI$, TRIM$(Left$(Text15, 24)), REGAF&, "NOMESS")
    End If

End Sub

Private Sub Text2_Change()
   Text4 = ""
   If XVALO(TRIM$(Text1)) > 0 Then
     NCLIE = XVALO(TRIM$(Text2))
     If NCLIE > 0 Then
       Text4 = rasocli(NCLIE)
     End If
   End If
End Sub

Private Sub Text2_DblClick()
   Call Command11_Click
End Sub

Private Sub Text35_Change()
   MODIFIC%(3) = 1
End Sub

Private Sub Text37_Change()
   MODIFIC%(3) = 1
End Sub

Private Sub Text43_Change()
   MODIFIC%(3) = 1
End Sub

Private Sub Text43_GotFocus()
   If List1.ListCount > 0 Then
     Call MENSERR(24, "Campo no Editable")
     Text45.SetFocus
   End If
End Sub

Private Sub Text44_Change()
   MODIFIC%(3) = 1
End Sub

Private Sub Text44_GotFocus()
   If List1.ListCount > 0 Then
     Call MENSERR(24, "Campo no Editable")
     Text45.SetFocus
   End If
End Sub

Private Sub Text45_Change()
  '
  
  Label60 = ""
  Label61 = ""
  Label62 = ""
  Label63 = ""
  Label64 = ""
  Label65 = ""
  Label66 = ""
  Label67 = ""
  Label69 = ""
  Text5 = ""
  '
  NTRA = XVALO(Text45)
  If NTRA > 0 Then
    If NTRA < 32767 Then
      NTRAI% = NTRA
      CONDI$ = "CODI_TRA = " & NTRAI%
       If CantRegistros("TRANSPORTES", CONDI$) > 0 Then
        Label60 = CUITTRANS$(NTRAI%)
        Label61 = DENOTRANS$(NTRAI%)
        Label62 = RASOTRANS$(NTRAI%)
        Label63 = DOMITRANS$(NTRAI%)
        Label64 = LOCATRANS$(NTRAI%)
        Label65 = TELETRANS$(NTRAI%)
        Label66 = FAXITRANS$(NTRAI%)
        Label67 = MAILTRANS$(NTRAI%)
        Label69 = CTACTRANS$(NTRAI%)
        COTEXT3 = PROVTRANS$(NTRAI%)
        '
        Text5 = Label61
        '
      End If
    End If
  End If
  '
  Call CargarInstrucciones

  MODIFIC%(3) = 1
  '
End Sub

Private Sub Text46_Change()
   MODIFIC%(3) = 1
End Sub

Private Sub Text45_GotFocus()
    If List1.ListCount > 0 Then
        On Error Resume Next
        Command26.SetFocus
        On Error GoTo 0
    End If
End Sub

Private Sub Text5_Change()
   MODIFIC%(3) = 1
End Sub

Private Sub Text5_GotFocus()
   If XVALO(Text45) > 0 Then
     Call MENSERR(24, "Imposible Editar")
     On Error Resume Next
     Text45.SetFocus
     On Error GoTo 0
   End If
End Sub

Private Sub Text6_Change()
   MODIFIC%(0) = 1
   If FOC6% = 1 Then
     If NUMECLI&(Text6) > 0 Then
       Text1 = TRIM$(Str$(NUMECLI&(Text6)))
     End If
   End If
End Sub

Private Sub Text6_DblClick()
   Call Command10_Click
End Sub

Private Sub COTEXT_Change(Index As Integer)
   '
   'VALIDACION DE DERECHOS DE USUARIO PARA LISTA DE PRECIOS, CONDICION DE PAGO Y VENDEDOR
   If Index >= 3 And Index <= 5 Then
      NC = XVALO(Text1)
      Select Case Index ' CASOS POSIBLES DE VALIDACION
        Case 3: VALORANT% = LIPRCLI%(NC)
        Case 4: VALORANT% = CPAGCLI%(NC)
        Case 5: VALORANT% = VENDCLI%(NC)
      End Select
      
      If COTEXT(Index) <> "" Then
        If XVALO(COTEXT(Index)) <> VALORANT% Then
          If DERACCE%("DERCOME", "Derecho de Datos Comerciales") < 2 Then
             COTEXT(Index) = TRIM$(FORMATNUM(VALORANT%, "F4.0"))
          End If
        End If
      End If
   End If
   'FIN VALIDACION DE DERECHOS DE USUARIO PARA LISTA DE PRECIOS, CONDICION DE PAGO Y VENDEDOR
   '
   ECO$ = ""
   If TRIM$(COTEXT(Index).TEXT) <> "" Then
      TV11% = TIPFIELD(ARCHEX$(Index), 1)
      LF11% = LENFIELD(ARCHEX$(Index), 1)
      ARGU$ = CBIN$(TRIM$(COTEXT(Index).TEXT), TV11%, LF11%)
      ECO$ = TRIM$(ECOARCH$(ARCHEX$(Index), ARGU$))
   End If
   '
   If ECO$ <> "" Then
        DETEXT(Index).Caption = ECO$
      Else
        DETEXT(Index).Caption = "No Definido"
   End If
   '
   If Index = 1 Then
     ' CAMBIA LEYENDA SEGÚN CAT.IVA DEL CLIENTE
     If XVALO(COTEXT(1)) = 0 Then 'SI ES EXENTO
         Label34 = "Doc.No:"
       ElseIf XVALO(COTEXT(1)) = 5 Then 'SI ES COMERCIO EXTERIOR
         Label34 = "R.U.C.:"
       Else
         Label34 = "CUIT:"  'CASO CONTRARIO
     End If
   End If
   '
   MODIFIC%(0) = 1
End Sub

Private Sub COTEXT_DblClick(Index As Integer)
   Call SELECHO(ARCHEX$(Index))
   VDEVU$ = VALACT1$(ARCHEX$(Index))
   If TRIM$(VDEVU$) <> "" Then
      COTEXT(Index).TEXT = VDEVU$
   End If
End Sub

Private Sub Command10_Click()
   '
   Call GRAGENCLI
   Call GRANOTACLI
   Call GRADAENTREGA
   '
   Screen.MousePointer = 1
   RSLX$ = REGSEL$("INDICLIE")
   If Len(RSLX$) > 4 Then
     NCLIE = CVI(Mid$(RSLX$, 63, 2))
     If NCLIE > 0 Then
       NC = NCLIE
       Text1.TEXT = TRIM$(Str$(NC))
     End If
   End If
   '
End Sub

Sub CARDATCLIE()
    '
    NCLIE = XVALO(TRIM$(Text1.TEXT))
    STCL% = STATCLI%(NCLIE)
    If NCLIE < 1 Or Abs(STCL%) < 1 Then
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    Text3.TEXT = rasocli$(NCLIE)
    Text24.TEXT = TEL2CLI$(NCLIE)
    Text23.TEXT = NPROCLI$(NCLIE)
    Check1.Value = 0
    If TRIM$(Text24) <> "" Then
      If TRIM$(ECOARCH$("DESTFAB", Text24)) <> "" Then
        Check1.Value = 1
      End If
    End If
    MODIFIC(0) = 0
    MODIFIC(4) = 0
    '
    Select Case SSTab1.Tab
      Case 0
        Call CARGENCLI
      Case 1
        Call CARNOTACLI
      Case 2
        Call CARESCUENTA
      Case 3
        Call CARDAENTREGA
      Case 4
        Call CARDACOBRANZA
      Case 5
        Call CARGAPEDIDOSPENDIENTES
    End Select
    Screen.MousePointer = 1
    '
End Sub
'
Sub CARGAPEDIDOSPENDIENTES()

   Screen.MousePointer = 11
   NCLI0 = XVALO(Text1)
   '
   SQLX$ = " SELECT DISTINCT PEDENCA.NroPed, PEDENCA.NroClie, PEDENCA.FechEmis, PEDENCA.NroOcom "
   SQLX$ = SQLX$ + " FROM PEDENCA "
   SQLX$ = SQLX$ + " INNER JOIN PEDDETA "
   SQLX$ = SQLX$ + " ON (PEDENCA.NroPed = PEDDETA.NroPed) "
   SQLX$ = SQLX$ + " WHERE PEDENCA.NroClie = " & CStr(NCLI0) & " "
   SQLX$ = SQLX$ + " AND PEDDETA.Status < 8 "   ' para que no este anulada ni cumplida
   SQLX$ = SQLX$ + " AND PEDDETA.CantEnt < (PEDDETA.CanPed - 0.005) "
   SQLX$ = SQLX$ + " AND PEDENCA.Status < 9 "
   SQLX$ = SQLX$ + " AND PEDENCA.CODIEMPR = " & CodiEmp% & " "
   SQLX$ = SQLX$ + " AND PEDDETA.CODIEMPR = " & CodiEmp% & " "
   SQLX$ = SQLX$ + " ORDER BY PEDENCA.NroPed ASC "
   
   Dim PediModifTemp As ADODB.Recordset
   Set PediModifTemp = New ADODB.Recordset
   '
   PediModifTemp.CursorLocation = adUseClient
25 On Error Resume Next
   PediModifTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   U& = 1
   With PediModifTemp
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
         Do While Not .EOF
             U& = U& + 1
             MSFlexGrid1.Rows = U&
             '
             NROP& = XVALO(!NROPED)
             FEX = CVINT(!FechEmis)
             COPAGI% = XVALO(VALOBADA("PEDENCA", "CONDPAG", "NROPED=" & CStr(NROP&)))
             CPA$ = "": If COPAGI% <> 0 Then CPA$ = TRIM$(ECOARCH("CONPAG", Chr$(COPAGI%)))
             NOCO$ = SAFETEXT$(VALOBADA("PEDENCA", "NROOCOM", "NROPED=" & CStr(NROP&)))
             OBSER$ = SAFETEXT$(VALOBADA("PEDENCA", "OBSERVA", "NROPED=" & CStr(NROP&)))
             '
             MSFlexGrid1.TextMatrix(U& - 1, 0) = CStr(NROP&)
             MSFlexGrid1.TextMatrix(U& - 1, 1) = NOCO$
             MSFlexGrid1.TextMatrix(U& - 1, 2) = FECHATEX$(FEX)
             MSFlexGrid1.TextMatrix(U& - 1, 3) = CPA$
             MSFlexGrid1.TextMatrix(U& - 1, 4) = OBSER$
         .MoveNext
         
         Loop
      End If
   End With
   PediModifTemp.Close
   Set PediModifTemp = Nothing
   '
   If U& > 1 Then
           POSIVA% = PIVACLI%(NCLI0)
   
           SQLX$ = " SELECT * FROM PEDDETA "
           TTXX$ = ""
           '
           For ii = 1 To MSFlexGrid1.Rows - 1
                MSFlexGrid1.Row = ii
                NPESEL& = XVALO(MSFlexGrid1.TextMatrix(ii, 0))
                TTXX$ = TTXX$ + " NroPed = " & CStr(NPESEL&) & " OR "
           Next ii
           TTXX$ = Mid$(TTXX$, 1, Len(TTXX$) - 4)
           SQLX$ = SQLX$ & " WHERE (" & TTXX$ & ") "
           SQLX$ = SQLX$ & " AND STATUS<8 AND (CANPED-CANTENT)>0.05"
           '
           Dim PEDTEMP As ADODB.Recordset
           Set PEDTEMP = New ADODB.Recordset
           PEDTEMP.CursorLocation = adUseClient
26         On Error Resume Next
           PEDTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
           If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
              On Error GoTo 0
              Call CapturaErrorOpen
              GoTo 26
           End If
           On Error GoTo 0
           '
           TOTALPESOS# = 0
           With PEDTEMP
             On Error Resume Next
             .MoveFirst
             Do While .EOF = False
                '
                CODIMAT% = XVALO(!CodiInt)
                CANP# = XVALO(!CanPed)
                CANTEN# = XVALO(!CantEnt)
                SALDO# = CANP - CANTEN
                PRECI1# = XVALO(FORMATNUM(!PreUnid, "F10.2"))
                MONEMI% = XVALO(!MoneEmis)
                TICA# = TICAMONE(MONEMI%)
                GIVA% = GRUPIVA%(CODIMAT%)
                POIVA# = PORCEIVA(GIVA%)
                '
                IMPORPESOS# = SALDO# * PRECI1# * TICA#
                If POSIVA% <> 5 And POSIVA% <> 7 Then IMPORPESOS# = IMPORPESOS# * (1 + POIVA# / 100)
                TOTALPESOS# = TOTALPESOS# + IMPORPESOS#
                '
                .MoveNext
             Loop
           End With
           PEDTEMP.Close
           Set PEDTEMP = Nothing
           Label73.Caption = "$ " & TRIM$(FORMATNUM$(TOTALPESOS#, "F10.2"))
   End If
           
        '   TOTALPEDIDOSPENDIENTES# = 0
        '   For II = 1 To MSFlexGrid1.Rows - 1
        '        MSFlexGrid1.Row = II
        '        Call MSFlexGrid1_Click
        '        TOTALPEDIDOSPENDIENTES# = TOTALPEDIDOSPENDIENTES# + XVALO(Mid$(TRIM$(Label51.Caption), 2))
        '   Next II
        '   MSFlexGrid1.Row = 1: Call MSFlexGrid1_Click
        '   Label73.Caption = "$ " & TRIM$(FORMATNUM$(TOTALPEDIDOSPENDIENTES#, "F10.2"))
   '
   Screen.MousePointer = 1

End Sub

Sub CARGENCLI()
    '
    If YACAR%(0) > 0 Then Exit Sub
    NCLIE = XVALO(TRIM$(Text1.TEXT))
    '
    Text6 = CODICLI$(NCLIE)
    Text25.TEXT = DOMICLI$(NCLIE)
    Text25.Refresh
    Text26.TEXT = CPOSCLI$(NCLIE)
    Text26.Refresh
    Text27.TEXT = LOCACLI$(NCLIE)
    Text27.Refresh
    If LIMICRE(NCLIE) < 0 Then
        Text20.TEXT = " Suspendido"
      Else
        Text20.TEXT = CSTRING$(MKS$(LIMICRE(NCLIE)), 3, 12, 2, 2)
    End If
    Text20.Refresh
    Text8.TEXT = CSTRING$(MKS$(LCred_CCte(NCLIE)), 3, 12, 2, 2)
    Text8.Refresh
    Text23.TEXT = NPROCLI$(NCLIE)
    Text23.Refresh
    Text29.TEXT = PAISCLI$(NCLIE)
    Text29.Refresh
    Text30.TEXT = TEL1CLI$(NCLIE)
    Text30.Refresh
    Text31.TEXT = FAXCLI$(NCLIE)
    Text31.Refresh
    Text32.TEXT = EMAILCLI$(NCLIE)
    Text32.Refresh
    Text19.TEXT = CONTACTCLI$(NCLIE)
    Text19.Refresh
    Text36.TEXT = HORARCLI$(NCLIE)
    Text36.Refresh
    Text33.TEXT = CuitCli$(NCLIE)
    Text33.Refresh
    chkRegimenFacturaCredito.Value = RegimenFacturaCredito%(NCLIE)
    Text34.TEXT = NIBRCLI$(NCLIE)
    Text34.Refresh
    Text12.TEXT = JURISIIBBCLI$(NCLIE)
    Text12.Refresh
    Text21.TEXT = TRIM$(DESCUCLI$(NCLIE))
    Text21.Refresh
    Check2.Value = ENVMAIDEUDA%(NCLIE) 'Marca de envio de mail de deuda
    '
    If ULTREG&("CAPOSIV") < 8 Then
      IQ& = 0
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(0) + AJUSTI$("IVA Consumidor Final            ", 31) + AJUSTI$("Cons.Fin", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(1) + AJUSTI$("IVA Responsable Inscripto   ", 31) + AJUSTI$("Resp.Ins", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(2) + AJUSTI$("IVA Responsable no Inscripto", 31) + AJUSTI$("R.No Ins", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(3) + AJUSTI$("IVA No Responsable          ", 31) + AJUSTI$("No Resp.", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(4) + AJUSTI$("IVA Responsable Exento      ", 31) + AJUSTI$("Exento  ", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(5) + AJUSTI$("IVA Operac.Comercio Exterior", 31) + AJUSTI$("Com.Ext. ", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(6) + AJUSTI$("IVA Contrib.Monotributo     ", 31) + AJUSTI$("Monotrib", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(7) + AJUSTI$("IVA Exento Zona Promovida   ", 31) + AJUSTI$("Monotrib", 8) + String$(24, 0), IQ&)
      Call SETULTREG("CAPOSIV", IQ&)
      Call CIERRARCH("CAPOSIV")
    End If
    '
    COTEXT(0).TEXT = TRIM$(PROVCLI$(NCLIE))
    COTEXT(0).Refresh
    COTEXT(1).TEXT = TRIM$(Str$(PIVACLI%(NCLIE)))
    COTEXT(1).Refresh
    COTEXT(2).TEXT = TRIM$(CATIBRU$(NCLIE))
    COTEXT(2).Refresh
    COTEXT(3).TEXT = TRIM$(Str$(LIPRCLI%(NCLIE)))
    COTEXT(3).Refresh
    COTEXT(4).TEXT = TRIM$(Str$(CPAGCLI%(NCLIE)))
    COTEXT(4).Refresh
    COTEXT(5).TEXT = TRIM$(Str$(VENDCLI%(NCLIE)))
    COTEXT(5).Refresh
    COTEXT(6).TEXT = TRIM$(CUEMADRE$(NCLIE))
    COTEXT(6).Refresh
    COTEXT(7).TEXT = TRIM$(Str$(GRUCOCLI%(NCLIE)))
    COTEXT(7).Refresh
    '
    'grupo divisional
    GRUDI = GRUDIVCLI(NCLIE)
    Text2 = ""
    If GRUDI <> NCLIE Then Text2 = TRIM$(Str$(GRUDI))
    '
    YACAR%(0) = 1: MODIFIC%(0) = 0
    '
End Sub
'
Sub GRAGENCLI()
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))

    NCLIE = XVALO(TRIM$(Text1.TEXT))
    If NCLIE < 1 Then
      If GRABATODO% = 1 Then
        Call MENSERR(24, "Imposible Registrar !!!\Número de Cliente no Válido.")
      End If
      Exit Sub
    End If
    '
    If Abs(STATCLI%(NCLIE)) < 1 Then
      If GRABATODO% = 1 Then
        Call MENSERR(24, "Imposible Registrar !!!\Numero de Cliente Inexistente")
        Screen.MousePointer = 1
        Exit Sub
      End If
    End If
    '
    If STATCLI(NCLIE) < 1 Then
      If GRABATODO% = 1 Then
        If COMALTER%("El Cliente Elegido Figura\Marcado Como 'Dado de Baja'.\Desea Re-Activar Esta Cuenta ?\\No, Dejar de Baja\Si, Re-Activar") = 2 Then
          Call APERBASDAT("CLIEN")
          'aca
          strsql = "Update Deudor12 set stat_cli = 1 where nume_cli =" & Abs(NCLIE)
          FLEXCONN.Execute (FILTSQL$(strsql))
          
          UPDATCLI% = 1
          Call GENECOCLI
          Call CARGENCLI
        End If
      End If
      Screen.MousePointer = 1
      Exit Sub
    End If
    '
    If MODIFIC%(0) < 1 Then
      If GRABATODO% <> 1 Then
        Exit Sub
      End If
    End If
    '
    If GRABATODO% = 1 Then
      If SSTab1.Tab <> 0 Then
        Exit Sub
      End If
    End If
    '
    If DERACCE%("MODATCLI", "Modificar Datos de Clientes") < 2 Then
       Exit Sub
    End If
    '
    If GRABATODO% <> 1 Then
      Command9.SetFocus
      If COMALTER%("Se han Modificado Datos del Cliente\'" + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
        YACAR%(0) = 0
        Screen.MousePointer = 11
        Text3.TEXT = rasocli$(NCLIE)
        Text24.TEXT = TEL2CLI$(NCLIE)
        MODIFIC(0) = 0
        Call CARGENCLI
        Screen.MousePointer = 1
        Exit Sub
      End If
    End If
    '
    If EmpresaNoRealizaCambios% > 0 Then Exit Sub

    Cuit$ = TRIM$(Text33.TEXT)
    PIVA% = VALIMI(COTEXT(1).TEXT, 0, 7)
    If PIVA% <> 5 And PIVA% <> 0 And Cuit$ = "" Then
          Screen.MousePointer = 1
          Call MENSERR(24, "Falta Ingresar un Número de C.U.I.T.  Válido")
          SSTab1.Tab = 0
          Call PINTATEXT(Text33)
          Exit Sub
    End If
    
    If Cuit$ <> "" Then
'      PIVA% = VALIMI(COTEXT(1).TEXT, 0, 7)
      '\\\OT-08-1054-OD-19/11/08///
      '"""MEDITEA" ' EXCLUYE COMERCIO EXTERIOR E IVA CONSUMIDOR FINAL
      If PIVA% <> 5 And PIVA% <> 0 Then
      '\\\OT-08-1054-OD-FIN///
        If VALICUIT%(Cuit$) <> 1 Then
          Screen.MousePointer = 1
          Call MENSERR(24, "Imposible Registrar !!!\Número de C.U.I.T. no Válido")
          SSTab1.Tab = 0
          Call PINTATEXT(Text33)
          Exit Sub
        End If
      End If
    End If
    '
    If Cuit$ <> "" Then
        If VALICUIT%(Cuit$) = 1 Then Cuit$ = FormatearCuitValido$(Cuit$)
        CONDI$ = "Cuit_Cli = '" & Cuit$ & "' AND nume_cli <> " & NCLIE
        razonSocial$ = TRIM$(VALOBADA("DEUDOR12", "RASO_CLI", CONDI$, "NOMESS"))
        If razonSocial$ <> "" Then
            Screen.MousePointer = 1
            OPX% = COMALTER%("Ciente: " & razonSocial$ & " Posee el Mismo Número de C.U.I.T.\\Cancelar\Grabar Igual")
            If OPX% < 2 Then
              Call PINTATEXT(Text33)
              Exit Sub
            End If
            'SE QUITA POR QUE SI TIENE CONFIGURACION DE VALIDACION DE DERECHOS DE CUIT DESHABILITA LA MODIFICACION
'            If DERACCE%("REPICLIE", "Repetir C.U.I.T. en Diferentes Clientes") < 2 Then
'              Call PINTATEXT(Text33)
'              Exit Sub
'            End If
               
        End If
    End If
    '
    If NCLIE = 0 Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Imposible Registrar Cuenta Número '0'")
        Exit Sub
    End If
    '
    If TRIM$(Text6) = "" Then
      Call MENSERR(24, "Imposible Registrar !!!\Falta Código de Cuenta")
      Screen.MousePointer = 1
      Exit Sub
    End If
    '
    NPRII = NUMECLI&(Text6)
    If NPRII > 0 Then
      If NPRII <> Abs(NCLIE) Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Imposible Registrar !!!\Codigo de Cuenta Repetido o no Válido.")
        Exit Sub
      End If
    End If
    '
    If TRIM$(Text3.TEXT) = "" Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Imposible Registrar !!!\Falta Razón Social del Cliente.")
        Exit Sub
    End If
    '
    If XVALO(Text8) > XVALO(Text20) Then
        Screen.MousePointer = 1
        Call COMUNI("Imposible Asignar Limite de Crédito en C.Cte\Superior al Limite de Crédito General")
        Exit Sub
    End If
    '
    LIPRE% = VALIMI(COTEXT(3).TEXT, 0, 255)
    If LIPRE% > 0 Then
        If ECOARCH("LISTAS", Chr$(LIPRE%)) = "" Then
          Call MENSERR(24, "Lista de Precios no Válida")
          LIPRE% = 0
        End If
    End If
    '
    COPAGI% = VALIMI(COTEXT(4).TEXT, 0, 255)
      If COPAGI% > 0 Then
        If ECOARCH("CONPAG", Chr$(COPAGI%)) = "" Then
          Call MENSERR(24, "Condición de Pago no Válida")
          COPAGI% = 0
        End If
    End If
    '
    VENDE% = VALIMI(COTEXT(5).TEXT, 0, 255)
      If VENDE% > 0 Then
        If ECOARCH("VENDEDOR", Chr$(VENDE%)) = "" Then
          Call MENSERR(24, "Vendedor no Válido")
          VENDE% = 0
        End If
    End If
    '
    POSIVA = XVALO(COTEXT(1))
    If POSIVA = 0 Then
        If Len(TRIM$(Text33)) > 12 Then
            Call MENSERR(24, "El Límite Máximo Para Número de Documento es de 12 Dígitos")
            Exit Sub
        End If
    End If
    '
    If InStr(UCase$(EMPREAC$), "MEDITEA") > 0 Then
      If TRIM$(UCase$(DETEXT(0))) = "NO DEFINIDO" Then
            Call MENSERR(24, "Falta Código de Provincia o Código Inválido")
            Call PINTATEXT(COTEXT(0))
            Exit Sub
      End If
    End If
    '
    If Control("MASTER", "DOMIMAND") = "SI" Then
       If TRIM$(Text25) = "" Then
            Call MENSERR(24, "Falta Ingresar el Domicilio")
            Call PINTATEXT(Text25)
            Exit Sub
       End If
    End If

    Screen.MousePointer = 11
    '
    'Grabacion con Autoregistro, en Tabla AUDITACT
    CONDI$ = "nume_cli = " & NCLIE
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    SQLX$ = "Select * from DEUDOR12"
    SQLX$ = SQLX$ + " WHERE " & CONDI$
25  On Error Resume Next
   rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0
    '
    With rs
       Do While Not .EOF

       !Codi_Cli = (UCase$(Text6))
       !raso_cli = Text3
       !Stat_Cli = 1
       !domi_cli = Text25
       !Cpos_Cli = Text26
       !loca_Cli = Text27
       !Tele_Cli = Text30
       !Fant_Cli = Text24
       !Cuit_Cli = TRIM$(Cuit$)
       !RegimenFacturaCredito = chkRegimenFacturaCredito.Value
       !Nibr_Cli = TRIM$(Text34)
       !JurisdiccioIIBB_CLI = TRIM$(Text12)
       !Caib_Cli = TRIM$(COTEXT(2))
       !Prov_Cli = TRIM$(COTEXT(0))
       .Update
       .MoveNext
       Loop
     End With
     rs.Close
     Set rs = Nothing
       
'    Call ACTUREGISTRO("DEUDOR12", "Codi_Cli", CONDI$, (UCase$(Text6)), REGAF&)
'    Call ACTUREGISTRO("DEUDOR12", "Raso_Cli", CONDI$, Text3, REGAF&)
'    Call ACTUREGISTRO("DEUDOR12", "Stat_Cli", CONDI$, 1, REGAF&)
'    Call ACTUREGISTRO("DEUDOR12", "Domi_Cli", CONDI$, Text25, REGAF&)
'    Call ACTUREGISTRO("DEUDOR12", "Cpos_Cli", CONDI$, Text26, REGAF&)
'    Call ACTUREGISTRO("DEUDOR12", "Loca_Cli", CONDI$, Text27, REGAF&)
'    Call ACTUREGISTRO("DEUDOR12", "Tele_Cli", CONDI$, Text30, REGAF&)
'    Call ACTUREGISTRO("DEUDOR12", "Fant_Cli", CONDI$, Text24, REGAF&)
'    Call ACTUREGISTRO("DEUDOR12", "Cuit_Cli", CONDI$, TRIM$(CUIT$), REGAF&)
'    Call ACTUREGISTRO("DEUDOR12", "Nibr_Cli", CONDI$, TRIM$(Text34), REGAF&)
'    Call ACTUREGISTRO("DEUDOR12", "Caib_Cli", CONDI$, TRIM$(COTEXT(2)), REGAF&)
'    Call ACTUREGISTRO("DEUDOR12", "Prov_Cli", CONDI$, TRIM$(COTEXT(0)), REGAF&)
    '
    MODLIMCRE% = 0
    LIMCRE# = XVALO(VALOBADA("Deudor12", "Lcre_Cli", CONDI$))
    If Abs(LIMCRE# - XVALO(Text20.TEXT)) > 0.05 Or InStr(UCase$(Text20.TEXT), "SUSPEN") > 0 Then
        If DERACCE%("MODLICRE", "MODIFICACION LIMITE DE CREDITO") >= 1 Then
           Call ACTUREGISTRO("DEUDOR12", "Lcre_Cli", CONDI$, XVALO(Text20), REGAF&)
              If TRIM$(UCase$(Text20.TEXT)) = "SUSPENDIDO" Then
                 Call ACTUREGISTRO("DEUDOR12", "Lcre_Cli", CONDI$, -1, REGAF&)
              End If
        Else
          Limite = XVALO(VALOBADA("Deudor12", "Lcre_Cli", CONDI$))
          Text20.TEXT = FORMATNUM(Limite, "F12.2")
          If LMCR < 0 Then Text20.TEXT = " Suspendido"
          Text20.Refresh
          MODLIMCRE% = 1
        End If
    End If
    '
    LMCRCCTE = XVALO(Text8)
    If LMCRCCTE < 0 Then LMCRCCTE = 0
    '
    CONDI$ = "nume_cli = " & NCLIE
    Dim rs1 As ADODB.Recordset
    Set rs1 = New ADODB.Recordset
    SQLX$ = "Select * from DEUDOR12"
    SQLX$ = SQLX$ + " WHERE " & CONDI$
30  On Error Resume Next
    rs1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 30
    End If
    On Error GoTo 0
    '
    With rs1
       Do While Not .EOF
       !LCred_CCte = LMCRCCTE
          PIVA% = VALIMI(COTEXT(1).TEXT, 0, 7)
       !Piva_Cli = PIVA%
          CATIB% = VALIMI(COTEXT(2).TEXT, 0, 2)
       !Caib_Cli = CATIB%
       !Lpre_Cli = LIPRE%
       !Cpag_Cli = COPAGI%
       !Vend_Cli = VENDE%
           CUECO% = VALIMI(COTEXT(7).TEXT, 0, 32767)
           If ECOARCH$("GRUCOCLI", MKI$(CUECO%)) = "" Then
               CUECO% = 0
           End If
       !GrCo_Cli = CUECO%
          CCX% = REGBUS&("CUENTAS", 1, AJUSTI$(COTEXT(6).TEXT, 12))
       !CuMa_Cli = CCX% ''''
          DCC$ = TRIM$(Text21.TEXT)
       !Dsta_Cli = DCC$ '''''
       !Pais_Cli = TRIM$(Text29) '''''
       !Faxi_Cli = TRIM$(Text31) ''''''
       !Mail_Cli = TRIM$(Text32) ''''''''
       !CTAC_CLI = TRIM$(Text19) '''''''
       
       
       !Prvd_Cli = TRIM$(Text23) ''''''
       !Hora_Cli = TRIM$(Text36) ''''''''
       !ENVMAIDEUDA = Check2.Value ''''''''
       !Gru_Div = XVALO(Text2) '''''''
       .Update
       .MoveNext
       Loop
     End With
     rs1.Close
     Set rs1 = Nothing
     If InStr(EPAC$, "MINEPARTS") < 1 And InStr(EPAC$, "AHP") > 0 Then
       CONDI$ = "nume_cli = " & NCLIE
       Call DuplicarRegistroDeUnaBaseAOtra("AHP", "DEUDOR12", "MINEPARTS", "DEUDOR12", CONDI$)
     End If
     
'    Call ACTUREGISTRO("DEUDOR12", "LCred_CCte", CONDI$, LMCRCCTE, REGAF&)
    '
'    PIVA% = VALIMI(COTEXT(1).TEXT, 0, 7)
    '
'    Call ACTUREGISTRO("DEUDOR12", "Piva_Cli", CONDI$, PIVA%, REGAF&)
    '
'    CATIB% = VALIMI(COTEXT(2).TEXT, 0, 2)
    '
'    Call ACTUREGISTRO("DEUDOR12", "Caib_Cli", CONDI$, CATIB%, REGAF&)
    '
'    Call ACTUREGISTRO("DEUDOR12", "Lpre_Cli", CONDI$, LIPRE%, REGAF&)
    '
'    Call ACTUREGISTRO("DEUDOR12", "Cpag_Cli", CONDI$, COPAGI%, REGAF&)
    '
'    Call ACTUREGISTRO("DEUDOR12", "Vend_Cli", CONDI$, VENDE%, REGAF&)
    '
'    CUECO% = VALIMI(COTEXT(7).TEXT, 0, 32767)
'    If ECOARCH$("GRUCOCLI", MKI$(CUECO%)) = "" Then
'       CUECO% = 0
'    End If
    '
'    Call ACTUREGISTRO("DEUDOR12", "Grco_Cli", CONDI$, CUECO%, REGAF&)
    '
'    Call ACTUREGISTRO("DEUDOR12", "Vend_Cli", CONDI$, VENDE%, REGAF&)
    '
'    CCX% = REGBUS&("CUENTAS", 1, AJUSTI$(COTEXT(6).TEXT, 12))
    '
'    Call ACTUREGISTRO("DEUDOR12", "CuMa_Cli", CONDI$, CCX%, REGAF&)
    '
'    DCC$ = TRIM$(Text21.TEXT)
    '
'    Call ACTUREGISTRO("DEUDOR12", "Dsta_Cli", CONDI$, TRIM$(Text21), REGAF&)
'    Call ACTUREGISTRO("DEUDOR12", "Pais_Cli", CONDI$, Text29, REGAF&)
'    Call ACTUREGISTRO("DEUDOR12", "Faxi_Cli", CONDI$, TRIM$(Text31), REGAF&)
'    Call ACTUREGISTRO("DEUDOR12", "Mail_Cli", CONDI$, TRIM$(Text32), REGAF&)
'    Call ACTUREGISTRO("DEUDOR12", "Ctac_Cli", CONDI$, TRIM$(Text19), REGAF&)
'    Call ACTUREGISTRO("DEUDOR12", "Prvd_Cli", CONDI$, TRIM$(Text23), REGAF&)
'    Call ACTUREGISTRO("DEUDOR12", "Hora_Cli", CONDI$, TRIM$(Text36), REGAF&)
'    Call ACTUREGISTRO("DEUDOR12", "ENVMAIDEUDA", CONDI$, Check2.Value, REGAF&)
    'AGREGADO POR QUE NO SE ESTABA GUARDANDO
'    Call ACTUREGISTRO("DEUDOR12", "Gru_Div", CONDI$, XVALO(Text2), REGAF&)
    '
    
' *******************************************************

'    Call APERBASDAT("CLIEN")
'    Dim Rdeudor12 As ADODB.Recordset
'    Set Rdeudor12 = New ADODB.Recordset
'    Rdeudor12.CursorType = adOpenDynamic
'    Rdeudor12.LockType = adLockPessimistic
'    Rdeudor12.Open FILTSQL$("Select * from Deudor12 where nume_cli=" & NCLIE& & "and stat_cli=1"), FLEXCONN, , , adCmdText
'    If (Rdeudor12.EOF And Rdeudor12.BOF) Then 'si esta vacio
'       Call MENSERR(24, "Imposible Registrar")    ' no existe
'       Rdeudor12.Close
'       Set Rdeudor12 = Nothing
'       Exit Sub
'    End If
'    '
'10  On Error GoTo 0
'    With Rdeudor12
'      On Error Resume Next
'      ABC1$ = Left$(!Codi_Cli, 12)
'      If Err Then
'        ABC1$ = ""
'      End If
'      On Error GoTo 0
'      '.Edit
'      If TRIM$(Text3) <> TRIM$(!Raso_Cli) Or UCase$(Text6) <> ABC1$ Then UPDATCLI% = 1
'        !Codi_Cli = Left$(UCase$(Text6), 12)
'        !Raso_Cli = Left$(Text3, 64)
'        !Stat_Cli = 1
'        !Domi_Cli = Left$(Text25, 64)
'        !Cpos_Cli = Left$(Text26, 16)
'        !Loca_Cli = Left$(Text27, 48)
'        !Tele_Cli = Left$(Text30, 32)
'        !Fant_Cli = Left$(Text24, 32)
'        !Cuit_Cli = Left$(TRIM$(CUIT$), 24)
'        !Nibr_Cli = Left$(TRIM$(Text34), 24)
'        !Caib_Cli = Left$(TRIM$(COTEXT(2)), 2)
'        !Prov_Cli = Left$(TRIM$(COTEXT(0)), 2)
         '
'        MODLIMCRE% = 0
'        On Error Resume Next
'        LMCRA = CDbl(!Lcre_Cli)
'        If Err Then
'        LMCRA = 0: !Lcre_Cli = 0
'      End If
'      On Error GoTo 0
      '
'      If Abs(CDbl(!Lcre_Cli) - XVALO(Text20.TEXT)) > 0.05 Or InStr(UCase$(Text20.TEXT), "SUSPEN") > 0 Then
'       If Abs(LIMCRE# - XVALO(Text20.TEXT)) > 0.05 Or InStr(UCase$(Text20.TEXT), "SUSPEN") > 0 Then
'        If DERACCE%("MODLICRE", "MODIFICACION LIMITE DE CREDITO") >= 1 Then
'            !Lcre_Cli = XVALO(Text20)
'             If TRIM$(UCase$(Text20.TEXT)) = "SUSPENDIDO" Then
'              !Lcre_Cli = (-1)
'             End If
'        Else
'           Text20.TEXT = CSTRING$(MKS$(!Lmcr_Cli), 3, 12, 2, 2)
'            If LMCR < 0 Then Text20.TEXT = " Suspendido"
'               Text20.Refresh
'               MODLIMCRE% = 1
'            End If
'       End If
'
'      LMCRCCTE = XVALO(Text8)
'      If LMCRCCTE < 0 Then LMCRCCTE = 0
'      !LCred_CCte = LMCRCCTE
'      PIVA% = VALIMI(COTEXT(1).TEXT, 0, 7)
'      !Piva_Cli = PIVA%
'      CATIB% = VALIMI(COTEXT(2).TEXT, 0, 2)
'      !Caib_Cli = CATIB%
'      LIPRE% = VALIMI(COTEXT(3).TEXT, 0, 255)
       '
'      If LIPRE% > 0 Then
'        If ECOARCH("LISTAS", Chr$(LIPRE%)) = "" Then
'          Call MENSERR(24, "Lista de Precios no Válida")
'          LIPRE% = 0
'        End If
'      End If
'     !Lpre_Cli = LIPRE%
      '
'      COPAGI% = VALIMI(COTEXT(4).TEXT, 0, 255)
'      If COPAGI% > 0 Then
'        If ECOARCH("CONPAG", Chr$(COPAGI%)) = "" Then
'          Call MENSERR(24, "Condición de Pago no Válida")
'          COPAGI% = 0
'        End If
'      End If
'     !Cpag_Cli = COPAGI%
       '
'      VENDE% = VALIMI(COTEXT(5).TEXT, 0, 255)
'      If VENDE% > 0 Then
'        If ECOARCH("VENDEDOR", Chr$(VENDE%)) = "" Then
'          Call MENSERR(24, "Vendedor no Válido")
'          VENDE% = 0
'        End If
'      End If
'     !Vend_Cli = VENDE%
      '
'      CUECO% = VALIMI(COTEXT(7).TEXT, 0, 32767)
'      If ECOARCH$("GRUCOCLI", MKI$(CUECO%)) = "" Then
'        CUECO% = 0
'      End If
'     !GrCo_Cli = CUECO%
      '
'     !Gru_Div = XVALO(Text2)
'     '
'      CCX% = REGBUS&("CUENTAS", 1, AJUSTI$(COTEXT(6).TEXT, 12))
      '
'     !CuMa_Cli = CCX%
'     '
'      DCC$ = TRIM$(Text21.TEXT)
'     !Dsta_Cli = Left$(TRIM$(Text21), 32)
'     !Pais_Cli = Left$(Text29, 48)
'     !Faxi_Cli = Left$(Text31, 32)
'     !Mail_Cli = Left$(Text32, 64)
'     !Ctac_Cli = Left$(Text19, 48)
'     !Prvd_Cli = Left$(Text23, 24)
'     !Hora_Cli = Left$(Text36, 32)
'     !ENVMAIDEUDA = Check2.Value
'     '
'     .Update
'    End With

' ******************************************************
    
    If Check1.Value = 1 Then
        If TRIM$(Text24) <> "" Then
          If ULTREG&("DESTFAB") < 2 Then Call VERIDEST
          'NFANCLI$ = AJUSTI$(UCase$(Text24), 8)
          NFANCLI$ = AJUSTI$(CStr(Text1), 8)
          RASOCIA$ = AJUSTI$(Text3, 24)
          Call FILTERPUTRECORD("DESTFAB", 1, TRIM$(NFANCLI$), NFANCLI$ + RASOCIA$)
        End If
      Else
        NFANCLI$ = AJUSTI$(CStr(Text1), 8)
        Call FILTERDELETE("DESTFAB", 1, NFANCLI$)
    End If
    '
'    Rdeudor12.Close
    Set Rdeudor12 = Nothing
    ' ACTUALIZA GRUPO DIVISIONAL
    GRDCLI& = XVALO(Text2)
    If GRDCLI& > 0 Then
      If GRDCLI& <> NCLIE Then
        QTR% = CantRegistros("CAJABANC", "NUCLIEN = " & NCLIE)
        If QTR% > 0 Then
          If COMALTER%("¿ Desea Aplicar El Grupo Divisional\a los Movimientos Anteriores de Esta Cuenta ?\\No, Conservar Anteriores\Si, Actualizar con Nuevo G.D") = 2 Then
            If DERACCE%("ACMOGRUD", "Actualizar Movimientos por Grupo Divisional") = 2 Then
              Screen.MousePointer = 11
              SQLX$ = "UPDATE CAJABANC SET NUCLIEN = " & GRDCLI& & " WHERE NUCLIEN = " & NCLIE
              FLEXCONN.Execute (FILTSQL$(SQLX$))
              SQLX$ = "UPDATE SADEUPEN SET NUCLIEN = " & GRDCLI& & " WHERE NUCLIEN = " & NCLIE
              FLEXCONN.Execute (FILTSQL$(SQLX$))
              Call COMUNI("Realizado")
              Screen.MousePointer = 1
            End If
          End If
        End If
      End If
    End If
    '
    'UPDATCLI% = 1
    'Call GENECOCLI
    RasCli$ = TRIM$(Text3)
    REGCOMPLETO$ = FILTERGETRECORD$("DEUDOR12", 1, MKS$(NCLIE))
    RasCliDat$ = TRIM$(Mid$(REGCOMPLETO$, 5))
    If TRIM$(REGCOMPLETO$) <> "" Then
      If RasCli$ <> RasCliDat$ Then
         Call CARGALISEL("DEUDOR1", "Deudor12", "Nume_Cli/F6.0>;Raso_Cli/A48", "Stat_Cli >= 0 and Nume_Cli>0 ORDER BY Raso_Cli")
         Call CARGALISEL("INDICLIE", "Deudor12", "Codi_Cli/A12<;Raso_Cli/A50", "Stat_Cli >= 0 AND Codi_Cli <> '' ORDER BY Raso_Cli")
         Call REPLA(REGCOMPLETO$, RasCli$, 5, 28)
         Call FILTERPUTRECORD("DEUDOR12", 1, MKS$(NCLIE), REGCOMPLETO$)
         'Call FILESORT("DEUDOR12", "", 1, 1) comentado por que estaba dando error
         Call FRESHECHO("DEUDOR12")
      End If
    Else
      If ULTREG&("DEUDOR12") > 0 Then
        XX$ = REGBLAN$("DEUDOR12")
        Call REPLA(XX$, MKS$(NCLIE), 1, 4)
        Call REPLA(XX$, RasCli$, 5, 28)
        Call REGAPP("DEUDOR12", XX$)
        'Call FILESORT("DEUDOR12", "", 1, 1) comentado por que estaba dando error
      Else
        Dim Rst As ADODB.Recordset
        SQLX$ = "select nume_cli, raso_cli from deudor12 with(nolock) where stat_cli>0"
        Set Rst = READSET(SQLX$)
        '
        With Rst
          jj& = 0
          Do While Not .EOF
              NCLIEN = XVALO(!nume_cli)
              RASOC$ = SAFETEXT$(!raso_cli)
              '
              XX$ = REGBLAN$("DEUDOR12")
              Call REPLA(XX$, MKS$(NCLIEN), 1, 4)
              Call REPLA(XX$, RASOC$, 5, 28)
              jj& = jj& + 1
              Call GRAREG("DEUDOR12", XX$, jj&)
              .MoveNext
          Loop
        End With
        On Error Resume Next
        Rst.Close
        Set Rst = Nothing
        On Error GoTo 0
        Call SETULTREG("DEUDOR12", jj&)
        Call FRESHECHO("DEUDOR12")
        '
      End If
     End If
    '
    MODIFIC%(0) = 0
    Screen.MousePointer = 1
    '
End Sub
'
Sub CARNOTACLI() 'ver
   '
   If YACAR%(1) > 0 Then Exit Sub
   NCLIE = XVALO(Text1.TEXT)
   Text28 = ""
   '
   If NCLIE > 0 Then
     If REGICLI(NCLIE) > 0 Then
       Call APERBASDAT("CLIEN")
       'Set ClienTemp = Badaclie.OpenRecordset("SELECT * FROM Deudor12 WHERE Nume_Cli = " & NCLIE, dbOpenSnapshot)
       '
       Set ClienTemp = FLEXCONN.Execute(FILTSQL$("SELECT * FROM Deudor12 WHERE Nume_Cli = " & NCLIE))
       On Error Resume Next '
       ClienTemp.MoveFirst
       If Err < 1 Then
         Text28 = ClienTemp!nota_Cli
       End If
     End If
   End If
   Text28.Refresh
   YACAR%(1) = 1: MODIFIC%(1) = 0
   '
   ClienTemp.Close
   Set ClienTemp = Nothing
  
End Sub
'
Sub GRANOTACLI()
   '
   NCLIE = XVALO(TRIM$(Text1.TEXT))
   If NCLIE < 1 Then Exit Sub
   '
   If REGICLI(NCLIE) < 1 Then Exit Sub
   If YACAR%(1) < 1 Then Exit Sub
   If MODIFIC%(1) < 1 Then Exit Sub
   '
   If DERACCE%("ANOTACLI", "Anotaciones de de Clientes") < 2 Then
      Exit Sub
   End If
   '
   If GRABATODO% <> 1 Then
     Command9.SetFocus
     If COMALTER%("Se han Modificado las Notas del Cliente\'" + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
       YACAR%(1) = 0
       Call CARNOTACLI
       Exit Sub
     End If
     Command9.SetFocus
   End If
   '
   Call APERBASDAT("CLIEN")
     Screen.MousePointer = 11
     TEXTO1$ = Text28.TEXT + Chr$(13) + Chr$(10)
     '
     Dim RSTDEUDOR12 As ADODB.Recordset
     Set RSTDEUDOR12 = New ADODB.Recordset
     RSTDEUDOR12.CursorType = adOpenDynamic
     RSTDEUDOR12.LockType = adLockPessimistic
     RSTDEUDOR12.Open FILTSQL$("Select * from Deudor12 where nume_cli=" & Abs(NCLIE)), FLEXCONN, , , adCmdText
     On Error Resume Next
     RSTDEUDOR12!nota_Cli = TEXTO1$
     RSTDEUDOR12.Update
     MODIFIC%(1) = 0
     If Err Then
       On Error GoTo 0
       Call MENSERR(24, "Imposible Registrar")
     End If
     '
   
10   On Error GoTo 0
     Screen.MousePointer = 1
   
   
   'Screen.MousePointer = 11
   'TEXTO$ = Text28.TEXT + Chr$(13) + Chr$(10)
   '
   'With Deudor12
   '  On Error Resume Next
   '  .MoveFirst
   '  .FindFirst "Nume_Cli = " & Abs(NCLIE)
   '  If .NoMatch = False Then GoTo 10
   '  On Error GoTo 0
   '  Call MENSERR(24, "Imposible Registrar")
   '  Screen.MousePointer = 1
   '  Exit Sub
   'End With
   '
'10 'On Error GoTo 0
   'With Deudor12
   '  .Edit
   '  !Nota_Cli = TEXTO$
   '  .Update
   'End With
   'Deudor12.Close
   'MODIFIC%(1) = 0
   'Screen.MousePointer = 1
   '
End Sub
'
Sub CARDAENTREGA()
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If YACAR%(3) > 0 Then Exit Sub
   NCLIE = XVALO(Text1.TEXT)
   Text45 = ""
   Text43 = ""
   Text44 = ""
   Text5 = ""
   Text35 = ""
   Text37 = ""
   Text46 = ""
   DETEXT2 = ""
   COTEXT2 = ""
   List1.Clear
   '
   If NCLIE > 0 Then
      Text45 = TRIM$(Str$(NUTRANCLI%(NCLIE)))
      If TRIM$(Text45) = "0" Then Text45 = ""
      
      Text43 = TRIM$(DOMIENT$(NCLIE))
      Text44 = TRIM$(LOCAENT$(NCLIE))
      Text11 = TRIM$(CODPOSTALENT$(NCLIE))
      Text5 = TRIM$(TRANENT$(NCLIE))
      Text35 = TRIM$(HORAENT$(NCLIE))
      Text37 = TRIM$(ZONACLI$(NCLIE))
      Text46 = OBSENTRE$(NCLIE)
      Text7 = FORMATNUM$(POREMBAL(NCLIE), "F5.1")
      COTEXT2 = TRIM$(codProvinciaEntrega$(NCLIE))
      
      'grilla de detalle de instrucciones de entrega AHP
      Campos$ = "id/A0;Instrucción/A80;NC/A0;NSUCURSAL/A0;CODITRANSP/A0"
      Call CARGA_ENCABEZADO(Me.MSFINSTRUC, Campos$, ACLI_GES07)

      ' Arma la lista con las sucursales
      'Call CARGALISTA(List1, "SUCENTRE", "DENO_SUC/A62;DOMI_SUC/A48;LOCA_SUC/A48;TELE_SUC/A32;NUME_SUC/I4;CODI_TRA/I4;CODPOSTAL_SUC/A10", "Nume_Cli =" & NCLIE) 'MAL LA POSICION
      Call CARGALISTA(List1, "SUCENTRE", "DENO_SUC/A62;DOMI_SUC/A48;LOCA_SUC/A48;TELE_SUC/A32;NUME_SUC/I4;CODI_TRA/I4;CODPOSTAL_SUC/A10", "Nume_Cli =" & NCLIE & " AND STATUS >=0 ")
      If InStr(EPAC$, "AHP") > 0 Then Call OrdenarListBox(List1)
      ' LO QUE SIGUE ES PARA QUE EL LUGAR DE ENTREGA
        ' SEA EL DE LA PRIMERA SUCURSAL
      If List1.ListCount > 0 Then
          List1.ListIndex = 0
          Call LIST1_Click
          'Text13 = 1
      End If
  

'      Call FILTERFILE("SUCENTRE", "", 5, MKI$(Int(NCLIE)))
'      URE& = ULTREG&("!SUCENTRE")
'      For KKU& = 1 To URE&
'        NUMSUC% = CVI(Mid$(REGLEIDO$("!SUCENTRE", KKU&), 1, 2))
'        TTXX1$ = TRIM$(Str(NUMSUC%))
'        TTXX2$ = REGLEIDO$("!SUCENTRE", KKU&)
'        NUTRAI% = CVI(Mid$(TTXX2$, 195, 2))         'NUMERO DE TRANSPORTE
'        NUMESUCUR$ = FORMATNUM(NUMSUC%, "I4")       'NUMERO DE SUCURSAL
'        'TTXX$ = Mid$(TTXX2$, 3, 158) + Mid$(TTXX2$, 163, 32) + TTXX1$
'        TTXX$ = Mid$(TTXX2$, 3, 158) + Mid$(TTXX2$, 163, 32) + NUMESUCUR$ + FORMATNUM$(NUTRAI%, "I4")
'        List1.AddItem TTXX$
'      Next KKU&
'      If URE& > 0 Then
'        List1.ListIndex = 0
'        Call List1_Click
'      End If
      '
   End If
   '
   YACAR%(3) = 1: MODIFIC%(3) = 0
   '
End Sub


Sub GRADAENTREGA()
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))

   NCLIE = XVALO(TRIM$(Text1.TEXT))
   If NCLIE < 1 Then Exit Sub
   '
   If REGICLI(NCLIE) < 1 Then Exit Sub
   If YACAR%(3) < 1 Then Exit Sub
   If MODIFIC%(3) < 1 Then Exit Sub
   '
   If GRABATODO% <> 1 Then
     'Command9.SetFocus
     If COMALTER%("Se han Modificado los Datos de Entrega\del Cliente '" + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
       YACAR%(3) = 0
       Call CARDAENTREGA
       Exit Sub
     End If
     Command9.SetFocus
   End If
   '
   Call Command33_Click
   ' LO QUE SIGUE ES PARA QUE EL LUGAR DE ENTREGA
   ' SEA EL DE LA PRIMERA SUCURSAL
   If List1.ListCount > 0 Then
     List1.ListIndex = 0
     Call LIST1_Click
'     Text13 = 1
   End If
   '
   Call APERBASDAT("CLIEN")
   Screen.MousePointer = 11
   '
   SQLX$ = "SELECT * FROM Deudor12"
   SQLX$ = SQLX$ + " WHERE Nume_Cli = " & Abs(NCLIE) & ""
   Dim DEUDDDOR12 As ADODB.Recordset
   Set DEUDDDOR12 = New ADODB.Recordset
   DEUDDDOR12.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
   'With Deudor12
   With DEUDDDOR12
     On Error Resume Next
     .MoveFirst
     '.FindFirst "Nume_Cli = " & Abs(NCLIE)
     'If .NoMatch = False Then GoTo 10
     If Err < 1 Then GoTo 10
     On Error GoTo 0
     Call MENSERR(24, "Imposible Registrar")
     Screen.MousePointer = 1
     Exit Sub
   End With
   '
10 On Error GoTo 0
   'With Deudor12
   With DEUDDDOR12
     '.Edit
     !Lentre0_Cli = Left$(Text43, 48)
     !Lentre1_Cli = Left$(Text44, 48)
     !Lentre2_Cli = Left$(Text5, 48)
     !Lentre3_Cli = Left$(Text11, 10)
     If Abs(XVALO(Text45)) < 32767 Then !NTrans_Cli = XVALO(Text45)
     !Hora_Ent = Left$(Text35, 48)
     !Zona_Cli = Left$(Text37, 32)
     !Obser_Entre = Text46
     !Recar_Embal = XVALO(Text7)
     !ProvEntre_Cli = Left$(TRIM$(COTEXT2), 2)
     .Update
   End With
   Set DEUDDDOR12 = Nothing
   '

  
   MODIFIC%(3) = 0
   If InStr(EPAC$, "MINEPARTS") < 1 And InStr(EPAC$, "AHP") > 0 Then
       CONDI$ = "Nume_Cli = " & Abs(NCLIE)
       Call DuplicarRegistroDeUnaBaseAOtra("AHP", "DEUDOR12", "MINEPARTS", "DEUDOR12", CONDI$)
   End If
   '
   
   Screen.MousePointer = 1
   '
End Sub
'
Sub BLANFORMU()
   '
   Text2 = ""
   Text3 = ""
   Text6 = ""
   Text24 = ""
   Text25.TEXT = ""
   Text26.TEXT = ""
   Text27.TEXT = ""
   Text20.TEXT = ""
   Text29.TEXT = ""
   Text30.TEXT = ""
   Text31.TEXT = ""
   Text32.TEXT = ""
   Text19.TEXT = ""
   Text33.TEXT = ""
   Text34.TEXT = ""
   Text36.TEXT = ""
   Text21.TEXT = ""
   Text8.TEXT = ""
   Check1.Value = 0
   Check2.Value = 0
   '
   For kk% = 0 To 7
     COTEXT(kk%).TEXT = ""
   Next kk%
   '
   Text28.TEXT = ""
   '
   Text23 = ""
   Text35 = ""
   Text36 = ""
   Text37 = ""
   mfgCuentaCorriente.Rows = 1
   mfgComposicionDeuda.Rows = 1
   '
   Label60 = ""
   Label61 = ""
   Label62 = ""
   Label63 = ""
   Label64 = ""
   Label65 = ""
   Label66 = ""
   Label67 = ""
   Label69 = ""
   '\\\OT-08-1054-OD-19/11/08///
   Label34 = "CUIT:" 'VUELVE A LA LEYENDA ORIGINAL
   '\\\OT-08-1054-OD-FIN///
   '
   Text45 = ""
   Text43 = ""
   Text44 = ""
   Text11 = ""
   Text5 = ""
   Text35 = ""
   Text37 = ""
   Text46 = ""
   Text7 = ""
   COTEXT2 = ""
   DETEXT2 = ""
   
   '
   'SOLAPA DE COBRANZA
   For i& = 0 To 4
     Text10(i&) = ""
   Next i&
   '
   Text14 = ""
   List4.Clear
   '
   MSFlexGrid1.Rows = 1
   MSFlexGrid2.Rows = 1
   Label11.Caption = ""
   Label51.Caption = ""
   Label73.Caption = ""
   '
   Text12 = ""
   Label85 = ""
   chkRegimenFacturaCredito.Value = 0
   Call VolverTamañoCompactoGrillaInstrucciones
   MSFINSTRUC.Rows = 1
   Text13 = "" 'nro de sucursal
End Sub

'
Private Sub Form_Unload(Cancel As Integer)
   '
   Call GRAGENCLI
   Call GRANOTACLI
   Call GRADAENTREGA
   Call CARDATCLIE
   '
   Call CIERRARCH("*.*")
   If AGRECLI% > 0 Then
      Call SORTCLI(1)
      AGRECLI% = 0
   End If
   Call BAJALDISCO
   '
   If APLINVO$ <> "" Then
    Hide
   Else
    On Error Resume Next
    Unload Me
    MARCH_GES07.Enabled = True
    MARCH_GES07.SetFocus
    On Error GoTo 0
   End If
   '
End Sub

Private Sub GruCli_Click()
With OPMASCL07
  .Option1(9).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub GuaCli_Click()
   Call Command1_Click
End Sub

Private Sub LiPreCli_Click()
   With OPMASCL07
     .Option1(7).Value = True
     Call .Command16_Click
   End With
End Sub

Private Sub List_Click()
   Call Command13_Click
End Sub

Private Sub List3_DblClick()
   '
   OPX% = 0: TIPOCOM% = 0
   RELIS$ = List3.TEXT
   FE% = FECHANUM(Mid$(RELIS$, 1, 8))
   TIDOCUM$ = TRIM$(Mid$(RELIS$, 11, 16))
   If TIDOCUM$ = "Recibo Cobranza" Then TIDOCUM$ = "Recibo"
   If TIDOCUM$ = "Ajuste de Saldo" Then TIDOCUM$ = "Ajuste Saldo Cliente"
   If TIDOCUM$ = "Comp.Ctas." Then TIDOCUM$ = "Compensación de Cuentas"
   NUDOCU$ = " " + TRIM$(Mid$(RELIS$, 27, 8)) + " "
   NUDOCU1$ = Str$(XVALO(NUDOCU$))
   If InStr(UCase$(TIDOCUM$), "FACTURA") > 0 Or InStr(UCase$(TIDOCUM$), "NOTA DE") > 0 Then
     NUDOCU$ = TRIM$(NUDOCU$)
     While Len(NUDOCU$) < 8: NUDOCU$ = "0" + NUDOCU$: Wend
     If TRIM$(Mid$(RELIS$, 257, 18)) <> "" Then
        NUDOCU$ = TRIM$(Mid$(RELIS$, 257, 18))
        NUDOCU2$ = Mid$(NUDOCU$, 4, 13)
     End If
     TIPOCOM% = 1
   ElseIf InStr(UCase$(TIDOCUM$), "AJUSTE SALDO") > 0 Or InStr(UCase$(TIDOCUM$), "COMPENSACIÓN DE CUENTAS") > 0 Then
     If TRIM$(Mid$(RELIS$, 257, 18)) <> "" Then
        NUDOCU$ = TRIM$(Mid$(RELIS$, 257, 18))
        NUDOCU1$ = ""
        NUDOCU2$ = NUDOCU$
     End If
     TIPOCOM% = 2
   ElseIf InStr(UCase$(TIDOCUM$), "RECIBO") > 0 Then
     NUDOCU1$ = NUDOCU$
     NUDOCU2$ = NUDOCU$
     TIPOCOM% = 3
   End If
   OPX% = COMALTER%("Opciones de Trabajo:\\Documento Digital\Re-Emitir Comprobante")
   If OPX% = 2 Then
      Call REPRIFAC(NUDOCU$)
      Exit Sub
   End If
   '
   Screen.MousePointer = 11
   DOCUNU& = DOCID&(TIDOCUM$, NUDOCU1$, NUDOCU2$, FECHATEX$(FE%))
   If DOCUNU& < 1 Then
        If OPX% > 0 And TIPOCOM% = 1 Then ' SI OPX% ES MAYOR A 0 POR QUE ES FACTURA O NOTA DE ..
            Call MUESTRASTSCREEN(XVALO(Text1.TEXT), NUDOCU$)
        Else
            Call COMUNI("Documento No Encontrado")
        End If
   Else
        OPPXX$ = ""
        If NUDOCU$ = "00000032" Then OPPXX$ = "RECUP"
        Call MUESTRADOC(DOCUNU&, OPPXX$, TIDOCUM$, NUDOCU1$, NUDOCU2$)
   End If
   Screen.MousePointer = 1

'   VUELBUS% = 1
'   '
'3  RELIS$ = List3.TEXT
'   TIDOCUM$ = TRIM$(Mid$(RELIS$, 11, 16))
'   If TIDOCUM$ = "Recibo Cobranza" Then TIDOCUM$ = "Recibo"
'   If TIDOCUM$ = "Ajuste de Saldo" Then TIDOCUM$ = "Ajuste Saldo Cliente"
'   NUDOCU$ = " " + TRIM$(Mid$(RELIS$, 27, 8)) + " "
'   If InStr(UCase$(TIDOCUM$), "FACTURA") > 0 Or InStr(UCase$(TIDOCUM$), "NOTA DE") > 0 Then
'     NUDOCU$ = TRIM$(NUDOCU$)
'     While Len(NUDOCU$) < 8: NUDOCU$ = "0" + NUDOCU$: Wend
'   End If
'   RACLI$ = TRIM$(Left$(RASOCLI$(XVALO(Text1.TEXT)), 8))
'   Screen.MousePointer = 11
'   For UI& = ULTREG&("PDOCLST") To 1 Step -1
'     XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
'     If InStr(XX$, TIDOCUM$) > 4 Then
'       If InStr(XX$, NUDOCU$) > 4 Then
'         If InStr(XX$, RACLI$) > 16 Or VUELBUS% = 2 Then
'           DOCUNU& = CVS(Left$(XX$, 4))
'           GoTo 5
'         End If
'       End If
'     End If
'   Next UI&
'   '
'   If VUELBUS% < 2 Then
'     VUELBUS% = 2
'     GoTo 3
'   End If
'   '
'   Screen.MousePointer = 1
'   Call COMUNI("Documento no Encontrado")
'   Exit Sub
'   '
'   'Se reemplaza por este rutina el codigo subsiguiente.
'5  Call MUESTRADOC(DOCUNU&)
'   '
End Sub

Private Sub LTEXT36_Change()
   DISPONI = XVALO(LTEXT36.Caption) - XVALO(LTEXT37.Caption) - XVALO(LTEXT35.Caption)
   LTEXT23.ForeColor = &H80000008
   If DISPONI < 0 Then LTEXT23.ForeColor = &HFF&
   LTEXT23.Caption = CSTRING$(MKS$(DISPONI), 4, 11, 2, 2)
   LTEXT23.Refresh
End Sub

Private Sub NueCli_Click()
   Call Command21_Click
End Sub

Private Sub OrDBCli_Click()
Call Command15_Click
End Sub

Private Sub print_Click()
   Call Command25_Click
End Sub

Private Sub PrZoVeCli_Click()
With OPMASCL07
  .Option1(5).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Tab <> PreviousTab Then
      Call GRAGENCLI
      Call GRANOTACLI
      Call GRADAENTREGA
      Call FRESHALL
      Call CARDATCLIE
      Call CARDACOBRANZA
   End If
End Sub

Private Sub Susp_Cli_Click()
   Call Command23_Click
End Sub

Private Sub Text1_Change()
   For Index = 0 To 8: YACAR%(Index) = 0: Next Index
   Call BLANFORMU
   If XVALO(Text1.TEXT) > 0 Then
     If XVALO(Text1.TEXT) <= 999999 Then
       NCLIE = XVALO(Text1.TEXT)
       If STATCLI(NCLIE) > 0 Then
         'NC% = NCLIE
         Call CARDATCLIE
       End If
     End If
   End If
End Sub
'
Private Sub Text1_DblClick()
   '
   Call GRAGENCLI
   Call GRANOTACLI
   Call GRADAENTREGA
   '
   'Call SELECHO("DEUDOR1")
   'NCLIE = XVALO(TRIM$(VALACT1$("DEUDOR1")))
   NCLIE = ListaSeleccionClientes("", True)
   If NCLIE < 1 Then Exit Sub
   '
   NC% = NCLIE
   Text1.TEXT = TRIM$(Str$(NC%))
   '
End Sub

Private Sub Text1_GotFocus()
   '
   Call GRAGENCLI
   Call GRANOTACLI
   'Call GRADAENTREGA       'Se comento para evitar un loop del proceso de grabacion
   '
End Sub

Private Sub Text19_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text20_Change()
   '
   Call ABRECLIEN 'SE AGREGO POR QUE SI NO DABA ERROR LA CONSUTA VALOBADA
   '
   CONDI$ = "Nume_Cli = " & XVALO(Text1)
   LICRANT = XVALO(VALOBADA("Deudor12", "Lcre_Cli", CONDI$))
   If TRIM(Text20) <> "" Then
     If XVALO(Text20) <> LICRANT Then
       If DERACCE%("LIMICRE", "Modificar Limite de Credito") >= 2 Then
           MODIFIC%(0) = 1
           If TRIM$(UCase$(Text20.TEXT)) = "SUSPENDIDO" Then
               LTEXT36.Caption = Text20.TEXT
             Else
               LTEXT36.Caption = CSTRING$(MKS$(XVALO(Text20.TEXT)), 4, 11, 2, 2)
           End If
         Else
           Text20 = FORMATNUM(LICRANT, "F12.2")
       End If
     End If
   End If
End Sub

Private Sub Text20_LostFocus()
  If TRIM$(UCase$(Text20.TEXT)) <> "SUSPENDIDO" Then
    Text20.TEXT = CSTRING$(MKS$(XVALO(Text20.TEXT)), 3, 12, 2, 2)
  End If
End Sub

Private Sub Text21_Change()
   '
   Call ABRECLIEN 'SE AGREGO POR QUE SI NO DABA ERROR LA CONSUTA VALOBADA
   '
   CONDI$ = "Nume_Cli = " & XVALO(Text1)
   DESCANT = XVALO(VALOBADA("Deudor12", "Dsta_Cli", CONDI$))
   If TRIM(Text21) <> "" Then
     If XVALO(Text21) <> DESCANT Then
       If DERACCE%("DESCUSTA", "Modificar Descuento Standard") >= 2 Then
           MODIFIC%(0) = 1
         Else
           Text21 = FORMATNUM(DESCANT, "F12.2")
       End If
     End If
   End If
End Sub

Private Sub Text23_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text23_GotFocus()
   If SSTab1.Tab <> 0 Then
      SSTab1.Tab = 0
   End If
End Sub

Private Sub Text24_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text24_GotFocus()
   If SSTab1.Tab <> 0 Then
      SSTab1.Tab = 0
   End If
End Sub

Private Sub Text25_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text26_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub TEXT27_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text28_Change()
   MODIFIC(1) = 1
End Sub

Private Sub Text29_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub TEXT3_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text3_GotFocus()
   If SSTab1.Tab <> 0 Then
      SSTab1.Tab = 0
   End If
End Sub

Private Sub Text30_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text31_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text32_Change()
   MODIFIC%(0) = 1
End Sub
'\\\OT-06-0429-MH-11/10/06///
Private Sub text32_DblClick()
   If InStr(Text32.TEXT, "@") <> 0 Then
      FORMAI04.Text28 = ""
      FORMAI04.Text9 = ""
      FORMAI04.Label3.Caption = Text32.TEXT
      FORMAI04.Show 1
   End If
End Sub
'\\\OT-06-0429-MH-FIN///

Private Sub Text33_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text34_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub LText35_DblClick()
   Call RECUEC07.Text6_DblClick
End Sub
'
Sub CARESCUENTA_VIE()
    '
    If YACAR%(2) > 0 Then Exit Sub
    '
50  Screen.MousePointer = 11
    NC& = XVALO(Text1.TEXT)
    If NC& > 0 Then
        mfgCuentaCorriente.Rows = 1
        mfgComposicionDeuda.Rows = 1
        List2.Clear
        List2.Refresh
        LTEXT23.Caption = ""
        LTEXT37.Caption = ""
        LTEXT36.Caption = ""
        LTEXT35.Caption = ""
        LMCR = LIMICRE(NC&)
        LTEXT36.Caption = CSTRING$(MKS$(LMCR), 4, 11, 2, 2)
        If LMCR < 0 Then LTEXT36.Caption = " Suspendido"
        Refresh
    End If
    '
100 RECUEC07.Text1 = Text1
    List2.Clear: List3.Clear
    With RECUEC07
      For U& = 1 To .List1.ListCount
        List3.AddItem .List1.List(U& - 1)
      Next U&
      '
      For U& = 1 To .List2.ListCount
        List2.AddItem .List2.List(U& - 1)
      Next U&
      '
      Label25 = TRIM$(.Label6) 'agrego trim por que si no no se ve en el label
      Label28 = TRIM$(.Label5) 'con borderStyle = 1
      LTEXT37 = .Text10
      LTEXT35 = .Text6
    End With
    '
    DISPONI = XVALO(LTEXT36.Caption) - XVALO(LTEXT37.Caption) - XVALO(LTEXT35.Caption)
    LTEXT23.ForeColor = &H80000008
    If DISPONI < 0 Then LTEXT23.ForeColor = &HFF&
    LTEXT23.Caption = CSTRING$(MKS$(DISPONI), 4, 11, 2, 2)
    '
    '\\\OT-08-0886-OD-24/09/08-
    On Error Resume Next
    List3.TopIndex = List3.ListCount - 9
    On Error GoTo 0
    '\\\OT-08-0886-OD-FIN
    YACAR%(2) = 1
    Screen.MousePointer = 1
    '
End Sub

Private Sub Command31_Click()
    Command31.Enabled = False
    If Me.MSFINSTRUC.Height < 2350 Then
       MSFINSTRUC.Height = 2355
       MSFINSTRUC.Top = 885
       Command31.Caption = "Cerrar"
    Else
       Call VolverTamañoCompactoGrillaInstrucciones
    End If
    Command31.Enabled = True
    Call AutoAjustarFila(MSFINSTRUC, "2", 64)


End Sub

Sub AutoAjustarFila(GRID As msFlexGrid, Columnas$, CantCaracteres%)
    Dim alto As Single
    Dim ANCHO As Single
    Dim margen As Integer
    Dim Texto As String
    
    For Fila = 1 To GRID.Rows - 1
      alto = GRID.RowHeight(Fila)
      For COL = 1 To GRID.Cols - 1
        If existeValorEnString%(COL, Columnas$) > 0 Then
            Texto = GRID.TextMatrix(Fila, COL) 'obtener el texto en la celda
            LONGTEXTO% = Len(Texto)
            alto = 250
            If LONGTEXTO% > CantCaracteres Then
               Call FRATEXTO(Texto, CantCaracteres%)
               For i = 1 To CARETEX%
                TTYY$ = RETEX$(i)
                If i = 1 Then
                  GRID.TextMatrix(Fila, COL) = TTYY$ + vbCrLf
                Else
                  GRID.TextMatrix(Fila, COL) = GRID.TextMatrix(Fila, COL) + TTYY$ + vbCrLf
                End If
               Next i
               cantReng = (i - 1)
               alto = alto * cantReng
            End If
            'establecer la altura de la fila
            If alto > GRID.RowHeight(Fila) Then
               GRID.RowHeight(Fila) = alto
            End If
        End If
      Next COL
    Next Fila
    
End Sub
Sub VolverTamañoCompactoGrillaInstrucciones()
       MSFINSTRUC.Height = 670
       MSFINSTRUC.Top = 2970
       Command31.Caption = "Editar"
End Sub
Sub CARESCUENTA()
    '
    If YACAR%(2) > 0 Then Exit Sub
    '
50  Screen.MousePointer = 11
    NC& = XVALO(Text1.TEXT)
    If NC& > 0 Then
        mfgCuentaCorriente.Rows = 1
        mfgComposicionDeuda.Rows = 1
        LTEXT23.Caption = ""
        LTEXT37.Caption = ""
        LTEXT36.Caption = ""
        LTEXT35.Caption = ""
        LMCR = LIMICRE(NC&)
        LTEXT36.Caption = CSTRING$(MKS$(LMCR), 4, 11, 2, 2)
        If LMCR < 0 Then LTEXT36.Caption = " Suspendido"
        Refresh
    End If
    '
100 RECUEC07.Text1 = Text1
    With RECUEC07
      mfgCuentaCorriente.Rows = .mfgCuentaCorriente.Rows
      For FI& = 1 To .mfgCuentaCorriente.Rows - 1
        For co& = 1 To .mfgCuentaCorriente.Cols - 1
            mfgCuentaCorriente.TextMatrix(FI&, co&) = .mfgCuentaCorriente.TextMatrix(FI&, co&)
        Next co&
      Next FI&
      '
      mfgComposicionDeuda.Rows = .mfgComposicionDeuda.Rows
      For FI& = 1 To .mfgComposicionDeuda.Rows - 1
        For co& = 1 To .mfgComposicionDeuda.Cols - 1
            mfgComposicionDeuda.TextMatrix(FI&, co&) = .mfgComposicionDeuda.TextMatrix(FI&, co&)
        Next co&
      Next FI&
      '
      Label25 = TRIM$(.Label6) 'agrego trim por que si no no se ve en el label
      Label28 = TRIM$(.Label5) 'con borderStyle = 1
      LTEXT37 = .Text10
      LTEXT35 = .Text6
    End With
    '
    Call POSICIONAR_ULTIMAFILA(mfgCuentaCorriente)
    Call POSICIONAR_ULTIMAFILA(mfgComposicionDeuda)
    
    Call POSICIONAR_FILA(mfgCuentaCorriente, mfgCuentaCorriente.Rows - 1)
    Call POSICIONAR_FILA(mfgComposicionDeuda, 2)
    '
    DISPONI = XVALO(LTEXT36.Caption) - XVALO(LTEXT37.Caption) - XVALO(LTEXT35.Caption)
    LTEXT23.ForeColor = &H80000008
    If DISPONI < 0 Then LTEXT23.ForeColor = &HFF&
    LTEXT23.Caption = CSTRING$(MKS$(DISPONI), 4, 11, 2, 2)
    '
    '\\\OT-08-0886-OD-24/09/08-
    On Error Resume Next
    List3.TopIndex = List3.ListCount - 9
    On Error GoTo 0
    '\\\OT-08-0886-OD-FIN
    YACAR%(2) = 1
    Screen.MousePointer = 1
    '
End Sub

'
Private Sub Text36_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text6_GotFocus()
   FOC6% = 1
End Sub

Private Sub Text6_LostFocus()
   FOC6% = 0
End Sub

Private Sub Text7_Change()
   MODIFIC(3) = 1
End Sub

Private Sub TEXT7_LostFocus()
   Text7 = FORMATNUM$(Text7, "F5.1")
End Sub

Private Sub Text8_Change()
   CONDI$ = "Nume_Cli = " & XVALO(Text1)
   LICRCCANT = XVALO(VALOBADA("Deudor12", "LCred_CCte", CONDI$))
   If TRIM(Text8) <> "" Then
     If XVALO(Text8) <> LICRCCANT Then
       If DERACCE%("CREDCCTE", "Modificar Limite de Credito en Cuenta Corriente") >= 2 Then
           MODIFIC%(0) = 1
         Else
           Text8 = FORMATNUM(LICRCCANT, "F12.2")
       End If
     End If
   End If

End Sub

Private Sub Text8_LostFocus()
    '
    Text8.TEXT = CSTRING$(MKS$(XVALO(Text8.TEXT)), 3, 12, 2, 2)
    '
End Sub


Private Sub Text9_DblClick()
   Text9 = "DETENER"
End Sub




Private Sub Timer1_Timer()

    Timer1.Enabled = False
    
    Dim oAltaCliente As New DatosEmpresaAfip
    R = oAltaCliente.ObtenerCredencial
    'r = oAltaCliente.ObtenerDatosEmpresaAfip("30715781634", response$)
    
End Sub

Private Sub VenCli_Click()
   With OPMASCL07
     .Option1(4).Value = True
     Call .Command16_Click
   End With
End Sub

Sub GENECOCLI(Optional LONGINT%)
   If UPDATCLI% > 0 Then
      Screen.MousePointer = 11
      Call CARGALISEL("DEUDOR1", "Deudor12", "Nume_Cli/F6.0>;Raso_Cli/A48", "Stat_Cli >= 0 and Nume_Cli>0 ORDER BY Raso_Cli")
      Call CARGALISEL("INDICLIE", "Deudor12", "Codi_Cli/A12<;Raso_Cli/A50", "Stat_Cli >= 0 AND Codi_Cli <> '' ORDER BY Raso_Cli")
      '
      Dim Rst As ADODB.Recordset
      SQLX$ = "select nume_cli, raso_cli from deudor12 with(nolock) where stat_cli>0"
      Set Rst = READSET(SQLX$)
      '
      With Rst
        jj& = 0
        Do While Not .EOF
            NCLIEN = XVALO(!nume_cli)
            RASOC$ = SAFETEXT$(!raso_cli)
            '
            XX$ = REGBLAN$("DEUDOR12")
            Call REPLA(XX$, MKS$(NCLIEN), 1, 4)
            Call REPLA(XX$, RASOC$, 5, 28)
            jj& = jj& + 1
            Call GRAREG("DEUDOR12", XX$, jj&)
            .MoveNext
        Loop
      End With
      On Error Resume Next
      Rst.Close
      Set Rst = Nothing
      On Error GoTo 0
      Call SETULTREG("DEUDOR12", jj&)
      Call FRESHECHO("DEUDOR12")
      '
      UPDATCLI% = 0
      Screen.MousePointer = 1
   End If
   '
End Sub
'
Sub GENCLIANU()
   '
   Call COPYSTRU("INDIPROA", "INDICLIA")
   Call APERBASDAT("CLIEN")
   ' Busco los que estan dados de baja
   Dim ABRO1TMP As ADODB.Recordset
   Set ABRO1TMP = New ADODB.Recordset
     
   SQLX$ = "SELECT * FROM DEUDOR12 WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE Stat_Cli = -1 "
   Set ABRO1TMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With ABRO1TMP
     jj& = 0
     Do While ABRO1TMP.EOF = False
       EL2$ = REBLB$
       Call REPLA(EL2$, !Codi_Cli, 1, 12)
       Call REPLA(EL2$, !raso_cli, 13, 50)
       Call REPLA(EL2$, MKI$(!nume_cli), 63, 2)
       jj& = jj& + 1
       Call GRAREG("INDICLIA", EL2$, jj&)
     .MoveNext
     Loop
   End With
   '
80 ABRO1TMP.Close
   Set ABRO1TMP = Nothing
   '
   Call SETULTREG("INDICLIA", jj&)
   Call FILESORT("INDICLIA", "", 1, 1, "ASC")
   Call FRESHECHO("INDICLIA")
   '
End Sub

Sub GENDAPIBCLI()
    '
    Screen.MousePointer = 11
    '
    FIN& = ULTREG&("DAPIBCLI")
    If FIN& > 0 Then
        For ii& = 1 To FIN&
            XX$ = REGLEIDO$("DAPIBCLI", ii&)
            RECORD$ = REGBLAN$("DAPIBCLN")
            '
              NC = CVI(Mid$(XX$, 1, 2))
            Call REPLA(RECORD$, MKS$(NC), 1, 4)
            Call REPLA(RECORD$, Mid$(XX$, 3, 1), 5, 1)
            Call REPLA(RECORD$, Mid$(XX$, 5, 4), 6, 4)
            Call REPLA(RECORD$, Mid$(XX$, 9, 32), 10, 32)
            Call REPLA(RECORD$, Mid$(XX$, 41, 20), 42, 20)
            Call REPLA(RECORD$, Mid$(XX$, 61, 2), 62, 2)
            Call REPLA(RECORD$, Mid$(XX$, 63, 2), 64, 2)
            Call REPLA(RECORD$, Mid$(XX$, 65, 2), 66, 2)
            Call REPLA(RECORD$, Mid$(XX$, 67, 16), 68, 16)
            Call REPLA(RECORD$, Mid$(XX$, 83, 23), 84, 23)
            Call FILTERPUTRECORD("DAPIBCLN", 1, MKS$(NCLIE), RECORD$)
    
        Next ii&
        Call CIERRARCH("DAPIBCLN")
        Call CIERRARCH("DAPIBCLI")
        Exit Sub
    End If
    
    '
    SQLX$ = "SELECT * FROM DEUDOR12 WITH(NOLOCK) WHERE Stat_Cli>0"
    Dim rst1 As ADODB.Recordset
    Set rst1 = READSET(SQLX$)
    FIN& = CantRegistros("DEUDOR12", "STAT_CLI>0")
    With rst1
        Do While Not .EOF
          Call TRACE(20, U&, FIN&)
          NCLIX& = XVALO(nume_cli)
          If NCLIX& > 0 Then
            PRC$ = UCase$(PROVCLI$(NCLIX&))
            RECORD$ = REGBLAN$("DAPIBCLN")
            Call REPLA(RECORD$, MKS$(NCLIX&), 1, 4)
            If PRC$ = "B" Then
                Call REPLA(RECORD$, Chr$(1), 4, 1)
                Call REPLA(RECORD$, MKS$(100), 5, 4)
              Else
                Call REPLA(RECORD$, Chr$(0), 4, 1)
                Call REPLA(RECORD$, MKS$(0), 5, 4)
                Call REPLA(RECORD$, "EXTRAZONA", 9, 32)
            End If
            Call FILTERPUTRECORD("DAPIBCLN", 1, MKS$(NCLIX&), RECORD$)
          End If
          .MoveNext
        Loop
        rst1.Close
        Set rst1 = Nothing
        Call CIERRARCH("DAPIBCLN")
    End With

'    REBLA$ = REGBLAN$("DAPIBCLI")
'    Call REPLA(REBLA$, USERNAME$, 41, 20)
'    Call REPLA(REBLA$, MKI$(HOY(HOS$)), 61, 2)
'    Call REPLA(REBLA$, MKI$(Int(HORANUM(Time$))), 63, 2)
'    '
'    For U& = 1 To FIN&
'      Call TRACE(20, U&, FIN&)
'      XX$ = REGLEIDO$("DEUDOR1", U&)
'      NCLIX% = CVI(Left$(XX$, 2))
'      If NCLIX% > 0 Then
'        PRC$ = UCase$(PROVCLI$(NCLIX%))
'        RECORD$ = REBLA$
'        Call REPLA(RECORD$, MKI$(NCLIX%), 1, 2)
'        If PRC$ = "B" Then
'            Call REPLA(RECORD$, Chr$(1), 3, 1)
'            Call REPLA(RECORD$, MKS$(100), 5, 4)
'          Else
'            Call REPLA(RECORD$, Chr$(0), 3, 1)
'            Call REPLA(RECORD$, MKS$(0), 5, 4)
'            Call REPLA(RECORD$, "EXTRAZONA", 9, 32)
'        End If
'        Call FILTERPUTRECORD("DAPIBCLI", 1, MKI$(NCLIX%), RECORD$)
'      End If
'    Next U&
'    Call CIERRARCH("DAPIBCLI")
    Screen.MousePointer = 1
    '
End Sub

Sub GENDAPIBCA()
    '
    Screen.MousePointer = 11
    
    SQLX$ = "SELECT * FROM DEUDOR12 WITH(NOLOCK) WHERE Stat_Cli>0"
    Dim rst1 As ADODB.Recordset
    Set rst1 = READSET(SQLX$)
    FIN& = CantRegistros("DEUDOR12", "STAT_CLI>0")
    With rst1
        Do While Not .EOF
          Call TRACE(20, U&, FIN&)
          NCLIX& = XVALO(nume_cli)
          If NCLIX& > 0 Then
            PRC$ = UCase$(PROVCLI$(NCLIX&))
            RECORD$ = REGBLAN$("DAPIBCLN")
            Call REPLA(RECORD$, MKS$(NCLIX&), 1, 4)
            If PRC$ = "B" Then
                Call REPLA(RECORD$, Chr$(1), 4, 1)
                Call REPLA(RECORD$, MKS$(100), 5, 4)
              Else
                Call REPLA(RECORD$, Chr$(0), 4, 1)
                Call REPLA(RECORD$, MKS$(0), 5, 4)
                Call REPLA(RECORD$, "EXTRAZONA", 9, 32)
            End If
            Call FILTERPUTRECORD("DAPIBCA", 1, MKS$(NCLIX&), RECORD$)
          End If
          .MoveNext
        Loop
        rst1.Close
        Set rst1 = Nothing
        Call CIERRARCH("DAPIBCLN")
    End With

    Screen.MousePointer = 1

End Sub

Sub GENDAPIBMI()
    '
    Screen.MousePointer = 11
    
    SQLX$ = "SELECT * FROM DEUDOR12 WITH(NOLOCK) WHERE Stat_Cli>0"
    Dim rst1 As ADODB.Recordset
    Set rst1 = READSET(SQLX$)
    FIN& = CantRegistros("DEUDOR12", "STAT_CLI>0")
    With rst1
        Do While Not .EOF
          Call TRACE(20, U&, FIN&)
          NCLIX& = XVALO(nume_cli)
          If NCLIX& > 0 Then
            PRC$ = UCase$(PROVCLI$(NCLIX&))
            RECORD$ = REGBLAN$("DAPIBCLN")
            Call REPLA(RECORD$, MKS$(NCLIX&), 1, 4)
            If PRC$ = "B" Then
                Call REPLA(RECORD$, Chr$(1), 4, 1)
                Call REPLA(RECORD$, MKS$(100), 5, 4)
              Else
                Call REPLA(RECORD$, Chr$(0), 4, 1)
                Call REPLA(RECORD$, MKS$(0), 5, 4)
                Call REPLA(RECORD$, "EXTRAZONA", 9, 32)
            End If
            Call FILTERPUTRECORD("DAPIBMI", 1, MKS$(NCLIX&), RECORD$)
          End If
          .MoveNext
        Loop
        rst1.Close
        Set rst1 = Nothing
        Call CIERRARCH("DAPIBCLN")
    End With

    Screen.MousePointer = 1

End Sub
Sub GENDAPIBTU()
    '
    Screen.MousePointer = 11
    
    SQLX$ = "SELECT * FROM DEUDOR12 WITH(NOLOCK) WHERE Stat_Cli>0"
    Dim rst1 As ADODB.Recordset
    Set rst1 = READSET(SQLX$)
    FIN& = CantRegistros("DEUDOR12", "STAT_CLI>0")
    With rst1
        Do While Not .EOF
          Call TRACE(20, U&, FIN&)
          NCLIX& = XVALO(!nume_cli)
          If NCLIX& > 0 Then
            PRC$ = UCase$(SAFETEXT$(!Prov_Cli))
            RECORD$ = REGBLAN$("DAPIBCLN")
            Call REPLA(RECORD$, MKS$(NCLIX&), 1, 4)
            If PRC$ = "T" Then
                Call REPLA(RECORD$, Chr$(1), 4, 1)
                Call REPLA(RECORD$, MKS$(100), 5, 4)
              Else
                Call REPLA(RECORD$, Chr$(0), 4, 1)
                Call REPLA(RECORD$, MKS$(0), 5, 4)
                Call REPLA(RECORD$, "EXTRAZONA", 9, 32)
            End If
            Call FILTERPUTRECORD("DAPIBTU", 1, MKS$(NCLIX&), RECORD$)
          End If
          .MoveNext
        Loop
        rst1.Close
        Set rst1 = Nothing
        Call CIERRARCH("DAPIBTU")
    End With

    Screen.MousePointer = 1

End Sub
Sub VERIDEST()
    '
    STO1% = 0: STO2% = 0
    For i& = 1 To ULTREG&("DESTFAB")
       X$ = REGLEIDO$("DESTFAB", i&)
       CODEST$ = TRIM$(UCase$(Left$(X$, 8)))
       If CODEST$ = "STOCK" Then STO1% = 1
       If CODEST$ = "SEMIELA" Then STO2% = 1
    Next i&
    '
    If STO1% = 0 Then
       X$ = REGBLAN$("DESTFAB")
       Call REPLA(X$, "STOCK", 1, 8)
       Call REPLA(X$, "Stock de Producto Terminado", 9, 32)
       Call REGAPP("DESTFAB", X$)
    End If
    '
    If STO2% = 0 Then
       X$ = REGBLAN$("DESTFAB")
       Call REPLA(X$, "SEMIELA", 1, 8)
       Call REPLA(X$, "Stock de Semielaborados", 9, 32)
       Call REGAPP("DESTFAB", X$)
    End If
    '
    Call CIERRARCH("DESTFAB")
    '
End Sub

'Sub MUESTRASTSCREEN(NC%, NUDOCUM$, FE%, OK%)
'    '
'    FEX = FETEXCOR1$(FE%)
'    CONDI$ = "TIPOMOVIM = 'FVEN' AND fechemisi = '" & FEX & "' AND NUCLIEN = " & NC% & "  AND  codicom_lar like '%" & NUDOCUM$ & "%'"
'    FACTURAX$ = VALOBADA("CAJABANC", "CODICOM_LAR", CONDI$, "NOMESS")
'    '
'    If TRIM$(FACTURAX$) <> "" Then
'        TIPOCOM$ = Mid$(FACTURAX$, 1, 2)
'        LETR$ = Right(FACTURAX$, 1)
'        '
'        CONDI$ = "DOSECLAR='" & FACTURAX$ & "' AND NUME_CLIE= " & NC% & " AND COD_INTE > 0"
'        CANTREG& = XVALO(CantRegistros("MOVISTO", CONDI$, "NOMESS"))
'        EMP$ = "1-MDT": If CANTREG& = 0 Then EMP$ = "T-MDT"
'        ARCSQR$ = TIPOCOM$ & LETR$ & EMP$
'         '
'        If EXISTE%(LUDAT$ + ARCSQR$ + ".SQR") = 0 Then
'            Call COMUNI("Falta Archivo " & ARCSQR$ + ".SQR")
'            Exit Sub
'        End If
'       '
'       OK% = 1
'       FILTX$ = TRIM$(FACTURAX$) & ";" & TRIM$(Str$(NC%))
'       Call STDFORM(ARCSQR$, FILTX$, MODO$)
'    End If
'
'End Sub

'
Sub REPRIFAC(NUDOCU$)
   Call COMUNI("No Disponible por el Momento")
End Sub
Sub GRADACOBRANZA()
   '

   If MODIFIC%(4) < 1 Then Exit Sub
   '
   NCLIE = XVALO(TRIM$(Text1.TEXT))
   If NCLIE < 1 Then Exit Sub
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
      
   If REGICLI(NCLIE) < 1 Then Exit Sub
   '
   If GRABATODO% <> 1 Then
     If COMALTER%("Se han Modificado los Datos de Cobranzas\del Cliente '" + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
       YACAR%(4) = 0
       Call CARDACOBRANZA
       Exit Sub
     End If
     Command9.SetFocus
   End If
   '
   If DERACCE%("SAVECOB", "Grabación de Datos de Cobranzas") >= 2 Then
        CONDI$ = "Nume_Cli =" & NCLIE
        Dim rs1 As ADODB.Recordset
        Set rs1 = New ADODB.Recordset
        SQLX$ = "Select * from DEUDOR12"
        SQLX$ = SQLX$ + " WHERE " & CONDI$
30      On Error Resume Next
        rs1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
          On Error GoTo 0
          Call CapturaErrorOpen
          GoTo 30
        End If
        On Error GoTo 0
        
        '
        With rs1
           Do While Not .EOF
              !Ctact_Cob = Left$(TRIM$(Text10(0)), 48)
              !Tele_Cob = Left$(TRIM$(Text10(1)), 48)
              !Mail_Cob = Left$(TRIM$(Text10(2)), 64)
              !DiasRec_Cob = Left$(TRIM$(Text10(3)), 64)
              !DiasPago_Cob = Left$(TRIM$(Text10(4)), 64)
              !MailFactura_Cob = Left$(TRIM$(Text10(5)), 128)
              .Update
              .MoveNext
           Loop
        End With
        rs1.Close
        Set rs1 = Nothing
        If InStr(EPAC$, "MINEPARTS") < 1 And InStr(EPAC$, "AHP") > 0 Then
            CONDI$ = "Nume_Cli =" & NCLIE
            Call DuplicarRegistroDeUnaBaseAOtra("AHP", "DEUDOR12", "MINEPARTS", "DEUDOR12", CONDI$)
        End If
        
    '    Call ACTUREGISTRO("DEUDOR12", "Ctact_Cob", CONDI$, TRIM$(Text10(0)), REGAF&)
    '    Call ACTUREGISTRO("DEUDOR12", "Tele_Cob", CONDI$, TRIM$(Text10(1)), REGAF&)
    '    Call ACTUREGISTRO("DEUDOR12", "Mail_Cob", CONDI$, TRIM$(Text10(2)), REGAF&)
    '    Call ACTUREGISTRO("DEUDOR12", "DiasRec_Cob", CONDI$, TRIM$(Text10(3)), REGAF&)
    '    Call ACTUREGISTRO("DEUDOR12", "DiasPago_Cob", CONDI$, TRIM$(Text10(4)), REGAF&)
    '    Call ACTUREGISTRO("DEUDOR12", "MailFactura_Cob", CONDI$, TRIM$(Text10(5)), REGAF&)
   End If
   
   '
   MODIFIC%(4) = 0
   '
End Sub
