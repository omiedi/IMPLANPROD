VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form KARCORTE 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   0  'None
   Caption         =   "Ficha Kardex Histórico de Cortes."
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   17175
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
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   17175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Left            =   1020
      Picture         =   "KARCORTE.frx":0000
      TabIndex        =   41
      Top             =   1560
      Visible         =   0   'False
      Width           =   175
   End
   Begin VB.TextBox Text4 
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
      Left            =   495
      TabIndex        =   40
      Text            =   " "
      Top             =   1560
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   5280
      TabIndex        =   17
      Top             =   0
      Width           =   9615
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2D3C0&
         Height          =   1335
         Left            =   240
         TabIndex        =   44
         Top             =   120
         Width           =   4815
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackColor       =   &H00E2D3C0&
            Caption         =   "FICHA KARDEX DE CORTES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   480
            TabIndex        =   45
            Top             =   180
            Width           =   3855
         End
      End
      Begin VB.CommandButton CMDInfoXOF 
         Caption         =   "+"
         Height          =   285
         Left            =   9240
         Picture         =   "KARCORTE.frx":030A
         TabIndex        =   39
         ToolTipText     =   "Más Datos de Materiales para esta Orden"
         Top             =   660
         Width           =   175
      End
      Begin VB.ComboBox Combo1 
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
         ForeColor       =   &H00000000&
         Height          =   285
         ItemData        =   "KARCORTE.frx":0614
         Left            =   7080
         List            =   "KARCORTE.frx":0616
         Sorted          =   -1  'True
         TabIndex        =   38
         Top             =   280
         Width           =   2400
      End
      Begin VB.CommandButton BTNQuitaFiltro 
         Caption         =   "X"
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
         Left            =   5160
         Picture         =   "KARCORTE.frx":0618
         TabIndex        =   37
         ToolTipText     =   "Quitar Filtros"
         Top             =   1650
         Width           =   615
      End
      Begin VB.CommandButton Command43 
         Height          =   705
         Left            =   5160
         Picture         =   "KARCORTE.frx":0922
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Filtrar Según Selección Realizada"
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox Combo5 
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
         ForeColor       =   &H00000000&
         Height          =   285
         ItemData        =   "KARCORTE.frx":0A6C
         Left            =   7080
         List            =   "KARCORTE.frx":0A6E
         Sorted          =   -1  'True
         TabIndex        =   34
         Top             =   1740
         Width           =   2400
      End
      Begin VB.ComboBox Combo4 
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
         ForeColor       =   &H00000000&
         Height          =   285
         ItemData        =   "KARCORTE.frx":0A70
         Left            =   7080
         List            =   "KARCORTE.frx":0A72
         Sorted          =   -1  'True
         TabIndex        =   32
         Top             =   1380
         Width           =   2400
      End
      Begin VB.ComboBox Combo3 
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
         ForeColor       =   &H00000000&
         Height          =   285
         ItemData        =   "KARCORTE.frx":0A74
         Left            =   7080
         List            =   "KARCORTE.frx":0A76
         Sorted          =   -1  'True
         TabIndex        =   30
         Top             =   1020
         Width           =   2400
      End
      Begin VB.ComboBox Combo2 
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
         ForeColor       =   &H00000000&
         Height          =   285
         ItemData        =   "KARCORTE.frx":0A78
         Left            =   7080
         List            =   "KARCORTE.frx":0A7A
         Sorted          =   -1  'True
         TabIndex        =   28
         Top             =   660
         Width           =   1920
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2D3C0&
         Caption         =   "RANGO DE FECHAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   240
         TabIndex        =   18
         Top             =   1440
         Visible         =   0   'False
         Width           =   4740
         Begin VB.TextBox Text12 
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
            Height          =   300
            Left            =   2865
            TabIndex        =   23
            Text            =   " "
            Top             =   240
            Width           =   1065
         End
         Begin VB.TextBox Text11 
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
            Height          =   300
            Left            =   720
            TabIndex        =   22
            Text            =   " "
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton BOTSEL 
            Caption         =   "..."
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
            Left            =   1785
            Picture         =   "KARCORTE.frx":0A7C
            TabIndex        =   21
            Top             =   240
            Width           =   175
         End
         Begin VB.CommandButton BOTSEL 
            Caption         =   "..."
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
            Left            =   3960
            Picture         =   "KARCORTE.frx":0D86
            TabIndex        =   20
            Top             =   240
            Width           =   175
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Incluir Antiguos"
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
            Left            =   0
            TabIndex        =   19
            Top             =   600
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta"
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
            Left            =   2280
            TabIndex        =   25
            Top             =   315
            Width           =   510
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Desde"
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
            Left            =   -240
            TabIndex        =   24
            Top             =   315
            Width           =   855
         End
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Operario:  "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   5520
         TabIndex        =   35
         Top             =   1800
         Width           =   1395
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ubicación: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   5520
         TabIndex        =   33
         Top             =   1440
         Width           =   1395
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Operación: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   5520
         TabIndex        =   31
         Top             =   1080
         Width           =   1395
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "O.Fabricación:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   5520
         TabIndex        =   29
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Partidas: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   5870
         TabIndex        =   26
         Top             =   285
         Width           =   1065
      End
      Begin VB.Image Image1 
         Height          =   360
         Left            =   5280
         Picture         =   "KARCORTE.frx":1090
         Stretch         =   -1  'True
         ToolTipText     =   "Detalle Movimientos por Lote Activo"
         Top             =   1080
         Width           =   360
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   14
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture999 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   105
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   11
      Top             =   735
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      Caption         =   "DESCRIPCIÓN"
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
      Left            =   105
      TabIndex        =   8
      Top             =   600
      Width           =   4755
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
         Height          =   615
         Left            =   75
         TabIndex        =   9
         Top             =   240
         Width           =   4605
      End
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Left            =   3570
      Picture         =   "KARCORTE.frx":139A
      TabIndex        =   6
      Top             =   215
      Width           =   175
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
      Left            =   1020
      TabIndex        =   3
      Top             =   210
      Width           =   2535
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      FillColor       =   &H00400000&
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
      Height          =   2115
      Left            =   14920
      ScaleHeight     =   2085
      ScaleWidth      =   2235
      TabIndex        =   0
      Top             =   50
      Width           =   2265
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
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
         Height          =   200
         Left            =   240
         ScaleHeight     =   165
         ScaleWidth      =   1575
         TabIndex        =   15
         Top             =   1080
         Width           =   1605
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   0
            TabIndex        =   16
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   50
         Picture         =   "KARCORTE.frx":16A4
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Búsqueda por Texto "
         Top             =   150
         Width           =   600
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
         Height          =   620
         Left            =   1500
         Picture         =   "KARCORTE.frx":2A0A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Terminar - Salir de Consulta de Stocks"
         Top             =   150
         Width           =   600
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Print"
         Height          =   620
         Left            =   780
         Picture         =   "KARCORTE.frx":2B54
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Imprime Ficha de Movimientos de Material (Ficha Kardex)"
         Top             =   150
         Width           =   600
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3480
      OleObjectBlob   =   "KARCORTE.frx":31BE
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   -240
      OleObjectBlob   =   "KARCORTE.frx":33F2
      TabIndex        =   12
      Top             =   840
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   6885
      Left            =   120
      TabIndex        =   27
      Top             =   2160
      Width           =   17055
      _ExtentX        =   30083
      _ExtentY        =   12144
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
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
      Height          =   405
      Left            =   480
      TabIndex        =   43
      Top             =   1680
      Visible         =   0   'False
      Width           =   4365
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dep"
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
      TabIndex        =   42
      Top             =   1635
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
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
      Left            =   720
      TabIndex        =   10
      Top             =   16120
      Width           =   870
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
      Height          =   315
      Left            =   4395
      TabIndex        =   7
      Top             =   210
      Width           =   405
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
      Left            =   3830
      TabIndex        =   5
      Top             =   290
      Width           =   510
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
      Left            =   -30
      TabIndex        =   4
      Top             =   290
      Width           =   960
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
   End
   Begin VB.Menu edicion 
      Caption         =   "Edición"
   End
   Begin VB.Menu filtros 
      Caption         =   "Utilidades"
      Begin VB.Menu mnuAusteA0BobinasCorte 
         Caption         =   "Ajustar Bobinas de Corte a Fecha"
      End
   End
   Begin VB.Menu mnuRepEstad 
      Caption         =   "Reportes y Estadísticas"
      Begin VB.Menu MatTransfXSOF 
         Caption         =   "Materiales Transf. a Dep.2 y Cortes Informados x O.F."
      End
      Begin VB.Menu MatTransfXSPEDIDO 
         Caption         =   "Materiales Transf. a Dep.2 y Cortes Informados x Nro.Pedido"
      End
      Begin VB.Menu MatTransfXSCodigo 
         Caption         =   "Materiales Transf. a Dep.2 y Cortes Informados x Código"
      End
      Begin VB.Menu EstadisticoKgIngSal 
         Caption         =   "Estadístico de Kg Ing Vs. Kg Producidos"
      End
      Begin VB.Menu EstadConsVsProducidoKG 
         Caption         =   " Kg Cons.Vs. Kg Prod. x Fechas de Cierre"
      End
   End
End
Attribute VB_Name = "KARCORTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARBUS$, ABUSQUE$

Sub CargaElEncabezado()
   Campos$ = "Fecha/D8;"  '1
   Campos$ = Campos$ + "O.Fabric./A12;" '2
   Campos$ = Campos$ + "Prod.Term./A14; " '3
   Campos$ = Campos$ + "Prod.Term./A24;" ' 4
   Campos$ = Campos$ + "OP/F3;" '  5
   Campos$ = Campos$ + "Desc.OP/A12;" '     6
   Campos$ = Campos$ + "Lote/A10;" '  7
   
   Campos$ = Campos$ + "Cortes/F7;" ' 8
   Campos$ = Campos$ + "Consumo/F7;" ' 9
   Campos$ = Campos$ + "Saldo/F8;" '  10
   
   Campos$ = Campos$ + "CoorUbic/A0;" '  11
   Campos$ = Campos$ + "Ubicación/A12;" '  12
   Campos$ = Campos$ + "N.Op./F3;" '  13
   Campos$ = Campos$ + "Operario/A12;" '  14
   
   Campos$ = Campos$ + "Id/A0;" '  15
   Campos$ = Campos$ + "Transf.Linea/D8;" '  16
   Campos$ = Campos$ + "Peso/F9.2;" '  17
   Campos$ = Campos$ + "FACTOR_CORTE/A0;" '  18
   Campos$ = Campos$ + "CI/A0" '19

   
   
   Call CARGA_ENCABEZADO(KARCORTE.GRID, Campos$, KARCORTE)

End Sub

Private Sub BOTSEL_Click(Index As Integer)
    If Index < 2 Then
         Call SELECHO("SELFECHA")
         If VALACT1$("SELFECHA") <> "" Then
           If Index = 0 Then
                Text11.TEXT = VALACT1$("SELFECHA")
              Else
                Text12.TEXT = VALACT1$("SELFECHA")
           End If
         End If
      ElseIf Index = 2 Then
        Call SELECHO("TADEPOSI")
        If VALACT1$("TADEPOSI") <> "" Then
          Text4 = TRIM$(VALACT1$("TADEPOSI"))
        End If
      ElseIf Index = 3 Then
        Call SELECHO("MASTER")
        If VALACT1$("MASTER") <> "" Then
          Text1 = VALACT1$("MASTER")
        End If
    End If
End Sub

Private Sub BTNQuitaFiltro_Click()
   Combo1.Clear: Combo1.AddItem ""
   Combo2.Clear: Combo2.AddItem ""
   Combo3.Clear: Combo3.AddItem ""
   Combo4.Clear: Combo4.AddItem ""
   Combo5.Clear: Combo5.AddItem ""
   Call Text1_Change
    
End Sub


Private Sub Check1_Click()
   Call GENKARDEXCTP
End Sub

Private Sub CMDInfoXOF_Click()
   CMDInfoXOF.Enabled = False
   NORFA$ = TRIM$(Combo2.TEXT)
   If NORFA$ <> "" Then
      Call ReporteMaterialesXOF(NORFA$, 1)
   End If
   CMDInfoXOF.Enabled = True
End Sub
Private Sub Combo1_Click()

'  For i = 1 To GRID.Rows - 1
'     LOTMOV$ = TRIM$(GRID.TextMatrix(i, 4))
'     GRID.Row = i
'     For J& = 0 To GRID.Cols - 1
'        GRID.COL = J
'        If LOTMOV$ = TRIM$(Combo1.TEXT) And TRIM$(Combo1.TEXT) <> "" Then
'             GRID.CellBackColor = RGB(0, 0, 255)
'             GRID.CellForeColor = RGB(255, 255, 255)
'           Else
'             GRID.CellBackColor = RGB(255, 255, 255)
'             GRID.CellForeColor = RGB(0, 0, 0)
'        End If
'     Next J&
'   Next i
End Sub

Private Sub Command1_Click()
   Hide
End Sub

Private Sub Command2_Click()
   
   ABUSQUE$ = UCase$(TRIM$(InputBox$("Busqueda por Texto", , ARBUS$)))
   If ABUSQUE$ <> "" Then
     ARBUS$ = ABUSQUE$
     For U& = 1 To GRID.Rows - 1
      For i& = 0 To GRID.Cols - 1
       If InStr(UCase$(GRID.TextMatrix(U&, i&)), ABUSQUE$) > 0 Then
        GRID.Row = U&: GRID.COL = i&
        If GRID.TopRow > GRID.Row Then GRID.TopRow = GRID.Row
        On Error Resume Next
        VISIROWS = (GRID.Height / GRID.RowHeight(J&)) - 4
        If Err > 0 Then
         VISIROWS = 1
        End If
        On Error GoTo 0
         If VISIROWS > 0 Then
            If GRID.TopRow < GRID.Row - VISIROWS Then
              GRID.TopRow = GRID.Row - VISIROWS
            End If
         End If
        GRID.SetFocus
        Exit Sub
       End If
       Next i&
     Next U&
     '
     Call MENSERR(24, "Argumento no Encontrado.")
   End If

End Sub

Private Sub Command3_Click()
'    '
'    Call ABRECONEXION
'    '
'    For ii% = 1 To NUMAS%
'        CIXI% = ii%
'        If STATUIT(CIXI%) < 0 Then GoTo 80
'        If CantRegistros("movisto", "cod_inte=" & CIXI% & " and refercor='' and cantsalid < 0.05 and cantingre < 0.05") > 1 Then
'            OPX% = COMALTER%("Codigo " & CODEXT$(CIXI%) & " a Depurar\\No, Cancelar\Si, Continuar")
'            If OPX% = 1 Then GoTo 80
'        End If
'        '
'        Call BORRAREGISTROS("movisto", "cod_inte=" & CIXI% & " and refercor='' and cantsalid < 0.05 and cantingre < 0.05", REGAF&)
'        '
'80  Next ii%

'    Call ABRECONEXION
'    '
'    For ii% = 1 To NUMAS%
'        CIXI% = ii%
'        If STATUIT(CIXI%) < 0 Then GoTo 80
'        If CantRegistros("movisto", "cod_inte=" & CIXI% & " and refercor='' and (cantsalid > 0.05 or cantingre > 0.05)") > 1 Then
'            OPX% = COMALTER%("Codigo " & CODEXT$(CIXI%) & " a Depurar\\No, Cancelar\Si, Continuar")
'            If OPX% = 1 Then GoTo 80
'        End If
'        sqlx$ = "select * from movisto where cod_inte=" & CIXI% & " and refercor='' and (cantsalid > 0.05 or cantingre > 0.05)"
'        '
'        Dim rst As ADODB.Recordset
'        Set rst = New ADODB.Recordset
'        rst.Open FILTSQL$(sqlx$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
'        '
'        With rst
'            Do While Not .EOF
'                dopricor$ = TRIM$(SAFETEXT$(!dopricor))
'                DEPO% = XVALO(!depomovi)
'                CANTI# = XVALO(!cantsalid)
'                '
'                CONDI$ = "cod_inte=" & CIXI% & " and depomovi=" & DEPO% & " and dopricor='" & dopricor$ & "'"
'                careg& = CantRegistros("movisto", CONDI$)
'                '
'                If careg& > 1 Then
'                    !cantsalid = 0
'                    !cantingre = 0
'                    .Update
'                     MsgBox "entro " & careg& & vbCrLf & CONDI$
'                End If
'                .MoveNext
'            Loop
'        End With
'        On Error Resume Next
'        rst.Close
'        Set rst = Nothing
'        On Error GoTo 0
'80  Next ii%
    '
    Call COMUNI("proceso finalizado")
    '
End Sub

Sub ReporteMaterialesXOF(NORFA$, Optional VieneConSoloUnaOFab%, Optional VieneDePedido%, Optional CI1%)
'HABRIA QUE AGREGARLE EL CONSUMO DEL MOVISTO
10  J& = 0
    NORFA$ = TRIM$(NORFA$)
    Campos$ = "CI./a0;" '1
    Campos$ = Campos$ + "Código/a16;" '2
    Campos$ = Campos$ + "Descripción/A32;" '3
    Campos$ = Campos$ + "Lote./A12;"      '4
    Campos$ = Campos$ + "N°OP/f6;"       '5
    Campos$ = Campos$ + "Desc.OP./A24;"   '6
    Campos$ = Campos$ + "Metros.Ingres./F12.1;"      '7
    Campos$ = Campos$ + "Corte/F10.1;"       '8
    Campos$ = Campos$ + "Cant.Cortes/f8;"        '9
    Campos$ = Campos$ + "Metros_Cortes./F12;"     '10
    Campos$ = Campos$ + "Kg.Transf/F10.2;" '11
    Campos$ = Campos$ + "Kg.Cortes/F10.2;" '12
'    Campos$ = Campos$ + "Importe/f12.2" '12
'EN ESTA CONSULTA MUESTRO LO TRANSFERIDO PARA  LA OF Y EL RESTO LO TRAIGO POR VALOBADA
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
    DEPOCORTE% = XVALO(Control("", "DEPODEST"))
    SQLX$ = "SELECT m.cod_inte,m.cod_exte,m.descrip, m.idenlote,m.ID_OPERACION" 'REALIZA EL SUM SOLO EN EL CASO QUE SEAN TRANSFERENCIAS
    SQLX$ = SQLX$ + " ,SUM(CASE WHEN m.CODIMOVI = 'TR' THEN m.cantingre - m.cantsalid ELSE 0 END) AS Material_Transferido"
    SQLX$ = SQLX$ + " ,SUM(CASE WHEN m.CODIMOVI = 'TR' AND Cantingre > 0 THEN pesokg ELSE -pesokg END) AS SumaPesokg"

    SQLX$ = SQLX$ + "  FROM movisto m WITH(NOLOCK)"
    SQLX$ = SQLX$ + "  WHERE depomovi = " & DEPOCORTE%
    'SQLX$ = SQLX$ + "  AND DOSECLAR = '" & NORFA$ & "' AND CODIMOVI = 'TR'"
    If VieneDePedido% > 0 Then
       SQLX$ = SQLX$ + "  AND DOSECLAR IN (" & NORFA$ & ")" 'ya viene con comillas simples
    ElseIf VieneConSoloUnaOFab% > 0 Then
       SQLX$ = SQLX$ + "  AND DOSECLAR IN ('" & NORFA$ & "')" 'le agrego las comillas simples
    ElseIf CI1% > 0 Then
       SQLX$ = SQLX$ + "  AND COD_INTE=  " & CI1% 'FILTRA POR CODIGO
    End If
    SQLX$ = SQLX$ + "  AND CODIMOVI = 'TR'"
    SQLX$ = SQLX$ + "  Group By m.cod_inte, m.cod_exte, m.idenlote,m.descrip,m.ID_OPERACION"
    SQLX$ = SQLX$ + "   HAVING SUM(CASE WHEN m.CODIMOVI = 'TR' THEN m.cantingre - m.cantsalid ELSE 0 END) > 0"
    SQLX$ = SQLX$ + "  ORDER BY m.cod_inte,m.ID_OPERACION,m.idenlote"
    Set rs = READSET(SQLX$)
    With rs
     Do While Not .EOF
        CI1% = XVALO(!cod_inte)
        Codigo$ = SAFETEXT$(!Cod_Exte)
        DESXRI$ = SAFETEXT$(!Descrip)
        LOTE$ = SAFETEXT$(!idenlote)
        CantTransferidaAlDepo2 = XVALO(!Material_Transferido)
        NOPER% = XVALO(!ID_OPERACION)
         CONDI$ = "idsubor  = '" & NORFA$ & "'  and codielem = " & CI1% & " AND IDENLOTE = '" & LOTE$ & "' AND NUMEOPER = " & NOPER% & " AND CANTSALID <= 0 "
        Descrioper$ = VALOBADA("INFOFAB", "DESCOPER", CONDI$, "NOMESS")
        CORTES = XVALO(VALOBADA("INFOFAB", "CORTE", CONDI$, "NOMESS"))
        CantidadRealizadas = XVALO(VALOBADA("INFOFAB", "sum(Cantreal)", CONDI$, "NOMESS"))
        CANTIMETROS = FORMATNUM$(CORTES * CantidadRealizadas, "F12.1")
        PesoInfofab = FORMATNUM$(XVALO(VALOBADA("INFOFAB", "SUM(PESO)", CONDI$, "NOMESS")), "F12.1")
        SumaPesokg = FORMATNUM$(XVALO(!SumaPesokg), "F12.1") 'del movisto
        
        J& = J& + 1
        GRILLAGRAL.GRID.Rows = J& + 1
        GRILLAGRAL.GRID.TextMatrix(J&, 1) = CI1%
        GRILLAGRAL.GRID.TextMatrix(J&, 2) = Codigo$
        GRILLAGRAL.GRID.TextMatrix(J&, 3) = DESXRI$
        GRILLAGRAL.GRID.TextMatrix(J&, 4) = LOTE$
        GRILLAGRAL.GRID.TextMatrix(J&, 5) = NOPER%
        GRILLAGRAL.GRID.TextMatrix(J&, 6) = Descrioper$
        GRILLAGRAL.GRID.TextMatrix(J&, 7) = FORMATNUM$(CantTransferidaAlDepo2, "F12.1")
  
        GRILLAGRAL.GRID.TextMatrix(J&, 8) = CORTES
        GRILLAGRAL.GRID.TextMatrix(J&, 9) = CantidadRealizadas
        GRILLAGRAL.GRID.TextMatrix(J&, 10) = CANTIMETROS
        GRILLAGRAL.GRID.TextMatrix(J&, 11) = SumaPesokg
        GRILLAGRAL.GRID.TextMatrix(J&, 12) = PesoInfofab
       .MoveNext
     Loop
    End With
    rs.Close
    Set rs = Nothing
    GRILLAGRAL.Width = 12500
    If J& = 0 Then Call COMUNI("No Hay Registros Con la Condición Seleccionada"): Exit Sub
    '*******************************************
    GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.TXTBUSCAR = ""
    
    GRILLAGRAL.Caption = "Detalle de Material Transferido para la Orden: " & NORFA$
    
    GRILLAGRAL.Show 1
    Dim MSF As msFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(MSF, Campos$, "INFOSCRP", TX$, "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
    End If
    
'   NORFA$ = OrdenFabricacionAbiertas$(1)
'   If TRIM$(NORFA$) <> "" Then GoTo 10
End Sub

Private Sub Command43_Click()
   Call Text1_Change
End Sub








Private Sub Command8_Click()
   '
   Command8.Enabled = False
   ZBLA$ = REGBLAN$("COKARDEX")
   JJ& = 0
   FIN& = GRID.Rows - 1 ' List1.ListCount
   '\\\OT-06-0390-MH-13/09/06///
   TODO% = 0
'   For KU& = 1 To FIN&
'     If List1.Selected(KU& - 1) = True Then GoTo 10
'   Next KU&
   TODO% = 1
   GoTo 20
   '
10 OPXX% = COMALTER%("Seleccione Tipo de Listado\\Texto Seleccionado\Todo\Cancelar")
   TODO% = OPXX% - 1
   If TODO% > 1 Then GoTo 99
   '
20 For J% = 1 To GRID.Rows - 1   'List1.ListCount
     '
     J2& = J% + 1
     Call TRACE(24, J2&, FIN&)
     '
'     If TODO% = 1 Or List1.Selected(J% - 1) Then
        '
        'X$ = List1.List(J% - 1)
        FEC1% = FECHANUM(GRID.TextMatrix(J%, 0))  ' Left$(X$, 8))
        CM1$ = GRID.TextMatrix(J%, 1)             ' Mid$(X$, 11, 2)
        DOPRI1$ = GRID.TextMatrix(J%, 2)          ' Mid$(X$, 15, 12)
        DOSEC1$ = GRID.TextMatrix(J%, 3)          ' Mid$(X$, 29, 12)
        REFE1$ = GRID.TextMatrix(J%, 5)           ' Mid$(X$, 43, 26)
        DEP1% = XVALO(GRID.TextMatrix(J%, 6))     ' Mid$(X$, 70, 3))
        ENTRA1# = XVALO(GRID.TextMatrix(J%, 7))   ' Mid$(X$, 75, 10))
        SALE1# = XVALO(GRID.TextMatrix(J%, 8))    ' Mid$(X$, 85, 10))
        SALD1# = XVALO(GRID.TextMatrix(J%, 9))    ' Mid$(X$, 95, 12))
        '
        YY$ = ZBLA$
        Call REPLA(YY$, MKI$(FEC1%), 1, 2)
        CI1% = CODINT%(Text1.TEXT)
        Call REPLA(YY$, MKI$(CI1%), 3, 2)
        Call REPLA(YY$, CM1$, 21, 2)
        Call REPLA(YY$, DOPRI1$, 23, 10)
        Call REPLA(YY$, DOSEC1$, 33, 10)
        Call REPLA(YY$, REFE1$, 43, 30)
        Call REPLA(YY$, Chr$(DEP1%), 83, 1)
        Call REPLA(YY$, MKD$(ENTRA1#), 88, 8)
        Call REPLA(YY$, MKD$(SALE1#), 96, 8)
        Call REPLA(YY$, MKD$(SALD1#), 104, 8)
        JJ& = JJ& + 1
        Call GRAREG("COKARDEX", YY$, JJ&)
        '
'     End If
     '
   Next J%
   Call SETULTREG("COKARDEX", JJ&)
  
   '\\\OT-06-0390-MH-FIN///
   '
   Call HEADERS("COKARDEX", "")
   If Val(Text4) > 0 Or ECOARCH$("TADEPOSI", Chr$(Val(Text4))) <> "" Then
      Call HEADERS("COKARDEX", "Deposito: " & TRIM$(Text4) & " - " & ECOARCH$("TADEPOSI", Chr$(Val(Text4))))
   End If
   Call HEADERS("COKARDEX", "Artículo: " + Text1.TEXT + " - " + Label6.Caption)
   Call HEADERS("COKARDEX", "Período: " + Text11.TEXT + " - " + Text12.TEXT)
   Call FINAL("*LICARDEX")
   '
99 Command8.Enabled = True
   '
 End Sub

Private Sub Filtroporproveedor_Click()
    CI1% = CODINT%(Text1)
    If CI1% < 1 Then Exit Sub
    
    Campos$ = "Cta.Prov/F9;R.Social/A40"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO, , , ANTOT)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO, , , ANTOT)
'    If ANTOT < 5500 Then ANTOT = 5500
'    FRMZELECHO.Width = ANTOT + 1500

    SQLX$ = " EXECUTE ProveedorSegunCodigo  " & CI1%
    Dim rs As New ADODB.Recordset
    Set rs = FLEXCONN.Execute(SQLX$)
    FRMZELECHO.MSF1.Rows = 1
    J& = 0
    '
    With rs
      Do While Not .EOF
        NPROV = XVALO(!NumProv)
        RAZONPROVEEDOR$ = rasocli$(-1 * NPROV)
        '
        J& = J& + 1
        FRMZELECHO.MSF2.Rows = J& + 1
        FRMZELECHO.MSF2.TextMatrix(J&, 1) = NPROV
        FRMZELECHO.MSF2.TextMatrix(J&, 2) = RAZONPROVEEDOR$
        .MoveNext
      Loop
    End With
    rs.Close
    Set rs = Nothing
    If J& = 0 Then
       Call COMUNI("No se Detectaron Entregas Ingresos Con Proveedor Para el Código Activo")
       Exit Sub
    End If
    
    FRMZELECHO.Caption = "Proveedores de Código: " & CODEXT$(CI1%)
    FRMZELECHO.Show 1
    
    PROVED = XVALO(RESP_ZELE_VECT(1))
    If PROVED > o Then
       Label8 = PROVED
       Label9 = rasocli$(-1 * PROVED)
    Else
       Label8 = ""
    End If

End Sub

Private Sub EstadConsVsProducidoKG_Click()

'CONSIDERA ORDENES DE FABRICACION CERRADAS QUE SU FECHA DE CIERRE ESTE DENTRO DEL RANGO
'select * from ordefabr where (StatuOrf = 3  OR CanProy <= CanApro)
    'and FechCierr >= '01/JUL/2025' AND FechCierr <= '31/07/2025'




         Call DESHASFECHA(Des%, Has%, PERIO$)
         If PERIO$ = "" Then Exit Sub
         '
         Desde = FETEXCOR1$(Des%) ' SE MODIFICO EL FORMATO FECHA DE NUMERICO A FECHA CORTA
         Hasta = FETEXCOR1$(Has%)
         '
         
         Call ABRECONEXION
        '
        StringOF$ = StringDeOrdenesFabricacionCerradasEnRangodeFechas$(Des%, Has%)
        If StringOF$ = "" Then
           Call COMUNI("No se encontraron Ordenes de Fabricación en el rango seleccionado")
           Exit Sub
        End If
        
        Screen.MousePointer = 11
20      Campos$ = "O.Fabricación/A12"  '1
        Campos$ = Campos$ + ";Pedido/F10" '2
        Campos$ = Campos$ + ";Código/A16" '3
        Campos$ = Campos$ + ";Descripción/A32" '4
'        Campos$ = Campos$ + ";KG.Transf./F10" '5
        Campos$ = Campos$ + ";KG.Cons.Corte/F12" '5
        Campos$ = Campos$ + ";KG.Prod.Term./F12" '6
        Campos$ = Campos$ + ";CI/A0" '7
        Campos$ = Campos$ + ";%/F6.2" '8
        
        

        J& = 0
        Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
        Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
        
        Call TEXTOLOTES$(StringOF$, ",")
        For i = 0 To UBound(CADENATEX)
            NORFA$ = TRIM$(CADENATEX(i))
            NORFA$ = REPLACAR(NORFA$, "'", "")
            CONSUMIDO = PesoKg_Consumidos(NORFA$)
            Producido = PesoKg_Producidos(NORFA$)
            NP& = NroPedidoSegunOf&(NORFA$)
            CI1% = CODIFAB%(NORFA$)
            Codigo$ = CODEXT$(CI1%)
            DESC$ = DESCRIT0$(CI1)
            J& = J& + 1
            GRILLAGRAL.GRID.Rows = J& + 1
            GRILLAGRAL.GRID.TextMatrix(J&, 1) = NORFA$
            GRILLAGRAL.GRID.TextMatrix(J&, 2) = NP&
            GRILLAGRAL.GRID.TextMatrix(J&, 3) = Codigo$
            GRILLAGRAL.GRID.TextMatrix(J&, 4) = DESC$
            GRILLAGRAL.GRID.TextMatrix(J&, 5) = FORMATNUM$(CONSUMIDO, "F10.1")
            GRILLAGRAL.GRID.TextMatrix(J&, 6) = FORMATNUM$(Producido, "F10.1")
            GRILLAGRAL.GRID.TextMatrix(J&, 7) = CI1%
            GRILLAGRAL.GRID.TextMatrix(J&, 8) = FORMATNUM$(PORCENT_ENTRE_2_VALORES(Producido, CONSUMIDO), "F6.2")
        Next i

   GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
   GRILLAGRAL.mnuMenu1.Visible = True
   GRILLAGRAL.GRID.ZOrder 0
   GRILLAGRAL.TXTBUSCAR = ""
   TX$ = "Kg. Consumidos vs. Kg. Producidos (Solo Considera O.Fab. Cerradas) en este Rango Desde: " & FECHATEX$(Des%) & " Hasta: " & FECHATEX$(Has%)
   
   GRILLAGRAL.Caption = TX$
   Screen.MousePointer = 1
   NORFA$ = ""
   GRILLAGRAL.Show 1

   
   Dim MSF As msFlexGrid
   If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(MSF, Campos$, "CONVSPRO", " CONSUMOS VS PRODUCIDO", "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
        Exit Sub
   End If
99 Exit Sub




End Sub

Private Sub EstadisticoKgIngSal_Click()
        Call SELECHO("SELFECHA")
        Desde% = FECHANUM(VALACT1$("SELFECHA"))
        If Desde% < 1 Then Exit Sub
        DesdeFecha$ = FETEXCOR1$(Desde%)
        
20      Campos$ = "O.Fabricación/A12"  '1
        Campos$ = Campos$ + ";Pedido/F10" '2
        Campos$ = Campos$ + ";Código/A16" '3
        Campos$ = Campos$ + ";Descripción/A32" '4
        Campos$ = Campos$ + ";KG.Transf./F10" '5
        Campos$ = Campos$ + ";KG.Inf.Corte/F12" '6
        Campos$ = Campos$ + ";KG.Prod.Term./F12" '7
        Campos$ = Campos$ + ";CI/A0" '8
        Campos$ = Campos$ + ";%/F6.2" '9
        
        

        J& = 0
        Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
        Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)

        'ESTA CONSULTA TRAE TODAS LAS OF, QUE SU PRIMER CIERRE EN EL MOVISTO SE PRODUJO DENTRO DE LAS FECHA
        'EJ, fecha a partir de 01/01/24 SI LA OF SE GENERO ANTES DEL DESDE PERO SU PRIMER CIERRE FUE
        'DENTRO DEL RANGO LA TRAE,SI SU PRIMER CIERRE FUE FUERA DEL RANGO NO LA TRAE
        '  ******   SOLO ORDENES CERRADAS   ******
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
             NORFA$ = SAFETEXT$(!doprilar)
             If STATORFA%(NORFA$) <> 3 Then GoTo 60
             CI1% = XVALO(!cod_inte)
             CONDI$ = "depomovi = 2   AND CODIMOVI = 'TR' and DOSECLAR = '" & NORFA$ & "'"
              KilosGramosTransferidosAlDeposito2XOF = XVALO(VALOBADA("MOVISTO", "SUM(CASE WHEN CODIMOVI = 'TR' AND Cantingre > 0 THEN pesokg ELSE -pesokg END)", CONDI$, "NOMESS"))
              'AGREGO ESTA CONDICION AND DescOper <> 'LINER' PARA QUE NO TOME EN CUENTA LOS LINER QUE ESTAN CON PESO EN INFOFAB
             CONDI$ = "IDSUBOR = '" & NORFA$ & "' AND DescOper <> 'LINER'"
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
             GRILLAGRAL.GRID.TextMatrix(J&, 9) = FORMATNUM$(PORCENT_ENTRE_2_VALORES(KilosGramosTransferidosAlDeposito2XOF, KilosGramosConsumidosXOF), "F6.2")


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
   NORFA$ = ""
   GRILLAGRAL.Show 1
   
   NORFA$ = TRIM$(RESP_ZELE_VECT(1))
   If NORFA$ <> "" Then
     'PARA QUE SE VEA DIFERENTE Y NO SEA CONFUSO
     GRILLAGRAL.Frame1.BackColor = &H9FBE85
     GRILLAGRAL.Frame2.BackColor = &H9FBE85
     GRILLAGRAL.Option1.BackColor = &H9FBE85
     GRILLAGRAL.Option2.BackColor = &H9FBE85
     GRILLAGRAL.Option3.BackColor = &H9FBE85
     GRILLAGRAL.BackColor = &H9FBE85
     Call ReporteMaterialesXOF(NORFA$, 1)
     For i& = 0 To UBound(RESP_ZELE_VECT)
       RESP_ZELE_VECT(i&) = ""
     Next i&
     GRILLAGRAL.Frame1.BackColor = &H8000000F
     GRILLAGRAL.Frame2.BackColor = &H8000000F
     GoTo 20
   End If
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

Private Sub Form_Load()

   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me, Picture3)
   '
   Text11 = "01/01/80"
'   FEINIKAR$ = TRIM$(CONTROL("", "FEINIKAR"))
'   FF% = FECHANUM(FEINIKAR$)
'   If FF% > 0 Then Text11 = FECHATEX$(FF%)
   HOII% = HOY(HOS$)
   Text12 = HOS$
   Call Text4_Change
   '\\\OT-08-0990-OD-30/10/08///
   LOMA% = XVALO(Control$("MASTER", "LOMAXCOD"))
   If LOMA% > 15 Then Text1.MaxLength = LOMA%
   '\\\OT-08-0990-OD-FIN///
   '
   
   '
   Call CargaElEncabezado
   Combo2.Clear: Combo2.AddItem ""
   Combo3.Clear: Combo3.AddItem ""
   Combo4.Clear: Combo4.AddItem ""
   Combo5.Clear: Combo5.AddItem ""
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   Me.GRID.Width = Me.Width - 300
   Me.GRID.Height = Me.Height - (Frame2.Height + 700)
   Picture3.Left = Me.Width - Picture3.Width
   Frame2.Left = Me.Width - (Picture3.Width + Frame2.Width)
   On Error GoTo 0
End Sub


Private Sub GRID_DblClick()
    '
    iRow& = GRID.Row
    CANTING = XVALO(GRID.TextMatrix(iRow&, 7))
    CANTSAL = XVALO(GRID.TextMatrix(iRow&, 8))
    COMO$ = TRIM$(GRID.TextMatrix(iRow&, 1))
    BRE$ = TRIM$(GRID.TextMatrix(iRow&, 2))
    CI1% = CODINT%(TRIM$(Text1))
    
    '
    If CANTING > 0 Then
       IDLOT$ = TRIM$(GRID.TextMatrix(iRow&, 4))
       If IDLOT$ <> "" Then
         OPX% = COMALTER("Opciones de Trabajo\\Consulta de Consumos Por Lote\Ver Datos Relacionados Según Tipo de Movimiento")
         If OPX% < 1 Then Exit Sub
         If OPX% = 1 Then
             On Error Resume Next
             CONSUMLOT07.Label4 = Text1
             If Err Then Exit Sub
             On Error GoTo 0
             CONSUMLOT07.Label1 = IDLOT$
             CONSUMLOT07.Show 1
             Exit Sub
         End If
       End If
     ElseIf CANTSAL > 0 Then
       On Error Resume Next
       LOCONSUM09.Label4 = TRIM$(GRID.TextMatrix(iRow&, 2))
       If Err Then Exit Sub
       On Error GoTo 0
       OPX% = COMALTER("Opciones de Trabajo\\Consulta de Consumos Por Lote\Ver Datos Relacionados Según Tipo de Movimiento")
       If OPX% < 1 Then Exit Sub
       If OPX% = 1 Then
        LOCONSUM09.Label5 = TRIM$(GRID.TextMatrix(iRow&, 0))
        LOCONSUM09.Label6 = TRIM$(GRID.TextMatrix(iRow&, 4))
        LOCONSUM09.Show 1
        Exit Sub
       End If
    End If
    '
    Select Case TRIM$(UCase(COMO$))
      Case "BR"
        Campos$ = "O.Compra/F9;Cta./F8;Proveedor/A32;Id.Lote/A16"
        Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO, , , ANTOT)
        Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO, , , ANTOT)
        FRMZELECHO.Caption = "Orden de Compra Según Recepcion: " & BRE$
        FRMZELECHO.Width = 8600
        '
        NBR& = XVALO(Mid$(BRE$, 4))
        SQLX$ = "SELECT * FROM BOLRECEP WITH(NOLOCK) WHERE NUMEBORE = " & NBR& & " AND COD_INTE = " & CI1%
        Set rs = READSET(SQLX$): k& = 0
        Do While Not rs.EOF
          k& = k& + 1
          FRMZELECHO.MSF2.Rows = k& + 1
          FRMZELECHO.MSF2.TextMatrix(k&, 1) = XVALO(rs!NUMOC)
          FRMZELECHO.MSF2.TextMatrix(k&, 2) = XVALO(rs!NumProv)
          FRMZELECHO.MSF2.TextMatrix(k&, 3) = SAFETEXT$(rs!Raso_Prov)
          FRMZELECHO.MSF2.TextMatrix(k&, 4) = SAFETEXT$(rs!idenlote)
          rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        FRMZELECHO.Show 1
        NUOC& = XVALO(RESP_ZELE_VECT(1))
        DOCUNU& = NUOC&
        TIDOCU$ = "NOTA DE PEDIDO"
        NUDOCU$ = TRIM$(str$(DOCUNU&))
        Call MUESTRADOC(0, "", TIDOCU$, NUDOCU$, NUDOCU$)

    End Select
    
End Sub

Private Sub GRID_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   i& = GRID.MouseRow
   J& = GRID.MouseCol
   If i& = 0 Or i& > GRID.Rows Then Exit Sub
   
   If Button = 2 Then
       '
       GRID.Row = i&
       GRID.COL = J&
       GRID.CellBackColor = &H80FFFF   ' LA PINTA DE AMARILLA
       strContenidoCelda$ = GRID.TextMatrix(GRID.MouseRow, GRID.MouseCol)
       Clipboard.Clear
       Clipboard.SetText strContenidoCelda
   End If
   
End Sub

Sub AJUASEMBLI(CIXI%, CANTING, CANTSAL, TIMOI$, DOPRIX$, FEMOX$, LOTEA$)
'         TTXX$ = List1.TEXT
         CANTING = XVALO(CANTING)
         CANTSAL = XVALO(CANTSAL)
         '
         If CIXI% < 1 Then Exit Sub
         '
         If CANTSAL > 0 Or TIMOI$ = "AJ" Or TIMOI$ = "TR" Then
            Call ARMALILOT(CIXI%)
            Call FRESHECHO("!LILOTSE")
            Screen.MousePointer = 1
            If ULTREG("!LILOTSE") < 1 Then
              Call MENSERR(24, "No Hay Lotes Disponibles\para el Ajuste Propuesto.")
              Exit Sub
            End If
            '
            Call SELECHO("!LILOTSE")
            VDEVU$ = VALACT1$("!LILOTSE")
            VDEV2$ = VALACT2$("!LILOTSE")
            If TRIM$(VDEVU$) = "" Then Exit Sub
            '
            CAMAXI = XVALO(Mid$(VDEV2$, 21, 12))
            If CANTSAL > CAMAXI Then
               Call MENSERR(24, "Cantidad del Movimiento Supera\Saldo Fisico del Lote")
               Exit Sub
            End If
            '
            FEMOI% = FECHANUM(FEMOX$)
            TTZZ$ = "Aprueba Aplicar el Lote '" + VDEVU$ + "'\al Comprobante '" + TIMOI$ + "' - '" + DOPRIX$ + "'"
            TTZZ$ = TTZZ$ + "\de Fecha '" + FEMOX$ + "' ?\\No, Cancelar\Si, Aplicar"
            If COMALTER%(TTZZ$) = 2 Then
               SQLX$ = "Cod_Inte = " & CIXI%
               SQLX$ = SQLX$ + " AND CodiMovi = '" & TIMOI$ & "'"
               SQLX$ = SQLX$ + " AND DoPrilar = '" & DOPRIX$ & "'"
               SQLX$ = SQLX$ + " AND FechMov= '" & FETEXCOR1$(FEMOI%) & "'"
               If CANTSAL >= 0.05 Then SQLX$ = SQLX$ + " AND CantSalid = " & TRIM$(FORMATNUM(CANTSAL, "F12.4"))
               SQLX$ = SQLX$ + " AND IdenLote = '" & LOTEA$ & "' "
               '
               Screen.MousePointer = 11
               Call ACTUREGISTRO("MOVISTO", "IdenLote", SQLX$, VDEVU$, REGAF&, "NOMESS")
               If REGAF& > 0 Then
                    Call ACSALOTE(CIXI%, VDEVU$, "NOMESS")
                    Call GENKARDEXCTP
               Else
                    Call MENSERR(24, "Actualización Imposible")
               End If
               Screen.MousePointer = 1
            End If
         End If
End Sub
Private Sub Image1_Click()
   IDLOT$ = TRIM$(Combo1.TEXT)
   If IDLOT$ <> "" Then
     On Error Resume Next
     CONSUMLOT07.Label4 = Text1
     If Err Then Exit Sub
     On Error GoTo 0
     CONSUMLOT07.Label1 = IDLOT$
     CONSUMLOT07.Show 1
   End If
End Sub


Private Sub List1_DblClick()
   TTXX$ = List1.TEXT
   CANTING = XVALO(Mid$(TTXX$, 75, 10))
   CANTSAL = XVALO(Mid$(TTXX$, 85, 10))
   '
   If CANTING > 0 Then
       IDLOT$ = TRIM$(Mid$(TTXX$, 129, 16))
       If IDLOT$ <> "" Then
         On Error Resume Next
         CONSUMLOT07.Label4 = Text1
         If Err Then Exit Sub
         On Error GoTo 0
         CONSUMLOT07.Label1 = IDLOT$
         CONSUMLOT07.Show 1
       End If
     ElseIf CANTSAL > 0 Then
       On Error Resume Next
       LOCONSUM09.Label4 = TRIM$(Mid$(TTXX$, 15, 14))
       If Err Then Exit Sub
       On Error GoTo 0
       LOCONSUM09.Label5 = Mid$(TTXX$, 1, 8)
       LOCONSUM09.Label6 = TRIM$(Mid$(TTXX$, 43, 26))
       LOCONSUM09.Show 1
   End If
   '
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
      '
      Linea$ = List1.List(List1.ListIndex)
      FECHA1$ = Mid$(Linea$, 1, 8)
      FF% = FECHANUM(FECHA1$)
      FECHA2$ = FETEXCOR1$(FF%)
      LOTE$ = Mid(Linea$, 15, 12)
      '
      CONDI$ = "DOPRICOR ='" & TRIM$(LOTE$) & "' AND FECHMOV='" & FECHA2$ & "'"
      '
      CODUSU% = XVALO(VALOBADA("MOVISTO", "NUMUSUAR", CONDI$))
      USU$ = USERNAM$(CODUSU%)
      FechaReal$ = Mid(SAFETEXT$(VALOBADA("MOVISTO", "FEREMOVI", CONDI$)), 1, 10)
      HORA$ = Mid(SAFETEXT$(VALOBADA("MOVISTO", "FEREMOVI", CONDI$)), 12, 8)
      '
      Call REPLA(VDEF$, USU$, 1, 20)
      Call REPLA(VDEF$, FechaReal$, 21, 10)
      Call REPLA(VDEF$, HORA$, 31, 8)
      Datos$ = OBJPLA("DATOSMOV", VDEF$)
   End If
   '
End Sub


Private Sub Label8_Change()
   Call GENKARDEXCTP
End Sub

Private Sub Label8_DblClick()
   Label8 = ""
   Label9 = ""
End Sub

Private Sub MatTransfXSCodigo_Click()
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, "Código/A18;Descripción/a48", FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Código/A18;Descripción/a48", FRMZELECHO)
    FRMZELECHO.Caption = "Códigos:"
    '****************************************************************
    
    Dim obj As New RECORXET
    Set rs1 = obj.RS_MAESTRODEARTICULOS()
    Call Carga_MSF_Recordset(rs1, "Código/A18;Descripción/a48", FRMZELECHO.MSF2)
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    
    FRMZELECHO.Width = 8500
    FRMZELECHO.Show 1
    Codigo$ = RESP_ZELE_VECT(1)
    CI1% = CODINT%(Codigo$)
    If CI1% > 0 Then
        If CI1% > 0 Then Call ReporteMaterialesXOF(NORFAS$, 0, 0, CI1%)
    End If
End Sub

Private Sub MatTransfXSOF_Click()
   NORFA$ = OrdenFabricacionAbiertas$(1)
   Call ReporteMaterialesXOF(NORFA$, 1)

End Sub

Private Sub MatTransfXSPEDIDO_Click()
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, "Nro.Pedido/f9;Referencia/a48", FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Nro.Pedido/f9;Referencia/a48", FRMZELECHO)
    FRMZELECHO.Caption = "Pedidos Todos:"
    '****************************************************************
    
    Dim obj As New RECORXET
    Set rs1 = obj.RS_PEDIDOS_TODOS()
    Call Carga_MSF_Recordset(rs1, "Nro.Pedido/f9;Referencia/a48", FRMZELECHO.MSF2)
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    
    FRMZELECHO.Width = 8500
    FRMZELECHO.Show 1
    NRO& = XVALO(RESP_ZELE_VECT(1))
    If NRO& > 0 Then
        NORFAS$ = StringDeOrdenesFabricacionSegunPedido$(NRO&)
        If NORFAS$ <> "" Then
           Call ReporteMaterialesXOF(NORFAS$, 0, 1)
        Else
           Call COMUNI("No Hay Ordenes de Fabricación Asociadas a Este Pedido")
        End If
    End If
End Sub

Sub Text1_Change()
   '
   CI1% = CODINT%(Text1.TEXT)
   Label6.Caption = DESCRIT0$(CI1%)
   Label5.Caption = Unimed$(CI1%)
   Call FRESHALL
   '
   Call CargaElEncabezado
   
   Call GENKARDEXCTP
End Sub

Private Sub Text1_DblClick()
   Call BOTSEL_Click(3)
End Sub

Private Sub Text11_Change()
   'List1.Clear
   If FECHANUM(Text11) <> 0 Then
     Call GENKARDEXCTP
   End If
End Sub

Private Sub Text11_DblClick()
   Call BOTSEL_Click(0)
End Sub

Private Sub Text12_Change()
   'List1.Clear
   If FECHANUM(Text12) <> 0 Then
     Call GENKARDEXCTP
   End If
End Sub

Private Sub Text12_DblClick()
   Call BOTSEL_Click(0)
End Sub

Private Sub TEXT3_Change()

End Sub

Private Sub Text4_Change()
   'List1.Clear
   If Val(Text4) < 256 Then
     If Val(Text4) < 1 Or ECOARCH$("TADEPOSI", Chr$(Val(Text4))) = "" Then
         Label4 = "Todos los Depósitos"
       Else
         Label4 = ECOARCH$("TADEPOSI", Chr$(Val(Text4)))
     End If
     Call GENKARDEXCTP
   End If
End Sub

Private Sub Text4_DblClick()
   Call BOTSEL_Click(2)
End Sub

Sub GENKARDEXCTP()
   '
Call GENKARDEXCTP2
Exit Sub
   Static DEPXX%, PRIDEP%
   If PRIDEP% < 1 Then
     Call CARDEPOS
     PRIDEP% = COSU%(1)
   End If
   '
   Combo1.Clear
   VECOMBO$ = Space$(1024)
   PINI& = 1
   '
   LU$ = LUDAT$
   If EXISTE%(LU$ + "COKARDEX.RFS") < 1 Then
     Call COPYSTRU("MOVISTO", "COKARDEX")
   End If
   REBLA$ = REGBLAN$("COKARDEX")
   '
   'List1.Clear
   CODL$ = Text1
5  CIXX% = CODINT%(CODL$)
   If CIXX% < 1 Then Exit Sub
   '
   DEPXX% = Val(Text4)
   If DEPXX% < 0 Then DEPXX% = 0
   '
   DESFE% = FECHANUM(Text11)
   Desde = CDbl(CVDAT(DESFE%))
   HASFE% = FECHANUM(Text12)
   Hasta = CDbl(CVDAT(HASFE%))
   If Desde > Hasta Then
     Call MENSERR(24, "Rango de Fechas no Válido")
     Exit Sub
   End If
   '
   INCLUVIE% = 0
   If Check1.Value = 1 Then INCLUVIE% = 1
   '
6  CIDD% = CIXX%
   JJ& = 0
   SARRAS# = 0
   INCLUPOS% = 0
   GRID.Redraw = False
   If CIDD% > 0 Then
     Call ABRESTOCKS
'     Call ABRESTOCKVIE
     iRow& = 0: GRID.Rows = 23
     For i& = 1 To GRID.Rows - 1
        For J& = 0 To GRID.Cols - 1
           GRID.TextMatrix(i&, J&) = ""
        Next J&
     Next i&
     If INCLUVIE% = 1 Then
     '
        sqlx1$ = " SELECT FECHMOV,  CANTINGRE, CANTSALID, COD_INTE, IdenLote, CodiMovi, DoPriCor, DoSecCor, ReferCor, VALOUNIT, MoneVal, "
        sqlx1$ = sqlx1$ + " DEPOMOVI,  Nume_Clie, Nume_Prov, FeReMovi,  NumUsuar "

        sqlx1$ = sqlx1$ + " FROM MOVIVIE WITH(NOLOCK) "
        sqlx1$ = sqlx1$ + "WHERE Cod_Inte = " & CIDD% & " "
        If DEPXX% > 0 Then
          If DEPXX% = PRIDEP% Then
              sqlx1$ = sqlx1$ + "AND (DepoMovi = 1 OR DepoMovi = 0 OR DepoMovi IS NULL)"
            Else
              sqlx1$ = sqlx1$ + "AND DepoMovi = " & DEPXX% & " "
          End If
        End If
        sqlx1$ = sqlx1$ + "AND CodiMovi <> 'KU' "
        sqlx1$ = sqlx1$ + "ORDER BY FechMov , FeReMovi ASC "
        '
'        Dim MOVIVIETMP As ADODB.Recordset
'        Set MOVIVIETMP = FLEXCONN.Execute(FILTSQL$(SQLX1$))
        Set MOVIVIETMP = READSET(sqlx1$)
        With MOVIVIETMP
            Do While Not .EOF
              If CVINT(!FechMov) > HASFE% Then
                If INCLUPOS% = 0 Then
                  OPXX% = COMALTER%("Se Han Detectado Movimientos POSTERIORES\a la Fecha Límite Seleccionada.\\No Considerarlos\Incluirlos")
                  If OPXX% < 2 Then INCLUPOS% = (-1)
                  If OPXX% > 1 Then INCLUPOS% = 1
                End If
              End If
              '
              If CVINT(!FechMov) < DESFE% Then
                  SARRAS# = SARRAS# + XVALO(!CANTINGRE) - XVALO(!CANTSALID)
                ElseIf CVINT(!FechMov) <= HASFE% Or INCLUPOS% = 1 Then
                  YY$ = REBLA$
                  Call REPLA(YY$, MKI$(CVINT(!FechMov)), 1, 2)
                  Call REPLA(YY$, MKI$(XVALO(!cod_inte)), 3, 2)
                  Call REPLA(YY$, SAFETEXT$(!idenlote), 5, 16)
                  Call REPLA(YY$, SAFETEXT$(!CODIMOVI), 21, 2)
                  Call REPLA(YY$, SAFETEXT$(!DoPriCor), 23, 10)
                  Call REPLA(YY$, SAFETEXT$(!DoSecCor), 33, 10)
                  Call REPLA(YY$, SAFETEXT$(!ReferCor), 43, 30)
                  Call REPLA(YY$, MKD$(XVALO(!VALOUNIT)), 73, 8)
                    MONEX$ = "$": If XVALO(!MoneVal) = 2 Then MONEX$ = "U$"
                  Call REPLA(YY$, MONEX$, 81, 2)
                  Call REPLA(YY$, Chr$(XVALO(!depomovi)), 83, 1)
                  Call REPLA(YY$, MKI$(XVALO(!NUME_CLIE)), 84, 2)
                  Call REPLA(YY$, MKI$(XVALO(!Nume_Prov)), 86, 2)
                  Call REPLA(YY$, MKD$(XVALO(!CANTINGRE)), 88, 8)
                  Call REPLA(YY$, MKD$(XVALO(!CANTSALID)), 96, 8)
                  Call REPLA(YY$, MKD$(0), 104, 8)
                  Call REPLA(YY$, Mid$(SAFETEXT$(!DoPriCor), 11, 4), 112, 4)
                  Call REPLA(YY$, Mid$(SAFETEXT$(!DoSecCor), 11, 4), 116, 4)
                  Call REPLA(YY$, CVINT(!FeReMovi), 122, 2)
                  Call REPLA(YY$, Chr$(XVALO(!NumUsuar)), 126, 1)
                  JJ& = JJ& + 1
                  Call GRAREG("#COKARDEX", YY$, JJ&)
              End If
            .MoveNext
            Loop
        End With
        MOVIVIETMP.Close
        Set MOVIVIETMP = Nothing
     End If
     
     SQLX$ = " SELECT FECHMOV,  CANTINGRE, CANTSALID, COD_INTE, IdenLote, CodiMovi, DoPriCor, DoSecCor, ReferCor, VALOUNIT, MoneVal, "
     SQLX$ = SQLX$ + " DEPOMOVI,  Nume_Clie, Nume_Prov, FeReMovi,  NumUsuar "
     SQLX$ = SQLX$ + " FROM MOVISTO WITH(NOLOCK) "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIDD% & " AND (VERIFICADO <> -1 OR VERIFICADO IS NULL) "
     If DEPXX% > 0 Then
       If DEPXX% = PRIDEP% Then
           SQLX$ = SQLX$ + "AND (DepoMovi = 1 OR DepoMovi = 0 OR DepoMovi IS NULL)"
         Else
           SQLX$ = SQLX$ + "AND DepoMovi = " & DEPXX% & " "
       End If
     End If
     If INCLUVIE% = 1 Then
       SQLX$ = SQLX$ + "AND CodiMovi <> 'KU' "
     End If
     SQLX$ = SQLX$ + "ORDER BY FechMov , FeReMovi ASC "

     '
'     Dim MoviTemp As ADODB.Recordset
'     Set MoviTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
     Set MoviTemp = READSET(SQLX$)
     With MoviTemp
         Do While Not .EOF
           If CVINT(!FechMov) > HASFE% Then
             If INCLUPOS% = 0 Then
               OPXX% = COMALTER%("Se Han Detectado Movimientos POSTERIORES\a la Fecha Límite Seleccionada.\\No Considerarlos\Incluirlos")
               If OPXX% < 2 Then INCLUPOS% = (-1)
               If OPXX% > 1 Then INCLUPOS% = 1
             End If
           End If
           '
           If CVINT(!FechMov) < DESFE% Then
               SARRAS# = SARRAS# + XVALO(!CANTINGRE) - XVALO(!CANTSALID)
             ElseIf CVINT(!FechMov) <= HASFE% Or INCLUPOS% = 1 Then
               YY$ = REBLA$
               Call REPLA(YY$, MKI$(CVINT(!FechMov)), 1, 2)
               Call REPLA(YY$, MKI$(XVALO(!cod_inte)), 3, 2)
               Call REPLA(YY$, SAFETEXT$(!idenlote), 5, 16)
               Call REPLA(YY$, SAFETEXT$(!CODIMOVI), 21, 2)
               Call REPLA(YY$, SAFETEXT$(!DoPriCor), 23, 10)
               Call REPLA(YY$, SAFETEXT$(!DoSecCor), 33, 10)
               Call REPLA(YY$, SAFETEXT$(!ReferCor), 43, 30)
               Call REPLA(YY$, MKD$(XVALO(!VALOUNIT)), 73, 8)
                 MONEX$ = "$": If XVALO(!MoneVal) = 2 Then MONEX$ = "U$"
               Call REPLA(YY$, MONEX$, 81, 2)
               Call REPLA(YY$, Chr$(XVALO(!depomovi)), 83, 1)
               Call REPLA(YY$, MKI$(XVALO(!NUME_CLIE)), 84, 2)
               Call REPLA(YY$, MKI$(XVALO(!Nume_Prov)), 86, 2)
               Call REPLA(YY$, MKD$(XVALO(!CANTINGRE)), 88, 8)
               Call REPLA(YY$, MKD$(XVALO(!CANTSALID)), 96, 8)
               Call REPLA(YY$, MKD$(0), 104, 8)
               Call REPLA(YY$, Mid$(SAFETEXT$(!DoPriCor), 11, 4), 112, 4)
               Call REPLA(YY$, Mid$(SAFETEXT$(!DoSecCor), 11, 4), 116, 4)
               Call REPLA(YY$, CVINT(!FeReMovi), 122, 2)
               Call REPLA(YY$, Chr$(XVALO(!NumUsuar)), 126, 1)
               JJ& = JJ& + 1
               ' Call GRAREG("!COKARDEX", YY$, JJ&)
               Call GRAREG("#COKARDEX", YY$, JJ&)
           End If
         .MoveNext
         Loop
     End With
     MoviTemp.Close
     Set MoviTemp = Nothing
     'Call SETULTREG("!COKARDEX", JJ&)
     Call SETULTREG("#COKARDEX", JJ&)
     Screen.MousePointer = 11
     '
     If INCLUVIE% = 1 Then
       Call BLOQARCH("COKARDEX")
       'Call TRACENTRAL("COKARDEX", "")
       Call FILEDUMP("COKARDEX")
       Call FILESORT("COKARDEX", "", 1, 1, "ASC")
       'Call TRALOCAL("COKARDEX", "")
       For KKUL& = 1 To JJ&
         CKX$ = REGLEIDO$("COKARDEX", KKUL&)
         Call GRAREG("#COKARDEX", CKX$, KKUL&)
       Next KKUL&
       Call CIERRARCH("COKARDEX")
     End If
     '
     iRow& = 0: GRID.Rows = 23
     For i& = 1 To GRID.Rows - 1
        For J& = 0 To GRID.Cols - 1
           GRID.TextMatrix(i&, J&) = ""
        Next J&
     Next i&
     Combo1.TEXT = "": Call Combo1_Click
     ' FIN& = ULTREG&("!COKARDEX")
     FIN& = ULTREG&("#COKARDEX")
     CASALDO# = 0
     If Abs(SARRAS#) >= 0.05 Then
        ' TTXX$ = Space$(256)
        FEX = DIANTE(DESFE%)
        ' Call REPLA(TTXX$, FECHATEX$(FEX), 1, 8)
        ' IROW& = IROW& + 1
        ' GRID.Row = IROW&
        ' GRID.Col = 0
        GRID.TextMatrix(iRow, 0) = FECHATEX$(FEX)
        REFEM$ = "Saldo Anterior al " + FECHATEX$(DESFE%)
        ' Call REPLA(TTXX$, REFEM$, 43, 36)
        ' GRID.Col = 5: GRID.TEXT = REFEM$
        GRID.TextMatrix(iRow, 5) = REFEM$
        '
        CASALDO# = SARRAS#
        ' Call REPLA(TTXX$, FORMATNUM$(CASALDO#, "F12.1"), 95, 12)
        ' Call REPLA(TTXX$, FORMATNUM$(CASALDO#, "F12." & CNTDC$(CIXX%) & ""), 95, 12)
        ' GRID.Col = 9: GRID.TEXT = TRIM$(FORMATNUM$(CASALDO#, "F12." & CNTDC$(CIXX%) & ""))
        GRID.TextMatrix(iRow, 9) = TRIM$(FORMATNUM$(CASALDO#, "F12." & CNTDC$(CIXX%) & ""))
        ' List1.AddItem TTXX$
     End If
     '
     GRID.Rows = FIN& + 1
     For KU& = 1 To FIN&
       Call TRACE(20, KU&, FIN&)
       ' YY$ = REGLEIDO$("!COKARDEX", KU&)
       YY$ = REGLEIDO$("#COKARDEX", KU&)
       ' TTXX$ = Space$(256)
       FEX = CVI(Left$(YY$, 2))
       ' Call REPLA(TTXX$, FECHATEX$(FEX), 1, 8)
       iRow& = iRow& + 1
       If iRow > GRID.Rows - 1 Then GRID.Rows = iRow& + 1
       ' GRID.Row = IROW&
       ' GRID.Col = 0:
       GRID.TextMatrix(iRow, 0) = FECHATEX$(FEX)
         CMOVI$ = Mid$(YY$, 21, 2)
       GRID.TextMatrix(iRow&, 1) = CMOVI$
         DOPRI$ = Mid$(YY$, 23, 10) + Mid$(YY$, 112, 4)
       GRID.TextMatrix(iRow&, 2) = DOPRI$
         DOSEC$ = Mid$(YY$, 33, 10) + Mid$(YY$, 116, 4)
       GRID.TextMatrix(iRow&, 3) = DOSEC$
         REFEM$ = Mid$(YY$, 43, 30)
       GRID.TextMatrix(iRow&, 5) = REFEM$
       ' Call REPLA(TTXX$, REFEM$, 43, 26)
         DEPOM$ = FORMATNUM$(Asc(Mid$(YY$, 83, 1)), "I3")
       GRID.TextMatrix(iRow&, 6) = TRIM$(DEPOM$)
         CAINGRE# = CVD(Mid$(YY$, 88, 8))
       GRID.TextMatrix(iRow&, 7) = TRIM$(FORMATNUM$(CAINGRE#, "F12." & CNTDC$(CIXX%) & ""))
         CASALI# = CVD(Mid$(YY$, 96, 8))
       GRID.TextMatrix(iRow&, 8) = TRIM$(FORMATNUM$(CASALI#, "F12." & CNTDC$(CIXX%) & ""))
         CASALDO# = CASALDO# + CAINGRE# - CASALI#
       GRID.TextMatrix(iRow&, 9) = TRIM$(FORMATNUM$(CASALDO#, "F12." & CNTDC$(CIXX%) & ""))
         IDLOT$ = Mid$(YY$, 5, 16)
       GRID.TextMatrix(iRow&, 4) = TRIM$(IDLOT$)
       '
       If TRIM$(IDLOT$) <> "" Then
          If InStr(VECOMBO$, IDLOT$) < 1 Then
             If PINI& > Len(VECOMBO$) Then VECOMBO$ = VECOMBO$ + Space$(1024)
             Mid$(VECOMBO$, PINI&, 16) = IDLOT$
             PINI& = PINI& + 16
          End If
       End If
       '
       Call CompletaCombo(IDLOT$)
'       List1.AddItem TTXX$
     Next KU&
     Call FREEMEM("COKARDEX")
     '
     GRID.Redraw = True
     
     For KU& = 1 To PINI& - 1 Step 16
        Combo1.AddItem Mid$(VECOMBO$, KU&, 16)
     Next KU&
     '
     ' ESTO ES PARA POSICIONARSE AL FINAL DE LA LISTA
     'TPXI& = List1.ListCount - 19
     'If TPXI& < 0 Then TPXI& = 0
     'List1.TopIndex = TPXI&
     '
     TPXI& = GRID.Rows - 22: If TPXI& < 1 Then TPXI& = 1
     On Error Resume Next
     GRID.TopRow = TPXI&
     On Error GoTo 0
     '
     Screen.MousePointer = 1
     '
Exit Sub
'       Call SETULTREG("COKARDEX", JJ&)
'       Call CIERRARCH("COKARDEX")
'       Screen.MousePointer = 1
'       Call CONECRUN("*COKARDEX/" + CODD$, "Ficha Kardex Indirecta") ' + "/WAIT")
'       Screen.MousePointer = 1
'    Else
'       Call MENSERR(24, "Codigo: " + TRIM$(CODD$) + "\sin Movimientos Según Condición SQL.")
   End If
   '
   Screen.MousePointer = 1
   GoTo 5
   '
End Sub

Sub CompletaCombo(IDLOT$)
    '
    If TRIM$(IDLOT$) <> "" Then
       IDLOT1$ = TRIM$(IDLOT$)
       Dim EXISTE As Boolean
       EXISTE = False
       '
'       For i = 0 To Combo1.ListCount - 1
'          If Combo1.List(i) = IDLOT1$ Then
'             EXISTE = True
'             Exit Sub
'          End If
'       Next i
       '
       If EXISTE = False Then
          Combo1.AddItem IDLOT1$
       End If
    End If
    '
End Sub

Sub GENKARDEXCTP2()
   '
   Call ABRECONEXION
   '
   FILTROLOTE$ = TRIM$(Combo1.TEXT)
   FILTRONORFA$ = TRIM$(Combo2.TEXT)
   FILTRONOPER% = XVALO(Combo3.TEXT)
   FILTROUBICA% = XVALO(Combo4.TEXT)
   FILTROOPERARIO% = XVALO(Combo5.TEXT)
      
   Static DEPXX%, PRIDEP%
   If PRIDEP% < 1 Then
      Call CARDEPOS
      PRIDEP% = COSU%(1)
   End If
   '
   VECOMBO$ = Space$(1024)
   PINI& = 1
   '
   iRow& = 1: GRID.Rows = 23
   For i& = 1 To GRID.Rows - 1
     GRID.Row = i
      For J& = 1 To GRID.Cols - 1
          GRID.COL = J
         GRID.TextMatrix(i&, J&) = ""
      Next J&
   Next i&
   '
   CODL$ = Text1
5  CI1% = CODINT%(CODL$)
   If CI1% < 1 Then Exit Sub
   '
   DEPXX% = XVALO(Text4)
   If DEPXX% < 0 Then DEPXX% = 0
   '
   DESFE% = FECHANUM(Text11): Desde = FETEXCOR1$(DESFE%)
   ' DESDE = CDbl(CVDAT(DESFE%))
   HASFE% = FECHANUM(Text12): Hasta = FETEXCOR1$(HASFE%, "HASTA")
   ' HASTA = CDbl(CVDAT(HASFE%))
   If DESFE% > HASFE% Then
     Call MENSERR(24, "Rango de Fechas no Válido")
     Exit Sub
   End If
   '
   INCLUVIE% = 0:   If Check1.Value = 1 Then INCLUVIE% = 1
   '
6  CIDD% = CI1%
'   condi0$ = "Cod_Inte = " & CIDD% & " "
'   If DEPXX% > 0 Then
'      If DEPXX% = PRIDEP% Then
'          condi0$ = condi0$ + "AND (DepoMovi = 1 OR DepoMovi = 0 OR DepoMovi IS NULL)"
'        Else
'          condi0$ = condi0$ + "AND DepoMovi = " & DEPXX% & " "
'      End If
'   End If
   '
   Screen.MousePointer = 11

   Call ABRESTOCKS
   iRow& = 0: GRID.Rows = 30
   For i& = 1 To GRID.Rows - 1
      For J& = 0 To GRID.Cols - 1
         GRID.TextMatrix(i&, J&) = ""
      Next J&
   Next i&
   ACUIN = 0: ACUSA = 0
   'CONSIDERO LOS FILTROS DE LOS COMBOS
   CONDI$ = " INFOFAB_ID > 0"
   If FILTROLOTE$ <> "" Then CONDI$ = CONDI$ + " AND IDENLOTE = '" & FILTROLOTE$ & "'"
   If FILTRONORFA$ <> "" Then CONDI$ = CONDI$ + " AND IDSUBOR = '" & FILTRONORFA$ & "'"
   If FILTRONOPER% > 0.05 Then CONDI$ = CONDI$ + " AND NUMEOPER = " & FILTRONOPER%
   If FILTROUBICA% > 0.05 Then CONDI$ = CONDI$ + " AND IDCOORDUBIC = " & FILTROUBICA%
   If FILTROOPERARIO% > 0.05 Then CONDI$ = CONDI$ + " AND NUMOPERARIO = " & FILTROOPERARIO%
   


'   ACUIN = ACUIN + XVALO(VALOBADA("INFOFAB", "SUM(CANTREAL)", CONDI$ & " AND FEHORINF < '" & Desde & "'", "NOMESS"))
'   ACUSA = ACUSA + XVALO(VALOBADA("INFOFAB", "SUM(CANTSALID)", CONDI$ & " AND FEHORINF < '" & Desde & "'", "NOMESS"))
   '
   
'   SARRAS# = ACUIN - ACUSA
   
'   If Abs(SARRAS#) >= 0.05 Then
'       iRow& = iRow& + 1
'       FEX = DIANTE(DESFE%)
'       GRID.TextMatrix(iRow&, 1) = FECHATEX$(FEX)
'       REFEM$ = "Saldo Anterior al " + FECHATEX$(DESFE%)
'       GRID.TextMatrix(iRow, 4) = REFEM$
'       CASALDO# = SARRAS#
'       GRID.TextMatrix(iRow, 10) = TRIM$(FORMATNUM$(CASALDO#, "F12." & CNTDC$(CI1%) & ""))
'       iRow& = iRow& + 1  ' PARA QUE DEJE UNO LIBRE
'   End If
   JJ& = 0

'   '

   GRID.Redraw = False
   
20  sqlx1$ = " SELECT codielem,IDSUBOR,COD_INTE,COD_EXTE,DESCRIP,NUMEOPER,DESCOPER,CANTREAL,CANTSALID,IDENLOTE,FEHORINF,IDCOORDUBIC,COORDUBIC,NUMOPERARIO,NAMEOPERARIO,INFOFAB_ID,FeTranLinea,FACTORCORTE"
    sqlx1$ = sqlx1$ + ",CASE WHEN CANTSALID > 0 THEN -PESO ELSE PESO END AS PESO"
    sqlx1$ = sqlx1$ + " FROM INFOFAB WITH(NOLOCK) "
    sqlx1$ = sqlx1$ + " WHERE  " & CONDI$ ' TRAE LA CONDICION DE LOS FILTROS DE LOS COMBOS
    sqlx1$ = sqlx1$ + " AND codielem = " & CI1%
    sqlx1$ = sqlx1$ + " AND FEHORINF >= '" & Desde & "' "
    sqlx1$ = sqlx1$ + " AND FEHORINF <= '" & Hasta & "' "
    sqlx1$ = sqlx1$ + " ORDER BY FEHORINF,IDSUBOR,NUMEOPER,IDENLOTE"
    
    'LIMPIO LOS CASILLEROS ANTES DE VOLVERLOS A SELECCIONAR
    Combo1.Clear:  Combo1.AddItem ""
    Combo2.Clear: Combo2.AddItem ""
    Combo3.Clear: Combo3.AddItem ""
    Combo4.Clear: Combo4.AddItem ""
    Combo5.Clear: Combo5.AddItem ""


   Set MOVIVIETMP = READSET(sqlx1$)
   GRID.Rows = MOVIVIETMP.RecordCount + 1

   With MOVIVIETMP
      Do While Not .EOF
              If CVINT(!FEHORINF) > HASFE% Then
                If INCLUPOS% = 0 Then
                  OPXX% = COMALTER%("Se Han Detectado Movimientos POSTERIORES\a la Fecha Límite Seleccionada.\\No Considerarlos\Incluirlos")
                  If OPXX% < 2 Then INCLUPOS% = (-1)
                  If OPXX% > 1 Then INCLUPOS% = 1
                End If
              End If
         IDLOT$ = SAFETEXT$(!idenlote)
'         If XVALO(Label8) > 0 Then
'            If PROVELOTE%(IDLOT$) <> XVALO(Label8) Then GoTo 65
'         End If
             '
         NORFA$ = SAFETEXT$(!idsubor)
         NOPER% = XVALO(!NUMEOPER)
         DESOPER$ = SAFETEXT$(!descoper)
         Call REPLA(TXOPEACION$, TRIM$(CStr(NOPER%)), 1, 6)
         Call REPLA(TXOPEACION$, TRIM$(DESOPER$), 8, 16)
         NUBIC% = XVALO(!IDCOORDUBIC)
         DESCUBIC$ = SAFETEXT$(!COORDUBIC):
         Call REPLA(TXUBICACION$, TRIM$(CStr(NUBIC%)), 1, 6)
         Call REPLA(TXUBICACION$, TRIM$(DESCUBIC$), 8, 16)
         NUMOPERARIO% = XVALO(!NUMOPERARIO)
         NameOperario$ = SAFETEXT$(!NameOperario)
         Call REPLA(TXOPERARIO$, TRIM$(CStr(NUMOPERARIO%)), 1, 6)
         Call REPLA(TXOPERARIO$, TRIM$(NameOperario$), 8, 16)

         iRow& = iRow& + 1
         If iRow& > GRID.Rows - 1 Then
            GRID.Rows = iRow& + 1
         End If
         GRID.TextMatrix(iRow&, 1) = FECHATEX$(CVINT(!FEHORINF))
         GRID.TextMatrix(iRow&, 2) = NORFA$
         GRID.TextMatrix(iRow&, 3) = SAFETEXT$(!Cod_Exte)
         GRID.TextMatrix(iRow&, 4) = SAFETEXT$(!Descrip)
         GRID.TextMatrix(iRow&, 5) = NOPER%
         GRID.TextMatrix(iRow&, 6) = DESOPER$
         GRID.TextMatrix(iRow&, 7) = SAFETEXT$(!idenlote)
            CAINGRE# = XVALO(!CANTREAL)
            CASALI# = XVALO(!CANTSALID)
            CASALDO# = CASALDO# + XVALO(!CANTREAL) - XVALO(!CANTSALID)
         GRID.TextMatrix(iRow&, 8) = TRIM$(FORMATNUM$(CAINGRE#, "F12." & CNTDC$(CI1%) & ""))
         GRID.TextMatrix(iRow&, 9) = TRIM$(FORMATNUM$(CASALI#, "F12." & CNTDC$(CI1%) & ""))
         GRID.TextMatrix(iRow&, 10) = TRIM$(FORMATNUM$(CASALDO#, "F12." & CNTDC$(CI1%) & ""))
         
         GRID.TextMatrix(iRow&, 11) = NUBIC%
         GRID.TextMatrix(iRow&, 12) = DESCUBIC$
         GRID.TextMatrix(iRow&, 13) = NUMOPERARIO%
         
         GRID.TextMatrix(iRow&, 14) = NameOperario$
         GRID.TextMatrix(iRow&, 15) = SAFETEXT$(!INFOFAB_ID)
         GRID.TextMatrix(iRow&, 16) = FECHATEX$(CVINT(!FeTranLinea))
         
         GRID.TextMatrix(iRow&, 17) = FORMATNUM$(XVALO(!Peso), "F12.2")
         GRID.TextMatrix(iRow&, 18) = XVALO(!FACTORCORTE)
         GRID.TextMatrix(iRow&, 19) = XVALO(!codielem)
         
         Call AgregarSinRepetirValorComboBox(Combo1, IDLOT$)
         Call AgregarSinRepetirValorComboBox(Combo2, NORFA$)
         Call AgregarSinRepetirValorComboBox(Combo3, TXOPEACION$)
         Call AgregarSinRepetirValorComboBox(Combo4, TXUBICACION$)
         Call AgregarSinRepetirValorComboBox(Combo5, TXOPERARIO$)
         
         
65    .MoveNext
      Loop
   End With
   MOVIVIETMP.Close
   Set MOVIVIETMP = Nothing
   If BASEAC$ = "MOVIVIE" Then
      BASEAC$ = "MOVISTO"
      GoTo 20
   End If
   '
   Call Duplicados1(Combo1)
   GRID.Redraw = True
   ' ESTO ES PARA POSICIONARSE AL FINAL DE LA LISTA
   TPXI& = GRID.Rows - 22: If TPXI& < 1 Then TPXI& = 1
   On Error Resume Next
   GRID.TopRow = TPXI&
   On Error GoTo 0
   '
   'RECUPERO LOS FILTROS PARA DEJARLOS VISIBLES EN LOS COMBOS
   Combo1.TEXT = FILTROLOTE$
   Combo2.TEXT = FILTRONORFA$
   Combo3.TEXT = FILTRONOPER%
   Combo4.TEXT = FILTROUBICA%
   Combo5.TEXT = FILTROOPERARIO%



   Screen.MousePointer = 1
'   Call AUTOADJUST
   '
End Sub


Sub AUTOADJUST()
   '
   TEPRUEBA.Font = GRID.Font
   TEPRUEBA.FontSize = GRID.FontSize
   TEPRUEBA.FontBold = GRID.FontBold
   '
   For COLUI& = 1 To GRID.Cols
      ACOLU = 0
      For RENGI& = 1 To GRID.Rows
         TEPRUEBA.Caption = TRIM$(GRID.TextMatrix(RENGI& - 1, COLUI& - 1)) + "."
         If TEPRUEBA.Width > ACOLU Then ACOLU = TEPRUEBA.Width
      Next RENGI&
      GRID.ColWidth(COLUI& - 1) = ACOLU
   Next COLUI&
   '
End Sub

Private Sub GRID_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then ' BUSQUEDA  POR TEXTO CON F3
       If ARBUS$ <> "" Then
          VUELTA% = 0
          J& = GRID.Row

          H& = 0: If GRID.COL < GRID.Cols Then H& = GRID.COL + 1
10        For U& = J& To GRID.Rows - 1
             HI& = H&: If U& > J& Then HI& = 0
             For i& = HI& To GRID.Cols - 1
              If InStr(UCase$(GRID.TextMatrix(U&, i&)), ABUSQUE$) > 0 Then
                GRID.Row = U&: GRID.COL = i&
                On Error Resume Next
                 VISIROWS = (GRID.Height / GRID.RowHeight(J&)) - 4
                If Err > 0 Then
                 VISIROWS = 1
                End If
                On Error GoTo 0
                 If VISIROWS > 0 Then
                    If GRID.TopRow < GRID.Row - VISIROWS Then
                      GRID.TopRow = GRID.Row - VISIROWS
                    End If
                 End If
                GRID.SetFocus
                Exit Sub
              End If
              Next i&
          Next U&
          '
          Call MENSERR(24, "Argumento no Encontrado.")
       End If
   End If

End Sub

