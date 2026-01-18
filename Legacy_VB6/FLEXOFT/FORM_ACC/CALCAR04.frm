VERSION 5.00
Begin VB.Form CALCAR04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EFFFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Re-Cálculo de Carga de Máquinas"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11175
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00606000&
      Height          =   4530
      Left            =   10290
      ScaleHeight     =   4470
      ScaleWidth      =   1635
      TabIndex        =   41
      Top             =   0
      Width           =   1695
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
         Height          =   635
         Left            =   105
         Picture         =   "CALCAR04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Listados de Control"
         Top             =   1680
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   44
         Top             =   3570
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   45
            Top             =   0
            Width           =   12000
         End
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
         Height          =   635
         Left            =   105
         Picture         =   "CALCAR04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Inicia Proceso de Cálculo"
         Top             =   940
         Width           =   650
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
         Height          =   640
         Left            =   105
         Picture         =   "CALCAR04.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Cancelar Proceso de Cálculo"
         Top             =   300
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "CALCAR04.frx":0EDE
         Top             =   3885
         Width           =   1515
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00EFFFE0&
      Caption         =   "Avance de Proceso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4110
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   6000
      Begin VB.Frame Frame6 
         BackColor       =   &H00EFFFE0&
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   1600
         Left            =   350
         TabIndex        =   35
         Top             =   2380
         Width           =   1065
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nro.O/F"
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
            Index           =   0
            Left            =   0
            TabIndex        =   40
            Top             =   0
            Width           =   1000
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Detalle"
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
            Index           =   1
            Left            =   0
            TabIndex        =   39
            Top             =   315
            Width           =   1000
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Operación"
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
            Index           =   3
            Left            =   0
            TabIndex        =   38
            Top             =   630
            Width           =   1000
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Desde"
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
            Index           =   4
            Left            =   0
            TabIndex        =   37
            Top             =   945
            Width           =   1000
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta"
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
            Index           =   5
            Left            =   0
            TabIndex        =   36
            Top             =   1240
            Width           =   1000
         End
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00EFFFE0&
         Caption         =   "Calculando Fechas Proyectadas de Entrega"
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
         Left            =   525
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1745
         Width           =   5055
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
         Index           =   8
         Left            =   4305
         TabIndex        =   17
         Text            =   " "
         Top             =   3570
         Width           =   645
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
         Index           =   7
         Left            =   1470
         TabIndex        =   16
         Text            =   " "
         Top             =   3570
         Width           =   2220
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
         Index           =   6
         Left            =   4305
         TabIndex        =   15
         Text            =   " "
         Top             =   3255
         Width           =   645
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
         Index           =   5
         Left            =   1470
         TabIndex        =   14
         Text            =   " "
         Top             =   3255
         Width           =   2220
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
         Index           =   4
         Left            =   1995
         TabIndex        =   13
         Text            =   " "
         Top             =   2940
         Width           =   3690
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
         Left            =   1470
         TabIndex        =   12
         Text            =   " "
         Top             =   2940
         Width           =   435
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
         Left            =   4515
         TabIndex        =   11
         Text            =   " "
         Top             =   2310
         Width           =   1170
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
         Left            =   1470
         TabIndex        =   10
         Text            =   " "
         Top             =   2625
         Width           =   4215
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
         Index           =   0
         Left            =   1470
         TabIndex        =   9
         Text            =   " "
         Top             =   2310
         Width           =   1275
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00EFFFE0&
         Caption         =   "Verificando Tiempos Standard"
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
         Left            =   525
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1215
         Width           =   5160
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00EFFFE0&
         Caption         =   "Ordenamiento Ordenes de Fabricación por Prioridad"
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
         Left            =   525
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   950
         Width           =   5055
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00EFFFE0&
         Caption         =   "Selección de Operaciones  Pendientes"
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
         Left            =   525
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   685
         Width           =   4110
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00EFFFE0&
         Caption         =   "Selección de Ordenes de Fabricación Lanzadas"
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
         Left            =   525
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   420
         Width           =   4950
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00EFFFE0&
         Caption         =   "Secuenciando Carga de Maquinas"
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
         Left            =   525
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1480
         Width           =   5055
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "a las"
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
         Index           =   9
         Left            =   3465
         TabIndex        =   19
         Top             =   3630
         Width           =   750
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "a las"
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
         Index           =   8
         Left            =   3465
         TabIndex        =   18
         Top             =   3300
         Width           =   750
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "horas"
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
         Index           =   7
         Left            =   4725
         TabIndex        =   8
         Top             =   3630
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "horas"
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
         Index           =   6
         Left            =   4830
         TabIndex        =   7
         Top             =   3300
         Width           =   750
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
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
         Index           =   2
         Left            =   3020
         TabIndex        =   6
         Top             =   2385
         Width           =   1380
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00EFFFE0&
      Caption         =   "Parámetros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4110
      Left            =   6300
      TabIndex        =   22
      Top             =   210
      Width           =   3795
      Begin VB.Frame Frame4 
         BackColor       =   &H00EFFFE0&
         Caption         =   "O/F's Suspendidas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   210
         TabIndex        =   33
         Top             =   2940
         Width           =   3375
         Begin VB.CheckBox Check7 
            BackColor       =   &H00EFFFE0&
            Caption         =   "Incluir O/F's Suspendidas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   525
            TabIndex        =   34
            Top             =   315
            Width           =   2640
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EFFFE0&
         Caption         =   "Ordenamiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   210
         TabIndex        =   29
         Top             =   1680
         Width           =   3390
         Begin VB.OptionButton Option1 
            BackColor       =   &H00EFFFE0&
            Caption         =   "Por Prioridades"
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
            Left            =   525
            TabIndex        =   32
            Top             =   480
            Value           =   -1  'True
            Width           =   1905
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00EFFFE0&
            Caption         =   "Por Secuencia"
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
            Left            =   525
            TabIndex        =   31
            Top             =   735
            Width           =   1905
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
            Height          =   540
            Left            =   2625
            MaskColor       =   &H8000000F&
            Picture         =   "CALCAR04.frx":2E00
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Re-Ordenar Secuenciaa"
            Top             =   380
            Width           =   540
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00EFFFE0&
         Caption         =   "Inicio de Carga"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   210
         TabIndex        =   23
         Top             =   315
         Width           =   3390
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
            Height          =   285
            Index           =   9
            Left            =   630
            TabIndex        =   26
            Top             =   735
            Width           =   1065
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
            Height          =   285
            Index           =   10
            Left            =   2415
            TabIndex        =   25
            Top             =   735
            Width           =   750
         End
         Begin VB.CommandButton Command3 
            Caption         =   "."
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
            Left            =   1680
            TabIndex        =   24
            Top             =   735
            Width           =   175
         End
         Begin VB.Label Label2 
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
            Index           =   10
            Left            =   -1260
            TabIndex        =   28
            Top             =   480
            Width           =   2220
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hora:"
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
            Index           =   11
            Left            =   1050
            TabIndex        =   27
            Top             =   480
            Width           =   1590
         End
      End
   End
   Begin VB.ListBox LIORFALA 
      Height          =   450
      Left            =   5985
      Sorted          =   -1  'True
      TabIndex        =   20
      Top             =   3675
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "CALCAR04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim HAYFAB%(32767), NIVEFAB%(32767)
Dim FEFINANT%(10), HOFINANT%(10)
'
Dim ORFA$(8192)
Dim ORFB$(8192)
Dim OPRIN$(8192)
Dim ORFASEQ$(4096), STATU%(4096), PRIVI%(4096)
Dim REGOF&(1024)
Dim PRIO%(1024)
Dim FENEC%(1024)
Dim DNEC%(1024)
Dim HNEC%(1024)
Dim MQNEC$(1024)
Dim MQBAS$(1024)
Dim RGOPX&(1024)
Dim REGOPE&(256)
'
Dim HODE%(64), HOHA%(64), OPEDIS(64)
Dim EQ$(10)
'
Dim NUMAQ%, VECHORES$(512)
Dim MAQUI$(), MAEQUI$(), COEFTIE(), TIEPRO()
'
Dim VECTOMAQ As String
Dim VECTOSEQ As String
'
Dim MQQ%(8192, 3)
Dim MAQUISE As String
'
Dim VECMODRES$(), CAVECMOD%
'
Dim CAOR As Integer
Dim CADIA As Integer
Dim TXT As String * 255

Dim IFORCA%
'
' TODO ESTA DIMENSIONADO PARA 64 SECCIONES PRODUCTIVAS
Dim COMTUR%(8), FINTUR%(8), CATURNOS%
Dim MODISEC(), CASECCI%
Dim MODISPO(7, 8, 64)
Dim VECMODIR$(64)
Dim DIACTU%

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    Call CALCUMAQ
    Command2.Enabled = True
    Hide
End Sub

Private Sub Command3_Click()
    Call SELECHO("SELFECHA")
    FECHANUE$ = TRIM$(VALACT1("SELFECHA"))
    If FECHANUE$ <> "" Then
        Text1(9).TEXT = FECHANUE$
    End If
End Sub

Private Sub Command4_Click()
    Option2.Value = True
    Call SECUORF
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   TILISTA% = ALTERNA%("Secuencia O/F's por Prioridad\Secuencia Operaciones por Linea")
   Select Case TILISTA%
      Case 1: Call FINAL("*IORFPRI")
      Case 2: Call FINAL("*OPERLIN")
   End Select
   Command5.Enabled = True
End Sub

Private Sub Form_Load()
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    HOX% = HOY(HO$)
    Text1(9).TEXT = HO$
    Text1(10).TEXT = Left$(Time$, 5)
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
End Sub

Private Sub Text1_Change(Index As Integer)
    Static HOX%
    If HOX% = 0 Then
        HOX% = HOY(HO$)
    End If
    If Index = 9 Then
      If FECHANUM(Text1(9).TEXT) > HOX% Then
        Text1(10).TEXT = "0:00"
      End If
    End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index = 9 Then
            Text1(10).SetFocus
          ElseIf Index = 10 Then
            Command2.SetFocus
        End If
    End If
End Sub

Sub CARCONFICAL()
   '
   DIACTU% = HOY(HOS$)
   '
   ' TURNOS DIARIOS DE PRODUCCION
   DESDA% = 0: HASDA% = 0: CATURNOS% = 0
   For U& = 1 To ULTREG&("TURDIAR")
     XX$ = REGLEIDO$("TURDIAR", U&)
     DESDE% = CVI(Mid$(XX$, 25, 2))
     HASTA% = CVI(Mid$(XX$, 27, 2))
     If DESDE% < HASDA% Then ' Or HASTA% <= DESDE% Then
       Call MENSERR(24, "Error de Configuración\de Turnos Diarios de Producción")
       Call FINAL("")
     End If
     CATURNOS% = CATURNOS% + 1
     If CATURNOS% > 8 Then GoTo 25
     Call REPLA(XX$, Chr$(CATURNOS%), 1, 1)
     Call GRAREG("TURDIAR", XX$, U&)
     '
     COMTUR%(CATURNOS%) = DESDE%: FINTUR%(CATURNOS%) = HASTA%
     DESDA% = DESDE%: HASDA% = HASTA%
   Next U&
   '
25 If CATURNOS% < 1 Or CATURNOS% > 8 Then
     Call MENSERR(24, "Falta Configurar\Turnos Diarios de Producción")
     Call FINAL("")
   End If
   '
   CASECCI% = ULTREG&("SECPROD")
   If CASECCI% > 64 Then
     Call MENSERR(24, "Proceso Imposible\Demasiadas Secciones Productivas. Consulte.")
     Call FINAL("")
   End If
   ReDim Preserve MODISEC(CASECCI%, 0 To 0)
   Dim TIEMODIS%(64)
   '
   'MANO DE OBRA DISPONIBLE POR DIA DE SEMANA
   MODISPO(0, 0, 0) = 0 ' todos los dias, todos los turnos, TODAS LAS SECCIONES
   For DISEMA% = 1 To 7
     For ITURNO% = 1 To CATURNOS%
       For ISECCI% = 1 To CASECCI%
         MODISPO(DISEMA%, ITURNO%, ISECCI%) = 0
       Next ISECCI%
     Next ITURNO
   Next DISEMA%
   '
   VECTOSEQ$ = ""
   For U& = 1 To ULTREG&("SECPROD")
     VECTOSEQ$ = VECTOSEQ$ + Left$(REGLEIDO$("SECPROD", U&), 4)
   Next U&
   For U& = 1 To ULTREG&("DOTASEG")
     XX$ = REGLEIDO$("DOTASEG", U&)
     For KI1% = CATURNOS% + 1 To 8
       Call REPLA(XX$, MKI$(0), 5 + 2 * KI1%, 2)
     Next KI1%
     Call GRAREG("DOTASEG", XX$, U&)
     SECCI = CVS(Left$(XX$, 4))
     ISECCI% = 0
     For KI% = 1 To Len(VECTOSEQ$) Step 4
       If CVS(Mid$(VECTOSEQ$, KI%, 4)) = SECCI Then
         ISECCI% = (KI% + 3) / 4
         Exit For
       End If
     Next KI%
     DISEMA% = Asc(Mid$(XX$, 5, 1))
     If ISECCI% > 0 Then
       If ISECCI% <= 64 Then
         If DISEMA% > 0 Then
           If ISEMA% <= 7 Then
             For ITURNO% = 1 To CATURNOS%
               MODISPO(DISEMA%, ITURNO%, ISECCI%) = CVI(Mid$(XX$, 5 + 2 * ITURNO%, 2))
               If MODISPO(DISEMA%, ITURNO%, ISECCI%) > 0.05 Then TIEMODIS%(ISECCI%) = 1
             Next ITURNO%
           End If
         End If
       End If
     End If
   Next U&
   '
   FALTAMOD% = 0
   For U& = 1 To CASECCI%
     If TIEMODIS%(U&) < 1 Then
       DENOSE$ = TRIM$(Mid$(REGLEIDO$("SECPROD", U&), 5, 32))
       Call MENSERR(24, "No Hay M.O.D. Disponible\para la Sección " + DENOSE$)
       FALTAMOD% = 1
     End If
   Next U&
   '
   If FALTAMOD% > 0 Then Call FINAL("")
   '
End Sub
'
Sub CALCUMAQ()   ' recalculo de carga de maquinas segun
                    ' ordenes de fabricacion en curso
                 
' VECTOMAQ$ es un string que contiene las maquinas definidas Y LAS SECCIONES en bloques de 10 bytes

    Call CARCONFICAL   ' CARGA CONFIGURACION DE CALCULO
    '
'0.- Dimensionamiento de vectores
   Dim ORFA$(), PRIO%(), CIOF%(), NVFB%(), PUNTOPER$()
   '
'1.- Carga vectores de maquinas y equipos
    NUMAQ% = ULTREG&("PADMAQ")
    ReDim MAQUI$(NUMAQ%), MAEQUI$(NUMAQ%), COEFTIE(NUMAQ%), TIEPRO(NUMAQ%)
    JJ& = 0: VECTOMAQ$ = "": VECTOSEQ$ = ""
    For U& = 1 To ULTREG&("SECPROD")
      VECTOSEQ$ = VECTOSEQ$ + Left$(REGLEIDO$("SECPROD", U&), 4)
    Next U&
    For U& = 1 To NUMAQ%
      REGMA$ = REGLEIDO$("PADMAQ", U&)
        If TRIM$(Left$(REGMA$, 6)) <> "" Then
        SECCIX = CVS(Mid$(REGMA$, 37, 4))
        If SECCIX < 1 Or ECOARCH$("SECPROD", MKS$(SECCIX)) = "" Then
          Call MENSERR(24, "Falta Definir Sección Productiva\para Equipo '" + TRIM$(Left$(REGMA$, 6)) + " - " + TRIM$(Mid$(REGMA$, 7, 30)) + "'")
          Exit Sub
        End If
        '
        For ISEQ% = 1 To Len(VECTOSEQ$) / 4
          If CVS(Mid$(VECTOSEQ$, 4 * ISEQ% - 3, 4)) = SECCIX Then GoTo 15
        Next ISEQ%
        ISEQ% = 0
        '
15      JJ& = JJ& + 1
        MAQUI$(JJ&) = Left$(REGMA$, 6)
        MAEQUI$(JJ&) = Mid$(REGMA$, 43, 6)
        COEFTIE(JJ&) = CVI(Mid$(REGMA$, 49, 2))
          If COEFTIE(JJ&) < 1 Then COEFI = 100
        TIEPRO(JJ&) = 0
        VECTOMAQ$ = VECTOMAQ$ + "@@" + Left$(REGMA$, 6) + MKI$(ISEQ%)
      End If
    Next U&
    NUMAQ% = JJ&
    '
    Dim COLITRA$(32), MALITRA$(32), CAEQLIN%(32), CAEQSEL%(32)
    If ULTREG&("LINTRAN") > 32 Then
      Call MENSERR(24, "Proceso Imposible.\  \Capacidad de Procesamiento\de 32 Lineas Transferenciales Excedida.\  \Consulte a su Soporte de Sistemas.")
    End If
    CALITRA% = 0
    For U& = 1 To ULTREG&("LINTRAN")
      XX$ = REGLEIDO$("LINTRAN", U&)
      COLIT$ = TRIM$(Left$(XX$, 6))
      If COLIT$ <> "" Then
        If ECOARCH$("PADMAQ", COLIT$) <> "" Then
          SQN0% = SEQNU%(COLIT$)
          MALIT$ = ""
          For KI% = 7 To 54 Step 6
            MAIN$ = TRIM$(Mid$(XX$, KI%, 6))
            If MAIN$ <> "" Then
              If ECOARCH$("PADMAQ", MAIN$) <> "" Then
                If SEQNU%(MAIN$) <> SQN0% Then
                  Call MENSERR(24, "Proceso Imposible.\Equipo '" + TRIM$(MAIN$) + "' en Distinta Sección\de Línea Transferencial '" + TRIM$(COLIT$) + "'.")
                  Exit Sub
                End If
                MALIT$ = MALIT$ + AJUSTI$(MAIN$, 6)
              End If
            End If
          Next KI%
          If MALIT$ <> "" Then
            CALITRA% = CALITRA% + 1
            COLITRA$(CALITRA%) = AJUSTI$(COLIT$, 6)
            MALITRA$(CALITRA%) = MALIT$
            CAEQLIN%(CALITRA%) = Len(MALIT$) / 6
            CAEQSEL%(CALITRA%) = CVI(Mid$(XX$, 121, 2))
              If CAEQSEL%(CALITRA%) < 1 Then CAEQSEL%(CALITRA%) = 1
          End If
        End If
      End If
    Next U&
    '
    CAORFA% = 8192
    CAPRO% = 128
    CAOPSELMAX% = 256
    '
'2.- Verifica y valida fecha y hora de comienzo de calculo
    FCAL% = HOY(HO$)
    FECOMPRO% = FECHANUM(Text1(9))
    HOCOMPRO% = HORANUM(Text1(10))
    If FECOMPRO% < FCAL% Then
      Call MENSERR(24, "Fecha no válida")
      Command2.Enabled = True
      Text1(9) = HO$
      Text1(9).SetFocus
      Exit Sub
    End If
    '
'2.- Verifica y valida modo de calculo en MODOCA%
    If Option1.Value = True Then
         MODOCA% = 1   ' por prioridades
      ElseIf Option2.Value = True Then
         MODOCA% = 2   ' por secuencia
      Else
        Call MENSERR(24, "Falta Seleccionar Modo de Ordenamiento")
        Command2.Enabled = True
        Option2.SetFocus
        Exit Sub
    End If
    '
'3.- Determina si incluye o no ordenes suspendidas INCLUSUS%
    INCLUSUS% = 0
    If Check7.Value = 1 Then
      INCLUSUS% = 1
    End If
    '
'5.- Carga VECHORES$ con horarios reservados por paradas programadas
    For U& = 1 To ULTREG("PARAPROG")
      XXX$ = REGLEIDO$("PARAPROG", U&)
      MQX% = MAQNU%(Left$(XXX$, 6))
      FECOM% = CVI(Mid$(XXX$, 7, 2))
      HOCOM% = CVI(Mid$(XXX$, 9, 2))
      FEFIN% = CVI(Mid$(XXX$, 11, 2))
      HOFIN% = CVI(Mid$(XXX$, 13, 2))
      STH$ = MKI$(FECOM%) + MKI$(HOCOM%) + MKI$(FEFIN%) + MKI$(HOFIN%) + MKI$(FECOM%) + MKI$(HOCOM%)
      VECHORES$(MQX%) = VECHORES$(MQX%) + STH$
    Next U&
    '
'6.- Seleccionando Ordenes Lanzadas y Prioridades Correspondientes
    '
    Check1.Value = 2
    CAOR% = 0: VECOR% = 0
    
    SECUORFA$ = "": CAORFSEQ% = 0
    If MODOCA% = 2 Then
      FIN& = ULTREG&("SECUORF")
      For I& = 1 To FIN&
        X$ = REGLEIDO$("SECUORF", I&)
        NORFAX$ = AJUSTI$(Left$(X$, 12), 12)
        CAOR% = CAOR% + 1
        If CAOR% > VECOR% Then
          VECOR% = VECOR% + 64
          ReDim Preserve ORFA$(VECOR%), PRIO%(VECOR%)
        End If
        ORFA$(CAOR%) = NORFAX$
        PRIO%(CAOR%) = CAOR%
        SECUORFA$ = SECUORFA$ + "@@" + AJUSTI$(NORFAX$, 14)
      Next I&
      CAORFSEQ% = CAOR%
    End If
    '
    Call COPYSTRU("ECORDEF", "ECORDEP")
    LIORFALA.Clear
    STATUMI% = 1: STATUMA% = 1
    If INCLUSUS% = 1 Then STATUMA% = 2
    '
    Call ABREORFAB
    SQLX$ = " SELECT * FROM ORDEFABR "
    SQLX$ = SQLX$ + " WHERE StatuOrf >= " & STATUMI% & " "
    SQLX$ = SQLX$ + " AND StatuOrf <= " & STATUMA% & " "
    SQLX$ = SQLX$ + " ORDER BY IdSubOr "
    '
    Set ORDEFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
    With ORDEFABTMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        JJ& = 0
        Do While (.EOF = False)
          STATUOF% = !Statuorf
          If STATUOF% >= STATUMI% And STATUOF% <= STATUMA% Then ' esta lanzada
            PRIORI% = !Priorid: If PRIORI% < 1 Or PRIORI% > 9 Then PRIORI% = 10
            PRIORI% = CAORFSEQ% + PRIORI% + 10000 * (STATUOF% - 1)
            ORFASE$ = "@@" + AJUSTI$(!IdSubOr, 14)
            PPXX% = InStr(SECUORFA$, ORFASE$)
            If PPXX% < 1 Then
              CAOR% = CAOR% + 1
              If CAOR% > VECOR% Then
                VECOR% = VECOR% + 64
                ReDim Preserve ORFA$(VECOR%), PRIO%(VECOR%)
              End If
              ORFA$(CAOR%) = !IdSubOr
              PRIO%(CAOR%) = PRIORI%
              SECUORFA$ = SECUORFA$ + "@@" + AJUSTI$(NORFAX$, 14)
            End If
          End If
        .MoveNext
        Loop
      End If
    End With
    '
'7.- Agregando a Cada Orden el Codigo de Producto Correspondiente
    '
    ReDim CIOF%(CAOR%), NVFB%(CAOR%), PUNTOPER$(CAOR%)
    JJ& = 0
    For IOR% = 1 To CAOR%
      SQLX$ = " SELECT * FROM ORDEFABR "
      SQLX$ = SQLX$ + " WHERE IdSubOr = '" & ORFA$(IOR%) & "' "
      '
      Set ORDEFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
      With ORDEFABTMP
        On Error Resume Next
        .MoveLast
        If Err < 1 Then
          On Error GoTo 0
          CIOF%(IOR%) = !Cod_Inte
          NVFB%(IOR%) = CANIVEXPLO%(CIOF%(IOR%))
        End If
      End With
      '
      Y$ = REGBLAN("ECORDEP")
      Call REPLA(Y$, ORFA$(IOR%), 1, 12)
      Call REPLA(Y$, ECOARCH$("ECORDEF", ORFA$(IOR%)), 13, 52)
      JJ& = JJ& + 1
      Call GRAREG("ECORDEP", Y$, JJ&)
      '
    Next IOR%
    Call SETULTREG("ECORDEP", JJ&)
    '
'8.- Ordenando O/F's por Prioridad
    '
    Call ABREORFAB
    FIN& = CAOR%
    For II% = 1 To CAOR%
      ORFX$ = TRIM$(Left$(ORFA$(II%), 12))
      SQLX$ = " SELECT * FROM OPERFAB "
      SQLX$ = SQLX$ + " WHERE IdSubOr = '" & ORFX$ & "' "
      SQLX$ = SQLX$ + " AND StatOper <= " & STATUMA% & " "
      SQLX$ = SQLX$ + " AND CantReal < CantPend "
      '
      Set OPERFAPTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
      With OPERFAPTMP
        On Error Resume Next
        .MoveFirst
        If Err < 1 Then
          Do While (.EOF = False)
            CAREA# = !CantReal
            If CAREA# > 0 Then     ' Las O/F's ya comenzadas tienen prioridad
              If PRIO%(II%) > 0 Then
                PRIO%(II%) = PRIO%(II%) - 32767
              End If
            End If
          .MoveNext
          Loop
        End If
        On Error GoTo 0
      End With
    Next II%
    '
    JJ& = 0
    For II% = 1 To CAOR%
      ZZ$ = REGBLAN$("IORFPRI")
      Call REPLA(ZZ$, MKI$(PRIO%(II%)), 1, 2)
      Call REPLA(ZZ$, MKS$(Val(Mid$(ORFA$(II%), 4, 6))), 3, 4)
      Call REPLA(ZZ$, MKI$(NVFB%(II%)), 7, 2)
      Call REPLA(ZZ$, ORFA$(II%), 9, 12)
      Call REPLA(ZZ$, MKI$(CIOF%(II%)), 21, 2)
      JJ& = JJ& + 1
      Call GRAREG("IORFPRI", ZZ$, JJ&)
    Next II%
    Call SETULTREG("IORFPRI", JJ&)
    Call CIERRARCH("IORFPRI")
    Call FILESORT("IORFPRI", "", 1, 3, "ASC")
    '
'9.- Re-Carga del Vector de O/F's
    II% = 0
    For JJ& = 1 To ULTREG("IORFPRI")
      ZZ$ = REGLEIDO$("IORFPRI", JJ&)
      II% = JJ&
      PRIO%(II%) = CVI(Mid$(ZZ$, 1, 2))
      NVFB%(II%) = CVI(Mid$(ZZ$, 7, 2))
      ORFA$(II%) = Mid$(ZZ$, 9, 12)
      CIOF%(II%) = CVI(Mid$(ZZ$, 21, 2))
    Next JJ&
    '
    For II% = 1 To CAOR%
      NUMOF1& = Val(Mid$(ORFA$(II%), 4, 6))
      For JJI% = II% + 1 To CAOR%
        If NUMOF1& = Val(Mid$(ORFA$(JJI%), 4, 6)) Then
          If Val(Mid$(ORFA$(II%), 13, 2)) > Val(Mid$(ORFA$(JJI%), 13, 2)) Then
            Call MENSERR(24, "Proceso Imposible.\Prioridades No Compatibles\con Niveles de Fabricación\en O/F " + TRIM$(ORFA$(II%)))
            Exit Sub
          End If
        End If
      Next JJI%
    Next II%
    '
'10.- Lectura de operaciones pendientes cargando archivo OPERFAP
    Check2.Value = 2
    '
    Call ABREORFAB
    Screen.MousePointer = 11
    REBLA$ = REGBLAN$("OPERFAT")
    JJ& = 0
    '
    For II% = 1 To CAOR%
      NOFA$ = TRIM$(ORFA$(II%))
      SQLX$ = " SELECT * FROM OPERFAB "
      SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NOFA$ & "' "
      SQLX$ = SQLX$ + " AND StatOper <= " & STATUMA% & " "
      SQLX$ = SQLX$ + " AND CantReal < CantProg "
      SQLX$ = SQLX$ + " ORDER BY NumeOper ASC "
      '
      Set OPERFAPTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
      With OPERFAPTMP
        On Error Resume Next
        .MoveFirst
        If Err < 1 Then
          Do While Not .EOF
            CAPROG# = !CantProg
            CAREA# = !CantReal
            CAPEN# = CAPROG# - CAREA#
            CIXII% = !Cod_Inte
            NOPEII% = !NumeOper
            '
            If CAPEN# >= 0.05 Then
              If CAPROG# >= 0.05 Then
'                COEF = CAPEN# / CAPROG#
'                    If COEF > 1 Then COEF = 1
'                TPREP = !TiemPrep: If CAREA# >= 0.05 Then TPREP = 0
'                TFIJO = !TiemFijo
'                TPROD = TIEMAQUI(CIXII%, NOPEII%, CAPEN#, 1) ' MODO 1 PARA NO INCLUIR PREPARACION
'                TVARI = TPROD - TFIJO: If TVARI < 0 Then TVARI = 0
'                TDEMO = 0: TMANT = 0
'                  If (!TiemFijo + !TiemVari) > 0.1 Then
'                    TDEMO = TVARI * !TiemDemo / (!TiemFijo + !TiemVari)
'                  End If
'                TMANT = COEF * CVS(Mid$(O$, 119, 4))
'                TLIMP = CVS(Mid$(O$, 123, 4))
'                TOTRO = COEF * CVS(Mid$(O$, 127, 4))
                TPREP = !TiemPrep: If CAREA# >= 0.05 Then TPREP = 0
                TFIJO = !TiemFijo
                TVARI = !TiemVari * CAPEN# / CAPROG# 'TPROD - TFIJO: If TVARI < 0 Then TVARI = 0
                TDEMO = !TiemDemo * CAPEN# / CAPROG#
                TMANT = !TiemMHer * CAPEN# / CAPROG#
                TLIMP = !TiemLimp
                TOTRO = !TiemOtro * CAPEN# / CAPROG#
                TTOTA = TPREP + TFIJO + TVARI + TDEMO + TMANT + TLIMP + TOTRO
                '
                O$ = REBLA$
                Call REPLA(O$, !IdSubOr, 1, 12)
                Call REPLA(O$, MKI$(!Cod_Inte), 13, 2)
                Call REPLA(O$, MKI$(!NumeOper), 15, 2)
                  DOPEXI$ = TRIM$(Str$(!NumeOper)) + "." + !DescOper
                Call REPLA(O$, DOPEXI$, 17, 48)
                'Call REPLA(O$, !Destino, 65, 8)  ' FALTA COMPLETAR ESTE ELEMENTO
                Call REPLA(O$, !Maquina, 73, 6)
                Call REPLA(O$, MKD$(!CantProg), 79, 8)
                Call REPLA(O$, MKD$(!CantReal), 87, 8)
                Call REPLA(O$, MKD$(!CantPend), 95, 8)
                Call REPLA(O$, MKS$(TPREP), 103, 4)
                Call REPLA(O$, MKS$(TFIJO), 107, 4)
                Call REPLA(O$, MKS$(TVARI), 111, 4)
                Call REPLA(O$, MKS$(TDEMO), 115, 4)
                Call REPLA(O$, MKS$(TMANT), 119, 4)
                Call REPLA(O$, MKS$(TLIMP), 123, 4)
                Call REPLA(O$, MKS$(TOTRO), 127, 4)
                Call REPLA(O$, MKS$(TTOTA), 131, 4)
                  HORACO$ = Format(!ComiePro, "hh:mm:ss")
                Call REPLA(O$, MKI$(0), 135, 2)
                Call REPLA(O$, MKI$(0), 137, 2)
                Call REPLA(O$, String$(8, 0), 139, 8)
                Call REPLA(O$, MKS$(!CantiOps), 147, 4)
                Call REPLA(O$, Space$(48), 151, 48)
                Call REPLA(O$, !EquiAux, 203, 18)
                Call REPLA(O$, MKS$(I&), 221, 4)
                '
                JJ& = JJ& + 1
                Call GRAREG("OPERFAT", O$, JJ&)
                PUNTOPER$(II%) = PUNTOPER$(II%) + MKS$(JJ&)
              End If
            End If
            '
          .MoveNext
          Loop
          '
        End If
        On Error GoTo 0
      End With
      '
    Next II%
    Call SETULTREG("OPERFAT", JJ&)
    Check2.Value = 1
    '
    '
'11.- Verificando Fecha de Comienzo Proyectada
    '
 ' estan todas en cero
 ' ajustar de acuerdo con fecha de lanzamiento de ordefabr
 ' luego con fecha y hora de la ultima operacion informada
 ' los que todavia quedan en cero con la fecha y hora de programacion
400 Check4.Value = 2
    '
    FIN& = ULTREG&("OPERFAT")
    FIN1& = FIN&: If FIN1& < 1 Then FIN1& = 1
    '
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      O$ = REGLEIDO$("OPERFAT", I&)
      '
      If CVI(Mid$(O$, 135, 2)) = 0 Then
        NORFAX$ = Left$(O$, 12)
        SQLX$ = " SELECT * FROM ORDEFABR "
        SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFAX$ & "' "
        '
        Set ORDEFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
        With ORDEFABTMP
          On Error Resume Next
          .MoveLast
          If Err < 1 Then
            On Error GoTo 0
            FEPROLA% = CVINT%(!FechLan)
            If FEPROLA% > FECOMPRO% Then
              Call REPLA(O$, MKI$(FEPROLA%), 135, 2)
              Call REPLA(O$, MKI$(0), 137, 2)
              Call GRAREG("OPERFAT", O$, I&)
            End If
          End If
          On Error GoTo 0
        End With
      End If
      '
    Next I&
    '
    'AQUI AJUSTAR CON LA FECHA Y HORA DE LA ULTIMA OPERACION INFORMADA
    '
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      O$ = REGLEIDO$("OPERFAT", I&)
      '
      If CVI(Mid$(O$, 135, 2)) = 0 Then
        Call REPLA(O$, MKI$(FECOMPRO%), 135, 2)
        Call REPLA(O$, MKI$(HOCOMPRO%), 137, 2)
        Call GRAREG("OPERFAT", O$, I&)
      End If
      '
    Next I&
    Call CIERRARCH("OPERFAT")
    Check4.Value = 1
    '
500 Call BLANARCH("CRONOPE")
    Call BLANARCH("GANTMAQ")
    Call BLANARCH("PERIPRO")
    '
    '
'11.- Secuenciando Carga de Maquinas
    '
   Check5.Value = 2
'LO QUE SIGUE ESTA DESHABILITADO PORQUE NO HAY REEMPLAZO POR MAQUINAS EQUIVALENTES
'*********************************************************************************
'    FIN& = ULTREG&("OPERFAP")
'    FIN1& = FIN&: If FIN1& < 1 Then FIN1& = 1
'    '
'    NUOTA = 0
'    CAOPERF& = ULTREG&("OPERFAP")            ' Cantidad de Operaciones Pendientes
'    '
'    For I& = 1 To FIN&
'       '
'       Call TRACE(20, I&, FIN&)
'       O$ = REGLEIDO$("OPERFAP", I&)
'       NUOTI = Val(Mid$(O$, 4, 9))            ' numero de operacion
'       If NUOTI <> NUOTA Then MAQUISE$ = ""
''LO ANTERIOR NO SE POR QUE ES
'       CAREA# = CVD(Mid$(O$, 87, 8))
'       TTOT = CVS(Mid$(O$, 131, 4))
'       MAQUIS$ = Mid$(O$, 73, 6) + Mid$(O$, 203, 18)
'       '
'' OJO - ACA SE TOMAN LAS MAQUINAS POR GRUPO
'       For IMQ% = 0 To 3
'          MQQ%(I&, IMQ%) = MAQNU%(Mid$(MAQUIS$, 6 * IMQ% + 1, 6))
'       Next IMQ%
'       '
'       If CAREA# < 0.1 Then
'          Call REPLAMAQ(MAQUIS$, TTOT, MAQUIAL$, TALT)
'          '
'          COEFTIPRO = 1: If TTOT > 0 Then COEFTIPRO = TALT / TTOT
'          STTIE$ = ""
'          For IMQ% = 103 To 134 Step 4
'            STTIE$ = STTIE$ + MKS$(COEFTIPRO * CVS(Mid$(O$, IMQ%, 4)))
'          Next IMQ%
'          Call REPLA(O$, Left$(MAQUIAL$, 6), 73, 6)
'          Call REPLA(O$, Mid$(MAQUIAL$, 7, 18), 203, 18)
'          Call REPLA(O$, STTIE$, 103, 32)
'          Call GRAREG("OPERFAP", O$, I&)
'       End If
'       '
'       For IMQ% = 0 To 3
'          MAQUIBA$ = Mid$(MAQUIAL$, 6 * IMQ% + 1, 6)
'          If TRIM$(MAQUIBA$) <> "" Then
'            For JJI% = 1 To NUMAQ%
'              If MAQUI$(JJI%) = MAQUIBA$ Then
'                TIEPRO(JJI%) = TIEPRO(JJI%) + TALT
'              End If
'            Next JJI%
'          End If
'       Next IMQ%
'       '
'       NUOTA = NUOTI
'       '
'503 Next I&
'    '
'    Call CIERRARCH("OPERFAP")
    Check5.Value = 1
    '
    FIN& = ULTREG&("OPERFAT")
    CAOPERF& = ULTREG&("OPERFAT")            ' Cantidad de Operaciones Pendientes
    '
    '\\\OT-08-0971-OD-24/10/08///
    Call SETULTREG&("CRPMENSA", 0)
    '\\\OT-08-0971-OD-24/10/08///
    For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       O$ = REGLEIDO$("OPERFAT", I&)
       NOFA$ = Left$(O$, 12)
       NUMEOP% = CVI(Mid$(O$, 15, 2))
       CAREA# = CVD(Mid$(O$, 87, 8))
       '
       ' LO QUE SIGUE ES PARA PONER LA FECHA MAS TEMPRANA QUE PUEDE EMPEZAR LA OPERACION
'ABC$ = NOFA$
'ABC$ = FECHATEX$(CVI(Mid$(O$, 135, 2)))
'DEF = CVI(Mid$(O$, 137, 2))
       If CVI(Mid$(O$, 135, 2)) < 1 Then
         Call REPLA(O$, MKI$(FECOMPRO%), 135, 2)
         Call REPLA(O$, MKI$(HOCOMPRO%), 137, 2)
         Call GRAREG("OPERFAT", O$, I&)
       End If
       '
       Text1(0).TEXT = NOFA$
       CIII% = CVI(Mid$(O$, 13, 2))
       CAPROGG = CVD(Mid$(O$, 79, 8))
       Text1(1).TEXT = TRIM$(CODEXT$(CIII%)) + " - " + TRIM$(DESCRIT$(CIII%)) + " (" + TRIM$(Str$(CAPROGG)) + ")"
       '
       FEFINA% = CVI(Mid$(O$, 135, 2))
       HOFINA% = CVI(Mid$(O$, 137, 2))
       FEX = FEFINA%
       Text1(2).TEXT = FECHATEX$(FEX)
       Text1(3).TEXT = TRIM$(CVI(Mid$(O$, 15, 2)))
       Text1(4).TEXT = TRIM$(Mid$(O$, 17, 48))
       '
       FECOM% = FEFINA%
       HOCOM% = HOFINA%
       '
       FEFINABAS% = FEFINA%
       HOFINABAS% = HOFINA%
       '
       TIENEC = CVS(Mid$(O$, 131, 4)) / 60
       TIENEPREP = CVS(Mid$(O$, 103, 4)) / 60
       EQUIPROP$ = AJUSTI$(Mid$(O$, 73, 6), 6)
       CANOPNEC = CVS(Mid$(O$, 147, 4))
       FECOMPROP% = FECOM%
       HOCOMPROP% = HOCOM%
       '
       If TRIM$(EQUIPROP$) = "" Or ECOARCH$("PADMAQ", EQUIPROP) = "" Then
         '\\\OT-08-0971-OD-24/10/08///
         'Call MENSERR(24, "Equipo no Válido en Operación " + TRIM$(Text1(4)) + "\para Orden de Fabricación " + TRIM$(NOFA$) + ".")
         MENSAJE$ = "Equipo no Válido en Operación " + TRIM$(Text1(4))
         Call REPLA(TX$, NOFA$, 1, 12)
         Call REPLA(TX$, MKI$(CIII%), 13, 2)
         Call REPLA(TX$, MENSAJE$, 15, 48)
         Call REGAPP("CRPMENSA", TX$)
         GoTo 139
         '\\\OT-08-0978-OD-FIN///

       End If
       '
       For KI1% = 1 To CALITRA%
         If COLITRA$(KI1%) = EQUIPROP$ Then
           ' ES UNA LINEA TRANSFERENCIAL
           MALIT$ = MALITRA$(KI1%)
           QLIN% = CAEQLIN%(KI1%)
           QSEL% = CAEQSEL%(KI1%)
           MEJFINPOS# = 0
           GoTo 1200
         End If
       Next KI1%
       '
       ' es un equipo aislado
       QLIN% = 1
       QSEL% = 1
       MALIT$ = AJUSTI$(EQUIPROP$, 6)
       MEJFINPOS# = 0
       '
1200   For KI2% = 1 To QLIN% - QSEL% + 1
         FECOMPROP% = FECOM%
         HOCOMPROP% = HOCOM%
1230     COMIENPROPU# = 10000 * CDbl(FECOMPROP%) + CDbl(HOCOMPROP%)
         '
         For KI3% = KI2% To KI2% + QSEL% - 1
           EQPX$ = Mid$(MALIT$, 6 * KI3% - 5, 6)
           'aqui EQPX$ es un equipo individual
           Call MEJORPRO(TIENEC, EQPX$, CANOPNEC, FECOMPROP%, HOCOMPROP%, FECOMREAL%, HOCOMREAL%, TURNOINI%, FEFINREAL%, HOFINREAL%, TURNOFIN%)
           COMIENPOSIB# = 10000 * CDbl(FECOMREAL%) + CDbl(HOCOMREAL%)
           If COMIENPOSIB# > COMIENPROPU# Then
             FECOMPROP% = FECOMREAL%
             HOCOMPROP% = HOCOMREAL%
             GoTo 1230
           End If
         Next KI3%
         FINPOSIBLE# = 10000 * CDbl(FEFINREAL%) + CDbl(HOFINREAL%)
         If MEJFINPOS# = 0 Or FINPOSIBLE# < MEJFINPOS# Then
           MEJFINPOS# = FINPOSIBLE#
           MALITSEL$ = Mid$(MALIT$, 6 * KI2% - 5, 6 * QSEL%)
           MEJFECOM% = FECOMREAL%
           MEJHOCOM% = HOCOMREAL%
           MEJTUCOM% = TURNOINI%
           MEJFEFIN% = FEFINREAL%
           MEJHOFIN% = HOFINREAL%
           MEJTUFIN% = TURNOINI%
         End If
       Next KI2%
       '
       '
       ' REPITE PARA EL TIEMPO DE PREPARACION
       EQPX$ = Left$(MALITSEL$, 6)
       FECOMPROP% = MEJFECOM%
       HOCOMPROP% = MEJHOCOM%
       Call MEJORPRO(TIENEPREP, EQPX$, CANOPNEC, FECOMPROP%, HOCOMPROP%, FECOMREAL%, HOCOMREAL%, TURNOINI%, FEFINREAL%, HOFINREAL%, TURNOFIN%)
       FEFINPREP% = FEFINREAL%
       HOFINPREP% = HOFINREAL%
       STHH$ = MKI$(MEJFECOM%) + MKI$(MEJHOCOM%) + MKI$(MEJFEFIN%) + MKI$(MEJHOFIN%) + MKI$(FEFINPREP%) + MKI$(HOFINPREP%)
       O$ = REGLEIDO$("OPERFAT", I&)
       Call REPLA(O$, STHH$, 135, 12)
       Call REPLA(O$, AJUSTI$(MALITSEL$, 48), 151, 48)
       Call GRAREG("OPERFAT", O$, I&)
       '
       HOI1 = HOY(HOS$)
       FEX2 = MEJFECOM%
       If DIENTRE(HOI1, FEX2) >= 366 Then
         If COMALTER%("Imposible Secuenciar Operación\" + TRIM$(Text1(4)) + " de " + Text1(0) + " por Falta Absoluta Disponibilidad M.O.D.\\Continuar Cálculo\Abortar") = 2 Then
           Call FINAL("")
         End If
       End If
       '
       CANOPNE = CVS(Mid$(O$, 147, 4))
       '
       For KI4% = 1 To Len(MALITSEL$) Step 6
         EQPX$ = Mid$(MALITSEL$, KI4%, 6)
         MQX% = MAQNU%(EQPX$)
         VECHORES$(MQX%) = VECHORES$(MQX%) + STHH$
         '
         'reserva mano de obra directa en la seccion
         If KI4% = 1 Then
           SQX% = SEQNU%(EQPX$)
           PUNTINI& = 24& * 60& * (MEJFECOM% - DIACTU%) + 60 * (MEJHOCOM% \ 100) + (MEJHOCOM% Mod 100)
           PUNTFIN& = 24& * 60& * (MEJFEFIN% - DIACTU%) + 60 * (MEJHOFIN% \ 100) + (MEJHOFIN% Mod 100)
           Call RESERMOD(SQX%, PUNTINI&, PUNTFIN&, CANOPNE)
         End If
         '
       Next KI4%
       '
         FEX1 = MEJFECOM%
       Text1(5).TEXT = FECHATEX$(FEX1)
       Text1(6).TEXT = HORATEXNOR$(MEJHOCOM%)
         FEX2 = MEJFEFIN%
       Text1(7).TEXT = FECHATEX$(FEX2)
       Text1(8).TEXT = HORATEXNOR$(MEJHOFIN%)
       '
       For III% = 0 To 8
          Text1(III%).Refresh
       Next III%
       '
       NORFAXI$ = Text1(0)
       With OrdeFabr
         .FindFirst "IdSubOr = '" & NORFAXI$ & "'"
         If .NoMatch = False Then
           .Edit
           FEFINII% = FEX2
           !FechProyEntre = CVDAT(FEFINII%)
           !FechRecal = Now
           .Update
         End If
       End With
       '
       ' AQUI MARCAR FECHA Y HORA MAS TEMPRANA QUE PUEDEN EMPEZAR LAS
       ' OPERACIONES SIGUIENTES DE LA MISMA ORDEN DE FABRICACION
       COMSIG$ = Mid(STHH$, 5, 4)
       COMISIGUIEN# = 10000 * CDbl(CVI(Left$(COMSIG$, 2))) + CDbl(CVI(Mid$(COMSIG$, 3, 2)))
       For J& = 1 To FIN&
         If J& <> I& Then
           O1$ = REGLEIDO$("OPERFAT", J&)
           If Left$(O1$, 12) = NOFA$ Then
             If CVI(Mid$(O1, 15, 2)) > NUMEOP% Then
               FEHOPROGRA# = 10000 * CDbl(CVI(Mid$(O1$, 135, 2))) + CDbl(CVI(Mid$(O1$, 137, 2)))
               If COMISIGUIEN# > FEHOPROGRA# Then
                 Call REPLA(O1$, COMSIG$, 135, 4)
                 Call GRAREG("OPERFAT", O1$, J&)
               End If
             End If
           End If
         End If
       Next J&
       '
       ' AQUI MARCAR FECHA Y HORA MAS TEMPRANA QUE PUEDEN EMPEZAR LAS
       ' OPERACIONES DE ORDENES DE FABRICACION ASOCIADAS DE NIVEL DE FABRICACION SUPERIOR
       '
       O$ = REGLEIDO$("OPERFAT", I&)
       CODXX% = CVI(Mid$(O$, 13, 2))
       NIVXX% = CANIVEXPLO%(CODXX%)
       '
       For J& = 1 To FIN&
         If J& <> I& Then
           O1$ = REGLEIDO$("OPERFAT", J&)
           If Left$(O1$, 12) <> NOFA$ Then
             If Left$(O1$, 9) = Left$(NOFA$, 9) Then
               CODYY% = CVI(Mid$(O1$, 13, 2))
               If CODYY% <> CODXX% Then
                 If CANIVEXPLO%(CODYY%) > NIVXX% Then
                   FEHOPROGRA# = 10000 * CDbl(CVI(Mid$(O1$, 135, 2))) + CDbl(CVI(Mid$(O1$, 137, 2)))
                   If COMISIGUIEN# > FEHOPROGRA# Then
                     Call REPLA(O1$, COMSIG$, 135, 4)
                     Call GRAREG("OPERFAT", O1$, J&)
                   End If
                 End If
               End If
             End If
           End If
         End If
       Next J&
       '
139 Next I&
    Call CIERRARCH("*.*")
    '
    'GENERA ARCHIVO OPERFAP
    JJ& = 0
    For U& = 1 To ULTREG&("OPERFAT")
      O$ = REGLEIDO$("OPERFAT", U&)
      MALITSEL$ = Mid$(O$, 151, 48)
      For KI% = 1 To 48 Step 6
        MAQUIX$ = TRIM$(Mid$(MALITSEL$, KI%, 6))
        If MAQUIX$ <> "" Then
          Call REPLA(O$, MAQUIX$, 73, 6)
          Call REPLA(O$, Space$(18), 203, 18)
          JJ& = JJ& + 1
          Call GRAREG("OPERFAP", O$, JJ&)
        End If
      Next KI%
    Next U&
    Call SETULTREG("OPERFAP", JJ&)
    Screen.MousePointer = 1
    '
    Call ABREORFAB
    For KU& = 1 To ULTREG&("OPERFAP")
      XX$ = REGLEIDO$("OPERFAP", KU&)
      IDENOF$ = Left$(XX$, 12)
      NOPEI% = CVI(Mid$(XX$, 15, 2))
      COMIP = CDate(CVDAT(CVI(Mid$(XX$, 135, 2))) & " " & HORATEXNOR$(CVI(Mid$(XX$, 137, 2))))
      FINIP = CDate(CVDAT(CVI(Mid$(XX$, 139, 2))) & " " & HORATEXNOR$(CVI(Mid$(XX$, 141, 2))))
      With OperFab
        .FindFirst "IdSubOr = '" & IDENOF$ & "' AND NumeOper = " & NOPEI%
        If .NoMatch = False Then
          .Edit
          !ComiePro = COMIP
          !TermiPro = FINIP
          .Update
        End If
      End With
      '
    Next KU&
    '
    Check6.Value = 2
    Call GENORDEFDAT(0)
    Call GENEOPERFAP
    Call CIERRARCH("*.*")
    Check6.Value = 1
    '
    '\\\OT-08-0971-OD-24/10/08///
    If ULTREG&("CRPMENSA") > 0 Then
      OPP% = COMALTER%("\Se han Encontrado Errores Durante el Proceso de Cálculo\¿ Desea Visualizarlos ?\\Si \No")
      If OPP% = 1 Then
        Call FINAL("*CRPMENSA")
      End If
    End If
    '\\\OT-08-0971-OD-FIN///
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    COMPLECALCU% = 1
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Hide
    '
End Sub

Sub HODISTUR(EQ$, DIA%, TURNO%, DEHOR%, HAHOR%)
     '
     ' devuelve sobre DEHOR% y HAHOR% los horarios de disponibilidad del
     ' equipo EQ$ en la fecha DIA%, en base a los archivos de
     ' disponibilidad horaria
     '
     Static CADIA%, DHORT%(7, 8), HHORT%(7, 8)
     Static DHORM%(7, 8, 512), HHORM%(7, 8, 512)
     '
'If TRIM$(EQ$) = "SPW-2" Then
'AA1$ = FECHATEX$(DIA%)
'AAA = BBB
'End If
     If CADIA% = 0 Then
         '
         CADIA% = 7
         For KI% = 1 To 7
           For KJ% = 1 To 8
             DHORT%(KI%, KJ%) = COMTUR%(KJ%)
             HHORT%(KI%, KJ%) = FINTUR%(KJ%)
             If KJ% > CATURNOS% Then
               DHORT%(KI%, KJ%) = FINTUR%(CATURNOS%)
               HHORT%(KI%, KJ%) = FINTUR%(CATURNOS%)
             End If
           Next KJ%
         Next KI%
         '
         For KKI% = 1 To 8
            DHORT%(0, KKI%) = DHORT%(7, KKI%)
            HHORT%(0, KKI%) = HHORT%(7, KKI%)
         Next KKI%
         '
         TSEM = 0
         For U& = 1 To 7
            For KKI% = 1 To 8
               TTUR = HHORT%(U&, KKI%) - DHORT%(U&, KKI%)
               If TTUR < 0 Then TTUR = TTUR + 2400
               TSEM = TSEM + TTUR
            Next KKI%
         Next U&
         '
         If TSEM < 1 Then
            OPP% = COMALTER%("Proceso Imposible.\Falta Definir Horarios\y Turnos Diarios de Trabajo.\\Definir Horarios\Terminar.")
            If OPP% = 1 Then
                Call FINAL("*TURGEN")
            End If
            Call CIERRARCH("*.*")
            Call FINAL("")
         End If
         '
         For MQPX% = 0 To NUMAQ%
           MQX$ = MAQUI$(MQPX%)
           SQX% = SEQNU%(MQX$)
           For DSX% = 0 To 7
             For KKI% = 1 To 8
               DHORM%(DSX%, KKI%, MQPX%) = DHORT%(DSX%, KKI%)
               HHORM%(DSX%, KKI%, MQPX%) = HHORT%(DSX%, KKI%)
               If MQPX% > 0 Then
                 If DSX% > 0 Then
                   If MODISPO(DSX%, KKI%, SQX%) < 1 Then
                     DHORM%(DSX%, KKI%, MQPX%) = 0
                     HHORM%(DSX%, KKI%, MQPX%) = 0
                   End If
                 End If
               End If
             Next KKI%
           Next DSX%
         Next MQPX%
         '
         For U& = 1 To ULTREG&("TURMAQ")
            X$ = REGLEIDO$("TURMAQ", U&)
            DSX% = Asc(Left$(X$, 1))
            MQPX% = MAQNU%(Mid$(X$, 33, 6))
            For KKI% = 1 To 8
               DHORM%(DSX%, KKI%, MQPX%) = 0
               DHR% = CVI(Mid$(X$, 4 * KKI% - 2, 2))
               HHR% = CVI(Mid$(X$, 4 * KKI%, 2))
               If DHR% >= 0 Then
                  If DHR% <= 2400 Then
                     DHORM%(DSX%, KKI%, MQPX%) = DHR%
                  End If
               End If
               HHORM%(DSX%, KKI%, MQPX%) = DHORM%(DSX%, KKI%, MQPX%)
               If HHR% > DHORM%(DSX%, KKI%, MQPX%) Then
                  If HHR% <= 2400 Then
                     HHORM%(DSX%, KKI%, MQPX%) = HHR%
                  End If
               End If
               If DSX% = 7 Then
                  DHORM%(0, KKI%, MQPX%) = DHORM%(7, KKI%, MQPX%)
                  HHORM%(0, KKI%, MQPX%) = HHORM%(7, KKI%, MQPX%)
               End If
            Next KKI%
         Next U&
         '
     End If
     '
     DEHOR% = 0: HAHOR% = 0
     FECHA = DIA%
     DS% = DIASEM%(FECHA)
     If DS% = 0 Then DS% = 7
     MQP% = MAQNU%(EQ$)
     '
     ' VERIFICAR TURNOS ESPECIALES POR FECHA
     
     'If TURNO% > 0 Then
     '   If TURNO% <= 6 Then
     '      DEHOR% = DHORT%(DS%, TURNO%)
     '      HAHOR% = HHORT%(DS%, TURNO%)
     '   End If
     'End If
     '
     For U& = 1 To ULTREG&("CALFERI")
         If CVI(Left$(REGLEIDO$("CALFERI", U&), 2)) = DIA% Then
             DEHOR% = 0
             HAHOR% = 0
             GoTo 99
         End If
     Next U&
     '
     If DS% >= 0 Then
       If DS% <= 7 Then
         If TURNO% > 0 Then
           If TURNO% <= 6 Then
             If MQP% >= 0 Then
               If MQP% <= 512 Then
                 DEHOR% = DHORM%(DS%, TURNO%, MQP%)
                 HAHOR% = HHORM%(DS%, TURNO%, MQP%)
               End If
             End If
           End If
         End If
       End If
     End If
     '
     EQP$ = TRIM$(EQ$)
     For U& = 1 To ULTREG&("TURESP")
        X$ = REGLEIDO$("TURESP", U&)
        If CVI(Mid$(X$, 31, 2)) = DIA% Then
           If TRIM$(Mid$(X$, 33, 6)) = EQP$ Then
              KKI% = TURNO%
              DHR% = CVI(Mid$(X$, 4 * KKI% - 2, 2))
              HHR% = CVI(Mid$(X$, 4 * KKI%, 2))
              DEHOR% = 0
              If DHR% >= 0 Then
                 If DHR% <= 2400 Then
                    DEHOR% = DHR%
                 End If
              End If
              HAHOR% = DEHOR%
              If HHR% > DEHOR% Then
                 If HHR% <= 2400 Then
                    HAHOR% = HHR%
                 End If
              End If
           End If
        End If
     Next U&
     '
'AB1$ = FECHATEX$(DIA)
99 End Sub

Sub HODISTURVIE(EQ$, DIA%, TURNO%, DEHOR%, HAHOR%)
     '
     ' devuelve sobre DEHOR% y HAHOR% los horarios de disponibilidad del
     ' equipo EQ$ en la fecha DIA%, en base a los archivos de
     ' disponibilidad horaria
     '
     Static CADIA%, DHORT%(7, 6), HHORT%(7, 6)
     Static DHORM%(7, 6, 512), HHORM%(7, 6, 512)
     '
     If CADIA% = 0 Then
         '
         CADIA% = 7
         For U& = 1 To ULTREG&("TURGEN")
             X$ = REGLEIDO$("TURGEN", U&)
             DSX% = Asc(Left$(X$, 1))
             If DSX% > 0 Then
                 If DSX% <= 7 Then
                    For KKI% = 1 To 6
                       DHR% = CVI(Mid$(X$, 4 * KKI% - 2, 2))
                       HHR% = CVI(Mid$(X$, 4 * KKI%, 2))
                       If DHR% >= 0 Then
                         If DHR% <= 2400 Then
                             DHORT%(DSX%, KKI%) = DHR%
                         End If
                       End If
                       If DHORT%(DSX%, KKI%) < HHORT%(DSX%, KKI% - 1) Then
                          DHORT%(DSX%, KKI%) = HHORT%(DSX%, KKI% - 1)
                       End If
                       '
                       HHORT%(DSX%, KKI%) = DHORT%(DSX%, KKI%)
                       If HHR% > DHORT%(DSX%, KKI%) Then
                         If HHR% <= 2400 Then
                             HHORT%(DSX%, KKI%) = HHR%
                         End If
                       End If
                    Next KKI%
                 End If
             End If
         Next U&
         For KKI% = 1 To 6
            DHORT%(0, KKI%) = DHORT%(7, KKI%)
            HHORT%(0, KKI%) = HHORT%(7, KKI%)
         Next KKI%
         '
         TSEM = 0
         For U& = 1 To 7
            For KKI% = 1 To 6
               TSEM = TSEM + HHORT%(U&, KKI%) - DHORT%(U&, KKI%)
            Next KKI%
         Next U&
         '
         If TSEM < 1 Then
            OPP% = COMALTER%("Proceso Imposible.\Falta Definir Horarios\y Turnos Diarios de Trabajo.\\Definir Horarios\Terminar.")
            If OPP% = 1 Then
                Call FINAL("*TURGEN")
            End If
            Call CIERRARCH("*.*")
            Call FINAL("")
         End If
         '
         For MQPX% = 0 To NUMAQ%
            For DSX% = 0 To 7
               For KKI% = 1 To 6
                  DHORM%(DSX%, KKI%, MQPX%) = DHORT%(DSX%, KKI%)
                  HHORM%(DSX%, KKI%, MQPX%) = HHORT%(DSX%, KKI%)
               Next KKI%
            Next DSX%
         Next MQPX%
         '
         For U& = 1 To ULTREG&("TURMAQ")
            X$ = REGLEIDO$("TURMAQ", U&)
            DSX% = Asc(Left$(X$, 1))
            MQPX% = MAQNU%(Mid$(X$, 33, 6))
            For KKI% = 1 To 6
               DHORM%(DSX%, KKI%, MQPX%) = 0
               DHR% = CVI(Mid$(X$, 4 * KKI% - 2, 2))
               HHR% = CVI(Mid$(X$, 4 * KKI%, 2))
               If DHR% >= 0 Then
                  If DHR% <= 2400 Then
                     DHORM%(DSX%, KKI%, MQPX%) = DHR%
                  End If
               End If
               HHORM%(DSX%, KKI%, MQPX%) = DHORM%(DSX%, KKI%, MQPX%)
               If HHR% > DHORM%(DSX%, KKI%, MQPX%) Then
                  If HHR% <= 2400 Then
                     HHORM%(DSX%, KKI%, MQPX%) = HHR%
                  End If
               End If
               If DSX% = 7 Then
                  DHORM%(0, KKI%, MQPX%) = DHORM%(7, KKI%, MQPX%)
                  HHORM%(0, KKI%, MQPX%) = HHORM%(7, KKI%, MQPX%)
               End If
            Next KKI%
         Next U&
         '
     End If
     '
     DEHOR% = 0: HAHOR% = 0
     FECHA = DIA%
     DS% = DIASEM%(FECHA)
     MQP% = MAQNU%(EQ$)
     '
     ' VERIFICAR TURNOS ESPECIALES POR FECHA
     
     'If TURNO% > 0 Then
     '   If TURNO% <= 6 Then
     '      DEHOR% = DHORT%(DS%, TURNO%)
     '      HAHOR% = HHORT%(DS%, TURNO%)
     '   End If
     'End If
     '
     For U& = 1 To ULTREG&("CALFERI")
         If CVI(Left$(REGLEIDO$("CALFERI", U&), 2)) = DIA% Then
             DEHOR% = 0
             HAHOR% = 0
             GoTo 99
         End If
     Next U&
     '
     If DS% >= 0 Then
       If DS% <= 7 Then
         If TURNO% > 0 Then
           If TURNO% <= 6 Then
             If MQP% >= 0 Then
               If MQP% <= 512 Then
                 DEHOR% = DHORM%(DS%, TURNO%, MQP%)
                 HAHOR% = HHORM%(DS%, TURNO%, MQP%)
               End If
             End If
           End If
         End If
       End If
     End If
     '
     For U& = 1 To ULTREG&("TURESP")
        X$ = REGLEIDO$("TURESP", U&)
        If CVI(Mid$(X$, 31, 2)) = DIA% Then
           If TRIM$(Mid$(X$, 33, 6)) = EQP$ Then
              KKI% = TURNO%
              DHR% = CVI(Mid$(X$, 4 * KKI% - 2, 2))
              HHR% = CVI(Mid$(X$, 4 * KKI%, 2))
              DEHOR% = 0
              If DHR% >= 0 Then
                 If DHR% <= 2400 Then
                    DEHOR% = DHR%
                 End If
              End If
              HAHOR% = DEHOR%
              If HHR% > DEHOR% Then
                 If HHR% <= 2400 Then
                    HAHOR% = HHR%
                 End If
              End If
           End If
        End If
     Next U&
     '
99 End Sub

Function MAQNU%(MQS$)
    '
    MAQNU% = 0
    PPXY% = InStr(VECTOMAQ$, MQS$)
    If PPXY% Mod 10 = 3 Then
       MAQNU% = (PPXY% + 7) / 10
    End If
    '
End Function
'
Function SEQNU%(MQS$)
    '
    SEQNU% = 0
    PPXY% = InStr(VECTOMAQ$, MQS$)
    If PPXY% Mod 10 = 3 Then
       SEQNU% = CVI(Mid$(VECTOMAQ$, PPXY% + 6, 2))
    End If
    '
End Function
'
Sub SECUORF()  ' secuencia de ordenes de fabricacion
    '
    For UFORMU% = 1 To Forms.Count
      FORNA$ = TRIM$(UCase$(Forms(UFORMU% - 1).Name))
      If FORNA$ = "FORCAMAQ" Or FORNA$ = "CALCUCAR" Or FORNA$ = "CALCAR04" Then
        'On Error Resume Next
        FECOMPRO% = FECHANUM(Forms(UFORMU% - 1).Text1(9).TEXT)
        If Err < 1 Then GoTo 5
      End If
    Next UFORMU%
    Call MENSERR(24, "Proceso Imposible")
    Exit Sub
        
5   IFORCA% = UFORMU% - 1
    On Error GoTo 0
    '
    FIN& = ULTREG&("SECUORF")
    If FIN& > 4096 Then
        Call MENSERR(24, "Proceso Imposible - Vector fuera de Rango - Consulte")
        Exit Sub
    End If
    '
    SECCI = (-1): SEPROD$ = ""
    URESEC& = ULTREG&("SECPROD")
    If URESEC& > 0 Then
      If URESEC& = 1 Then
          XX$ = REGLEIDO$("SECPROD", 1)
        Else
          XX$ = REGSEL$("SECPROD")
      End If
      If Len(XX$) >= 4 Then
        SECCI = CVS(Left$(XX$, 4))
        SEPROD$ = Mid$(XX$, 5, 28)
      End If
    End If
    '
    Screen.MousePointer = 11
    For I& = 1 To FIN&
        X$ = REGLEIDO$("SECUORF", I&)
        ORFASEQ$(I&) = Left$(X$, 12)
        PRIVI%(I&) = 0
        STATU%(I&) = 0
    Next I&
    CAORFA% = FIN&
    CAORFB& = CAORFA%
    '
    FIN& = ULTREG&("ORDEFABR")
    For I& = 1 To FIN&
        F$ = REGLEIDO$("ORDEFABR", I&)
        ST% = Asc(Mid$(F$, 108, 1))
        If ST% = 1 Or ST% = 2 Then ' esta lanzada
            ORFASE$ = Mid$(F$, 69, 12)
            For J& = 1 To CAORFA%
                If ORFASEQ$(J&) = ORFASE$ Then
                    STATU%(J&) = ST%
                    If CVS(Mid$(F$, 113, 4)) = SECCI Then PRIVI%(J&) = 1
                    GoTo 20
                End If
            Next J&
            If ST% = 1 Or ST% = 2 Then
                DORFA$ = Mid$(F$, 5, 64)
                CAORFB& = CAORFB& + 1
                ORFASEQ$(CAORFB&) = ORFASE$
                STATU%(CAORFB&) = ST%
                PRIVI%(CAORFB&) = 0
                If CVS(Mid$(F$, 113, 4)) = SECCI Then PRIVI%(CAORFB&) = 1
                Call GRAREG("SECUORF", ORFASE$ + DORFA$, CAORFB&)
            End If
        End If
20  Next I&
    '
    J& = 0
    LISECUORF.LISECU.Clear
    For I& = 1 To CAORFB&
        If STATU%(I&) = 1 Then
            F$ = REGLEIDO$("SECUORF", I&)
            F$ = Left$(F$, 12) + Space$(6) + Mid$(F$, 13)
            If PRIVI%(I&) > 0 Then
              Call REPLA(F$, "<=>", 14, 4)
            End If
            LISECUORF.LISECU.AddItem F$
        End If
    Next I&
    '
    If Forms(IFORCA%).Check7.Value = 1 Then
       For I& = 1 To CAORFB&   ' incluye suspendidas
          If STATU%(I&) = 2 Then
             F$ = REGLEIDO$("SECUORF", I&)
             F$ = Left$(F$, 12) + Space$(6) + Mid$(F$, 13)
             If PRIVI%(I&) > 0 Then
               Call REPLA(F$, "<=>", 14, 4)
             End If
             LISECUORF.LISECU.AddItem F$
          End If
       Next I&
    End If
    '
    ' ESTO ES PARA UN ORDEN LOGICO
    For I& = 0 To LISECUORF.LISECU.ListCount - 1
30    NOFA$ = Left$(LISECUORF.LISECU.List(I&), 10)
      IOFA% = Val(Mid$(LISECUORF.LISECU.List(I&), 11, 4))
      For J& = I& + 1 To LISECUORF.LISECU.ListCount - 1
        NOFB$ = Left$(LISECUORF.LISECU.List(J&), 10)
        IOFB% = Val(Mid$(LISECUORF.LISECU.List(J&), 11, 4))
        If NOFB$ = NOFA$ Then
          If IOFB% > IOFA% Then
            TT1$ = LISECUORF.LISECU.List(I&)
            TT2$ = LISECUORF.LISECU.List(J&)
            LISECUORF.LISECU.RemoveItem (J&)
            LISECUORF.LISECU.RemoveItem (I&)
            LISECUORF.LISECU.AddItem TT2$, I&
            LISECUORF.LISECU.AddItem TT1$, J&
            GoTo 30
          End If
        End If
      Next J&
    Next I&
    '
    Screen.MousePointer = 1
    LISECUORF.Caption = "Secuenciación de Ordenes de Fabricación"
    If SEPROD$ <> "" Then
      LISECUORF.Caption = LISECUORF.Caption + " - Resalta: " + SEPROD$
    End If
    LISECUORF.Show 1
    '
End Sub
'

Sub CARHORAEQ()
    '
10  OBX$ = OBJPLA$("ENTRAEQUIPO", "")
    If OBX$ = "" Then
        Exit Sub
    End If
    '
    EQUI$ = AJUSTI$(OBX$, 6)
    YAESTA% = 0
    Call BLANARCH("!HORAEQUI")
    For U& = 1 To ULTREG&("HORAEQUI")
        X$ = REGLEIDO$("HORAEQUI", U&)
        If Left$(X$, 6) = EQUI$ Then
            YAESTA% = 1
            Call REGAPP("!HORAEQUI", X$)
        End If
    Next U&
    '
    If YAESTA% < 1 Then
        For UI% = 1 To 7
            X$ = REGBLAN$("HORAEQUI")
            Call REPLA(X$, Chr$(UI%), 7, 1)
            Call REPLA(X$, ECOARCH$("DIASEM", Chr$(UI%)), 8, 15)
            Call REGAPP("!HORAEQUI", X$)
        Next UI%
    End If
    Call CIERRARCH("!HORAEQUI")
    '
    VTB% = VENTABU%("HORAESP/TITUPAN=Horarios Especiales: " + TRIM$(EQUI$) + " (" + TRIM$(ECOARCH$("PADMAQ", EQUI$)) + ")")
    If VTB% < 0 Then GoTo 10
    '
    V& = 0
    For U& = 1 To ULTREG&("HORAEQUI")
        X$ = REGLEIDO$("HORAEQUI", U&)
        If Left$(X$, 6) <> EQUI$ Then
            V& = V& + 1
            Call GRAREG("HORAEQUI", X$, V&)
        End If
    Next U&
    Call SETULTREG("HORAEQUI", V&)
    Call CIERRARCH("HORAEQUI")
    '
    For U& = 1 To 7
        X$ = REGLEIDO$("!HORAEQUI", U&)
        TDES% = CVI(Mid$(X$, 23, 2))
        THAS% = CVI(Mid$(X$, 25, 2))
        If THAS% > TDES% Then GoTo 100
    Next U&
    GoTo 10
    '
100 For U& = 1 To 7
        X$ = REGLEIDO$("!HORAEQUI", U&)
        Call REPLA(X$, EQUI$, 1, 6)
        Call REGAPP("HORAEQUI", X$)
    Next U&
    Call CIERRARCH("HORAEQUI")
    GoTo 10
    '
End Sub
'
Sub REPLAMAQ(MAQUIBAS$, TTOT, MAQUIALT$, TALT)
    '
    MAQUIALT$ = ""
    TALT = TTOT
    COEFMAX = 0
    For IMQ% = 0 To 3
       '
       MAQUIBA$ = Mid$(MAQUIBAS$, 6 * IMQ% + 1, 6)
       maquiasi$ = MAQUIBA$
       COEFITIE = 0
       '
5      If TRIM$(MAQUIBA$) <> "" Then
          '
          ALTERMAQ$ = ""
          For JJ% = 1 To NUMAQ%
            If MAQUI$(JJ%) = MAQUIBA$ Then
              TIMAQBA = TIEPRO(JJ%)
              ALTERMAQ$ = MAEQUI$(JJ%)
              COEFALT = COEFTIE(JJ%)
              GoTo 10
            End If
          Next JJ%
          '
10        COEFITIE = 100
          If TRIM$(ALTERMAQ$) <> "" Then
             For JJ% = 1 To NUMAQ%
                If MAQUI$(JJ%) = ALTERMAQ$ Then
                   TIMAQAL = TIEPRO(JJ%)
                End If
             Next JJ%
             '
             If TIMAQAL + TTOT * COEFALT / 100 < TIMAQBA + TTOT Then
                maquiasi$ = ALTERMAQ$
                COEFITIE = COEFALT
                MAQUIBA$ = maquiasi$
                GoTo 5
             End If
             '
          End If
          '
       End If
       '
90     MAQUIALT$ = MAQUIALT$ + AJUSTI$(maquiasi$, 6)
       If TRIM$(maquiasi$) <> "" Then
          For JJ% = 1 To Len(MAQUISE$) Step 6
             If Mid$(MAQUISE$, JJ%, 6) = maquiasi$ Then GoTo 95
          Next JJ%
          MAQUISE$ = MAQUISE$ + maquiasi$
       End If
       '
95     If COEFITIE > COEFMAX Then
          COEFMAX = COEFITIE
       End If
       '
    Next IMQ%
    '
    TALT = TTOT
    If COEFMAX > 0 Then
       TALT = TTOT * COEFMAX / 100
    End If
    '
End Sub
       
Sub MEJORPRO(TIENEC, EQUIPROP$, CANOPNEC, FECOMPROP%, HOCOMPROP%, FECOMREAL%, HOCOMREAL%, TURNOINI%, FEFINREAL%, HOFINREAL%, TURNOFIN%)
    '
    Static HOXI
    If HOXI = 0 Then HOXI = HOY(HOS$)
    '
    TNEC0 = TIENEC
    TINEC = TIENEC
    MAQUIS$ = EQUIPROP$
    '
    TURNO% = 0
    FECOMI% = FECOMPROP%
    '
    FECOMIEN% = FECOMPROP%: HOCOMIEN% = HOCOMPROP%: TURCOMIEN% = 1
    '
511 YACOMIEN% = 0
    While TINEC > 0
      TURNO% = TURNO% + 1
      If TURNO% > CATURNOS% Then
        FECOMI% = DIPOST(FECOMI%)
        TURNO% = 1
      End If
      DIASIG% = 0
      Call HODISTUR(MAQUIS$, FECOMI%, TURNO%, DEHOR%, HAHOR%)
      If HAHOR% < DEHOR% Then DIASIG% = 1
      If FECOMI% = FECOMIEN% Then
        If DEHOR% < HOCOMIEN% Then
          DEHOR% = HOCOMIEN%
        End If
      End If
      TDIS = 0
      If DEHOR% <= HAHOR% Then
           TDIS = DIFHORA(DEHOR%, HAHOR%)
         Else
           If DIASIG% = 1 Then
             TDIS = DIFHORA(DEHOR%, 2400) + DIFHORA(0, HAHOR%)
           End If
      End If
      '
      If TDIS > 0 Then
        If YACOMIEN% = 0 Then
          FECOMIEN% = FECOMI%
          HOCOMIEN% = DEHOR%
          TURCOMIEN% = TURNO%
          YACOMIEN% = 1
        End If
        '
      End If
      '
      If TINEC <= TDIS Then
           MINUDI% = 60 * Int(TINEC) + 60 * (TINEC - Int(TINEC))
           HOFIN% = DEHOR%
           While MINUDI% > 0
             HOFIN% = HOFIN% + 1
             If (HOFIN% Mod 100) > 59 Then HOFIN% = HOFIN% + 40
             MINUDI% = MINUDI% - 1
           Wend
           TINEC = 0
         Else
           HOFIN% = HAHOR%
           TINEC = TINEC - TDIS
      End If
      FEFINAL% = FECOMI%
      TURFINAL% = TURNO%
      If HOFIN% > 2400 Then
        FECOMX1 = FECOMI%
        FEFINAL% = DIPOST(FECOMX1)
        HOFIN% = HOFIN% - 2400
      End If
      HOFINAL% = HOFIN%
    Wend
    '
    ' verificar si hay interferencias - AQUI ES DONDE TENGO QUE ACTUAR 31/05/2004
    '
    DESPRO# = 10000# * CDbl(FECOMIEN%) + CDbl(HOCOMIEN%)
    HASPRO# = 10000# * CDbl(FEFINAL%) + CDbl(HOFINAL%)
      If HASPRO# < DESPRO# Then
        FEFINAL% = FECOMIEN%
        HOFINAL% = HOCOMIEN%
        TURFINAL% = TURCOMIEN%
      End If
      '
    MQX% = MAQNU%(MAQUIS$)
    For IBLO% = 1 To Len(VECHORES$(MQX%)) Step 12
      STH$ = Mid$(VECHORES$(MQX%), IBLO%, 12)
      DESDE# = 10000# * CDbl(CVI(Mid$(STH$, 1, 2))) + CDbl(CVI(Mid$(STH$, 3, 2)))
      HASTA# = 10000# * CDbl(CVI(Mid$(STH$, 5, 2))) + CDbl(CVI(Mid$(STH$, 7, 2)))
      If DESPRO# < HASTA# And HASPRO# > DESDE# Then
        FECOMIEN% = CVI(Mid$(STH$, 5, 2))
        FECOMI% = FECOMIEN%
        HOCOMIEN% = CVI(Mid$(STH$, 7, 2))
        TURCOMIEN% = 1
        TURNO% = 0
        TINEC = TIENEC
        GoTo 511
      End If
    Next IBLO%
    '
    SQX% = SEQNU%(MAQUIS$)
    PUNTINI& = 24& * 60& * (FECOMIEN% - DIACTU%) + 60 * (HOCOMIEN% \ 100) + (HOCOMIEN% Mod 100)
    If PUNTINI& < LBound(MODISEC, 2) Then
      PUNTINIX& = LBound(MODISEC, 2)
      PUNTFINX& = UBound(MODISEC, 2)
      ReDim Preserve MODISEC(CASECCI%, PUNTINI& To PUNTFINX&)
      Call CARMODIS(PUNTINI&, PUNTINIX& - 1)
      ' CARGAR TIEMPOS DISPONIBLES
    End If
    PUNTFIN& = 24& * 60& * (FEFINAL% - DIACTU%) + 60 * (HOFINAL% \ 100) + (HOFINAL% Mod 100)
    If PUNTFIN& > UBound(MODISEC, 2) Then
      PUNTINIX& = LBound(MODISEC, 2)
      PUNTFINX& = UBound(MODISEC, 2)
      ReDim Preserve MODISEC(CASECCI%, PUNTINIX& To PUNTFIN&)
      Call CARMODIS(PUNTFINX&, PUNTFIN&)
      ' CARGAR TIEMPOS DISPONIBLES
    End If
    '
    ATXX% = 0
    For UKU& = PUNTINI& To PUNTFIN&
      If MODISEC(SQX%, UKU&) < CANOPNEC Then
        ATXX% = 1
        For UKK& = UKU& + 1 To PUNTFIN&
          If MODISEC(SQX%, UKK&) >= CANOPNEC Then GoTo 540
        Next UKK&
        UKK& = PUNTFIN&
        '
540     DIAXJ& = Int(UKK& / (24 * 60))
        FECOMIEN% = DIAXJ& + DIACTU%
        HOXU& = 100 * Int((UKK& - 24 * 60 * DIAXJ&) / 60)
        HOCOMIEN% = HOXU& + ((UKK& - 24 * 60 * DIAXJ&) Mod 60)
        FECOMI% = FECOMIEN%
        TURCOMIEN% = 1
        TURNO% = 0
        TINEC = TIENEC
        '
        FEX1 = FECOMPROP%
        FEX2 = FECOMIEN%
        If DIENTRE%(FEX1, FEX2) > 100 Then
          HOI1 = HOY(HOS$)
          FEX2 = HOI1
          While DIENTRE(HOI1, FEX2) <= 366
            FEX2 = DIPOST(FEX2)
          Wend
          FECOMIEN% = FEX2
          HOCOMIEN% = 0
          TURCOMIEN% = 1
          FEFINAL% = FEX2
          HOFINAL% = 0
          TURFINAL% = 1
          GoTo 90
        End If
        '
        GoTo 511
      End If
    Next UKU&
    '
90  FECOMREAL% = FECOMIEN%
    HOCOMREAL% = HOCOMIEN%
    TURNOINI% = TURCOMIEN%
    FEFINREAL% = FEFINAL%
    HOFINREAL% = HOFINAL%
    TURNOFIN% = TURFINAL%
    '
End Sub

Sub CARMODIS(COMIE&, FINIE&)
   '
   If FINIE& >= COMIE& Then
     For UKU& = COMIE& To FINIE&
       DIAXJ& = Int(UKU& / (24 * 60))
       DIAXI% = DIAXJ& + DIACTU%
       HOXU& = 100 * Int((UKU& - 24 * 60 * DIAXJ&) / 60)
       HOXI% = HOXU& + ((UKU& - 24 * 60 * DIAXJ&) Mod 60)
       '
       ITURNO% = 0
       For KKI% = 1 To CATURNOS%
         If HOXI% >= COMTUR%(KKI%) Then
           If HOXI% <= FINTUR%(KKI%) Then
             ITURNO% = KKI%
             Exit For
           End If
         End If
       Next KKI%
       '
       DISEMA% = DIASEM%(DIAXI%)
       '
       For ISECCI% = 1 To CASECCI%
         MDSX = 9999999
         If ITURNO% > 0 Then
           If DISEMA% > 0 Then
             MDSX = MODISPO(DISEMA%, ITURNO%, ISECCI%)
             If MDSX < 0.1 Then
               MDSX = 9999999
             End If
           End If
         End If
         MODISEC(ISECCI%, UKU&) = MDSX
       Next ISECCI%
     Next UKU&
   End If
   '
End Sub

Sub RESERMOD(ISECCI%, COMIE&, FINIE&, CANOPNE)
   '
   If FINIE& >= COMIE& Then
     For UKU& = COMIE& To FINIE&
       MODISEC(ISECCI%, UKU&) = MODISEC(ISECCI%, UKU&) - CANOPNE
     Next UKU&
   End If
   '
End Sub

Function MODISPON(MAQX$, FECHA%, TURNO%)
MODISPON = 999999999
Exit Function
   ' devuelve cantidad de operarios disponibles en fecha FECHA%,MAQX$,TURNO%
   SEMA = SECMAQ(MAQX$)
   '
   For UI% = 1 To CAVECMOD%
     If CVI(Left$(VECMODRES$(UI%), 2)) = FECHA% Then
       If CVS(Mid$(VECMODRES$(UI%), 3, 4)) = SEMA Then
         MODISPON = CVS(Mid$(VECMODRES$(UI%), 3 + 4 * TURNO%, 4))
         GoTo 99
       End If
     End If
   Next UI%
   '
   CAVECMOD% = CAVECMOD% + 1
   ReDim Preserve VECMODRES$(CAVECMOD%)
   VECMODRES$(CAVECMOD%) = MKI$(FECHA%) + MKS$(SEMA)
   '
   DIMOBRA$ = String$(12, 0)
   FEX = FECHA%
   DISE% = DIASEM%(FEX)
   If DISE% < 1 Then DISE% = DISE% + 7
   For U& = 1 To ULTREG&("DOTASEG")
     XX$ = REGLEIDO$("DOTASEG", U&)
     If CVS(Left$(XX$, 4)) = SEMA Then
       If Asc(Mid$(XX$, 5, 1)) = DISE% Then
         DIMOBRA$ = Mid$(XX$, 7, 12)
         Exit For
       End If
     End If
   Next U&
   '
   For U& = 1 To ULTREG&("DOTASEF")
     XX$ = REGLEIDO$("DOTASEG", U&)
     If CVS(Left$(XX$, 4)) = SEMA Then
       If CVI(Mid$(XX$, 5, 2)) = FECHA% Then
         DIMOBRA$ = Mid$(XX$, 7, 12)
         Exit For
       End If
     End If
   Next U&
   '
   For U& = 1 To 12 Step 2
     CAOPDI = CVI(Mid$(DIMOBRA$, U&, 2))
     VECMODRES$(CAVECMOD%) = VECMODRES$(CAVECMOD%) + MKS$(CAOPDI)
   Next U&
   MODISPON = CVS(Mid$(VECMODRES$(UI%), 3 + 4 * TURNO%, 4))
   '
99 End Function

Sub GRARESMOD(MAQX$, CANOPS, FECHINI%, TURINI%, FECHFIN%, TURFIN%)
   '
Exit Sub
   If TRIM$(MAQX$) = "" Then Exit Sub
   '
   SEMA = SECMAQ(MAQX$)
   FEXI% = FECHINI%
   TURNO% = TURINI%
   '
10 If FEXI% > FECHFIN% Then GoTo 99
   If FEXI% = FECHFIN% Then
     If TURNO% > TURFIN% Then
       GoTo 99
     End If
   End If
   '
   For UI% = 1 To CAVECMOD%
     If CVI(Left$(VECMODRES$(UI%), 2)) = FEXI% Then
       If CVS(Mid$(VECMODRES$(UI%), 3, 4)) = SEMA Then
         MODISPA = CVS(Mid$(VECMODRES$(UI%), 3 + 4 * TURNO%, 4))
         MODISPX = MODISPA - CANOPS
         If MODISPX < 0 Then MODISPX = 0
         Mid$(VECMODRES$(UI%), 3 + 4 * TURNO%, 4) = MKS$(MODISPX)
       End If
     End If
   Next UI%
   '
   TURNO% = TURNO% + 1
   If TURNO% > CATURNOS% Then
     FEX = FEXI%
     FEXI% = DIPOST(FEX)
     TURNO% = 1
     GoTo 10
   End If
   '
99 End Sub

Function SECMAQ(MAQX$)
   FGR$ = FILTERGETRECORD$("PADMAQ", 1, AJUSTI$(MAQX$, 6))
   SECMAQ = CVS(Mid$(FGR$, 37, 4))
End Function
