VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form CALCAR07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0D0D0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Re-Cálculo de Carga de Máquinas"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11130
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   11130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   1
      Left            =   4680
      ScaleHeight     =   375
      ScaleWidth      =   1215
      TabIndex        =   45
      Top             =   0
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   4530
      Left            =   10290
      ScaleHeight     =   4500
      ScaleWidth      =   1665
      TabIndex        =   41
      Top             =   0
      Width           =   1695
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   105
         ScaleHeight     =   225
         ScaleWidth      =   615
         TabIndex        =   47
         Top             =   3720
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   48
            Top             =   0
            Width           =   12000
         End
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
         Height          =   635
         Left            =   105
         Picture         =   "CALCAR07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Listados de Control"
         Top             =   1680
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
         Height          =   635
         Left            =   105
         Picture         =   "CALCAR07.frx":030A
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
         Picture         =   "CALCAR07.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Cancelar Proceso de Cálculo"
         Top             =   300
         Width           =   650
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0D0D0&
      Caption         =   "AVANCE DE PROCESO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4110
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6000
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0D0D0&
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
            Left            =   0
            TabIndex        =   36
            Top             =   1240
            Width           =   1000
         End
      End
      Begin VB.CheckBox Check6 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0D0D0&
         Caption         =   "Calculando Fechas Proyectadas de Entrega"
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
         Left            =   525
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1745
         Width           =   5055
      End
      Begin VB.TextBox Text1 
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
         Index           =   8
         Left            =   3345
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   " "
         Top             =   3570
         Width           =   645
      End
      Begin VB.TextBox Text1 
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
         Height          =   285
         Index           =   7
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   " "
         Top             =   3570
         Width           =   900
      End
      Begin VB.TextBox Text1 
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
         Index           =   6
         Left            =   3345
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   " "
         Top             =   3255
         Width           =   645
      End
      Begin VB.TextBox Text1 
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
         Height          =   285
         Index           =   5
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   " "
         Top             =   3255
         Width           =   900
      End
      Begin VB.TextBox Text1 
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
         Index           =   4
         Left            =   1995
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   " "
         Top             =   2940
         Width           =   3690
      End
      Begin VB.TextBox Text1 
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
         Index           =   3
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   " "
         Top             =   2940
         Width           =   435
      End
      Begin VB.TextBox Text1 
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
         Index           =   2
         Left            =   4515
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   " "
         Top             =   2310
         Width           =   1170
      End
      Begin VB.TextBox Text1 
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
         Index           =   1
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   " "
         Top             =   2625
         Width           =   4215
      End
      Begin VB.TextBox Text1 
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
         Index           =   0
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   " "
         Top             =   2310
         Width           =   1275
      End
      Begin VB.CheckBox Check4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0D0D0&
         Caption         =   "Verificando Tiempos Standard"
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
         Left            =   525
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1215
         Width           =   5160
      End
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0D0D0&
         Caption         =   "Ordenamiento Ordenes de Fabricación por Prioridad"
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
         Left            =   525
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   950
         Width           =   5055
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0D0D0&
         Caption         =   "Selección de Operaciones  Pendientes"
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
         Left            =   525
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   685
         Width           =   4110
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0D0D0&
         Caption         =   "Selección de Ordenes de Fabricación Lanzadas"
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
         Left            =   525
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   420
         Width           =   4950
      End
      Begin VB.CheckBox Check5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0D0D0&
         Caption         =   "Secuenciando Carga de Maquinas"
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
         Left            =   2505
         TabIndex        =   19
         Top             =   3630
         Width           =   750
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "a las"
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
         Left            =   2505
         TabIndex        =   18
         Top             =   3300
         Width           =   750
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "horas"
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
         Left            =   3765
         TabIndex        =   8
         Top             =   3630
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "horas"
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
         Left            =   3870
         TabIndex        =   7
         Top             =   3300
         Width           =   750
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
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
         Left            =   3020
         TabIndex        =   6
         Top             =   2385
         Width           =   1380
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0D0D0&
      Caption         =   "PARÁMETROS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4110
      Left            =   6300
      TabIndex        =   22
      Top             =   120
      Width           =   3795
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0D0D0&
         Caption         =   "O/F's SUSPENDIDAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
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
            Appearance      =   0  'Flat
            BackColor       =   &H00C0D0D0&
            Caption         =   "Incluir O/F's Suspendidas"
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
            Height          =   540
            Left            =   525
            TabIndex        =   34
            Top             =   315
            Width           =   2640
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0D0D0&
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
         Height          =   1170
         Left            =   210
         TabIndex        =   29
         Top             =   1680
         Width           =   3390
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0D0D0&
            Caption         =   "Por Prioridades"
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
            Left            =   525
            TabIndex        =   32
            Top             =   480
            Value           =   -1  'True
            Width           =   1905
         End
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0D0D0&
            Caption         =   "Por Secuencia"
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
            Picture         =   "CALCAR07.frx":0EDE
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Re-Ordenar Secuenciaa"
            Top             =   360
            Width           =   540
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0D0D0&
         Caption         =   "INICIO DE CARGA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
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
            Height          =   285
            Index           =   9
            Left            =   630
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   735
            Width           =   945
         End
         Begin VB.TextBox Text1 
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
            Height          =   285
            Index           =   10
            Left            =   2415
            Locked          =   -1  'True
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
            Visible         =   0   'False
            Width           =   175
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha"
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
            Index           =   10
            Left            =   630
            TabIndex        =   28
            Top             =   480
            Width           =   660
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Hora"
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
            Left            =   2430
            TabIndex        =   27
            Top             =   480
            Width           =   630
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
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4020
      OleObjectBlob   =   "CALCAR07.frx":1028
      Top             =   2100
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   4080
      OleObjectBlob   =   "CALCAR07.frx":125C
      TabIndex        =   46
      Top             =   420
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "CALCAR07"
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
Public CONSIMULACION%

Sub AGRESECOPERPEN()
   '
   Call ABREORFAB
   '
   'RECORRER EL OPERFAB TOMAR EL CODIGO DE LA MAQUINA
   SQLX$ = "SELECT * FROM OPERFAB "
   SQLX$ = SQLX$ + " WHERE CANTPEND > 0 "
   Dim OPERTMP As ADODB.Recordset
   Set OPERTMP = New ADODB.Recordset
   OPERTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   With OPERTMP
     Do While Not .EOF
       CODMAQ$ = SAFETEXT$(!Maquina)
       'TRAIGO EL SECTOR SEGUN LA MAQUINA DE LA TABLA EQUIPOS
       CONDI$ = "CODIGO_EQU = '" & CODMAQ$ & "'"
       SECT = XVALO(VALOBADA("EQUIPOS", "UBICA_EQU", CONDI$, "N0MESS"))
       '
       If SECT > 0 Then ' SI LO ENCUENTRA LO GRABA
          !SecciPro = SECT
          .Update
          GoTo 20
       Else
         SECT = SECMAQ(CODMAQ$)
         !SecciPro = SECT
         .Update
       End If
       '
20    .MoveNext
     Loop
   End With
   OPERTMP.Close
   Set OPERTMP = Nothing
End Sub

Sub FechaCalculadaFinOperacion(Maquina$, TiempoNecesario#, FechaComienzo%, HoraComienzo%, FechaFin%, HoraFin%)

    MAQUIS$ = Maquina$
    FECOMIEN% = FechaComienzo%
    HOCOMIEN% = HoraComienzo%
    TINEC = TiempoNecesario#
    FECOMI% = FECOMIEN%
    
511 YACOMIEN% = 0
    TURNO% = 0
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
    FechaFin% = FEFINAL%
    HoraFin% = HOFINAL%
    If HASPRO# < DESPRO# Then
        FechaFin% = FECOMIEN%
        HoraFin% = HOCOMIEN%
    End If
    '
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    CONSIMULACION% = 0
    Hide
End Sub

Sub Command2_Click()
    Command2.Enabled = False
    If CONSIMULACION% = 0 Then
        OPX% = COMALTER%("Opciones de Cálculo:\\Normal\Simulación")
        If OPX% < 1 Or OPX% > 2 Then GoTo 99
        If OPX% = 2 Then
            Call CARGALISEL("LISTSIMU", "SIMU_ORDEFABR", "IdSimulOR/F6;Origen/A48", "IdSimulOR > 0 ORDER BY IdSimulOR DESC", "DISTINCT")
            Call SELECHO("LISTSIMU/LISTA DE SIMULACIONES")
            SIMULA& = XVALO(VALACT1$("LISTSIMU"))
            If SIMULA& < 1 Then GoTo 99
            CONSIMULACION% = SIMULA&
            Call ACTUREGISTRO("SIMU_ORDEFABR", "STATUORF", "NUMEORFA>0", 0, REGAF&)
        End If
    End If
    Call CALCUMAQ
    If CONSIMULACION% > 0 Then
       Call COMUNI("Se Re-Iniciará Sistema de Carga de Máquinas Para su Refresco")
       Call FINAL("")
    End If
99  CONSIMULACION% = 0
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
'    EPAC$ = TRIM$(EMPREAC$)
'    If EPAC$ <> "" Then
'        Caption = Caption + "  -  " + EPAC$
'    End If
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
    End If
    '
    Call APLICACIONSKINS(Me, Picture1)
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
Sub CALCUMAQ1()   ' recalculo de carga de maquinas segun
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
      For i& = 1 To FIN&
        X$ = REGLEIDO$("SECUORF", i&)
        NORFAX$ = AJUSTI$(Left$(X$, 12), 12)
        CAOR% = CAOR% + 1
        If CAOR% > VECOR% Then
          VECOR% = VECOR% + 64
          ReDim Preserve ORFA$(VECOR%), PRIO%(VECOR%)
        End If
        ORFA$(CAOR%) = NORFAX$
        PRIO%(CAOR%) = CAOR%
        SECUORFA$ = SECUORFA$ + "@@" + AJUSTI$(NORFAX$, 14)
      Next i&
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
    'Set ORDEFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim ORDEFABTMP As ADODB.Recordset
    Set ORDEFABTMP = New ADODB.Recordset
    ORDEFABTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

    With ORDEFABTMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        JJ& = 0
        Do While (.EOF = False)
          STATUOF% = XVALO(!StatuOrf)
          If STATUOF% >= STATUMI% And STATUOF% <= STATUMA% Then ' esta lanzada
            PRIORI% = XVALO(!Priorid): If PRIORI% < 1 Or PRIORI% > 9 Then PRIORI% = 10
            PRIORI% = CAORFSEQ% + PRIORI% + 10000 * (STATUOF% - 1)
            ORFASE$ = "@@" + AJUSTI$(SAFETEXT$(!idsubor), 14)
            PPXX% = InStr(SECUORFA$, ORFASE$)
            If PPXX% < 1 Then
              CAOR% = CAOR% + 1
              If CAOR% > VECOR% Then
                VECOR% = VECOR% + 64
                ReDim Preserve ORFA$(VECOR%), PRIO%(VECOR%)
              End If
              ORFA$(CAOR%) = SAFETEXT$(!idsubor)
              PRIO%(CAOR%) = PRIORI%
              SECUORFA$ = SECUORFA$ + "@@" + AJUSTI$(NORFAX$, 14)
            End If
          End If
        .MoveNext
        Loop
      End If
      On Error GoTo 0
    End With
    ORDEFABTMP.Close
    Set ORDEFABTMP = Nothing
    '
'7.- Agregando a Cada Orden el Codigo de Producto Correspondiente
    '
    ReDim CIOF%(CAOR%), NVFB%(CAOR%), PUNTOPER$(CAOR%)
    JJ& = 0
    For IOR% = 1 To CAOR%
      SQLX$ = " SELECT * FROM ORDEFABR "
      SQLX$ = SQLX$ + " WHERE IdSubOr = '" & ORFA$(IOR%) & "' "
      '
      'Set ORDEFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
      Dim ORFABTEMP1 As ADODB.Recordset
      Set ORFABTEMP1 = FLEXCONN.Execute(FILTSQL$(SQLX$))

      With ORDEFABTMP
        On Error Resume Next
        .MoveLast
        If Err < 1 Then
          On Error GoTo 0
          CIOF%(IOR%) = XVALO(!COD_INTE)
          NVFB%(IOR%) = CANIVEXPLO%(CIOF%(IOR%))
        End If
        On Error GoTo 0
      End With
      Set ORFABTEMP1 = Nothing
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
      'Set OPERFAPTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
      Dim OPERFAPTMP2 As ADODB.Recordset
      Set OPERFAPTMP2 = New ADODB.Recordset
      OPERFAPTMP2.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

      With OPERFAPTMP2
        On Error Resume Next
        .MoveFirst
        If Err < 1 Then
          Do While (.EOF = False)
            CAREA# = XVALO(!CANTREAL)
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
      Set OPERFAPTMP2 = Nothing
      
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
      'Set OPERFAPTMP3 = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
      Dim OPERFAPTMP3 As ADODB.Recordset
      Set OPERFAPTMP3 = New ADODB.Recordset
      OPERFAPTMP3.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

      With OPERFAPTMP3
        On Error Resume Next
        .MoveFirst
        If Err < 1 Then
          Do While Not .EOF
            CAPROG# = XVALO(!CANTPROG)
            CAREA# = XVALO(!CANTREAL)
            CAPEN# = CAPROG# - CAREA#
            CIXII% = XVALO(!COD_INTE)
            NOPEII% = XVALO(!NUMEOPER)
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
                TPREP = XVALO(!TiemPrep): If CAREA# >= 0.05 Then TPREP = 0
                TFIJO = XVALO(!TiemFijo)
                TVARI = XVALO(!TiemVari) * CAPEN# / CAPROG# 'TPROD - TFIJO: If TVARI < 0 Then TVARI = 0
                TDEMO = XVALO(!TiemDemo) * CAPEN# / CAPROG#
                TMANT = XVALO(!TiemMHer) * CAPEN# / CAPROG#
                TLIMP = XVALO(!TiemLimp)
                TOTRO = XVALO(!TiemOtro) * CAPEN# / CAPROG#
                TTOTA = TPREP + TFIJO + TVARI + TDEMO + TMANT + TLIMP + TOTRO
                '
                O$ = REBLA$
                Call REPLA(O$, SAFETEXT$(!idsubor), 1, 12)
                Call REPLA(O$, MKI$(!COD_INTE), 13, 2)
                Call REPLA(O$, MKI$(!NUMEOPER), 15, 2)
                  DOPEXI$ = TRIM$(Str$(XVALO(!NUMEOPER))) + "." + SAFETEXT$(!DESCOPER)
                Call REPLA(O$, DOPEXI$, 17, 48)
                'Call REPLA(O$, !Destino, 65, 8)  ' FALTA COMPLETAR ESTE ELEMENTO
                Call REPLA(O$, SAFETEXT$(!Maquina), 73, 6)
                Call REPLA(O$, MKD$(!CANTPROG), 79, 8)
                Call REPLA(O$, MKD$(!CANTREAL), 87, 8)
                Call REPLA(O$, MKD$(!CANTPEND), 95, 8)
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
                Call REPLA(O$, SAFETEXT$(!EquiAux), 203, 18)
                Call REPLA(O$, MKS$(i&), 221, 4)
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
      Set OPERFAPTMP3 = Nothing
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
    For i& = 1 To FIN&
      Call TRACE(20, i&, FIN&)
      O$ = REGLEIDO$("OPERFAT", i&)
      '
      If CVI(Mid$(O$, 135, 2)) = 0 Then
        NORFAX$ = Left$(O$, 12)
        SQLX$ = " SELECT * FROM ORDEFABR "
        SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFAX$ & "' "
        '
        'Set OPERFAPTMP4 = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
        Dim OPERFAPTMP4 As ADODB.Recordset
        Set OPERFAPTMP4 = FLEXCONN.Execute(FILTSQL$(SQLX$))
        With OPERFAPTMP4
          On Error Resume Next
          .MoveLast
          If Err < 1 Then
            On Error GoTo 0
            FEPROLA% = CVINT%(!FechLan)
            If FEPROLA% > FECOMPRO% Then
              Call REPLA(O$, MKI$(FEPROLA%), 135, 2)
              Call REPLA(O$, MKI$(0), 137, 2)
              Call GRAREG("OPERFAT", O$, i&)
            End If
          End If
          On Error GoTo 0
        End With
        Set OPERFAPTMP4 = Nothing
      End If
      '
    Next i&
    '
    'AQUI AJUSTAR CON LA FECHA Y HORA DE LA ULTIMA OPERACION INFORMADA
    '
    For i& = 1 To FIN&
      Call TRACE(20, i&, FIN&)
      O$ = REGLEIDO$("OPERFAT", i&)
      '
      If CVI(Mid$(O$, 135, 2)) = 0 Then
        Call REPLA(O$, MKI$(FECOMPRO%), 135, 2)
        Call REPLA(O$, MKI$(HOCOMPRO%), 137, 2)
        Call GRAREG("OPERFAT", O$, i&)
      End If
      '
    Next i&
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
    For i& = 1 To FIN&
       Call TRACE(20, i&, FIN&)
       O$ = REGLEIDO$("OPERFAT", i&)
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
         Call GRAREG("OPERFAT", O$, i&)
       End If
       '
       Text1(0).TEXT = NOFA$
       CIII% = CVI(Mid$(O$, 13, 2))
       CAPROGG = CVD(Mid$(O$, 79, 8))
       Text1(1).TEXT = TRIM$(CODEXT$(CIII%)) + " - " + TRIM$(DESCRIT$(CIII%)) + " (" + TRIM$(Str$(CAPROGG)) + ")"
       '
       FEFINA% = CVI(Mid$(O$, 135, 2))
       HOFINA% = CVI(Mid$(O$, 137, 2))
       fex = FEFINA%
       Text1(2).TEXT = FECHATEX$(fex)
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
         Call MENSERR(24, "Equipo no Válido en Operación " + TRIM$(Text1(4)) + "\para Orden de Fabricación " + TRIM$(NOFA$) + ".")
         GoTo 139
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
       O$ = REGLEIDO$("OPERFAT", i&)
       Call REPLA(O$, STHH$, 135, 12)
       Call REPLA(O$, AJUSTI$(MALITSEL$, 48), 151, 48)
       Call GRAREG("OPERFAT", O$, i&)
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
       SQLX$ = " SELECT * FROM ORDEFABR"
       SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFAXI$ & "'"
       Dim ORDEFFFABR As ADODB.Recordset
       Set ORDEFFFABR = New ADODB.Recordset
       ORDEFFFABR.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
       With ORDEFFFABR 'OrdeFabr
'         .FindFirst "IdSubOr = '" & NORFAXI$ & "'"
'         If .NoMatch = False Then
           '.Edit
         If Not (.EOF And .BOF) Then
           FEFINII% = FEX2
           !FechProyEntre = CVDAT(FEFINII%)
           !FechRecal = Now
           .Update
         End If
       End With
       ORDEFFFABR.Close
       Set ORDEFFFABR = Nothing
       '
       ' AQUI MARCAR FECHA Y HORA MAS TEMPRANA QUE PUEDEN EMPEZAR LAS
       ' OPERACIONES SIGUIENTES DE LA MISMA ORDEN DE FABRICACION
       COMSIG$ = Mid(STHH$, 5, 4)
       COMISIGUIEN# = 10000 * CDbl(CVI(Left$(COMSIG$, 2))) + CDbl(CVI(Mid$(COMSIG$, 3, 2)))
       For J& = 1 To FIN&
         If J& <> i& Then
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
       O$ = REGLEIDO$("OPERFAT", i&)
       CODXX% = CVI(Mid$(O$, 13, 2))
       NIVXX% = CANIVEXPLO%(CODXX%)
       '
       For J& = 1 To FIN&
         If J& <> i& Then
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
139 Next i&
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
      SQLX$ = "SELECT * FROM OPERFAB"
      SQLX$ = SQLX$ + " WHERE IdSubOr = '" & IDENOF$ & "'"
      SQLX$ = SQLX$ + " AND NumeOper = " & NOPEI% & ""
      Dim OPERFFFAB As ADODB.Recordset
      Set OPERFFFAB = New ADODB.Recordset
      OPERFFFAB.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
      With OPERFFFAB 'OperFab
'        .FindFirst "IdSubOr = '" & IDENOF$ & "' AND NumeOper = " & NOPEI%
'        If .NoMatch = False Then
        If Not (.EOF And .BOF) Then
          '.Edit
          !ComiePro = COMIP
          !TermiPro = FINIP
          .Update
        End If
      End With
      OPERFFFAB.Close
      Set OPERFFFAB = Nothing
      '
    Next KU&
    '
    Check6.Value = 2
    Call GENORDEFDAT(0)
    Call GENEOPERFAP
    Call CIERRARCH("*.*")
    Check6.Value = 1
    '
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
         For MQPX% = 0 To NUMAQ%    ' MAQUINA
           MQX$ = MAQUI$(MQPX%)
           SQX% = SEQNU%(MQX$)
           For DSX% = 0 To 7        ' DIA DE SEMANA
             For KKI% = 1 To 8      ' NUMERO DE TURNO
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
      If FORNA$ = "FORCAMAQ" Or FORNA$ = "CALCUCAR" Or FORNA$ = "CALCAR07" Then
        'On Error Resume Next
        FECOMPRO% = FECHANUM(Forms(UFORMU% - 1).Text1(9).TEXT)
        If Err < 1 Then GoTo 5
      End If
    Next UFORMU%
    Call MENSERR(24, "Proceso Imposible")
    Exit Sub
    
5   CONSIMULACION% = 0
    If EXISTE_TABLA("SIMU_ORDEFABR") > 0 Then
       If CantRegistros("SIMU_ORDEFABR", , "NOMESS") > 0 Then
         OPX% = COMALTER%("Opciones de Trabajo\\No Incluir Última Simulación\Sí Incluir Última Simulación")
         If OPX% > 1 Then
           CONSIMULACION% = XVALO(CONTROL("CARMAQUI", "ULTISIMU"))
         End If
       End If
    End If

    IFORCA% = UFORMU% - 1
    On Error GoTo 0
    '
    ARCHORDEFABR$ = "ORDEFABR"
    ARCHSECUORF$ = "SECUORF"
    If CONSIMULACION% > 0 Then
      ARCHORDEFABR$ = "SRDEFABR"
      Call COPYSTRU("SECUORF", "SSECUORF")
      ARCHSECUORF$ = "SSECUORF"
    End If
    
    FIN& = ULTREG&(ARCHSECUORF$)
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
    
    
    For i& = 1 To FIN&
        X$ = REGLEIDO$(ARCHSECUORF$, i&)
        ORFASEQ$(i&) = Left$(X$, 12)
        PRIVI%(i&) = 0
        STATU%(i&) = 0
    Next i&
    CAORFA% = FIN&
    CAORFB& = CAORFA%
    '
    '
    FIN& = ULTREG&(ARCHORDEFABR$)
    For i& = 1 To FIN&
        f$ = REGLEIDO$(ARCHORDEFABR$, i&)
        ST% = Asc(Mid$(f$, 108, 1))
        If ST% = 1 Or ST% = 2 Then ' esta lanzada
            ORFASE$ = Mid$(f$, 69, 12)
            For J& = 1 To CAORFA%
                If ORFASEQ$(J&) = ORFASE$ Then
                    STATU%(J&) = ST%
                    If CVS(Mid$(f$, 113, 4)) = SECCI Then PRIVI%(J&) = 1
                    GoTo 20
                End If
            Next J&
            If ST% = 1 Or ST% = 2 Then
                DORFA$ = Mid$(f$, 5, 64)
                CAORFB& = CAORFB& + 1
                ORFASEQ$(CAORFB&) = ORFASE$
                STATU%(CAORFB&) = ST%
                PRIVI%(CAORFB&) = 0
                If CVS(Mid$(f$, 113, 4)) = SECCI Then PRIVI%(CAORFB&) = 1
                Call GRAREG(ARCHSECUORF$, ORFASE$ + DORFA$, CAORFB&)
            End If
        End If
20  Next i&
    '
    J& = 0
    LISECUORF.LISECU.Clear
    For i& = 1 To CAORFB&
        If STATU%(i&) = 1 Then
            f$ = REGLEIDO$(ARCHSECUORF$, i&)
            f$ = Left$(f$, 12) + Space$(6) + Mid$(f$, 13)
            If PRIVI%(i&) > 0 Then
              Call REPLA(f$, "<=>", 14, 4)
            End If
            LISECUORF.LISECU.AddItem f$
        End If
    Next i&
    '
    If Forms(IFORCA%).Check7.Value = 1 Then
       For i& = 1 To CAORFB&   ' incluye suspendidas
          If STATU%(i&) = 2 Then
             f$ = REGLEIDO$(ARCHSECUORF$, i&)
             f$ = Left$(f$, 12) + Space$(6) + Mid$(f$, 13)
             If PRIVI%(i&) > 0 Then
               Call REPLA(f$, "<=>", 14, 4)
             End If
             LISECUORF.LISECU.AddItem f$
          End If
       Next i&
    End If
    '
    ' ESTO ES PARA UN ORDEN LOGICO
    For i& = 0 To LISECUORF.LISECU.ListCount - 1
30    NOFA$ = Left$(LISECUORF.LISECU.LIST(i&), 10)
      IOFA% = Val(Mid$(LISECUORF.LISECU.LIST(i&), 11, 4))
      For J& = i& + 1 To LISECUORF.LISECU.ListCount - 1
        NOFB$ = Left$(LISECUORF.LISECU.LIST(J&), 10)
        IOFB% = Val(Mid$(LISECUORF.LISECU.LIST(J&), 11, 4))
        If NOFB$ = NOFA$ Then
          If IOFB% > IOFA% Then
            TT1$ = LISECUORF.LISECU.LIST(i&)
            TT2$ = LISECUORF.LISECU.LIST(J&)
            LISECUORF.LISECU.RemoveItem (J&)
            LISECUORF.LISECU.RemoveItem (i&)
            LISECUORF.LISECU.AddItem TT2$, i&
            LISECUORF.LISECU.AddItem TT1$, J&
            GoTo 30
          End If
        End If
      Next J&
    Next i&
    '
    Screen.MousePointer = 1
    LISECUORF.Caption = "Secuenciación de Ordenes de Fabricación"
    If SEPROD$ <> "" Then
      LISECUORF.Caption = LISECUORF.Caption + " - Resalta: " + SEPROD$
    End If
    LISECUORF.CONSIMULACION% = CONSIMULACION%
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
       
'
Sub MEJORPRO(TIENEC, EQUIPROP$, CANOPNEC, FECOMPROP%, HOCOMPROP%, FECOMREAL%, HOCOMREAL%, TURNOINI%, FEFINREAL%, HOFINREAL%, TURNOFIN%, Optional FECHAFINOPERANT%, Optional HORAFINOPERANT%, Optional OperacionTercerizada%)
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
    If TINEC = 0 Then GoTo 90
    '
    ' *** VERIFICA SI ES UNA OPERACION TERCERIZADA. SI ES ASI, SE TOMA QUE LA CAPACIDAD DE LA 'MAQUINA'
    ' *** ES INFINITA. O SEA, SIEMPRE ESTÁ DISPONIBLE.
    If OperacionTercerizada% = 1 Then
        GoTo 511
    End If
    
    ' ANTES DE SEGUIR, VERIFICA SI EXISTEN MAQUINAS ALTERNATIVAS CUYA FECHA Y HORA DE LIBERACION
    ' ES ANTES DE LA FECHA PROPUESTA (FECOMPROP% Y HOCOMPROP%)
    ' NUEVA_FECOMPRO% = FECOMPROP%: NUEVA_HOCOMPRO% = HOCOMPROP%: NUEVO_MAQUINA$ = MAQUIS$
    EQUIPO$ = MAQUIS$: EQU% = 0
    FinMaquinasAlternativas% = 0
    
    ' UTILIZADO PARA LA BUSQUEDA DE HUECOS
    EncontroHueco% = 0
    RevisaMaqDefault% = 1
    EquipoActivo$ = EQUIPO$
    EquipoOptimo$ = EQUIPO$
    FechaComienzoOperacion% = FECOMIEN%
    HoraComienzoOperacion% = HOCOMIEN%
    FechaInicioEncontrada% = 32767
    HoraComienzoEncontrada% = 0
    FECHAHOY% = HOY(HO$)
       
    ' *** VERIFICA QUE LA MAQUINA DEFAULT YA TIENE TIEMPO RESERVADO. SI NO TIENE RESERVA, NO BUSCA EN LAS ALTERNATIVAS
    If VECHORES(MAQNU%(EquipoActivo$)) <> "" And TINEC > 0 Then
        
        While (FinMaquinasAlternativas% = 0 Or RevisaMaqDefault% = 1)
            NumeroEquipo% = MAQNU%(EquipoActivo$) ' NUMERO DE MAQUINA
            RESERVAEQUIPOS$ = VECHORES(NumeroEquipo%) ' EQUIPO DEFAULT
            
                FechaInicioReservas% = CVI(Mid$(RESERVAEQUIPOS, 1, 2))
                If FechaInicioReservas% < FECHAHOY Then
                    VECHORES(NumeroEquipo%) = ""
                    RESERVAEQUIPOS$ = ""
                End If
                
                If TRIM$(EquipoActivo$) = "12" Then
                A = B
                End If

                ' TIENE DISPONIBILIDAD COMPLETA. NO CONTINUA BUSCANDO, SE QUEDA CON ESTA MAQUINA.
                If TRIM$(RESERVAEQUIPOS$) = "" Then
                    EquipoOptimo$ = EquipoActivo$
                    FechaInicioEncontrada% = FECOMIEN%
                    HoraComienzoEncontrada% = HOCOMIEN%
                    EncontroHueco% = 1
                End If
                
                ' VERIFICA SI HAY DISPONIBILIDAD ENTRE HUECOS DE OPERACIONES PROGRAMADAS POR MAQUINA
                For II& = 1 To Len(RESERVAEQUIPOS$) Step 12
                    Reserva1$ = Mid$(RESERVAEQUIPOS$, II&, 12)
                    Reserva2$ = Mid$(RESERVAEQUIPOS$, II& + 12, 12)

                    ' HUECO ENTRE OPERACIONES YA PROGRAMADAS
                    FechaComienzo% = CVI(Mid$(Reserva1$, 5, 2))      ' FECHA FIN DE LA PRIMERA OPERACION
                    HoraComienzo% = CVI(Mid$(Reserva1$, 7, 2))      ' HORA FIN DE LA PRIMERA OPERACION
                    FechaFin% = CVI(Mid$(Reserva2$, 1, 2))           ' FECHA INICIO DE LA SIGUIENTE OPERACION
                    HoraFin% = CVI(Mid$(Reserva2$, 3, 2))           ' HORA INICIO DE LA SIGUIENTE OPERACION

                    If FechaComienzo% < FECHAHOY% Then
                        FechaComienzo% = FECHAHOY%
                    End If
                    If FechaFin% = 0 Then FechaFin% = 32767
                    
                    ' ES LA ULTIMA OPERACION PROGRAMADA EN LA MAQUINA
                    fch1# = (10000# * CDbl(FechaComienzoOperacion%) + CDbl(HoraComienzoOperacion%))
                    fch2# = (10000# * CDbl(FechaFin%) + CDbl(HoraFin%))
                    
                    If FechaComienzoOperacion% < FechaComienzo% And II& = 1 Then
                        
                        Call FechaCalculadaFinOperacion(EquipoActivo$, CDbl(TINEC), FechaComienzoOperacion%, HoraComienzoOperacion%, FechaFinReal%, HoraFinReal%)
                        
                        fch1# = (10000# * CDbl(FechaFinReal%) + CDbl(HoraFinReal%))
                        fch2# = (10000# * CDbl(FechaComienzo%) + CDbl(HoraComienzo%))
                        If fch1# <= fch2# Then
                            'NUEVA_FECOMPRO% = FechaComienzoOperacion%
                            'NUEVA_HOCOMPRO% = 0
                            EquipoOptimo$ = EquipoActivo$
                            FechaInicioEncontrada% = FechaComienzoOperacion%
                            HoraComienzoEncontrada% = HoraComienzoOperacion%
                            EncontroHueco% = 1
                            'GoTo 50
                        End If
                    ElseIf FechaComienzoOperacion% < FechaFin% Then
                        
                        fch1# = (10000# * CDbl(FechaComienzo%) + CDbl(HoraComienzo%))
                        fch2# = (10000# * CDbl(FechaComienzoOperacion%) + CDbl(HoraComienzoOperacion%))
                        If (fch1# < fch2#) Then
                            FechaComienzo% = FechaComienzoOperacion%
                            HoraComienzo% = HoraComienzoOperacion%
                        End If
                    
                        fch1# = (10000# * CDbl(FechaComienzo%) + CDbl(HoraComienzo%))
                        fch2# = (10000# * CDbl(FechaInicioEncontrada%) + CDbl(HoraComienzoOperacion%))
                        'If (fch1# < fch2#) Or (FechaComienzo% = FechaInicioEncontrada% And RevisaMaqDefault% = 2) Then
                        If (fch1# < fch2#) Then
                        'If (FechaComienzo% < FechaInicioEncontrada%) Then
                            Call FechaCalculadaFinOperacion(EquipoActivo$, CDbl(TINEC), FechaComienzo%, HoraComienzo%, FechaFinReal%, HoraFinReal%)
                            
                            fch1# = (10000# * CDbl(FechaFinReal%) + CDbl(HoraFinReal%))
                            fch2# = (10000# * CDbl(FechaFin%) + CDbl(HoraFin%))
                            If (fch1# < fch2#) Then
                                'NUEVA_FECOMPRO% = FechaComienzo%
                                'NUEVA_HOCOMPRO% = HoraComienzo%
                                FechaInicioEncontrada% = FechaComienzo%
                                HoraComienzoEncontrada% = HoraComienzo%
                                EquipoOptimo$ = EquipoActivo$
                                EncontroHueco% = 1
                                'GoTo 50
                            End If
                        End If
                        
                        'NUEVO_MAQUINA$ = EQUIPO$: If RevisaMaqDefault = 3 Then NUEVO_MAQUINA$ = MAEQUI(EQU%)
                    ElseIf FechaFin% = 0 Then
                        'NUEVA_FECOMPRO% = FechaComienzo%
                        'NUEVA_HOCOMPRO% = HoraComienzo%
                        If FechaComienzo% < FechaInicioEncontrada% Then
                            If FechaComienzo% > FechaComienzoOperacion% Then
                                FechaInicioEncontrada% = FechaComienzo%
                                HoraComienzoEncontrada% = HoraComienzo%
                                EquipoOptimo$ = EquipoActivo$
                                EncontroHueco% = 1
                            End If
                        Else
                            
                        End If

                    End If
                Next II&
                
45              RevisaMaqDefault% = 2
                EquipoActivo$ = TRIM$(MAEQUI(NumeroEquipo%))
                If EquipoActivo$ = "" Then
                    FinMaquinasAlternativas% = 1
                    ElseIf EquipoActivo$ = "05" Then
                    A = B
                End If
            'Else
            '    FinMaquinasAlternativas% = 1
            'End If
        Wend
        
50 '      If FechaComienzoOperacion% < FECOMPROP% Or (FechaComienzoOperacion% = FECOMPROP% And HoraComienzoOperacion% < HOCOMPROP%) Then
'            MAQUIS$ = EquipoOptimo$
'            EQUIPROP$ = MAQUIS$
'        End If
        
        If EncontroHueco% = 1 Then
            FECOMIEN% = FechaInicioEncontrada%
            HOCOMIEN% = HoraComienzoEncontrada%
            FECOMI% = FechaInicioEncontrada%
            MAQUIS$ = EquipoOptimo$
            EQUIPROP$ = MAQUIS$
        End If
        
        
'        While FinMaquinasAlternativas% = 0
'            EQU% = MAQNU%(EQUIPO$)
'            If TRIM$(MAEQUI(EQU%)) <> "" And MAEQUI(EQU%) <> MAQUIS$ Then
'                MaquinaDefault# = 10000# * CDbl(NUEVA_FECOMPRO%) + CDbl(NUEVA_HOCOMPRO%)
'                ' *** FECHA Y HORA DE FINALIZACION DE LA MAQUINA ALTERNATIVA
'                RESERVAEQUIPOS$ = VECHORES(MAQNU%(MAEQUI$(EQU%)))
'                STH$ = Right$(RESERVAEQUIPOS$, 12)
'                MaquinaAlternativa# = 10000# * CDbl(CVI(Mid$(STH$, 5, 2))) + CDbl(CVI(Mid$(STH$, 7, 2)))
'                '
'                If MaquinaAlternativa# < MaquinaDefault# Then
'                     If MaquinaAlternativa# > 0 Then
'                        NUEVA_FECOMPRO% = CVI(Mid$(STH$, 5, 2))
'                        NUEVA_HOCOMPRO% = CVI(Mid$(STH$, 7, 2))
'                        NUEVO_MAQUINA$ = MAEQUI(EQU%)
'                     Else
'                        NUEVA_FECOMPRO% = CVI(Format$(Now, "dd/mm/yy"))
'                        NUEVA_HOCOMPRO% = HORANUM(Format$(Now, "hh:mm"))
'                        NUEVO_MAQUINA$ = MAEQUI(EQU%)
'                     End If
'                End If
'
'                EQUIPO$ = MAEQUI(EQU%)
'            Else
'                FinMaquinasAlternativas% = 1
'            End If
'        Wend
'
'        If NUEVA_FECOMPRO% < FECOMPROP% Or (NUEVA_FECOMPRO% = FECOMPROP% And NUEVA_HOCOMPRO% > HOCOMPROP%) Then
'            MAQUIS$ = NUEVO_MAQUINA$
'            EQUIPROP$ = MAQUIS$
'        End If
    End If
    
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
    
    ' *** ESTO ES PARA OBVIAR LAS MAQUINAS QUE CORRESPONDEN A TERCERIZACIONES
    ' *** MAS ABAJO CONTROLA SUPERPOSICION DE OPERACIONES EN UNA MISMA MAQUINA
    If OperacionTercerizada% = 1 Then GoTo 530
      
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
530 SQX% = SEQNU%(MAQUIS$)
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
        If TINEC = 0 Then GoTo 90     ' VER SI ESTO RESUELVE LA COLGADA - epb - 8/10/2010
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
   fex = FECHA%
   DISE% = DIASEM%(fex)
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
     fex = FEXI%
     FEXI% = DIPOST(fex)
     TURNO% = 1
     GoTo 10
   End If
   '
99 End Sub

Function SECMAQ(MAQX$)
   FGR$ = FILTERGETRECORD$("PADMAQ", 1, AJUSTI$(MAQX$, 6))
   SECMAQ = CVS(Mid$(FGR$, 37, 4))
End Function

Sub CALCUMAQ()   ' recalculo de carga de maquinas segun
                    ' ordenes de fabricacion en curso
' VECTOMAQ$ es un string que contiene las maquinas definidas Y LAS SECCIONES en bloques de 10 bytes
    archoperfat$ = "OPERFAT"
    ARCHOPERFAP$ = "OPERFAP"
    ARCHECORDEP$ = "ECORDEP"
    ARCHIORFPRI$ = "IORFPRI"

    If CONSIMULACION% > 0 Then ' SIMULACION
      Call COPYSTRU("OPERFAT", "SOPERFAT")
      Call COPYSTRU("OPERFAP", "SOPERFAP")
      Call COPYSTRU("ECORDEP", "SECORDEP")
      Call COPYSTRU("IORFPRI", "SIORFPRI")
      archoperfat$ = "SOPERFAT"
      ARCHOPERFAP$ = "SOPERFAP"
      ARCHECORDEP$ = "SECORDEP"
      ARCHIORFPRI$ = "SIORFPRI"
    End If
    '
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
      For i& = 1 To FIN&
        X$ = REGLEIDO$("SECUORF", i&)
        NORFAX$ = AJUSTI$(Left$(X$, 12), 12)
        CAOR% = CAOR% + 1
        If CAOR% > VECOR% Then
          VECOR% = VECOR% + 64
          ReDim Preserve ORFA$(VECOR%), PRIO%(VECOR%)
        End If
        ORFA$(CAOR%) = NORFAX$
        PRIO%(CAOR%) = CAOR%
        SECUORFA$ = SECUORFA$ + "@@" + AJUSTI$(NORFAX$, 14)
      Next i&
      CAORFSEQ% = CAOR%
    End If
    '
    Call COPYSTRU("ECORDEF", ARCHECORDEP$)
    LIORFALA.Clear
    STATUMI% = 1: STATUMA% = 1
    If INCLUSUS% = 1 Then STATUMA% = 2
    '
    Call ABREORFAB
    SQLX$ = " SELECT StatuOrf, Priorid, IDSUBOR, FechEnSol  FROM ORDEFABR "
    SQLX$ = SQLX$ + " WHERE StatuOrf >= " & STATUMI% & " "
    SQLX$ = SQLX$ + " AND StatuOrf <= " & STATUMA% & " "
    
    If CONSIMULACION% > 0 Then ' SIMULACION
        SQLX$ = SQLX$ + " UNION ALL"
        SQLX$ = SQLX$ + " SELECT  StatuOrf, Priorid, IDSUBOR, FechEnSol FROM SIMU_ORDEFABR "
'        SQLX$ = SQLX$ + " WHERE StatuOrf >= " & STATUMI% & " "
'        SQLX$ = SQLX$ + " AND StatuOrf <= " & STATUMA% & " "
        SQLX$ = SQLX$ + " WHERE IdSimulOR = " & CONSIMULACION%
    End If
    '
    'SQLX$ = SQLX$ + " ORDER BY IdSubOr "
    SQLX$ = SQLX$ + " ORDER BY FechEnSol, IdSubOr"
    '
    Dim ORDEFABTMP As ADODB.Recordset
    Set ORDEFABTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With ORDEFABTMP
'      On Error Resume Next
'      .MoveFirst
'      If Err < 1 Then
'        On Error GoTo 0
        JJ& = 0
        Do While (.EOF = False)
          STATUOF% = XVALO(!StatuOrf)
          If STATUOF% >= STATUMI% And STATUOF% <= STATUMA% Then ' esta lanzada
            PRIORI% = XVALO(!Priorid): If PRIORI% < 1 Or PRIORI% > 9 Then PRIORI% = 10
            PRIORI% = CAORFSEQ% + PRIORI% + 10000 * (STATUOF% - 1)
            ORFASE$ = "@@" + AJUSTI$(SAFETEXT$(!idsubor), 14)
            PPXX% = InStr(SECUORFA$, ORFASE$)
            If PPXX% < 1 Then
              CAOR% = CAOR% + 1
              If CAOR% > VECOR% Then
                VECOR% = VECOR% + 64
                ReDim Preserve ORFA$(VECOR%), PRIO%(VECOR%)
              End If
              ORFA$(CAOR%) = SAFETEXT$(!idsubor)
              PRIO%(CAOR%) = PRIORI%
              SECUORFA$ = SECUORFA$ + "@@" + AJUSTI$(NORFAX$, 14)
            End If
          End If
        .MoveNext
        Loop
'      End If
    End With
    '
'7.- Agregando a Cada Orden el Codigo de Producto Correspondiente
    '
    ReDim CIOF%(CAOR%), NVFB%(CAOR%), PUNTOPER$(CAOR%)
    JJ& = 0
    For IOR% = 1 To CAOR%
      SQLX$ = " SELECT COD_INTE FROM ORDEFABR "
      SQLX$ = SQLX$ + " WHERE IdSubOr = '" & ORFA$(IOR%) & "' "
      If CONSIMULACION% > 0 Then ' SIMULACION
        SQLX$ = SQLX$ + " UNION ALL SELECT COD_INTE FROM SIMU_ORDEFABR "
        SQLX$ = SQLX$ + " WHERE IdSubOr = '" & ORFA$(IOR%) & "' "
        SQLX$ = SQLX$ + " AND IdSimulOR = " & CONSIMULACION%
      End If
      
      '
      'Dim ORDEFABTMP As ADODB.Recordset
      Set ORDEFABTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
      With ORDEFABTMP
        On Error Resume Next
'        .MoveLast
        If Err < 1 Then
          On Error GoTo 0
          CIOF%(IOR%) = XVALO(!COD_INTE)
          NVFB%(IOR%) = CANIVEXPLO%(CIOF%(IOR%))
        End If
      End With
      '
      Y$ = REGBLAN(ARCHECORDEP$)
      Call REPLA(Y$, ORFA$(IOR%), 1, 12)
      Call REPLA(Y$, ECOARCH$("ECORDEF", ORFA$(IOR%)), 13, 52)
      JJ& = JJ& + 1
      Call GRAREG(ARCHECORDEP$, Y$, JJ&)
      '
    Next IOR%
    Call SETULTREG(ARCHECORDEP$, JJ&)
    '
'8.- Ordenando O/F's por Prioridad
    '
    Call ABREORFAB
    FIN& = CAOR%
    For II% = 1 To CAOR%
      ORFX$ = TRIM$(Left$(ORFA$(II%), 12))
      SQLX$ = " SELECT CANTREAL FROM OPERFAB "
      SQLX$ = SQLX$ + " WHERE IdSubOr = '" & ORFX$ & "' "
      SQLX$ = SQLX$ + " AND StatOper <= " & STATUMA% & " "
      SQLX$ = SQLX$ + " AND CantReal < CantPend "
      If CONSIMULACION% > 0 Then ' SIMULACION
            SQLX$ = SQLX$ + " UNION ALL "
            SQLX$ = SQLX$ + " SELECT  CANTREAL FROM SIMU_OPERFAB "
            SQLX$ = SQLX$ + " WHERE IdSubOr = '" & ORFX$ & "' "
            SQLX$ = SQLX$ + " AND StatOper <= " & STATUMA% & " "
            SQLX$ = SQLX$ + " AND CantReal < CantPend "
            SQLX$ = SQLX$ + " AND IdSimulOP = " & CONSIMULACION%
      End If
      
      '
      Set OPERFAPTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
      With OPERFAPTMP
        'On Error Resume Next
        '.MoveFirst
        'If Err < 1 Then
          Do While (.EOF = False)
            CAREA# = XVALO(!CANTREAL)
            If CAREA# > 0 Then     ' Las O/F's ya comenzadas tienen prioridad
              If PRIO%(II%) > 0 Then
                PRIO%(II%) = PRIO%(II%) - 32767
              End If
            End If
          .MoveNext
          Loop
        'End If
        On Error GoTo 0
      End With
    Next II%
    '
    JJ& = 0
    For II% = 1 To CAOR%
      ZZ$ = REGBLAN$(ARCHIORFPRI$)
      Call REPLA(ZZ$, MKI$(PRIO%(II%)), 1, 2)
      Call REPLA(ZZ$, MKS$(Val(Mid$(ORFA$(II%), 4, 6))), 3, 4)
      Call REPLA(ZZ$, MKI$(NVFB%(II%)), 7, 2)
      Call REPLA(ZZ$, ORFA$(II%), 9, 12)
      Call REPLA(ZZ$, MKI$(CIOF%(II%)), 21, 2)
      JJ& = JJ& + 1
      Call GRAREG(ARCHIORFPRI$, ZZ$, JJ&)
    Next II%
    Call SETULTREG(ARCHIORFPRI$, JJ&)
    Call CIERRARCH(ARCHIORFPRI$)
    
    ' *** SE ANULA EL ORDENAMIENTO POR PRIORIDAD Y SE DEJA EL ORDEN POR FECHA DE SOLICITUD DE ENTREGA
    ' *** QUE SE HACE MAS ARRIBA
    ' Call FILESORT(ARCHIORFPRI$, "", 1, 3, "ASC")
    '
    ' ***********************************************************************************************
    
    '
'9.- Re-Carga del Vector de O/F's
    II% = 0
    For JJ& = 1 To ULTREG(ARCHIORFPRI$)
      ZZ$ = REGLEIDO$(ARCHIORFPRI$, JJ&)
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
    REBLA$ = REGBLAN$(archoperfat$)
    JJ& = 0
    '
    For II% = 1 To CAOR%
      NOFA$ = TRIM$(ORFA$(II%))
      SQLX$ = " SELECT EquiAux, CantiOps, ComiePro, CANTPEND, CANTREAL, CANTPROG, MAQUINA, NUMEOPER"
      SQLX$ = SQLX$ + ", COD_INTE, IDSUBOR, TiemOtro, TiemLimp, TiemMHer, TiemDemo, TiemVari, TiemFijo, TiemPrep, DescOper"
      SQLX$ = SQLX$ + " FROM OPERFAB "
      SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NOFA$ & "' "
      SQLX$ = SQLX$ + " AND StatOper <= " & STATUMA% & " "
      SQLX$ = SQLX$ + " AND CantReal < CantProg "
      If CONSIMULACION% > 0 Then ' SIMULACION
        SQLX$ = SQLX$ + " UNION ALL "
        SQLX$ = SQLX$ + " SELECT EquiAux, CantiOps, ComiePro, CANTPEND, CANTREAL, CANTPROG, MAQUINA, NUMEOPER"
        SQLX$ = SQLX$ + ", COD_INTE, IDSUBOR, TiemOtro, TiemLimp, TiemMHer, TiemDemo, TiemVari, TiemFijo, TiemPrep, DescOper"
        SQLX$ = SQLX$ + " FROM SIMU_OPERFAB "
        SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NOFA$ & "' "
        SQLX$ = SQLX$ + " AND StatOper <= " & STATUMA% & " "
        SQLX$ = SQLX$ + " AND CantReal < CantProg "
        SQLX$ = SQLX$ + " AND IdSimulOP = " & CONSIMULACION%
      End If
      SQLX$ = SQLX$ + " ORDER BY NumeOper ASC "
      '
      Set OPERFAPTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
      With OPERFAPTMP
        On Error Resume Next
       ' .MoveFirst
       ' If Err < 1 Then
          Do While Not .EOF
            CAPROG# = XVALO(!CANTPROG)
            CAREA# = XVALO(!CANTREAL)
            CAPEN# = CAPROG# - CAREA#
            CIXII% = XVALO(!COD_INTE)
            NOPEII% = XVALO(!NUMEOPER)
            '
            If CAPEN# >= 0.05 Then
              If CAPROG# >= 0.05 Then
    
                    ' *** OBTIENE MARCA SECUENCIACION POR UNIDAD O POR LOTE
                    ' *** 0 = POR FINALIZACION DE LOTE DE OPERACION; 1= POR FINALIZACION DE UNA UNIDAD
                    MODOSECUENCIA% = 0
                    If XVALO(VALOBADA("SECOPER", "SecuenciaFabricacion", "CodIConj='" & CStr(CIXII%) & "' AND NumeOper=" & CStr(NOPEII%))) = 1 Then MODOSECUENCIA% = 1
                    '
                    TPREP = XVALO(!TiemPrep): If CAREA# >= 0.05 Then TPREP = 0
                    TFIJO = XVALO(!TiemFijo)
                        CAPEN_AUX# = CAPEN#
                        If MODOSECUENCIA% = 1 Then CAPEN_AUX# = 1 ' SECUENCIA DE FABRICACION POR UNIDAD
                    TVARI = XVALO(!TiemVari) * (CAPEN_AUX# / CAPROG#) 'TPROD - TFIJO: If TVARI < 0 Then TVARI = 0
                    TDEMO = XVALO(!TiemDemo) * (CAPEN_AUX# / CAPROG#)
                    TMANT = XVALO(!TiemMHer) * (CAPEN_AUX# / CAPROG#)
                    TLIMP = XVALO(!TiemLimp)
                    TOTRO = XVALO(!TiemOtro) * (CAPEN_AUX# / CAPROG#)
                    TTOTA = TPREP + TFIJO + TVARI + TDEMO + TMANT + TLIMP + TOTRO
                    '
                    O$ = REBLA$
                    Call REPLA(O$, SAFETEXT$(!idsubor), 1, 12)
                    Call REPLA(O$, MKI$(!COD_INTE), 13, 2)
                    Call REPLA(O$, MKI$(!NUMEOPER), 15, 2)
                      'DOPEXI$ = TRIM$(Str$(XVALO(!NUMEOPER))) + "." + SAFETEXT$(!DescOper)
                      DOPEXI$ = SAFETEXT$(!DESCOPER)
                    Call REPLA(O$, DOPEXI$, 17, 48)
                    'Call REPLA(O$, !Destino, 65, 8)  ' FALTA COMPLETAR ESTE ELEMENTO
                    Call REPLA(O$, SAFETEXT$(!Maquina), 73, 6)
                    Call REPLA(O$, MKD$(!CANTPROG), 79, 8)
                    Call REPLA(O$, MKD$(!CANTREAL), 87, 8)
                    Call REPLA(O$, MKD$(!CANTPEND), 95, 8)
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
                    Call REPLA(O$, SAFETEXT$(!EquiAux), 203, 18)
                    Call REPLA(O$, MKS$(i&), 221, 4)
                    '
                    ' *** SI MODOSECUENCIA% ES 1, INICIA LOS NIVELES PARA LA SECUENCIA POR UNIDAD DE FABRICACION
                    ' *** A CADA UNIDAD DE FABRICACION SE LE ASIGNA UN NIVEL DISTINTO Y TODA LA SECUENCIA DE FABRICACION
                    ' *** DE ESA UNIDAD TIENE EL MISMO NIVEL. ES UNA FORMA DE AGRUPAR LA FABRICACION DE CADA UNIDAD EN
                    ' *** FORMA INDEPENDIENTE
                    NRONIVEL& = 0
                    If MODOSECUENCIA% = 1 Then NRONIVEL& = 1
                    Call REPLA(O$, MKS$(NRONIVEL&), 199, 4)
                    '
                    JJ& = JJ& + 1
                    Call GRAREG(archoperfat$, O$, JJ&)
                    PUNTOPER$(II%) = PUNTOPER$(II%) + MKS$(JJ&)
                    '
                    ' *** REPITE OPERACION TANTAS COMO CANTIDADES UNITARIAS PENDIENTE HAYA
                    If MODOSECUENCIA% = 1 Then
                        For B& = 2 To CLng(CAPEN#)
                                XXT$ = O$
                                NRONIVEL& = NRONIVEL& + 1
                                Call REPLA(XXT$, MKS$(NRONIVEL&), 199, 4)
                                '
                                JJ& = JJ& + 1
                                Call GRAREG(archoperfat$, XXT$, JJ&)
                                PUNTOPER$(II%) = PUNTOPER$(II%) + MKS$(JJ&)
                        Next B&
                    End If
              End If
            End If
            '
          .MoveNext
          Loop
          '
        On Error GoTo 0
      End With
      '
    Next II%
    Call SETULTREG(archoperfat$, JJ&)
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
    FIN& = ULTREG&(archoperfat$)
    FIN1& = FIN&: If FIN1& < 1 Then FIN1& = 1
    '
    For i& = 1 To FIN&
      Call TRACE(20, i&, FIN&)
      O$ = REGLEIDO$(archoperfat$, i&)
      '
      If CVI(Mid$(O$, 135, 2)) = 0 Then
        NORFAX$ = Left$(O$, 12)
        SQLX$ = " SELECT FechLan FROM ORDEFABR "
        SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFAX$ & "' "
        If CONSIMULACION% > 0 Then ' SIMULACION
            SQLX$ = SQLX$ + " UNION ALL "
            SQLX$ = SQLX$ + " SELECT  FechLan FROM SIMU_ORDEFABR "
            SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFAX$ & "' "
            SQLX$ = SQLX$ + " AND IdSimulOR = " & CONSIMULACION%
        End If
        '
        Set ORDEFABTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))       ' , dbOpenSnapshot)
        With ORDEFABTMP
'          On Error Resume Next
'          .MoveLast
'          If Err < 1 Then
            On Error GoTo 0
            FEPROLA% = CVINT%(!FechLan)
            If FEPROLA% > FECOMPRO% Then
              Call REPLA(O$, MKI$(FEPROLA%), 135, 2)
              Call REPLA(O$, MKI$(0), 137, 2)
              Call GRAREG(archoperfat$, O$, i&)
            End If
'          End If
          On Error GoTo 0
        End With
      End If
      '
    Next i&
    '
    'AQUI AJUSTAR CON LA FECHA Y HORA DE LA ULTIMA OPERACION INFORMADA
    '
    For i& = 1 To FIN&
      Call TRACE(20, i&, FIN&)
      O$ = REGLEIDO$(archoperfat$, i&)
      '
      If CVI(Mid$(O$, 135, 2)) = 0 Then
        Call REPLA(O$, MKI$(FECOMPRO%), 135, 2)
        Call REPLA(O$, MKI$(HOCOMPRO%), 137, 2)
        Call GRAREG(archoperfat$, O$, i&)
      End If
      '
    Next i&
    Call CIERRARCH(archoperfat$)
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
    FIN& = ULTREG&(archoperfat$)
    CAOPERF& = ULTREG&(archoperfat$)            ' Cantidad de Operaciones Pendientes
    '
    '\\\OT-08-0978-OD-27/10/08///
    Call ACTUREGISTRO("ORDEFABR", "FechProyEntre", "NUMEORFA>0", CVDAT(FECHANUM("01/01/80")), REGAF&)
    '
    Call SETULTREG&("CRPMENSA", 0)
    '\\\OT-08-0978-OD-FIN///
    For i& = 1 To FIN&
       Call TRACE(20, i&, FIN&)
       O$ = REGLEIDO$(archoperfat$, i&)
       NOFA$ = Left$(O$, 12)
       If NOFAANT$ <> NOFA$ Then
            FECHLANZA% = CVI(Mid$(O$, 135, 2))
            NOFAANT$ = NOFA$
            PRINUMEOP% = CVI(Mid$(O$, 15, 2))
       End If
       
       If InStr(1, NOFA$, "200") > 0 Then
       A = B
       End If
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
         Call GRAREG(archoperfat$, O$, i&)
       End If
       '
       Text1(0).TEXT = NOFA$
       CIII% = CVI(Mid$(O$, 13, 2))
       CAPROGG = CVD(Mid$(O$, 79, 8))
       Text1(1).TEXT = TRIM$(CODEXT$(CIII%)) + " - " + TRIM$(DESCRIT$(CIII%)) + " (" + TRIM$(Str$(CAPROGG)) + ")"
       '
       FEFINA% = CVI(Mid$(O$, 135, 2))
       HOFINA% = CVI(Mid$(O$, 137, 2))
       fex = FEFINA%
       Text1(2).TEXT = FECHATEX$(fex)
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
       '\\\OT-08-0978-OD-27/10/08///

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
       ' *** OBTIENE LA FECHA Y HORA DE FINALIZACION DE LA OPERACION ANTERIOR
       ' *** PARA PASARSELO LUEGO MAS ABAJO A LA RUTINA MEJORPRO
'       FECHFINOPERANT% = 0
'       HORAFINOPERANT% = 0
'       PASO% = 0
'       NRONIVEL& = CVS(Mid$(O$, 199, 4))
'       For IXX& = I& - 1 To 1 Step -1
'           REG$ = REGLEIDO$(archoperfat$, IXX&)
'           If Left$(REG$, 12) = NOFA$ Then
'               '
'               ' SI ENCONTRO LA MISMA OPERACION EN LOS PASOS ANTERIORES, TOMA SU FECHA Y HORA DE FIN
'               ' (LIBERACION DE LA MAQUINA) Y LA CONSIDERA COMO FECHA Y HORA DE COMIENZO PARA LA ACTUAL
'               NOPERA% = CVI(Mid$(REG$, 15, 2))
'               NRONIVEL_ANT& = CVS(Mid$(REG$, 199, 4))
'
'               NIVELANT& = NRONIVEL& - 1
'               If NRONIVEL_ANT& > NRONIVEL& Then NIVELANT& = NRONIVEL_ANT&
'
'               If NOPERA% < NUMEOP% And (NRONIVEL_ANT& = NIVELANT& Or NRONIVEL_ANT& = 0) Then
'                   DDD$ = Mid$(REG$, 135, 12)
'                   COMSIG_OPERANT$ = Mid(DDD$, 5, 4) ' FECHA Y HORA FIN DE LA OPERACION ANTERIOR S/SECUENCIA DEL MAESTRO
'                   PASO% = 1
'                   Exit For
'               End If
'
''               If NOPERA% < NUMEOP% And (NRONIVEL_ANT& = NRONIVEL& Or NRONIVEL_ANT& = 0) Then
''                   DDD$ = Mid$(REG$, 135, 12)
''                   COMSIG_OPERANT$ = Mid(DDD$, 5, 4) ' FECHA Y HORA FIN DE LA OPERACION ANTERIOR S/SECUENCIA DEL MAESTRO
''                   PASO% = 1
''                   Exit For
''               End If
'           End If
'       Next IXX&
'       '
'       If PASO% > 0 Then
'            FECHFINOPERANT% = CVI(Mid$(COMSIG_OPERANT$, 1, 2))
'            HORAFINOPERANT% = CVI(Mid$(COMSIG_OPERANT$, 3, 2))
'       End If
       ' *** FIN DE OBTENCION DE LA FECHA Y HORA DE FINALIZACION DE LA OPERACION ANTERIOR
       '
       CodigoInterno% = CIII%
       NumeroOperacion% = NUMEOP%
       OperacionTercerizada% = XVALO(VALOBADA("SECOPER", "OPERACIONTERCEROS", "CODICONJ=" & CStr(CodigoInterno%) & " AND NUMEOPER=" & CStr(NumeroOperacion%)))
       '
       If (NumeroOperacion% = 60 Or NumeroOperacion% = 100) And InStr(1, NOFA$, "264") > 0 Then
       A = B
       End If
1200   For KI2% = 1 To QLIN% - QSEL% + 1
         FECOMPROP% = FECOM%
         HOCOMPROP% = HOCOM%
1230     COMIENPROPU# = 10000 * CDbl(FECOMPROP%) + CDbl(HOCOMPROP%)
         '
         For KI3% = KI2% To KI2% + QSEL% - 1
            EQPX$ = Mid$(MALIT$, 6 * KI3% - 5, 6)
            ' *** SE SEPARA LA VARIABLE CON EL NOMBRE DE LA MAQUINA, POR SI CAMBIA DENTRO DE LA RUTINA MEJOPRO
            ' *** TENER EN CUENTA QUE EN MEJORPRO SE CONTEPLAN LAS MAQUINAS ALTERNATIVAS
            EQPXX$ = EQPX$
            NRONIVEL_OPERANT& = CVS(Mid$(O$, 199, 4))
            'aqui EQPX$ es un equipo individual
            NRONIVEL_ANT& = CVS(Mid$(O$, 199, 4))
            Call MEJORPRO(TIENEC, EQPXX$, CANOPNEC, FECOMPROP%, HOCOMPROP%, FECOMREAL%, HOCOMREAL%, TURNOINI%, FEFINREAL%, HOFINREAL%, TURNOFIN%, FECHFINOPERANT%, HORAFINOPERANT%, OperacionTercerizada%)
            COMIENPOSIB# = 10000 * CDbl(FECOMREAL%) + CDbl(HOCOMREAL%)
'            If COMIENPOSIB# > COMIENPROPU# Then
'              FECOMPROP% = FECOMREAL%
'              HOCOMPROP% = HOCOMREAL%
'              GoTo 1230
'            End If
'
'            ' SI CAMBIO LA MAQUINA POR DEFAULT SEGUN SECUENCIA DE OPERACIONES DEL MAESTRO
'            ' POR UNA ALTERNATIVA DEFINIDA EN LA TABLA DE MAQUINAS Y EQUIPOS
            If EQPX$ <> EQPXX$ Then
                MALIT$ = EQPXX$
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
       O$ = REGLEIDO$(archoperfat$, i&)
       Call REPLA(O$, STHH$, 135, 12)
       Call REPLA(O$, AJUSTI$(MALITSEL$, 48), 151, 48)
       Call GRAREG(archoperfat$, O$, i&)
       '
       HOI1 = HOY(HOS$)
       FEX2 = MEJFECOM%
' *************************************************************************
' SE COMENTA EL CONTROL DE LA MANO DE OBRA
' *************************************************************************
'       If DIENTRE(HOI1, FEX2) >= 366 Then
'         If COMALTER%("Imposible Secuenciar Operación\" + TRIM$(Text1(4)) + " de " + Text1(0) + " por Falta Absoluta Disponibilidad M.O.D.\\Continuar Cálculo\Abortar") = 2 Then
'           Call FINAL("")
'         End If
'       End If
       '
       CANOPNE = CVS(Mid$(O$, 147, 4))
       '
       For KI4% = 1 To Len(MALITSEL$) Step 6
         EQPX$ = Mid$(MALITSEL$, KI4%, 6)
         MQX% = MAQNU%(EQPX$)
         
         ' BUSCA UBICACION
         RESERVAEQUIPO$ = VECHORES$(MQX%)
         
         STHH1$ = Right$(RESERVAEQUIPO$, 12)
         FechaFinUltimaReserva% = CVI(Mid$(STHH1$, 5, 2))
         HoraFinUltimaReserva% = CVI(Mid$(STHH1$, 7, 2))
         
         FechaComienzoNuevaReserva% = CVI(Mid$(STHH$, 5, 2))
         HoraComienzoNuevaReserva% = CVI(Mid$(STHH$, 7, 2))
         
         FechaHora1# = (10000# * CDbl(FechaFinUltimaReserva%) + CDbl(HoraFinUltimaReserva%))
         FechaHora2# = (10000# * CDbl(FechaComienzoNuevaReserva%) + CDbl(HoraComienzoNuevaReserva%))
         If (FechaHora1# > FechaHora2#) Then
'
'                    Debug.Print "ANTES DE INSERTAR"
'                    For iControl% = 1 To Len(VECHORES$(MQX%)) Step 12
'                        DatoReserva$ = Mid$(VECHORES$(MQX%), iControl%, 12)
'                        FechaFinReserva% = CVI(Mid$(DatoReserva$, 5, 2))
'                        HoraFinReserva% = CVI(Mid$(DatoReserva$, 7, 2))
'                        FechaHora3# = (10000# * CDbl(FechaFinReserva%) + CDbl(HoraFinReserva%))
'                        Debug.Print CStr(FechaHora3#)
'                    Next iControl%
'            A = B
            For iReserva% = Len(RESERVAEQUIPO$) To 1 Step -12
                DatoReserva$ = Mid$(RESERVAEQUIPO$, iReserva% + 1, 12)
                FechaFinReserva% = CVI(Mid$(DatoReserva$, 5, 2))
                HoraFinReserva% = CVI(Mid$(DatoReserva$, 7, 2))
                FechaHora3# = (10000# * CDbl(FechaFinReserva%) + CDbl(HoraFinReserva%))
                If FechaHora3# < FechaHora2# And FechaHora3# > 0 Then
                    VECHORES$(MQX%) = Mid$(RESERVAEQUIPO$, 1, iReserva% + 12) + STHH$ + Mid$(RESERVAEQUIPO$, iReserva% + 12 + 1)
                    
'                    Debug.Print "DESPUES DE INSERTAR"
'                    For iControl% = 1 To Len(VECHORES$(MQX%)) Step 12
'                        DatoReserva$ = Mid$(VECHORES$(MQX%), iControl%, 12)
'                        FechaFinReserva% = CVI(Mid$(DatoReserva$, 5, 2))
'                        HoraFinReserva% = CVI(Mid$(DatoReserva$, 7, 2))
'                        FechaHora3# = (10000# * CDbl(FechaFinReserva%) + CDbl(HoraFinReserva%))
'                        Debug.Print CStr(FechaHora3#)
'                    Next iControl%
                    
                    
                    Exit For
                End If
            Next iReserva%
         Else
            VECHORES$(MQX%) = VECHORES$(MQX%) + STHH$
         End If
         'VECHORES$(MQX%) = VECHORES$(MQX%) + STHH$
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
       SQLX$ = " SELECT * FROM ORDEFABR "
       SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFAXI$ & "'"
       Dim CORDEFABR As ADODB.Recordset
       Set CORDEFABR = New ADODB.Recordset
       CORDEFABR.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     '
       With CORDEFABR
         '.FindFirst "IdSubOr = '" & NORFAXI$ & "'"
         If Not (.BOF And .EOF) Then
           '.Edit
           FEFINII% = FEX2
           If CVINT(!FechProyEntre) < FEFINII% Then
               !FechProyEntre = CVDAT(FEFINII%)
           End If
           !FechRecal = Now
           .Update
         End If
       End With
       '
       If CONSIMULACION% > 0 Then ' SIMULACION
            SQLX$ = " SELECT * FROM SIMU_ORDEFABR "
            SQLX$ = SQLX$ + " WHERE IdSimulOR = " & CONSIMULACION%
    
            Set CORDEFABR = New ADODB.Recordset
            CORDEFABR.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
            '
            With CORDEFABR
             If Not (.BOF And .EOF) Then
               FEFINII% = FEX2
               If CVINT(!FechProyEntre) < FEFINII% Then
                   !FechProyEntre = CVDAT(FEFINII%)
               End If
               !FechRecal = Now
               !StatuOrf = 1
               .Update
             End If
            End With
       End If

       ' AQUI MARCAR FECHA Y HORA MAS TEMPRANA QUE PUEDEN EMPEZAR LAS
       ' OPERACIONES SIGUIENTES DE LA MISMA ORDEN DE FABRICACION
       Dim FINOPE As String
       COMSIG$ = Mid(STHH$, 5, 4)
       COMISIGUIEN# = 10000 * CDbl(CVI(Left$(COMSIG$, 2))) + CDbl(CVI(Mid$(COMSIG$, 3, 2)))
       NRONIVEL_OPERANT& = CVS(Mid$(O$, 199, 4))
       VECES% = 0: ENCONTRO_ORDEN% = 0
       For J& = 1 To FIN&
         'If j& <> I& Then
         If J& > i& Then
           O1$ = REGLEIDO$(archoperfat$, J&)
           If Left$(O1$, 12) = NOFA$ Then
             ENCONTRO_ORDEN% = 1
             NRONIVEL& = CVS(Mid$(O1, 199, 4))
             'Debug.Print CStr(CVI(Mid$(O1$, 15, 2))) & "    " & CStr(NRONIVEL&)
             OperacionActual% = CVI(Mid$(O1$, 15, 2))
             If OperacionActual% > NUMEOP% Or (OperacionActual% = NUMEOP% And NRONIVEL& > NRONIVEL_OPERANT&) Then
               FEHOPROGRA# = 10000 * CDbl(CVI(Mid$(O1$, 135, 2))) + CDbl(CVI(Mid$(O1$, 137, 2)))
               If COMISIGUIEN# > FEHOPROGRA# Then
                 Call REPLA(O1$, COMSIG$, 135, 4)
                 'Call GRAREG(ARCHOPERFAT$, O1$, j&)
               End If
               
               ' *** ESTABLECE LA FECHA DE COMIENZO DE LAS OPERACIONES AGREGADAS
               ' *** POR ESTAR LA OPERACION CONFIGURADA QUE TRABAJA CON SECUENCIA POR UNIDAD DE FABRICACION
               If NRONIVEL& > 0 Then
                 VECES% = VECES% + 1
                 Operacion% = CVI(Mid$(O1, 15, 2))
                 COMSIG_MISMAOPER$ = "    "
                 EQUIP$ = AJUSTI$(Mid$(O1$, 73, 6), 6)
                 EQU% = MAQNU%(EQUIP$)
                 
                 ' *** BUSCA EL FIN DE LA MISMA OPERACION EN PASOS ANTERIORES
                 For IXX& = J& - 1 To 1 Step -1
                    REG$ = REGLEIDO$(archoperfat$, IXX&)
                    If Left$(REG$, 12) = NOFA$ Then
                        '
                        ' SI ENCONTRO LA MISMA OPERACION EN LOS PASOS ANTERIORES, TOMA SU FECHA Y HORA DE FIN
                        ' (LIBERACION DE LA MAQUINA) Y LA CONSIDERA COMO FECHA Y HORA DE COMIENZO PARA LA ACTUAL
                        NOPERA% = CVI(Mid$(REG$, 15, 2))
                        NRONIVEL_ANT& = CVS(Mid$(REG$, 199, 4))
                        If NOPERA% = Operacion% And NRONIVEL_ANT& = (NRONIVEL& - 1) Then
                            DDD$ = Mid$(REG$, 135, 12)
                            COMSIG_MISMAOPER$ = Mid(DDD$, 5, 4) ' FECHA Y HORA FIN DE LA MISMA OPERACION ANTERIOR S/SECUENCIA POR UNIDAD
                            COMIENZO_MISMAOPER$ = Mid$(DDD$, 1, 4)
                            Exit For
                        End If
                    End If
                 Next IXX&
                 '
'                 ' *** SI TIENE MAQUINA ALTERNATIVA, EN PRINCIPIO NO UTILIZA EL FIN DE LA MISMA OPERACION DEL NIVEL ANTERIOR
'                 ' *** SINO QUE SE BASA EN LA FINALIZACION DE LA OPERACION ANTERIOR DEL MISMO NIVEL. SI ESTA ULTIMA NO EXISTE,
'                 ' *** ENTONCES LA MISMA OPERACION.
'                 EQUIP$ = AJUSTI$(Mid$(O1$, 73, 6), 6)
'                 EQU% = MAQNU%(EQUIP$)
'                 If TRIM$(MAEQUI(EQU%)) <> "" Then GoTo 544
                 '
                 COMSIG_OPERA$ = COMSIG_MISMAOPER$
                 ' *** VERIFICA SI LA FECHA Y HORA DE LA OPERACION ANTERIOR (SEGUN LA SECUENCIA EN EL MAESTRO DE ARTICULOS),
                 ' *** TIENE UNA FECHA Y HORA DE FINALIZACION POSTERIOR A LA LIBERACION DE LA MAQUINA DEL PASO ANTERIOR.
                 ' *** EN CASO DE SER MAYOR, SERA ESTA ULTIMA EL INICIO PARA LA OPERACION ACTUAL
544:             If J& > 1 Then
                    ' *** BUSCA EL FIN DE LA MISMA OPERACION EN PASOS ANTERIORES
                    PASO% = 0
                    For IXX& = J& - 1 To 1 Step -1
                       REG$ = REGLEIDO$(archoperfat$, IXX&)
                       If Left$(REG$, 12) = NOFA$ Then
                           ' SI ENCONTRO LA MISMA OPERACION EN LOS PASOS ANTERIORES, TOMA SU FECHA Y HORA DE FIN
                           ' (LIBERACION DE LA MAQUINA) Y LA CONSIDERA COMO FECHA Y HORA DE COMIENZO PARA LA ACTUAL
                           NOPERA% = CVI(Mid$(REG$, 15, 2))
                           NRONIVEL_ANT& = CVS(Mid$(REG$, 199, 4))
                           If NOPERA% < Operacion% And (NRONIVEL_ANT& = NRONIVEL& Or NRONIVEL_ANT& = 0) Then
                               DDD$ = Mid$(REG$, 135, 12)
                               COMSIG_OPERANT$ = Mid(DDD$, 5, 4) ' FECHA Y HORA FIN DE LA OPERACION ANTERIOR S/SECUENCIA DEL MAESTRO
                               PASO% = 1
                               Exit For
                           End If
                       End If
                    Next IXX&
                    '
                    If PASO% > 0 Then
                        FECHFINPADRE% = CVI(Mid$(COMSIG_MISMAOPER$, 1, 2))
                        HORAFINPADRE% = CVI(Mid$(COMSIG_MISMAOPER$, 3, 2))
                            
                        FECHFINOPERANT% = CVI(Mid$(COMSIG_OPERANT$, 1, 2))
                        HORAFINOPERANT% = CVI(Mid$(COMSIG_OPERANT$, 3, 2))
                            
                        ' *** SI ENCONTRO EL FIN DE LA OPERACION ANTERIOR DEL MISMO NIVEL (POR ESO ENTRO) Y
                        ' *** TIENE MAQUINA ALTERNATIVA, TOMA ESTA ULTIMA Y ANULA LA QUE VIENE DE LA MISMA OPERACION
                        ' *** PERO UNO NIVEL ANTERIOR
                        If TRIM$(MAEQUI(EQU%)) <> "" Then
                            FECHFINPADRE% = 0: HORAFINPADRE% = 0
                        End If
                        
                        If FECHFINOPERANT% > FECHFINPADRE% Or (FECHFINOPERANT% = FECHFINPADRE% And HORAFINOPERANT% > HORAFINPADRE%) Then
                           COMSIG_OPERA$ = COMSIG_OPERANT$
                        End If
                    Else
                        If NUMEOP% > PRINUMEOP% Then
                            FechaInicio% = HOY(HO$)
                            If FECHLANZA% > FechaInicio% Then FechaInicio% = FECHLANZA%
                            
                            COMSIG_OPERA$ = MKI$(FechaInicio%)
                            COMSIG_OPERA$ = COMSIG_OPERA$ & MKI$(HORANUM("00:00"))
                        ElseIf (NUMEOP% = PRINUMEOP% And OperacionTercerizada% = 1) Then
                            COMSIG_OPERA$ = COMIENZO_MISMAOPER$
                        End If
                    End If
                 End If
                 
                 Call REPLA(O1$, COMSIG_OPERA$, 135, 4) ' FECHA Y HORA DE COMIENZO
               End If
               '
               ' *** GRABA LA FECHA DE COMIENZO DE LAS OPERACIONES POSTERIORES A I&
               Call GRAREG(archoperfat$, O1$, J&)
               '
             End If
           End If
          
           ' *** FIN DE ASIGNACION DE FECHA DE COMIENZO PARA LA MODALIDAD
           ' *** DE OPERACIONES POR UNIDAD DE FABRICACION
           If VECES% > 3 Then Exit For
          
           ' *** SALE DEL BUCLE CUANDO HABIENDO ENCONTRADO LA ORDEN DE FABRICACION, LLEGA UN PUNTO
           ' *** DONDE ENCUENTRA QUE SE ACABARON LOS REGISTROS QUE CORRESPONDEN A DICHA ORDEN (NOFA$)
           If Left$(O1$, 12) <> NOFA$ And ENCONTRO_ORDEN% = 1 Then
                Exit For
           End If
           
         End If
       Next J&
       '
       ' AQUI MARCAR FECHA Y HORA MAS TEMPRANA QUE PUEDEN EMPEZAR LAS
       ' OPERACIONES DE ORDENES DE FABRICACION ASOCIADAS DE NIVEL DE FABRICACION SUPERIOR
       '
       O$ = REGLEIDO$(archoperfat$, i&)
       CODXX% = CVI(Mid$(O$, 13, 2))
       NIVXX% = CANIVEXPLO%(CODXX%)
       '
       For J& = 1 To FIN&
         If J& <> i& Then
           O1$ = REGLEIDO$(archoperfat$, J&)
           If Left$(O1$, 12) <> NOFA$ Then
             If Left$(O1$, 9) = Left$(NOFA$, 9) Then
               CODYY% = CVI(Mid$(O1$, 13, 2))
               If CODYY% <> CODXX% Then
                 If CANIVEXPLO%(CODYY%) > NIVXX% Then
                   FEHOPROGRA# = 10000 * CDbl(CVI(Mid$(O1$, 135, 2))) + CDbl(CVI(Mid$(O1$, 137, 2)))
                   If COMISIGUIEN# > FEHOPROGRA# Then
                     Call REPLA(O1$, COMSIG$, 135, 4)
                     Call GRAREG(archoperfat$, O1$, J&)
                   End If
                 End If
               End If
             End If
           End If
         End If
       Next J&
       '
139 Next i&
    Call CIERRARCH("*.*")
    '
    'GENERA ARCHIVO OPERFAP
    ARCHOPERFAP_AUXILIAR$ = "OPERFAP0"
    Call COPYSTRU(ARCHOPERFAP$, ARCHOPERFAP_AUXILIAR$)
    JJ& = 0
    For U& = 1 To ULTREG&(archoperfat$)
      O$ = REGLEIDO$(archoperfat$, U&)
      MALITSEL$ = Mid$(O$, 151, 48)
      For KI% = 1 To 48 Step 6
        MAQUIX$ = TRIM$(Mid$(MALITSEL$, KI%, 6))
        If MAQUIX$ <> "" Then
          Call REPLA(O$, MAQUIX$, 73, 6)
          Call REPLA(O$, Space$(18), 203, 18)
          JJ& = JJ& + 1
          Call GRAREG(ARCHOPERFAP$, O$, JJ&)
          Call GRAREG(ARCHOPERFAP_AUXILIAR$, O$, JJ&)
        End If
      Next KI%
    Next U&
    Call SETULTREG(ARCHOPERFAP$, JJ&)
    Call SETULTREG(ARCHOPERFAP_AUXILIAR$, JJ&)
    '
    Screen.MousePointer = 1
    '
    Call ABREORFAB
    Dim COPERFAB1 As ADODB.Recordset
    Dim COPERFAB As ADODB.Recordset

    For KU& = 1 To ULTREG&(ARCHOPERFAP$)
      XX$ = REGLEIDO$(ARCHOPERFAP$, KU&)
      IDENOF$ = Left$(XX$, 12)
      NOPEI% = CVI(Mid$(XX$, 15, 2))
      COMIP = CDate(CVDAT(CVI(Mid$(XX$, 135, 2))) & " " & HORATEXNOR$(CVI(Mid$(XX$, 137, 2))))
      FINIP = CDate(CVDAT(CVI(Mid$(XX$, 139, 2))) & " " & HORATEXNOR$(CVI(Mid$(XX$, 141, 2))))
      
      Set COPERFAB = New ADODB.Recordset
      SQLX$ = "SELECT * FROM OPERFAB WHERE IdSubOr = '" & IDENOF$ & "' AND NumeOper = " & NOPEI%
      COPERFAB.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
      '
      With COPERFAB
        '.FindFirst "IdSubOr = '" & IDENOF$ & "' AND NumeOper = " & NOPEI%
        If Not (.EOF And .BOF) Then
          '.Edit
          !ComiePro = COMIP
          !TermiPro = FINIP
          .Update
        End If
      End With
      '
      If CONSIMULACION% > 0 Then ' SIMULACION
        Set COPERFAB1 = New ADODB.Recordset
        SQLX$ = "SELECT * FROM SIMU_OPERFAB WHERE IdSubOr = '" & IDENOF$ & "' AND NumeOper = " & NOPEI%
        SQLX$ = SQLX$ + " AND IdSimulOP = " & CONSIMULACION%
        
        COPERFAB1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        '
        With COPERFAB1
          If Not (.EOF And .BOF) Then
            !ComiePro = COMIP
            !TermiPro = FINIP
            .Update
          End If
        End With
      End If
    Next KU&
    '
    Check6.Value = 2
    If CONSIMULACION% > 0 Then ' SIMULACION
      SIMUCARM.CONSIMULACION% = CONSIMULACION%
      Call SIMUCARM.GENORDEFDAT_SIMU(0)
      Call SIMUCARM.GENEOPERFAP_SIMU
    Else
      Call GENORDEFDAT(0)
      Call GENEOPERFAP
    End If
    
    Call CIERRARCH("*.*")
    Check6.Value = 1
    '
    Call AGRESECOPERPEN
    '
    '\\\OT-08-0978-OD-27/10/08///
    If ULTREG&("CRPMENSA") > 0 Then
      OPP% = COMALTER%("\Se han Encontrado Errores Durante el Proceso de Cálculo\¿ Desea Visualizarlos ?\\Si \No")
      If OPP% = 1 Then
        Call FINAL("*CRPMENSA")
      End If
    End If
    '\\\OT-08-0978-OD-FIN///
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    COMPLECALCU% = 1
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    If IsFormLoaded(SIMUCARM) Then
       Unload SIMUCARM
    End If
    Hide
    '
End Sub


Sub CALCUMAQ_VIEJO()   ' recalculo de carga de maquinas segun
                    ' ordenes de fabricacion en curso
' VECTOMAQ$ es un string que contiene las maquinas definidas Y LAS SECCIONES en bloques de 10 bytes
    archoperfat$ = "OPERFAT"
    ARCHOPERFAP$ = "OPERFAP"
    ARCHECORDEP$ = "ECORDEP"
    ARCHIORFPRI$ = "IORFPRI"

    If CONSIMULACION% > 0 Then ' SIMULACION
      Call COPYSTRU("OPERFAT", "SOPERFAT")
      Call COPYSTRU("OPERFAP", "SOPERFAP")
      Call COPYSTRU("ECORDEP", "SECORDEP")
      Call COPYSTRU("IORFPRI", "SIORFPRI")
      archoperfat$ = "SOPERFAT"
      ARCHOPERFAP$ = "SOPERFAP"
      ARCHECORDEP$ = "SECORDEP"
      ARCHIORFPRI$ = "SIORFPRI"
    End If
    '
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
      For i& = 1 To FIN&
        X$ = REGLEIDO$("SECUORF", i&)
        NORFAX$ = AJUSTI$(Left$(X$, 12), 12)
        CAOR% = CAOR% + 1
        If CAOR% > VECOR% Then
          VECOR% = VECOR% + 64
          ReDim Preserve ORFA$(VECOR%), PRIO%(VECOR%)
        End If
        ORFA$(CAOR%) = NORFAX$
        PRIO%(CAOR%) = CAOR%
        SECUORFA$ = SECUORFA$ + "@@" + AJUSTI$(NORFAX$, 14)
      Next i&
      CAORFSEQ% = CAOR%
    End If
    '
    Call COPYSTRU("ECORDEF", ARCHECORDEP$)
    LIORFALA.Clear
    STATUMI% = 1: STATUMA% = 1
    If INCLUSUS% = 1 Then STATUMA% = 2
    '
    Call ABREORFAB
    SQLX$ = " SELECT StatuOrf, Priorid, IDSUBOR  FROM ORDEFABR "
    SQLX$ = SQLX$ + " WHERE StatuOrf >= " & STATUMI% & " "
    SQLX$ = SQLX$ + " AND StatuOrf <= " & STATUMA% & " "
    
    If CONSIMULACION% > 0 Then ' SIMULACION
        SQLX$ = SQLX$ + " UNION ALL"
        SQLX$ = SQLX$ + " SELECT  StatuOrf, Priorid, IDSUBOR FROM SIMU_ORDEFABR "
'        SQLX$ = SQLX$ + " WHERE StatuOrf >= " & STATUMI% & " "
'        SQLX$ = SQLX$ + " AND StatuOrf <= " & STATUMA% & " "
        SQLX$ = SQLX$ + " WHERE IdSimulOR = " & CONSIMULACION%
    End If
    '
    SQLX$ = SQLX$ + " ORDER BY IdSubOr "
    '
    Dim ORDEFABTMP As ADODB.Recordset
    Set ORDEFABTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With ORDEFABTMP
'      On Error Resume Next
'      .MoveFirst
'      If Err < 1 Then
'        On Error GoTo 0
        JJ& = 0
        Do While (.EOF = False)
          STATUOF% = XVALO(!StatuOrf)
          If STATUOF% >= STATUMI% And STATUOF% <= STATUMA% Then ' esta lanzada
            PRIORI% = XVALO(!Priorid): If PRIORI% < 1 Or PRIORI% > 9 Then PRIORI% = 10
            PRIORI% = CAORFSEQ% + PRIORI% + 10000 * (STATUOF% - 1)
            ORFASE$ = "@@" + AJUSTI$(SAFETEXT$(!idsubor), 14)
            PPXX% = InStr(SECUORFA$, ORFASE$)
            If PPXX% < 1 Then
              CAOR% = CAOR% + 1
              If CAOR% > VECOR% Then
                VECOR% = VECOR% + 64
                ReDim Preserve ORFA$(VECOR%), PRIO%(VECOR%)
              End If
              ORFA$(CAOR%) = SAFETEXT$(!idsubor)
              PRIO%(CAOR%) = PRIORI%
              SECUORFA$ = SECUORFA$ + "@@" + AJUSTI$(NORFAX$, 14)
            End If
          End If
        .MoveNext
        Loop
'      End If
    End With
    '
'7.- Agregando a Cada Orden el Codigo de Producto Correspondiente
    '
    ReDim CIOF%(CAOR%), NVFB%(CAOR%), PUNTOPER$(CAOR%)
    JJ& = 0
    For IOR% = 1 To CAOR%
      SQLX$ = " SELECT COD_INTE FROM ORDEFABR "
      SQLX$ = SQLX$ + " WHERE IdSubOr = '" & ORFA$(IOR%) & "' "
      If CONSIMULACION% > 0 Then ' SIMULACION
        SQLX$ = SQLX$ + " UNION ALL SELECT COD_INTE FROM SIMU_ORDEFABR "
        SQLX$ = SQLX$ + " WHERE IdSubOr = '" & ORFA$(IOR%) & "' "
        SQLX$ = SQLX$ + " AND IdSimulOR = " & CONSIMULACION%
      End If
      
      '
      'Dim ORDEFABTMP As ADODB.Recordset
      Set ORDEFABTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
      With ORDEFABTMP
        On Error Resume Next
'        .MoveLast
        If Err < 1 Then
          On Error GoTo 0
          CIOF%(IOR%) = XVALO(!COD_INTE)
          NVFB%(IOR%) = CANIVEXPLO%(CIOF%(IOR%))
        End If
      End With
      '
      Y$ = REGBLAN(ARCHECORDEP$)
      Call REPLA(Y$, ORFA$(IOR%), 1, 12)
      Call REPLA(Y$, ECOARCH$("ECORDEF", ORFA$(IOR%)), 13, 52)
      JJ& = JJ& + 1
      Call GRAREG(ARCHECORDEP$, Y$, JJ&)
      '
    Next IOR%
    Call SETULTREG(ARCHECORDEP$, JJ&)
    '
'8.- Ordenando O/F's por Prioridad
    '
    Call ABREORFAB
    FIN& = CAOR%
    For II% = 1 To CAOR%
      ORFX$ = TRIM$(Left$(ORFA$(II%), 12))
      SQLX$ = " SELECT CANTREAL FROM OPERFAB "
      SQLX$ = SQLX$ + " WHERE IdSubOr = '" & ORFX$ & "' "
      SQLX$ = SQLX$ + " AND StatOper <= " & STATUMA% & " "
      SQLX$ = SQLX$ + " AND CantReal < CantPend "
      If CONSIMULACION% > 0 Then ' SIMULACION
            SQLX$ = SQLX$ + " UNION ALL "
            SQLX$ = SQLX$ + " SELECT  CANTREAL FROM SIMU_OPERFAB "
            SQLX$ = SQLX$ + " WHERE IdSubOr = '" & ORFX$ & "' "
            SQLX$ = SQLX$ + " AND StatOper <= " & STATUMA% & " "
            SQLX$ = SQLX$ + " AND CantReal < CantPend "
            SQLX$ = SQLX$ + " AND IdSimulOP = " & CONSIMULACION%
      End If
      
      '
      Set OPERFAPTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
      With OPERFAPTMP
        'On Error Resume Next
        '.MoveFirst
        'If Err < 1 Then
          Do While (.EOF = False)
            CAREA# = XVALO(!CANTREAL)
            If CAREA# > 0 Then     ' Las O/F's ya comenzadas tienen prioridad
              If PRIO%(II%) > 0 Then
                PRIO%(II%) = PRIO%(II%) - 32767
              End If
            End If
          .MoveNext
          Loop
        'End If
        On Error GoTo 0
      End With
    Next II%
    '
    JJ& = 0
    For II% = 1 To CAOR%
      ZZ$ = REGBLAN$(ARCHIORFPRI$)
      Call REPLA(ZZ$, MKI$(PRIO%(II%)), 1, 2)
      Call REPLA(ZZ$, MKS$(Val(Mid$(ORFA$(II%), 4, 6))), 3, 4)
      Call REPLA(ZZ$, MKI$(NVFB%(II%)), 7, 2)
      Call REPLA(ZZ$, ORFA$(II%), 9, 12)
      Call REPLA(ZZ$, MKI$(CIOF%(II%)), 21, 2)
      JJ& = JJ& + 1
      Call GRAREG(ARCHIORFPRI$, ZZ$, JJ&)
    Next II%
    Call SETULTREG(ARCHIORFPRI$, JJ&)
    Call CIERRARCH(ARCHIORFPRI$)
    Call FILESORT(ARCHIORFPRI$, "", 1, 3, "ASC")
    '
'9.- Re-Carga del Vector de O/F's
    II% = 0
    For JJ& = 1 To ULTREG(ARCHIORFPRI$)
      ZZ$ = REGLEIDO$(ARCHIORFPRI$, JJ&)
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
    REBLA$ = REGBLAN$(archoperfat$)
    JJ& = 0
    '
    For II% = 1 To CAOR%
      NOFA$ = TRIM$(ORFA$(II%))
      SQLX$ = " SELECT EquiAux, CantiOps, ComiePro, CANTPEND, CANTREAL, CANTPROG, MAQUINA, NUMEOPER"
      SQLX$ = SQLX$ + ", COD_INTE, IDSUBOR, TiemOtro, TiemLimp, TiemMHer, TiemDemo, TiemVari, TiemFijo, TiemPrep"
      SQLX$ = SQLX$ + " FROM OPERFAB "
      SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NOFA$ & "' "
      SQLX$ = SQLX$ + " AND StatOper <= " & STATUMA% & " "
      SQLX$ = SQLX$ + " AND CantReal < CantProg "
      If CONSIMULACION% > 0 Then ' SIMULACION
        SQLX$ = SQLX$ + " UNION ALL "
        SQLX$ = SQLX$ + " SELECT EquiAux, CantiOps, ComiePro, CANTPEND, CANTREAL, CANTPROG, MAQUINA, NUMEOPER"
        SQLX$ = SQLX$ + ", COD_INTE, IDSUBOR, TiemOtro, TiemLimp, TiemMHer, TiemDemo, TiemVari, TiemFijo, TiemPrep"
        SQLX$ = SQLX$ + " FROM SIMU_OPERFAB "
        SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NOFA$ & "' "
        SQLX$ = SQLX$ + " AND StatOper <= " & STATUMA% & " "
        SQLX$ = SQLX$ + " AND CantReal < CantProg "
        SQLX$ = SQLX$ + " AND IdSimulOP = " & CONSIMULACION%
      End If
      SQLX$ = SQLX$ + " ORDER BY NumeOper ASC "
      '
      'Dim OPERFAPTMP As ADODB.Recordset
      Set OPERFAPTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
      With OPERFAPTMP
        On Error Resume Next
       ' .MoveFirst
       ' If Err < 1 Then
          Do While Not .EOF
            CAPROG# = XVALO(!CANTPROG)
            CAREA# = XVALO(!CANTREAL)
            CAPEN# = CAPROG# - CAREA#
            CIXII% = XVALO(!COD_INTE)
            NOPEII% = XVALO(!NUMEOPER)
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
                TPREP = XVALO(!TiemPrep): If CAREA# >= 0.05 Then TPREP = 0
                TFIJO = XVALO(!TiemFijo)
                TVARI = XVALO(!TiemVari) * CAPEN# / CAPROG# 'TPROD - TFIJO: If TVARI < 0 Then TVARI = 0
                TDEMO = XVALO(!TiemDemo) * CAPEN# / CAPROG#
                TMANT = XVALO(!TiemMHer) * CAPEN# / CAPROG#
                TLIMP = XVALO(!TiemLimp)
                TOTRO = XVALO(!TiemOtro) * CAPEN# / CAPROG#
                TTOTA = TPREP + TFIJO + TVARI + TDEMO + TMANT + TLIMP + TOTRO
                '
                O$ = REBLA$
                Call REPLA(O$, SAFETEXT$(!idsubor), 1, 12)
                Call REPLA(O$, MKI$(!COD_INTE), 13, 2)
                Call REPLA(O$, MKI$(!NUMEOPER), 15, 2)
                  DOPEXI$ = TRIM$(Str$(XVALO(!NUMEOPER))) + "." + SAFETEXT$(!DESCOPER)
                Call REPLA(O$, DOPEXI$, 17, 48)
                'Call REPLA(O$, !Destino, 65, 8)  ' FALTA COMPLETAR ESTE ELEMENTO
                Call REPLA(O$, SAFETEXT$(!Maquina), 73, 6)
                Call REPLA(O$, MKD$(!CANTPROG), 79, 8)
                Call REPLA(O$, MKD$(!CANTREAL), 87, 8)
                Call REPLA(O$, MKD$(!CANTPEND), 95, 8)
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
                Call REPLA(O$, SAFETEXT$(!EquiAux), 203, 18)
                Call REPLA(O$, MKS$(i&), 221, 4)
                '
                JJ& = JJ& + 1
                Call GRAREG(archoperfat$, O$, JJ&)
                PUNTOPER$(II%) = PUNTOPER$(II%) + MKS$(JJ&)
              End If
            End If
            '
          .MoveNext
          Loop
          '
        'End If
        On Error GoTo 0
      End With
      '
    Next II%
    Call SETULTREG(archoperfat$, JJ&)
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
    FIN& = ULTREG&(archoperfat$)
    FIN1& = FIN&: If FIN1& < 1 Then FIN1& = 1
    '
    For i& = 1 To FIN&
      Call TRACE(20, i&, FIN&)
      O$ = REGLEIDO$(archoperfat$, i&)
      '
      If CVI(Mid$(O$, 135, 2)) = 0 Then
        NORFAX$ = Left$(O$, 12)
        SQLX$ = " SELECT FechLan FROM ORDEFABR "
        SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFAX$ & "' "
        If CONSIMULACION% > 0 Then ' SIMULACION
            SQLX$ = SQLX$ + " UNION ALL "
            SQLX$ = SQLX$ + " SELECT  FechLan FROM SIMU_ORDEFABR "
            SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFAX$ & "' "
            SQLX$ = SQLX$ + " AND IdSimulOR = " & CONSIMULACION%
        End If
        '
        Set ORDEFABTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))       ' , dbOpenSnapshot)
        With ORDEFABTMP
'          On Error Resume Next
'          .MoveLast
'          If Err < 1 Then
            On Error GoTo 0
            FEPROLA% = CVINT%(!FechLan)
            If FEPROLA% > FECOMPRO% Then
              Call REPLA(O$, MKI$(FEPROLA%), 135, 2)
              Call REPLA(O$, MKI$(0), 137, 2)
              Call GRAREG(archoperfat$, O$, i&)
            End If
'          End If
          On Error GoTo 0
        End With
      End If
      '
    Next i&
    '
    'AQUI AJUSTAR CON LA FECHA Y HORA DE LA ULTIMA OPERACION INFORMADA
    '
    For i& = 1 To FIN&
      Call TRACE(20, i&, FIN&)
      O$ = REGLEIDO$(archoperfat$, i&)
      '
      If CVI(Mid$(O$, 135, 2)) = 0 Then
        Call REPLA(O$, MKI$(FECOMPRO%), 135, 2)
        Call REPLA(O$, MKI$(HOCOMPRO%), 137, 2)
        Call GRAREG(archoperfat$, O$, i&)
      End If
      '
    Next i&
    Call CIERRARCH(archoperfat$)
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
    FIN& = ULTREG&(archoperfat$)
    CAOPERF& = ULTREG&(archoperfat$)            ' Cantidad de Operaciones Pendientes
    '
    '\\\OT-08-0978-OD-27/10/08///
    Call SETULTREG&("CRPMENSA", 0)
    '\\\OT-08-0978-OD-FIN///
    For i& = 1 To FIN&
       Call TRACE(20, i&, FIN&)
       O$ = REGLEIDO$(archoperfat$, i&)
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
         Call GRAREG(archoperfat$, O$, i&)
       End If
       '
       Text1(0).TEXT = NOFA$
       CIII% = CVI(Mid$(O$, 13, 2))
       CAPROGG = CVD(Mid$(O$, 79, 8))
       Text1(1).TEXT = TRIM$(CODEXT$(CIII%)) + " - " + TRIM$(DESCRIT$(CIII%)) + " (" + TRIM$(Str$(CAPROGG)) + ")"
       '
       FEFINA% = CVI(Mid$(O$, 135, 2))
       HOFINA% = CVI(Mid$(O$, 137, 2))
       fex = FEFINA%
       Text1(2).TEXT = FECHATEX$(fex)
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
       '\\\OT-08-0978-OD-27/10/08///

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
       O$ = REGLEIDO$(archoperfat$, i&)
       Call REPLA(O$, STHH$, 135, 12)
       Call REPLA(O$, AJUSTI$(MALITSEL$, 48), 151, 48)
       Call GRAREG(archoperfat$, O$, i&)
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
       SQLX$ = " SELECT * FROM ORDEFABR "
       SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFAXI$ & "'"
       Dim CORDEFABR As ADODB.Recordset
       Set CORDEFABR = New ADODB.Recordset
       CORDEFABR.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     '
       With CORDEFABR
         '.FindFirst "IdSubOr = '" & NORFAXI$ & "'"
         If Not (.BOF And .EOF) Then
           '.Edit
           FEFINII% = FEX2
           !FechProyEntre = CVDAT(FEFINII%)
           !FechRecal = Now
           .Update
         End If
       End With
       '
       If CONSIMULACION% > 0 Then ' SIMULACION
        SQLX$ = " SELECT * FROM SIMU_ORDEFABR "
        SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFAXI$ & "'"
        SQLX$ = SQLX$ + " AND IdSimulOR = " & CONSIMULACION%

        Set CORDEFABR = New ADODB.Recordset
        CORDEFABR.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        '
        With CORDEFABR
         If Not (.BOF And .EOF) Then
           FEFINII% = FEX2
           !FechProyEntre = CVDAT(FEFINII%)
           !FechRecal = Now
           .Update
         End If
        End With

    End If

       ' AQUI MARCAR FECHA Y HORA MAS TEMPRANA QUE PUEDEN EMPEZAR LAS
       ' OPERACIONES SIGUIENTES DE LA MISMA ORDEN DE FABRICACION
       COMSIG$ = Mid(STHH$, 5, 4)
       COMISIGUIEN# = 10000 * CDbl(CVI(Left$(COMSIG$, 2))) + CDbl(CVI(Mid$(COMSIG$, 3, 2)))
       For J& = 1 To FIN&
         If J& <> i& Then
           O1$ = REGLEIDO$(archoperfat$, J&)
           If Left$(O1$, 12) = NOFA$ Then
             If CVI(Mid$(O1, 15, 2)) > NUMEOP% Then
               FEHOPROGRA# = 10000 * CDbl(CVI(Mid$(O1$, 135, 2))) + CDbl(CVI(Mid$(O1$, 137, 2)))
               If COMISIGUIEN# > FEHOPROGRA# Then
                 Call REPLA(O1$, COMSIG$, 135, 4)
                 Call GRAREG(archoperfat$, O1$, J&)
               End If
             End If
           End If
         End If
       Next J&
       '
       ' AQUI MARCAR FECHA Y HORA MAS TEMPRANA QUE PUEDEN EMPEZAR LAS
       ' OPERACIONES DE ORDENES DE FABRICACION ASOCIADAS DE NIVEL DE FABRICACION SUPERIOR
       '
       O$ = REGLEIDO$(archoperfat$, i&)
       CODXX% = CVI(Mid$(O$, 13, 2))
       NIVXX% = CANIVEXPLO%(CODXX%)
       '
       For J& = 1 To FIN&
         If J& <> i& Then
           O1$ = REGLEIDO$(archoperfat$, J&)
           If Left$(O1$, 12) <> NOFA$ Then
             If Left$(O1$, 9) = Left$(NOFA$, 9) Then
               CODYY% = CVI(Mid$(O1$, 13, 2))
               If CODYY% <> CODXX% Then
                 If CANIVEXPLO%(CODYY%) > NIVXX% Then
                   FEHOPROGRA# = 10000 * CDbl(CVI(Mid$(O1$, 135, 2))) + CDbl(CVI(Mid$(O1$, 137, 2)))
                   If COMISIGUIEN# > FEHOPROGRA# Then
                     Call REPLA(O1$, COMSIG$, 135, 4)
                     Call GRAREG(archoperfat$, O1$, J&)
                   End If
                 End If
               End If
             End If
           End If
         End If
       Next J&
       '
139 Next i&
    Call CIERRARCH("*.*")
    '
    'GENERA ARCHIVO OPERFAP
    JJ& = 0
    For U& = 1 To ULTREG&(archoperfat$)
      O$ = REGLEIDO$(archoperfat$, U&)
      MALITSEL$ = Mid$(O$, 151, 48)
      For KI% = 1 To 48 Step 6
        MAQUIX$ = TRIM$(Mid$(MALITSEL$, KI%, 6))
        If MAQUIX$ <> "" Then
          Call REPLA(O$, MAQUIX$, 73, 6)
          Call REPLA(O$, Space$(18), 203, 18)
          JJ& = JJ& + 1
          Call GRAREG(ARCHOPERFAP$, O$, JJ&)
        End If
      Next KI%
    Next U&
    Call SETULTREG(ARCHOPERFAP$, JJ&)
    Screen.MousePointer = 1
    '
    Call ABREORFAB
    Dim COPERFAB1 As ADODB.Recordset
    Dim COPERFAB As ADODB.Recordset

    For KU& = 1 To ULTREG&(ARCHOPERFAP$)
      XX$ = REGLEIDO$(ARCHOPERFAP$, KU&)
      IDENOF$ = Left$(XX$, 12)
      NOPEI% = CVI(Mid$(XX$, 15, 2))
      COMIP = CDate(CVDAT(CVI(Mid$(XX$, 135, 2))) & " " & HORATEXNOR$(CVI(Mid$(XX$, 137, 2))))
      FINIP = CDate(CVDAT(CVI(Mid$(XX$, 139, 2))) & " " & HORATEXNOR$(CVI(Mid$(XX$, 141, 2))))
      
      Set COPERFAB = New ADODB.Recordset
      SQLX$ = "SELECT * FROM OPERFAB WHERE IdSubOr = '" & IDENOF$ & "' AND NumeOper = " & NOPEI%
      COPERFAB.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
      '
      With COPERFAB
        '.FindFirst "IdSubOr = '" & IDENOF$ & "' AND NumeOper = " & NOPEI%
        If Not (.EOF And .BOF) Then
          '.Edit
          !ComiePro = COMIP
          !TermiPro = FINIP
          .Update
        End If
      End With
      '
      If CONSIMULACION% > 0 Then ' SIMULACION
        Set COPERFAB1 = New ADODB.Recordset
        SQLX$ = "SELECT * FROM SIMU_OPERFAB WHERE IdSubOr = '" & IDENOF$ & "' AND NumeOper = " & NOPEI%
        SQLX$ = SQLX$ + " AND IdSimulOP = " & CONSIMULACION%
        
        COPERFAB1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        '
        With COPERFAB1
          If Not (.EOF And .BOF) Then
            !ComiePro = COMIP
            !TermiPro = FINIP
            .Update
          End If
        End With
      End If
    Next KU&
    
    '
    Check6.Value = 2
    If CONSIMULACION% > 0 Then ' SIMULACION
      SIMUCARM.CONSIMULACION% = CONSIMULACION%
      Call SIMUCARM.GENORDEFDAT_SIMU(0)
      Call SIMUCARM.GENEOPERFAP_SIMU
    Else
      Call GENORDEFDAT(0)
      Call GENEOPERFAP
    End If
    Call CIERRARCH("*.*")
    Check6.Value = 1
    '
    Call AGRESECOPERPEN
    '
    '\\\OT-08-0978-OD-27/10/08///
    If ULTREG&("CRPMENSA") > 0 Then
      OPP% = COMALTER%("\Se han Encontrado Errores Durante el Proceso de Cálculo\¿ Desea Visualizarlos ?\\Si \No")
      If OPP% = 1 Then
        Call FINAL("*CRPMENSA")
      End If
    End If
    '\\\OT-08-0978-OD-FIN///
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    COMPLECALCU% = 1
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    If IsFormLoaded(SIMUCARM) Then
       Unload SIMUCARM
    End If
    Hide
    '
End Sub



