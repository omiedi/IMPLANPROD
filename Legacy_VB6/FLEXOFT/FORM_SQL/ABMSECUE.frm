VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ABMSECUE 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   15840
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   15840
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Height          =   540
      Left            =   12240
      Picture         =   "ABMSECUE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   7320
      Width           =   330
   End
   Begin VB.CommandButton Command9 
      Height          =   540
      Left            =   11880
      Picture         =   "ABMSECUE.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   7320
      Width           =   330
   End
   Begin MSFlexGridLib.MSFlexGrid GRID2 
      Height          =   7185
      Left            =   0
      TabIndex        =   63
      ToolTipText     =   "Doble Click Pasa Directamente a Grilla Sec.Seleccionadas (Permite Duplicación)"
      Top             =   240
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   12674
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Roboto Mono"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox CheckMantenerSeleccion 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mantener Selección"
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
      Left            =   120
      TabIndex        =   61
      Top             =   7515
      Width           =   1995
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   6120
      MaxLength       =   511
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   51
      Top             =   7875
      Width           =   6495
   End
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   7005
      Left            =   6720
      TabIndex        =   10
      Top             =   240
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   12356
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Roboto Mono"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Nueva Lista de Secuencias"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8040
      TabIndex        =   47
      Top             =   7365
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Grabar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7200
      TabIndex        =   46
      Top             =   7365
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Limpiar Grilla"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6360
      TabIndex        =   45
      Top             =   7320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   6280
      TabIndex        =   44
      ToolTipText     =   "Imprime las Filas Selccionadas en Grilla de Secuencias Seleccionadas "
      Top             =   2280
      Width           =   435
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Filtro de Secuencias"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8655
      Left            =   12720
      TabIndex        =   12
      Top             =   0
      Width           =   3135
      Begin VB.CommandButton Command7 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   2840
         TabIndex        =   60
         Top             =   1080
         Width           =   175
      End
      Begin VB.TextBox TxtCodCliente 
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
         Height          =   280
         Left            =   120
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   1635
         Width           =   2880
      End
      Begin VB.TextBox txtCodDaichi 
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
         Height          =   280
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   1080
         Width           =   2660
      End
      Begin VB.TextBox TXTREFERSEC 
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
         Height          =   560
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   240
         Width           =   1965
      End
      Begin VB.CommandButton Command6 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   855
         TabIndex        =   49
         Top             =   240
         Width           =   175
      End
      Begin VB.TextBox Text18 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   120
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   240
         Width           =   675
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2D3C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Index           =   2
         Left            =   1800
         TabIndex        =   41
         Top             =   7755
         Width           =   1275
         Begin VB.CheckBox CheckGeneral 
            BackColor       =   &H00E2D3C0&
            Caption         =   "Test Eléct."
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
            Index           =   24
            Left            =   80
            TabIndex        =   43
            Top             =   440
            Width           =   1250
         End
         Begin VB.CheckBox CheckGeneral 
            BackColor       =   &H00E2D3C0&
            Caption         =   "Armado"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   23
            Left            =   80
            TabIndex        =   42
            Top             =   200
            Width           =   1000
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Especiales"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1830
         Index           =   1
         Left            =   120
         TabIndex        =   34
         Top             =   6765
         Width           =   1575
         Begin VB.CheckBox CheckGeneral 
            BackColor       =   &H00E2D3C0&
            Caption         =   "Goma"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   20
            Left            =   40
            TabIndex        =   40
            Top             =   960
            Width           =   1390
         End
         Begin VB.CheckBox CheckGeneral 
            BackColor       =   &H00E2D3C0&
            Caption         =   "Niquel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   21
            Left            =   40
            TabIndex        =   39
            Top             =   1200
            Width           =   1390
         End
         Begin VB.CheckBox CheckGeneral 
            BackColor       =   &H00E2D3C0&
            Caption         =   "Niquel - Oro"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   22
            Left            =   40
            TabIndex        =   38
            Top             =   1440
            Width           =   1390
         End
         Begin VB.CheckBox CheckGeneral 
            BackColor       =   &H00E2D3C0&
            Caption         =   "Fondo Negro"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   17
            Left            =   40
            TabIndex        =   37
            Top             =   240
            Width           =   1390
         End
         Begin VB.CheckBox CheckGeneral 
            BackColor       =   &H00E2D3C0&
            Caption         =   "Fondo Blanco"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   18
            Left            =   40
            TabIndex        =   36
            Top             =   480
            Width           =   1390
         End
         Begin VB.CheckBox CheckGeneral 
            BackColor       =   &H00E2D3C0&
            Caption         =   "Grafito"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   19
            Left            =   40
            TabIndex        =   35
            Top             =   720
            Width           =   1390
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Corte Placa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   0
         Left            =   1800
         TabIndex        =   30
         Top             =   6765
         Width           =   1275
         Begin VB.CheckBox CheckGeneral 
            BackColor       =   &H00E2D3C0&
            Caption         =   "Router"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   14
            Left            =   40
            TabIndex        =   33
            Top             =   240
            Width           =   1190
         End
         Begin VB.CheckBox CheckGeneral 
            BackColor       =   &H00E2D3C0&
            Caption         =   "V-Scoring"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   15
            Left            =   40
            TabIndex        =   32
            Top             =   480
            Width           =   1190
         End
         Begin VB.CheckBox CheckGeneral 
            BackColor       =   &H00E2D3C0&
            Caption         =   "Guillotina"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   16
            Left            =   40
            TabIndex        =   31
            Top             =   720
            Width           =   1190
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Terminación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   5220
         Width           =   2850
         Begin VB.OptionButton OPTIONGENERAL 
            BackColor       =   &H00E2D3C0&
            Caption         =   "Cobre"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   12
            Left            =   120
            TabIndex        =   29
            Top             =   960
            Width           =   1980
         End
         Begin VB.OptionButton OPTIONGENERAL 
            BackColor       =   &H00E2D3C0&
            Caption         =   "FLUX"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   13
            Left            =   120
            TabIndex        =   28
            Top             =   1200
            Value           =   -1  'True
            Width           =   1500
         End
         Begin VB.OptionButton OPTIONGENERAL 
            BackColor       =   &H00E2D3C0&
            Caption         =   "Estaño Electrolítico"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   10
            Left            =   120
            TabIndex        =   27
            Top             =   480
            Width           =   1980
         End
         Begin VB.OptionButton OPTIONGENERAL 
            BackColor       =   &H00E2D3C0&
            Caption         =   "H.A.L."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   9
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   1500
         End
         Begin VB.OptionButton OPTIONGENERAL 
            BackColor       =   &H00E2D3C0&
            Caption         =   "Estaño Químico"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   11
            Left            =   120
            TabIndex        =   25
            Top             =   720
            Width           =   1500
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Máscaras"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   4365
         Width           =   2850
         Begin VB.OptionButton OPTIONGENERAL 
            BackColor       =   &H00E2D3C0&
            Caption         =   "Componentes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   8
            Left            =   240
            TabIndex        =   23
            Top             =   500
            Width           =   1500
         End
         Begin VB.OptionButton OPTIONGENERAL 
            BackColor       =   &H00E2D3C0&
            Caption         =   "Antisoldante"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   7
            Left            =   240
            TabIndex        =   22
            Top             =   240
            Width           =   1500
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Tipo Fabricación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   3015
         Width           =   2850
         Begin VB.OptionButton OPTIONGENERAL 
            BackColor       =   &H00E2D3C0&
            Caption         =   "Pintura Fotoimageable"
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
            Index           =   5
            Left            =   240
            TabIndex        =   20
            Top             =   480
            Width           =   1740
         End
         Begin VB.OptionButton OPTIONGENERAL 
            BackColor       =   &H00E2D3C0&
            Caption         =   "Fotoproceso"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   4
            Left            =   240
            TabIndex        =   19
            Top             =   240
            Width           =   1500
         End
         Begin VB.OptionButton OPTIONGENERAL 
            BackColor       =   &H00E2D3C0&
            Caption         =   "MULTICAPA"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   6
            Left            =   240
            TabIndex        =   18
            Top             =   1000
            Value           =   -1  'True
            Width           =   1500
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Circuito"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   2025
         Width           =   2850
         Begin VB.OptionButton OPTIONGENERAL 
            BackColor       =   &H00E2D3C0&
            Caption         =   "MULTICAPA"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   240
            TabIndex        =   16
            Top             =   720
            Value           =   -1  'True
            Width           =   1200
         End
         Begin VB.OptionButton OPTIONGENERAL 
            BackColor       =   &H00E2D3C0&
            Caption         =   "PTH"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   240
            TabIndex        =   15
            Top             =   240
            Width           =   1200
         End
         Begin VB.OptionButton OPTIONGENERAL 
            BackColor       =   &H00E2D3C0&
            Caption         =   "SEMI"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   240
            TabIndex        =   14
            Top             =   480
            Width           =   1200
         End
      End
      Begin VB.Label Label71 
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Código Cliente:"
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
         Index           =   3
         Left            =   120
         TabIndex        =   56
         Top             =   1395
         Width           =   1905
      End
      Begin VB.Label Label71 
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Código Dai-Ichi:"
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
         Index           =   2
         Left            =   120
         TabIndex        =   55
         Top             =   840
         Width           =   1425
      End
   End
   Begin VB.Frame Frame1 
      Height          =   840
      Left            =   0
      TabIndex        =   6
      Top             =   7800
      Width           =   5655
      Begin VB.TextBox TXTBUSCAR 
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
         Left            =   120
         TabIndex        =   0
         Top             =   405
         Width           =   2535
      End
      Begin VB.ComboBox CMBFILTRO 
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
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   405
         Width           =   2775
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar:"
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
         Index           =   0
         Left            =   0
         TabIndex        =   8
         Top             =   165
         Width           =   585
      End
      Begin VB.Label Label71 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Campo:"
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
         Index           =   0
         Left            =   2640
         TabIndex        =   7
         Top             =   165
         Width           =   705
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7680
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   1320
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "ABMSECUE.frx":0614
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   120
      OleObjectBlob   =   "ABMSECUE.frx":0848
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid MSF_2 
      Height          =   7185
      Left            =   0
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   12674
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   8421631
      BackColorBkg    =   12632256
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Roboto Mono"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   0
      TabIndex        =   62
      Top             =   0
      Visible         =   0   'False
      Width           =   5655
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label71 
      BackColor       =   &H00D8BD05&
      BackStyle       =   0  'Transparent
      Caption         =   "Maestro de Secuencias:"
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
      Index           =   5
      Left            =   0
      TabIndex        =   59
      Top             =   30
      Width           =   3345
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      BeginProperty Font 
         Name            =   "Roboto Mono"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   58
      Top             =   240
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Label71 
      BackColor       =   &H00D8BD05&
      BackStyle       =   0  'Transparent
      Caption         =   "Secuencias Seleccionadas:"
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
      Index           =   4
      Left            =   6720
      TabIndex        =   57
      Top             =   30
      Width           =   3465
   End
   Begin VB.Label Label71 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D8BD05&
      BackStyle       =   0  'Transparent
      Caption         =   "Observaciones:"
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
      Left            =   6120
      TabIndex        =   52
      Top             =   7635
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   8160
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      Height          =   240
      Left            =   0
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Transacciones"
      Begin VB.Menu mnuNuevaLisSec 
         Caption         =   "Nueva Lista de Secuencias"
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGrabarSecuencia 
         Caption         =   "Grabar Secuencia Activa"
      End
      Begin VB.Menu s3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAltaNuevaSecuencia 
         Caption         =   "Alta Nueva Secuencia"
      End
   End
   Begin VB.Menu mnuArchivoNV 
      Caption         =   "Archivo"
      Visible         =   0   'False
      Begin VB.Menu mnucopiar 
         Caption         =   "Copiar Celda Seleccionada"
         Visible         =   0   'False
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuitar 
         Caption         =   "Quitar"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu MnuEdición 
      Caption         =   "Edición"
      Begin VB.Menu mnuClearGrillaSeleccionados 
         Caption         =   "Limpiar Grilla de Seleccionados"
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMenu3 
         Caption         =   "Revertir Selección"
      End
      Begin VB.Menu mnuOrdSecProd 
         Caption         =   "Ordenar Por Secuencia Productiva (Secuencias Seleccionadas)"
      End
   End
End
Attribute VB_Name = "ABMSECUE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GLOBALFRM_strContenidoCelda$
Dim GLOBALFRM_FILA&
Dim MSFGLOBALFRM As MSFlexGrid
Public MENU1%
Public MENU2%
Public MENU3%
Public NOPINTAR%
Public SeleccionaFilaDobleClick%
Public MODIFIC%

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub BOTEXIT_Click()
  Unload Me
End Sub

Function GeneraStringFiltros$()
   GeneraStringFiltros$ = ""
   cantioption% = OPTIONGENERAL.UBound
   For i% = 1 To cantioption%
      valor% = Abs(OPTIONGENERAL(i%).Value)
      GeneraStringFiltros$ = GeneraStringFiltros$ & valor%
   Next i%
   
   cantiCheck% = CheckGeneral.UBound
   For i% = cantioption% + 1 To cantiCheck%
      valor% = Abs(CheckGeneral(i%).Value)
      GeneraStringFiltros$ = GeneraStringFiltros$ & valor%
   Next i%
   
End Function

Sub AplicarSecuenciasGrabadas(NroConjSec&, StSecuen$, MSF As MSFlexGrid)
     MSF.Rows = 1
     Dim strArray() As String
     On Error Resume Next
     strArray = Split(StSecuen$, " ")
     For i% = 0 To UBound(strArray)
        IDSEC% = XVALO(strArray(i%))
        'LEO UNO A UNO PARA MANTENER EL ORDEN  EN EL CUAL SE HABIA GUARDADO
        SQLX$ = "SELECT * FROM SECTORES WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE ID = " & IDSEC%
        Set RS = READSET(SQLX$)
        J% = J% + 1
        MSF.Rows = J% + 1
        MSF.TextMatrix(J%, 1) = XVALO(RS!ID)
        MSF.TextMatrix(J%, 2) = SAFETEXT$(RS!SECTOR)
        MSF.TextMatrix(J%, 3) = XVALO(RS!SECUENPROD)
        MSF.TextMatrix(J%, 4) = XVALO(RS!SECTORES_ID)
        RS.Close
        Set RS = Nothing
     Next i%
    
   
End Sub

Sub AplicarOpcionesSeleccionadas(NroConjSec&, StFiltros$)
     For i% = 1 To Len(StFiltros$) 'RECORRE TODO EL STRING GUARDADO CON LOS VALORES DE LOS OPTION Y CHECK
         ValLeido% = XVALO(Mid(StFiltros$, i%, 1)) 'LEE EL VALOR DE CADA POSICION
         Descrip$ = "": GRUPO$ = "": T_OBJ$ = "" 'SON VALORES DE RETORNO
         If SecuenciasOptionCheck%(i%, Descrip$, GRUPO$, T_OBJ$) > 0 Then 'esta funcion retorna 1 si encontro info en datasql
            If T_OBJ$ = "OptionButton" Then 'en T_OBJ$ retorna que objeto es check u option
              OPTIONGENERAL(i%) = ValLeido% 'asigna valor al objeto segun corresponda
            ElseIf T_OBJ$ = "CheckBox" Then
              CheckGeneral(i%) = ValLeido%  'asigna valor al objeto segun corresponda
            End If
         End If
     Next i%
End Sub

Private Sub CheckGeneral_Click(Index As Integer)
  MODIFIC% = 1
End Sub

Private Sub CMBFILTRO_Change()
   TXTBUSCAR_Change
End Sub

Private Sub CMBFILTRO_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
     CODNUEVASEC& = XVALO(Text18)
     If CODNUEVASEC& < 1 Then
      Call COMUNI("Falta Nro. de Grupo de Secuencias de Operaciones")
      GoTo 99
     End If
     '
     If GRID.Rows < 2 Then
      Call COMUNI("No Se Han Seleccionado Secuencias")
      GoTo 99
     End If
     '
     'VALIDA QUE NO SE REPITA
     NIVELEXISTENCIA% = 0
     StSecuen$ = GeneraStringSecuencias$(GRID)
     StFiltros$ = GeneraStringFiltros$
     CONDI$ = " StringSecuen = '" & StSecuen$ & "'"
     CONDI$ = CONDI$ + " AND Filtros = '" & StFiltros$ & "'"
     
     If CantRegistros&("SecuenMaster", CONDI$, "NOMESS") > 0 Then
        NIVELEXISTENCIA% = 1
        GRUPSECNRO& = XVALO(VALOBADA("SecuenMaster", "CODSECUEN", CONDI$, "NOMESS"))
        If GRUPSECNRO& = CODNUEVASEC& Then GoTo 50 'SI ES EL MISMO CODIGO SALTEA LAS VALIDACIONES
        
        CONDI$ = CONDI$ + "  AND  CodCliente = '" & TxtCodCliente & "'"
        If CantRegistros&("SecuenMaster", CONDI$, "NOMESS") > 0 Then
           CODCLIENTE$ = VALOBADA("SecuenMaster", "CODCLIENTE", CONDI$, "NOMESS")
           NIVELEXISTENCIA% = 2
        End If
        CONDI$ = CONDI$ + "  AND  CodExtRela = '" & txtCodDaichi & "'"
        If CantRegistros&("SecuenMaster", CONDI$, "NOMESS") > 0 Then
           CODDAICHI$ = TRIM$(VALOBADA("SecuenMaster", "CodExtRela", CONDI$, "NOMESS"))
           If CODDAICHI$ <> "" Then
               NIVELEXISTENCIA% = 3
           End If
        End If
     End If
     TTTX1$ = "Grupo de Secuencias Existente N° Secuencia: " & GRUPSECNRO&
     TTTX2$ = "\Código Cliente: " & CODCLIENTE$
     TTTX3$ = "\Código Dai Ichi: " & CODDAICHI$
     '
     If Len(StSecuen$) >= 1024 Then
        Call COMUNI("Cantidad de Secuencias Superan la Cantidad Máxima")
        GoTo 99
     End If
     If Len(StFiltros$) >= 1024 Then
        Call COMUNI("Cantidad de Filtros Superan la Cantidad Máxima")
        GoTo 99
     End If

     If NIVELEXISTENCIA% = 3 Then
       Call COMUNI(TTTX1$ + TTTX2$ + TTTX3$)
       GoTo 99
     ElseIf NIVELEXISTENCIA% = 2 Then
       Call COMUNI(TTTX1$ + TTTX2$)
       GoTo 99
     ElseIf NIVELEXISTENCIA% = 1 Then
       Call COMUNI(TTTX1$)
       GoTo 99
     End If
     
50   On Error GoTo ERROR_GUARDAR
     FLEXCONN.BeginTrans
     ORDGUAR% = 0
     Call BORRAREGISTROS("SecuenMaster", "CodSecuen = " & CODNUEVASEC&, REGAF&, "NOMESS")
     '
     SQLX$ = "SELECT * FROM SecuenMaster"
     Dim RS As ADODB.Recordset
     Set RS = New ADODB.Recordset
     RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     '
     With RS
        .AddNew   ' YA CONVERTIDO
        On Error Resume Next
        !CODIEMPR = CodiEmp%
        On Error GoTo 0
        !REFERENCIA = Left(TXTREFERSEC, 128)
        !CODSECUEN = CODNUEVASEC&
        !StringSecuen = StSecuen$
        !Filtros = StFiltros$
        !CIRela = CODINT%(Me.txtCodDaichi)
        !CodExtRela = Me.txtCodDaichi
        !CODCLIENTE = Me.TxtCodCliente
        !Observaciones = Left(Text1, 512)
        !MARBORRA = 0
        ORDGUAR% = ORDGUAR% + 1
        !OrdGuardado = ORDGUAR%
        .Update
     End With
     RS.Close
     Set RS = Nothing
   
ERROR_GUARDAR:
    If Err.Number <> 0 Then
       FLEXCONN.RollbackTrans
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       RS.Close
       Set RS = Nothing
       On Error GoTo 0
       GoTo 99
    Else
      FLEXCONN.CommitTrans
      Call COMUNI("Transacción Realizada")
    End If


99 Command2.Enabled = True
End Sub
Private Sub Command3_Click()
   GRID.Rows = 1
   CODNUEVASEC& = XVALO(VALOBADA("SecuenMaster", "MAX(CodSecuen)", "CodSecuen > 0 ", "NOMESS"))
   Text18 = 1 + CODNUEVASEC&
End Sub

Function GeneraStringSecuencias$(MSF As MSFlexGrid)
   Set MSF = GRID
   AcumulaCodSecuencia$ = ""
   For i& = 1 To MSF.Rows - 1
     AcumulaCodSecuencia$ = AcumulaCodSecuencia$ & MSF.TextMatrix(i&, 1) + " "
   Next i&
   GeneraStringSecuencias$ = TRIM$(AcumulaCodSecuencia$)
End Function

Private Sub Command4_Click()
  Command4.Enabled = False
  Dim MSF As MSFlexGrid
  Set MSF = GRID2
  If MSF_2.Visible = True Then
     Set MSF = MSF_2
  End If
  With MSF
  
   For i& = 1 To .Rows - 1
      MARCA$ = .TextMatrix(i&, 0)
      If MARCA$ = "*" Then
        For J& = 1 To GRID.Rows - 1
            If GRID.TextMatrix(J&, 3) = .TextMatrix(i&, 3) Then GoTo 30 ' NEXT
        Next J&
        Call AGREGAR_FILA_OTRO_MSF(MSF, Me.GRID, i&)
        MODIFIC% = 1
      End If
30 Next i&
  End With
  Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    REG& = GRID.RowSel
    If REG& = 0 Then Exit Sub
    If REG& > 0 Then Call SELECCIONARFILA(GRID, REG&)
    Call INSERTYDESPLAZAR(REG&, GRID, 1)
    MODIFIC% = 1
End Sub
Private Sub Command6_Click()
   Command6.Enabled = False
   NROGROPOSECUENCIAS& = LISTAGRUPOSECUENCIASPRODUCCION&
   If NROGROPOSECUENCIAS& > 0 Then
      Text18 = NROGROPOSECUENCIAS&
   End If
   

99 Command6.Enabled = True
   

End Sub
Function LISTAGRUPOSECUENCIASPRODUCCION&()
   CAMPOS$ = "ID/F7;Referencia/A32;Cod.Cliente/A20;Cod.Dai Ichi/A20;CI/A0"
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, CAMPOS$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO)
   FRMZELECHO.Caption = "Listas de Secuencias de Producción Generadas"
    
    Dim obj As New RECORXET
    Set RS1 = obj.RS_ListaGrupoSecuencias()
    Call Carga_MSF_Recordset(RS1, CAMPOS$, FRMZELECHO.MSF2, 1)
    On Error Resume Next
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 10000
50  FRMZELECHO.Show 1

    LISTAGRUPOSECUENCIASPRODUCCION& = XVALO(RESP_ZELE_VECT(1))
   
End Function

Private Sub Command7_Click()
   Command7.Enabled = False
   txtCodDaichi = ListaSeleccionArticulos("", True)
   Command7.Enabled = True


End Sub

Private Sub Command8_Click()

End Sub

Private Sub Command9_Click()
    REG& = GRID.RowSel
    Call INSERTYDESPLAZAR(REG&, GRID, -1)
    MODIFIC% = 1
End Sub

Private Sub Form_Activate()
MENU1% = 0
MENU2% = 0
MENU3% = 0
SeleccionaFilaDobleClick% = 0
   Call refresca_botones(Me)

'FALTA LEVANTAR INFORMACION CUANDO SE SELECCIONA UNA EXISTENTE.
'VER SI COMBIENE QUE EL ORDEN SEA DE 10 EN 10 PARA INTERCALAR NUEVAS SECUENCIAS
'PODER DAR DE ALTA NUEVAS.
'PODER CAMBIAR ORDEN SOLO EN GRID DE DERECHA.


End Sub

Private Sub Form_Load()
    Call APLICACIONSKINS(Me)
    Call MOSTRAR_SECUENCIAS
    CODNUEVASEC& = XVALO(VALOBADA("SecuenMaster", "MAX(CodSecuen)", "CodSecuen > 0 ", "NOMESS"))
    Text18 = 1 + CODNUEVASEC&

End Sub
Sub MOSTRAR_SECUENCIAS()
    If CantRegistros&("SECTORES", "ID > 0 ", "NOMESS") < 1 Then
       Call COMUNI("Falta Importar Sectores")
    End If
    '
    CAMPOS$ = "ID/F4;SECTOR/A48;Sec.Prod./F7;ID/A0;"
    Call CARGA_ENCABEZADO(GRID2, CAMPOS$, Me)
    Call CARGA_ENCABEZADO(GRID, CAMPOS$, Me, 1)
    Call CARGA_ENCABEZADO(MSF_2, CAMPOS$, Me, 1)
    
    Screen.MousePointer = 11
    i& = 0
    GRID2.Rows = 1
    SQLX$ = "SELECT * FROM SECTORES WITH(NOLOCK)"
    SQLX$ = SQLX$ + " ORDER BY ID"
    Set RS = READSET(SQLX$)
    With RS
      Do While Not .EOF
        i& = i& + 1
        GRID2.Rows = i& + 1
        GRID2.TextMatrix(i&, 1) = XVALO(!ID)
        GRID2.TextMatrix(i&, 2) = SAFETEXT$(!SECTOR)
        GRID2.TextMatrix(i&, 3) = XVALO(!SECUENPROD)
        GRID2.TextMatrix(i&, 4) = XVALO(!SECTORES_ID)
       .MoveNext
      Loop
    End With
99  Screen.MousePointer = 1
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If Cancel = 0 Then
     Unload Me
   End If
End Sub

Private Sub Form_Resize()
   On Error Resume Next
'   Me.MSF_2.Width = Me.Width - 300
'   Me.GRID.Width = Me.Width - 300
'   Me.MSF_2.Height = Me.Height - 1600
'   Me.GRID.Height = Me.Height - 1600
   On Error GoTo 0
End Sub

Private Sub GRID_DblClick()
     
    i& = GRID.MouseRow
    J& = GRID.MouseCol
    If i& = 0 Or i& >= GRID.Rows Then Exit Sub
    Call SELECCIONARFILA(GRID, i&)
    REG& = MSF.RowSel
'    '
'    On Error Resume Next
'    GRID.Row = i&
'    GRID.COL = J&
'    On Error Resume Next
'
'    REG& = i&
'    If REG& < 1 Then
'        Exit Sub
'    End If
'    Call COLOREAR_GRILLA_SOLAFILA(Me.GRID, &HFFFF80, REG&)
'    '
'30  SeleccionaFilaDobleClick% = 1
'    If GRID.Rows <= 1 Then Exit Sub
'    GRID.COL = 0
'    RESP_ZELE$ = GRID.TEXT
'    REG& = GRID.MouseRow
'    If REG& < 1 Then Exit Sub
'    For i& = 1 To GRID.Cols - 1
'     RESP_ZELE_VECT(i&) = GRID.TextMatrix(REG&, i&)
'    Next i&
'    Call T_Espera(1 / 10) 'le agrego un tiempo de espera antes que desaparezca por que si no toma el foco en otro lugar
'    Me.Hide

End Sub

Private Sub GRID_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
     Set MSFGLOBALFRM = GRID
     GLOBALFRM_INDIX% = Index
     REG& = MSFGLOBALFRM.MouseRow
     COL% = MSFGLOBALFRM.MouseCol
     Call MENU_CELDA(MSFGLOBALFRM, REG&, COL%)
  End If
  
End Sub
Sub MENU_CELDA(MSF As MSFlexGrid, REG&, COL%)
        If MSF.Rows <= 1 Then Exit Sub
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbYellow, REG&, Int(COL%))   ' LA PINTA DE AMARILLA
        Call INVICTRL
        GLOBALFRM_strContenidoCelda$ = TRIM$(MSF.TextMatrix(REG&, COL%))  'CELDA SELECCIONADA
        GLOBALFRM_FILA& = REG&
        PopupMenu mnuArchivoNV
        Call INVICTRL
        On Error Resume Next
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbWhite, REG&, COL%)  ' LA PINTA DE AMARILLA
        On Error GoTo 0
End Sub

Private Sub GRID2_DblClick()
  i& = GRID2.MouseRow
  J& = GRID2.MouseCol
  If i& = 0 Or i& >= GRID2.Rows Then Exit Sub

  Dim MSF As MSFlexGrid
  Set MSF = GRID2
  If MSF_2.Visible = True Then
     Set MSF = MSF_2
  End If
  
  With MSF
    Call AGREGAR_FILA_OTRO_MSF(MSF, Me.GRID, i&)
    MODIFIC% = 1
  End With
End Sub

Private Sub mnuAltaNuevaSecuencia_Click()
        CAMPOS$ = "Sector/A128;ULT/F0"
        VENTANA.Inicializar = 0
        VENTANA.UtilizaFormula = 1
        VENTANA.HabilitaInsertar = 0
        VENTANA.NoMostrarBotonTilde = 0
        VENTANA.TITULO = "Alta de Nuevo Sector"
        VENTANA.CampEditables = "1"
        VENTANA.Campox = CAMPOS$
        VENTANA.EnumerarRepUltimo(9) = "S"
        VENTANA.Inicializar = 1
30      VENTABNEW.Show 1
        LEIDO$ = ""
        If VENTABNEW.APROBADO% > 0 Then
            LEIDO$ = TRIM$(VENTABNEW.MSF.TextMatrix(1, 1))
        End If
        If LEIDO$ <> "" Then
           OPX% = COMALTER%("Confirma Agregar un Nuevo Sector al Maestro de Secuencias\\Cancelar\Agregar")
           If OPX% < 2 Then GoTo 99
           Call AltaSecuencia(LEIDO$)
        End If
        Call MOSTRAR_SECUENCIAS

99
End Sub

Private Sub mnuClearGrillaSeleccionados_Click()
  GRID.Rows = 1
  MODIFIC% = 1
End Sub

Private Sub mnucopiar_Click()
       Clipboard.Clear
       Clipboard.SetText GLOBALFRM_strContenidoCelda$
End Sub

 Sub INVICTRL()
   mnucopiar.Visible = Not (mnucopiar.Visible)
   mnuQuitar.Visible = Not (mnuQuitar.Visible)
End Sub

Private Sub mnuListSecuencia_Click()

End Sub

Private Sub mnuGrabarSecuencia_Click()
   Call Command2_Click
   MODIFIC% = 0
End Sub

Private Sub mnuNuevaLisSec_Click()
   Call Command3_Click
   If CheckMantenerSeleccion.Value < 1 Then
      Call ColorearMismoColorQuitarAzterizco(GRID2, vbWhite)
   End If

End Sub

Private Sub mnuOrdSecProd_Click()
   Call Command5_Click
   MODIFIC% = 1
End Sub

Private Sub mnuquitar_Click()
       FILX& = GLOBALFRM_FILA&
       '
       IDTRAX = XVALO(MSFGLOBALFRM.TextMatrix(GLOBALFRM_FILA&, 4))
       Call cmdEliminaItem(MSFGLOBALFRM, GLOBALFRM_FILA&)
       MODIFIC% = 1

Exit Sub


99
    
End Sub


Private Sub GRID_Click()
    Static MODO  As Boolean
    i& = GRID.MouseRow
    J& = GRID.MouseCol
    If i& = 0 Or i& >= GRID2.Rows Then Exit Sub

    Call SELECCIONARFILA(GRID, i&)
    REG& = GRID.RowSel

    If (GRID.MouseRow = 0) Then
          
        ' Ordena en forma ascendente
        If MODO Then
            GRID.COL = GRID.MouseCol
            GRID.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            GRID.COL = GRID.MouseCol
            GRID.Sort = 1
            MODO = True
        End If
    ElseIf NOPINTAR% > 0 Then Exit Sub
    End If
    If J& = 0 Then
       Call COLOREAR_GRILLA_SOLAFILA(GRID, vbYellow, GRID.Row, 0, 0, 0)
    End If
End Sub

Private Sub GRID_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If GRID.Rows <= 1 Then Exit Sub
  With GRID
    ' verifica que el GRIDlexgrid contiene filas
    If .MouseRow > 0 And .MouseCol >= 0 Then
        ' le asigna el dato con TextMatrix
        .ToolTipText = VectorToolTipText(.MouseCol)
    Else
        ' No muestra ningún ToolTip
        .ToolTipText = vbNullString
          
    End If
  End With
    
End Sub

Private Sub MC_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then MC.Visible = False
End Sub

Private Sub GRID2_Click()
    i& = GRID2.MouseRow
    J& = GRID2.MouseCol
    If i& = 0 Or i& > GRID2.Rows Then Exit Sub

    Call COLOREAR_GRILLA_SOLAFILA(Me.GRID2, vbYellow, GRID2.Row, 0, 0, 0)
    Exit Sub
    
'    REGORIG& = GRID2.MouseRow
'    J& = GRID2.MouseCol
'    If J& > 0 Then Exit Sub
'    If REGORIG& < 1 Then Exit Sub
'    '
'    IDORIGEN = GRID2.TextMatrix(REGORIG&, 4)
'
'    'VALIDA SI NO EXISTE EL MISMO ID
'    For i& = 1 To GRID.Rows - 1
'         If GRID.TextMatrix(i, 4) = IDORIGEN Then
'            Exit Sub
'         End If
'    Next i&
'    '
'
'    REG& = GRID.Rows
'    GRID.Rows = GRID.Rows + 1
'    GRID.TextMatrix(REG&, 1) = GRID2.TextMatrix(REGORIG&, 1)
'    GRID.TextMatrix(REG&, 2) = GRID2.TextMatrix(REGORIG&, 2)
'    GRID.TextMatrix(REG&, 3) = GRID2.TextMatrix(REGORIG&, 3)
'    GRID.TextMatrix(REG&, 4) = GRID2.TextMatrix(REGORIG&, 4)
End Sub

Private Sub mnuMenu1_Click()
  MENU1% = 1
  Hide
End Sub
Private Sub mnuMenu2_Click()
  MENU2% = 2
  Hide
End Sub
Private Sub mnuMenu3_Click()
   Screen.MousePointer = 11
   Dim MSFF As MSFlexGrid
   Set MSFF = Me.GRID2
   If TRIM$(TXTBUSCAR.TEXT) <> "" Then Set MSFF = Me.MSF_2
   FIN& = MSFF.Rows - 1
   MSFF.Redraw = False
   For i& = 1 To MSFF.Rows - 1
       Call TRACE(20, i&, FIN&)
       Call COLOREAR_GRILLA_SOLAFILA(MSFF, vbYellow, i&)
   Next i&
   MSFF.Redraw = True
   Screen.MousePointer = 1
End Sub


Private Sub MSF_2_Click()
    Static MODO  As Boolean
    i& = MSF_2.MouseRow
    J& = MSF_2.MouseCol

    If (MSF_2.MouseRow = 0) Then
        ' Ordena en forma ascendente
        If MODO Then
            MSF_2.COL = MSF_2.MouseCol
            MSF_2.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            MSF_2.COL = MSF_2.MouseCol
            MSF_2.Sort = 1
            MODO = True
        End If
    End If
    Call COLOREAR_GRILLA_SOLAFILA(MSF_2, vbYellow, MSF_2.Row, 0, 0, 0)
     
End Sub




Private Sub MSF_2_DblClick()
      i& = MSF_2.MouseRow
      J& = MSF_2.MouseCol
      If i& = 0 Or i& >= MSF_2.Rows Then Exit Sub
    
      Dim MSF As MSFlexGrid
      Set MSF = MSF_2
      If MSF_2.Visible = True Then
         Set MSF = MSF_2
      End If
      
      With MSF
        Call AGREGAR_FILA_OTRO_MSF(MSF, Me.GRID, i&)
        MODIFIC% = 1
      End With

'    If NOPINTAR% > 0 Then Exit Sub
'    i& = MSF_2.MouseRow
'    J& = MSF_2.MouseCol
'    If i& = 0 Or i& > MSF_2.Rows Then Exit Sub
'    '
'    On Error Resume Next
'    MSF_2.Row = i&
'    MSF_2.COL = J&
'    On Error Resume Next
'
'    REG& = i&
'    If REG& < 1 Then
'        Exit Sub
'    End If
'    Call COLOREAR_GRILLA_SOLAFILA(Me.MSF_2, &HFFFF80, REG&)

End Sub

Private Sub Option1_Click()

End Sub

Private Sub SelecSecuencia_Click()

End Sub

Private Sub OPTIONGENERAL_Click(Index As Integer)
   MODIFIC% = 1
End Sub

Private Sub Text18_Change()
     If MODIFIC% = 1 Then
        If COMALTER%("Se han Modificado Datos de La Secuencia Nro: " & XVALO(Text18) & " \\Registrar Cambios\Cancelar") = 1 Then
          Call Command2_Click
          MODIFIC% = 0
        End If
     End If
     
     NroConjSec& = XVALO(Text18)
     If NroConjSec& < 1 Then Exit Sub
     Call BLANOPTIONCHECK
     GRID.Rows = 1
     
     If CheckMantenerSeleccion.Value < 1 Then
        Call ColorearMismoColorQuitarAzterizco(GRID2, vbWhite)
     End If
     
     SQLX$ = "SELECT * FROM SecuenMaster WITH(NOLOCK)"
     SQLX$ = SQLX$ + " WHERE CodSecuen = " & NroConjSec&
     Set RS = READSET(SQLX$)
     '
     With RS
        TXTREFERSEC = SAFETEXT$(!REFERENCIA)
        StSecuen$ = SAFETEXT$(!StringSecuen)
        StFiltros$ = SAFETEXT$(!Filtros)
        Me.txtCodDaichi = SAFETEXT$(!CodExtRela)
        Me.TxtCodCliente = SAFETEXT$(!CODCLIENTE)
        Text1 = SAFETEXT$(!Observaciones)
     End With
     RS.Close
     Set RS = Nothing
     '
     Call AplicarSecuenciasGrabadas(NroConjSec&, StSecuen$, GRID)
     Call AplicarOpcionesSeleccionadas(NroConjSec&, StFiltros$)
     MODIFIC% = 0
     
End Sub
Sub BLANOPTIONCHECK()
     'PONE EN FALSE TODOS LOS CHECK Y LOS OPTION DEL SECTOR FILTRO
     CNREG& = CantRegistros&("DATASQL", "LISTA = 'FILTROSEC' AND CAMPO2 <> '' AND CAMPO3 <> '' AND CAMPO4 <> ''", "NOMESS")
     For i% = 1 To CNREG& 'RECORRE TODO EL STRING GUARDADO CON LOS VALORES DE LOS OPTION Y CHECK
         ValLeido% = XVALO(Mid(StFiltros$, i%, 1)) 'LEE EL VALOR DE CADA POSICION
         If SecuenciasOptionCheck%(i%, Descrip$, GRUPO$, T_OBJ$) > 0 Then 'esta funcion retorna 1 si encontro info en datasql
            If T_OBJ$ = "OptionButton" Then 'en T_OBJ$ retorna que objeto es check u option
              OPTIONGENERAL(i%) = 0 'asigna valor al objeto segun corresponda
            ElseIf T_OBJ$ = "CheckBox" Then
              CheckGeneral(i%) = 0  'asigna valor al objeto segun corresponda
            End If
         End If
     Next i%
    


End Sub
Private Sub TXTBUSCAR_Change()
    '
    
    If TRIM$(TXTBUSCAR) = "" Then
      MSF_2.Visible = False
      GRID2.Visible = True
      Exit Sub
    End If
    '
    MSF_2.Rows = 2
    For i& = 1 To MSF_2.Cols - 1
       MSF_2.TextMatrix(1, i&) = ""
    Next i&
    REG& = 0
    MSF_2.Visible = True
    MSF_2.ZOrder 0
    MSF_2.Cols = GRID2.Cols
    MSF_2.Redraw = False
    GRID2.Visible = False
    '
    'RECORRO EL MSF QUE TIENE TODO E IMPRIMO EN msf  LOS QUE MACHEAN
    For i& = 1 To GRID2.Rows - 1
      For J& = 1 To GRID2.Cols - 1
        If (InStr(UCase$(GRID2.TextMatrix(i&, J&)), UCase$(TXTBUSCAR))) Then
          If (TRIM$(CMBFILTRO.TEXT)) = "" Or TRIM$(UCase$(CMBFILTRO)) = TRIM$(UCase$(GRID2.TextMatrix(0, J&))) Then
            REG& = REG& + 1
            MSF_2.Rows = REG& + 1
            For h& = 1 To MSF_2.Cols - 1
              MSF_2.TextMatrix(REG&, h&) = GRID2.TextMatrix(i&, h&)
            Next h&
          End If
          If (TRIM$(CMBFILTRO.TEXT)) = "" Then Exit For
        End If
      Next J&
    Next i&
    '
    MSF_2.Redraw = True
    
    '
End Sub

Private Sub TXTBUSCAR_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      TXTBUSCAR.TEXT = ""
   End If
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub

Private Sub TXTCAMPODESCRIPTIVO_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub TXTCANFILA_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If
End Sub

Sub ACOMODARMSF(MSF1 As MSFlexGrid)
     MSF1.RowHeight(MSF1.Row) = Label2.Height + 100
     MSF1.TextMatrix(MSF1.Row, 2) = Label2
End Sub





Private Sub TxtCodCliente_Change()
     MODIFIC% = 1
End Sub

Private Sub TXTREFERSEC_Change()
    MODIFIC% = 1
End Sub


