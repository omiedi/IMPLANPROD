VERSION 5.00
Begin VB.Form OPTRAJUST 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración General de Movimientos de Stocks"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11775
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   11775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   10770
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   73
      Top             =   7245
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   74
         Top             =   105
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Depurar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10770
      Picture         =   "OPTRAJUST.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   72
      ToolTipText     =   "Depurar Movimientos Viejos de B.Datos Movimientos de Stock"
      Top             =   4095
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Transferencias de Stocks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2910
      Left            =   5460
      TabIndex        =   3
      Top             =   210
      Width           =   5160
      Begin VB.CommandButton Command15 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   4095
         Picture         =   "OPTRAJUST.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   1260
         Width           =   855
      End
      Begin VB.TextBox Text14 
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
         Left            =   210
         TabIndex        =   69
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1665
         Width           =   3750
      End
      Begin VB.TextBox Text13 
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
         Left            =   2415
         TabIndex        =   68
         Text            =   " "
         Top             =   1260
         Width           =   1380
      End
      Begin VB.CommandButton Command14 
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
         Left            =   3780
         TabIndex        =   67
         Top             =   1260
         Width           =   175
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4080
         TabIndex        =   12
         Top             =   2145
         Width           =   855
         Begin VB.TextBox Text6 
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
            Left            =   190
            TabIndex        =   13
            Text            =   "TR"
            Top             =   240
            Width           =   450
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Numeración"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   210
         TabIndex        =   9
         Top             =   2145
         Width           =   3750
         Begin VB.OptionButton Option4 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Automatica"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2160
            TabIndex        =   11
            Top             =   240
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Manual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   480
            TabIndex        =   10
            Top             =   240
            Width           =   1380
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
         Height          =   360
         Left            =   3780
         TabIndex        =   7
         Top             =   315
         Width           =   175
      End
      Begin VB.TextBox Text3 
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
         Left            =   2415
         TabIndex        =   6
         Text            =   " "
         Top             =   315
         Width           =   1380
      End
      Begin VB.TextBox Text4 
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
         Left            =   210
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   " "
         Top             =   720
         Width           =   3750
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   4095
         Picture         =   "OPTRAJUST.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   315
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario Múltiple:"
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
         TabIndex        =   71
         Top             =   1365
         Width           =   2325
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Item Unico:"
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
         TabIndex        =   8
         Top             =   420
         Width           =   2325
      End
   End
   Begin VB.Frame Frame17 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Integracion Remota"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3220
      Left            =   5460
      TabIndex        =   60
      Top             =   3150
      Width           =   5160
      Begin VB.FileListBox File1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2040
         Left            =   2940
         TabIndex        =   64
         Top             =   360
         Width           =   1995
      End
      Begin VB.DriveListBox Drive1 
         Enabled         =   0   'False
         Height          =   360
         Left            =   210
         TabIndex        =   63
         Top             =   360
         Width           =   1275
      End
      Begin VB.DirListBox Dir1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1440
         Left            =   210
         TabIndex        =   62
         Top             =   930
         Width           =   2535
      End
      Begin VB.CommandButton Command13 
         Enabled         =   0   'False
         Height          =   330
         Left            =   2205
         Picture         =   "OPTRAJUST.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta de Acceso:"
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
         Left            =   210
         TabIndex        =   66
         Top             =   2520
         Width           =   1590
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Enabled         =   0   'False
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
         TabIndex        =   65
         Top             =   2790
         Width           =   4680
      End
      Begin VB.Line Line2 
         X1              =   210
         X2              =   2760
         Y1              =   2370
         Y2              =   2370
      End
      Begin VB.Line Line3 
         X1              =   2745
         X2              =   2745
         Y1              =   930
         Y2              =   2370
      End
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Descartes - Vale de Scrap"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   210
      TabIndex        =   43
      Top             =   4400
      Width           =   5160
      Begin VB.Frame Frame15 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4080
         TabIndex        =   51
         Top             =   1200
         Width           =   855
         Begin VB.TextBox Text12 
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
            Left            =   190
            TabIndex        =   52
            Text            =   "SC"
            Top             =   240
            Width           =   450
         End
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Numeración"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   48
         Top             =   1200
         Width           =   3750
         Begin VB.OptionButton Option10 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Automatica"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2205
            TabIndex        =   50
            Top             =   240
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option9 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Manual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   525
            TabIndex        =   49
            Top             =   240
            Width           =   1380
         End
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
         Height          =   360
         Left            =   3780
         TabIndex        =   47
         Top             =   315
         Width           =   175
      End
      Begin VB.TextBox Text11 
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
         Left            =   2415
         TabIndex        =   46
         Text            =   " "
         Top             =   315
         Width           =   1380
      End
      Begin VB.TextBox Text10 
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
         Left            =   210
         TabIndex        =   45
         TabStop         =   0   'False
         Text            =   " "
         Top             =   720
         Width           =   3750
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   4095
         Picture         =   "OPTRAJUST.frx":0D98
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   315
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
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
         TabIndex        =   53
         Top             =   420
         Width           =   2325
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Vale de Consumo de Materiales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   210
      TabIndex        =   32
      Top             =   2280
      Width           =   5160
      Begin VB.CommandButton Command9 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   4095
         Picture         =   "OPTRAJUST.frx":10A2
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   315
         Width           =   855
      End
      Begin VB.TextBox Text9 
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
         Left            =   210
         TabIndex        =   40
         TabStop         =   0   'False
         Text            =   " "
         Top             =   720
         Width           =   3750
      End
      Begin VB.TextBox Text8 
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
         Left            =   2415
         TabIndex        =   39
         Text            =   " "
         Top             =   315
         Width           =   1380
      End
      Begin VB.CommandButton Command8 
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
         Left            =   3780
         TabIndex        =   38
         Top             =   315
         Width           =   175
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Numeración"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   210
         TabIndex        =   35
         Top             =   1200
         Width           =   3750
         Begin VB.OptionButton Option8 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Manual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   525
            TabIndex        =   37
            Top             =   240
            Width           =   1380
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Automatica"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2205
            TabIndex        =   36
            Top             =   240
            Value           =   -1  'True
            Width           =   1380
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4095
         TabIndex        =   33
         Top             =   1200
         Width           =   855
         Begin VB.TextBox Text7 
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
            Left            =   190
            TabIndex        =   34
            Text            =   "VM"
            Top             =   240
            Width           =   450
         End
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
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
         TabIndex        =   42
         Top             =   420
         Width           =   2325
      End
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Depos."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10770
      Picture         =   "OPTRAJUST.frx":13AC
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Definición de Depósitos y Sucursales"
      Top             =   3150
      Width           =   855
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Consumo de Componentes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1640
      Left            =   210
      TabIndex        =   25
      Top             =   6480
      Width           =   10410
      Begin VB.Frame Frame16 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Depósito de Consumo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   4935
         TabIndex        =   54
         Top             =   315
         Width           =   5265
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
            Height          =   285
            Left            =   4935
            TabIndex        =   58
            Top             =   350
            Width           =   170
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Automático se- gún Existen- cias de Stock"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   315
            TabIndex        =   55
            Top             =   315
            Width           =   1695
         End
         Begin VB.Label Label7 
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
            Height          =   285
            Left            =   2205
            TabIndex        =   59
            Top             =   700
            Width           =   2900
         End
         Begin VB.Label Label6 
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
            Height          =   285
            Left            =   4410
            TabIndex        =   57
            Top             =   350
            Width           =   540
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Concentrar Negativos:"
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
            Left            =   2310
            TabIndex        =   56
            Top             =   420
            Width           =   2010
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Controles Operativos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   2415
         TabIndex        =   29
         Top             =   315
         Width           =   2430
         Begin VB.CheckBox Check5 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Presenta Pantalla de Componentes  a  Consumir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   315
            TabIndex        =   30
            Top             =   315
            Width           =   1905
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Fórmula de Consumo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   210
         TabIndex        =   26
         Top             =   315
         Width           =   2115
         Begin VB.OptionButton Option5 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Multinivel Linealizada"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   525
            TabIndex        =   27
            Top             =   630
            Width           =   1485
         End
         Begin VB.OptionButton Option6 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Lineal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   525
            TabIndex        =   28
            Top             =   315
            Value           =   -1  'True
            Width           =   1380
         End
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Ajustes de Stocks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1965
      Left            =   210
      TabIndex        =   14
      Top             =   210
      Width           =   5160
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
         Left            =   2415
         TabIndex        =   24
         Text            =   " "
         Top             =   315
         Width           =   1380
      End
      Begin VB.CommandButton Command3 
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
         Left            =   3780
         TabIndex        =   23
         Top             =   315
         Width           =   175
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
         Left            =   210
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   " "
         Top             =   720
         Width           =   3750
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   4095
         Picture         =   "OPTRAJUST.frx":16B6
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   315
         Width           =   855
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Numeración"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   210
         TabIndex        =   17
         Top             =   1200
         Width           =   3750
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Manual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   525
            TabIndex        =   19
            Top             =   240
            Width           =   1380
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Automatica"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2205
            TabIndex        =   18
            Top             =   240
            Value           =   -1  'True
            Width           =   1380
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4095
         TabIndex        =   15
         Top             =   1200
         Width           =   855
         Begin VB.TextBox Text5 
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
            Left            =   190
            TabIndex        =   16
            Text            =   "AJ"
            Top             =   240
            Width           =   450
         End
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
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
         TabIndex        =   20
         Top             =   420
         Width           =   2325
      End
   End
   Begin VB.CommandButton Command7 
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
      Height          =   855
      Left            =   10770
      Picture         =   "OPTRAJUST.frx":19C0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2205
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10770
      Picture         =   "OPTRAJUST.frx":1CCA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1260
      Width           =   855
   End
   Begin VB.CommandButton command1 
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
      Height          =   855
      Left            =   10770
      Picture         =   "OPTRAJUST.frx":1FD4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   315
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   10245
      Picture         =   "OPTRAJUST.frx":211E
      Top             =   7590
      Width           =   2010
   End
End
Attribute VB_Name = "OPTRAJUST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command10_Click()
   ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
   Call CONECRUN(ACONEC$ + "/" + TRIM$(Text11.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command11_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text11.TEXT = VDEVU$
        Text10.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   Call SELECHO("ECODEP")
   DPX1% = Val(VALACT1$("ECODEP"))
   Label6 = VALACT1$("ECODEP")
   Label7 = ""
   If DPX1% > 0 Then
     Label7 = ECOARCH$("ECODEP", Chr$(DPX1%))
   End If
   Command12.Enabled = True
End Sub

Private Sub Command14_Click()
    '\\\OT-5-0403-FG-15/06/05///
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text13.TEXT = VDEVU$
        Text14.TEXT = VALACT2$("INDIFRM")
    End If
    '\\\OT-5-0403-FG-FIN///
End Sub

Private Sub Command15_Click()
    '\\\OT-5-0403-FG-15/06/05///
   ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
   Call CONECRUN(ACONEC$ + "/" + TRIM$(Text13.TEXT), "Formularios de Impresion")
    '\\\OT-5-0403-FG-FIN///
End Sub

Private Sub Command16_Click()
    Command16.Enabled = False
    Call CARTADEPOSI
    Command16.Enabled = True
End Sub

 Sub Command17_Click()
  '\\\OT-06-0097-WAR-17/03/06///
  Command17.Enabled = False
    '
    If DERACCE%("DESCAMOV", "Descarga de Movimientos de Stock") < 2 Then
      GoTo 99
    End If
    'MENSAJE DE ADVERTENCIA PARA QUE SALGAN TODOS!!!
    RTA1% = COMALTER("Para Ejectuar este Proceso\debe Cerrar FLEXOFT\en Todas las Estaciones de Trabajo\\Cancelar\Continuar")
    If RTA1% <> 2 Then GoTo 99
    '
10  Dim VESAL$()
    NMS% = NUMAS%
    ReDim VESAL$(NMS%)
    For UA& = 1 To NMS%
      VESAL$(UA&) = String$(256, 0)
    Next UA&
    '
    HOII% = HOY(HOS$)
13  FECHINVE$ = InputBox$("Ingrese la Fecha de Corte", "Depuración de los Movimientos de Stocks")
    FF% = FECHANUM(FECHINVE$)
    If FF% < 1 Or FF% >= HOY(HOS$) Then
      Call MENSERR(24, "Fecha no Válida")
      GoTo 13
    End If
    '
    Screen.MousePointer = 11
    '
    Call CARDEPOS
    '
    ' ****************** INICIALIZANDO ARCHIVO DE TRANSFERENCIA
    '
    FIN& = NUMAS% + 1
    '
    ' ***************** RE-CALCULANDO STOCK AL  FFS$
    '
    Screen.MousePointer = 11
    '
    Call COPYSTRU("MOVISTO", "MOVISVIE")
    Call COPYSTRU("MOVISTO", "MOVISNUE")
    Call BLOQARCH("MOVISTO")
    CONT1& = ULTREG&("MOVISVIE")
    JNUE& = 0
    '
    FIN& = ULTREG&("MOVISTO")
    For J& = 1 To FIN&
      Call TRACE(20, J&, FIN&)
      x$ = REGLEIDO$("MOVISTO", J&)
      F1% = CVI(Mid$(x$, 1, 2))
      CM1$ = Mid$(x$, 21, 2)
      If F1% <= FF% Then
            'ACUMULO EN VESAL$
            CI% = CVI(Mid$(x$, 3, 2))
            DEP% = Asc(Mid$(x$, 83, 1))
105         SUCX% = 1
            For U% = 1 To NUSU%
              If COSU%(U%) = DEP% Then SUCX% = U%: GoTo 110
            Next U%
            '
110         CANTI = CVD(Mid$(x$, 88, 8)) - CVD(Mid$(x$, 96, 8))
            If CI% > 0 Then
              If CI% <= NUMAS% Then
                VECSALDO$ = VESAL$(CI%)
                CANTIACU = CVS(Mid$(VECSALDO$, 4 * SUCX% - 3, 4)) + CANTI
                Call REPLA(VECSALDO$, MKS$(CANTIACU), 4 * SUCX% - 3, 4)
                VESAL$(CI%) = VECSALDO$
              End If
            End If
            '
            'GRABO REGISTRO EN MOVIMIENTOS VIEJOS
            If CM1$ <> "KU" Then
              CONT1& = CONT1& + 1
              Call GRAREG("MOVISVIE", x$, CONT1&)
            End If
         Else
            'GRABO REGISTRO EN MOVIMIENTOS NUEVOS
            JNUE& = JNUE& + 1
            Call GRAREG("MOVISNUE", x$, JNUE&)
      End If
    Next J&
    '
    Call SETULTREG("MOVISVIE", CONT1&)
    Call CIERRARCH("MOVISVIE")
    Call CIERRARCH("MOVISTO")
115 On Error Resume Next
    Kill LUDAT$ + "MOVISTO.DAT"
    If Err Then
      On Error GoTo 0
      If COMALTER%("Imposible Completar Proceso\Archivo 'MOVISTO' en Uso.\\Reintentar\Cancelar") <> 2 Then
        GoTo 115
      End If
      GoTo 99
    End If
    On Error GoTo 0
    '
    'AQUI GRABAR REGISTROS CORRESPONDIENTES A MOVIMIENTOS ACUMULADOS
    REBLA$ = REGBLAN$("MOVISNUE")
    FIN& = NMS%
    For XI% = 1 To NMS%
      JU& = XI%
      Call TRACE(20, JU&, FIN&)
      For U% = 1 To NUSU%
        SALTOTU = CVS(Mid$(VESAL$(XI%), 4 * U% - 3, 4))
        If Abs(SALTOTU) >= 0.05 Then 'GRABO REGISTRO ACUMULACIÓN
          '
          Y$ = REBLA$
          Call REPLA(Y$, MKI(FF%), 1, 2)
          CIXI% = XI%
          Call REPLA(Y$, MKI(CIXI%), 3, 2)
          DEPO1% = COSU%(U%)
          Call REPLA(Y$, Chr$(DEPO1%), 83, 1)
          CANSA1# = SALTOTU
          Call REPLA(Y$, String$(16, 0), 88, 8)
          If CANSA1# < 0 Then
              Call REPLA(Y$, MKD$(Abs(CANSA1#)), 96, 8) 'SALIDA
             Else
              Call REPLA(Y$, MKD$(CANSA1#), 88, 8) 'ENTRADA
          End If
          '
          Call REPLA(Y$, "KU", 21, 2) 'MOVIMIENTO ACUMULADOs
          FEX = FF%: FEXU$ = FECHATEX$(FEX)
          REFE1$ = "Acumulados al " & FEXU$ & " "
          Call REPLA(Y$, REFE1$, 43, 30)
          Call REPLA(Y$, "", 5, 16) 'LOTE
            DOPRI1$ = "QM-" + Mid$(FEXU$, 7, 2) + Mid$(FEXU$, 4, 2) + Left$(FEXU$, 2)
            DOSEC1$ = DOPRI1$
          Call REPLA(Y$, DOPRI1$, 23, 10)
          Call REPLA(Y$, DOSEC1$, 33, 10)
          HOII% = HOY(HOS$)
          Call REPLA(Y$, MKI$(HOII%), 122, 2)
          Call REPLA(STRI$, MKI$(HORANUM(Time$)), 124, 2)
          Call REPLA(Y$, Chr$(USNBR%), 126, 1)
          JNUE& = JNUE& + 1
          Call GRAREG("MOVISNUE", Y$, JNUE&)
          '
        End If
      Next U%
    Next XI%
    Call SETULTREG("MOVISNUE", JNUE&)
    Call CIERRARCH("MOVISNUE")
    '
    FIN& = ULTREG&("MOVISNUE")
    JJ& = 0
    For KU& = 1 To FIN&
      Call TRACE(20, KU&, FIN&)
      x$ = REGLEIDO$("MOVISNUE", KU&)
      JJ& = JJ& + 1
      Call GRAREG("MOVISTO", x$, JJ&)
    Next KU&
    Call SETULTREG("MOVISTO", JJ&)
    Call CIERRARCH("MOVISTO")
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Screen.MousePointer = 1
    '
    Call COMUNI("Fin deL Proceso de Depuración")
    '
99 Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    Command17.Enabled = True
  '\\\OT-06-0097-WAR-FIN///
End Sub

Private Sub Command2_Click()
    '
    Call GRACONTROL("", "FORAJUST", Text1.TEXT)
    Call GRACONTROL("", "FORTRAST", Text3.TEXT)
    '\\\OT-5-0403-FG-15/06/05///
    Call GRACONTROL("", "FORTRSTM", Text13.TEXT)
    '\\\OT-5-0403-FG-FIN///
    Call GRACONTROL("", "FORCONSU", Text8.TEXT)
    Call GRACONTROL("", "FORSCRAP", Text11.TEXT)
    Call GRACONTROL("", "CODAJUST", Text5.TEXT)
    Call GRACONTROL("", "CODTRAST", Text6.TEXT)
    Call GRACONTROL("", "CODCONSU", Text7.TEXT)
    Call GRACONTROL("", "CODSCRAP", Text12.TEXT)
    '
    MONUME$ = "AUTO": If Option1.Value = True Then MONUME$ = "MANUAL"
    Call GRACONTROL("", "NUMAJUST", MONUME$)
    MONUME$ = "AUTO": If Option3.Value = True Then MONUME$ = "MANUAL"
    Call GRACONTROL("", "NUMTRAST", MONUME$)
    MONUME$ = "AUTO": If Option8.Value = True Then MONUME$ = "MANUAL"
    Call GRACONTROL("", "NUMVALMA", MONUME$)
    MONUME$ = "AUTO": If Option9.Value = True Then MONUME$ = "MANUAL"
    Call GRACONTROL("", "NUMSCRAP", MONUME$)
    '
    MONUME$ = "LINEAL": FORMUCON% = 1
    If Option5.Value = True Then
       MONUME$ = "MULTILIN"
       FORMUCON% = 2
    End If
    Call GRACONTROL("", "FORMUCON", MONUME$)
    '
    MONUME$ = "NO": PRECOCIE% = (-1)
    If Check5.Value = 1 Then
        MONUME$ = "SI"
        PRECOCIE% = 1
    End If
    Call GRACONTROL("", "PRECOCIE", MONUME$)
    '
    MONUME$ = "NO": DECONSUM$ = "AUTOM"
    If Check1.Value = 1 Then
        MONUME$ = "AUTOM"
        DECONSUM$ = "AUTOM"
    End If
    Call GRACONTROL("", "DECONSUM", MONUME$)
    Call GRACONTROL("", "DEPONEGA", TRIM$(Label6))
    '
    DPXX$ = Label11
    If Left$(DPXX$, 2) = Left$(LUDAT$, 2) Then
      DPXX$ = Mid$(DPXX$, 3)
    End If
    Call GRACONTROL("", "INTREMOT", DPXX$)
    '
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command3_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub
'

Private Sub Command4_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text3.TEXT = VDEVU$
        Text4.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
   Call CONECRUN(ACONEC$ + "/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
   Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
   Call CONECRUN(ACONEC$ + "/" + TRIM$(Text3.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command8_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
       Text8.TEXT = VDEVU$
       Text9.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command9_Click()
   ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
   Call CONECRUN(ACONEC$ + "/" + TRIM$(Text8.TEXT), "Formularios de Impresion")
End Sub

Private Sub Dir1_Change()
   File1.Path = Dir1.Path
   Label11 = Dir1.Path
End Sub

Private Sub Drive1_Change()
   On Error Resume Next
   Dir1.Path = Drive1.Drive
   If Err Then
      Drive1.Drive = Left$(CurDir, 2)
      Exit Sub
   End If
End Sub

Private Sub File1_Click()
   Label11.Caption = Dir1.Path
   Label11.Refresh
End Sub

Private Sub Form_Load()
    '
    Screen.MousePointer = 11
    Call TRANSLABELS(Name)
    Refresh
    DoEvents: DoEvents: DoEvents
    '
    Screen.MousePointer = 11
    Text1.TEXT = CONTROL$("", "FORAJUST")
    Text2.TEXT = ECOARCH$("INDIFRM", Text1.TEXT)
    Text3.TEXT = CONTROL$("", "FORTRAST")
    Text4.TEXT = ECOARCH$("INDIFRM", Text3.TEXT)
    '\\\OT-5-0403-FG-15/06/05///
    Text13.TEXT = CONTROL$("", "FORTRSTM")
    Text14.TEXT = ECOARCH$("INDIFRM", Text13.TEXT)
    '\\\OT-5-0403-FG-FIN///
    Text8.TEXT = CONTROL$("", "FORCONSU")
    Text9.TEXT = ECOARCH$("INDIFRM", Text8.TEXT)
    Text11.TEXT = CONTROL$("", "FORSCRAP")
    Text10.TEXT = ECOARCH$("INDIFRM", Text11.TEXT)
    Label6 = CONTROL$("", "DEPONEGA")
    '
    If CONTROL$("", "NUMAJUST") = "MANUAL" Then Option1.Value = True
    If CONTROL$("", "NUMTRAST") = "MANUAL" Then Option3.Value = True
    If CONTROL$("", "NUMVALMA") = "MANUAL" Then Option8.Value = True
    If CONTROL$("", "NUMSCRAP") = "MANUAL" Then Option9.Value = True
    '
    If CONTROL$("", "FORMUCON") = "MULTILIN" Then Option5.Value = True
    If CONTROL$("", "PRECOCIE") = "SI" Then Check5.Value = 1
    If CONTROL$("", "DECONSUM") = "AUTOM" Then Check1.Value = 1
    '
    CMO$ = TRIM$(CONTROL$("", "CODAJUST")): If CMO$ <> "" Then Text5.TEXT = CMO$
    CMO$ = TRIM$(CONTROL$("", "CODTRAST")): If CMO$ <> "" Then Text6.TEXT = CMO$
    CMO$ = TRIM$(CONTROL$("", "CODCONSU")): If CMO$ <> "" Then Text7.TEXT = CMO$
    CMO$ = TRIM$(CONTROL$("", "CODSCRAP")): If CMO$ <> "" Then Text12.TEXT = CMO$
    '
    'DPXX$ = TRIM$(CONTROL("", "INTREMOT"))
    'If DPXX$ <> "" Then
    '  If Mid$(DPXX$, 2, 1) <> ":" Then
    '    DPXX$ = Left$(LUDAT$, 2) + DPXX$
    '  End If
    'End If
    'On Error Resume Next
    'Dir1.Path = DPXX$
    'Label11.Caption = DPXX$
    'Label11.Refresh
    'On Error GoTo 0
    Screen.MousePointer = 1
    '
End Sub

Private Sub Label11_DblClick()
   Label11 = ""
End Sub

Private Sub Label6_Change()
   Label7 = ""
   If Val(Label6) > 0 Then
     Label7 = ECOARCH$("ECODEP", Chr$(Val(Label6)))
   End If
End Sub

Private Sub Text1_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text11_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text11.TEXT = VDEVU$
        Text10.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text2_CLICK()
    Text1.SetFocus
End Sub

Private Sub Text4_CLICK()
    Text2.SetFocus
End Sub

Private Sub Text3_DBLCLICK()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text3.TEXT = VDEVU$
        Text4.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text8_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text8.TEXT = VDEVU$
        Text9.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

