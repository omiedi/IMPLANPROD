VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form OPTRAJU04 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración General de Movimientos de Stocks"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11580
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
   ScaleHeight     =   8310
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture6 
      Height          =   8350
      Left            =   11160
      ScaleHeight     =   8295
      ScaleWidth      =   1035
      TabIndex        =   94
      Top             =   0
      Visible         =   0   'False
      Width           =   1100
      Begin VB.Image Image3 
         Height          =   8400
         Left            =   600
         Picture         =   "OPTRAJU04.frx":0000
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   92
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame20 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Nota de Entrega"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1180
      Left            =   210
      TabIndex        =   81
      Top             =   5490
      Width           =   5160
      Begin VB.CommandButton Command21 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   4095
         Picture         =   "OPTRAJU04.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   315
         Width           =   855
      End
      Begin VB.TextBox Text18 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   210
         TabIndex        =   84
         TabStop         =   0   'False
         Text            =   " "
         Top             =   720
         Width           =   3750
      End
      Begin VB.TextBox Text17 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2415
         TabIndex        =   83
         Text            =   " "
         Top             =   315
         Width           =   1380
      End
      Begin VB.CommandButton Command20 
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
         TabIndex        =   82
         Top             =   315
         Width           =   175
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
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
         Left            =   0
         TabIndex        =   86
         Top             =   420
         Width           =   2325
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   8520
      Left            =   10680
      ScaleHeight     =   8460
      ScaleWidth      =   885
      TabIndex        =   69
      Top             =   -120
      Width           =   945
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
         Height          =   650
         Left            =   120
         Picture         =   "OPTRAJU04.frx":26A2
         TabIndex        =   90
         Top             =   6600
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command22 
         Caption         =   "C-Cos"
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
         Left            =   120
         Picture         =   "OPTRAJU04.frx":29AC
         Style           =   1  'Graphical
         TabIndex        =   87
         ToolTipText     =   "Base de Datos de Centros de Costos"
         Top             =   3440
         Width           =   650
      End
      Begin VB.CommandButton Command19 
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
         Left            =   120
         Picture         =   "OPTRAJU04.frx":2CB6
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Recuento de Saldo de Stock"
         Top             =   5840
         Width           =   650
      End
      Begin VB.CommandButton Command18 
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
         Left            =   120
         Picture         =   "OPTRAJU04.frx":2FC0
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "Correción Movimientos Deposito 99"
         Top             =   5040
         Width           =   650
      End
      Begin VB.CommandButton command1 
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
         Height          =   650
         Left            =   120
         Picture         =   "OPTRAJU04.frx":32CA
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   240
         Width           =   650
      End
      Begin VB.CommandButton Command2 
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
         Left            =   120
         Picture         =   "OPTRAJU04.frx":3414
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   1040
         Width           =   650
      End
      Begin VB.CommandButton Command7 
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
         Left            =   120
         Picture         =   "OPTRAJU04.frx":371E
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   1840
         Width           =   650
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Depo."
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
         Left            =   120
         Picture         =   "OPTRAJU04.frx":3A28
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Definición de Depósitos y Sucursales"
         Top             =   2640
         Width           =   650
      End
      Begin VB.CommandButton Command14 
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
         Left            =   120
         Picture         =   "OPTRAJU04.frx":3D32
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "Depurar Movimientos Viejos de B.Datos Movimientos de Stock"
         Top             =   4240
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   70
         Top             =   7800
         Width           =   650
         Begin VB.Frame Frame18 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   71
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -315
         Picture         =   "OPTRAJU04.frx":403C
         Top             =   8040
         Width           =   1515
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Fórmula de Consumo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   5460
      TabIndex        =   65
      Top             =   5990
      Width           =   5160
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Lineal"
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
         Left            =   600
         TabIndex        =   67
         Top             =   315
         Value           =   -1  'True
         Width           =   1380
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Multinivel Linealizada"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2280
         TabIndex        =   66
         Top             =   240
         Width           =   2685
      End
   End
   Begin VB.Frame Frame17 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Integracion Remota"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2715
      Left            =   5460
      TabIndex        =   53
      Top             =   3060
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
         Height          =   1455
         Left            =   2940
         TabIndex        =   57
         Top             =   420
         Width           =   1995
      End
      Begin VB.DriveListBox Drive1 
         Enabled         =   0   'False
         Height          =   360
         Left            =   210
         TabIndex        =   56
         Top             =   420
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
         Height          =   990
         Left            =   210
         TabIndex        =   55
         Top             =   900
         Width           =   2535
      End
      Begin VB.CommandButton Command13 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2205
         Picture         =   "OPTRAJU04.frx":5F5E
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   420
         Width           =   540
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta de Acceso:"
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
         Left            =   120
         TabIndex        =   59
         Top             =   1995
         Width           =   1590
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Enabled         =   0   'False
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
         Left            =   240
         TabIndex        =   58
         Top             =   2280
         Width           =   4680
      End
      Begin VB.Line Line2 
         X1              =   210
         X2              =   2760
         Y1              =   1890
         Y2              =   1890
      End
      Begin VB.Line Line3 
         X1              =   2745
         X2              =   2745
         Y1              =   855
         Y2              =   1890
      End
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Descartes - Vale de Scrap"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   210
      TabIndex        =   36
      Top             =   3720
      Width           =   5160
      Begin VB.Frame Frame15 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cód."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4080
         TabIndex        =   44
         Top             =   1080
         Width           =   855
         Begin VB.TextBox Text12 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   190
            TabIndex        =   45
            Text            =   "SC"
            Top             =   240
            Width           =   450
         End
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Numeración"
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
         Left            =   240
         TabIndex        =   41
         Top             =   1080
         Width           =   3750
         Begin VB.OptionButton Option10 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Automatica"
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
            Left            =   2205
            TabIndex        =   43
            Top             =   220
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option9 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Manual"
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
            Left            =   525
            TabIndex        =   42
            Top             =   220
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
         TabIndex        =   40
         Top             =   315
         Width           =   175
      End
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
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
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   210
         TabIndex        =   38
         TabStop         =   0   'False
         Text            =   " "
         Top             =   720
         Width           =   3750
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   4095
         Picture         =   "OPTRAJU04.frx":63D8
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   315
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
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
         Left            =   0
         TabIndex        =   46
         Top             =   420
         Width           =   2325
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Vale de Consumo de Materiales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Left            =   210
      TabIndex        =   25
      Top             =   1870
      Width           =   5160
      Begin VB.CommandButton Command9 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   4095
         Picture         =   "OPTRAJU04.frx":66E2
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   315
         Width           =   855
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   210
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   " "
         Top             =   720
         Width           =   3750
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2415
         TabIndex        =   32
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
         TabIndex        =   31
         Top             =   315
         Width           =   175
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Numeración"
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
         Left            =   210
         TabIndex        =   28
         Top             =   1150
         Width           =   3750
         Begin VB.OptionButton Option8 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Manual"
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
            Left            =   525
            TabIndex        =   30
            Top             =   220
            Width           =   1380
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Automatica"
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
            Left            =   2205
            TabIndex        =   29
            Top             =   220
            Value           =   -1  'True
            Width           =   1380
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cód."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4095
         TabIndex        =   26
         Top             =   1080
         Width           =   855
         Begin VB.TextBox Text7 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   190
            TabIndex        =   27
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   35
         Top             =   420
         Width           =   2325
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Consumo de Componentes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   210
      TabIndex        =   22
      Top             =   6720
      Width           =   10410
      Begin VB.Frame Frame19 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cant.Dec."
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
         Left            =   4700
         TabIndex        =   77
         Top             =   240
         Width           =   1095
         Begin VB.TextBox Text15 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   330
            MaxLength       =   1
            TabIndex        =   78
            Text            =   "1"
            Top             =   410
            Width           =   450
         End
      End
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
         Left            =   5880
         TabIndex        =   47
         Top             =   240
         Width           =   4425
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
            Left            =   4110
            TabIndex        =   51
            Top             =   345
            Width           =   170
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Automático se- gún Existen- cias de Stock"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   105
            TabIndex        =   48
            Top             =   315
            Width           =   1590
         End
         Begin VB.Label Label7 
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
            Height          =   285
            Left            =   1815
            TabIndex        =   52
            Top             =   705
            Width           =   2475
         End
         Begin VB.Label Label6 
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
            Height          =   285
            Left            =   3570
            TabIndex        =   50
            Top             =   345
            Width           =   540
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Concentrar Negativos:"
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
            Left            =   2100
            TabIndex        =   49
            Top             =   210
            Width           =   1170
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Controles Operativos"
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
         Left            =   135
         TabIndex        =   23
         Top             =   240
         Width           =   4455
         Begin VB.CheckBox Check6 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Transf. Con Aprobación"
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
            Left            =   1525
            TabIndex        =   91
            ToolTipText     =   "Al Realizar un Ajuste el Mismo Deberá ser Aprobado Posteriormente"
            Top             =   685
            Width           =   1365
         End
         Begin VB.CheckBox Check4 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Ajuste con Aprobación"
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
            Left            =   3000
            TabIndex        =   89
            ToolTipText     =   "Al Realizar un Ajuste el Mismo Deberá ser Aprobado Posteriormente"
            Top             =   685
            Width           =   1400
         End
         Begin VB.CheckBox Check3 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Transf Valid. Fecha"
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
            Left            =   3000
            TabIndex        =   88
            Top             =   210
            Width           =   1400
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Bloqueo de Negativos"
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
            Left            =   80
            TabIndex        =   68
            Top             =   685
            Width           =   1305
         End
         Begin VB.CheckBox Check5 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Presenta Pantalla de Componentes  a  Consumir"
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
            Left            =   80
            TabIndex        =   24
            Top             =   210
            Width           =   2685
         End
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Ajustes de Stocks"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   210
      TabIndex        =   11
      Top             =   120
      Width           =   5160
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2415
         TabIndex        =   21
         Text            =   " "
         Top             =   240
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
         TabIndex        =   20
         Top             =   240
         Width           =   175
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   210
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   " "
         Top             =   650
         Width           =   3750
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   4095
         Picture         =   "OPTRAJU04.frx":69EC
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   195
         Width           =   855
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Numeración"
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
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   3750
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Manual"
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
            Left            =   525
            TabIndex        =   16
            Top             =   220
            Width           =   1380
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Automatica"
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
            Left            =   2205
            TabIndex        =   15
            Top             =   220
            Value           =   -1  'True
            Width           =   1380
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cód."
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
         TabIndex        =   12
         Top             =   960
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
            TabIndex        =   13
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   17
         Top             =   360
         Width           =   2325
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Transferencias de Stocks"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2805
      Left            =   5460
      TabIndex        =   0
      Top             =   120
      Width           =   5160
      Begin VB.CommandButton Command17 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   4095
         Picture         =   "OPTRAJU04.frx":6CF6
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text14 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   210
         TabIndex        =   62
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1605
         Width           =   3750
      End
      Begin VB.TextBox Text13 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2415
         TabIndex        =   61
         Text            =   " "
         Top             =   1200
         Width           =   1380
      End
      Begin VB.CommandButton Command15 
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
         TabIndex        =   60
         Top             =   1200
         Width           =   175
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cód."
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
         TabIndex        =   9
         Top             =   2040
         Width           =   855
         Begin VB.TextBox Text6 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   190
            TabIndex        =   10
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
         TabIndex        =   6
         Top             =   2040
         Width           =   3750
         Begin VB.OptionButton Option4 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Automatica"
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
            Left            =   2160
            TabIndex        =   8
            Top             =   240
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Manual"
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
            Left            =   480
            TabIndex        =   7
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
         TabIndex        =   4
         Top             =   315
         Width           =   175
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2415
         TabIndex        =   3
         Text            =   " "
         Top             =   315
         Width           =   1380
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   210
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   " "
         Top             =   720
         Width           =   3750
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   4095
         Picture         =   "OPTRAJU04.frx":7000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   315
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario Múltiple:"
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
         Left            =   0
         TabIndex        =   64
         Top             =   1305
         Width           =   2325
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Item Unico:"
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
         Left            =   0
         TabIndex        =   5
         Top             =   420
         Width           =   2325
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "OPTRAJU04.frx":730A
      TabIndex        =   93
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "OPTRAJU04.frx":7384
      Top             =   0
   End
   Begin VB.Menu menu 
      Caption         =   "Menú"
      Begin VB.Menu grabar 
         Caption         =   "Grabar"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu tabla 
      Caption         =   "Tablas"
      Begin VB.Menu dep 
         Caption         =   "Depósitos"
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu cencos 
         Caption         =   "Centro de Costos"
      End
   End
   Begin VB.Menu util 
      Caption         =   "Útilies"
      Begin VB.Menu depuramovimientos 
         Caption         =   "Depurar Movimientos Viejos de B.Datos Movimientos de Stock"
      End
      Begin VB.Menu corrigemovimientosdeposito99 
         Caption         =   "Correción Movimientos Deposito 99"
      End
      Begin VB.Menu recuentosaldostock 
         Caption         =   "Recuento de Saldo de Stock"
      End
   End
End
Attribute VB_Name = "OPTRAJU04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub BARRATRAZA_DragDrop(Source As CONTROL, x As Single, y As Single)

End Sub

Private Sub cencos_Click()
    Call Command22_Click
End Sub

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
    Command14.Enabled = False
    '
    If DERACCE%("DESCAMOV", "Descarga de Movimientos de Stock") < 2 Then
      GoTo 99
    End If
    '
    Dim VESAL$()
    ReDim VESAL$(NUMAS%)
    '
    HOII% = HOY(HOS$)
13  FECHINVE$ = InputBox$("Ingrese la Fecha de Corte", "Depuración de los Movimientos de Stocks")
    If TRIM$(FECHINVE$) = "" Then GoTo 99
    FF% = FECHANUM(FECHINVE$)
    If FF% < 1 Or FF% >= HOY(HOS$) Then
      Call MENSERR(24, "Fecha no Válida")
      GoTo 13
    End If
    '
    If COMALTER%("Está ABSOLUTAMENTE SEGURO de Depurar\los Movimientos de Stock Anteriores al " + FECHINVE$ + "\\No, Conservarlos\Si, Depurar") <> 2 Then
      GoTo 99
    End If
    '
    Screen.MousePointer = 11
    '
    Dim COSU%(62), SUCU$(62), DISTKK%(62)
    NUSU% = 1
    For I& = 1 To ULTREG&("SUCURSAL")
      x$ = REGLEIDO$("SUCURSAL", I&)
      If I& <= 62 Then
        If Asc(Left$(x$, 1)) > 32 Then
          NUSU% = I&
          COSU%(I&) = Asc(Mid$(x$, 64, 1))
        End If
      End If
    Next I&
    Call CIERRARCH("SUCURSAL")
    '
    ' ****************** INICIALIZANDO ARCHIVO DE TRANSFERENCIA
    '
    FIN& = NUMAS% + 1
    For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        SSQ$ = String$(255, 0) + Chr$(1)
        Call GRAREG("SALDEPO", SSQ$, I&)
    Next I&
    Call SETULTREG("SALDEPO", FIN&)
    '
    ' ***************** RE-CALCULANDO STOCK AL  FFS$
    '
    Screen.MousePointer = 11
    Call ABRESTOCKS
    FEC1 = CDbl(CVDAT(FF%))
    '
    SQLX$ = " SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE FechMov <= " & FEC1 & " "
    SQLX$ = SQLX$ & "AND ucase(CodiMovi)<> 'KU' "
    '
    Set GENIVITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    With GENIVITMP
      On Error Resume Next
      .MoveLast
      FIN& = .RecordCount
      .MoveFirst
      If Err Then
        Call MENSERR(24, "No Existen Movimientos Anteriores\a la Fecha: " + FECHATEX$(FF%) + ".")
        On Error GoTo 0
        Screen.MousePointer = 1
        Command14.Enabled = True
        Exit Sub
      End If
      On Error GoTo 0
      '
      CONTI& = 0
      Call ABRESTOCKVIE
      '
      Screen.MousePointer = 11
      Do While .EOF = False
        '
        CONTI& = CONTI& + 1
        Call TRACE(20, CONTI&, FIN&)
        '
        Movisto1.AddNew
        Movisto1.Fields("FechMov") = !FECHMOV
        Movisto1.Fields("Cod_Inte") = !COD_INTE
        Movisto1.Fields("Cod_Exte") = !Cod_Exte
        Movisto1.Fields("DESCRIP") = !Descrip
        Movisto1.Fields("IdenLote") = !IdenLote
        Movisto1.Fields("CodiMovi") = !CodiMovi
        Movisto1.Fields("DoPriCor") = !DoPriCor
        Movisto1.Fields("DoPriLar") = !DOPRILAR
        Movisto1.Fields("DoSecCor") = !DoSecCor
        Movisto1.Fields("DoSecLar") = !DoSecLar
        Movisto1.Fields("NuorItem") = !NuorItem
        Movisto1.Fields("NuPedCli") = !NuPedCli
        Movisto1.Fields("ReferCor") = !ReferCor
        Movisto1.Fields("ValoUnit") = !VALOUNIT
        Movisto1.Fields("MoneVal") = !MoneVal
        Movisto1.Fields("DepoMovi") = !DEPOMOVI
        Movisto1.Fields("Nume_Clie") = !Nume_Clie
        Movisto1.Fields("Nume_Prov") = !Nume_Prov
        Movisto1.Fields("CantIngre") = !CANTINGRE
        Movisto1.Fields("CantSalid") = !CANTSALID
        Movisto1.Fields("CantSaldo") = !CantSaldo
        Movisto1.Fields("FeReMovi") = !FeReMovi
        Movisto1.Fields("NumUsuar") = !NumUsuar
        '
        Movisto1.Update
      .MoveNext
      Loop
      End With
      '
      Screen.MousePointer = 11
      Call ABRESTOCKS
      '
      SQLX$ = " SELECT * FROM MOVISTO "
      SQLX$ = SQLX$ + " WHERE FechMov <= " & FEC1 & " "
      '
      Set CALCTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
      With CALCTMP
      On Error Resume Next
      .MoveLast
      FIN& = .RecordCount
      .MoveFirst
      If Err Then
        Call MENSERR(24, "No Existen Movimientos Anteriores\a la Fecha: " + FECHATEX$(FF%) + ".")
        On Error GoTo 0
        Screen.MousePointer = 1
        Command14.Enabled = True
        Exit Sub
      End If
      On Error GoTo 0
      CONTI& = 0
      '
      'ACA GRABO EL ARCHIVO CON LOS TOTALES
      '(DEBO INCLUIR LOS MOV KU)
103   Do While .EOF = False
        '
        Screen.MousePointer = 11
        CONTI& = CONTI& + 1
        Call TRACE(20, CONTI&, FIN&)
        '
        CI% = !COD_INTE
        dep1% = !DEPOMOVI
        '
105     SUCX% = 1
        For U% = 1 To NUSU%
          If COSU%(U%) = dep1% Then SUCX% = U%: GoTo 110
        Next U%
        Call MENSERR(24, "Error Fatal. El Déposito " & dep1% & " no Existe. Debe Crearlo\y Ejecutar nuevamente el Proceso.")
        GoTo 99
        '
110     CANTI = !CANTINGRE - !CANTSALID
        RESAL& = CI% + 1
        If CI% > 0 Then
          If Len(VESAL$(CI%)) < 4 * NUSU% Then
            VESAL$(CI%) = Left$(VESAL$(CI%) + String$(4 * NUSU%, 0), 4 * NUSU%)
          End If
          SSQ$ = VESAL$(CI%)
          CANUE = CANTI + (Int(10 * CVS(Mid$(SSQ$, 4 * SUCX% - 3, 4)) + 0.5)) / 10
          Mid$(SSQ$, 4 * SUCX% - 3, 4) = MKS$(CANUE)
          VESAL$(CI%) = SSQ$
        End If
      .MoveNext
      Loop
    End With
    '
    DoEvents
    Screen.MousePointer = 1
    '
    ' guardar en saldepo completando el vesal$ con chr$(0)'s
    For CIXO% = 1 To NUMAS%
      If VESAL$(CIXO%) <> "" Then
        CIXS& = CIXO%
        SSQ1$ = REGLEIDO("SALDEPO", CIXS&)
        Call GRAREG("SALDEPO", VESAL$(CIXO%) + String$(256, 0), CIXS&)
      End If
    Next CIXO%
    '
    Call CIERRARCH("SALDEPO")
    '
' ******************** GENERANDO INVENTARIO ANALITICO Y GLOBAL
    '
120 j& = 0: TTI% = 0
    URE& = NUMAS%
    '
    REBLA$ = REGBLAN$("INVEDEP")
    For REG& = 1 To URE&
      Call TRACE(20, REG&, URE&)
      '
      CI% = REG&
      SSQ$ = REGLEIDO$("SALDEPO", REG&)
      '
      SALTOTU = 0
      For U% = 1 To NUSU%
        SALTOTU = SALTOTU + Abs(CVS(Mid$(SSQ$, 4 * U% - 3, 4)))
      Next U%
      '
      If Abs(SALTOTU) >= 0.05 Then
        TOTAL = 0
        For U% = 1 To NUSU%
          SALDE = CVS(Mid$(SSQ$, 4 * U% - 3, 4))
          If Abs(SALDE) >= 0.05 Then
            '
            TOTAL = TOTAL + SALDE
            '
            x$ = REBLA$
            Call REPLA(x$, MKI$(CI%), 1, 2)
            Call REPLA(x$, Chr$(COSU%(U%)), 3, 1)
            Call REPLA(x$, MKS$(SALDE), 5, 4)
            '
            j& = j& + 1
            Call GRAREG("INVEDEP", x$, j&)
            '
          End If
        Next U%
      End If
199 Next REG&
    Call SETULTREG("INVEDEP", j&)
    '
    DoEvents
    '
    'ANTES ESTABA ACA EL PROCESO DE GRABACION DE ACUMULADOS
    '
    'ELIMINACIÓN DE REGISTROS DE LA BASE ORIGINAL
    Call ABRESTOCKS
    FEC1 = CDbl(CVDAT(FF%))
    '
    SQLX$ = " SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE FechMov <= " & FEC1 & " "
    'SQLX$ = SQLX$ + " AND CODIMOVI <> 'KU' "
    Set ELITMP = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
    '
    DoEvents
    With ELITMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, "No Existen Movimientos Anteriores\a la Fecha: " + FECHATEX$(FF%) + ".")
        On Error GoTo 0
        GoTo 99
      End If
      On Error GoTo 0
      On Error Resume Next
      .MoveLast
      Screen.MousePointer = 11
      FIN& = .RecordCount
      JU& = 0
      While .RecordCount > 0  'limpiar los registros
        JU& = JU& + 1
        Call TRACE(20, JU&, FIN&)
        .MoveLast
        .Delete
      Wend
    End With
    '
    'ACA GRABO UN REGISTRO EN LA BASE STOCKS/MOVISTO PARA
    'CADA UNO DE LOS REGISTROS GRABADOS EN EL ARCHIVO INVEDEP
    Screen.MousePointer = 11
    Call ABRESTOCKS
    NORIT& = 0
    For II& = 1 To ULTREG("INVEDEP")
        Call TRACE(20, II&, j&)
        JJ$ = REGLEIDO("INVEDEP", II&)
        With Movisto
          .AddNew
          !FECHMOV = CVDAT(FF%)
          !COD_INTE = CVI(Mid$(JJ$, 1, 2))
            CIXI% = !COD_INTE
          !Cod_Exte = CODEXT$(CIXI%)
          !Descrip = Left$(DESCRIT$(CIXI%), 64)
          !DEPOMOVI = Asc(Mid$(JJ$, 3, 1))
          CANSA = CVS(Mid$(JJ$, 5, 4))
          If CANSA < 0 Then
               !CANTSALID = Abs(CANSA)
               !CANTINGRE = 0
             Else
               !CANTINGRE = CANSA
               !CANTSALID = 0
          End If
          !CodiMovi = "KU" 'MOVIMIENTO ACUMULADOs
          !ReferCor = "Acumulados al " & CVDAT(FF%) & " "
          !IdenLote = ""
          !DoPriCor = "QM" + FECHATEX$(FF%)
          !DOPRILAR = "QM" + FECHATEX$(FF%)
          !DoSecCor = "QM" + FECHATEX$(FF%)
          !DoSecLar = "QM" + FECHATEX$(FF%)
          !NuorItem = NORIT& Mod 10000
          !NuPedCli = 0
          !VALOUNIT = 0
          !MoneVal = 0 '"" cambie a 0 xq es campo integer
          !Nume_Clie = 0
          !Nume_Prov = 0
          !FeReMovi = Now
          !NumUsuar = USNBR%
          .Update
        End With
    Next II&
    '
999 Call CIERRARCH("*.*")
    Call COMUNI("Fin de la Depuración")
    Screen.MousePointer = 1
    '
99 Screen.MousePointer = 1
   Command14.Enabled = True
End Sub
Function NUMAS%()
   '
   Call ABREMASTER
   On Error Resume Next
   Master.MoveLast
   NUMAS% = Master.RecordCount
   If NUMAS% > 0 Then
     If Master!CODINT > NUMAS% Then
       Call MENSERR(24, "La Base de Datos Maestro de Artículos\está Dañada !!! - Consulte Urgente !!!")
       Call FINAL("")
     End If
   End If
   '
End Function

Private Sub Command15_Click()
 '\\\OT-05-0477-WAR-08/07/05///
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text13.TEXT = VDEVU$
        Text14.TEXT = VALACT2$("INDIFRM")
    End If
  '\\\OT-05-0477-WAR-FIN///
End Sub

Private Sub Command16_Click()
    Command16.Enabled = False
    Call CARTADEPOSI
    Command16.Enabled = True
End Sub

Private Sub Command17_Click()
  '\\\OT-05-0477-WAR-08/07/05///
  ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
  Call CONECRUN(ACONEC$ + "/" + TRIM$(Text13.TEXT), "Formularios de Impresion")
  '\\\OT-05-0477-WAR-FIN///
End Sub

Private Sub Command18_Click()
     'Corrige Deposito 99
     Command18.Enabled = False
     Screen.MousePointer = 11
     '
     Dim MoviTemp As DAO.Recordset
     Call ABRESTOCKS
     '
     SQLX$ = "SELECT * FROM MOVISTO WHERE DEPOMOVI = 99"
     '
     Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenDynaset) 'dbOpenDynaset  dbOpenSnapshot
     On Error Resume Next
     '
     MoviTemp.MoveFirst
     '
     If Err > 0 Then
        On Error GoTo 0
        GoTo 99
     End If
     '
     On Error GoTo 0
     '
     Do While Not (MoviTemp.EOF)
         CIXI% = XVALO(MoviTemp!COD_INTE)
         DEPO% = DEPOPRE%(CIXI%)
         MOV$ = SAFETEXT(MoviTemp!DOPRILAR)
         MoviTemp.Edit
         If DEPO% = 0 Then
            COD$ = SAFETEXT(MoviTemp!Cod_Exte)
            COMUNI ("No se Ha Encontrado Deposito para Movimiento " & MOV$ & "\Correspondiente a Codigo " & COD$)
         Else
            COMUNI ("Movimiento " & MOV$ & " fue Cambiado del Deposito 99 al Deposito " & DEPO%)
            MoviTemp!DEPOMOVI = DEPO%
         End If
         MoviTemp.Update
         MoviTemp.MoveNext
     Loop
     MoviTemp.Close
     Screen.MousePointer = 1
    
     '
99   Call COMUNI("Proceso Finalizado")
     Screen.MousePointer = 1
     Command18.Enabled = True
     '
End Sub

Private Sub Command19_Click()
    Command19.Enabled = False
    Screen.MousePointer = 11
    If TIBASE$ <> "MDB" Then
       Call COMUNI("No Disponible en La Presente Versión")
       Screen.MousePointer = 1
       GoTo 99
    End If
    '
    'SOLO PARA MDB
    Call RECUSALSTOK
    Screen.MousePointer = 1
    Call COMUNI("Proceso Finalizado")
99  Command19.Enabled = True
End Sub

Private Sub Command2_Click()
    '
    Call GRACONTROL("", "FORAJUST", Text1.TEXT)
    Call GRACONTROL("", "FORTRAST", Text3.TEXT)
    Call GRACONTROL("", "FORCONSU", TexT8.TEXT)
    Call GRACONTROL("", "FORSCRAP", Text11.TEXT)
    Call GRACONTROL("", "CODAJUST", Text5.TEXT)
    Call GRACONTROL("", "CODTRAST", Text6.TEXT)
    Call GRACONTROL("", "CODCONSU", Text7.TEXT)
    Call GRACONTROL("", "CODSCRAP", Text12.TEXT)
    '\\\OT-5-0447-WAR-08/07/05///
    Call GRACONTROL("", "FORTRSTM", Text13.TEXT)
    '\\\OT-5-0447-WAR-FIN///
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
    '
    Call GRACONTROL("", "DECONSUM", MONUME$)
    Call GRACONTROL("", "DEPONEGA", TRIM$(Label6))
    '
    ' VALIDA FECHA DE TRANSFERENCIA
    MONUME$ = ""
    If Check3.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL$("STOCKS", "VALFETRA", MONUME$)
    '
    '\\\OT-05-0730-WAR-09/11/05///
    BLONEG$ = "NO"
    If Check2.Value = 1 Then
      BLONEG$ = "SI"
    End If
    Call GRACONTROL("STOCKS", "CONEGATI", BLONEG$)
    '\\\OT-05-0730-WAR-FIN///
    '
    DPXX$ = Label11
    If Left$(DPXX$, 2) = Left$(LUDAT$, 2) Then
      DPXX$ = Mid$(DPXX$, 3)
    End If
    Call GRACONTROL("", "INTREMOT", DPXX$)
    '
    CANDECI$ = TRIM(Text15)
    If XVALO(CANDECI$) < 1 Then CANDECI$ = 1
    If XVALO(CANDECI$) > 5 Then CANDECI$ = 5
    Call GRACONTROL("STOCKS", "CANDECI", CANDECI$)
    '
    Call GRACONTROL("", "FORNOTEN", Text17.TEXT)
    '
    'CONTROL DE AJUSTE CON APROBACION
    MONUME$ = ""
    If Check4.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL$("STOCKS", "AJUSAPRO", MONUME$)
    
    'CONTROL DE TRANSFERENCIA CON APROBACION
    MONUME$ = ""
    If Check6.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL$("STOCKS", "TRANSAPRO", MONUME$)
    '
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command20_Click()
  Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text17.TEXT = VDEVU$
        Text18.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command21_Click()
   ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
   Call CONECRUN(ACONEC$ + "/" + TRIM$(Text17.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command22_Click()
    '
    Command22.Enabled = False
    '
    TO1& = ULTREG("TACENCOS") '\\\OT-06-0244-WAR-28/06/06///
    Call TRALOCAL("TACENCOS", "")
10  VTB% = VENTABU("TACENCOS")
    If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!TACENCOS")
        x$ = REGLEIDO$("!TACENCOS", U&)
        cd$ = TRIM$(UCase$(Mid$(x$, 1, 8)))
        Call REPLA(x$, cd$, 1, 8)
        Call GRAREG("!TACENCOS", x$, U&)
        '
        ' Valida que el codigo no sea un valor nulo o menor a dos caracteres
        If Len(cd$) < 2 Then
          Call MENSERR(24, "Código de Centro Costos Menor a dos Caracteres")
          GoTo 10
        End If
        ' Valida que el CC posea descripcion
        DESC$ = TRIM$(Mid$(x$, 9, 40))
        If Len(DESC$) < 1 Then
          Call MENSERR(24, "Centro de Costos sin Descripción")
          GoTo 10
        End If
        ' Valida que el CC no este repetido
        For U2& = 1 To U& - 1
          y$ = REGLEIDO$("!TACENCOS", U2&)
          If TRIM$(Mid$(x$, 1, 8)) = TRIM$(Mid$(y$, 1, 8)) Then
            Call MENSERR(24, "Código de Centro de Costos Repetido")
            GoTo 10
          End If
        Next U2&
        '
      Next U&
      Call TRACENTRAL("TACENCOS", "")
      '\\\OT-06-0244-WAR-28/06/06///
      TO2& = ULTREG&("TACENCOS")
      If TO2& = TO1& + 1 Then 'Si Agrego un registo
        'Le coloco la fecha 'desde' de la Vigencia
        HOII% = HOY(HO$)
        XPLA$ = REGLEIDO$("TACENCOS", TO2&)
        Call REPLA(XPLA$, MKI$(HOII%), 45, 2)
        Call GRAREG("TACENCOS", XPLA$, TO2&)
      End If
      '\\\OT-06-0244-WAR-FIN///
    End If
    '
99  Command22.Enabled = True
    '
End Sub

Private Sub Command23_Click()
    '
    
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
        TexT8.TEXT = VDEVU$
        Text9.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command9_Click()
   ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
   Call CONECRUN(ACONEC$ + "/" + TRIM$(TexT8.TEXT), "Formularios de Impresion")
End Sub

Private Sub corrigemovimientosdeposito99_Click()
    Call Command18_Click
End Sub

Private Sub dep_Click()
   Call Command16_Click
End Sub

Private Sub depuramovimientos_Click()
    Call Command14_Click
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
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture6.Visible = True
      Me.BackColor = &H8000000F
      Me.Frame1.BackColor = &H8000000F
      Me.Frame2.BackColor = &H8000000F
      Me.Frame3.BackColor = &H8000000F
      Me.Frame4.BackColor = &H8000000F
      Me.Frame5.BackColor = &H8000000F
      Me.Frame6.BackColor = &H8000000F
      Me.Frame7.BackColor = &H8000000F
      Me.Frame8.BackColor = &H8000000F
      Me.Frame9.BackColor = &H8000000F
      Me.Frame10.BackColor = &H8000000F
      Me.Frame11.BackColor = &H8000000F
      Me.Frame12.BackColor = &H8000000F
      Me.Frame13.BackColor = &H8000000F
      Me.Frame14.BackColor = &H8000000F
      Me.Frame15.BackColor = &H8000000F
      Me.Frame16.BackColor = &H8000000F
      Me.Frame17.BackColor = &H8000000F
      Me.Frame18.BackColor = &H8000000F
      Me.Frame19.BackColor = &H8000000F
      Me.Frame20.BackColor = &H8000000F
      
      Me.Check1.BackColor = &H8000000F
      Me.Check2.BackColor = &H8000000F
      Me.Check3.BackColor = &H8000000F
      Me.Check4.BackColor = &H8000000F
      Me.Check5.BackColor = &H8000000F
      Me.Check6.BackColor = &H8000000F
      
      
      
      Me.Option1.BackColor = &H8000000F
      Me.Option2.BackColor = &H8000000F
      Me.Option3.BackColor = &H8000000F
      Me.Option4.BackColor = &H8000000F
      Me.Option5.BackColor = &H8000000F
      Me.Option6.BackColor = &H8000000F
      Me.Option7.BackColor = &H8000000F
      Me.Option8.BackColor = &H8000000F
      Me.Option9.BackColor = &H8000000F
      Me.Option10.BackColor = &H8000000F
      MARCOBARRA.Top = 8000
      Set MARCOBARRA.Container = Me.Picture6
      
    End If

    Screen.MousePointer = 11
    Call TRANSLABELS(Name)
    Refresh
    DoEvents: DoEvents: DoEvents
    '
    '
    Screen.MousePointer = 11
    Text1.TEXT = CONTROL$("", "FORAJUST")
    Text2.TEXT = ECOARCH$("INDIFRM", Text1.TEXT)
    Text3.TEXT = CONTROL$("", "FORTRAST")
    Text4.TEXT = ECOARCH$("INDIFRM", Text3.TEXT)
    TexT8.TEXT = CONTROL$("", "FORCONSU")
    Text9.TEXT = ECOARCH$("INDIFRM", TexT8.TEXT)
    Text11.TEXT = CONTROL$("", "FORSCRAP")
    Text10.TEXT = ECOARCH$("INDIFRM", Text11.TEXT)
    Text17.TEXT = CONTROL$("", "FORNOTEN")
    Text18.TEXT = ECOARCH$("INDIFRM", Text17.TEXT)

    Label6 = CONTROL$("", "DEPONEGA")
    '
    '\\\OT-5-0447-WAR-08/07/05///
    Text13.TEXT = CONTROL$("", "FORTRSTM")
    Text14.TEXT = ECOARCH$("INDIFRM", Text13.TEXT)
    '\\\OT-5-0447-WAR-FIN///
    '
    If CONTROL$("", "NUMAJUST") = "MANUAL" Then Option1.Value = True
    If CONTROL$("", "NUMTRAST") = "MANUAL" Then Option3.Value = True
    If CONTROL$("", "NUMVALMA") = "MANUAL" Then Option8.Value = True
    If CONTROL$("", "NUMSCRAP") = "MANUAL" Then Option9.Value = True
    '
    If CONTROL$("", "FORMUCON") = "MULTILIN" Then Option5.Value = True
    If CONTROL$("", "PRECOCIE") = "SI" Then Check5.Value = 1
    If CONTROL$("", "DECONSUM") = "AUTOM" Then Check1.Value = 1
    
    ' VALIDA FECHA DE TRANSFERENCIA
    Check3.Value = 0
    If CONTROL$("STOCKS", "VALFETRA") = "SI" Then Check3.Value = 1
    '
    'VALIDA AJUSTE CON APROBACION
    Check4.Value = 0
    If CONTROL$("STOCKS", "AJUSAPRO") = "SI" Then Check4.Value = 1
    '
    'VALIDA TRANSFERENCIA CON APROBACION
    Check6.Value = 0
    If CONTROL$("STOCKS", "TRANSAPRO") = "SI" Then Check6.Value = 1
    
    '
    '\\\OT-05-0730-WAR-09/11/05///
    If TRIM$(CONTROL$("STOCKS", "CONEGATI")) = "SI" Then Check2.Value = 1
    '\\\OT-05-0730-WAR-FIN///
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
    
    CANDECI$ = CONTROL("STOCKS", "CANDECI")
    If XVALO(CANDECI$) < 1 Then CANDECI$ = "1"
    Text15 = TRIM$(CANDECI$)
    Screen.MousePointer = 1
    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me, Picture1)

End Sub

Private Sub grabar_Click()
   Call Command2_Click
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

Private Sub recuentosaldostock_Click()
    Call Command19_Click
End Sub

Private Sub salir_Click()
   Call Command1_Click
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

Private Sub Text13_DBLCLICK()
  '\\\OT-05-0477-WAR-08/07/05///
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text13.TEXT = VDEVU$
        Text14.TEXT = VALACT2$("INDIFRM")
    End If
  '\\\OT-05-0477-WAR-FIN//
End Sub

Private Sub Text17_DblClick()
 Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text17.TEXT = VDEVU$
        Text18.TEXT = VALACT2$("INDIFRM")
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
        TexT8.TEXT = VDEVU$
        Text9.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

