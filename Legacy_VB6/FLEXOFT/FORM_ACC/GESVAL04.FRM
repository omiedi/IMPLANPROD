VERSION 5.00
Begin VB.Form GESVAL04 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gestion de Cartera de Valores"
   ClientHeight    =   3795
   ClientLeft      =   525
   ClientTop       =   2790
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture16 
      Height          =   3855
      Left            =   0
      ScaleHeight     =   3795
      ScaleMode       =   0  'User
      ScaleWidth      =   9435
      TabIndex        =   60
      Top             =   0
      Visible         =   0   'False
      Width           =   9495
      Begin VB.Image Image1 
         Height          =   3800
         Left            =   0
         Picture         =   "GESVAL04.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   9500
      End
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00004080&
      Height          =   1065
      Left            =   5355
      ScaleHeight     =   1005
      ScaleWidth      =   4140
      TabIndex        =   40
      Top             =   0
      Width           =   4200
      Begin VB.CommandButton Command19 
         Caption         =   "Check"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   1470
         Picture         =   "GESVAL04.frx":4100
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Control de Cartera de Cheques"
         Top             =   210
         Width           =   600
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Check"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         Picture         =   "GESVAL04.frx":440A
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Control de Cartera de Cheques"
         Top             =   840
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   160
         ScaleHeight     =   75
         ScaleWidth      =   840
         TabIndex        =   46
         Top             =   210
         Width           =   900
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   47
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00004080&
         BorderStyle     =   0  'None
         Caption         =   "Frame9"
         Height          =   960
         Left            =   1050
         TabIndex        =   45
         Top             =   105
         Width           =   280
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00004080&
         BorderStyle     =   0  'None
         Caption         =   "Frame9"
         Height          =   960
         Left            =   0
         TabIndex        =   44
         Top             =   105
         Width           =   150
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   3360
         Picture         =   "GESVAL04.frx":4714
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   210
         Width           =   600
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   2730
         Picture         =   "GESVAL04.frx":485E
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   210
         Width           =   600
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Setup"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   2100
         Picture         =   "GESVAL04.frx":4B68
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Configuración de Funcionamiento"
         Top             =   210
         Width           =   600
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -105
         Picture         =   "GESVAL04.frx":4FAA
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1410
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Otros Movimientos"
      Height          =   2525
      Left            =   5355
      TabIndex        =   33
      Top             =   1155
      Width           =   4005
      Begin VB.PictureBox Picture8 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1260
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   48
         Top             =   1680
         Width           =   2325
         Begin VB.CommandButton Command12 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "GESVAL04.frx":6ECC
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Transferencia de Fondos"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Conciliación Bancaria"
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
            Left            =   840
            TabIndex        =   50
            Top             =   105
            Width           =   1380
         End
      End
      Begin VB.PictureBox Picture9 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   945
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   52
         Top             =   1260
         Width           =   2325
         Begin VB.CommandButton Command16 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   1710
            Picture         =   "GESVAL04.frx":71D6
            Style           =   1  'Graphical
            TabIndex        =   56
            ToolTipText     =   "Entrega de Valores Sin Orden de Pago"
            Top             =   0
            Width           =   560
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
            Height          =   600
            Left            =   0
            Picture         =   "GESVAL04.frx":74E0
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Recepción de Cheques sin Recibo Cobranza"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Movimiento de Cheques"
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
            Left            =   630
            TabIndex        =   54
            Top             =   0
            Width           =   1050
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   34
         Top             =   840
         Width           =   2325
         Begin VB.CommandButton Command10 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "GESVAL04.frx":762A
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Transferencia de Fondos"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Transferencia de Fondos"
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
            Left            =   840
            TabIndex        =   36
            Top             =   0
            Width           =   1380
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   37
         Top             =   420
         Width           =   2325
         Begin VB.CommandButton Command11 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "GESVAL04.frx":7774
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Depósito de Valores"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Depósito Cuen- ta Bancaria"
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
            Left            =   840
            TabIndex        =   39
            Top             =   0
            Width           =   1380
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Consultas - Listados"
      Height          =   3565
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   5160
      Begin VB.OptionButton Option12 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Búsqueda de Cheque por CUIT "
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
         Left            =   525
         TabIndex        =   59
         Top             =   600
         Width           =   3300
      End
      Begin VB.OptionButton Option11 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cheques Pendientes de Acreditación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Left            =   525
         TabIndex        =   58
         Top             =   2280
         Width           =   3900
      End
      Begin VB.OptionButton Option10 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Historico de Cheques "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   525
         TabIndex        =   51
         Top             =   1200
         Width           =   3855
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Resumen Diario de Tesorería"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   525
         TabIndex        =   29
         Top             =   1995
         Width           =   3900
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cheques Entregados por Debitar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   525
         TabIndex        =   18
         Top             =   1740
         Width           =   3900
      End
      Begin VB.CommandButton Command7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   4095
         Picture         =   "GESVAL04.frx":7A7E
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   420
         Width           =   855
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
         Height          =   630
         Left            =   4095
         Picture         =   "GESVAL04.frx":7D88
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2730
         Width           =   855
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Búsqueda de Cheque por Número"
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
         Left            =   525
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   3375
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Salida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   210
         TabIndex        =   2
         Top             =   2670
         Width           =   3690
         Begin VB.OptionButton Option5 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Pantalla"
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
            Left            =   315
            TabIndex        =   4
            Top             =   300
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Impresión"
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
            Left            =   2100
            TabIndex        =   3
            Top             =   300
            Width           =   1275
         End
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Detalle de Valores en Cartera"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   525
         TabIndex        =   1
         Top             =   1470
         Width           =   3900
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         BorderWidth     =   2
         X1              =   0
         X2              =   5145
         Y1              =   1050
         Y2              =   1050
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Movimientos Varios"
      Height          =   2535
      Left            =   420
      TabIndex        =   7
      Top             =   4515
      Visible         =   0   'False
      Width           =   4110
      Begin VB.OptionButton Option7 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Conciliación Bancaria por Fecha Registración"
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
         Height          =   330
         Left            =   105
         TabIndex        =   28
         Top             =   1575
         Width           =   4530
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Posición Diferida"
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
         Height          =   330
         Left            =   105
         TabIndex        =   27
         ToolTipText     =   "Doble-Click para Conciliar"
         Top             =   1260
         Width           =   3900
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1575
         ScaleHeight     =   585
         ScaleWidth      =   1950
         TabIndex        =   11
         Top             =   1575
         Width           =   2010
         Begin VB.CommandButton Command4 
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
            Height          =   600
            Left            =   -20
            Picture         =   "GESVAL04.frx":8092
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Entrega  de Valores"
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
            Left            =   735
            TabIndex        =   13
            Top             =   105
            Width           =   960
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1155
         ScaleHeight     =   585
         ScaleWidth      =   1950
         TabIndex        =   14
         Top             =   1050
         Width           =   2010
         Begin VB.CommandButton Command3 
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
            Height          =   600
            Left            =   -20
            Picture         =   "GESVAL04.frx":839C
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Carga Manual de Cartera"
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
            Left            =   735
            TabIndex        =   16
            Top             =   50
            Width           =   1380
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   1950
         TabIndex        =   8
         Top             =   525
         Width           =   2010
         Begin VB.CommandButton Command5 
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
            Height          =   600
            Left            =   0
            Picture         =   "GESVAL04.frx":86A6
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Cheques Rechazados"
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
            Left            =   735
            TabIndex        =   10
            Top             =   50
            Width           =   1590
         End
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   855
      Left            =   6855
      TabIndex        =   26
      Top             =   420
      Width           =   840
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   855
      Left            =   4515
      TabIndex        =   25
      Top             =   420
      Width           =   940
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Informes de Caja y Bancos"
      Height          =   1905
      Left            =   3885
      TabIndex        =   19
      Top             =   4620
      Width           =   4215
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   420
         TabIndex        =   30
         Top             =   1155
         Width           =   3375
         Begin VB.CommandButton Command8 
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
            Height          =   480
            Left            =   2835
            Picture         =   "GESVAL04.frx":89B0
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   105
            Width           =   520
         End
         Begin VB.OptionButton Option9 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Asiento Resumen de CyB"
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
            Left            =   315
            TabIndex        =   31
            Top             =   250
            Width           =   2640
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   3315
         TabIndex        =   20
         Top             =   420
         Width           =   3375
         Begin VB.CommandButton Command15 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "GESVAL04.frx":8CBA
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Emision de Subdiario de Caja y Bancos"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command17 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   2760
            Picture         =   "GESVAL04.frx":8FC4
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Resumen de Cuenta de Caja y Bancos"
            Top             =   0
            Width           =   560
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
            Height          =   600
            Left            =   2200
            Picture         =   "GESVAL04.frx":92CE
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Subdiario de Caja (Ingresos y Egresos)"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Reportes de Caja y Bancos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   840
            TabIndex        =   24
            Top             =   120
            Width           =   1425
         End
      End
   End
   Begin VB.Shape Shape1 
      Height          =   525
      Left            =   4515
      Top             =   4305
      Width           =   1380
   End
   Begin VB.Menu menu 
      Caption         =   "Menú"
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu consultasylistados 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu busquedadechequepornumero 
         Caption         =   "Búsqueda de Cheque por Número"
      End
      Begin VB.Menu BUSQUEDADECHEQUEPORCUIT 
         Caption         =   "Búsqueda de Cheque por CUIT"
      End
      Begin VB.Menu S0 
         Caption         =   "-"
      End
      Begin VB.Menu historicodecheques 
         Caption         =   "Histórico de Cheques"
         Begin VB.Menu pantalla1 
            Caption         =   "Pantalla"
         End
         Begin VB.Menu impresora1 
            Caption         =   "Impresora"
         End
      End
      Begin VB.Menu chequesentregadosparadebitar 
         Caption         =   "Cheques Entregados para Debitar"
         Begin VB.Menu pantalla2 
            Caption         =   "Pantalla"
         End
         Begin VB.Menu impresora2 
            Caption         =   "Impresora"
         End
      End
      Begin VB.Menu pendientesdeaprobacion 
         Caption         =   "Cheques Pendientes de Acreditación"
         Begin VB.Menu pantalla3 
            Caption         =   "Pantalla"
         End
         Begin VB.Menu impresora3 
            Caption         =   "Impresora"
         End
      End
      Begin VB.Menu detallevaloresencartera 
         Caption         =   "Detalle de Valores en Cartera"
         Begin VB.Menu pantalla5 
            Caption         =   "Pantalla"
         End
         Begin VB.Menu impresora5 
            Caption         =   "Impresora"
         End
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu resumendiariotesoreria 
         Caption         =   "Resumen Diario de Tesorería"
      End
   End
   Begin VB.Menu banco 
      Caption         =   "Bancos"
      Begin VB.Menu depositos 
         Caption         =   "Depósitos "
      End
      Begin VB.Menu conciliacionbancaria 
         Caption         =   "Conciliación Bancaria"
      End
   End
   Begin VB.Menu utiles 
      Caption         =   "Útiles"
      Begin VB.Menu transferenciadefondos 
         Caption         =   "Transferencias de Fondos"
      End
      Begin VB.Menu movmientodecheques 
         Caption         =   "Movimientos de Cheques"
         Begin VB.Menu recepciondechequessinrecibo 
            Caption         =   "Recepción de Cheques sin Recibo Cobranza"
         End
         Begin VB.Menu entregadevaloressinop 
            Caption         =   "Entrega de Valores Sin Orden de Pago"
         End
      End
      Begin VB.Menu ss 
         Caption         =   "-"
      End
      Begin VB.Menu controldecarteradecheques 
         Caption         =   "Control de Cartera de Cheques"
      End
   End
   Begin VB.Menu configuracion 
      Caption         =   "Configuración"
      Begin VB.Menu config 
         Caption         =   "Configuración"
      End
   End
End
Attribute VB_Name = "GESVAL04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BUSQUEDADECHEQUEPORCUIT_Click()
    Option12.Value = True
    Call Command7_Click

End Sub

Private Sub busquedadechequepornumero_Click()
    Option6.Value = True
    Call Command7_Click
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Hide
End Sub

Sub Command10_Click()
   Command10.Enabled = False
   OPX% = ALTERNA%("Procesar Transferencia\Listado de Transferencia")
   If OPX% = 1 Then
        Call TRANFONDOS
   ElseIf OPX% = 2 Then
        Call LISTRAFON
   Else: GoTo 10
   End If
10 Command10.Enabled = True
End Sub

Sub Command11_Click()
   Command1.Enabled = False
   '\\\OT-05-0495-WAR-11/07/05///
   OE% = COMALTER("Opciones de Trabajo:\\Efectuar Déposito\Anular Déposito\Corrige Imputacion")
   If OE% = 1 Then
       Call ASIDECHE 'NUEVO DEPOSITO
     ElseIf OE% = 2 Then
      If DERACCE%("ANUDEBA", "Anulación de Deposito Bancario") >= 2 Then
        Call ANUDECHE 'ANULACION
      End If
     ElseIf OE% = 3 Then
        Call CORRIGEIMPU ' CORRIGE IMPUTACION
   End If
   '\\\OT-05-0495-WAR-FIN///
99 Command1.Enabled = True
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   '
   VDEF$ = Space$(12) + String$(16, 0)
   Call REPLA(VDEF$, MKI$(COMESANT%), 13, 2)
   FEXI% = DIANTE(COMESACT%)
   Call REPLA(VDEF$, MKI$(FEXI%), 15, 2)
   FEXI% = FECHANUM("01/01/80")
   Call REPLA(VDEF$, MKI$(FEXI%), 25, 2)
   Call REPLA(VDEF$, MKI$(FINMESACT%), 27, 2)
   '
10 OBX$ = OBJPLA$("CONCIBANCO", VDEF$)
   If OBX$ = "" Then GoTo 99
   '
   VDEF$ = OBX$
   CUEBA$ = Left$(OBX$, 12)
   CUEII% = CUEINT%(CUEBA$)
   If CUEII% < 1 Then GoTo 10
   DESTICHE$ = ECOARCH$("CUEBANC", CUEBA$)
   If DESTICHE$ = "" Then
     Call MENSERR(24, "Cuenta " + CUEBA$ + " no Definida\como Cuenta Bancaria")
     GoTo 10
   End If
   '
   DEXTRACT% = CVI(Mid$(OBX$, 13, 2))
   HAXTRACT% = CVI(Mid$(OBX$, 15, 2))
   If HAXTRACT% < 1 Or DEXTRACT% > HAXTRACT% Then
     Call MENSERR(24, "Rango de Fechas Extracto no Procesable")
     GoTo 10
   End If
   '
   SALICON# = CVD(Mid$(OBX$, 17, 8))
   '
   DESPEN% = CVI(Mid$(OBX$, 25, 2))
   HASPEN% = CVI(Mid$(OBX$, 27, 2))
   If HASPEN% < 1 Or DESPEN% > HASPEN% Then
     Call MENSERR(24, "Rango de Fechas Movimientos Pendientes no Procesable")
     GoTo 10
   End If
   '
   FEX = DEXTRACT%: CONCIBAN.Label8 = FECHATEX$(FEX)
   FEX = HAXTRACT%: CONCIBAN.Label6 = FECHATEX$(FEX)
   CONCIBAN.Label3 = FORMATNUM$(SALICON#, "F12.2")
   '
   Call BLOQARCH("LIMAYOAC")
   Screen.MousePointer = 11
   '
   FEINI = CDbl(CVDAT(DESPEN%))
   FEFIN = CDbl(CVDAT(HASPEN%))
   '
   Call ABREPLANCUEN
   Screen.MousePointer = 11
   SQLX$ = "SELECT * FROM DetaAsien "
   SQLX$ = SQLX$ + "WHERE (CueCoInt=" & CUEII%
   SQLX$ = SQLX$ + " AND SucuAsien = " & SUCUI%
   SQLX$ = SQLX$ + " AND CDBL(Fechasien) >= " & FEINI
   SQLX$ = SQLX$ + " AND CDBL(Fechasien) <= " & FEFIN
   SQLX$ = SQLX$ + " AND statpase > 0 ) "
   SQLX$ = SQLX$ + " AND (MarConsi = '' or ISNULL(MarConsi) = True or IsEmpty(MarConsi) = True) "
   SQLX$ = SQLX$ + "ORDER BY FECHASIEN, ORFEASIEN ASC;"
   Set PASEASIEN = Contable.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
   '
   JJ& = 0
   TRANSPOR# = 0
20 With PASEASIEN
     On Error Resume Next
     .MoveFirst
     If Err Then GoTo 30
     On Error GoTo 0
     Do While Not .EOF
       NUAS$ = !NUIDASIEN
       TIASI% = 0
       Encasien.FindFirst "NUIDASIEN = '" & NUAS$ & "'"
       If Encasien.NoMatch = False Then
         TIASI% = Encasien!TIPOASIEN
       End If
       If TIASI% = 0 Then  ' DESCARTA ASIENTOS RESUMEN
         Z$ = REGBLAN$("LIMAYOAC")
         Call REPLA(Z$, MKI$(CVINT(!FECHASIEN)), 1, 2)
         Call REPLA(Z$, !NUIDASIEN, 3, 12)
         Call REPLA(Z$, MKI$(CVINT(!FECHASIEN)), 15, 2)
           DEPASE$ = !REFEPASE
         Call REPLA(Z$, DEPASE$, 17, 32)
           IDEBE# = !IDEBEPASE
         Call REPLA(Z$, MKD$(IDEBE#), 49, 8)
           IHABE# = !IHABEPASE
         Call REPLA(Z$, MKD$(IHABE#), 57, 8)
         Call REPLA(Z$, "DIAR", 73, 8)
         If IDEBE# >= 0.005 Or IHABE# >= 0.005 Then
           JJ& = JJ& + 1
           Call GRAREG("LIMAYOAC", Z$, JJ&)
         End If
       End If
     .MoveNext
     Loop
   End With
   '
30 On Error GoTo 0
   Call APERBASDAT("CABAN")
   Screen.MousePointer = 11
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "where (CueContaI = " & CUEII% & " "
   SQLX$ = SQLX$ + " AND (MarConsi = '' or isnull(MarConsi) = True OR IsEmpty(MarConsi) = True) "
   SQLX$ = SQLX$ + " AND CDBL(FechConta) >= " & FEINI
   SQLX$ = SQLX$ + " AND CDBL(FechConta) <= " & FEFIN & ") "
   SQLX$ = SQLX$ + " ORDER BY FECHCONTA ASC, FEHOREAL ASC;"
   '
   Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   With CABATEMP
     On Error Resume Next
     .MoveFirst
     If Err Then
        GoTo 40
     End If
     On Error GoTo 0
     Do While Not .EOF
         IDEBE# = !IDEBECOMP
         IHABE# = !IHABECOMP
       If Int(CDbl(!FechConta)) < FEINI Then
           TRANSPOR# = TRANSPOR# + IDEBE# - IHABE#
         ElseIf Int(CDbl(!FechConta)) <= FEFIN Then
           Z$ = REGBLAN$("LIMAYOAC")
           Call REPLA(Z$, MKI$(CVINT(!FechConta)), 1, 2)
           Call REPLA(Z$, !CodiCom_Cor, 3, 12)
           Call REPLA(Z$, MKI$(CVINT(!FechAcred)), 15, 2)
             DEPASE$ = Left$(!Descrimov, 32)
           Call REPLA(Z$, DEPASE$, 17, 32)
           Call REPLA(Z$, MKD$(IDEBE#), 49, 8)
           Call REPLA(Z$, MKD$(IHABE#), 57, 8)
           Call REPLA(Z$, "CABAN", 73, 8)
           If IDEBE# >= 0.005 Or IHABE# >= 0.005 Then
             JJ& = JJ& + 1
             Call GRAREG("LIMAYOAC", Z$, JJ&)
           End If
       End If
     .MoveNext
     Loop
   End With
   '
40 Call SETULTREG("LIMAYOAC", JJ&)
   Call FILESORT("LIMAYOAC", "", 1, 3, "ASC")
   '
   Screen.MousePointer = 11
   CONCIBAN.List2.Clear
   Z$ = Space$(128)
   For U& = 1 To ULTREG("LIMAYOAC")
     X$ = REGLEIDO$("LIMAYOAC", U&)
     Z$ = Space$(128)
       FEX = CVI(Left$(X$, 2))
       FEXA$ = FECHATEX$(FEX)
     Call REPLA(Z$, FEXA$, 1, 8)
       COCOM$ = TRIM$(Mid$(X$, 3, 12))
     Call REPLA(Z$, COCOM$, 11, 15)
       FEX = CVI(Mid$(X$, 15, 2))
       FEXA$ = FECHATEX$(FEX)
     Call REPLA(Z$, FEXA$, 27, 8)
       DEPASE$ = TRIM$(Mid$(X$, 17, 32))
     Call REPLA(Z$, DEPASE$, 38, 28)
       IDEBE# = CVD(Mid$(X$, 49, 8))
       IHABE# = CVD(Mid$(X$, 57, 8))
     If IDEBE# > 0 Then
         Call REPLA(Z$, CSTRING$(MKD$(IDEBE#), 3, 14, 2, 2), 66, 14)
       Else
         Call REPLA(Z$, CSTRING$(MKD$(IHABE#), 3, 14, 2, 2), 80, 14)
     End If
       TRANSPOR# = TRANSPOR# + IDEBE# - IHABE#
     Call REPLA(Z$, CSTRING$(MKD$(TRANSPOR#), 3, 14, 2, 2), 94, 14)
       ORIMO$ = TRIM$(Mid$(X$, 73, 8))
     Call REPLA(Z$, ORIMO$, 121, 8)
     CONCIBAN.List2.AddItem Z$
     DoEvents
   Next U&
   Call CIERRARCH("LIMAYOAC")
   '
   '
   ' CARGA LISTA DE MOVIMIENTOS CONCILIADOS
   FEINI = CDbl(CVDAT(DEXTRACT%))
   FEFIN = CDbl(CVDAT(HAXTRACT%))
   '
   Call ABREPLANCUEN
   Screen.MousePointer = 11
   SQLX$ = "SELECT * FROM DetaAsien "
   SQLX$ = SQLX$ + "WHERE (CueCoInt=" & CUEII%
   SQLX$ = SQLX$ + " AND SucuAsien = " & SUCUI%
   SQLX$ = SQLX$ + " AND CDBL(Fechasien) >= " & FEINI
   SQLX$ = SQLX$ + " AND CDBL(Fechasien) <= " & FEFIN
   SQLX$ = SQLX$ + " AND statpase >0 ) "
   SQLX$ = SQLX$ + " AND MarConsi = '*' "
   SQLX$ = SQLX$ + "ORDER BY FECHASIEN, ORFEASIEN ASC;"
   Set PASEASIEN = Contable.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
   '
   JJ& = 0
   TRANSPOR# = 0
70 With PASEASIEN
     On Error Resume Next
     .MoveFirst
     If Err Then GoTo 80
     On Error GoTo 0
     Do While Not .EOF
       NUAS$ = !NUIDASIEN
       TIASI% = 0
       Encasien.FindFirst "NUIDASIEN = '" & NUAS$ & "'"
       If Encasien.NoMatch = False Then
         TIASI% = Encasien!TIPOASIEN
       End If
       If TIASI% = 0 Then  ' DESCARTA ASIENTOS RESUMEN
         Z$ = REGBLAN$("LIMAYOAC")
         Call REPLA(Z$, MKI$(CVINT(!FECHASIEN)), 1, 2)
         Call REPLA(Z$, !NUIDASIEN, 3, 12)
         Call REPLA(Z$, MKI$(CVINT(!FECHASIEN)), 15, 2)
           DEPASE$ = !REFEPASE
         Call REPLA(Z$, DEPASE$, 17, 32)
           IDEBE# = !IDEBEPASE
         Call REPLA(Z$, MKD$(IDEBE#), 49, 8)
           IHABE# = !IHABEPASE
         Call REPLA(Z$, MKD$(IHABE#), 57, 8)
         Call REPLA(Z$, "DIAR", 73, 6)
         Call REPLA(Z$, "*", 80, 1)
         If IDEBE# >= 0.005 Or IHABE# >= 0.005 Then
           JJ& = JJ& + 1
           Call GRAREG("LIMAYOAC", Z$, JJ&)
         End If
       End If
     .MoveNext
     Loop
   End With
   '
80 On Error GoTo 0
   Call APERBASDAT("CABAN")
   Screen.MousePointer = 11
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "where (CueContaI = " & CUEII% & " "
   SQLX$ = SQLX$ + " AND MarConsi = '*' "
   SQLX$ = SQLX$ + " AND CDBL(FechAcreD) >= " & FEINI
   SQLX$ = SQLX$ + " AND CDBL(FechAcreD) <= " & FEFIN & ") "
   SQLX$ = SQLX$ + " ORDER BY FECHACRED ASC, FEHOREAL ASC;"
   '
   Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   With CABATEMP
     On Error Resume Next
     .MoveFirst
     If Err Then
        GoTo 90
     End If
     On Error GoTo 0
     Do While Not .EOF
         IDEBE# = !IDEBECOMP
         IHABE# = !IHABECOMP
       If Int(CDbl(!FechAcred)) < FEINI Then
           TRANSPOR# = TRANSPOR# + IDEBE# - IHABE#
         ElseIf Int(CDbl(!FechAcred)) <= FEFIN Then
           Z$ = REGBLAN$("LIMAYOAC")
           Call REPLA(Z$, MKI$(CVINT(!FechConta)), 1, 2)
           Call REPLA(Z$, !CodiCom_Cor, 3, 12)
           Call REPLA(Z$, MKI$(CVINT(!FechAcred)), 15, 2)
             DEPASE$ = Left$(!Descrimov, 32)
           Call REPLA(Z$, DEPASE$, 17, 32)
           Call REPLA(Z$, MKD$(IDEBE#), 49, 8)
           Call REPLA(Z$, MKD$(IHABE#), 57, 8)
           Call REPLA(Z$, "CABAN", 73, 6)
           Call REPLA(Z$, "*", 80, 1)
           If IDEBE# >= 0.005 Or IHABE# >= 0.005 Then
             JJ& = JJ& + 1
             Call GRAREG("LIMAYOAC", Z$, JJ&)
           End If
       End If
     .MoveNext
     Loop
   End With
   '
90 Call SETULTREG("LIMAYOAC", JJ&)
   Call FILESORT("LIMAYOAC", "", 3, 3, "ASC")
   '
   CONCIBAN.List1.Clear
   Z$ = Space$(128)
   For U& = 1 To ULTREG("LIMAYOAC")
     X$ = REGLEIDO$("LIMAYOAC", U&)
     Z$ = Space$(128)
       FEX = CVI(Left$(X$, 2))
       FEXA$ = FECHATEX$(FEX)
     Call REPLA(Z$, FEXA$, 1, 8)
       COCOM$ = TRIM$(Mid$(X$, 3, 12))
     Call REPLA(Z$, COCOM$, 11, 15)
       FEX = CVI(Mid$(X$, 15, 2))
       FEXA$ = FECHATEX$(FEX)
     Call REPLA(Z$, FEXA$, 27, 8)
       DEPASE$ = TRIM$(Mid$(X$, 17, 32))
     Call REPLA(Z$, DEPASE$, 38, 28)
       IDEBE# = CVD(Mid$(X$, 49, 8))
       IHABE# = CVD(Mid$(X$, 57, 8))
     If IDEBE# > 0 Then
         Call REPLA(Z$, CSTRING$(MKD$(IDEBE#), 3, 14, 2, 2), 66, 14)
       Else
         Call REPLA(Z$, CSTRING$(MKD$(IHABE#), 3, 14, 2, 2), 80, 14)
     End If
       TRANSPOR# = TRANSPOR# + IDEBE# - IHABE#
     Call REPLA(Z$, CSTRING$(MKD$(TRANSPOR#), 3, 14, 2, 2), 94, 14)
       ORIMO$ = TRIM$(Mid$(X$, 73, 8))
     Call REPLA(Z$, ORIMO$, 121, 8)
     CONCIBAN.List1.AddItem Z$
     DoEvents
   Next U&
   Call CIERRARCH("LIMAYOAC")
   '
   '
   TRANSPOR# = Val(CONCIBAN.Label3)
   For KU& = 1 To CONCIBAN.List1.ListCount
     TTXX$ = CONCIBAN.List1.List(KU& - 1)
     TRANSPOR# = TRANSPOR# + Val(Mid$(TTXX$, 66, 14)) - Val(Mid$(TTXX$, 80, 14))
     Call REPLA(TTXX$, FORMATNUM$(TRANSPOR#, "F14.2"), 94, 14)
     CONCIBAN.List1.RemoveItem (KU& - 1)
     CONCIBAN.List1.AddItem TTXX$, KU& - 1
   Next KU&
   CONCIBAN.Label14 = FORMATNUM(TRANSPOR#, "F12.2")
   '
   TRANSPOR# = 0
   For KU& = 1 To CONCIBAN.List2.ListCount
     TTXX$ = CONCIBAN.List2.List(KU& - 1)
     TRANSPOR# = TRANSPOR# + Val(Mid$(TTXX$, 66, 14)) - Val(Mid$(TTXX$, 80, 14))
   Next KU&
   CONCIBAN.Label15 = FORMATNUM(TRANSPOR#, "F12.2")
   '
   CONCIBAN.Ltext2 = CODCUE$(CUEII%)
   CONCIBAN.Label4 = DENOCUE$(CUEII%)
   Screen.MousePointer = 1
   CONCIBAN.Show 1
   GoTo 10
   '
99 Command12.Enabled = True
End Sub

Private Sub Command13_Click()
    Command13.Enabled = False
    Call RECICHEQ
    Command13.Enabled = True
End Sub

Private Sub Command14_Click()
' este es el original con el que recuperamos RVC
   Command14.Enabled = False
   HOII% = HOY(HOS$)
   Screen.MousePointer = 11
   '
   HUBOCONCI% = 0
   VUELTA% = 1
   '
15 Call APERBASDAT("CABAN")
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "where TipoMovim = 'CHRECI' "
   'SQLX$ = SQLX$ + "AND Status = 1 "     ' aun en cartera
   Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
   '
   With CABATEMP
     On Error Resume Next
     .MoveLast
     If Err Then
         On Error GoTo 0
         Call MENSERR(24, "No Hay Ingresos Registrados\a la Cartera de Cheques y Valores.")
         GoTo 99
       Else
         On Error GoTo 0
         FIN& = .RecordCount
25       .MoveFirst
         JJ& = 0
         Do While Not .EOF
           JJ& = JJ& + 1
           Call TRACE(20, JJ&, FIN&)
           If TRIM$(!OpciMovim) <> "CH" Then GoTo 19 ' no es un cheque
           If IsNull(!FECHENTRE) = False Then
             If CVINT(!FECHENTRE) > 33 Then
               .Edit
               !Status = 2
               !FECHENTRE = !FEDEPOSENT
               .Update
               GoTo 19    ' salio de cartera
             End If
           End If
           '
           .Edit
           !Status = 1
           .Update
           FEACRED% = CVINT(!FechAcred)
           BANUFE$ = !Descrimov
           ICHEQ# = !ValAbsComp
           IMCHEX$ = TRIM$(FORMATNUM$(ICHEQ#, "F14.2"))
           '
           If TRIM$(Left$(BANUFE$, 12)) = "" Then
             Call REPLA(BANUFE$, "Cheque Nro:", 1, 11)
           End If
           '
           FEACRED% = CVINT%(!FechAcred)
           FEX = FEACRED%
           MESA$ = Mid$(FECHATEX$(FEX), 4, 5)
           FEMISI% = CVINT%(!FechEmisi)
           FERECEP% = CVINT(!FechConta)
           NURECIB$ = !CodiCom_Cor
           ORIGCHE$ = !REFECOM
           CUEII% = !CueContaI
           CUENTRA$ = TRIM$(CODCUE$(CUEII%)) + " - " + DENOCUE$(CUEII%)
           DEMOV$ = !Descrimov
             If VUELTA% = 2 Then DEMOV$ = Left$(!Descrimov, 32)
           TXD$ = REPLACAR$(DEMOV$, " ", "")
           '
           ' buscar por importe haber
           SQLY$ = "TipoMovim = 'CHENTR' AND ABS(IHabeComp - " & IMCHEX$ & ") < 0.005 "
           CajaBanc.FindFirst SQLY$
16         If CajaBanc.NoMatch = False Then
             TXH$ = REPLACAR$(CajaBanc!Descrimov, " ", "")
             If VUELTA% = 2 Then TXH$ = REPLACAR$(Left$(CajaBanc!Descrimov, 32), " ", "")
             If TXD$ = TXH$ Then
               CUEII% = CajaBanc!CueContaI
               CUENSAL$ = TRIM$(CODCUE$(CUEII%)) + " - " + DENOCUE$(CUEII%)
               FESALI% = CVINT(CajaBanc!FechConta)
               DOCUSA$ = TRIM$(CajaBanc!CodiCom_Cor)
               DESTIN$ = TRIM$(CajaBanc!REFECOM)
               TTXX$ = "Desea Conciliar la Salida del Cheque\"
               TTXX$ = TTXX$ + "'" + TRIM$(BANUFE$) + "'\"
               TTXX$ = TTXX$ + "por un Importe de " + TRIM$(IMCHEX$) + "\"
               TTXX$ = TTXX$ + "Entregado a " + DESTIN$ + " ?"
               If COMALTER%(TTXX$ + "\\Si, Conciliar\No, Mantener") = 1 Then
                 .Edit
                 !FEDEPOSENT = CVDAT(FESALI%)
                 !FECHENTRE = CVDAT(FESALI%)
                 !Document = DOCUSA$
                 !DESTICHEQ = DESTIN
                 !REFESALCHE = "Conciliación Documento " + DOCUSA$
                 !IDENSALCHE = AJUSTI$(HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 48)
                 !UsuaSalChe = AJUSTI$("Usuario: " + TRIM$(USERNAME$), 48)
                 !Status = 2                     ' Marca Salida de Cartera
                 !IdenBaja = IDACTU$             ' Marca Modificacion de registro - FECHA - HORA - USUARIO - TERMINAL
                 .Update
                 Call COMUNI("Se ha Conciliado el Valor '" + TRIM$(BANUFE$) + "'")
                 HUBOCONCI% = 1
               End If
               GoTo 19
             End If
             CajaBanc.FindNext SQLY$
             GoTo 16
           End If
19       .MoveNext
         Loop
         If VUELTA% = 1 Then
           VUELTA% = 2
           GoTo 15
         End If
     End If
   End With
   '
   Screen.MousePointer = 1
   TTXX$ = "Control Completo"
   If HUBOCONCI% = 1 Then TTXX$ = TTXX$ + " con Conciliaciones."
   Call COMUNI(TTXX$)
   '
99 Command14.Enabled = True
   '
End Sub

Private Sub Command19_Click()
' este es el modificado el 23 de agosto
   Dim IDENCHE0$(), IDENCHE1$(), IDENCHE2$(), IMPOCHE#(), FESALID%(), DOCUSAL$(), DESTINX$()
   Dim DEMECO$(), CUENCONCO%()
   '
   Command14.Enabled = False
   HOII% = HOY(HOS$)
   CUECA% = CUEICARCH%
   Screen.MousePointer = 11
   '
   CUECHECA$ = CONTROL$("", "CUECHECA")
   CUEICHEQ% = CUEINT%(CUECHECA$)
   If TRIM$(CUECHECA$) = "" Or CUEICHEQ% < 1 Then
      Screen.MousePointer = 1
      Command14.Enabled = True
      Call MENSERR(24, "Falta Definir Cuenta Contable\Asociada a Cartera de Valores")
      OPCHEQ04.Show 1
      Exit Sub
   End If
   '
   CAMECO& = ULTREG&("MECOBRA")
   ReDim DEMECO$(CAMECO&), CUENCONCO%(CAMECO&)
   For U& = 1 To CAMECO&
      XX$ = REGLEIDO$("MECOBRA", U&)
      DEMECO$(U&) = TRIM$(Mid$(XX$, 3, 12))
      CUENCONCO%(U&) = CUEINT%(TRIM$(Mid$(XX$, 45, 12)))
   Next U&
   '
   HUBOCONCI% = 0
   VUELTA% = 1
   '
   Call APERBASDAT("CABAN")
   SQLZ$ = "SELECT * FROM CAJABANC "
   SQLZ$ = SQLZ$ + "where TipoMovim = 'CHENTR' "
   Set cabachentr = CueCorri.OpenRecordset(SQLZ$, dbOpenDynaset)
   '
   With cabachentr
     On Error Resume Next
     .MoveFirst
     If Err Then
       On Error GoTo 0
       GoTo 10
     End If
     On Error GoTo 0
     Do While Not .EOF
        If !CueContaI <> CUECA% Then GoTo 9 ' SOLO SALIDAS DE CARTERA DE CHEQUES
        '
        ' enjuaque para guidi
        If !Descrimov = "PROVINC.          109307        04/07/05" Then
          If CDbl(!FechEmisi) = CDbl(CVDAT(FECHANUM("09/09/05"))) Then
            .Edit
            !Descrimov = "PROVINC.          10947         01/09/05"
            .Update
          End If
        End If
        ' fin enjuague para guidi
        '
        CASALI& = CASALI& + 1
        ReDim Preserve IDENCHE0$(CASALI&), IDENCHE1$(CASALI&), IDENCHE2$(CASALI&), IMPOCHE#(CASALI&), FESALID%(CASALI&), DOCUSAL$(CASALI&), DESTINX$(CASALI&)
        IDENCHE0$(CASALI&) = !Descrimov
        IDENCHE1$(CASALI&) = REPLACAR$(!Descrimov, " ", "")
        IDENCHE2$(CASALI&) = REPLACAR$(Left$(!Descrimov, 32), " ", "")
        IMPOCHE#(CASALI&) = !IHABECOMP
        FESALID%(CASALI&) = CVINT(!FechConta)
        DOCUSAL$(CASALI&) = TRIM$(!CodiCom_Cor)
        DESTINX$(CASALI&) = TRIM$(!REFECOM)
9    .MoveNext
     Loop
   End With
   '
10
'           ' buscar por importe haber
'           SQLY$ = "ABS(IHabeComp - " & IMCHEX$ & ") < 0.005 "
'           cabachentr.FindFirst SQLY$
'           If cabachentr.NoMatch = False Then
'             TXH$ = REPLACAR$(cabachentr!DescriMov, " ", "")
'             If vuelta% = 2 Then TXH$ = REPLACAR$(Left$(cabachentr!DescriMov, 32), " ", "")
'             If TXD$ = TXH$ Then
'               CUEII% = cabachentr!CueContaI
'               CUENSAL$ = TRIM$(CODCUE$(CUEII%)) + " - " + DENOCUE$(CUEII%)
'               FESALI% = CVINT(cabachentr!FechConta)
'               DOCUSA$ = TRIM$(cabachentr!CodiCom_Cor)
'               DESTIN$ = TRIM$(cabachentr!RefeCom)
15 Call APERBASDAT("CABAN")
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "where TipoMovim = 'CHRECI' "
   'SQLX$ = SQLX$ + "AND Status = 1 "     ' aun en cartera
   Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
   '
   With CABATEMP
     On Error Resume Next
     .MoveLast
     If Err Then
         On Error GoTo 0
         Call MENSERR(24, "No Hay Ingresos Registrados\a la Cartera de Cheques y Valores.")
         GoTo 99
       Else
         On Error GoTo 0
         FIN& = .RecordCount
25       .MoveFirst
         JJ& = 0
         Do While Not .EOF
           JJ& = JJ& + 1
           Call TRACE(20, JJ&, FIN&)
           If TRIM$(!OpciMovim) <> "CH" Then GoTo 29 ' no es un cheque
           If !Status < 1 Then GoTo 29 ' no entro a cartera, p.ej una transferencia directa al banco
           '
           ' BUSCAR LA SALIDA
           DEMOV$ = !Descrimov
           '
           ' ESTO ES UN ENJUAGUE PARA TANIT
           If DEMOV$ = "Prov.Bs.As.       49623014      12/04/06" Then
              DEMOV$ = "Prov.Bs.As.         49623013    07/04/06"
              .Edit
              !Descrimov = DEMOV$
              .Update
           End If
           ' FIN ENJUAGUE PARA TANIT
           '
           DEMOV1$ = REPLACAR$(DEMOV$, " ", "")
           DEMOV2$ = REPLACAR$(Left$(DEMOV$, 32), " ", "")
           ICHEQ# = !ValAbsComp
           '
           VUELTA% = 1
26         For KU& = 1 To CASALI&
             If IDENCHE1$(KU&) = DEMOV1$ Then
               If Abs(IMPOCHE#(KU&) - ICHEQ#) <= 0.05 Then    ' ajustado a 0.05 por diferencia asembli 27/11/09 - EPB
                 GoTo 27
               End If
             End If
             '
             If Abs(IMPOCHE#(KU&) - ICHEQ#) <= 0.005 Then    ' ajustado a 0.02 por diferencia asembli
               If IDENCHE1$(KU&) = DEMOV1$ Or (VUELTA% = 2 And IDENCHE2(KU&) = DEMOV2$) Then
                 ' ENCONTRO LA SALIDA
27               IDENCHE1$(KU&) = ""
                 IDENCHE2$(KU&) = ""
                 IMPOCHE#(KU&) = 0
                 If !Status <> 2 Or (CVINT(!FECHENTRE) <= 33 And CVINT(!FEDEPOSENT) <= 33) Then
                   .Edit
                   !FEDEPOSENT = CVDAT(FESALID%(KU&))
                   !FECHENTRE = CVDAT(FESALID%(KU&))
                   !Document = DOCUSAL$(KU&)
                   !DESTICHEQ = DESTINX$(KU&)
                   !REFESALCHE = "Conciliación Documento " + DOCUSAL$(KU&)
                   !IDENSALCHE = AJUSTI$(HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 48)
                   !UsuaSalChe = AJUSTI$("Usuario: " + TRIM$(USERNAME$), 48)
                   !Status = 2                     ' Marca Salida de Cartera
                   !IdenBaja = IDACTU$             ' Marca Modificacion de registro - FECHA - HORA - USUARIO - TERMINAL
                   .Update
                   MsgBox "Cheque " + DEMOV$ + " Recuperada Marca de Salida"
                   HUBOCONCI% = 1
                 End If
                 GoTo 29
               End If
             End If
           Next KU&
           If VUELTA% = 1 Then
             VUELTA% = 2
             GoTo 26
           End If
           '
           ' NO ENCONTRO LA SALIDA
           If !Status <> 1 Or CVINT(!FECHENTRE) > 33 Then
              .Edit
              !FEDEPOSENT = CVDAT(0)
              !FECHENTRE = CVDAT(0)
              !Document = ""
              !DESTICHEQ = ""
              !REFESALCHE = ""
              !IDENSALCHE = ""
              !UsuaSalChe = ""
              !Status = 1                     ' Marca está en Cartera
              !IdenBaja = ""                  ' Marca Modificacion de registro - FECHA - HORA - USUARIO - TERMINAL
              .Update
             MsgBox "Cheque " + DEMOV$ + " Recuperada Marca de 'En Cartera'"
             HUBOCONCI% = 1
           End If
           '
           ' esto es para sacar de cartera los registros de retenciones IB e IVA
           If !Status = 1 Then      ' FIGURA COMO QUE ESTÁ EN CARTERA
              If !CueContaI <> CUEICHEQ% Then
                 .Edit
                 !Status = 0
                 .Update
                 HUBOCONCI% = 1
              End If
           End If
           '
29       .MoveNext
         Loop
     End If
   End With
   '
   For KU& = 1 To CASALI&
     If Abs(IMPOCHE#(KU&)) >= 0.005 Then
        MsgBox "Cheque " + IDENCHE0$(KU&) + " Figura Salido sin Entrada - Consulte"
        HUBOCONCI% = 1
     End If
   Next KU&
   '
   Screen.MousePointer = 1
   TTXX$ = "Control Completo"
   If HUBOCONCI% = 1 Then
        TTXX$ = TTXX$ + " con Conciliaciones."
      Else
        TTXX$ = TTXX$ + " - O.K."
   End If
   Call COMUNI(TTXX$)
   '
99 Command14.Enabled = True
   '
End Sub

Sub Command15_Click()
    Command15.Enabled = False
    Call SUDIARCAJ   ' Impresion Subdiario de Caja y Bancos
    Command15.Enabled = True
End Sub

Private Sub Command16_Click()
    Command16.Enabled = False
    Call ENTRECHEQ
    Command16.Enabled = True
End Sub

Sub Command17_Click()
   Command17.Enabled = False
   Call RESCAJAPAN
   Command17.Enabled = True
End Sub

Private Sub Command18_Click()
   Command18.Enabled = False
   Call OPNOIN("")
   Command18.Enabled = True
End Sub


Sub Command2_Click()
    If Option1.Value = True Then
        Call LICHECAR
      ElseIf Option2.Value = True Then
        Call LICHEQEM
      ElseIf Option3.Value = True Then
        Call PODIFER
      ElseIf Option7.Value = True Then
        Call CONCIBANC
      ElseIf Option8.Value = True Then
        'Call SALTESOR
        '\\\OT-5-661-FG-06/10/05///
        Call RESUTES
        '\\\OT-5-661-FG-FIN///
      '\\\OT-05-0378-WAR-06/06/05///
      ElseIf Option10.Value = True Then
        Call LICHFERE
      '\\\OT-05-0378-WAR-FIN///
      ElseIf Option11.Value = True Then
        Call CHEQPENAC
    End If
End Sub

Private Sub Command20_Click()
    '
    FEXI% = DIANTE(DIANTE(HOY(HOS$)))
    AYER = CDbl(CVDAT(FEXI%))
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "where TipoMovim = 'CHRECI' "
    SQLX$ = SQLX$ + "AND Status > 0 "
    SQLX$ = SQLX$ + "AND CDBL(FechAcred) >= " & AYER & " "
    SQLX$ = SQLX$ + "ORDER BY NuClien ASC, FechAcred ASC;"
    
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        CONT& = 0
        Do While Not .EOF
          If TRIM(!OpciMovim) <> "CH" Then GoTo 19 ' no es un cheque
            If IsNull(!FECHENTRE) = False Then
                If CVINT(!FECHENTRE) > 0 Then
                  GoTo 19    ' salio de cartera
                End If
            End If
            
            CUECLI% = !NuClien
            FEACRED% = CVINT(!FechAcred)
            BANUFE$ = !Descrimov
            FEEMI% = CVINT(!FechEmisi)
            FEREC% = CVINT(!FechConta)
            DOCU$ = !CodiCom_Cor
            ICHEQ# = !ValAbsComp
          
            Call REPLA(X$, MKI$(CUECLI%), 1, 2)
            Call REPLA(X$, MKI$(FEACRED%), 3, 2)
            Call REPLA(X$, BANUFE$, 5, 40)
            Call REPLA(X$, MKI$(FEEMI%), 45, 2)
            Call REPLA(X$, MKI$(FEREC%), 47, 2)
            Call REPLA(X$, DOCU$, 49, 12)
            Call REPLA(X$, MKD$(ICHEQ#), 61, 8)
          
            CONT& = CONT& + 1
    
            Call GRAREG("CHEPEAC", X$, CONT&)
19          .MoveNext
        Loop
      End If
    End With
    Call SETULTREG("CHEPEAC", CONT&)
    Call FINAL("*CHEPEAC")
    '
End Sub




Private Sub Command3_Click()
    Command3.Enabled = False
    Call CARGACHECAR
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    Call ENTREGACHE
    Command4.Enabled = True
End Sub


Private Sub Command5_Click()
   Command5.Enabled = False
   Call CHERECHAZA
   Command5.Enabled = True
End Sub

Private Sub Command6_Click()
    Call HELPONLINE("GECHECAR")
End Sub

Sub Command7_Click()
   If Option6.Value = True Then
      Call BUSCACHEQ
   Else
     Call BUSCHECUIT
   End If
End Sub


Sub PODIFER()

   Static CAMP%, CUEMP%(2048), MEPA%(2048)
   '
   Dim REFE$(2048)
   Dim FECHAMO%(2048)
   Dim IPODIF#(2048)
   '
   If EXTE$ <> "DAT" Then
      Call MENSERR(24, "Para Consultar Posición Diferida\Debe Pasar al Ejercicio ACTUAL")
      Exit Sub
   End If
   '
   If CAMP% <= 0 Then
      For I& = 1 To ULTREG&("MEPAGO")
         F0$ = REGLEIDO$("MEPAGO", I&)
         CAMP% = I&
         MEPA%(CAMP%) = CVI(Left$(F0$, 2))
         CUEX$ = Mid$(F0$, 45, 12)
         CUEMP%(CAMP%) = CUEINT%(CUEX$)
         If TRIM$(CODCUE$(CUEMP%(CAMP%))) <> TRIM$(CUEX$) Then
            Call MENSERR(24, "Cuenta Contable Incorrecta\para Medio de Pago '" + TRIM$(Mid$(F0$, 3, 12)) + "'\-\Resultados Imprevisibles.")
            Call FINAL("")
         End If
     Next I&
   End If
   '
   HOII% = HOY(HOS$)
   DES% = FECHANUM("01" + Mid$(HOS$, 3))
   VDEF$ = MKI$(DES%) + MKI$(HOII%)
   '
5  OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
   If OBX$ = "" Then Exit Sub
   '
   DES% = CVI(Left$(OBX$, 2))
   HAS% = CVI(Mid$(OBX$, 3, 2))
   If HAS% < DES% Then Exit Sub
   '
   If DES% < FECHANUM("01/10/01") Then
      Call MENSERR(24, "Fecha no Válida")
      GoTo 5
   End If
   '
   If HAS% > FINEJ% Then
      Call MENSERR(24, "Fecha 'Hasta' es Posterior\al Fin del Ejercicio.")
      GoTo 5
   End If
   '
10 'Call SELECHO("ECUECON")
   'CUEX$ = VALACT1$("ECUECON")
   'If CUEX$ = "" Then Exit Sub
   '
   Screen.MousePointer = 11
   KLM& = 0
   '
   For cuebal& = 1 To ULTREG&("CUEBANC")
   CBB$ = REGLEIDO$("CUEBANC", cuebal&)
   DCBB$ = Left$(CBB$, 32)
   CUEX$ = Left$(CBB$, 12)
   CUE% = CUEINT%(CUEX$)
   If CUE% < 1 Then
      Call MENSERR(24, "Cuenta Inexistente")
      GoTo 80
   End If
   '
   Frame1.Refresh
   Call FRESHALL
   CUEBIN$ = MKI$(CUE%)
   CODENOCUE$ = TRIM$(CODCUE$(CUE%)) + " - " + TRIM$(DENOCUE$(CUE%))
   JJ& = 0
   SALINI# = 0
   '
   If DES% >= FECHANUM("01/10/01") Then
      CCCC$ = TRIM$(CODCUE$(CUE%))
      If CCCC$ = "111401" Then
         SALINI# = 10886.85   ' 11409.33
         SALINI# = SALINI# + 114.95 ' para tener en cuenta cheque duplicado flama
      End If
      If CCCC$ = "111502" Then SALINI# = 11791.1
      If CCCC$ = "111403" Then SALINI# = 13039.69   ' 2414.3
      If CCCC$ = "111406" Then SALINI# = 30026.22   ' 29769.7
      If CCCC$ = "111506" Then SALINI# = 381.86
   End If
   '
   ' CARGAR MOVIMIENTOS DE DIARIO
   EXTEA$ = EXTE$
   If EXTEA$ = "DAT" Then
     Close #NX9%: NX9% = 0
     Close #NX10%: NX10% = 0
     Close #NX11%: NX11% = 0
     EXTE$ = "OLD"
   End If
   '
12 PMC$ = PUNMODIAR$(CUE%)
   FIN& = Len(PMC$)
   For KK% = 1 To Len(PMC$) Step 2
      UU& = KK% + 1
      Call TRACE(20, UU&, FIN&)
      REGDIA% = CVI(Mid$(PMC$, KK%, 2))
      Call LEEDIARIO(REGDIA%)
      If CVS(E12$) < 1.05 Then GoTo 19 ' SALTEA ASIENTO APERTURA
      If Asc(E18$) <> 1 Then  ' SALTEA ASIENTOS RESUMEN DE CAJA
         '
         Call LEEPASDIA(REGDIA%, 2)
         DETAPAS$ = TRIM$(DETAPAS$)
         If DETAPAS$ <> "" Then
            DETAPAS$ = TRIM$(DETAPAS$) + " - "
         End If
         DETAPAS$ = DETAPAS$ + REFASIEN$
         '
         ACIERR% = 0
         If InStr(UCase$(DETAPAS$), "(CIERR") > 0 Or InStr(UCase$(DETAPAS$), "(*") > 0 Then
           ACIERR% = 1
         End If
         MARCADIA$ = "*"
         FEMOREA% = CVI(E15$)
         IMPO# = CVD(E16$)
         If FEMOREA% < DES% Then
             If ACIERR% = 1 Then IMPO# = 0
             If FEMOREA% > FECHANUM("31/09/01") Then GoSub 90 'TOTALIZAR INICIAL
          ElseIf FEMOREA% <= HAS% Then
            FEMORE% = FEMOREA%
            FEREGI% = FEMOREA%
            GoSub 95
         End If
      End If
19 Next KK%
   '
   If EXTE$ = "OLD" Then
     If EXTEA$ = "DAT" Then
       Close #NX9%: NX9% = 0
       Close #NX10%: NX10% = 0
       Close #NX11%: NX11% = 0
       EXTE$ = "DAT"
       EXTEA$ = ""
       GoTo 12
     End If
   End If
   '
   '
20 ' CARGAR MOVIMIENTOS DE CAJA Y BANCOS
   EXTEA$ = EXTE$
   If EXTEA$ = "DAT" Then
     Close #NX5%: NX5% = 0
     Close #NX6%: NX6% = 0
     Close #NX7%: NX7% = 0
     EXTE$ = "OLD"
   End If
   '
22 PMC$ = PUNMOCAJA$(CUE%)
   FIN& = Len(PMC$)
   For KK% = 1 To Len(PMC$) Step 2
      UU& = KK% + 1
      Call TRACE(20, UU&, FIN&)
      REGDIAL& = CVI(Mid$(PMC$, KK%, 2))
      Call LEEPASCAJ(REGDIAL&, 1)
      FEMOVI% = CVI(D15$)
      IMPO# = CVD(D16$)
      '
      If TRIM$(DETAPAS$) = "" Then
         Call LEEPASCAJ(REGDIAL&, 2)
         DETAPAS$ = REFASIEN$
      End If
      MARCADIA$ = ""
      If InStr(DETAPAS$, "- OP") > 0 Then GoTo 29
      If Left$(DETAPAS$, 8) = "DEPOSITO" Then
         FEMORE% = FECHACRE%(FEMOVI%, Val(Mid$(DETAPAS$, 10, 3)))
         If FEMORE% < DES% Then
            If FEMORE% > FECHANUM("31/09/01") Then GoSub 90
          ElseIf FEMORE% <= HAS% Then
            GoSub 95
         End If
         GoTo 29
      End If
      If FEMOVI% < DES% Then
        If FEMOVI% > FECHANUM("31/09/01") Then GoSub 90
        GoTo 29
      End If
      If FEMOVI% >= DES% Then
        If FEMOVI% <= HAS% Then
          FEMORE% = FEMOVI%
          GoSub 95
        End If
      End If
      '
29 Next KK%
   '
   If EXTE$ = "OLD" Then
     If EXTEA$ = "DAT" Then
       Close #NX5%: NX5% = 0
       Close #NX6%: NX6% = 0
       Close #NX7%: NX7% = 0
       EXTE$ = "DAT"
       EXTEA$ = ""
       GoTo 22
     End If
   End If
   '
   For U& = 1 To ULTREG&("CHEQUEMI")
     X$ = REGLEIDO$("CHEQUEMI", U&)
      PROV% = CVI(Left$(X$, 2))
      NUMERO$ = Mid$(X$, 3, 12)
      banco1$ = Mid$(X$, 15, 16)
      EMIS% = CVI(Mid$(X$, 47, 2)) ' toma fecha de O/Pago
      IMPO# = (-1) * CVD(Mid$(X$, 33, 8))
      NOPAGO& = CVS(Mid$(X$, 41, 4))
      VENCI% = CVI(Mid$(X$, 49, 2))
      '
      If VENCI% < EMIS% Then  ' modificacion del 28/01/02. Si
         VENCI% = EMIS%       ' la fecha del valor es anterior
      End If                  ' a la OP, toma la de la OP
      '
      QUE% = CVI(Mid$(X$, 45, 2))
            For UX% = 1 To CAMP%
               If MEPA%(UX%) = QUE% Then GoTo 38
            Next UX%
            Call MENSERR(24, "Medio de Pago" + Str$(QUE%) + " no Definido")
            Call FINAL("")
            '
38    CUEH% = CUEMP%(UX%)
      If CUEH% = CUE% Then
        FEMORE% = VENCI%
If EMIS% <= FECHANUM("31/09/01") And VENCI% > FECHANUM("31/09/01") Then
SALINICA# = SALINICA# + IMPO#
AAA = BBB
End If
        If FEMORE% < DES% Then
          If FEMORE% > FECHANUM("31/09/01") Then GoSub 90
        End If
        If FEMORE% >= DES% Then
          If FEMORE% <= HAS% Then
            FEX = EMIS%
            DETAPAS$ = "Ch." + TRIM$(NUMERO$)
            DETAPAS$ = DETAPAS$ + " - OP." + TRIM$(Str$(NOPAGO&))
            DETAPAS$ = DETAPAS$ + " - " + TRIM$(RASOCLI$((-1) * PROV%))
            GoSub 95
          End If
        End If
      End If
   Next U&
   '
   FEX = DES%
   FECHAMO%(0) = DIANTE(FEX)
   FEX = FECHAMO(0)
   REFE$(0) = "Saldo al " + FECHATEX$(FEX)
   IPODIF#(0) = SALINI#
   '
   For KKL& = 0 To JJ&
     For KKL1& = KKL& + 1 To JJ&
       If FECHAMO%(KKL1&) < FECHAMO%(KKL&) Then
         FX1% = FECHAMO%(KKL&)
         FX2$ = REFE$(KKL&)
         FX3# = IPODIF#(KKL&)
         FECHAMO%(KKL&) = FECHAMO%(KKL1&)
         REFE$(KKL&) = REFE$(KKL1&)
         IPODIF#(KKL&) = IPODIF#(KKL1&)
         FECHAMO%(KKL1&) = FX1%
         REFE$(KKL1&) = FX2$
         IPODIF#(KKL1&) = FX3#
        End If
     Next KKL1&
     GoSub 96
   Next KKL&
AAA = SALINICA#
ABC$ = CCCC$
SALINICA# = 0
80 Next cuebal&
   '
   DCBB$ = "Cartera de Valores"
   JJ& = (-1)
   For U& = 1 To ULTREG&("BADACHE")
     X$ = REGLEIDO$("BADACHE", U&)
     If CVI(Mid$(X$, 133, 2)) < 1 Then  ' aun en cartera
       FEMORE% = CVI(Mid$(X$, 59, 2))
       If FEMORE% <= HAS% Then
         DETAPAS$ = Left$(X$, 40)
         IMPO# = CVD(Mid$(X$, 49, 8))
         GoSub 95
       End If
     End If
   Next U&
   '
   For KKL& = 0 To JJ&
     For KKL1& = KKL& + 1 To JJ&
       If FECHAMO%(KKL1&) < FECHAMO%(KKL&) Then
         FX1% = FECHAMO%(KKL&)
         FX2$ = REFE$(KKL&)
         FX3# = IPODIF#(KKL&)
         FECHAMO%(KKL&) = FECHAMO%(KKL1&)
         REFE$(KKL&) = REFE$(KKL1&)
         IPODIF#(KKL&) = IPODIF#(KKL1&)
         FECHAMO%(KKL1&) = FX1%
         REFE$(KKL1&) = FX2$
         IPODIF#(KKL1&) = FX3#
        End If
     Next KKL1&
     GoSub 96
   Next KKL&
   '
   Call SETULTREG("PODIFER", KLM&)
   Call CIERRARCH("PODIFER")
   Call CIERRARCH("*.*")
   '
   Screen.MousePointer = 1
   Call FINAL("*PODIFER")
   '
   Exit Sub
   
90 SALINI# = SALINI# + IMPO#
   Return
   '
95 If Abs(IMPO#) >= 0.005 Then
      JJ& = JJ& + 1
      REFE$(JJ&) = DETAPAS$
      FECHAMO%(JJ&) = FEMORE%
      IPODIF#(JJ&) = IMPO#
   End If
   Return
   '
96 ZZZ$ = REGBLAN$("PODIFER")
   If DCBB$ <> DCBBA$ Then
     SARRAS# = 0
     DCBBA$ = DCBB$
   End If
   Call REPLA(ZZZ$, DCBB$, 1, 32)
   Call REPLA(ZZZ$, MKI$(FECHAMO%(KKL&)), 33, 2)
   Call REPLA(ZZZ$, REFE$(KKL&), 35, 30)
   ACIERR% = 0
   If InStr(UCase$(REFE$(KKL&)), "(CIERR") > 0 Or InStr(REFE$(KKL&), "(*") > 0 Then
      ACIERR% = 1
   End If
   SARRAS# = SARRAS# + IPODIF#(KKL&)
   If ACIERR% > 0 Then SARRAS# = SARRAS# - IPODIF#(KKL&)
   If IPODIF#(KKL&) > 0 Then
      Call REPLA(ZZZ$, MKD$(IPODIF#(KKL&)), 65, 8)
      If ACIERR% > 0 Then Call REPLA(ZZZ$, MKD$(Abs(IPODIF#(KKL&))), 73, 8)
     Else
      Call REPLA(ZZZ$, MKD$(Abs(IPODIF#(KKL&))), 73, 8)
      If ACIERR% > 0 Then Call REPLA(ZZZ$, MKD$(Abs(IPODIF#(KKL&))), 65, 8)
   End If
   Call REPLA(ZZZ$, MKD$(SARRAS#), 81, 8)
   KLM& = KLM& + 1
   Call GRAREG("PODIFER", ZZZ$, KLM&)
   Return
   '
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   Call SELMESA(DES%, HAS%, CORRESP$)
   If CORRESP$ <> "" Then
     Dim CUEK%(256), SUDEBE#(256), SUHABE#(256)
     CACUEN% = 0
     FIN& = ULTREG&("CAJABANC")
     '
     TINOCON$ = "/AS/FC/ND/NC/CO/RH/"
     For UU& = 1 To FIN&
       Call TRACE(20, UU&, FIN&)
       X$ = REGLEIDO$("CAJABANC", UU&)
       STATU% = Asc(Mid$(X$, 121, 1))
       '
       If STATU% < 9 Then      ' no está anulado
         FEAS% = CVI(Mid$(X$, 61, 2))
         If FEAS% >= DES% Then
           If FEAS% <= HAS% Then
             NUCO$ = Left$(X$, 12)
             TICO$ = UCase$(Left$(NUCO$, 2))
             '
             If TICO$ = "TF" Then
               If Mid$(X$, 69, 5) <> "TFDe " Then
                 If Mid$(X$, 69, 4) <> "TFA " Then
                   GoTo 19
                 End If
               End If
             End If
             '
             If InStr(TINOCON$, "/" + TICO$ + "/") < 1 Then
               IDEBE# = CVD(Mid$(X$, 153, 8))
               IHABE# = CVD(Mid$(X$, 161, 8))
               CUE% = CVI(Mid$(X$, 65, 2))
               '
               For KI% = 1 To CACUEN%
                 If CUEK%(KI%) = CUE% Then
                   SUDEBE#(KI%) = SUDEBE#(KI%) + IDEBE#
                   SUHABE#(KI%) = SUHABE#(KI%) + IHABE#
                   GoTo 19     ' NEXT UU&
                 End If
               Next KI%
               CACUEN% = CACUEN% + 1
               KI% = CACUEN%
               CUEK%(KI%) = CUE%
               SUDEBE#(KI%) = IDEBE#
               SUHABE#(KI%) = IHABE#
               '
             End If
           End If
         End If
       End If
19   Next UU&
     '
     TTXXY$ = Chr$(1) + Chr$(3) + MKI$(HAS%)
     J& = 0
     For IL& = 1 To ULTREG&("ASIDIAR")
       X$ = REGLEIDO$("ASIDIAR", IL&)
       If Left$(X$, 4) <> TTXXY$ Then
         J& = J& + 1
         Call GRAREG("ASIDIAR", X$, J&)
       End If
     Next IL&
     '
     For KI% = 1 To CACUEN%
       If Abs(SUDEBE#(KI%) - SUHABE#(KI%)) >= 0.005 Then
         X$ = REGBLAN$("ASIDIAR")
         Call REPLA(X$, TTXXY$, 1, 4)
         Call REPLA(X$, MKI$(CUEK%(KI%)), 5, 2)
         Call REPLA(X$, DENOCUE$(CUEK%(KI%)), 9, 30)
         Call REPLA(X$, MKD$(SUDEBE#(KI%) - SUHABE#(KI%)), 39, 8)
         J& = J& + 1
         Call GRAREG("ASIDIAR", X$, J&)
       End If
     Next KI%
     Call SETULTREG("ASIDIAR", J&)
     Call CIERRARCH("ASIDIAR")
     Call CONASIDIA(HAS%, CORRESP$)
     '
   End If
   Screen.MousePointer = 1
   Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   OPCHEQ04.Show 1
   Command9.Enabled = True
End Sub







Private Sub conciliacionbancaria_Click()
    Call Command12_Click
End Sub

Private Sub config_Click()
     Call Command9_Click
End Sub

Private Sub controldecarteradecheques_Click()
    Call Command19_Click
End Sub

Private Sub depositos_Click()
   Call Command11_Click
End Sub

Private Sub entregadevaloressinop_Click()
    Call Command16_Click
End Sub

Private Sub impresora1_Click()
    Option4.Value = True
    Option10.Value = True
    Call Command2_Click

End Sub

Private Sub impresora2_Click()
    Option4.Value = True
    Option2.Value = True
    Call Command2_Click

End Sub

Private Sub impresora3_Click()
    Option4.Value = True
    Option11.Value = True
    Call Command2_Click
End Sub

Private Sub impresora5_Click()
    Option4.Value = True
    Option1.Value = True
    Call Command2_Click

End Sub

Private Sub Option3_DblClick()
   Call INDEXCAJA(MAXDIA&)
   '
   DES% = COMEJ%
   HAS% = FINEJ%
   '
10 Call SELECHO("ECUECON")
   CUEX$ = VALACT1$("ECUECON")
   If CUEX$ = "" Then Exit Sub
   '
   CUE% = CUEINT%(CUEX$)
   If CUE% < 1 Then
      Call MENSERR(24, "Cuenta Inexistente")
      GoTo 10
   End If
   '
   Call FRESHALL
   CUEBIN$ = MKI$(CUE%)
   CODENOCUE$ = TRIM$(CODCUE$(CUE%)) + " - " + TRIM$(DENOCUE$(CUE%))
   JJ& = 0: TRANSPORTE# = 0
   '
   ' CARGAR MOVIMIENTOS DE DIARIO
   PMC$ = PUNMODIAR$(CUE%)
   FIN& = Len(PMC$)
   For KK% = 1 To Len(PMC$) Step 2
      UU& = KK% + 1
      Call TRACE(20, UU&, FIN&)
      REGDIA% = CVI(Mid$(PMC$, KK%, 2))
      Call LEEDIARIO(REGDIA%)
      If Asc(E18$) <> 1 Then  ' SALTEA ASIENTOS RESUMEN DE CAJA
         Call LEEPASDIA(REGDIA%, 1)
         If TRIM$(DETAPAS$) = "" Then
            Call LEEPASDIA(REGDIA%, 2)
            DETAPAS$ = REFASIEN$
         End If
         MARCADIA$ = "*"
         If FECHASIEN% < DES% Then
             TRANSPORTE# = TRANSPORTE# + IMPOPAS#
             GoTo 19
           ElseIf FECHASIEN% <= HAS% Then
             GoSub 90
         End If
      End If
19 Next KK%
   '
   ' CARGAR MOVIMIENTOS DE CAJA Y BANCOS
   PMC$ = PUNMOCAJA$(CUE%)
   FIN& = Len(PMC$)
   For KK% = 1 To Len(PMC$) Step 2
      UU& = KK% + 1
      Call TRACE(20, UU&, FIN&)
      REGDIAL& = CVI(Mid$(PMC$, KK%, 2))
      Call LEEPASCAJ(REGDIAL&, 1)
      If TRIM$(DETAPAS$) = "" Then
         Call LEEPASCAJ(REGDIAL&, 2)
         DETAPAS$ = REFASIEN$
      End If
      MARCADIA$ = ""
      If FECHASIEN% < DES% Then
          TRANSPORTE# = TRANSPORTE# + IMPOPAS#
          GoTo 29
        ElseIf FECHASIEN% <= HAS% Then
          GoSub 90
      End If
29 Next KK%
   '
   If Abs(TRANSPORTE#) >= 0.005 Then
      FEX = DES%
      FECHASIEN% = DIANTE(FEX)
      NROASIEN = 0
      NROASDEF% = 0
      MARCADIA$ = " "
      DETAPAS$ = "SALDO ANTERIOR"
      IMPOPAS# = TRANSPORTE#
      GoSub 90
   End If
   '
   Call SETULTREG("RESUCON", JJ&)
   Call CIERRARCH("RESUCON")
   '
   ' hasta aqui genero resume de cuenta de C&B en RESUCON.WKF
   '
   '
   Call GEPODIF(CUE%)
   '
   'hasta aqui genero posicion diferida en PODIFER.WKF
   '
   '
   UPO& = ULTREG&("PODIFER")
   FIN& = ULTREG&("RESUCON")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("RESUCON", U&)
     IDERE# = CVD(Mid$(XX$, 105, 8))
     IHARE# = CVD(Mid$(XX$, 113, 8))
     '
     For KKU& = 1 To UPO&
       YY$ = REGLEIDO$("PODIFER", KKU&)
       If Abs(CVD(Mid$(YY$, 65, 8)) - IDERE#) < 0.005 Then
         If Abs(CVD(Mid$(YY$, 73, 8)) - IHARE#) < 0.005 Then
           Call GRAREG("PODIFER", String$(96, 0), KKU&)
           Call GRAREG("RESUCON", String$(128, 0), U&)
           GoTo 59
         End If
       End If
     Next KKU&
     '
59 Next U&
   '
   JJ& = 0
   For U& = 1 To ULTREG&("PODIFER")
     XX$ = REGLEIDO$("PODIFER", U&)
     If XX$ <> String$(96, 0) Then
       JJ& = JJ& + 1
       Call GRAREG("PODIFER", XX$, JJ&)
     End If
   Next U&
   Call SETULTREG("PODIFER", JJ&)
   Call CIERRARCH("PODIFER")
   '
   JJ& = 0
   For U& = 1 To ULTREG&("RESUCON")
     XX$ = REGLEIDO$("RESUCON", U&)
     If XX$ <> String$(128, 0) Then
       JJ& = JJ& + 1
       Call GRAREG("RESUCON", XX$, JJ&)
     End If
   Next U&
   Call SETULTREG("RESUCON", JJ&)
   Call CIERRARCH("RESUCON")
   '
   Call CIERRARCH("*.*")
80 OPX% = COMALTER%("Conciliacion Completa\\Pos.Diferida\Res.Cuenta\Terminar")
   If OPX% = 1 Then
       Call FINAL("*CPODIFE")
       GoTo 80
     ElseIf OPX% = 2 Then
       Call FINAL("*CRESUCO")
       GoTo 80
   End If
   Exit Sub
   '
   '
90 XX$ = REGBLAN$("RESUCON")
   Call REPLA(XX$, CODENOCUE$, 1, 64)
   Call REPLA(XX$, MKI$(FECHASIEN%), 65, 2)
   Call REPLA(XX$, MKS$(NROASIEN), 67, 4)
   Call REPLA(XX$, MKI$(NROASDEF%), 71, 2)
   Call REPLA(XX$, MARCADIA$, 73, 2)
   Call REPLA(XX$, DETAPAS$, 75, 30)
   If IMPOPAS# >= 0 Then
         Call REPLA(XX$, MKD$(IMPOPAS#), 105, 8)
      Else
         Call REPLA(XX$, MKD$(Abs(IMPOPAS#)), 113, 8)
   End If
   JJ& = JJ& + 1
   Call GRAREG("RESUCON", XX$, JJ&)
   '
Return
'
End Sub

Private Sub Form_Load()
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Picture16.ZOrder 0
      Me.BackColor = &HFCD7B6
    End If

   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
   End If
   If COMEJ% < 1 Then
      Call CARDATEJ
   End If

   Option6.Value = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
End Sub

Sub BUSCACHEQ()
   '
30 NCHE$ = OBJPLA$("CARNUCHE", NCHE$)
   If NCHE$ = "" Then
      Exit Sub
   End If
   '
   NCHE$ = TRIM$(NCHE$)
   Screen.MousePointer = 11
   Call FRESHECHO("CHESELEC")
   '
   Call APERBASDAT("CABAN")
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "where TipoMovim = 'CHRECI' "
   SQLX$ = SQLX$ + "and Status > 0 "
   SQLX$ = SQLX$ + "AND DescriMov LIKE  '*" & NCHE$ & "*'"
   '
   Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   JJ& = 0
   With CABATEMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do While Not .EOF
         If TRIM(!OpciMovim) <> "CH" Then GoTo 39 ' no es un cheque
         ICHX$ = MKD$(!ValAbsComp)
         X$ = AJUSTI$(!Descrimov, 40) + String$(8, 0) + ICHX$
         NCHX$ = Mid$(X$, 19, 12)
         '
         JJ& = JJ& + 1
         Call GRAREG("CHESELEC", X$, JJ&)
39     .MoveNext
       Loop
     End If
   End With
   '
   Call SETULTREG("CHESELEC", JJ&)
   Call CIERRARCH("CHESELEC")
   Screen.MousePointer = 1
   '
   CASEL& = ULTREG&("CHESELEC")
   If CASEL& < 1 Then
         Call COMUNI("Número no Encontrado")
         GoTo 30
       ElseIf CASEL& = 1 Then
         Y$ = REGLEIDO$("CHESELEC", 1&)
         GoTo 45
   End If
   '
   Y$ = REGSEL$("CHESELEC")
   If Len(Y$) <= 4 Then GoTo 30
   '
45 BANUFE$ = AJUSTI$(Y$, 40)
   IPASE# = CVD(Mid$(Y$, 49, 8))
   IPASX$ = TRIM$(FORMATNUM$(IPASE#, "F14.2"))
   Call APERBASDAT("CABAN")
   SQLZ$ = "SELECT * FROM CAJABANC "
   SQLZ$ = SQLZ$ + "where TipoMovim = 'CHRECI' "
   SQLZ$ = SQLZ$ + "and Status > 0 "
   SQLZ$ = SQLZ$ + "AND ABS(ValAbsComp - " & IPASX$ & ") < 0.005 "
   '
   Set CABASELE = CueCorri.OpenRecordset(SQLZ$, dbOpenDynaset)
   '
   With CABASELE    ' CajaBanc
     'IPASX$ = TRIM$(FORMATNUM$(IPASE#, "F14.2"))
     SQLX$ = "TipoMovim = 'CHRECI' "
     SQLX$ = SQLX$ + "AND Status > 0 "
     SQLX$ = SQLX$ + "AND DescriMov = '" & BANUFE$ & "' "
     SQLX$ = SQLX$ + "AND ABS(ValAbsComp - " & IPASX$ & ") < 0.005 "
     '
     .FindFirst SQLX$
     If .NoMatch = False Then
       '
       'ESTO ES PARA RECUPERAR CHEQUES QUE SALIERON EQUIVOCADOS
       If IsNull(!FECHENTRE) = True Or CVINT(!FECHENTRE) <= 33 Then
         BANUFE$ = !Descrimov
         If Mid$(BANUFE$, 19, 2) = "  " Then
           If Mid$(BANUFE$, 21, 1) <> " " Then
             BANUFE$ = Left$(BANUFE$, 18) + Mid$(BANUFE$, 21, 12) + "  " + Mid$(BANUFE$, 33, 8)
             .Edit
             !Descrimov = BANUFE$
             .Update
             '
             SQLY$ = "select * from CAJABANC "
             SQLY$ = SQLY$ + "WHERE TipoMovim = 'CHENTR' "
             SQLY$ = SQLY$ + "AND DescriMov = '" & BANUFE$ & "' "
             SQLY$ = SQLY$ + "AND  MID$(CodiCom_Cor,11,2) <> 'nl'"
             Set CABATEMP1 = CueCorri.OpenRecordset(SQLY$, dbOpenSnapshot)
             On Error Resume Next
             CABATEMP1.MoveFirst
             If Err < 1 Then
               On Error GoTo 0
               RCX$ = CABATEMP1!REFECOM
               FCX = CABATEMP1!FechConta
               CCX$ = CABATEMP1!CodiCom_Cor
               RFX$ = CABATEMP1!OBSERGEN
               FHX$ = CABATEMP1!FeHoReal
               '
               .Edit
               !DESTICHEQ = RCX$
               !FEDEPOSENT = FCX
               !FECHENTRE = FCX
               !Document = CCX$
               !REFESALCHE = RFX$
               !IDENSALCHE = FHX$
               !Status = 2
               .Update
             End If
           End If
         End If
       End If
       ' FIN RECUPERACION CHEQUES EQUIVOCADOS
       '
       YZ$ = String$(512, 0)
       Call REPLA(YZ$, Mid$(!Descrimov, 1, 32), 71, 32)
       Call REPLA(YZ$, Chr$(!Status), 121, 1)
       Call REPLA(YZ$, MKD$(!ValAbsComp), 145, 8)
       Call REPLA(YZ$, MKI$(CVINT(!FechEmisi)), 111, 2)
       Call REPLA(YZ$, MKI$(CVINT(!FechConta)), 61, 2)
       Call REPLA(YZ$, MKI$(CVINT(!FechAcred)), 63, 2)
       Call REPLA(YZ$, !CodiCom_Cor, 1, 12)
       Call REPLA(YZ$, !REFECOM, 13, 48)
       If IsNull(!DESTICHEQ) = False Then Call REPLA(YZ$, !DESTICHEQ, 293, 30)
       Call REPLA(YZ$, MKI$(CVINT(!FEDEPOSENT)), 275, 2)
       If IsNull(!Document) = False Then Call REPLA(YZ$, !Document, 277, 16)
       If IsNull(!DESTICHEQ) = False Then Call REPLA(YZ$, !DESTICHEQ, 323, 40)
       If IsNull(!REFESALCHE) = False Then Call REPLA(YZ$, !REFESALCHE, 363, 40)
       If IsNull(!IDENSALCHE) = False Then Call REPLA(YZ$, !IDENSALCHE, 403, 40)
       GoTo 50
     End If
     '
     Call MENSERR(24, "Imposible Presentar Datos del Cheque")
     GoTo 30
     '
   End With
     '
50 BANUFEA$ = BANUFE$
   X$ = OBJPLA$("MUESTRACHEGES", YZ$)
   If X$ <> "" Then
     If CVI(Mid$(X$, 63, 2)) <> CVI(Mid$(YZ$, 63, 2)) Then
       If DERACCE%("MOFEACRE", "Modificar Fecha de Acreditación Cheques") = 2 Then
         If COMALTER%("Confirme Modificación de Fecha de Acreditación\\No, Conservar\Si, Cambiar") = 2 Then
           IPASE# = CVD(Mid$(X$, 145, 8))
           IPASX$ = TRIM$(FORMATNUM$(IPASE#, "F14.2"))
           NUEFECHA% = CVI(Mid$(X$, 63, 2))
           FEX = NUEFECHA%
           NUFECHX$ = FECHATEX$(FEX)
           '
           Call APERBASDAT("CABAN")
           With CABASELE   ' CajaBanc
             SQLX$ = "TipoMovim = 'CHRECI' "
             SQLX$ = SQLX$ + "AND Status > 0 "
             SQLX$ = SQLX$ + "AND DescriMov = '" & BANUFEA$ & "' "
             SQLX$ = SQLX$ + "AND ABS(ValAbsComp - " & IPASX$ & ") < 0.005 "
             '
             .FindFirst SQLX$
             If .NoMatch = False Then
               .Edit
               !FechAcred = CVDAT(NUEFECHA%)
                 BANUFEX$ = BANUFEA$
                 Call REPLA(BANUFEX$, NUFECHX$, 33, 8)
               !Descrimov = BANUFEX$
               .Update
             End If
             '
             SQLY$ = "TipoMovim = 'CHENTR' "
             SQLY$ = SQLY$ + "AND DescriMov = '" & BANUFEA$ & "' "
             SQLY$ = SQLY$ + "AND  MID$(CodiCom_Cor,11,2) <> 'nl'"
             '
             .FindFirst SQLY$
60           If .NoMatch = False Then
               .Edit
                 BANUFEX$ = BANUFEA$
                 Call REPLA(BANUFEX$, NUFECHX$, 33, 8)
               !Descrimov = BANUFEX$
               .Update
               .FindNext SQLY$
               GoTo 60
             End If
           End With
           '
         End If
       End If
     End If
   End If
   '
   GoTo 30
   '
End Sub
'

Sub CARGACHECAR()
   '
10 CHENUE$ = OBJPLA$("CARGACHEQUE", CHENUE$)
   If CHENUE$ = "" Then Exit Sub
   '
   REFE$ = TRIM$(Left$(CHENUE$, 30))
   If REFE$ <> "" Then
     If CVD(Mid$(CHENUE$, 49, 8)) > 0.01 Then
       If CVI(Mid$(CHENUE$, 57, 2)) > 0 Then
         If CVI(Mid$(CHENUE$, 59, 2)) > 0 Then
           If CVI(Mid$(CHENUE$, 61, 2)) > 0 Then
             If TRIM$(Mid$(CHENUE$, 63, 10)) <> "" Then
               If TRIM$(Mid$(CHENUE$, 73, 18)) <> "" Then
                 RECIDE$ = TRIM$(Mid$(CHENUE$, 73, 17)) + " (Ing.Manual)"
                 Call REPLA(CHENUE$, RECIDE$, 73, 30)
                 FEX = CVI(Mid$(CHENUE$, 59, 2))
                 FEVE$ = FECHATEX$(FEX)
                 Call REPLA(CHENUE$, FEVE$, 33, 8)
                 Call REGAPP("BADACHE", CHENUE$)
                 Call CIERRARCH("BADACHE")
                 CHENUE$ = ""
                 Call COMUNI("Carga Completa")
                 GoTo 10
               End If
             End If
           End If
         End If
       End If
     End If
   End If
   '
   Call MENSERR(24, "Datos Incompletos !!!\  \Corrija o Agregue los Faltantes.")
   GoTo 10
   '
End Sub
'

Sub ENTREGACHE()
   '
30 NCHE$ = OBJPLA$("CARNUCHE", NCHE$)
   If NCHE$ = "" Then
      Exit Sub
   End If
   '
   NCHE$ = TRIM$(NCHE$)
   Screen.MousePointer = 11
   Call FRESHECHO("CHESELEC")
   JJ& = 0
   FIN& = ULTREG&("BADACHE")
   '
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      X$ = REGLEIDO$("BADACHE", I&)
      NCHX$ = Mid$(X$, 19, 12)
      If CVI(Mid$(X$, 133, 2)) < 1 Then  'aun en cartera
         '
         If NCHE$ = "" Then
            If TRIM$(NCHX$) = "" Then
               Call REPLA(X$, MKS$(I&), 45, 4)
               JJ& = JJ& + 1
               Call GRAREG("CHESELEC", X$, JJ&)
            End If
            GoTo 39
         End If
         '
         If InStr(NCHX$, NCHE$) > 0 Then
            Call REPLA(X$, MKS$(I&), 45, 4)
            JJ& = JJ& + 1
            Call GRAREG("CHESELEC", X$, JJ&)
         End If
         '
      End If
39 Next I&
   '
   Call SETULTREG("CHESELEC", JJ&)
   Call CIERRARCH("CHESELEC")
   Screen.MousePointer = 1
   '
   CASEL& = ULTREG&("CHESELEC")
   If CASEL& < 1 Then
         Call COMUNI("Número no Encontrado en Cartera Actual.")
         GoTo 30
       ElseIf CASEL& = 1 Then
         Y$ = REGLEIDO$("CHESELEC", 1&)
         GoTo 50
   End If
   '
   Y$ = REGSEL$("CHESELEC")
   If Len(Y$) <= 4 Then GoTo 30
   '
50 HOII% = HOY(HOS$)
   REFESAL$ = MKI$(HOII%) + Space$(120)
   Call REPLA(REFESAL$, "Salida Manual de Cartera.", 3, 40)
   Call REPLA(REFESAL$, HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 43, 40)
   Call REPLA(REFESAL$, "Usuario: " + TRIM$(USERNAME$), 83, 40)
   Call REPLA(Y$, REFESAL$, 133, 122)
   '
60 X$ = OBJPLA$("MUESTRACHEQUE", Y$)
   If X$ <> "" Then
      If TRIM$(Mid$(X$, 103, 30)) = "" Then
         Call MENSERR(24, "Falta Indicar Destino del Valor.")
         GoTo 60
      End If
      '
      If Mid$(X$, 135, 120) <> Mid$(REFESAL$, 3, 120) Then
         Call MENSERR(24, "Referencias no Modificables")
         GoTo 60
      End If
      '
      If CVI(Mid$(X$, 133, 2)) > 0 Then
         REG& = CVS(Mid$(X$, 45, 4))
         Call REPLA(X$, String$(4, 0), 45, 4)
         If REG& > 0 Then
           If REG& <= ULTREG&("BADACHE") Then
             Call GRAREG("BADACHE", X$, REG&)
           End If
         End If
      End If
   End If
   GoTo 30
   '
End Sub
'\\\OT-05-0378-WAR-06/06/05///
Sub LICHFERE()
   '
   HAS% = HOY(HOS$)
   FEX = HAS%
   HASS$ = FECHATEX$(FEX)
   DESS$ = "01/01/80"
   DES% = FECHANUM(DESS$)
   VDEF$ = MKI$(DES%) + MKI$(HAS%) + MKI$(0)
   '
   OBX$ = OBJPLA$("DESHASFECHCLI", VDEF$)
10 'OBX$ = OBJPLA$("DESHESFECHREC", VDEF$)
   If OBX$ = "" Then
      Exit Sub
   End If
   DES% = CVI(Left$(OBX$, 2))
   DESDEFE = CDbl(CVDAT(DES%))
   HAS% = CVI(Mid$(OBX$, 3, 2))
   HASTAFE = CDbl(CVDAT(HAS%))
   CLISEL% = CVI(Mid$(OBX$, 5, 2))
   '
   Screen.MousePointer = 11
   ASALID$ = "BADACHE1"
   '
   Call APERBASDAT("CABAN")
   If CLISEL% < 1 Then
      SQLX$ = "SELECT * FROM CAJABANC "
      SQLX$ = SQLX$ + "where TipoMovim = 'CHRECI' "
      SQLX$ = SQLX$ + "AND STATUS >= 1 AND STATUS <=2 "
      SQLX$ = SQLX$ + "AND CDBL(FechConta) >= " & DESDEFE & " "
      SQLX$ = SQLX$ + "AND CDBL(FechConta) <= " & HASTAFE & " "
      SQLX$ = SQLX$ + "ORDER BY FechAcred ASC;"
   Else
      SQLX$ = "SELECT * FROM CAJABANC "
      SQLX$ = SQLX$ + "where NuClien = " & CLISEL% & " "
      SQLX$ = SQLX$ + "AND TipoMovim = 'CHRECI' "
      SQLX$ = SQLX$ + "AND STATUS >= 1 AND STATUS <=2 "
      SQLX$ = SQLX$ + "AND CDBL(FechConta) >= " & DESDEFE & " "
      SQLX$ = SQLX$ + "AND CDBL(FechConta) <= " & HASTAFE & " "
      SQLX$ = SQLX$ + "ORDER BY FechAcred ASC;"
   End If
    '
   Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   JJ& = 0
   Z0$ = REGBLAN$(ASALID$)
   With CABATEMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do While Not .EOF
         If TRIM(!OpciMovim) <> "CH" Then GoTo 19 ' no es un cheque
         '
         FEACRED% = CVINT(!FechAcred)
         BANUFE$ = !Descrimov
         ICHEQ# = !ValAbsComp
         '
         If TRIM$(Left$(BANUFE$, 12)) = "" Then
           Call REPLA(BANUFE$, "Cheque Nro:", 1, 11)
         End If
         '
         BANC1$ = Mid$(BANUFE$, 1, 18)
         NUMECH1$ = Mid$(BANUFE$, 19, 14)
         FEVTO1$ = Mid$(BANUFE$, 33, 8)
         '
         FEX = FEACRED%
         FEMISI% = CVINT%(!FechEmisi)
         FERECEP% = CVINT(!FechConta)
         NURECIB$ = !CodiCom_Cor
         ORIGCHE$ = !REFECOM
           DESTICHE$ = ""
           On Error Resume Next
         DESTICHE$ = !DESTICHEQ
           On Error GoTo 0
         FENDEP% = CVINT(!FEDEPOSENT)
         '
         Z$ = Z0$
         Call REPLA(Z$, BANC1$, 1, 18)
         Call REPLA(Z$, NUMECH1$, 19, 14)
         Call REPLA(Z$, FEVTO1$, 33, 8)
         Call REPLA(Z$, MKD$(ICHEQ#), 41, 8)
         Call REPLA(Z$, MKI$(FEMISI%), 49, 2)
         Call REPLA(Z$, FEVTO1$, 51, 8)
         'Call REPLA(Z$, MKI$(FEACRED%), 51, 2)
         Call REPLA(Z$, MKI$(FERECEP%), 59, 2)
         Call REPLA(Z$, NURECIB$, 61, 10)
         Call REPLA(Z$, ORIGCHE$, 71, 30)
         Call REPLA(Z$, DESTICHE$, 101, 30)
         Call REPLA(Z$, MKI$(FENDEP%), 131, 2)
         JJ& = JJ& + 1
         Call GRAREG(ASALID$, Z$, JJ&)
         '
19     .MoveNext
       Loop
     End If
   End With
   '
   Call SETULTREG(ASALID$, JJ&)
   Call FILESORT(ASALID$, ASALID$, 5, 5, "ASC")
   Call CIERRARCH(ASALID$)
   Screen.MousePointer = 1
   '
   Call FINAL("*LICHEFRE")
   '
End Sub

Sub CHEQPENAC()
     '
     If Option5.Value = True Then
        Call COMUNI("Solo Disponible como Listado")
        Option4.Value = True
     End If
     '
     Call SELRACLIEN(NCL1%, NCL2%, VENDSEL%)
     If (NCL1% + NCL2%) < 1 Then Exit Sub
     '
     RANCLI$ = "Rango Clientes: " + TRIM$(Str$(NCL1%)) + " - " + TRIM$(Str$(NCL2%))
     If VENDSEL% > 0 Then
        RANCLI$ = RANCLI$ + " ..... Vendedor: " + ECOARCH$("VENDEDOR", Chr$(VENDSEL%))
     End If
     '
     Call DESHASFECHA(DESD%, HAST%, PERIOD$)
     FEXI1% = DIANTE(DIANTE(DESD%))
     FEXI2% = DIANTE(DIANTE(HAST%))
     FEXI% = DIANTE(DIANTE(HOY(HOS$)))
     DESDE = CDbl(CVDAT(FEXI1%))
     HASTA = CDbl(CVDAT(FEXI2%))
     Call APERBASDAT("CABAN")
     SQLX$ = "SELECT * FROM CAJABANC "
     SQLX$ = SQLX$ + "where TipoMovim = 'CHRECI' "
     SQLX$ = SQLX$ + "AND Status >0 "
     SQLX$ = SQLX$ + "AND CDBL(FechAcred) >= " & DESDE & " "
     SQLX$ = SQLX$ + "AND CDBL(FechAcred) <= " & HASTA & " "
     SQLX$ = SQLX$ + "AND NuClien BETWEEN " & NCL1% & " AND " & NCL2% & " "
     SQLX$ = SQLX$ + "ORDER BY NuClien ASC, FechAcred ASC;"
     
     Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
     With CABATEMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
         CONT& = 0
         Do While Not .EOF
           If TRIM(!OpciMovim) <> "CH" Then GoTo 19 ' no es un cheque
             CUECLI% = !NuClien
             ' Selección por Vendedor
             If VENDSEL% > 0 Then
                If VENDCLI%(CUECLI%) <> VENDSEL% Then
                   GoTo 19
                End If
             End If
             '
             FEACRED% = CVINT(!FechAcred)
             BANUFE$ = !Descrimov
             FEEMI% = CVINT(!FechEmisi)
             FEREC% = CVINT(!FechConta)
             DOCU$ = !CodiCom_Cor
             ICHEQ# = !ValAbsComp
           
             Call REPLA(X$, MKI$(CUECLI%), 1, 2)
             Call REPLA(X$, MKI$(FEACRED%), 3, 2)
             Call REPLA(X$, BANUFE$, 5, 40)
             Call REPLA(X$, MKI$(FEEMI%), 45, 2)
             Call REPLA(X$, MKI$(FEREC%), 47, 2)
             Call REPLA(X$, DOCU$, 49, 12)
             Call REPLA(X$, MKD$(ICHEQ#), 61, 8)
           
             CONT& = CONT& + 1
     
             Call GRAREG("CHEPEAC", X$, CONT&)
19              .MoveNext
         Loop
       End If
     End With
     Call SETULTREG("CHEPEAC", CONT&)
     Call HEADERS("CHEPEAC", "")
     Call HEADERS("CHEPEAC", RANCLI$)
     Call HEADERS("CHEPEAC", "Periodo: " + PERIOD$)
     '
     Call FINAL("*CHEPEAC")
     '
End Sub



'\\\OT-05-0378-WAR-FIN///
Sub LICHECAR()
   '
   If DERACCE%("VECARCHE", "Visualizar Cartera de Cheques y Valores") < 2 Then Exit Sub
   '
   OPX% = 1
   If Option4.Value = True Then OPX% = 2
   '
   Call GENLICHECA(OPX%)
   '
   If OPX% = 1 Then
        VTB% = VENTABU%("COCARCHE")
      ElseIf OPX% = 2 Then
        If EXISTE%(LUDAT$ + "LISCARTE.RFS") > 0 Then
          If EXISTE%("LISCARTE.PRF") > 0 Then
            GoTo 50       ' OTRO FORMATO DE LISTADO
          End If
        End If
        Call FINAL("*SOCARCHE")
   End If
   Exit Sub
   '
50 URE& = ULTREG&("LICARCHE")
   For KKU& = 1 To URE&
     XX$ = REGLEIDO$("LICARCHE", KKU&)
     YY$ = Mid$(XX$, 113, 6) + Left$(XX$, 112) + Mid$(XX$, 119)
     Call GRAREG("LISCARTE", YY$, KKU&)
   Next KKU&
   Call SETULTREG("LISCARTE", URE&)
   Call FILESORT("LISCARTE", "", 5, 6, "ASC")
   Call FINAL("*LISCARTE")
   '
End Sub

Sub GENLICHECA(OPX%)
   '
   HAS% = HOY(HOS$)
   FEX = HAS%
   HASS$ = FECHATEX$(FEX)
   DESS$ = "01/01/80"
   DES% = FECHANUM(DESS$)
   VDEF$ = MKI$(DES%) + MKI$(HAS%)
   '
   If OPX% = 3 Then    ' LLAMADO DESDE CONCILIACION DE CARTERA
     OBX$ = VDEF$
     GoTo 12
   End If
   '
10 OBX$ = OBJPLA$("DESHESFECHREC", VDEF$)
   If OBX$ = "" Then
      OPX% = (-1)
      Exit Sub
   End If
   '
12 DES% = CVI(Left$(OBX$, 2))
   DESDEFE = CDbl(CVDAT(DES%))
   HAS% = CVI(Mid$(OBX$, 3, 2))
   HASTAFE = CDbl(CVDAT(HAS%))
   '
   Screen.MousePointer = 11
   ASALID$ = "LICARCHE"
   If OPX% = 1 Or OPX% = 3 Then ASALID$ = "!" + ASALID$
   '
15 Call APERBASDAT("CABAN")
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "where TipoMovim = 'CHRECI' "
   SQLX$ = SQLX$ + "AND Status = 1 "
   SQLX$ = SQLX$ + "AND CDBL(FechConta) >= " & DESDEFE & " "
   SQLX$ = SQLX$ + "AND CDBL(FechConta) <= " & HASTAFE & " "
   SQLX$ = SQLX$ + "ORDER BY FechAcred ASC;"
   '
   Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   JJ& = 0: TRANSPOR# = 0
   Z0$ = REGBLAN$(ASALID$)
   With CABATEMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do While Not .EOF
         If TRIM$(!OpciMovim) <> "CH" Then GoTo 19 ' no es un cheque
         If IsNull(!FECHENTRE) = False Then
           If CVINT(!FECHENTRE) > 33 Then
             GoTo 19    ' salio de cartera
           End If
         End If
         '
         FEACRED% = CVINT(!FechAcred)
         BANUFE$ = !Descrimov
         ICHEQ# = !ValAbsComp
         '
         If TRIM$(Left$(BANUFE$, 12)) = "" Then
           Call REPLA(BANUFE$, "Cheque Nro:", 1, 11)
         End If
         '
         FEACRED% = CVINT%(!FechAcred)
         FEX = FEACRED%
         MESA$ = Mid$(FECHATEX$(FEX), 4, 5)
         FEMISI% = CVINT%(!FechEmisi)
         FERECEP% = CVINT(!FechConta)
         NURECIB$ = !CodiCom_Cor
         ORIGCHE$ = !REFECOM
         '
         Z$ = Z0$
         Call REPLA(Z$, BANUFE$, 1, 40)
         Call REPLA(Z$, MKD$(ICHEQ#), 49, 8)
         Call REPLA(Z$, MKI$(FEMISI%), 57, 2)
         Call REPLA(Z$, MKI$(FEACRED%), 59, 2)
         Call REPLA(Z$, MKI$(FERECEP%), 61, 2)
         Call REPLA(Z$, NURECIB$, 63, 12)
         Call REPLA(Z$, ORIGCHE$, 75, 38)
         Call REPLA(Z$, MESA$, 113, 6)
         Call REPLA(Z$, MKD$(0), 119, 8)
         NORDEN% = 1: If OPX% = 1 Then NORDEN% = JJ& + 1
         Call REPLA(Z$, MKI$(NORDEN%), 127, 2)
         JJ& = JJ& + 1
         Call GRAREG(ASALID$, Z$, JJ&)
         '
19     .MoveNext
       Loop
     End If
   End With
   '
   Call SETULTREG(ASALID$, JJ&)
   Call CIERRARCH(ASALID$)
   Screen.MousePointer = 1
   '
End Sub
'
Sub LICHEQEM()
   '////***ESTE CODIGO ES NUEVO***//////
   Dim CCUE%(1024)
   '
   Screen.MousePointer = 11
   Call COPYSTRU("CHEQUEMI", "CHEQUEMJ")
   Call COPYSTRU("GRUCOVEN", "ICUEBAN")
   ASALID$ = "LICHEQEM"
   '
   HOII% = HOY(HOS$)
   JJ& = 0
   FIN& = ULTREG&("CAJABANC")
   Screen.MousePointer = 1
   '
   Call SELECHO("SELFECHA/Emitidos al")
   FF11$ = VALACT1$("SELFECHA")
   If FF11$ = "" Then GoTo 99
   DESEM% = FECHANUM(FF11$)
   DES# = CDbl(CVDAT(DESEM%))
   DESDA = CVDAT(DESEM%)
   '
   OPX% = COMALTER%("Criterio de Selección:\\Solo Propios\Terceros\Todos Juntos\Cancelar")
   If OPX% < 1 Or OPX% > 3 Then GoTo 99
   '
   'leer cuentas bancarias y armar indice
   ICUEB% = 0
   CUECAR$ = TRIM$(CUECARCHE$)
   YACARCHE% = 0
   '
   If OPX% = 2 Then GoTo 20 ' SOLO CHEQUES DE TERCEROS
   '
   For UB& = 1 To ULTREG&("CUEBANC")
     RECUEBA$ = REGLEIDO$("CUEBANC", UB&)
     CUEXTE$ = TRIM$(Left$(RECUEBA$, 12))
     DENOBA$ = TRIM$(Mid$(RECUEBA$, 13, 48))
     If CUEXTE$ <> "" Then
       If CUEXTE$ = CUECAR$ Then GoTo 19 ' QUE NO CARGUE CUENTA CARTERA
       '
       CUINTE% = CUECOINT%(CUEXTE$)
       If CUINTE% > 0 Then
         ICUEB% = ICUEB% + 1
         TZX$ = MKI$(ICUEB%) + DENOBA$ + Space$(32)
         REGB& = ICUEB%
         Call GRAREG("ICUEBAN", TZX$, REGB&)
         CCUE%(ICUEB%) = CUINTE%
       End If
     End If
19 Next UB&
   '
   If OPX% = 1 Then GoTo 30  ' SOLO CHEQUES PROPIOS
   '
20 If YACARCHE% = 0 Then
      ICUEB% = ICUEB% + 1
      REGB& = ICUEB%
      CUINTE% = CUECOINT%(CUECAR$)
      TTXX1$ = MKI$(ICUEB%)
      TTXX2$ = DENOCUE$(CUECOINT%(CUECAR$))
      TZX$ = TTXX1$ + TTXX2$ + Space$(32)
      Call GRAREG("ICUEBAN", TZX$, REGB&)
      CCUE%(ICUEB%) = CUINTE%
   End If
   '
30 Call SETULTREG("ICUEBAN", REGB&)
   Call FRESHECHO("ICUEBAN")
   '
   Screen.MousePointer = 11
   Call APERBASDAT("CABAN")
   SQLX$ = " SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + " WHERE CDBL(FechAcred)>= " & DES# & " "
   ' OJO, TIENE QUE QUEDAR ASÍ CON EL CDBL, PORQUE SI NO NO FILTRA CORRECTAMENTE - EPB - 22/06/06
   SQLX$ = SQLX$ + " AND TIPOMOVIM = 'CHENTR' "
   SQLX$ = SQLX$ + " ORDER BY FechAcred "
   Set CHTMP = CueCorri.OpenRecordset(SQLX$, 4)
   With CHTMP
     On Error Resume Next
     .MoveLast
     If Err Then
       On Error GoTo 0
       Call COMUNI("No hay Movimientos para la Fecha Seleccionada.")
       Exit Sub
     End If
     On Error GoTo 0
     FIN& = .RecordCount
     '
     .MoveFirst
     KU& = 0
     Do While .EOF = False
       KU& = KU& + 1
       Call TRACE(20, KU&, FIN&)
       '
       For UB& = 1 To REGB&
          CUINTE% = CCUE%(UB&)
          If !CueContaI = CUINTE% Then
             ICUEB% = UB&
             FACRE% = CVINT(!FechAcred)
             PROV% = !NuProve 'Abs(CVI(Mid$(X$, 67, 2)))
             '
             If PROV% > 0 Then
               NUMERO$ = TRIM$(Mid$(!Descrimov, 17, 16)) 'TRIM$(Mid$(X$, 89, 14))
               banco1$ = TRIM$(Left$(!Descrimov, 16)) 'TRIM$(Mid$(X$, 71, 18))
               EMIS% = CVINT(!FechEmisi) 'CVI(Mid$(X$, 111, 2))   ' EMISION=FECHA O/PAGO
               FENTREG% = CVINT(!FECHENTRE) 'CVI(Mid$(X$, 61, 2))
               If EMIS% < 1 Then
                 EMIS% = FENTREG%
               End If
               IMPO# = !IHABECOMP 'IMPEX#
               VENCI% = FACRE%
               'NOPA& = Val(Mid$(X$, 4, 6))
               '
               Y$ = REGBLAN$("LICHEQEM")
               Call REPLA(Y$, MKI$(VENCI%), 1, 2)
               Call REPLA(Y$, banco1$, 3, 16)
               Call REPLA(Y$, NUMERO$, 19, 12)
               Call REPLA(Y$, MKD$(IMPO#), 31, 8)
               Call REPLA(Y$, MKI$(EMIS%), 39, 2)
               Call REPLA(Y$, MKI$(PROV%), 41, 2)
               '
               Z$ = REGBLAN$("CHEQUEMJ")
               Call REPLA(Z$, MKI$(PROV%), 1, 2)
               Call REPLA(Z$, NUMERO$, 3, 12)
               Call REPLA(Z$, banco1$, 15, 16)
               Call REPLA(Z$, MKI$(EMIS%), 31, 2)
               Call REPLA(Z$, MKD$(IMPO#), 33, 8)
               NOPA& = Val(!NUMECOMP)
               Call REPLA(Z$, MKS$(NOPA&), 41, 4)
               Call REPLA(Z$, MKI$(ICUEB%), 45, 2)
               'Call REPLA(Z$, MKI$(FENTREG%), 47, 2)
               Call REPLA(Z$, MKI$(EMIS%), 47, 2)
               Call REPLA(Z$, MKI$(FACRE%), 49, 2)
               '
               JJ& = JJ& + 1
               Call GRAREG(ASALID$, Y$, JJ&)
               Call GRAREG("CHEQUEMJ", Z$, JJ&)
               '
             End If
             '
          End If
       Next UB&
       '
     .MoveNext
     Loop
   End With
   Call SETULTREG(ASALID$, JJ&)
   Call CIERRARCH(ASALID$)
   Call SETULTREG("CHEQUEMJ", JJ&)
   Call CIERRARCH("CHEQUEMJ")
   Screen.MousePointer = 1
   '
   If Option5.Value = True Then
       Call CONECRUN("*COCHEMJ", "Consulta Cheques Emitidos")
     Else
       Call CONECRUN("*LICHEMI", "Listado de Cheques Emitidos")
   End If
   '
99 End Sub

Sub CHERECHAZA()
   '
   Dim TOFAC#(10), IMV#(10), F%(10), GRU%(10), IPAR#(10), DATCLI$(32)
   Dim CTA%(128), IMPO#(128), DETA$(128)

10 NCHE$ = OBJPLA$("CARNUCHE", NCHE$)
   If NCHE$ = "" Then Exit Sub
   NCHE$ = TRIM$(NCHE$)
   '
   Screen.MousePointer = 11
   '
   Call BLANARCH("!CHESELEC")
   FIN& = ULTREG&("BADACHE")
   '
   For I& = 1 To FIN&
     Call TRACE(20, I&, FIN&)
     X$ = REGLEIDO$("BADACHE", I&)
     If CVI(Mid$(X$, 255, 2)) < 1 Then  ' SALTEA RECHAZADOS
       NCHX$ = Mid$(X$, 19, 12)
        '
        If NCHE$ = "" Then
          If TRIM$(NCHX$) = "" Then
            Call REPLA(X$, MKS$(I&), 45, 4)
            Call REGAPP("!CHESELEC", X$)
          End If
          GoTo 19
        End If
        '
        If InStr(NCHX$, NCHE$) > 0 Then
          Call REPLA(X$, MKS$(I&), 45, 4)
          Call REGAPP("!CHESELEC", X$)
        End If
      End If
19 Next I&
   '
   Call CIERRARCH("!CHESELEC")
   Screen.MousePointer = 1
   '
   CASEL& = ULTREG&("!CHESELEC")
   If CASEL& < 1 Then
         Call COMUNI("Número no Encontrado")
         Exit Sub
       ElseIf CASEL& = 1 Then
         Y$ = REGLEIDO$("!CHESELEC", 1&)
         GoTo 25
   End If
   '
   Call FRESHECHO("!CHESELEC")
   Y$ = REGSEL$("!CHESELEC")
   If Len(Y$) <= 4 Then GoTo 10
   '
25 X$ = OBJPLA$("MUESTRACHEQUE", Y$)
   If X$ = "" Then GoTo 10
   '
   If CVI(Mid$(Y$, 133, 2)) < 1 Then
       Call COMUNI("Cheque Aún en Cartera")
       GoTo 10
   End If
   '
   ICHE# = CVD(Mid$(Y$, 49, 8))
   REFECHE$ = TRIM$(TRIM$(Left$(Y$, 16)) + " " + TRIM$(Mid$(Y$, 19, 12)))
   NCHX$ = Mid$(Y$, 19, 12)
   REBADA& = CVS(Mid$(Y$, 45, 4))
   '
30 PROVE$ = Mid$(Y$, 103, 30)
   For U& = 1 To ULTREG&("PROVED1")
      Z$ = REGLEIDO$("PROVED1", U&)
      If Mid$(Z$, 3) = PROVE$ Then
         PROV% = CVI(Left$(Z$, 2))
         GoTo 31
      End If
   Next U&
   PROV% = 0
   '
31 CLIEN$ = Mid$(Y$, 73, 30)
   For U& = 1 To ULTREG&("DEUDOR1")
      Z$ = REGLEIDO$("DEUDOR1", U&)
      If Mid$(Z$, 3) = CLIEN$ Then
         CLI% = CVI(Left$(Z$, 2))
         GoTo 33
      End If
   Next U&
   CLI% = 0
   '
33 modo% = COMALTER%("Indique el Origen del Rechazo:\\Proveedor\Banco\Abandonar")
   On modo% GoTo 64, 70
   GoTo 10
   '
64 HO% = HOY(HOS$)
   VDEF$ = MKI$(PROV%) + MKI$(CLI%) + MKI$(HO%)
   Z$ = OBJPLA$("DEVUCHEPRO", VDEF$)
   If Z$ = "" Then GoTo 10
   '
   PROV% = CVI(Left$(Z$, 2)): If PROV% < 1 Then GoTo 64
   CLI% = CVI(Mid$(Z$, 3, 2)): If CLI% < 1 Then GoTo 64
   FF% = CVI(Mid$(Z$, 5, 2))
   If FF% < COMEJ% Or FF% > HO% Then GoTo 64
   '
   RGXX& = REGICLI&(CLI%)
   If RGXX& < 1 Or RGXX& > ULTREG&("DEUDOR1") Then
      Call MENSERR(24, "Imposible Registrar Rechazo.\Cliente no Encontrado.")
      GoTo 10
   End If
   '
   RGYY& = REGICLI&((-1) * PROV%)
   If RGYY& < 1 Or RGYY& > ULTREG&("PROVED1") Then
      Call MENSERR(24, "Imposible Registrar Rechazo.\Proveedor no Encontrado.")
      GoTo 10
   End If
   '
   Call BLOQARCH("FACTUR")
   Call BLOQARCH("SVD202")
   Call BLOQARCH("DEUDOR2")
   Call BLOQARCH("CUECORR")
   Call BLOQARCH("PROVED2")
   '
   OKEY% = 0
   If CONTROL$("", "CHERECHA") = "NOTADEBI" Then
        CLIX% = CLI%
        Call COMUNI("Recuerde Hacer la NOTA DE DEBITO POR RECHAZO\en la Cuenta " + TRIM$(Str$(CLIX%)) + " - " + TRIM$(RASOCLI$(CLIX%)))
        OKEY% = 1
      Else
        Call GRAJUCLI(CLI%, FF%, ICHE#, OKEY%)
   End If
   Call GRAJUPRO(PROV%, FF%, ICHE#, NCHX$, OKEY%)
   If OKEY% = 2 Then
        Call COMUNI("Rechazo Registrado")
      Else
        Call COMUNI("Rechazo mal Registrado.\Consulte !!!")
   End If
   If REBADA& > 0 Then
     If REBADA& <= ULTREG&("BADACHE") Then
       X$ = REGLEIDO$("BADACHE", REBADA&)
       Call REPLA(X$, MKI$(FF%), 255, 2)
       Call GRAREG("BADACHE", X$, REBADA&)
     End If
   End If
   Call CIERRARCH("BADACHE")
   '
   GoTo 10
   '
70 HO% = HOY(HOS$)
   CLIEN$ = Mid$(Y$, 73, 30)
   For U& = 1 To ULTREG&("DEUDOR1")
      Z$ = REGLEIDO$("DEUDOR1", U&)
      If Mid$(Z$, 3) = CLIEN$ Then
         CLI% = CVI(Left$(Z$, 2))
         GoTo 71
      End If
   Next U&
   CLI% = 0
   '
71 VDEF$ = Space$(24) + MKI$(HO%) + MKI$(CLI%)
   CUECHERA$ = AJUSTI$(CONTROL$("", "CUECHERA"), 12)
   Call REPLA(VDEF$, CUECHERA$, 13, 12)
72 Z$ = OBJPLA$("RECHACHEBA", VDEF$)
   VDEF$ = Z$
   If Z$ = "" Then Exit Sub
   '
   CLI% = CVI(Mid$(Z$, 27, 2)): If CLI% < 1 Then GoTo 72
   FF% = CVI(Mid$(Z$, 25, 2)): If FF% < 1 Then GoTo 72
   '
   For II% = 1 To 128
       CTA%(II%) = 0
       IMPO#(II%) = 0
       DETA$(II%) = ""
   Next II%
   '
   CUEBANCO$ = Left$(Z$, 12): CTA%(2) = 0
   CUECHERA$ = Mid$(Z$, 13, 12): CTA%(1) = 0
   '
   For U& = 1 To ULTREG&("ECUECON")
      X$ = REGLEIDO$("ECUECON", U&)
      If Left$(X$, 12) = CUEBANCO$ Then CTA%(2) = CVI(Mid$(X$, 41, 2))
      If Left$(X$, 12) = CUECHERA$ Then CTA%(1) = CVI(Mid$(X$, 41, 2))
   Next U&
   '
   FF% = CVI(Mid$(Z$, 25, 2)): If FF% < 1 Then Exit Sub
   If CTA%(1) < 1 Then GoTo 72
   If CTA%(2) < 1 Then GoTo 72
   Call GRACONTROL("", "CUECHERA", CUECHERA$)
   '
   IMPO#(1) = ICHE#
   IMPO#(2) = (-1) * ICHE#
   DETA$(1) = REFECHE$
   DETA$(2) = "Cheque Rechazado"
   REFE$ = "Ch.Rech." + REFECHE$
   CONTA% = 2
   '
   FEX = FF%
   If FF% < COMEJ% Then
      Call MENSERR(24, "Fecha Anterior al Comienzo del Ejercicio")
      VDEF$ = Z$
      GoTo 72
   End If
   If FF% > FINEJ% Then
      Call MENSERR(24, "Fecha Posterior al Cierre del Ejercicio")
      VDEF$ = Z$
      GoTo 72
   End If
   '
   If FF% >= FULTAS% Then
      NROAS = Int(ULTAS + 1)
      GoTo 90
   End If
   '
   Screen.MousePointer = 11
   For UI% = MAXENC% To 2 Step -1
      Call LEENCAJA(UI%)
      NAS2 = CVS(D1$): FAS2% = CVI(D3$)
      If FAS2% < FF% Then GoTo 80
      Call LEENCAJA(UI% - 1)
      NAS1 = CVS(D1$): FAS1% = CVI(D3$)
      If FAS1% <= FF% Then
        If FAS2% >= FF% Then
          If (Int(10 * NAS2 + 0.5) - Int(10 * NAS1 + 0.5)) >= 2 Then
            NROAS = (1 + Int(10 * NAS1 + 0.5)) / 10
            Screen.MousePointer = 1
            GoTo 90
          End If
        End If
      End If
   Next UI%
   '
80 Screen.MousePointer = 1
   Call MENSERR(24, "Imposible Registrar Rechazo\en la Fecha Elegida.\  \Fecha Propuesta " + FECHATEX$(FEX))
   VDEF$ = Z$
   Call REPLA(Z$, MKI$(FULTAS%), 25, 2)
   GoTo 72
   '
90 Call GRASICAJA(NROAS, FF%, REFE$, CONTA%, CTA%(), IMPO#(), DETA$())
   If CONTROL$("", "CHERECHA") = "NOTADEBI" Then
        CLIX% = CLI%
        Call COMUNI("Recuerde Hacer la NOTA DE DEBITO POR RECHAZO\en la Cuenta " + TRIM$(Str$(CLIX%)) + " - " + TRIM$(RASOCLI$(CLIX%)))
      Else
        Call GRAJUCLI(CLI%, FF%, ICHE#, OKEY%)
   End If
   If REBADA& > 0 Then
     If REBADA& <= ULTREG&("BADACHE") Then
       X$ = REGLEIDO$("BADACHE", REBADA&)
       Call REPLA(X$, MKI$(FF%), 255, 2)
       Call GRAREG("BADACHE", X$, REBADA&)
     End If
   End If
   Call CIERRARCH("BADACHE")
   '
   Call COMUNI("Rechazo Cheque " + REFECHE$ + "\Registrado y Contabilizado.")
   GoTo 10
   '
End Sub

Sub GRAJUCLI(CLI%, FF%, ICHE#, OKEY%)
   '
   RGXX& = REGICLI&(CLI%)
   If RGXX& < 1 Or RGXX& > ULTREG&("DEUDOR1") Then
      Call MENSERR(24, "Imposible Registrar Rechazo.\Cliente no Encontrado.")
      Exit Sub
   End If
   '
   NAJU& = 1 + ULTREG&("FACTUR")
   '
   Z$ = REGBLAN$("FACTUR")
   Call REPLA(Z$, MKI$(CLI%), 1, 2)
   Call REPLA(Z$, MKS$(NAJU&), 3, 4)
   Call REPLA(Z$, MKI$(FF%), 7, 2)
   Call REPLA(Z$, MKI$(FF%), 9, 2)
   Call REPLA(Z$, MKD$(ICHE#), 15, 8)
   Call REPLA(Z$, MKD$(ICHE#), 39, 8)
   Call REGAPP("FACTUR", Z$)
   '
   Z$ = REGBLAN$("SVD202")
   Call REPLA(Z$, MKI$(CLI%), 1, 2)
   Call REPLA(Z$, MKS$(NAJU&), 5, 4)
   Call REPLA(Z$, MKI$(FF%), 9, 2)
   Call REPLA(Z$, MKD$(ICHE#), 11, 8)
   Call REPLA(Z$, MKI$(FF%), 19, 2)
   Call REPLA(Z$, MKD$(ICHE#), 21, 8)
   Call REGAPP("SVD202", Z$)
   '
   TTXS$ = REGLEIDO$("DEUDOR2", RGXX&)
   SALNUE# = (Int(100 * ICHE# + 0.5)) / 100 + (Int(100 * CVD(Mid$(TTXS$, 109, 8)) + 0.5)) / 100
   Call REPLA(TTXS$, MKD$(SALNUE#), 109, 8)
   Call GRAREG("DEUDOR2", TTXS$, RGXX&)
   '
   Call CIERRARCH("FACTUR")
   Call CIERRARCH("SVD202")
   Call CIERRARCH("DEUDOR2")
   '
   OKEY% = OKEY% + 1
   '
End Sub

Sub GRAJUPRO(PROV%, FF%, ICHE#, NCHE$, OKEY%)
   '
   RGYY& = REGICLI&((-1) * PROV%)
   If RGYY& < 1 Or RGYY& > ULTREG&("PROVED1") Then
      Call MENSERR(24, "Imposible Registrar Rechazo.\Proveedor no Encontrado.")
      Exit Sub
   End If
   '
   NCHEQUE$ = Right$("00000000" + TRIM$(NCHE$), 8)
   TINUVA$ = "AS-X-ReCh-" + NCHEQUE$
   '
39 RECORD$ = MKI$(FF%)
   RECORD$ = RECORD$ + MKI$(0) + MKI$(PROV%)
   RECORD$ = RECORD$ + TINUVA$
   '
   REFE$ = Left$("Ch.Rech." + NCHE$ + Space$(24), 24)
   RECORD$ = RECORD$ + REFE$
   '
   MO$ = "$ "
   RECORD$ = RECORD$ + MO$

   TOTAL$ = MKD$(ICHE#) + String$(56, 0)
   RECORD$ = RECORD$ + TOTAL$
   RECORD$ = RECORD$ + String$(8, 0) + MKD$(ICHE#) + String$(8, 0)

   VENCI$ = MKI$(FF%) + MKD$(ICHE#) + String$(50, 0)
   RECORD$ = RECORD$ + VENCI$
   '
   IMPU$ = String$(80, 9)
   RECORD$ = RECORD$ + IMPU$
   '
   RECORD$ = RECORD$ + String$(11, 0)
   HOO = HOY(HO$)
   While HOO > 32767: HOO = HOO - 10000:   Wend
   HOI% = HOO
   RECORD$ = RECORD$ + MKI$(HOI%)
   '
   HP = 100 * Val(Left$(Time$, 2)) + Val(Mid$(Time$, 4, 2))
   While HP > 32767: HP = HP - 10000: Wend
   HPI% = HP
   RECORD$ = RECORD$ + MKI$(HPI%)
   RECORD$ = RECORD$ + Chr$(User%)
   RECORD$ = RECORD$ + String$(6, 0)
   '
   Call REGAPP("CUECORR", RECORD$)
   '
   TTXS$ = REGLEIDO$("PROVED2", RGYY&)
   SALNUE# = (Int(100 * CVD(Mid$(TTXS$, 109, 8)) + 0.5) - (Int(100 * ICHE# + 0.5))) / 100
   Call REPLA(TTXS$, MKD$(SALNUE#), 109, 8)
   Call GRAREG("PROVED2", TTXS$, RGYY&)
   '
   Call CIERRARCH("CUECORR")
   Call LIBARCH("CUECORR")
   Call CIERRARCH("PROVED2")
   Call LIBARCH("PROVED2")
   '
   OKEY% = OKEY% + 1
   '
End Sub

Sub CONCIBANC()

   Static CAMP%, CUEMP%(2048), MEPA%(2048)
   '
   Dim REFE$(2048)
   Dim FEREGIS%(2048)
   Dim FECHAMO%(2048)
   Dim IPODIF#(2048)
   '
   If CAMP% <= 0 Then
      For I& = 1 To ULTREG&("MEPAGO")
         F0$ = REGLEIDO$("MEPAGO", I&)
         CAMP% = I&
         MEPA%(CAMP%) = CVI(Left$(F0$, 2))
         CUEX$ = Mid$(F0$, 45, 12)
         CUEMP%(CAMP%) = CUEINT%(CUEX$)
         If TRIM$(CODCUE$(CUEMP%(CAMP%))) <> TRIM$(CUEX$) Then
            Call MENSERR(24, "Cuenta Contable Incorrecta\para Medio de Pago '" + TRIM$(Mid$(F0$, 3, 12)) + "'\-\Resultados Imprevisibles.")
            Call FINAL("")
         End If
     Next I&
   End If
   '
   'CMJJ% = COMEJ%
   'If CMJJ% < FECHANUM("30/09/01") Then CMJJ% = FECHANUM("30/09/01")
   CMJJ% = FECHANUM("30/09/01")
   '
   HOII% = HOY(HOS$)
   DES% = FECHANUM("01" + Mid$(HOS$, 3))
   VDEF$ = MKI$(DES%) + MKI$(HOII%)
   '
5  OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
   If OBX$ = "" Then Exit Sub
   '
   DES% = CVI(Left$(OBX$, 2))
   HAS% = CVI(Mid$(OBX$, 3, 2))
   If HAS% < DES% Then Exit Sub
   '
   If DES% < FECHANUM("01/10/01") Then
      Call MENSERR(24, "Fecha no Válida")
      GoTo 5
   End If
   '
   If HAS% > FINEJ% Then
      Call MENSERR(24, "Fecha 'Hasta' es Posterior\al Fin del Ejercicio.")
      GoTo 5
   End If
   '
  'For CUEBAL& = 1 To ULTREG&("CUEBANC")
10 CBB$ = REGSEL$("CUEBANC")
   If Len(CBB$) <= 4 Then Exit Sub
   DCBB$ = Left$(CBB$, 32)
   CUEX$ = Left$(CBB$, 12)
   CUE% = CUEINT%(CUEX$)
   If CUE% < 1 Then
      Call MENSERR(24, "Cuenta Inexistente")
      GoTo 10
   End If
   '
   Screen.MousePointer = 11
   KLM& = 0
   '
   Frame1.Refresh
   Call FRESHALL
   CUEBIN$ = MKI$(CUE%)
   CODENOCUE$ = TRIM$(CODCUE$(CUE%)) + " - " + TRIM$(DENOCUE$(CUE%))
   JJ& = 0
   SALINI# = 0
   SARRAS# = 0
   '
   '
   ' CARGAR MOVIMIENTOS DE DIARIO
   EXTEA$ = EXTE$
   If EXTEA$ = "DAT" Then
     Close #NX9%: NX9% = 0
     Close #NX10%: NX10% = 0
     Close #NX11%: NX11% = 0
     EXTE$ = "OLD"
   End If
   '
12 PMC$ = PUNMODIAR$(CUE%)
   FIN& = Len(PMC$)
   For KK% = 1 To Len(PMC$) Step 2
      UU& = KK% + 1
      Call TRACE(20, UU&, FIN&)
      REGDIA% = CVI(Mid$(PMC$, KK%, 2))
      Call LEEDIARIO(REGDIA%)
      If CVS(E12$) < 1.05 Then
        If EXTE$ = "DAT" Then
          GoTo 19 ' SALTEA ASIENTO APERTURA
        End If
      End If
      If Asc(E18$) <> 1 Then  ' SALTEA ASIENTOS RESUMEN DE CAJA
         Call LEEPASDIA(REGDIA%, 1)
         If TRIM$(DETAPAS$) = "" Then
            Call LEEPASDIA(REGDIA%, 2)
            DETAPAS$ = REFASIEN$
         End If
         MARCADIA$ = "*"
         FEMOREA% = CVI(E15$)
         IMPO# = CVD(E16$)
         If FEMOREA% < DES% Then
            If FEMOREA% >= CMJJ% Then GoSub 90 'TOTALIZAR INICIAL
          ElseIf FEMOREA% <= HAS% Then
            FEMORE% = FEMOREA%
            FEREGI% = FEMOREA%
            GoSub 95
         End If
      End If
19 Next KK%
   '
   If EXTE$ = "OLD" Then
     If EXTEA$ = "DAT" Then
       Close #NX9%: NX9% = 0
       Close #NX10%: NX10% = 0
       Close #NX11%: NX11% = 0
       EXTE$ = "DAT"
       EXTEA$ = ""
       GoTo 12
     End If
   End If
   '
   '
   ' CARGAR MOVIMIENTOS DE CAJA Y BANCOS
   EXTEA$ = EXTE$
   If EXTEA$ = "DAT" Then
     Close #NX5%: NX5% = 0
     Close #NX6%: NX6% = 0
     Close #NX7%: NX7% = 0
     EXTE$ = "OLD"
   End If
   '
22 PMC$ = PUNMOCAJA$(CUE%)
   FIN& = Len(PMC$)
   For KK% = 1 To Len(PMC$) Step 2
      UU& = KK% + 1
      Call TRACE(20, UU&, FIN&)
      REGDIAL& = CVI(Mid$(PMC$, KK%, 2))
      Call LEEPASCAJ(REGDIAL&, 1)
      FEMOVI% = CVI(D15$)
      IMPO# = CVD(D16$)
      '
      If TRIM$(DETAPAS$) = "" Then
         Call LEEPASCAJ(REGDIAL&, 2)
         DETAPAS$ = REFASIEN$
      End If
      MARCADIA$ = ""
      If InStr(DETAPAS$, "- OP") > 0 Then GoTo 29
      If Left$(DETAPAS$, 8) = "DEPOSITO" Then
         FEREGI% = FEMOVI%
         FEMORE% = FECHACRE%(FEMOVI%, Val(Mid$(DETAPAS$, 10, 3)))
         If FEREGI% < DES% Then
            If FEREGI% >= CMJJ% Then GoSub 90
          ElseIf FEREGI% <= HAS% Then
            GoSub 95
         End If
         GoTo 29
      End If
      If FEMOVI% < DES% Then
        GoSub 90
        GoTo 29
      End If
      If FEMOVI% >= DES% Then
        If FEMOVI% <= HAS% Then
          FEREGI% = FEMOVI%
          FEMORE% = FEMOVI%
          GoSub 95
        End If
      End If
      '
29 Next KK%
   '
   If EXTE$ = "OLD" Then
     If EXTEA$ = "DAT" Then
       Close #NX5%: NX5% = 0
       Close #NX6%: NX6% = 0
       Close #NX7%: NX7% = 0
       EXTE$ = "DAT"
       EXTEA$ = ""
       GoTo 22
     End If
   End If
   '
   For U& = 1 To ULTREG&("CHEQUEMI")
     X$ = REGLEIDO$("CHEQUEMI", U&)
      PROV% = CVI(Left$(X$, 2))
      NUMERO$ = Mid$(X$, 3, 12)
      banco1$ = Mid$(X$, 15, 16)
      EMIS% = CVI(Mid$(X$, 47, 2)) ' toma fecha de O/Pago
      IMPO# = (-1) * CVD(Mid$(X$, 33, 8))
      NOPAGO& = CVS(Mid$(X$, 41, 4))
      VENCI% = CVI(Mid$(X$, 49, 2))
      QUE% = CVI(Mid$(X$, 45, 2))
            For UX% = 1 To CAMP%
               If MEPA%(UX%) = QUE% Then GoTo 38
            Next UX%
            Call MENSERR(24, "Medio de Pago" + Str$(QUE%) + " no Definido")
            Call FINAL("")
            '
38    CUEH% = CUEMP%(UX%)
      If CUEH% = CUE% Then
        FEREGI% = EMIS%
        If FEREGI% < DES% Then
           If FEREGI% >= CMJJ% Then GoSub 90
        End If
        FEMORE% = VENCI%
        If FEREGI% >= DES% Then
          If FEREGI% <= HAS% Then
            FEX = EMIS
            DETAPAS$ = "Ch." + TRIM$(NUMERO$)
            DETAPAS$ = DETAPAS$ + " - OP." + TRIM$(Str$(NOPAGO&))
            DETAPAS$ = DETAPAS$ + " - " + TRIM$(RASOCLI$((-1) * PROV%))
            GoSub 95
          End If
        End If
      End If
   Next U&
   '
   FEX = DES%
   FECHAMO%(0) = DIANTE(FEX)
   FEREGIS%(0) = FECHAMO%(0)
   FEX = FECHAMO%(0)
   REFE$(0) = "Saldo al " + FECHATEX$(FEX)
   IPODIF#(0) = SALINI#
   '
   For KKL& = 0 To JJ&
     For KKL1& = KKL& + 1 To JJ&
       If FEREGIS%(KKL1&) < FEREGIS%(KKL&) Then
         FX1% = FECHAMO%(KKL&)
         FX4% = FEREGIS%(KKL&)
         FX2$ = REFE$(KKL&)
         FX3# = IPODIF#(KKL&)
         FECHAMO%(KKL&) = FECHAMO%(KKL1&)
         FEREGIS%(KKL&) = FEREGIS%(KKL1&)
         REFE$(KKL&) = REFE$(KKL1&)
         IPODIF#(KKL&) = IPODIF#(KKL1&)
         FECHAMO%(KKL1&) = FX1%
         FEREGIS%(KKL1&) = FX4%
         REFE$(KKL1&) = FX2$
         IPODIF#(KKL1&) = FX3#
        End If
     Next KKL1&
     GoSub 96
   Next KKL&
   '
   Call SETULTREG("CONCIBAN", KLM&)
   Call CIERRARCH("CONCIBAN")
   Call CIERRARCH("*.*")
   '
   Screen.MousePointer = 1
   If Option5.Value = True Then
        Call CONECRUN("*CONCIBA/" + AJUSTI$(DCBB$, 30), "Conciliacion Bancaria (Pantalla)")
      Else
        Call HEADERS("CONCIBAN", "")
        Call HEADERS("CONCIBAN", "Banco: " + DCBB$)
        Call CONECRUN("*LCONBAN", "Conciliacion Bancaria (Listado)")
   End If
   GoTo 10
   '
   Exit Sub
   
90 SALINI# = SALINI# + IMPO#
   Return
   '
95 If Abs(IMPO#) >= 0.005 Then
      JJ& = JJ& + 1
      REFE$(JJ&) = DETAPAS$
      FECHAMO%(JJ&) = FEMORE%
      FEREGIS%(JJ&) = FEREGI%
      IPODIF#(JJ&) = IMPO#
   End If
   Return
   '
96 ZZZ$ = REGBLAN$("CONCIBAN")
   If DCBB$ <> DCBBA$ Then
     SARRAS# = 0
     DCBBA$ = DCBB$
   End If
   Call REPLA(ZZZ$, DCBB$, 1, 30)
   Call REPLA(ZZZ$, MKI$(FEREGIS%(KKL&)), 31, 2)
   Call REPLA(ZZZ$, MKI$(FECHAMO%(KKL&)), 33, 2)
   Call REPLA(ZZZ$, REFE$(KKL&), 35, 30)
   SARRAS# = SARRAS# + IPODIF#(KKL&)
   If IPODIF#(KKL&) > 0 Then
      Call REPLA(ZZZ$, MKD$(IPODIF#(KKL&)), 65, 8)
     Else
      Call REPLA(ZZZ$, MKD$(Abs(IPODIF#(KKL&))), 73, 8)
   End If
   Call REPLA(ZZZ$, MKD$(SARRAS#), 81, 8)
   KLM& = KLM& + 1
   Call GRAREG("CONCIBAN", ZZZ$, KLM&)
   Return
   '
End Sub

Sub GEPODIF(CUEKK%)
'
   Static CAMP%, CUEMP%(2048), MEPA%(2048)
   '
   Dim REFE$(2048)
   Dim FECHAMO%(2048)
   Dim IPODIF#(2048)
   '
   If CAMP% <= 0 Then
      For I& = 1 To ULTREG&("MEPAGO")
         F0$ = REGLEIDO$("MEPAGO", I&)
         CAMP% = I&
         MEPA%(CAMP%) = CVI(Left$(F0$, 2))
         CUEX$ = Mid$(F0$, 45, 12)
         CUEMP%(CAMP%) = CUEINT%(CUEX$)
         If TRIM$(CODCUE$(CUEMP%(CAMP%))) <> TRIM$(CUEX$) Then
            Call MENSERR(24, "Cuenta Contable Incorrecta\para Medio de Pago '" + TRIM$(Mid$(F0$, 3, 12)) + "'\-\Resultados Imprevisibles.")
            Call FINAL("")
         End If
     Next I&
   End If
   '
   HOII% = HOY(HOS$)
   DES% = COMEJ%
   HAS% = FINEJ%
   '
10 Screen.MousePointer = 11
   KLM& = 0
   '
   For cuebal& = 1 To ULTREG&("CUEBANC")
     CBB$ = REGLEIDO$("CUEBANC", cuebal&)
     DCBB$ = Left$(CBB$, 32)
     CUEX$ = Left$(CBB$, 12)
     CUE% = CUEINT%(CUEX$)
     If CUE% = CUEKK% Then GoTo 11
   Next cuebal&
   Call MENSERR(24, "Cuenta Inexistente")
   Call FINAL("")
   '
11 CUEBIN$ = MKI$(CUE%)
   CODENOCUE$ = TRIM$(CODCUE$(CUE%)) + " - " + TRIM$(DENOCUE$(CUE%))
   JJ& = 0
   SALINI# = 0
   '
   If DES% >= FECHANUM("01/10/01") Then
      CCCC$ = TRIM$(CODCUE$(CUE%))
      If CCCC$ = "111401" Then
         SALINI# = 10886.85   ' 11409.33
         SALINI# = SALINI# + 114.95 ' para tener en cuenta cheque duplicado flama
      End If
      If CCCC$ = "111502" Then SALINI# = 11791.1
      If CCCC$ = "111403" Then SALINI# = 13039.69   ' 2414.3
      If CCCC$ = "111406" Then SALINI# = 30026.22   ' 29769.7
      If CCCC$ = "111506" Then SALINI# = 381.86
   End If
   '
   ' CARGAR MOVIMIENTOS DE DIARIO
   EXTEA$ = EXTE$
   If EXTEA$ = "DAT" Then
     Close #NX9%: NX9% = 0
     Close #NX10%: NX10% = 0
     Close #NX11%: NX11% = 0
     EXTE$ = "OLD"
   End If
   '
12 PMC$ = PUNMODIAR$(CUE%)
   FIN& = Len(PMC$)
   For KK% = 1 To Len(PMC$) Step 2
      UU& = KK% + 1
      Call TRACE(20, UU&, FIN&)
      REGDIA% = CVI(Mid$(PMC$, KK%, 2))
      Call LEEDIARIO(REGDIA%)
      If CVS(E12$) < 1.05 Then GoTo 19 ' SALTEA ASIENTO APERTURA
      If Asc(E18$) <> 1 Then  ' SALTEA ASIENTOS RESUMEN DE CAJA
         '
         Call LEEPASDIA(REGDIA%, 2)
         DETAPAS$ = TRIM$(DETAPAS$)
         If DETAPAS$ <> "" Then
            DETAPAS$ = TRIM$(DETAPAS$) + " - "
         End If
         DETAPAS$ = DETAPAS$ + REFASIEN$
         '
         ACIERR% = 0
         If InStr(UCase$(DETAPAS$), "(CIERR") > 0 Or InStr(UCase$(DETAPAS$), "(*") > 0 Then
           ACIERR% = 1
         End If
         MARCADIA$ = "*"
         FEMOREA% = CVI(E15$)
         IMPO# = CVD(E16$)
         If FEMOREA% < DES% Then
             If ACIERR% = 1 Then IMPO# = 0
             If FEMOREA% > FECHANUM("31/09/01") Then GoSub 90 'TOTALIZAR INICIAL
          ElseIf FEMOREA% <= HAS% Then
            FEMORE% = FEMOREA%
            FEREGI% = FEMOREA%
            GoSub 95
         End If
      End If
19 Next KK%
   '
   If EXTE$ = "OLD" Then
     If EXTEA$ = "DAT" Then
       Close #NX9%: NX9% = 0
       Close #NX10%: NX10% = 0
       Close #NX11%: NX11% = 0
       EXTE$ = "DAT"
       EXTEA$ = ""
       GoTo 12
     End If
   End If
   '
   '
20 ' CARGAR MOVIMIENTOS DE CAJA Y BANCOS
   EXTEA$ = EXTE$
   If EXTEA$ = "DAT" Then
     Close #NX5%: NX5% = 0
     Close #NX6%: NX6% = 0
     Close #NX7%: NX7% = 0
     EXTE$ = "OLD"
   End If
   '
22 PMC$ = PUNMOCAJA$(CUE%)
   FIN& = Len(PMC$)
   For KK% = 1 To Len(PMC$) Step 2
      UU& = KK% + 1
      Call TRACE(20, UU&, FIN&)
      REGDIAL& = CVI(Mid$(PMC$, KK%, 2))
      Call LEEPASCAJ(REGDIAL&, 1)
      FEMOVI% = CVI(D15$)
      IMPO# = CVD(D16$)
      '
      If TRIM$(DETAPAS$) = "" Then
         Call LEEPASCAJ(REGDIAL&, 2)
         DETAPAS$ = REFASIEN$
      End If
      MARCADIA$ = ""
      If InStr(DETAPAS$, "- OP") > 0 Then GoTo 29
      If Left$(DETAPAS$, 8) = "DEPOSITO" Then
         FEMORE% = FECHACRE%(FEMOVI%, Val(Mid$(DETAPAS$, 10, 3)))
         If FEMORE% < DES% Then
            If FEMORE% > FECHANUM("31/09/01") Then GoSub 90
          ElseIf FEMORE% <= HAS% Then
            GoSub 95
         End If
         GoTo 29
      End If
      If FEMOVI% < DES% Then
        If FEMOVI% > FECHANUM("31/09/01") Then GoSub 90
        GoTo 29
      End If
      If FEMOVI% >= DES% Then
        If FEMOVI% <= HAS% Then
          FEMORE% = FEMOVI%
          GoSub 95
        End If
      End If
      '
29 Next KK%
   '
   If EXTE$ = "OLD" Then
     If EXTEA$ = "DAT" Then
       Close #NX5%: NX5% = 0
       Close #NX6%: NX6% = 0
       Close #NX7%: NX7% = 0
       EXTE$ = "DAT"
       EXTEA$ = ""
       GoTo 22
     End If
   End If
   '
   For U& = 1 To ULTREG&("CHEQUEMI")
     X$ = REGLEIDO$("CHEQUEMI", U&)
      PROV% = CVI(Left$(X$, 2))
      NUMERO$ = Mid$(X$, 3, 12)
      banco1$ = Mid$(X$, 15, 16)
      EMIS% = CVI(Mid$(X$, 47, 2)) ' toma fecha de O/Pago
      IMPO# = (-1) * CVD(Mid$(X$, 33, 8))
      NOPAGO& = CVS(Mid$(X$, 41, 4))
      VENCI% = CVI(Mid$(X$, 49, 2))
      '
      If VENCI% < EMIS% Then  ' modificacion del 28/01/02. Si
         VENCI% = EMIS%       ' la fecha del valor es anterior
      End If                  ' a la OP, toma la de la OP
      '
      QUE% = CVI(Mid$(X$, 45, 2))
            For UX% = 1 To CAMP%
               If MEPA%(UX%) = QUE% Then GoTo 38
            Next UX%
            Call MENSERR(24, "Medio de Pago" + Str$(QUE%) + " no Definido")
            Call FINAL("")
            '
38    CUEH% = CUEMP%(UX%)
      If CUEH% = CUE% Then
        FEMORE% = VENCI%
If EMIS% <= FECHANUM("31/09/01") And VENCI% > FECHANUM("31/09/01") Then
SALINICA# = SALINICA# + IMPO#
AAA = BBB
End If
        If FEMORE% < DES% Then
          If FEMORE% > FECHANUM("31/09/01") Then GoSub 90
        End If
        If FEMORE% >= DES% Then
          If FEMORE% <= HAS% Then
            FEX = EMIS%
            DETAPAS$ = "Ch." + TRIM$(NUMERO$)
            DETAPAS$ = DETAPAS$ + " - OP." + TRIM$(Str$(NOPAGO&))
            DETAPAS$ = DETAPAS$ + " - " + TRIM$(RASOCLI$((-1) * PROV%))
            GoSub 95
          End If
        End If
      End If
   Next U&
   '
   FEX = DES%
   FECHAMO%(0) = DIANTE(FEX)
   FEX = FECHAMO(0)
   REFE$(0) = "Saldo al " + FECHATEX$(FEX)
   IPODIF#(0) = SALINI#
   '
   For KKL& = 0 To JJ&
     For KKL1& = KKL& + 1 To JJ&
       If FECHAMO%(KKL1&) < FECHAMO%(KKL&) Then
         FX1% = FECHAMO%(KKL&)
         FX2$ = REFE$(KKL&)
         FX3# = IPODIF#(KKL&)
         FECHAMO%(KKL&) = FECHAMO%(KKL1&)
         REFE$(KKL&) = REFE$(KKL1&)
         IPODIF#(KKL&) = IPODIF#(KKL1&)
         FECHAMO%(KKL1&) = FX1%
         REFE$(KKL1&) = FX2$
         IPODIF#(KKL1&) = FX3#
        End If
     Next KKL1&
     GoSub 96
   Next KKL&
80 'Next cuebal&
   '
   Call SETULTREG("PODIFER", KLM&)
   Call CIERRARCH("PODIFER")
   Call CIERRARCH("*.*")
   '
   Screen.MousePointer = 1
   '
   Exit Sub
   
90 SALINI# = SALINI# + IMPO#
   Return
   '
95 If Abs(IMPO#) >= 0.005 Then
      JJ& = JJ& + 1
      REFE$(JJ&) = DETAPAS$
      FECHAMO%(JJ&) = FEMORE%
      IPODIF#(JJ&) = IMPO#
   End If
   Return
   '
96 ZZZ$ = REGBLAN$("PODIFER")
   If DCBB$ <> DCBBA$ Then
     SARRAS# = 0
     DCBBA$ = DCBB$
   End If
   Call REPLA(ZZZ$, DCBB$, 1, 32)
   Call REPLA(ZZZ$, MKI$(FECHAMO%(KKL&)), 33, 2)
   Call REPLA(ZZZ$, REFE$(KKL&), 35, 30)
   ACIERR% = 0
   If InStr(UCase$(REFE$(KKL&)), "(CIERR") > 0 Or InStr(REFE$(KKL&), "(*") > 0 Then
      ACIERR% = 1
   End If
   SARRAS# = SARRAS# + IPODIF#(KKL&)
   If ACIERR% > 0 Then SARRAS# = SARRAS# - IPODIF#(KKL&)
   If IPODIF#(KKL&) > 0 Then
      Call REPLA(ZZZ$, MKD$(IPODIF#(KKL&)), 65, 8)
      If ACIERR% > 0 Then Call REPLA(ZZZ$, MKD$(Abs(IPODIF#(KKL&))), 73, 8)
     Else
      Call REPLA(ZZZ$, MKD$(Abs(IPODIF#(KKL&))), 73, 8)
      If ACIERR% > 0 Then Call REPLA(ZZZ$, MKD$(Abs(IPODIF#(KKL&))), 65, 8)
   End If
   Call REPLA(ZZZ$, MKD$(SARRAS#), 81, 8)
   KLM& = KLM& + 1
   Call GRAREG("PODIFER", ZZZ$, KLM&)
   Return
   '

End Sub

Sub SUDIARCAJ()
   '
5  HAS% = HOY(HOS$)
   DES% = COMESACT%
   VDEF$ = MKI$(DES%) + MKI$(HAS%)
   OBXX$ = OBJPLA$("DESHASFECHA", VDEF$)
   If OBXX$ = "" Then Exit Sub
   '
   DES% = CVI(Left$(OBXX$, 2))
   HAS% = CVI(Mid$(OBXX$, 3, 2))
   If HAS% < DES% Then GoTo 5
   If DES% < 1 Or HAS% > HOY(HOS$) Then GoTo 5
   FEX = DES%: MMXX1$ = FECHATEX$(FEX)
   FEX = HAS%: MMXX1$ = MMXX1$ + " - " + FECHATEX$(FEX)
   '
   CORRESP$ = MMXX1$
   Screen.MousePointer = 11
   '
   ' CARGAR MOVIMIENTOS DE CAJA Y BANCOS
   FIN& = ULTREG&("CAJABANC")
   Z$ = REGBLAN$("SUCAJBAN")
   FEX = DES%
   DIA0% = DIANTE(FEX)
   Call REPLA(Z$, MKI$(DIA0%), 1, 2)
   Call REPLA(Z$, MKS$(0), 3, 4)
   Call REPLA(Z$, "            ", 7, 12)
   Call REPLA(Z$, "Transporte", 19, 48)
   Call REPLA(Z$, MKI$(DIA0%), 67, 2)
   Call REPLA(Z$, Space$(36), 69, 36)
   Call REPLA(Z$, MKD$(0), 105, 8)
   Call REPLA(Z$, MKD$(0), 113, 8)
   Call REPLA(Z$, MKD$(0), 121, 8)
   Call GRAREG("SUCAJBAN", Z$, 1&)
   '
   Z$ = REGBLAN$("SUCAJBAN")
   Call REPLA(Z$, MKI$(DIA0%), 1, 2)
   Call GRAREG("SUCAJBAN", Z$, 2)
   '
   JJ& = 2: ASIANT = 0: NUCOANT$ = ""
   SUDEBE# = 0: SUHABE# = 0
   '
   For UU& = 1 To FIN&
      '
      Call TRACE(20, UU&, FIN&)
      X$ = REGLEIDO$("CAJABANC", UU&)
      STATU% = Asc(Mid$(X$, 121, 1))
      '
      If STATU% < 9 Then      ' no está anulado
        FEAS% = CVI(Mid$(X$, 61, 2))
        IDEBE# = CVD(Mid$(X$, 153, 8))
        IHABE# = CVD(Mid$(X$, 161, 8))
        '
        If FEAS% < DES% Then
             '
             SUDEBE# = SUDEBE# + IDEBE#
             SUHABE# = SUHABE# + IHABE#
             '
          ElseIf FEAS% <= HAS% Then
             '
             NUCO$ = Left$(X$, 12)
             NUAS = CVS(Mid$(X$, 125, 4))
             CUE% = CVI(Mid$(X$, 65, 2))
             REFASIEN$ = Mid$(X$, 13, 48)
             '
             If NUCO$ <> NUCOANT$ Then
               Z$ = REGBLAN$("SUCAJBAN")
               Call REPLA(Z$, MKI$(FEAS%), 1, 2)
               Call REPLA(Z$, MKS$(NUAS), 3, 4)
               Call REPLA(Z$, NUCO$, 7, 12)
               Call REPLA(Z$, REFASIEN$, 19, 48)
               JJ& = JJ& + 1
               Call GRAREG("SUCAJBAN", Z$, JJ&)
               NUCOANT$ = NUCO$: ASIANT = NUAS
             End If
             '
             FACRE% = CVI(Mid$(X$, 63, 2))
             '
             REFEL$ = TRIM$(Mid$(X$, 71, 40))
20           PPXX% = InStr(REFEL$, "  ")
             If PPXX% > 0 Then
               REFEL$ = TRIM$(Left$(REFEL$, PPXX%) + Mid$(REFEL$, PPXX% + 2))
               GoTo 20
             End If
             REFEL$ = "- " + REFEL$
             '
             REFE$ = REPLACAR$(REFEL$, "@", "A")
             If TRIM$(REFE$) = "-" Then
               REFE$ = "- " + TRIM$(DENOCUE$(CUE%))
             End If
             DENOMI$ = TRIM$(CODCUE$(CUE%)) + " - " + TRIM$(DENOCUE$(CUE%))
             '
             Z$ = REGBLAN$("SUCAJBAN")
             Call REPLA(Z$, MKI$(FEAS%), 1, 2)
             Call REPLA(Z$, MKS$(NUAS), 3, 4)
             Call REPLA(Z$, NUCO$, 7, 12)
             Call REPLA(Z$, REFE$, 19, 48)
             Call REPLA(Z$, MKI$(FACRE%), 67, 2)
             Call REPLA(Z$, DENOMI$, 69, 36)
             '
             Call REPLA(Z$, MKD$(IDEBE#), 105, 8)
             Call REPLA(Z$, MKD$(IHABE#), 113, 8)
             JJ& = JJ& + 1
             Call GRAREG("SUCAJBAN", Z$, JJ&)
             '
        End If
      End If
      '
90 Next UU&
   '
   Z$ = REGLEIDO$("SUCAJBAN", 1)
   Call REPLA(Z$, MKD$(SUDEBE#), 105, 8)
   Call REPLA(Z$, MKD$(SUHABE#), 113, 8)
   Call GRAREG("SUCAJBAN", Z$, 1)
   '
   Call SETULTREG("SUCAJBAN", JJ&)
   Call CIERRARCH("SUCAJBAN")
   '
   Call HEADERS("SUCAJBAN", "")
   Call HEADERS("SUCAJBAN", "Corresponde a: " + CORRESP$)
   '
   Screen.MousePointer = 1
   Call CONECRUN("*SOSUCAJ", "Libro Subdiario de Caja y Bancos")
   '
End Sub
'
Sub RESCAJAPAN()
   '
5  HAS% = HOY(HOS$)
   DES% = COMESACT%
   VDEF$ = MKI$(DES%) + MKI$(HAS%)
   OBXX$ = OBJPLA$("DESHASFECHA", VDEF$)
   If OBXX$ = "" Then Exit Sub
   '
   DES% = CVI(Left$(OBXX$, 2))
   HAS% = CVI(Mid$(OBXX$, 3, 2))
   If HAS% < DES% Then GoTo 5
   If DES% < 1 Or HAS% > HOY(HOS$) Then GoTo 5
   FEX = DES%: MMXX1$ = FECHATEX$(FEX)
   FEX = HAS%: MMXX1$ = MMXX1$ + " - " + FECHATEX$(FEX)
   '
   CORRESP$ = MMXX1$
   '
10 Call SELECHO("ECUECON")
   CUEX$ = VALACT1$("ECUECON")
   If CUEX$ = "" Then GoTo 5
   '
   CUE% = CUEINT%(CUEX$)
   If CUE% < 1 Then
      Call MENSERR(24, "Cuenta Inexistente")
      GoTo 10
   End If
   '
   CODECUE$ = CUEX$
   DENECUE$ = DENOCUE$(CUE%)
   JJ& = 0: ASIANT = 0: NUCOANT$ = ""
   '
   XX$ = REGBLAN$("RESUCON")
   Call REPLA(XX$, CODECUE$, 1, 12)
   Call REPLA(XX$, DENECUE$, 13, 40)
   Call REPLA(XX$, "Transporte", 75, 30)
   JJ& = JJ& + 1
   Call GRAREG("RESUTES", XX$, JJ&)
   '
   XX$ = REGBLAN$("RESUCON")
   Call REPLA(XX$, CODECUE$, 1, 12)
   Call REPLA(XX$, DENECUE$, 13, 40)
   Call REPLA(XX$, "", 75, 30)
   JJ& = JJ& + 1
   Call GRAREG("RESUTES", XX$, JJ&)
   '
   SUDEBE# = 0: SUHABE# = 0
   RFF$ = RECFILT$("CAJABANC", 5, MKI$(CUE%))
   FIN& = Len(RFF$)
   For U& = 1 To FIN& Step 4
      Call TRACE(20, U&, FIN&)
      UU& = CVS(Mid$(RFF$, U&, 4))
      X$ = REGLEIDO$("CAJABANC", UU&)
      STATU% = Asc(Mid$(X$, 121, 1))
      '
      If STATU% < 9 Then      ' no está anulado
        FEAS% = CVI(Mid$(X$, 61, 2))
        IDEBE# = CVD(Mid$(X$, 153, 8))
        IHABE# = CVD(Mid$(X$, 161, 8))
        '
        If FEAS% < DES% Then
             '
             SUDEBE# = SUDEBE# + IDEBE#
             SUHABE# = SUHABE# + IHABE#
             '
          ElseIf FEAS% <= HAS% Then
             '
             NUCO$ = Left$(X$, 12)
             NUAS = CVS(Mid$(X$, 125, 4))
             CUE% = CVI(Mid$(X$, 65, 2))
             REFASIEN$ = Mid$(X$, 13, 48)
             '
             FACRE% = CVI(Mid$(X$, 63, 2))
             '
             REFEL$ = TRIM$(Mid$(X$, 177, 32))
             If REFEL$ = "" Then
               REFEL$ = TRIM$(Mid$(X$, 71, 40))
             End If
             '
20           PPXX% = InStr(REFEL$, "  ")
             If PPXX% > 0 Then
               REFEL$ = TRIM$(Left$(REFEL$, PPXX%) + Mid$(REFEL$, PPXX% + 2))
               GoTo 20
             End If
             '
             REFE$ = REPLACAR$(REFEL$, "@", "A")
             If TRIM$(REFE$) = "-" Then
               REFE$ = "- " + TRIM$(DENOCUE$(CUE%))
             End If
             DENOMI$ = TRIM$(CODCUE$(CUE%)) + " - " + TRIM$(DENOCUE$(CUE%))
             '
             XX$ = REGBLAN$("RESUCON")
             Call REPLA(XX$, CODECUE$, 1, 12)
             Call REPLA(XX$, DENECUE$, 13, 40)
             Call REPLA(XX$, NUCO$, 53, 12)
             Call REPLA(XX$, MKI$(FEAS%), 65, 2)
             Call REPLA(XX$, MKS$(NUAS), 67, 4)
             Call REPLA(XX$, MKI$(NROASDEF%), 71, 2)
             Call REPLA(XX$, MARCADIA$, 73, 2)
             Call REPLA(XX$, REFEL$, 75, 30)
             Call REPLA(XX$, MKD$(IDEBE#), 105, 8)
             Call REPLA(XX$, MKD$(IHABE#), 113, 8)
             JJ& = JJ& + 1
             Call GRAREG("RESUTES", XX$, JJ&)
        End If
      End If
   Next U
   '
   XX$ = REGLEIDO$("RESUTES", 1)
   Call REPLA(XX$, MKD$(SUDEBE#), 105, 8)
   Call REPLA(XX$, MKD$(SUHABE#), 113, 8)
   Call GRAREG("RESUTES", XX$, 1)
   '
   Call SETULTREG("RESUTES", JJ&)
   Call CIERRARCH("RESUTES")
   If JJ& < 1 Then
        Call MENSERR(24, "Cuenta " + CODECUE$ + "\sin Movimientos en el Ejercicio.")
      Else
        Call FILESORT("RESUTES", "", 4, 4, "ASC")
        Call CONECRUN("*RESUTES/" + CODECUE$, "Resumen de Cuenta de Caja y Bancos")
   End If
   GoTo 10
   '
Return
'
End Sub

Sub SALTESOR()
5 JJ& = 0
  HOII% = HOY(HOS$)
  Screen.MousePointer = 11
  For UK& = 1 To ULTREG&("SALTESOR")
    XX$ = REGLEIDO$("SALTESOR", UK&)
    CUEX$ = TRIM$(Left$(XX$, 12))
    If CUEX$ <> "" Then
      CUE% = CUEINT%(CUEX$)
      If CUE% > 0 Then
        SALAC# = 0: SALDIF# = 0
        RFF$ = RECFILT$("CAJABANC", 5, MKI$(CUE%))
        FIN& = Len(RFF$)
        For U& = 1 To FIN& Step 4
          Call TRACE(20, U&, FIN&)
          UU& = CVS(Mid$(RFF$, U&, 4))
          X$ = REGLEIDO$("CAJABANC", UU&)
          STATU% = Asc(Mid$(X$, 121, 1))
          '
          If STATU% < 9 Then      ' no está anulado
            FEAS% = CVI(Mid$(X$, 61, 2))
            FACRE% = CVI(Mid$(X$, 63, 2))
            IDEBE# = CVD(Mid$(X$, 153, 8))
            IHABE# = CVD(Mid$(X$, 161, 8))
            '
            If FACRE% <= HOII% Then
                SALAC# = SALAC# + IDEBE# - IHABE#
              Else
                SALDIF# = SALDIF# + IDEBE# - IHABE#
            End If
          End If
        Next U&
        JJ& = JJ& + 1
        Call REPLA(XX$, MKD$(SALAC#), 13, 8)
        Call REPLA(XX$, MKD$(SALDIF#), 21, 8)
        Call REPLA(XX$, MKD$(SALAC# + SALDIF#), 29, 8)
        Call GRAREG("SALTESOR", XX$, JJ&)
      End If
    End If
  Next UK&
  Call SETULTREG("SALTESOR", JJ&)
  '
  FIANT# = FIDATIM#("SALTESOR")
  Screen.MousePointer = 1
  Call FINAL("*SALTESOR")
  If FIDATIM#("SALTESOR") <> FIANT# Then GoTo 5
  '
End Sub


Private Sub Option9_Click()
   Command8.Enabled = False
   If Option9.Value = True Then
     Command8.Enabled = True
   End If
End Sub

Sub CONASIDIA(HAS%, CORRESP$)
    '
    Screen.MousePointer = 11
    J& = 0
    Call BLANARCH("!ASIDIAFI")
    TTXXY$ = Chr$(1) + Chr$(3) + MKI$(HAS%)
    For IL& = 1 To ULTREG&("ASIDIAR")
        X$ = REGLEIDO$("ASIDIAR", IL&)
        If Left$(X$, 4) = TTXXY$ Then
            CUENU& = CVI(Mid$(X$, 5, 2))
            IDEB# = CVD(Mid$(X$, 39, 8))
            If Abs(IDEB#) >= 0.005 Then
                IHAB# = 0
                If IDEB# < 0 Then
                    IHAB# = Abs(IDEB#)
                    IDEB# = 0
                End If
                CUES$ = Space$(12)
                If CUENU& > 0 Then
                    If CUENU& <= ULTREG&("CUENTAS") Then
                        CUES$ = Left$(REGLEIDO$("CUENTAS", CUENU&), 12)
                    End If
                End If
                DENOCUES$ = ECOARCH$("ECUECON", CUES$)
                '
                Z$ = REGBLAN$("ASIDIAFI")
                Call REPLA(Z$, CUES$, 1, 12)
                Call REPLA(Z$, DENOCUES$, 13, 30)
                Call REPLA(Z$, MKD$(IDEB#), 43, 8)
                Call REPLA(Z$, MKD$(IHAB#), 51, 8)
                '
                J& = J& + 1
                Call GRAREG("!ASIDIAFI", Z$, J&)
            End If
        End If
    Next IL&
    Call SETULTREG("!ASIDIAFI", J&)
    '
    For K1& = 1 To J&
        For K2& = K1& + 1 To J&
            X1$ = REGLEIDO$("!ASIDIAFI", K1&)
            X2$ = REGLEIDO$("!ASIDIAFI", K2&)
            If Left$(X1$, 12) > Left$(X2$, 12) Then
                X0$ = X1$
                X1$ = X2$
                X2$ = X0$
                Call GRAREG("!ASIDIAFI", X1$, K1&)
                Call GRAREG("!ASIDIAFI", X2$, K2&)
            End If
        Next K2&
    Next K1&
    Call CIERRARCH("!ASIDIAFI")
    '
    Screen.MousePointer = 1
    VTB% = VENTABU%("ASIDIAFI/TITUPAN=Resumen Mensual Caja y Bancos - " + CORRESP$)
    '
End Sub
'
Sub ASIDECHE()
   '
   HOII% = HOY(HOS$)
   CUECHECA$ = CONTROL$("", "CUECHECA")
   CUEICHEQ% = CUEINT%(CUECHECA$)
   If TRIM$(CUECHECA$) = "" Or CUEICHEQ% < 1 Then
      Call MENSERR(24, "Falta Definir Cuenta Contable\Asociada a Cartera de Valores")
      OPCHEQ04.Show 1
      Exit Sub
   End If
   '
   If ULTREG&("CUEBANC") < 1 Then
      Call MENSERR(24, "Falta Definir Cuentas Bancarias\para Deposito de Cheques.")
      Call CARCUEBANC
      Exit Sub
   End If
   '
5  VDEF$ = Space$(80)
   Call REPLA(VDEF$, MKI$(HOII%), 13, 2)
   Call REPLA(VDEF$, "DEPOSITO CHEQUES ", 15, 40)
   Call REPLA(VDEF$, CUECHECA$, 55, 12)
   Call REPLA(VDEF$, "48 Hs", 77, 10)
10 OBX$ = OBJPLA$("DEPOCHECAR", VDEF$)
   If OBX$ = "" Then Exit Sub
   '
   VDEF$ = OBX$
   CUEBA$ = Left$(OBX$, 12)
   If CUEINT%(CUEBA$) < 1 Then GoTo 10
   DESTICHE$ = ECOARCH$("CUEBANC", CUEBA$)
   If DESTICHE$ = "" Then
     Call MENSERR(24, "Cuenta " + CUEBA$ + " no Definida\como Cuenta Bancaria")
     GoTo 10
   End If
   '
   CUECHE$ = Mid$(OBX$, 55, 12)
   FEDEPO% = CVI(Mid$(OBX$, 13, 2))
   If FEDEPO% < COMEJ% Or FEDEPO% > FINEJ% Then
     Call COMUNI("Fecha de Depósito Fuera del Ejercicio Contable")
     GoTo 10
   End If
   PLADEPO$ = TRIM$(Mid$(OBX$, 77, 10))
   HORACRE% = Val(Left$(PLADEPO$, 3))
   If HORACRE% < 24 Then
     If Left$(PLADEPO$, 2) <> "0 " Then
       Call MENSERR(24, "Plazo de Acreditación no Válido.")
       GoTo 10
     End If
   End If
   If HORACRE% Mod 24 <> 0 Then
     Call MENSERR(24, "Plazo de Acreditación no Válido.")
     GoTo 10
   End If
   '
   Screen.MousePointer = 11
   FEX = FEDEPO%
   REFE$ = TRIM$(Mid$(OBX$, 15, 38)) + " " + FECHATEX$(FEX)
   Call GRACONTROL("", "CUECHE", CUECHE$)
   '
   SELCHE04.List1.Clear
   SELCHE04.LICHESEL.Clear
   '
   Call APERBASDAT("CABAN")
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "where TipoMovim = 'CHRECI' "
   SQLX$ = SQLX$ + "AND Status = 1 "
   '
   Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   JJ& = 0: TRANSPOR# = 0
   FORMALTER.LINOSORT.Clear
   With CABATEMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do While Not .EOF
         If TRIM(!OpciMovim) <> "CH" Then GoTo 19 ' no es un cheque
         If IsNull(!FECHENTRE) = False Then
           If CVINT(!FECHENTRE) > 0 Then
             GoTo 19    ' salio de cartera
           End If
         End If
         '
         FEACRED% = CVINT(!FechAcred)
         BANUFE$ = !Descrimov
         ICHEQ# = !ValAbsComp
         '
         If TRIM$(Left$(BANUFE$, 12)) = "" Then
           Call REPLA(BANUFE$, "Cheque Nro:", 1, 11)
         End If
         RECHE& = !NUMECOMP
         '
         Z$ = "A" + Space$(80)
         Call REPLA(Z$, CSTRING$(MKI$(FEACRED%), 1, 5, 0, 2), 2, 5)
         Call REPLA(Z$, TRIM$(Left$(BANUFE$, 40)), 9, 40)
         Call REPLA(Z$, CSTRING$(MKD$(ICHEQ#), 4, 12, 2, 2), 49, 12)
         Call REPLA(Z$, CSTRING$(MKS$(RECHE&), 3, 8, 0, 2), 73, 8)
         SELCHE04.List1.AddItem Z$
         '
19     .MoveNext
       Loop
     End If
     On Error GoTo 0
   End With
   '
   SUCHETO# = 0
   For I& = 1 To SELCHE04.List1.ListCount
      Z$ = Mid$(SELCHE04.List1.List(I& - 1), 9)
      SELCHE04.LICHESEL.AddItem Z$
      SUCHETO# = SUCHETO# + Val(Mid$(Z$, 41, 12))
   Next I&
   '
   SELCHE04.Label4.Caption = CSTRING$(MKD$(SUCHETO#), 4, 11, 2, 2)
   SELCHE04.Label4.Refresh
   '
   Screen.MousePointer = 1
   SELCHE04.Show 1
   DoEvents
   '
   IPASE# = Val(SELCHE04.Label2.Caption)
   FEACR% = FECHACRE%(FEDEPO%, HORACRE%)
   If IPASE# > 0.005 Then
      Screen.MousePointer = 11
      NOCLOSE% = 1
      FEX = FEDEPO%
      FETX$ = FECHATEX$(FEX)
      NRECIB$ = "DB-" + Mid$(FETX$, 7, 2) + Mid$(FETX$, 4, 2) + Mid$(FETX$, 1, 2) + "-"
      NRECIBX$ = NRECIB$ + "??"
      ORMA% = (-1)
      '
      Call APERBASDAT("CABAN")
      SQLX$ = "SELECT * FROM CAJABANC "
      SQLX$ = SQLX$ + "where CodiCom_Cor LIKE '" & NRECIBX$ & "' "
      Set CajaBanc1 = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
      With CajaBanc1
        On Error Resume Next
        .FindFirst "CodiCom_Cor LIKE '" & NRECIBX$ & "'"
30      If .NoMatch = False Then
          If Val(Mid$(!CodiCom_Cor, 11)) > ORMA% Then
            ORMA% = Val(Mid$(!CodiCom_Cor, 11))
          End If
          .FindNext "CodiCom_Cor LIKE '" & NRECIBX$ & "'"
          GoTo 30
        End If
      End With
      ORMX$ = TRIM$(Str$(ORMA% + 1))
      While Len(ORMX$) < 2: ORMX$ = "0" + ORMX$: Wend
      NRECIB$ = NRECIB$ + Left$(ORMX$, 2)
      '
      CoComCor$ = NRECIB$
      CoComLar$ = NRECIB$
      CodiEmp% = 0
      CodiSuc% = SUC%
      NumeClie% = 0
      NumeProv% = 0
      FechEmi = CVDAT(FEDEPO%)
      '
      Referen$ = REFE$
      FeContab = CVDAT(FEDEPO%)
      FeAcredi = CVDAT(FEACR%)
      NumeTal = Val(Mid$(NRECIB$, 11, 2))
      NumeCom = Val(Mid$(NRECIB$, 4, 6))
      Observa$ = "Depósito Cheques " + PLADEPO$
      '
      CuentaCo% = CUEINT%(CUEBA$)
      TiMovi = "DEPOBAN"            ' DEPOSITO BANCARIO
      OpMovi = PLADEPO$             ' PLAZO DE ACREDITACION
      VaMovi = ""                   '
      Descrip = "Depósito Cheques " + PLADEPO$
      ImpoDeb# = IPASE#
      StatInit = 0
      Call GRACABAN
      '
      SUMACHESEL# = 0
      FEX = FEDEPO%: HOS$ = FECHATEX$(FEX)
      Call APERBASDAT("CABAN")
      SQLX$ = "SELECT * FROM CAJABANC "
      SQLX$ = SQLX$ + "where TipoMovim = 'CHRECI' "
      SQLX$ = SQLX$ + "AND Status = 1 "
      Set CajaBanc2 = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
      '
      For I& = 1 To SELCHE04.LICHESEL.ListCount
        If SELCHE04.LICHESEL.Selected(I& - 1) = True Then
          X$ = SELCHE04.LICHESEL.List(I& - 1)
          BANUFE$ = Left$(X$, 40)
          BACH$ = TRIM$(Left$(X$, 16))
          NCHE$ = Mid$(X$, 19, 12)
          IPASE# = Val(Mid$(X$, 41, 12))
          SUMACHESEL# = SUMACHESEL# + IPASE#
          REGBA& = Val(Mid$(X$, 65, 8))
          '
          SQLX$ = "DescriMov = '" & BANUFE$ & "' "
          SQLX$ = SQLX$ + "AND ValAbsComp = " & Str$(IPASE#) & " "
          SQLX$ = SQLX$ + "AND NumeComp = " & REGBA& & " "
          SQLX$ = SQLX$ + "AND Status = 1"
          '
          With CajaBanc2
            .FindFirst SQLX$
            If .NoMatch = False Then
              On Error GoTo 0
              .Edit
              !FEDEPOSENT = CVDAT(FEDEPO%)
              !Document = NRECIB$
              !DESTICHEQ = DESTICHE$
              CUECONTABINT% = !CueContaI
              !REFESALCHE = Left$(REFE$, 48)
              !IDENSALCHE = AJUSTI$(HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 48)
              !UsuaSalChe = AJUSTI$("Usuario: " + TRIM$(USERNAME$), 48)
              !Status = 2                     ' Marca Salida de Cartera
              !IdenBaja = IDACTU$             ' Marca Modificacion de registro - FECHA - HORA - USUARIO - TERMINAL
              .Update
            End If
          End With
          '
          CoComCor$ = NRECIB$
          CoComLar$ = NRECIB$
          CodiEmp% = 0
          CodiSuc% = SUC%
          NumeClie% = 0
          NumeProv% = 0
          FechEmi = CVDAT(FEDEPO%)
          '
          Referen$ = REFE$
          FeContab = CVDAT(FEDEPO%)
          FeAcredi = CVDAT(FEDEPO%)
          NumeTal = Val(Mid$(NRECIB$, 11, 2))
          NumeCom = Val(Mid$(NRECIB$, 4, 6))
          '
          CuentaCo% = CUECONTABINT%
          TiMovi$ = "CHENTR"
          OpMovi$ = "DP"
          VaMovi$ = PLADEPO$
          Descrip = BANUFE$
          ImpoDeb# = (-1) * IPASE#
          Call GRACABAN
          '
        End If
      Next I&
      NOCLOSE% = 0
      Set CajaBanc2 = Nothing
      '
      '\\\OT-05-0475-WAR-07/07/05///
      'IMPRESION DE DEPOSITO BANCARIO
      FORIPRE$ = TRIM$(CONTROL$("", "FORDEPBA"))
237   If TRIM$(FORIPRE$) <> "" Then
         LU$ = LUDAT$
         CODFOR$ = TRIM$(UCase$(FORIPRE$))
         '
         YAC = Val(ATRIFORM$(CODFOR$, "DETDESMM"))
         MARIZW = Val(ATRIFORM$(CODFOR$, "MARIZQMM"))
         MARSUW = Val(ATRIFORM$(CODFOR$, "MARSUPMM"))
         '
         INTERLI = Val(ATRIFORM$(CODFOR$, "INTERLIN"))
         If INTERLI < 1 Then
            Call COMUNI("Error de Configuración de Interlinea\en Formulario '" + CODFOR$ + ".FRM'.")
            INTERLI = 1
         End If
         '
         CACOPI% = Val(ATRIFORM$(CODFOR$, "CAN-COP"))
         If CACOPI% < 1 Or CACOPI% > 6 Then
            Call COMUNI("Error de Configuración de Cantidad de Copias\en Formulario '" + CODFOR$ + ".FRM'.")
            CACOPI% = 1
         End If
         '
         Screen.MousePointer = 11
         FEX = FEDEPO%
         FECHDOC$ = FECHATEX(FEX)
         NUMEDOC$ = NRECIB$
         TIPODOC$ = "Deposito Bancario"
         FEHORE$ = Format(Now, " dd/mm/yy  hh:mm")
         '
         CODPARAM$(1) = "NUME-COM" 'NUMERO DE DEPOSITO BANCARIO
         CODPARAM$(2) = "FECH-EMI" 'FECHA DE EMISION
         CODPARAM$(3) = "COD-CUEN" 'CUENTA CONTABLE-BANCO
         CODPARAM$(4) = "DEN-CUEN" 'DESCRIPCION CUENTA
         CODPARAM$(5) = "REF-ASIE" 'REFERENCIA
         CODPARAM$(6) = "TIE-TOTA" 'PLAZO DE DEPOSITO /CLEARING
         CODPARAM(7) = "F-PLANO" 'FECHA Y HORA REAL
         CODPARAM(8) = "IMP-TOTA" 'SUMA TOTAL DE IMPORTES CHEQUES
    
         CAPARDOC% = 8
    
         DOCPARAM(1) = NUME1
         DOCPARAM(2) = FECHDOC
         DOCPARAM(3) = CUEBA$
         DOCPARAM(4) = ECOARCH("ECUECON", CUEBA$) 'DESTICHE$
         DOCPARAM(5) = REFE$
         DOCPARAM(6) = PLADEPO$
         DOCPARAM(7) = FEHORE$
         '
         CAITAB1% = 0
         CAITCRU% = 0
         '
         'AQUI PONER LOS CODIGOS DE COLUMNAS A LISTAR
         CODTABU1$(1) = "REF-MAT1" 'BANCO
         CODTABU1$(2) = "REF-MAT2" 'NUMERO
         CODTABU1$(3) = "DES-MAT" 'OBSERVACIONES
         CODTABU1$(4) = "SUB-TOTA" 'IMPORTE
         CODTABU1$(5) = "FECH-REC" 'FECHA DE ACREDITACION
         CACOLTAB1% = 5
         '
         'ACA PONGO EL DETALLE
         SUMACHESEL# = 0
         FEX = FEDEPO%: HOS$ = FECHATEX$(FEX)
         Call APERBASDAT("CABAN")
         For I& = 1 To SELCHE04.LICHESEL.ListCount
           If SELCHE04.LICHESEL.Selected(I& - 1) = True Then
             X$ = SELCHE04.LICHESEL.List(I& - 1)
             BANUFE$ = Left$(X$, 40)
             BAN1$ = TRIM(Mid(BANUFE, 1, 19))
             NUMCHE1$ = TRIM(Mid(BANUFE, 19, 14))
             FECHE1$ = TRIM(Mid(BANUFE, 31, 12))
             BACH$ = TRIM$(Left$(X$, 16))
             NCHE$ = Mid$(X$, 19, 12)
             IPASE# = Val(Mid$(X$, 41, 12))
             SUMACHESEL# = SUMACHESEL# + IPASE#
             '
             CAITAB1% = CAITAB1% + 1
             TETABU1$(1, CAITAB1%) = BAN1$
             TETABU1$(2, CAITAB1%) = NUMCHE1$
             TETABU1$(3, CAITAB1%) = OBSERVA1$ 'FALTA
             ' TETABU1$(4, CAITAB1%) = IPASE#   ' Walter, aqui hay un error - ver conmigo
             TETABU1$(4, CAITAB1%) = FORMATNUM$(IPASE#, "F12.2")
             TETABU1$(5, CAITAB1%) = FECHE1$
           End If
         Next I&
         'DOCPARAM(8) = SUMACHESEL#  ' WALTER, AQUI HAY UN ERROR
         DOCPARAM(8) = FORMATNUM$(SUMACHESEL#, "F14.2")
         '
         Call PRINTDOC("FORDEPBA") ' DEPOSITO BANCARIO
         '
         'LIMPIA LA LISTA
         For KK1% = 1 To CACOLTAB1%
           For KK2% = 1 To CAITAB1%
             TETABU1$(KK1%, KK2%) = ""
           Next KK2%
         Next KK1%
         '
199      Call CIERRARCH("*.*")
         Call BAJALDISCO
      End If
      '\\\OT-05-0475-WAR-FIN///
      Screen.MousePointer = 1
      Call COMUNI("Depósito Procesado")
      '
   End If
   '
   Screen.MousePointer = 1
   GoTo 10
   '
End Sub

Sub ANUDECHE() 'ANULACION
  '
  Call COMUNI("Recuerde que no Podrán Anularse\los Depósitos ya Conciliados.")
  '
  '1.- SELECCION MOVI CAJABANC SEGUN BANCO Y SIN REPETIR Y CN
  'MARCONSI IGUAL NADA
3 Screen.MousePointer = 11
  Call APERBASDAT("CABAN")
  Screen.MousePointer = 11
  SQLX$ = " SELECT *  FRoM CAJABANC"
  SQLX$ = SQLX$ + " WHERE TIPOMOVIM = 'DEPOBAN'"
  SQLX$ = SQLX$ + " AND(MarConsi = '' or ISNULL(MarConsi) = True or IsEmpty(MarConsi) = True)"
  SQLX$ = SQLX$ + " ORDER BY CODICOM_COR ASC"
  Set DEPTMP = CueCorri.OpenRecordset(SQLX$, 4)
  With DEPTMP
    On Error Resume Next
    .MoveFirst
    If Err Then
      Call COMUNI("No Existen Depositos\que Puedan Anularse")
      On Error GoTo 0
      GoTo 99
    End If
    On Error GoTo 0
    '
    '2.- GRABO ARCHIVO SELDEPO
    J& = 0
    ZZ$ = REGBLAN$("SELDEPO")
    NUDEPAN$ = "": TXBUS$ = ""
    Do While .EOF = False
      If NUDEPAN$ <> !CodiCom_Cor Then
        Call REPLA(ZZ$, !CodiCom_Cor, 1, 12)
        Call REPLA(ZZ$, MKI(CVINT(!FechConta)), 13, 2)
        'Call REPLA(ZZ$, MKD(!ValAbsComp), 15, 8)
        J& = J& + 1
        Call GRAREG("SELDEPO", ZZ$, J&)
        NUDEPAN$ = !CodiCom_Cor
        TXBUS$ = TXBUS$ + AJUSTI$(NUDEPAN$, 16)
      End If
      .MoveNext
    Loop
  End With
  '
  ' ESTO ES PARA AGREGAR LOS QUE SALIERON PERO SIN ENTRADA
  SQLX$ = " SELECT *  FRoM CAJABANC"
  SQLX$ = SQLX$ + " WHERE TIPOMOVIM = 'CHRECI' AND DocumEnt LIKE 'DB-*' "
  SQLX$ = SQLX$ + " ORDER BY DocumEnt ASC"
  Set DEPTMP = CueCorri.OpenRecordset(SQLX$, 4)
  With DEPTMP
    On Error Resume Next
    .MoveFirst
    If Err Then GoTo 5
    On Error GoTo 0
    Do While .EOF = False
      NUDEPAN$ = !Document
      If InStr(TXBUS$, NUDEPAN$) < 1 Then
        Call REPLA(ZZ$, !Document, 1, 12)
        Call REPLA(ZZ$, MKI(CVINT(!FEDEPOSENT)), 13, 2)
        'Call REPLA(ZZ$, MKD(!ValAbsComp), 15, 8)
        J& = J& + 1
        Call GRAREG("SELDEPO", ZZ$, J&)
        TXBUS$ = TXBUS$ + AJUSTI$(NUDEPAN$, 16)
      End If
      .MoveNext
    Loop
  End With
  '
5 On Error GoTo 0
  Call SETULTREG("SELDEPO", J&)
  Call FILESORT("SELDEPO", "", 1, 1, "ASC")
  Call FRESHECHO("SELDEPO")
  Screen.MousePointer = 1
  '
  '3.-REGSEL("SELDEPO")
6 Y1$ = REGSEL("SELDEPO")
  If Len(Y1$) <= 4 Then GoTo 99
  '
  Screen.MousePointer = 11
  NUDEP1$ = Mid$(Y1$, 1, 12)
  '4.- SELECCIONO Y BORRRO TODOS LOS ARHIVOS DEL CAJABANC
  If TRIM(NUDEP1$) <> "" Then
     Screen.MousePointer = 11
     SQLY$ = " SELECT *  FRoM CAJABANC"
     SQLY$ = SQLY$ + " WHERE CODICOM_LAR = '" & NUDEP1$ & "'"
     SQLY$ = SQLY$ + " OR (TipoMovim = 'CHRECI' AND DocumEnt = '" & NUDEP1$ & "') "
     Set CajaBanSel = CueCorri.OpenRecordset(SQLY$, dbOpenDynaset)
     '
     With CajaBanSel
       SQLX$ = "CODICOM_LAR = '" & NUDEP1$ & "' "
       .FindFirst SQLX$
7      If .NoMatch = False Then
         If IsNull(!MARCONSI) = False And IsEmpty(!MARCONSI) = False Then
           If TRIM$(!MARCONSI) <> "" Then
             Call MENSERR(24, "Imposible Anular Movimiento Conciliado.")
             GoTo 99
           End If
         End If
         .FindNext SQLX$
         GoTo 7
       End If
       '
10     .FindFirst SQLX$
       If .NoMatch = False Then
         .Delete
         GoTo 10
       End If
       '
       .FindFirst "TipoMovim = 'CHRECI' AND DocumEnt = '" & NUDEP1$ & "' "
22     If .NoMatch = False Then
         .Edit
         !Status = 1     'LO PONGO DEVUELTA COMO PENDIENTE EN CARTERA DE VALORES
         !Document = ""  'QUITO EL NUMERO DE DEPOSITO ASIGNADO
         !DESTICHEQ = ""
         !REFESALCHE = ""
         !IDENSALCHE = ""
         !FEDEPOSENT = CVDAT(0) 'fecha de deposito
         .Update
         .FindNext "TipoMovim = 'CHRECI' AND DocumEnt = '" & NUDEP1$ & "' "
         GoTo 22
       End If
     End With
     Set CajaBanSel = Nothing
     '
     Screen.MousePointer = 1
     Call COMUNI("Se ha Anulado el Déposito Bancario: '" & NUDEP1$ & "'.")
     GoTo 3
  End If
  '
99 Screen.MousePointer = 1
End Sub

Sub CORRIGEIMPU() 'REIMPUTACION
  '
     Screen.MousePointer = 11
     Call ABRECABAN
'     SQLY$ = " SELECT *  FROM CAJABANC"
'     SQLY$ = SQLY$ + " WHERE TipoMovim = 'CHRECI' AND (Nuclien=18 or Nuclien=1003)"
'     Set CajaBanSel = CueCorri.OpenRecordset(SQLY$, dbOpenDynaset)
'     '
'     With CajaBanSel
'         .FindFirst "Descrimov Like '*7425*' AND nuclien=18"
'         If .NoMatch = False Then
'            .Edit
'            !CueContaI = 52
'            !Status = 1
'            .Update
'         End If
'         .FindFirst "Descrimov Like '*128526*' AND Nuclien=1003"
'         If .NoMatch = False Then
'            .Edit
'            !CueContaI = 52
'            !Status = 1
'            .Update
'         End If
'         .FindFirst "Descrimov Like '*2432*' AND Nuclien=18"
'         If .NoMatch = False Then
'            .Edit
'            !CueContaI = 52
'            !Status = 1
'            .Update
'         End If
'         .FindFirst "Descrimov Like '*7229*' AND Nuclien=18"
'         If .NoMatch = False Then
'            .Edit
'            !CueContaI = 52
'            !Status = 1
'            .Update
'         End If
'         .FindFirst "Descrimov Like '*128339*' AND Nuclien=1003"
'         If .NoMatch = False Then
'            .Edit
'            !CueContaI = 52
'            !Status = 1
'            .Update
'         End If
'         .FindFirst "Descrimov Like '*9155*' AND Nuclien=18"
'         If .NoMatch = False Then
'            .Edit
'            !CueContaI = 52
'            !Status = 1
'            .Update
'         End If
'         .FindFirst "Descrimov Like '*9196*' AND Nuclien=18"
'         If .NoMatch = False Then
'            .Edit
'            !CueContaI = 52
'            !Status = 1
'            .Update
'         End If
'   End With
'   Set CajaBanSel = Nothing
'   '
'   Screen.MousePointer = 1
'   Call COMUNI("Reimputacion de Cheques Finalizada")
'  '
'99 Screen.MousePointer = 1
'Exit Sub
  
  
  '1.- SELECCION MOVI CAJABANC SEGUN BANCO Y SIN REPETIR Y CN
  'MARCONSI IGUAL NADA
3 Screen.MousePointer = 11
  Call APERBASDAT("CABAN")
  Screen.MousePointer = 11
  SQLX$ = " SELECT *  FRoM CAJABANC"
  SQLX$ = SQLX$ + " WHERE TIPOMOVIM = 'DEPOBAN'"
  SQLX$ = SQLX$ + " AND (MarConsi = '' or ISNULL(MarConsi) = True or IsEmpty(MarConsi) = True)"
  SQLX$ = SQLX$ + " ORDER BY CODICOM_COR ASC"
  Set DEPTMP = CueCorri.OpenRecordset(SQLX$, 4)
  With DEPTMP
    On Error Resume Next
    .MoveFirst
    If Err Then
      Call COMUNI("No Existen Depositos\que Puedan Re-Imputarse")
      On Error GoTo 0
      GoTo 99
    End If
    On Error GoTo 0
    '
    '2.- GRABO ARCHIVO SELDEPO
    J& = 0
    ZZ$ = REGBLAN$("SELDEPO")
    NUDEPAN$ = "": TXBUS$ = ""
    Do While .EOF = False
      If NUDEPAN$ <> !CodiCom_Cor Then
        Call REPLA(ZZ$, !CodiCom_Cor, 1, 12)
        Call REPLA(ZZ$, MKI(CVINT(!FechConta)), 13, 2)
        'Call REPLA(ZZ$, MKD(!ValAbsComp), 15, 8)
        J& = J& + 1
        Call GRAREG("SELDEPO", ZZ$, J&)
        NUDEPAN$ = !CodiCom_Cor
        TXBUS$ = TXBUS$ + AJUSTI$(NUDEPAN$, 16)
      End If
      .MoveNext
    Loop
  End With
  '
  ' ESTO ES PARA AGREGAR LOS QUE SALIERON PERO SIN ENTRADA
  SQLX$ = " SELECT *  FRoM CAJABANC"
  SQLX$ = SQLX$ + " WHERE TIPOMOVIM = 'CHRECI' AND DocumEnt LIKE 'DB-*' "
  SQLX$ = SQLX$ + " ORDER BY DocumEnt ASC"
  Set DEPTMP = CueCorri.OpenRecordset(SQLX$, 4)
  With DEPTMP
    On Error Resume Next
    .MoveFirst
    If Err Then GoTo 5
    On Error GoTo 0
    Do While .EOF = False
      NUDEPAN$ = !Document
      If InStr(TXBUS$, NUDEPAN$) < 1 Then
        Call REPLA(ZZ$, !Document, 1, 12)
        Call REPLA(ZZ$, MKI(CVINT(!FEDEPOSENT)), 13, 2)
        'Call REPLA(ZZ$, MKD(!ValAbsComp), 15, 8)
        J& = J& + 1
        Call GRAREG("SELDEPO", ZZ$, J&)
        TXBUS$ = TXBUS$ + AJUSTI$(NUDEPAN$, 16)
      End If
      .MoveNext
    Loop
  End With
  '
5 On Error GoTo 0
  Call SETULTREG("SELDEPO", J&)
  Call FILESORT("SELDEPO", "", 1, 1, "ASC")
  Call FRESHECHO("SELDEPO")
  Screen.MousePointer = 1
  '
  '3.-REGSEL("SELDEPO")
6 Y1$ = REGSEL("SELDEPO")
  If Len(Y1$) <= 4 Then GoTo 99
  '
  Call SELECHO("ECUECON")
  VDEVU$ = VALACT1$("ECUECON")
  If VDEVU$ = "" Then GoTo 99
  CUCON% = CUEINT%(VDEVU$)

  Screen.MousePointer = 11
  NUDEP1$ = Mid$(Y1$, 1, 12)
  '4.- SELECCIONO Y BORRRO TODOS LOS ARHIVOS DEL CAJABANC
  If TRIM(NUDEP1$) <> "" Then
     Screen.MousePointer = 11
     SQLY$ = " SELECT *  FRoM CAJABANC"
     SQLY$ = SQLY$ + " WHERE CODICOM_LAR = '" & NUDEP1$ & "'"
     SQLY$ = SQLY$ + " OR (TipoMovim = 'CHRECI' AND DocumEnt = '" & NUDEP1$ & "') "
     Set CajaBanSel = CueCorri.OpenRecordset(SQLY$, dbOpenDynaset)
     '
     With CajaBanSel
       SQLX$ = "CODICOM_LAR = '" & NUDEP1$ & "' "
       .FindFirst SQLX$
7      If .NoMatch = False Then
         If IsNull(!MARCONSI) = False And IsEmpty(!MARCONSI) = False Then
           If TRIM$(!MARCONSI) <> "" Then
             Call MENSERR(24, "Imposible Anular Movimiento Conciliado.")
             GoTo 99
           End If
         End If
         .FindNext SQLX$
         GoTo 7
       End If
       '
       .FindFirst "TipoMovim = 'DEPOBAN' AND Codicom_Lar = '" & NUDEP1$ & "' "
22     If .NoMatch = False Then
         .Edit
         !CueContaI = CUCON%
         .Update
         .FindNext "TipoMovim = 'DEPOBAN' AND Codicom_Lar = '" & NUDEP1$ & "' "
         GoTo 22
       End If
     End With
     Set CajaBanSel = Nothing
     '
     Screen.MousePointer = 1
     Call COMUNI("Se ha Reimputador el Déposito Bancario: '" & NUDEP1$ & "'.")
     GoTo 3
  End If
  '
99 Screen.MousePointer = 1
End Sub

Sub TRANFONDOS()
   '
   HOII% = HOY(HOS$)
   CUECHECA$ = CONTROL$("", "CUECHECA")
   CUEICHEQ% = CUECOINT%(CUECHECA$)
   If TRIM$(CUECHECA$) = "" Or CUEICHEQ% < 1 Then
      Call MENSERR(24, "Falta Definir Cuenta Contable\Asociada a Cartera de Valores")
      OPCHEQ00.Show 1
      Exit Sub
   End If
   '
   If ULTREG&("CUEBANC") < 1 Then
      Call MENSERR(24, "Falta Definir Cuentas Bancarias\para Deposito de Cheques.")
      Call CARCUEBANC
      Exit Sub
   End If
   '
5  VDEF$ = Space$(96)
   Call REPLA(VDEF$, MKI$(HOII%), 1, 2)
   Call REPLA(VDEF$, "TRANSFERENCIA DE FONDOS ", 35, 40)
   '
10 OBX$ = OBJPLA$("TRANFONDOS", VDEF$)
   If OBX$ = "" Then Exit Sub
   '
   VDEF$ = OBX$
   CUEDEBE$ = TRIM$(Mid$(OBX$, 23, 12))
   CUDEB1% = CUECOINT%(CUEDEBE$)
   If CUEDEBE$ = "" Or CUDEB1% < 1 Then
     Call MENSERR(24, "Cuenta 'Debe' no Válida o Inexistente.")
     GoTo 10
   End If
   '
   CUEHABE$ = TRIM$(Mid$(OBX$, 11, 12))
   CUHAB1% = CUECOINT%(CUEHABE$)
   If CUEHABE$ = "" Or CUHAB1% < 1 Then
     Call MENSERR(24, "Cuenta 'Haber' no Válida o Inexistente.")
     GoTo 10
   End If
   '
   If CUDEB1% = CUEICHEQ% Or CUHAB1% = CUEICHEQ% Then
     Call MENSERR(24, "No Válidas Transferencias que Tocan\Cuenta Asociada a la Cartera de Valores.")
     GoTo 10
   End If
   
   FEDEP% = CVI(Mid$(OBX$, 1, 2))
   If FEDEP% < 1 Then
     Call MENSERR(24, "Fecha no Válida")
     GoTo 10
   End If
   '
   ITRAN# = CVD(Mid$(OBX$, 3, 8))
   If ITRAN# < 0.005 Then
     Call MENSERR(24, "Falta Importe")
     GoTo 10
   End If
   '
   If TRIM$(Mid$(OBX$, 35, 40)) = "" Then
     Call MENSERR(24, "Falta Referencia")
     GoTo 10
   End If
   REFE$ = TRIM$(Mid$(OBX$, 35, 40)) + " " + FECHATEX$(FEX)
   '
   Screen.MousePointer = 11
   FEX = FEDEP%
   FETX$ = FECHATEX$(FEX)
   NRECIB$ = "TF-" + Mid$(FETX$, 7, 2) + Mid$(FETX$, 4, 2) + Mid$(FETX$, 1, 2) + "-"
   NRECIBX$ = NRECIB$ + "??"
   ORMA% = (-1)
   Call APERBASDAT("CABAN")
   Screen.MousePointer = 11
   With CajaBanc
     On Error Resume Next
     .FindFirst "CodiCom_Cor LIKE '" & NRECIBX$ & "'"
30   If .NoMatch = False Then
       If Val(Mid$(!CodiCom_Cor, 11)) > ORMA% Then
         ORMA% = Val(Mid$(!CodiCom_Cor, 11))
       End If
       .FindNext "CodiCom_Cor LIKE '" & NRECIBX$ & "'"
       GoTo 30
     End If
   End With
   ORMX$ = TRIM$(Str$(ORMA% + 1))
   While Len(ORMX$) < 2: ORMX$ = "0" + ORMX$: Wend
   NRECIB$ = NRECIB$ + Left$(ORMX$, 2)
   '
   CoComCor$ = NRECIB$
   CoComLar$ = NRECIB$
   CodiEmp% = 0
   CodiSuc% = SUC%
   NumeClie% = 0
   NumeProv% = 0
   FechEmi = CVDAT(FEDEP%)
   '
   Referen$ = REFE$
   FeContab = CVDAT(FEDEP%)
   FeAcredi = CVDAT(FEDEP%)
   NumeTal = Val(Mid$(NRECIB$, 11, 2))
   NumeCom = Val(Mid$(NRECIB$, 4, 6))
   Observa$ = "Transferencia de Fondos"
   '
   CuentaCo% = CUDEB1%
   TiMovi = "TRAFON"            ' TRANSFERENCIA DE FONDOS
   OpMovi = ""                  '
   VaMovi = ""                  '
   Descrip = "De " + TRIM$(ECOARCH$("ECUECON", CUEHABE$))
   ImpoDeb# = ITRAN#
   StatInit = 0
   Call GRACABAN
   '
   CuentaCo% = CUHAB1%
   Descrip = "A " + TRIM$(ECOARCH$("ECUECON", CUEDEBE$))
   ImpoDeb# = (-1) * ITRAN#
   Call GRACABAN
   '\\\OT-05-0440-WAR-04/07/05///
   'IMPRIMIR
FORIPRE$ = TRIM$(CONTROL$("", "FORTRAFO"))
237 If TRIM$(FORIPRE$) <> "" Then
    LU$ = LUDAT$
    CODFOR$ = TRIM$(UCase$(FORIPRE$))
    '
    YAC = Val(ATRIFORM$(CODFOR$, "DETDESMM"))
    MARIZW = Val(ATRIFORM$(CODFOR$, "MARIZQMM"))
    MARSUW = Val(ATRIFORM$(CODFOR$, "MARSUPMM"))
    '
    INTERLI = Val(ATRIFORM$(CODFOR$, "INTERLIN"))
    If INTERLI < 1 Then
       Call COMUNI("Error de Configuración de Interlinea\en Formulario '" + CODFOR$ + ".FRM'.")
       INTERLI = 1
    End If
    '
    CACOPI% = Val(ATRIFORM$(CODFOR$, "CAN-COP"))
    If CACOPI% < 1 Or CACOPI% > 6 Then
       Call COMUNI("Error de Configuración de Cantidad de Copias\en Formulario '" + CODFOR$ + ".FRM'.")
       CACOPI% = 1
    End If
    '
    For COP% = 1 To CACOPI%
      '
      Call CARFORWIN
      '
      Call PRINTFORWIN(FECHATEX$(FEX), "FECH-EMI", 0, 0)
      Call PRINTFORWIN(NRECIB$, "NUME-COM", 0, 0)
      Call PRINTFORWIN(USERNAME$, "USUARIO", 0, 0)
      Call PRINTFORWIN(REFE$, "REF-ASIE", 0, 0)
      Call PRINTFORWIN(CODCUE$(CUHAB1%), "COD-CUEN", 0, YAC)
      Call PRINTFORWIN(CODCUE$(CUDEB1%), "COD-CUEN", 0, YAC + INTERLI)
      Call PRINTFORWIN(TRIM$(ECOARCH$("ECUECON", CUEHABE$)), "DEN-CUEN", 0, YAC)
      Call PRINTFORWIN(TRIM$(ECOARCH$("ECUECON", CUEDEBE$)), "DEN-CUEN", 0, YAC + INTERLI)
      CANS$ = CSTRING$(MKS$(ITRAN#), 3, 9, 2, 2)
      Call PRINTFORWIN(CANS$, "IMP-DEBE", 0, YAC)
      Call PRINTFORWIN(REFE0$, "REF-MAT1", 0, YAC)
      '
      Call SETSPOOL("ENDDOC", "")
      '
    Next COP%
    'VALIDACION DE MODO DE IMPRESION
    '*******************************
    CANCELPRINT% = 0
    MODOIMPRE04.Check7.Caption = "Imprime Transferencia"
    MODOIMPRE04.Label1.Caption = "" 'LIMPIA EL LABEL DONDE PONE OK, SI ACEPTA
    MODOIMPRE04.Check8.Value = 0
    Screen.MousePointer = 1
    MODOIMPRE04.Show 1
    If MODOIMPRE04.Label1.Caption = "OK" Then
       If MODOIMPRE04.Check8.Value > 0 Then CANCELPRINT% = 1
    End If
    Screen.MousePointer = 11
    Call SAVESPOOL("Transferencia de Fondo", OKX%)
    End If
    '
   '\\\OT-05-0440-WAR-FIN///
   On Error Resume Next
   CajaBanc.Close
   On Error GoTo 0
   '
   CANCELPRINT% = 0
   Screen.MousePointer = 1
   Call COMUNI("Transferencia Completa y Verificada")
   GoTo 5
   '
End Sub

Sub RECICHEQ()
   '
'\\\OT-5-602-FG-29/08/05///
   HOII% = HOY(HOS$)
   CUECHECA$ = CONTROL$("", "CUECHECA")
   CUEICHEQ% = CUEINT%(CUECHECA$)
   If TRIM$(CUECHECA$) = "" Or CUEICHEQ% < 1 Then
      Call MENSERR(24, "Falta Definir Cuenta Contable\Asociada a Cartera de Valores")
      OPCHEQ04.Show 1
      Exit Sub
   End If
   '
5  VDEF$ = Space$(80)
   Call REPLA(VDEF$, MKI$(HOII%), 1, 2)
   Call REPLA(VDEF$, MKI$(HOII%), 39, 2)
   Call REPLA(VDEF$, MKI$(HOII%), 41, 2)
   Call REPLA(VDEF$, "VALOR SIN RECIBO DE COBRANZA", 55, 40)
10 OBX$ = OBJPLA$("RECICHEQ", VDEF$)
   If OBX$ = "" Then Exit Sub
   '
   VDEF$ = OBX$
   FEDEPO% = CVI(Mid$(OBX$, 1, 2))
   FEMISI% = CVI(Mid$(OBX$, 39, 2))
   FEACRE% = CVI(Mid$(OBX$, 41, 2))
   banco1$ = TRIM$(Mid$(OBX$, 3, 16))
   NUCHE$ = TRIM$(Mid$(OBX$, 19, 12))
   IMCHE# = CVD(Mid$(OBX$, 31, 8))
   CONTRA$ = TRIM$(Mid$(OBX$, 43, 12))
   REFEX8$ = TRIM$(Mid$(OBX$, 55, 42))
   
   '
   '\\\ VALIDACIONES
   If FEDEPO% = 0 Then
     Call COMUNI("Fecha de Recepción no Válida")
     GoTo 10
   ElseIf FEMISI% = 0 Then
     Call COMUNI("Fecha de Emisión no Válida")
     GoTo 10
   ElseIf FEACRE% = 0 Then
     Call COMUNI("Fecha de Acreditación no Válida")
     GoTo 10
   ElseIf FEMISI% > FEACRE% Then
     Call COMUNI("La Fecha de Emisión no puede Superar a la de Acreditación")
     GoTo 10
   End If
   '
   If TRIM$(banco1$) = "" Then
     Call COMUNI("Falta Denominación del Banco")
     GoTo 10
   End If
   '
   If TRIM$(NUCHE$) = "" Then
     Call COMUNI("Falta Número de Cheque")
     GoTo 10
   End If
   '
   If IMCHE# < 0.01 Then
     Call COMUNI("Falta Importe del Cheque")
     GoTo 10
   End If
   '
   If CUEINT(CONTRA$) < 1 Then
     Call COMUNI("Cuenta Contable de Contrapartida no Válida")
     GoTo 10
   End If
   '
   If TRIM$(REFEX8$) = "" Then
     Call COMUNI("Falta Referencia del Movimiento")
     GoTo 10
   End If
   '
   VeriChek% = YAESTACHEQUE%(banco1$, NUCHE$)
   If VeriChek% > 0 Then
      Call MENSERR(24, "Cheque N.: " + NUCHE$ + "  Banco : " + banco1$ + " Ya Ingresado")
      GoTo 10
   End If

   '\\\FIN VALIDACIONES
   '
   Screen.MousePointer = 11
   FEX = FEDEPO%
   FETX$ = FECHATEX$(FEX)
   NRECIB$ = "CH-" + Mid$(FETX$, 7, 2) + Mid$(FETX$, 4, 2) + Mid$(FETX$, 1, 2) + "-"
   NRECIBX$ = NRECIB$ + "??"
   ORMA% = (-1)
   Call APERBASDAT("CABAN")
   With CajaBanc
     On Error Resume Next
     .FindFirst "CodiCom_Cor LIKE '" & NRECIBX$ & "'"
30   If .NoMatch = False Then
       If Val(Mid$(!CodiCom_Cor, 11)) > ORMA% Then
         ORMA% = Val(Mid$(!CodiCom_Cor, 11))
       End If
       .FindNext "CodiCom_Cor LIKE '" & NRECIBX$ & "'"
       GoTo 30
     End If
   End With
   ORMX$ = TRIM$(Str$(ORMA% + 1))
   While Len(ORMX$) < 2: ORMX$ = "0" + ORMX$: Wend
   NRECIB$ = NRECIB$ + Left$(ORMX$, 2)
   '
   CoComCor$ = NRECIB$
   CoComLar$ = NRECIB$
   Referen$ = REFEX8$
   FeContab = CVDAT(FEDEPO%)
   '
   FEX = FEACRE%: FEVECHE$ = FECHATEX$(FEX)
   CuentaCo% = CUEICHEQ%
   NumeClie% = 0
   NumeProv% = 0
   TiMovi$ = "CHRECI"
   OpMovi$ = "CH"
   VaMovi$ = ""
     IDEVA$ = AJUSTI$(banco1$, 18)
     IDEVA$ = IDEVA$ + AJUSTI$(NUCHE$, 14)
     IDEVA$ = IDEVA$ + FEVECHE$
   Descrip$ = IDEVA$
     If FECHE% < 1 Then FECHE% = FC%
   FechEmi = CVDAT(FEMISI%)
     If FEACRE% < FEMISI% Then FEACRE% = FEMISI%
     'If FEACRE% < FEDEPO% Then FEACRE% = FEDEPO%
   FeAcredi = CVDAT(FEACRE%)
   NumeTal = 0
   CodiEmp% = 0
   CodiSuc% = SUCUI%
   NumeCom = 0
   StatInit = 1
   ImpoDeb# = IMCHE#
   Observa$ = ""
   Call GRACABAN
   '
   CuentaCo% = CUEINT%(CONTRA$)
   TiMovi$ = "RECICH"
   OpMovi$ = ""
   VaMovi$ = ""
   StatInit = 0
   ImpoDeb# = (-1) * IMCHE#
   Call GRACABAN
   '
   FORIPRE$ = TRIM$(CONTROL$("", "MOVICHEQ"))
   If TRIM$(FORIPRE$) <> "" Then
    LU$ = LUDAT$
    CODFOR$ = TRIM$(UCase$(FORIPRE$))
    '
    CACOPI% = Val(ATRIFORM$(CODFOR$, "CAN-COP"))
    If CACOPI% < 1 Or CACOPI% > 6 Then
       Call COMUNI("Error de Configuración de Cantidad de Copias\en Formulario '" + CODFOR$ + ".FRM'.")
      CACOPI% = 1
    End If
    '
    For COP% = 1 To CACOPI%
      '
      Call CARFORWIN
      '
      Call PRINTFORWIN("Recepcion de Valores", "NUME-COM", 0, 0)
      Call PRINTFORWIN(banco1$, "BANC-CHE", 0, 0)
      Call PRINTFORWIN(NUCHE$, "NUME-CHE", 0, 0)
      IMPORTE$ = FORMATNUM$(IMCHE#, "F10.2")
      Call PRINTFORWIN(IMPORTE$, "IMPO-CHE", 0, 0)
      Call PRINTFORWIN(FECHATEX$(HOY(HO$)), "FECH-EMI", 0, 0)
      Call PRINTFORWIN(FECHATEX$(FEDEPO%), "FECH-REC", 0, 0)
      Call PRINTFORWIN(FECHATEX$(FEMISI%), "FECH-VEN", 0, 0)
      Call PRINTFORWIN(FECHATEX$(FEACRE%), "FECH-CHE", 0, 0)
      Call PRINTFORWIN(REFEX8$, "REF-MAT1", 0, 0)
      '
      'Printer.EndDoc
      Call SETSPOOL("ENDDOC", "")
      '
    Next COP%
    Call SAVESPOOL("Recepcion de Valores " & NRECIB$, OKX%)
   End If
   
   Screen.MousePointer = 1
   Call COMUNI("Se ha Procesado\la Recepción del Cheque / Valor")
   '
   GoTo 5
   '
'\\\OT-5-602-FG-FIN///
   '
End Sub

Sub ENTRECHEQ()
   '
   HOII% = HOY(HOS$)
   CUECHECA$ = CONTROL$("", "CUECHECA")
   CUEICHEQ% = CUEINT%(CUECHECA$)
   If TRIM$(CUECHECA$) = "" Or CUEICHEQ% < 1 Then
      Call MENSERR(24, "Falta Definir Cuenta Contable\Asociada a Cartera de Valores")
      OPCHEQ04.Show 1
      Exit Sub
   End If
   '
   If ULTREG&("CUEBANC") < 1 Then
      Call MENSERR(24, "Falta Definir Cuentas Bancarias\para Deposito de Cheques.")
      Call CARCUEBANC
      Exit Sub
   End If
   '
5  VDEF$ = Space$(80)
   Call REPLA(VDEF$, MKI$(HOII%), 13, 2)
   Call REPLA(VDEF$, "ENTREGA DE CHEQUES ", 15, 40)
   Call REPLA(VDEF$, CUECHECA$, 55, 12)
   Call REPLA(VDEF$, "", 77, 10)
10 OBX$ = OBJPLA$("ENTRECHEQ", VDEF$)
   If OBX$ = "" Then Exit Sub
   '
   VDEF$ = OBX$
   CUEBA$ = Left$(OBX$, 12)
   If CUEINT%(CUEBA$) < 1 Then GoTo 10
   DESTICHE$ = ECOARCH$("CUEBANC", CUEBA$)
   '
   CUECHE$ = Mid$(OBX$, 55, 12)
   FEDEPO% = CVI(Mid$(OBX$, 13, 2))
   If FEDEPO% < COMEJ% Or FEDEPO% > FINEJ% Then
     Call COMUNI("Fecha de Entrega Fuera del Ejercicio Contable")
     GoTo 10
   End If
   PLADEPO$ = ""
   '
   Screen.MousePointer = 11
   FEX = FEDEPO%
   REFE$ = TRIM$(Mid$(OBX$, 15, 38))
   '
   SELCHE04.List1.Clear
   SELCHE04.LICHESEL.Clear
   '
   Call APERBASDAT("CABAN")
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "where TipoMovim = 'CHRECI' "
   SQLX$ = SQLX$ + "AND Status = 1 "
   '
   Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   JJ& = 0: TRANSPOR# = 0
   FORMALTER.LINOSORT.Clear
   With CABATEMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do While Not .EOF
         If TRIM(!OpciMovim) <> "CH" Then GoTo 19 ' no es un cheque
         If IsNull(!FECHENTRE) = False Then
           If CVINT(!FECHENTRE) > 0 Then
             GoTo 19    ' salio de cartera
           End If
         End If
         '
         FEACRED% = CVINT(!FechAcred)
         BANUFE$ = !Descrimov
         ICHEQ# = !ValAbsComp
         '
         If TRIM$(Left$(BANUFE$, 12)) = "" Then
           Call REPLA(BANUFE$, "Cheque Nro:", 1, 11)
         End If
         RECHE& = !NUMECOMP
         '
         Z$ = "A" + Space$(80)
         Call REPLA(Z$, CSTRING$(MKI$(FEACRED%), 1, 5, 0, 2), 2, 5)
         Call REPLA(Z$, TRIM$(Left$(BANUFE$, 40)), 9, 40)
         Call REPLA(Z$, CSTRING$(MKD$(ICHEQ#), 4, 12, 2, 2), 49, 12)
         Call REPLA(Z$, CSTRING$(MKS$(RECHE&), 3, 8, 0, 2), 73, 8)
         SELCHE04.List1.AddItem Z$
         '
19     .MoveNext
       Loop
     End If
     On Error GoTo 0
   End With
   '
   SUCHETO# = 0
   For I& = 1 To SELCHE04.List1.ListCount
      Z$ = Mid$(SELCHE04.List1.List(I& - 1), 9)
      SELCHE04.LICHESEL.AddItem Z$
      SUCHETO# = SUCHETO# + Val(Mid$(Z$, 41, 12))
   Next I&
   '
   SELCHE04.Label4.Caption = CSTRING$(MKD$(SUCHETO#), 4, 11, 2, 2)
   SELCHE04.Label4.Refresh
   '
   Screen.MousePointer = 1
   SELCHE04.Show 1
   DoEvents
   '
   IPASE# = Val(SELCHE04.Label2.Caption)
   If IPASE# > 0.005 Then
      FEX = FEDEPO%
      FETX$ = FECHATEX$(FEX)
      NRECIB$ = "EV-" + Mid$(FETX$, 7, 2) + Mid$(FETX$, 4, 2) + Mid$(FETX$, 1, 2) + "-"
      NRECIBX$ = NRECIB$ + "??"
      ORMA% = (-1)
      Call APERBASDAT("CABAN")
      SQLX$ = "SELECT * FROM CAJABANC"
      SQLX$ = SQLX$ + " WHERE CodiCom_Cor LIKE '" & NRECIBX$ & "'"
      Set CAJABTEMP = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
      With CAJABTEMP
        .FindFirst "CodiCom_Cor LIKE '" & NRECIBX$ & "'"
30      If .NoMatch = False Then
          If Val(Mid$(!CodiCom_Cor, 11)) > ORMA% Then
            ORMA% = Val(Mid$(!CodiCom_Cor, 11))
          End If
          .FindNext "CodiCom_Cor LIKE '" & NRECIBX$ & "'"
          GoTo 30
        End If
      End With
      ORMX$ = TRIM$(Str$(ORMA% + 1))
      While Len(ORMX$) < 2: ORMX$ = "0" + ORMX$: Wend
      NRECIB$ = NRECIB$ + Left$(ORMX$, 2)
      '
      CoComCor$ = NRECIB$
      CoComLar$ = NRECIB$
      CodiEmp% = 0
      CodiSuc% = SUC%
      NumeClie% = 0
      NumeProv% = 0
      FechEmi = CVDAT(FEDEPO%)
      '
      Referen$ = REFE$
      FeContab = CVDAT(FEDEPO%)
      FeAcredi = CVDAT(FEDEPO%)
      NumeTal = Val(Mid$(NRECIB$, 11, 2))
      NumeCom = Val(Mid$(NRECIB$, 4, 6))
      Observa$ = "Entrega Cheques / Valores "
      '
      CuentaCo% = CUEINT%(CUEBA$)
      TiMovi = "ENTRECHE"           ' ENTREGA DE VALORES
      OpMovi = ""                   ' PLAZO DE ACREDITACION
      VaMovi = ""                   '
      Descrip = "Entrega Cheques / Valores "
      ImpoDeb# = IPASE#
      StatInit = 0
      Call GRACABAN
      '
      SUMACHESEL# = 0
      FEX = FEDEPO%: HOS$ = FECHATEX$(FEX)
      Call APERBASDAT("CABAN")
      For I& = 1 To SELCHE04.LICHESEL.ListCount
        If SELCHE04.LICHESEL.Selected(I& - 1) = True Then
          X$ = SELCHE04.LICHESEL.List(I& - 1)
          BANUFE$ = Left$(X$, 40)
          BACH$ = TRIM$(Left$(X$, 16))
          NCHE$ = Mid$(X$, 19, 12)
          IPASE# = Val(Mid$(X$, 41, 12))
          SUMACHESEL# = SUMACHESEL# + IPASE#
          REGBA& = Val(Mid$(X$, 65, 8))
          '
          SQLX$ = "DescriMov = '" & BANUFE$ & "' "
          SQLX$ = SQLX$ + "AND ValAbsComp = " & Str$(IPASE#) & " "
          SQLX$ = SQLX$ + "AND NumeComp = " & REGBA& & " "
          SQLX$ = SQLX$ + "AND Status = 1"
          '
          Call APERBASDAT("CABAN")
          Set CAJABTEMP = CueCorri.OpenRecordset("SELECT * FROM CAJABANC WHERE DescriMov = '" & BANUFE$ & "'", dbOpenDynaset)
          With CAJABTEMP
            .FindFirst SQLX$
            If .NoMatch = False Then
              On Error GoTo 0
              .Edit
              !FEDEPOSENT = CVDAT(FEDEPO%)
              !Document = NRECIB$
              !DESTICHEQ = DESTICHE$
              CUECONTABINT% = !CueContaI
              !REFESALCHE = Left$(REFE$, 48)
              !IDENSALCHE = AJUSTI$(HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 48)
              !UsuaSalChe = AJUSTI$("Usuario: " + TRIM$(USERNAME$), 48)
              !Status = 2                     ' Marca Salida de Cartera
              !IdenBaja = IDACTU$             ' Marca Modificacion de registro - FECHA - HORA - USUARIO - TERMINAL
              .Update
            End If
          End With
          '
          CoComCor$ = NRECIB$
          CoComLar$ = NRECIB$
          CodiEmp% = 0
          CodiSuc% = SUC%
          NumeClie% = 0
          NumeProv% = 0
          FechEmi = CVDAT(FEDEPO%)
          '
          Referen$ = REFE$
          FeContab = CVDAT(FEDEPO%)
          FeAcredi = CVDAT(FEDEPO%)
          NumeTal = Val(Mid$(NRECIB$, 11, 2))
          NumeCom = Val(Mid$(NRECIB$, 4, 6))
          '
          CuentaCo% = CUECONTABINT%
          TiMovi$ = "CHENTR"
          OpMovi$ = "SV"
          VaMovi$ = ""
          Descrip = BANUFE$
          ImpoDeb# = (-1) * IPASE#
          Call GRACABAN
          '
        End If
      Next I&
      '
199   Call CIERRARCH("*.*")
      Call BAJALDISCO
      Screen.MousePointer = 1
      Call COMUNI("Movimiento Procesado")
   End If
    '\\\OT-05-0475-WAR-FIN///
    
   Screen.MousePointer = 1
   GoTo 5
   '
End Sub

'\\\OT-5-661-FG-06/10/05///
Sub RESUTES()
   '
   Dim CUETE$(), SINI#(), SDEBE#(), SHABE#(), SFIN#()
   '
   If Option4.Value <> True Then
     Call COMUNI("Solo Disponible por Impresión")
     Option4.Value = True
   End If
   '
   If COMEJ% < 1 Then Call CARDATEJ
   If COMEJ% < 1 Then
      Call MENSERR(24, "Imposible Determinar Datos Ejercicio Activo.")
      Call FINAL("")
   End If
   '
   Call DESHASFECHA(DES%, HAS%, PERIO$)
   If PERIO$ = "" Then GoTo 99
   If DES% < COMEJ% Or HAS% > FINEJ% Then
      Call MENSERR(24, "Período Fuera de Límites\del Ejercicio Activo.")
      GoTo 99
   End If
   '
'   FEINI# = CDbl(CVDAT(DES%))
'   FEFIN# = CDbl(CVDAT(HAS%))
   CORRESP$ = FECHATEX$(DES%) + " - " + FECHATEX$(HAS%)
   '
   FIN& = ULTREG&("TACUECAJ")
   If FIN& < 1 Then
     Call MENSERR(24, "Falta Definir Cuentas de Tesorería")
     GoTo 99 ' NO EXISTEN CUENTAS DE TESORERIA
   End If
   ReDim CUETE$(FIN&), SINI#(FIN&), SDEBE#(FIN&), SHABE#(FIN&), SFIN#(FIN&)
   '
   Call BLANARCH("REDIATES")
   Call ABREPLANCUEN
   If NUCUEN% < 1 Then
     NUCUEN% = ULTREG&("CUENTAS")
   End If
   REBLA$ = REGBLAN$("REDIATES")
   '
   Screen.MousePointer = 11
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("TACUECAJ", U&)
     CUENTA$ = TRIM$(Mid$(X$, 1, 12))
     CUEII% = CUECOINT%(CUENTA$) ' CUENTA CONTABLE DE TESORERIA
     '
     ' BUSQUEDA DE MOVIMIENTOS EN TABLA DETASIEN
     ' *****************************************
     '
FEINI = Int(CDbl(CVDAT(DES%)))
FEINICONSUL = Int(CDbl(CVDAT(COMEJ%)))
FEFIN = Int(CDbl(CVDAT(HAS%)))
'
     SUCUI% = 0
     SQLX$ = "SELECT * FROM DetaAsien AND "
     SQLX$ = SQLX$ + "WHERE (CueCoInt=" & CUEII%
     SQLX$ = SQLX$ + " AND SucuAsien = " & SUCUI%
        SQLX$ = SQLX$ + " AND Fechasien >= " & FEINICONSUL
        SQLX$ = SQLX$ + " AND Fechasien <= " & FEFIN
     SQLX$ = SQLX$ + " AND statpase >0 ) "
     SQLX$ = SQLX$ + "ORDER BY FECHASIEN, ORFEASIEN ASC;"
     Set PASEASIEN = Contable.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
     '
     JJ& = 0
     TRANSPOR# = 0
10   With PASEASIEN
       On Error Resume Next
       .MoveFirst
       If Err Then GoTo 30
       On Error GoTo 0
       Do While Not .EOF
         IDEBE# = !IDEBEPASE
         IHABE# = !IHABEPASE
         If Int(CDbl(!FECHASIEN)) < FEINI Then
             TRANSPOR# = TRANSPOR# + IDEBE# - IHABE#
           ElseIf Int(CDbl(!FECHASIEN)) <= FEFIN Then
             Z$ = REBLA$
             Call REPLA(Z$, CODCUE$(!CUECOINT), 1, 12)
               FECH% = CVINT%(!FECHASIEN)
             Call REPLA(Z$, MKI$(FECH%), 13, 2)
             Call REPLA(Z$, !NUIDASIEN, 15, 12)
               REFEX8$ = !REFEPASE
             Call REPLA(Z$, REFEX8$, 27, 48)
             Call REPLA(Z$, REFEX8$, 75, 48)
             Call REPLA(Z$, MKD$(IDEBE#), 123, 8)
             Call REPLA(Z$, MKD$(IHABE#), 139, 8)
             JJ& = JJ& + 1
             Call GRAREG("!REDIATES", Z$, JJ&)
         End If
       .MoveNext
       Loop
     End With
     '
30   On Error GoTo 0
     
     ' BUSQUEDA DE MOVIMIENTOS EN TABLA CAJABANC
     ' *****************************************
     '
     EPAC$ = TRIM$(UCase$(EMPREAC$))
     '
     Call APERBASDAT("CABAN")
     SQLX$ = "SELECT * FROM CAJABANC "
     SQLX$ = SQLX$ + "where (CueContaI = " & CUEII% & " "
     SQLX$ = SQLX$ + " AND (NuAsidiar = '' or isnull(NuAsiDiar) = true OR IsEmpty(NuAsiDiar) = true))"
      SQLX$ = SQLX$ + " AND FechConta >= " & FEINICONSUL
      SQLX$ = SQLX$ + " AND FechConta <= " & FEFIN & " "
     SQLX$ = SQLX$ + " ORDER BY FECHCONTA ASC, FEHOREAL ASC;"
     '
     Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
     Set CABATEMP2 = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     With CABATEMP
       On Error Resume Next
       .MoveFirst
       If Err Then
          GoTo 40
       End If
       On Error GoTo 0
       Do While Not .EOF
         IDEBE# = !IDEBECOMP
         IHABE# = !IHABECOMP
If InStr(EPAC$, "TANIT") > 0 Then
  If CUEII% = CUEICARCH% Then
    If IDEBE# < 0.005 Then GoTo 17
  End If
End If
         If Int(CDbl(!FechConta)) < FEINI Then
             TRANSPOR# = TRANSPOR# + IDEBE# - IHABE#
           ElseIf Int(CDbl(!FechConta)) <= FEFIN Then
             Z$ = REBLA$
             Call REPLA(Z$, CODCUE(!CueContaI), 1, 12)
               FECH% = CVINT%(!FechConta)
             Call REPLA(Z$, MKI$(FECH%), 13, 2)
               COMPROB$ = !CodiCom_Cor
             Call REPLA(Z$, COMPROB$, 15, 12)
               REFEX8$ = !REFECOM
             Call REPLA(Z$, REFEX8$, 27, 48)
               DEPASE$ = !Descrimov
             Call REPLA(Z$, DEPASE$, 75, 48)
             Call REPLA(Z$, MKD$(IDEBE#), 123, 8)
             Call REPLA(Z$, MKD$(IHABE#), 139, 8)
             '
             JJ& = JJ& + 1
             Call GRAREG("!REDIATES", Z$, JJ&)
             '
If InStr(EPAC$, "TANIT") > 0 Then
  If CUEII% = CUEICARCH% Then
    IDEB1$ = TRIM$(FORMATNUM(IDEBE#, "F15.3"))
    SQLX$ = "DescriMov = '" & DEPASE$ & "' AND (ABS(IHabeComp - " & IDEB1$ & ")) < 0.005 "
    CABATEMP2.FindFirst SQLX$
    If CABATEMP2.NoMatch = False Then
        FEDEPOI = CDbl(!FEDEPOSENT)
        If FEDEPOI < FEINI Or FEDEPOI > FEFIN Then GoTo 17
      Z$ = REBLA$
         IDEBE# = CABATEMP2!IDEBECOMP
         IHABE# = CABATEMP2!IHABECOMP
      Call REPLA(Z$, CODCUE(CABATEMP2!CueContaI), 1, 12)
        FECH% = CVINT%(CABATEMP2!FechConta)
      Call REPLA(Z$, MKI$(FECH%), 13, 2)
        COMPROB$ = CABATEMP2!CodiCom_Cor
      Call REPLA(Z$, COMPROB$, 15, 12)
        REFEX8$ = CABATEMP2!REFECOM
      Call REPLA(Z$, REFEX8$, 27, 48)
        DEPASE$ = CABATEMP2!Descrimov
      Call REPLA(Z$, DEPASE$, 75, 48)
      Call REPLA(Z$, MKD$(IDEBE#), 123, 8)
      Call REPLA(Z$, MKD$(IHABE#), 139, 8)
        '
      JJ& = JJ& + 1
      Call GRAREG("!REDIATES", Z$, JJ&)
        '
    End If
  End If
End If
'
         End If
17     .MoveNext
       Loop
     End With
     '
     ' MUEVE LOS MOVIMIENTOS ENCONTRADOS AL ARCHIVO CENTRAL
     ' ****************************************************
     '
40   CUETE$(U&) = CODCUE$(CUEII%)
     SINI#(U&) = 0
If InStr(EPAC$, "TANIT") > 0 Then TRANSPOR# = 0
     If Abs(TRANSPOR#) > 0 Then
        X$ = REBLA$
        Call REPLA(X$, CODCUE$(CUEII%), 1, 12)
          FEX = DES%
          FEXI% = DIANTE(FEX)
        Call REPLA(X$, MKI$(FEXI%), 13, 2)
        Call REPLA(X$, "Saldo Anterior (Transporte) ..............................", 27, 48)
        If TRANSPOR# > 0 Then
            Call REPLA(X$, MKD$(TRANSPOR#), 123, 8)
          Else
            Call REPLA(X$, MKD$(TRANSPOR#), 139, 8)
        End If
        Call REGAPP("REDIATES", X$)
        SINI#(U&) = TRANSPOR#
        '
        'Pone un renglon en blanco
        X$ = REBLA$
        Call REPLA(X$, CODCUE$(CUEII%), 1, 12)
        Call REPLA(X$, MKI$(FEXI%), 13, 2)
        Call REGAPP("REDIATES", X$)
      End If
      '
      For UU& = 1 To JJ&
        X$ = REGLEIDO$("!REDIATES", UU&)
        SDEBE#(U&) = SDEBE#(U&) + CVD(Mid$(X$, 123, 8))
        SHABE#(U&) = SHABE#(U&) + CVD(Mid$(X$, 139, 8))
        Call REGAPP("REDIATES", X$)
      Next UU&
      '
      Z$ = REGBLAN$("RETESOR")
      Call REPLA(Z$, CUETE$(U&), 1, 12)
      Call REPLA(Z$, MKD$(SINI#(U&)), 13, 8)
      Call REPLA(Z$, MKD$(SDEBE#(U&)), 21, 8)
      Call REPLA(Z$, MKD$(SHABE#(U&)), 29, 8)
      Call REPLA(Z$, MKD$(SINI#(U&) + SDEBE#(U&) - SHABE#(U&)), 37, 8)
      Call GRAREG("RETESOR", Z$, U&)
      '
   Next U&
   Call SETULTREG("RETESOR", FIN&)
   Call CIERRARCH("RETESOR")
   Call FILESORT("RETESOR", "", 1, 1, "ASC")
   '
   Call FILESORT("REDIATES", "", 3, 3, "ASC")
   Call FILESORT("REDIATES", "", 1, 1, "ASC")
   '
   Call HEADERS("REDIATES", "")
   Call HEADERS("REDIATES", "Corresponde a: " + CORRESP$)
   '
   Call HEADERS("RETESOR", "")
   Call HEADERS("RETESOR", "Corresponde a: " + CORRESP$)
   '
   Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   '
90 OPX% = COMALTER%("Opciones de Listado:\\Detalle\Resumen\Cancelar")
   If OPX% = 1 Then
       Call FINAL("*LREDIATE")
       GoTo 90
     ElseIf OPX% = 2 Then
       Call FINAL("*RETESOR")
       GoTo 90
   End If
   '
99 End Sub



Sub BUSCHECUIT()
   '
30 CUITLIB$ = OBJPLA$("CACUITCH", CUITLIB$)
   If CUITLIB$ = "" Then
      Exit Sub
   End If
   '
   CUITLIB$ = TRIM$(CUITLIB$)
   Screen.MousePointer = 11
   Call FRESHECHO("CHESELEC")
   '
   Call APERBASDAT("CABAN")
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "where TipoMovim = 'CHRECI' "
   SQLX$ = SQLX$ + "and Status > 0 "
   SQLX$ = SQLX$ + "AND ObserGen LIKE  '*" & TRIM$(CUITLIB$) & "*'"
   '
   Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   JJ& = 0
   With CABATEMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do While Not .EOF
         If TRIM(!OpciMovim) <> "CH" Then GoTo 39 ' no es un cheque
         ICHX$ = MKD$(!ValAbsComp)
         X$ = AJUSTI$(!Descrimov, 40) + String$(8, 0) + ICHX$
         NCHX$ = Mid$(X$, 19, 12)
         '
         JJ& = JJ& + 1
         Call GRAREG("CHESELEC", X$, JJ&)
39     .MoveNext
       Loop
     End If
   End With
   '
   Call SETULTREG("CHESELEC", JJ&)
   Call CIERRARCH("CHESELEC")
   Screen.MousePointer = 1
   '
   CASEL& = ULTREG&("CHESELEC")
   If CASEL& < 1 Then
         Call COMUNI("Número no Encontrado")
         GoTo 30
       ElseIf CASEL& = 1 Then
         Y$ = REGLEIDO$("CHESELEC", 1&)
         GoTo 45
   End If
   '
   Y$ = REGSEL$("CHESELEC")
   If Len(Y$) <= 4 Then GoTo 30
   '
45 BANUFE$ = AJUSTI$(Y$, 40)
   IPASE# = CVD(Mid$(Y$, 49, 8))
   IPASX$ = TRIM$(FORMATNUM$(IPASE#, "F14.2"))
   Call APERBASDAT("CABAN")
   SQLZ$ = "SELECT * FROM CAJABANC "
   SQLZ$ = SQLZ$ + "where TipoMovim = 'CHRECI' "
   SQLZ$ = SQLZ$ + "and Status > 0 "
   SQLZ$ = SQLZ$ + "AND ABS(ValAbsComp - " & IPASX$ & ") < 0.005 "
   '
   Set CABASELE = CueCorri.OpenRecordset(SQLZ$, dbOpenDynaset)
   '
   With CABASELE    ' CajaBanc
     'IPASX$ = TRIM$(FORMATNUM$(IPASE#, "F14.2"))
     SQLX$ = "TipoMovim = 'CHRECI' "
     SQLX$ = SQLX$ + "AND Status > 0 "
     SQLX$ = SQLX$ + "AND DescriMov = '" & BANUFE$ & "' "
     SQLX$ = SQLX$ + "AND ABS(ValAbsComp - " & IPASX$ & ") < 0.005 "
     '
     .FindFirst SQLX$
     If .NoMatch = False Then
       '
       'ESTO ES PARA RECUPERAR CHEQUES QUE SALIERON EQUIVOCADOS
       If IsNull(!FECHENTRE) = True Or CVINT(!FECHENTRE) <= 33 Then
         BANUFE$ = !Descrimov
         If Mid$(BANUFE$, 19, 2) = "  " Then
           If Mid$(BANUFE$, 21, 1) <> " " Then
             BANUFE$ = Left$(BANUFE$, 18) + Mid$(BANUFE$, 21, 12) + "  " + Mid$(BANUFE$, 33, 8)
             .Edit
             !Descrimov = BANUFE$
             .Update
             '
             SQLY$ = "select * from CAJABANC "
             SQLY$ = SQLY$ + "WHERE TipoMovim = 'CHENTR' "
             SQLY$ = SQLY$ + "AND DescriMov = '" & BANUFE$ & "' "
             SQLY$ = SQLY$ + "AND  MID$(CodiCom_Cor,11,2) <> 'nl'"
             Set CABATEMP1 = CueCorri.OpenRecordset(SQLY$, dbOpenSnapshot)
             On Error Resume Next
             CABATEMP1.MoveFirst
             If Err < 1 Then
               On Error GoTo 0
               RCX$ = CABATEMP1!REFECOM
               FCX = CABATEMP1!FechConta
               CCX$ = CABATEMP1!CodiCom_Cor
               RFX$ = CABATEMP1!OBSERGEN
               FHX$ = CABATEMP1!FeHoReal
               '
               .Edit
               !DESTICHEQ = RCX$
               !FEDEPOSENT = FCX
               !FECHENTRE = FCX
               !Document = CCX$
               !REFESALCHE = RFX$
               !IDENSALCHE = FHX$
               !Status = 2
               .Update
             End If
           End If
         End If
       End If
       ' FIN RECUPERACION CHEQUES EQUIVOCADOS
       '
       YZ$ = String$(512, 0)
       Call REPLA(YZ$, Mid$(!Descrimov, 1, 32), 71, 32)
       Call REPLA(YZ$, Chr$(!Status), 121, 1)
       Call REPLA(YZ$, MKD$(!ValAbsComp), 145, 8)
       Call REPLA(YZ$, MKI$(CVINT(!FechEmisi)), 111, 2)
       Call REPLA(YZ$, MKI$(CVINT(!FechConta)), 61, 2)
       Call REPLA(YZ$, MKI$(CVINT(!FechAcred)), 63, 2)
       Call REPLA(YZ$, !CodiCom_Cor, 1, 12)
       Call REPLA(YZ$, !REFECOM, 13, 48)
       If IsNull(!DESTICHEQ) = False Then Call REPLA(YZ$, !DESTICHEQ, 293, 30)
       Call REPLA(YZ$, MKI$(CVINT(!FEDEPOSENT)), 275, 2)
       If IsNull(!Document) = False Then Call REPLA(YZ$, !Document, 277, 16)
       If IsNull(!DESTICHEQ) = False Then Call REPLA(YZ$, !DESTICHEQ, 323, 40)
       If IsNull(!REFESALCHE) = False Then Call REPLA(YZ$, !REFESALCHE, 363, 40)
       If IsNull(!IDENSALCHE) = False Then Call REPLA(YZ$, !IDENSALCHE, 403, 40)
       GoTo 50
     End If
     '
     Call MENSERR(24, "Imposible Presentar Datos del Cheque")
     GoTo 30
     '
   End With
     '
50 BANUFEA$ = BANUFE$
   X$ = OBJPLA$("MUESTRACHEGES", YZ$)
   If X$ <> "" Then
     If CVI(Mid$(X$, 63, 2)) <> CVI(Mid$(YZ$, 63, 2)) Then
       If DERACCE%("MOFEACRE", "Modificar Fecha de Acreditación Cheques") = 2 Then
         If COMALTER%("Confirme Modificación de Fecha de Acreditación\\No, Conservar\Si, Cambiar") = 2 Then
           IPASE# = CVD(Mid$(X$, 145, 8))
           IPASX$ = TRIM$(FORMATNUM$(IPASE#, "F14.2"))
           NUEFECHA% = CVI(Mid$(X$, 63, 2))
           FEX = NUEFECHA%
           NUFECHX$ = FECHATEX$(FEX)
           '
           Call APERBASDAT("CABAN")
           With CABASELE   ' CajaBanc
             SQLX$ = "TipoMovim = 'CHRECI' "
             SQLX$ = SQLX$ + "AND Status > 0 "
             SQLX$ = SQLX$ + "AND DescriMov = '" & BANUFEA$ & "' "
             SQLX$ = SQLX$ + "AND ABS(ValAbsComp - " & IPASX$ & ") < 0.005 "
             '
             .FindFirst SQLX$
             If .NoMatch = False Then
               .Edit
               !FechAcred = CVDAT(NUEFECHA%)
                 BANUFEX$ = BANUFEA$
                 Call REPLA(BANUFEX$, NUFECHX$, 33, 8)
               !Descrimov = BANUFEX$
               .Update
             End If
             '
             SQLY$ = "TipoMovim = 'CHENTR' "
             SQLY$ = SQLY$ + "AND DescriMov = '" & BANUFEA$ & "' "
             SQLY$ = SQLY$ + "AND  MID$(CodiCom_Cor,11,2) <> 'nl'"
             '
             .FindFirst SQLY$
60           If .NoMatch = False Then
               .Edit
                 BANUFEX$ = BANUFEA$
                 Call REPLA(BANUFEX$, NUFECHX$, 33, 8)
               !Descrimov = BANUFEX$
               .Update
               .FindNext SQLY$
               GoTo 60
             End If
           End With
           '
         End If
       End If
     End If
   End If
   '
   GoTo 30
   '
End Sub
'
Sub LISTRAFON()

     Call DESHASFECHA(DESD%, HAST%, PERIOD$)
 
     DESDE = CDbl(CVDAT(DESD%))
     HASTA = CDbl(CVDAT(HAST%))
     Call APERBASDAT("CABAN")
     SQLX$ = "SELECT * FROM CAJABANC "
     SQLX$ = SQLX$ + "where TipoMovim = 'TRAFON' "
     SQLX$ = SQLX$ + "AND CDBL(FechEmisi) >= " & DESDE & " "
     SQLX$ = SQLX$ + "AND CDBL(FechEmisi) <= " & HASTA & " "
     SQLX$ = SQLX$ + "ORDER BY FechEmisi"
     
     Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
     With CABATEMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
         CONT& = 0
         Do While Not .EOF
             '
             FECEMIS% = CVINT(!FechEmisi)
             COMPRO$ = !CodiCom_Cor
             REFERE$ = !REFECOM
             CUENTA$ = CODCUE$(!CueContaI)
             IMPDEB# = !IDEBECOMP
             IMPHAB# = !IHABECOMP
           
             Call REPLA(X$, MKI$(FECEMIS%), 1, 2)
             Call REPLA(X$, COMPRO$, 3, 12)
             Call REPLA(X$, REFERE$, 15, 48)
             Call REPLA(X$, CUENTA$, 63, 12)

             Call REPLA(X$, MKD$(IMPDEB#), 75, 8)
             Call REPLA(X$, MKD$(IMPHAB#), 83, 8)
           
             CONT& = CONT& + 1
     
             Call GRAREG("TRANSFON", X$, CONT&)
             .MoveNext
         Loop
       End If
     End With
     Call SETULTREG("TRANSFON", CONT&)
     Call HEADERS("TRANSFON", "")
     Call HEADERS("TRANSFON", "Periodo: " + PERIOD$)
     '
     Call FINAL("*TRANSFON")



End Sub

Private Sub pantalla1_Click()
    Option5.Value = True
    Option10.Value = True
    Call Command2_Click

End Sub

Private Sub pantalla2_Click()
    Option5.Value = True
    Option2.Value = True
    Call Command2_Click

End Sub

Private Sub pantalla3_Click()
    Option5.Value = True
    Option11.Value = True
    Call Command2_Click

End Sub

Private Sub pantalla5_Click()
    Option5.Value = True
    Option1.Value = True
    Call Command2_Click

End Sub

Private Sub recepciondechequessinrecibo_Click()
    Call Command13_Click
End Sub

Private Sub resumendiariotesoreria_Click()
    Option4.Value = True
    Option8.Value = True
    Call Command2_Click

End Sub

Private Sub salir_Click()
   Call Command1_Click
End Sub


Private Sub transferenciadefondos_Click()
    Call Command10_Click

End Sub


