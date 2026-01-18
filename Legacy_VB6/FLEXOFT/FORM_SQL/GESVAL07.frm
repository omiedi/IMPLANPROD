VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form GESVAL07 
   BackColor       =   &H0095C5E1&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gestion de Cartera de Valores"
   ClientHeight    =   3795
   ClientLeft      =   525
   ClientTop       =   2805
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
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4515
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   55
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H002C3D46&
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   5355
      ScaleHeight     =   1035
      ScaleWidth      =   4395
      TabIndex        =   40
      Top             =   0
      Width           =   4425
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   855
         TabIndex        =   61
         Top             =   240
         Width           =   885
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   62
            Top             =   0
            Width           =   12000
         End
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
         Height          =   650
         Left            =   1150
         Picture         =   "GESVAL07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   120
         Visible         =   0   'False
         Width           =   550
      End
      Begin VB.CommandButton AGenRep 
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
         Left            =   1740
         Picture         =   "GESVAL07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   120
         Visible         =   0   'False
         Width           =   550
      End
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
         Height          =   650
         Left            =   2330
         Picture         =   "GESVAL07.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Control de Cartera de Cheques"
         Top             =   120
         Visible         =   0   'False
         Width           =   550
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
         Picture         =   "GESVAL07.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Control de Cartera de Cheques"
         Top             =   840
         Visible         =   0   'False
         Width           =   600
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
         Height          =   650
         Left            =   3500
         Picture         =   "GESVAL07.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   120
         Width           =   550
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
         Height          =   650
         Left            =   2910
         Picture         =   "GESVAL07.frx":0D72
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Configuración de Funcionamiento"
         Top             =   120
         Visible         =   0   'False
         Width           =   550
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H0095C5E1&
      Caption         =   "OTROS MOVIMIENTOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2525
      Left            =   5355
      TabIndex        =   33
      Top             =   1155
      Width           =   4005
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   1260
         ScaleHeight     =   615
         ScaleWidth      =   2265
         TabIndex        =   44
         Top             =   1680
         Width           =   2295
         Begin VB.CommandButton Command20 
            Height          =   630
            Left            =   1740
            Picture         =   "GESVAL07.frx":11B4
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "Depósito Cupones de Crédito"
            Top             =   0
            Width           =   560
         End
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
            Height          =   630
            Left            =   0
            Picture         =   "GESVAL07.frx":23BE
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Transferencia de Fondos"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Conciliación Bancaria"
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
            Left            =   720
            TabIndex        =   46
            Top             =   75
            Width           =   1020
         End
      End
      Begin VB.PictureBox Picture9 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   945
         ScaleHeight     =   615
         ScaleWidth      =   2295
         TabIndex        =   48
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
            Height          =   630
            Left            =   1750
            Picture         =   "GESVAL07.frx":26C8
            Style           =   1  'Graphical
            TabIndex        =   52
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
            Height          =   630
            Left            =   0
            Picture         =   "GESVAL07.frx":29D2
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Recepción de Cheques sin Recibo Cobranza"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Movimiento  de Cheques"
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
            Left            =   720
            TabIndex        =   50
            Top             =   0
            Width           =   1050
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   630
         ScaleHeight     =   615
         ScaleWidth      =   2295
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
            Height          =   630
            Left            =   0
            Picture         =   "GESVAL07.frx":2B1C
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
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   720
            TabIndex        =   36
            Top             =   0
            Width           =   1380
         End
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   315
         ScaleHeight     =   615
         ScaleWidth      =   2295
         TabIndex        =   37
         Top             =   420
         Width           =   2325
         Begin VB.CommandButton Command22 
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
            Left            =   1750
            Picture         =   "GESVAL07.frx":2C66
            Style           =   1  'Graphical
            TabIndex        =   59
            ToolTipText     =   "Cheque Rechazado"
            Top             =   0
            Visible         =   0   'False
            Width           =   560
         End
         Begin VB.CommandButton Command21 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   1720
            Picture         =   "GESVAL07.frx":2F70
            Style           =   1  'Graphical
            TabIndex        =   58
            ToolTipText     =   "Visualizar Depósitos Bancarios"
            Top             =   0
            Width           =   560
         End
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
            Height          =   630
            Left            =   0
            Picture         =   "GESVAL07.frx":327A
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Depósito de Valores"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Dep. Cuenta Bancaria"
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
            Left            =   720
            TabIndex        =   39
            Top             =   0
            Width           =   1020
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0095C5E1&
      Caption         =   "CONSULTAS - LISTADOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3565
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   5160
      Begin VB.OptionButton Option12 
         Appearance      =   0  'Flat
         BackColor       =   &H0095C5E1&
         Caption         =   "Búsqueda de Cheque por CUIT "
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
         Height          =   525
         Left            =   525
         TabIndex        =   60
         Top             =   600
         Width           =   3300
      End
      Begin VB.OptionButton Option10 
         Appearance      =   0  'Flat
         BackColor       =   &H0095C5E1&
         Caption         =   "Historico de Cheques "
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
         Height          =   255
         Left            =   525
         TabIndex        =   47
         Top             =   1335
         Width           =   3855
      End
      Begin VB.OptionButton Option8 
         Appearance      =   0  'Flat
         BackColor       =   &H0095C5E1&
         Caption         =   "Resumen Diario de Tesorería"
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
         Height          =   330
         Left            =   525
         TabIndex        =   29
         Top             =   2220
         Width           =   3900
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H0095C5E1&
         Caption         =   "Cheques Entregados por Debitar"
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
         Height          =   330
         Left            =   525
         TabIndex        =   18
         Top             =   1905
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
         Picture         =   "GESVAL07.frx":3584
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
         Picture         =   "GESVAL07.frx":388E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2760
         Width           =   855
      End
      Begin VB.OptionButton Option6 
         Appearance      =   0  'Flat
         BackColor       =   &H0095C5E1&
         Caption         =   "Búsqueda de Cheque por Número"
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
         Height          =   330
         Left            =   525
         TabIndex        =   5
         Top             =   360
         Width           =   3375
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H0095C5E1&
         Caption         =   "SALIDA"
         BeginProperty Font 
            Name            =   "Arial"
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
            Appearance      =   0  'Flat
            BackColor       =   &H0095C5E1&
            Caption         =   "Pantalla"
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
            Left            =   600
            TabIndex        =   4
            Top             =   320
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option4 
            Appearance      =   0  'Flat
            BackColor       =   &H0095C5E1&
            Caption         =   "Impresión"
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
            Left            =   2100
            TabIndex        =   3
            Top             =   320
            Width           =   1275
         End
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H0095C5E1&
         Caption         =   "Detalle de Valores en Cartera"
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
         Height          =   330
         Left            =   525
         TabIndex        =   1
         Top             =   1605
         Width           =   3900
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         BorderWidth     =   2
         X1              =   0
         X2              =   5145
         Y1              =   1155
         Y2              =   1155
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
            Picture         =   "GESVAL07.frx":3B98
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
            Picture         =   "GESVAL07.frx":3EA2
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
            Picture         =   "GESVAL07.frx":41AC
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
      BackColor       =   &H0095C5E1&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   855
      Left            =   6855
      TabIndex        =   26
      Top             =   420
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0095C5E1&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   855
      Left            =   4515
      TabIndex        =   25
      Top             =   420
      Visible         =   0   'False
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
            Picture         =   "GESVAL07.frx":44B6
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
            Picture         =   "GESVAL07.frx":47C0
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
            Picture         =   "GESVAL07.frx":4ACA
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
            Picture         =   "GESVAL07.frx":4DD4
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
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "GESVAL07.frx":50DE
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "GESVAL07.frx":5312
      TabIndex        =   56
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Shape Shape1 
      Height          =   525
      Left            =   4515
      Top             =   4305
      Width           =   1380
   End
   Begin VB.Menu mnuTransacciones 
      Caption         =   "Transacciones"
      Begin VB.Menu mnuConvertirAjusteDebitosPorRechazos 
         Caption         =   "Convertir Ajuste en Debito por Rechazo"
      End
   End
   Begin VB.Menu mnuConsultasListados 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu mnuReportesVarios 
         Caption         =   "Reportes Varios"
      End
   End
   Begin VB.Menu mnuUtilidades 
      Caption         =   "Utilidades"
   End
   Begin VB.Menu mnuConfiguracion 
      Caption         =   "Configuración"
      Begin VB.Menu mnuVincularReporte 
         Caption         =   "Vincular Reporte con Formulario"
      End
      Begin VB.Menu linea1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConfiguraConciliacionAutomatica 
         Caption         =   "Conciliación Automática"
         Visible         =   0   'False
      End
      Begin VB.Menu linea2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuConfiguracionGeneral 
         Caption         =   "Configuración General"
      End
   End
End
Attribute VB_Name = "GESVAL07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Dim objexcel As Excel.Application
   Dim Libro As Excel.Workbook
   Dim HOJA As Excel.Worksheet


Sub ANULA_IMPORTACION()
   If DERACCE%("BORRACON", "Anulación de Importación") >= 2 Then
       Call SELECHO("CUEBANC")
       CUENTA_SELECCIONADA$ = TRIM$(VALACT1$("CUEBANC"))
       '
       If CUENTA_SELECCIONADA$ = "" Then Exit Sub
       CUEI% = CUEINT%(CUENTA_SELECCIONADA$)
    
      'PRESENTO LA LISTA DE SELECCION DE TODAS LAS CONCILIACIONES
      Call CARGALISEL("CONCITODO", "CONCILIA", "CODICONCI/A12;REFECONCI/A32", "CODICONCI <> '' AND CUENTACONTBANCO = " & CUEI% & " AND STATUS < 1 ", "NOMESS/DISTINCT")
      Call SELECHO("CONCITODO/SELECCION DE EXTRACTOS IMPORTADOS")
      CODICONCI$ = TRIM$(VALACT1$("CONCITODO"))
      If CODICONCI$ = "" Then Exit Sub
      '
      If CantRegistros&("CONCILIA", "STATUS > 0 ", "NOMESS") < 1 Then
        Call BORRAREGISTROS("CONCILIA", "CODICONCI = '" & CODICONCI$ & "'", REGAF&, "NOMESS")
        Call COMUNI("Se ha Anulado La Importación Seleccionada")
      Else
        Call COMUNI("Imposible Borrar\La Conciliación Elegida Se Encuentra Procesada")
      End If
   End If

End Sub



Sub CONCILIACION_AUTOMATICA()

   Call CREACAMPOS_CONCILIA

   Call LECTURAEXTRACTO(CODICONCI$)

   If CODICONCI$ <> "" Then
     Call GASTOEXTRACTO(CODICONCI$)
     Call CONCILIAEXTRACTO(CODICONCI$)
     Call PRESENTA_CONCILIA(CODICONCI$)
   End If
   
End Sub
Sub CONCILIAEXTRACTO(CODICONCI$)

    Screen.MousePointer = 11
    DES% = COMEJ%
    '
    CUECHECA$ = Control$("", "CUECHECA")
    CUEICHEQ% = CUEINT%(CUECHECA$)
    '
    'RECORRO EL CONCILIA
    FECH_COM_EJERCICIO = FETEXCOR1$(DES%)
    Condi$ = " CODICONCI = '" & CODICONCI$ & "'"
    Condi$ = Condi$ + " AND (COMPROBANTE = '' OR COMPROBANTE IS NULL)"
    Condi$ = Condi$ + " AND status < 1"
    Condi$ = Condi$ + " ORDER BY FECHA"
    SQLX$ = "SELECT * FROM CONCILIA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE " & Condi$
    '
    
    Dim RSTTMP As ADODB.Recordset
    Set RSTTMP = READSET(SQLX$)
    With RSTTMP
      Do While Not .EOF
         IMPODEBEX# = XVALO(!DEBE)
         IMPOHABEX# = XVALO(!HABER)
         If Abs(IMPODEBEX# + IMPOHABEX#) < 0.005 Then
           GoTo 80 'MOVENEXT
         End If
         CUEII% = XVALO(!CUENTACONTBANCO)
         IDREFE$ = SAFETEXT$(!ID_REFERENCIA)
         FE% = CVINT(!FECHA)
         FE1% = DIASPOST(FE%, 30)
         FECH_EXTRACTO = FETEXCOR1$(FE1%)
         CONCEP$ = SAFETEXT$(!CONCEPTO)
         REFECON$ = SAFETEXT$(!REFECONCI)
         NUORD& = XVALO(!ORDCONCILIA)
         Encontrado% = 0
         ' EN CASO QUE NO SEA 0 APLICO FORMANTNUM PARA QUE REEMPLACE LA COMA POR UN PUNTO
         ' CUANDO ES 0 NO POR QUE SI NO DA ERROR EN LA CONSULTA
         IMPODEBE_PTO = 0: IMPOHABE_PTO = 0
         If IMPODEBEX# <> 0 Then IMPODEBE_PTO = TRIM$(FORMATNUM$(IMPODEBEX#, "F16.5"))
         If IMPOHABEX# <> 0 Then IMPOHABE_PTO = TRIM$(FORMATNUM$(IMPOHABEX#, "F16.5"))
         ' BUSCA LA COINCIDENCIA EN MOVIMIENTOS QUE ESTEN HASTA 4 MESES HACIAS ATRAS DE LA FECHA QUE FIGURA EN EL EXTRACTO
         FechaInicioBuscaMovimientoAConciliar% = DIASANTES(FE%, 30)
         If IMPODEBEX# > 0 Then
           Condi$ = "IHABECOMP = " & IMPODEBE_PTO & " AND FECHACRED <= '" & FECH_EXTRACTO & "'"
           Condi$ = Condi$ + " AND FECHACRED >= '" & FETEXCOR1$(FechaInicioBuscaMovimientoAConciliar%) & "'"
           Condi$ = Condi$ + " AND (TIPOMOVIM = 'CHENTR' or TIPOMOVIM = 'TRAFON' or TIPOMOVIM = 'DEPOBAN' or TIPOMOVIM='CHRECI')"
           Condi$ = Condi$ + " AND (MarConsi = '' or MarConsi IS NULL )"
           Condi$ = Condi$ + " AND CueContaI = " & CUEII%
         ElseIf IMPOHABEX# > 0 Then
            Condi$ = "IDEBECOMP = " & IMPOHABE_PTO & " AND FECHACRED <= '" & FECH_EXTRACTO & "'"
            Condi$ = Condi$ + " AND FECHACRED >= '" & FETEXCOR1$(FechaInicioBuscaMovimientoAConciliar%) & "'"
            Condi$ = Condi$ + " AND (TIPOMOVIM = 'CHENTR' or TIPOMOVIM = 'TRAFON' or TIPOMOVIM = 'DEPOBAN' or TIPOMOVIM='CHRECI')"
            Condi$ = Condi$ + " AND (MarConsi = '' or MarConsi IS NULL )"
            Condi$ = Condi$ + " AND CueContaI = " & CUEII%
         End If
         '
         'VERIFICO SI MACHEA CON LA CONDICION
         COINCIDE% = CantRegistros&("CAJABANC", Condi$, "NOMESS")
         '
         'SI HAY MAS DE UN REGISTRO BUSCO POR EL NUMERO DE CHEQUE
         If COINCIDE% > 1 Then
           Condi$ = Condi$ + " AND SUBSTRING(DESCRIMOV,19,14)  Like '%" & IDREFE$ & "%'"
           If CantRegistros("CAJABANC", Condi$, "NOMESS") = 1 Then Encontrado% = 1
         'CASO QUE YA SE DETECTO UN SOLO REGISTRO COINCIDENTE
         ElseIf COINCIDE% = 1 Then
           Encontrado% = 1
         End If
         '
         'GRABO NUMERO DE COMPRBANTE Y FECHACONTABLE
         If Encontrado% = 1 Then
            COMPROBTE$ = VALOBADA("CAJABANC", "CODICOM_COR", Condi$, "NOMESS")
            FECONTA = FETEXCOR1$(CVINT(VALOBADA("CAJABANC", "FECHCONTA", Condi$, "NOMESS")))
            DESCIRIMOVIMIENTO$ = VALOBADA("CAJABANC", "DESCRIMOV", Condi$, "NOMESS")
            CONDI1$ = "CODICONCI = '" & CODICONCI$ & "' AND OrdConcilia = " & NUORD&
            Call ACTUREGISTRO("CONCILIA", "FECHACOMPROBANTE", CONDI1$, FECONTA, REGAF&, "NOMESS")
            Call ACTUREGISTRO("CONCILIA", "COMPROBANTE", CONDI1$, COMPROBTE$, REGAF&, "NOMESS")
            Call ACTUREGISTRO("CONCILIA", "COMPROB_DE_TABLA", CONDI1$, "CABAN", REGAF&, "NOMESS")
            Call ACTUREGISTRO("CONCILIA", "CONCEPTO_SISTEMA", CONDI1$, DESCIRIMOVIMIENTO$, REGAF&, "NOMESS")
            GoTo 80 ' SI LO GRABO PASA AL SIGUIENTE
         End If
         '
         '*-----------------------------------------------------------------------
         
         'SI NO ENCONTRO EN CAJABANC BUSCA EN DETASIEN
         Condi$ = "FECHASIEN <= '" & FECH_EXTRACTO & "'"
         Condi$ = Condi$ + " AND FECHASIEN >= '" & FETEXCOR1$(FechaInicioBuscaMovimientoAConciliar%) & "'"
'         CONDI$ = CONDI$ + " AND FECHASIEN >= '" & FECH_COM_EJERCICIO & "'"
         Condi$ = Condi$ + " AND CUECOINT = " & CUEII%
         Condi$ = Condi$ + " AND statpase > 0"
         Condi$ = Condi$ + " AND (MarConsi = '' or MarConsi IS NULL ) "
         IMPODEBE_ASIEN# = XVALO(VALOBADA("DETAASIEN", "IDEBEPASE", Condi$, "NOMESS"))
         IMPOHABE_ASIEN# = XVALO(VALOBADA("DETAASIEN", "IHABEPASE", Condi$, "NOMESS"))
         '
         IMPODEBE_PTO = 0: IMPOHABE_PTO = 0
         If IMPODEBEX# <> 0 Then IMPODEBE_PTO = TRIM$(FORMATNUM$(IMPODEBEX#, "F16.5"))
         If IMPOHABEX# <> 0 Then IMPOHABE_PTO = TRIM$(FORMATNUM$(IMPOHABEX#, "F16.5"))
         
         If IMPODEBEX# > 0 Then
            Condi$ = Condi$ + " AND IHABEPASE = " & IMPODEBE_PTO
         ElseIf IMPOHABEX# > 0 Then
            Condi$ = Condi$ + " AND IDEBEPASE = " & IMPOHABE_PTO
         End If
         '
         COINCIDE% = CantRegistros&("DETAASIEN", Condi$, "NOMESS")
         If COINCIDE% = 1 Then
            CONDI1$ = "CODICONCI = '" & CODICONCI$ & "' AND OrdConcilia = " & NUORD&
            IDASIEN$ = VALOBADA("DETAASIEN", "NUIDASIEN", Condi$, "NOMESS")
            FEASIEN = FETEXCOR1$(CVINT(VALOBADA("DETAASIEN", "FECHASIEN", Condi$, "NOMESS")))
            DESCIRIMOVIMIENTO$ = VALOBADA("DETAASIEN", "REFEPASE", Condi$, "NOMESS")
            Call ACTUREGISTRO("CONCILIA", "COMPROBANTE", CONDI1$, IDASIEN$, REGAF&, "NOMESS")
            Call ACTUREGISTRO("CONCILIA", "FECHACOMPROBANTE", CONDI1$, FEASIEN, REGAF&, "NOMESS")
            Call ACTUREGISTRO("CONCILIA", "COMPROB_DE_TABLA", CONDI1$, "DIAR", REGAF&, "NOMESS")
            Call ACTUREGISTRO("CONCILIA", "CONCEPTO_SISTEMA", CONDI1$, DESCIRIMOVIMIENTO$, REGAF&, "NOMESS")
         End If
         '
80      .MoveNext
      Loop
    End With
    RSTTMP.Close
    Set RSTTMP = Nothing
    Screen.MousePointer = 1
End Sub

Sub DESPEJA_PENDIENTES()

   With CONCIMAN07
        .MSF1.Redraw = False
        .MSF3.Redraw = False
       i& = 0
       Do While i& < .MSF3.Rows - 1
             i& = i& + 1
             COMPROB$ = TRIM$(.MSF3.TextMatrix(i&, 2))
             FECON$ = TRIM$(.MSF3.TextMatrix(i&, 1))
             DETREF$ = TRIM$(.MSF3.TextMatrix(i&, 4))
             DEBEX# = XVALO(.MSF3.TextMatrix(i&, 5))
             HABEX# = XVALO(.MSF3.TextMatrix(i&, 6))
             For J& = 1 To .MSF1.Rows - 1
               If TRIM$(.MSF1.TextMatrix(J&, 0)) <> "*" Then
                 If TRIM$(.MSF1.TextMatrix(J&, 2)) = COMPROB$ Then
                   If TRIM$(.MSF1.TextMatrix(J&, 1)) = FECON$ Then
                     If XVALO(.MSF1.TextMatrix(J&, 5)) = DEBEX# Then
                       If XVALO(.MSF1.TextMatrix(J&, 6)) = HABEX# Then
                         .MSF1.TextMatrix(J&, 0) = "*"
                         Call borrafila(.MSF3, i&) 'BORRO LA FILA DE PENDIENTES
                         i& = i& - 1 'RESTO UNO POR QUE SE ELIMINO LA FILA LEIDA
                         Exit For
                       End If
                     End If
                   End If
                 End If
               End If
            Next J&
        Loop
        '
        'BORRO LOS MARCADOS COMO COINCIDENTES
        For k& = 1 To .MSF1.Rows - 1
          If .MSF1.TextMatrix(k&, 0) = "*" Then .MSF1.TextMatrix(k&, 0) = ""
        Next k&
        '
       .MSF1.Redraw = True
       .MSF3.Redraw = True
   End With

End Sub

Function borrafila(MSF As MSFlexGrid, REG&)

        For i& = REG& To MSF.Rows - 2
          For J& = 1 To MSF.Cols - 1
            MSF.TextMatrix(i, J) = MSF.TextMatrix(i + 1, J)
          Next
        Next
        MSF.Rows = MSF.Rows - 1
        
End Function

Sub GENERA_PENDIENTE_EN_LIMAYO01(CUEII%, DESPEN%, HASPEN%)
   
   If DESPEN% < 1 Then
      DESPEN% = FECHANUM("01/01/80")
   End If
   If HASPEN% < 1 Then
      HASPEN% = (FINMESACT%)
   End If
       
   FEINISQL = FETEXCOR1$(DESPEN%)
   FEFINSQL = FETEXCOR1$(HASPEN%)
   FEINI = CDbl(CVDAT(DESPEN%))
   FEFIN = CDbl(CVDAT(HASPEN%))

   Screen.MousePointer = 11
   SQLX$ = "SELECT * FROM DetaAsien "
   SQLX$ = SQLX$ + "WHERE (CueCoInt=" & CUEII%
   ' SQLX$ = SQLX$ + " AND SucuAsien = " & SUCUI%
   SQLX$ = SQLX$ + " AND Fechasien >= '" & FEINISQL & "'"
   SQLX$ = SQLX$ + " AND Fechasien <= '" & FEFINSQL & "'"
   SQLX$ = SQLX$ + " AND statpase > 0 ) "
   'SQLX$ = SQLX$ + " AND (MarConsi = '' or ISNULL(MarConsi) = True or IsEmpty(MarConsi) = True) "
   SQLX$ = SQLX$ + " AND (MarConsi = '' or MarConsi IS NULL ) "
   SQLX$ = SQLX$ + "ORDER BY FECHASIEN, ORFEASIEN ASC;"
   'Set PASEASIEN = Contable.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
   Dim ENCASI As ADODB.Recordset
   Dim PASEASIEN As ADODB.Recordset
   Set PASEASIEN = New ADODB.Recordset
24  On Error Resume Next
    PASEASIEN.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 24
   End If
   On Error GoTo 0

   JJ& = 0
   TRANSPOR# = 0
20 With PASEASIEN
     On Error Resume Next
     .MoveFirst
     If Err Then GoTo 30
     On Error GoTo 0
     Do While Not .EOF
       NUAS$ = !nuidasien
       TIASI% = 0
       'Encasien.FindFirst "NUIDASIEN = '" & NUAS$ & "'"
       SQLX$ = "SELECT * FROM ENCABEASIEN "
       SQLX$ = SQLX$ + " WHERE NUIDASIEN = '" & NUAS$ & "'"
       Dim Encasien As ADODB.Recordset
       Set Encasien = FLEXCONN.Execute(FILTSQL$(SQLX$))

       If Encasien.EOF = False Then
         TIASI% = Encasien!TIPOASIEN
       End If
       '
       If TIASI% = 0 Then  ' DESCARTA ASIENTOS RESUMEN
         Z$ = REGBLAN$("!LIMAYOA1")
         Call REPLA(Z$, MKI$(CVINT(!FECHASIEN)), 1, 2)
         Call REPLA(Z$, !nuidasien, 3, 12)
         Call REPLA(Z$, MKI$(CVINT(!FECHASIEN)), 15, 2)
           DEPASE$ = !RefePase
         Call REPLA(Z$, DEPASE$, 17, 32)
           IDEBE# = !idebepase
         Call REPLA(Z$, MKD$(IDEBE#), 49, 8)
           IHABE# = !ihabepase
         Call REPLA(Z$, MKD$(IHABE#), 57, 8)
         Call REPLA(Z$, "DIAR", 73, 8)
         If IDEBE# >= 0.005 Or IHABE# >= 0.005 Then
           JJ& = JJ& + 1
           Call GRAREG("!LIMAYOA1", Z$, JJ&)
         End If
       End If
     .MoveNext
     Loop
   End With
   Set PASEASIEN = Nothing
30 On Error GoTo 0
   Call APERBASDAT("CABAN")
   Screen.MousePointer = 11
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "where (CueContaI = " & CUEII% & " "
   'SQLX$ = SQLX$ + " AND (MarConsi = '' or isnull(MarConsi) = True OR IsEmpty(MarConsi) = True) "
   SQLX$ = SQLX$ + " AND (MarConsi = '' OR MarConsi IS NULL ) "
   SQLX$ = SQLX$ + " AND FechConta >= '" & FEINISQL & "'"
   SQLX$ = SQLX$ + " AND FechConta <= '" & FEFINSQL & "') "
   SQLX$ = SQLX$ + " ORDER BY FECHCONTA ASC, FEHOREAL ASC;"
   '
   'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim CABATEMP As ADODB.Recordset
   Set CABATEMP = New ADODB.Recordset
26 On Error Resume Next
   CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 26
   End If
   On Error GoTo 0
   '
   With CABATEMP
     On Error Resume Next
     .MoveFirst
     If Err Then
        GoTo 40
     End If
     On Error GoTo 0
     Do While Not .EOF
         IDEBE# = !idebecomp
         IHABE# = !ihabecomp
       If Int(CDbl(!fechconta)) < FEINI Then
           TRANSPOR# = TRANSPOR# + IDEBE# - IHABE#
         ElseIf Int(CDbl(!fechconta)) <= FEFIN Then
           Z$ = REGBLAN$("LIMAYOA1")
           Call REPLA(Z$, MKI$(CVINT(!fechconta)), 1, 2)
           Call REPLA(Z$, !CodiCom_Cor, 3, 12)
           Call REPLA(Z$, MKI$(CVINT(!fechacred)), 15, 2)
             DEPASE$ = Left$(!DESCRIMOV, 32)
           Call REPLA(Z$, DEPASE$, 17, 32)
           Call REPLA(Z$, MKD$(IDEBE#), 49, 8)
           Call REPLA(Z$, MKD$(IHABE#), 57, 8)
           Call REPLA(Z$, "CABAN", 73, 8)
           If IDEBE# >= 0.005 Or IHABE# >= 0.005 Then
             JJ& = JJ& + 1
             Call GRAREG("!LIMAYOA1", Z$, JJ&)
           End If
       End If
     .MoveNext
     Loop
   End With
   '
40 Call SETULTREG("!LIMAYOA1", JJ&)
   'Call FILESORT("LIMAYOAC", "", 1, 3, "ASC")
   Call TRACENTRAL("LIMAYOA1", "")
   Call FILESORT("LIMAYOA1", "", 1, 3, "ASC")
   Call TRALOCAL("LIMAYOA1", "")
End Sub

Sub PRESENTA_CONCILIA(CODICONCI$, Optional MODO%, Optional FechaDesdeMovNoConciliados%)
    'MODO% = 0 CUANDO VIENE DE LA IMPORTACION
    'MODO% = 1 CUANDO VISUALIZA DESDE LA OPCION DE VER O MODIFICAR UN EXTRACTO IMPORTADO
    i& = 0
    '
    CONCIMAN07.MSF1.Cols = 11
    CONCIMAN07.msf2.Cols = 10
    CONCIMAN07.MSF3.Cols = 10
    CONCIMAN07.MSF1.Rows = 1
    CONCIMAN07.msf2.Rows = 1
    CONCIMAN07.MSF3.Rows = 1
    Screen.MousePointer = 11
    '
    Call PoneEnCeroNull("CONCILIA", "COMPROBANTE", REGAF&)
    'RECORRO EL CONCILIA
    'MUESTRO EN LA GRILLA 1 LOS QUE TIENEN ASIGNADO UN COMPROBANTE
    SQLX$ = "SELECT * FROM CONCILIA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE CODICONCI = '" & CODICONCI$ & "'"
    SQLX$ = SQLX$ + " AND COMPROBANTE <> ''"
    SQLX$ = SQLX$ + " AND status < 1"
    SQLX$ = SQLX$ + " ORDER BY FECHA"
    '
    CONCIMAN07.MSF1.Redraw = False
    Dim RSTTMP As ADODB.Recordset
    Set RSTTMP = READSET(SQLX$)
    With RSTTMP
      Do While Not .EOF
         IMPODEBEX# = XVALO(!DEBE)
         IMPOHABEX# = XVALO(!HABER)
         CUEII% = XVALO(!CUECONTAI)
         IDREFE$ = SAFETEXT$(ID_REFERENCIA)
         FECONTA% = CVINT(!FechaComprobante)
         FECH1 = FETEXCOR1$(FECONTA%)
         FECH_EXTR% = CVINT(!FECHA)
         COMPRBTE$ = SAFETEXT$(!COMPROBANTE)
         CUEIIGASTO% = XVALO(!CUECONTAI)
         CONCEP$ = SAFETEXT$(!CONCEPTO)
         REFECON$ = SAFETEXT$(!ID_REFERENCIA)
         NUORD% = XVALO(!ORDCONCILIA)
         DETABLA$ = SAFETEXT$(!COMPROB_DE_TABLA)
         ORDVINCULO% = XVALO(!ORDENVINCULO)
         With CONCIMAN07.MSF1
            REG& = REG& + 1
            .Rows = REG& + 1
            .TextMatrix(REG&, 1) = FECHATEX$(FECONTA%)
            .TextMatrix(REG&, 2) = COMPRBTE$
            .TextMatrix(REG&, 3) = FECHATEX$(FECH_EXTR%) 'FECHA DEL COMPROBANTE
            .TextMatrix(REG&, 4) = CONCEP$
            .TextMatrix(REG&, 6) = TRIM$(FORMATNUM$(IMPODEBEX#, "F12.2"))
            .TextMatrix(REG&, 5) = TRIM$(FORMATNUM$(IMPOHABEX#, "F12.2"))
            .TextMatrix(REG&, 7) = NUORD%
            .TextMatrix(REG&, 8) = DETABLA$
            .TextMatrix(REG&, 9) = SAFETEXT$(CUEIIGASTO%)
            .TextMatrix(REG&, 10) = CStr(ORDVINCULO%)
         End With
      .MoveNext
      Loop
    End With
    CONCIMAN07.MSF1.Redraw = True
    RSTTMP.Close
    Set RSTTMP = Nothing
    REG& = 0
    '
    'RECORRO EL CONCILIA
    'MUESTRO EN LA GRILLA 2 LOS QUE NO TIENEN COMPROBANTE
    SQLX$ = "SELECT * FROM CONCILIA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE CODICONCI = '" & CODICONCI$ & "'"
    SQLX$ = SQLX$ + " AND (COMPROBANTE = '' OR COMPROBANTE IS NULL)"
    SQLX$ = SQLX$ + " AND status < 1"
    SQLX$ = SQLX$ + " ORDER BY FECHA"
    '
    CONCIMAN07.msf2.Redraw = False
    Dim RSTTMP2 As ADODB.Recordset
    Set RSTTMP2 = READSET(SQLX$)
    With RSTTMP2
      Do While Not .EOF
         IMPODEBEX# = XVALO(!DEBE)
         IMPOHABEX# = XVALO(!HABER)
         CUEII% = XVALO(!CUECONTAI)
         IDREFE$ = SAFETEXT$(ID_REFERENCIA)
         FECONTA% = CVINT(!FechaComprobante)
         FECH1 = FETEXCOR1$(FECONTA%)
         FECH_EXTR% = CVINT(!FECHA)
         CONCEP$ = SAFETEXT$(!CONCEPTO)
         REFECON$ = SAFETEXT$(!ID_REFERENCIA)
         COMPRBTE$ = SAFETEXT$(!COMPROBANTE)
         NUORD% = XVALO(!ORDCONCILIA)
         DETABLA$ = SAFETEXT$(!COMPROB_DE_TABLA)
         With CONCIMAN07.msf2
            REG& = REG& + 1
            .Rows = REG& + 1
            .TextMatrix(REG&, 1) = FECHATEX$(FECONTA%) 'FECHA DEL COMPROBANTE
            .TextMatrix(REG&, 2) = COMPRBTE$
            .TextMatrix(REG&, 3) = FECHATEX$(FECH_EXTR%)
            .TextMatrix(REG&, 4) = CONCEP$
            .TextMatrix(REG&, 6) = TRIM$(FORMATNUM$(IMPODEBEX#, "F12.2"))
            .TextMatrix(REG&, 5) = TRIM$(FORMATNUM$(IMPOHABEX#, "F12.2"))
            .TextMatrix(REG&, 7) = NUORD%
            .TextMatrix(REG&, 8) = DETABLA$
            .TextMatrix(REG&, 9) = ""
         End With
      .MoveNext
      Loop
    End With
    CONCIMAN07.msf2.Redraw = True
    RSTTMP2.Close
    Set RSTTMP2 = Nothing
    REG& = 0
    '
    Call RecFechAcredCheques 'PROCESO DE RECUPERACION DE FECHA DE ACRDIDTACION DE CHEQUES
    CUEII% = CUENTA_BCO_CONCILIA%(CODICONCI$)
    DES% = COMEJ%
    If FechaDesdeMovNoConciliados% > 0 Then DES% = FechaDesdeMovNoConciliados%
    Call GENERA_PENDIENTE_EN_LIMAYO01(CUEII%, DES%, 0)
    '
    'RECORRO EL ARCHIVO LIMAYO01 CON LOS PENDIENTES DEL SISTEMA PARA MOSTRARLOS EN LA GRILLA 3
    With CONCIMAN07.MSF3
        .Redraw = False
        Z$ = Space$(128)
        .Rows = 1
        k& = 0
        For U& = 1 To ULTREG&("!LIMAYOA1")
         X$ = REGLEIDO$("!LIMAYOA1", U&)
         Z$ = Space$(128)
           FEX = CVI(Left$(X$, 2))
           FECONTA1 = FECHATEX$(FEX)
           COCOM$ = TRIM$(Mid$(X$, 3, 12))
           FEX = CVI(Mid$(X$, 15, 2))
           FEACREDIT$ = FECHATEX$(FEX)
           DESCRIMOVIMIENTO$ = TRIM$(Mid$(X$, 17, 32))
           IDEBE# = CVD(Mid$(X$, 49, 8))
           IHABE# = CVD(Mid$(X$, 57, 8))
           TABLA$ = TRIM$(Mid$(X$, 73, 8))
           
'           'VALIDO EL CASO QUE SEA UN DEPOSITO BANCARIO CON MAS DE UN CHEQUE
'           'PARA QUE PRESENTE LOS CHEQUES INDIVIDUALMENTE POR QUE EN EL ARCHIVO ESTAN ACUMULADOS
'           If UCase$(Left$(COCOM$, 3)) = "DB-" Then
'              CONDI$ = "CODICOM_LAR = '" & COCOM$ & "' AND TIPOMOVIM = 'CHENTR' "
'              If CantRegistros("CAJABANC", CONDI$, "NOMESS") > 1 Then
'                SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK)"
'                SQLX$ = SQLX$ + " WHERE " & CONDI$
'                SQLX$ = SQLX$ + " AND (MarConsi = '' OR MarConsi IS NULL)"
'                Dim DEPOTMP As New ADODB.Recordset
'                Set DEPOTMP = READSET(SQLX$)
'                With DEPOTMP
'                  Do While Not .EOF
'                    DESCRIMOVIMIENTO$ = SAFETEXT$(!DescriMov$)
'                    IDEBE# = XVALO(!IDEBECOMP)
'                    IHABE# = XVALO(!IHABECOMP)
'                    With CONCIMAN07.MSF3
'                        CONDI$ = "COMPROBANTE = '" & COCOM$ & "' AND STATUS < 1 "
'                        CONDI$ = CONDI$ + " AND ID_REFERENCIA LIKE '%" & TRIM$(Mid$(DESCRIMOVIMIENTO$, 19, 14)) & "%'"
'                        IMPDEBE_PTO = 0: IMPHABE_PTO = 0
'                        If IDEBE# > 0 Then
'                          IMPDEBE_PTO = TRIM$(FORMATNUM$(IDEBE#, "F16.5"))
'                          CONDI$ = CONDI$ + " AND DEBE = " & IMPDEBE_PTO
'                        Else
'                         IMPHAEBE_PTO = TRIM$(FORMATNUM$(IHABE#, "F16.5"))
'                         CONDI$ = CONDI$ + " AND HABER = " & IMPHAEBE_PTO
'                        End If
'                        If CantRegistros&("CONCILIA", CONDI$, "NOMESS") < 1 Then
'                            K& = K& + 1
'                            .Rows = K& + 1
'                            .TextMatrix(K&, 1) = FECONTA1
'                            .TextMatrix(K&, 2) = COCOM$
'                            .TextMatrix(K&, 3) = FEACREDIT$
'                            .TextMatrix(K&, 4) = DESCRIMOVIMIENTO$
'                            If IDEBE# > 0 Then
'                              .TextMatrix(K&, 5) = FORMATNUM$(IDEBE#, "F14.2")
'                            Else
'                              .TextMatrix(K&, 6) = FORMATNUM$(IHABE#, "F14.2")
'                            End If
'                            .TextMatrix(K&, 7) = ""
'                            .TextMatrix(K&, 8) = TABLA$
'                            .TextMatrix(K&, 9) = ""
'                        End If
'                    End With
'                  .MoveNext
'                  Loop
'
'                End With
'                GoTo 50 'NEXT
'              End If
'           End If
           k& = k& + 1
           .Rows = k& + 1
           .TextMatrix(k&, 1) = FECONTA1
           .TextMatrix(k&, 2) = COCOM$
           .TextMatrix(k&, 3) = FEACREDIT$
           .TextMatrix(k&, 4) = DESCRIMOVIMIENTO$
           If IMPODEBEX# <> 0 Then IMPODEBE_PTO = TRIM$(FORMATNUM$(IMPODEBEX#, "F16.5"))
           If IMPOHABEX# <> 0 Then IMPOHABE_PTO = TRIM$(FORMATNUM$(IMPOHABEX#, "F16.5"))
           If IDEBE# > 0 Then
               .TextMatrix(k&, 5) = FORMATNUM$(IDEBE#, "F14.2")
           Else
               .TextMatrix(k&, 6) = FORMATNUM$(IHABE#, "F14.2")
           End If

          .TextMatrix(k&, 7) = ""
          .TextMatrix(k&, 8) = TABLA$
          .TextMatrix(k&, 9) = ""

50     Next U&
       .Redraw = True
   End With
   
   Call CIERRARCH("!LIMAYOA1")
   '
   'ELIMINAR DE PENDIENTES LAS FILAS CONCILIADAS
   Call DESPEJA_PENDIENTES
   '
   CUENTABANCARIA$ = TRIM$(DENOCUE$(CUEII%))
   REFERENCIA1$ = TRIM$(VALOBADA("CONCILIA", "REFECONCI", "CODICONCI = '" & CODICONCI$ & "'", "NOMESS"))
   TITU$ = "Situación de Conciliación del Banco: " & CUENTABANCARIA$ & " (" & REFERENCIA1$ & ")"
   CONCIMAN07.Caption = TITU$
   CONCIMAN07.CODICONCI$ = CODICONCI$
   If MODO% = 1 Then
     If CantRegistros("CONCILIA", "CODICONCI = '" & CODICONCI$ & "' AND STATUS > 0 ", "NOMESS") > 0 Then
        CONCIMAN07.Command2.Enabled = False
        CONCIMAN07.Command3.Enabled = False
        CONCIMAN07.Command5.Enabled = False
        CONCIMAN07.Command7.Enabled = False
        CONCIMAN07.Command6.Enabled = False
        Screen.MousePointer = 1
        Call COMUNI("Extracto Conciliado No Podra Realizar Modificaciones")
        Screen.MousePointer = 11
     End If
   End If
   '
   Call CONCIMAN07.SUMAMSF(CONCIMAN07.MSF1)
   Call CONCIMAN07.SUMAMSF(CONCIMAN07.msf2)
   Call CONCIMAN07.SUMAMSF(CONCIMAN07.MSF3)
   
   '
   Screen.MousePointer = 1
   CONCIMAN07.Show 1
   '
End Sub

Sub GASTOEXTRACTO(CODICONCI$)
    '
    'RECORRE LA TABLA GASEXTRA Y SI ES LA MISMA CUENTA INTENTA GRABAR LA CUENTA CONTABLE
    'Y GASTOS EN EL CAMPO COMPROBANTE
    Screen.MousePointer = 11
    CUEI% = CUENTA_BCO_CONCILIA%(CODICONCI$)
    CUENTA_SELECCIONADA$ = CODCUE$(CUEI%)
    '
    RFF$ = RECFILT$("GASEXTRA", 1, AJUSTI$(CUENTA_SELECCIONADA$, 12))
    For U& = 1 To Len(RFF$) Step 4
        REG& = CVS(Mid$(RFF$, U&, 4))
        XX$ = REGLEIDO$("GASEXTRA", REG&)
        CONCEP$ = TRIM$(Mid$(XX$, 13, 128))
        CUENTACONT% = CUEINT%(TRIM$(Mid$(XX$, 141, 12)))
        Condi$ = "(COMPROBANTE = '' OR COMPROBANTE IS NULL) AND CONCEPTO LIKE '%" & CONCEP$ & "%' AND STATUS < 1 "
        Condi$ = Condi$ + " AND CODICONCI = '" & CODICONCI$ & "'"
        Call ACTUREGISTRO("CONCILIA", "CUECONTAI", Condi$, CUENTACONT%, REGAF&, "NOMESS")
        Call ACTUREGISTRO("CONCILIA", "COMPROBANTE", Condi$, "GASTOS", REGAF&, "NOMESS")
    Next U&
    Screen.MousePointer = 1
    '
End Sub

Sub LECTURAEXTRACTO(CODICONCI$)
        RNDX$ = ""
        J& = 0
        '
        'MENSAJE DE AVISO
        Call COMUNI("Los Archivos Excel a Importar No Deberan Tener Filas Vacias\Y Los Mismos Deberan Tener Configuradas Las Posiciones de Lectura")
        '
        'LISTA DE SELECCION DE BANCOS
10      Call SELECHO("CUEBANC/BANCOS")
        CUENTA_SELECCIONADA$ = TRIM$(VALACT1$("CUEBANC"))
        '
        If CUENTA_SELECCIONADA$ = "" Then Exit Sub
        '
        'FUNCION DEVUELVE POSICIONES DE LECTURA DE COLUMNAS DE PLANILLA EXCEL
        'VALOR DE FUNCION (0) FALTA INFORMACION MANDATARIA, (1) TODAS - HABER (2) TODAS
        VALIDA_POS% = POSICIONES_CFGCONCI%(CUENTA_SELECCIONADA$, FILA_INICIO%, COL_INICIO%, COL_FECHA%, COL_CONCEPTO%, COL_IDENT%, COL_DEBE%, COL_HABER%)
        If VALIDA_POS% < 1 Then GoTo 10
        '
        DEBE_Y_HABER% = 1 'INFORAMCION DE IMPORTES EN DEBE
        If VALIDA_POS% = 2 Then DEBE_Y_HABER% = 2 'INFORAMCION DE IMPORTES EN DEBE Y HABER
        '
        'SELECCION DE PLANILLA EXCEL
15      CD1.FILTER = "ARCHIVOS EXCEL(*.XLS)|*.XLS|(*.XLSX)|*.XLSX|"
        CD1.FilterIndex = 1
        CD1.ShowOpen
        Ruta$ = CD1.FileName
        If TRIM(Ruta$) = "" Then GoTo 10
        '
        'INGRESO DE REFERENCIA DE EXTRACTO
20      DESCRI$ = TRIM$(InputBox$("Ingrese Texto ", "Referencia de Conciliación", DESCRI$))
        
        
        If DESCRI$ = "" Then GoTo 10
        '
        Screen.MousePointer = 11
        Set objexcel = New Excel.Application
        Set Libro = objexcel.Workbooks.Open(Ruta$)
        Set HOJA = Libro.Sheets(1)
        '
        'SOBREESCRIBE$ = VALIDA_CONCILIACION$(CUENTA_SELECCIONADA$, COL_FECHA%, FILA_INICIO%, COL_INICIO%, COL_CONCEPTO%, COL_IDENT%, COL_DEBE%, COL_HABER%)
        '
        'LECTURA  Y GRABACION DE EXTRACTO EN TABLA CONCILIA
        Dim CONCTMP As ADODB.Recordset
        Set CONCTMP = New ADODB.Recordset
        SQLX$ = "SELECT * FROM CONCILIA WHERE CODICONCI <> 'A' " 'CONDICION INVALIDA
25      On Error Resume Next
        CONCTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
          On Error GoTo 0
          Call CapturaErrorOpen
          GoTo 25
        End If
        On Error GoTo 0
        CUEI% = CUEINT%(CUENTA_SELECCIONADA$)
        FECHA_MAX% = MAXFECHA_CONCILIA%(CUEI%)
        CANFIL& = 1000
        CANCOL& = COL_DEBE% + 4
        GRABO_CONCI% = 0
        HAY_REGISTROS_CON_FECHA_ANTERIOR% = 0
        TOT_REGISTROS% = 0
        '
50      RNDX$ = RANDSTRING$(-10)
        If CantRegistros("CONCILIA", "CODICONCI = '" & RNDX$ & "'", "NOMESS") > 0 Then GoTo 50

        For i& = FILA_INICIO% To CANFIL&
          'ASIGNACION DE CAMPOS DE PLANILLA A VARIABLES
          DEBEX# = 0: HABEX# = 0
          FECHAX = HOJA.Cells(i&, COL_FECHA%)
          CONCEPTO$ = HOJA.Cells(i&, COL_CONCEPTO%)
          IDENTIFI$ = HOJA.Cells(i&, COL_IDENT%)
          DEBEX# = XVALO(HOJA.Cells(i&, COL_DEBE%))
          'VALIDACION POR QUE HABEX# PUEDE SER 0
          If COL_HABER% > 0 Then HABEX# = XVALO(HOJA.Cells(i&, COL_HABER%))
          '
          'SI TIENE UNA FECHA VALIDA GRABA EN CONCILIA
          If CVINT(FECHAX) > 0 Then
                TOT_REGISTROS% = TOT_REGISTROS% + 1
                If CVINT(FECHAX) <= FECHA_MAX% Then
                   HAY_REGISTROS_CON_FECHA_ANTERIOR% = HAY_REGISTROS_CON_FECHA_ANTERIOR% + 1
                   GoTo 90
                End If
                '
                With CONCTMP
                .AddNew   ' YA CONVERTIDO
                J& = J& + 1
                On Error Resume Next
                !CodiEmpr = CodiEmp%
                On Error GoTo 0
                    !CODICONCI = RNDX$
                    !FECHA = FECHAX
                    !CONCEPTO = Left$(CONCEPTO$, 128)
                    'CASO QUE SEA SOLO TENGA IMPORTE NEGATIVO O POSITIVO EN UNA SOLA COLUMNA
                    If DEBE_Y_HABER% = 1 Then
                      If DEBEX# < 0.005 Then
                         DEBEX# = Abs(DEBEX#): HABEX# = 0
                      Else
                         HABEX# = DEBEX#: DEBEX# = 0
                      End If
                    End If
                    '
                    'GRABA
                    !CUENTACONTBANCO = CUEINT%(CUENTA_SELECCIONADA$)
                    !CUENTACONTBANCO_EXT = CUENTA_SELECCIONADA$
                    !DEBE = DEBEX#
                    !HABER = HABEX#
                    !Status = 0
                    !REFECONCI = DESCRI$
                    !ID_REFERENCIA = IDENTIFI$
                    !ORDCONCILIA = J&
                    !CONCEPTO_SISTEMA = ""
                    
                    GRABO_CONCI% = 1
                    
                    '
                .Update
               End With
          End If
90     Next i&
       '
       CONCTMP.Close
       Set CONCTMP = Nothing
       CODICONCI$ = RNDX$
       '
99     Libro.Close (False)
       objexcel.Quit ' Cierra Excel
       Set objexcel = Nothing
       Set Libro = Nothing
       Set HOJA = Nothing
       '
       Screen.MousePointer = 1
       'AVISO QUE SE ENCONTRARON REGISTROS CON FECHA ANTERIOR AL ULTIMO REGISTRADO
       If HAY_REGISTROS_CON_FECHA_ANTERIOR% > 0 Then
          Call COMUNI("Se Han Detectado " & HAY_REGISTROS_CON_FECHA_ANTERIOR% & " de " & TOT_REGISTROS% & " Movimientos Bancarios Con Fecha\Anterior a la Ultima Importación\Los Mismos No Fueron Incluidos.")
          'SI TAMPOCO GRABO NINGUN REGISTRO SALE PARA NO PRESENTAR EL MENSAJE SIGUIENTE
          If GRABO_CONCI% < 1 Then
          CODICONCI$ = ""
            Exit Sub
          End If
       End If
       '
       'AVISO QUE NO SE REALIZO NINGUNA GRABACION
       If GRABO_CONCI% < 1 Then
          Call COMUNI("No Se Ha Podido Realizar La Importación\Verifique Que La Planilla Excel Sea La Correcta\Y Que Tenga Correctamente Configurada Las Posiciones\Para el Banco Seleccionado")
          CODICONCI$ = ""
          Exit Sub
       End If
       '
       'SI LLEGA ACA MUESTRA MENSAJE DE AVISO DE FIN DE IMPORTACION
       Call COMUNI("Importación Finalizada - Comienzo de Conciliación Automática")
       
End Sub

Function MAXFECHA_CONCILIA%(CUEI%)
    MAXFECHA_CONCILIA% = CVINT(VALOBADA("CONCILIA", "FECHA", "CODICONCI <> '' AND CUENTACONBANCO = " & CUEI% & " ORDER BY FECHA DESC", "NOMESS"))
End Function
Function POSICIONES_CFGCONCI%(CUENTA_SELECCIONADA$, FILA_INICIO%, COL_INICIO%, COL_FECHA%, COL_CONCEPTO%, COL_IDENT%, COL_DEBE%, COL_HABER%)
     '
     'POSICIONES_CFGCONCI% = 0 FALTAN DATOS
     'POSICIONES_CFGCONCI% = 1 TODOS LOS DATOS Y SOLO DEBE
     'POSICIONES_CFGCONCI% = 2 TODOS LOS DATOS CARGADOS
     '
     'LEO EL ARCHIVO QUE MACHEA CON LA CUENTA CONTABLE
     Call TRALOCAL("CFGCONCI", "")
     XX$ = FILTERGETRECORD$("!CFGCONCI", 1, AJUSTI$(CUENTA_SELECCIONADA$, 12))
     '
     'ASIGNO LOS VALORES A LOS PARAMETROS DE DEVOLUCION
     CUENTACON$ = Mid$(XX$, 1, 12)
     If CUEINT%(CUENTACON$) > 0 Then
       FILA_INICIO% = CVI(Mid$(XX$, 13, 2))
       COL_INICIO% = CVI(Mid$(XX$, 15, 2))
       COL_FECHA% = CVI(Mid$(XX$, 17, 2))
       COL_CONCEPTO% = CVI(Mid$(XX$, 19, 2))
       COL_IDENT% = CVI(Mid$(XX$, 21, 2))
       COL_DEBE% = CVI(Mid$(XX$, 23, 2))
       COL_HABER% = CVI(Mid$(XX$, 25, 2))
    Else
       'VALIDO QUE LA CUENTA CONTABLE SEA VALIDA
       Call COMUNI("Cuenta Contable de Banco No Válida o Inexistente")
       Exit Function
    End If
    '
    'VALIDO QUE TODOS TENGAN VALOR CON EXCEPCION DEL HABER
    TX$ = ""
    If FILA_INICIO% <= 0 Then TX$ = "FILA DE INICIO": GoTo 50
    If COL_INICIO% <= 0 Then TX$ = "COLUMNA DE INICIO": GoTo 50
    If COL_FECHA% <= 0 Then TX$ = "COLUMNA DE FECHA": GoTo 50
    If COL_CONCEPTO% <= 0 Then TX$ = "COLUMNA DE CONCEPTO": GoTo 50
    If COL_IDENT% <= 0 Then TX$ = "COLUMNA IDENTIFICADOR": GoTo 50
    If COL_DEBE% <= 0 Then TX$ = "COLUMNA DEBE": GoTo 50
    '
    GoTo 60 'SI LLEGO ACA ESTA TODO BIEN Y VALIDO LA DEVOLUCION DE LA FUNCION
    
50  Call COMUNI("Falta Número de: " & TX$)
    Exit Function
    '
    'VALIDO LA DEVOLUCION DE LA FUNCION
60  If COL_HABER% < 1 Then
      POSICIONES_CFGCONCI% = 1
    Else
      POSICIONES_CFGCONCI% = 2
    End If
    '
End Function

Function VALIDA_CONCILIACION$(CUENTA_SELECCIONADA$, COL_FECHA%, FILA_INICIO%, COL_INICIO%, COL_CONCEPTO%, COL_IDENT%, COL_DEBE%, COL_HABER%)
        
        VALIDA_CONCILIACION$ = ""
        CANFIL& = 1000
        CANCOL& = COL_DEBE% + 4
        'RECORRO LA HOJA EXCEL Y VERIFICO QUE NO HAYA MOVMIMENTOS YA GUARDADEOS
        For i& = FILA_INICIO% To CANFIL&
          FECHA = "": CONCEPTO$ = "": IDINTIFI$ = "": DEBEX# = 0: HABER# = 0
          
          For J% = COL_INICIO% To CANCOL&
            If COL_FECHA% = J% Then FECHAX = HOJA.Cells(i&, J%)
            If COL_CONCEPTO% = J% Then CONCEPTO$ = HOJA.Cells(i&, J%)
            If COL_IDENT% = J% Then IDENTIFI$ = HOJA.Cells(i&, J%)
            If COL_DEBE% = J% Then DEBEX = XVALO(HOJA.Cells(i&, J%))
            If COL_HABER% = J% Then HABEX# = XVALO(HOJA.Cells(i&, J%))
          Next J%
          If TRIM$(FECHAX) <> "" Then
             Condi$ = " FECHA = '" & FETEXCOR1$(CVINT(FECHAX)) & "'"
             Condi$ = Condi$ + " AND CONCEPTO = '" & CONCEPTO$ & "'"
'             CONDI$ = CONDI$ + " CONCEPTO$ = '" & IDENTIFI$ & "'"
             Condi$ = Condi$ + " AND CUENTACONTBANCO = " & CUEINT%(CUENTA_SELECCIONADA$)
             Condi$ = Condi$ + " AND STATUS < 1"
             
             If COL_HABER% < 1 Then
                HABEX# = DEBEX#
                If DEBEX# < 0 Then DEBEX# = Abs(DEBEX#): HABEX# = 0
                Condi$ = Condi$ + " AND DEBE = " & DEBEX#
             Else
                Condi$ = Condi$ + " AND DEBE = " & XVALO(DEBEX#)
                Condi$ = Condi$ + " AND HABER = " & XVALO(HABEX#)
             End If
             YA_INGRESADO% = CantRegistros&("CONCILIA", Condi$, "NOMESS")
             If YA_INGRESADO% > 0 Then
                OPX% = COMALTER%("Se han Detectado Datos De Esta Conciliacion Ya Ingresados\\Cancelar\Sobre-escribir")
                If OPX% < 1 Or OPX% > 2 Then Exit Function
                If OPX% = 2 Then
                   VALIDA_CONCILIACION$ = VALOBADA("CONCILIA", "CODICONCI", Condi$, "NOMESS")
                End If
             End If
          End If

       Next i&

End Function


Sub VIEWTRANFON()

   Condi$ = "TIPOMOVIM = 'TRAFON'  ORDER BY CODICOM_COR DESC "
   Call CARGALISEL("ANTRAFON", "CAJABANC", "CODICOM_COR/A12;FECHCONTA/D8", Condi$, "DISTINCT")

10 Call SELECHO("ANTRAFON/Visualizar Transferencias")
   VDEVU$ = VALACT1$("ANTRAFON")
   If VDEVU$ = "" Then Exit Sub
   '
   NPX& = XVALO(Mid$(VDEVU$, 4))
   VDEVU$ = Mid$(VDEVU$, 4)
   TIDOCUM$ = "Transferencia de Fondo"
   DOCUNU& = 0
   Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, Str$(NPX&), VDEVU$)
   GoTo 10

End Sub

Sub VISUALI_EXTRACT_IMPOR()
10  Call SELECHO("CUEBANC/BANCOS")
    CUENTA_SELECCIONADA$ = TRIM$(VALACT1$("CUEBANC"))
        '
    If CUENTA_SELECCIONADA$ = "" Then Exit Sub
    CUEI% = CUEINT%(CUENTA_SELECCIONADA$)
    
   'PRESENTO LA LISTA DE SELECCION DE TODAS LAS CONCILIACIONES
   Call CARGALISEL("CONCITODO", "CONCILIA", "CODICONCI/A12;REFECONCI/A32", "CODICONCI <> '' AND CUENTACONTBANCO = " & CUEI% & " AND STATUS < 1 ", "NOMESS/DISTINCT")
   Call SELECHO("CONCITODO/SELECCION DE EXTRACTOS IMPORTADOS")
   CODICONCI$ = TRIM$(VALACT1$("CONCITODO"))
   If CODICONCI$ = "" Then Exit Sub
   
   On Error Resume Next
   FechaDesdeMovNoConciliados% = 0
   A$ = INIGES07.Caption        ' A PARTIR DEL ERROR DETECTA SI ESTA TRABAJANDO DESDE EL FLGEST07 O EL CONTAB07
   If Err.Number = 0 Then
        On Error GoTo 0
        VDEF$ = MKI$(COMEJ%)
        OBX$ = OBJPLA$("ENTRAFECHA1", VDEF$)
        If OBX$ = "" Then
           FechaDesdeMovNoConciliados% = 0
        Else
           FechaDesdeMovNoConciliados% = CVI(Mid$(OBX$, 1, 2))
        End If
   End If
   On Error GoTo 0
   Call PRESENTA_CONCILIA(CODICONCI$, 1, FechaDesdeMovNoConciliados%)
End Sub

Private Sub AGenRep_Click()
    '
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub

Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Hide
End Sub

Sub Command10_Click()
   Command10.Enabled = False
10 OPX% = ALTERNA%("Transferencia de Fondos\Anulación de Transferencia de Fondos\Visualizar T.Fondos")
   If OPX% = 1 Then
      If DERACCE%("TRAFONDO", "Transferencia de Fondos") >= 2 Then
         Call TRANFONDOS
         GoTo 10
      End If
   ElseIf OPX% = 2 Then
      If DERACCE%("ANTRAFO", "Anulación de Transferencia de Fondos") >= 2 Then
         Call ANUTRANFONDOS
         GoTo 10
      End If
   ElseIf OPX% = 3 Then
         Call VIEWTRANFON
         GoTo 10
   End If
   Command10.Enabled = True
End Sub

Sub Command11_Click()
   Command1.Enabled = False
   '\\\OT-05-0495-WAR-11/07/05///
   OE% = COMALTER("Opciones de Trabajo:\\Efectuar Déposito\Anular Déposito ")
   If OE% = 1 Then
      If DERACCE%("DEPOBAN", "Deposito Bancario") >= 2 Then
         Call ASIDECHE 'NUEVO DEPOSITO
      End If
     ElseIf OE% = 2 Then
      If DERACCE%("ANUDEBA", "Anulación de Deposito Bancario") >= 2 Then
         Call ANUDECHE 'ANULACION
      End If
   End If
   '\\\OT-05-0495-WAR-FIN///
99 Command1.Enabled = True
End Sub

Sub Command12ORIGINAL_Click()
   '
   'GUARDO UNA COPIA POR LAS DUDAS POR QUE SE HICIERON CAMBIOS VER COMMAND12_CLICK
   '******************************************************************************
   
   Command12.Enabled = False
   '
5  OPX% = COMALTER%("Modo de Conciliación\\Manual\Automática")
   If OPX% < 1 Or OPX% > 2 Then GoTo 99
   If OPX% = 2 Then
      Call CONCILIACION_AUTOMATICA
      GoTo 5
   End If
   
   Call RecFechAcredCheques  'PROCESO DE ASIGNAR FECHA DE ACREDITACION
                             'A LOS CHEQUES QUE NO TIENEN FECHA INGRESADA
                             'ASIGNANDOLE LA MISMA FECHA DE EMISION
                              '
   Call COPYSTRU("LIMAYOAC", "LIMAYOA1")
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
   FEX = DEXTRACT%: CONCIBAN07.Label8 = FECHATEX$(FEX)
   FEX = HAXTRACT%: CONCIBAN07.Label6 = FECHATEX$(FEX)
   CONCIBAN07.Label3 = FORMATNUM$(SALICON#, "F12.2")
   '
   'Call BLOQARCH("LIMAYOAC1")
   Screen.MousePointer = 11
   '
   FEINI = CDbl(CVDAT(DESPEN%))
   FEFIN = CDbl(CVDAT(HASPEN%))
   FEINISQL = FETEXCOR1$(DESPEN%)
   FEFINSQL = FETEXCOR1$(HASPEN%)

   '
   Call ABREPLANCUEN
   Screen.MousePointer = 11
   SQLX$ = "SELECT * FROM DetaAsien "
   SQLX$ = SQLX$ + "WHERE (CueCoInt=" & CUEII%
   ' SQLX$ = SQLX$ + " AND SucuAsien = " & SUCUI%
   SQLX$ = SQLX$ + " AND Fechasien >= '" & FEINISQL & "'"
   SQLX$ = SQLX$ + " AND Fechasien <= '" & FEFINSQL & "'"
   SQLX$ = SQLX$ + " AND statpase > 0 ) "
   'SQLX$ = SQLX$ + " AND (MarConsi = '' or ISNULL(MarConsi) = True or IsEmpty(MarConsi) = True) "
   SQLX$ = SQLX$ + " AND (MarConsi = '' or MarConsi IS NULL ) "
   SQLX$ = SQLX$ + "ORDER BY FECHASIEN, ORFEASIEN ASC;"
   'Set PASEASIEN = Contable.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
   Dim ENCASI As ADODB.Recordset
   Dim PASEASIEN As ADODB.Recordset
   Set PASEASIEN = New ADODB.Recordset
24  On Error Resume Next
    PASEASIEN.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 24
   End If
   On Error GoTo 0

   JJ& = 0
   TRANSPOR# = 0
20 With PASEASIEN
     On Error Resume Next
     .MoveFirst
     If Err Then GoTo 30
     On Error GoTo 0
     Do While Not .EOF
       NUAS$ = !nuidasien
       TIASI% = 0
       'Encasien.FindFirst "NUIDASIEN = '" & NUAS$ & "'"
       SQLX$ = "SELECT * FROM ENCABEASIEN "
       SQLX$ = SQLX$ + " WHERE NUIDASIEN = '" & NUAS$ & "'"
       Dim Encasien As ADODB.Recordset
       Set Encasien = FLEXCONN.Execute(FILTSQL$(SQLX$))

       If Encasien.EOF = False Then
         TIASI% = Encasien!TIPOASIEN
       End If
       '
       If TIASI% = 0 Then  ' DESCARTA ASIENTOS RESUMEN
         Z$ = REGBLAN$("!LIMAYOA1")
         Call REPLA(Z$, MKI$(CVINT(!FECHASIEN)), 1, 2)
         Call REPLA(Z$, !nuidasien, 3, 12)
         Call REPLA(Z$, MKI$(CVINT(!FECHASIEN)), 15, 2)
           DEPASE$ = !RefePase
         Call REPLA(Z$, DEPASE$, 17, 32)
           IDEBE# = !idebepase
         Call REPLA(Z$, MKD$(IDEBE#), 49, 8)
           IHABE# = !ihabepase
         Call REPLA(Z$, MKD$(IHABE#), 57, 8)
         Call REPLA(Z$, "DIAR", 73, 8)
         If IDEBE# >= 0.005 Or IHABE# >= 0.005 Then
           JJ& = JJ& + 1
           Call GRAREG("!LIMAYOA1", Z$, JJ&)
         End If
       End If
     .MoveNext
     Loop
   End With
   Set PASEASIEN = Nothing
30 On Error GoTo 0
   Call APERBASDAT("CABAN")
   Screen.MousePointer = 11
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "where (CueContaI = " & CUEII% & " "
   'SQLX$ = SQLX$ + " AND (MarConsi = '' or isnull(MarConsi) = True OR IsEmpty(MarConsi) = True) "
   SQLX$ = SQLX$ + " AND (MarConsi = '' OR MarConsi IS NULL ) "
   SQLX$ = SQLX$ + " AND FechConta >= '" & FEINISQL & "'"
   SQLX$ = SQLX$ + " AND FechConta <= '" & FEFINSQL & "') "
   SQLX$ = SQLX$ + " ORDER BY FECHCONTA ASC, FEHOREAL ASC;"
   '
   'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim CABATEMP As ADODB.Recordset
   Set CABATEMP = New ADODB.Recordset
26 On Error Resume Next
   CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 26
   End If
   On Error GoTo 0
   '
   With CABATEMP
     On Error Resume Next
     .MoveFirst
     If Err Then
        GoTo 40
     End If
     On Error GoTo 0
     Do While Not .EOF
         IDEBE# = !idebecomp
         IHABE# = !ihabecomp
       If Int(CDbl(!fechconta)) < FEINI Then
           TRANSPOR# = TRANSPOR# + IDEBE# - IHABE#
         ElseIf Int(CDbl(!fechconta)) <= FEFIN Then
           Z$ = REGBLAN$("LIMAYOA1")
           Call REPLA(Z$, MKI$(CVINT(!fechconta)), 1, 2)
           Call REPLA(Z$, !CodiCom_Cor, 3, 12)
           Call REPLA(Z$, MKI$(CVINT(!fechacred)), 15, 2)
             DEPASE$ = Left$(!DESCRIMOV, 32)
           Call REPLA(Z$, DEPASE$, 17, 32)
           Call REPLA(Z$, MKD$(IDEBE#), 49, 8)
           Call REPLA(Z$, MKD$(IHABE#), 57, 8)
           Call REPLA(Z$, "CABAN", 73, 8)
           If IDEBE# >= 0.005 Or IHABE# >= 0.005 Then
             JJ& = JJ& + 1
             Call GRAREG("!LIMAYOA1", Z$, JJ&)
           End If
       End If
     .MoveNext
     Loop
   End With
   '
40 Call SETULTREG("!LIMAYOA1", JJ&)
   'Call FILESORT("LIMAYOAC", "", 1, 3, "ASC")
   Call TRACENTRAL("LIMAYOA1", "")
   Call FILESORT("LIMAYOA1", "", 1, 3, "ASC")
   Call TRALOCAL("LIMAYOA1", "")
   '
   Screen.MousePointer = 11
   CONCIBAN07.List2.Clear
   Z$ = Space$(128)
   For U& = 1 To ULTREG("!LIMAYOA1")
     X$ = REGLEIDO$("!LIMAYOA1", U&)
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
     CONCIBAN07.List2.AddItem Z$
     DoEvents
   Next U&
   Call CIERRARCH("!LIMAYOA1")
   '
   '
   ' CARGA LISTA DE MOVIMIENTOS CONCILIADOS
   FEINI = CDbl(CVDAT(DEXTRACT%))
   FEFIN = CDbl(CVDAT(HAXTRACT%))
   FEINISQL1 = FETEXCOR1$(DEXTRACT%)
   FEFINSQL1 = FETEXCOR1$(HAXTRACT%)

   '
   Call ABREPLANCUEN
   Screen.MousePointer = 11
   SQLX$ = "SELECT * FROM DetaAsien "
   SQLX$ = SQLX$ + " WHERE (CueCoInt=" & CUEII%
   ' SQLX$ = SQLX$ + " AND SucuAsien = " & SUCUI%
   SQLX$ = SQLX$ + " AND Fechasien >= '" & FEINISQL1 & "'"
   SQLX$ = SQLX$ + " AND Fechasien <= '" & FEFINSQL1 & "'"
   SQLX$ = SQLX$ + " AND statpase >0 ) "
   SQLX$ = SQLX$ + " AND MarConsi = '*' "
   SQLX$ = SQLX$ + "ORDER BY FECHASIEN, ORFEASIEN ASC;"
   'Set PASEASIEN = Contable.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
   'Dim PASEASIEN As ADODB.Recordset
   Set PASEASIEN = New ADODB.Recordset
27  On Error Resume Next
   PASEASIEN.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 27
   End If
   On Error GoTo 0


   JJ& = 0
   TRANSPOR# = 0
70 With PASEASIEN
     On Error Resume Next
     .MoveFirst
     If Err Then GoTo 80
     On Error GoTo 0
     Do While Not .EOF
       NUAS$ = !nuidasien
       TIASI% = 0
       '
       SQLY$ = "SELECT TIPOASIEN FROM ENCABEASIEN "
       SQLY$ = SQLY$ + " WHERE NUIDASIEN = '" & NUAS$ + "'"
       Set ENCASI = New ADODB.Recordset
       Set ENCASI = FLEXCONN.Execute(FILTSQL$(SQLY$))
       '
       If Not (ENCASI.BOF And ENCASI.EOF) Then
         TIASI% = ENCASI!TIPOASIEN
       End If
       If TIASI% = 0 Then  ' DESCARTA ASIENTOS RESUMEN
         Z$ = REGBLAN$("LIMAYOAC")
         Call REPLA(Z$, MKI$(CVINT(!FECHASIEN)), 1, 2)
         Call REPLA(Z$, !nuidasien, 3, 12)
         Call REPLA(Z$, MKI$(CVINT(!FECHASIEN)), 15, 2)
           DEPASE$ = !RefePase
         Call REPLA(Z$, DEPASE$, 17, 32)
           IDEBE# = !idebepase
         Call REPLA(Z$, MKD$(IDEBE#), 49, 8)
           IHABE# = !ihabepase
         Call REPLA(Z$, MKD$(IHABE#), 57, 8)
         Call REPLA(Z$, "DIAR", 73, 6)
         Call REPLA(Z$, "*", 80, 1)
         If IDEBE# >= 0.005 Or IHABE# >= 0.005 Then
           JJ& = JJ& + 1
           Call GRAREG("!LIMAYOA1", Z$, JJ&)
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
   SQLX$ = SQLX$ + " AND FechAcreD >= '" & FEINISQL1 & "'"
   SQLX$ = SQLX$ + " AND FechAcreD <= '" & FEFINSQL1 & "') "
   SQLX$ = SQLX$ + " ORDER BY FECHACRED ASC, FEHOREAL ASC;"
   'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   'Dim CABATEMP As ADODB.Recordset
   Set CABATEMP = New ADODB.Recordset
28 On Error Resume Next
   CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 28
   End If
   On Error GoTo 0
   
   With CABATEMP
     On Error Resume Next
     .MoveFirst
     If Err Then
        GoTo 90
     End If
     On Error GoTo 0
     Do While Not .EOF
         IDEBE# = !idebecomp
         IHABE# = !ihabecomp
       If Int(CDbl(!fechacred)) < FEINI Then
           TRANSPOR# = TRANSPOR# + IDEBE# - IHABE#
         ElseIf Int(CDbl(!fechacred)) <= FEFIN Then
           Z$ = REGBLAN$("LIMAYOAC")
           Call REPLA(Z$, MKI$(CVINT(!fechconta)), 1, 2)
           Call REPLA(Z$, !CodiCom_Cor, 3, 12)
           Call REPLA(Z$, MKI$(CVINT(!fechacred)), 15, 2)
             DEPASE$ = Left$(!DESCRIMOV, 32)
           Call REPLA(Z$, DEPASE$, 17, 32)
           Call REPLA(Z$, MKD$(IDEBE#), 49, 8)
           Call REPLA(Z$, MKD$(IHABE#), 57, 8)
           Call REPLA(Z$, "CABAN", 73, 6)
           Call REPLA(Z$, "*", 80, 1)
           If IDEBE# >= 0.005 Or IHABE# >= 0.005 Then
             JJ& = JJ& + 1
             Call GRAREG("!LIMAYOA1", Z$, JJ&)
           End If
       End If
     .MoveNext
     Loop
   End With
   Set CABATEMP = Nothing
90 Call SETULTREG("!LIMAYOA1", JJ&)
   'Call FILESORT("LIMAYOAC", "", 3, 3, "ASC")
   '\\\OT-08-0609-OD-29/07/08///
   Call TRACENTRAL("LIMAYOA1", "")
   Call FILESORT("LIMAYOA1", "", 3, 3, "ASC")
   Call TRALOCAL("LIMAYOA1", "")
   '\\\OT-08-0609-OD-FIN///

   '
   CONCIBAN07.List1.Clear
   Z$ = Space$(128)
   For U& = 1 To ULTREG("!LIMAYOA1")
     X$ = REGLEIDO$("!LIMAYOA1", U&)
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
     CONCIBAN07.List1.AddItem Z$
     DoEvents
   Next U&
   Call CIERRARCH("!LIMAYOA1")
   '
   '
   TRANSPOR# = Val(CONCIBAN07.Label3)
   For KU& = 1 To CONCIBAN07.List1.ListCount
     TTXX$ = CONCIBAN07.List1.List(KU& - 1)
     TRANSPOR# = TRANSPOR# + Val(Mid$(TTXX$, 66, 14)) - Val(Mid$(TTXX$, 80, 14))
     Call REPLA(TTXX$, FORMATNUM$(TRANSPOR#, "F14.2"), 94, 14)
     CONCIBAN07.List1.RemoveItem (KU& - 1)
     CONCIBAN07.List1.AddItem TTXX$, KU& - 1
   Next KU&
   CONCIBAN07.Label14 = FORMATNUM(TRANSPOR#, "F12.2")
   '
   TRANSPOR# = 0
   For KU& = 1 To CONCIBAN07.List2.ListCount
     TTXX$ = CONCIBAN07.List2.List(KU& - 1)
     TRANSPOR# = TRANSPOR# + Val(Mid$(TTXX$, 66, 14)) - Val(Mid$(TTXX$, 80, 14))
   Next KU&
   CONCIBAN07.Label15 = FORMATNUM(TRANSPOR#, "F12.2")
   '
   CONCIBAN07.Ltext2 = CODCUE$(CUEII%)
   CONCIBAN07.Label4 = DENOCUE$(CUEII%)
   Screen.MousePointer = 1
   CONCIBAN07.Show 1
   '
99 Command12.Enabled = True
End Sub

Sub Command12_Click()
   '
   'SE MODIFICO   CORTANDO CODIGO Y LLEVANDOLO A  RUTINA : GENERA_PENDIENTE_EN_LIMAYO01
   'PARA PODER LLAMARLO DESDE OTRO LUGAR SIN CAMBIAR EL CONCEPTO
   'CODIGO ORIGINAL EN  COMMAND12ORIGINAL
   
   Command12.Enabled = False
   '
   If Control$("", "CONCIAU") <> "SI" Then GoTo 9
   '
5  OPX% = COMALTER%("Modo de Conciliación\\Manual\Automática")

   If OPX% < 1 Or OPX% > 3 Then GoTo 99
   If OPX% = 2 Then
8     OPX2% = COMALTER%("Modo de Conciliación\\Nueva Importación\Visualizar Extractos Importados\Anular Importación")
      If OPX2% = 1 Then
        Call CONCILIACION_AUTOMATICA
        GoTo 8
      ElseIf OPX2% = 2 Then
        Call VISUALI_EXTRACT_IMPOR
        GoTo 8
      ElseIf OPX2% = 3 Then
        Call ANULA_IMPORTACION
        GoTo 8
      End If
      GoTo 5 ' si no elige se va atras
   End If
   
9  Call RecFechAcredCheques  'PROCESO DE ASIGNAR FECHA DE ACREDITACION
                             'A LOS CHEQUES QUE NO TIENEN FECHA INGRESADA
                             'ASIGNANDOLE LA MISMA FECHA DE EMISION
                              '
   Call COPYSTRU("LIMAYOAC", "LIMAYOA1")
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
   FEX = DEXTRACT%: CONCIBAN07.Label8 = FECHATEX$(FEX)
   FEX = HAXTRACT%: CONCIBAN07.Label6 = FECHATEX$(FEX)
   CONCIBAN07.Label3 = FORMATNUM$(SALICON#, "F12.2")
   '
   'Call BLOQARCH("LIMAYOAC1")
   Screen.MousePointer = 11
   '
   FEINI = CDbl(CVDAT(DESPEN%))
   FEFIN = CDbl(CVDAT(HASPEN%))
   FEINISQL = FETEXCOR1$(DESPEN%)
   FEFINSQL = FETEXCOR1$(HASPEN%)

   '
   Call ABREPLANCUEN
   'SE GENERO ESTA RUTINA PARA PODER CARGAR LOS PENDIENTES Y LLAMARLA DESDE OTRAS APLICACIONES
   'EN LA MISMA SE CARGA LA INFO DE LOS PENDIENTES AL ARCHIVO LIMAYO01
   Call GENERA_PENDIENTE_EN_LIMAYO01(CUEII%, DESPEN%, HASPEN%)

   '
   Screen.MousePointer = 11
   CONCIBAN07.List2.Clear
   Z$ = Space$(128)
   For U& = 1 To ULTREG("!LIMAYOA1")
     X$ = REGLEIDO$("!LIMAYOA1", U&)
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
     CONCIBAN07.List2.AddItem Z$
     DoEvents
   Next U&
   Call CIERRARCH("!LIMAYOA1")
   '
   '
   ' CARGA LISTA DE MOVIMIENTOS CONCILIADOS
   FEINI = CDbl(CVDAT(DEXTRACT%))
   FEFIN = CDbl(CVDAT(HAXTRACT%))
   FEINISQL1 = FETEXCOR1$(DEXTRACT%)
   FEFINSQL1 = FETEXCOR1$(HAXTRACT%)

   '
   Call ABREPLANCUEN
   Screen.MousePointer = 11
   SQLX$ = "SELECT * FROM DetaAsien "
   SQLX$ = SQLX$ + " WHERE (CueCoInt=" & CUEII%
   ' SQLX$ = SQLX$ + " AND SucuAsien = " & SUCUI%
   SQLX$ = SQLX$ + " AND Fechasien >= '" & FEINISQL1 & "'"
   SQLX$ = SQLX$ + " AND Fechasien <= '" & FEFINSQL1 & "'"
   SQLX$ = SQLX$ + " AND statpase >0 ) "
   SQLX$ = SQLX$ + " AND MarConsi = '*' "
   SQLX$ = SQLX$ + "ORDER BY FECHASIEN, ORFEASIEN ASC;"
   'Set PASEASIEN = Contable.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
   'Dim PASEASIEN As ADODB.Recordset
   Set PASEASIEN = New ADODB.Recordset
27  On Error Resume Next
   PASEASIEN.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 27
   End If
   On Error GoTo 0


   JJ& = 0
   TRANSPOR# = 0
70 With PASEASIEN
     On Error Resume Next
     .MoveFirst
     If Err Then GoTo 80
     On Error GoTo 0
     Do While Not .EOF
       NUAS$ = !nuidasien
       TIASI% = 0
       '
       SQLY$ = "SELECT TIPOASIEN FROM ENCABEASIEN "
       SQLY$ = SQLY$ + " WHERE NUIDASIEN = '" & NUAS$ + "'"
       Set ENCASI = New ADODB.Recordset
       Set ENCASI = FLEXCONN.Execute(FILTSQL$(SQLY$))
       '
       If Not (ENCASI.BOF And ENCASI.EOF) Then
         TIASI% = ENCASI!TIPOASIEN
       End If
       If TIASI% = 0 Then  ' DESCARTA ASIENTOS RESUMEN
         Z$ = REGBLAN$("LIMAYOAC")
         Call REPLA(Z$, MKI$(CVINT(!FECHASIEN)), 1, 2)
         Call REPLA(Z$, !nuidasien, 3, 12)
         Call REPLA(Z$, MKI$(CVINT(!FECHASIEN)), 15, 2)
           DEPASE$ = !RefePase
         Call REPLA(Z$, DEPASE$, 17, 32)
           IDEBE# = !idebepase
         Call REPLA(Z$, MKD$(IDEBE#), 49, 8)
           IHABE# = !ihabepase
         Call REPLA(Z$, MKD$(IHABE#), 57, 8)
         Call REPLA(Z$, "DIAR", 73, 6)
         Call REPLA(Z$, "*", 80, 1)
         If IDEBE# >= 0.005 Or IHABE# >= 0.005 Then
           JJ& = JJ& + 1
           Call GRAREG("!LIMAYOA1", Z$, JJ&)
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
   SQLX$ = SQLX$ + " AND FechAcreD >= '" & FEINISQL1 & "'"
   SQLX$ = SQLX$ + " AND FechAcreD <= '" & FEFINSQL1 & "') "
   SQLX$ = SQLX$ + " ORDER BY FECHACRED ASC, FEHOREAL ASC;"
   'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   'Dim CABATEMP As ADODB.Recordset
   Set CABATEMP = New ADODB.Recordset
28 On Error Resume Next
   CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 28
   End If
   On Error GoTo 0
   
   With CABATEMP
     On Error Resume Next
     .MoveFirst
     If Err Then
        GoTo 90
     End If
     On Error GoTo 0
     Do While Not .EOF
         IDEBE# = !idebecomp
         IHABE# = !ihabecomp
       If Int(CDbl(!fechacred)) < FEINI Then
           TRANSPOR# = TRANSPOR# + IDEBE# - IHABE#
         ElseIf Int(CDbl(!fechacred)) <= FEFIN Then
           Z$ = REGBLAN$("LIMAYOAC")
           Call REPLA(Z$, MKI$(CVINT(!fechconta)), 1, 2)
           Call REPLA(Z$, !CodiCom_Cor, 3, 12)
           Call REPLA(Z$, MKI$(CVINT(!fechacred)), 15, 2)
             DEPASE$ = Left$(!DESCRIMOV, 32)
           Call REPLA(Z$, DEPASE$, 17, 32)
           Call REPLA(Z$, MKD$(IDEBE#), 49, 8)
           Call REPLA(Z$, MKD$(IHABE#), 57, 8)
           Call REPLA(Z$, "CABAN", 73, 6)
           Call REPLA(Z$, "*", 80, 1)
           If IDEBE# >= 0.005 Or IHABE# >= 0.005 Then
             JJ& = JJ& + 1
             Call GRAREG("!LIMAYOA1", Z$, JJ&)
           End If
       End If
     .MoveNext
     Loop
   End With
   Set CABATEMP = Nothing
90 Call SETULTREG("!LIMAYOA1", JJ&)
   'Call FILESORT("LIMAYOAC", "", 3, 3, "ASC")
   '\\\OT-08-0609-OD-29/07/08///
   Call TRACENTRAL("LIMAYOA1", "")
   Call FILESORT("LIMAYOA1", "", 3, 3, "ASC")
   Call TRALOCAL("LIMAYOA1", "")
   '\\\OT-08-0609-OD-FIN///

   '
   CONCIBAN07.List1.Clear
   Z$ = Space$(128)
   For U& = 1 To ULTREG("!LIMAYOA1")
     X$ = REGLEIDO$("!LIMAYOA1", U&)
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
     CONCIBAN07.List1.AddItem Z$
     DoEvents
   Next U&
   Call CIERRARCH("!LIMAYOA1")
   '
   '
   TRANSPOR# = Val(CONCIBAN07.Label3)
   For KU& = 1 To CONCIBAN07.List1.ListCount
     TTXX$ = CONCIBAN07.List1.List(KU& - 1)
     TRANSPOR# = TRANSPOR# + Val(Mid$(TTXX$, 66, 14)) - Val(Mid$(TTXX$, 80, 14))
     Call REPLA(TTXX$, FORMATNUM$(TRANSPOR#, "F14.2"), 94, 14)
     CONCIBAN07.List1.RemoveItem (KU& - 1)
     CONCIBAN07.List1.AddItem TTXX$, KU& - 1
   Next KU&
   CONCIBAN07.Label14 = FORMATNUM(TRANSPOR#, "F12.2")
   '
   TRANSPOR# = 0
   For KU& = 1 To CONCIBAN07.List2.ListCount
     TTXX$ = CONCIBAN07.List2.List(KU& - 1)
     TRANSPOR# = TRANSPOR# + Val(Mid$(TTXX$, 66, 14)) - Val(Mid$(TTXX$, 80, 14))
   Next KU&
   CONCIBAN07.Label15 = FORMATNUM(TRANSPOR#, "F12.2")
   '
   CONCIBAN07.Ltext2 = CODCUE$(CUEII%)
   CONCIBAN07.Label4 = DENOCUE$(CUEII%)
   Screen.MousePointer = 1
   CONCIBAN07.Show 1
   '
99 Command12.Enabled = True
End Sub
'
Sub RecFechAcredCheques()
   ' PROCESO DE ASIGNAR FECHA DE ACREDITACION A LOS CHEQUES QUE NO
   ' TIENEN FECHA INGRESADA ASIGNANDOLE LA MISMA FECHA DE EMISION
   '
   Call ABRECABAN
   FE% = FECHANUM("01/01/1980")
   fech = FETEXCOR1$(FE%)
   SQLX$ = " SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + " WHERE (FECHACRED <= ' " & fech & "'"
   SQLX$ = SQLX$ + " OR FECHACRED IS NULL)"
   SQLX$ = SQLX$ + " AND OPCIMOVIM = 'CH'"
   SQLX$ = SQLX$ + " AND TIPOMOVIM = 'CHRECI'"
   
   Dim CABTMP As ADODB.Recordset
   Set CABTMP = New ADODB.Recordset
25 On Error Resume Next
   CABTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   With CABTMP
     Do While Not .EOF
        !fechacred = CVDAT(CVINT(!FECHEMISI))
        .Update
     .MoveNext
     Loop
   End With
   CABTMP.Close
   Set CABTMP = Nothing
   '
End Sub

Sub Command13_Click()
    Command13.Enabled = False
    If DERACCE%("RECECHE", "Recepción de Cheques") >= 2 Then
       Call RECICHEQ
    End If
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
   'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
   Dim CABATEMP As ADODB.Recordset
   Set CABATEMP = New ADODB.Recordset
24 On Error Resume Next
   CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 24
   End If
   On Error GoTo 0
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
               '.Edit
               !Status = 2
               !FECHENTRE = !FEDEPOSENT
               .Update
               GoTo 19    ' salio de cartera
             End If
           End If
           '
           '.Edit
           !Status = 1
           .Update
           FEACRED% = CVINT(!fechacred)
           BANUFE$ = !DESCRIMOV
           ICHEQ# = !ValAbsComp
           IMCHEX$ = TRIM$(FORMATNUM$(ICHEQ#, "F14.2"))
           '
           If TRIM$(Left$(BANUFE$, 12)) = "" Then
             Call REPLA(BANUFE$, "Cheque Nro:", 1, 11)
           End If
           '
           FEACRED% = CVINT%(!fechacred)
           FEX = FEACRED%
           MESA$ = Mid$(FECHATEX$(FEX), 4, 5)
           FEMISI% = CVINT%(!FECHEMISI)
           FERECEP% = CVINT(!fechconta)
           NURECIB$ = !CodiCom_Cor
           ORIGCHE$ = !refecom
           CUEII% = !CUECONTAI
           CUENTRA$ = TRIM$(CODCUE$(CUEII%)) + " - " + DENOCUE$(CUEII%)
           DEMOV$ = !DESCRIMOV
             If VUELTA% = 2 Then DEMOV$ = Left$(!DESCRIMOV, 32)
           TXD$ = REPLACAR$(DEMOV$, " ", "")
           '
           ' buscar por importe haber
           SQLY$ = "SELECT * FROM CAJABANC"
           SQLY$ = SQLY$ + " TipoMovim = 'CHENTR'"
           SQLY$ = SQLY$ + " AND (ABS(IHabeComp - " & IMCHEX$ & ") < 0.005) "
           '"TipoMovim = 'CHENTR' AND ABS(IHabeComp - " & IMCHEX$ & ") < 0.005 "
           Dim CAJJJABANC As ADODB.Recordset
           Set CAJJJABANC = New ADODB.Recordset
26         On Error Resume Next
           CAJJJABANC.Open FILTSQL$(SQLY$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
           If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
              On Error GoTo 0
              Call CapturaErrorOpen
              GoTo 26
           End If
           On Error GoTo 0

           With CAJJJABANC
           'CajaBanc.FindFirst SQLY$
'16         'If CajaBanc.NoMatch = False Then
             On Error Resume Next
             .MoveFirst
             If Err < 1 Then
              Do While Not .EOF
                TXH$ = REPLACAR$(!DESCRIMOV, " ", "")
                If VUELTA% = 2 Then TXH$ = REPLACAR$(Left$(!DESCRIMOV, 32), " ", "")
                If TXD$ = TXH$ Then
                  CUEII% = !CUECONTAI
                  CUENSAL$ = TRIM$(CODCUE$(CUEII%)) + " - " + DENOCUE$(CUEII%)
                  FESALI% = CVINT(!fechconta)
                  DOCUSA$ = TRIM$(!CodiCom_Cor)
                  DESTIN$ = TRIM$(!refecom)
                  TTXX$ = "Desea Conciliar la Salida del Cheque\"
                  TTXX$ = TTXX$ + "'" + TRIM$(BANUFE$) + "'\"
                  TTXX$ = TTXX$ + "por un Importe de " + TRIM$(IMCHEX$) + "\"
                  TTXX$ = TTXX$ + "Entregado a " + DESTIN$ + " ?"
                  If COMALTER%(TTXX$ + "\\Si, Conciliar\No, Mantener") = 1 Then
                    '.Edit
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
                'CajaBanc.FindNext SQLY$
                .MoveNext
              'GoTo 16
              Loop
             End If
19          .MoveNext
           End With
         Loop
         If VUELTA% = 1 Then
           VUELTA% = 2
           GoTo 15
         End If
     End If
   End With
   CABATEMP.Close
   Set CABATEMP = Nothing
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
    If InStr(1, EMPREAC$, "FERVI") > 0 Then
        Call APERBASDAT("CABAN")
        '
        Condi$ = "NUIDASIEN='130610-002'"
        SQLX$ = "SELECT * FROM ENCABEASIEN WITH(NOLOCK) WHERE " & Condi$
        '
        Set rst = READSET(SQLX$)
        With rst
            Do While Not .EOF
                XXY$ = Mid$(SAFETEXT$(!refeasien), 10)
                NN% = InStr(1, XXY$, "-"): BANCO$ = AJUSTI$(Mid$(XXY$, 1, NN% - 1), 18)
                XXY$ = TRIM$(Mid$(XXY$, NN% + 1))
                NN% = InStr(1, XXY$, "-"): NUMCHE$ = AJUSTI$(Mid$(XXY$, 1, NN% - 1), 14)
                XXY$ = TRIM$(Mid$(XXY$, NN% + 1))
                FECHA$ = XXY$
                CHEQUE$ = BANCO$ & NUMCHE$ & FECHA$
                Condi$ = "TIPOMOVIM='CHRECI' AND DescriMov LIKE '" & CHEQUE$ & "%'"
                If CantRegistros("CAJABANC", Condi$) > 0 Then
                    Call ACTUREGISTRO("CAJABANC", "Status", Condi$, 2, REGAF&)
                Else
                    Call COMUNI("No Se Encontro el Cheque '" & CHEQUE$ & "'")
                End If
                .MoveNext
            Loop
        End With
        Condi$ = "NUIDASIEN='130610-002'"
        Call BORRAREGISTROS("ENCABEASIEN", Condi$, REGAF&)
        Call BORRAREGISTROS("DETAASIEN", Condi$, REGAF&)
        
        Call COMUNI("Proceso Finalizado")
        
'        FF% = FECHANUM("31/12/01")
'        HASTAFE = FETEXCOR1$(FF%)
'        CONDI$ = "TipoMovim = 'CHRECI' "
'        CONDI$ = CONDI$ + "AND Status = 1 "
'        CONDI$ = CONDI$ + "AND FechConta <= '" & HASTAFE & "' "
'        '
'        Call ACTUREGISTRO("CAJABANC", "STATUS", CONDI$, 9, REGAF&)
'        '
'        Call ACTUREGISTRO("CAJABANC", "STATUS", "DOCUMENT='OP-075157-00'", 1, REGAF&)
'        Call ACTUREGISTRO("CAJABANC", "STATUS", "DOCUMENT='OP-075158-00'", 1, REGAF&)
'        '
'        Call COMUNI("Proceso Finalizado")
        '
    End If
    Exit Sub
'
' este es el modificado el 23 de agosto
   Dim IDENCHE0$(), IDENCHE1$(), IDENCHE2$(), IMPOCHE#(), FESALID%(), DOCUSAL$(), DESTINX$()
   Dim DEMECO$(), CUENCONCO%()
   '
   Command19.Enabled = False
   HOII% = HOY(HOS$)
   CUECA% = CUEICARCH%
   Screen.MousePointer = 11
   '
   CUECHECA$ = Control$("", "CUECHECA")
   CUEICHEQ% = CUEINT%(CUECHECA$)
   If TRIM$(CUECHECA$) = "" Or CUEICHEQ% < 1 Then
      Screen.MousePointer = 1
      Command19.Enabled = True
      Call MENSERR(24, "Falta Definir Cuenta Contable\Asociada a Cartera de Valores")
      OPCHEQ07.Show 1
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
   'Set cabachentr = CueCorri.OpenRecordset(SQLZ$, dbOpenDynaset)
   Dim CABACHENTR As ADODB.Recordset
   Set CABACHENTR = New ADODB.Recordset
27 On Error Resume Next
   CABACHENTR.Open FILTSQL$(SQLZ$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 27
   End If
   On Error GoTo 0
   '
   With CABACHENTR
     On Error Resume Next
     .MoveFirst
     If Err Then
       On Error GoTo 0
       GoTo 10
     End If
     On Error GoTo 0
     Do While Not .EOF
        If !CUECONTAI <> CUECA% Then GoTo 9 ' SOLO SALIDAS DE CARTERA DE CHEQUES
        '
        ' enjuaque para guidi
        If !DESCRIMOV = "PROVINC.          109307        04/07/05" Then
          'If CDbl(!FechEmisi) = CDbl(CVDAT(FECHANUM("09/09/05"))) Then
          If !FECHEMISI = CVDAT(FECHANUM("09/09/05")) Then
            '.Edit
            !DESCRIMOV = "PROVINC.          10947         01/09/05"
            .Update
          End If
        End If
        ' fin enjuague para guidi
        '
        CASALI& = CASALI& + 1
        ReDim Preserve IDENCHE0$(CASALI&), IDENCHE1$(CASALI&), IDENCHE2$(CASALI&), IMPOCHE#(CASALI&), FESALID%(CASALI&), DOCUSAL$(CASALI&), DESTINX$(CASALI&)
        IDENCHE0$(CASALI&) = !DESCRIMOV
        IDENCHE1$(CASALI&) = REPLACAR$(!DESCRIMOV, " ", "")
        IDENCHE2$(CASALI&) = REPLACAR$(Left$(!DESCRIMOV, 32), " ", "")
        IMPOCHE#(CASALI&) = !ihabecomp
        FESALID%(CASALI&) = CVINT(!fechconta)
        DOCUSAL$(CASALI&) = TRIM$(!CodiCom_Cor)
        DESTINX$(CASALI&) = TRIM$(!refecom)
9    .MoveNext
     Loop
   End With
   Set CABACHENTR = Nothing
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
   'SQLX$ = SQLX$ + "AND Status = 1 "     ' aun en cartera ** YA ESTABA COMENTADO
   'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
   Dim CABATEMP As ADODB.Recordset
   Set CABATEMP = New ADODB.Recordset
24 On Error Resume Next
   CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 24
   End If
   On Error GoTo 0
   With CABATEMP
     On Error Resume Next
     .MoveFirst
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
           DEMOV$ = !DESCRIMOV
           '
           ' ESTO ES UN ENJUAGUE PARA TANIT
           If DEMOV$ = "Prov.Bs.As.       49623014      12/04/06" Then
              DEMOV$ = "Prov.Bs.As.         49623013    07/04/06"
              '.Edit
              !DESCRIMOV = DEMOV$
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
             If Abs(IMPOCHE#(KU&) - ICHEQ#) < 0.005 Then
               If IDENCHE1$(KU&) = DEMOV1$ Or (VUELTA% = 2 And IDENCHE2(KU&) = DEMOV2$) Then
                 ' ENCONTRO LA SALIDA
                 IDENCHE1$(KU&) = ""
                 IDENCHE2$(KU&) = ""
                 IMPOCHE#(KU&) = 0
                 If !Status <> 2 Or (CVINT(!FECHENTRE) <= 33 And CVINT(!FEDEPOSENT) <= 33) Then
                   '.Edit
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
              '.Edit
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
              If !CUECONTAI <> CUEICHEQ% Then
                 '.Edit
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
   Set CABATEMP = Nothing
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
99 Command19.Enabled = True
   '
End Sub

Sub Command15_Click()
    Command15.Enabled = False
    Call SUDIARCAJ   ' Impresion Subdiario de Caja y Bancos
    Command15.Enabled = True
End Sub

Sub Command16_Click()
    Command16.Enabled = False
    OE% = COMALTER("Opciones de Trabajo:\\Entregar Valores\Anular Entrega de Valores\Imprimir Comprobante")
    If OE% = 1 Then
        If DERACCE%("ENTRECHE", "Entrega de Valores") >= 2 Then
           Call ENTRECHEQ 'NUEVO DEPOSITO
        End If
      ElseIf OE% = 2 Then
        If DERACCE%("ANUENCHE", "Anulación de Entrega de Valores") >= 2 Then
          Call ANUENCHE 'ANULACION
        End If
      ElseIf OE% = 3 Then
        Call ImprimeEntregaDeValores
    End If
    Command16.Enabled = True
End Sub

Sub ImprimeEntregaDeValores()
      
    ARCSQR$ = TRIM$(Control("", "IMPENTVA"))
    If ARCSQR$ = "" Then
        Call COMUNI("Imposible Imprimir Entrega de Valores.\ \Falta Configurar Formulario de Impresión.")
        GoTo 99
    End If
    
    TituloLista$ = "Listado de Entrega de Valores"
    
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Nro Entrega/A18;Fecha/D8;Cuenta/I0", FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Nro Entrega/A18;Fecha/D8;Cuenta/I0", FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$
    
    Dim obj As ADODB.Recordset
    consultasql$ = "SELECT CODICOM_LAR, FECHCONTA, CUECONTAI FROM CAJABANC WITH(NOLOCK) WHERE TIPOMOVIM='ENTRECHE' ORDER BY FECHCONTA DESC"
    Set RS1 = READSET(consultasql$)
    Call Carga_MSF_Recordset(RS1, "Nro Entrega/A18;Fecha/D8;Cuenta/I0", FRMZELECHO.msf2, 1)
    On Error Resume Next
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 5000
50  FRMZELECHO.Show 1
    
    NumeroEntrega$ = TRIM(RESP_ZELE_VECT(1))
    CuentaContable% = XVALO(RESP_ZELE_VECT(3))

    If NumeroEntrega$ <> "" Then
        FILTX$ = NumeroEntrega$ & ";" & CStr(CuentaContable%)
        Call STDFORM(ARCSQR$, FILTX$, "PRINT")
        
        GoTo 50
    End If

99  Exit Sub

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
   End If
End Sub

Sub Command20_Click()

    Call CONCITCRE  'Llama al formulario CONCITC07 con el listado de cuentas de tarjetas de creditos

End Sub

Private Sub Command21_Click()
   'MUESTRA LA LISTA DIRECTAMENTE LEYENDO EL PDOC_TBL
   'ESTO SE AGREGO PARA TANIT PARA REALIZAR UN RASTREO DE CHEQUES EN LAS IMAGENES DIGITALES
   Command21.Enabled = False
   Condi$ = "tipodocu = 'Deposito Bancario'"
   Call CARGALISEL("INDIDEBA", "PDOC_TBL", "numecomple/A12; tipodocu/A20", Condi$, "NOMESS")
   Call SELECHO("INDIDEBA/Imagenes Digitalizadas (Dep.Bancarios)")
   NumeComp$ = TRIM$(VALACT1$("INDIDEBA"))
   If NumeComp$ = "" Then GoTo 99
   '
   TIPODOCU$ = "Deposito Bancario"
   NUMEBUSQ$ = Mid(NumeComp$, 4, 6)
   Call MUESTRADOC(DOCUNU&, "", TIPODOCU$, NUMEBUSQ$, NumeComp$)
   
99 Command21.Enabled = True

End Sub

Private Sub Command22_Click()
   Command22.Enabled = False
   
   RESPUESTA% = COMALTER%("Gestión de Rechazo de Cheques\\Informar Rechazo\Re-imprimir Débito por Rechazo")
   If RESPUESTA% = 1 Then
        Call CHERECHAZA
        
   ElseIf RESPUESTA% = 2 Then
        ARCSQR$ = Control("", "DEBRECHE")
        If ARCSQR$ = "" Then
            Call COMUNI("Falta Configurar Formulario de Débito por Rechazo")
            GoTo 99
        End If
   
        campos$ = "Fecha/D8;Comprobante/A16;Cta/I10;Cliente/A24;Importe/F14.2"
        Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, campos$, FRMZELECHO, , , ANTOT)
        Call CARGA_ENCABEZADO(FRMZELECHO.msf2, campos$, FRMZELECHO, , , ANTOT)
        
        consultasql$ = "select fechconta, codicom_lar, nume_cli, raso_cli, valabscomp from" _
        & " cajabanc c with(nolock) inner join deudor12 d with(nolock) on (c.nuclien=d.nume_cli)" _
        & " where c.tipomovim='ajud' and c.modoopcimovim='chrecha' order by c.fechconta, codicom_lar"
        Dim rst As ADODB.Recordset
        Set rst = READSET(consultasql$)
        
        Call Carga_MSF_Recordset(rst, campos$, FRMZELECHO.msf2)
        
        On Error Resume Next
        rst.Close
        Set rst = Nothing
        On Error GoTo 0
        
        FRMZELECHO.Caption = "Lista de Débitos por Rechazo"
        FRMZELECHO.Width = 9700
        FRMZELECHO.Show 1
    
        COMPROBANTE$ = RESP_ZELE_VECT(2)
        If COMPROBANTE$ = "" Then GoTo 99
        NumeroCliente& = XVALO(RESP_ZELE_VECT(3))
        
        FILTX$ = COMPROBANTE$ & ";" & Str$(NumeroCliente&) & ";" & "ORIGINAL"
        Call STDFORM(ARCSQR$, FILTX$)
        
   End If
    
99 Command22.Enabled = True

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
      For i& = 1 To ULTREG&("MEPAGO")
         F0$ = REGLEIDO$("MEPAGO", i&)
         CAMP% = i&
         MEPA%(CAMP%) = CVI(Left$(F0$, 2))
         CUEX$ = Mid$(F0$, 45, 12)
         CUEMP%(CAMP%) = CUEINT%(CUEX$)
         If TRIM$(CODCUE$(CUEMP%(CAMP%))) <> TRIM$(CUEX$) Then
            Call MENSERR(24, "Cuenta Contable Incorrecta\para Medio de Pago '" + TRIM$(Mid$(F0$, 3, 12)) + "'\-\Resultados Imprevisibles.")
            Call FINAL("")
         End If
     Next i&
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
   Has% = CVI(Mid$(OBX$, 3, 2))
   If Has% < DES% Then Exit Sub
   '
   If DES% < FECHANUM("01/10/01") Then
      Call MENSERR(24, "Fecha no Válida")
      GoTo 5
   End If
   '
   If Has% > FINEJ% Then
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
          ElseIf FEMOREA% <= Has% Then
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
          ElseIf FEMORE% <= Has% Then
            GoSub 95
         End If
         GoTo 29
      End If
      If FEMOVI% < DES% Then
        If FEMOVI% > FECHANUM("31/09/01") Then GoSub 90
        GoTo 29
      End If
      If FEMOVI% >= DES% Then
        If FEMOVI% <= Has% Then
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
      BANCO$ = Mid$(X$, 15, 16)
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
          If FEMORE% <= Has% Then
            FEX = EMIS%
            DETAPAS$ = "Ch." + TRIM$(NUMERO$)
            DETAPAS$ = DETAPAS$ + " - OP." + TRIM$(Str$(NOPAGO&))
            DETAPAS$ = DETAPAS$ + " - " + TRIM$(rasocli$((-1) * PROV%))
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
       If FEMORE% <= Has% Then
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
   Call SELMESA(DES%, Has%, CORRESP$)
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
           If FEAS% <= Has% Then
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
     TTXXY$ = Chr$(1) + Chr$(3) + MKI$(Has%)
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
     Call CONASIDIA(Has%, CORRESP$)
     '
   End If
   Screen.MousePointer = 1
   Command8.Enabled = True
End Sub

Sub Command9_Click()
   Command9.Enabled = False
   If DERACCE%("STPGVAL", "Configuración de Funcionamiento de Valores") = 2 Then
       OPCHEQ07.Show 1
   End If
   Command9.Enabled = True
End Sub

Private Sub mnuConfiguracionGeneral_Click()
    Call Command9_Click
End Sub

Private Sub mnuConfiguraConciliacionAutomatica_Click()
    Call OPCHEQ07.CONFI_CONCI_AUTO
End Sub

Private Sub mnuConvertirAjusteDebitosPorRechazos_Click()
    
    Call COMUNI("Opción No Disponible")
    Exit Sub
    
    IdComprobante$ = TRIM$(InputBox("Id comprobante", "Convertir Ajuste en Débito"))
    CondicionSql$ = "codicom_lar='" & IdComprobante$ & "' and nuclien>0"
    If CantRegistros("cajabanc", CondicionSql$) > 0 Then
        REFERENCIA$ = TRIM$(InputBox("Referencia", "Convertir Ajuste en Débito"))
        REFERENCIA$ = "Ch.Rech. " & REFERENCIA$
        Call ACTUREGISTRO("cajabanc", "modoopcimovim", CondicionSql$, "CHRECHA", REGAF&)
        Call ACTUREGISTRO("cajabanc", "obsergen", CondicionSql$ & " and tipomovim='ajusald' and varimovim='no-imp'", REFERENCIA$, REGAF&)
    End If
    
End Sub

Private Sub mnuReportesVarios_Click()
    Call AGenRep_Click
End Sub

Private Sub mnuVincularReporte_Click()
    Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
End Sub

Private Sub Option3_DblClick()
   Call INDEXCAJA(MAXDIA&)
   '
   DES% = COMEJ%
   Has% = FINEJ%
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
           ElseIf FECHASIEN% <= Has% Then
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
        ElseIf FECHASIEN% <= Has% Then
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
'   EPAC$ = TRIM$(EMPREAC$)
'   If EPAC$ <> "" Then
'       Caption = Caption + "  -  " + EPAC$
'   End If

    EPAC$ = UCase$(TRIM$(EMPREAC$))
    If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
    End If
    
   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me, Picture7)
   If COMEJ% < 1 Then
      Call CARDATEJ
   End If
   Option6.Value = True
   '
   Call CREACAMPOS_CONCILIA
   If Control$("", "CONCIAU") = "SI" Then
        mnuConfiguraConciliacionAutomatica.Visible = True
        linea2.Visible = True
   End If
   '
   If InStr(1, EPAC$, "FERVI") Or InStr(1, EPAC$, "AMBEST") > 0 Then
        Command22.Visible = True
        Call ABRECONEXION
        FLEXCONN.Execute ("update cajabanc set cuecontai=195 where codicom_cor='db-130115-00' and tipomovim='depoban'")
   End If
   '
   If InStr(1, EPAC$, "MEDITEA") < 1 Then
        Command20.Visible = False
   End If
   
   '
End Sub
Sub CREACAMPOS_CONCILIA()
   Call CREACAMPO("", "CONCILIA", "CUENTACONTBANCO", 3, 0, 1)
   Call CREACAMPO("", "CONCILIA", "CODICONCI", 10, 10)
   Call CREACAMPO("", "CONCILIA", "REFECONCI", 10, 64)
   Call CREACAMPO("", "CONCILIA", "FECHA", 8, 0)
   Call CREACAMPO("", "CONCILIA", "CONCEPTO", 10, 128)
   Call CREACAMPO("", "CONCILIA", "DEBE", 7, 0)
   Call CREACAMPO("", "CONCILIA", "HABER", 7, 0)
   Call CREACAMPO("", "CONCILIA", "STATUS", 3, 0, 1)
   Call CREACAMPO("", "CONCILIA", "COMPROBANTE", 10, 24)
   
   Call CREACAMPO("", "CONCILIA", "FECHACOMPROBANTE", 8, 0)
   Call CREACAMPO("", "CONCILIA", "CUECONTAI", 3, 0)
   Call CREACAMPO("", "CONCILIA", "ID_REFERENCIA", 10, 64)
   Call CREACAMPO("", "CONCILIA", "ORDCONCILIA", 3, 0)
   Call CREACAMPO("", "CONCILIA", "COMPROB_DE_TABLA", 10, 24)
   Call CREACAMPO("", "CONCILIA", "CONCEPTO_SISTEMA", 10, 128)
   Call CREACAMPO("", "CONCILIA", "CUENTACONTBANCO_EXT", 10, 16)

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
   SQLX$ = SQLX$ + "AND DescriMov LIKE  '%" & NCHE$ & "%'"
   '
   Dim CABATEMP As ADODB.Recordset
   Set CABATEMP = New ADODB.Recordset
   Set CABATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
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
         X$ = AJUSTI$(!DESCRIMOV, 40) + String$(8, 0) + ICHX$
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
   CABATEMP.Close
   Set CABATEMP = Nothing
   Screen.MousePointer = 1
   '
   CASEL1& = ULTREG&("CHESELEC")
   If CASEL1& < 1 Then
         Call COMUNI("Número no Encontrado")
         GoTo 30
       ElseIf CASEL1& = 1 Then
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
     '
     Dim CABASEL1 As ADODB.Recordset
     Set CABASEL1 = New ADODB.Recordset
     SQLX$ = "SELECT * FROM CAJABANC "
     SQLX$ = SQLX$ + "where TipoMovim = 'CHRECI' "
     SQLX$ = SQLX$ + "and Status > 0 "
     SQLX$ = SQLX$ + "AND DescriMov = '" & BANUFE$ & "' "
     SQLX$ = SQLX$ + "AND ABS(ValAbsComp - " & IPASX$ & ") < 0.005 "
25   On Error Resume Next
     CABASEL1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
        GoTo 25
     End If
     On Error GoTo 0
     '
     With CABASEL1
        If Not (.EOF And .BOF) Then 'si existe, actualiza
           'ESTO ES PARA RECUPERAR CHEQUES QUE SALIERON EQUIVOCADOS
           If IsNull(!FECHENTRE) = True Or CVINT(!FECHENTRE) <= 33 Then
              BANUFE$ = !DESCRIMOV
              If Mid$(BANUFE$, 19, 2) = "  " Then
                 If Mid$(BANUFE$, 21, 1) <> " " Then
                    BANUFE$ = Left$(BANUFE$, 18) + Mid$(BANUFE$, 21, 12) + "  " + Mid$(BANUFE$, 33, 8)
                    !DESCRIMOV = BANUFE$
                    .Update
                    '
                    SQLY$ = "select * from CAJABANC "
                    SQLY$ = SQLY$ + "WHERE TipoMovim = 'CHENTR' "
                    SQLY$ = SQLY$ + "AND DescriMov = '" & BANUFE$ & "' "
                    SQLY$ = SQLY$ + "AND  MID$(CodiCom_Cor,11,2) <> 'nl'"
                    Set CABATEMP1 = FLEXCONN.Execute(FILTSQL$(SQLX$))           'CueCorri.OpenRecordset(SQLY$, dbOpenSnapshot
                    On Error Resume Next
                    CABATEMP1.MoveFirst
                    If Err < 1 Then
                      On Error GoTo 0
                      RCX$ = "": FCX = "01/01/1980": CCX$ = "": RFX$ = "": FHX$ = "01/01/1980"
                      On Error Resume Next
                      RCX$ = CABATEMP1!refecom
                      FCX = CABATEMP1!fechconta
                      CCX$ = CABATEMP1!CodiCom_Cor
                      RFX$ = CABATEMP1!OBSERGEN
                      FHX$ = CABATEMP1!FeHoReal
                      On Error GoTo 0
                      '
       '               .Edit
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
        End If
     '
        YZ$ = String$(512, 0)
        Call REPLA(YZ$, Mid$(!DESCRIMOV, 1, 32), 71, 32)
        Call REPLA(YZ$, Chr$(!Status), 121, 1)
        Call REPLA(YZ$, MKD$(!ValAbsComp), 145, 8)
        Call REPLA(YZ$, MKI$(CVINT(!FECHEMISI)), 111, 2)
        Call REPLA(YZ$, MKI$(CVINT(!fechconta)), 61, 2)
        Call REPLA(YZ$, MKI$(CVINT(!fechacred)), 63, 2)
        Call REPLA(YZ$, !CodiCom_Cor, 1, 12)
        Call REPLA(YZ$, !refecom, 13, 48)
      
        If IsNull(!DESTICHEQ) = False Then Call REPLA(YZ$, !DESTICHEQ, 293, 30)
        Call REPLA(YZ$, MKI$(CVINT(!FEDEPOSENT)), 275, 2)
        If IsNull(!Document) = False Then Call REPLA(YZ$, !Document, 277, 16)
        If IsNull(!DESTICHEQ) = False Then Call REPLA(YZ$, !DESTICHEQ, 323, 40)
        If IsNull(!REFESALCHE) = False Then Call REPLA(YZ$, !REFESALCHE, 363, 40)
        If IsNull(!IDENSALCHE) = False Then Call REPLA(YZ$, !IDENSALCHE, 403, 40)
        GoTo 50
        '
        Call MENSERR(24, "Imposible Presentar Datos del Cheque")
        GoTo 30
       '
     End With
     CABASEL1.Close
     Set CABASEL1 = Nothing
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
           FECACRE$ = FETEXCOR1(NUEFECHA%)
           '
           BANUFEX$ = BANUFEA$
           Call REPLA(BANUFEX$, NUFECHX$, 33, 8)
           Call APERBASDAT("CABAN")
           SQLX$ = "TipoMovim = 'CHRECI' "
           SQLX$ = SQLX$ + "AND Status > 0 "
           SQLX$ = SQLX$ + "AND DescriMov = '" & BANUFEA$ & "' "
           SQLX$ = SQLX$ + "AND ABS(IDEBEComp - " & IPASX$ & ") < 0.005 "
           Dim CABASELE As ADODB.Recordset
           Set CABASELE = New ADODB.Recordset
           CABASELE.Open (FILTSQL$("SELECT * FROM CAJABANC WHERE " + SQLX$)), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
           If Not (CABASELE.BOF And CABASELE.EOF) Then
'             CABASELE!FechAcred = CDate(NUEFECHA%)
              CABASELE!fechacred = FETEXCOR1$(NUEFECHA%)
              CABASELE!DESCRIMOV = BANUFEX$
              CABASELE.Update
           End If
           CABASELE.Close
           Set CABASELE = Nothing
           '
           SQLY$ = "TipoMovim = 'CHENTR' "
           SQLY$ = SQLY$ + "AND DescriMov = '" & BANUFEA$ & "' "
           SQLY$ = SQLY$ + "AND  SUBSTRING(CodiCom_Cor,11,2) <> 'nl'"
           SQLY$ = SQLY$ + "AND ABS(IHABEComp - " & IPASX$ & ") < 0.005 "
           Set CABASELE = New ADODB.Recordset
           CABASELE.Open (FILTSQL$("SELECT * FROM CAJABANC WHERE " + SQLY$)), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
           If Not (CABASELE.BOF And CABASELE.EOF) Then
              CABASELE!DESCRIMOV = BANUFEX$
              CABASELE.Update
           End If
           CABASELE.Close
           Set CABASELE = Nothing
           '
         End If
       End If
     End If
   End If
   '
   GoTo 30
   '
End Sub
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
   SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK)"
   SQLX$ = SQLX$ + " where TipoMovim = 'CHRECI' "
   SQLX$ = SQLX$ + " and Status > 0 "
   SQLX$ = SQLX$ + " AND ObserGen LIKE  '%" & TRIM$(CUITLIB$) & "%'"
   '
   Set CABATEMP = READSET(FILTSQL$(SQLX$))
   '
   JJ& = 0
   With CABATEMP
       Do While Not .EOF
         If TRIM(!OpciMovim) <> "CH" Then GoTo 39 ' no es un cheque
         ICHX$ = MKD$(!ValAbsComp)
         X$ = AJUSTI$(!DESCRIMOV, 40) + String$(8, 0) + ICHX$
         NCHX$ = Mid$(X$, 19, 12)
         '
         JJ& = JJ& + 1
         Call GRAREG("CHESELEC", X$, JJ&)
39     .MoveNext
       Loop
   End With
   '
   CABATEMP.Close
   Set CABATEMP = Nothing
   '
   Call SETULTREG("CHESELEC", JJ&)
   Call CIERRARCH("CHESELEC")
   Screen.MousePointer = 1
   '
   CASEL1& = ULTREG&("CHESELEC")
   If CASEL1& < 1 Then
         Call COMUNI("Número no Encontrado")
         GoTo 30
       ElseIf CASEL1& = 1 Then
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
   Dim CABASELE As ADODB.Recordset
   Set CABASELE = New ADODB.Recordset
25 On Error Resume Next
   CABASELE.Open FILTSQL$(SQLZ$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   With CABASELE    ' CajaBanc
     'IPASX$ = TRIM$(FORMATNUM$(IPASE#, "F14.2"))
     SQLX$ = "TipoMovim = 'CHRECI' "
     SQLX$ = SQLX$ + "AND Status > 0 "
     SQLX$ = SQLX$ + "AND DescriMov = '" & BANUFE$ & "' "
     SQLX$ = SQLX$ + "AND ABS(ValAbsComp - " & IPASX$ & ") < 0.005 "
     '
     If Not .EOF Then
       'ESTO ES PARA RECUPERAR CHEQUES QUE SALIERON EQUIVOCADOS
      If IsNull(!FECHENTRE) = True Or CVINT(!FECHENTRE) <= 33 Then
         BANUFE$ = !DESCRIMOV
         If Mid$(BANUFE$, 19, 2) = "  " Then
           If Mid$(BANUFE$, 21, 1) <> " " Then
             BANUFE$ = Left$(BANUFE$, 18) + Mid$(BANUFE$, 21, 12) + "  " + Mid$(BANUFE$, 33, 8)
             !DESCRIMOV = BANUFE$
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
               RCX$ = CABATEMP1!refecom
               FCX = CABATEMP1!fechconta
               CCX$ = CABATEMP1!CodiCom_Cor
               RFX$ = CABATEMP1!OBSERGEN
               FHX$ = CABATEMP1!FeHoReal
               '
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
      Call REPLA(YZ$, Mid$(!DESCRIMOV, 1, 32), 71, 32)
      Call REPLA(YZ$, Chr$(!Status), 121, 1)
      Call REPLA(YZ$, MKD$(!ValAbsComp), 145, 8)
      Call REPLA(YZ$, MKI$(CVINT(!FECHEMISI)), 111, 2)
      Call REPLA(YZ$, MKI$(CVINT(!fechconta)), 61, 2)
      Call REPLA(YZ$, MKI$(CVINT(!fechacred)), 63, 2)
      Call REPLA(YZ$, !CodiCom_Cor, 1, 12)
      Call REPLA(YZ$, !refecom, 13, 48)
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
   CABASELE.Close
   Set CABASELE = Nothing
   
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
           FECACRE$ = FETEXCOR1(NUEFECHA%)
           '
           BANUFEX$ = BANUFEA$
           Call REPLA(BANUFEX$, NUFECHX$, 33, 8)
           Call APERBASDAT("CABAN")
           SQLX$ = "TipoMovim = 'CHRECI' "
           SQLX$ = SQLX$ + "AND Status > 0 "
           SQLX$ = SQLX$ + "AND DescriMov = '" & BANUFEA$ & "' "
           SQLX$ = SQLX$ + "AND ABS(IDEBEComp - " & IPASX$ & ") < 0.005 "
           Dim CABASELE2 As ADODB.Recordset
           Set CABASELE2 = New ADODB.Recordset
           CABASELE2.Open (FILTSQL$("SELECT * FROM CAJABANC WHERE " + SQLX$)), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
           If Not (CABASELE2.BOF And CABASELE2.EOF) Then
'             CABASELE!FechAcred = CDate(NUEFECHA%)
              CABASELE2!fechacred = FETEXCOR1$(NUEFECHA%)
              CABASELE2!DESCRIMOV = BANUFEX$
              CABASELE2.Update
           End If
           CABASELE2.Close
           Set CABASELE2 = Nothing
           '
           SQLY$ = "TipoMovim = 'CHENTR' "
           SQLY$ = SQLY$ + "AND DescriMov = '" & BANUFEA$ & "' "
           SQLY$ = SQLY$ + "AND  SUBSTRING(CodiCom_Cor,11,2) <> 'nl'"
           SQLY$ = SQLY$ + "AND ABS(IHABEComp - " & IPASX$ & ") < 0.005 "
           Set CABASELE = New ADODB.Recordset
           CABASELE.Open (FILTSQL$("SELECT * FROM CAJABANC WHERE " + SQLY$)), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
           With CABASELE
             Do While Not .EOF
              CABASELE!DESCRIMOV = BANUFEX$
              CABASELE.Update
              .MoveNext
             Loop
           End With
           CABASELE.Close
           Set CABASELE = Nothing
           '
         End If
       End If
     End If
   End If
   '
   GoTo 30
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
   For i& = 1 To FIN&
      Call TRACE(20, i&, FIN&)
      X$ = REGLEIDO$("BADACHE", i&)
      NCHX$ = Mid$(X$, 19, 12)
      If CVI(Mid$(X$, 133, 2)) < 1 Then  'aun en cartera
         '
         If NCHE$ = "" Then
            If TRIM$(NCHX$) = "" Then
               Call REPLA(X$, MKS$(i&), 45, 4)
               JJ& = JJ& + 1
               Call GRAREG("CHESELEC", X$, JJ&)
            End If
            GoTo 39
         End If
         '
         If InStr(NCHX$, NCHE$) > 0 Then
            Call REPLA(X$, MKS$(i&), 45, 4)
            JJ& = JJ& + 1
            Call GRAREG("CHESELEC", X$, JJ&)
         End If
         '
      End If
39 Next i&
   '
   Call SETULTREG("CHESELEC", JJ&)
   Call CIERRARCH("CHESELEC")
   Screen.MousePointer = 1
   '
   CASEL1& = ULTREG&("CHESELEC")
   If CASEL1& < 1 Then
         Call COMUNI("Número no Encontrado en Cartera Actual.")
         GoTo 30
       ElseIf CASEL1& = 1 Then
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
   Has% = HOY(HOS$)
   FEX = Has%
   HASS$ = FECHATEX$(FEX)
   DESS$ = "01/01/80"
   DES% = FECHANUM(DESS$)
   VDEF$ = MKI$(DES%) + MKI$(Has%)
   '
10 OBX$ = OBJPLA$("DESHESFECHREC", VDEF$)
   If OBX$ = "" Then
      Exit Sub
   End If
   DES% = CVI(Left$(OBX$, 2))
   DESDEFE = CVDAT(DES%)
   Has% = CVI(Mid$(OBX$, 3, 2))
   HASTAFE = CVDAT(Has%)
   '
   Screen.MousePointer = 11
   ASALID$ = "BADACHE1"
   '
   Call APERBASDAT("CABAN")
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "where TipoMovim = 'CHRECI' "
   SQLX$ = SQLX$ + "AND STATUS >= 1 AND STATUS <=2 "
   SQLX$ = SQLX$ + "AND FechConta >= '" & DESDEFE & "' "
   SQLX$ = SQLX$ + "AND FechConta <= '" & HASTAFE & "' "
   SQLX$ = SQLX$ + "ORDER BY FechConta ASC;"
   '
   'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim CABATEMP As ADODB.Recordset
   Set CABATEMP = New ADODB.Recordset
25 On Error Resume Next
   CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0

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
         FEACRED% = CVINT(!fechacred)
         BANUFE$ = !DESCRIMOV
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
         FEMISI% = CVINT%(!FECHEMISI)
         FERECEP% = CVINT(!fechconta)
         NURECIB$ = !CodiCom_Cor
         ORIGCHE$ = !refecom
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
   Set CABATEMP = Nothing
   '
   Call SETULTREG(ASALID$, JJ&)
   Call FILESORT(ASALID$, ASALID$, 5, 5, "ASC")
   Call CIERRARCH(ASALID$)
   Screen.MousePointer = 1
   '
   Call FINAL("*LICHEFRE")
   '
End Sub
'\\\OT-05-0378-WAR-FIN///
Sub LICHECAR()
   '
   If DERACCE%("VECARCHE", "Visualizar Cartera de Cheques y Valores") < 2 Then Exit Sub
   '
   OPXX% = COMALTER%("Opciones de Consulta\\Rango de Fecha de Recepción\Corte por Fecha de Acreditación")
   If OPXX% = 0 Then Exit Sub
   
   OPX% = 1
   If Option4.Value = True Then OPX% = 2
   If OPXX% = 2 Then Option4.Value = True: OPX% = 4
   '
   Call GENLICHECA(OPX%)
   '
   If OPX% = 1 Then
        VTB% = VENTABU%("COCARCHE")
      ElseIf OPX% = 2 Or OPX% = 4 Then
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
   Has% = HOY(HOS$)
   FEX = Has%
   HASS$ = FECHATEX$(FEX)
   DESS$ = "01/01/80"
   DES% = FECHANUM(DESS$)
   VDEF$ = MKI$(DES%) + MKI$(Has%)
   '
   If OPX% = 3 Then    ' LLAMADO DESDE CONCILIACION DE CARTERA
     OBX$ = VDEF$
     GoTo 12
   End If
   '
   If OPX% = 4 Then
     'PRESENTA OBJETO PARA SELECCIONAR FECHA
     VDEF$ = MKI$(HOY(HO$))
8    OBX$ = OBJPLA$("ENTRAFECHA1", VDEF$)
     If OBX$ = "" Then
        OPX% = (-1)
        Exit Sub
     End If
     FE% = CVI(Mid$(OBX$, 1, 2))
     HASTAFE = FETEXCOR1$(FE%)
     GoTo 13
   End If
   
10 OBX$ = OBJPLA$("DESHESFECHREC", VDEF$)
   If OBX$ = "" Then
      OPX% = (-1)
      Exit Sub
   End If
   '
12 DES% = CVI(Left$(OBX$, 2))
   DESDEFE = FETEXCOR1$(DES%)
   Has% = CVI(Mid$(OBX$, 3, 2))
   HASTAFE = FETEXCOR1$(Has%)
   '
13 Screen.MousePointer = 11
   ASALID$ = "LICARCHE"
   If OPX% = 1 Or OPX% = 3 Then ASALID$ = "!" + ASALID$
   '
15 Call APERBASDAT("CABAN")
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "where TipoMovim = 'CHRECI' "
   SQLX$ = SQLX$ + "AND Status = 1 "
   If OPX% <> 4 Then
      SQLX$ = SQLX$ + "AND FechConta >= '" & DESDEFE & "' "
      SQLX$ = SQLX$ + "AND FechConta <= '" & HASTAFE & "' "
   Else
      SQLX$ = SQLX$ + "AND FechAcred <= '" & HASTAFE & "' "
   End If
   SQLX$ = SQLX$ + "ORDER BY FechAcred ASC;"
   '
   'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim CABATEMP As ADODB.Recordset
   Set CABATEMP = New ADODB.Recordset
25 On Error Resume Next
   CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0

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
         FEACRED% = CVINT(!fechacred)
         BANUFE$ = !DESCRIMOV
         ICHEQ# = !ValAbsComp
         '
         If TRIM$(Left$(BANUFE$, 12)) = "" Then
           Call REPLA(BANUFE$, "Cheque Nro:", 1, 11)
         End If
         '
         FEACRED% = CVINT%(!fechacred)
         FEX = FEACRED%
         MESA$ = Mid$(FECHATEX$(FEX), 4, 5)
         FEMISI% = CVINT%(!FECHEMISI)
         FERECEP% = CVINT(!fechconta)
         NURECIB$ = !CodiCom_Cor
         ORIGCHE$ = SAFETEXT$(!refecom)
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
     On Error GoTo 0
   End With
   Set CABATEMP = Nothing
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
   'DESDA = CVDAT(DESEM%)
   DESDA = FETEXCOR1$(DESEM%)
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
   SQLX$ = SQLX$ + " WHERE FechAcred >= '" & DESDA & "' "
   ' OJO, TIENE QUE QUEDAR ASÍ CON EL CDBL, PORQUE SI NO NO FILTRA CORRECTAMENTE - EPB - 22/06/06
   SQLX$ = SQLX$ + " AND TIPOMOVIM = 'CHENTR' "
   SQLX$ = SQLX$ + " ORDER BY FechAcred "
   'Set CHTMP = CueCorri.OpenRecordset(SQLX$, 4)
   Dim CHTMP As ADODB.Recordset
   Set CHTMP = New ADODB.Recordset
25 On Error Resume Next
   CHTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
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
          If !CUECONTAI = CUINTE% Then
             ICUEB% = UB&
             FACRE% = CVINT(!fechacred)
             PROV% = !NUPROVE 'Abs(CVI(Mid$(X$, 67, 2)))
             '
             If PROV% > 0 Then
               NUMERO$ = TRIM$(Mid$(!DESCRIMOV, 17, 16)) 'TRIM$(Mid$(X$, 89, 14))
               BANCO$ = TRIM$(Left$(!DESCRIMOV, 16)) 'TRIM$(Mid$(X$, 71, 18))
               EMIS% = CVINT(!FECHEMISI) 'CVI(Mid$(X$, 111, 2))   ' EMISION=FECHA O/PAGO
               FENTREG% = CVINT(!FECHENTRE) 'CVI(Mid$(X$, 61, 2))
               If EMIS% < 1 Then
                 EMIS% = FENTREG%
               End If
               IMPO# = !ihabecomp 'IMPEX#
               VENCI% = FACRE%
               'NOPA& = Val(Mid$(X$, 4, 6))
               '
               Y$ = REGBLAN$("LICHEQEM")
               Call REPLA(Y$, MKI$(VENCI%), 1, 2)
               Call REPLA(Y$, BANCO$, 3, 16)
               Call REPLA(Y$, NUMERO$, 19, 12)
               Call REPLA(Y$, MKD$(IMPO#), 31, 8)
               Call REPLA(Y$, MKI$(EMIS%), 39, 2)
               Call REPLA(Y$, MKI$(PROV%), 41, 2)
               '
               Z$ = REGBLAN$("CHEQUEMJ")
               Call REPLA(Z$, MKI$(PROV%), 1, 2)
               Call REPLA(Z$, NUMERO$, 3, 12)
               Call REPLA(Z$, BANCO$, 15, 16)
               Call REPLA(Z$, MKI$(EMIS%), 31, 2)
               Call REPLA(Z$, MKD$(IMPO#), 33, 8)
               NOPA& = Val(!NumeComp)
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
   Set CHTMP = Nothing
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
   NCHE$ = TRIM$(NCHE$)
   If NCHE$ = "" Then Exit Sub
   '
   Screen.MousePointer = 11
   '
   Call APERBASDAT("CABAN")
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "where TipoMovim = 'CHRECI' "
   SQLX$ = SQLX$ + "and (Status = 2 or status = 9) "
   SQLX$ = SQLX$ + "AND DescriMov LIKE  '%" & NCHE$ & "%'"
   '
   Dim CABATEMP As ADODB.Recordset
   Set CABATEMP = New ADODB.Recordset
   Set CABATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   JJ& = 0
   With CABATEMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do While Not .EOF
         If TRIM(!OpciMovim) <> "CH" Then GoTo 19 ' no es un cheque
         ICHX$ = MKD$(!ValAbsComp)
         X$ = AJUSTI$(!DESCRIMOV, 40) + String$(8, 0) + ICHX$
         NCHX$ = Mid$(X$, 19, 12)
         '
         JJ& = JJ& + 1
         Call GRAREG("CHESELEC", X$, JJ&)
19     .MoveNext
       Loop
     End If
   End With
   '
   Call SETULTREG("CHESELEC", JJ&)
   Call CIERRARCH("CHESELEC")
   Call FRESHECHO("CHESELEC")
   CABATEMP.Close
   Set CABATEMP = Nothing
   Screen.MousePointer = 1
   '
   CASEL1& = ULTREG&("CHESELEC")
   If CASEL1& < 1 Then
         Call COMUNI("Número no Encontrado")
         GoTo 10
       ElseIf CASEL1& = 1 Then
         Y$ = REGLEIDO$("CHESELEC", 1&)
         GoTo 45
   End If
   '
   Y$ = REGSEL$("CHESELEC")
   If Len(Y$) <= 4 Then GoTo 10
   '
45 BANUFE$ = AJUSTI$(Y$, 40)
   IPASE# = CVD(Mid$(Y$, 49, 8))
   IPASX$ = TRIM$(FORMATNUM$(IPASE#, "F14.2"))
   Call APERBASDAT("CABAN")
   Dim CABASEL1 As ADODB.Recordset
   Set CABASEL1 = New ADODB.Recordset
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "where TipoMovim = 'CHRECI' "
   SQLX$ = SQLX$ + "and Status > 0 "
   SQLX$ = SQLX$ + "AND DescriMov = '" & BANUFE$ & "' "
   SQLX$ = SQLX$ + "AND ABS(ValAbsComp - " & IPASX$ & ") < 0.005 "
47 On Error Resume Next
   CABASEL1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
        GoTo 47
   End If
   On Error GoTo 0
   '
   With CABASEL1
        If Not (.EOF And .BOF) Then 'si existe, actualiza
          YZ$ = String$(512, 0)
          Call REPLA(YZ$, Mid$(!DESCRIMOV, 1, 32), 71, 32)
          Call REPLA(YZ$, Chr$(!Status), 121, 1)
          Call REPLA(YZ$, MKD$(!ValAbsComp), 145, 8)
          Call REPLA(YZ$, MKI$(CVINT(!FECHEMISI)), 111, 2)
          Call REPLA(YZ$, MKI$(CVINT(!fechconta)), 61, 2)
          Call REPLA(YZ$, MKI$(CVINT(!fechacred)), 63, 2)
          Call REPLA(YZ$, !CodiCom_Cor, 1, 12)
          Call REPLA(YZ$, !refecom, 13, 48)
        
          Call REPLA(YZ$, SAFETEXT$(!DESTICHEQ), 293, 30)
          Call REPLA(YZ$, MKI$(CVINT(!FEDEPOSENT)), 275, 2)
          Call REPLA(YZ$, SAFETEXT$(!Document), 277, 16)
          Call REPLA(YZ$, SAFETEXT$(!DESTICHEQ), 323, 40)
          Call REPLA(YZ$, SAFETEXT$(!REFESALCHE), 363, 40)
          Call REPLA(YZ$, SAFETEXT$(!IDENSALCHE), 403, 40)
          CLI% = XVALO(!NUCLIEN)
        Else
          Call MENSERR(24, "Imposible Presentar Datos del Cheque")
          GoTo 10
        End If
   End With
   CABASEL1.Close
   Set CABASEL1 = Nothing
   '
50 X$ = OBJPLA$("MUESTRACHEGES", YZ$)
   '
33 MODO% = COMALTER%("Indique el Origen del Rechazo:\\Proveedor\Banco\Abandonar")
   On MODO% GoTo 64, 70
   GoTo 10
   '
64 HO% = HOY(HOS$)
   Condi$ = "TipoMovim = 'CHENTR' AND DescriMov = '" & BANUFE$ & "' "
   PROV% = XVALO(VALOBADA("CAJABANC", "NUPROVE", Condi$, "NOMBESS"))
   
   VDEF$ = MKI$(PROV%) + MKI$(CLI%) + MKI$(HO%)
   Z$ = OBJPLA$("DEVUCHEPRO", VDEF$)
   If Z$ = "" Then GoTo 10
   '
   PROV% = CVI(Left$(Z$, 2))
   CLI% = CVI(Mid$(Z$, 3, 2)): If CLI% < 1 Then GoTo 64
   FF% = CVI(Mid$(Z$, 5, 2))
   If FF% < COMEJ% Or FF% > HO% Then GoTo 64
   '
   If CantRegistros("DEUDOR12", "NUME_CLI=" & CLI%, "NOMESS") < 1 Then
      Call MENSERR(24, "Imposible Registrar Rechazo.\Cliente no Encontrado.")
      GoTo 10
   End If
   '
   If PROV% > 0 Then
        If CantRegistros("PROVED12", "NUME_CLI=" & PROV%, "NOMESS") < 1 Then
           Call MENSERR(24, "Imposible Registrar Rechazo.\Proveedor No Encontrado.")
           GoTo 10
        End If
   End If
   '
   OKEY% = 0
   DebitoRechazo% = COMALTER%("¿Desea Generar el Débito por Rechazo?\\Si, Generar\No, Continuar")
   If Control$("", "CHERECHA") = "NOTADEBI" Then
        CLIX% = CLI%
        If DebitoRechazo% = 2 Then Call COMUNI("Recuerde Hacer la NOTA DE DEBITO POR RECHAZO\en la Cuenta " + TRIM$(Str$(CLIX%)) + " - " + TRIM$(rasocli$(CLIX%)))
        OKEY% = 1
      Else
        CLIX% = CLI%
        If DebitoRechazo% = 2 Then Call COMUNI("Recuerde Hacer la NOTA DE DEBITO POR RECHAZO\en la Cuenta " + TRIM$(Str$(CLIX%)) + " - " + TRIM$(rasocli$(CLIX%)))
        OKEY% = 1
        'Call GRAJUCLI(CLI%, FF%, ICHE#, OKEY%)
   End If
   CreditoRechazo% = 2
   If PROV% > 0 Then
        CreditoRechazo% = COMALTER%("¿Desea Generar el Movimiento en el PROVEEDOR?\\Si, Generar\No, Continuar")
   End If
    
   If CreditoRechazo% = 2 And DebitoRechazo% = 2 Then GoTo 10
   
   Call SELECHO("ECUECON/C. Contable Correspondiente al Rechazo")
   PRICUEX$ = VALACT1$("ECUECON")
   If PRICUEX$ <> "" Then
       CUECON% = XVALO(VALOBADA("PLANCUEN", "CODINCUE", "CODEXCUE='" & TRIM$(PRICUEX$) & "'", "NOMESS"))
       If CUECON% = 0 Then GoTo 10
   Else
       GoTo 10
   End If
   CCONTABLE_RECHAZO% = CUECON%
   
   If CreditoRechazo% = 1 Then
        Call GRAJUPRO(PROV%, FF%, IPASE#, BANUFE$, OKEY%, CCONTABLE_RECHAZO%)
   End If
   If DebitoRechazo% = 1 Then
        Call GRAJUCLI(CLng(CLI%), FF%, IPASE#, BANUFE$, CCONTABLE_RECHAZO%, OKEY%)
   End If
   If OKEY% = 2 Then
        Call COMUNI("Rechazo Registrado")
      Else
        Call COMUNI("Rechazo mal Registrado.\Consulte !!!")
        GoTo 10
   End If
   '
   Condi$ = "TipoMovim = 'CHRECI' AND DescriMov = '" & BANUFE$ & "' AND NUCLIEN=" & CLI%
   Call ACTUREGISTRO("CAJABANC", "STATUS", Condi$, 8, REGAF&, "NOMESS")
   '
   GoTo 10
   '
70 HO% = HOY(HOS$)
   Condi$ = "TipoMovim = 'CHRECI' AND DescriMov = '" & BANUFE$ & "' "
   CLI% = XVALO(VALOBADA("CAJABANC", "NUCLIEN", Condi$, "NOMBESS"))
   '
71 VDEF$ = Space$(24) + MKI$(HO%) + MKI$(CLI%)
   CUECHERA$ = AJUSTI$(Control$("SVD451S", "CUEVE-8"), 12) ' CUENTA DE RECHAZO DE CHEQUES
   Call REPLA(VDEF$, CUECHERA$, 13, 12)
72 Z$ = OBJPLA$("RECHACHEBA", VDEF$)
   VDEF$ = Z$
   If Z$ = "" Then Exit Sub
   '
   CLI% = CVI(Mid$(Z$, 27, 2)): If CLI% < 1 Then GoTo 72
   FF% = CVI(Mid$(Z$, 25, 2)): If FF% < 1 Then GoTo 72
   '
   For II% = 1 To 2
       CTA%(II%) = 0
       IMPO#(II%) = 0
       DETA$(II%) = ""
   Next II%
   '
   CUEBANCO$ = Left$(Z$, 12): CTA%(2) = 0
   CUECHERA$ = Mid$(Z$, 13, 12): CTA%(1) = 0
   '
   CTA%(1) = XVALO(VALOBADA("PLANCUEN", "CODINCUE", "CODEXCUE='" & TRIM$(CUECHERA$) & "'", "NOMESS"))
   CTA%(2) = XVALO(VALOBADA("PLANCUEN", "CODINCUE", "CODEXCUE='" & TRIM$(CUEBANCO$) & "'", "NOMESS"))
   '
   FF% = CVI(Mid$(Z$, 25, 2)): If FF% < 1 Then Exit Sub
   If CTA%(1) < 1 Then GoTo 72
   If CTA%(2) < 1 Then GoTo 72
   'Call GRACONTROL("", "CUECHERA", CUECHERA$)
   '
   IMPO#(1) = IPASE#
   IMPO#(2) = (-1) * IPASE#
   DETA$(1) = DENOCUE$(CTA%(1))
   DETA$(2) = "Cheque Rechazado"
   REFE$ = "Ch.Rech. " + TRIM$(Mid$(BANUFE$, 1, 18)) & "-" & TRIM$(Mid$(BANUFE$, 18, 14)) & "-" & TRIM$(Mid$(BANUFE$, 32))
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
90 Call GRASICAJA(NROAS, FF%, REFE$, CONTA%, CTA%(), IMPO#(), DETA$())
   DebitoRechazo% = COMALTER%("¿Desea Generar el Débito por Rechazo?\\Si, Generar\No, Continuar")
   If DebitoRechazo% = 1 Then
        CCONTABLE_RECHAZO% = CTA%(1)
        Call GRAJUCLI(CLng(CLI%), FF%, IPASE#, BANUFE$, CCONTABLE_RECHAZO%, OKEY%)
   End If
   If Control$("", "CHERECHA") = "NOTADEBI" Then
        CLIX% = CLI%
        If DebitoRechazo% = 2 Then Call COMUNI("Recuerde Hacer la NOTA DE DEBITO POR RECHAZO\en la Cuenta " + TRIM$(Str$(CLIX%)) + " - " + TRIM$(rasocli$(CLIX%)))
      Else
        CLIX% = CLI%
        If DebitoRechazo% = 2 Then Call COMUNI("Recuerde Hacer la NOTA DE DEBITO POR RECHAZO\en la Cuenta " + TRIM$(Str$(CLIX%)) + " - " + TRIM$(rasocli$(CLIX%)))
        'Call GRAJUCLI(CLI%, FF%, ICHE#, OKEY%)
   End If
   '
   Condi$ = "TipoMovim = 'CHRECI' AND DescriMov = '" & BANUFE$ & "' AND NUCLIEN=" & CLI%
   Call ACTUREGISTRO("CAJABANC", "STATUS", Condi$, 8, REGAF&, "NOMESS")
   '
   Call COMUNI("Rechazo Cheque " + BANUFE$ + "\Registrado y Contabilizado.")
   GoTo 10
   '
End Sub

'Sub GRAJUCLI(CLI%, FF%, ICHE#, OKEY%)
'   '
'   RGXX& = REGICLI&(CLI%)
'   If RGXX& < 1 Or RGXX& > ULTREG&("DEUDOR1") Then
'      Call MENSERR(24, "Imposible Registrar Rechazo.\Cliente no Encontrado.")
'      Exit Sub
'   End If
'   '
'   NAJU& = 1 + ULTREG&("FACTUR")
'   '
'   Z$ = REGBLAN$("FACTUR")
'   Call REPLA(Z$, MKI$(CLI%), 1, 2)
'   Call REPLA(Z$, MKS$(NAJU&), 3, 4)
'   Call REPLA(Z$, MKI$(FF%), 7, 2)
'   Call REPLA(Z$, MKI$(FF%), 9, 2)
'   Call REPLA(Z$, MKD$(ICHE#), 15, 8)
'   Call REPLA(Z$, MKD$(ICHE#), 39, 8)
'   Call REGAPP("FACTUR", Z$)
'   '
'   Z$ = REGBLAN$("SVD202")
'   Call REPLA(Z$, MKI$(CLI%), 1, 2)
'   Call REPLA(Z$, MKS$(NAJU&), 5, 4)
'   Call REPLA(Z$, MKI$(FF%), 9, 2)
'   Call REPLA(Z$, MKD$(ICHE#), 11, 8)
'   Call REPLA(Z$, MKI$(FF%), 19, 2)
'   Call REPLA(Z$, MKD$(ICHE#), 21, 8)
'   Call REGAPP("SVD202", Z$)
'   '
'   TTXS$ = REGLEIDO$("DEUDOR2", RGXX&)
'   SALNUE# = (Int(100 * ICHE# + 0.5)) / 100 + (Int(100 * CVD(Mid$(TTXS$, 109, 8)) + 0.5)) / 100
'   Call REPLA(TTXS$, MKD$(SALNUE#), 109, 8)
'   Call GRAREG("DEUDOR2", TTXS$, RGXX&)
'   '
'   Call CIERRARCH("FACTUR")
'   Call CIERRARCH("SVD202")
'   Call CIERRARCH("DEUDOR2")
'   '
'   OKEY% = OKEY% + 1
'   '
'End Sub

Sub GRAJUPRO(PROV%, FF%, ICHE#, NCHE$, OKEY%, Optional CCONTABLE_RECHAZO%)
    '
    FEX = FF%
    FETX$ = FECHATEX$(FEX)
    NRECIB$ = "AS-" + Mid$(FETX$, 7, 2) + Mid$(FETX$, 4, 2) + Mid$(FETX$, 1, 2) + "-"
    DIFESAL# = ICHE#
    '
    OKEY% = 1
    '
'    Call SELECHO("ECUECON/C. Contable Correspondiente al Rechazo")
'    PRICUEX$ = VALACT1$("ECUECON")
'    If PRICUEX$ <> "" Then
'        CUECON% = XVALO(VALOBADA("PLANCUEN", "CODINCUE", "CODEXCUE='" & TRIM$(PRICUEX$) & "'", "NOMESS"))
'        If CUECON% = 0 Then Exit Sub
'    Else
'        Exit Sub
'    End If
'    CCONTABLE_RECHAZO% = CUECON%
    '
    ORMA% = (-1)
    Dim RECITEMP As ADODB.Recordset
    SQLX$ = "SELECT max(CodiCom_Cor) as Maximo FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'AJUP' "
    SQLX$ = SQLX$ + "AND CodiCom_Cor LIKE '" & NRECIB$ + "%' "
    Set RECITEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    If Not IsNull(RECITEMP!maximo) Then
        ORMA% = XVALO(Mid$(RECITEMP!maximo, 11, 2))
    End If
    RECITEMP.Close
    Set RECITEMP = Nothing
    
    ORMX$ = TRIM$(Str$(ORMA% + 1))
    While Len(ORMX$) < 2: ORMX$ = "0" + ORMX$: Wend
    NRECIB$ = NRECIB$ + Left$(ORMX$, 2)
    '
20  NC1% = (-1) * PROV%
    Call APERBASDAT("CABAN")
    '
    CoComCor$ = NRECIB$
    CoComLar$ = NRECIB$
    NumeClie = 0
    NumeProv% = PROV%
    FechEmi = CVDAT(FF%)
    '
    Referen$ = rasocli$(NC1%)
    FeContab = CVDAT(FF%)
    FeAcredi = CVDAT(FF%)
    NumeTal = 0
    NumeCom = XVALO(Mid$(NRECIB$, 4, 6))
    REFE$ = Left$("Ch.Rech. " + NCHE$ + Space$(64), 64)
    Observa$ = REFE$
    '
    CuentaCo% = CUECOINT%(CUEMADRE$(NC1%))
    TiMovi = "AJUP"            ' AJUSTE DE SALDO PROVEEDORES
    OpMovi = "AS"              ' AJUSTE DE SALDO
    VaMovi = "CCORR"           ' CUENTA CORRIENTE
    Descrip = "Aj.Saldo - " + rasocli$(NC1%)
    ImpoDeb# = (-1) * DIFESAL#
    Call GRACABAN
    '
'4.- Graba la Contra-Partida en ICONI%
    'CUE$ = TRIM$(CONTROL("SVD451S", "CUEVE-8"))
    ICONI% = CCONTABLE_RECHAZO%           ' CUEINT%(CUE$)
    CuentaCo% = ICONI%
    TiMovi = "AJUSALPR"        ' AJUSTE DE SALDO PROVEEDORES
    OpMovi = "AS"              ' AJUSTE DE SALDO
    VaMovi = "NO-IMP"          ' NO IMPUTADO
    Descrip = "Aj.Saldo - " + rasocli$(NC1%)
    ImpoDeb# = DIFESAL#
    Call GRACABAN
    '
'5.- Grabar Vencimientos del Ajuste ******************************************
    SIGNO% = 1
    IBruTot# = DIFESAL#
    IIvaTot# = 0
    '
    NuorVen% = 1
    FeVenc = CVDAT(FF%)
    ImVenc# = DIFESAL#
    Call GRACREPEN
    '
    OKEY% = OKEY% + 1
    '
End Sub

Sub GRAJUCLI(CLI&, FF%, ICHE#, NCHE$, CCONTABLE_RECHAZO%, OKEY%)
    '
    FEX = FF%
    FETX$ = FECHATEX$(FEX)
    NRECIB$ = "AS-" + Mid$(FETX$, 7, 2) + Mid$(FETX$, 4, 2) + Mid$(FETX$, 1, 2) + "-"
    DIFESAL# = ICHE#
    '
    OKEY% = 1
    '
    CUECON% = CCONTABLE_RECHAZO%
    '
    GastoAdminGravado# = XVALO(InputBox("Gastos Adm. (Iva 21%)", "Gastos Administrativos"))
    If GastoAdminGravado# > 0.005 Then
        IvaGastoAdminGravado# = GastoAdminGravado# * 0.21
    End If
    
    ORMA% = (-1)
    Dim RECITEMP As ADODB.Recordset
    SQLX$ = "SELECT max(CodiCom_Cor) as Maximo FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'AJUD' "
    SQLX$ = SQLX$ + "AND CodiCom_Cor LIKE '" & NRECIB$ + "%' "
    Set RECITEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    If Not IsNull(RECITEMP!maximo) Then
        ORMA% = XVALO(Mid$(RECITEMP!maximo, 11, 2))
    End If
    RECITEMP.Close
    Set RECITEMP = Nothing
    
    ORMX$ = TRIM$(Str$(ORMA% + 1))
    While Len(ORMX$) < 2: ORMX$ = "0" + ORMX$: Wend
    NRECIB$ = NRECIB$ + Left$(ORMX$, 2)
    '
20  NC1& = CLI&
    Call APERBASDAT("CABAN")
    '
    CoComCor$ = NRECIB$
    CoComLar$ = NRECIB$
    NumeClie = NC1&
    NumeProv% = 0
    FechEmi = CVDAT(FF%)
    '
    Referen$ = rasocli$(NC1&)
    FeContab = CVDAT(FF%)
    FeAcredi = CVDAT(FF%)
    NumeTal = 0
    NumeCom = XVALO(Mid$(NRECIB$, 4, 6))
    REFE$ = Left$("Ch.Rech. " + NCHE$ + Space$(64), 64)
    Observa$ = REFE$
    '
    CuentaCo% = CUECOINT%(CUEMADRE$(NC1&))
    TiMovi = "AJUD"            ' AJUSTE DE SALDO CLIENTES
    OpMovi = "AS"              ' AJUSTE DE SALDO
    VaMovi = "CCORR"           ' CUENTA CORRIENTE
    Descrip = "Aj.Saldo - " + rasocli$(NC1&)
    ImpoDeb# = DIFESAL# + GastoAdminGravado# + IvaGastoAdminGravado#
    ModoOpMovi$ = "CHRECHA"
    Call GRACABAN
    '
'4.- Graba la Contra-Partida en ICONI%
    ICONI% = CUECON%           ' CUEINT%(CUE$)
    CuentaCo% = ICONI%
    TiMovi = "AJUSALD"         ' AJUSTE DE SALDO CLIENTES
    OpMovi = "AS"              ' AJUSTE DE SALDO
    VaMovi = "NO-IMP"          ' NO IMPUTADO
    Descrip = "Aj.Saldo - " + rasocli$(NC1&)
    ImpoDeb# = (-1) * DIFESAL#
    Call GRACABAN
    
    If GastoAdminGravado# > 0.005 Then
        ICONI% = CUECON%           ' CUEINT%(CUE$)
        CuentaCo% = ICONI%
        TiMovi = "AJUSALD"         ' AJUSTE DE SALDO CLIENTES
        OpMovi = "AS"              ' AJUSTE DE SALDO
        VaMovi = "GGLES"           ' NO IMPUTADO
        Descrip = "Aj.Saldo - " + rasocli$(NC1&)
        Observa$ = "Gastos Generales"
        ImpoDeb# = (-1) * GastoAdminGravado#
        Call GRACABAN
    End If
    
    If IvaGastoAdminGravado# > 0.005 Then
        ICONI% = CUECON%          ' IVA DEBITO 21%
        CuentaCo% = ICONI%
        TiMovi = "AJUSALD"        ' AJUSTE DE SALDO CLIENTES
        OpMovi = "AS"             ' AJUSTE DE SALDO
        VaMovi = "IVA"            ' NO IMPUTADO
        Descrip = "Aj.Saldo - " + rasocli$(NC1&)
        Observa$ = "IVA"
        ImpoDeb# = (-1) * IvaGastoAdminGravado#
        Call GRACABAN
    End If
    '
'5.- Grabar Vencimientos del Ajuste ******************************************
    SIGNO% = 1
    IBruTot# = DIFESAL# + GastoAdminGravado# + IvaGastoAdminGravado#
    IIvaTot# = 0
    '
    NuorVen% = 1
    FeVenc = CVDAT(FF%)
    ImVenc# = DIFESAL# + GastoAdminGravado# + IvaGastoAdminGravado#
    Call GRADEUPEN
    '
    ModoOpMovi$ = ""
    OKEY% = OKEY% + 1
    '
    ARCSQR$ = Control("", "DEBRECHE")
    If ARCSQR$ <> "" Then
        FILTX$ = NRECIB$ & ";" & Str$(NC1&) & ";" & "ORIGINAL"
        Call STDFORM(ARCSQR$, FILTX$)
    End If
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
      For i& = 1 To ULTREG&("MEPAGO")
         F0$ = REGLEIDO$("MEPAGO", i&)
         CAMP% = i&
         MEPA%(CAMP%) = CVI(Left$(F0$, 2))
         CUEX$ = Mid$(F0$, 45, 12)
         CUEMP%(CAMP%) = CUEINT%(CUEX$)
         If TRIM$(CODCUE$(CUEMP%(CAMP%))) <> TRIM$(CUEX$) Then
            Call MENSERR(24, "Cuenta Contable Incorrecta\para Medio de Pago '" + TRIM$(Mid$(F0$, 3, 12)) + "'\-\Resultados Imprevisibles.")
            Call FINAL("")
         End If
     Next i&
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
   Has% = CVI(Mid$(OBX$, 3, 2))
   If Has% < DES% Then Exit Sub
   '
   If DES% < FECHANUM("01/10/01") Then
      Call MENSERR(24, "Fecha no Válida")
      GoTo 5
   End If
   '
   If Has% > FINEJ% Then
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
          ElseIf FEMOREA% <= Has% Then
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
          ElseIf FEREGI% <= Has% Then
            GoSub 95
         End If
         GoTo 29
      End If
      If FEMOVI% < DES% Then
        GoSub 90
        GoTo 29
      End If
      If FEMOVI% >= DES% Then
        If FEMOVI% <= Has% Then
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
      BANCO$ = Mid$(X$, 15, 16)
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
          If FEREGI% <= Has% Then
            FEX = EMIS
            DETAPAS$ = "Ch." + TRIM$(NUMERO$)
            DETAPAS$ = DETAPAS$ + " - OP." + TRIM$(Str$(NOPAGO&))
            DETAPAS$ = DETAPAS$ + " - " + TRIM$(rasocli$((-1) * PROV%))
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
      For i& = 1 To ULTREG&("MEPAGO")
         F0$ = REGLEIDO$("MEPAGO", i&)
         CAMP% = i&
         MEPA%(CAMP%) = CVI(Left$(F0$, 2))
         CUEX$ = Mid$(F0$, 45, 12)
         CUEMP%(CAMP%) = CUEINT%(CUEX$)
         If TRIM$(CODCUE$(CUEMP%(CAMP%))) <> TRIM$(CUEX$) Then
            Call MENSERR(24, "Cuenta Contable Incorrecta\para Medio de Pago '" + TRIM$(Mid$(F0$, 3, 12)) + "'\-\Resultados Imprevisibles.")
            Call FINAL("")
         End If
     Next i&
   End If
   '
   HOII% = HOY(HOS$)
   DES% = COMEJ%
   Has% = FINEJ%
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
          ElseIf FEMOREA% <= Has% Then
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
          ElseIf FEMORE% <= Has% Then
            GoSub 95
         End If
         GoTo 29
      End If
      If FEMOVI% < DES% Then
        If FEMOVI% > FECHANUM("31/09/01") Then GoSub 90
        GoTo 29
      End If
      If FEMOVI% >= DES% Then
        If FEMOVI% <= Has% Then
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
      BANCO$ = Mid$(X$, 15, 16)
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
          If FEMORE% <= Has% Then
            FEX = EMIS%
            DETAPAS$ = "Ch." + TRIM$(NUMERO$)
            DETAPAS$ = DETAPAS$ + " - OP." + TRIM$(Str$(NOPAGO&))
            DETAPAS$ = DETAPAS$ + " - " + TRIM$(rasocli$((-1) * PROV%))
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
5  Has% = HOY(HOS$)
   DES% = COMESACT%
   VDEF$ = MKI$(DES%) + MKI$(Has%)
   OBXX$ = OBJPLA$("DESHASFECHA", VDEF$)
   If OBXX$ = "" Then Exit Sub
   '
   DES% = CVI(Left$(OBXX$, 2))
   Has% = CVI(Mid$(OBXX$, 3, 2))
   If Has% < DES% Then GoTo 5
   If DES% < 1 Or Has% > HOY(HOS$) Then GoTo 5
   FEX = DES%: MMXX1$ = FECHATEX$(FEX)
   FEX = Has%: MMXX1$ = MMXX1$ + " - " + FECHATEX$(FEX)
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
          ElseIf FEAS% <= Has% Then
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
5  Has% = HOY(HOS$)
   DES% = COMESACT%
   VDEF$ = MKI$(DES%) + MKI$(Has%)
   OBXX$ = OBJPLA$("DESHASFECHA", VDEF$)
   If OBXX$ = "" Then Exit Sub
   '
   DES% = CVI(Left$(OBXX$, 2))
   Has% = CVI(Mid$(OBXX$, 3, 2))
   If Has% < DES% Then GoTo 5
   If DES% < 1 Or Has% > HOY(HOS$) Then GoTo 5
   FEX = DES%: MMXX1$ = FECHATEX$(FEX)
   FEX = Has%: MMXX1$ = MMXX1$ + " - " + FECHATEX$(FEX)
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
          ElseIf FEAS% <= Has% Then
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

Sub CONASIDIA(Has%, CORRESP$)
    '
    Screen.MousePointer = 11
    J& = 0
    Call BLANARCH("!ASIDIAFI")
    TTXXY$ = Chr$(1) + Chr$(3) + MKI$(Has%)
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
   CUECHECA$ = Control$("", "CUECHECA")
   CUEICHEQ% = CUEINT%(CUECHECA$)
   If TRIM$(CUECHECA$) = "" Or CUEICHEQ% < 1 Then
      Call MENSERR(24, "Falta Definir Cuenta Contable\Asociada a Cartera de Valores")
      OPCHEQ07.Show 1
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
   SELCHE07.List1.Clear
   SELCHE07.LICHESEL.Clear
   '
   Call APERBASDAT("CABAN")
   SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) "
   SQLX$ = SQLX$ + "where TipoMovim = 'CHRECI' "
   SQLX$ = SQLX$ + "AND Status = 1 "
   '
   'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim CABATEMP As ADODB.Recordset
   Set CABATEMP = READSET(SQLX$)
'   Set CABATEMP = New ADODB.Recordset
'15 On Error Resume Next
'   CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
'   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'      On Error GoTo 0
'      Call CapturaErrorOpen
'      GoTo 15
'   End If
'   On Error GoTo 0
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
         FEACRED% = CVINT(!fechacred)
         BANUFE$ = !DESCRIMOV
         ICHEQ# = !ValAbsComp
         '
         If TRIM$(Left$(BANUFE$, 12)) = "" Then
           Call REPLA(BANUFE$, "Cheque Nro:", 1, 11)
         End If
         RECHE& = !NumeComp
         '
         Z$ = "A" + Space$(80)
         Call REPLA(Z$, CSTRING$(MKI$(FEACRED%), 1, 5, 0, 2), 2, 5)
         Call REPLA(Z$, TRIM$(Left$(BANUFE$, 40)), 9, 40)
         Call REPLA(Z$, CSTRING$(MKD$(ICHEQ#), 4, 12, 2, 2), 49, 12)
         Call REPLA(Z$, CSTRING$(MKS$(RECHE&), 3, 8, 0, 2), 73, 8)
         SELCHE07.List1.AddItem Z$
         '
19     .MoveNext
       Loop
     End If
     On Error GoTo 0
   End With
   Set CABATEMP = Nothing
   '
   SUCHETO# = 0
   For i& = 1 To SELCHE07.List1.ListCount
      Z$ = Mid$(SELCHE07.List1.List(i& - 1), 9)
      SELCHE07.LICHESEL.AddItem Z$
      SUCHETO# = SUCHETO# + Val(Mid$(Z$, 41, 12))
   Next i&
   '
   SELCHE07.Label4.Caption = CSTRING$(MKD$(SUCHETO#), 4, 11, 2, 2)
   SELCHE07.Label4.Refresh
   '
   Screen.MousePointer = 1
   SELCHE07.Show 1
   DoEvents
   '
   IPASE# = Val(SELCHE07.Label2.Caption)
   FEACR% = FECHACRE%(FEDEPO%, HORACRE%)
   If IPASE# > 0.005 Then
      Screen.MousePointer = 11
      NOCLOSE% = 1
      FEX = FEDEPO%
      FETX$ = FECHATEX$(FEX)
      NRECIB$ = "DB-" + Mid$(FETX$, 7, 2) + Mid$(FETX$, 4, 2) + Mid$(FETX$, 1, 2) + "-"
      NRECIBX$ = NRECIB$ + "__"
      ORMA% = (-1)
      '
      Call APERBASDAT("CABAN")
      SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) "
      SQLX$ = SQLX$ + "where CodiCom_Cor LIKE '" & NRECIBX$ & "' "
      'Set CajaBanc1 = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
      Dim CajaBanc1 As ADODB.Recordset
      Set CajaBanc1 = READSET(SQLX$)
'      Set CajaBanc1 = New ADODB.Recordset
'26    On Error Resume Next
'      CajaBanc1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
'      If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'         On Error GoTo 0
'         Call CapturaErrorOpen
'         GoTo 26
'      End If
'      On Error GoTo 0

      With CajaBanc1
        On Error Resume Next
        '.FindFirst "CodiCom_Cor LIKE '" & NRECIBX$ & "'"
        .MoveFirst
30      'If .NoMatch = False Then
        If Err < 1 Then
          On Error GoTo 0
          Do While Not .EOF
            If Val(Mid$(!CodiCom_Cor, 11)) > ORMA% Then
              ORMA% = Val(Mid$(!CodiCom_Cor, 11))
            End If
            '.FindNext "CodiCom_Cor LIKE '" & NRECIBX$ & "'"
            'GoTo 30
            .MoveNext
          Loop
        End If
        On Error GoTo 0
      End With
      Set CajaBanc1 = Nothing
      ORMX$ = TRIM$(Str$(ORMA% + 1))
      While Len(ORMX$) < 2: ORMX$ = "0" + ORMX$: Wend
      NRECIB$ = NRECIB$ + Left$(ORMX$, 2)
      '
      CoComCor$ = NRECIB$
      CoComLar$ = NRECIB$
'      CodiEmp% = 0
'      CodiSuc% = SUC%
      NumeClie = 0
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
      'SQLX$ = "SELECT * FROM CAJABANC "
      'SQLX$ = SQLX$ + "where TipoMovim = 'CHRECI' "
      'SQLX$ = SQLX$ + "AND Status = 1 "
      'Set CajaBanc2 = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
      '
      Dim CajaBanc2 As ADODB.Recordset

      For i& = 1 To SELCHE07.LICHESEL.ListCount
        If SELCHE07.LICHESEL.Selected(i& - 1) = True Then
          X$ = SELCHE07.LICHESEL.List(i& - 1)
          BANUFE$ = Left$(X$, 40)
          BACH$ = TRIM$(Left$(X$, 16))
          NCHE$ = Mid$(X$, 19, 12)
          IPASE# = Val(Mid$(X$, 41, 12))
          SUMACHESEL# = SUMACHESEL# + IPASE#
          REGBA& = Val(Mid$(X$, 65, 8))
          '
'          SQLX$ = "DescriMov = '" & BANUFE$ & "' "
'          SQLX$ = SQLX$ + "AND ValAbsComp = " & Str$(IPASE#) & " "
'          SQLX$ = SQLX$ + "AND NumeComp = " & REGBA& & " "
'          SQLX$ = SQLX$ + "AND Status = 1"
          SQLX$ = "SELECT * FROM CAJABANC "
          SQLX$ = SQLX$ + "where TipoMovim = 'CHRECI' "
          SQLX$ = SQLX$ + " AND Status = 1 "
          SQLX$ = SQLX$ + "AND DescriMov = '" & BANUFE$ & "' "
          SQLX$ = SQLX$ + "AND ValAbsComp = " & Str$(IPASE#) & " "
          SQLX$ = SQLX$ + "AND NumeComp = " & REGBA& & " "
          Set CajaBanc2 = New ADODB.Recordset
25        On Error Resume Next
          CajaBanc2.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
          If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
            On Error GoTo 0
            Call CapturaErrorOpen
            GoTo 25
          End If
          On Error GoTo 0

          On Error Resume Next
          With CajaBanc2
            .MoveFirst
            If Err < 1 Then
            '.FindFirst SQLX$
            'If .NoMatch = False Then
              On Error GoTo 0
              '.Edit
              !FEDEPOSENT = CVDAT(FEDEPO%)
              !Document = NRECIB$
              !DESTICHEQ = DESTICHE$
              CUECONTABINT% = !CUECONTAI
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
'          CodiEmp% = 0
'          CodiSuc% = SUC%
          NumeClie = 0
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
      Next i&
      NOCLOSE% = 0
      Set CajaBanc2 = Nothing
      '
      '\\\OT-05-0775-WAR-07/07/05///
      'IMPRESION DE DEPOSITO BANCARIO
      FORIPRE$ = TRIM$(Control$("", "FORDEPBA"))
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
         For i& = 1 To SELCHE07.LICHESEL.ListCount
           If SELCHE07.LICHESEL.Selected(i& - 1) = True Then
             X$ = SELCHE07.LICHESEL.List(i& - 1)
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
         Next i&
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
  Call APERBASDAT("CABAN")
  Screen.MousePointer = 11
  '
  ' ENJUAGUE GABAL
  HOII% = HOY(HOS$)
'  If FECHATEX$(HOII%) = "21/12/09" Then
'     Call ABRECONEXION
'     FLEXCONN.Execute ("UPDATE CajaBanc set CodiEmpr = 1 WHERE CodiEmpr < 1")
'  End If
  '
  Call SETULTREG("SELDEPO", 0)
  SQLX$ = " SELECT *  FROM CAJABANC"
  SQLX$ = SQLX$ + " WHERE TIPOMOVIM = 'DEPOBAN'"
  'SQLX$ = SQLX$ + " AND(MarConsi = '' or ISNULL(MarConsi) = True or IsEmpty(MarConsi) = True)"
  SQLX$ = SQLX$ + " AND(MarConsi = '' or MarConsi IS NULL )"
  SQLX$ = SQLX$ + " ORDER BY CODICOM_COR ASC"
  'Set DEPTMP = CueCorri.OpenRecordset(SQLX$, 4)
  Dim DEPTMP As ADODB.Recordset

   Set DEPTMP = New ADODB.Recordset
25 On Error Resume Next
   DEPTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
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
        Call REPLA(ZZ$, MKI(CVINT(!fechconta)), 13, 2)
        J& = J& + 1
        Call GRAREG("SELDEPO", ZZ$, J&)
        NUDEPAN$ = !CodiCom_Cor
        TXBUS$ = TXBUS$ + AJUSTI$(NUDEPAN$, 16)
      End If
      .MoveNext
    Loop
  End With
  Set DEPTMP = Nothing
  '
  ' ESTO ES PARA AGREGAR LOS QUE SALIERON PERO SIN ENTRADA
  SQLX$ = " SELECT *  FROM CAJABANC"
  SQLX$ = SQLX$ + " WHERE TIPOMOVIM = 'CHRECI' AND DocumEnt LIKE 'DB-%' "
  SQLX$ = SQLX$ + " ORDER BY DocumEnt ASC"
  Set DEPTMP = New ADODB.Recordset
26  On Error Resume Next
    DEPTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 26
   End If
   On Error GoTo 0


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
  Set DEPTMP = Nothing
  '
5 On Error GoTo 0

  Call SETULTREG("SELDEPO", J&)
  Call FILESORT("SELDEPO", "", 1, 1, "ASC")
  Call FRESHECHO("SELDEPO")
  Call TRALOCAL("SELDEPO", "")
  Call FRESHECHO("!SELDEPO")
  Screen.MousePointer = 1
  '
  '3.-REGSEL("SELDEPO")
  Y1$ = REGSEL("!SELDEPO")
 
  If Len(Y1$) <= 4 Then GoTo 99
  '
  Screen.MousePointer = 11
  NUDEP1$ = Mid$(Y1$, 1, 12)
  '4.- SELECCIONO Y BORRRO TODOS LOS ARHIVOS DEL CAJABANC
  If TRIM(NUDEP1$) <> "" Then
     Screen.MousePointer = 1
     SQLY$ = " SELECT *  FRoM CAJABANC"
     SQLY$ = SQLY$ + " WHERE CODICOM_LAR = '" & NUDEP1$ & "'"
     SQLY$ = SQLY$ + "OR (TipoMovim = 'CHRECI' "
     SQLY$ = SQLY$ + " AND DocumEnt = '" & NUDEP1$ & "') "
     Dim CajaBanSel As ADODB.Recordset
     Set CajaBanSel = New ADODB.Recordset
27   On Error Resume Next
     CajaBanSel.Open FILTSQL$(SQLY$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
        GoTo 27
     End If
     On Error GoTo 0
     With CajaBanSel
       On Error Resume Next
7      .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
          If IsNull(!MARCONSI) = False And IsEmpty(!MARCONSI) = False Then
            If TRIM$(!MARCONSI) <> "" Then
              Call MENSERR(24, "Imposible Anular Movimiento Conciliado.")
              GoTo 99
            End If
          End If
          .MoveNext
         Loop
       End If
       On Error GoTo 0
     End With
     Set CajaBanSel = Nothing
       '
10     SQLD$ = " DELETE FROM CAJABANC"
       SQLD$ = SQLD$ + " WHERE CODICOM_LAR = '" & NUDEP1$ & "'"
       FLEXCONN.Execute (FILTSQL$(SQLD$))
       '
22     SQLX$ = " SELECT * FROM CAJABANC "
       SQLX$ = SQLX$ + " WHERE TipoMovim = 'CHRECI' "
       SQLX$ = SQLX$ + " AND DocumEnt = '" & NUDEP1$ & "' "
       Dim CAJJJABAN As ADODB.Recordset
       Set CAJJJABAN = New ADODB.Recordset
28     On Error Resume Next
       CAJJJABAN.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
          On Error GoTo 0
          Call CapturaErrorOpen
          GoTo 28
       End If
       On Error GoTo 0
       On Error Resume Next
       With CAJJJABAN
         .MoveFirst
         If Err < 1 Then
            On Error GoTo 0
            Do While Not .EOF
               !Status = 1     'LO PONGO DEVUELTA COMO PENDIENTE EN CARTERA DE VALORES
               !Document = ""  'QUITO EL NUMERO DE DEPOSITO ASIGNADO
               !DESTICHEQ = ""
               !REFESALCHE = ""
               !IDENSALCHE = ""
               !FEDEPOSENT = CVDAT(0) 'fecha de deposito
               .Update
               .MoveNext
            Loop
         End If
       End With
       Set CAJJJABAN = Nothing
     '
     Screen.MousePointer = 1
     Call COMUNI("Se ha Anulado el Déposito Bancario: '" & NUDEP1$ & "'.")
  End If
  '
99 Screen.MousePointer = 1
End Sub

Sub ANUENCHE() 'ANULACION
  '
  '1.- SELECCION MOVI CAJABANC SEGUN BANCO Y SIN REPETIR Y CON MARCONSI IGUAL NADA
  Screen.MousePointer = 11
  Call APERBASDAT("CABAN")
  Screen.MousePointer = 11
  SQLX$ = " SELECT *  FRoM CAJABANC"
  SQLX$ = SQLX$ + " WHERE TIPOMOVIM = 'ENTRECHE'"
  SQLX$ = SQLX$ + " AND(MarConsi = '' or MarConsi IS NULL )"
  SQLX$ = SQLX$ + " ORDER BY CODICOM_COR ASC"
  Dim DEPTMP As ADODB.Recordset
  
  Set DEPTMP = New ADODB.Recordset
25 On Error Resume Next
   DEPTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0


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
    ZZ$ = REGBLAN$("!SELDEPO")
    NUDEPAN$ = "": TXBUS$ = ""
    Do While .EOF = False
      If NUDEPAN$ <> !CodiCom_Cor Then
        Call REPLA(ZZ$, !CodiCom_Cor, 1, 12)
        Call REPLA(ZZ$, MKI(CVINT(!fechconta)), 13, 2)
        J& = J& + 1
        Call GRAREG("!SELDEPO", ZZ$, J&)
        NUDEPAN$ = !CodiCom_Cor
        TXBUS$ = TXBUS$ + AJUSTI$(NUDEPAN$, 16)
      End If
      .MoveNext
    Loop
  End With
  Set DEPTMP = Nothing
  '
  ' ESTO ES PARA AGREGAR LOS QUE SALIERON PERO SIN ENTRADA
  SQLX$ = " SELECT *  FRoM CAJABANC"
  SQLX$ = SQLX$ + " WHERE TIPOMOVIM = 'CHRECI' AND DocumEnt LIKE 'EV-%' "
  SQLX$ = SQLX$ + " ORDER BY DocumEnt ASC"
  Set DEPTMP = New ADODB.Recordset
29 On Error Resume Next
   DEPTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 29
   End If
   On Error GoTo 0

 
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
        J& = J& + 1
        Call GRAREG("!SELDEPO", ZZ$, J&)
        TXBUS$ = TXBUS$ + AJUSTI$(NUDEPAN$, 16)
      End If
      .MoveNext
    Loop
  End With
  Set DEPTMP = Nothing
  '
5 On Error GoTo 0
  Call SETULTREG("!SELDEPO", J&)
  Call TRACENTRAL("SELDEPO", "")
  Call FILESORT("SELDEPO", "", 1, 1, "ASC")
  Call FRESHECHO("SELDEPO")
  Call TRALOCAL("SELDEPO", "")
  Call FRESHECHO("!SELDEPO")
  Screen.MousePointer = 1
  '
  '3.-REGSEL("SELDEPO")
  Y1$ = REGSEL("!SELDEPO")
  If Len(Y1$) <= 4 Then GoTo 99
  '
  Screen.MousePointer = 11
  NUDEP1$ = Mid$(Y1$, 1, 12)
  '4.- SELECCIONO Y BORRRO TODOS LOS ARHIVOS DEL CAJABANC
  If TRIM(NUDEP1$) <> "" Then
     Screen.MousePointer = 11
     SQLY$ = " SELECT *  FRoM CAJABANC"
     SQLY$ = SQLY$ + " WHERE CODICOM_LAR = '" & NUDEP1$ & "'"
     SQLY$ = SQLY$ + "OR (TipoMovim = 'CHRECI' "
     SQLY$ = SQLY$ + " AND DocumEnt = '" & NUDEP1$ & "') "
     Dim CajaBanSel As ADODB.Recordset
     Set CajaBanSel = New ADODB.Recordset
30   On Error Resume Next
     CajaBanSel.Open FILTSQL$(SQLY$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
        GoTo 30
     End If
     On Error GoTo 0

     With CajaBanSel
       On Error Resume Next
7      .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
          If IsNull(!MARCONSI) = False And IsEmpty(!MARCONSI) = False Then
            If TRIM$(!MARCONSI) <> "" Then
              Call MENSERR(24, "Imposible Anular Movimiento Conciliado.")
              GoTo 99
            End If
          End If
          .MoveNext
         Loop
       End If
       On Error GoTo 0
     End With
     Set CajaBanSel = Nothing
     '
10   SQLD$ = " DELETE FROM CAJABANC"
     SQLD$ = SQLD$ + " WHERE CODICOM_LAR = '" & NUDEP1$ & "'"
      FLEXCONN.Execute (FILTSQL$(SQLD$))
     '
22   SQLX$ = " SELECT * FROM CAJABANC "
       SQLX$ = SQLX$ + " WHERE TipoMovim = 'CHRECI' "
       SQLX$ = SQLX$ + " AND DocumEnt = '" & NUDEP1$ & "' "
       Dim CAJJJABAN As ADODB.Recordset
       Set CAJJJABAN = New ADODB.Recordset
27     On Error Resume Next
       CAJJJABAN.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
          On Error GoTo 0
          Call CapturaErrorOpen
          GoTo 27
       End If
       On Error GoTo 0

       On Error Resume Next
       With CAJJJABAN
         .MoveFirst
         If Err < 1 Then
            On Error GoTo 0
            Do While Not .EOF
               !Status = 1     'LO PONGO DEVUELTA COMO PENDIENTE EN CARTERA DE VALORES
               !Document = ""  'QUITO EL NUMERO DE DEPOSITO ASIGNADO
               !DESTICHEQ = ""
               !REFESALCHE = ""
               !IDENSALCHE = ""
               !FEDEPOSENT = CVDAT(0) 'fecha de deposito
               .Update
               .MoveNext
            Loop
         End If
       End With
       Set CAJJJABAN = Nothing
     '
     Screen.MousePointer = 1
     Call COMUNI("Se ha Anulado la Entrega de Valores: '" & NUDEP1$ & "'.")
  End If
  '
99 Screen.MousePointer = 1
End Sub

Sub TRANFONDOS()
   '
   HOII% = HOY(HOS$)
   CUECHECA$ = Control$("", "CUECHECA")
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
   If CUEDEBE$ = "" Or CUDEB1% < 1 Or CUDEB1% = CUEICHEQ% Then
     Call MENSERR(24, "Cuenta no Válida")
     GoTo 10
   End If
   '
   CUEHABE$ = TRIM$(Mid$(OBX$, 11, 12))
   CUHAB1% = CUECOINT%(CUEHABE$)
   If CUEHABE$ = "" Or CUHAB1% < 1 Or CUHAB1% = CUEICHEQ% Then
     Call MENSERR(24, "Cuenta no Válida")
     GoTo 10
   End If
   '
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
   NRECIBX$ = NRECIB$ + "__"
   ORMA% = (-1)
   Call APERBASDAT("CABAN")
   Screen.MousePointer = 11
   'With CajaBanc
   SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE CodiCom_Cor LIKE '" & NRECIBX$ & "'"
   Dim CAJABTEMP As ADODB.Recordset
   Set CAJABTEMP = READSET(SQLX$)
'   Set CAJABTEMP = New ADODB.Recordset
'25 On Error Resume Next
'   CAJABTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
'   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'      On Error GoTo 0
'      Call CapturaErrorOpen
'      GoTo 25
'   End If
'   On Error GoTo 0

   '.FindFirst "CodiCom_Cor LIKE '" & NRECIBX$ & "'"
   On Error Resume Next
   With CAJABTEMP
30   'If .NoMatch = False Then
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do While Not .EOF
          If Val(Mid$(!CodiCom_Cor, 11)) > ORMA% Then
            ORMA% = Val(Mid$(!CodiCom_Cor, 11))
          End If
          '.FindNext "CodiCom_Cor LIKE '" & NRECIBX$ & "'"
          'GoTo 30
          .MoveNext
       Loop
     End If
     On Error GoTo 0
   End With
   Set CAJABTEMP = Nothing
   ORMX$ = TRIM$(Str$(ORMA% + 1))
   While Len(ORMX$) < 2: ORMX$ = "0" + ORMX$: Wend
   NRECIB$ = NRECIB$ + Left$(ORMX$, 2)
   '
   CoComCor$ = NRECIB$
   CoComLar$ = NRECIB$
'   CodiEmp% = 0
'   CodiSuc% = SUC%
   NumeClie = 0
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
   Descrip = "De " + TRIM$(ECOARCH$("ECUECON", CUEDEBE$))
   ImpoDeb# = ITRAN#
   StatInit = 0
   Call GRACABAN
   '
   CuentaCo% = CUHAB1%
   Descrip = "A " + TRIM$(ECOARCH$("ECUECON", CUEHABE$))
   ImpoDeb# = (-1) * ITRAN#
   Call GRACABAN
   '\\\OT-05-0440-WAR-04/07/05///
   'IMPRIMIR
    MODOIMPRE07.Check7.Caption = "Imprime Transferencia"
    MODOIMPRE07.Label1.Caption = "" 'LIMPIA EL LABEL DONDE PONE OK, SI ACEPTA
    MODOIMPRE07.Check8.Value = 0
    Screen.MousePointer = 1
    MODOIMPRE07.Show 1
    If MODOIMPRE07.Label1.Caption = "OK" Then
     If MODOIMPRE07.Check8.Value > 0 Then CANCELPRINT% = 1
    End If
    Screen.MousePointer = 11

    FORIPRE$ = TRIM$(Control$("", "FORTRAFO"))
    If FORIPRE$ <> "" Then
237     LU$ = LUDAT$
        CODFOR$ = TRIM$(UCase$(FORIPRE$))
        'Call SAVESPOOL("Orden de Fabricación " + NORFA$, OKX%)
        TIPODOC$ = "Transferencia de Fondo"
        'NUMEDOC = FECHA INVERTIDA MAS EL NUMERO DE ORDEN 130116-01
        NUMEDOC$ = SAFETEXT$(XVALO(Mid$(NRECIB$, 4, 6))) & "-" & Right("00" & SAFETEXT$(XVALO(Mid$(NRECIB$, 11, 2))), 2)
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
          Call PRINTFORWIN(NUMEDOC$, "NUME-COM", 0, 0)
          Call PRINTFORWIN(USERNAME$, "USUARIO", 0, 0)
          Call PRINTFORWIN(REFE$, "REF-ASIE", 0, 0)
          Call PRINTFORWIN(CODCUE$(CUDEB1%), "COD-CUEN", 0, YAC)
          Call PRINTFORWIN(CODCUE$(CUHAB1%), "COD-CUEN", 0, YAC + INTERLI)
          Call PRINTFORWIN(TRIM$(ECOARCH$("ECUECON", CUEDEBE$)), "DEN-CUEN", 0, YAC)
          Call PRINTFORWIN(TRIM$(ECOARCH$("ECUECON", CUEHABE$)), "DEN-CUEN", 0, YAC + INTERLI)
    
          'CANS$ = CSTRING$(MKS$(ITRAN#), 3, 12, 2, 2)
          CANS$ = FORMATNUM$((ITRAN#), "F12.2")
          Call PRINTFORWIN(CANS$, "IMP-DEBE", 0, YAC)
          Call PRINTFORWIN(REFE0$, "REF-MAT1", 0, YAC)
          '
          'Printer.EndDoc
          Call SETSPOOL("ENDDOC", "")
          '
        Next COP%
        Call SAVESPOOL("Transferencia de Fondo", OKX%)
   End If
   '
   Screen.MousePointer = 1
   CANCELPRINT% = 0
   Condi$ = "TIPOMOVIM = 'TRAFON'  ORDER BY CODICOM_COR DESC "
   Call CARGALISEL("ANTRAFON", "CAJABANC", "CODICOM_COR/A12;FECHCONTA/D8", Condi$, "DISTINCT")

   Call COMUNI("Transferencia Completa y Verificada")
   GoTo 5
   '
End Sub

Sub ANUTRANFONDOS()
   '
   Call COMUNI("Recuerde que no Podrán Anularse\las Transferencias ya Conciliadas.")
   '
   Condi$ = "TIPOMOVIM = 'TRAFON'"
   '
   REGAF& = CantRegistros&("CAJABANC", Condi$)
   If REGAF& < 1 Then
      Call COMUNI("No Existen Transferencias de Fondos que Puedan Anularse.")
      Exit Sub
   End If
   '
   Condi$ = "TIPOMOVIM = 'TRAFON'  ORDER BY CODICOM_COR DESC "
   Call CARGALISEL("ANTRAFON", "CAJABANC", "CODICOM_COR/A12;FECHCONTA/D8", Condi$, "DISTINCT")
   '
10 Call SELECHO("ANTRAFON/Anulacion de Transferencias")
   VDEVU$ = VALACT1$("ANTRAFON")
   If VDEVU$ = "" Then Exit Sub
   '
   Condi$ = "TIPOMOVIM = 'TRAFON' AND CODICOM_COR = '" & VDEVU$ & "' and MarConsi = '*' "
   REGAF& = CantRegistros&("CAJABANC", Condi$)
   If REGAF& > 0 Then
      Call COMUNI("Imposible Anular Movimiento Conciliado.")
      GoTo 10
   End If
   '
   OPX% = COMALTER%("¿Está Seguro que Desea Anular la Transferencia " & VDEVU$ & "?\\No, Cancelar\Si, Anular")
   If OPX% = 2 Then
     Condi$ = "TIPOMOVIM = 'TRAFON' AND CODICOM_COR = '" & VDEVU$ & "'"
     Call BORRAREGISTROS("CAJABANC", Condi$, REGAF&, "NOMESS")
     Call COMUNI("Se Ha Anulado la Transferencia de Fondos " & VDEVU$)
   End If
   '
   Condi$ = "TIPOMOVIM = 'TRAFON'  ORDER BY CODICOM_COR DESC "
   Call CARGALISEL("ANTRAFON", "CAJABANC", "CODICOM_COR/A12;FECHCONTA/D8", Condi$, "DISTINCT")
   '
End Sub

Sub RECICHEQ()
   '
'\\\OT-5-602-FG-29/08/05///
   HOII% = HOY(HOS$)
   CUECHECA$ = Control$("", "CUECHECA")
   CUEICHEQ% = CUEINT%(CUECHECA$)
   If TRIM$(CUECHECA$) = "" Or CUEICHEQ% < 1 Then
      Call MENSERR(24, "Falta Definir Cuenta Contable\Asociada a Cartera de Valores")
      OPCHEQ07.Show 1
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
   BANCO$ = TRIM$(Mid$(OBX$, 3, 16))
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
   If TRIM$(BANCO$) = "" Then
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
   '\\\FIN VALIDACIONES
   '
   Screen.MousePointer = 11
   FEX = FEDEPO%
   FETX$ = FECHATEX$(FEX)
   NRECIB$ = "CH-" + Mid$(FETX$, 7, 2) + Mid$(FETX$, 4, 2) + Mid$(FETX$, 1, 2) + "-"
   NRECIBX$ = NRECIB$ + "__"
   ORMA% = (-1)
   Call APERBASDAT("CABAN")
   'With CajaBanc
   SQLX$ = "SELECT * FROM CajaBanc "
   SQLX$ = SQLX$ + " WHERE CodiCom_Cor LIKE '" & NRECIBX$ & "'"
   Dim CAJABTEMP As ADODB.Recordset
   Set CAJABTEMP = New ADODB.Recordset
25 On Error Resume Next
   CAJABTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0

   On Error Resume Next
   With CAJABTEMP
     .MoveFirst
     '.FindFirst "CodiCom_Cor LIKE '" & NRECIBX$ & "'"
     'If .NoMatch = False Then
     If Err < 1 Then
       On Error GoTo 0
       Do While Not .EOF
       '
       If Val(Mid$(!CodiCom_Cor, 11)) > ORMA% Then
          ORMA% = Val(Mid$(!CodiCom_Cor, 11))
       End If
       ' Validacion por Nº de Cheque del mismo Banco
       If TRIM$(Mid$(SAFETEXT$(!DESCRIMOV), 1, 18)) = UCase$(BANCO$) Then
           If TRIM$(Mid$(SAFETEXT$(!DESCRIMOV), 19, 14)) = UCase$(NUCHE$) Then
              Call COMUNI("El Numero de Cheque para ese Banco es Existente en la Cartera de Valores\Por favor ingresarlo nuevamente")
              GoTo 10
           End If
       End If
        '.FindNext "CodiCom_Cor LIKE '" & NRECIBX$ & "'"
        'GoTo 30
        .MoveNext
       Loop
     End If
     On Error GoTo 0
   End With
   Set CAJABTEMP = Nothing
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
   NumeClie = 0
   NumeProv% = 0
   TiMovi$ = "CHRECI"
   OpMovi$ = "CH"
   VaMovi$ = ""
     IDEVA$ = AJUSTI$(UCase$(BANCO$), 18)
     IDEVA$ = IDEVA$ + AJUSTI$(UCase$(NUCHE$), 14)
     IDEVA$ = IDEVA$ + FEVECHE$
   Descrip$ = IDEVA$
     If FECHE% < 1 Then FECHE% = FC%
   FechEmi = CVDAT(FEMISI%)
     If FEACRE% < FEMISI% Then FEACRE% = FEMISI%
     'If FEACRE% < FEDEPO% Then FEACRE% = FEDEPO%
   FeAcredi = CVDAT(FEACRE%)
   NumeTal = 0
'   CodiEmp% = 0
'   CodiSuc% = SUCUI%
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
   'IMPRESION DE FRM DE RECEPCION  DE CHEQUES
   FORMIMPRESION$ = Control("", "IMPRECHE")
   If FORMIMPRESION$ <> "" And EXISTE%(LUDAT$ + FORMIMPRESION$ + ".SQR") > 0 Then
      Call IMPRIRECECHEQUE(CoComLar$)
   End If
   '
   Screen.MousePointer = 1
   Call COMUNI("Se ha Procesado\la Recepción del Cheque / Valor")
   '
   GoTo 5
   '
'\\\OT-5-602-FG-FIN///
   '
End Sub

Sub IMPRIENTREVALO(DOCUME$)

    FILTX$ = TRIM$(DOCUME$)
    Call STDFORM("ENTVALO", FILTX$, "PRINT")


End Sub
Sub IMPRIRECECHEQUE(DOCUME$)
 
    FILTX$ = TRIM$(DOCUME$)
    Call STDFORM("VALSINRE", FILTX$, "PRINT")
    
 End Sub

Sub ENTRECHEQ()
   '
   HOII% = HOY(HOS$)
   CUECHECA$ = Control$("", "CUECHECA")
   CUEICHEQ% = CUEINT%(CUECHECA$)
   If TRIM$(CUECHECA$) = "" Or CUEICHEQ% < 1 Then
      Call MENSERR(24, "Falta Definir Cuenta Contable\Asociada a Cartera de Valores")
      OPCHEQ07.Show 1
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
   SELCHE07.List1.Clear
   SELCHE07.LICHESEL.Clear
   '
   Call APERBASDAT("CABAN")
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "where TipoMovim = 'CHRECI' "
   SQLX$ = SQLX$ + "AND Status = 1 "
   Dim CABATEMP As ADODB.Recordset
   Set CABATEMP = New ADODB.Recordset
25  On Error Resume Next
    CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0

   'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
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
         FEACRED% = CVINT(!fechacred)
         BANUFE$ = !DESCRIMOV
         ICHEQ# = !ValAbsComp
         '
         If TRIM$(Left$(BANUFE$, 12)) = "" Then
           Call REPLA(BANUFE$, "Cheque Nro:", 1, 11)
         End If
         RECHE& = !NumeComp
         '
         Z$ = "A" + Space$(80)
         Call REPLA(Z$, CSTRING$(MKI$(FEACRED%), 1, 5, 0, 2), 2, 5)
         Call REPLA(Z$, TRIM$(Left$(BANUFE$, 40)), 9, 40)
         Call REPLA(Z$, CSTRING$(MKD$(ICHEQ#), 4, 12, 2, 2), 49, 12)
         Call REPLA(Z$, CSTRING$(MKS$(RECHE&), 3, 8, 0, 2), 73, 8)
         SELCHE07.List1.AddItem Z$
         '
19     .MoveNext
       Loop
     End If
     On Error GoTo 0
   End With
   Set CABATEMP = Nothing
   '
   SUCHETO# = 0
   For i& = 1 To SELCHE07.List1.ListCount
      Z$ = Mid$(SELCHE07.List1.List(i& - 1), 9)
      SELCHE07.LICHESEL.AddItem Z$
      SUCHETO# = SUCHETO# + Val(Mid$(Z$, 41, 12))
   Next i&
   '
   SELCHE07.Label4.Caption = CSTRING$(MKD$(SUCHETO#), 4, 11, 2, 2)
   SELCHE07.Label4.Refresh
   '
   Screen.MousePointer = 1
   SELCHE07.Show 1
   DoEvents
   '
   IPASE# = Val(SELCHE07.Label2.Caption)
   If IPASE# > 0.005 Then
      FEX = FEDEPO%
      FETX$ = FECHATEX$(FEX)
      NRECIB$ = "EV-" + Mid$(FETX$, 7, 2) + Mid$(FETX$, 4, 2) + Mid$(FETX$, 1, 2) + "-"
      NRECIBX$ = NRECIB$ + "__"
      ORMA% = (-1)
      Call APERBASDAT("CABAN")
      'With CajaBanc
        On Error Resume Next
        '.FindFirst "CodiCom_Cor LIKE '" & NRECIBX$ & "'"
      SQLX$ = "SELECT * FROM CAJABANC"
      SQLX$ = SQLX$ + " WHERE CodiCom_Cor LIKE '" & NRECIBX$ & "'"
      Dim CAJABTEMP As ADODB.Recordset
      Set CAJABTEMP = New ADODB.Recordset
28    On Error Resume Next
      CAJABTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
      If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 28
      End If
      On Error GoTo 0

      On Error Resume Next
      With CAJABTEMP
          .MoveFirst
          'If .NoMatch = False Then
          If Err < 1 Then
            On Error GoTo 0
            Do While Not .EOF
              If Val(Mid$(!CodiCom_Cor, 11)) > ORMA% Then
                ORMA% = Val(Mid$(!CodiCom_Cor, 11))
              End If
              '.FindNext "CodiCom_Cor LIKE '" & NRECIBX$ & "'"
              'GoTo 30
              .MoveNext
            Loop
          End If
      End With
      Set CAJABTEMP = Nothing
      ORMX$ = TRIM$(Str$(ORMA% + 1))
      While Len(ORMX$) < 2: ORMX$ = "0" + ORMX$: Wend
      NRECIB$ = NRECIB$ + Left$(ORMX$, 2)
      '
      CoComCor$ = NRECIB$
      CoComLar$ = NRECIB$
'      CodiEmp% = 0
'      CodiSuc% = SUC%
      NumeClie = 0
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
      For i& = 1 To SELCHE07.LICHESEL.ListCount
        If SELCHE07.LICHESEL.Selected(i& - 1) = True Then
          X$ = SELCHE07.LICHESEL.List(i& - 1)
          BANUFE$ = TRIM$(Left$(X$, 40))
          BACH$ = TRIM$(Left$(X$, 16))
          NCHE$ = TRIM$(Mid$(X$, 19, 12))
          IPASE# = Val(Mid$(X$, 41, 12))
          SUMACHESEL# = SUMACHESEL# + IPASE#
          REGBA& = Val(Mid$(X$, 65, 8))
          '
          SQLX$ = "SELECT * FROM CAJABANC "
          SQLX$ = SQLX$ + "WHERE DescriMov = '" & BANUFE$ & "' "
          SQLX$ = SQLX$ + "AND ABS(ValAbsComp - " & FORMATNUM$(IPASE#, "F15.6") & ") < 0.005 "
          SQLX$ = SQLX$ + "AND NumeComp = " & REGBA& & " "
          SQLX$ = SQLX$ + "AND Status = 1"
          '
          Call APERBASDAT("CABAN")
          'With CajaBanc
            '.FindFirst SQLX$
          Set CAJABTEMP = New ADODB.Recordset
29        On Error Resume Next
          CAJABTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
          If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
             On Error GoTo 0
             Call CapturaErrorOpen
             GoTo 29
          End If
          On Error GoTo 0

          On Error Resume Next
          With CAJABTEMP
            .MoveFirst
            'If .NoMatch = False Then
            If Err < 1 Then
              On Error GoTo 0
              '.Edit
              !FEDEPOSENT = CVDAT(FEDEPO%)
              !Document = NRECIB$
              !DESTICHEQ = DESTICHE$
              CUECONTABINT% = !CUECONTAI
              !REFESALCHE = Left$(REFE$, 48)
              !IDENSALCHE = AJUSTI$(HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 48)
              !UsuaSalChe = AJUSTI$("Usuario: " + TRIM$(USERNAME$), 48)
              !Status = 2                     ' Marca Salida de Cartera
              !IdenBaja = IDACTU$             ' Marca Modificacion de registro - FECHA - HORA - USUARIO - TERMINAL
              .Update
            End If
            On Error GoTo 0
          End With
          Set CAJABTEMP = Nothing
          '
          CoComCor$ = NRECIB$
          CoComLar$ = NRECIB$
'          CodiEmp% = 0
'          CodiSuc% = SUC%
          NumeClie = 0
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
      Next i&
      '
      
      'IMPRESION DE FRM DE ENTREGA DE CHEQUES
      FORMIMPRESION$ = Control("", "IMPENTVA")
      If FORMIMPRESION$ <> "" And EXISTE%(LUDAT$ + FORMIMPRESION$ + ".SQR") > 0 Then
         Call IMPRIENTREVALO(CoComLar$)
      End If

      
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
   Call DESHASFECHA(DES%, Has%, PERIO$)
   If PERIO$ = "" Then GoTo 99
   '
   CORRESP$ = FECHATEX$(DES%) + " - " + FECHATEX$(Has%)
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
     Cuenta$ = TRIM$(Mid$(X$, 1, 12))
     CUEII% = CUECOINT%(Cuenta$) ' CUENTA CONTABLE DE TESORERIA
     '
     ' BUSQUEDA DE MOVIMIENTOS EN TABLA DETASIEN
     ' *****************************************
     '
     'FEINI = Int(CDbl(CVDAT(DES%)))
     'FEFIN = Int(CDbl(CVDAT(HAS%)))
     FEINI = DES%
     FEFIN = Has%
     FEINISQL = FETEXCOR1$(DES%) '**ESTOS 2 LOS CREE NUEVOS POR QUE 'FEINI' Y
     FEFINSQL = FETEXCOR1$(DIPOST(Has%)) '** 'FEFIN' SE UTILIZAN MAS ABAJO
     '
     SQLX$ = "SELECT * FROM DetaAsien WITH(NOLOCK)"
     SQLX$ = SQLX$ + "WHERE CueCoInt=" & CUEII% & " "
        SQLX$ = SQLX$ + " AND Fechasien >= '" & FEINISQL & "'"
        SQLX$ = SQLX$ + " AND Fechasien < '" & FEFINSQL & "'"
     SQLX$ = SQLX$ + " AND statpase >0 "
     SQLX$ = SQLX$ + "ORDER BY FECHASIEN, ORFEASIEN ASC;"
     Dim PASEASIEN As ADODB.Recordset
     Set PASEASIEN = New ADODB.Recordset
25   Set PASEASIEN = FLEXCONN.Execute(FILTSQL$(SQLX$))
     '
     JJ& = 0
     TRANSPOR# = 0
10   With PASEASIEN
       Do While Not .EOF
         IDEBE# = XVALO(!idebepase)
         IHABE# = XVALO(!ihabepase)
         If CVINT(!FECHASIEN) < FEINI Then
             TRANSPOR# = TRANSPOR# + IDEBE# - IHABE#
           ElseIf CVINT(!FECHASIEN) <= FEFIN Then
             Z$ = REBLA$
             Call REPLA(Z$, CODCUE$(XVALO(!CUECOINT)), 1, 12)
               fech% = CVINT%(!FECHASIEN)
             Call REPLA(Z$, MKI$(fech%), 13, 2)
             On Error Resume Next 'POR QUE SI ES NULL DA ERROR
             Call REPLA(Z$, SAFETEXT(!nuidasien), 15, 12)
             On Error GoTo 0
               REFEX8$ = SAFETEXT(!RefePase)
             Call REPLA(Z$, REFEX8$, 27, 48)
             Call REPLA(Z$, REFEX8$, 75, 48)
             Call REPLA(Z$, MKD$(IDEBE#), 123, 8)
             Call REPLA(Z$, MKD$(IHABE#), 139, 8)
             JJ& = JJ& + 1
             Call GRAREG("!REDIATES", Z$, JJ&)
         End If
         On Error GoTo 0
       .MoveNext
       Loop
     End With
     Set PASEASIEN = Nothing
     '
30   On Error GoTo 0
     
     ' BUSQUEDA DE MOVIMIENTOS EN TABLA CAJABANC
     ' *****************************************
     '
     EPAC$ = TRIM$(UCase$(EMPREAC$))
     '
     Call APERBASDAT("CABAN")
     SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK)"
     SQLX$ = SQLX$ + "where (CueContaI = " & CUEII% & " "
     SQLX$ = SQLX$ + " AND (NuAsidiar = '' or NuAsiDiar is null ))"
     SQLX$ = SQLX$ + " AND FechConta >= '" & FEINISQL & "'"
     SQLX$ = SQLX$ + " AND FechConta < '" & FEFINSQL & "'"
     SQLX$ = SQLX$ + " ORDER BY FECHCONTA, FEHOREAL ASC "
     '
     Dim CABATEMP As ADODB.Recordset
     Set CABATEMP = New ADODB.Recordset
32   Set CABATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
     '
     With CABATEMP
       Do While Not .EOF
         IDEBE# = XVALO(!idebecomp)
         IHABE# = XVALO(!ihabecomp)
         '
         If InStr(EPAC$, "TANIT") > 0 Then
           If CUEII% = CUEICARCH% Then
             If IDEBE# < 0.005 Then GoTo 17
           End If
         End If
         '
         If CVINT(!fechconta) < FEINI Then
             TRANSPOR# = TRANSPOR# + IDEBE# - IHABE#
           ElseIf CVINT(!fechconta) <= FEFIN Then
             Z$ = REBLA$
             Call REPLA(Z$, CODCUE(XVALO(!CUECONTAI)), 1, 12)
               fech% = CVINT%(!fechconta)
             Call REPLA(Z$, MKI$(fech%), 13, 2)
               COMPROB$ = SAFETEXT(!CodiCom_Cor)
             Call REPLA(Z$, COMPROB$, 15, 12)
               REFEX8$ = SAFETEXT(!refecom)
             Call REPLA(Z$, REFEX8$, 27, 48)
               DEPASE$ = SAFETEXT(!DESCRIMOV)
             Call REPLA(Z$, DEPASE$, 75, 48)
             Call REPLA(Z$, MKD$(IDEBE#), 123, 8)
             Call REPLA(Z$, MKD$(IHABE#), 139, 8)
             '
             JJ& = JJ& + 1
             Call GRAREG("!REDIATES", Z$, JJ&)
             '
             ' CASO PARTICULAR TANIT
             If InStr(EPAC$, "TANIT") > 0 Then
               If CUEII% = CUEICARCH% Then
                 IDEB1$ = TRIM$(FORMATNUM(IDEBE#, "F15.3"))
                 SQLX$ = "SELECT * FROM CAJABANC  WITH(NOLOCK)"
                 SQLX$ = SQLX$ + " WHERE DescriMov = '" & DEPASE$ & "'"
                 SQLX$ = SQLX$ + " AND IHabeComp > 0"
                 '
                 Dim CABATEMP2 As ADODB.Recordset
                 Set CABATEMP2 = New ADODB.Recordset
31               Set CABATEMP2 = FLEXCONN.Execute(FILTSQL$(SQLX$))
                 With CABATEMP2
                   If Not (.BOF And .EOF) Then
                       'FEDP% = CVINT(!FEDEPOSENT)
                       FEDP% = CVINT(!fechconta)
                       'FEDEPOI = CDbl(CVDAT(FEDP%))
                       If FEDP% < FEINI Or FEDP% > FEFIN Then GoTo 17
                       Z$ = REBLA$
                       IDEBE# = XVALO(!idebecomp)
                       IHABE# = XVALO(!ihabecomp)
                       'If Abs(IHABE# - XVALO(IDEB1$)) < 0.005 Then GoTo 17
                       Call REPLA(Z$, CODCUE(XVALO(CABATEMP2!CUECONTAI)), 1, 12)
                         fech% = CVINT%(!fechconta)
                       Call REPLA(Z$, MKI$(fech%), 13, 2)
                         COMPROB$ = SAFETEXT(!CodiCom_Cor)
                       Call REPLA(Z$, COMPROB$, 15, 12)
                         REFEX8$ = SAFETEXT(!refecom)
                       Call REPLA(Z$, REFEX8$, 27, 48)
                         DEPASE$ = SAFETEXT(!DESCRIMOV)
                       Call REPLA(Z$, DEPASE$, 75, 48)
                       Call REPLA(Z$, MKD$(IDEBE#), 123, 8)
                       Call REPLA(Z$, MKD$(IHABE#), 139, 8)
                         '
                       JJ& = JJ& + 1
                       Call GRAREG("!REDIATES", Z$, JJ&)
                       '
                   End If
                 End With
                 Set CABATEMP2 = Nothing
               End If
             End If
             ' FIN CASO PARTICULAR TANIT
             '
         End If
17     .MoveNext
       Loop
     End With
     Set CABATEMP = Nothing
     '
     ' MUEVE LOS MOVIMIENTOS ENCONTRADOS AL ARCHIVO CENTRAL
     ' ****************************************************
     '
40   CUETE$(U&) = CODCUE$(CUEII%)
     SINI#(U&) = 0
     '
     TRANSPOR# = SARRASCUE(CUEII%, DIANTE(DES%))
     If InStr(EPAC$, "TANIT") > 0 Then TRANSPOR# = 0
     '
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
            Call REPLA(X$, MKD$(Abs(TRANSPOR#)), 139, 8)
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


Sub CONCITCRE()

    Call COPYSTRU("MECOBRA", "MECOBTC")
    F& = 0
    For U& = 1 To ULTREG&("MECOBRA")
        XX$ = REGLEIDO$("MECOBRA", U&)
        If Mid$(XX$, 3, 2) = "TC" Then
            F& = F& + 1
            Call GRAREG("MECOBTC", XX$, F&)
        End If
     Next U&
     Call SETULTREG("MECOBTC", F&)
     Call CIERRARCH("MECOBTC")
     Call FRESHECHO("MECOBTC")
     
     CONCITC07.Show 1
     
End Sub

