VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form RECIBO08 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recibo de Cobranza en Cuenta Corriente"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11805
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11805
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox LINDICE 
      Height          =   1035
      Left            =   1260
      Sorted          =   -1  'True
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   12000
      Begin VB.PictureBox Picture999 
         Height          =   540
         Index           =   0
         Left            =   2280
         ScaleHeight     =   480
         ScaleWidth      =   1110
         TabIndex        =   86
         Top             =   420
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   0
         Top             =   0
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E2D3C0&
         Caption         =   "NETO GRAVADO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   210
         TabIndex        =   58
         Top             =   7190
         Width           =   1815
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   180
            TabIndex        =   59
            Top             =   315
            Width           =   1485
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00E2D3C0&
         Caption         =   "APLICADOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   6960
         TabIndex        =   56
         Top             =   7190
         Width           =   1815
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   180
            TabIndex        =   57
            Top             =   315
            Width           =   1485
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00E2D3C0&
         Caption         =   "ANTICIPOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   8880
         TabIndex        =   54
         Top             =   7190
         Width           =   1815
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   180
            TabIndex        =   55
            Top             =   315
            Width           =   1485
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00E2D3C0&
         Caption         =   "SALDO CANCELAR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   2100
         TabIndex        =   50
         Top             =   7190
         Width           =   1830
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   160
            TabIndex        =   51
            Top             =   315
            Width           =   1485
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00E2D3C0&
         Caption         =   "PROM. DE PAGOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1250
         Left            =   6195
         TabIndex        =   43
         Top             =   105
         Width           =   1665
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   1050
            TabIndex        =   49
            Top             =   315
            Width           =   540
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   1050
            TabIndex        =   48
            Top             =   630
            Width           =   540
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   1050
            TabIndex        =   47
            Top             =   945
            Width           =   540
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2D3C0&
            Caption         =   "Prom. Venc"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   185
            Left            =   100
            TabIndex        =   46
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2D3C0&
            Caption         =   "Prom. Val"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   185
            Left            =   100
            TabIndex        =   45
            Top             =   670
            Width           =   855
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2D3C0&
            Caption         =   "Dias Atraso"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   185
            Left            =   100
            TabIndex        =   44
            Top             =   980
            Width           =   855
         End
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
         Left            =   5880
         TabIndex        =   39
         Top             =   3950
         Width           =   1275
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2355
         Left            =   240
         ScaleHeight     =   2325
         ScaleWidth      =   10410
         TabIndex        =   38
         Top             =   1440
         Width           =   10440
         Begin VB.Frame Frame17 
            Caption         =   "Frame6"
            Height          =   2900
            Left            =   10000
            TabIndex        =   72
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame16 
            Caption         =   "Frame6"
            Height          =   2900
            Left            =   8500
            TabIndex        =   71
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame15 
            Caption         =   "Frame6"
            Height          =   2900
            Left            =   7000
            TabIndex        =   70
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame14 
            Caption         =   "Frame6"
            Height          =   2900
            Left            =   5700
            TabIndex        =   69
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame13 
            Caption         =   "Frame6"
            Height          =   2900
            Left            =   5000
            TabIndex        =   68
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame6 
            Caption         =   "Frame6"
            Height          =   2900
            Left            =   4100
            TabIndex        =   67
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame4 
            Caption         =   "Frame6"
            Height          =   2900
            Left            =   2600
            TabIndex        =   66
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame3 
            Caption         =   "Frame6"
            Height          =   2900
            Left            =   600
            TabIndex        =   65
            Top             =   0
            Width           =   15
         End
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2280
            ItemData        =   "RECIBO08.frx":0000
            Left            =   -10
            List            =   "RECIBO08.frx":0002
            TabIndex        =   64
            Top             =   320
            Width           =   10450
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   $"RECIBO08.frx":0004
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
            Height          =   200
            Left            =   0
            TabIndex        =   89
            Top             =   70
            Width           =   1.04500e5
         End
         Begin VB.Label Label36 
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
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   0
            TabIndex        =   92
            Top             =   0
            Width           =   1.04505e5
         End
      End
      Begin VB.CommandButton Command23 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   11000
         Picture         =   "RECIBO08.frx":00A8
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Ver e Imprimir Copia de Recibo ya Emitido"
         Top             =   6480
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
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
         Left            =   11000
         Picture         =   "RECIBO08.frx":03B2
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Aprueba, Graba e Imprime Documento"
         Top             =   7270
         Width           =   650
      End
      Begin VB.TextBox Text2 
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
         Left            =   1470
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "TBOX"
         Text            =   " "
         Top             =   265
         Width           =   3060
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   5160
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "X"
         ToolTipText     =   "Configuracion de Funcionamiento - Doble-Click Aquí"
         Top             =   265
         Width           =   435
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00624E28&
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
         Height          =   6225
         Left            =   10800
         ScaleHeight     =   6195
         ScaleWidth      =   1035
         TabIndex        =   23
         Top             =   0
         Width           =   1065
         Begin VB.PictureBox MARCOBARRA 
            Appearance      =   0  'Flat
            BackColor       =   &H002B9973&
            ForeColor       =   &H80000008&
            Height          =   200
            Left            =   170
            ScaleHeight     =   165
            ScaleWidth      =   615
            TabIndex        =   96
            Top             =   5880
            Width           =   650
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H0048CC9D&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   -120
               TabIndex        =   97
               Top             =   0
               Width           =   12000
            End
         End
         Begin VB.CommandButton Command24 
            Caption         =   "B.Ch."
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
            Left            =   170
            Picture         =   "RECIBO08.frx":06BC
            Style           =   1  'Graphical
            TabIndex        =   61
            ToolTipText     =   "Busqueda de Cheque x Número"
            Top             =   4260
            Visible         =   0   'False
            Width           =   650
         End
         Begin VB.CommandButton Command2 
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
            Left            =   170
            Picture         =   "RECIBO08.frx":0806
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Abandona la Aplicación"
            Top             =   120
            Width           =   650
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Notas"
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
            Left            =   170
            Picture         =   "RECIBO08.frx":0950
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Notas de Clientes"
            Top             =   3570
            Visible         =   0   'False
            Width           =   650
         End
         Begin VB.CommandButton Command4 
            Caption         =   "C.Cte."
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
            Left            =   170
            Picture         =   "RECIBO08.frx":0C5A
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Resumen de Cuenta Corriente"
            Top             =   1500
            Width           =   650
         End
         Begin VB.CommandButton Command11 
            Caption         =   "New"
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
            Left            =   170
            Picture         =   "RECIBO08.frx":0F64
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Apertura de Nueva Cuenta de Cliente"
            Top             =   2880
            Visible         =   0   'False
            Width           =   650
         End
         Begin VB.CommandButton Command6 
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
            Left            =   170
            Picture         =   "RECIBO08.frx":13A6
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Configuración de Funcionamiento"
            Top             =   4950
            Visible         =   0   'False
            Width           =   650
         End
         Begin VB.CommandButton BOTNEXT 
            Caption         =   "Next"
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
            Left            =   170
            Picture         =   "RECIBO08.frx":17E8
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Continúa Carga de Recibo."
            Top             =   810
            Width           =   650
         End
         Begin VB.CommandButton Command8 
            Caption         =   "More"
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
            Left            =   170
            Picture         =   "RECIBO08.frx":1AF2
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Anular, Re-Imprimir, Imputación de Créditos, Tablas"
            Top             =   2190
            Width           =   650
         End
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
         Left            =   10500
         TabIndex        =   21
         Top             =   630
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2800
         Left            =   240
         ScaleHeight     =   2775
         ScaleWidth      =   10410
         TabIndex        =   16
         Top             =   4300
         Width           =   10440
         Begin VB.Frame Frame19 
            Caption         =   "Frame6"
            Height          =   2800
            Left            =   3720
            TabIndex        =   91
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame23 
            Caption         =   "Frame6"
            Height          =   2400
            Left            =   10100
            TabIndex        =   85
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame22 
            Caption         =   "Frame6"
            Height          =   2400
            Left            =   8500
            TabIndex        =   84
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame21 
            Caption         =   "Frame6"
            Height          =   2400
            Left            =   6600
            TabIndex        =   83
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame20 
            Caption         =   "Frame6"
            Height          =   2400
            Left            =   5400
            TabIndex        =   82
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame24 
            Caption         =   "Frame6"
            Height          =   2400
            Left            =   2500
            TabIndex        =   81
            Top             =   0
            Width           =   15
         End
         Begin VB.ListBox List2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            ItemData        =   "RECIBO08.frx":1C3C
            Left            =   0
            List            =   "RECIBO08.frx":1C3E
            TabIndex        =   80
            Top             =   330
            Width           =   10440
         End
         Begin VB.ListBox List3 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1860
            Left            =   0
            TabIndex        =   79
            Top             =   360
            Width           =   10450
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   5630
            Left            =   0
            TabIndex        =   73
            Top             =   360
            Width           =   10515
            Begin VB.Line Line15 
               X1              =   0
               X2              =   11295
               Y1              =   2415
               Y2              =   2415
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Total General"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   45
               TabIndex        =   78
               Top             =   2130
               Width           =   2055
            End
            Begin VB.Label Label8 
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
               Height          =   330
               Left            =   6800
               TabIndex        =   77
               Top             =   2080
               Width           =   1700
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   0
               TabIndex        =   76
               Top             =   120
               Visible         =   0   'False
               Width           =   1695
            End
            Begin VB.Label Label10 
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
               Height          =   330
               Left            =   8500
               TabIndex        =   75
               Top             =   2080
               Width           =   1610
            End
            Begin VB.Label Label4 
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
               Height          =   330
               Left            =   3730
               TabIndex        =   74
               Top             =   2080
               Width           =   1680
            End
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   $"RECIBO08.frx":1C40
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
            Height          =   200
            Left            =   0
            TabIndex        =   90
            Top             =   90
            Width           =   10410
         End
         Begin VB.Label Label37 
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
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   0
            TabIndex        =   93
            Top             =   0
            Width           =   10410
         End
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
         Left            =   9030
         TabIndex        =   1
         Text            =   " "
         Top             =   1000
         Width           =   1650
      End
      Begin VB.TextBox Text9 
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
         Left            =   9345
         TabIndex        =   3
         Text            =   " "
         Top             =   265
         Width           =   1170
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
         Left            =   10500
         TabIndex        =   4
         Top             =   265
         Width           =   180
      End
      Begin VB.TextBox Text5 
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
         Left            =   2310
         TabIndex        =   10
         TabStop         =   0   'False
         Tag             =   "TBOX"
         Text            =   " "
         Top             =   1000
         Width           =   2220
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   1470
         TabIndex        =   9
         TabStop         =   0   'False
         Tag             =   "TBOX"
         Text            =   " "
         Top             =   1000
         Width           =   750
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
         Height          =   300
         Left            =   1470
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   "TBOX"
         Text            =   " "
         Top             =   630
         Width           =   3060
      End
      Begin VB.CommandButton Command1 
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
         Left            =   1050
         TabIndex        =   2
         Top             =   630
         Width           =   180
      End
      Begin VB.TextBox Text1 
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
         Left            =   240
         TabIndex        =   0
         Text            =   " "
         Top             =   630
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid GRID2 
         Height          =   765
         Left            =   240
         TabIndex        =   62
         Top             =   1920
         Visible         =   0   'False
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   1349
         _Version        =   393216
         Rows            =   23
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   8433840
         BackColorSel    =   -2147483643
         ForeColorSel    =   12582912
         BackColorBkg    =   16777215
         Redraw          =   -1  'True
         FocusRect       =   2
         AllowUserResizing=   1
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
      Begin MSFlexGridLib.MSFlexGrid GRID1 
         Height          =   765
         Left            =   3000
         TabIndex        =   63
         Top             =   1920
         Visible         =   0   'False
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   1349
         _Version        =   393216
         Rows            =   23
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   8433840
         BackColorSel    =   -2147483643
         ForeColorSel    =   12582912
         BackColorBkg    =   16777215
         Redraw          =   -1  'True
         FocusRect       =   2
         AllowUserResizing=   1
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
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
         Height          =   435
         Index           =   0
         Left            =   60
         OleObjectBlob   =   "RECIBO08.frx":1CEE
         TabIndex        =   87
         Top             =   0
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Descuento (N./Cred)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   4035
         TabIndex        =   52
         Top             =   7190
         Visible         =   0   'False
         Width           =   2820
         Begin VB.TextBox DESCUEN 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   1200
            TabIndex        =   95
            Top             =   315
            Width           =   1515
         End
         Begin VB.TextBox Text7 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            TabIndex        =   94
            Top             =   315
            Width           =   675
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   840
            TabIndex        =   53
            Top             =   300
            Width           =   330
         End
      End
      Begin VB.Label Label35 
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
         Height          =   255
         Left            =   240
         TabIndex        =   88
         Top             =   3960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label34 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   7875
         TabIndex        =   60
         Top             =   210
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
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
         Left            =   8760
         TabIndex        =   42
         Top             =   3950
         Width           =   1520
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ajuste de Centavos"
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
         Left            =   3885
         TabIndex        =   41
         ToolTipText     =   "Doble Click Para Ajuste Automatico"
         Top             =   4050
         Width           =   1905
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Valores"
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
         Left            =   7400
         TabIndex        =   40
         Top             =   4050
         Width           =   1275
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Cambio"
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
         Left            =   10850
         TabIndex        =   36
         ToolTipText     =   "Doble Click para Configurar Cliente"
         Top             =   6300
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label MONEMIS 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pesos"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   11.25
            Charset         =   0
            Weight          =   900
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   10815
         TabIndex        =   35
         ToolTipText     =   "T.C.: 1.0000 - Doble-Click para Cambiar Moneda"
         Top             =   6810
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label13 
         BorderStyle     =   1  'Fixed Single
         Height          =   120
         Left            =   3360
         TabIndex        =   34
         Top             =   7980
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label TICAMBIS 
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
         Height          =   280
         Left            =   10815
         TabIndex        =   33
         ToolTipText     =   "Doble Click para Ajustar Tipo de Cambio"
         Top             =   6500
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label26 
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
         Left            =   240
         TabIndex        =   32
         Top             =   1000
         Width           =   980
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Documento No Válido Como Factura"
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
         Left            =   4560
         TabIndex        =   15
         ToolTipText     =   "Configuracion de Funcionamiento - Doble-Click Aquí"
         Top             =   1005
         Width           =   1680
      End
      Begin VB.Line Line30 
         X1              =   0
         X2              =   12000
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label22 
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
         Left            =   8520
         TabIndex        =   20
         Top             =   600
         Width           =   1995
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Talo- nario"
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
         Left            =   7875
         TabIndex        =   19
         Top             =   585
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3780
         TabIndex        =   18
         Top             =   5775
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Valores:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1470
         TabIndex        =   17
         Top             =   5880
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.Label Label6 
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
         Height          =   210
         Left            =   8040
         TabIndex        =   13
         Top             =   1110
         Width           =   900
      End
      Begin VB.Label Label3 
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
         Height          =   210
         Left            =   8760
         TabIndex        =   12
         Top             =   360
         Width           =   750
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   250
         TabIndex        =   6
         ToolTipText     =   "Doble - Click para visualzar aplicación de Créditos y Anticipos"
         Top             =   390
         Width           =   960
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "RECIBO08.frx":1D68
      Top             =   1680
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1050
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Pos.IVA:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Menu mnuTransacciones 
      Caption         =   "Transacciones"
      Begin VB.Menu ReciCobran 
         Caption         =   "Nuevo Recibo de Cobranza"
      End
      Begin VB.Menu mnuAnuRec 
         Caption         =   "Anulación Recibo de Cobranza"
      End
      Begin VB.Menu separa1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReimputacionContable 
         Caption         =   "Reimputacion Contable de Recibos"
      End
      Begin VB.Menu mnuImpCre 
         Caption         =   "Imputación  de Créditos"
      End
      Begin VB.Menu separa2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnuConsuList 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu mnuCuentaCorriente 
         Caption         =   "Resumen de Cuenta Corriente"
      End
      Begin VB.Menu separa6 
         Caption         =   "-"
      End
      Begin VB.Menu LiReReg 
         Caption         =   "Listado de Recibos Registrados"
      End
      Begin VB.Menu mnuRecibosConDetalleDeMedios 
         Caption         =   "Listado de Recibos con Detalle de Medios"
      End
      Begin VB.Menu mnuEstadisticasCobranza 
         Caption         =   "Estadisticas de Cobranza"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuApliCredAnt 
         Caption         =   "Aplicación de Créditos y Anticipos"
      End
      Begin VB.Menu mnuDiferenciaCambio 
         Caption         =   "Listado Aplicación Diferencias de Cambio"
      End
      Begin VB.Menu mnuAplicaVentasConCobranza 
         Caption         =   "Aplicaciones de Ventas y Cobranza con Detalle de Medios"
      End
      Begin VB.Menu mnuBusquedaCheque 
         Caption         =   "Búsqueda de Cheque por Número"
      End
      Begin VB.Menu separa3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReportesVarios 
         Caption         =   "Reportes Varios"
      End
   End
   Begin VB.Menu mnuEdi 
      Caption         =   "Ver"
      Begin VB.Menu mnuVerReimp 
         Caption         =   "Ver Recibos y Re-Imprimir"
      End
   End
   Begin VB.Menu mnuUtilidades 
      Caption         =   "Utilidades"
      Begin VB.Menu mnuAjuSalDeu 
         Caption         =   "Ajuste de Saldos Deudores"
      End
      Begin VB.Menu mnuCompensaciones 
         Caption         =   "Compensación de Cliente-Proveedor"
      End
   End
   Begin VB.Menu mnuConfi 
      Caption         =   "Configuración"
      Begin VB.Menu mnuAsociarReporte 
         Caption         =   "Vincular Reportes al Formulario"
      End
      Begin VB.Menu separa4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTalonariosRecibos 
         Caption         =   "Talonario de Recibos"
      End
      Begin VB.Menu mnuMediosCobranza 
         Caption         =   "Tabla de Medios de Cobranza"
      End
      Begin VB.Menu mnuretencionIIBBProvincia 
         Caption         =   "Retenciones de Ing. Brutos por Provincia"
      End
      Begin VB.Menu mnuCotizacionesMoneda 
         Caption         =   "Cotizaciones de Moneda Extranjera"
      End
      Begin VB.Menu separa5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConfiGral 
         Caption         =   "Configuración General"
      End
   End
   Begin VB.Menu mnuAccDir 
      Caption         =   "Accesos Directos"
      Begin VB.Menu mnuMaeClie 
         Caption         =   "Maestro de Clientes"
      End
      Begin VB.Menu mnuNoCli 
         Caption         =   "Notas de Clientes"
      End
      Begin VB.Menu mnuPlanDeCuentas 
         Caption         =   "Plan de Cuentas Contables"
      End
      Begin VB.Menu calculodefcambio 
         Caption         =   "Modulo de Cálculo de Dif. de Cambio x Nota de Débito"
      End
   End
End
Attribute VB_Name = "RECIBO08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LIST2ANT$, AUTONUME%
Dim MODORECI%
Dim MODOMONEI%
Dim ctrl As Object
Dim CtaRetIva$


Sub CARGA_ENCABEZADO_GRILLA_MESES()
  With INFO_GRILLA09
  
    .MSF.Row = 1
    .MSF.Cols = 21
    NUCOLU% = 1: .TEPRUEBA = "90": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = ""
    
    NUCOLU% = 2: .TEPRUEBA = "1234567890123456789023456789012345": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 1: .MSF.TextMatrix(0, NUCOLU% - 1) = "Cliente"
    NUCOLU% = 3: .TEPRUEBA = "1234567890234567890": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 1: .MSF.TextMatrix(0, NUCOLU% - 1) = "Cond.Pago"
    NUCOLU% = 4: .TEPRUEBA = "123456": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Plazo"
    '
'    NUCOLU% = 5: .TEPRUEBA = "12345678": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Sem.1"
'    NUCOLU% = 6: .TEPRUEBA = "1234567890123456": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Cob. $$"
'    NUCOLU% = 7: .TEPRUEBA = "1234567890123456": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Num"
'    '
'    NUCOLU% = 8: .TEPRUEBA = "12345678": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Sem.2"
'    NUCOLU% = 9: .TEPRUEBA = "1234567890123456": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Cob. $$"
'    NUCOLU% = 10: .TEPRUEBA = "1234567890123456": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Num"
'    '
'    NUCOLU% = 11: .TEPRUEBA = "12345678": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Sem.3"
'    NUCOLU% = 12: .TEPRUEBA = "1234567890123456": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Cob. $$"
'    NUCOLU% = 13: .TEPRUEBA = "1234567890123456": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Num"
'    '
'    NUCOLU% = 14: .TEPRUEBA = "12345678": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Sem.4"
'    NUCOLU% = 15: .TEPRUEBA = "1234567890123456": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Cob. $$"
'    NUCOLU% = 16: .TEPRUEBA = "1234567890123456": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Num"
'    '
'    NUCOLU% = 17: .TEPRUEBA = "12345678": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Sem.5"
'    NUCOLU% = 18: .TEPRUEBA = "1234567890123456": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Cob. $$"
'    NUCOLU% = 19: .TEPRUEBA = "1234567890123456": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Num"
'    '
'    NUCOLU% = 20: .TEPRUEBA = "1234567890123456": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "P.Pond."
'    NUCOLU% = 21: .TEPRUEBA = "12345678": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Desvío"
  End With
    '
End Sub

Sub CARGA_ENCABEZADO_GRILLA()
  With INFO_GRILLA07
  
    .MSF.Row = 1
    .MSF.Cols = 21
    NUCOLU% = 1: .TEPRUEBA = "90": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = ""
    
    NUCOLU% = 2: .TEPRUEBA = "1234567890123456789023456789012345": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 1: .MSF.TextMatrix(0, NUCOLU% - 1) = "Cliente"
    NUCOLU% = 3: .TEPRUEBA = "1234567890234567890": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 1: .MSF.TextMatrix(0, NUCOLU% - 1) = "Cond.Pago"
    NUCOLU% = 4: .TEPRUEBA = "123456": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Plazo"
    '
    NUCOLU% = 5: .TEPRUEBA = "12345678": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Sem.1"
    NUCOLU% = 6: .TEPRUEBA = "1234567890123456": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Cob. $$"
    NUCOLU% = 7: .TEPRUEBA = "1234567890123456": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Num"
    '
    NUCOLU% = 8: .TEPRUEBA = "12345678": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Sem.2"
    NUCOLU% = 9: .TEPRUEBA = "1234567890123456": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Cob. $$"
    NUCOLU% = 10: .TEPRUEBA = "1234567890123456": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Num"
    '
    NUCOLU% = 11: .TEPRUEBA = "12345678": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Sem.3"
    NUCOLU% = 12: .TEPRUEBA = "1234567890123456": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Cob. $$"
    NUCOLU% = 13: .TEPRUEBA = "1234567890123456": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Num"
    '
    NUCOLU% = 14: .TEPRUEBA = "12345678": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Sem.4"
    NUCOLU% = 15: .TEPRUEBA = "1234567890123456": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Cob. $$"
    NUCOLU% = 16: .TEPRUEBA = "1234567890123456": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Num"
    '
    NUCOLU% = 17: .TEPRUEBA = "12345678": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Sem.5"
    NUCOLU% = 18: .TEPRUEBA = "1234567890123456": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Cob. $$"
    NUCOLU% = 19: .TEPRUEBA = "1234567890123456": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Num"
    '
    NUCOLU% = 20: .TEPRUEBA = "1234567890123456": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "P.Pond."
    NUCOLU% = 21: .TEPRUEBA = "12345678": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Desvío"
  End With
    '
End Sub

'
Sub ClearControls()
    For Each ctrl In Me.Controls
        If ctrl.Tag = "TBOX" Then ctrl.TEXT = ""
    Next
    Label26 = ""
    Label22 = ""
    
End Sub

Sub DIASDECOBRANZA_VENCIFACTURA()
        ' CALCULO DE DIAS DE COBRANZA A PARTIR DE LA FECHA DE VENCIMIENTO
        ' DE LAS FACTURAS CANCELADAS
        '
        ' CALCULO DE DIAS DE COBRANZA
        JJ% = 0: ACRED# = 0: TOTALVAL# = 0
        FechaRecibo% = FECHANUM(Text9)
        For II% = 0 To List1.ListCount - 1
             ' VERIFICA QUE TIPO DE MEDIO ES. PARA ESO LEE EL CODIGO DE MEDIO
             ' (LA ULTIMA COLUMNA DE LA TABLA DE MEDIOS DE COBRANZA)
             ' OBVIA A LAS TRANSFERENCIAS (TR)
             MPAGO% = XVALO(Mid$(List1.List(II%), 1, 4))
             RGMP$ = FILTERGETRECORD$("MECOBRA", 1, MKI$(MPAGO%))
             CODMP$ = TRIM$(Mid$(RGMP$, 57, 4))
             '
             FVENCHE% = FECHANUM(Mid$(List1.List(II%), 62, 8))   ' fecha de vencimiento del cheque
             If FVENCHE% < FechaRecibo% And CODMP$ <> "TR" Then FVENCHE% = FechaRecibo%
             IMPOCHE# = XVALO(Mid$(List1.List(II%), 70, 14))   ' importe del cheque
             
5            If ACRED# < 0.005 Then     ' SI HAY IMPORTE DE FACTURA PENDIENTE
10               If JJ% < List2.ListCount Then ' RECORRE LOS COMPROBANTES DE LA COMPOSICION DE DEUDA
                     REGFACTURA$ = List2.List(JJ%)
                     FVENFAC% = FECHANUM(Mid$(REGFACTURA$, 23, 8)) ' FECHA DE EMISION DE LA FACTURA
                     ACRED# = XVALO(Mid$(REGFACTURA$, 70, 14))
                     If ACRED# < 0.005 Then    ' VERIFICA SI EL COMPROBANTE FUE ELEGIDO
                         JJ% = JJ% + 1
                         GoTo 10
                     End If
                 Else ' SI NO HAY MAS COMPROBANTES AGREGA EL SALDO DEL ULTIMO CHEQUE CON PONDERACION 1 (UNO)
                     DIFEDIA% = 1
                     TOTALVAL# = TOTALVAL# + (DIFEDIA% * IMPOCHE#)
                     Exit For
                 End If
             End If
                 
             If IMPOCHE# >= ACRED# Then ' EL IMPORTE DEL CHEQUE ES MAYOR AL DE LA FACTURA
                 IMPOCHE# = IMPOCHE# - ACRED#
                 '
                 DIFEDIA% = DIENTRE(FVENFAC%, FVENCHE%)
                 If ACRED# < 0 Then DIFEDIA% = 1 ' SI EL COMPROBANTE ES UN CREDITO, DIFEDIA% = 1
                 '
                 TOTALVAL# = TOTALVAL# + (DIFEDIA% * ACRED#)
                 ACRED# = 0: JJ% = JJ% + 1
             Else                       ' EL IMPORTE DE LA FACTURA ES MAYOR AL DEL CHEQUE
                 DIFEDIA% = DIENTRE(FVENFAC%, FVENCHE%)
                 TOTALVAL# = TOTALVAL# + (DIFEDIA% * IMPOCHE#)
                 ACRED# = ACRED# - IMPOCHE#
                 IMPOCHE# = 0
             End If
             
             If IMPOCHE# > 0.005 Then GoTo 5 ' SI TODAVIA HAY IMPORTE DEL CHEQUE, BUSCA UN NUEVO COMPROBANTE
        Next II%
        '
        Label35 = FORMATNUM(0, "I1")
        TOTAVALO# = XVALO(Label14) ' TOTAL DE VALORES RECIBIDOS
        If TOTAVALO# > 0.005 Then Label35 = FORMATNUM((TOTALVAL# / TOTAVALO#), "I5") ' CALCULO DE DIAS PONDERADO
        '
End Sub

Function PROMEDIO_CONDI_PAGOS(CONDIPAG%)
      '
      PROMEDIO_CONDI_PAGOS = 0
      CANTI_CONDICIONES% = 0
      '
      'LEO EL REGISTRO SEGUN CONDICION DE PAGO Y TOMO LAS CONDICIONES
      'SUMO CUANTAS CONDICIONES TIENE PARA LUEGO  PODER CALCULAR EL PROMEDIO
      '
      x$ = FILTERGETRECORD$("CONPAG", 1, MKI$(CONDIPAG%))
      P1% = Asc(Mid$(x$, 38, 1)): If P1% > 0 Then CANTI_CONDICIONES% = CANTI_CONDICIONES% + 1
      P2% = Asc(Mid$(x$, 39, 1)): If P2% > 0 Then CANTI_CONDICIONES% = CANTI_CONDICIONES% + 1
      P3% = Asc(Mid$(x$, 40, 1)): If P3% > 0 Then CANTI_CONDICIONES% = CANTI_CONDICIONES% + 1
      P4% = Asc(Mid$(x$, 41, 1)): If P4% > 0 Then CANTI_CONDICIONES% = CANTI_CONDICIONES% + 1
      P5% = Asc(Mid$(x$, 42, 1)): If P5% > 0 Then CANTI_CONDICIONES% = CANTI_CONDICIONES% + 1
      P6% = Asc(Mid$(x$, 43, 1)): If P6% > 0 Then CANTI_CONDICIONES% = CANTI_CONDICIONES% + 1
      '
      'SUMO LOS PLAZOS DE PAGO
      SUMACONDI = P1% + P2% + P3% + P4% + P5% + P6%
      '
      'SI NO TIENE PLAZOS DEVUELVE 0 (CONTADO)
      If SUMACONDI <= 0 Then Exit Function
      '
      If CANTI_CONDICIONES% < 1 Then CANTI_CONDICIONES% = 1
      PROMEDIO_SUMACONDI = SUMACONDI / CANTI_CONDICIONES%
      '
      PROMEDIO_CONDI_PAGOS = PROMEDIO_SUMACONDI
End Function

Private Sub BOTNEXT_Click()
    '
    NRONUE& = XVALO(Mid$(Text11, 4))
    If NRONUE& < 1 Or NRONUE& > 999999 Then
      Call MENSERR(24, "Número de Recibo Fuera de Rango")
      Call NURECIBO(NROPRO&)
      Call PRENURE(NROPRO&)
      Exit Sub
    End If
    '
    FECHALIM% = XVALO(Control$("", "CIEMEVEN"))
    If FECHANUM(Text9) <= FECHALIM% Then
      Call MENSERR(24, "Fecha no Válida.\  \Ese Período Mensual Está Cerrado.")
      Exit Sub
    End If
    '
    Call GRACONTROL("RECISTA", "COBRACTI", TRIM$(Label22.Caption))
    Call SIGUERECI1
    Command5.Enabled = True
    '
End Sub



Private Sub calculodefcambio_Click()
     EPAC$ = TRIM$(UCase$(EMPREAC$))
     If InStr(EPAC$, "AHP") > 0 Then
        SOLNOTADEB.Show 1
     Else
        Call MENSERR(24, "No Disponible por el Momento")
     End If

End Sub

Private Sub Command1_Click()
    '
    'Call SELECHO("DEUDOR1")
    'NCLI = XVALO(TRIM$(VALACT1$("DEUDOR1")))
    
    NCLI = ListaSeleccionClientes()
    If NCLI > 0 Then
        NCLIE = NCLI
       '\\\OT-06-0607-OD-28/07/08///
       'EVALUA SI LA CUENTA NO ESTA SUSPENDIDA
       '
       If LIMICRE(NCLI) < (-0.5) Then
          Call BorraEncabezado
          Call MENSERR(24, "Cuenta Suspendida.")
          Exit Sub
       End If
       '\\\OT-06-0607-OD-FIN///
        Text1.TEXT = TRIM$(Str$(NCLIE))
        Text2.TEXT = rasocli$(NCLIE)
        Text3.TEXT = DOMICLI$(NCLIE)
        Text4.TEXT = CPOSCLI$(NCLIE)
        Text5.TEXT = LOCACLI$(NCLIE)
        Label26 = CODICLI$(NCLIE)
        Call NotaCli(NCLIE, 1)
    End If
    '
End Sub

Sub BorraEncabezado()
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Label26 = ""
End Sub

Sub ESTADISTICADECOBRANZA()
'COLUMNA 1 RAZON SOCIAL
'2 CONDICION DE PAGO FORMATO TEXTO
'3 CONDICION DE PAGOS (PROMEDIO DE DIAS)
'4 CANTIDAD DE DIAS DE COBRANZA PONDERADO(CAMPO ATRACHEQUE) PROMEDIO (SE REPITE PARA CADA SEMANA COLUMNA 4,7,10,13,16)
'5 ACUMULADO DEL IMPORTE DE LOS RECIBOS POR CLIENTE                  (SE REPITE PARA CADA SEMANA)COLUMNA 5,8,11,14,17)
'6 VALOR DE COLUMNA 4 X VALOR DE LA COLUMNA 5                        (SE REPITE PARA CADA SEMANA)COLUMNA 6,9,12,15,18)
'COLUMNA 19 SUMA DE LAS COLUMNAS 6 DIVIDIDO LA SUMA  DE LAS COLUMNA 5
'COLUMNA 20 PROMEDIO DE LAS OCOLUMNAS 4  O SEA LA SUMA DE LA COLUMNA 4 DIVIDIDO LA CANTIDAD DE SEMANAS
'
'TOTALES A PIE DE GRILLA.

'TOTALES PROMEDIO DE LA COLUNA 4 (PROMEDIO DE ATRAZO)
'SUMA DE LA COLUMNA 5 (IMPORTES DE COBRANZA)
'SUMA DE LA COLUMNA 6 (NUM)

'RESULTADO DE LA FORMULA: (COLUMNA 3 * 100) / COLUMNA 2
  '**************** FUNCIONA CON CONFIGURACION POR VENCIMIENTO DE FACTURA  **************************************

   '
   Call SELMESA(DESDE%, HASTA%, CORRESP$)
'   CORRESP$ = MESANO$(MESAN$, DESDE%, HASTA%)
   If CORRESP$ = "" Then Exit Sub
   '
   INFO_GRILLA07.Caption = "Corresponde al Mes: " & CORRESP$
   
   Call CARGA_ENCABEZADO_GRILLA
   '
   'VEO COMO OBTENGO LAS SEMANAS
   
   PRIMER_DIA_MES% = DESDE%: PRIMER_DIA_SEMANA% = DESDE%
   DIATEXTO$ = DSEM$(PRIMER_DIA_SEMANA%)
   If DIATEXTO$ = "DOMINGO" Then PRIMER_DIA_SEMANA% = DIPOST(PRIMER_DIA_SEMANA%)
   If DIATEXTO$ = "SABADO" Then PRIMER_DIA_SEMANA% = DIPOST(DIPOST(PRIMER_DIA_SEMANA%))
   '
   'OBTENGO LOS DATOS DE LA PRIMER SEMANA CONTANDO A PARTIR DEL PRIMER DIA
   '***********************************************************************
   ULTIMO_DIA_MES% = HASTA%
   'FEX = PRIMER_DIA_MES%
   FEX = PRIMER_DIA_SEMANA%

     For i% = PRIMER_DIA_SEMANA% To (PRIMER_DIA_SEMANA% + 7)
       FEX = DIPOST(FEX)
       DIATEXTO$ = DSEM$(FEX)
       If DIATEXTO$ = "DOMINGO" Then
          ULTIMO_DIA_SEMANA% = FEX
          Exit For
       End If
     Next i%
DIAA = DSEM$(FEX)
   'FIN -- OBTENGO LOS DATOS DE LA PRIMER SEMANA CONTANDO A PARTIR DEL PRIMER DIA
   '******************************************************************************
   '
   'HAGO UNA CONSULTA SOBRE CAJABANC DE LOS CLIENTES QUE TIENEN RECIBOS EN EL RANGO SELECCIONADO
   'RECORRO LA LISTA DE LOS CLIENTES QUE SE LE REALIZARON COBRANZAS DENTRO DEL MES SELECCIONADO
   '******************************************************
   DESFE = FETEXCOR1$(DESDE%)
   HASFE = FETEXCOR1$(HASTA%)
   SQLX$ = "SELECT NUCLIEN, REFECOM FROM CAJABANC WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE TIPOMOVIM = 'RCOB' "
   SQLX$ = SQLX$ + " and fechemisi >= '" & DESFE & "'"
   SQLX$ = SQLX$ + " and fechemisi <= '" & HASFE & "'"
   SQLX$ = SQLX$ + " GROUP BY NUCLIEN, REFECOM "
   SQLX$ = SQLX$ + " ORDER BY REFECOM "
   
   DIAX$ = DSEM$(DESDE%)
   '
   JJ& = 0: VUELTA% = 0
   YA_CARGO_ENCABEZADO% = 0
   Dim ACUMULA_PROM_ATRA(5), ACUMULA_IMPORTEX_TOTAL(5), ACUMU_PROM_X_IMPORTEX_TOTAL(5)

   '
   Screen.MousePointer = 11
   INFO_GRILLA07.MSF.Redraw = False 'MEJORA VELOCIDAD DE CARGA DE GRILLA
   '
   Dim REG_CON_VALOR&(5)
   For i% = 1 To 5
     REG_CON_VALOR(i%) = 0
   Next i%

   Set RSBANC = READSET(SQLX$)
   With RSBANC
    Do While Not .EOF
      REG& = REG& + 1
      '
      NC% = XVALO(!NUCLIEN)
      CONDICION_DE_PAGO% = CPAGCLI%(NC%)
      PROMEDIO_DE_PAGOS = PROMEDIO_CONDI_PAGOS(CONDICION_DE_PAGO%) 'DEVUELVE EL PROMEDIO DE LAS CONDICIONES DE PAGO
      REFECONPAGO$ = TRIM$(ECOARCH$("CONPAG", MKI$(CONDICION_DE_PAGO%)))
      
      '
      INFO_GRILLA07.MSF.Rows = INFO_GRILLA07.MSF.Rows + 1
      '
      INFO_GRILLA07.MSF.TextMatrix(REG&, 1) = rasocli$(NC%)
      INFO_GRILLA07.MSF.TextMatrix(REG&, 2) = REFECONPAGO$
      INFO_GRILLA07.MSF.TextMatrix(REG&, 3) = XVALO(PROMEDIO_DE_PAGOS)
      '
      Condi$ = "TIPOMOVIM = 'RCOB' AND NUCLIEN = " & NC%
      '
      PRI_D_SEMANA% = PRIMER_DIA_SEMANA% 'ACTUALIZO LA VARIABLE
      ULT_D_SEMANA% = ULTIMO_DIA_SEMANA% 'ACTUALIZO LA VARIABLE
      ACUMU_PROM_ATRAZ = 0
      ACUMU_PROM_X_IMPORTEX = 0
      M3% = 0:
      CANTI_SEMANAS% = 0
      ACUMU_IMPORTEX = 0
      '
      For i% = 1 To 5
         DIAX$ = DSEM$(PRI_D_SEMANA%)
         If DIAX$ <> "DOMINGO" Then
           'SI EL PRIMER DIA DEL MES ES DOMINGO ESA SEMANA NO SE TOMA EN CUENTA, SERIA CASO DE 6 SEMANAS
           
           If PRI_D_SEMANA% > HASTA% Then Exit For 'SI EL PRIMER DIA DE LA SEMANA ES MAYOR TERMINA
           If ULT_D_SEMANA% > HASTA% Then ULT_D_SEMANA% = HASTA% ' SI EL ULTMO DIA > A FIN DE MES ASIGNO FIN DE MES
           P_D_SEM = FETEXCOR1$(PRI_D_SEMANA%)
           U_D_SEM = FETEXCOR1$(ULT_D_SEMANA%)
           CONDI2$ = Condi$ + " AND FECHEMISI >= '" & P_D_SEM & "' AND FECHEMISI <= '" & U_D_SEM & "'"
           importex# = XVALO(VALOBADA("CAJABANC", "Sum(IHABECOMP)", CONDI2$, "NOMESS"))  'COBRANZA
           IMPORTEP# = XVALO(VALOBADA("CAJABANC", "Sum(IHABECOMP * atracheque)", CONDI2$, "NOMESS"))  'COBRANZA PONDERADA
           PROMEDIO_ATRAZO1 = 0 '   'PARA EVITAR QUE REALICE UNA LECTURA SIN SENTIDO
           ACUMU_IMPORTEX = ACUMU_IMPORTEX + importex#
           If importex# > 0.05 Then 'SI IMPORTEX ES 0 NO VA A TENER PROMEDIO DE ATRAZO - EVITO VALOBADA
              'PROMEDIO_ATRAZO1 = XVALO(VALOBADA("CAJABANC", "AVG(AtraCheque)", CONDI2$, "NOMESS"))
              'CANTIDAD_RECIBOS = CantRegistros("cajabanc", condi2$)
              'CANTIDAD_DIAS = XVALO(VALOBADA("CAJABANC", "sum(AtraCheque)", condi2$, "NOMESS"))
              PROMEDIO_ATRAZO1 = 0
              If importex# > 0.05 Then
                PROMEDIO_ATRAZO1 = XVALO(FORMATNUM$((IMPORTEP# / importex#), "F12.0"))
              End If
              CANTI_SEMANAS% = CANTI_SEMANAS% + 1
           End If
           '
           
           'COMENTADO POR QUE NO SE ESTA PONIENDO QUE DIAS ABARCA LA SEMANA EN EL ENCABEZADO
'           'DOY FORMATO 01/07 PARA PRESENTAR EN EL ENCABEZADO DE LA COLUMNA SOLO LA PRIMERA VUELTA
'           If YA_CARGO_ENCABEZADO% < 1 Then
'            PDS$ = FECHATEX$(PRI_D_SEMANA%): PDS$ = Left$(PDS$, 3)
'            UDS$ = FECHATEX$(ULT_D_SEMANA%): UDS$ = Left$(UDS$, 2)
'            SEMANAX$ = "Sem. " & PDS$ & UDS$
'            INFO_GRILLA07.MSF.TextMatrix(0, 4 + M3%) = SEMANAX$
'            INFO_GRILLA07.MSF.TextMatrix(0, 5 + M3%) = "Cobr.-" & SEMANAX$
'            INFO_GRILLA07.MSF.TextMatrix(0, 6 + M3%) = "Prom.-" & SEMANAX$
'           End If
           '
           INFO_GRILLA07.MSF.TextMatrix(REG&, 4 + M3%) = FORMATNUM$(PROMEDIO_ATRAZO1, "G12.0")
           If PROMEDIO_ATRAZO1 > 0 Then REG_CON_VALOR(i%) = REG_CON_VALOR&(i%) + 1

           INFO_GRILLA07.MSF.TextMatrix(REG&, 5 + M3%) = FORMATNUM$(importex#, "G12.2")
           INFO_GRILLA07.MSF.TextMatrix(REG&, 6 + M3%) = FORMATNUM$((PROMEDIO_ATRAZO1 * importex#), "G12.2")
           '
           'VARIABLES CON VECTOR PARA ACUMULAR EL TOTAL DE CADA COLMNA
           ACUMULA_PROM_ATRA(i%) = ACUMULA_PROM_ATRA(i%) + PROMEDIO_ATRAZO1
           ACUMULA_IMPORTEX_TOTAL(i%) = ACUMULA_IMPORTEX_TOTAL(i%) + importex#
           '
           ACUMU_PROM_X_IMPORTEX = ACUMU_PROM_X_IMPORTEX + (PROMEDIO_ATRAZO1 * importex#) 'num 'ACUMULA TOTAL DE TODAS LAS FILAS
           ACUMU_PROM_X_IMPORTEX_TOTAL(i%) = ACUMU_PROM_X_IMPORTEX_TOTAL(i%) + (PROMEDIO_ATRAZO1 * importex#)
           '
           ACUMU_PROM_ATRAZ = ACUMU_PROM_ATRAZ + PROMEDIO_ATRAZO1
         End If
         '
         M3% = M3% + 3 ' ACUMULADOR PARA IR MOVIENDOSE A LA COLUMNA CORRESPONDIENTE
         '
         'ARMA LA NUEVA SEMANA
         PRI_D_SEMANA% = ULT_D_SEMANA% + 1
         ULT_D_SEMANA% = ULT_D_SEMANA% + 7
      Next i%
      '
      If ACUMU_IMPORTEX < 0.005 Then ACUMU_IMPORTEX = 1
      PROME_PONDERADO = ACUMU_PROM_X_IMPORTEX / ACUMU_IMPORTEX  'COLUMNA NUM TOTAL DE TODAS LAS FILAS/ COLUMNA $$ COBRANZA TOTAL DE TODAS LAS FILAS
      INFO_GRILLA07.MSF.TextMatrix(REG&, 19) = FORMATNUM$(PROME_PONDERADO, "F6.0")
      If CANTI_SEMANAS% < 1 Then CANTI_SEMANAS% = 1 '
      INFO_GRILLA07.MSF.TextMatrix(REG&, 20) = FORMATNUM$(PROME_PONDERADO - PROMEDIO_DE_PAGOS, "F6.0")
      YA_CARGO_ENCABEZADO% = 1
      '
      .MoveNext
    Loop
    '
   End With
   RSBANC.Close
   Set RSBANC = Nothing
   Screen.MousePointer = 1
   INFO_GRILLA07.MSF.Redraw = True
   '
   If REG& < 1 Then Exit Sub
   
   Call COLOREAR_GRILLASQL(INFO_GRILLA07.MSF, &HEFF1F1)
   '
   'IMPRIMIR TOTALES
   '*****************
   
   ACUMULADO_IMPORTEX = 0
   INFO_GRILLA07.MSF.Rows = INFO_GRILLA07.MSF.Rows + 3
   '
   M3% = 0
   '
   INFO_GRILLA07.MSF.TextMatrix(INFO_GRILLA07.MSF.Rows - 3, 2) = "Totales: "
   '
   'CALCULA TOTALES
   For i% = 1 To 5
        If REG_CON_VALOR&(i%) < 1 Then REG_CON_VALOR&(i%) = 1
        'INFO_GRILLA07.MSF.TextMatrix(INFO_GRILLA07.MSF.Rows - 3, 4 + M3%) = TRIM$(FORMATNUM$((ACUMULA_PROM_ATRA(I%) / REG_CON_VALOR&(I%)), "F14.0")) ' PROMEDIO)
        INFO_GRILLA07.MSF.TextMatrix(INFO_GRILLA07.MSF.Rows - 3, 5 + M3%) = FORMATNUM$(ACUMULA_IMPORTEX_TOTAL(i%), "G14.2")
        ACUMULADO_IMPORTEX = ACUMULADO_IMPORTEX + ACUMULA_IMPORTEX_TOTAL(i%)
        INFO_GRILLA07.MSF.TextMatrix(INFO_GRILLA07.MSF.Rows - 3, 6 + M3%) = FORMATNUM$(ACUMU_PROM_X_IMPORTEX_TOTAL(i%), "G14.2")
        ACUMULADO_NUM = ACUMULADO_NUM + ACUMU_PROM_X_IMPORTEX_TOTAL(i%)
        '
        'ASIGNO 0 POR SI ACUMULADO_IMPORTEX ES < 1
'        INFO_GRILLA07.MSF.TextMatrix(INFO_GRILLA07.MSF.Rows - 3, 6 + M3%) = 0
'        If ACUMULA_IMPORTEX_TOTAL(I%) > 0 Then
''           INFO_GRILLA07.MSF.TextMatrix(INFO_GRILLA07.MSF.Rows - 3, 6 + M3%) = TRIM$(FORMATNUM$((ACUMU_PROM_X_IMPORTEX_TOTAL(I%) / ACUMULA_IMPORTEX_TOTAL(I%)), "G14.0"))
'        End If
        M3% = M3% + 3
   Next i%
   '
   'IMPRIME LOS VALORES DEBAJO DE LA COLUMNA PROMEDIO PONDERADO A PIE DE PAGINA UNO DEBAJO DE OTRO
   INFO_GRILLA07.MSF.TextMatrix(INFO_GRILLA07.MSF.Rows - 3, 19) = ACUMULADO_NUM
   INFO_GRILLA07.MSF.TextMatrix(INFO_GRILLA07.MSF.Rows - 2, 19) = ACUMULADO_IMPORTEX
'
   If ACUMULADO_IMPORTEX < 0.005 Then ACUMULADO_IMPORTEX = 1
   INFO_GRILLA07.MSF.TextMatrix(INFO_GRILLA07.MSF.Rows - 1, 19) = FORMATNUM$((ACUMULADO_NUM / ACUMULADO_IMPORTEX), "F6.0")
   '
   'COLOREA LOS TOTALES
   Call COLOREAR_GRILLA_SOLO_UNA_AFILA(INFO_GRILLA07.MSF, &HC0E0FF, INFO_GRILLA07.MSF.Rows - 3)
   '
   Call COLOREAR_GRILLA_SOLO_UNA_AFILA(INFO_GRILLA07.MSF, &H80C0FF, INFO_GRILLA07.MSF.Rows - 1)
   '
   INFO_GRILLA07.Show 1
   '
End Sub

Sub ESTADISTICA_DE_COBRANZA_X_MES()
'COLUMNA 1 RAZON SOCIAL
'2 CONDICION DE PAGO FORMATO TEXTO
'3 CONDICION DE PAGOS (PROMEDIO DE DIAS)
'4 CANTIDAD DE DIAS DE COBRANZA PONDERADO(CAMPO ATRACHEQUE) PROMEDIO (SE REPITE PARA CADA SEMANA COLUMNA 4,7,10,13,16)
'5 ACUMULADO DEL IMPORTE DE LOS RECIBOS POR CLIENTE                  (SE REPITE PARA CADA SEMANA)COLUMNA 5,8,11,14,17)
'6 VALOR DE COLUMNA 4 X VALOR DE LA COLUMNA 5                        (SE REPITE PARA CADA SEMANA)COLUMNA 6,9,12,15,18)
'COLUMNA 19 SUMA DE LAS COLUMNAS 6 DIVIDIDO LA SUMA  DE LAS COLUMNA 5
'COLUMNA 20 PROMEDIO DE LAS OCOLUMNAS 4  O SEA LA SUMA DE LA COLUMNA 4 DIVIDIDO LA CANTIDAD DE SEMANAS
'
'TOTALES A PIE DE GRILLA.

'TOTALES PROMEDIO DE LA COLUNA 4 (PROMEDIO DE ATRAZO)
'SUMA DE LA COLUMNA 5 (IMPORTES DE COBRANZA)
'SUMA DE LA COLUMNA 6 (NUM)

'RESULTADO DE LA FORMULA: (COLUMNA 3 * 100) / COLUMNA 2
  '**************** FUNCIONA CON CONFIGURACION POR VENCIMIENTO DE FACTURA  **************************************
   '
   DESHASMES.Show 1
   If XVALO(DESHASMES.DIADESDE) < 1 Or XVALO(DESHASMES.DIAHASTA) < 1 Then Exit Sub
   DIA1_MES% = XVALO(DESHASMES.DIADESDE)
   DIA_HASTA_SELEC% = XVALO(DESHASMES.DIAHASTA)
   '
   If DESHASMES.Label3 = DESHASMES.Label5 Then
      INFO_GRILLA09.Caption = "Corresponde a: " & DESHASMES.Label3
   Else
      INFO_GRILLA09.Caption = "Desde: " & DESHASMES.Label3 & " Hasta: " & DESHASMES.Label5
   End If
   '
   NCLIE = XVALO(DESHASMES.Label7)
   '
   Call CARGA_ENCABEZADO_GRILLA_MESES
   '
   'VEO COMO OBTENGO LAS SEMANAS
   
   DIA_ULT_MES% = FINMESACTUAL%(DIA1_MES%)
   '
   DESFE = FETEXCOR1$(DIA1_MES%)
   HASFE = FETEXCOR1$(DIA_HASTA_SELEC%)
   CANTIMESES% = 1 + DateDiff("m", FECHATEXLAR$(DIA1_MES%), FECHATEXLAR$(DIA_HASTA_SELEC%))
   If CANTIMESES% > 12 Then
      Call COMUNI("Rango de Meses Seleccionado Supera Cantidad Máxima (12)")
      Exit Sub
   End If
   
   '
   SQLX$ = "SELECT NUCLIEN, REFECOM FROM CAJABANC WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE TIPOMOVIM = 'RCOB' "
   SQLX$ = SQLX$ + " and fechemisi >= '" & DESFE & "'"
   SQLX$ = SQLX$ + " and fechemisi <= '" & HASFE & "'"
   
   If NCLIE > 0 Then SQLX$ = SQLX$ & " AND Nuclien=" & CStr(NCLIE)
   
   SQLX$ = SQLX$ + " GROUP BY NUCLIEN, REFECOM "
   SQLX$ = SQLX$ + " ORDER BY REFECOM "
   
   '
   JJ& = 0: VUELTA% = 0
   Dim ACUMULA_PROM_ATRA(12), ACUMULA_IMPORTEX_TOTAL(12), ACUMU_PROM_X_IMPORTEX_TOTAL(12)
   '
   Screen.MousePointer = 11
   INFO_GRILLA09.MSF.Redraw = False 'MEJORA VELOCIDAD DE CARGA DE GRILLA
   '
   Dim REG_CON_VALOR&(12)
   For i% = 1 To 12
     REG_CON_VALOR(i%) = 0
   Next i%
   '
   YA_CARGO_ENCABEZADO% = 0
   Set RSBANC = READSET(SQLX$)
   With RSBANC
    Do While Not .EOF
      REG& = REG& + 1
      INFO_GRILLA09.MSF.Cols = 6 + (CANTIMESES% * 3)
      '
      NC% = XVALO(!NUCLIEN)
      CONDICION_DE_PAGO% = CPAGCLI%(NC%)
      PROMEDIO_DE_PAGOS = PROMEDIO_CONDI_PAGOS(CONDICION_DE_PAGO%) 'DEVUELVE EL PROMEDIO DE LAS CONDICIONES DE PAGO
      REFECONPAGO$ = TRIM$(ECOARCH$("CONPAG", MKI$(CONDICION_DE_PAGO%)))
      
      '
      INFO_GRILLA09.MSF.Rows = INFO_GRILLA09.MSF.Rows + 1
      '
      INFO_GRILLA09.MSF.TextMatrix(REG&, 1) = rasocli$(NC%)
      INFO_GRILLA09.MSF.TextMatrix(REG&, 2) = REFECONPAGO$
      INFO_GRILLA09.MSF.TextMatrix(REG&, 3) = XVALO(PROMEDIO_DE_PAGOS)
      '
      Condi$ = "TIPOMOVIM = 'RCOB' AND NUCLIEN = " & NC%
      '
      ACUMU_PROM_ATRAZ = 0
      ACUMU_PROM_X_IMPORTEX = 0
      M3% = 0:
      ACUMU_IMPORTEX = 0
      DIA1_MES% = XVALO(DESHASMES.DIADESDE) 'ACTUALIZO LA VARIABLE
      DIA_ULT_MES% = FINMESACTUAL%(DIA1_MES%) 'ACTUALIZO LA VARIABLE

      '
      For i% = 1 To CANTIMESES%
           
           If DIA1_MES% > DIA_HASTA_SELEC% Then Exit For 'SI EL PRIMER DIA DE LA SEMANA ES MAYOR TERMINA
           If DIA_ULT_MES% > DIA_HASTA_SELEC% Then DIA_ULT_MES% = DIA_HASTA_SELEC% ' SI EL ULTMO DIA > A FIN DE MES ASIGNO FIN DE MES
           P_D_MES = FETEXCOR1$(DIA1_MES%)
           U_D_MES = FETEXCOR1$(DIA_ULT_MES%)
           CONDI2$ = Condi$ + " AND FECHEMISI >= '" & P_D_MES & "' AND FECHEMISI <= '" & U_D_MES & "'"
           importex# = XVALO(VALOBADA("CAJABANC", "Sum(IHABECOMP)", CONDI2$, "NOMESS"))  'COBRANZA
           IMPORTEP# = XVALO(VALOBADA("CAJABANC", "Sum(IHABECOMP * atracheque)", CONDI2$, "NOMESS"))  'COBRANZA PONDERADA
           PROMEDIO_ATRAZO1 = 0 '   'PARA EVITAR QUE REALICE UNA LECTURA SIN SENTIDO
           ACUMU_IMPORTEX = ACUMU_IMPORTEX + importex#
           If importex# > 0.05 Then 'SI IMPORTEX ES 0 NO VA A TENER PROMEDIO DE ATRAZO - EVITO VALOBADA
              'PROMEDIO_ATRAZO1 = XVALO(VALOBADA("CAJABANC", "AVG(AtraCheque)", condi2$, "NOMESS"))
              'CANTIDAD_RECIBOS = CantRegistros("cajabanc", condi2$)
              'CANTIDAD_DIAS = XVALO(VALOBADA("CAJABANC", "sum(AtraCheque)", condi2$, "NOMESS"))
              PROMEDIO_ATRAZO1 = 0
              If IMPORTEP# > 0.05 Then
                PROMEDIO_ATRAZO1 = XVALO(FORMATNUM$((IMPORTEP# / importex#), "F12.0"))
              End If
           End If
           '
           
           'COMENTADO POR QUE NO SE ESTA PONIENDO QUE DIAS ABARCA LA SEMANA EN EL ENCABEZADO
           'DOY FORMATO 01/07 PARA PRESENTAR EN EL ENCABEZADO DE LA COLUMNA SOLO LA PRIMERA VUELTA
           If YA_CARGO_ENCABEZADO% < 1 Then
            MEX$ = METEXTO$(DIA1_MES%)
            MEX1$ = Left$(MEX$, 5)
            ANOX$ = Right(MEX$, 2)
            MEX2$ = MEX1$ & "-" & ANOX$
            INFO_GRILLA09.MSF.TextMatrix(0, 4 + M3%) = MEX2$
            INFO_GRILLA09.MSF.TextMatrix(0, 5 + M3%) = "Cob. $$"
            INFO_GRILLA09.MSF.TextMatrix(0, 6 + M3%) = "Num."
           End If
           '
           INFO_GRILLA09.MSF.TextMatrix(REG&, 4 + M3%) = TRIM$(FORMATNUM$(PROMEDIO_ATRAZO1, "G12.0"))
           If PROMEDIO_ATRAZO1 > 0 Then REG_CON_VALOR(i%) = REG_CON_VALOR&(i%) + 1

           INFO_GRILLA09.MSF.TextMatrix(REG&, 5 + M3%) = FORMATNUM$(importex#, "G12.2")
           INFO_GRILLA09.MSF.TextMatrix(REG&, 6 + M3%) = FORMATNUM$((PROMEDIO_ATRAZO1 * importex#), "G12.2")
           '
           'VARIABLES CON VECTOR PARA ACUMULAR EL TOTAL DE CADA COLMNA
           ACUMULA_PROM_ATRA(i%) = ACUMULA_PROM_ATRA(i%) + PROMEDIO_ATRAZO1
           ACUMULA_IMPORTEX_TOTAL(i%) = ACUMULA_IMPORTEX_TOTAL(i%) + importex#
           '
           ACUMU_PROM_X_IMPORTEX = ACUMU_PROM_X_IMPORTEX + (PROMEDIO_ATRAZO1 * importex#) 'num 'ACUMULA TOTAL DE TODAS LAS FILAS
           ACUMU_PROM_X_IMPORTEX_TOTAL(i%) = ACUMU_PROM_X_IMPORTEX_TOTAL(i%) + (PROMEDIO_ATRAZO1 * importex#)
           '
           ACUMU_PROM_ATRAZ = ACUMU_PROM_ATRAZ + PROMEDIO_ATRAZO1
         
           '
           M3% = M3% + 3 ' ACUMULADOR PARA IR MOVIENDOSE A LA COLUMNA CORRESPONDIENTE
           '
           'ARMA EL  RANGO DEL NUEVO MES
         DIA1_MES% = DIASPOST(DIA_ULT_MES%, 1)
         DIA_ULT_MES% = FINMESACTUAL%(DIA1_MES%)

      Next i%
      '
      With INFO_GRILLA09
        .MSF.Cols = 6 + M3%
        If ACUMU_IMPORTEX < 0.05 Then ACUMU_IMPORTEX = 1
        PROME_PONDERADO = ACUMU_PROM_X_IMPORTEX / ACUMU_IMPORTEX  'COLUMNA NUM TOTAL DE TODAS LAS FILAS/ COLUMNA $$ COBRANZA TOTAL DE TODAS LAS FILAS
        .MSF.TextMatrix(REG&, 4 + M3%) = FORMATNUM$(PROME_PONDERADO, "F6.0")
        .MSF.TextMatrix(REG&, 5 + M3%) = FORMATNUM$(PROME_PONDERADO - PROMEDIO_DE_PAGOS, "F6.0")
        
      End With
      '
      YA_CARGO_ENCABEZADO% = 1
      .MoveNext
    Loop
    '
    With INFO_GRILLA09
        NUCOLU% = 5 + M3%: .TEPRUEBA = "1234567890123456": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "P.Pond."
        NUCOLU% = 6 + M3%: .TEPRUEBA = "12345678": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Desvío"
    End With
   End With
   RSBANC.Close
   Set RSBANC = Nothing
   Screen.MousePointer = 1
   INFO_GRILLA09.MSF.Redraw = True
   '
   If REG& < 1 Then Exit Sub

   Call COLOREAR_GRILLASQL(INFO_GRILLA09.MSF, &HEFF1F1)
'   '
'   'IMPRIMIR TOTALES
'   '*****************
'
   ACUMULADO_IMPORTEX = 0
   INFO_GRILLA09.MSF.Rows = INFO_GRILLA09.MSF.Rows + 4
   '
   M3% = 0
   '
   INFO_GRILLA09.MSF.TextMatrix(INFO_GRILLA09.MSF.Rows - 3, 2) = "Totales: "
   '
   M3% = 0
   '
   INFO_GRILLA09.MSF.TextMatrix(INFO_GRILLA09.MSF.Rows - 3, 2) = "Totales: "
   'CALCULA TOTALES
   For i% = 1 To CANTIMESES%
        If REG_CON_VALOR&(i%) < 1 Then REG_CON_VALOR&(i%) = 1
        'INFO_GRILLA09.MSF.TextMatrix(INFO_GRILLA09.MSF.Rows - 3, 4 + M3%) = TRIM$(FORMATNUM$((ACUMULA_PROM_ATRA(I%) / REG_CON_VALOR&(I%)), "F14.0")) ' PROMEDIO)
        INFO_GRILLA09.MSF.TextMatrix(INFO_GRILLA09.MSF.Rows - 3, 5 + M3%) = FORMATNUM$(ACUMULA_IMPORTEX_TOTAL(i%), "G14.2")
        ACUMULADO_IMPORTEX = ACUMULADO_IMPORTEX + ACUMULA_IMPORTEX_TOTAL(i%)
        INFO_GRILLA09.MSF.TextMatrix(INFO_GRILLA09.MSF.Rows - 3, 6 + M3%) = FORMATNUM$(ACUMU_PROM_X_IMPORTEX_TOTAL(i%), "G14.2")
        ACUMULADO_NUM = ACUMULADO_NUM + ACUMU_PROM_X_IMPORTEX_TOTAL(i%)
        '
        M3% = M3% + 3
   Next i%
   
   
   With INFO_GRILLA09
    NUCOLU% = 6 + M3%
    .MSF.Cols = NUCOLU% + 1
    .TEPRUEBA = "1234567890123456": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "P.Pond."
    .TEPRUEBA = "12345678": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU%) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Desvío"
   End With
'   '
'   'IMPRIME LOS VALORES DEBAJO DE LA COLUMNA PROMEDIO PONDERADO A PIE DE PAGINA UNO DEBAJO DE OTRO

   INFO_GRILLA09.MSF.TextMatrix(INFO_GRILLA09.MSF.Rows - 3, 4 + M3%) = ACUMULADO_NUM
   INFO_GRILLA09.MSF.TextMatrix(INFO_GRILLA09.MSF.Rows - 2, 4 + M3%) = ACUMULADO_IMPORTEX
'
   If ACUMULADO_IMPORTEX < 0.005 Then ACUMULADO_IMPORTEX = 1
   INFO_GRILLA09.MSF.TextMatrix(INFO_GRILLA09.MSF.Rows - 1, 4 + M3%) = FORMATNUM$((ACUMULADO_NUM / ACUMULADO_IMPORTEX), "F6.0")
   '
   'COLOREA LOS TOTALES
   Call COLOREAR_GRILLA_SOLO_UNA_AFILA(INFO_GRILLA09.MSF, &HC0E0FF, INFO_GRILLA09.MSF.Rows - 3)
   '
   Call COLOREAR_GRILLA_SOLO_UNA_AFILA(INFO_GRILLA09.MSF, &H80C0FF, INFO_GRILLA09.MSF.Rows - 1)
   '
   INFO_GRILLA09.Show 1
   '
End Sub



Private Sub Command11_Click()
    Command11.Enabled = False
    Call CIERRARCH("*.*")
    Call CONECRUN("ARCHIG07/AMACLI", "Archivos Maestros")
99  Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   For U& = 1 To ULTREG&("MECOBRA")
     XX$ = REGLEIDO$("MECOBRA", U&)
     TTXX$ = TTXX$ + Mid$(XX$, 45, 12)
   Next U&
   MsgBox TTXX$
End Sub

Private Sub Command2_Click()
    TXTRAZA$ = TXTRAZA$ & "PUNTO RECI.1 = " & Now & vbCrLf
    If List1.ListCount > 0 Or List2.ListCount > 0 Then
          Call CIERRARCH("!CARMECO")
          Call BLANARCH("!CARMECO")
          Call CIERRARCH("!APLIRECI")
          Call BLANARCH("!APLIRECI")
          List1.Clear
          List2.Clear
          Label4.Caption = ""
          Label8.Caption = ""
          Label10.Caption = ""
          Label14.Caption = ""
          Label18.Caption = ""
          Label20.Caption = ""
          '\\\OT-07-0029-RG-26/01/07///
          DESCUEN(0) = ""
          Text7 = ""
          Frame8.Caption = "Descuento (N./Cred)"
          '\\\OT-07-0029-RG-FIN///
          BOTNEXT.Enabled = True
          Command5.Enabled = False
          Text1.SetFocus
    TXTRAZA$ = TXTRAZA$ & "PUNTO RECI.2 = " & Now & vbCrLf
          
          Call LIMPPROM
        Else
    TXTRAZA$ = TXTRAZA$ & "PUNTO RECI.3 = " & Now & vbCrLf
        
          If TRIM$(Control$("RECISTA", "AVAHOJA")) = "NEWPAGE" Then
              'Printer.EndDoc
              Call SETSPOOL("ENDDOC", "")
          End If
    TXTRAZA$ = TXTRAZA$ & "PUNTO RECI.4 = " & Now & vbCrLf
          
          Call CIERRARCH("*.*")
     TXTRAZA$ = TXTRAZA$ & "PUNTO RECI.5 = " & Now & vbCrLf
         
          Call BAJALDISCO
    TXTRAZA$ = TXTRAZA$ & "PUNTO RECI.6 = " & Now & vbCrLf
          Hide
    TXTRAZA$ = TXTRAZA$ & "PUNTO RECI.7 = " & Now & vbCrLf
          
    End If
    Call REMPLAGRID(Me, GRID1, Picture1, Label12, List1)
    Call REMPLAGRID(Me, GRID2, Picture2, Label9, List2)
    TXTRAZA$ = TXTRAZA$ & "PUNTO RECI.8 = " & Now & vbCrLf
    
End Sub

Private Sub Command23_Click()
   Command23.Enabled = False
   Call RECICO07.Command23_Click
   Command23.Enabled = True
End Sub

Sub Command24_Click()
   Command24.Enabled = False
   'llamada a busqueda de cheques por numero
   Call GESVAL07.BUSCACHEQ
   Command24.Enabled = True

End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    NCLI$ = TRIM$(Text1.TEXT)
    'Call CONECRUN("*NOTACLI/" + NCLI$, "Notas de Clientes")
    NC% = XVALO(Text1.TEXT)
    If NC% > 0 Then Call NotaCli(NC%, 2)
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    RECUEC07.Text1.TEXT = Text1.TEXT
    RECUEC07.Show 1
    Command4.Enabled = True
End Sub

Function VALIDARECIBO%()
    '
    Label34 = 1: Label34.Refresh
    Screen.MousePointer = 11   ' Grabando Recibo - Espere un Momento, por Favor"
    VDRX% = 0
    '
    ' NUMERO DE CLIENTE EN TEXT1
    NCLI = XVALO(Text1)
    If NCLI <> XVALO(Text1.TEXT) Then
       Call MENSERR(24, "Imposible Grabar - Intente de Nuevo")
       Call CIERRARCH("*.*")
       Call FINAL("")
    End If
    APRORECIB.Label1 = TRIM$(Str$(NCLI))
    APRORECIB.Label2 = rasocli$(NCLI)
    '
    IREDON# = XVALO(Text8)
    If Abs(IREDON#) >= 0.005 Then
       CUERED% = 0
       CUEREDO$ = TRIM$(Control$("RECISTA", "CUEREDON"))
       If CUEREDO$ <> "" Then CUERED% = CUEINT%(CUEREDO$)
       If CUERED% < 1 Then
          Call MENSERR(24, "Falta Cuenta Contable Redondeo.")
          VDRX% = (-1)
          GoTo 99
       End If
    End If
    '
    ' VERIFICA QUE EL RECIBO POSEA CHEQUES
    If XVALO(Label14.Caption) < 0.01 Then
        Call COMUNI("Imposible Grabar - Falta Detalle de Valores")
        VDRX% = (-1)
        GoTo 99
    End If
    '
    ' BALANCE ENTRE VALORES Y CANCELACIONES
    TOTAVAL# = CDbl(XVALO(Label14.Caption)) + CDbl(XVALO(DESCUEN(0)))
    TOTACAN# = CDbl(XVALO(Label18.Caption))
    If ROUND#(TOTACAN#, 2) > ROUND#(TOTAVAL#, 2) Then
        Call COMUNI("Operación Incorrecta.\  \Las Cancelaciones no Pueden\Exceder la Suma de Valores Recibidos.")
        VDRX% = (-1)
        GoTo 99
    End If
    '
    If TOTACAN# < (-0.005) Then
        Call MENSERR(24, "Valores no Procesables")
        VDRX% = (-1)
        GoTo 99
    End If
    '
    If XVALO(Label20) >= 0.005 Then
      If XVALO(DESCUEN(0)) >= 0.005 Then
        Call MENSERR(24, "Para Generar Descuento\Debe Aplicarse la Totalidad del Pago.")
        VDRX% = (-1)
        GoTo 99
      End If
    End If
    '
    APRORECIB.Label12 = Label14
    '
    ANTI# = CDbl(XVALO(Label20.Caption))
    APRORECIB.Label14 = FORMATNUM$(ANTI#, "F12.2")
    '
    ' FECHA DE EMISIÓN DEL RECIBO
    FECHALIM% = XVALO(Control$("", "CIEMEVEN"))
    FLIMI = CVDAT(FECHALIM%)
    FF% = FECHANUM(Text9.TEXT)
    If FF% < 1 Then
        Call MENSERR(24, "Fecha de Emisión no Válida")
        VDRX% = (-1)
        GoTo 99
    End If
    '
    If FF% <= FECHALIM% Then
        Call MENSERR(24, "Fecha no Válida.\  \Ese Período Mensual Está Cerrado.")
        VDRX% = (-1)
        GoTo 99
    End If
    FEX = FF%
    APRORECIB.Label6 = FECHATEX$(FEX)
    '
    Call SETULTREG("!PRINRECI", 0)
    RENDISPO& = XVALO(Control$("RECISTA", "CARENDET"))
    If RENDISPO& < 1 Then RENDISPO& = 9999
    If RENGLONEC& > RENDISPO& Then
      Call MENSERR(24, "Imposible Registrar este Recibo.\   \Excede Renglones Disponibles\en Formulario de Impresión.")
      VDRX% = (-1)
      GoTo 99
    End If
    '
    NROO& = XVALO(Mid$(TRIM$(Text11.TEXT), 4))
    If Val(NROO&) < 1 Then
      Call MENSERR(24, "Número de Recibo no Válido")
      VDRX% = (-1)
      GoTo 99
    End If
    APRORECIB.Label4 = TRIM$(Str$(NROO&))
    '
    COBRACT% = 0
    On Error Resume Next
    COBRACT% = XVALO(Label22.Caption)
    On Error GoTo 0
    APRORECIB.Label3 = TRIM$(Str$(COBRACT%))
    '
    CUECARTE% = CUEMECOB%(-1)
    APRORECIB.Label16 = TRIM$(Str$(CUECARTE%))
    APRORECIB.Label17 = ECOARCH$("CUECOT", MKI$(CUECARTE%))
    '
    APRORECIB.Label29 = TRIM$(Mid$(Frame8.Caption, 9))
      NFCY$ = APRORECIB.Label29
    APRORECIB.Label31 = Left$(NFCY$, 3) + Mid$(NFCY$, 11)
      NEFA7# = ROUND#(XVALO(DESCUEN(0)) / (1 + ALIVA / 100), 2)
      NIVA8# = XVALO(DESCUEN(0)) - NEFA7#
    APRORECIB.Label27 = TRIM$(FORMATNUM$(NEFA7#, "F12.2"))
    APRORECIB.Label25 = TRIM$(FORMATNUM$(NIVA8#, "F12.2"))
      CUEDESCU$ = TRIM$(Control$("", "CUEDESCU"))
      CUECONDE% = CUECOINT%(CUEDESCU$)
    APRORECIB.Label22 = TRIM$(Str$(CUECONDE%))
    APRORECIB.Label21 = CODCUE$(CUECONDE%)
    '
    If NEFA7# >= 0.005 Then
      If CUECONDE% < 1 Or TRIM$(APRORECIB.Label21) = "" Then
        Call MENSERR(24, "Imposible Grabar N/Cred. Descuento.\   \Falta Definir Cuenta Contable.")
        VDRX% = (-1)
        GoTo 99
      End If
    End If
    '
    'BLOQUEA ARCHIVO DE COBRANZA
    Call BLOQARCH("CUECORR")
    '
    RECICOR$ = TRIM$(Str$(COBRACT%))
    While Len(RECICOR$) < 2: RECICOR$ = "0" + RECICOR$: Wend
    RECICOR$ = TRIM$(Str$(NROO&)) + "-" + RECICOR$
    While Len(RECICOR$) < 9: RECICOR$ = "0" + RECICOR$: Wend
    RECICOR$ = "RB-" + RECICOR$
    APRORECIB.Label8 = RECICOR$
    '
    RECILAR$ = TRIM$(Str$(NROO&))
    While Len(RECILAR$) < 8: RECILAR$ = "0" + RECILAR$: Wend
    RECILAR$ = TRIM$(Str$(COBRACT%)) + "-" + RECILAR$
    While Len(RECILAR$) < 13: RECILAR = "0" + RECILAR$: Wend
    RECILAR$ = "RB-X-" + RECILAR$
    APRORECIB.Label10 = RECILAR$
    '
    'JANDY SQL
    
    FENUE$ = Control$("", "FENURECI")
    If FENUE$ = "" Then FENUE$ = "01/01/80"
    FEXI% = FECHANUM(FENUE$)
    'FLIMI = CVDAT(FECHANUM(FENUE$))
    FLIMI = FETEXCOR1$(FEXI%)
    
    'If FLIMI > Now Then FLIMI = CVDAT(HOY(HO$))
    'FLIMI1 = CDbl(FLIMI)
     
 
    '
    SQLX$ = "SELECT NumeComp FROM CajaBanc WITH(NOLOCK)"
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'RCOB' "
    SQLX$ = SQLX$ + "AND (CodiCom_Cor = '" & RECICOR$ & "' "
    SQLX$ = SQLX$ + "OR CodiCom_Lar = '" & RECILAR$ & "') "
    SQLX$ = SQLX$ + "AND Status < 9 "
    'SQLX$ = SQLX$ + "AND CDBL(FechEmisi) >= " & FLIMI1 & " "
    SQLX$ = SQLX$ + "AND FechEmisi >= '" & FLIMI & "' "
    
    Dim OrpaTemp As ADODB.Recordset
    
    'Set OrpaTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    'Set OrpaTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
    'se modifica por que la rutina readset captura el error de tiempo de espera
    Set OrpaTemp = READSET(SQLX$)
    With OrpaTemp
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Call LIBARCH("CUECORR")
        Call MENSERR(24, "Al Grabar este Recibo\el Número Propuesto Estaba Ocupado.\  \Verifique el Nuevo Número Propuesto\y Vuelva a Grabar e Imprimir.")
        Call NURECIBO(NROPRO&)
        Call PRENURE(NROPRO&)
        VDRX% = (-1)
        GoTo 99
      End If
      On Error GoTo 0
    End With
    OrpaTemp.Close
    Set OrpaTemp = Nothing
    '
    If NEFA7# >= 0.005 Then
      If FORFASQL.VERINUFA%(NFCY$) < 1 Then
        Call LIBARCH("CUECORR")
        Call MENSERR(24, "Al Grabar este Recibo\el Número de N/Cred. Estaba Ocupado.\  \Verifique el Nuevo Número Propuesto\y Vuelva a Grabar e Imprimir.")
        Frame8.Caption = "N./Cred. " + PRONOCRE$
        VDRX% = (-1)
        GoTo 99
      End If
    End If
    '
    VDRX% = 1
    '
99  VALIDARECIBO% = VDRX%
    Label34 = 4: Label34.Refresh
    Screen.MousePointer = 1
    '
End Function
'
Private Sub Command5_Click()
    '
    Command5.Enabled = False
    If VALIDARECIBO% <> 1 Then
       Call LIBARCH("CUECORR")
       Command5.Enabled = True
       Exit Sub
    End If
    '
    Call GRARECISTA(OKEY%)
    Call LIBARCH("CUECORR")
    '
    If OKEY% = 0 Then
        Command5.Enabled = True
        Exit Sub
      ElseIf OKEY% <> 6 Then
        Call COMUNI("La Registración de este Recibo\no Pudo Completarse. (" + TRIM$(Str$(OKEY%)) + ")\  \Consulte a su Soporte de Sistemas FLEXOFT.")
        Call FINAL("")
    End If
    
    On Error GoTo 0
    HOJASUEL% = 1
    If AUTONUME% = 1 Then HOJASUEL% = 0
    Call PRIRECISTA
    '
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 1 = " & Now & vbCrLf

    If Val(DESCUEN(0)) >= 0.005 Then
       Call PRICREDESC
    End If
    '
    Call CIERRARCH("!CARMECO")
    Call BLANARCH("!CARMECO")
    Call CIERRARCH("!APLIRECI")
    Call BLANARCH("!APLIRECI")
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 1 = " & Now & vbCrLf
    List1.Clear
    List2.Clear
    Label4.Caption = ""
    Label8.Caption = ""
    Label10.Caption = ""
    Label14.Caption = ""
    Label18.Caption = ""
    Label20.Caption = ""
    '\\\OT-07-0029-RG-26/01/07///
    DESCUEN(0) = ""
    Text7 = ""
    Text8 = ""
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 2 = " & Now & vbCrLf
    Call LIMPPROM
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 3 = " & Now & vbCrLf
    Frame8.Caption = "Descuento (N./Cred)"
    '\\\OT-07-0029-RG-FIN///
    BOTNEXT.Enabled = True
    Command5.Enabled = False
    FF% = HOY(HO$)
    Text9.TEXT = HO$
    '
    NROPRO& = 0
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 4 = " & Now & vbCrLf
    
    Call NURECIBO(NROPRO&)
    
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 5 = " & Now & vbCrLf

    Call PRENURE(NROPRO&)
    '
    Text11 = "RB.      "
    Text11.SelStart = 3
    Text11.SelLength = Len(Text11.TEXT) - 3
    Text11.SetFocus
    
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 6 = " & Now & vbCrLf

    If EsPorPedido& > 0 Then
      Command2_Click ' CERRAR LA VENTANA
    End If
    
'GRABO EL ARCHIVO DE TRAZABILIDAD
'For I1% = 1 To 32767
'     ARCH$ = LUDAT$ & "TRAZARECI" & TRIM$(Str$(I1%)) & ".TXT"
'     If EXISTE%(ARCH$) < 1 Then
'       NARCHI2% = FILEOPEN%(ARCH$, 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA
'       Print #NARCHI2%, TXTRAZA_BR$ ' ESCRIBE EN EL ARCHIVO EL NUEVO TEXTO
'       Close #NARCHI2%
'       Exit For
'     End If
'Next I1%
    
    '
    'Call REMPLAGRID(Me, GRID1, Picture1, Label12, List1)
    'Call REMPLAGRID(Me, GRID2, Picture2, Label9, List2)
    '
End Sub

Sub PRENURE(NROPRO&)
    '
    Static MODONUM$
    If MODONUM$ = "" Then
      MODONUM$ = Control$("", "NURECIST")
    End If
    '
    TNUME$ = "RB.      "
    If MODONUM$ = "AUTONUME" Then TNUME$ = TRIM$(TNUME$) + TRIM$(Str$(NROPRO&))
    Text11.TEXT = TNUME$
    '
End Sub
'
Sub Command6_Click()
    Command6.Enabled = False
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    If DERACCE%("STPRECIB", "Configuración de Funcionamiento Recibos") > 1 Then
       OPRECIB07.Show 1
       Call COMUNI("Reconfiguracion Completada\   \Re-Inicie Ahora la Aplicacion.")
       Call FINAL("")
    End If
    Command6.Enabled = True
End Sub

Private Sub Command7_Click()
    '
    If CTX% = 0 Then
       CTX% = 1
       If ULTREG&("COBRADOR") > 1 Then
          XXX$ = REGSEL$("COBRADOR")
          If Len(XXX$) < 64 Then
             Call FINAL("")
          End If
          Label22.Caption = Left$(TRIM$(Asc(Left$(XXX$, 1))) + " - " + Mid$(XXX$, 2, 30), 19)
          NROPRO& = 0
          Call NURECIBO(NROPRO&)
          Call PRENURE(NROPRO&)
          'Text11.TEXT = "RB." + TRIM$(Str$(NROPRO&))
          Text11.SelStart = 3
          Text11.SelLength = Len(Text11.TEXT) - 3
       End If
    End If
    '
End Sub

Sub Command8_Click()
   Command8.Enabled = False
   RECICO07.Show 1
   '
   NROPRO& = 0
   Call NURECIBO(NROPRO&)
   Call PRENURE(NROPRO&)
   'Text11.TEXT = "RB." + TRIM$(Str$(NROPRO&))
   Text11.SelStart = 3
   Text11.SelLength = Len(Text11.TEXT) - 3
   '
   Command8.Enabled = True
End Sub

Private Sub Command9_Click()
    If DERACCE%("CAMFERE", "Cambiar Fecha del Recibo") >= 2 Then
        Call SELECHO("SELFECHA")
        VDEVV$ = VALACT1$("SELFECHA")
        If VDEVV$ <> "" Then
            Text9.TEXT = VDEVV$
            FF% = FECHANUM(Text9.TEXT)
        End If
    Else
        Text1.SetFocus
    End If
End Sub

Private Sub DESCUEN_Change(Index As Integer)
    '
    Label20.Caption = TRIM$(CSTRING$(MKD$(XVALO(Label14) + XVALO(DESCUEN(0)) - XVALO(Label18)), 4, 14, 2, 2))
    If XVALO(Label20) < 0 Then
          Frame10.Caption = "Descuentos:"
        Else
          Frame10.Caption = "Anticipos:"
    End If
    Call Timer1_Timer
    '
End Sub

Private Sub DESCUEN_GotFocus(Index As Integer)
    If TRIM$(Text1) = "" Then Text1.SetFocus
    NCX1 = XVALO(Text1)
    If NCX1 < 1 Or NCX1 > 32767 Then Text1.SetFocus
    NCX2% = NCX1
    If rasocli$(NCX2%) = "" Then Text1.SetFocus
    DESCUEN(Index) = TRIM$(DESCUEN(Index))
End Sub

Private Sub DESCUEN_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
      Call DESCUEN_LostFocus(Index)
    End If
End Sub

Private Sub DESCUEN_LostFocus(Index As Integer)
    FDESCU$ = TRIM$(FORMATNUM(XVALO(DESCUEN(Index)), "F12.2"))
    PDX = 0
    IDES# = XVALO(DESCUEN(0))
    L14# = XVALO(Label14)
    If L14# >= 1 Then
      PDX = 100 - 100 / (IDES# / L14# + 1)
    End If
    Text7 = TRIM$(FORMATNUM$(PDX, "F6.1"))
    DESCUEN(Index) = FDESCU$
End Sub

Private Sub Form_Load()
    '
    Dim BUFLE As String * 64
    '
    Call CENTRAFORM(Me)
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
    End If
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)

    Frame1.Refresh
    DoEvents
    '\\\OT-07-0029-RG-26/01/07///
    If Control$("RECISTA", "DESCAJA") = "CREDESCU" Then
       'Frame9.Left = 5930
       Frame8.Visible = True
    End If
    '\\\OT-07-0029-RG-FIN///
    '
    MONEMI% = 1
    MONEMIS = DEMONECO$(1)
    TICAMBIS = "1.0000"
    MODOMONEI% = 0
    MODOMONE$ = Left$(TRIM$(UCase$(Control$("GESVENTA", "MODOMONE"))), 8)
    If MODOMONE$ = "MULTIMON" Then
       MODOMONEI% = 2
       Command23.Visible = False
       MONEMIS.Visible = True
       Label24.Visible = True
       TICAMBIS.Visible = True
    ElseIf MODOMONE$ = "BIMONETA" Then
       MODOMONEI% = 1
    End If
    '
10  FORIPRE$ = TRIM$(Control$("", "FORECIST"))
    If FORIPRE$ <> "" Then
      LU$ = LUDAT$
      If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
        RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
        Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Recibo Standard de Cobranza:\  \'" + RMCC$ + "'.")
        If DERACCE%("STPRECIB", "Configuración de Funcionamiento Recibos") > 1 Then OPRECIB07.Show 1
        Call CIERRARCH("*.*")
        Call FINAL("")
      End If
    End If
    '
    CUEDEUVE$ = TRIM$(Control$("AMACLI", "ECUECON"))
    If CUEDEUVE$ = "" Then
      Call MENSERR(24, "Falta Definir Cuenta Contable\Correspondiente a 'Deudores por Ventas'.")
      If DERACCE%("STPRECIB", "Configuración de Funcionamiento Recibos") > 1 Then OPRECIB07.Show 1
      Call CIERRARCH("*.*")
      Call FINAL("")
    End If
    If ECOARCH$("ECUECON", CUEDEUVE$) = "" Then
      Call MENSERR(24, "Falta Definir Cuenta Contable\Correspondiente a 'Deudores por Ventas'.")
      If DERACCE%("STPRECIB", "Configuración de Funcionamiento Recibos") > 1 Then OPRECIB07.Show 1
      Call CIERRARCH("*.*")
      Call FINAL("")
    End If
    '
    CUECHECA$ = TRIM$(Control$("", "CUECHECA"))
    If CUECHECA$ = "" Then
      Call MENSERR(24, "Falta Definir Cuenta Contable\Asociada a la Cartera de Valores.")
      If DERACCE%("STPRECIB", "Configuración de Funcionamiento Recibos") > 1 Then OPRECIB07.Show 1
      Call CIERRARCH("*.*")
      Call FINAL("")
    End If
    If ECOARCH$("ECUECON", CUECHECA$) = "" Then
      Call MENSERR(24, "Falta Definir Cuenta Contable\Asociada a la Cartera de Valores.")
      If DERACCE%("STPRECIB", "Configuración de Funcionamiento Recibos") > 1 Then OPRECIB07.Show 1
      Call CIERRARCH("*.*")
      Call FINAL("")
    End If
    '
    If Frame8.Visible = True Then
       CUEDESCU$ = TRIM$(Control$("", "CUEDESCU"))
       If CUEDESCU$ = "" Then
         Call MENSERR(24, "Falta Definir Cuenta Contable\Asociada a Descuentos de Caja.")
         If DERACCE%("STPRECIB", "Configuración de Funcionamiento Recibos") > 1 Then OPRECIB07.Show 1
         Call CIERRARCH("*.*")
         Call FINAL("")
       End If
       If ECOARCH$("ECUECON", CUEDESCU$) = "" Then
         Call MENSERR(24, "Falta Definir Cuenta Contable\Asociada a Descuentos de Caja.")
         If DERACCE%("STPRECIB", "Configuración de Funcionamiento Recibos") > 1 Then OPRECIB07.Show 1
         Call CIERRARCH("*.*")
         Call FINAL("")
       End If
     End If
     '

20  If ULTREG&("MECOBRA") < 1 Then
      TTXX$ = "Falta Tabla de Medios de Cobranza\Puede Instalar una Tabla Preconfigurada"
      OPX% = COMALTER%(TTXX$ + "\\Instalar Ahora\Carga Manual")
      If OPX% = 1 Then
         NX1% = FILEOPEN%(CurDir$ + "\MECOBRA.DRV", 0, 64)
         For U& = 1 To LOF(NX1%) / 64
           Get #NX1%, U&, BUFLE$
           NMP% = CVI(Left$(BUFLE$, 2))
           If NMP% > 0 Then
             If NMP% < 900 Then
               XXX$ = BUFLE$
               Call REPLA(XXX$, CUECHECA$, 45, 12)
               Call REGAPP("MECOBRA", XXX$)
             End If
           End If
         Next U&
         Call CIERRARCH("MECOBRA")
         Close #NX1%
         If DERACCE%("STPRECIB", "Configuración de Funcionamiento Recibos") > 1 Then OPRECIB07.Show 1
         Call CIERRARCH("*.*")
         Call FINAL("")
       ElseIf OPX% = 2 Then
         If DERACCE%("STPRECIB", "Configuración de Funcionamiento Recibos") > 1 Then OPRECIB07.Show 1
         Call CIERRARCH("*.*")
         Call FINAL("")
      End If
      GoTo 20
    End If
    '
    FF% = HOY(HO$)
    Text9.TEXT = HO$
    '
    If ULTREG&("COBRADOR") > 0 Then
       Label21.Visible = True
       Label22.Visible = True
       Command7.Visible = True
       XXX$ = REGLEIDO$("COBRADOR", 1)
       RECIBO08.Label22.Caption = Left$(TRIM$(Asc(Left$(XXX$, 1))) + " - " + Mid$(XXX$, 2, 30), 19)
    End If
    '
    AUTONUME% = 0
    If Control$("", "NURECIST") = "AUTONUME" Then
      AUTONUME% = 1
    End If
    '
    '
    CtaRetIva$ = Control$("OPRECIBO", "CURETIVA") ' cuenta de retencion de iva
    
    NROPRO& = 0
    Call NURECIBO(NROPRO&)
    Call PRENURE(NROPRO&)
    'Text11.TEXT = "RB." + TRIM$(Str$(NROPRO&))
    Text11.SelStart = 3
    Text11.SelLength = Len(Text11.TEXT) - 3
    '
    If AUTONUME% <> 0 Then
      Text1.TEXT = "     "
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
    End If
    '
'    Refresh
    '
    Call REMPLAGRID(Me, GRID1, Picture1, Label12, List1)
    Call REMPLAGRID(Me, GRID2, Picture2, Label9, List2)
    '
    Call AplicaFuenteAnchoFijo(List1, , 10)
    Call AplicaFuenteAnchoFijo(List2, , 10)
    '
    Call AplicaFuenteAnchoFijo(Text8, [3. Negrita], 9)
    Call AplicaFuenteAnchoFijo(Label30, [3. Negrita], 9)
    '
    Call AplicaFuenteAnchoFijo(Label4, [3. Negrita], 10)
    Call AplicaFuenteAnchoFijo(Label8, [3. Negrita], 10)
    Call AplicaFuenteAnchoFijo(Label10, [3. Negrita], 10)
    Call AplicaFuenteAnchoFijo(Label14, [3. Negrita], 10)
    Call AplicaFuenteAnchoFijo(Label29, [3. Negrita], 10)
    Call AplicaFuenteAnchoFijo(Text7, [3. Negrita], 10)
    Call AplicaFuenteAnchoFijo(DESCUEN(0), [3. Negrita], 10)
    Call AplicaFuenteAnchoFijo(Label18, [3. Negrita], 10)
    Call AplicaFuenteAnchoFijo(Label20, [3. Negrita], 10)

End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BLOQEXE("*")
     Call LIBARCH("*.*")
     Reset
     Call BLOQEXE(App.EXEName)
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
   If EsPorPedido& > 0 Then Unload Me: Exit Sub
   '
   Hide
End Sub

Private Sub GRID1_DblClick()
    If GRID2.Row > 0 Then
       On Error Resume Next
       List1.ListIndex = GRID2.Row - 1
       Call List1_DblClick
       On Error GoTo 0
    End If
End Sub


Private Sub GRID2_DblClick()
    If GRID2.Row > 0 Then
       On Error Resume Next
       List2.ListIndex = GRID2.Row - 1
       Call List2_Click
       On Error GoTo 0
    End If
End Sub

Private Sub Label1_DblClick()
   '\\\OT-07-0189-OD-24/04/08///
   If TRIM$(Text1) = "" Or TRIM$(Text2) = "" Then
     GoTo 99
   End If
   NC% = XVALO(Text1)
   Call RECICO07.VAPLIREC(NC%)
99
  '\\\OT-07-0189-OD-FIN///
End Sub

Private Sub Label10_Change()
   '\\\OT-08-0203-OD-29/04/08///
   PROVENCI# = 0
   FIN = List2.ListCount - 1
   For U& = 0 To FIN
     Z$ = List2.List(U&)
     FEVEN$ = Mid$(Z$, 47, 8)
     IMPRT# = XVALO(Mid$(Z$, 70, 14))
     ResultDientre = DIENTRE(FECHANUM(FEVEN$), HOY(HO$))
     If ResultDientre > 0 Then
       PROVENCI# = PROVENCI + (DIENTRE(FECHANUM(FEVEN$), HOY(HO$)) * IMPRT#)
      Else
       PROVENCI# = PROVENCI + (DIENTRE(HOY(HO$), FECHANUM(FEVEN$)) * IMPRT#) * (-1)
     End If
   Next U&
   TOTAPLI# = XVALO(TRIM$(Label10))
   PROMVEN# = 0
   If Abs(TOTAPLI#) > 0.1 Then
     PROMVEN# = (PROVENCI# / TOTAPLI#)
   End If
   Label17 = FORMATNUM(PROMVEN#, "F4.0")
   If Control("", "OPCVALOR") = "NO" Then
      PORDEX = XVALO(Text7): If PORDEX > 99 Then PORDEX = 99
      If PORDEX < 0 Then PORDEX = 0
      IDES# = XVALO(Label10) * PORDEX / 100
      DESCUEN(0) = TRIM$(CSTRING$(MKD$(IDES#), 4, 14, 2, 2))
   End If
   Call REMPLAGRID(Me, GRID1, Picture1, Label12, List1)
   Call REMPLAGRID(Me, GRID2, Picture2, Label9, List2)
   '\\\OT-08-0203-OD-FIN///
   '
   ' CALCULO DE DIAS DE COBRANZA A PARTIR DE LA FECHA DE VENCIMIENTO
   ' DE LAS FACTURAS CANCELADAS
   If Control$("RECISTA", "DIASVEN") = "FECHAFA" Then
        Call DIASDECOBRANZA_VENCIFACTURA
   End If
   '
End Sub
 
Private Sub Label14_Change()
   '\\\OT-07-0029-RG-26/01/07///
   TT77$ = REPLACAR$(TRIM$(Text7), ",", ".")
   PORDEX = Val(Text7): If PORDEX > 99 Then PORDEX = 99
   If PORDEX < 0 Then PORDEX = 0
   IDES# = Val(Label14) * (100 / (100 - PORDEX) - 1)
   DESCUEN(0) = TRIM$(CSTRING$(MKD$(IDES#), 4, 14, 2, 2))
   '\\\OT-07-0029-RG-FIN///
   Label20 = TRIM$(FORMATNUM$(XVALO(Label14) + XVALO(DESCUEN(0)) - XVALO(Label18), "F12.2"))
End Sub

Private Sub Label17_Change()
Label23 = FORMATNUM$(XVALO(Label17) + XVALO(Label19), "F4.0")
If XVALO(Label23) < 0 Then Label23 = 0
End Sub

Private Sub Label19_Change()
Label23 = FORMATNUM$(XVALO(Label17) + XVALO(Label19), "F4.0")
End Sub

Private Sub Label2_DblClick()
   If DERACCE%("STPRECIB", "Configuración de Funcionamiento Recibos") > 1 Then OPRECIB07.Show 1
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Label22_DblClick()
    '
    If CTX% = 0 Then
       CTX% = 1
       If ULTREG&("COBRADOR") > 1 Then
          XXX$ = REGSEL$("COBRADOR")
          If Len(XXX$) < 64 Then
             Call FINAL("")
          End If
          Label22.Caption = Left$(TRIM$(Asc(Left$(XXX$, 1))) + " - " + Mid$(XXX$, 2, 30), 19)
          NROPRO& = 0
          Call NURECIBO(NROPRO&)
          Call PRENURE(NROPRO&)
          'Text11.TEXT = "RB." + TRIM$(Str$(NROPRO&))
          Text11.SelStart = 3
          Text11.SelLength = Len(Text11.TEXT) - 3
       End If
    End If
    '
End Sub

Private Sub Label26_DblClick()
   Call SELECHO("INDICLIE")
   CPROX$ = TRIM$(VALACT1$("INDICLIE"))
   If CPROX$ <> "" Then
     NCXI% = NUMECLI(CPROX$)
     If NCXI% > 0 Then
       If REGICLI&(NCXI%) > 0 Then
         Label26 = CPROX$
         NC% = NCXI%
         Text1.TEXT = TRIM$(Str$(NC%))
         Text2.TEXT = rasocli$(NC%)
         Text3.TEXT = DOMICLI$(NC%)
         Text4.TEXT = CPOSCLI$(NC%)
         Text5.TEXT = LOCACLI$(NC%)
         Call NotaCli(NC%, 1)
       End If
     End If
   End If
End Sub

Private Sub Label32_DblClick()
   If Abs(XVALO(Label29)) < 1 Then
      Text8 = TRIM$(Label29)
   End If
End Sub

Private Sub Label35_Change()
    Label19 = Label35
End Sub
Private Sub Label8_Change()
    Call Timer1_Timer
End Sub

Private Sub LiReReg_Click()
    Call RECICO07.Command8_Click
End Sub

Private Sub List1_DblClick()
    '
    Call SIGUERECI2
    '
End Sub

Private Sub List2_Click()
    '
    If VACON% <> 0 Then
        Exit Sub
    End If
    '
    Z$ = List2.TEXT
    LIST2ANT$ = Z$
    ILISTA& = List2.ListIndex
    IACRE$ = TRIM$(Mid$(Z$, 70, 14))
    ACREMAX# = CDbl(XVALO(Label14.Caption)) + CDbl(XVALO(DESCUEN(0))) - CDbl(XVALO(Label18.Caption))
    If ACREMAX# < 0 Then ACREMAX# = 0
    IPEN$ = Mid$(Z$, 56, 14)
    If XVALO(IPEN$) > ACREMAX# Then
        IPEN$ = CSTRING$(MKD$(ACREMAX#), 3, 14, 2, 2)
    End If
    SUMACRE# = XVALO(Label18.Caption)
    If IACRE$ = "" Then
         Call REPLA(Z$, IPEN$, 70, 14)
         SUMACRE# = SUMACRE# + CDbl(XVALO(IPEN$))
       Else
         Call REPLA(Z$, Space$(14), 70, 14)
         SUMACRE# = SUMACRE# - CDbl(XVALO(IACRE$))
    End If
    VACON% = 1
    List2.AddItem Z$, ILISTA&
    List2.RemoveItem ILISTA& + 1
    List2.Refresh
    List2.ListIndex = ILISTA&
    VACON% = 0
    '
    Label10.Caption = TRIM$(CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2))
    Label10.Refresh
    Label18.Caption = TRIM$(CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2))
    Label18.Refresh
    '
    '\\\OT-07-0029-RG-26/01/07///
    ANTI# = CDbl(XVALO(Label14.Caption)) + XVALO(DESCUEN(0)) - SUMACRE#
    If ANTI# < 0 Then
          Frame10.Caption = "Descuentos:"
        Else
          Frame10.Caption = "Anticipos:"
    End If
    '\\\OT-07-0029-RG-FIN///
    Label20.Caption = TRIM$(CSTRING$(MKD$(ANTI#), 4, 14, 2, 2))
    Label20.Refresh
    '
    Call REMPLAGRID(Me, GRID2, Picture2, Label9, List2)
End Sub
'

Private Sub List2_DblClick()
    '
    VACON% = 1
    ILIST2& = List2.ListIndex
    TOLIST2& = List2.TopIndex
    '
    Z$ = List2.TEXT
    ILISTA& = List2.ListIndex
    IACRE$ = TRIM$(Mid$(Z$, 70, 14))
    IPEN$ = Mid$(Z$, 56, 14)
    SUMACRE# = XVALO(Label18.Caption)
    SUMACRE# = SUMACRE# - CDbl(XVALO(IACRE$))
    SUMACRE# = SUMACRE# + CDbl(XVALO(Mid$(LIST2ANT$, 70, 14)))
    List2.Visible = False
    List2.AddItem LIST2ANT$, ILISTA&
    List2.RemoveItem ILISTA& + 1
    '
    Call SIGUERECI3
    '
    VACON% = 1
    List2.TopIndex = TOLIST2&
    List2.Visible = True
    If List2.ListCount > ILIST2& Then
          List2.ListIndex = ILIST2&
        Else
          Call CIERRARCH("!CARMECO")
          Call BLANARCH("!CARMECO")
          Call CIERRARCH("!APLIRECI")
          Call BLANARCH("!APLIRECI")
          List1.Clear
          List2.Clear
          Label4.Caption = ""
          Label8.Caption = ""
          Label10.Caption = ""
          Label14.Caption = ""
          Label18.Caption = ""
          Label20.Caption = ""
          '\\\OT-07-0029-RG-26/01/07///
          DESCUEN(0) = ""
          Text7 = ""
          Frame8.Caption = "Descuento (N./Cred)"
          '\\\OT-07-0029-RG-FIN///
          BOTNEXT.Enabled = True
          Command5.Enabled = False
          Text1.SetFocus
    End If
    VACON% = 0
    '
    Call REMPLAGRID(Me, GRID1, Picture1, Label12, List1)
End Sub

Private Sub mnuAjuSalDeu_Click()
    Call RECICO07.Command4_Click
End Sub

Private Sub mnuAnuRec_Click()
    Call RECICO07.Command25_Click
End Sub

Private Sub mnuAplicaVentasConCobranza_Click()

    Call DESHASFECHA(DES%, HAS%, PERIO$)
    If PERIO$ = "" Then Exit Sub
    
    DESDE = FETEXCOR1$(DES%)
    HASTA = FETEXCOR1$(HAS%)
    
    Campos$ = "Fecha/D8;Factura/A18;Razón Social/A32;Importe/F14.2;Recibo/A18;Aplicado/F14.2;Medio/A18;Detalle/A24;Importe/F14.2"
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
    
    Screen.MousePointer = 11
    ConsultaSql$ = "SELECT FACTURA, FECHAFACTURA, CLIENTE, IMPOFACTURA, RECIBO, FECHCONTA AS FECHARECIBO, VALABSCOMP AS IMPORECIBO, DESCRIMOV, IMPOAPLI FROM" _
                & " (SELECT CODICOM_LAR AS FACTURA, FECHCONTA AS FECHAFACTURA, CAJABANC.NUCLIEN AS CLIENTE, VALABSCOMP AS IMPOFACTURA, ISNULL(CODOCAP_LAR,'SIN APLICACION') AS RECIBO, ISNULL(IMPOAPLI,0) AS IMPOAPLI FROM CAJABANC WITH(NOLOCK) LEFT JOIN APLICOBRA WITH(NOLOCK) ON CODICOM_LAR=CODOCOR_LAR WHERE TIPOMOVIM='FVEN' AND (OPCIMOVIM='FC' OR OPCIMOVIM='ND' OR OPCIMOVIM='NC') AND CAJABANC.FECHCONTA >='" & DESDE & "' AND CAJABANC.FECHCONTA <='" & HASTA & "') AS RECIBOS" _
                & " LEFT JOIN CAJABANC WITH(NOLOCK) ON RECIBO=CODICOM_LAR" _
                & " WHERE TIPOMOVIM='CHRECI' OR TIPOMOVIM IS NULL" _
                & " ORDER BY FECHAFACTURA, FACTURA"
    CAITAB1% = 1
    Dim RST As ADODB.Recordset
    Set RST = READSET(ConsultaSql$)
    
    If RST.RecordCount > 0 Then GRILLAGRAL.GRID.Rows = RST.RecordCount + 4
    
    TotalFacturado# = 0
    TotalAplicado# = 0
    With RST
        Do While Not .EOF
        
            CAITAB1% = CAITAB1% + 1
            
            GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 1) = FECHATEX$(CVINT(!FechaFactura))
            GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 2) = SAFETEXT$(!FACTURA)
            GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 3) = rasocli$(!Cliente)
                SIGNO% = 1
                If Mid$(!FACTURA, 1, 2) = "NC" Then
                    SIGNO% = -1
                End If
                ImporteFactura# = !IMPOFACTURA * SIGNO%
            GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 4) = FORMATNUM$(ImporteFactura#, "F14.2")
            GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 5) = SAFETEXT$(!recibo)
            
            GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 6) = FORMATNUM$(!IMPOAPLI, "F14.2")
                DescripcionMovimiento$ = SAFETEXT$(!DESCRIMOV)
            GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 7) = Mid$(DescripcionMovimiento$, 1, 18)
            GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 8) = Mid$(DescripcionMovimiento$, 19)
            GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 9) = FORMATNUM$(!IMPORECIBO, "F14.2")
            
            FacturaActual$ = SAFETEXT$(!FACTURA)
            If FacturaActual$ <> FacturaAnterior$ Then
                TotalFacturado# = TotalFacturado# + ImporteFactura#
                FacturaAnterior$ = FacturaActual$
            Else
                GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 4) = FORMATNUM$(0, "F14.2")
            End If
            
            ReciboActual$ = SAFETEXT$(!recibo)
            ImporteRecibo# = XVALO(!IMPOAPLI)
            If ReciboActual$ <> ReciboAnterior$ Then
                TotalAplicado# = TotalAplicado# + ImporteRecibo#
                ReciboAnterior$ = ReciboActual$
            Else
                GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 6) = FORMATNUM$(0, "F14.2")
            End If
            
            .MoveNext
        Loop
    End With
    
    If TotalFacturado# > 0.005 Then
        CAITAB1% = CAITAB1% + 2
        GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 3) = "TOTALES"
        GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 4) = FORMATNUM$(TotalFacturado#, "F14.2")
        GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 6) = FORMATNUM$(TotalAplicado#, "F14.2")
    End If
            
    On Error Resume Next
    RST.Close
    Set RST = Nothing
    On Error GoTo 0
    Screen.MousePointer = 1
    GRILLAGRAL.Caption = "APLICACION DE VENTAS Y COBRANZA CON DETALLE DE MEDIOS (" & PERIO$ & ")"
    GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.TXTBUSCAR = ""

    GRILLAGRAL.Show 1
    Dim MSF As MSFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(MSF, Campos$, "APVENCO", GRILLAGRAL.Caption, "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
    End If
    Set obj = Nothing
       
End Sub

Private Sub MnuApliCredAnt_Click()
    '\\\OT-07-0189-OD-24/04/08///
    Call RECICO07.Command11_Click
    '\\\OT-07-0189-OD-FIN///
End Sub

Private Sub mnuAsociarReporte_Click()
    Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
End Sub

Private Sub mnuBusquedaCheque_Click()
    Call Command24_Click
End Sub

Private Sub mnuCompensaciones_Click()
    Call RECICO07.Command10_Click
End Sub

Private Sub mnuConfiGral_Click()
    Call Command6_Click
End Sub

Private Sub mnuCotizacionesMoneda_Click()
    Call OPRECIB07.Command7_Click
End Sub

Private Sub mnuCuentaCorriente_Click()
    Call Command4_Click
End Sub

Private Sub mnuDiferenciaCambio_Click()
    Call RECICO07.Command6_Click
End Sub

Private Sub mnuEstadisticasCobranza_Click()
    Call RECICO07.Command20_Click
End Sub

Private Sub mnuImpCre_Click()
    Call RECICO07.Command3_Click
End Sub

Private Sub mnuMaeCli_Click()
    Call Command11_Click
End Sub

Private Sub mnuMaeClie_Click()
    Call Command11_Click
End Sub

Private Sub mnuMediosCobranza_Click()
    Call CIERRARCH("MECOBRA")
    Call OPRECIB07.Command6_Click
End Sub

Private Sub mnuNoCli_Click()
    Call Command3_Click
End Sub

Private Sub mnuOtAccDir_Click()
    Call COMUNI("No Existen Accesos Directos \ para esta Aplicación")
End Sub

Private Sub mnuPlanDeCuentas_Click()
    Call OPRECIB07.Command11_Click
End Sub

Private Sub mnuRecibosConDetalleDeMedios_Click()

    Call DESHASFECHA(DES%, HAS%, PERIO$)
    If PERIO$ = "" Then Exit Sub
    
    DESDE = FETEXCOR1$(DES%)
    HASTA = FETEXCOR1$(HAS%)
    
    RESPUESTA% = COMALTER%("Opciones\\Todos los Recibos\Solo los Aplicados")
    
    Campos$ = "Fecha/D8;Recibo/A18;Cliente/A32;Medio/A18;Detalle/A24;Imp. Parcial/F14.2"
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
    
    Screen.MousePointer = 11
    If RESPUESTA% = 1 Then
        ConsultaSql$ = "select nuclien, fechconta, codicom_lar, valabscomp, descrimov from cajabanc with(nolock) where fechconta>='" & DESDE & "' and fechconta<='" & HASTA & "' and tipomovim='chreci'" _
        & " order by fechconta, codicom_lar"
    ElseIf RESPUESTA% = 2 Then
        ConsultaSql$ = "select distinct cajabanc.nuclien, fechconta, codicom_lar, valabscomp, descrimov from aplicobra with(nolock) inner join cajabanc with(nolock) on codocap_lar=codicom_lar where fechconta>='" & DESDE & "' and fechconta<='" & HASTA & "' and tipomovim='chreci'" _
        & " order by fechconta, codicom_lar"
    End If
    CAITAB1% = 1
    Dim RST As ADODB.Recordset
    Set RST = READSET(ConsultaSql$)
    
    If RST.RecordCount > 0 Then GRILLAGRAL.GRID.Rows = RST.RecordCount + 4
    
    TotalRecibo# = 0
    With RST
        Do While Not .EOF
        
            CAITAB1% = CAITAB1% + 1
            
            GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 1) = FECHATEX$(CVINT(!fechconta))
            GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 2) = SAFETEXT$(!codicom_lar)
            GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 3) = rasocli$(!NUCLIEN)
                DescripcionMovimiento$ = SAFETEXT$(!DESCRIMOV)
            GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 4) = Mid$(DescripcionMovimiento$, 1, 18)
            GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 5) = Mid$(DescripcionMovimiento$, 19)
            GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 6) = FORMATNUM$(!ValAbsComp, "F14.2")
            
            TotalRecibo# = TotalRecibo# + !ValAbsComp
                        
            .MoveNext
        Loop
    End With
    
    If TotalRecibo# > 0.005 Then
        CAITAB1% = CAITAB1% + 2
        GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 3) = "TOTALES"
        GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 6) = FORMATNUM$(TotalRecibo#, "F14.2")
    End If
            
    On Error Resume Next
    RST.Close
    Set RST = Nothing
    On Error GoTo 0
    Screen.MousePointer = 1
    FiltroRecibos$ = "(TODOS"
    If RESPUESTA% = 2 Then FiltroRecibos$ = "(APLICADOS"
    GRILLAGRAL.Caption = "LISTADO DE RECIBOS CON DETALLE DE MEDIOS " & FiltroRecibos$ & " - " & PERIO$ & ")"
    GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.TXTBUSCAR = ""

    GRILLAGRAL.Show 1
    Dim MSF As MSFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(MSF, Campos$, "LIREDEME", GRILLAGRAL.Caption, "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
    End If
    Set obj = Nothing
       
End Sub

Private Sub mnuReimputacionContable_Click()
    Call RECICO07.Command1_Click
End Sub

Private Sub mnuReportesVarios_Click()
    Call RECICO07.AGenRep_Click
End Sub

Private Sub mnuResCuCo_Click()
    Call Command4_Click
End Sub

Private Sub mnuSaInCuCor_Click()
    Call RECICO00.Command5_Click
End Sub

Private Sub mnuretencionIIBBProvincia_Click()
    Call OPRECIB07.Command21_Click
End Sub

Private Sub mnuSalir_Click()
    Call Command2_Click
End Sub

Private Sub mnuTalonariosRecibos_Click()
    Call OPRECIB07.Command10_Click
End Sub

Private Sub mnuVerReimp_Click()
    Call RECICO07.Command23_Click
End Sub

Private Sub MONEMIS_Change()
    '
    '\\\OT-07-0086-RG-06/03/07///
    ENTRATICA.Label6.Caption = MONEMIS.Caption
    ENTRATICA.Text4 = TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4"))
    '\\\OT-07-0086-RG-FIN///
    '
End Sub

Private Sub MONEMIS_DblClick()
  Static CAMON%, DMONE$(), TCMONE()
   ' Exit Sub    ' PORQUE TODAVÍA NO ESTA TERMINADO BIMONETARIO
   If CAMON% < 1 Then
     CAMON% = ULTREG&("TAMONED")
10   If CAMON% < 1 Then
         CAMON% = 1
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         DMONE$(1) = "Pesos"
         TCMONE(1) = 1
       Else
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         U2& = 0
         For U& = 1 To CAMON%
           XX$ = REGLEIDO$("TAMONED", U&)
           DMX$ = TRIM$(Left$(XX$, 8)) + " "
           PPXX% = InStr(DMX$, " ")
           DMX$ = Left$(DMX$, PPXX% - 1)
           '
           If TRIM$(DMX$) <> "" Then
             VHX = CVS(Mid$(XX$, 29, 4)): If VHX < 0.005 Then VHX = 1
             U2& = U2& + 1
             DMONE$(U2&) = DMX$
             TCMONE(U2&) = VHX
           End If
         Next U&
         CAMON% = U2&
         If CAMON% < 1 Then GoTo 10
     End If
   End If
   '
   MONEMIA% = MONEMI%
   MONEMI% = MONEMI% + 1
   If MONEMI% > CAMON% Then MONEMI% = 1
   MONEMIS = DMONE$(MONEMI%)
   MONEMIS.ForeColor = RGB(0, 0, 255)
   MONEMIS.FontSize = 11
   If MONEMI% > 1 Then
      MONEMIS.ForeColor = RGB(255, 0, 0)
      MONEMIS.FontSize = 10
   End If
   MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
   '
   COEFCON = TICAMONE(MONEMIA%) / TICAMONE(MONEMI%)
   TICAMBX# = TICAMONE(MONEMI%)
   TICAMBIS = FORMATNUM(TICAMBX#, "F9.4")
   '
   List1.Clear
   List2.Clear
   Call REMPLAGRID(Me, GRID1, Picture1, Label12, List1)
   Call REMPLAGRID(Me, GRID2, Picture2, Label9, List2)
   Label4.Caption = ""
   Label8.Caption = ""
   Label10.Caption = ""
   Label14.Caption = ""
   Label18.Caption = ""
   Label20.Caption = ""
   DESCUEN(0) = ""
   Text7 = ""
   Frame8.Caption = "Descuento (N./Cred)"
   BOTNEXT.Enabled = True
   '
End Sub




Private Sub Picture4_Click()

End Sub

Private Sub Picture5_Click()

End Sub

Private Sub ReciCobran_Click()
    Call Command1_Click
End Sub

Private Sub Text1_Change()
   'jandy SQL
    Call ClearControls

    List1.Clear
    List2.Clear
    Call REMPLAGRID(Me, GRID2, Picture2, Label9, List2)
    If Picture1.Visible = True Then Call REMPLAGRID(Me, GRID1, Picture1, Label12, List1)
    Label4.Caption = ""
    Label8.Caption = ""
    Label10.Caption = ""
    Label14.Caption = ""
    Label18.Caption = ""
    Label20.Caption = ""
    '\\\OT-07-0029-RG-26/01/07///
    DESCUEN(0) = ""
    Text7 = ""
    Text8 = ""
    Call LIMPPROM
    Frame8.Caption = "Descuento (N./Cred)"
    '\\\OT-07-0029-RG-FIN///
    BOTNEXT.Enabled = True
    '
    If TRIM$(Text1) = "" Then
       A% = PIVACLI%(0)  ' ES PARA REFRESCAR LAS RUTINAS DEL TIPO RASOCLI, PIVACLI, ETC.
    End If
    '
    Command5.Enabled = False
End Sub

Private Sub Text1_DblClick()
    '
    'Call SELECHO("DEUDOR12")
    'NCXI = XVALO(TRIM$(VALACT1$("DEUDOR12")))
    
    NCLI = ListaSeleccionClientes()
    If NCXI > 0 Then
        'NC% = NCXI%
        '\\\OT-06-0607-OD-28/07/08///
        'EVALUA SI LA CUENTA NO ESTA SUSPENDIDA
        '
        If LIMICRE(NCXI) < (-0.5) Then
          Call BorraEncabezado
          Call MENSERR(24, "Cuenta Suspendida.")
          Exit Sub
        End If
        '\\\OT-06-0607-OD-FIN///

        Text1.TEXT = TRIM$(Str$(NCXI))
        Text2.TEXT = rasocli$(NCXI)
        Text3.TEXT = DOMICLI$(NCXI)
        Text4.TEXT = CPOSCLI$(NCXI)
        Text5.TEXT = LOCACLI$(NCXI)
        Label26 = CODICLI$(NCXI)
        Call NotaCli(NCXI, 1)
    End If
    '
End Sub

Private Sub Text1_GotFocus()
    Static CTXX%
    If CTXX% = 0 Then
        Refresh
        CTXX% = 1
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
       If XVALO(Text1.TEXT) > 0 Then
          If XVALO(Text1.TEXT) <= 999999 Then
             NCLI = XVALO(Text1.TEXT)
             '\\\OT-06-0607-OD-28/07/08///
             'EVALUA SI LA CUENTA NO ESTA SUSPENDIDA
             '
              If LIMICRE(NCLI) < (-0.5) Then
                  Call BorraEncabezado
                  Call MENSERR(24, "Cuenta Suspendida.")
                  Exit Sub
              End If
              '\\\OT-06-0607-OD-FIN///

             If rasocli$(NCLI) <> "" Then
                Text2.TEXT = rasocli$(NCLI)
                Text3.TEXT = DOMICLI$(NCLI)
                Text4.TEXT = CPOSCLI$(NCLI)
                Text5.TEXT = LOCACLI$(NCLI)
                Label26 = CODICLI$(NCLI)
                Call NotaCli(NCLI, 1)
                On Error Resume Next
                BOTNEXT.SetFocus
                On Error GoTo 0
                Exit Sub
             End If
          End If
       End If
       '
       Beep
       Text1.TEXT = ""
       '
    End If
    '
End Sub

Private Sub Text10_DblClick()
    Call SELECHO("ECODEP")
    DPX1% = XVALO(VALACT1$("ECODEP"))
    If DPX1% > 0 Then
        DP1% = DPX1%
        Text10.TEXT = TRIM$(Str$(DP1%))
        Text7.TEXT = ECOARCH$("ECODEP", Chr$(DP1%))
    End If
End Sub

Private Sub Text11_GotFocus()
    If AUTONUME% <> 0 Then
      NROPRO& = 0
      Call NURECIBO(NROPRO&)
      Call PRENURE(NROPRO&)
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
    End If
    '
    Text11.SelStart = 3
    On Error Resume Next
    Text11.SelLength = Len(Text11.TEXT) - 3
    On Error GoTo 0
End Sub

Private Sub Text11_LostFocus()
    '
    VDEFF$ = TRIM$(Text11.TEXT)
    Text11.TEXT = VDEFF$
    If Left$(VDEFF$, 3) <> "RB." Then GoTo 90
    NRONUE& = XVALO(Mid$(VDEFF$, 4))
    If NRONUE& > 999999 Then
        Call MENSERR(24, "Número de Recibo Fuera de Rango")
        Call NURECIBO(NROPRO&)
        GoTo 95
    End If
    '
    NROPRO& = NRONUE&
    If NROPRO& > 0 Then
      Call NURECIBO(NROPRO&)
      If NROPRO& <> NRONUE& Then
        Beep
        Call MENSERR(24, "Número no Válido o Pre-existente")
        GoTo 95
      End If
    End If
    Exit Sub
    '
90  Beep
    If FORMALTER.Visible = True Then Exit Sub
    'COMENTE LA SIG. LINEA XQ EL OPRECIB00 NO EXISTE
    'If OPRECIB00.Visible = True Then Exit Sub
    Call MENSERR(24, "Número no Válido")
    NROPRO& = 0
    Call NURECIBO(NROPRO&)
95  Call PRENURE(NROPRO&)
    'Text11.TEXT = "RB." + TRIM$(Str$(NROPRO&))
    On Error Resume Next
    Text11.SelStart = 3
    Text11.SelLength = Len(Text11.TEXT) - 3
    Text11.SetFocus
    On Error GoTo 0
    '
End Sub

Private Sub Text2_gotfocus()
    Text1.SetFocus
End Sub

Private Sub Text3_GotFocus()
    On Error Resume Next
    Text1.SetFocus
    On Error GoTo 0
End Sub

Private Sub Text4_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text5_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text6_Click()
    Text1.SetFocus
End Sub

Private Sub Text6_DBLCLICK()
   If DERACCE%("STPRECIB", "Configuración de Funcionamiento Recibos") > 1 Then OPRECIB07.Show 1
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Text7_Change()
   '
  '\\\OT-07-0029-RG-26/01/07///
   Static OPCIVALO$
   If OPCIVALO$ = "" Then
      OPCIVALO$ = "SI"
      If Control("", "OPCVALOR") = "NO" Then OPCIVALO$ = "NO"
   End If
   '
   If OPCIVALO$ = "NO" Then
      PORDEX = XVALO(Text7): If PORDEX > 99 Then PORDEX = 99
      If PORDEX < 0 Then PORDEX = 0
      IDES# = XVALO(Label10) * PORDEX / 100
      DESCUEN(0) = TRIM$(CSTRING$(MKD$(IDES#), 4, 14, 2, 2))
    Else
      TT77$ = REPLACAR$(TRIM$(Text7), ",", ".")
      PORDEX = Val(Text7): If PORDEX > 99 Then PORDEX = 99
      If PORDEX < 0 Then PORDEX = 0
      IDES# = Val(Label14) * (100 / (100 - PORDEX) - 1)
      DESCUEN(0) = TRIM$(CSTRING$(MKD$(IDES#), 4, 14, 2, 2))
      '\\\OT-07-0029-RG-FIN///
      'Call CONDOLA
   End If
   '
   If Text7 <> "" Then
     If Left$(Frame8.Caption, 7) = "Descuen" Then
       Frame8.Caption = "N./Cred. " + PRONOCRE$
     End If
   End If
   '
End Sub

Private Sub Text7_GotFocus()
    If TRIM$(Text1) = "" Then Text1.SetFocus: Exit Sub
    NCX1 = XVALO(Text1)
    If NCX1 < 1 Or NCX1 > 32767 Then Text1.SetFocus: Exit Sub
    NCX2% = NCX1
    If rasocli$(NCX2%) = "" Then Text1.SetFocus
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call Text7_LostFocus
    End If
End Sub

Private Sub Text7_LostFocus()
    Text7 = FORMATNUM(Text7, "F4.1")
End Sub

Private Sub Text8_Change()
   Call Timer1_Timer
End Sub

Private Sub Text8_LostFocus()
    Text8 = FORMATNUM$(XVALO(Text8), "F10.2")
End Sub

Private Sub Text9_DblClick()
    If DERACCE%("CAMFERE", "Cambiar Fecha del Recibo") >= 1 Then
        Call SELECHO("SELFECHA")
        VDEVV$ = VALACT1$("SELFECHA")
        If VDEVV$ <> "" Then
            Text9.TEXT = VDEVV$
            FF% = FECHANUM(Text9.TEXT)
        End If
    Else
        Text1.SetFocus
    End If
End Sub

Private Sub Text9_GotFocus()

    If DERACCE%("CAMFERE/NOMESS", "Cambiar Fecha del Recibo") < 1 Then
        Text1.SetFocus
    End If

End Sub

Private Sub Text9_LostFocus()

    If DERACCE%("CAMFERE/NOMESS", "Cambiar Fecha del Recibo") >= 1 Then
        VDEVV$ = TRIM$(Text9.TEXT)
        If FECHANUM(VDEVV$) > 0 Then
              FF% = FECHANUM(VDEVV$)
            Else
              Call MENSERR(24, "Fecha no Válida")
              FEX = HOY(HO$)
              Text9.TEXT = HO$
              Text9.SetFocus
        End If
    End If
End Sub

Sub NURECIBO(NROPRO&)
    Label34 = 2: Label34.Refresh
    'ver ernesto fechas
    Static FLIMI, FENUE$, AUTONURE%
    '
    Screen.MousePointer = 11
    '
    If FENUE$ = "" Then
      FENUE$ = Control$("", "FENURECI")
      If FENUE$ = "" Then FENUE$ = "01/01/80"
      FEXI% = FECHANUM(FENUE$)
      'FLIMI = CVDAT(FECHANUM(FENUE$))
      FLIMI = FETEXCOR1$(FEXI%)
      
      'If FLIMI > Now Then FLIMI = CVDAT(HOY(HO$))
      '
      AUTONURE% = 0
      If Control$("", "NURECIST") = "AUTONUME" Then
        AUTONURE% = 1
      End If
    End If
    '
    COBRACT% = 0
    On Error Resume Next
    COBRACT% = XVALO(Label22.Caption)
    On Error GoTo 0
    '
    NROMA& = 0
    
    COBRACT% = XVALO(Label22)
    Call APERBASDAT("CABAN")
    '
    EXISTEX% = 0
    If NROPRO& > 0 Then
      Condi$ = " FechEmisi >= '" & FLIMI & "' AND CodiCom_Cor LIKE 'RB%' AND  NumeComp = " & NROPRO& & "  AND Status < 9  AND NumeTalo = " & COBRACT%
      EXISTEX% = CantRegistros&("CAJABANC", Condi$, "NOMESS")
      If EXISTEX% > 0 Then NROPRO& = 0
    End If
    
    'TOMADO DE LA NUEVA TABLA DE CONTROL DE N. RECIBO ULTRECIBO
    If EXISTE_TRIGGER%("ULTIMORECI_INS") Then
      SQLX$ = "SELECT TOP 1 NUMECOMP FROM ULTRECIBO WITH(NOLOCK) WHERE NumeTalo =" & COBRACT%
      SQLX$ = SQLX$ + " AND  TIPOMOVIM = 'RCOB' AND Codiempr = " & CodiEmp%
      SQLX$ = SQLX$ + " ORDER BY NUMECOMP DESC"
      Set RS1 = READSET(SQLX$)
      NROMA& = XVALO(RS1!NumeComp)
      If NROMA& > 0 Then GoTo 85 ' SI OBTIENE NUMERO SIGUE SI NO VA POR LAS CONSULTAS ANTERIORES
    End If
    'FIN TABLA DE CONTROL ULTRECIBO
    
    'INTERVENCION SP
    SQLX$ = "DECLARE @MAXIMO INT;"
    SQLX$ = SQLX$ + " EXECUTE PROXIMO_N_RECIBO " & COBRACT%
    SQLX$ = SQLX$ + ",'" & FLIMI & "'"
    Set RS = FLEXCONN.Execute(SQLX$)
    NROMA& = XVALO(RS.Fields(0))
    If NROMA& > 0 Then GoTo 85
    'FIN INTERVENCION SP
    
    'CODIGO ANTERIOR
    SQLX$ = "SELECT NumeComp as maximo FROM CajaBanc WITH(NOLOCK)"
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'RCOB' "
    SQLX$ = SQLX$ + " AND CodiCom_Cor LIKE 'RB%' "
    SQLX$ = SQLX$ + " AND FechEmisi >= '" & FLIMI & "'"
    SQLX$ = SQLX$ + " AND NumeTalo = " & COBRACT%
    SQLX$ = SQLX$ + " AND Status < 9 "
    SQLX$ = SQLX$ + " ORDER BY NUMECOMP DESC"
    Dim MIRS As ADODB.Recordset
    'MODIFCADO PARA CAPTURAR ERROR DE TIEMPO DE ESPERA DENTRO DE FUNCION READSET
    'Set MIRS = FLEXCONN.Execute(FILTSQL$(SQLX$))
    Set MIRS = READSET(SQLX$)
    If Not IsNull(MIRS!maximo) Then
        NROMA& = XVALO(MIRS!maximo)
    End If
    MIRS.Close
    Set MIRS = Nothing
    
85  If NROPRO& < 1 Then NROPRO& = 1 + NROMA&
    Screen.MousePointer = 1
    Label34 = 3: Label34.Refresh
    '
End Sub
'
Sub SIGUERECI1()
    '
    If MODORECI% = 0 Then
        MODORECI% = 1       ' por emision
        If Control$("RECISTA", "MODORDE") = "VENCIM" Then
           MODORECI% = 2
        End If
    End If
    '
    NCXI = XVALO(Text1.TEXT)
    If NCXI < 1 Or NCXI > 999999 Then
        Call MENSERR(24, "Numero de Cuenta no Válido")
        Text1.SetFocus
        Exit Sub
    End If
    '
    NCLI = NCXI
    If rasocli$(NCLI) = "" Then
        Call MENSERR(24, "Cuenta no Figura\en Base de Datos de Clientes")
        Text1.SetFocus
        Exit Sub
    End If
    '
    NCX1 = Int(GRUDIVCLI(NCLI))
    If NCX1 <> NCLI Then
      Call MENSERR(24, "Sub-Cuenta Divisional - Utilice " + TRIM$(Str$(NCX1X%)))
      Text1 = TRIM$(Str$(NCX1))
      Text1.SetFocus
      Exit Sub
    End If
    '
    FE$ = TRIM$(Text9.TEXT)
    FF% = FECHANUM(FE$)
    HOI% = HOY(HO$)
    If FF% < 1 Or FF% > HOI% + 1 Then
        Call MENSERR(24, "Fecha no Válida")
        Text9.SetFocus
        Exit Sub
    End If
    '
    NoValidaFechaRecibo$ = TRIM$(Control("RECISTA", "NOVAFECH"))
    If DIENTRE%(FF%, HOI%) > 30 And NoValidaFechaRecibo$ <> "SI" Then
      EPAC$ = TRIM$(UCase$(EMPREAC$))
      If InStr(EPAC$, "NEUMAN") < 1 Or XVALO(Left$(TRIM$(Label22.Caption), 2)) <> 1 Then
        Call MENSERR(24, "Fecha no Válida")
        Text9.SetFocus
        Exit Sub
      End If
    End If
    '
    Call CIERRARCH("*.*")
    '
    Screen.MousePointer = 11
    NCLI = XVALO(Text1.TEXT)
    If NCLI > 0 Then
        Text2.TEXT = rasocli$(NCLI)
        Text3.TEXT = DOMICLI$(NCLI)
        Text4.TEXT = CPOSCLI$(NCLI)
        Text5.TEXT = LOCACLI$(NCLI)
        'Text7.TEXT = TEL1CLI$(NCLI)
        'Text10.TEXT = CSTRING$(MKD$(SALDCLI#(NCLI)), 4, 13, 2, 2)
        Refresh
    End If
    '
100 List1.Clear
    List1.Refresh
    List2.Clear
    List2.Refresh
    SUMADEU# = 0: SUIMPO# = 0: SUSALDO# = 0
    '
    Z$ = Space$(128)
    '
    'Jandy SQL
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM SADEUPEN "
    SQLX$ = SQLX$ + "WHERE NuClien = " & NCLI & " "
    SQLX$ = SQLX$ + "ORDER BY FeVencim ASC;"
    
    'Set FacPenTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    Dim FacPenTemp As ADODB.Recordset
    Set FacPenTemp = New ADODB.Recordset
    FacPenTemp.CursorLocation = adUseClient
25  On Error Resume Next
    FacPenTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0
    '
    With FacPenTemp
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        On Error GoTo 0
        .MoveFirst
        Do While Not .EOF
          TIX$ = UCase$(Left$(!codicom_lar, 2))
          x$ = !codicom_lar
          SG% = (-1)
          TIPOS$ = "Factura": If TIX$ = "AS" Then TIPOS$ = "Ajuste de Saldo"
                              If TIX$ = "ND" Then TIPOS$ = "Nota de Debito"
                              If TIX$ = "NC" Then TIPOS$ = "Nota de Credito": SG% = 1
                              If TIX$ = "CC" Then TIPOS$ = "Comp.Ctas.": SG% = 1 '\\\OT-06-0112-WAR-30/03/06///
                              If TIX$ = "RB" Then TIPOS$ = "Anticipo": SG% = 1 '\\\OT-06-0112-WAR-30/03/06///
                              '
          If TIX$ = "FC" Or TIX$ = "SC" Or TIX$ = "ND" Or TIX$ = "AS" Or TIX$ = "RB" Or TIX$ = "NC" Then
            SALDO# = !SaldDebe - !SALDAcre
            If MODOMONEI% = 2 Then
               SALDOLA# = !SaldDebe_DOLA - !SaldAcre_DOLA
               MONEDOCU% = !Mone_Emis
               If MONEMI% > 1 Then
                 If MONEMI% <> MONEDOCU% Then GoTo 95
                 SALDO# = SALDOLA#
               End If
            End If
            '
            'If SALDO# >= 0.005 Then
            If Abs(SALDO#) >= 0.005 Then
              IMPO# = (-1) * !IBruComp * SG%
              SUMAFAC# = SUMAFAC# + IMPO#
              IMPO# = Abs(IMPO#)
              '
              NROX$ = !codicom_lar
              FEMI = CVINT(!FECHEMISI)
              FEVE = CVINT(!FEVENCIM)
              '
              IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 13, 2, 2)
              IVATEX$ = CSTRING$(MKD$(IVAO#), 3, 13, 2, 2)
              SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 14, 2, 2)
              '
                    REGTEX$ = CSTRING$(MKS$(REG&), 3, 8, 0, 2)
                    IRETEX$ = CSTRING$(MKI$(KKI%), 1, 2, 0, 2)
              '
              Z$ = Space$(128)
              Call REPLA(Z$, NROX$, 1, 18)
              Call REPLA(Z$, FECHATEX$(FEMI), 23, 8)
              Call REPLA(Z$, IMPTEX$, 32, 13)
              Call REPLA(Z$, FECHATEX$(FEVE), 47, 8)
              Call REPLA(Z$, SALTEX$, 56, 14)
              Call REPLA(Z$, REGTEX$, 119, 8)
              Call REPLA(Z$, IRETEX$, 127, 2)
              '
              If MODOMONEI% = 2 Then GoTo 80
              ' ESTO ES PARA PROCESO BIMONETARIO
              MONEMORI% = 0: SALDOLAX# = 0
'              RFF$ = RECFILT$("CARBIMON", 2, MKI$(NCLI))
'              DOCUCABE$ = TRIM$(NROX$)
'              For JKLY& = 1 To Len(RFF$) Step 4
'                RECABI& = CVS(Mid$(RFF$, JKLY&, 4))
'                YCB$ = REGLEIDO$("CARBIMON", RECABI&)
'                If TRIM$(Mid$(YCB$, 13, 18)) = DOCUCABE$ Then
'                  If TRIM$(Mid$(YCB$, 45, 18)) = DOCUCABE$ Then
'                    MONEMORI% = CVI(Mid$(YCB$, 31, 2))
'                    If MONEMORI% > ULTREG&("TAMONED") Then
'                       MONEMORI% = MONEMORI% Mod 6         ' ENJUAGUE ASEMBLI - GRAMY
'                    End If
'                  End If
'                  SALDOLAX# = SALDOLAX# + CVD(Mid$(YCB$, 97, 8)) - CVD(Mid$(YCB$, 105, 8))
'                  Call REPLA(YCB$, MKD$(0), 113, 8)
'                End If
'              Next JKLY&
'              If MONEMORI% > 1 Then
'                SALDO1# = SALDOLAX# * TICAMONE(MONEMORI%)
'                If SALDO1# < 0 Then SALDO1# = 0
'                If SALDO1# > SALDO# Then SALDO1 = SALDO#
'                SALTEX$ = CSTRING$(MKD$(SALDO1#), 3, 14, 2, 2)
'                Call REPLA(Z$, SALTEX$, 56, 14)
'              End If
              GoTo 90
              ' FIN ESTO ES PARA PROCESO BIMONETARIO
              '
80            ' esto es para MULTIMONEDA
'              MONEMORI% = 0: SALDOLAX# = 0: IMPODOL# = 0
'              RFF$ = RECFILT$("CARBIMON", 2, MKI$(NCLI))
'              DOCUCABE$ = TRIM$(NROX$)
'              For JKLY& = 1 To Len(RFF$) Step 4
'                RECABI& = CVS(Mid$(RFF$, JKLY&, 4))
'                YCB$ = REGLEIDO$("CARBIMON", RECABI&)
'                If TRIM$(Mid$(YCB$, 13, 18)) = DOCUCABE$ Then
'                  If TRIM$(Mid$(YCB$, 45, 18)) = DOCUCABE$ Then
'                    MONEMORI% = CVI(Mid$(YCB$, 31, 2))
'                    If MONEMORI% > ULTREG&("TAMONED") Then
'                       MONEMORI% = MONEMORI% Mod 6         ' ENJUAGUE ASEMBLI - GRAMY
'                    End If
'                    IMPODOL# = Abs(CVD(Mid$(YCB$, 97, 8)) - CVD(Mid$(YCB$, 105, 8)))
'                  End If
'                  SALDOLAX# = SALDOLAX# + CVD(Mid$(YCB$, 97, 8)) - CVD(Mid$(YCB$, 105, 8))
'                  Call REPLA(YCB$, MKD$(0), 113, 8)
'                End If
'              Next JKLY&
'              If MONEMORI% <> MONEMI% Then GoTo 95
'              SALDO1# = SALDOLAX#
'              If SALDO1# < 0 Then SALDO1# = 0
'              SALTEX$ = CSTRING$(MKD$(SALDO1#), 3, 14, 2, 2)
'              Call REPLA(Z$, SALTEX$, 56, 14)
'              IMPTEX$ = CSTRING$(MKD$(IMPODOL#), 3, 13, 2, 2)
'              Call REPLA(Z$, IMPTEX$, 32, 13)
'              ' FIN PROCESO MULTIMONEDA
              
90            SUMADEU# = SUMADEU# + XVALO(SALTEX$)
              SUIMPO# = SUIMPO# + XVALO(IMPTEX$)
              '
              List2.AddItem Z$
              '
            End If
          End If
95      .MoveNext
        Loop
      End If
      On Error GoTo 0
    End With
    FacPenTemp.Close
    Set FacPenTemp = Nothing
    '
200 List2.ListIndex = -1
    List2.Refresh
    Label4.Caption = TRIM$(CSTRING$(MKD$(SUIMPO#), 4, 14, 2, 2))
    Label4.Refresh
    Label8.Caption = TRIM$(CSTRING$(MKD$(SUMADEU#), 4, 14, 2, 2))
    Label8.Refresh
    Label10.Caption = ""
    Label10.Refresh
    List2.Visible = True
    List2.SetFocus
    If List2.ListCount > 0 Then
        VACON% = 1
        List2.ListIndex = 0
        VACON% = 0
    End If
    '
    BOTNEXT.Enabled = False
    Call BLANARCH("!CARMECO")
    Screen.MousePointer = 1
    '
    Call REMPLAGRID(Me, GRID1, Picture1, Label12, List1)
    Call REMPLAGRID(Me, GRID2, Picture2, Label9, List2)
    Call SIGUERECI2
    '
End Sub
'

Sub SIGUERECI2()
    '
    HOI% = HOY(HO$)
    NCLI = XVALO(Text1)
    '
200 ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Ingreso de Medios de Pago Codificados"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Picture1.Top + 100))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture1.Left))
    '
    Timer1.Enabled = True
    VTB% = VENTABU%("CARMECOW" + ATRI$)
    Timer1.Enabled = False
    Call CIERRARCH("!CARMECO")
    AAA = ULTREG&("!CARMECO")
    Call Timer1_Timer
    If VTB% < 0 Then
        BOTNEXT.Enabled = True
        Exit Sub
    End If
    '
    'VALIDO SI HAY CUENTAS DE RETENCIONES DE IVA O DE INGRESOS BRUTOS SIN NUMERO DE COMPROBANTE
    CTARETENCIONIVA% = CUEINT%(CtaRetIva$)
    PROVINCIAX$ = PROVCLI$(NCLI)
    XX = CodigoRegimenRetIBXProvincia(PROVINCIAX$, RET_REIB%) 'EL PARAMETRO RET_REIB% ES DE RETORNO
    CtaRetIIBB% = RET_REIB%
    
'    For U1& = 1 To ULTREG&("!CARMECO")
'       TXT$ = REGLEIDO$("!CARMECO", U1&)
'       Tipo% = CVI(Left$(TXT$, 2))
'       CUECO% = CUEMECOB%(Tipo%)
'       BANCO$ = TRIM$(ECOARCH$("MECOBRA", MKI$(Tipo%)))
'
'       NUCHE$ = TRIM$(Mid$(TXT$, 19, 12))
'
'       If NUCHE$ = "" Then
'         If CUECO% = CTARETENCIONIVA% Then
'                Call COMUNI("Cuenta de Retención de Iva (" & BANCO$ & "), Debe Ingresar (en Casillero de Número de Cheque)\El Número de Comprobante")
'                Screen.MousePointer = 1
'                GoTo 200
'         End If
'         If CUECO% = CtaRetIIBB% Then
'                Call COMUNI("Cuenta de Retención de Ingresos Brutos (" & BANCO$ & "), Debe Ingresar (en Casillero de Número de Cheque)\El Número de Comprobante")
'                Screen.MousePointer = 1
'                GoTo 200
'         End If
'
'       End If
'     Next U1&
'
    'HASTA AQUI VALIDACION DE CUENTAS DE RETENCION
    ATXY% = 0
    For U& = 1 To ULTREG&("!CARMECO")
      x$ = REGLEIDO$("!CARMECO", U&)
      PROTER$ = UCase$(TRIM$(Mid$(x$, 43, 2)))
      If PROTER$ = "" Or Left$(PROTER$, 1) = "P" Then
           PROTER$ = "  "
         Else
           PROTER$ = "3 "
      End If
      If TRIM$(PROTER$) <> "" Then ATXY% = 1
    Next U&
    '
210 If ATXY% > 0 Then
      ATRI$ = "/NC"
      ATRI$ = ATRI$ + "/TITUPAN=Ingreso CUIT Cheques de Terceros"
      ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Picture1.Top + 100))
      ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture1.Left))
      '
      VTB% = VENTABU%("CARMECOX" + ATRI$)
      Call CIERRARCH("!CARMECO")
      AAA = ULTREG&("!CARMECO")
      If VTB% < 0 Then
          BOTNEXT.Enabled = True
          Exit Sub
      End If
    End If
    '
    List1.Clear
    SUMAVA# = 0
    PROVALO# = 0
    TOTVALO# = 0
    Label14 = "": Label30 = "": Text8 = ""
    FALTACUIT% = 0
    '
    ' SI ESTA CONFIGURADO PARA QUE CALCULE LOS DIAS DE COBRANZA
    ' A PARTIR DEL VENCIMIENTO DE LA FACTURA, ORDENA LOS VALORES
    ' POR SU VENCIMIENTO
    If Control$("RECISTA", "DIASVEN") = "FECHAFA" Then
        Call COPYSTRU("CARMECO", "CARMECF")
        Call TRACENTRAL("CARMECF", "CARMECO")
        Call FILESORT("CARMECF", "", 7, 7, "ASC")
        Call TRALOCAL("CARMECF", "CARMECO")
    End If
    '
    For U& = 1 To ULTREG&("!CARMECO")
        x$ = REGLEIDO$("!CARMECO", U&)
        If x$ = REBLA$ Then GoTo 290
        MECO% = CVI(Left$(x$, 2))
        BANCO$ = TRIM$(ECOARCH$("MECOBRA", MKI$(MECO%)))
        '
        If BANCO$ = "" Then
            Call MENSERR(24, "Medio de Pago no Válido")
            GoTo 200
        End If
        '
        '\\\OT-08-0622-OD-12/08/08///
        'MEDIOCOBRAN$ = ECOARCH$("MECOBRA", MKI$(MECO%))
        'Valida que no este marcado el medio de cobranza como no Vigiente (N/V)
        If InStr(1, BANCO$, "N/V") > 0 Then
           Call MENSERR(24, "Medio de Pago " + "'" + BANCO$ + "'" + "  No Vigente")
           GoTo 200
        End If
        '\\\OT-08-0622-OD-FIN///
        '
        NUVA$ = Mid$(x$, 19, 12)
        '
        'VERIFICA QUE SEA CHEQUE
        CUECO% = CUEMECOB%(MECO%)
        If CUECO% <> CUEICARCH% Then GoTo 19
        
        'VERIFICA SI EL CHEQUE  FUE INGRESADO
        BANCO1$ = Mid$(BANCO$, 4, 18)
        VeriChek% = YAESTACHEQUE%(BANCO1$, NUVA$)
        If VeriChek% > 0 Then
          TX1$ = "Cheque N.: " + NUVA$ + "  Banco : " + BANCO1$ + " Ya Ingresado"
          TX1$ = TX1$ + "\Ante Cualquier Duda Utilice la Opción\Consulta de Cheques 'B.Ch.'"
          Call MENSERR(24, TX1$)
          GoTo 200
        End If
        '
        'Verifica si el cheque esta repetido
        For i& = 1 To (U& - 1)
          X1$ = REGLEIDO$("!CARMECO", i&)
          If Mid$(X1$, 1, 32) = Mid$(x$, 1, 32) Then
            Call MENSERR(24, "Cheque N.: " + NUVA$ + "  Banco : " + BANCO1$ + " Repetido")
            GoTo 200
          End If
        Next i&
        '
        'BANDERA VALIDACION DE INGRESO DE CUIT
        CUITLIB$ = TRIM$(Mid$(x$, 47, 16))
        If CUITLIB$ = "" Then FALTACUIT% = 1
        '
19      IMVA# = CVD(Mid$(x$, 31, 8))
        IMVABS# = IMVA#: If UCase$(BANCO$) = "EFECTIVO" Then IMVABS# = Abs(IMVA#)
        ' LO ANTERIOR ES PARA ACEPTAR EFECTIVO EN NEGATIVO
        If IMVABS# < 0.005 Then
            Call MENSERR(24, "Falta Importe")
            GoTo 200
        End If
        '
        ' Control de fecha de acreditacion para TODOS ... no solo para cheques
        FEMI = CVI(Mid$(x$, 39, 2))
        FEVE = CVI(Mid$(x$, 41, 2))
        '
        ' ESTO ES PARA EL CASO DE QUE NO SE INGRESA FECHA DE EMISIÓN Y
        ' SI DE VENCIMIENTO - PEDIDO POR ANTONIO DE TANIT EL 16/05/2007
        If Abs(FEMI) < 0.5 And FEVE > 0.5 Then
           FEMII% = FEVE
           If FEMII% > HOI% Then FEMII% = HOI%
           Call REPLA(x$, MKI$(FEMII%), 39, 2)
           Call GRAREG("!CARMECO", x$, U&)
           FEMI = FEMII%
        End If
        '
        If Abs(FEMI) < 0.5 Then
           Call REPLA(x$, MKI$(HOI%), 39, 2)
           Call GRAREG("!CARMECO", x$, U&)
           FEMI = HOI%
        End If
        F1 = FEMI: F2 = HOI%
        If FEMI < HOI% - 32 * 13 Or FEMI > HOI% + 30 Or DIENTRE%(F1, F2) > 365 Then
           Call MENSERR(24, "Fecha de Emisión no Válida")
           GoTo 200
        End If
        '
        FEVE = CVI(Mid$(x$, 41, 2))
        If Abs(FEVE) < 0.5 Then
            Call REPLA(x$, MKI$(HOI%), 41, 2)
            Call GRAREG("!CARMECO", x$, U&)
            FEVE = HOI%
        End If
        F1 = FEVE: F2 = HOI%
        If FEVE < HOI% - 64 Or DIENTRE%(F1, F2) > 30 Or FEVE < FEMI - 0.001 Then
            RESPUESTA% = COMALTER%("Se Ha Detectado que Alguno de los Valores\Posee Fecha de Vencimiento con 30 Días de Vencida\o Menor a la Fecha de Emisión\ \¿Desea Continuar?\\Si, Continuar\No, Cancelar")
            If RESPUESTA% = 1 Then GoTo 250
            GoTo 200
        End If
        '
250     PROTER$ = UCase$(TRIM$(Mid$(x$, 43, 2)))
        If PROTER$ = "" Or Left$(PROTER$, 1) = "P" Then
             PROTER$ = "  "
           Else
             PROTER$ = "3 "
        End If
        Call REPLA(x$, PROTER$, 43, 2)
        Call GRAREG("!CARMECO", x$, U&)
        '
        CLEA% = CVI(Mid$(x$, 45, 2))
        If Mid$(x$, 45, 2) = "  " Then
          Call REPLA(x$, MKI$(24), 45, 2)
          Call GRAREG("!CARMECO", x$, U&)
          CLEA% = 24
        End If
        '
        If TRIM$(NUVA$) = "" Then
         If CUECO% = CTARETENCIONIVA% Then
                Call COMUNI("Cuenta de Retención de Iva (" & BANCO$ & "), Debe Ingresar (en Casillero de Número de Cheque)\El Número de Comprobante")
                Screen.MousePointer = 1
                GoTo 200
         End If
         If CUECO% = CtaRetIIBB% Then
                Call COMUNI("Cuenta de Retención de Ingresos Brutos (" & BANCO$ & "), Debe Ingresar (en Casillero de Número de Cheque)\El Número de Comprobante")
                Screen.MousePointer = 1
                GoTo 200
         End If
         
       End If
        
        '
        If CLEA% < 24 Then CLEA% = 24
        If (CLEA% Mod 24) <> 0 Then
            Call MENSERR(24, "Clearing no Válido")
            GoTo 200
        End If
        Call REPLA(x$, MKI$(CLEA%), 45, 2)
        Call GRAREG("!CARMECO", x$, U&)
        '
        Z$ = Space$(96)
        SUMAVA# = SUMAVA# + IMVA#
        IMPTEX$ = CSTRING$(MKD$(IMVA#), 3, 14, 2, 2)
        Call REPLA(Z$, CSTRING$(MKI$(MECO%), 1, 4, 0, 2), 1, 4)
        Call REPLA(Z$, BANCO$, 7, 14)
        Call REPLA(Z$, NUVA$, 23, 12)
        Call REPLA(Z$, PROTER$, 38, 2)
        Call REPLA(Z$, CSTRING$(MKI$(CLEA%), 1, 4, 0, 2), 42, 4)
        Call REPLA(Z$, FECHATEX$(FEMI), 50, 8)
        Call REPLA(Z$, FECHATEX$(FEVE), 62, 8)
        Call REPLA(Z$, IMPTEX$, 70, 14)
        '\\\OT-08-0203-OD-29/04/08///
        'CALCULO PARA PROMEDIO DE VALORES
        PROVALO# = PROVALO# + (DIENTRE(HOY(HO$), FEVE) * XVALO(IMPTEX$))
        TOTVALO# = TOTVALO# + XVALO(IMPTEX$)
        '
        List1.AddItem Z$
        '
290 Next U&
    '
    ' EN CASO QUE LOS CUIT QUE FALTEN SEAN DE CHEQUES DE TERCEROS CONSULTA
    If FALTACUIT% > 0 And ATXY% = 1 Then
      OPX% = COMALTER%("Se han Detectado Valores (Cheques de Terceros) sin Cuit del Librador\¿Desea Aplicar a Los Mismos El CUIT Del Cliente Activo? \\No, Ingreso Manual\Sí, Aplicar")
      If OPX% <= 1 Then
         GoTo 200 ' VUELVE A LA VENTANA
      End If
    End If
    
    'AGREGA LOS CUIT DEL CLIENTE AUTOMATICAMENTE
    If FALTACUIT% > 0 Then
        CARTVAL% = CUEICARCH%
        For U& = 1 To ULTREG&("!CARMECO")
          x$ = REGLEIDO$("!CARMECO", U&)
          MECO1% = CVI(Left$(x$, 2))
          '
          'VERIFICA QUE SEA CHEQUE
          CUECO1% = CUEMECOB%(MECO1%)
          If CUECO1% = CARTVAL% Then 'SI ES CHEQUE
              CUITLIB$ = TRIM$(Mid$(x$, 47, 16))
              If CUITLIB$ <> "" Then GoTo 300 ' SI YA TIENE PASA AL SIGUIENTE
              CUITLIB$ = CuitCli$(NCLI)
              Call REPLA(x$, CUITLIB$, 47, 16)
              Call GRAREG("!CARMECO", x$, U&)
         End If
300     Next U&
        Call CIERRARCH("!CARMECO")
        If ATXY% = 1 Then GoTo 210 'SI TIENE CHEQUES DE TERCERO VUELVE A LA SEGUNDA VENTANA
        
    End If
    'FIN AGREGADO DE CUIT PARA CHEQUES
    '
    If Abs(TOTVALO#) > 0.001 Then
      Label19 = FORMATNUM(PROVALO# / TOTVALO#, "F4.0")
     Else
      Label19 = 0
    End If
    '
      Label23 = FORMATNUM$(XVALO(Label17) + XVALO(Label19), "F4.0")
    '\\\OT-08-0203-OD-FIN///
    '
    If List1.ListCount > 0 Then
        List1.ListIndex = 0
    End If
    '
    Label30.Caption = TRIM$(CSTRING$(MKD$(SUMAVA#), 4, 14, 2, 2))
    Label30.Refresh
    Label14.Caption = TRIM$(CSTRING$(MKD$(SUMAVA#), 4, 14, 2, 2))
    Label14.Refresh
    '\\\OT-07-0029-RG-26/01/07///
    ANTI# = SUMAVA# + XVALO(DESCUEN(0)) - CDbl(XVALO(Label18.Caption))
    If ANTI# < 0 Then
          Frame10.Caption = "Descuentos:"
        Else
          Frame10.Caption = "Anticipos:"
    End If
    '\\\OT-07-0029-RG-FIN///
    Label20.Caption = TRIM$(CSTRING$(MKD$(ANTI#), 4, 14, 2, 2))
    Label20.Refresh
    '
    If Abs(XVALO(Label29)) < 1 Then
       On Error Resume Next
       Text8.SetFocus
       On Error GoTo 0
    End If
    Call REMPLAGRID(Me, GRID1, Picture1, Label12, List1)
    '
    ' CALCULO DE DIAS DE COBRANZA A PARTIR DE LA FECHA DE VENCIMIENTO
    ' DE LAS FACTURAS CANCELADAS
    If Control$("RECISTA", "DIASVEN") = "FECHAFA" Then
        Call DIASDECOBRANZA_VENCIFACTURA
    End If
    '
End Sub
'

Sub SIGUERECI3()
    '
    '
    If List2.ListCount < 1 Then
        Exit Sub
    End If
    '
    Call BLANARCH("!CARGAPLI")
    For U& = 1 To List2.ListCount
        List2.ListIndex = U& - 1
        Z$ = List2.TEXT
        NROX$ = Mid$(Z$, 1, 20)
        FEMII% = FECHANUM(Mid$(Z$, 23, 8))
        IPED# = CDbl(XVALO(Mid$(Z$, 32, 13)))
        FEVEI% = FECHANUM(Mid$(Z$, 47, 8))
        SALD# = CDbl(XVALO(Mid$(Z$, 56, 14)))
        ACRE# = CDbl(XVALO(Mid$(Z$, 70, 14)))
        DOCU$ = Mid$(Z$, 1, 18)
        REG2& = XVALO(Mid$(Z$, 119, 8))
        IRE% = XVALO(Mid$(Z$, 127, 2))
        '
        Y$ = REGBLAN$("!CARGAPLI")
        Call REPLA(Y$, NROX$, 1, 20)
        Call REPLA(Y$, MKI$(FEMII%), 21, 2)
        Call REPLA(Y$, MKD$(IPED#), 23, 8)
        Call REPLA(Y$, MKI$(FEVEI%), 31, 2)
        Call REPLA(Y$, MKD$(SALD#), 33, 8)
        Call REPLA(Y$, MKD$(ACRE#), 41, 8)
        Call REPLA(Y$, DOCU$, 49, 18)
        Call REPLA(Y$, MKS$(REG2&), 67, 4)
        Call REPLA(Y$, MKI$(IRE%), 71, 2)
        '
        Call REGAPP("!CARGAPLI", Y$)
    Next U&
    Call CIERRARCH("!CARGAPLI")
    '
200 ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Aplicación de Cobranza"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Picture2.Top))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture2.Left))
    '
    VTB% = VENTABU%("CARGAPLI" + ATRI$)
    '
    VACON% = 1
    If VTB% >= 0 Then
        List2.Clear
        For U& = 1 To ULTREG&("!CARGAPLI")
            Y$ = REGLEIDO$("!CARGAPLI", U&)
            NROX$ = Mid$(Y$, 1, 20)
            If TRIM$(NROX$) = "Anticipo no Aplicado" Then
               HOI% = HOY(HO$)
               ACRE# = CVD(Mid$(Y$, 41, 8))
               Call REPLA(Y$, MKI$(HOI%), 31, 2)
               Call REPLA(Y$, MKD$(ACRE#), 23, 8)
               Call REPLA(Y$, MKD$(ACRE#), 33, 8)
               Call GRAREG("!CARGAPLI", Y$, U&)
            End If
            FEMI = CVI(Mid$(Y$, 21, 2))
            IPED# = CVD(Mid$(Y$, 23, 8))
            IMPTEX$ = CSTRING$(MKD$(IPED#), 4, 13, 2, 2)
            FEVE = CVI(Mid$(Y$, 31, 2))
            SALD# = CVD(Mid$(Y$, 33, 8))
            SALTEX$ = CSTRING$(MKD$(SALD#), 4, 14, 2, 2)
            ACRE# = CVD(Mid$(Y$, 41, 8))
            '
            'If ACRE# < 0 Then
            If Abs(ACRE#) >= 0.005 And Sgn(ACRE#) <> Sgn(SALD#) Then
                Call MENSERR(24, "No se Permite Aplicación Negativa")
                GoTo 200
            End If
            If U& > 1 Then
              If Abs(ACRE#) > Abs(SALD#) Then
                Call MENSERR(24, "El Importe Acreditado Debe Ser\Inferior al Saldo del Documento")
                 GoTo 200
              End If
            End If
            ACRTEX$ = CSTRING$(MKD$(ACRE#), 4, 14, 2, 2)
            DOCU$ = Mid$(Y$, 49, 18)
            REG2& = CVS(Mid$(Y$, 67, 4))
            IRE% = CVI(Mid$(Y$, 71, 2))
            REGTEX$ = CSTRING$(MKS$(REG2&), 3, 8, 0, 2)
            IRETEX$ = CSTRING$(MKI$(IRE%), 1, 2, 0, 2)
            '
            Z$ = Space$(128)
            Call REPLA(Z$, NROX$, 1, 20)
            Call REPLA(Z$, FECHATEX$(FEMI), 23, 8)
            Call REPLA(Z$, IMPTEX$, 32, 13)
            Call REPLA(Z$, FECHATEX$(FEVE), 47, 8)
            Call REPLA(Z$, SALTEX$, 56, 14)
            Call REPLA(Z$, ACRTEX$, 70, 14)
            Call REPLA(Z$, REGTEX$, 119, 8)
            Call REPLA(Z$, IRETEX$, 127, 2)
            List2.AddItem Z$
            '
        Next U&
    End If
    '
    SUMACRE# = 0
    For U& = 1 To List2.ListCount
        List2.ListIndex = U& - 1
        Z$ = List2.TEXT
        SUMACRE# = SUMACRE# + CDbl(XVALO(Mid$(Z$, 70, 14)))
    Next U&
    '
    Label10.Caption = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
    Label10.Refresh
    Label18.Caption = TRIM$(CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2))
    Label18.Refresh
    '\\\OT-07-0029-RG-26/01/07///
    Label20.Caption = TRIM$(CSTRING$(MKD$(XVALO(Label14.Caption) + XVALO(DESCUEN(0)) - XVALO(Label18.Caption)), 4, 14, 2, 2))
    '\\\OT-07-0029-RG-FIN///
    Label20.Refresh
    '
    List2.Refresh
    VACON% = 1
    '
    Call REMPLAGRID(Me, GRID2, Picture2, Label9, List2)
    '
End Sub
'

Sub GRARECISTA(OKEY%)
    '
700 OKEY% = 0
    'APRORECIB.Show 1
    'If APRORECIB.Label19 <> "OK" Then Exit Sub
    Screen.MousePointer = 11   ' Grabando Recibo - Espere un Momento, por Favor"
    '
    
    EXISTEACRED% = 0
    Condi$ = "LISTA = 'CONFIMECO'"
    If CantRegistros&("DATASQL", Condi$, "NOMESS") > 0 Then
       EXISTEACRED% = 1
    End If
 TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 1 = " & Now & vbCrLf
 
    NC = XVALO(Text1)
    NoHayFactAnt% = 0: SiHayFactAnt% = 0
    For U& = 1 To List2.ListCount
       Z$ = List2.List(U& - 1)
       DOCU$ = Left$(Z$, 20)
       If DOCU$ <> "Anticipo no Aplicado" Then
          IMPOX# = ROUND#((XVALO(Mid$(Z$, 70, 14))))
          If IMPOX# > 0 Then
            If EsFacturaAnticipo%(NC, DOCU$) > 0 Then
               SiHayFactAnt% = 1
            Else
               NoHayFactAnt% = 1
            End If
          End If
       End If
    Next U&
    If SiHayFactAnt% + NoHayFactAnt% >= 2 Then
       Call COMUNI("Imposible Aplicar Recibo a Facturas Anticipo con Facturas Comunes\Debe Seleccionar Solo un Modo por Vez")
       Screen.MousePointer = 1
       Exit Sub
    End If
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 2 = " & Now & vbCrLf
    COTIMONEM# = XVALO(TICAMBIS)
    NCLIE = XVALO(APRORECIB.Label1)
    
       NROO& = XVALO(Mid$(TRIM$(Text11.TEXT), 4))
    FC% = FECHANUM(TRIM$(APRORECIB.Label6))
    TOTAVAL# = ROUND#(CDbl(XVALO(APRORECIB.Label12)) * COTIMONEM#)
    ANTI# = ROUND#(CDbl(XVALO(APRORECIB.Label14)) * COTIMONEM#)
    COBRACT% = XVALO(APRORECIB.Label3)
    CUECARTE% = XVALO(APRORECIB.Label16)
    '
    RECICORTO$ = TRIM$(APRORECIB.Label8)
    RECILARGO$ = TRIM$(APRORECIB.Label10)
    '
    ' LIMPIA LAS VARIABLEES DE LA FACTURA ELECTRONICA
    NumeroCae$ = ""
    IdentiCae$ = ""
    '
    CoComCor$ = RECICORTO$
    CoComLar$ = RECILARGO$
    Referen$ = rasocli$(NCLIE)
    FeContab = CVDAT(FC%)
    FeAcredi = CVDAT(FC%)
    '
840 ' ***** aqui grabar cheques en CAJABANC.DAT
    NOCLOSE% = 1
    SUDEBE# = 0
    Label34 = 5: Label34.Refresh
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 3 = " & Now & vbCrLf
    For U& = 1 To ULTREG&("!CARMECO")
       x$ = REGLEIDO$("!CARMECO", U&)
       Tipo% = CVI(Left$(x$, 2))
       BANCO$ = TRIM$(Mid$(x$, 3, 16))
       '
       OpMovivi$ = "CH" 'ASIGNO VALORES POR DEFECTO
       VaMovivi$ = ""
       '
       If BANCO$ = "" Then
         BANCO$ = TRIM$(ECOARCH$("MECOBRA", MKI$(Tipo%)))
       End If
       If Left$(UCase$(BANCO$), 3) = "CH." Then
         If Len(BANCO$) > 3 Then
           BANCO$ = Mid$(BANCO$, 4)
         End If
       End If
       '
       If Left$(UCase$(BANCO$), 3) = "TC." Then
        TIPOSS$ = Space$(4)
        If Len(BANCO$) > 3 Then
            MCB$ = FILTERGETRECORD$("MECOBRA", 1, MKI$(Tipo%))
            TIPOSS$ = TRIM$(Mid$(MCB$, 57, 4))
            VaMovivi$ = TIPOSS$
            OpMovivi$ = "TC" ' SE LE QUITO EL PUNTO A TC. EN  LA GRABACION
            BANCO$ = Mid$(BANCO$, 4)
        End If
       End If
       If Left$(UCase$(BANCO$), 3) = "ME." Then
        TIPOSS$ = Space$(4)
        If Len(BANCO$) > 3 Then
            MCB$ = FILTERGETRECORD$("MECOBRA", 1, MKI$(Tipo%))
            TIPOSS$ = TRIM$(Mid$(MCB$, 57, 4))
            VaMovivi$ = TIPOSS$
            OpMovivi$ = "ME" ' MEDIO ELECTRONICO
            BANCO$ = Mid$(BANCO$, 4)
        End If
       End If

       '
       NUCHE$ = TRIM$(Mid$(x$, 19, 12))
       
       IMCHE# = ROUND(COTIMONEM# * CVD(Mid$(x$, 31, 8)))
       FECHE% = CVI(Mid$(x$, 39, 2))
       FEVEN% = CVI(Mid$(x$, 41, 2))
       If EXISTEACRED% > 0 Then
          FEVEN% = FECHAACRED%(Tipo%, FECHE%, FEVEN%)
       End If

       
       CUITLIB$ = TRIM$(Mid$(x$, 47, 16))
       '
       If Abs(IMCHE#) > 0 Then
          SUDEBE# = SUDEBE# + IMCHE#
             FEX = FEVEN%: FEVECHE$ = FECHATEX$(FEX)
          FeContab = CVDAT(FC%)
             CUECO% = CUEMECOB%(Tipo%)
          CuentaCo% = CUECO%
          NumeClie = NCLIE
          NumeProv% = 0
          TiMovi$ = "CHRECI"
          OpMovi$ = OpMovivi$
          VaMovi$ = VaMovivi$
             IDEVA$ = AJUSTI$(BANCO$, 18)
             IDEVA$ = IDEVA$ + AJUSTI$(NUCHE$, 14)
             IDEVA$ = IDEVA$ + FEVECHE$
          Descrip$ = IDEVA$
             If FECHE% < 1 Then FECHE% = FC%
          FechEmi = CVDAT(FECHE%)
          FeAcredi = CVDAT(FEVEN%)
          NumeTal = COBRACT%
'          CodiEmp% = 0
'          CodiSuc% = SUCUI%
          NumeCom = NROO&
             STATUI% = 0
             If CUECO% = CUECARTE% Then STATUI% = 1
          StatInit = STATUI%
          ImpoDeb# = IMCHE#
          Observa$ = ""
            If CUITLIB$ <> "" Then
              Observa$ = "CUIT Libr: " + CUITLIB$
            End If
          Call GRACABAN
       End If
    Next U&
    Label34 = 6: Label34.Refresh
    OKEY% = 6
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 4 = " & Now & vbCrLf
    ' fin graba de cheques en CAJABANC **********************************
    '
    ' AQUI GRABA AJUSTE DE REDONDEO
    IREDON# = XVALO(Text8)
    If Abs(IREDON#) >= 0.005 Then
       CUEREDO$ = TRIM$(Control$("RECISTA", "CUEREDON"))
       If CUEREDO$ <> "" Then CUERED% = CUEINT%(CUEREDO$)
       '
       SUDEBE# = SUDEBE# + IREDON#
          FEX = FC%: FEVECHE$ = FECHATEX$(FEX)
       FeContab = CVDAT(FC%)
       CuentaCo% = CUERED%
       NumeClie = NCLIE
       NumeProv% = 0
       TiMovi$ = "CHRECI"
       OpMovi$ = "REDON"
       VaMovi$ = ""
       Descrip$ = "Ajuste Redondeo"
       FechEmi = CVDAT(FC%)
       FeAcredi = CVDAT(FC%)
       NumeTal = COBRACT%
'       CodiEmp% = 0
'       CodiSuc% = SUCUI%
       NumeCom = NROO&
       StatInit = 0
       ImpoDeb# = IREDON#
       Observa$ = ""
       Call GRACABAN
    End If
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 5 = " & Now & vbCrLf
    '
843 ' acumula anticipo entre detalle y casillero
    For U& = 1 To List2.ListCount
      Z$ = List2.List(U& - 1)
      DOCU$ = Left$(Z$, 20)
      If DOCU$ = "Anticipo no Aplicado" Then
        ANTI# = ANTI# + ROUND#(CDbl(XVALO(Mid$(Z$, 70, 14))) * COTIMONEM#)
        List2.RemoveItem (U& - 1)
        GoTo 843
      End If
    Next U&
    '
    ' graba el anticipo como deuda pendiente
    If ANTI# >= 0.005 Then
      IBruTot# = (-1) * ANTI#
      IIvaTot# = 0
      NuorVen% = 0
      FechEmi = CVDAT(FC%)
      FeVenc = CVDAT(FC%)
      ImVenc# = (-1) * ANTI#
      Call GRADEUPEN
    End If
    '
    OKEY% = 1
    '
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 6 = " & Now & vbCrLf
    
800 ' *********************************** aqui blanquear facturas canceladas
    CoComCor$ = RECICORTO$
    CoComLar$ = RECILARGO$
'    CodiEmp% = 0
'    CodiSuc% = SUC%
    NumeClie = NCLIE
    NumeProv% = 0
    FechEmi = CVDAT(FC%)
    RefApli$ = "Cobranza " + RECICORTO$
    '
    XX% = 0
    VACON% = 1
    SUCANCE# = 0: ENCUENDE% = 0
    For U& = 1 To List2.ListCount
       Z$ = List2.List(U& - 1)
       DOCU$ = Left$(Z$, 20)
       If DOCU$ <> "Anticipo no Aplicado" Then
          CANCE# = ROUND#(COTIMONEM# * CDbl(XVALO(Mid$(Z$, 70, 14))))
          If Abs(CANCE#) > 0 Then
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 7 = " & Now & vbCrLf
             DOCUCA$ = TRIM$(Left$(Z$, 18))
             ICAN# = CANCE#
             SICAN% = Sgn(CANCE#)
             TOTAPLI# = 0
             '
             DocOrCor$ = ""
             TotOrig# = 0
             IvaOrig# = 0
             '
             Call APERBASDAT("CABAN") 'WAR MODIFICADO X CT-101463-FG
             '
             'jandy sql ver order by fechavencimiento y tipo de recordset
             
             strSQL = "Select * from SaDeuPen where Nuclien =" & NCLIE
             strSQL = strSQL & " And CODICOM_LAR = '" & DOCUCA$ & "'"
             If SICAN% > 0 Then
                 strSQL = strSQL & " And SALDDEBE > 0 "
               Else
                 strSQL = strSQL & " And SALDACRE > 0 "
             End If
             '
             Dim SaDeuPen As ADODB.Recordset
             Set SaDeuPen = New ADODB.Recordset
25           On Error Resume Next
             SaDeuPen.Open FILTSQL$(strSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
             If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
               On Error GoTo 0
               Call CapturaErrorOpen
               GoTo 25
             End If
             On Error GoTo 0
             '
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 8 = " & Now & vbCrLf
             With SaDeuPen
             If Not (.EOF And .BOF) Then
               .MoveFirst
               Do While Not (.EOF)
                   ENCUENDE% = 1
                   PARCAN# = ICAN#
                   If SICAN% * (!SaldDebe - !SALDAcre) < SICAN% * ICAN# Then PARCAN# = !SaldDebe - !SALDAcre
                   If PARCAN# > 0 Then
                        !SaldDebe = !SaldDebe - PARCAN#
                      Else
                        !SALDAcre = !SALDAcre + PARCAN#
                   End If
                   .Update
                   TOTAPLI# = TOTAPLI# + PARCAN#   ' SUMA APLICACIONES -  agregado por tema bimonetario 22/03
                   '
                   DocOrCor$ = !CodiCom_Cor
                   TotOrig# = !IBruComp
                   IIVAORIG# = 0
                   On Error Resume Next
                   IvaOrig# = !IIvaComp
                   On Error GoTo 0
                   '
                   ICAN# = ICAN# - PARCAN#
                   If Abs(ICAN#) >= 0.005 Then  'me indico Flavio
                      .MoveNext
                     Else
                       Exit Do
                   End If
                Loop
              End If
             End With
             SaDeuPen.Close
             Set SaDeuPen = Nothing
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 9 = " & Now & vbCrLf
'             With SaDeuPen
'                .FindFirst "NuClien = " & NCLIE & " AND CODICOM_LAR = '" & DOCUCA$ & "' AND SALDDEBE > 0"
'850             If .NoMatch = False Then
'                   ENCUENDE% = 1
'                   PARCAN# = ICAN#: If !SaldDebe < ICAN# Then PARCAN# = !SaldDebe
'                   .Edit
'                   !SaldDebe = !SaldDebe - PARCAN#
'                   .Update
'                   TOTAPLI# = TOTAPLI# + PARCAN#   ' SUMA APLICACIONES -  agregado por tema bimonetario 22/03
'                   '
'                   DocOrCor$ = !CodiCom_Cor
'                   TotOrig# = !IBruComp
'                   IIVAORIG# = 0
'                   On Error Resume Next
'                   IvaOrig# = !IIvaComp
'                   On Error GoTo 0
'                   '
'                   ICAN# = ICAN# - PARCAN#
'                   If ICAN# >= 0.005 Then
'                     .FindNext "Nuclien = " & NumeClie & " AND CODICOM_LAR = '" & DOCUCA$ & "' AND SALDDEBE > 0"
'                     GoTo 850
'                   End If
'                End If
'             End With
             
             '
             DocOrLar$ = TRIM$(Left$(Z$, 18))
             ImpApli# = TOTAPLI#
             SUCANCE# = SUCANCE# + TOTAPLI#
             If XVALO(DESCUEN(0)) < 0.005 Then
                  Call GRAPLICOB
               Else
                  ImpApli# = ROUND#(TOTAPLI# * TOTAVAL# / (TOTAVAL# + XVALO(DESCUEN(0))), 2)
                  Call GRAPLICOB
                  ImpApli# = TOTAPLI# - ImpApli#
                  CoComCor$ = APRORECIB.Label31
                  CoComLar$ = APRORECIB.Label29
                  Call GRAPLICOB
                  ' LO QUE SIQUE ES PARA RESTAURAR LOS VALORES ORIGINALES
                  CoComCor$ = RECICORTO$
                  CoComLar$ = RECILARGO$
             End If
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 10 = " & Now & vbCrLf
             '
             If TOTAPLI# * SICAN% < CANCE# * SICAN% - 0.005 Then
                 ' acumula la diferencia para ajuste de saldo
                 Z1$ = Left$(Z$ + Space$(256), 256)
                 DIFAPLI$ = FORMATNUM$(CANCE# - TOTAPLI#, "F12.2")
                 Call REPLA(Z1$, DIFAPLI$, 241, 12)
                 List2.RemoveItem U& - 1
                 List2.AddItem Z1$, U& - 1
             End If
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 11 = " & Now & vbCrLf
          End If
       End If
    Next U&
    
    VACON% = 0
    '
    Label34 = 7: Label34.Refresh
    ' aqui graba el importe bruto del recibo
    FeContab = CVDAT(FC%)
    '
    CUECO% = CUECOCLI%(NCLIE)
'    If SiHayFactAnt% > 0 Then
'        CUECO1% = CUECOINT%(Control$("SVD451S", "CUEVE-" + TRIM$(Str$(17)))) ' ANTICIPOS A COBRAR
'        If CUECO1% > 0 Then CUECO% = CUECO1%
'    End If
    '
    CuentaCo% = CUECO%
    NumeClie = NCLIE
    NumeProv% = 0
    TiMovi$ = "RCOB"
    OpMovi$ = "CCORR"
    VaMovi$ = ""
    Descrip$ = "Cobranza Cta.Corriente " + RECICORTO$
    FechEmi = CVDAT(FC%)
    FeAcredi = CVDAT(FC%)
    NumeTal = COBRACT%
'    CodiEmp% = 0
'    CodiSuc% = SUCUI%
    NumeCom = NROO&
    StatInit = 0
    ImpoDeb# = (-1) * SUDEBE#
    AtraCance& = XVALO(Label17)
    AtraCheque& = XVALO(Label19)
    Call GRACABAN
    OKEY% = 4
    '
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 12 = " & Now & vbCrLf
    Label34 = 8: Label34.Refresh
    ' aqui graba la nota de credito por descuento
    DESCAJ# = XVALO(DESCUEN(0))
    If DESCAJ# >= 0.005 Then
       GoSub GRANOCRED
    End If
    '
    DIFECAN# = SUDEBE# - SUCANCE# - ANTI#
    If DIFECAN# >= 0.005 Then
       If MODOMONEI% = 2 Then
           IAJU% = CUEINT%(Control$("SVD451S", "CUEVE-9"))
          Else
           IAJU% = CUEINT%(Control$("SVD451S", "CUEVE-11"))
       End If
       If IAJU% < 1 Then
          Call MENSERR(24, "Errores al Grabar este Recibo.\Debe Anularlo y Cargarlo de Nuevo.")
          GoTo 890
       End If
       '
       If MODOMONEI% <> 2 Then
          TTXX$ = "Ocurrio Un Error al Grabar este Recibo.\  \"
          TTXX$ = TTXX$ + "Comunique los Siguientes Valores\"
          TTXX$ = TTXX$ + "a su Soporte de Sistemas:\   \"
          TTXX$ = TTXX$ + "Sudebe = " + TRIM$(FORMATNUM$(SUDEBE#, "F10.4")) + "; "
          TTXX$ = TTXX$ + "Sucance = " + TRIM$(FORMATNUM$(SUCANCE#, "F10.4")) + "; "
          TTXX$ = TTXX$ + "Anti = " + TRIM$(FORMATNUM$(ANTI#, "F10.4")) + "; "
          TTXX$ = TTXX$ + "CotiMonem = " + TRIM$(FORMATNUM$(COTIMONEM#, "F10.4")) + "; "
          TTXX$ = TTXX$ + "Encuende = " + TRIM$(FORMATNUM$(ENCUENDE%, "F10.4")) + "."
          Call MENSERR(24, TTXX$)
          GoTo 49
       End If
       '
       MONEMANT% = MONEMI%
       MONEMI% = 1
       ' graba ajuste diferencia de cambio
       CoComCor$ = "AS" + Mid$(RECICORTO$, 3)
       CoComLar$ = "AS" + Mid$(RECILARGO$, 3)
       FeContab = CVDAT(FC%)
       CUECO% = CUECOCLI%(NCLIE)
'       If SiHayFactAnt% > 0 Then
'           CUECO1% = CUECOINT%(Control$("SVD451S", "CUEVE-" + TRIM$(Str$(17)))) ' ANTICIPOS A COBRAR
'           If CUECO1% > 0 Then CUECO% = CUECO1%
'       End If
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 13 = " & Now & vbCrLf
       For U5& = 1 To List2.ListCount
         DIFE1# = XVALO(Mid$(List2.List(U5& - 1), 241, 12))
         If DIFE1# >= 0.005 Then
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 14 = " & Now & vbCrLf
             CuentaCo% = CUECO%
             NumeClie = NCLIE
             NumeProv% = 0
             TiMovi = "AJUD"            ' AJUSTE DE SALDO DEUDORES
             OpMovi = "AS"              ' AJUSTE DE SALDO
             VaMovi = "CCORR"           ' CUENTA CORRIENTE
             Descrip = "Aj.Saldo - " + rasocli$(NCLIE)
             FechEmi = CVDAT(FC%)
             FeAcredi = CVDAT(FC%)
             NumeTal = COBRACT%
'             CodiEmp% = 0
'             CodiSuc% = SUCUI%
             NumeCom = NROO&
             StatInit = 0
             ImpoDeb# = DIFE1#
             Call GRACABAN
             '
             ' graba contrapartida del ajuste
             CuentaCo% = IAJU%
             TiMovi = "AJUSALD"         ' AJUSTE DE SALDO DEUDORES
             OpMovi = "AS"              ' AJUSTE DE SALDO
             VaMovi = "NO-IMP"          ' NO IMPUTADO
             Descrip = "Aj.Saldo - " + rasocli$(NCLIE)
             ImpoDeb# = (-1) * DIFE1#
             Call GRACABAN
         End If
       Next U5&
       '
       MONEMI% = MONEMANT%
       '
    End If
    '
49  OKEY% = 5
    '
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 15 = " & Now & vbCrLf
    MODOMONE$ = Left$(TRIM$(UCase$(Control$("GESVENTA", "MODOMONE"))), 8)
    If MODOMONE$ <> "BIMONETA" Then
       If MODOMONE$ <> "MULTIMON" Then
          GoTo 880
       End If
    End If
    '
    RFF$ = RECFILT$("CARBIMON", 2, MKI$(NCLIE))
    If Len(RFF$) > 0 Then
      For U& = 1 To List2.ListCount
        x$ = List2.List(U& - 1)
        CANCE# = CDbl(XVALO(Mid$(x$, 70, 14)))
        If CANCE# >= 0.005 Then
          ABUSQUE$ = TRIM$(Left$(x$, 18))
          For JKLX& = 1 To Len(RFF$) Step 4
            RECABI& = CVS(Mid$(RFF$, JKLX&, 4))
            YCB$ = REGLEIDO$("CARBIMON", RECABI&)
            If TRIM$(Mid$(YCB$, 13, 18)) = ABUSQUE$ Or TRIM$(Mid$(YCB$, 45, 18)) = ABUSQUE$ Then
              DOCUCABE$ = TRIM$(Mid$(YCB$, 13, 18))
              For JKLY& = 1 To Len(RFF$) Step 4
                RECABI& = CVS(Mid$(RFF$, JKLY&, 4))
                YCB$ = REGLEIDO$("CARBIMON", RECABI&)
                If TRIM$(Mid$(YCB$, 13, 18)) = DOCUCABE$ Then
                  MONEMORI% = CVI(Mid$(YCB$, 31, 2))
                  TICAMBX# = COTIMONEM#     '    TICAMONE(MONEMORI%)
                  Call REPLA(YCB$, MKI$(FC%), 33, 2)
                  Call REPLA(YCB$, MKI$(NCLIE), 35, 2)
                  Call REPLA(YCB$, Chr$(4), 37, 1)
                  Call REPLA(YCB$, Chr$(1), 38, 1)
                  Call REPLA(YCB$, MKS$(NROO&), 39, 4)
                  Call REPLA(YCB$, MKI$(SUC%), 43, 2)
                  Call REPLA(YCB$, RECILARGO$, 45, 18)
                  Call REPLA(YCB$, MKI$(1), 63, 2) ' RECIBO SIEMPRE EN PESOS
                  Call REPLA(YCB$, MKD$(0), 65, 8)
                  Call REPLA(YCB$, MKD$(ROUND#(CANCE# * COTIMONEM#)), 73, 8)
                  Call REPLA(YCB$, MKD$(0), 81, 8)
                  Call REPLA(YCB$, MKD$(TICAMBX#), 89, 8)
                  Call REPLA(YCB$, MKD$(0), 97, 8)
                  Call REPLA(YCB$, MKD$(CANCE#), 105, 8)
                  Call REPLA(YCB$, MKD$(0), 113, 8)
                     ' ESTO ES PARA QUE SI SE APLICA DOCUMENTO EN PESOS,
                     ' SOLO CANCELE EL SALDO EN PESOS
                     'If MODOMONEI% = 2 And MONEMI% <= 1 Then
                     If (MODOMONEI% = 1 Or MODOMONEI% = 2) And MONEMI% <= 1 Then
                        Call REPLA(YCB$, String$(32, 0), 89, 32)
                     End If
                  Call REGAPP("CARBIMON", YCB$)
                  '
                  DIFE1# = XVALO(Mid$(List2.List(U& - 1), 241, 12))
                  If DIFE1# >= 0.005 Then
                     Call REPLA(YCB$, "AS" + Mid$(RECILARGO$, 3), 45, 18)
                     Call REPLA(YCB$, MKI$(1), 63, 2) ' RECIBO SIEMPRE EN PESOS
                     Call REPLA(YCB$, MKD$(DIFE1#), 65, 8)
                     Call REPLA(YCB$, MKD$(0), 73, 8)
                     Call REPLA(YCB$, MKD$(0), 81, 8)
                     Call REPLA(YCB$, MKD$(TICAMBX#), 89, 8)
                     Call REPLA(YCB$, MKD$(0), 97, 8)
                     Call REPLA(YCB$, MKD$(0), 105, 8)
                     Call REPLA(YCB$, MKD$(0), 113, 8)
                     Call REGAPP("CARBIMON", YCB$)
                     '
                     NumeClie = NCLIE
                     NumeProv% = 0
                     CoComLar = RECILARGO$
                     CoComCor$ = RECICORTO$
                     ImpApli# = DIFE1#
'                     CodiEmp% = 0
'                     CodiSuc% = SUC%
                     RefApli$ = "Cobranza " + RECICORTO$
                     DocOrLar$ = "AS" + Mid$(RECILARGO$, 3)
                     DocOrCor$ = "AS" + Mid$(RECICORTO$, 3)
                     TotOrig# = DIFE1#
                     IvaOrig# = 0
                     Call GRAPLICOB
                     '
                  End If
                  GoTo 835
                End If
              Next JKLY&
            End If
          Next JKLX&
        End If
835   Next U&
    End If
    Call CIERRARCH("CARBIMON")
    '
    ' RECUPERAR EL SALDO EN DOLARES
    For U& = 1 To List2.ListCount
      x$ = List2.List(U& - 1)
      CANCE# = CDbl(XVALO(Mid$(x$, 70, 14)))
      If CANCE# >= 0.005 Then
        ABUSQUE$ = TRIM$(Left$(x$, 18))
        Call RECUSALDOL(ABUSQUE$, CInt(NCLIE))
      End If
    Next U&
    '
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 16 = " & Now & vbCrLf
880 OKEY% = 6 'VOLVER A COMENTARIO(SOLO PARA PROBAR IMPRESION Y REIMPRESION DE REC. COBRANZAS)
'    '
890  'SE MOVIO A RUTINA GRAPLICOB
     '2.-CUANDO ES ASI EN EL CASO QUE QUEDE UN SALDO SE DEBERIA REALIZAR EL MOVIMIENTO TIPO TRANSFERENCIA DE FONDOS POR EL SALDO A DEUDORES POR VENTA.
     'SI ES FACTURA ANTICIPO Y ANTI# > =0.005 THEN GENERAR EL ASIENTO.
'     If SiHayFactAnt% > 0 And ANTI# >= 0.005 Then
'        Call GRABA_ASIENTO_SALDO_RECIBO(NCLIE, ANTI#, RECILARGO$)
'     End If

'GRABO EL ARCHIVO DE TRAZABILIDAD
'For I1% = 1 To 32767
'     ARCH$ = LUDAT$ & "TRAZABR" & TRIM$(Str$(I1%)) & ".TXT"
'     If EXISTE%(ARCH$) < 1 Then
'       NARCHI2% = FILEOPEN%(ARCH$, 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA
'       Print #NARCHI2%, TXTRAZA_BR$ ' ESCRIBE EN EL ARCHIVO EL NUEVO TEXTO
'       Close #NARCHI2%
'       Exit For
'     End If
'Next I1%

    Screen.MousePointer = 1
    
    Call CIERRARCH("*.*")
    Exit Sub
    'Call BAJALDISCO
    '
GRANOCRED:                  ' GRABA NOTA DE CREDITO
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 90 = " & Now & vbCrLf
   NFCY$ = TRIM$(Mid$(Frame8.Caption, 9))
30 CoComCor$ = Left$(NFCY$, 3) + Mid$(NFCY$, 11)
   CoComLar$ = NFCY$
'   CodiEmp% = 0
'   CodiSuc% = 0
   NumeClie = Int(GRUDIVCLI(NCLIE))
   NumeProv% = 0
   FechEmi = CVDAT(FC%)
   Referen$ = rasocli$(NCLIE)
   FeContab = CVDAT(FC%)
   FeAcredi = CVDAT(FC%)
   '
   CuentaCo% = CUECOCLI%(NCLIE)
'   If SiHayFactAnt% > 0 Then
'       CUECO1% = CUECOINT%(Control$("SVD451S", "CUEVE-" + TRIM$(Str$(17)))) ' ANTICIPOS A COBRAR
'       If CUECO1% > 0 Then CuentaCo% = CUECO1%
'   End If
   TiMovi$ = "FVEN"
   OpMovi$ = "NC"
   VaMovi$ = Mid$(NFCY$, 18, 1)
   Descrip$ = "Descuento de Caja " + NFCY$
   NumeTal = XVALO(Mid$(NFCY$, 4, 4))
   NumeCom = XVALO(Mid$(NFCY$, 9, 8))
   StatInit = 0
   ImpoDeb = (-1) * ROUND#(XVALO(DESCUEN(0)) * COTIMONEM#)
   Observa$ = rasocli$(NCLIE)
   Call GRACABAN
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 91 = " & Now & vbCrLf
   '
   NEFA7# = ROUND#((COTIMONEM# * XVALO(DESCUEN(0))) / (1 + PORCEIVA(0) / 100), 2)
   NIVA8# = ROUND#(COTIMONEM# * XVALO(DESCUEN(0))) - NEFA7#
   'graba venta neta - Verificar que QQQ sea la IC de combinacion Cliente-Tipo de Comprobante
   TOTANETX# = NEFA7#
   If TOTANETX# >= 0.005 Then
     CUEDESCU$ = TRIM$(Control$("", "CUEDESCU"))
     CuentaCo% = CUECOINT%(CUEDESCU)
     TiMovi$ = "VENTA"
     StatInit = 0
     ImpoDeb = TOTANETX#
     Call GRACABAN
   End If
'
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 92 = " & Now & vbCrLf
   'graba iva inscripto
   TOTANETX# = NIVA8#
   If TOTANETX# >= 0.005 Then
      CUEIVX$ = Control$("SVD451S", "CUEIVA")
      CuentaCo% = CUECOINT%(CUEIVX$)
      TiMovi$ = "IVA"
      OpMovi$ = "DeFi"
      VaMovi$ = TRIM$(FORMATNUM$(PORCEIVA(0), "F6.2")) + " %"
      Descrip$ = "Iva Debito Fiscal " + NFCX$
      StatInit = 0
      ImpoDeb = TOTANETX#
      BasImpo# = NEFA7#
      Call GRACABAN
   End If
   BasImpo# = 0
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 93 = " & Now & vbCrLf
   '
Return                      ' ***********************************************************
'
End Sub


'
Function RENGLONEC&()
    '
    Call SETULTREG("!PRINRECI", 0)
    Call BLANARCH("!PRINRECI")
    For U& = 1 To ULTREG&("!CARMECO")
      x$ = REGLEIDO$("!CARMECO", U&)
      If x$ = REBLA$ Then GoTo 19
      NUVA$ = Mid$(x$, 19, 12)
      IMVA# = CVD(Mid$(x$, 31, 8))
      FEMII% = CVI(Mid$(x$, 39, 2))
      FEVEI% = CVI(Mid$(x$, 41, 2))
      PROTER$ = UCase$(TRIM$(Mid$(x$, 43, 2)))
      CLEA% = CVI(Mid$(x$, 45, 2))
      '
      TIPOSS$ = Space$(4)
      '
10    Z$ = REGBLAN$("!PRINRECI")
      Call REPLA(Z$, TIPOSS$, 1, 4)
      Call REPLA(Z$, NUVA$, 5, 12)
      Call REPLA(Z$, BANCO$, 17, 16)
      Call REPLA(Z$, MKD$(IMVA#), 33, 8)
      Call REGAPP("!PRINRECI", Z$)
19  Next U&
    '
    V0& = ULTREG&("!PRINRECI")
    V& = 0
    For U& = 1 To List2.ListCount
      x$ = List2.List(U& - 1)
      DOCU$ = Mid$(x$, 1, 18)
      CANCE# = CDbl(XVALO(Mid$(x$, 70, 14)))
      
      
      If Abs(CANCE#) >= 0.005 Then
        V& = V& + 1
        If V& <= V0& Then
              Z$ = REGLEIDO$("!PRINRECI", V&)
           Else
              Z$ = REGBLAN$("!PRINRECI")
        End If
        Call REPLA(Z$, DOCU$, 53, 18)
        Call REPLA(Z$, MKD$(CANCE#), 71, 8)
       
        Call GRAREG("!PRINRECI", Z$, V&)
      End If
    Next U&
    '
    RENGLONEC& = ULTREG&("!PRINRECI")
    Call CIERRARCH("!PRINRECI")
    '
End Function
'
Sub PRIRECISTA()
    '
    Dim ORCOP$(6)
    '
    On Error GoTo 0
    NCLI& = XVALO(Text1)
    '
    ORCOP$(1) = "ORIGINAL"
    ORCOP$(2) = "DUPLICADO"
    ORCOP$(3) = "TRIPLICADO"
    ORCOP$(4) = "CUADRUPLICADO"
    ORCOP$(5) = "QUINTUPLICADO"
    ORCOP$(6) = "SEXTUPLICADO"
    '
    FORIPRE$ = TRIM$(Control$("", "FORECIST"))
    'If FORIPRE$ = "" Then
    '    Call COMUNI("Error de Configuración:\  \Falta Definir Formulario\de Impresión de Recibo.")
    '    Exit Sub
    'End If
    '
    If FORIPRE$ <> "" Then
        LU$ = LUDAT$
        If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
            RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
            Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Recibo de Cobranza:\  \'" + RMCC$ + "'.")
            Exit Sub
        End If
    End If
    '
    CONTA% = ULTREG&("!CARMECO")
    If CONTA% < 1 Then
        Call COMUNI("No se ha Ingresado Detalle de Valores")
        Exit Sub
    End If
    '
    AVAHOJA$ = TRIM$(Control$("RECISTA", "AVAHOJA"))
    CANCELPRINT% = 0
    If Control$("RECISTA", "SUPRINT") = "SI" Then CANCELPRINT% = 1
    '
    Screen.MousePointer = 11
    Call BLANARCH("!PRINRECI")
    For U& = 1 To ULTREG&("!CARMECO")
        x$ = REGLEIDO$("!CARMECO", U&)
        If x$ = REBLA$ Then GoTo 19
        MECO% = CVI(Left$(x$, 2))
           
        BANCO$ = TRIM$(ECOARCH$("MECOBRA", MKI$(MECO%)))
        NUVA$ = Mid$(x$, 19, 12)
        IMVA# = CVD(Mid$(x$, 31, 8))
        FEMII% = CVI(Mid$(x$, 39, 2))
        FEVEI% = CVI(Mid$(x$, 41, 2))
        PROTER$ = UCase$(TRIM$(Mid$(x$, 43, 2)))
        CLEA% = CVI(Mid$(x$, 45, 2))
        CUITLIB$ = TRIM$(Mid$(x$, 47, 16))
        '
        TIPOSS$ = Space$(4)
        If MECO% > 0 Then
            For V& = 1 To ULTREG&("MECOBRA")
                MCB$ = REGLEIDO$("MECOBRA", V&)
                If CVI(Left$(MCB$, 2)) = MECO% Then
                    TIPOSS$ = Mid$(MCB$, 57, 4)
                    GoTo 10
                End If
            Next V&
        End If
        '
10      Z$ = REGBLAN$("!PRINRECI")
        Call REPLA(Z$, TIPOSS$, 1, 4)
        Call REPLA(Z$, NUVA$, 5, 12)
        Call REPLA(Z$, BANCO$, 17, 16)
        Call REPLA(Z$, MKD$(IMVA#), 33, 8)
        Cuenta$ = CODCUE$(CUEMECOB%(MECO%))
        Call REPLA(Z$, Cuenta$, 111, 12)
        FEX = FEVEI%
        If FEX > 0 Then
            Call REPLA(Z$, FECHATEX$(FEX), 41, 8)
        End If
        Call REGAPP("!PRINRECI", Z$)
        '
        If CUITLIB$ <> "" Then
          FORPRICU$ = TRIM$(UCase$(ATRIFORM$(FORIPRE$, "CUIT-TRA/FORMATO")))
          If FORPRICU$ <> "" Then
              URE& = ULTREG&("!PRINRECI")
              Z$ = REGLEIDO$("!PRINRECI", URE&)
              Call REPLA(Z$, CUITLIB$, 79, 16)
              Call GRAREG("!PRINRECI", Z$, URE&)
            Else
              Z$ = REGBLAN$("!PRINRECI")
              Call REPLA(Z$, ".CUIT Librad", 5, 12)
              Call REPLA(Z$, CUITLIB$, 17, 16)
              Call REGAPP("!PRINRECI", Z$)
          End If
        End If
        '
19  Next U&
    '
    If Abs(XVALO(Text8)) >= 0.005 Then ' ajuste redondeo
       Z$ = REGBLAN$("!PRINRECI")
       Call REGAPP("!PRINRECI", Z$)
       Z$ = REGBLAN$("!PRINRECI")
       Call REPLA(Z$, "AJU", 1, 4)
       Call REPLA(Z$, "REDONDEO", 5, 12)
       Call REPLA(Z$, MKD$(XVALO(Text8)), 33, 8)
       Call REGAPP("!PRINRECI", Z$)
    End If
    '
    V0& = ULTREG&("!PRINRECI")
    '
    VACON% = 1
    'List2.Visible = False
    ILISTA& = List2.ListIndex
    TOLISTA& = List2.TopIndex
    '
    V& = 0
    For U& = 1 To List2.ListCount
        List2.ListIndex = U& - 1
        x$ = List2.TEXT
        DOCU$ = Mid$(x$, 1, 18)
        CANCE# = CDbl(XVALO(Mid$(x$, 70, 14)))
        FECEMI$ = Mid(x$, 23, 8)
        FECVEN$ = Mid(x$, 47, 8)
        
        If Abs(CANCE#) >= 0.005 Then
            V& = V& + 1
            If V& <= V0& Then
                  Z$ = REGLEIDO$("!PRINRECI", V&)
               Else
                  Z$ = REGBLAN$("!PRINRECI")
            End If
            Z$ = Left$(Z$ + Space(128), 128)
            Call REPLA(Z$, DOCU$, 53, 18)
            Call REPLA(Z$, MKD$(CANCE#), 71, 8)
            Call REPLA(Z$, FECEMI$, 95, 8)
            Call REPLA(Z$, FECVEN$, 103, 8)
            Call GRAREG("!PRINRECI", Z$, V&)
        End If
    Next U&
    VACON% = 0
    '
    If XVALO(DESCUEN(0)) >= 0.005 Then
        V& = V& + 1
        If V& <= V0& Then
              Z$ = REGLEIDO$("!PRINRECI", V&)
           Else
              Z$ = REGBLAN$("!PRINRECI")
        End If
        NUNOCRE$ = TRIM$(Mid$(Frame8.Caption, 9))
        Call REPLA(Z$, NUNOCRE$, 53, 18)
        Call REPLA(Z$, MKD$((-1) * XVALO(DESCUEN(0))), 71, 8)
        Call GRAREG("!PRINRECI", Z$, V&)
    End If
    '
    V0& = ULTREG&("!PRINRECI")
    '
    TOTAVAL# = CDbl(XVALO(Label14.Caption))
    TOTACAN# = CDbl(XVALO(Label18.Caption)) + CDbl(XVALO(Label20.Caption)) - CDbl(XVALO(DESCUEN(0)))
    ITEXTO$ = NUMTEX$(TOTAVAL#)
    LOTEX% = XVALO(ATRIFORM$("", "LOIMPLET")): If LOTEX% < 20 Then LOTEX% = 20
    Call FRATEX(ITEXTO$, NRX1$, NRX2$, NRX3$, LOTEX%)
    IMPLET1$ = TRIM$(NRX1$) + " " + "***************************************************************************************************************"
    IMPLET2$ = TRIM$(NRX2$) + " " + TRIM$(NRX3$) + " ctvs. *****************************************************************************************"
    '
    ANTI# = CDbl(XVALO(Label20.Caption))
    '
    If Abs(ANTI#) >= 0.005 Then
        '
        V& = V& + 1
        If V& <= V0& Then
             Z$ = REGLEIDO$("!PRINRECI", V&)
           Else
             Z$ = REGBLAN$("!PRINRECI")
        End If
        Call REPLA(Z$, Space$(18), 53, 18)
        Call REPLA(Z$, MKD$(0), 71, 8)
        Call GRAREG("!PRINRECI", Z$, V&)
        '
        V0& = ULTREG&("!PRINRECI")
        '
        V& = V& + 1
        If V& <= V0& Then
             Z$ = REGLEIDO$("!PRINRECI", V&)
           Else
             Z$ = REGBLAN$("!PRINRECI")
        End If
        Call REPLA(Z$, "Anticipo N/Aplic.", 53, 18)
        If ANTI# < 0 Then
            Call REPLA(Z$, "Descuento de Caja", 53, 18)
        End If
        Call REPLA(Z$, MKD$(ANTI#), 71, 8)
        Call GRAREG("!PRINRECI", Z$, V&)
        '
    End If
    '
    Call CIERRARCH("!PRINRECI")
    '
30  If TRIM$(FORIPRE$) <> "" Then
        '
        CODFOR$ = TRIM$(UCase$(FORIPRE$))
        MARIZW = XVALO(ATRIFORM$(CODFOR$, "MARIZQMM"))
        MARSUW = XVALO(ATRIFORM$(CODFOR$, "MARSUPMM"))
        '
        INTERLI = XVALO(ATRIFORM$(CODFOR$, "INTERLIN"))
        If INTERLI < 1 Then
            Call COMUNI("Error de Configuración de Interlinea\en Formulario '" + CODFOR$ + ".FRM'.")
            INTERLI = 1
        End If
        '
        DESMM = XVALO(ATRIFORM$(CODFOR$, "DETDESMM"))
        HASMM = XVALO(ATRIFORM$(CODFOR$, "DETHASMM"))
        '
        
        CACOPI% = XVALO(ATRIFORM$(CODFOR$, "CAN-COP"))
        If CACOPI% < 1 Or CACOPI% > 6 Then
            Call COMUNI("Error de Configuración de Cantidad de Copias\en Formulario '" + CODFOR$ + ".FRM'.")
            CACOPI% = 1
        End If
        '
        ORIENTA% = 1
        PORI$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "ORIENTA")))
        If PORI$ = "PAISAJE" Or PORI$ = "2" Then
            ORIENTA% = 2
        End If
        '
        PORTUREC$ = TRIM$(Control$(IDENTER$, "PORTREC"))
        If PORTUREC$ = "" Then PORTUREC$ = Control$("*", "PORTREC")
        If PORTUREC$ <> "" Then
          NEWPORT$ = PORTUREC$ + "/NOMESS"
          Call SETPRINPORT(NEWPORT$)
          If NEWPORT$ <> "" Then
            Call SETPRINPORT("RESTORE")
          End If
          PRINTERPORT$ = PORTUREC$
        End If
        '
        If TRIM$(PRINTERPORT$) <> "" Then
           Call SETPRINPORT(PRINTERPORT$)
        End If
        '
        PORIAN = Printer.Orientation  ' retiene orientacion
        '
        NRORECI$ = TRIM$(Mid$(Text11.TEXT, 4))
        '
        Screen.MousePointer = 11
        If AVAHOJA$ <> "NEWPAGE" Then
            'Printer.Orientation = 1
            Call SETSPOOL("ORIENTATION", "1")
            If ORIENTA% = 2 Then
                'Printer.Orientation = 2
                Call SETSPOOL("ORIENTATION", "2")
            End If
        End If
        '
        TIPODOC$ = "Recibo ": NUMEDOC$ = NRORECI$

        For COP% = 1 To CACOPI%
            '
            If AVAHOJA$ = "NEWPAGE" Then
               If PRINTERCONTROL% = 1 Then
                  'Printer.NewPage
                  Call SETSPOOL("NEWPAGE", "")
               End If
            End If
            PRINTERCONTROL% = 1
            '
            SUMAVA# = 0: SUMACAN# = 0: NUHOJ% = 0
            Call CARFORWIN
            '
            'AGREGADO PARA QUE IMPRIMA LAS OBSERVACIONES 13/06/12 - OD
            YAC = 0
            For U& = 1 To ULTREG&("!OBSERVOF")
               TXOB$ = TRIM$(REGLEIDO$("!OBSERVOF", U&))
               Call PRINTFORWIN(TXOB$, "OBSERVA", 0, YAC)
               YAC = YAC + INTERLI
            Next U&
            Call CIERRARCH("!OBSERVOF")

            Call PRINTFORWIN(TRIM$(Str$(NCLI&)), "CUEN-CLI", 0, 0)
            Call PRINTFORWIN(TRIM$(rasocli$(NCLI&)), "RASO-CLI", 0, 0)
            Call PRINTFORWIN(DOMICLI$(NCLI&), "DOMI-CLI", 0, 0)
            Call PRINTFORWIN(CPOSCLI$(NCLI&), "CPOS-CLI", 0, 0)
            Call PRINTFORWIN(LOCACLI$(NCLI&), "LOCA-CLI", 0, 0)
            Call PRINTFORWIN(DEPROVI$(PROVCLI$(NCLI&)), "PCIA-CLI", 0, 0)
            Call PRINTFORWIN(CuitCli$(NCLI&), "CUIT-CLI", 0, 0)
            Call PRINTFORWIN(POSIVAL$(NCLI&), "PIVA-CLI", 0, 0)
            '\\\OT-08-0182-OD-21/04/8///
            Call PRINTFORWIN(Str$(VENDCLI%(NCLI&)), "VENDEDOR", 0, 0)
                CPAGO$ = ECOARCH$("CONPAG", Chr$(CPAGCLI%(NCLI&)))
            Call PRINTFORWIN(TRIM$(CPAGO$), "COND-PAG", 0, 0)
            '\\\OT-08-0182-OD-FIN///
            FF% = FECHANUM(TRIM$(Text9.TEXT))
            Call PRINTFORWIN(FETEXTO$(FF%), "FECH-EMI", 0, 0)
            Call PRINTFORWIN(NRORECI$, "NUME-COM", 0, 0)
            Call PRINTFORWIN(ORCOP$(COP%), "ORD-COPI", 0, 0)
            Call PRINTFORWIN(MONEMIS, "MONEDEMI", 0, 0)
            '
            YAC = DESMM
            CANRE1% = ULTREG&("!PRINRECI")
            '
            For YX& = 1 To ULTREG&("!PRINRECI")
               '
               YAC = YAC + INTERLI
               '
               If YAC > HASMM - 2 * INTERLI Then
                   YAC = YAC + INTERLI
                   Call PRINTFORWIN("Transporte", "BANC-CHE", 0, YAC)
                   Call PRINTFORWIN(CSTRING$(MKD$(SUMAVA#), 4, 12, 2, 2), "IMPO-CHE", 0, YAC)
                   Call PRINTFORWIN("Transporte", "DOCU-APL", 0, YAC)
                   Call PRINTFORWIN(CSTRING$(MKD$(SUMACAN#), 4, 12, 2, 2), "IMPO-CAN", 0, YAC)
                   '
                   Call PRINTFORWIN(String$(20, "*"), "TOTA-VAL", 0, 0)
                   Call PRINTFORWIN(String$(20, "*"), "TOTA-CAN", 0, 0)
                   Call PRINTFORWIN(String$(96, "*"), "IMP-LET1", 0, 0)
                   Call PRINTFORWIN(String$(96, "*"), "IMP-LET2", 0, 0)
                   '
                   NUHOJ% = NUHOJ% + 1
                   If NUHOJ% = 1 Then
                     On Error Resume Next
                     CANHOJ& = 1 + Int((CANRE1% / (YX& - 1)) - 0.05)
                     On Error GoTo 0
                   End If
                   Call PRINTFORWIN(TRIM$(Str$(NUHOJ%)), "HOJA-NUM", 0, 0)
                   Call PRINTFORWIN(TRIM$(Str$(CANHOJ&)), "HOJA-TOT", 0, 0)
                   '
                   If AVAHOJA$ = "NEWPAGE" Then
                        'Printer.NewPage
                        Call SETSPOOL("NEWPAGE", "")
                      Else
                        'Printer.EndDoc
                        Call SETSPOOL("ENDDOC", "")
                   End If
                   '
                   Call CARFORWIN
                   Call PRINTFORWIN(TRIM$(Str$(NCLI&)), "CUEN-CLI", 0, 0)
                   Call PRINTFORWIN(TRIM$(rasocli$(NCLI&)), "RASO-CLI", 0, 0)
                   Call PRINTFORWIN(DOMICLI$(NCLI&), "DOMI-CLI", 0, 0)
                   Call PRINTFORWIN(CPOSCLI$(NCLI&), "CPOS-CLI", 0, 0)
                   Call PRINTFORWIN(LOCACLI$(NCLI&), "LOCA-CLI", 0, 0)
                   Call PRINTFORWIN(CuitCli$(NCLI&), "CUIT-CLI", 0, 0)
                   Call PRINTFORWIN(POSIVAL$(NCLI&), "PIVA-CLI", 0, 0)
                   FEX = FF%
                   Call PRINTFORWIN(FETEXTO$(FF%), "FECH-EMI", 0, 0)
                   Call PRINTFORWIN(NRORECI$, "NUME-COM", 0, 0)
                   '
                   YAC = DESMM + INTERLI
                   Call PRINTFORWIN("Transporte", "BANC-CHE", 0, YAC)
                   Call PRINTFORWIN(CSTRING$(MKD$(SUMAVA#), 4, 12, 2, 2), "IMPO-CHE", 0, YAC)
                   Call PRINTFORWIN("Transporte", "DOCU-APL", 0, YAC)
                   Call PRINTFORWIN(CSTRING$(MKD$(SUMACAN#), 4, 12, 2, 2), "IMPO-CAN", 0, YAC)
                   YAC = YAC + 2 * INTERLI
                   '
                End If
                '
                YY$ = REGLEIDO$("!PRINRECI", YX&)
                Call PRINTFORWIN(Left$(YY$, 4), "TIPO-VAL", 0, YAC)
                Call PRINTFORWIN(Mid$(YY$, 5, 12), "NUME-CHE", 0, YAC)
                Call PRINTFORWIN(Mid$(YY$, 17, 16), "BANC-CHE", 0, YAC)
                Call PRINTFORWIN(CSTRING$(Mid$(YY$, 33, 8), 4, 12, 2, 2), "IMPO-CHE", 0, YAC)
                SUMAVA# = SUMAVA# + CVD(Mid$(YY$, 33, 8))
                Call PRINTFORWIN(CSTRING$(Mid$(YY$, 41, 8), 6, 8, 0, 1), "FECH-CHE", 0, YAC)
                Call PRINTFORWIN(Mid$(YY$, 53, 18), "DOCU-APL", 0, YAC)
                Call PRINTFORWIN(CSTRING$(Mid$(YY$, 71, 8), 4, 12, 2, 2), "IMPO-CAN", 0, YAC)
                SUMACAN# = SUMACAN# + CVD(Mid$(YY$, 71, 8))
                Call PRINTFORWIN(TRIM$(Mid$(YY$, 79, 16)), "CUIT-TRA", 0, YAC)
                  FEMIS$ = TRIM$(Mid$(YY$, 95, 8))
                Call PRINTFORWIN(FEMIS$, "MEDIDA-1", 0, YAC) 'Fecha de Emision
                  FEMVE$ = TRIM$(Mid$(YY$, 103, 8))
                Call PRINTFORWIN(FEMVE$, "MEDIDA-2", 0, YAC) 'Fecha de Vencimiento
                Cuenta$ = TRIM$(Mid$(YY$, 111, 12))
                Call PRINTFORWIN(Cuenta$, "DEN-CUEN", 0, YAC) 'Fecha de Vencimiento
                  CUE$ = TRIM(CUEMADRE(XVALO(Text1))): DENO$ = TRIM(DENOCUE$(CUEINT(CUE$)))
                  DET$ = CUE$ & " - " & DENO$    'Cuenta Madre y denominacion de la misma
                Call PRINTFORWIN(DET$, "DET-PASE", 0, YAC)
                '
                'IMPRESION DE SALDO DE FACTURA
                NUMFACTU$ = TRIM$(Mid$(YY$, 53, 18))
                TIPX$ = UCase$(Left$(NUMFACTU$, 2))
                SALDO1# = 0
                If TIPX$ = "ND" Or TIPX$ = "FC" Then
                   Condi$ = "CODICOM_LAR = '" & NUMFACTU$ & "'"
                   SALDO1# = XVALO(VALOBADA("SADEUPEN", "SALDDEBE", Condi$, "NOMESS"))
                   Call PRINTFORWIN(CSTRING$(MKD$(SALDO1#), 4, 12, 2, 2), "IMPO-PEN", 0, YAC)
                End If
                '
                '//OT-8-870-09/10/08
                For i = 0 To List2.ListCount - 1
                  TXTX$ = List2.List(i)
                  DOCUMENTO$ = Mid$(YY$, 53, 18)
                  If Mid$(TXTX$, 1, 18) = DOCUMENTO$ Then
                    FECHADOCU$ = Mid$(TXTX$, 23, 8)
                    Call PRINTFORWIN(FECHADOCU$, "FECH-REC", 0, YAC)
                    Exit For
                  End If
                Next i
                '//OT-8-870-FIN
                '
            Next YX&
            '
            ENTRAOBS.Text28 = "" ' LIMPIO EL TEXTO DE OBSERVACIONES (QUEDABA CON LO ANTERIOR)

            Call PRINTFORWIN(CSTRING$(MKD$(TOTAVAL#), 4, 12, 2, 2), "TOTA-VAL", 0, 0)
            Call PRINTFORWIN(CSTRING$(MKD$(TOTACAN#), 4, 12, 2, 2), "TOTA-CAN", 0, 0)
            Call PRINTFORWIN(IMPLET1$, "IMP-LET1", 0, 0)
            Call PRINTFORWIN(IMPLET2$, "IMP-LET2", 0, 0)
            '
            NUHOJ% = NUHOJ% + 1
            If NUHOJ% = 1 Then
              On Error Resume Next
              CANHOJ& = 1 + Int((CANRE1% / (YX& - 1)) - 0.05)
              On Error GoTo 0
            End If
            Call PRINTFORWIN(TRIM$(Str$(NUHOJ%)), "HOJA-NUM", 0, 0)
            Call PRINTFORWIN(TRIM$(Str$(CANHOJ&)), "HOJA-TOT", 0, 0)
            '
            If AVAHOJA$ <> "NEWPAGE" Then
                 'Printer.EndDoc
                 Call SETSPOOL("ENDDOC", "")
            End If
            '
        Next COP%
        '
        Call CIERRARCH("*.*")
        DESCRIDOC$ = "Recibo Nro. " + NRORECI$ + " - " + TRIM$(rasocli$(NCLI&))
        FEX = FF%
        DESCRIDOC$ = DESCRIDOC$ + " - " + FECHATEX$(FEX)
        
        'TIPODOC$ LO LEO ARRIBA POR QUE SI NO CUANDO PRESENTA LA VENTANA DE LAS OBSERVACIONES
        'TOMA EL DEL REMITO QUE ES EL ULTIMO DATO ASIGNADO
'        TIPODOC$ = "Recibo ": NUMEDOC$ = NRORECI$
        '
        Call SAVESPOOL(DESCRIDOC$, OKX%)
        If TRIM$(PRINTERPORT$) <> "" Then
          Call SETPRINPORT("RESTORE")
        End If
        '
        Screen.MousePointer = 1
        If COMALTER%("Recibo Numero\'" + NRORECI$ + "'\Emitido y Registrado\\Repetir Impresion\Otro Recibo") = 1 Then GoTo 30
        '
        If AVAHOJA$ <> "NEWPAGE" Then
            'Printer.Orientation = PORIAN
            Call SETSPOOL("ORIENTATION", Str$(PORIAN))
        End If
        '
    End If
    Call CIERRARCH("*.*")
    '
End Sub

Sub APLIPAG(FE%, NC%, TI%, VA%, PRF%, NROF&, IMAPLI#, REFE$, DOCORI$, ITOTORI#, IVAORI#)
    '
    ARAPLI$ = "APLICOB": If NC% < 0 Then ARAPLI$ = "APLIPAG"
    PRFS$ = TRIM$(Str$(PRF%)): While Len(PRFS$) < 4: PRFS$ = "0" + PRFS$: Wend
    NROS$ = TRIM$(Str$(NROF&)): While Len(NROS$) < 8: NROS$ = "0" + NROS$: Wend
    '
    TIX% = TI% Mod 10
    TINUVA$ = "FC"                                               ' por default
    If TIX% = 0 Then TINUVA$ = "AS"
    If TIX% = 1 Then TINUVA$ = "VC"
    If TIX% = 3 Then TINUVA$ = "ND"
        If TIX% = 4 Then
        TINUVA$ = "RB"
        If NC% < 0 Then
            TINUVA$ = "OP"
            If TI% = 14 Then TINUVA$ = "FF"
        End If
    End If
    '
    If TIX% = 5 Then TINUVA$ = "DC"
    If TIX% = 6 Then TINUVA$ = "NC"
    If TIX% = 7 Then TINUVA$ = "DP"
    '
    VAD$ = Chr$(64 + VA%): If VAD$ < "A" Or VAD$ > "Z" Then VAD$ = "X"
    TINUVA$ = TINUVA$ + "-" + VAD$ + "-" + PRFS$ + "-" + NROS$
    '
    RECORD$ = MKI$(FE%)
    RECORD$ = RECORD$ + MKI$(Abs(NC%))
    '
    RECORD$ = RECORD$ + TINUVA$
    RECORD$ = RECORD$ + MKD$(IMAPLI#)
    '
    REFE$ = Left$(REFE$ + Space$(24), 24)
    RECORD$ = RECORD$ + REFE$
    '
    RECORD$ = RECORD$ + DOCORI$ + MKD$(ITOTORI#) + MKD$(IVAORI#)
    '
    RECORD$ = RECORD$ + String$(8, 0)
    '
    Call REGAPP(ARAPLI$, RECORD$)
    'Call CIERRARCH(ARAPLI$)
    '
End Sub
'

Private Sub TICAMBIS_Change()
   '
   If TICAMBIS.Enabled = True Then
      TICACH(MONEMI%) = XVALO(TICAMBIS)
   End If
   TICAMBX# = 1
   If MONEMI% > 1 Then
     TICAMBX# = XVALO(TICAMBIS)
   End If
   '
End Sub

Private Sub TICAMBIS_DblClick()
   'Call CARTAMONED
   '
   '\\\OT-07-0086-RG-06/03/07///
   If MONEMI% > 1 Then
     If DERACCE%("COTILOCA", "Ajustar Tipo de Cambio Local") >= 2 Then
        ENTRATICA.Frame4.Caption = "Cotización Local para " + MONEMIS.Caption
        ENTRATICA.Text4 = FORMATNUM(XVALO(TICAMBIS.Caption), "F10.4")
        ENTRATICA.Show 1
        If TRIM$(ENTRATICA.Text4) <> "" Then
          TICAMBIS = TRIM$(ENTRATICA.Text4)
        End If
     End If
     List1.Clear
     List2.Clear
     Label4.Caption = ""
     Label8.Caption = ""
     Label10.Caption = ""
     Label14.Caption = ""
     Label18.Caption = ""
     Label20.Caption = ""
     DESCUEN(0) = ""
     Text7 = ""
     Frame8.Caption = "Descuento (N./Cred)"
     BOTNEXT.Enabled = True
   End If
   Call REMPLAGRID(Me, GRID1, Picture1, Label12, List1)
   '\\\OT-07-0086-RG-FIN///
   '
End Sub

Sub CARTAMONED()
   '
   If DERACCE%("TICAMBIO", "Ajustar Tipo de Cambio Central") < 2 Then
     Exit Sub
   End If
   '
   If ULTREG&("TAMONED") > 6 Then
      Call SETULTREG("TAMONED", 6)
      Call COMUNI("Tabla de Monedas Truncada\al Máximo de 6 Registros")
   End If
   '
   Call TRALOCAL("TAMONED", "")
   VTB% = VENTABU%("TAMONED/NC/TITUPAN=Monedas y Tipo de Cambio")
   If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!TAMONED")
         x$ = REGLEIDO$("!TAMONED", U&)
         Call REPLA(x$, Chr$(U&), 1, 1)
         Call GRAREG("TAMONED", x$, U&)
      Next U&
      FIN& = ULTREG&("!TAMONED")
      Call SETULTREG("TAMONED", FIN&)
   End If
   Call CIERRARCH("TAMONED")
   '
End Sub

Function PRONOCRE$()
   
   'VER ERNESTO flimi MODORAPIDO
   
   '
   Static FLINUMERA%
   Static NUMECORR%
   Static MONUMERA$
   '
   Screen.MousePointer = 11
   If FLINUMERA% = 0 Then
      FLINUMERA% = FECHANUM("01/01/99")
      If ULTREG&("FORMFAC") > 0 Then
         FFS$ = REGLEIDO$("FORMFAC", 2&)
         FLINUMERA% = CVI(Mid$(FFS$, 19, 2))
      End If
   End If
   '
   If NUMECORR% = 0 Then
      NUMECORR% = 1
      QTRX$ = Control$("NUMEFACT", "NUMECORR")
      If QTRX$ = "INDEPEN" Then
           NUMECORR% = (-1)
         ElseIf QTRX$ = "UNIFICA" Then
           NUMECORR% = (-2)
      End If
   End If
   '
   MONUMERA$ = "AUTO"
   '
   TIPDOCX% = 6
   OPMOVIX$ = "NC"
   '
   VACREDI$ = "A"
   NCLIE% = XVALO(Text1)
   PIVACRE% = PIVACLI%(NCLIE%)
   If PIVACRE% = 0 Or PIVACRE% = 3 Or PIVACRE% = 4 Or PIVACRE% = 6 Then
        VACREDI$ = "B"
      ElseIf PIVACRE% = 5 Then
        VACREDI$ = "X"
        If TRIM$(LETREXPO$) <> "" Then
          VACREDI$ = LETREXPO$
        End If
   End If
   '
   DATEDOC% = FECHANUM(Text9.TEXT)
   If DATEDOC% < FLINUMERA% Then
      Call MENSERR(24, "Imposible Registrar en la Fecha Elegida.\  \La Fecha de Recibo Debe Ser POSTERIOR\a la Fecha de Nueva Numeración.")
      Call CIERRARCH("*.*")
      Call FINAL("")
   End If
   FLINUME = CDbl(CVDAT(FLINUMERA%))
   '
   FEXUL = DATEDOC%
   FLIDIAS = CDbl(CVDAT(DIASANTES(FEXUL, 15)))
   '//////////*******NOTA*******////////
   'Con posterioridad se detecto un Error en la Impòrtacion
   '(VARIMOVIM(A,B), SE IMPORTO EN VEZ DE LA LETRA CON EL OPCIMOVIX$(ND,NC).
   'Por tal motivo el sistema no toma en cuenta los numero
   'de NC-ND inferiores al 01/11/04.
   'Pero si lo hace a futuro. es decir yo puedo reiniciar numeración
   'al 01/11/04
   'jandy sql
   
   'FLIMI = FLIDIAS: MODORAPIDO% = 1
   'If FLINUME > FLIMI Then
     ' FLIMI = FLINUME
'    MODORAPIDO% = 0
   'End If
   '
'22
   
   Call APERBASDAT("CABAN")
   'JANDY SQL
   'SE CAMBIO LA RUTINA POR EL MAX(NUMECOMP)
   
   FLIMI = FETEXCOR1$(FLINUMERA)
  
  'ASELE$ = "SELECT NumeComp FROM CajaBanc "
   ASELE$ = "SELECT MAX(NumeComp) AS Maximo FROM CajaBanc "
   '
   If OPMOVIX$ = "ND" Or OPMOVIX$ = "NC" Then
     If NUMECORR% = (-2) Then    ' Unificada debitos y creditos
       ASELE$ = ASELE$ + "WHERE ((TipoMovim = 'FVEN' AND (OpciMovim = 'ND' OR OpciMovim = 'NC')) "
       ASELE$ = ASELE$ + "OR (TipoMovim = 'FCOM' AND (OpciMovim = 'DP' OR OpciMovim = 'CP'))) "
       GoTo 25
     End If
   End If
   '
   ASELE$ = ASELE$ + "WHERE TipoMovim = 'FVEN' "
   If NUMECORR% <> 1 Then
       ASELE$ = ASELE$ + "AND OpciMovim = '" & OPMOVIX$ & "' "
     Else
       '       ASELE$ = ASELE$ + "AND (OpciMovim = 'FC' OR OpciMovim = 'ND' OR OpciMovim = 'NC') "
       '       //////****NUEVA CONDICION PARA CALCULAR PROX. Nº FAC***////
       '       ESTA CONDICION LA AGREGUE  PARA QUE TOME EN CTA. LOS Nº
       '       DE COMPROBANTES ARRUINADOS: OR OpciMovim = 'ANUL'
       ASELE$ = ASELE$ + "AND (OpciMovim = 'FC' OR OpciMovim = 'ND' OR OpciMovim = 'NC' OR OpciMovim = 'ANUL') "
   End If
   '
25 ASELE$ = ASELE$ + "AND VariMovim = '" & VACREDI$ & "' "
   ASELE$ = ASELE$ + "AND (NumeTalo = 0 OR NumeTalo = 1) "
   ASELE$ = ASELE$ + "AND FechEmisi >= '" & FLIMI & "' "
   
   'ASELE$ = ASELE$ & "ORDER BY NumeComp DESC;"
   '
   'Set NUFAXI = CueCorri.OpenRecordset(ASELE$, dbOpenDynaset, dbReadOnly)
   '
   NUMAXI& = 0
   Dim MIRS As ADODB.Recordset
   Set MIRS = FLEXCONN.Execute(FILTSQL$(ASELE$))
   If Not IsNull(MIRS!maximo) Then
        NUMAXI& = MIRS!maximo
   End If
   MIRS.Close
   Set MIRS = Nothing
   
   'On Error Resume Next
   'NUFAXI.MoveFirst
   'If Err < 1 Then
   '    On Error GoTo 0
   '    NUMAXI& = NUFAXI!NumeComp
   '  Else
   '    On Error GoTo 0
   '    If MODORAPIDO% = 1 Then
   '       FLIMI = FLINUME
   '       MODORAPIDO% = 0
   '       GoTo 22
   '    End If
   'End If
   '
99 If Visible = True Then
     If NUMAXI& < 1 Then
       NUMAXI& = XVALO(InputBox$("Ingrese Número de la Próxima Nota de Crédito '" + VACREDI$ + "'", "Incicialización de Talonarios")) - 1
       If NUMAXI& < 0 Then GoTo 99
     End If
   End If
   NPP& = NUMAXI& + 1
   NFCZ$ = TRIM$(Str$(NPP&)): While Len(NFCZ$) < 8: NFCZ$ = "0" + NFCZ$: Wend
   NFCY$ = OPMOVIX$ + "-0001-" + NFCZ$ + "-" + VACREDI$

   PRONOCRE$ = TRIM$(NFCY$)
   Screen.MousePointer = 1
   '
End Function
'
Sub PRICREDESC()
    '
    CAHOJAS% = 0
    TIPOCOM$ = "(T.Libre)"
    TIFADOC% = 6 ' NOTA DE CREDITO
    NFCY$ = TRIM$(Mid$(Frame8.Caption, 9))
    VARDOCX$ = TRIM$(Mid$(NFCY$, 18, 1))
    '
    FORIPRE$ = TRIM$(Control$("", "FORFAC-" + VARDOCX$))
    If VARDOCX$ = "A" Then
        FORIALT$ = Control$("", "FORNCR-A") 'NC-A
      ElseIf UCase$(Label13.Caption) = "B" Then
        FORIALT$ = Control$("", "FORNCR-B") 'NC-B
    End If
    If TRIM$(FORIALT$) <> "" Then
      FORIPRE$ = FORIALT$
    End If
    ''\\\OT-06-0038-WAR-FIN///
    If FORIPRE$ = "" Then
       Call MENSERR(24, "Falta Definir Controlador de Impresión")
       Exit Sub
    End If
    '
    If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
       Call MENSERR(24, "Imposible Imprimir Documento.\No Existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
       Exit Sub
    End If
    '
    CODFOR$ = FORIPRE$
    For KKI% = 0 To CACOLTAB1%
      For KKJ% = 0 To CAITAB1%
        TETABU1$(KKI%, KKJ%) = ""
      Next KKJ%
    Next KKI%
    '
    Printer.TrackDefault = False
    PORTUFAC$ = TRIM$(Control$(IDENTER$, "PORTFAC"))
      If PORTUFAC$ = "" Then PORTUFAC$ = Control$("*", "PORTFAC")
    PORTUFAC1$ = TRIM$(PORTUFAC$)
    If PORTUFAC1$ = "" Then PORTUFAC1$ = Printer.DeviceName
    PPORT$ = REPLACAR$(PORTUFAC1$, "\", "/")
    NNCRE$ = APRORECIB.Label29
    Call COMUNI("Se Emitirá Nota de Crédito '" + NNCRE$ + "'\por Impresora '" + PPORT$ + "'.")
    '
    Screen.MousePointer = 11
    ITPORHOJA% = 32767
    '
    CANCELPRINT% = 0
    FECHDOC$ = Text9.TEXT
    NUMEDOC$ = NFCY$
    '
    TIPODOC$ = "Nota de Credito": CODOCUM$ = "CRE"
    CPAG% = 0
    'DESTIDOC% = XVALO(Text1.TEXT)
    DESTIDOCNUE = XVALO(Text1.TEXT)
    '
    CODTABU1$(1) = "COD-MAT"
    CODTABU1$(2) = "DES-MAT"
    CODTABU1$(3) = "CANTIDAD"
    'CODTABU1$(4) = "PRE-LIST"
    'CODTABU1$(5) = "PORDESCU"
    'CODTABU1$(6) = "PRE-NETO"
    'CODTABU1$(7) = "NET-ITEM"
    'CODTABU1$(8) = "FECH-REC"
    CODTABU1$(9) = "ORD-ITEM"
    CODTABU1$(10) = "UNIMED"
    CODTABU1$(11) = "REF-MAT1"
    CODTABU1$(12) = "REF-MAT2"
    CODTABU1$(13) = "REF-MAT3"
    CODTABU1$(14) = "N-PLANO"
    CODTABU1$(15) = "N-ORIGIN"
    CODTABU1$(16) = "N-OCOMPR"
    CODTABU1$(17) = "N-REMITO"
    CODTABU1$(18) = "N-BOLREC"
    CODTABU1$(19) = "FECH-REC"
    CODTABU1$(20) = "PRE-LIST"
    CODTABU1$(21) = "PORDESCU"
    CODTABU1$(22) = "PRE-NETO"
    CODTABU1$(23) = "NET-ITEM"
    CODTABU1$(24) = "ALICUIVA"
    CODTABU1$(25) = "IVA-ITEM"
    CODTABU1$(26) = "BRU-ITEM"
    
    '
    CACOLTAB1% = 26
    '
    CAITAB1% = 0
    CAITCRU% = 0
    '
    SUMATOTA# = XVALO(DESCUEN(0).TEXT)
    IDESCUEN# = 0
    '
    NroRemi$ = ""
    NUOCOMPRA$ = ""
    NROPEDIDO$ = ""
    '
    COEFIVAINCLU# = 1
    If VARDOCX$ = "B" Then
      COEFIVAINCLU# = 1 + ALIVA / 100
    End If
    '
    CODPARAM$(1) = "CIVA-CLI": DOCPARAM$(1) = Label13.Caption
    CODPARAM$(2) = "NRO-REMI": DOCPARAM$(2) = NroRemi$
    CODPARAM$(3) = "NRO-OCOM":  DOCPARAM$(3) = NUOCOMPRA$
    CODPARAM$(4) = "NRO-AENT":  DOCPARAM$(4) = NROPEDIDO$
    CODPARAM$(5) = "IMP-NETO": DOCPARAM$(5) = ""
    'CODPARAM$(6) = "VENDEDOR": DOCPARAM$(6) = Eco21(2).TEXT
    'CODPARAM$(7) = "COND-PAG": DOCPARAM$(7) = Eco21(1).TEXT
    'CODPARAM$(8) = "POR-IVA": DOCPARAM$(8) = PORIVA(0).Caption
    CODPARAM$(9) = "IMP-IVA": DOCPARAM$(9) = ""
    'CODPARAM$(10) = "POR-IVA1": DOCPARAM$(10) = PORIVA(1).Caption
    CODPARAM$(11) = "IMP-IVA1": DOCPARAM$(11) = ""
    CODPARAM$(12) = "IMP-DESC": DOCPARAM$(12) = ""
    CODPARAM$(13) = "IMPNEDES": DOCPARAM$(13) = ""
    'CODPARAM$(4) = "NRO-OCOM": DOCPARAM$(4) = Text10.TEXT
    CODPARAM$(14) = "IMP-TOTA": DOCPARAM$(14) = ""
    'CODPARAM$(15) = "POR-DESC": DOCPARAM$(15) = TRIM$(PORDESCU.TEXT)
    CODPARAM$(16) = "IMP-RNI": DOCPARAM$(16) = ""
    CODPARAM$(17) = "IMP-LET1": DOCPARAM$(17) = ""
    CODPARAM$(18) = "IMP-LET2": DOCPARAM$(18) = ""
    CODPARAM$(19) = "IMP-LET3": DOCPARAM$(19) = ""
    CODPARAM$(23) = "TOTA-DOL": DOCPARAM$(23) = ""
    '
    'CODPARAM$(20) = "DOMI-TRA": DOCPARAM$(20) = LENTREGA(0).TEXT
    'CODPARAM$(21) = "DESTINO": DOCPARAM$(21) = LENTREGA(1).TEXT
    'CODPARAM$(22) = "TRANSPOR": DOCPARAM$(22) = LENTREGA(2).TEXT
    CODPARAM$(24) = "MONEDEMI": DOCPARAM$(24) = MONEMIS
    CODPARAM$(25) = "PERC-BRU": DOCPARAM$(25) = ""
    CODPARAM$(26) = "SUB-TOTA": DOCPARAM$(26) = ""
    CODPARAM$(27) = "PERC-IVA": DOCPARAM$(27) = "" 'PERCIVA
    CODPARAM$(28) = "FECH-VEN": DOCPARAM$(28) = Format$(FVENCIM, "dd/mm/yy") ' FECHA DE VENCIMIENTO
      On Error Resume Next
      FEXO = FECHANUM(FECHDOC$)
      FEXU = CVINT(FVENCIM)
      CADIVEN = DIENTRE(FEXO, FEXU)
      On Error GoTo 0
    CODPARAM$(32) = "FECH-LAN": DOCPARAM$(32) = FORMATNUM$(CADIVEN, "F5.0")
    '\\\OT-5-579-FG-16/08/05
    'CODPARAM$(29) = "TIPO-VAL": DOCPARAM$(29) = TRIM$(Text18) ' VALOR FOB
    CODPARAM$(30) = "TICAMBIO": DOCPARAM$(30) = FORMATNUM$(TICAMONE(2), "F7.4") ' TIPO DE CAMBIO
    CODPARAM$(31) = "CTAC-CLI": DOCPARAM$(31) = CONTACTCLI$(DESTIDOCNUE) '  CONTACTO
    '\\\OT-5-579-FG-FIN
    '\\\OT-06-0125-WAR-03/04/06///
    'CODPARAM$(33) = "NUBULTOS": DOCPARAM$(33) = TRIM$(FINFACTURA.Text2(2)) ' BULTOS
    'CODPARAM$(34) = "PES-CONJ": DOCPARAM$(34) = TRIM$(FINFACTURA.Text2(11)) + " Kg" ' PESO NETO
    'CODPARAM$(35) = "PES-COMP": DOCPARAM$(35) = TRIM$(FINFACTURA.Text2(4)) + " Kg" ' PESO BRUTO
    'CODPARAM$(36) = "BANC-CHE": DOCPARAM$(36) = TRIM$(FINFACTURA.Text2(5)) 'BANCO EXTERIOR
    'CODPARAM$(37) = "DOCU-APL": DOCPARAM$(37) = TRIM$(FINFACTURA.Text2(8)) ' BANCO LOCAL
    'CODPARAM$(38) = "NUME-CHE": DOCPARAM$(38) = TRIM$(FINFACTURA.Text2(10)) ' CREDITO
    'CODPARAM$(39) = "DEN-CUEN": DOCPARAM$(39) = TRIM$(FINFACTURA.Text2(9)) ' GUIA DE IMPORTACION
    'CODPARAM$(40) = "REF-ASIE": DOCPARAM$(40) = TRIM$(FINFACTURA.Text2(6)) ' PUERTO DE EMBARQUE
    'CODPARAM$(41) = "DET-PASE": DOCPARAM$(41) = TRIM$(FINFACTURA.Text2(7)) ' DESTINO
    'CODPARAM$(42) = "MEDIDA-1": DOCPARAM$(42) = TRIM$(IEMBALA)             ' gasto de embalaje
    CODPARAM$(43) = "ZONA-VEN": DOCPARAM$(43) = TRIM$(Str$(GRUCOCLI%(DESTIDOCNUE)))
    '\\\OT-06-0125-WAR-FIN///
    'AGREGUEMH04-10
    'CODPARAM$(44) = "NRO-PASE": DOCPARAM$(44) = TRIM$(Label27.Caption) 'Sucursal
    '
    CAPARDOC% = 44 '32 '\\\OT-06-0125-WAR-03/04/06///
    '
    Call SETULTREG("!OBSERVOF", 0)
    '
    PRINTERPORT$ = PORTUFAC$
    CACOPDOC% = XVALO(FORFINFAC.Text2(0).TEXT)
    If CACOPDOC% < 1 Then CACOPDOC% = 1
    '
    NOCONFIRM% = 0
    HOJASUEL% = 1
    '
    CIPRE% = 0
    CAPEX$ = 1       ' cantidad
    '
    NPEDIDD& = 0
    SUBCO$ = ""
    '
    CAITAB1% = CAITAB1% + 1
    CODDX$ = ""
    TETABU1$(1, CAITAB1%) = CODDX$
    TETABU1$(2, CAITAB1%) = "Desc." + TRIM$(Text7) + " % S/Rbo." + TRIM$(Text11)
    TETABU1$(3, CAITAB1%) = CAPEX$
    TETABU1$(9, CAITAB1%) = Str$(U&)
    TETABU1$(10, CAITAB1%) = ""    ' unidad de medida
    TETABU1$(11, CAITAB1%) = ""    ' referencia 1
    TETABU1$(12, CAITAB1%) = ""    ' referencia 1
    TETABU1$(13, CAITAB1%) = ""    ' referencia 1
    TETABU1$(14, CAITAB1%) = ""    ' numero de plano
    TETABU1$(15, CAITAB1%) = ""    ' numero original
    TETABU1$(16, CAITAB1%) = ""    ' numero de o.compra
    TETABU1$(17, CAITAB1%) = ""    ' numero de remito
    TETABU1$(18, CAITAB1%) = ""    ' numero de doc-recepcion
    TETABU1$(19, CAITAB1%) = ""    ' fecha de recepcion
    '
    FACIVITEM = 1: If VARDOCX$ = "B" Then FACIVITEM = 1 + ALIVA / 100
    TETABU1$(20, CAITAB1%) = ""    ' precio de lista
    TETABU1$(21, CAITAB1%) = ""    ' PORCENTAJE DE DESCUENTO
    '
    TETABU1$(22, CAITAB1%) = Str$(FACIVITEM * XVALO(DESCUEN(0)) / (1 + ALIVA / 100)) ' precio unitario
    TETABU1$(23, CAITAB1%) = Str$(XVALO(DESCUEN(0)) / (1 + ALIVA / 100))       ' total neto
    TETABU1$(24, CAITAB1%) = Str$(ALIVA)                         ' alicuota iva
    TETABU1$(25, CAITAB1%) = Str$(ALIVA / 100 * XVALO(DESCUEN(0)) / (1 + ALIVA / 100)) ' importe iva del item
    TETABU1$(26, CAITAB1%) = DESCUEN(0) ' total bruto
    '
    CODPARAM$(5) = "IMP-NETO": DOCPARAM$(5) = Str$(XVALO(DESCUEN(0)) / (1 + ALIVA / 100))
    CODPARAM$(9) = "IMP-IVA": DOCPARAM$(9) = Str$(ALIVA / 100 * XVALO(DESCUEN(0)) / (1 + ALIVA / 100))
    CODPARAM$(11) = "IMP-IVA1": DOCPARAM$(11) = ""
    CODPARAM$(12) = "IMP-DESC": DOCPARAM$(12) = ""
    CODPARAM$(13) = "IMPNEDES": DOCPARAM$(13) = Str$(XVALO(DESCUEN(0)) / (1 + ALIVA / 100))
    CODPARAM$(14) = "IMP-TOTA": DOCPARAM$(14) = DESCUEN(0)
    CODPARAM$(16) = "IMP-RNI": DOCPARAM$(16) = ""
    If TICAMBIO > 0 Then
       VADOLA# = XVALO(DESCUEN(0)) / TICAMBIO
       DOCPARAM$(23) = TRIM$(CSTRING$(MKD$(VADOLA#), 4, 12, 2, 2))
    End If
    CODPARAM$(25) = "PERC-BRU": DOCPARAM$(25) = ""
    DOCPARAM$(26) = ""
    DOCPARAM$(27) = ""                          'PARA IMPRIMIR PERC-IVA
    '
    If VARDOCX$ = "B" Then
      DOCPARAM$(5) = FORMATNUM$(COEFIVAINCLU# * (XVALO(DESCUEN(0)) / (1 + ALIVA / 100)), "F13.2")
      'DOCPARAM$(12) = FORMATNUM$(COEFIVAINCLU# * XVALO(Total(1).TEXT), "F13.2")
      DOCPARAM$(13) = FORMATNUM$(COEFIVAINCLU# * XVALO(DESCUEN(0)) / (1 + ALIVA / 100), "F13.2")
      'DOCPARAM$(42) = FORMATNUM$(COEFIVAINCLU# * XVALO(IEMBALA), "F13.2")          ' gasto de embalaje
    End If
    '
    CODFOR$ = FORIPRE$
    LOIMPLET% = XVALO(ATRIFORM$(CODFOR$, "LOIMPLET"))
    If LOIMPLET% < 16 Then LOIMPLET% = 32
    TTTT# = XVALO(DESCUEN(0))
    TOTFAC$ = NUMTEX$(TTTT#)
    If MONECO$ <> "" Then
      PPXX% = InStr(TOTFAC$, "PESOS")
      If PPXX% > 0 Then
        TOTFAC$ = Left$(TOTFAC$, PPXX% - 1) + UCase$(MONECO$) + Mid$(TOTFAC$, PPXX% + 5)
      End If
    End If
    Call FRATEX(TOTFAC$, NRX1$, NRX2$, NRX3$, LOIMPLET%)
    CODPARAM$(17) = "IMP-LET1": DOCPARAM$(17) = NRX1$
    CODPARAM$(18) = "IMP-LET2": DOCPARAM$(18) = NRX2$
    CODPARAM$(19) = "IMP-LET3": DOCPARAM$(19) = NRX3$
    '
    If CACOPDOC% > 0 Then
      Call PRINTDOC("@" + FORIPRE$)
    End If
    '
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    '
    For KKI% = 0 To CACOLTAB1%
      For KKJ% = 0 To CAITAB1%
        TETABU1$(KKI%, KKJ%) = ""
      Next KKJ%
    Next KKI%
    '
End Sub

Sub LIMPPROM()
  Label17 = ""
  Label19 = ""
  Label23 = ""
  Label35 = ""
  PROVENCI# = 0
End Sub

Private Sub Timer1_Timer()
   '
   SUVALO# = 0
   For U& = 1 To ULTREG&("!CARMECO")
      x$ = REGLEIDO$("!CARMECO", U&)
      SUVALO# = SUVALO# + ROUND#(CVD(Mid$(x$, 31, 8)))
   Next U&
   Label30 = FORMATNUM(SUVALO#, "F12.2")
   Label14 = FORMATNUM$(SUVALO# + XVALO(Text8), "F12.2")
   FALTAPACAN# = XVALO(Label8) * ((100 - XVALO(Text7)) / 100) - XVALO(Label14)
   Label29.ForeColor = RGB(0, 0, 0)
   If FALTAPACAN# < 0 Then
      Label29.ForeColor = RGB(255, 0, 0)
   End If
   Label29 = TRIM$(FORMATNUM$(FALTAPACAN#, "F10.2"))
   Label29.Refresh
   '
End Sub


