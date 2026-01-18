VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FAPECOB07 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturas Pendientes de Cobranza "
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   9240
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
   ScaleHeight     =   4755
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2400
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   55
      Top             =   120
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2D3C0&
      Caption         =   "SALIDA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   4250
      TabIndex        =   22
      Top             =   2310
      Width           =   1905
      Begin VB.CommandButton Command17 
         Caption         =   "!"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1575
         TabIndex        =   52
         ToolTipText     =   "Mail Aviso Vencimiento"
         Top             =   360
         Width           =   195
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
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
         Index           =   4
         Left            =   315
         TabIndex        =   24
         Top             =   270
         Width           =   1170
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
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
         Index           =   8
         Left            =   315
         TabIndex        =   23
         Top             =   590
         Value           =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E2D3C0&
      Caption         =   "FORMATO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   6370
      TabIndex        =   44
      Top             =   2310
      Width           =   1900
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Mora"
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
         Height          =   315
         Index           =   11
         Left            =   1050
         TabIndex        =   53
         Top             =   270
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Pl.Financiero"
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
         Index           =   10
         Left            =   90
         TabIndex        =   46
         Top             =   580
         Width           =   1485
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Tabla"
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
         Height          =   315
         Index           =   9
         Left            =   90
         TabIndex        =   45
         Top             =   270
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Cust."
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
      Left            =   1890
      Picture         =   "FAPECOB07.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Acceso a Maestro de Clientes"
      Top             =   5040
      Width           =   650
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      ForeColor       =   &H80000008&
      Height          =   5685
      Left            =   8400
      ScaleHeight     =   5655
      ScaleWidth      =   1560
      TabIndex        =   39
      Top             =   0
      Width           =   1590
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   90
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   64
         Top             =   4400
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   65
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command18 
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
         Left            =   90
         Picture         =   "FAPECOB07.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Configuración de Procesos y Formularios"
         Top             =   2760
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
         Left            =   90
         Picture         =   "FAPECOB07.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Emite Reporte Solicitado"
         Top             =   3660
         Width           =   650
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Mail"
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
         Left            =   90
         TabIndex        =   50
         Top             =   3420
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command13 
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
         Left            =   90
         Picture         =   "FAPECOB07.frx":0B8E
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Pendientes en Moneda Extranjera"
         Top             =   2110
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command5 
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
         Height          =   650
         Left            =   90
         Picture         =   "FAPECOB07.frx":0E98
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Opciones por Default"
         Top             =   810
         Width           =   650
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
         Height          =   650
         Left            =   90
         Picture         =   "FAPECOB07.frx":0FE2
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   170
         Width           =   650
      End
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Deuda"
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
      Left            =   5355
      Picture         =   "FAPECOB07.frx":112C
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Composición de Deuda a Fecha"
      Top             =   6195
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   540
      Left            =   2625
      Sorted          =   -1  'True
      TabIndex        =   34
      Top             =   6195
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame5 
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
      Height          =   2160
      Left            =   4250
      TabIndex        =   25
      Top             =   80
      Width           =   4005
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
         Left            =   2625
         MaxLength       =   6
         TabIndex        =   31
         Top             =   420
         Width           =   810
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
         Left            =   3450
         TabIndex        =   30
         Top             =   420
         Width           =   175
      End
      Begin VB.TextBox Text9 
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
         Left            =   315
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   " "
         Top             =   800
         Width           =   3330
      End
      Begin VB.TextBox Text8 
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
         Left            =   2625
         MaxLength       =   6
         TabIndex        =   28
         Text            =   " "
         Top             =   1290
         Width           =   840
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
         Height          =   300
         Left            =   3465
         TabIndex        =   27
         Top             =   1305
         Width           =   175
      End
      Begin VB.TextBox Text7 
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
         Left            =   315
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1665
         Width           =   3330
      End
      Begin VB.Label Label10 
         Caption         =   "Si existe FAPECOBM es visible la opcion Mora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   0
         TabIndex        =   54
         Top             =   1200
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label6 
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
         Left            =   210
         TabIndex        =   33
         Top             =   510
         Width           =   2325
      End
      Begin VB.Label Label5 
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
         Left            =   210
         TabIndex        =   32
         Top             =   1380
         Width           =   2325
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Formato de Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   525
      TabIndex        =   18
      Top             =   5985
      Width           =   4005
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Listado Tabulado"
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
         Index           =   7
         Left            =   420
         TabIndex        =   21
         Top             =   420
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Plan Financiero de Cobranza"
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
         Index           =   6
         Left            =   420
         TabIndex        =   20
         Top             =   735
         Width           =   3270
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Gráfico"
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
         Index           =   5
         Left            =   420
         TabIndex        =   19
         Top             =   1155
         Width           =   1380
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   3990
         Y1              =   1050
         Y2              =   1050
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      Caption         =   "FILTROS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4560
      Left            =   120
      TabIndex        =   5
      Top             =   80
      Width           =   4005
      Begin VB.Frame Frame8 
         BackColor       =   &H00E2D3C0&
         Caption         =   "FECHA DE EMISION"
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
         Left            =   120
         TabIndex        =   57
         Top             =   3600
         Width           =   3750
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
            Height          =   285
            Left            =   1680
            TabIndex        =   59
            Top             =   360
            Width           =   175
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
            Height          =   285
            Left            =   3465
            TabIndex        =   58
            Top             =   360
            Width           =   175
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            Left            =   705
            TabIndex        =   63
            Top             =   360
            Width           =   960
         End
         Begin VB.Label Label123 
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
            Left            =   15
            TabIndex        =   62
            Top             =   425
            Width           =   645
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            Left            =   2460
            TabIndex        =   61
            Top             =   350
            Width           =   1000
         End
         Begin VB.Label Label7 
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
            Left            =   1770
            TabIndex        =   60
            Top             =   425
            Width           =   645
         End
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Barrida Automática"
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
         Left            =   315
         TabIndex        =   48
         Top             =   1110
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Barrida Automática"
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
         Left            =   315
         TabIndex        =   47
         Top             =   240
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E2D3C0&
         Caption         =   "RANGO DE FECHAS"
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
         Left            =   135
         TabIndex        =   14
         Top             =   2640
         Width           =   3750
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
            Left            =   1680
            TabIndex        =   38
            Top             =   340
            Width           =   175
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
            Left            =   700
            TabIndex        =   36
            Top             =   340
            Width           =   960
         End
         Begin VB.CommandButton Command6 
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
            Left            =   3465
            TabIndex        =   16
            Top             =   340
            Width           =   175
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
            Left            =   2460
            TabIndex        =   15
            Top             =   340
            Width           =   1000
         End
         Begin VB.Label Label4 
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
            Left            =   1770
            TabIndex        =   37
            Top             =   440
            Width           =   645
         End
         Begin VB.Label Label3 
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
            Left            =   0
            TabIndex        =   17
            Top             =   440
            Width           =   645
         End
      End
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
         Left            =   315
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "Todos - Sin Clasificación"
         Top             =   1665
         Width           =   3330
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
         Height          =   300
         Left            =   3480
         TabIndex        =   11
         Top             =   1290
         Width           =   175
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
         Left            =   3060
         TabIndex        =   10
         Text            =   " "
         Top             =   1290
         Width           =   435
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
         Left            =   315
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " Todos - Sin Clasificación"
         Top             =   800
         Width           =   3330
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
         Height          =   300
         Left            =   3480
         TabIndex        =   7
         Top             =   420
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
         Height          =   300
         Left            =   3045
         TabIndex        =   6
         Text            =   " "
         Top             =   420
         Width           =   435
      End
      Begin VB.Label Label1 
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
         Left            =   645
         TabIndex        =   13
         Top             =   1380
         Width           =   2325
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo de Clientes"
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
         Left            =   630
         TabIndex        =   9
         Top             =   510
         Width           =   2325
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E2D3C0&
      Caption         =   "ORDENAMIENTO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1300
      Left            =   4250
      TabIndex        =   0
      Top             =   3330
      Width           =   4005
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Por Fecha de Vencimiento"
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
         Index           =   3
         Left            =   315
         TabIndex        =   4
         Top             =   945
         Width           =   2745
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Alfabético"
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
         Index           =   2
         Left            =   2205
         TabIndex        =   3
         Top             =   630
         Width           =   1380
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Numérico"
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
         Index           =   1
         Left            =   735
         TabIndex        =   2
         Top             =   630
         Value           =   -1  'True
         Width           =   1380
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Por Cliente"
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
         Index           =   0
         Left            =   315
         TabIndex        =   1
         Top             =   315
         Width           =   1380
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "FAPECOB07.frx":1276
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "FAPECOB07.frx":14AA
      TabIndex        =   56
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Menu OtrasConsultas 
      Caption         =   "Otras consultas"
      Begin VB.Menu ComposicionDeuda 
         Caption         =   "Composición de Deuda"
      End
      Begin VB.Menu ComposicionDeudaCliente 
         Caption         =   "Composición de Deuda por Cliente"
      End
      Begin VB.Menu linea1 
         Caption         =   "-"
      End
      Begin VB.Menu PendienteCobranzaFecha 
         Caption         =   "Pendientes de Cobrar por Rango de Fecha"
      End
      Begin VB.Menu mnuFacpropend 
         Caption         =   "Facturas Proforma Pend. x Rango de Emisión y Clientes Activos"
      End
   End
End
Attribute VB_Name = "FAPECOB07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TICUE%(32767)

Sub AVIVENCI()
   If DERACCE%("AVISOVEN", "Aviso de vencimiento de factura") = 2 Then ' Definir derecho de acceso
   FEX = HOY(HO$) ' tomo la fecha del dia
   FEXU% = DIASPOST(FEX, 5) ' tomo la fecha 5 dias posterior
   FEVE = FETEXCOR1$(FEXU%) ' cambio el formato para la consulta sql
   Dim CABATEMP As ADODB.Recordset ' dim a la base de datos
   Set CABATEMP = New ADODB.Recordset ' cargar bd
   SQLX$ = "SELECT SUM(salddebe-saldacre), nuclien FROM sadeupen"
   SQLX$ = SQLX$ + " WHERE fevencim <= '" & FEVE & " '"
   SQLX$ = SQLX$ + " GROUP BY nuclien Having Sum(salddebe-saldacre) > 0.005 ORDER BY Nuclien"
   ' consulta a la bd
   Set CABATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
     With CABATEMP
      Do While (Not .EOF) ' recorrer la bd
         NC% = CABATEMP!NUCLIEN ' asignar campo a la variable
         RAZO$ = AJUSTI(Left$(rasocli$(NC%), 30), 32) ' Ajustar a la izq, 30 posiones de la izq
         mail = AJUSTI(EMAILCLI(NC%), 35)
         SELCLI09.LISCLISE.AddItem FORMATNUM$(NC%, "F6") + " " + RAZO$ + mail ' mostrar en el list
         ' Muestra el checkbox On, solo los que fueron tildados previamente en Clientes
         If ENVMAIDEUDA%(NC) = 1 Then
         SELCLI09.LISCLISE.Selected(i) = True ' checkbox On
         End If
         i = i + 1
         .MoveNext
       Loop
      End With
      
      If SELCLI09.LISCLISE.ListCount > 0 Then SELCLI09.LISCLISE.ListIndex = 0
      SELCLI09.Show 1
   End If
End Sub

Sub CARGAINTMORA()
       If DERACCE%("INTMORA", "Mora en Cuentas a Cobrar") > 1 Then
           ' LEE EL CONTROL
           INTMORA$ = CONTROL("VENTAS", "INTMORA")
           
           INTERES# = XVALO(INTMORA$)
           ' CAMBIA LOS DATOS DE LA VARIABLE A BINARIO
           Call REPLA(VDEF$, MKS$(INTERES#), 1, 4)
           ' SE GUARDA EL VALOR DE OBJERTO EN UNA VARIABLE
           OBX$ = OBJPLA("INTEMORA", VDEF$)
           '
           If OBX$ = "" Then Exit Sub
           
           INTERES# = CVS(Mid(OBX$, 1, 4))
           TXT$ = TRIM(FORMATNUM(INTERES#, "F10.2"))
           ' GRABO EL CONTROL
           Call GRACONTROL("VENTAS", "INTMORA", TXT$)
       End If
End Sub

Sub RECONSTRUYE_DEUDA_CLIENTES_CON_DIFERENCIA()
'
    Dim DEUDA#(999999), SALDO#(999999)

    For i& = 1 To 999999
      SALDO#(i&) = 0
      DEUDA#(i&) = 0
    Next i&
    '
    Screen.MousePointer = 11
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT SUM(IDEBECOMP) AS DEBEX  ,SUM(IHABECOMP) AS HABEX  , NUCLIEN  FROM CAJABANC WITH(NOLOCK)"
    SQLX$ = SQLX$ + "WHERE (TipoMovim = 'FVEN' "
    SQLX$ = SQLX$ + "or TipoMovim = 'RCOB' "
    SQLX$ = SQLX$ + "or TipoMovim = 'AJUD') "
    SQLX$ = SQLX$ + "GROUP BY NUCLIEN "
    '
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = READSET(SQLX$)
    With CABATEMP
       Do While .EOF = False
        NCLI = Abs(XVALO(!NUCLIEN))
        SALDO#(NCLI) = XVALO(!DEBEX) - XVALO(!HABEX)
        .MoveNext
      Loop
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
    '
    SQLX$ = "SELECT SUM(salddebe) AS SADEBE ,SUM(saldacre) AS SAACRRE ,NUCLIEN FROM SADEUPEN WITH(NOLOCK)"
    SQLX$ = SQLX$ + "WHERE Status < 3 "
    SQLX$ = SQLX$ + "GROUP BY NUCLIEN "
    '
    Dim SACTEMP As ADODB.Recordset
    Set SACTEMP = READSET(SQLX$)
    With SACTEMP
       Do While .EOF = False
        NCLI = XVALO(!NUCLIEN)
        DEUDA#(NCLI) = XVALO(!SADEBE) - XVALO(!SAACRRE)
        .MoveNext
      Loop
    End With
    SACTEMP.Close
    Set SACTEMP = Nothing
    '
    FIN& = 0
    Dim rstdeudor As ADODB.Recordset
    Set rstdeudor = FLEXCONN.Execute("select max(nume_cli) as max from deudor12")
    If Not (rstdeudor.EOF And rstdeudor.BOF) Then FIN& = rstdeudor!Max
    rstdeudor.Close: Set rstdeudor = Nothing
    For i& = 1 To FIN&
       If Abs(SALDO#(i&) - DEUDA#(i&)) > 0.005 Then
            Call RECUEC07.REAPLICRE(i&)
            '
            CONDI2$ = " Status < 3  AND NUCLIEN = " & i&
            SADE# = XVALO(VALOBADA("SADEUPEN", "Sum(salddebe)", CONDI2$, "NOMESS"))
            SACRE# = XVALO(VALOBADA("SADEUPEN", "Sum(saldacre)", CONDI2$, "NOMESS"))
            DEUDAX# = SADE# - SACRE#
            '
            If Abs(SALDO#(i&) - DEUDAX#) > 0.005 Then
               Call COMUNI("No se ha Podido Recuperar la Diferencia (Saldo Y Deuda) Del Cliente N° : " & i& & "\" & rasocli$(i&) & "\ Consulte a su Soporte de Sistemas")
            End If
           
        End If
    Next i&
    '
    
End Sub

Private Sub Check1_Click()
   If Check1.Value = 1 Then
      Check2.Value = 0
      Text1 = ""
   End If
End Sub

Private Sub Check2_Click()
   If Check2.Value = 1 Then
      Check1.Value = 0
      Text3 = ""
   End If
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   Command2.Enabled = False
   Call COMPODEU
   Command2.Enabled = True
   Command10.Enabled = True
End Sub

Private Sub Command11_Click()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text5.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   Call CIERRARCH("*.*")
   Call CONECRUN("ARCHIG07/AMACLI", "Archivos Maestros")
   Command12.Enabled = True
End Sub

Private Sub Command13_Click()
    '
    Dim DOCULAR$(), SALPESO#(), SALDOLA#()
    '
    Command13.Enabled = False
    MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESVENTA", "MODOMONE"))), 8)
    If MODOMONE$ <> "BIMONETA" Then
       If MODOMONE$ <> "MULTIMON" Then
          Call MENSERR(24, "Solo Disponible en Sistemas Bi- O Multi-Monetarios")
          GoTo 99
       End If
    End If
    '
    Screen.MousePointer = 11
    J& = 0
    '
    NUDOCU& = 0: RAMAX& = 0
    FIN& = ULTREG&("CARBIMON")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("CARBIMON", U&)
      NUDOCX$ = TRIM$(Mid$(XX$, 13, 18))
      DEBEPES# = ROUND#(CVD(Mid$(XX$, 65, 8)))
      HABEPES# = ROUND#(CVD(Mid$(XX$, 73, 8)))
      DEBEDOL# = ROUND#(CVD(Mid$(XX$, 97, 8)))
      HABEDOL# = ROUND#(CVD(Mid$(XX$, 105, 8)))
      '
      If AJUSTI$(NUDOCX$, 18) = AJUSTI$(Mid$(XX$, 45, 18), 18) Then ' ES MOVIMIENTO CABEZA
        If Abs(DEBEPES# - DEBEDOL#) < 0.005 Then
          If Abs(HABEPES# - HABEDOL#) < 0.005 Then
            TTXX$ = "Se han Encontrado Registros Incorrectos\en la Tabla de Control Multimoneda.\  \¿ Desea Recuperar la Tabla Dañada ?"
            If COMALTER%(TTXX$ + "\\Si, Recuperar\No, Continuar") = 1 Then
              Call RECUBIMO
              Call COMUNI("Tabla Recuperada.\Vuelva Ahora a Pedir el Listado")
              GoTo 99
            End If
          End If
        End If
      End If
      '
      For JJ& = 1 To NUDOCU&
         If DOCULAR$(JJ&) = NUDOCX$ Then
            SALPESO#(JJ&) = SALPESO#(JJ&) + DEBEPES# - HABEPES#
            SALDOLA#(JJ&) = SALDOLA#(JJ&) + DEBEDOL# - HABEDOL#
            GoTo 19
         End If
      Next JJ&
      '
      NUDOCU& = NUDOCU& + 1
      If NUDOCU& > RAMAX& Then
         RAMAX& = RAMAX& + 128
         ReDim Preserve DOCULAR$(RAMAX&), SALPESO#(RAMAX&), SALDOLA#(RAMAX&)
      End If
      DOCULAR$(NUDOCU&) = NUDOCX$
      SALPESO#(NUDOCU&) = SALPESO#(NUDOCU&) + DEBEPES# - HABEPES#
      SALDOLA#(NUDOCU&) = SALDOLA#(NUDOCU&) + DEBEDOL# - HABEDOL#
      '
19  Next U&
    '
    JJ& = 0: TEBUX$ = ""
    For U& = 1 To NUDOCU&
       Call TRACE(20, U&, NUDOCU&)
       If Abs(SALPESO#(U&)) >= 0.005 Or Abs(SALDOLA#(U&)) >= 0.005 Then
          TEBUX$ = TEBUX$ + "**" + AJUSTI$(DOCULAR$(U&), 18)
       End If
    Next U&
    '
    FIN& = ULTREG&("CARBIMON")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       XX$ = REGLEIDO$("CARBIMON", U&)
       NUDOCX$ = "**" + AJUSTI$(Mid$(XX$, 13, 18), 18)
       If InStr(TEBUX$, NUDOCX$) > 0 Then
          '
          ' ENJUAGUE ASEMBLI-GRAMY POR ERROR DE MONEDA
          MONEDAPLI% = CVI(Mid$(XX$, 63, 2))
          If MONEDAPLI% > ULTREG&("TAMONED") Then
            MONEDAPLI% = MONEDAPLI% Mod 6
            Call REPLA(XX$, MKI$(MONEDAPLI%), 63, 2)
            Call GRAREG("CARBIMON", XX$, U&)
            If TRIM$(Mid$(XX$, 45, 18)) = TRIM$(Mid$(XX$, 13, 18)) Then
              Call REPLA(XX$, MKI$(MONEDAPLI%), 31, 2)
              Call GRAREG("CARBIMON", XX$, U&)
            End If
          End If
          ' FIN ENJUAGUE ASEMBLI-GRAMY
          '
          J& = J& + 1
          Call GRAREG("CARBIMOX", XX$, J&)
       End If
    Next U&
    '
    Call SETULTREG("CARBIMOX", J&)
    Call CIERRARCH("CARBIMOX")
    '
    Call FILESORT("CARBIMOX", "", 1, 6, "ASC")
    Call CONECRUN("*LISBIMON", "")
    '
99  Screen.MousePointer = 1
    Command13.Enabled = True

End Sub

Sub RECUBIMO()
    '
    FIN& = ULTREG&("CARBIMON")
    Call ABRECABAN
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("CARBIMON", U&)
      NUDOCX$ = AJUSTI$(Mid$(XX$, 13, 18), 18)
      If TRIM$(NUDOCX$) <> "" Then
         DEBEPES# = ROUND#(CVD(Mid$(XX$, 65, 8)))
         HABEPES# = ROUND#(CVD(Mid$(XX$, 73, 8)))
         DEBEDOL# = ROUND#(CVD(Mid$(XX$, 97, 8)))
         HABEDOL# = ROUND#(CVD(Mid$(XX$, 105, 8)))
         NCX% = CVI(Mid$(XX$, 3, 2))
         '
         If AJUSTI$(Mid$(XX$, 45, 18), 18) = NUDOCX$ Then ' ES MOVIMIENTO CABEZA
           If Abs(DEBEPES# - DEBEDOL#) < 0.005 Then
             If Abs(HABEPES# - HABEDOL#) < 0.005 Then
               YY$ = String$(256, 0)
               Call GRAREG("CARBIMON", YY$, U&)
               '
               For JJ& = U& + 1 To FIN&
                 ZZ$ = REGLEIDO$("CARBIMON", JJ&)
                 If AJUSTI$(Mid$(ZZ$, 13, 18), 18) = NUDOCX$ Then
                   YY$ = String$(256, 0)
                   Call GRAREG("CARBIMON", YY$, JJ&)
                 End If
               Next JJ&
               '
               DOCULA$ = TRIM$(NUDOCX$)
               SQLX$ = "SELECT * FROM CAJABANC "
               SQLX$ = SQLX$ + " WHERE NuClien = " & NCX% & " "
               SQLX$ = SQLX$ + " AND CODICOM_LAR = '" & DOCULA$ & "'"
               Dim CAJJJABANC As ADODB.Recordset
               Set CAJJJABANC = New ADODB.Recordset
25             On Error Resume Next
               CAJJJABANC.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
               If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                   On Error GoTo 0
                   Call CapturaErrorOpen
                   GoTo 25
               End If
               On Error GoTo 0
               '
               With CAJJJABANC
                  On Error Resume Next
                  '.FindFirst "NuClien = " & NCX% & " AND CODICOM_LAR = '" & DOCULA$ & "'"
'201               'If .NoMatch = False Then
                    '.Edit
                  If Not (.EOF And .BOF) Then
                      On Error GoTo 0
                      Do While (Not .EOF)
                        !Mone_Emis = 1
                        !Tipo_Cam = 1
                        !Impo_Dola = 0
                        .Update
                        '.FindNext "NuClien = " & NCX% & " AND CODICOM_LAR = '" & DOCULA$ & "'"
                        .MoveNext
                      Loop
                      'GoTo 201
                  End If
                  On Error GoTo 0
               End With
               '
               SQLX$ = "SELECT * FROM SaDeuPen "
               SQLX$ = SQLX$ + " WHERE NuClien = " & NCX% & " "
               SQLX$ = SQLX$ + " AND CODICOM_LAR = '" & DOCULA$ & "'"
               Dim SADDDEUPEN As ADODB.Recordset
               Set SADDDEUPEN = New ADODB.Recordset
26             On Error Resume Next
               SADDDEUPEN.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
               If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                   On Error GoTo 0
                   Call CapturaErrorOpen
                   GoTo 26
               End If
               On Error GoTo 0
               '
               With SADDDEUPEN
                  On Error Resume Next
                  '.FindFirst "NuClien = " & NCX% & " AND CODICOM_LAR = '" & DOCULA$ & "'"
202               'If .NoMatch = False Then
                    '.Edit
                  If Not (.EOF And .BOF) Then
                    On Error GoTo 0
                    Do While (Not .EOF)
                      !SaldDebe_DOLA = 0
                      !SaldAcre_DOLA = 0
                      !Mone_Emis = 1
                      !Ibru_Dola = 0
                      .Update
                      '.FindNext "NuClien = " & NCX% & " AND CODICOM_LAR = '" & DOCULA$ & "'"
                      .MoveNext
                    Loop
                    'GoTo 202
                 End If
               End With
               '
             End If
           End If
         End If
      End If
    Next U&
    '
    JJ& = 0
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("CARBIMON", U&)
      If XX$ <> String$(256, 0) Then
         JJ& = JJ& + 1
         Call GRAREG("CARBIMON", XX$, JJ&)
      End If
    Next U&
    CAJJJABANC.Close
    Set CAJJJABANC = Nothing
    SADDDEUPEN.Close
    Set SADDDEUPEN = Nothing
    Call SETULTREG("CARBIMON", JJ&)
    Call CIERRARCH("CARBIMON")
    '
End Sub

Private Sub Command14_Click()
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   Label8 = VALACT1$("SELFECHA")
End Sub

Private Sub Command15_Click()
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   Label9 = VALACT1$("SELFECHA")
End Sub

Private Sub Command16_Click()
    ' envio de deuda por mail
    Call COPYSTRU("FAPECOBA", "FAPECOX")
    For U& = 1 To ULTREG&("DEUDOR1")
       X1$ = REGLEIDO$("DEUDOR1", U&)
       NC% = CVI(Left$(X1$, 2))
       RFF$ = RECFILT$("FAPECOBA", 1, MKS$(NC%))
       If Len(RFF$) > 0 Then
         UJ& = 0
         SUMADEU# = 0
         For UL& = 1 To Len(RFF$) Step 4
           UK& = CVS(Mid$(RFF$, UL&, 4))
           X$ = REGLEIDO$("FAPECOBA", UK&)
           UJ& = UJ& + 1
           YY$ = MKI$(NC%) & Mid$(X$, 3)
           Call GRAREG("FAPECOX", YY$, UJ&)
           SUMADEU# = SUMADEU# + CVD(Mid$(X$, 73, 8))
         Next UL&
         Call SETULTREG("FAPECOX", UJ&)
         '
         If SUMADEU# > 0.005 Then
            JJ& = 0
            n1% = FILEOPEN%("C:\FLEXOFT\CODEUCLI.TXT", 2, 128)
            FIN& = ULTREG&("FAPECOX")
            For U3& = 1 To FIN&
               '
               Call TRACE(20, U3&, FIN&)
               X$ = REGLEIDO$("FAPECOX", U3&)
               '
               TTXX$ = Space$(80)
               NC% = CVI(Left$(X$, 2))
               NFS$ = Mid$(X$, 49, 12)
               FF% = CVI(Mid$(X$, 61, 2))
               FVI% = CVI(Mid$(X$, 71, 2))
               SALDEU# = CVD(Mid$(X$, 73, 8))
               IMORIG# = CVD(Mid$(X$, 63, 8))
               '
               Call REPLA(TTXX$, NFS$, 1, 12)
               FEX = FF%
               Call REPLA(TTXX$, FECHATEX$(FEX), 15, 8)
               Call REPLA(TTXX$, FORMATNUM$(IMORIG#, "F12.2"), 23, 12)
               FEX = FVI%
               Call REPLA(TTXX$, FECHATEX$(FEX), 37, 8)
               Call REPLA(TTXX$, FORMATNUM$(SALDEU#, "F12.2"), 45, 12)
               '
               If JJ& < 1 Then
                 HOII% = HOY(HOS$)
                 Print #n1%, "." + Space$(20) + "Buenos Aires, " + FETEXTO(HOII%)
                 Print #n1%, "Sres:"
                 Print #n1%, rasocli$(NC%)
                 Print #n1%, DOMICLI$(NC%)
                 Print #n1%, LOCACLI$(NC%)
                 Print #n1%, "   "
                 Print #n1%, "At: " + CONTACTCLI$(NC%)
                 Print #n1%, "   "
                 Print #n1%, "De nuestra consideración:"
                 Print #n1%, "                          Nos dirigimos a uds con el"
                 Print #n1%, "fin de recordarles el estado de su apreciada cuenta:"
                 Print #n1%, "   "
                 Print #n1%, "Factura     Fecha      Importe   Vencim.       Saldo"
                 Print #n1%, "-------------------------------------------------------"
               End If
               '
               JJ& = JJ& + 1
               Print #n1%, TTXX$
            Next U3&
            Print #n1%, "---------------------------------------------------------"
            Print #n1%, "Su Saldo Total a la Fecha: $ " + FORMATNUM$(SUMADEU#, "F12.2")
            Print #n1%, ""
            Print #n1%, "Agradecemos cancelar dicho saldo en cualquiera de las "
            Print #n1%, "formas siguientes:"
            Print #n1%, ""
            Print #n1%, "- Transferencia a nuestra CC 120-1160/1 del Banco Francés"
            Print #n1%, ". BBVA, Sucursal 120, Titular Ernesto Pablo Bauer, Número"
            Print #n1%, ". de CBU: 01701209  200  000  001  160  17"
            Print #n1%, ""
            Print #n1%, "- Cheque a la orden de Ernesto Pablo Bauer, comunicándonos"
            Print #n1%, ". por e-mail o fax (011-4524-1284) días y horarios de pago."
            Print #n1%, ""
            Print #n1%, "En caso de haber ya efectivizado el pago solicitado, favor"
            Print #n1%, "ignorar la presente y comunicarlo por email o fax a fin de"
            Print #n1%, "emitirles el recibo correspondiente."
            Print #n1%, ""
            Print #n1%, "Muchas gracias por su atención. Con cordiales saludos,"
            Print #n1%, ""
            Print #n1%, ".                                    María Florencia Ricci"
            Print #n1%, ".                                      FLEXOFT Soluciones"
            Close #n1%
            '
            '*******************************************************
            '
            TTYY$ = LOADFILE$("C:\FLEXOFT\CODEUCLI.TXT")
            MAI_DESTI$ = rasocli$(NC%)
            MAI_DIREL$ = EMAILCLI$(NC%)
            MAI_ASUNT$ = "Estado y Detalle de Deuda"
            MAI_TEXTO$ = TTYY$
            '
            JKI% = 0
            MAI_ADJUN$(JKI%) = ""
            '
            Call GENERAMAIL
            '
         End If
         Screen.MousePointer = 1
         '
       End If
    Next U&
    '
    Call COMUNI("Colocados en Bandeja de Salida")
    '
End Sub

Private Sub Command17_Click()
Call AVIVENCI
End Sub

Private Sub Command18_Click()
    '
    Command18.Enabled = False
    '
10  OPX% = COMALTER%("Opciones de Trabajo:\\Carga de Interes por Mora\Observaciones Mail Deuda\Mail Interno de Deudas")
    If OPX% < 1 Or OPX% > 3 Then GoTo 99
    '
    If OPX% = 1 Then 'Carga de Interes por Mora
        Call CARGAINTMORA
        GoTo 10
    ElseIf OPX% = 2 Then
        Call OBSERVAMAIL 'Observaciones Mail Deuda
        GoTo 10
    ElseIf OPX% = 3 Then
        Call INGRMAIDEUDAINT 'INGRESO DE MAIL DE DEUDA INTERNO
        GoTo 10
    End If
    '
99  Command18.Enabled = True
End Sub
Sub INGRMAIDEUDAINT()
       '
       'GRABACION DE MAIL INTERNO PARA PASAR REPORTE DIARIO DE DEUDORES
       '
       MAI_INT$ = CONTROL("FAPECOB", "MAIDEUDA")
       MAILINTERNO$ = TRIM$(InputBox("Ingrese Mail Interno de Deudas", "Mail", MAI_INT$))
       If MAILINTERNO$ = "" Then Exit Sub
       '
       'NO VALIDO SI TIENE @ EN ESTE PUNTO PARA QUE SI QUIERE PONER UN MAIL INVALIDO
       'PARA QUE NO SE LO ENVIE PUEDA INGRESARLO.
       Call GRACONTROL("FAPECOB", "MAIDEUDA", MAILINTERNO$)
       '
       
End Sub
Sub OBSERVAMAIL()
      '
      Call COPYSTRU("OBSERVOF", "OBSMAIDE")
      Call TRALOCAL("OBSMAIDE", "OBSERVOF")
      VTB% = VENTABU%("OBSERVOF/TITUPAN=Observaciones de Mail de Deuda")
      If VTB% < 0 Then Exit Sub
      '
      For U& = 1 To ULTREG&("!OBSERVOF")
           TX$ = REGLEIDO$("!OBSERVOF", U&)
           Call REPLA(X$, TX$, 1, 64)
           Call GRAREG("!OBSMAIDE", X$, U&)
      Next U&
      '
      Call SETULTREG("!OBSMAIDE", U& - 1)
      Call TRACENTRAL("OBSMAIDE", "")
      Call CIERRARCH("OBSMAIDE")
      Call CIERRARCH("OBSERVOF")
   
End Sub

Private Sub ComposicionDeudaPorClienteFecha()

    Call SELECHO("SELFECHA")
    
    VDEVU$ = VALACT1$("SELFECHA")
    If TRIM$(VDEVU$) = "" Then Exit Sub
    
    NCLI = ListaSeleccionClientes()
    If NCLI > 0 Then
        Call BLANARCH("FAPECOBA")
        FECHA$ = FETEXCOR1$(FECHANUM(VDEVU$))
        Call RECADEUDA(NCLI, SALIN#, VDEVU$)
        
        Call CIERRARCH("FAPECOBA")
        '
        Screen.MousePointer = 1
        Call HEADERS("FAPECOBA", "")
        Call HEADERS("FAPECOBA", "Fecha de Corte al: " + VDEVU$)
        '
        Call FINAL("*CODEUCLI")
    End If

End Sub

Private Sub Command2_Click()
    '
    Command2.Enabled = False
    Command10.Enabled = False
    '
    '\\\OT-06-0533-RG-23/11/06///
    GRUX$ = ""
    VENDX$ = ""
    n1% = Val(Text1)
    If n1% > 0 Then GRUX$ = Str(n1%) + " - " + ECOARCH("GRUCOCLI", MKI$(n1%))
    N2% = Val(Text3)
    If N2% > 0 Then VENDX$ = Str(N2%) + " - " + ECOARCH("VENDEDOR", Chr$(N2%))
    '\\\OT-06-0533-RG-FIN///
    '
    NCMIN = XVALO(Text10.TEXT)
    NCMAX = XVALO(Text8.TEXT)
    FINF = FECHANUM(Text5.TEXT)
    FLIM = FECHANUM(TRIM$(Text6.TEXT)) ': If FLIM < HOY(HOS$) Then FLIM = HOY(HOS$)
    FEMI1% = FECHANUM(TRIM$(Label9))
    FEMI2% = FECHANUM(TRIM$(Label8))
    '
    If NCMIN < 1 Then
       Call MENSERR(24, "Número de Cuenta no Válido")
       Text10.SetFocus
       GoTo 99
    End If
    '
    If NCMAX < 1 Then
       Call MENSERR(24, "Número de Cuenta no Válido")
       Text8.SetFocus
       GoTo 99
    End If
    '
    If rasocli$(NCMIN) = "" Then
       Call MENSERR(24, "Cuenta Inexistente")
       Text10.SetFocus
       GoTo 99
    End If
    '
    If rasocli$(NCMAX) = "" Then
       Call MENSERR(24, "Cuenta Inexistente")
       Text8.SetFocus
       GoTo 99
    End If
    '
    If FLIM < FINF Then
      Call MENSERR(24, "Rango de Fechas No Procesable")
      GoTo 99
    End If
    '
    ' VALIDA FECHA DE EMISION
    If FEMI2% < FEMI1% Then
      Call MENSERR(24, "Rango de Fechas de Emisión No Procesable")
      GoTo 99
    End If
    '
    Call RECONSTRUYE_DEUDA_CLIENTES_CON_DIFERENCIA
    Call FAPECOB2
    '
    '\\\OT-06-0533-RG-11/12/06///
    If GRUX$ <> "" Then
       Call HEADERS("FAPECOBD", "Grupo:" + GRUX$)
       Call HEADERS("FAPECOBA", "Grupo:" + GRUX$)
    End If
    '
    If VENDX$ <> "" Then
       Call HEADERS("FAPECOBD", "Vendedor:" + VENDX$)
       Call HEADERS("FAPECOBA", "Vendedor:" + VENDX$)
    End If
    '\\\OT-06-0533-RG-FIN///
    '
    If Option1(4).Value = True Then  ' pantalla
         If Option1(3).Value = False Then
            Call FINAL("*COFAPEA1")
            GoTo 99
          Else
            Call FINAL("*COFAPEA2")
            GoTo 99
         End If
       Else                          ' impresora
       '\\\OT-06-0503-RG-23/11/06///
         If Option1(10).Value = True Then
           Call PLANFINA
           GoTo 99
         End If
        '\\\OT-06-0503-RG-FIN///
         MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESVENTA", "MODOMONE"))), 8)
         If (MODOMONE$ = "BIMONETA" Or MODOMONE$ = "MULTIMON") Or InStr(1, EMPREAC$, "FERVI") > 0 Then
             ' Selecciona segun opcion los listados
             If Option1(3).Value = False Then
                AA$ = "*LIFAPE1D"
                If Option1(11) Then
                  AA$ = "*LIFAPE1M"
                  Call GENERAMORA
                End If
                Call FINAL(AA$)
                GoTo 99
              Else
                Call FINAL("*LIFAPE2D")
                GoTo 99
             End If
           Else
             If Option1(3).Value = False Then
                If Check1.Value = 1 Or Check2.Value = 1 Then
                     Call FILESORT("FAPECOBA", "", 12, 12, "ASC")
                     Call TRALOCAL("FAPECOBA", "FAPECOBA")
                     Call BLANARCH("FAPECOBA")
                     JJ& = 0: GRAX$ = ""
                     For U& = 1 To ULTREG&("!FAPECOBA")
                        X$ = REGLEIDO$("!FAPECOBA", U&)
                        GRACT$ = TRIM$(Mid$(X$, 101, 28))
                        If U& > 1 Then
                           If GRACT$ <> GRAX$ Then
                              JJ& = JJ& + 1
                              Call GRAREG("FAPECOBA", String$(512, "="), JJ&)
                           End If
                        End If
                        JJ& = JJ& + 1
                        Call GRAREG("FAPECOBA", X$, JJ&)
                        GRAX$ = GRACT$
                     Next U&
                     Call SETULTREG("FAPECOBA", JJ&)
                     Call FINAL("*LIFAPEG1")
                   Else
                     Call FINAL("*LIFAPEA1")
                End If
                GoTo 99
              Else
                If Check1.Value = 1 Or Check2.Value = 1 Then
                     Call FILESORT("FAPECOBA", "", 12, 12, "ASC")
                     Call TRALOCAL("FAPECOBA", "FAPECOBA")
                     Call BLANARCH("FAPECOBA")
                     JJ& = 0: GRAX$ = ""
                     For U& = 1 To ULTREG&("!FAPECOBA")
                        X$ = REGLEIDO$("!FAPECOBA", U&)
                        GRACT$ = TRIM$(Mid$(X$, 101, 28))
                        If U& > 1 Then
                           If GRACT$ <> GRAX$ Then
                              JJ& = JJ& + 1
                              Call GRAREG("FAPECOBA", String$(512, "="), JJ&)
                           End If
                        End If
                        JJ& = JJ& + 1
                        Call GRAREG("FAPECOBA", X$, JJ&)
                        GRAX$ = GRACT$
                     Next U&
                     Call SETULTREG("FAPECOBA", JJ&)
                     Call FINAL("*LIFAPEG2")
                   Else
                     Call FINAL("*LIFAPEA2")
                End If
                GoTo 99
             End If
         End If
    End If
    '
99  Command10.Enabled = True
    Command2.Enabled = True
    '
End Sub

Private Sub Command2_GotFocus()
    '
    If CARFAPEN% = 0 Then
        Refresh
        Screen.MousePointer = 11
        Call FAPECOB1
    End If
    '
End Sub

Private Sub Command3_Click()
    Call SELECHO("GRUCOCLI")
    Text1.TEXT = VALACT1$("GRUCOCLI")
    If TRIM$(Text1.TEXT) = "" Then
        Text2.TEXT = "Todos - Sin Clasificación"
    End If
End Sub

Private Sub Command4_Click()
    Call SELECHO("VENDEDOR")
    Text3.TEXT = VALACT1$("VENDEDOR")
    If TRIM$(Text3.TEXT) = "" Then
        Text4.TEXT = "Todos - Sin Clasificación"
    End If
End Sub

Private Sub Command5_Click()
    Call FAPECOB1
End Sub

Private Sub Command6_Click()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text6.TEXT = VALACT1$("SELFECHA")
    End If
End Sub


Sub ComposicionDeudaAFecha()
    '
    Dim SALINI#(999999)
    Call SELECHO("SELFECHA")
    
    VDEVU$ = VALACT1$("SELFECHA")
    If VDEVU$ <> "" Then
      FECOR% = FECHANUM(VDEVU$)
      If FECOR% > 0 Then
        Screen.MousePointer = 11
        FECORTE$ = FETEXCOR1$(FECOR%)
        '
        TABORI$ = "CAJABANC"
        'PONE UN SETULTREG DEL ARCHIVO QUE SE USA PARA GRABAR LA DEUDA EN CERO
        Call SETULTREG("FAPECOBA", 0)
        '
        ' DETERMINA SALDO A FECHA POR CUENTA
        Dim CABATEMP As ADODB.Recordset
        Set CABATEMP = New ADODB.Recordset
        CABATEMP.CursorLocation = adUseClient
        Call ABRECONEXION
10      SQLX$ = "SELECT NuClien, SUM(IDEBECOMP-IHABECOMP) AS SALDOIN FROM " + TABORI$ + " "
        SQLX$ = SQLX$ + "WHERE FechConta <= '" & FECORTE$ & "' "
        SQLX$ = SQLX$ + "AND (TipoMovim = 'FVEN' "
        SQLX$ = SQLX$ + "or TipoMovim = 'RCOB' "
        SQLX$ = SQLX$ + "or TipoMovim = 'AJUD') "
        SQLX$ = SQLX$ + "GROUP BY NuClien "
        '
25      On Error Resume Next
        CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 25
        End If
        On Error GoTo 0
        '
        On Error Resume Next
        If Err Then
           On Error GoTo 0
           Screen.MousePointer = 1
           Call MENSERR(24, "Imposible Determinar Saldos")
           GoTo 99
        End If
        On Error GoTo 0
        '
        JU& = 0: CLIMAX = 0
        With CABATEMP
           If Not .BOF And Not .EOF Then
             .MoveLast
             FIN& = .RecordCount
             .MoveFirst
             JU& = 0
             Do While Not .EOF
               JU& = JU& + 1
               Call TRACE(20, JU&, FIN&)
               NCLI = XVALO(!NUCLIEN)
               If NCLI > 0 Then
                 If TICUEPRO((-1) * NCLI) = 0 Then ' solo cuentas corrientes
                   If NCLI > CLIMAX Then CLIMAX = NCLI
                   SALINI#(NCLI) = SALINI#(NCLI) + !SALDOIN
                   If NCLI > CLIMAX Then CLIMAX = NCLI
                 End If
               End If
             .MoveNext
             Loop
           End If
        End With
        CABATEMP.Close
        '
        ' DETERMINA COMPOSICION DE DEUDA DE CADA CLIENTE
        JJ& = 0
        For NCLI = 1 To CLIMAX
           FIN& = CLIMAX
           JJ& = JJ& + 1
           Call TRACE(20, JJ&, FIN&)
           If Abs(SALINI#(NCLI)) >= 0.005 Then
              Call RECADEUDA(NCLI, SALINI#(NCLI), FECORTE$)
           End If
        Next NCLI
        
        Call CIERRARCH("FAPECOBA")
        '
        Screen.MousePointer = 1
        Call HEADERS("FAPECOBA", "")
        Call HEADERS("FAPECOBA", "Fecha de Corte al: " + FECHATEX$(FECOR%))
        '
        Call FINAL("*CODEUCLI")
      End If
    End If
    '
99 End Sub
'
Sub RECADEUDA(NCLI, SALIN#, FECORTE$)

    Call ABRECONEXION
    
    fechaCorte% = CVINT(FECORTE$)
    
    ' *** OBTIENE LOS COMPROBANTES DE LA COMPOSICION DE DEUDA CUYA FECHA DE EMISION SEA MENOR A LA DE CORTE
    ' ***
    Dim rst1 As ADODB.Recordset
    ConsultaSql$ = "SELECT * FROM SADEUPEN WITH(NOLOCK) WHERE NUCLIEN=" & CStr(NCLI)
    ConsultaSql$ = ConsultaSql$ & " AND FECHEMISI <='" & FECORTE$ & "'"
    
    Set rst1 = READSET(FILTSQL$(ConsultaSql$))
    With rst1
        Do While Not .EOF
            comprobanteNro$ = SAFETEXT$(!CodiCom_Cor)
            saldoComprobante# = XVALO(!SaldDebe) - XVALO(!SALDAcre)
            importeOriginal# = XVALO(!IBruComp)
            FECHACOMPROBANTE% = CVINT(!Fechemisi)
            
            GoSub GRADEURECU
            .MoveNext
        Loop
    End With
    
    On Error Resume Next
    rst1.Close
    Set rst1 = Nothing
    On Error GoTo 0


    ' *** OBTIENE LOS DEBITOS ANTERIORES A LA FECHA DE CORTE QUE FUERON CANCELADOS CON CREDITOS POSTERIORES A LA FECHA DE CORTE
    ' ***
    Dim rst2 As ADODB.Recordset
    ConsultaSql$ = "SELECT * FROM CAJABANC WITH(NOLOCK) WHERE NUCLIEN=" & CStr(NCLI)
    ConsultaSql$ = ConsultaSql$ & " AND (TIPOMOVIM='RCOB' OR TIPOMOVIM='AJUD' OR (TIPOMOVIM='FVEN' AND OPCIMOVIM='NC'))" ' FALTA COMPOSICION ENTRE CUENTAS
    ConsultaSql$ = ConsultaSql$ & " AND FECHCONTA >'" & FECORTE$ & "'"
    
    Set rst2 = READSET(FILTSQL$(ConsultaSql$))
    With rst2
        Do While Not .EOF
            reciboNro$ = SAFETEXT$(!codicom_lar)
            Dim rst3 As ADODB.Recordset
            ConsultaSql$ = "SELECT * FROM APLICOBRA WITH(NOLOCK) WHERE NUCLIEN=" & CStr(NCLI)
            ConsultaSql$ = ConsultaSql$ & " AND CODOCAP_LAR='" & reciboNro$ & "'"
            Set rst3 = READSET(FILTSQL$(ConsultaSql$))
            With rst3
                Do While Not .EOF
                    facturaNro$ = SAFETEXT$(!codocor_lar)
                    fechaFactura% = CVINT(VALOBADA("CAJABANC", "FECHCONTA", "CODICOM_LAR='" & facturaNro$ & "' AND NUCLIEN=" & CStr(NCLI)))
                    
                    If fechaFactura% <= fechaCorte% And fechaFactura% > 0 _
                       And Not (InStr(1, EMPREAC$, "FERVI") > 0 And (InStr(1, facturaNro$, "00002356") > 0 Or InStr(1, facturaNro$, "00002357") > 0 Or InStr(1, facturaNro$, "00002889") > 0 Or InStr(1, facturaNro$, "00002907") > 0)) Then
                        comprobanteNro$ = SAFETEXT$(!CODOCOR_COR)
                        saldoComprobante# = XVALO(!IMPOAPLI)
                        importeOriginal# = XVALO(!ITOTORIG)
                        FECHACOMPROBANTE% = fechaFactura%
                        
                        GoSub GRADEURECU
                    End If
                    
                    .MoveNext
                Loop
            End With
            
            On Error Resume Next
            rst3.Close
            Set rst3 = Nothing
            On Error GoTo 0
            
            .MoveNext
        Loop
    End With

    On Error Resume Next
    rst2.Close
    Set rst2 = Nothing
    On Error GoTo 0
    
    
    ' *** OBTIENE LOS CREDITOS ANTERIORES A LA FECHA DE CORTE QUE FUERON APLICADOS CON DEBITOS POSTERIORES A LA FECHA DE CORTE
    ' ***
    Dim rst4 As ADODB.Recordset
    ConsultaSql$ = "SELECT * FROM CAJABANC WITH(NOLOCK) WHERE NUCLIEN=" & CStr(NCLI)
    ConsultaSql$ = ConsultaSql$ & " AND ((TIPOMOVIM='FVEN' AND OPCIMOVIM='FC') OR (TIPOMOVIM='FVEN' AND OPCIMOVIM='ND') OR TIPOMOVIM='AJUD')" ' FALTA COMPOSICION ENTRE CUENTAS
    ConsultaSql$ = ConsultaSql$ & " AND FECHCONTA >'" & FECORTE$ & "'"
    
    Set rst4 = READSET(FILTSQL$(ConsultaSql$))
    With rst4
        Do While Not .EOF
            facturaNro$ = SAFETEXT$(!codicom_lar)
            Dim rst5 As ADODB.Recordset
            ConsultaSql$ = "SELECT * FROM APLICOBRA WITH(NOLOCK) WHERE NUCLIEN=" & CStr(NCLI)
            ConsultaSql$ = ConsultaSql$ & " AND CODOCOR_LAR='" & facturaNro$ & "'"
            Set rst5 = READSET(FILTSQL$(ConsultaSql$))
            With rst5
                Do While Not .EOF
                    reciboNro$ = SAFETEXT$(!codocap_lar)
                    FechaRecibo% = CVINT(VALOBADA("CAJABANC", "FECHCONTA", "CODICOM_LAR='" & reciboNro$ & "' AND NUCLIEN=" & CStr(NCLI)))
                    
                    If FechaRecibo% <= fechaCorte% And FechaRecibo% > 0 And _
                        (Not (InStr(1, EMPREAC$, "FERVI") > 0 And InStr(1, reciboNro$, "1788") > 0)) Then
                        comprobanteNro$ = SAFETEXT$(!CODOCAP_COR)
                        saldoComprobante# = (-1) * XVALO(!IMPOAPLI)
                        importeOriginal# = (-1) * XVALO(!ITOTORIG)
                        FECHACOMPROBANTE% = FechaRecibo%
                        
                        GoSub GRADEURECU
                    End If
                    
                    .MoveNext
                Loop
            End With
            
            On Error Resume Next
            rst5.Close
            Set rst5 = Nothing
            On Error GoTo 0
            
            .MoveNext
        Loop
    End With

    On Error Resume Next
    rst4.Close
    Set rst4 = Nothing
    On Error GoTo 0


    Exit Sub

GRADEURECU:
   SUMADEU# = 0

        Y$ = REGBLAN$("FAPECOBA")
        '
        Call REPLA(Y$, MKS$(NCLI), 1, 4)
        Call REPLA(Y$, rasocli$(NCLI), 5, 28)
        Call REPLA(Y$, TEL1CLI$(NCLI), 33, 16)
        Call REPLA(Y$, comprobanteNro$, 49, 12)
        Call REPLA(Y$, MKI$(FECHACOMPROBANTE%), 61, 2)
        Call REPLA(Y$, MKD$(importeOriginal#), 63, 8)
        Call REPLA(Y$, MKD$(saldoComprobante#), 73, 8)
        '
        If Abs(saldoComprobante#) >= 0.005 Then
           Call REGAPP("FAPECOBA", Y$)
           SUMADEU# = SUMADEU# + saldoComprobante#
        End If

    Return

'    '
'    Dim DOCULAR$(), DOCUCOR$(), IMPO#(), IORI#(), FFDOC%(), FVENC%(), STATDOC%(), CODIEM%(), CODISU%(), SALDOLA#()
'    RAMAX& = 0
'    ATENTI% = 1
'    '
'100 Screen.MousePointer = 11
'    '
'    Call APERBASDAT("CABAN")
'    SQLX$ = "SELECT * FROM CAJABANC "
'    SQLX$ = SQLX$ + "WHERE NuClien = " & NCLI & " "
'    SQLX$ = SQLX$ + "AND (TipoMovim = 'FVEN' "
'    SQLX$ = SQLX$ + "OR TipoMovim = 'AJUD' "
'    SQLX$ = SQLX$ + "OR TipoMovim = 'RCOB') "
'    SQLX$ = SQLX$ + " AND FECHEMISI <= '" & FECORTE$ & "' "
'    SQLX$ = SQLX$ + "ORDER BY FechEmisi ASC;"
'    '
'    Dim CuecoTemp As ADODB.Recordset
'    Set CuecoTemp = New ADODB.Recordset
'    CuecoTemp.CursorLocation = adUseClient
'    Set CuecoTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
'    '
'    SALDO# = 0: SARRAS# = 0: NUDOCU% = 0
'    With CuecoTemp
'      If Not (.BOF And .EOF) Then
'        Do While Not .EOF
'          NUDOCU% = NUDOCU% + 1
'          If NUDOCU% > RAMAX& Then
'            RAMAX& = RAMAX& + 128
'            ReDim Preserve DOCULAR$(RAMAX&), DOCUCOR$(RAMAX&), IORI#(RAMAX&), IMPO#(RAMAX&), FFDOC%(RAMAX&), FVENC%(RAMAX&), STATDOC%(RAMAX&), CODIEM%(RAMAX&), CODISU%(RAMAX&), SALDOLA#(RAMAX&)
'          End If
'          DOCULAR$(NUDOCU%) = TRIM$(!CODICOM_LAR)   ' ANTES ESTABA SIN TRIM - 24/03/07
'          DOCUCOR$(NUDOCU%) = !CODICOM_COR
'          IMPO#(NUDOCU%) = ROUND(!IDEBECOMP - !IHABECOMP)
'          IORI#(NUDOCU%) = !ValAbsComp
'          FFDOC%(NUDOCU%) = CVINT(!FECHEMISI)
'          CODIEM%(NUDOCU%) = !CodiEmpr
'          CODISU%(NUDOCU%) = !CodiSucu
'          SALDOLA#(NUDOCU%) = 0
'          If IsNull(!Status) = False Then STATDOC%(NUDOCU%) = !Status
'          '
'          NRODOC& = XVALO(Mid$(DOCULAR$(NUDOCU%), 11))
' ' pensar aqui como obtener la fecha de vencimiento de cada documento
''          For KKU& = 1 To List2.ListCount
''            TTXY$ = List2.List(KKU& - 1)
''            If XVALO(Mid$(TTXY$, 17, 8)) = NRODOC& Then
''              If FECHANUM(Mid$(TTXY$, 27, 8)) = FFDOC%(NUDOCU%) Then
''                If Abs(Abs(XVALO(Mid$(TTXY$, 37, 14))) - IORI#(NUDOCU%)) < 0.005 Then
''                  FVENC%(NUDOCU%) = FECHANUM(Mid$(TTXY$, 53, 8))
''                  GoTo 119
''                End If
''              End If
''            End If
''          Next KKU&
'          FVENC%(NUDOCU%) = 0
'          '
'119     .MoveNext
'        Loop
'      End If
'      On Error GoTo 0
'    End With
'    '
'    SQLX$ = "SELECT * FROM APLICOBRA "
'    SQLX$ = SQLX$ + "WHERE NuClien = " & NCLI & " "
'    Dim ApliTemp As ADODB.Recordset
'    Set ApliTemp = New ADODB.Recordset
'    ApliTemp.CursorLocation = adUseClient
'    Set ApliTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
'    '
'    With ApliTemp
'      If Not (.BOF And .EOF) Then
'        Do While Not .EOF
'          IAPLI# = XVALO(!IMPOAPLI)
'          RECI$ = TRIM$(SAFETEXT$(!codocap_lar))
'          FACU$ = TRIM$(SAFETEXT$(!codocor_lar))
'          '
'          DOCUAP% = 0: DOCUOR% = 0
'          For KKI% = 1 To NUDOCU%
'            If DOCULAR$(KKI%) = RECI$ And DOCUAP% < 1 Then DOCUAP% = KKI%
'            If DOCULAR$(KKI%) = FACU$ And DOCUOR% < 1 Then DOCUOR% = KKI%
'          Next KKI%
'          '
'          If DOCUAP% > 0 And DOCUOR% > 0 Then
'            SIGAP% = 1: If IAPLI# < 0 Then SIGAP% = (-1)
'            If SIGAP% * IAPLI# > SIGAP% * IMPO#(DOCUOR%) Then IAPLI# = IMPO#(DOCUOR%)
'            IMPO#(DOCUOR%) = IMPO#(DOCUOR%) - IAPLI#
'            IMPO#(DOCUAP%) = IMPO#(DOCUAP%) + IAPLI#
'          End If
'          '
'        .MoveNext
'        Loop
'      End If
'    End With
'    '
'    If InStr(1, EMPREAC$, "AHP") > 0 Or InStr(1, EMPREAC$, "CEYEQ") > 0 Then GoTo 24
'    If InStr(1, EMPREAC$, "FERVI") > 0 Then GoTo 241
'    '
'    SUMAPEN# = 0
'    FIN& = NUDOCU%
'    For KKI% = 1 To NUDOCU%
'      U& = KKI%: 'Call TRACE(20, U&, FIN&)
'      SUMAPEN# = SUMAPEN# + IMPO#(KKI%)
'      If Abs(SUMAPEN#) < 0.005 Then
'        For KKJ% = 1 To KKI%
'          IMPO#(KKJ%) = 0
'        Next KKJ%
'      End If
'    Next KKI%
'    GoTo 242
'    '
'    '
'241 SUMAPEN# = 0
'    FIN& = NUDOCU%
'    For KKI% = 1 To NUDOCU%
'      U& = KKI%: Call TRACE(20, U&, FIN&)
'      SUMAPEN# = SUMAPEN# + IMPO#(KKI%)
'      If FFDOC%(KKI%) > FECHANUM("01/10/15") Then Exit For
'      If Abs(SUMAPEN#) < 0.005 Then
'        For KKJ% = 1 To KKI%
'          IMPO#(KKJ%) = 0
'        Next KKJ%
'      End If
'    Next KKI%
'    '
'242 For KKI% = 1 To NUDOCU%
'      For KKJ% = KKI% + 1 To NUDOCU%
'        If (Abs(IMPO#(KKI%) + IMPO#(KKJ%))) < 0.005 Then
'          IMPO#(KKI%) = 0
'          IMPO#(KKJ%) = 0
'        End If
'      Next KKJ%
'    Next KKI%
'
'    KKM% = 0
'    SUMAPEN# = 0
'    For KKI% = 1 To NUDOCU%
'      If STATDOC%(KKI%) = 8 Then
'        SUMAPEN# = SUMAPEN# + IMPO#(KKI%)
'      End If
'    Next KKI%
'    '
'    If Abs(SUMAPEN#) < 0.005 Then
'      For KKI% = 1 To NUDOCU%
'        If STATDOC%(KKI%) = 8 Then
'          IMPO#(KKI%) = 0
'        End If
'      Next KKI%
'      GoTo 19
'    End If
'    '
'    For KKI% = NUDOCU% To 1 Step -1
'      If STATDOC%(KKI%) = 8 Then
'        If Sgn(IMPO#(KKI%)) <> Sgn(SUMAPEN#) Then
'            IMPO#(KKI%) = 0
'          Else
'            If Abs(SUMAPEN#) >= Abs(IMPO#(KKI%)) Then
'                SUMAPEN# = SUMAPEN# - IMPO#(KKI%)
'              Else
'                IMPO#(KKI%) = SUMAPEN#
'                SUMAPEN# = 0
'            End If
'        End If
'      End If
'    Next KKI%
'    '
'19  If MODOMONEI% = 2 Then       ' cargar saldo en dolares de carbimon.dat
'       RFF$ = RECFILT$("CARBIMON", 2, MKI$(NCLI))
'       FIN& = Len(RFF$)
'       For KKU& = 1 To FIN& Step 4
'          Call TRACE(20, KKU&, FIN&)
'          U& = CVS(Mid$(RFF$, KKU&, 4))
'          XX$ = REGLEIDO$("CARBIMON", U&)
'          NUDOCX$ = TRIM$(Mid$(XX$, 13, 18))
'          DEBEDOL# = ROUND#(CVD(Mid$(XX$, 97, 8)))
'          HABEDOL# = ROUND#(CVD(Mid$(XX$, 105, 8)))
'          '
'          For JJ% = 1 To NUDOCU%
'            If Left$(DOCULAR$(JJ%), 16) = Left$(NUDOCX$, 16) Then
'               SALDOLA#(JJ%) = SALDOLA#(JJ%) + DEBEDOL# - HABEDOL#
'               GoTo 191
'            End If
'          Next JJ%
'          '
'          NUDOCU% = NUDOCU% + 1
'          If NUDOCU% > RAMAX& Then
'            RAMAX& = RAMAX& + 128
'            ReDim Preserve DOCULAR$(RAMAX&), DOCUCOR$(RAMAX&), IORI#(RAMAX&), IMPO#(RAMAX&), FFDOC%(RAMAX&), FVENC%(RAMAX&), STATDOC%(RAMAX&), CODIEM%(RAMAX&), CODISU%(RAMAX&), SALDOLA#(RAMAX&)
'          End If
'          DOCULAR$(NUDOCU%) = NUDOCX$
'          DOCUCOR$(NUDOCU%) = Left$(NUDOCX$, 3) + Right$(NUDOCX$, 8)
'          IMPO#(NUDOCU%) = 0
'          IORI#(NUDOCU%) = CVD(Mid$(XX$, 65, 8))
'          FFDOC%(NUDOCU%) = CVI(Mid$(XX$, 33, 2))
'          CODIEM%(NUDOCU%) = 0
'          CODISU%(NUDOCU%) = 0
'          STATDOC%(NUDOCU%) = 1
'          SALDOLA#(NUDOCU%) = DEBEDOL# - HABEDOL#
'          '
'191    Next KKU&
'    End If
'    '
'    For JJ% = 1 To NUDOCU%
'      For JJ1% = JJ% + 1 To NUDOCU%
'        If Abs(IMPO#(JJ%) + IMPO#(JJ1%)) < 0.005 Then
'          If Abs(SALDOLA#(JJ%) + SALDOLA#(JJ1%)) < 0.005 Then
'            IMPO#(JJ%) = 0
'            SALDOLA#(JJ%) = 0
'            IMPO#(JJ1%) = 0
'            SALDOLA#(JJ1%) = 0
'          End If
'        End If
'      Next JJ1%
'    Next JJ%
'    '
'24  GoSub GRADEURECU
''AQUI GRABAR EN EL LISTADOR - IR A GRADEURECU
'    '
'999 Call CIERRARCH("*.*")
'    Call BAJALDISCO
'    Screen.MousePointer = 1
'    ATENTI% = 0
'    Exit Sub
'    '
'    '
'GRADEURECU:
'   SUMADEU# = 0
'   For i& = 1 To NUDOCU%
'        Y$ = REGBLAN$("FAPECOBA")
'        '
'        Call REPLA(Y$, MKS$(NCLI), 1, 4)
'        Call REPLA(Y$, rasocli$(NCLI), 5, 28)
'        Call REPLA(Y$, TEL1CLI$(NCLI), 33, 16)
'        Call REPLA(Y$, DOCUCOR$(i&), 49, 12)
'        Call REPLA(Y$, MKI$(FFDOC%(i&)), 61, 2)
'        Call REPLA(Y$, MKD$(IORI#(i&)), 63, 8)
'        Call REPLA(Y$, MKD$(IMPO#(i&)), 73, 8)
'        '
'        If Abs(IMPO#(i&)) >= 0.005 Then
'           Call REGAPP("FAPECOBA", Y$)
'           SUMADEU# = SUMADEU# + IMPO#(i&)
'        End If
'        '
'  Next i&
'  '
'  If Abs(SUMADEU# - SALIN#) > 0.005 Then
'     Call MENSERR(24, "No Coincide Saldo en Cuenta " & NCLI)
'  End If
'  '
'Return
''
End Sub

Private Sub Command8_Click()
    Call SELECHO("DEUDOR12")
    NCMAX = XVALO(VALACT1$("DEUDOR12"))
    If NCMAX > 0 Then
        Text8.TEXT = TRIM$(Str$(NCMAX))
        Text7.TEXT = rasocli$(NCMAX)
    End If
End Sub

Private Sub Command9_Click()
    Call SELECHO("DEUDOR12")
    NCMI = XVALO(VALACT1$("DEUDOR12"))
    If NCMI > 0 Then
        Text10.TEXT = TRIM$(Str$(NCMI))
        Text9.TEXT = rasocli$(NCMI)
    End If
End Sub

Sub PECORANFE()
    '
    Call DESHASFECHA(DES%, HAS%, PERIO$)
    If DES% < 1 Then GoTo 99
    
    DES1$ = FECHATEXLAR(DES%)
    HAS1$ = FECHATEXLAR(HAS%)
    '
    CORRESP$ = PERIO$
    Call HEADERS("VEXRAFE1", "")
    Call HEADERS("VEXRAFE1", "Corresponde a: " + CORRESP$)
    '
    
    Dim VEXRAFETMP As ADODB.Recordset
    Set VEXRAFETMP = New ADODB.Recordset
    VEXRAFETMP.CursorLocation = adUseClient
    
    SQLX$ = "SELECT * FROM SADEUPEN "
    SQLX$ = SQLX$ + "WHERE Status < 3 "
'    SQLX$ = SQLX$ + "AND INT(CDBL(FeVencim))<= " & HAS1# & " "
'    SQLX$ = SQLX$ + "AND INT(CDBL(FeVencim))>= " & DES1# & " "
    SQLX$ = SQLX$ + "AND FeVencim<= '" & HAS1$ & "' "
    SQLX$ = SQLX$ + "AND FeVencim>= '" & DES1$ & "' "
    SQLX$ = SQLX$ + "AND (SaldDebe)>= 0.005 "
    SQLX$ = SQLX$ + "ORDER BY FeVencim DESC "
    '
    Call APERBASDAT("CABAN")
    'Set VEXRAFETMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
25  On Error Resume Next
    VEXRAFETMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0
    '
    With VEXRAFETMP
      '
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        Call COMUNI("No Existen Vencimientos en el Rango de Fecha Seleccionado.")
        GoTo 99
      End If
      On Error GoTo 0
      J& = 0
      Y$ = REGBLAN("VEXRAFE1")
      Do While .EOF = False
          '
          NCLI = !NUCLIEN
          NFS$ = !codicom_lar
          FVI% = CVINT(!FEVENCIM)
          SALDEU# = !SaldDebe - !SALDAcre
          '
          Call REPLA(Y$, MKS$(NCLI), 1, 4)
          Call REPLA(Y$, CODICLI$(NCLI), 5, 24)
          Call REPLA(Y$, rasocli$(NCLI), 29, 28)
          Call REPLA(Y$, NFS$, 57, 18)
          Call REPLA(Y$, MKI$(FVI%), 75, 2)
          Call REPLA(Y$, MKD$(SALDEU#), 77, 8)
          '
          J& = J& + 1
          Call GRAREG("VEXRAFE1", Y$, J&)
          '
        .MoveNext
        '
      Loop
      '
      Call SETULTREG("VEXRAFE1", J&)
      Call FILESORT("VEXRAFE1", "VEXRAFE1", 5, 5, "ASC")
      Call FILESORT("VEXRAFE1", "VEXRAFE1", 1, 1, "ASC")
      '
    End With
    VEXRAFETMP.Close
    Set VEXRAFETMP = Nothing
    '
    Call FINAL("*LIVERAF1")
    '
99 End Sub

Private Sub ComposicionDeuda_Click()
    
    Call ComposicionDeudaAFecha
    
End Sub

Private Sub ComposicionDeudaCliente_Click()
    
    Call ComposicionDeudaPorClienteFecha
    
End Sub

Private Sub Form_Load()
    '
    Call APLICACIONSKINS(Me, Picture9)
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
    End If    '
    If EXISTE%(LUDAT$ + "FAPECOBM.RFS") > 0 Then
        Option1(11).Visible = True
    End If
    '
    If RECLEN("FAPECOBA") >= 128 Then
       Check1.Visible = True
       Check2.Visible = True
    End If
    '
    Call FAPECOB1
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
End Sub

Private Sub Label8_Click()
    Call Command14_Click
End Sub

Private Sub Label9_Click()
    Call Command15_Click
End Sub

Private Sub mnuFacpropend_Click()
   CAMPOS$ = "Cta./F7;R.Social/A32;Fecha/D8;N°Prof./F8;Referencia/A48;Imp.Neto/F12.2;Imp.Total/F12.2"
   NCDESDE = XVALO(Text10)
   NCHASTA = XVALO(Text8)
   DES% = FECHANUM(Label9)
   HAS% = FECHANUM(Label8)
   
   Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, CAMPOS$, GRILLAGRAL, 1)
   Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, CAMPOS$, GRILLAGRAL, 1)
   Screen.MousePointer = 11
   SQLX$ = "Select P.NUCLIEN,P.Fechemisi,P.NumeComp,P.REFERENCIA,P.TotNeto,P.ImpoTot,D.RASO_CLI"
   SQLX$ = SQLX$ + " FROM PROFORMTL P WITH(NOLOCK)  "
   SQLX$ = SQLX$ + " INNER JOIN DEUDOR12 D ON D.NUME_CLI = P.NUCLIEN "
   
   SQLX$ = SQLX$ + " WHERE NUCLIEN > 0"
   If NCDESDE > 0 Then SQLX$ = SQLX$ + " and NUCLIEN >= " & NCDESDE
   If NCHASTA > 0 Then SQLX$ = SQLX$ + " and  NUCLIEN <= " & NCHASTA
   
   If DES% > 0 Then SQLX$ = SQLX$ + " and fechemisi >= '" & FETEXCOR1$(DES%) & "'"
   If HAS% > 0 Then SQLX$ = SQLX$ + " and  fechemisi <= '" & FETEXCOR1$(HAS%) & "'"
   
   If MODO% = 0 Then SQLX$ = SQLX$ + " AND (STATUS = 0 OR STATUS IS NULL ) AND STATUS < 9"
  
   SQLX$ = SQLX$ + " ORDER BY FeHoReal ASC"
   Set RS = READSET(SQLX$)
   JJ& = 0
   With RS
        Do While Not RS.EOF
            JJ& = JJ& + 1
            GRILLAGRAL.GRID.Rows = JJ& + 1
            GRILLAGRAL.GRID.TextMatrix(JJ&, 1) = XVALO(!NUCLIEN)
            GRILLAGRAL.GRID.TextMatrix(JJ&, 2) = SAFETEXT$(!raso_cli)
            GRILLAGRAL.GRID.TextMatrix(JJ&, 3) = SAFETEXT$(!Fechemisi)
            GRILLAGRAL.GRID.TextMatrix(JJ&, 4) = XVALO(!NumeComp)
            GRILLAGRAL.GRID.TextMatrix(JJ&, 5) = SAFETEXT$(!REFERENCIA)
            GRILLAGRAL.GRID.TextMatrix(JJ&, 6) = FORMATNUM$(XVALO(!TotNeto), "F12.2")
            GRILLAGRAL.GRID.TextMatrix(JJ&, 7) = FORMATNUM$(XVALO(!ImpoTot), "F12.2")
            .MoveNext
       Loop
   End With
   RS.Close
   Set RS = Nothing
   
   TX$ = "Detalle de Facturas Proformas Pendientes"
   GRILLAGRAL.Caption = TX$
   '*******************************************
   GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
   GRILLAGRAL.mnuMenu1.Visible = True
   GRILLAGRAL.GRID.ZOrder 0
   GRILLAGRAL.TXTBUSCAR = ""
'   Call POSICIONAR_ULTIMAFILA(GRILLAGRAL.GRID)
   Screen.MousePointer = 1
   GRILLAGRAL.Show 1
   Dim MSF As MSFlexGrid
   If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(MSF, CAMPOS$, "PROFPEND", TX$, "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
   End If
   Set obj = Nothing
    
End Sub

Private Sub Option1_Click(Index As Integer)
    If Index = 0 Then
        Option1(0) = False
        Option1(1) = True
    End If
End Sub



Private Sub PendienteCobranzaFecha_Click()

    Call PECORANFE
    
End Sub

Private Sub Text1_DblClick()
    Call SELECHO("GRUCOCLI")
    Text1.TEXT = VALACT1$("GRUCOCLI")
    If TRIM$(Text1.TEXT) = "" Then
        Text2.TEXT = "Todos - Sin Clasificación"
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii% = 13 Then
       Text3.SetFocus
    End If
End Sub

Private Sub Text10_Change()
    If Abs(XVALO(Text10.TEXT)) < 999999 Then
        NCMIN = XVALO(Text10.TEXT)
        Text9.TEXT = rasocli$(NCMIN)
      Else
        Text10.TEXT = ""
        Text9 = ""
    End If
End Sub

Private Sub TEXT10_KeyPress(KeyAscii As Integer)
    If KeyAscii% = 13 Then
       Text8.SetFocus
    End If
End Sub

Private Sub Text10_LostFocus()
    NCMIN = XVALO(Text10.TEXT)
    If NCMIN < 1 Then
       Text9.TEXT = ""
       Exit Sub
    End If
    If rasocli$(NCMIN) = "" Then
       Text9.TEXT = ""
       Exit Sub
    End If
    Text9.TEXT = rasocli$(NCMIN)
End Sub

Private Sub Text2_CLICK()
    Text1.SetFocus
End Sub

Private Sub Text3_DBLCLICK()
    Call SELECHO("VENDEDOR")
    Text3.TEXT = VALACT1$("VENDEDOR")
    If TRIM$(Text3.TEXT) = "" Then
        Text4.TEXT = "Todos - Sin Clasificación"
    End If
End Sub

Private Sub Text1_LostFocus()
    If XVALO(Text1.TEXT) < 256 Then
        GRUCLI% = XVALO(Text1.TEXT)
        GRUX$ = ECOARCH$("GRUCOCLI", MKI$(GRUCLI%))
        If GRUX$ <> "" Then
            Text2.TEXT = GRUX$
            Exit Sub
        End If
    End If
    Text1.TEXT = ""
    Text2.TEXT = "Todos - Sin Clasificación"
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii% = 13 Then
       Text5.SetFocus
    End If
End Sub

Private Sub Text3_LostFocus()
    If XVALO(Text3.TEXT) < 256 Then
        VENDE% = XVALO(Text3.TEXT)
        VENDX$ = ECOARCH$("VENDEDOR", Chr$(VENDE%))
        If VENDX <> "" Then
            Text4.TEXT = VENDX$
            Exit Sub
        End If
    End If
    Text3.TEXT = ""
    Text4.TEXT = "Todos - Sin Clasificación"
End Sub

Private Sub Text4_CLICK()
    Text3.SetFocus
End Sub

Private Sub Text1_Change()
    If TRIM$(Text1.TEXT) = "" Then
        Text2.TEXT = "Todos - Sin Clasificación"
      ElseIf XVALO(Text1.TEXT) < 256 Then
        Check1.Value = 0
        GRUCLI% = XVALO(Text1.TEXT)
        GRUX$ = ECOARCH$("GRUCOCLI", MKI$(GRUCLI%))
        If GRUX$ <> "" Then
          Text2.TEXT = GRUX$
          Exit Sub
        End If
    End If
End Sub

Private Sub Text3_Change()
    If TRIM$(Text3.TEXT) = "" Then
        Text4.TEXT = "Todos - Sin Clasificación"
      ElseIf XVALO(Text3.TEXT) < 256 Then
        Check2.Value = 0
        VENDE% = XVALO(Text3.TEXT)
        VENDX$ = ECOARCH$("VENDEDOR", Chr$(VENDE%))
        If VENDX <> "" Then
          Text4.TEXT = VENDX$
          Exit Sub
        End If
    End If
End Sub

Private Sub Text5_DblClick()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text5.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    If KeyAscii% = 13 Then
       Text6.SetFocus
    End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    If KeyAscii% = 13 Then
       Text10.SetFocus
    End If
End Sub

Private Sub Text8_Change()
    If Abs(XVALO(Text8.TEXT)) <= 999999 Then
        NCMAX = XVALO(Text8.TEXT)
        Text7.TEXT = rasocli$(NCMAX)
      Else
        Text8.TEXT = ""
        Text7 = ""
    End If
End Sub

Private Sub Text10_DblClick()
    Call SELECHO("DEUDOR1")
    NCMIN = XVALO(VALACT1$("DEUDOR1"))
    If NCMIN > 0 Then
        Text10.TEXT = TRIM$(Str$(NCMIN))
        Text9.TEXT = rasocli$(NCMIN)
    End If
End Sub

Private Sub Text6_DBLCLICK()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text6.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Text7_CLICK()
    Text8.SetFocus
End Sub

Private Sub Text8_DblClick()
    Call SELECHO("DEUDOR1")
    NCMAX = XVALO(VALACT1$("DEUDOR1"))
    If NCMAX > 0 Then
        Text8.TEXT = TRIM$(Str$(NCMAX))
        Text7.TEXT = rasocli$(NCMAX)
    End If
End Sub

Private Sub TEXT8_KeyPress(KeyAscii As Integer)
    If KeyAscii% = 13 Then
       Command2.SetFocus
    End If
End Sub

Private Sub Text8_LostFocus()
    NCMAX = XVALO(Text8.TEXT)
    If NCMAX < 1 Then
       Text7.TEXT = ""
       Exit Sub
    End If
    If rasocli$(NCMAX) = "" Then
       Text7.TEXT = ""
       Exit Sub
    End If
    Text7.TEXT = rasocli$(NCMAX)
End Sub

Private Sub Text9_Click()
    Text10.SetFocus
End Sub

Sub COMPODEU()
    '
    Dim NROPO&(8196), CLIPO%(8196)
    Dim SARRACLI#(32767), SUDEUCLI#(32767)
    Dim TIS$(10)
    TDat$ = "AS,VC,FC,ND,RB,DC,NC,AS"
    For i% = 0 To 7
        TIS$(i%) = Mid$(TDat$, 3 * i% + 1, 2)
    Next i%
    '
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) < 1 Then
       Exit Sub
    End If
    '
    FLIM% = FECHANUM(VALACT1$("SELFECHA"))
    '
    FIN& = ULTREG&("FACTUR")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("FACTUR", U&)
      NCLIE% = CVI(Left$(X$, 2))
      FEDOC% = CVI(Mid$(X$, 7, 2))
      If FEDOC% <= FLIM% Then
        TI% = Asc(Mid$(X$, 13, 1))
        If TI% < 0 Or TI% > 7 Then TI% = 2
        SG% = 1: If TI% >= 4 Then If TI% <= 6 Then SG% = -1
        IMPO# = CVD(Mid$(X$, 15, 8))
        IMPO# = IMPO# + CVD(Mid$(X$, 23, 8))
        IMPO# = IMPO# + CVD(Mid$(X$, 31, 8))
        IMPO# = IMPO# + CVD(Mid$(X$, 49, 8))
        IMPO# = IMPO# + CVD(Mid$(X$, 57, 8))
        IMPO# = SG% * IMPO#
        SARRACLI#(NCLIE%) = SARRACLI#(NCLIE%) + IMPO#
        SARRATO# = SARRATO# + IMPO#
      End If
    Next U&
    '
    FIN& = ULTREG&("COBRAN")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("COBRAN", U&)
      NCLIE% = CVI(Left$(X$, 2))
      FEDOC% = CVI(Mid$(X$, 7, 2))
      If FEDOC% <= FLIM% Then
      IMPO# = (-1) * CVD(Mid$(X$, 15, 8))
        SARRACLI#(NCLIE%) = SARRACLI#(NCLIE%) + IMPO#
        SARRATO# = SARRATO# + IMPO#
      End If
    Next U&
    '
    ' LEER SVD202 - DESCARTAR DOCUMENTOS POSTERIORES A FLIM%
    UJ& = 0: SUDEUTO# = 0
    FIN& = ULTREG&("SVD202")
    For UI& = 1 To FIN&
       Call TRACE(20, UI&, FIN&)
       X$ = REGLEIDO$("SVD202", UI&)
       FF% = CVI(Mid$(X$, 19, 2))
       If FF% <= FLIM% Then
         NC% = CVI(Left$(X$, 2))
         If NC% > 0 Then
           If Abs(SARRACLI#(NC%)) >= 0.005 Then
             TI% = Asc(Mid$(X$, 3, 1))
             If TI% < 0 Or TI% > 10 Then TI% = 2
             TIX$ = TIS$(TI%): SIG% = 1
             If TI% > 3 Then SIG% = (-1)
             '
             NF& = CVS(Mid$(X$, 5, 4))
             While NF& > 999999
                NF& = NF& - 100000
             Wend
             NFS$ = TRIM$(Str$(NF&))
             While Len(NFS$) < 8
                NFS$ = " " + NFS$
             Wend
             NFS$ = TIX$ + NFS$
             '
15           FF% = CVI(Mid$(X$, 19, 2))
             FVI% = CVI(Mid$(X$, 9, 2))
             SALDEU# = SIG% * Abs(CVD(Mid$(X$, 11, 8)))
             If TI% = 0 Or TI% = 7 Then
                If Sgn(CVD(Mid$(X$, 21, 8))) < 0 Then
                   SALDEU# = (-1) * Abs(CVD(Mid$(X$, 11, 8)))
                End If
             End If
             If Abs(SALDEU#) < 0.005 Then GoTo 99
             '
             If FF% = 0 Then
                Screen.MousePointer = 1
                Call MENSERR(24, "Imposible Continuar Proceso de Cálculo.\Fecha de Factura no Válida en Cuenta '" + TRIM$(Str$(NC%)) + "'.\Revise la Composición de Deuda de Esta Cuenta.")
                Hide
                Exit Sub
             End If
             '
             IMORIG# = SIG% * CVD(Mid$(X$, 21, 8))
             ES$ = "    "
             If SALDEU# < 0 Then
                   ES$ = " CR "
                 Else
                   'If FVI% <= HOO% + 8 Then ES$ = " ** "
                   'If FVI% <= HOO% Then ES$ = "!!!!"
                   If FVI% < HOO% Then
                     FVIL& = FVI%: HOOL& = HOO%
                     ES$ = FORMATNUM$(DIENTRE%(FVIL&, HOOL&), "F4.0")
                     If Left$(ES$, 1) = " " Then Mid$(ES$, 1, 1) = "."
                   End If
             End If
             '
             Z$ = REGBLAN$("FAPECOB")
             Call REPLA(Z$, MKI$(NC%), 1, 2)
             Call REPLA(Z$, NFS$, 3, 10)
             Call REPLA(Z$, MKI$(FF%), 13, 2)
             Call REPLA(Z$, MKI$(FVI%), 15, 2)
             Call REPLA(Z$, MKD$(SALDEU#), 17, 8)
             Call REPLA(Z$, MKD$(IMORIG#), 25, 8)
             '
             UJ& = UJ& + 1
             Call GRAREG("FAPECOB", Z$, UJ&)
             SUDEUCLI#(NC%) = SUDEUCLI#(NC%) + SALDEU#
             SUDEUTO# = SUDEUTO# + SALDEU#
             '
           End If
         End If
       End If
       '
99   Next UI&
     '
     Call SETULTREG("FAPECOB", UJ&)
     Call CIERRARCH("FAPECOB")
     '
     Screen.MousePointer = 1
     '
    ' LEER FACTURAS - CARGAR VECTOR CON MOVIMIENTOS POSTERIORES A FLIM%
    CAMOPO% = 0
    FIN& = ULTREG&("FACTUR")
    For UI& = 1 To FIN&
      X$ = REGLEIDO$("FACTUR", UI&)
      Tipo% = Asc(Mid$(X$, 13, 1))
      If Tipo% > 0 Then
        If Tipo% < 4 Then
          FFAC% = CVI(Mid$(X$, 7, 2))
          If FFAC% > FLIM% Then
            CAMOPO% = CAMOPO% + 1
            If CAMOPO% > 8196 Then
              Call MENSERR(24, "Imposible Re-Calcular Composición de Deuda\en la Fecha Elegida.")
              Call FINAL("")
            End If
            NROPO&(CAMOPO%) = CVS(Mid$(X$, 3, 4))
            CLIPO%(CAMOPO%) = CVI(Left$(X$, 2))
          End If
        End If
      End If
    Next UI&
    '
    ' LEER APLICOB - DESCARTAR MOVIMIENTOS ANTERIORES A FLIM%
    '              - DESCARTAR MOVIMIENTOS QUE FIGURAN EN MOVIMIENTOS POSTERIORES
    '
    FIN& = ULTREG&("APLICOB")
    For UI& = 1 To FIN&
      X$ = REGLEIDO$("APLICOB", UI&)
      FAPLI% = CVI(Left$(X$, 2))
      If FAPLI% > FLIM% Then
        CLIAPLI% = CVI(Mid$(X$, 3, 2))
        If CLIAPLI% > 0 Then
          If Abs(SARRACLI#(CLIAPLI%) - SUDEUCLI#(CLIAPLI%)) >= 0.005 Then
            DOCUORI$ = Mid$(X$, 55, 18)
            If Mid$(DOCUORI$, 3, 1) <> "-" Or Mid$(DOCUORI$, 8, 1) <> "-" Or Mid$(DOCUORI$, 17, 1) <> "-" Then
   AAA = BBB
            End If
            NUDOAP& = XVALO(Mid$(DOCUORI$, 9, 8))
            For KKI% = 1 To CAMOPO%
              If NROPO&(KKI%) = NUDOAP& Then
                If CLIPO%(KKI%) = CLIAPLI% Then GoTo 29
              End If
            Next KKI%
            '
            Z$ = REGBLAN$("FAPECOB")
            Call REPLA(Z$, MKI$(CLIAPLI%), 1, 2)
             NF& = XVALO(Mid$(DOCUORI$, 10, 7))
             While NF& > 999999: NF& = NF& - 100000: Wend
             NFS$ = TRIM$(Str$(NF&))
             While Len(NFS$) < 8: NFS$ = " " + NFS$: Wend
             NFS$ = Left$(DOCUORI$, 2) + NFS$
            Call REPLA(Z$, NFS$, 3, 10)
            TDCC$ = UCase$(Left$(NFS$, 2))
            TIXA% = 2: If TDCC$ = "ND" Then TIXA% = 3
                       If TDCC$ = "NC" Then TIXA% = 6
                       If TDCC$ = "AJ" Then TIXA% = 0
            FF% = FLIM%
            RFFA$ = RECFILT$("FACTUR", 1, MKI$(CLIAPLI%))
            For KKLL& = 1 To Len(RFFA$) Step 4
              RELEL& = CVS(Mid$(RFFA$, KKLL&, 4))
              ZFF$ = REGLEIDO$("FACTUR", RELEL&)
              If CVS(Mid$(ZFF$, 3, 4)) = NF& Then
                If Asc(Mid$(ZFF$, 13, 1)) = TIXA% Then
                  FF% = CVI(Mid$(ZFF$, 7, 2))
                  GoTo 24
                End If
              End If
            Next KKLL&
            '
24          Call REPLA(Z$, MKI$(FF%), 13, 2)
            FVI% = FAPLI%
            Call REPLA(Z$, MKI$(FVI%), 15, 2)
            SALDEU# = CVD(Mid$(X$, 23, 8))
            Call REPLA(Z$, MKD$(SALDEU#), 17, 8)
            IMORIG# = CVD(Mid$(X$, 73, 8))
            Call REPLA(Z$, MKD$(IMORIG#), 25, 8)
            '
            UJ& = UJ& + 1
            Call GRAREG("FAPECOB", Z$, UJ&)
            SUDEUCLI#(CLIAPLI%) = SUDEUCLI#(CLIAPLI%) + SALDEU#
            SUDEUTO# = SUDEUTO# + SALDEU#
            '
          End If
        End If
      End If
29  Next UI&
    '
    FIN& = ULTREG&("APLICOB")
    For UI& = 1 To FIN&
      X$ = REGLEIDO$("APLICOB", UI&)
      FAPLI% = CVI(Left$(X$, 2))
      If FAPLI% <= FLIM% Then
        CLIAPLI% = CVI(Mid$(X$, 3, 2))
        DIFESAL# = SARRACLI#(CLIAPLI%) - SUDEUCLI#(CLIAPLI%)
        If Abs(DIFESAL#) >= 0.005 Then
          DOCUORI$ = Mid$(X$, 55, 18)
          If Mid$(DOCUORI$, 3, 1) <> "-" Or Mid$(DOCUORI$, 8, 1) <> "-" Or Mid$(DOCUORI$, 17, 1) <> "-" Then
AAA = BBB
          End If
          NUDOAP& = XVALO(Mid$(DOCUORI$, 9, 8))
          DOCUAP$ = Mid$(X$, 5, 3) + Mid$(X$, 16, 7)
          'For KKI% = 1 To CAMOPO%
          '  If NROPO&(KKI%) = NUDOAP& Then
          '    If CLIPO%(KKI%) = CLIAPLI% Then GoTo 39
          '  End If
          'Next KKI%
          '
          Z$ = REGBLAN$("FAPECOB")
          Call REPLA(Z$, MKI$(CLIAPLI%), 1, 2)
             NF& = XVALO(Mid$(DOCUAP$, 4, 7))
             While NF& > 999999: NF& = NF& - 100000: Wend
             NFS$ = TRIM$(Str$(NF&))
             While Len(NFS$) < 8: NFS$ = " " + NFS$: Wend
             NFS$ = Left$(DOCUAP$, 2) + NFS$
          Call REPLA(Z$, NFS$, 3, 10)
          FF% = FLIM%
          Call REPLA(Z$, MKI$(FF%), 13, 2)
          FVI% = FAPLI%
          Call REPLA(Z$, MKI$(FVI%), 15, 2)
          SALDEU# = CVD(Mid$(X$, 23, 8))
          Call REPLA(Z$, MKD$((-1) * SALDEU#), 17, 8)
          IMORIG# = CVD(Mid$(X$, 73, 8))
          Call REPLA(Z$, MKD$(IMORIG#), 25, 8)
          '
          UJ& = UJ& + 1
          Call GRAREG("FAPECOB", Z$, UJ&)
          SUDEUCLI#(CLIAPLI%) = SUDEUCLI#(CLIAPLI%) - SALDEU#
          SUDEUTO# = SUDEUTO# - SALDEU#
          '
        End If
      End If
39  Next UI&
    '
    FIN& = 32767
    For UI& = 1 To FIN&
      CLIAPLI% = UI&
      DIFESAL# = SARRACLI#(CLIAPLI%) - SUDEUCLI#(CLIAPLI%)
      If Abs(DIFESAL#) >= 0.005 Then
          '
          Z$ = REGBLAN$("FAPECOB")
          Call REPLA(Z$, MKI$(CLIAPLI%), 1, 2)
          NFS$ = "Dif.Saldo"
          Call REPLA(Z$, NFS$, 3, 10)
          FF% = FLIM%
          Call REPLA(Z$, MKI$(FF%), 13, 2)
          FVI% = FLIM%
          Call REPLA(Z$, MKI$(FVI%), 15, 2)
          SALDEU# = DIFESAL#
          Call REPLA(Z$, MKD$(SALDEU#), 17, 8)
          IMORIG# = CVD(Mid$(X$, 73, 8))
          Call REPLA(Z$, MKD$(DIFESAL#), 25, 8)
          '
          UJ& = UJ& + 1
          Call GRAREG("FAPECOB", Z$, UJ&)
          SUDEUCLI#(CLIAPLI%) = SUDEUCLI#(CLIAPLI%) + SALDEU#
          '
      End If
49  Next UI&
    '
    Call SETULTREG("FAPECOB", UJ&)
    Call CIERRARCH("FAPECOB")
    '              - AGREGAR A LISTA DE SVD202
    ' ORDENAR POR NUMERO O ALFABETICO Y POR FECHA EMISION
    '
    GRUCLI% = XVALO(Text1.TEXT)
    GRUX$ = TRIM$(Text2.TEXT)
    VENDE% = XVALO(Text3.TEXT)
    VENDX$ = TRIM$(Text4.TEXT)
    NCMIN% = XVALO(Text10.TEXT)
    NCMAX% = XVALO(Text8.TEXT)
    '
    MODORDE% = 1
    If Option1(2) = True Then MODORDE% = 2
    If Option1(3) = True Then MODORDE% = 3
    '
    FIN& = ULTREG&("FAPECOB")
    UJ& = 0
    List1.Clear
    '
    For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("FAPECOB", U&)
        NC% = CVI(Left$(X$, 2))
        '
        If NC% >= NCMIN% Then
          If NC% <= NCMAX% Then
            FVI% = CVI(Mid$(X$, 15, 2))
            If GRUCLI% = 0 Or GRUCOCLI%(NC%) = GRUCLI% Then
              If VENDE% = 0 Or VENDCLI%(NC%) = VENDE% Then
                '
                TLISTA$ = Space$(36)
                If MODORDE% = 1 Then
                     Call REPLA(TLISTA$, CSTRING$(MKI$(NC%), 1, 6, 0, 2), 1, 6)
                  ElseIf MODORDE% = 2 Then
                     Call REPLA(TLISTA$, rasocli$(NC%), 1, 30)
                  ElseIf MODORDE% = 3 Then
                     Call REPLA(TLISTA$, CSTRING$(MKI$(FVI%), 1, 6, 0, 2), 1, 6)
                End If
                Call REPLA(TLISTA$, CSTRING$(MKS$(U&), 3, 6, 0, 2), 31, 6)
                List1.AddItem TLISTA$
                '
              End If
            End If
          End If
        End If
        '
    Next U&
    '
    FIN& = List1.ListCount
    '
    For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        List1.ListIndex = U& - 1
        REG& = XVALO(Mid$(List1.TEXT, 31, 8))
        X$ = REGLEIDO$("FAPECOB", REG&)
        '
        NC% = CVI(Left$(X$, 2))
        NFS$ = Mid$(X$, 3, 10)
        FF% = CVI(Mid$(X$, 13, 2))
        FVI% = CVI(Mid$(X$, 15, 2))
        SALDEU# = CVD(Mid$(X$, 17, 8))
        IMORIG# = CVD(Mid$(X$, 25, 8))
        ES$ = ""   ': If FVI% < HOY(HOS$) Then ES$ = "!!!!"
        If FVI% < HOY(HOS$) Then
           FVIL& = FVI%: HOOL& = HOY(HOS$)
           ES$ = FORMATNUM$(DIENTRE%(FVIL&, HOOL&), "F4.0")
           If Left$(ES$, 1) = " " Then Mid$(ES$, 1, 1) = "."
        End If        '
        Y$ = REGBLAN$("FAPECOBA")
        Call REPLA(Y$, MKI$(NC%), 1, 2)
        Call REPLA(Y$, rasocli$(NC%), 3, 30)
        Call REPLA(Y$, TEL1CLI$(NC%), 33, 16)
        Call REPLA(Y$, NFS$, 49, 12)
        Call REPLA(Y$, MKI$(FF%), 61, 2)
        Call REPLA(Y$, MKD$(IMORIG#), 63, 8)
        Call REPLA(Y$, MKI$(FVI%), 71, 2)
        Call REPLA(Y$, MKD$(SALDEU#), 73, 8)
        Call REPLA(Y$, ES$, 97, 4)
        '
        UJ& = UJ& + 1
        Call GRAREG("FAPECOBA", Y$, UJ&)
        '          '
    Next U&
    '
    Call SETULTREG("FAPECOBA", UJ&)
    Call CIERRARCH("FAPECOBA")
    '
    FEX = FLIM%
    Call HEADERS("FAPECOBA", "")
    Call HEADERS("FAPECOBA", "Composición de Deuda al: " + FECHATEX$(FEX))
    '
    Call FINAL("*LIFAPEA1")
    '
    Screen.MousePointer = 1

End Sub

Sub FAPECOB1()
    '
    Screen.MousePointer = 11
    '
    NCMIN = 999999: NCMAX = 0
    For U& = 1 To 32767: TICUE%(U&) = 1: Next U&
    NCMIN = XVALO(VALOBADA("DEUDOR12", "MIN(Nume_Cli)", "Nume_Cli > 0"))
    NCMAX = XVALO(VALOBADA("DEUDOR12", "MAX(Nume_Cli)", "Nume_Cli <= 999999"))
    '
    Text10 = TRIM$(Str$(NCMIN))
    Text9 = rasocli$(NCMIN)
    Text8 = TRIM$(Str$(NCMAX))
    Text7 = rasocli$(NCMAX)
    '
    Text1 = ""
    Text3 = ""
    '
    SQLX$ = "SELECT * FROM SADEUPEN WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE Status < 3 "
    SQLX$ = SQLX$ + " ORDER BY FeVencim DESC "
    '
    Call APERBASDAT("CABAN")
    Dim SACTEMP As ADODB.Recordset
    Set SACTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    
    On Error Resume Next
    FLIM = HOY(HOS$)
    With SACTEMP
      .MoveFirst
      If Err < 1 Then
        FLIM = CVINT(!FEVENCIM)
      End If
    End With
    SACTEMP.Close
    Set SACTEMP = Nothing
    On Error GoTo 0
    'RECHA MAX
    
    Text6 = FECHATEX$(FLIM)
    Text6.Refresh
    'FECHA MIN
    Text5.TEXT = "01/01/80"
    Text5.Refresh
    '
    CARFAPEN% = 1
    Screen.MousePointer = 1
    '
End Sub
'
Sub FAPECOB2()
    '
    Screen.MousePointer = 11
    '
    HOII% = HOY(HOS$)
    MODOMONEI% = 0
    MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESVENTA", "MODOMONE"))), 8)
    If MODOMONE$ = "MULTIMON" Then
       MODOMONEI% = 2
    End If
    '
    FF1% = FECHANUM(Text5.TEXT)
    FDESD = FETEXCOR1$(FF1%)
    FF2% = FECHANUM(Text6.TEXT)
    FLIM = FETEXCOR1$(FF2%)
    FEMI1% = FECHANUM(TRIM$(Label9))
    FEMI2% = FECHANUM(TRIM$(Label8))
    '
    NCMIN = XVALO(Text10.TEXT)
    NCMAX = XVALO(Text8.TEXT)
    '
    GRUCLI% = XVALO(Text1.TEXT)
    GRUX$ = TRIM$(Text2.TEXT)
    VENDE% = XVALO(Text3.TEXT)
    VENDX$ = TRIM$(Text4.TEXT)
    MODORDE% = 1
    If Option1(2) = True Then MODORDE% = 2
    If Option1(3) = True Then MODORDE% = 3
    '
    On Error Resume Next
    Kill LUDAT$ + "FAPECOBA.DAT"
    On Error GoTo 0
    On Error Resume Next
    Kill LUDAT$ + "FAPECOBD.DAT"
    On Error GoTo 0
    On Error Resume Next
    Kill LUDAT$ + "FAPECOBR.DAT"
    On Error GoTo 0
    '
    BIMONET% = 0
    MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESVENTA", "MODOMONE"))), 8)
    If MODOMONE$ = "BIMONETA" Then
         BIMONET% = 1
       ElseIf MODOMONE$ = "MULTIMON" Then
         BIMONET% = 2
    End If
    '
    REBLX$ = REGBLAN$("FAPECOBA")
    REBLY$ = REGBLAN$("FAPECOBD")
    REBLZ$ = REGBLAN$("FAPECOBR")
    '
'    SQLX$ = "SELECT * FROM SADEUPEN "
'    SQLX$ = SQLX$ + "WHERE Status < 3 "
'    SQLX$ = SQLX$ + "AND NuClien >= " & NCLI & " "
'    SQLX$ = SQLX$ + "AND NuClien <= " & NCMAX & " "
'    SQLX$ = SQLX$ + "AND FeVencim <= '" & FLIM & "' "
'    SQLX$ = SQLX$ + "AND FeVencim >= '" & FDESD & "' "
'    SQLX$ = SQLX$ + "ORDER BY FeVencim DESC "
'    '
    SQLX$ = "SELECT * FROM SADEUPEN "
    If GRUCLI% > 0 Or VENDE% > 0 Then
       SQLX$ = SQLX$ + "INNER JOIN DEUDOR12 "
       SQLX$ = SQLX$ + "ON SADEUPEN.NUCLIEN = DEUDOR12.Nume_Cli "
    End If
    SQLX$ = SQLX$ + "WHERE SADEUPEN.Status < 3 "
    SQLX$ = SQLX$ + "AND (ABS(SADEUPEN.SALDdebe - SADEUPEN.SALDAcre) >= 0.005 "
    SQLX$ = SQLX$ + "OR ABS(SADEUPEN.SALDdebe_Dola - SADEUPEN.SALDAcre_Dola) >= 0.005) "
    SQLX$ = SQLX$ + "AND SADEUPEN.NuClien >= " & NCMIN & " "
    SQLX$ = SQLX$ + "AND SADEUPEN.NuClien <= " & NCMAX & " "
    SQLX$ = SQLX$ + "AND SADEUPEN.FeVencim <= '" & FLIM & "' "
    SQLX$ = SQLX$ + "AND SADEUPEN.FeVencim >= '" & FDESD & "' "
    SQLX$ = SQLX$ + "AND SADEUPEN.CODIEMPR = " & CodiEmp% & " "
      If GRUCLI% > 0 Then SQLX$ = SQLX$ + "AND DEUDOR12.GrCo_Cli = " & GRUCLI% & " "
      If VENDE% > 0 Then SQLX$ = SQLX$ + "AND DEUDOR12.VEND_Cli = " & VENDE% & " "
      If FEMI1% > 0 Then SQLX$ = SQLX$ + "AND FechEmisi >='" & FETEXCOR1$(FEMI1%) & "' "
      If FEMI2% > 0 Then SQLX$ = SQLX$ + "AND FechEmisi <='" & FETEXCOR1$(FEMI2%) & "' "
    SQLX$ = SQLX$ + "ORDER BY SADEUPEN.FeVencim DESC "
    '
    Call APERBASDAT("CABAN")
    '
    Dim SACTEMP As ADODB.Recordset
    Set SACTEMP = New ADODB.Recordset
    SACTEMP.CursorLocation = adUseClient
25  On Error Resume Next
    SACTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0
    '
    With SACTEMP
      On Error Resume Next
      .MoveLast
      FIN& = .RecordCount
      .MoveFirst
      If Err < 1 Then
        UJ& = 0: UK& = 0: UL& = 0
        Do While Not .EOF
          UL& = UL& + 1
          Call TRACE(20, UL&, FIN&)
          NCLI = !NUCLIEN
          NFS$ = !CodiCom_Cor
          COCOLAR$ = !codicom_lar
          FF% = CVINT(!Fechemisi)
          FVI% = CVINT(!FEVENCIM)
          SALDEU# = XVALO(!SaldDebe) - XVALO(!SALDAcre)
          SALDOL# = 0
          On Error Resume Next
          SALDOL# = XVALO(!SaldDebe_DOLA) - XVALO(!SaldAcre_DOLA)
          On Error GoTo 0
          '
          If MODOMONEI% < 2 Then
             If Abs(SALDEU#) < 0.005 Then
                SALDOL# = 0
                GoTo 29
             End If
          End If
          IMORIG# = !IBruComp
          ES$ = "    "
          MONEMX% = 1
          On Error Resume Next
          MONEMX% = !Mone_Emis
          On Error GoTo 0
          '
          If MODOMONEI% = 2 Then
             Call APERBASDAT("CABAN")
             ' AGREGUE LAS CONDICIONES DE TipoMovim =
             ' PORQUE SI NO A VECES TRAE EL NETO O EL IVA - EPB 24/05/07
             'CajaBanc.FindFirst "CodiCom_Lar = '" & COCOLAR$ & "' AND (TipoMovim = 'FVEN' OR TipoMovim = 'RCOB' OR TipoMovim = 'AJUD')"
             SQLX$ = " SELECT * FROM CAJABANC "
             SQLX$ = SQLX$ + " WHERE CodiCom_Lar = '" & COCOLAR$ & "'"
             SQLX$ = SQLX$ + " AND (TipoMovim = 'FVEN' "
             SQLX$ = SQLX$ + " OR TipoMovim = 'RCOB' "
             SQLX$ = SQLX$ + " OR TipoMovim = 'AJUD')"
             Dim CAJJJABANC As ADODB.Recordset
             Set CAJJJABANC = New ADODB.Recordset
26           On Error Resume Next
             CAJJJABANC.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
             If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
               On Error GoTo 0
               Call CapturaErrorOpen
               GoTo 26
             End If
             On Error GoTo 0
             '
             With CAJJJABANC
              'If CajaBanc.NoMatch = False Then
              On Error Resume Next
              If Not (.EOF And .BOF) Then
                 On Error GoTo 0
                 IMORIG# = !Impo_Dola
                 If Abs(IMORIG#) < 0.005 Or Sgn(IMORIG#) <> Sgn(SALDEU#) Then
                   IPESO# = ROUND(!Idebecomp - !IHABECOMP)
                   MONEMX% = !Mone_Emis
                   TICAM = !Tipo_Cam: If TICAM < 0.005 Then TICAM = 1
                   IMORIG# = ROUND#(IPESO# / TICAM)
                 End If
              End If
              On Error GoTo 0
             End With
             CAJJJABANC.Close
             Set CAJJJABANC = Nothing
          End If
          '
        Y$ = REBLX$
        Call REPLA(Y$, MKS$(NCLI), 1, 4)
        Call REPLA(Y$, rasocli$(NCLI), 5, 28)
        Call REPLA(Y$, TEL1CLI$(NCLI), 33, 16)
        Call REPLA(Y$, NFS$, 49, 12)
        Call REPLA(Y$, MKI$(FF%), 61, 2)
        Call REPLA(Y$, MKD$(IMORIG#), 63, 8)
        Call REPLA(Y$, MKI$(FVI%), 71, 2)
        Call REPLA(Y$, MKD$(SALDEU#), 73, 8)
        Call REPLA(Y$, ES$, 97, 4)
        '
        GRUVE$ = ""
        If Check1.Value = 1 Then
             GRUVE$ = "Grupo: " + ECOARCH$("GRUCOCLI", MKI$(GRUCOCLI%(NCLI)))
          ElseIf Check2.Value = 1 Then
             GRUVE$ = "Vendedor: " + ECOARCH$("VENDEDOR", Chr$(VENDCLI%(NCLI)))
        End If
        Call REPLA(Y$, GRUVE$, 101, 28)
        '
        UJ& = UJ& + 1
        Call GRAREG("FAPECOBA", Y$, UJ&)
        '
        If BIMONET% < 1 And InStr(1, EMPREAC$, "FERVI") < 1 Then GoTo 29
        '
        Y$ = REBLY$
        Z$ = REBLZ$
        '
        Call REPLA(Y$, MKS$(NCLI), 1, 4)
        Call REPLA(Y$, rasocli$(NCLI), 5, 28)
        Call REPLA(Y$, TEL1CLI$(NCLI), 33, 16)
        Call REPLA(Y$, NFS$, 49, 10)
        Call REPLA(Y$, MKI$(FF%), 59, 2)
20              If MONEMX% > 1 Then
           Call REPLA(Y$, SIMBOLPES$(MONEMX%), 61, 4)
        End If
        Call REPLA(Y$, MKD$(IMORIG#), 65, 8)
        Call REPLA(Y$, MKI$(FVI%), 73, 2)
        Call REPLA(Y$, MKD$(SALDEU#), 75, 8)
        Call REPLA(Y$, MKD$(SALDOL#), 91, 8)
        ES$ = ""
        If FVI% < HOY(HOS$) Then
           FVIL& = FVI%: HOOL& = HOY(HOS$)
           ES$ = FORMATNUM$(DIENTRE%(FVIL&, HOOL&), "F4.0")
           If Left$(ES$, 1) = " " Then Mid$(ES$, 1, 1) = "."
           ES$ = "Atr" + ES$
        End If
        If Abs(SALDEU#) >= 0.005 Then
          If Abs(SALDOL#) < 0.005 Then
            If MONEMX% > 1 Then
              ES$ = "Ajustar"
            End If
          End If
        End If
        Call REPLA(Y$, ES$, 107, 10)
        DESCRIPAGOX$ = ""
        If Left$(COCOLAR$, 3) = "FC-" Then
            CONDI$ = " CodiCom_Lar = '" & COCOLAR$ & "'"
            CONDI$ = CONDI$ + " AND (TipoMovim = 'FVEN')"
            CP% = XVALO(VALOBADA("CAJABANC", "CONDPAG", CONDI$, "NOMESS"))
            DESCRIPAGOX$ = Left$(ECOARCH$("CONPAG", MKI$(CP%)), 12)
        End If
        Call REPLA(Y$, DESCRIPAGOX$, 117, 12)
        '
        Call REPLA(Z$, Left$(Y$, 60), 1, 60)
        Call REPLA(Z$, MKD$(IMORIG#), 61, 8)
        Call REPLA(Z$, MKI$(FVI%), 69, 2)
        Call REPLA(Z$, MKD$(SALDEU#), 71, 8)
        Call REPLA(Z$, ES$, 87, 10)
        '
        UK& = UK& + 1
        Call GRAREG("FAPECOBR", Z$, UK&)
        Call GRAREG("FAPECOBD", Y$, UK&)
        '
29      .MoveNext
        Loop
        '
      End If
    End With
    SACTEMP.Close
    Set SACTEMP = Nothing
    '
    Call SETULTREG("FAPECOBA", UJ&)
    Call CIERRARCH("FAPECOBA")
    Call SETULTREG("FAPECOBR", UK&)
    Call CIERRARCH("FAPECOBR")
    Call SETULTREG("FAPECOBD", UK&)
    Call CIERRARCH("FAPECOBD")
    '
    Call HEADERS("FAPECOBA", "")
    Call HEADERS("FAPECOBA", "Vencimiento Hasta: " + Text6.TEXT)
    Call HEADERS("FAPECOBA", "Rango de Cuentas: " + TRIM$(Str$(NCLI)) + " - " + TRIM$(Str$(NCMAX)))
    Call HEADERS("FAPECOBR", "")
    Call HEADERS("FAPECOBR", "Vencimiento Hasta: " + Text6.TEXT)
    Call HEADERS("FAPECOBR", "Rango de Cuentas: " + TRIM$(Str$(NCLI)) + " - " + TRIM$(Str$(NCMAX)))
    Call HEADERS("FAPECOBD", "")
    Call HEADERS("FAPECOBD", "Vencimiento Hasta: " + Text6.TEXT)
    Call HEADERS("FAPECOBD", "Rango de Cuentas: " + TRIM$(Str$(NCLI)) + " - " + TRIM$(Str$(NCMAX)))
    '
    Call FILESORT("FAPECOBA", "", 7, 7, "ASC")
    Call FILESORT("FAPECOBD", "", 8, 8, "ASC")
    Select Case MODORDE%
      Case 1
        Call FILESORT("FAPECOBA", "", 1, 1, "ASC")
        Call FILESORT("FAPECOBD", "", 1, 1, "ASC")
      Case 2
        Call FILESORT("FAPECOBA", "", 2, 2, "ASC")
        Call FILESORT("FAPECOBD", "", 2, 2, "ASC")
    End Select
    Screen.MousePointer = 1
    '
End Sub

Private Sub GENERAMORA()

  NCMIN% = XVALO(Text10.TEXT)
  NCMAX% = XVALO(Text8.TEXT)

  Call SETULTREG("FAPECOBM", 0)
  
  JJ& = 0
  TXT$ = REGBLAN$("FAPECOBM")
  For UU& = 1 To ULTREG&("FAPECOBD")
        Y$ = REGLEIDO$("FAPECOBD", UU&)
    
        FVI% = CVI(Mid$(Y$, 71, 2))
        Call REPLA(TXT$, Mid$(Y$, 1, 83), 1, 83)
        ' CALCULAR LA MORA
        CANTDIAS# = DIENTRE(FVI%, HOY(HO$))
        IMPOPE# = CVD(Mid$(TXT$, 75, 8))
        
        'No calcula Mora en los recibos ni notas de credito
        If IMPOPE# > 0.05 Then
           INTERES# = XVALO(CONTROL("VENTAS", "INTMORA"))
           MORADI# = (IMPOPE# * INTERES#) / 100
           MORAT# = MORADI# * CANTDIAS#
           TOTACONMORA# = IMPOPE# + MORAT
        Else
           MORADI# = 0
           MORAT# = 0
           TOTACONMORA# = IMPOPE#
           CANTDIAS# = 0
        End If
        '
        JJ& = JJ& + 1
        Call REPLA(TXT$, MKD(CANTDIAS#), 91, 8)
        Call REPLA(TXT$, MKD(MORAT#), 99, 8)
        Call REPLA(TXT$, MKD(TOTACONMORA#), 115, 8)
        Call GRAREG("FAPECOBM", TXT$, JJ&)
  Next UU&
  
  Call SETULTREG("FAPECOBM", JJ&)
  
  Call HEADERS("FAPECOBM", "")
  Call HEADERS("FAPECOBM", "Vencimiento Hasta: " + Text6.TEXT)
  Call HEADERS("FAPECOBM", "Rango de Cuentas: " + TRIM$(Str$(NCMIN%)) + " - " + TRIM$(Str$(NCMAX%)))
  '
End Sub

Sub FAPECOB2_VIE()
    'NUEVA
    '/////**********//////
    'FALTA TESTEAR'
    '/////*****************//////////
    Screen.MousePointer = 11
    '
    MODOMONEI% = 0
    MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESVENTA", "MODOMONE"))), 8)
    If MODOMONE$ = "MULTIMON" Then
       MODOMONEI% = 2
    End If
    '
    FDESD = CDbl(CVDAT(Int(FECHANUM(Text5.TEXT))))
    FLIM = CDbl(CVDAT(Int(FECHANUM(Text6.TEXT))))
    NCMIN% = XVALO(Text10.TEXT)
    NCMAX% = XVALO(Text8.TEXT)
    '
    GRUCLI% = XVALO(Text1.TEXT)
    GRUX$ = TRIM$(Text2.TEXT)
    VENDE% = XVALO(Text3.TEXT)
    VENDX$ = TRIM$(Text4.TEXT)
    MODORDE% = 1
    If Option1(2) = True Then MODORDE% = 2
    If Option1(3) = True Then MODORDE% = 3
    '
    On Error Resume Next
    Kill LUDAT$ + "FAPECOBA.DAT"
    On Error GoTo 0
    On Error Resume Next
    Kill LUDAT$ + "FAPECOBD.DAT"
    On Error GoTo 0
    On Error Resume Next
    Kill LUDAT$ + "FAPECOBR.DAT"
    On Error GoTo 0
    '
    BIMONET% = 0
    MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESVENTA", "MODOMONE"))), 8)
    If MODOMONE$ = "BIMONETA" Then
         BIMONET% = 1
       ElseIf MODOMONE$ = "MULTIMON" Then
         BIMONET% = 2
    End If
    '
    REBLX$ = REGBLAN$("FAPECOBA")
    REBLY$ = REGBLAN$("FAPECOBD")
    REBLZ$ = REGBLAN$("FAPECOBR")
    '
    SQLX$ = "SELECT * FROM SADEUPEN "
    SQLX$ = SQLX$ + "WHERE Status < 3 "
    SQLX$ = SQLX$ + "AND NuClien >= " & NCMIN% & " "
    SQLX$ = SQLX$ + "AND NuClien <= " & NCMAX% & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FeVencim))<= " & FLIM & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FeVencim))>= " & FDESD & " "
    SQLX$ = SQLX$ + "ORDER BY FeVencim DESC "
    '
    Call APERBASDAT("CABAN")
    'Set SACTEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim SACTEMP As ADODB.Recordset
    Set SACTEMP = New ADODB.Recordset
25  On Error Resume Next
    SACTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0
    '
    On Error Resume Next
    With SACTEMP
      .MoveLast
      FIN& = .RecordCount
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        UJ& = 0: UK& = 0: UL& = 0
        Do While Not .EOF
          UL& = UL& + 1
          Call TRACE(20, UL&, FIN&)
          NC% = !NUCLIEN
          NFS$ = !CodiCom_Cor
          COCOLAR$ = !codicom_lar
          FF% = CVINT(!Fechemisi)
          FVI% = CVINT(!FEVENCIM)
          SALDEU# = !SaldDebe - !SALDAcre
          IMORIG# = !IBruComp
          ES$ = "    "
          '
          If MODOMONEI% = 2 Then
             Call APERBASDAT("CABAN")
             CajaBanc.FindFirst "CodiCom_Lar = '" & COCOLAR$ & "' "
             If CajaBanc.NoMatch = False Then
                On Error Resume Next
                IMORIG# = CajaBanc!Impo_Dola
                On Error GoTo 0
             End If
          End If
          '
          If GRUCLI% = 0 Or GRUCOCLI%(NC%) = GRUCLI% Then
            If VENDE% = 0 Or VENDCLI%(NC%) = VENDE% Then
              If Abs(SALDEU#) > 0.005 Then
                Y$ = REBLX$
                Call REPLA(Y$, MKI$(NC%), 1, 2)
                Call REPLA(Y$, rasocli$(NC%), 3, 30)
                Call REPLA(Y$, TEL1CLI$(NC%), 33, 16)
                Call REPLA(Y$, NFS$, 49, 12)
                Call REPLA(Y$, MKI$(FF%), 61, 2)
                Call REPLA(Y$, MKD$(IMORIG#), 63, 8)
                Call REPLA(Y$, MKI$(FVI%), 71, 2)
                Call REPLA(Y$, MKD$(SALDEU#), 73, 8)
                Call REPLA(Y$, ES$, 97, 4)
                '
                GRUVE$ = ""
                If Check1.Value = 1 Then
                     GRUVE$ = "Grupo: " + ECOARCH$("GRUCOCLI", MKI$(GRUCOCLI%(NC%)))
                  ElseIf Check2.Value = 1 Then
                     GRUVE$ = "Vendedor: " + ECOARCH$("VENDEDOR", Chr$(VENDCLI%(NC%)))
                End If
                Call REPLA(Y$, GRUVE$, 101, 28)
                '
                UJ& = UJ& + 1
                Call GRAREG("FAPECOBA", Y$, UJ&)
                '
                If BIMONET% < 1 Then GoTo 29
                '
                Y$ = REBLY$
                Z$ = REBLZ$
                '
                Call REPLA(Y$, MKI$(NC%), 1, 2)
                Call REPLA(Y$, rasocli$(NC%), 3, 30)
                Call REPLA(Y$, TEL1CLI$(NC%), 33, 16)
                Call REPLA(Y$, NFS$, 49, 10)
                Call REPLA(Y$, MKI$(FF%), 59, 2)
                Call REPLA(Y$, MKD$(IMORIG#), 65, 8)
                Call REPLA(Y$, MKI$(FVI%), 73, 2)
                Call REPLA(Y$, MKD$(SALDEU#), 75, 8)
                '
                Call REPLA(Z$, Left$(Y$, 60), 1, 60)
                Call REPLA(Z$, MKD$(IMORIG#), 61, 8)
                Call REPLA(Z$, MKI$(FVI%), 69, 2)
                Call REPLA(Z$, MKD$(SALDEU#), 71, 8)
                '
                MONEMX% = 0
                On Error Resume Next   ' por si no está el campo
                MONEMX% = !Mone_Emis
                On Error GoTo 0
                If MONEMX% < 1 Then
                  COCOLAR$ = TRIM$(!codicom_lar)
                  RFG$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
                  For KUL& = Len(RFG$) - 3 To 1 Step -4
                     RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
                     XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
                     If TRIM$(Mid$(XXC$, 45, 18)) = COCOLAR$ Then
                        MONEMX% = CVI(Mid$(XXC$, 31, 2))
                        GoTo 15
                     End If
                  Next KUL&
15                If MONEMX% < 1 Then MONEMX% = 1
                End If
                '
                If MONEMX% > 1 Then
                  SALDOL# = 0
                  RFG$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
                  For KUL& = Len(RFG$) - 3 To 1 Step -4
                    RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
                    XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
                    If Mid$(XXC$, 13, 18) = COCOLAR$ Then
                      SALDOL# = SALDOL# + CVD(Mid$(XXC$, 97, 8)) - CVD(Mid$(XXC$, 105, 8))
                    End If
                  Next KUL&
                  Call REPLA(Y$, MKD$(SALDOL# * TICAMONE(MONEMX%)), 75, 8)
                  Call REPLA(Y$, MKD$(SALDOL#), 91, 8)
                End If
                '
20              If MONEMX% > 1 Then
                   Call REPLA(Y$, SIMBOLPES$(MONEMX%), 61, 4)
                End If
                Call REPLA(Y$, ES$, 107, 10)
                Call REPLA(Z$, ES$, 87, 10)
                '
                UK& = UK& + 1
                Call GRAREG("FAPECOBR", Z$, UK&)
                Call GRAREG("FAPECOBD", Y$, UK&)
                '          '
29            End If
            End If
          End If
          '
        .MoveNext
        Loop
        '
      End If
      On Error GoTo 0
    End With
    Set SACTEMP = Nothing
    '
    Call SETULTREG("FAPECOBA", UJ&)
    Call CIERRARCH("FAPECOBA")
    Call SETULTREG("FAPECOBR", UK&)
    Call CIERRARCH("FAPECOBR")
    Call SETULTREG("FAPECOBD", UK&)
    Call CIERRARCH("FAPECOBD")
    '
    Call HEADERS("FAPECOBA", "")
    Call HEADERS("FAPECOBA", "Vencimiento Hasta: " + Text6.TEXT)
    Call HEADERS("FAPECOBA", "Rango de Cuentas: " + TRIM$(Str$(NCMIN%)) + " - " + TRIM$(Str$(NCMAX%)))
    Call HEADERS("FAPECOBR", "")
    Call HEADERS("FAPECOBR", "Vencimiento Hasta: " + Text6.TEXT)
    Call HEADERS("FAPECOBR", "Rango de Cuentas: " + TRIM$(Str$(NCMIN%)) + " - " + TRIM$(Str$(NCMAX%)))
    Call HEADERS("FAPECOBD", "")
    Call HEADERS("FAPECOBD", "Vencimiento Hasta: " + Text6.TEXT)
    Call HEADERS("FAPECOBD", "Rango de Cuentas: " + TRIM$(Str$(NCMIN%)) + " - " + TRIM$(Str$(NCMAX%)))
    '
    Call FILESORT("FAPECOBA", "", 7, 7, "ASC")
    Call FILESORT("FAPECOBD", "", 8, 8, "ASC")
    Select Case MODORDE%
      Case 1
        Call FILESORT("FAPECOBA", "", 1, 1, "ASC")
        Call FILESORT("FAPECOBD", "", 1, 1, "ASC")
      Case 2
        Call FILESORT("FAPECOBA", "", 2, 2, "ASC")
        Call FILESORT("FAPECOBD", "", 2, 2, "ASC")
    End Select
    Screen.MousePointer = 1
    '
End Sub

Private Sub PLANFINA()
  '\\\OT-06-0503-RG-23/11/06///
  Dim PLAFIN#(32767, 5) 'MATRIZ QUE CONTIENE EL PLAN FINANCIERO
  Dim FECOSE(5) As Integer 'FECHA COMIENZO DE SEMANA
  '
  FLIM1% = HOY(HO$)
  NUMSEM% = SEMANU%(FLIM1%, DELX, ALX)
  FECOSE%(1) = DELX 'FECHA COMIENZO SEMANA UNO
  FECOSE%(2) = DIPOST(ALX) 'FECHA COMIENZO SEMANA DOS
  NUSEM% = SEMANU%(FECOSE%(2), DELX, ALX)
  FECOSE%(3) = DIPOST(ALX) 'FECHA COMIENZO SEMANA TRES
  NUSEM% = SEMANU%(FECOSE%(3), DELX, ALX)
  FECOSE%(4) = DIPOST(ALX) 'FECHA COMIENZO SEMANA CUATRO
  NUSEM% = SEMANU%(FECOSE%(4), DELX, ALX)
  FECOSE%(5) = DIPOST(ALX) 'FECHA COMIENZO SEMANA CUATRO
  FECOSE%(0) = (-32767)
  '
'1.- GENERAR FACTURAS PENDIENTES DE COBRANZA (FABECOBR)
  'Call FAPECOB2
'2.- DISTRIBUIR PENDIENTES EN LA MATRIZ
  FIN& = ULTREG&("FAPECOBR")
  NCMAX% = 0
  For JX& = 1 To FIN&
    '
    Call TRACE(24, JX&, FIN&)
    '
    X$ = REGLEIDO$("FAPECOBR", JX&)
    '
    FV1% = CVI(Mid$(X$, 69, 2))
    NPRO1% = CVI(Mid$(X$, 1, 2))
    IMPEN1# = CVD(Mid$(X$, 71, 8))
    If NPRO1% > 0 Then
      If NPRO1% > NCMAX% Then NCMAX% = NPRO1%
    End If
    '
    For KKI% = 0 To 4
      If FV1% >= FECOSE%(KKI%) Then  'VENCIDAS ANTERIORES
        If FV1% < FECOSE%(KKI% + 1) Then
          PLAFIN#(NPRO1%, KKI%) = PLAFIN#(NPRO1%, KKI%) + IMPEN1#
          GoTo 19
        End If
      End If
    Next KKI%
    PLAFIN#(NPRO1%, 5) = PLAFIN#(NPRO1%, 5) + IMPEN1#
19 Next JX&
   '
'3.- GENERAR ARCHIVO  PLAN FINANCIERO A PARTIR DE LA MATRIZ
    '
  REBLAZ$ = REGBLAN$("PLANFIN")
  JJ& = 0
  FIN& = NCMAX%
  For JX1% = 1 To FIN&
    '
    REGACT1& = JX1%
    Call TRACE(24, REGACT1&, FIN&)
    '
    TOTAPRO1# = 0
    For XB% = 0 To 5
     TOTAPRO1# = TOTAPRO1# + PLAFIN#(JX1%, XB%)
    Next XB%
    '
    If Abs(TOTAPRO1#) > 0.005 Then 'SI TIENE PENDIENTES
      Y$ = REBLAZ$
      Call REPLA(Y$, MKI$(JX1%), 1, 2)
      Call REPLA(Y$, MKD$(PLAFIN#(JX1%, 0)), 3, 8) 'VEN
      Call REPLA(Y$, MKD$(PLAFIN#(JX1%, 1)), 11, 8) '1SEM
      Call REPLA(Y$, MKD$(PLAFIN#(JX1%, 2)), 19, 8) '2SEM
      Call REPLA(Y$, MKD$(PLAFIN#(JX1%, 3)), 27, 8) '3SEM
      Call REPLA(Y$, MKD$(PLAFIN#(JX1%, 4)), 35, 8) '4SEM
      Call REPLA(Y$, MKD$(PLAFIN#(JX1%, 5)), 43, 8) 'POST
      Call REPLA(Y$, MKD$(TOTAPRO1#), 51, 8)        'TOTAL
      JJ& = JJ& + 1
      Call GRAREG("PLANFIN", Y$, JJ&)
    End If
    '
  Next JX1%
  Call SETULTREG("PLANFIN", JJ&)
'5.- LISTAR
  Call FINAL("*PLANFIN")
  '\\\OT-06-0503-RG-FIN///
End Sub
