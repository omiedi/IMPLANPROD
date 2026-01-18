VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form OPPED07 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pedidos de Clientes - Modificaciones, Consultas y Listados"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5805
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
   ScaleHeight     =   6435
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   3720
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   56
      Top             =   840
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cumplimiento de Entregas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   210
      TabIndex        =   44
      Top             =   6510
      Width           =   4635
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Index           =   7
         Left            =   1050
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   52
         Top             =   480
         Width           =   2430
         Begin VB.CommandButton Command26 
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
            Picture         =   "OPPED07.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Control de Cumplimiento de Entregas por Rango de Fechas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Control  por Ran- go de Fechas"
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
            TabIndex        =   54
            Top             =   45
            Width           =   1530
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Listados de Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Left            =   2625
         TabIndex        =   45
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command25 
            Caption         =   "Análisis A-B-C"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   51
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command24 
            Caption         =   "Estadísticas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   50
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command19 
            Caption         =   "Críticos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   49
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command14 
            Caption         =   "Consignación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   48
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command15 
            Caption         =   "Negativos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   47
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command18 
            Caption         =   "Inventarios"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   46
            Top             =   420
            Width           =   1590
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      Caption         =   "GENERACION DE PEDIDOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   120
      TabIndex        =   34
      Top             =   2880
      Width           =   4635
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Index           =   8
         Left            =   1080
         ScaleHeight     =   615
         ScaleWidth      =   2400
         TabIndex        =   35
         Top             =   420
         Width           =   2430
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
            Picture         =   "OPPED07.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Copia de Pedido"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Generación por Copia de Pedido"
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
            Left            =   735
            TabIndex        =   37
            Top             =   80
            Width           =   1530
         End
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      ForeColor       =   &H80000008&
      Height          =   8415
      Left            =   4920
      ScaleHeight     =   8385
      ScaleWidth      =   1140
      TabIndex        =   27
      Top             =   0
      Width           =   1170
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   105
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   58
         Top             =   5640
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   59
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton AGenRep 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "OPPED07.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   2280
         Width           =   650
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Direct"
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
         Left            =   105
         Picture         =   "OPPED07.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Accesos Directos"
         Top             =   1600
         Width           =   650
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Stk."
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
         Left            =   105
         Picture         =   "OPPED07.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Listados Estadísticos de Ventas y Pedidos"
         Top             =   4740
         Width           =   650
      End
      Begin VB.CommandButton Command10 
         Caption         =   "List"
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
         Left            =   105
         Picture         =   "OPPED07.frx":0D72
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Histórico de Precios por Pedido"
         Top             =   4050
         Width           =   650
      End
      Begin VB.ListBox List1 
         Height          =   300
         Left            =   105
         Sorted          =   -1  'True
         TabIndex        =   40
         Top             =   5145
         Visible         =   0   'False
         Width           =   645
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
         Left            =   105
         Picture         =   "OPPED07.frx":107C
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Otras Opciones - Ranking de Pedidos por Rango de Fechas"
         Top             =   3045
         Width           =   650
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Check"
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
         Left            =   420
         Picture         =   "OPPED07.frx":11C6
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Control de Integridad de la B.Datos"
         Top             =   1890
         Visible         =   0   'False
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
         Height          =   700
         Left            =   105
         Picture         =   "OPPED07.frx":14D0
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Salir"
         Top             =   210
         Width           =   650
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   105
         Picture         =   "OPPED07.frx":161A
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Ayuda"
         Top             =   900
         Width           =   650
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2D3C0&
      Caption         =   "CONSULTAS Y LISTADOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   8
      Top             =   4200
      Width           =   4635
      Begin VB.CommandButton Command9 
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
         Left            =   4200
         Picture         =   "OPPED07.frx":1924
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Pedidos Pendientes por Cliente"
         Top             =   420
         Visible         =   0   'False
         Width           =   400
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   420
         ScaleHeight     =   615
         ScaleWidth      =   3765
         TabIndex        =   12
         Top             =   405
         Width           =   3795
         Begin VB.CommandButton Command64 
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
            Left            =   3180
            Picture         =   "OPPED07.frx":1D66
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Pedidos Pendientes por Código de Producto"
            Top             =   0
            Width           =   600
         End
         Begin VB.CommandButton Command62 
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
            Left            =   2570
            Picture         =   "OPPED07.frx":2070
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Pedidos Pendientes por Cliente"
            Top             =   0
            Width           =   600
         End
         Begin VB.CommandButton Command61 
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
            Picture         =   "OPPED07.frx":24B2
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Encabezados de Pedidos"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command63 
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
            Left            =   1970
            Picture         =   "OPPED07.frx":27BC
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Detalle por Número de Pedido / Listado Pedidos sin Facturar"
            Top             =   0
            Width           =   600
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Base de Datos de Pedidos"
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
            Left            =   630
            TabIndex        =   17
            Top             =   105
            Width           =   1185
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2D3C0&
         Caption         =   "DISPOSITIVO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   610
         Left            =   420
         TabIndex        =   9
         Top             =   1245
         Width           =   3795
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Pantalla"
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
            Height          =   225
            Left            =   525
            TabIndex        =   11
            Top             =   270
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Impresión"
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
            Height          =   225
            Left            =   2205
            TabIndex        =   10
            Top             =   270
            Width           =   1170
         End
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E2D3C0&
      Caption         =   "BASE DE DATOS DE PEDIDOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2730
      Left            =   120
      TabIndex        =   0
      Top             =   50
      Width           =   4635
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Index           =   4
         Left            =   1485
         ScaleHeight     =   615
         ScaleWidth      =   2400
         TabIndex        =   31
         Top             =   1840
         Width           =   2430
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
            Height          =   630
            Left            =   0
            Picture         =   "OPPED07.frx":2AC6
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Aprobar Pedidos de Clientes."
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Aprobar Pedidos de Clientes"
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
            Left            =   735
            TabIndex        =   33
            Top             =   105
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Index           =   1
         Left            =   1200
         ScaleHeight     =   615
         ScaleWidth      =   2400
         TabIndex        =   24
         Top             =   1360
         Width           =   2430
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
            Height          =   630
            Left            =   0
            Picture         =   "OPPED07.frx":2DD0
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Anulación de Pedidos"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Anulación por Items de Pedidos "
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
            TabIndex        =   26
            Top             =   50
            Width           =   1545
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Index           =   2
         Left            =   945
         ScaleHeight     =   615
         ScaleWidth      =   2400
         TabIndex        =   21
         Top             =   880
         Width           =   2430
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
            Left            =   1840
            Picture         =   "OPPED07.frx":2F1A
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Revertir Anulación de Pedidos"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command45 
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
            Picture         =   "OPPED07.frx":3064
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Anulación de Pedidos"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Anulación de Pedidos "
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
            Left            =   735
            TabIndex        =   22
            Top             =   45
            Width           =   1065
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Index           =   0
         Left            =   630
         ScaleHeight     =   615
         ScaleWidth      =   2445
         TabIndex        =   18
         Top             =   400
         Width           =   2475
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
            Height          =   600
            Left            =   1920
            Picture         =   "OPPED07.frx":336E
            Style           =   1  'Graphical
            TabIndex        =   60
            ToolTipText     =   "Recuperar Pedido Con Remito Anulado"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command46 
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
            Picture         =   "OPPED07.frx":34B8
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Modificación de Pedido en Curso."
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Modificación de Pedido en Curso"
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
            TabIndex        =   19
            Top             =   45
            Width           =   1440
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Listados de Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Left            =   2625
         TabIndex        =   1
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command16 
            Caption         =   "Análisis A-B-C"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   7
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command17 
            Caption         =   "Estadísticas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   6
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command20 
            Caption         =   "Críticos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   5
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command21 
            Caption         =   "Consignación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   4
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command22 
            Caption         =   "Negativos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   3
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command23 
            Caption         =   "Inventarios"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   2
            Top             =   420
            Width           =   1590
         End
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "OPPED07.frx":37C2
      Top             =   1800
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "OPPED07.frx":39F6
      TabIndex        =   57
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Line Line1 
      X1              =   5250
      X2              =   6120
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Menu mnuedit 
      Caption         =   "Edición"
      Visible         =   0   'False
      Begin VB.Menu agregarobservaciondeitem 
         Caption         =   "Agregar Observación a Item en Pedido"
      End
   End
End
Attribute VB_Name = "OPPED07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Sub PEDIPENGRILLA(XCLIENTE, XCODIGO%)
   Call ABREPEDCLI
   Screen.MousePointer = 11
   Campos$ = "NroPed/F8;Cta./F8;R.Social/A24;Fecha/d8;Código/A16;Descripción/A24;C.Pedida/f10.1;P.Unit./F12.2;Entreg./F10.1;Saldo/F10.1"
   IncluyeAtributoComercial% = 0
   If CantRegistros("datasql", "lista='ATRICOPE'") > 0 Then
        Dim da As New DatosAuxiliares
        Campos$ = Campos$ & ";Id/I5;Comercial/A32"
        IncluyeAtributoComercial% = 1
   End If
   
   Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
   Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
   J& = 0
   SQLX$ = " SELECT * "
   SQLX$ = SQLX$ + " FROM PEDDETA WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE Status < 8 "
   SQLX$ = SQLX$ + " AND STATUS <> -1 " ' EXCLUYE LAS NO APROBADAS
   SQLX$ = SQLX$ + " AND (((CanPed - CantEnt) >= 0.05) OR (TipoPed = 'ABIERTO' AND FinVigen >= getdate()))"
   SQLX$ = SQLX$ + " ORDER BY NROCLIE"
   Set RS = READSET(SQLX$)
   With RS
     Do While Not .EOF
       J& = J& + 1
       '
       NRO& = XVALO(!NROPED)
         FECHAPED% = CVINT(!FECHEMIS)
       FECHAPEX$ = FECHATEX$(FECHAPED%)
       NC = XVALO(!NROCLIE)
       RACLI$ = SAFETEXT$(!RasoClie)
       CI1% = XVALO(!CODIINT)
       CODI$ = SAFETEXT$(!CODIEXT)
       DESKRIP$ = SAFETEXT$(!Descrip)
       CAPE = XVALO(!CanPed)
       PREUNIT# = XVALO(!PreUnid)
       CAENTR = XVALO(!cantent)
       SALDO = CAPE - CAENTR

       If SALDO < 0 Then SALDO = 0
           GRILLAGRAL.GRID.Rows = J& + 1
           GRILLAGRAL.GRID.TextMatrix(J&, 1) = NRO&
            NumeroCliente$ = CStr(NC)
            While Len(NumeroCliente$) < 7: NumeroCliente$ = "0" & NumeroCliente$: Wend
           GRILLAGRAL.GRID.TextMatrix(J&, 2) = NumeroCliente$
           GRILLAGRAL.GRID.TextMatrix(J&, 3) = RACLI$
           GRILLAGRAL.GRID.TextMatrix(J&, 4) = FECHAPEX$
           GRILLAGRAL.GRID.TextMatrix(J&, 5) = CODI$
           
           GRILLAGRAL.GRID.TextMatrix(J&, 6) = DESKRIP$
           GRILLAGRAL.GRID.TextMatrix(J&, 7) = FORMATNUM$(CAPE, "F10.1")
           GRILLAGRAL.GRID.TextMatrix(J&, 8) = FORMATNUM$(PREUNIT#, "F12.2")
           GRILLAGRAL.GRID.TextMatrix(J&, 9) = FORMATNUM$(CAENTR, "F10.1")
           GRILLAGRAL.GRID.TextMatrix(J&, 10) = FORMATNUM$(SALDO, "F10.1")
           If IncluyeAtributoComercial% = 1 Then
                If NroPedidoAnterior& <> NRO& Then NroPedidoAnterior& = NRO&: idAtributo% = XVALO(VALOBADA("pedenca", "atributocomercial", "nroped=" & CStr(NRO&)))
                If idAtributo% > 0 Then
                    DenominacionAtributo$ = da.EcoTablaAuxiliar("ATRICOPE", CStr(idAtribut%), 1, 2)
                    GRILLAGRAL.GRID.TextMatrix(J&, 11) = CStr(idAtributo%)
                    GRILLAGRAL.GRID.TextMatrix(J&, 12) = DenominacionAtributo$
                End If
           End If
       .MoveNext
     Loop
   End With
   RS.Close
   Set RS = Nothing
   '
   Screen.MousePointer = SCMA
    
    TX$ = "Pedidos Pendientes" & PERIO$
    GRILLAGRAL.Caption = TX$
    '*******************************************
    GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.TXTBUSCAR = ""
    Call POSICIONAR_ULTIMAFILA(GRILLAGRAL.GRID)
    
    GRILLAGRAL.SinEfectoClickEnGrilla% = 1
    
    If XCODIGO% > 0 Then
      GRILLAGRAL.CMBFILTRO = "Código"
      GRILLAGRAL.TXTBUSCAR = CODEXT$(XCODIGO%)
    ElseIf XCLIENTE > 0 Then
      GRILLAGRAL.CMBFILTRO = "Cta."
      NumeroCliente$ = CStr(XCLIENTE)
      While Len(NumeroCliente$) < 7: NumeroCliente$ = "0" & NumeroCliente$: Wend
      GRILLAGRAL.TXTBUSCAR = NumeroCliente$
    End If
    
    GRILLAGRAL.Show 1
    Dim MSF As msFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(MSF, Campos$, "ENCAPEDI", TX$, "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
    End If

End Sub

 Sub AGenRep_Click()
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

Sub agregarobservaciondeitem_Click()
10 Call SELECHO("!INDIPEP")
   NPED& = XVALO(VALACT1$("!INDIPEP"))
        
  ' VALIDA SI ELIGE PEDIDO
  If NPED& < 1 Then Exit Sub
  '
  Call CARGA_ENCABEZADO(MUESTRADATOS8.MSF1, "Código/A16;Descripción/A28;Pedido/A8;Entregado/A8;Saldo/A8;F.Sol./a8;Observación/A36;Codint/F0;Nped/F0", MUESTRADATOS8, , , ANTOT)

  SQLX$ = "SELECT CODIINT,CodiExt ,Descrip, CanPed, CantEnt,ObservaItem,FeSolEnt FROM PEDDETA WITH(NOLOCK)"
  SQLX$ = SQLX$ + " WHERE NROPED = " & NPED&
  Set RS = READSET(SQLX$)
  With RS
    Do While Not .EOF
      CI1% = XVALO(!CODIINT)
      CANTIPEDI = FORMATNUM$(XVALO(!CanPed), "F10.1")
      CANTIENT = FORMATNUM$(XVALO(!cantent), "F10.1")
      SALDO = XVALO(CANTIPEDI) - XVALO(CANTIENT)
      If SALDO < 0.05 Then SALDO = 0
      '
      MUESTRADATOS8.MSF1.Rows = MUESTRADATOS8.MSF1.Rows + 1
      MUESTRADATOS8.MSF1.TextMatrix(MUESTRADATOS8.MSF1.Rows - 1, 1) = SAFETEXT$(!CODIEXT)
      MUESTRADATOS8.MSF1.TextMatrix(MUESTRADATOS8.MSF1.Rows - 1, 2) = SAFETEXT$(!Descrip)
      MUESTRADATOS8.MSF1.TextMatrix(MUESTRADATOS8.MSF1.Rows - 1, 3) = CANTIPEDI
      MUESTRADATOS8.MSF1.TextMatrix(MUESTRADATOS8.MSF1.Rows - 1, 4) = CANTIENT
      MUESTRADATOS8.MSF1.TextMatrix(MUESTRADATOS8.MSF1.Rows - 1, 5) = FORMATNUM$(SALDO, "F10.1")
      MUESTRADATOS8.MSF1.TextMatrix(MUESTRADATOS8.MSF1.Rows - 1, 6) = FECHATEX$(CVINT(!FeSolEnt))
      OBSERVACIONITEM$ = RTrim$(SAFETEXT$(!ObservaItem))
      MUESTRADATOS8.Label1 = OBSERVACIONITEM$
      MUESTRADATOS8.MSF1.RowHeight(MUESTRADATOS8.MSF1.Rows - 1) = MUESTRADATOS8.Label1.Height + 100
      MUESTRADATOS8.MSF1.TextMatrix(MUESTRADATOS8.MSF1.Rows - 1, 7) = OBSERVACIONITEM$
      MUESTRADATOS8.MSF1.TextMatrix(MUESTRADATOS8.MSF1.Rows - 1, 8) = CI1%
      MUESTRADATOS8.MSF1.TextMatrix(MUESTRADATOS8.MSF1.Rows - 1, 9) = NPED&
  
    .MoveNext
    Loop
  End With
  RS.Close
  Set RS = Nothing
  '
  MUESTRADATOS8.Caption = "Observaciones de Items de Pedido N°: " & NPED&
  MUESTRADATOS8.Command2.Visible = False
  MUESTRADATOS8.Show 1
  GoTo 10
  
End Sub


Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub


Sub Command10_Click()
   '
   '\\\OT-07-0092-RG-08/03/07///
   Command10.Enabled = False
   '
   If DERACCE%("HISPRECI", "Histórico de Precios por Pedido") >= 2 Then
       Call DESHASFECHA(Des%, Has%, PERIO$)
       If PERIO$ = "" Then GoTo 99
       DESDAT = CDbl(CVDAT(Des%))
       HASDAT = CDbl(CVDAT(Has%))
       '
       Call ABREPEDCLI
       'jandy
       Desde = FETEXCOR1$(Des%)
       Hasta = FETEXCOR1$(Has%)
       SQLX$ = "SELECT * FROM PEDDETA "
       'sqlx$ = sqlx$ + "WHERE INT(CDBL(FechEmis)) >= " & DESDAT & " "
       'sqlx$ = sqlx$ + "AND INT(CDBL(FechEmis)) <= " & HASDAT & " "
       SQLX$ = SQLX$ + "WHERE FechEmis >= '" & Desde & "' "
       SQLX$ = SQLX$ + "AND FechEmis <= '" & Hasta & "' "
       
       '
       'On Error GoTo 0
       'Set PEDDETTEMP = Ventas.OpenRecordset(sqlx$, dbOpenSnapshot)
       '
       Dim PEDDETTEMP As ADODB.Recordset
       Set PEDDETTEMP = New ADODB.Recordset
       PEDDETTEMP.CursorLocation = adUseClient
25     On Error Resume Next
       PEDDETTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
          On Error GoTo 0
          Call CapturaErrorOpen
          GoTo 25
       End If
       On Error GoTo 0
       With PEDDETTEMP
           On Error Resume Next
           .MoveFirst
           If Err > 0 Then
               On Error GoTo 0
               Call MENSERR(24, "Sin Movimientos en el Período Elegido")
               GoTo 99
           End If
           On Error GoTo 0
           '
           X& = 0
           '
           Do While Not .EOF
               NCLIE = !NROCLIE
               If NCLIE > 0 Then
                  NART% = !CODIINT
                  FEDOC% = CVINT(!FECHEMIS)
                  NPED& = !NROPED
                  ORCOM$ = !NROOCOM
                  CANTI# = !CanPed
                  PUNIT# = !PreUnid
                  '28/03/07 (OT-07-0092)
                  MONEMIS1% = !MoneEmis
                  PREORIG$ = ""
                  '
                  If MONEMIS1% > 1 Then
                     SIMBO$ = SIMBOLPES$(MONEMIS1%)
                     PREORIG$ = TRIM$(AJUSTI$(SIMBO$, 4) + FORMATNUM$(PUNIT#, "F12.4"))
                  End If
                  '28/03/07 (FIN)
                  '
                  Z$ = REGBLAN$("HISPRECI")
                  Call REPLA(Z$, MKS$(NCLIE), 1, 4)
                  Call REPLA(Z$, MKI$(NART%), 5, 2)
                  Call REPLA(Z$, MKI$(FEDOC%), 7, 2)
                  Call REPLA(Z$, MKS$(NPED&), 9, 4)
                  Call REPLA(Z$, ORCOM$, 13, 16)
                  PRECONV# = PUNIT#
                  If !MoneEmis > 1 Then
                     '28/03/07 (OT-07-0092)
                     PRECONV# = PUNIT# * !TIPOCAM
                     '28/03/07 (FIN)
                  End If
                  Call REPLA(Z$, MKD$(PRECONV#), 29, 8)
                  Call REPLA(Z$, MKS$(CANTI#), 37, 4)
                  IMPO# = CANTI# * PRECONV#
                  Call REPLA(Z$, MKD$(IMPO#), 41, 8)
                  Call REPLA(Z$, PREORIG$, 49, 20)
                  '
                  X& = X& + 1
                  Call GRAREG("HISPRECI", Z$, X&)
                  '
                End If
                '
           .MoveNext
           Loop
       End With
       PEDDETTEMP.Close
       Set PEDDETTEMP = Nothing
       '
       Call SETULTREG("HISPRECI", X&)
       '
       Call FILESORT("HISPRECI", "", 1, 5, "ASC") 'CLIENTE
       '
       '\\\OT-07-0092-RG-12/03/07///
       Call HEADERS("HISPRECI", "")
       Call HEADERS("HISPRECI", "Período: " + PERIO$)
       '\\\OT-07-0092-RG-FIN///
       '
       Call FINAL("*HISPRECI")
   End If
   '
99 Screen.MousePointer = 1
   Command10.Enabled = True
   '\\\OT-07-0092-RG-FIN///
   '
End Sub

Sub Command11_Click()
   'jandy sql
   Dim PEDENCTEMP As ADODB.Recordset
   
   Command11.Enabled = False
   Call ABREPEDCLI      ' SE ABRE LA CONEXION PARA EL CASO DEL ACCESO DIRECTO DESDE EL MENÚ
   SQLX$ = "SELECT NroPed, Status, Referencia FROM PEDENCA "
   If PROLLAMA$ = "ABIE" Then
        SQLX$ = SQLX$ + "WHERE TipoPed = 'ABIERTO' "
      Else
        SQLX$ = SQLX$ + "WHERE (TipoPed <> 'ABIERTO' "
        'sqlx$ = sqlx$ + "OR IsNull(TipoPed) = True "
        SQLX$ = SQLX$ + "OR TipoPed is null)"
   End If
   SQLX$ = SQLX$ + " ORDER BY NroPed ASC"
   Set PEDENCTEMP = New ADODB.Recordset
   PEDENCTEMP.CursorLocation = adUseClient
25 On Error Resume Next
   PEDENCTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   JJ& = 0
   Call COPYSTRU("INDIPEDI", "IPEDCOP")
   With PEDENCTEMP
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        FIN& = .RecordCount
        .MoveFirst
        JJ& = 0
        Screen.MousePointer = 11
        Do While Not .EOF
           NRODES& = !NROPED
           ANULX$ = "": If !Status = 9 Then ANULX$ = "(Anul)"
           TTXX$ = MKS$(NRODES&)
           TTXX$ = TTXX$ + ANULX$ + !REFERENCIA
           JJ& = JJ& + 1
           Call GRAREG("!IPEDCOP", TTXX$, JJ&)
           Call TRACE(20, JJ&, FIN&)
        .MoveNext
        Loop
      End If
   End With
   PEDENCTEMP.Close
   Set PEDENCTEMP = Nothing
   Screen.MousePointer = 1
   '
   Call SETULTREG("!IPEDCOP", JJ&)
   Call CIERRARCH("!IPEDCOP")
   Call FRESHECHO("!IPEDCOP")
   Screen.MousePointer = 1
   '
   Dim FORMORIG As Form
   Set FORMORIG = FORPED07
   If PROLLAMA = "ABIE" Then Set FORMORIG = PEDABIE07
   '
   Call FRESHECHO("!IPEDCOP")
30 Call SELECHO("!IPEDCOP/INDICE GENERAL DE PEDIDOS")
   NPSEL& = XVALO(VALACT1$("!IPEDCOP"))
   
'ver ernesto no seria un join?
'select * from pedenca inner join peddeta on
'PedEnca.nroped = PedDeta.nroped
'order by pedenca.nroped,peddeta.NuorItem


   If NPSEL& > 0 Then
40    CAMBIOCLIENTE% = 0
      OPX% = COMALTER%("Opciones de Trabajo\\Mismo Cliente\Seleccionar Otro Cliente")
      If OPX% < 1 Then GoTo 30
      If OPX% = 1 Then
        NCLIE = cliente_SEGUN_PEDIDO(NPSEL&)
        CAMBIOCLIENTE% = 0
      Else
        'OPCION DE SELECCIONAR OTRO CLIENTE
        Call SELECHO("DEUDOR1")
        NCLIE = XVALO(TRIM$(VALACT1$("DEUDOR1")))
        If NCLIE < 1 Then GoTo 40
        CAMBIOCLIENTE% = 1
      End If
      '
      SQLX$ = "SELECT * FROM PEDENCA "
      SQLX$ = SQLX$ + "WHERE NroPed = " & NPSEL& & " "
      
      'Set PEDENCTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
      '
     
      Set PEDENCTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
      
      With FORMORIG
        Call .BLANFORMU
        .Text1 = NCLIE
        Call .CARDACLI
        LISTAPRENCNEW$ = ""
        LISTAPRECIO$ = TRIM$(Str$(PEDENCTEMP!ListPre))
        If CAMBIOCLIENTE% > 0 Then
            LISTAPRENCNEW$ = TRIM$(.Text21(0).TEXT) ' TRAE LA LISTA DE PRECIOS DEL CLIENTE
        Else
            'NCLIE = PEDENCTEMP!NROCLIE
            '
            .Text21(0) = TRIM$(Str$(PEDENCTEMP!ListPre))
            .Text21(1) = TRIM$(Str$(PEDENCTEMP!CondPag))
            .Text21(2) = TRIM$(Str$(PEDENCTEMP!VENDED))
            .LENTREGA(0) = TRIM$(Str$(PEDENCTEMP!DOMIENT))
            .LENTREGA(1) = TRIM$(Str$(PEDENCTEMP!LOCAENT))
            .LENTREGA(2) = TRIM$(Str$(PEDENCTEMP!TranspEnt))
            .Text10 = PEDENCTEMP!NROOCOM
            .Text13 = PEDENCTEMP!Observa
        End If
        '
        OPX1% = 0
        'SI CAMBIO DE CLIENTE CHEQUEO SI HAY LISTAS DE PRECIOS DIFERENTES
         If CAMBIOCLIENTE% > 0 And XVALO(LISTAPRENCNEW$) > 0 Then
            If LISTAPRECIO$ <> LISTAPRENCNEW$ Then
               OPX1% = COMALTER%("Lista de Precios de Cliente Seleccionado\Difiere de la Lista del Pedido Original\\Mantener Lista " & LISTAPRECIO$ & " Pedido Original\Actualizar a Lista " & LISTAPRENCNEW$ & " De Cliente\No Actualizar (Mantener Precios De Pedido Original)  ")
               If OPX1% < 1 Then GoTo 40  'VUELVE HACIA ATRAS
               If OPX1% = 2 Then LISTAPRECIO$ = LISTAPRENCNEW$ 'UTILIZO LA LISTA DEL CLIENTE
               If OPX1% = 3 Then LISTAPRECIO$ = "": 'PONE LISTA 0 PARA QUE TRAIGA LOS PRECIOS DEL PEDIDO ORIGINAL
            End If
         End If
         'SI NO CAMBIO DE CLIENTE ENTRA EN OPX1% < 2
         'SI ELIGIO QUE MANTIENE LA LISTA DE PRECIOS ORIGINAL QUE CARGO EL PEDIDO PREGUNTA SI ACTUALIZA O NO
         If OPX1% < 2 Then 'SI ELIGIO QUE MANTIENE LA LISTA DE PRECIOS ORIGINAL QUE CARGO EL PEDIDO PREGUNTA SI ACTUALIZA O NO
            OPX1% = COMALTER%("Opciones de Actualización\\Actualizar Según Lista: " & LISTAPRECIO$ & "\Mantener Precios De Pedido Original  ")
            If OPX1% < 1 Then GoTo 40
            If OPX1% = 2 Then LISTAPRECIO$ = ""
         End If
        '\\\OT-05-0812-WAR-17/11/05///
        .PORDESCU = TRIM$(PEDENCTEMP!TDescApli)
        If Abs(XVALO(.PORDESCU)) > 0 Then
           Call COMUNI("Descuento General Según Pedido Original")
        End If
        '\\\OT-05-0812-WAR-FIN///
        '
        
        '
        JJ& = 0
        REBLA$ = REGBLAN$("CARDETPE")
        
        SQLX$ = "SELECT * FROM PEDDETA "
        SQLX$ = SQLX$ + "WHERE NroPed = " & NPSEL& & " "
        SQLX$ = SQLX$ + "ORDER BY NuorItem ASC "
        
        Set PEDDETTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
        'Set PEDDETTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
        '
        On Error Resume Next
        PEDDETTEMP.MoveFirst
        If Err < 1 Then
          '\\\OT-05-0812-WAR-17/11/05///
          MONEMI% = PEDDETTEMP!MoneEmis
          '\\\OT-05-0812-WAR-FIN///
          Do While Not PEDDETTEMP.EOF
           'For U& = 1 To ULTREG&("!CARDETPE")
            X$ = REBLA$
              
              CIXI% = PEDDETTEMP!CODIINT
            Call REPLA(X$, MKI$(CIXI%), 1, 2)
            Call REPLA(X$, PEDDETTEMP!TCOLOR, 3, 16)
            Call REPLA(X$, PEDDETTEMP!TTALLE, 19, 8)
              TLAR = PEDDETTEMP!TLargo
            Call REPLA(X$, MKS$(TLAR), 27, 4)
              TCHO = PEDDETTEMP!TAncho
            Call REPLA(X$, MKS$(TCHO), 31, 4)
              TESP = PEDDETTEMP!TEspesor
            Call REPLA(X$, MKS$(TESP), 35, 4)
             If XVALO(LISTAPRECIO$) > 0 Then    'Si no tiene lista de precio entonces toma la lista original (1)
                PREL = PRELISTA(LISTAPRECIO$, CIXI%)
               Else
                PREL = PEDDETTEMP!PreUnid
             End If
            Call REPLA(X$, MKS$(PREL), 39, 4)
              CAN = PEDDETTEMP!CanPed
            Call REPLA(X$, MKS$(CAN), 47, 4)
            Call REPLA(X$, MKI$(HOII%), 63, 2) ' FECHA SOLICITADA DE ENTREGA
              SUCOD$ = ""
              If IsNull(PEDDETTEMP!SubCod) = False Then SUCOD$ = PEDDETTEMP!SubCod
            Call REPLA(X$, SUCOD$, 109, 8)
            JJ& = JJ& + 1
            Call GRAREG("!CARDETPE", X$, JJ&)
          PEDDETTEMP.MoveNext
          Loop
        End If
        
        Call SETULTREG("!CARDETPE", JJ&)
        Call CIERRARCH("CARDETPE")
        On Error GoTo 0
        '
        On Error Resume Next
        Call .CARLISPE
        .MARCONOTA.Visible = False
        .MARSELPRESU.Visible = False
        Call .CALTOPED
        Call .CALPERIB
      End With
      
      Hide
    End If
    Command11.Enabled = True
    PEDDETTEMP.Close
    Set PEDDETTEMP = Nothing
End Sub

Sub Command12_Click()
    Command12.Enabled = False
    
    Dim TOTVENTA#(999999)
    Dim TOTIVA#(999999)
    '
    SELVENDE.BackColor = &HFFFFC0
    SELVENDE.Frame6.BackColor = &HFFFFC0
10  SELVENDE.Show 1
    If SELVENDE.Label2 <> "OK" Then GoTo 99
    '
    opt% = COMALTER%("Opciones de Listado\\Estadística de Facturación\Estadística de Pedidos")
    If opt% = 0 Then GoTo 99
    '
    Screen.MousePointer = 11
    CANSELE% = XVALO(SELVENDE.Label5)
    Des% = FECHANUM(SELVENDE.Label7)
    Has% = FECHANUM(SELVENDE.Label3)
    FEX = Des%: PERIO$ = FECHATEX$(FEX)
    FEX = Has%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    DESDAT = FETEXCOR1$(Des%)
    HASDAT = FETEXCOR1$(Has%)
    '
    If CANSELE% < 1 Then
       VENDSEL% = 0
       VENDDSC$ = "Todos Juntos"
       GoTo 15
    End If
    '
    For IVENDI% = 1 To SELVENDE.LICOPEN.ListCount
       If SELVENDE.LICOPEN.Selected(IVENDI% - 1) = True Then
          VENDSEL% = XVALO(Left$(SELVENDE.LICOPEN.List(IVENDI% - 1), 3))
          VENDDSC$ = TRIM$(Mid$(SELVENDE.LICOPEN.List(IVENDI% - 1), 11))
          '
15        J& = 0: TOTAGEN# = 0
          For U& = 0 To 999999
            TOTVENTA#(U&) = 0
            TOTIVA#(U&) = 0
          Next U&
          '
          NCMAXI = 0
          If opt% = 1 Then
               Call ABRECABAN
               SQLX$ = "SELECT * FROM CAJABANC "
               SQLX$ = SQLX$ + "where FechEmisi >= '" & DESDAT & "' "
               SQLX$ = SQLX$ + "AND FechEmisi <= '" & HASDAT & "' "
               'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
               Dim CABATEMP As ADODB.Recordset
               Set CABATEMP = New ADODB.Recordset
25             On Error Resume Next
               CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
               If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                  On Error GoTo 0
                  Call CapturaErrorOpen
                  GoTo 25
               End If
               On Error GoTo 0

               With CABATEMP
                  On Error Resume Next
                  .MoveLast
                  If Err > 0 Then
                    On Error GoTo 0
                    Call MENSERR(24, "Sin Movimientos en el Período Elegido")
                    GoTo 99
                  End If
                  On Error GoTo 0
                  FIN& = .RecordCount
                  .MoveFirst
                  '
                  U& = 0
                  Do While Not .EOF
                    U& = U& + 1
                    Call TRACE(20, U&, FIN&)
                    '
                    TIMO$ = TRIM$(UCase$(!TipoMovim))
                    NCLIE = !nuclien
                    If VENDSEL% = 0 Or VENDSEL% = VENDCLI%(NCLIE) Then
                       If NCLIE < 0 Then NCLIE = 0
                       If NCLIE > NCMAXI Then NCMAXI = NCLIE
                       ITOTA# = !IHABECOMP - !Idebecomp
                       If TIMO$ = "VENTA" Then
                            TOTVENTA#(NCLIE) = TOTVENTA#(NCLIE) + ITOTA#
                          ElseIf TIMO$ = "IVA" Then
                            TOTIVA#(NCLIE) = TOTIVA#(NCLIE) + ITOTA#
                       End If
                    End If
19                .MoveNext
                  Loop
               End With
               CABATEMP.Close
               Set CABATEMP = Nothing
               '
            ElseIf opt% = 2 Then
               Call ABREPEDCLI
               SQLX$ = "SELECT * FROM PEDDETA "
               SQLX$ = SQLX$ + "WHERE FechEmis >= '" & DESDAT & "' "
               SQLX$ = SQLX$ + "AND FechEmis <= '" & HASDAT & "' "
               '
               On Error GoTo 0
               'Set PEDDETTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
               Dim PEDDETTEMP As ADODB.Recordset
               Set PEDDETTEMP = New ADODB.Recordset
26             On Error Resume Next
               PEDDETTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
               If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                  On Error GoTo 0
                  Call CapturaErrorOpen
                  GoTo 26
               End If
               On Error GoTo 0

               With PEDDETTEMP
                  On Error Resume Next
                  .MoveLast
                  If Err > 0 Then
                     On Error GoTo 0
                     Call MENSERR(24, "Sin Movimientos en el Período Elegido")
                     GoTo 99
                  End If
                  On Error GoTo 0
                  FIN& = .RecordCount
                  .MoveFirst
                  U& = 0
                  Do While Not .EOF
                     U& = U& + 1
                     Call TRACE(20, U&, FIN&)
                     NCLIE = !NROCLIE
                     NART% = !CODIINT
                     If VENDSEL% = 0 Or VENDSEL% = VENDCLI%(NCLIE) Then
                       If NART% < 0 Then NART% = 0
                       If NART% > NCMAXI Then NCMAXI = NART%
                       NETO# = !CanPed * !PreUnid
                       If !MoneEmis > 1 Then
                          NETO# = NETO# * TICAMONE(!MoneEmis)
                       End If
                       NETO# = ROUND(NETO#, 2)
                       TOTVENTA#(NART%) = TOTVENTA#(NART%) + NETO#
                       ' AQUI VER COMO CALCULAR EL iva
                       NIVA# = (NETO# * PORCEIVA(NART%)) / 100
                       TOTIVA#(NART%) = TOTIVA#(NART%) + NIVA#
                     End If
                  .MoveNext
                  Loop
               End With
               PEDDETTEMP.Close
               Set PEDDETTEMP = Nothing
          End If
          '
          For U& = 1 To NCMAXI
             NC% = U&
             If Abs(TOTVENTA#(U&)) >= 0.005 Then
               W$ = REGBLAN$("RANMABRU")
               '
               J& = J& + 1: JI% = J&
               Call REPLA(W$, MKI$(NC%), 1, 2)
               Call REPLA(W$, MKD$(TOTVENTA#(NC%)), 13, 8)
               Call REPLA(W$, MKD$(TOTIVA#(NC%)), 53, 8)
                 ITOVE# = TOTVENTA#(NC%) + TOTIVA#(NC%)
               Call REPLA(W$, MKD$(ITOVE#), 61, 8)
               Call GRAREG("RANMABRU", W$, J&)
             End If
          Next U&
          '
          Call SETULTREG("RANMABRU", J&)
          Call CIERRARCH("RANMABRU")
          If CANSELE% > 1 Then
             If J& < 1 Then
                GoTo 90
             End If
          End If
          '
          Call HEADERS("RANMABRU", "")
          Call HEADERS("RANMABRU", "Periodo: " + PERIO$)
          '\\\OT-07-0010-FG-12/01/07///
          TEXVEN$ = "0 - Todos los Vendedores"
          If VENDSEL% <> 0 Then TEXVEN$ = Str(VENDSEL%) & " - " & VENDDSC$
          Call HEADERS("RANMABRU", "Vendedor: " + TEXVEN$)
          '\\\OT-07-0010-FG-FIN///
          '
          FORMAEM$ = ""
          If CANSELE% > 1 Then FORMAEM$ = "/LPT1"
          '
          If DERACCE%("VERCOSTO", "Visualizar Costos y Precios") > 1 Then
             If opt% = 1 Then
                  Call FINAL("*STATVENT" + FORMAEM$)
                ElseIf opt% = 2 Then
                  Call FINAL("*STATPEDI" + FORMAEM$)
             End If
          End If
          '
90     End If
    Next IVENDI%
    '
    Screen.MousePointer = 1
    If CANSELE% > 0 Then GoTo 10
    '
99  Screen.MousePointer = 1
    Call CIERRARCH("*.*")
    Command12.Enabled = True
    '
End Sub

Sub Command13_Click()
   Command13.Enabled = False
   If FORCAPED% < 1 Or FORCAPED% > 2 Then
      FORCAPED% = 1
      If Control$("ADMIPED", "FORMACAR") = "ITEMIZADO" Then FORCAPED% = 2
   End If
   '
   If FORCAPED% = 2 Then
      Call OPNOIN("Revertir Anulación de Pedido Itemizado")
      Exit Sub
   End If
   '
   If DERACCE%("ANULPEDI", "Anular Pedidos de Clientes") < 2 Then
      Exit Sub
   End If
   '
   Call COPYSTRU("PEDENCA", "PEDICLIE")
   Call BLOQARCH("PEDICLIE")
   Call DESANUPED
   Call CIERRARCH("*.*")
   Call LIBARCH("PEDICLIE")
   Command13.Enabled = True
End Sub







     
                       
'End Sub

Private Sub Command2_Click()
       Command2.Enabled = False
       Call COMUNI("Se Listarán Los Pedidos Que han Sido Remitidos\ y que su Remito ha sido Anulado Modo 'Anular con Reserva'\El Pedido Que Seleccione Se Liberará para Volver a Ser Remitido")
       Screen.MousePointer = 11
       Call ABREPEDCLI
       
10     Encabezado$ = "Pedido/I10;Código/A18;Cant/F8;Remito/a18;CI/a0;Cta/F7;R.Social/A24"
       Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Encabezado$, FRMZELECHO)
       Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Encabezado$, FRMZELECHO)
       
       SQLX$ = "SELECT * FROM  PEDDETA P WITH(NOLOCK) "
       SQLX$ = SQLX$ + " INNER JOIN MOVISTO M ON P.NROREMI = M.DOPRILAR "
       SQLX$ = SQLX$ + " WHERE P.Status < 8 "   '
       SQLX$ = SQLX$ + " AND P.CantEnt >=  P.CanPed"
       SQLX$ = SQLX$ + " AND P.CODIEMPR = " & CodiEmp% & " "
       SQLX$ = SQLX$ + " AND M.CANTSALID < 1"
       SQLX$ = SQLX$ + " ORDER BY P.NroPed DESC "
       Set RS = READSET(SQLX$)
       With RS
         Do While Not .EOF
            JJ& = JJ& + 1
            FRMZELECHO.MSF2.Rows = JJ& + 1
            FRMZELECHO.MSF2.TextMatrix(JJ&, 1) = XVALO(!NROPED)
            FRMZELECHO.MSF2.TextMatrix(JJ&, 2) = SAFETEXT$(!Cod_Exte)
            FRMZELECHO.MSF2.TextMatrix(JJ&, 3) = XVALO(!CanPed)
            FRMZELECHO.MSF2.TextMatrix(JJ&, 4) = SAFETEXT$(!DOPRILAR)
            FRMZELECHO.MSF2.TextMatrix(JJ&, 5) = XVALO(!CODIINT)
            FRMZELECHO.MSF2.TextMatrix(JJ&, 6) = SAFETEXT$(!NROCLIE)
            FRMZELECHO.MSF2.TextMatrix(JJ&, 7) = SAFETEXT$(!RasoClie)
            .MoveNext
         Loop
       End With
       RS.Close
       Set RS = Nothing
       
       JJ& = 0

     '
     Screen.MousePointer = 1
     FRMZELECHO.Caption = ""
     FRMZELECHO.Width = 14000
     FRMZELECHO.Show 1
    
     NUPEDIDO1& = XVALO(RESP_ZELE_VECT(1))
     Remito$ = RESP_ZELE_VECT(4)
     NCLI = XVALO(XVALO(RESP_ZELE_VECT(6)))
     If NUPEDIDO1& < 1 Then GoTo 99
     
    SQLX$ = "UPDATE PEDDETA "
    SQLX$ = SQLX$ + "SET CANTENT = 0, "
    SQLX$ = SQLX$ + " CANSALDO = 0, "
    SQLX$ = SQLX$ + " FEULTENT =  '" & CVDAT(0) & "', "
    SQLX$ = SQLX$ + " NROREMI = '' "
    SQLX$ = SQLX$ + " WHERE NROPED = " & NUPEDIDO1&
    SQLX$ = SQLX$ + " AND NroClie= " & NCLI
    SQLX$ = SQLX$ + " AND NROREMI = '" & Remito$ & "'"
    FLEXCONN.Execute (SQLX$)

    Call COMUNI("Transacción Realizada")
    GoTo 10
   
99   Command2.Enabled = True

End Sub


Private Sub Command3_Click()
   Command3.Enabled = False
   OPX% = ALTERNA%("Imputacion de Anticipos\Depuracion de Pedidos\Cambio Fecha de Regreso de Equipo\Otros Accesos Directos")
   If OPX% = 1 Then
     Call IMPUTAPEDI
   ElseIf OPX% = 2 Then
     Call ANUPEDMAS
   ElseIf OPX% = 3 Then
     Call CAMBIAFECHAEQUIPO
   Else
     Call ACCDIR("PEDIDO02")
   End If
   
   Command3.Enabled = True
   
End Sub

 Sub Command4_Click()
    Command4.Enabled = False
    If FORCAPED% < 1 Or FORCAPED% > 2 Then
       FORCAPED% = 1
       If Control$("ADMIPED", "FORMACAR") = "ITEMIZADO" Then FORCAPED% = 2
    End If
    '
    If FORCAPED% = 2 Then
       Call OPNOIN("Anulación de Pedido Itemizado")
       GoTo 99
    End If
    '
    If DERACCE%("ANULPEDI", "Anular Pedidos de Clientes") < 2 Then
       GoTo 99
    End If
    '
    Call COPYSTRU("PEDENCA", "PEDICLIE")
    Call BLOQARCH("PEDICLIE")
    Call ANULITPED
    Call CIERRARCH("*.*")
    Call LIBARCH("PEDICLIE")
    'Call RECAPEDPE
99  Command4.Enabled = True
End Sub

 Sub Command43_Click()
    Command43.Enabled = False
    If FORCAPED% < 1 Or FORCAPED% > 2 Then
       FORCAPED% = 1
       If Control$("ADMIPED", "FORMACAR") = "ITEMIZADO" Then FORCAPED% = 2
    End If
    '
    If FORCAPED% = 2 Then
       Call OPNOIN("Revertir Anulación de Pedido Itemizado")
       Exit Sub
    End If
    '
    If DERACCE%("ANULPEDI", "Anular Pedidos de Clientes") < 2 Then
       Exit Sub
    End If
    '
    Call COPYSTRU("PEDENCA", "PEDICLIE")
    Call BLOQARCH("PEDICLIE")
    Call DESANUPED
    Call CIERRARCH("*.*")
    Call LIBARCH("PEDICLIE")
    Command43.Enabled = True
End Sub

Sub Command45_Click()
    Command45.Enabled = False
    If FORCAPED% < 1 Or FORCAPED% > 2 Then
       FORCAPED% = 1
       If Control$("ADMIPED", "FORMACAR") = "ITEMIZADO" Then FORCAPED% = 2
    End If
    '
    If FORCAPED% = 2 Then
       Call OPNOIN("Anulación de Pedido Itemizado")
       GoTo 99
    End If
    '
    If DERACCE%("ANULPEDI", "Anular Pedidos de Clientes") < 2 Then
       GoTo 99
    End If
    '
    Call COPYSTRU("PEDENCA", "PEDICLIE")
    Call BLOQARCH("PEDICLIE")
    Call BLOQARCH("ODESEMI")
    Call ANULAPED
    
    Call CIERRARCH("*.*")
    Call LIBARCH("PEDICLIE")
    Call LIBARCH("ODESEMI")
99  Command45.Enabled = True
End Sub

 Sub Command46_Click()
    Command46.Enabled = False
    'Call MENSERR(24, "No Disponible en la Presente Versión.\Anule y Vuelva a Generar el Pedido Corregido.")
    'GoTo 99
    '
    If FORCAPED% < 1 Or FORCAPED% > 2 Then
       FORCAPED% = 1
       If Control$("ADMIPED", "FORMACAR") = "ITEMIZADO" Then FORCAPED% = 2
    End If
    '
    If FORCAPED% = 2 Then
       Call MENSERR(24, "Imposible Modificar Pedido\para Formato Itemizado.\ \Anule y Cargue un Nuevo Pedido.")
       GoTo 99
    End If
    '
    If DERACCE%("MODIPEDI", "Modificar Pedidos de Clientes") < 2 Then
       Exit Sub
    End If
    '
    Call COPYSTRU("PEDENCA", "PEDICLIE")
    Call BLOQARCH("PEDICLIE")
    Call CAMBIAPED
    Call CIERRARCH("*.*")
    Call LIBARCH("PEDICLIE")
99  Command46.Enabled = True
End Sub
Function MODIIFCAR_OF_DE_PEDIDO%(NPED&, MODO$)
    MODIIFCAR_OF_DE_PEDIDO% = 1
    'SE REALIZO PARA ENVAPLAST PARA VALIDAR SI EL PEDIDO TIENE CANTIDAD ENTREGADA O LA OF VINCULADA TIENE INFORMES DE PROD INFORMADOS
    SQLX$ = "SELECT * FROM PEDDETA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE NROPED = " & NPED&
    Set RS = READSET(SQLX$)
    With RS
      Do While Not .EOF
        CI1% = XVALO(!CODIINT)
        CANTI = XVALO(!CanPed)
        CONDI$ = "NROPED = " & NPED& & " AND COD_INTE = " & CI1%
        If CantRegistros&("ORDEFABR", CONDI$) > 0 Then 'SI EL PEDIDO TIENE ORDEN DE FABRICACION
           If XVALO(VALOBADA("PEDDETA", "CANTENT", "NROPED = " & NPED& & " AND CODIINT = " & CI1%, "NOMESS")) > 0 Then 'tiene cantidad entregada
              If TRIM$(UCase$(MODO$)) <> "NOMESS" Then Call COMUNI("Imposible Modificar Pedido N° " & NPED& & " \Por que Posee Cantidades Entregadas\Producto: " & CODEXT$(CI1%)): MODIIFCAR_OF_DE_PEDIDO% = -1: Exit Function
           End If
           NORFA$ = VALOBADA("ORDEFABR", "IDSUBOR", CONDI$, "NOMESS")
           If CantRegistros&("PROD1Y2", "IDSUBOR =  '" & NORFA$ & "'") > 0 Then 'TIENE OPERACIONES INFORMADAS
              If TRIM$(UCase$(MODO$)) <> "NOMESS" Then Call COMUNI("Imposible Modificar Pedido N° " & NPED& & " \Por que Posee Operaciones Informadas\Orden de Fabricación Asociada: " & NORFA$ & " Producto: " & CODEXT$(CI1%)): MODIIFCAR_OF_DE_PEDIDO% = -1: Exit Function
           End If
        End If
        .MoveNext
      Loop
    End With
     
     
     

End Function


Sub CAMBIAPED_VIE()
     '
2100 If YAINDI% = 0 Then
       '
       Screen.MousePointer = 11
       '
       EPAC$ = TRIM$(UCase$(EMPREAC$))
       '
       REFRESCARPED% = 0
       'OPCIONES DE REFRESCAR PEDIDOS PARA RECUPERAR IMAGEN DE LOS QUE ESTAN CUMPLIDOS O ANULADOS
10     OPX% = COMALTER%("Opciones:\\Modificar Pedido Pendiente\Refrescar Pedidos")
       If OPX% <= 1 Then GoTo 20
       If OPX% = 2 Then
         '
         If DERACCE%("Refrescar Pedidos", "REFRPEDI") < 2 Then
           GoTo 10
         End If
         '
         Call SELECHO("INDIPEDI/Indice General de Pedidos (Pendientes-Cumplidos-Anulados)")
         NUPEDIDO1& = XVALO(VALACT1$("INDIPEDI"))
         If NUPEDIDO1& < 1 Then GoTo 10
         REFRESCARPED% = 1
         GoTo 261
       End If
       '
20     Call ABREPEDCLI
       SQLX$ = " SELECT DISTINCT PEDENCA.NroPed, PEDENCA.NroClie, PEDENCA.FechEmis, PEDENCA.NroOcom, PEDENCA.Referencia "
       SQLX$ = SQLX$ + " FROM PEDENCA WITH(NOLOCK) "
       SQLX$ = SQLX$ + " INNER JOIN  PEDDETA "
       SQLX$ = SQLX$ + " ON (PEDENCA.NroPed = PEDDETA.NroPed) "
       SQLX$ = SQLX$ + " WHERE PEDDETA.Status < 8  "   ' para que no este anulada ni cumplida
       SQLX$ = SQLX$ + " AND PEDDETA.CantEnt < (PEDDETA.CanPed - 0.005) "
       SQLX$ = SQLX$ + " AND PEDENCA.Status < 9 "
       SQLX$ = SQLX$ + " AND PEDENCA.CODIEMPR = " & CodiEmp% & " "
       SQLX$ = SQLX$ + " AND PEDDETA.CODIEMPR = " & CodiEmp% & " "
       If CODCARPE$ <> "" Then
          SQLX$ = SQLX$ + " AND PEDENCA.CodcarpImp = '" & CODCARPE$ & "' "
       End If
       SQLX$ = SQLX$ + " ORDER BY PEDENCA.NroPed ASC "
       'jandy sql
       Dim PediModifTemp As ADODB.Recordset
       Set PediModifTemp = New ADODB.Recordset
       'PediModifTemp.CursorLocation = adUseClient
25     'On Error Resume Next
       'PediModifTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
       Set PediModifTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
       '
       JJ& = 0
       With PediModifTemp
           Do While Not .EOF
               NRODES& = !NROPED
             TTXX$ = MKS$(NRODES&)
             TTXX$ = TTXX$ + !REFERENCIA
             JJ& = JJ& + 1
             Call GRAREG("!INDIPEP", Left$(TTXX$, 52) + MKS$(REGENCA&), JJ&)
           .MoveNext
           Loop
       End With
       PediModifTemp.Close
       Set PediModifTemp = Nothing
       '
       Call SETULTREG("!INDIPEP", JJ&)
       Call CIERRARCH("!INDIPEP")
       Call FRESHECHO("!INDIPEP")
       '
       Screen.MousePointer = 1
       YAINDI% = 1
     End If
     '
2112 If ULTREG&("!INDIPEP") < 1 Then
       Call COMUNI("No hay Pedidos Abiertos\que Puedan Modificarse")
       GoTo 9999
     End If
     '
     Call LIBARCH("PEDICLIE")
     '
2110 X$ = REGSEL$("!INDIPEP/Pedidos Pendientes de Entrega")
     If Len(X$) <= 4 Then
        GoTo 9999
     End If
     '
     Call MODPED07.BLANFORMU

     NUPEDIDO1& = CVS(Left$(X$, 4))
     If ESTAPREPARADO%(NUPEDIDO1&) > 0 Then
        Call MENSERR(24, "Pedido ya Preparado - Imposible Modificar")
        GoTo 2110
     End If
     '
     ' VERIFICA SI HAY MATERIALES YA ENTREGADOS CON REMITO Y EN CASO AFIRMATIVO NO PERMITA
     ' REALIZAR LA MODIFICACION
     CANYADES = XVALO(VALOBADA("PEDDETA", "SUM(CantEnt)", " NroPed = " & NUPEDIDO1&))
     If CANYADES >= 0.05 And InStr(1, EMPREAC$, "FERVI") < 1 Then
        Call MENSERR(24, "Imposible Modificar este Pedido.\  \Ya hay Entregas o Anulaciones Parciales.")
        GoTo 9999 ' SE AGREGA PARA QUE VERIFICADA UNA ENTREGA PARCIAL NO PERMITA CONTINUAR
     End If
     '
     'VERIFICA QUE EL PEDIDO NO TENGA ASIGNADO UN ARTICULO CON NUMERO DE SERIE
     ' queda pendiente crear alguna funcion que pueda poner como no editable un campo de el prf
     PEDSER = XVALO(VALOBADA("PEDDETA", "count(NUMEROSERIE)", "(NUMEROSERIE <> '' or (NUMEROSERIE is not null and NUMEROSERIE <> ''))  and NROPED = " & NUPEDIDO1&))
     If PEDSER > 0 Then
        Call MENSERR(24, "Pedido con Asignación de Número de Serie \Solo se Podran Modificar Algunos Datos")
     End If
     '
261  Screen.MousePointer = 11
     SQLX$ = "SELECT * FROM PEDENCA "
     SQLX$ = SQLX$ + " WHERE NroPed = " & NUPEDIDO1&
     '
     Dim PEDENCATEMP As ADODB.Recordset
     Set PEDENCATEMP = New ADODB.Recordset
     PEDENCATEMP.CursorLocation = adUseClient
26   On Error Resume Next
     PEDENCATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
        GoTo 26
     End If
     On Error GoTo 0
     'jandy esto asi estaba'
     
     On Error Resume Next
     With PEDENCATEMP
       .MoveFirst
       If Err > 0 Then
         Call MENSERR(24, "Imposible Modificar Este Pedido")
         GoTo 9999
       End If
       '
       '23/03/07 (OT-07-0121)'NO PERMITE MODIFICAR UN PEDIDO A UN CLIENTE CON LA
       NC% = !NROCLIE        'CUENTA SUSPENDIDA
       LMCR# = LIMICRE(NC%)
       If LMCR# < 0 Then
          Call COMUNI("Imposible Modificar Pedido.\ \La Cuenta esta Suspendida")
          GoTo 9999
       End If
       '23/03/07 (FIN)
       '
         FEX = CVINT(!FECHEMIS)
       MODPED07.Text9.TEXT = FECHATEX$(FEX)
       MODPED07.Text11.TEXT = TRIM$(Str$(NUPEDIDO1&))
         NCLIE% = !NROCLIE
       MODPED07.LTEXT1 = TRIM$(Str$(NCLIE%))
       MODPED07.Text1.TEXT = TRIM$(Str$(NCLIE%))
       '
         NC% = NCLIE%
       MODPED07.Text2.TEXT = rasocli$(NC%)
       MODPED07.Text3.TEXT = DOMICLI$(NC%)
       MODPED07.Text4.TEXT = CPOSCLI$(NC%)
       MODPED07.Text5.TEXT = LOCACLI$(NC%)
       MODPED07.Text7.TEXT = CSTRING$(MKS$(LIMICRE(NC%)), 3, 13, 2, 2)
       MODPED07.Text12.TEXT = TRIM$(POSIVAC$(NC%)) + " - CUIT " + CuitCli$(NC%)
         CREDISPO# = LIMICRE(NC%) - SALDCLI#(NC%) - CHEPEACRE#(NC%)
       MODPED07.Text8.ForeColor = RGB(0, 0, 0)
         If CREDISPO# < 0 Then MODPED07.Text8.ForeColor = RGB(255, 0, 0)
       MODPED07.Text8.TEXT = CSTRING$(MKD$(Abs(CREDISPO#)), 3, 13, 2, 2)
       '
       MODPED07.Text10.TEXT = !NROOCOM
       MODPED07.Text21(0).TEXT = TRIM$(Str$(!ListPre))
       MODPED07.Text21(1).TEXT = TRIM$(Str$(!CondPag))
         If XVALO(MODPED07.Text21(1).TEXT) < 1 Then MODPED07.Text21(1).TEXT = Str$(CPAGCLI%(NC%))
       MODPED07.Text21(2).TEXT = TRIM$(Str$(!VENDED))
         If XVALO(MODPED07.Text21(2).TEXT) < 1 Then MODPED07.Text21(2).TEXT = Str$(VENDCLI%(NC%))
         If XVALO(MODPED07.Text21(2).TEXT) < 1 Then MODPED07.Eco21(2).TEXT = "GERENCIA"
       POCOMI = XVALO(!PorComVend)
         If Left$(UCase$(MODPED07.Label21(4)), 5) = "LETRA" Then
              MODPED07.Text44 = Chr$(64 + POCOMI)
            Else
              MODPED07.Text44 = TRIM(FORMATNUM$(POCOMI, "F4.1"))
         End If
       MODPED07.Text21(3).TEXT = TRIM$(Str$(!TIPOINGRE))
       '
       MODPED07.PORDESCU.TEXT = SAFETEXT$(!TDescApli)
       MODPED07.LENTREGA(0) = SAFETEXT$(!DOMIENT)
       MODPED07.LENTREGA(1) = SAFETEXT$(!LOCAENT)
       MODPED07.LENTREGA(2) = SAFETEXT$(!TranspEnt)
       '
       MODPED07.Text19.TEXT = SAFETEXT$(!NTranspo)
       MODPED07.Text20.TEXT = SAFETEXT$(!NSucent)
       '
       MODPED07.Text13.TEXT = SAFETEXT$(!Observa)
       MODPED07.Label11.ForeColor = RGB(0, 0, 255)
       MODPED07.Label11 = "Venta"
       '
       'MOTIVO DE PEDIDO
       MODPED07.Text16 = XVALO(!OV_Motivo)
       NUMOT% = XVALO(MODPED07.Text16)
       If EXISTE(LUDAT$ + "\TAMOTPRE.RFS") > 0 Then
          MODPED07.Label28 = ECOARCH$("TAMOTPRE", Chr$(NUMOT%))
       End If
       '
       If !VenRep = "R" Then
          MODPED07.Label11.ForeColor = RGB(255, 0, 0)
          MODPED07.Label11 = "Reparación"
          MODPED07.Label11.ToolTipText = "Doble-Click para 'Venta'"
       End If
       '
       
       If InStr(EPAC$, "GABAL") > 0 Then
          DESCURETIRA$ = TRIM$(SAFETEXT$(!Dto_Ret_Cte))
          If DESCURETIRA$ <> "" Then
             Call MODPED07.Image1_Click
             MODPED07.Check2.Value = 1
             MODPED07.Label31 = DESCURETIRA$
          End If
         DESCURETIRA$ = TRIM$(SAFETEXT$(!DtoxCant))
          If DESCURETIRA$ <> "" Then
             Call MODPED07.Image1_Click
             MODPED07.Text18 = DESCURETIRA$
             MODPED07.Labeldto.Visible = True
             MODPED07.Text17.Visible = True
             MODPED07.Text18.Visible = True
          End If
       End If
    

     End With
     PEDENCATEMP.Close
     Set PEDENCATEMP = Nothing
     
     '
     Call SETULTREG("!CARDETPE", 0)
     JJ& = 0
     REBLA$ = REGBLAN$("CARDETPE")
     
     SQLX$ = "SELECT * FROM PEDDETA "
     SQLX$ = SQLX$ + "WHERE NroPed = " & NUPEDIDO1& & " "
     '\\\OT-06-0520-RG-01/12/06///
     SQLX$ = SQLX$ + "AND Status < 8 "
     '\\\OT-06-0520-RG-FIN///
     SQLX$ = SQLX$ + "ORDER BY NuorItem ASC "
     'Set PEDDETTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     'jandy sql
     Dim PEDDETTEMP As ADODB.Recordset
     Set PEDDETTEMP = New ADODB.Recordset
     PEDDETTEMP.CursorLocation = adUseClient
27   On Error Resume Next
     PEDDETTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
        GoTo 27
     End If
     On Error GoTo 0
     'jandy esto asi estaba
     
     YAMONE% = 0
     On Error Resume Next
     PEDDETTEMP.MoveFirst
     If Err < 1 Then
       Do While Not PEDDETTEMP.EOF
         If YAMONE% = 0 Then
           YAMONE% = 1
           MONEMI% = PEDDETTEMP!MoneEmis
           MONEMA% = MONEMI%
           If MONEMI% < 1 Then MONEMI% = 1
           'If MONEMI% = 1 Then
               MODPED07.MONEMIS = DEMONECO$(MONEMI%)
               MODPED07.TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
           '  Else
           '    MONEMI% = 2
           '    MODPED07.MONEMIS = "Dólares"
           'End If
         End If
         X$ = REBLA$
           CIXI% = PEDDETTEMP!CODIINT
         Call REPLA(X$, MKI$(CIXI%), 1, 2)
         Call REPLA(X$, PEDDETTEMP!TCOLOR, 3, 16)
         Call REPLA(X$, PEDDETTEMP!TTALLE, 19, 8)
           TLAR = PEDDETTEMP!TLargo
         Call REPLA(X$, MKS$(TLAR), 27, 4)
           TCHO = PEDDETTEMP!TAncho
         Call REPLA(X$, MKS$(TCHO), 31, 4)
           TESP = PEDDETTEMP!TEspesor
         Call REPLA(X$, MKS$(TESP), 35, 4)
           'PREL = PEDDETTEMP!PreUnid * TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
           PREL = PEDDETTEMP!PreList * TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
         Call REPLA(X$, MKS$(PREL), 39, 4)
           PDESCU = 0
           If Abs(XVALO(PEDDETTEMP!PreList)) >= 0.005 Then
              PDESCU = 100 * (XVALO(PEDDETTEMP!PreList) - XVALO(PEDDETTEMP!PreUnid)) / XVALO(PEDDETTEMP!PreList)
           End If
           If Abs(PDESCU) < 0.05 Then PDESCU = 0
         Call REPLA(X$, MKS$(PDESCU), 69, 4)
           '
           ' SI NO HAY DESCUENTOS PREDEFINIDOS, LO APLICA AL PRECIO
           If CantRegistros("DESCOMBI") < 1 Then
              PREL = PREL * (100 - PDESCU) / 100
              PDESCU = 0
              Call REPLA(X$, MKS$(PREL), 39, 4)
              Call REPLA(X$, MKS$(PDESCU), 69, 4)
           End If
           '
           CAN = PEDDETTEMP!CanPed
         Call REPLA(X$, MKS$(CAN), 47, 4)
           FSE% = CVINT(PEDDETTEMP!FeSolEnt)
         Call REPLA(X$, MKI$(FSE%), 63, 2) ' FECHA SOLICITADA DE ENTREGA
           'CAENTAN = PEDDETTEMP!CantEnt
           'FULENTAN% = CVINT(PEDDETTEMP!FeUltEnt)
           IDIT$ = PEDDETTEMP!TPedItem
         Call REPLA(X$, IDIT$, 93, 16)
         Call REPLA(X$, MKS$(0), 117, 4)   ' ESTO SE ACTUALIZA MAS ABAJO
         Call REPLA(X$, MKI$(0), 121, 2)
           SUCOD$ = ""
           If IsNull(PEDDETTEMP!SubCod) = False Then SUCOD$ = PEDDETTEMP!SubCod
         Call REPLA(X$, SUCOD$, 109, 8)
         JJ& = JJ& + 1
         Call GRAREG("!CARDETPE", X$, JJ&)
       PEDDETTEMP.MoveNext
       Loop
     End If
     PEDDETTEMP.Close
     Set PEDDETTEMP = Nothing
     
     Call SETULTREG("!CARDETPE", JJ&)
     Call CIERRARCH("CARDETPE")
     On Error GoTo 0
     '
     ' REVISAR Y ACTUALIZAR CAENTAN
     Dim CANENTAN(32767), FULENT%(32767)
     CMO$ = TRIM$(Control$("", "CODREFAC"))
     If CMO$ = "" Then CMO$ = "RT"
     NMS% = NUMAS%
     '
     Call ABRESTOCKS
     SQLX$ = "SELECT * FROM MOVISTO "
     SQLX$ = SQLX$ + " WHERE CodiMovi = '" & CMO$ & "' "
     SQLX$ = SQLX$ + " AND NuPedCli = " & NUPEDIDO1& & " "
     'Set MOVITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     'jandy sql
     Dim MOVITMP As ADODB.Recordset
     Set PEDDETTEMP = New ADODB.Recordset
     PEDDETTEMP.CursorLocation = adUseClient
28   On Error Resume Next
     PEDDETTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
        GoTo 28
     End If
     On Error GoTo 0
     
     'jandy esto asi estaba
     On Error Resume Next
     With MOVITMP
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           CIXXI% = !cod_inte
           If CIXXI% > 0 Then
             If CIXXI% <= NMS% Then
               CANENTAN(CIXXI%) = CANENTAN(CIXXI%) + !Cantsalid - !CANTINGRE
               If !Cantsalid >= 0.0005 Then
                 If CVINT(!FechMov) > FULENT%(CIXXI%) Then
                   FULENT%(CIXXI%) = CVINT(!FechMov)
                 End If
               End If
             End If
           End If
         .MoveNext
         Loop
       End If
     End With
     PEDDETTEMP.Close
     Set PEDDETTEMP = Nothing
     '
     For JJ& = 1 To ULTREG&("!CARDETPE")
       X$ = REGLEIDO$("!CARDETPE", JJ&)
       CIXXI% = CVI(Left$(X$, 2))
       If CIXXI% > 0 Then
         If CIXXI% <= NMS% Then
           CANPE = CVS(Mid$(X$, 47, 4))
           CAYAEN = CVS(Mid$(X$, 117, 4))
           CASAL = CANPE - CAYAEN
           FULTIEN% = CVI(Mid$(X$, 121, 2))
           If CASAL >= CANENTAN(CIXXI%) Then
               CAYAEN = CAYAEN + CANENTAN(CIXXI%)
               CANENTAN(CIXXI%) = 0
             Else
               CAYAEN = CANPE
               CANENTAN(CIXXI%) = CANENTAN(CIXXI%) - CASAL
           End If
           If CASAL >= 0.0005 Then
             If FULENT%(CIXXI%) > FULTIEN% Then
               FULTIEN% = FULENT%(CIXXI%)
             End If
           End If
           Call REPLA(X$, MKS$(CAYAEN), 117, 4)
           Call REPLA(X$, MKI$(FULTIEN%), 121, 2)
           Call GRAREG("!CARDETPE", X$, JJ&)
         End If
       End If
     Next JJ&
     '
     For CIXI% = 1 To NUMAS%
       If CANENTAN(CIXI%) > 0.005 Then
         TTXX$ = "ATENCION: Código '" + TRIM$(CODEXT$(CIXI%)) + " - " + DESCRIT0$(CIXI%) '"
         TTXX$ = TTXX$ + "\  \Cantidad Ya Despachada Supera Pedido !!!"
         Call COMUNI(TTXX$)
       End If
     Next CIXI%
     '
     ' FIN DE RECARGA CANTIDADES YA ENTREGADAS
     '
     RNC$ = RECFILT$("NOTACLI", 1, MKI$(NC%))
     URN& = ULTREG&("NOTACLI")
     MODPED07.Text6.TEXT = ""
     '
     For U& = 1 To Len(RNC$) Step 4
       RENOTA& = CVS(Mid$(RNC$, U&, 4))
       If RENOTA& > 0 Then
         If RENOTA& <= URN& Then
           TTYY$ = RTrim$(Mid$(REGLEIDO$("NOTACLI", RENOTA&), 3))
           If MODPED07.Text6.TEXT <> "" Then
             MODPED07.Text6.TEXT = MODPED07.Text6.TEXT + Chr$(13) + Chr$(10)
           End If
           MODPED07.Text6.TEXT = MODPED07.Text6.TEXT + RTrim$(TTYY$)
         End If
       End If
     Next U&
     '
8000 MODPED07.MARCONOTA.Visible = True
     On Error Resume Next
     MODPED07.Command6.SetFocus
     On Error GoTo 0
     Screen.MousePointer = 1
     '
     NOFRESHPRE% = 1
     'ESTO SE ACTIVA PARA QUE SI DESDE UNA LISTA DE PRECIOS SE CAMBIA UN CODIGO O SE ELIGE EL MISMO
     'SE TOME EL PRECIO ACTUALIZADO DE DICHA LISTA
     If InStr(1, EMPREAC$, "FERVI") > 0 Then NOFRESHPRE% = 0
     
     MODPED07.Show 1
     NOFRESHPRE% = 0
     '
     GoTo 2110
     '
9999 Call CIERRARCH("*.*")
     Screen.MousePointer = 1
     '
End Sub
'
Sub CAMBIAPED(Optional NUPEDIDO&)
     '

     If NUPEDIDO& > 0 Then
        Call MODPED07.BLANFORMU
        GoTo 2150
     End If
     
2100 If YAINDI% = 0 Then
       '
       Screen.MousePointer = 11
       '
       EPAC$ = TRIM$(UCase$(EMPREAC$))
       '
'       REFRESCARPED% = 0
'       'OPCIONES DE REFRESCAR PEDIDOS PARA RECUPERAR IMAGEN DE LOS QUE ESTAN CUMPLIDOS O ANULADOS
'10     OPX% = COMALTER%("Opciones:\\Modificar Pedido Pendiente\Refrescar Pedidos")
'       If OPX% <= 1 Then GoTo 20
'       If OPX% = 2 Then
'         '
'         If DERACCE%("Refrescar Pedidos", "REFRPEDI") < 2 Then
'           GoTo 10
'         End If
         '
'       Call SELECHO("INDIPEDI/Indice General de Pedidos (Pendientes-Cumplidos-Anulados)")
'       NUPEDIDO& = XVALO(VALACT1$("INDIPEDI"))
'       If NUPEDIDO& < 1 Then GoTo 9999
''       End If
       '
20     Call ABREPEDCLI
       SQLX$ = " SELECT DISTINCT PEDENCA.NroPed, PEDENCA.NroClie, PEDENCA.FechEmis, PEDENCA.NroOcom, PEDENCA.Referencia "
       SQLX$ = SQLX$ + " FROM PEDENCA WITH(NOLOCK) "
       SQLX$ = SQLX$ + " INNER JOIN  PEDDETA "
       SQLX$ = SQLX$ + " ON (PEDENCA.NroPed = PEDDETA.NroPed) "
       SQLX$ = SQLX$ + " WHERE PEDDETA.Status < 8  "   ' para que no este anulada ni cumplida
       SQLX$ = SQLX$ + " AND PEDENCA.Status < 9 "
       SQLX$ = SQLX$ + " AND PEDDETA.CantEnt < (PEDDETA.CanPed - 0.005) "
       SQLX$ = SQLX$ + " AND PEDENCA.CODIEMPR = " & CodiEmp% & " "
       SQLX$ = SQLX$ + " AND PEDDETA.CODIEMPR = " & CodiEmp% & " "
       If CODCARPE$ <> "" Then
          SQLX$ = SQLX$ + " AND PEDENCA.CodcarpImp = '" & CODCARPE$ & "' "
       End If
       SQLX$ = SQLX$ + " ORDER BY PEDENCA.NroPed DESC "
       'jandy sql
       Dim PediModifTemp As ADODB.Recordset
       Set PediModifTemp = New ADODB.Recordset
       'PediModifTemp.CursorLocation = adUseClient
25     'On Error Resume Next
       'PediModifTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
       Set PediModifTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
       '
       JJ& = 0
       With PediModifTemp
           Do While Not .EOF
               NRODES& = !NROPED
             TTXX$ = MKS$(NRODES&)
             TTXX$ = TTXX$ + !REFERENCIA
             JJ& = JJ& + 1
             Call GRAREG("!INDIPEP", Left$(TTXX$, 52) + MKS$(REGENCA&), JJ&)
           .MoveNext
           Loop
       End With
       PediModifTemp.Close
       Set PediModifTemp = Nothing
       '
       Call SETULTREG("!INDIPEP", JJ&)
       Call CIERRARCH("!INDIPEP")
       Call FRESHECHO("!INDIPEP")
       '
       Screen.MousePointer = 1
       YAINDI% = 1
     End If
     '
2112 If ULTREG&("!INDIPEP") < 1 Then
       Call COMUNI("No hay Pedidos Abiertos\que Puedan Modificarse")
       GoTo 9999
     End If
     '
     Call LIBARCH("PEDICLIE")
     '
2110 X$ = REGSEL$("!INDIPEP/Pedidos Pendientes de Entrega")
     If Len(X$) <= 4 Then
        GoTo 9999
     End If
     '
     Call MODPED07.BLANFORMU

     NUPEDIDO& = CVS(Left$(X$, 4))
2150 If ESTAPREPARADO%(NUPEDIDO&) > 0 Then
        Call MENSERR(24, "Pedido ya Preparado - Imposible Modificar")
        GoTo 2110
     End If
     '
     ' VERIFICA SI HAY MATERIALES YA ENTREGADOS CON REMITO Y EN CASO AFIRMATIVO NO PERMITA
     ' REALIZAR LA MODIFICACION
     
     CANYADES = XVALO(VALOBADA("PEDDETA", "SUM(CantEnt)", " NroPed = " & NUPEDIDO&))
     If CANYADES >= 0.05 And (InStr(1, EPAC$, "FERVI") < 1 And InStr(1, EPAC$, "ASEMBLI") < 1 And InStr(1, EPAC$, "TANIT") < 1 And InStr(1, EPAC$, "CEYEQ") < 1 And InStr(1, EPAC$, "AMBEST") < 1 And InStr(1, EPAC$, "MECANIZADOS") < 1 And InStr(1, EPAC$, "EDR") < 1) Then
        Call MENSERR(24, "Imposible Modificar este Pedido.\  \Ya hay Entregas o Anulaciones Parciales.")
        GoTo 9999 ' SE AGREGA PARA QUE VERIFICADA UNA ENTREGA PARCIAL NO PERMITA CONTINUAR
     End If
     '
     'VERIFICA QUE EL PEDIDO NO TENGA ASIGNADO UN ARTICULO CON NUMERO DE SERIE
     ' queda pendiente crear alguna funcion que pueda poner como no editable un campo de el prf
     PEDSER = XVALO(VALOBADA("PEDDETA", "count(NUMEROSERIE)", "NROPED = " & NUPEDIDO& & " AND (NUMEROSERIE <> '' or (NUMEROSERIE is not null and NUMEROSERIE <> ''))"))
     If PEDSER > 0 Then
        Call MENSERR(24, "Pedido con Asignación de Número de Serie \Solo se Podran Modificar Algunos Datos")
     End If
     PEDLOT = XVALO(VALOBADA("PEDDETA", "count(IDENLOTE)", "NROPED = " & NUPEDIDO& & " AND (IDENLOTE <> '' or (IDENLOTE is not null and IDENLOTE <> ''))"))
     If PEDLOT > 0 Then
        If InStr(EPAC$, "MINEPARTS") < 1 And InStr(EPAC$, "AHP") < 1 Then
          Call MENSERR(24, "Pedido con Asignación de Lotes\Debe Desasignar Los Mismos Para Poder Modificar")
          GoTo 9999
        End If
     End If
     '
     'VALIDACION DE ENVAPLAST
     If InStr(EPAC$, "ENVAPLAST") > 0 Then
         If MODIIFCAR_OF_DE_PEDIDO%(NUPEDIDO&, "") < 0 Then
           If DERACCE%("MODIPEDI", "Modificar Pedidos de Clientes") >= 2 Then
                OPX% = COMALTER("Solos Cambios de Precios\\Cancelar\Cambiar Precios")
                If OPX% > 1 Then
                  GoTo 2200
                Else
                  Exit Sub
                End If
           Else
             Exit Sub
           End If
           
         End If
     End If

2200 Call CARGA_PEDIDO(NUPEDIDO&)
     
'261  Screen.MousePointer = 11
'     SQLX$ = "SELECT * FROM PEDENCA "
'     SQLX$ = SQLX$ + " WHERE NroPed = " & NUPEDIDO&
'     '
'     Dim PEDENCATEMP As ADODB.Recordset
'     Set PEDENCATEMP = New ADODB.Recordset
'     PEDENCATEMP.CursorLocation = adUseClient
'26   On Error Resume Next
'     PEDENCATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
'     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'        On Error GoTo 0
'        Call CapturaErrorOpen
'        GoTo 26
'     End If
'     On Error GoTo 0
'     'jandy esto asi estaba'
'
'     On Error Resume Next
'     With PEDENCATEMP
'       .MoveFirst
'       If Err > 0 Then
'         Call MENSERR(24, "Imposible Modificar Este Pedido")
'         GoTo 9999
'       End If
'       '
'       '23/03/07 (OT-07-0121)'NO PERMITE MODIFICAR UN PEDIDO A UN CLIENTE CON LA
'       NC% = !NroClie        'CUENTA SUSPENDIDA
'       LMCR# = LIMICRE(NC%)
'       If LMCR# < 0 Then
'          Call COMUNI("Imposible Modificar Pedido.\ \La Cuenta esta Suspendida")
'          GoTo 9999
'       End If
'       '23/03/07 (FIN)
'       '
'         FEX = CVINT(!FechEmis)
'       MODPED07.Text9.TEXT = FECHATEX$(FEX)
'       MODPED07.Text11.TEXT = TRIM$(Str$(NUPEDIDO&))
'         NCLIE% = !NroClie
'       MODPED07.LTEXT1 = TRIM$(Str$(NCLIE%))
'       MODPED07.Text1.TEXT = TRIM$(Str$(NCLIE%))
'       '
'         NC% = NCLIE%
'       MODPED07.Text2.TEXT = RASOCLI$(NC%)
'       MODPED07.Text3.TEXT = DOMICLI$(NC%)
'       MODPED07.Text4.TEXT = CPOSCLI$(NC%)
'       MODPED07.Text5.TEXT = LOCACLI$(NC%)
'       MODPED07.Text7.TEXT = CSTRING$(MKS$(LIMICRE(NC%)), 3, 13, 2, 2)
'       MODPED07.Text12.TEXT = TRIM$(POSIVAC$(NC%)) + " - CUIT " + CUITCLI$(NC%)
'         CREDISPO# = LIMICRE(NC%) - SALDCLI#(NC%) - CHEPEACRE#(NC%)
'       MODPED07.Text8.ForeColor = RGB(0, 0, 0)
'         If CREDISPO# < 0 Then MODPED07.Text8.ForeColor = RGB(255, 0, 0)
'       MODPED07.Text8.TEXT = CSTRING$(MKD$(Abs(CREDISPO#)), 3, 13, 2, 2)
'       '
'       MODPED07.Text10.TEXT = !NROOCOM
'       MODPED07.TEXT21(0).TEXT = TRIM$(Str$(!ListPre))
'       MODPED07.TEXT21(1).TEXT = TRIM$(Str$(!CONDPAG))
'         If XVALO(MODPED07.TEXT21(1).TEXT) < 1 Then MODPED07.TEXT21(1).TEXT = Str$(CPAGCLI%(NC%))
'       MODPED07.TEXT21(2).TEXT = TRIM$(Str$(!Vended))
'         If XVALO(MODPED07.TEXT21(2).TEXT) < 1 Then MODPED07.TEXT21(2).TEXT = Str$(VENDCLI%(NC%))
'         If XVALO(MODPED07.TEXT21(2).TEXT) < 1 Then MODPED07.Eco21(2).TEXT = "GERENCIA"
'       POCOMI = XVALO(!PorComVend)
'         If Left$(UCase$(MODPED07.Label21(4)), 5) = "LETRA" Then
'              MODPED07.Text44 = Chr$(64 + POCOMI)
'            Else
'              MODPED07.Text44 = TRIM(FORMATNUM$(POCOMI, "F4.1"))
'         End If
'       MODPED07.TEXT21(3).TEXT = TRIM$(Str$(!TIPOINGRE))
'       '
'       MODPED07.PORDESCU.TEXT = SAFETEXT$(!TDescApli)
'       MODPED07.LENTREGA(0) = SAFETEXT$(!DOMIENT)
'       MODPED07.LENTREGA(1) = SAFETEXT$(!LOCAENT)
'       MODPED07.LENTREGA(2) = SAFETEXT$(!TranspEnt)
'       '
'       MODPED07.Text19.TEXT = SAFETEXT$(!NTranspo)
'       MODPED07.Text20.TEXT = SAFETEXT$(!NSucent)
'       '
'       MODPED07.Text13.TEXT = SAFETEXT$(!Observa)
'       MODPED07.Label11.ForeColor = RGB(0, 0, 255)
'       MODPED07.Label11 = "Venta"
'       '
'       'MOTIVO DE PEDIDO
'       MODPED07.Text16 = XVALO(!OV_Motivo)
'       NUMOT% = XVALO(MODPED07.Text16)
'       If EXISTE(LUDAT$ + "\TAMOTPRE.RFS") > 0 Then
'          MODPED07.Label28 = ECOARCH$("TAMOTPRE", Chr$(NUMOT%))
'       End If
'       '
'       If !VenRep = "R" Then
'          MODPED07.Label11.ForeColor = RGB(255, 0, 0)
'          MODPED07.Label11 = "Reparación"
'          MODPED07.Label11.ToolTipText = "Doble-Click para 'Venta'"
'       End If
'       '
'
'       If InStr(EPAC$, "GABAL") > 0 Then
'          DESCURETIRA$ = TRIM$(SAFETEXT$(!Dto_Ret_Cte))
'          If DESCURETIRA$ <> "" Then
'             Call MODPED07.Image1_Click
'             MODPED07.Check2.Value = 1
'             MODPED07.Label31 = DESCURETIRA$
'          End If
'         DESCURETIRA$ = TRIM$(SAFETEXT$(!DtoxCant))
'          If DESCURETIRA$ <> "" Then
'             Call MODPED07.Image1_Click
'             MODPED07.Text18 = DESCURETIRA$
'             MODPED07.Labeldto.Visible = True
'             MODPED07.Text17.Visible = True
'             MODPED07.Text18.Visible = True
'          End If
'       End If
'
'
'     End With
'     PEDENCATEMP.Close
'     Set PEDENCATEMP = Nothing
'
'     '
'     Call SETULTREG("!CARDETPE", 0)
'     JJ& = 0
'     REBLA$ = REGBLAN$("CARDETPE")
'
'     SQLX$ = "SELECT * FROM PEDDETA "
'     SQLX$ = SQLX$ + "WHERE NroPed = " & NUPEDIDO& & " "
'     '\\\OT-06-0520-RG-01/12/06///
'     SQLX$ = SQLX$ + "AND Status < 8 "
'     '\\\OT-06-0520-RG-FIN///
'     SQLX$ = SQLX$ + "ORDER BY NuorItem ASC "
'     'Set PEDDETTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
'     '
'     'jandy sql
'     Dim PEDDETTEMP As ADODB.Recordset
'     Set PEDDETTEMP = New ADODB.Recordset
'     PEDDETTEMP.CursorLocation = adUseClient
'27   On Error Resume Next
'     PEDDETTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
'     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'        On Error GoTo 0
'        Call CapturaErrorOpen
'        GoTo 27
'     End If
'     On Error GoTo 0
'     'jandy esto asi estaba
'
'     YAMONE% = 0
'     On Error Resume Next
'     PEDDETTEMP.MoveFirst
'     If Err < 1 Then
'       Do While Not PEDDETTEMP.EOF
'         If YAMONE% = 0 Then
'           YAMONE% = 1
'           MONEMI% = PEDDETTEMP!MoneEmis
'           MONEMA% = MONEMI%
'           If MONEMI% < 1 Then MONEMI% = 1
'           'If MONEMI% = 1 Then
'               MODPED07.MONEMIS = DEMONECO$(MONEMI%)
'               MODPED07.TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
'           '  Else
'           '    MONEMI% = 2
'           '    MODPED07.MONEMIS = "Dólares"
'           'End If
'         End If
'         X$ = REBLA$
'           CIXI% = PEDDETTEMP!CodiInt
'         Call REPLA(X$, MKI$(CIXI%), 1, 2)
'         Call REPLA(X$, PEDDETTEMP!TColor, 3, 16)
'         Call REPLA(X$, PEDDETTEMP!TTalle, 19, 8)
'           TLAR = PEDDETTEMP!TLargo
'         Call REPLA(X$, MKS$(TLAR), 27, 4)
'           TCHO = PEDDETTEMP!TAncho
'         Call REPLA(X$, MKS$(TCHO), 31, 4)
'           TESP = PEDDETTEMP!TEspesor
'         Call REPLA(X$, MKS$(TESP), 35, 4)
'           'PREL = PEDDETTEMP!PreUnid * TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
'           PREL = PEDDETTEMP!PreList * TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
'         Call REPLA(X$, MKS$(PREL), 39, 4)
'           PDESCU = 0
'           If Abs(XVALO(PEDDETTEMP!PreList)) >= 0.005 Then
'              PDESCU = 100 * (XVALO(PEDDETTEMP!PreList) - XVALO(PEDDETTEMP!PreUnid)) / XVALO(PEDDETTEMP!PreList)
'           End If
'           If Abs(PDESCU) < 0.05 Then PDESCU = 0
'         Call REPLA(X$, MKS$(PDESCU), 69, 4)
'           '
'           ' SI NO HAY DESCUENTOS PREDEFINIDOS, LO APLICA AL PRECIO
'           If CantRegistros("DESCOMBI") < 1 Then
'              PREL = PREL * (100 - PDESCU) / 100
'              PDESCU = 0
'              Call REPLA(X$, MKS$(PREL), 39, 4)
'              Call REPLA(X$, MKS$(PDESCU), 69, 4)
'           End If
'           '
'           CAN = PEDDETTEMP!CanPed
'         Call REPLA(X$, MKS$(CAN), 47, 4)
'           FSE% = CVINT(PEDDETTEMP!FeSolEnt)
'         Call REPLA(X$, MKI$(FSE%), 63, 2) ' FECHA SOLICITADA DE ENTREGA
'           'CAENTAN = PEDDETTEMP!CantEnt
'           'FULENTAN% = CVINT(PEDDETTEMP!FeUltEnt)
'           IDIT$ = PEDDETTEMP!TPedItem
'         Call REPLA(X$, IDIT$, 93, 16)
'         Call REPLA(X$, MKS$(0), 117, 4)   ' ESTO SE ACTUALIZA MAS ABAJO
'         Call REPLA(X$, MKI$(0), 121, 2)
'           SUCOD$ = ""
'           If IsNull(PEDDETTEMP!SubCod) = False Then SUCOD$ = PEDDETTEMP!SubCod
'         Call REPLA(X$, SUCOD$, 109, 8)
'         JJ& = JJ& + 1
'         Call GRAREG("!CARDETPE", X$, JJ&)
'       PEDDETTEMP.MoveNext
'       Loop
'     End If
'     PEDDETTEMP.Close
'     Set PEDDETTEMP = Nothing
'
'     Call SETULTREG("!CARDETPE", JJ&)
'     Call CIERRARCH("CARDETPE")
'     On Error GoTo 0
'     '
'     ' REVISAR Y ACTUALIZAR CAENTAN
'     Dim CANENTAN(32767), FULENT%(32767)
'     CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
'     If CMO$ = "" Then CMO$ = "RT"
'     NMS% = NUMAS%
'     '
'     Call ABRESTOCKS
'     SQLX$ = "SELECT * FROM MOVISTO "
'     SQLX$ = SQLX$ + " WHERE CodiMovi = '" & CMO$ & "' "
'     SQLX$ = SQLX$ + " AND NuPedCli = " & NUPEDIDO& & " "
'     'Set MOVITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
'     '
'     'jandy sql
'     Dim MOVITMP As ADODB.Recordset
'     Set PEDDETTEMP = New ADODB.Recordset
'     PEDDETTEMP.CursorLocation = adUseClient
'28   On Error Resume Next
'     PEDDETTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
'     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'        On Error GoTo 0
'        Call CapturaErrorOpen
'        GoTo 28
'     End If
'     On Error GoTo 0
'
'     'jandy esto asi estaba
'     On Error Resume Next
'     With MOVITMP
'       .MoveFirst
'       If Err < 1 Then
'         Do While Not .EOF
'           CIXXI% = !COD_INTE
'           If CIXXI% > 0 Then
'             If CIXXI% <= NMS% Then
'               CANENTAN(CIXXI%) = CANENTAN(CIXXI%) + !CANTSALID - !CANTINGRE
'               If !CANTSALID >= 0.0005 Then
'                 If CVINT(!FECHMOV) > FULENT%(CIXXI%) Then
'                   FULENT%(CIXXI%) = CVINT(!FECHMOV)
'                 End If
'               End If
'             End If
'           End If
'         .MoveNext
'         Loop
'       End If
'     End With
'     PEDDETTEMP.Close
'     Set PEDDETTEMP = Nothing
'     '
'     For JJ& = 1 To ULTREG&("!CARDETPE")
'       X$ = REGLEIDO$("!CARDETPE", JJ&)
'       CIXXI% = CVI(Left$(X$, 2))
'       If CIXXI% > 0 Then
'         If CIXXI% <= NMS% Then
'           CANPE = CVS(Mid$(X$, 47, 4))
'           CAYAEN = CVS(Mid$(X$, 117, 4))
'           CASAL = CANPE - CAYAEN
'           FULTIEN% = CVI(Mid$(X$, 121, 2))
'           If CASAL >= CANENTAN(CIXXI%) Then
'               CAYAEN = CAYAEN + CANENTAN(CIXXI%)
'               CANENTAN(CIXXI%) = 0
'             Else
'               CAYAEN = CANPE
'               CANENTAN(CIXXI%) = CANENTAN(CIXXI%) - CASAL
'           End If
'           If CASAL >= 0.0005 Then
'             If FULENT%(CIXXI%) > FULTIEN% Then
'               FULTIEN% = FULENT%(CIXXI%)
'             End If
'           End If
'           Call REPLA(X$, MKS$(CAYAEN), 117, 4)
'           Call REPLA(X$, MKI$(FULTIEN%), 121, 2)
'           Call GRAREG("!CARDETPE", X$, JJ&)
'         End If
'       End If
'     Next JJ&
'     '
'     For CIXI% = 1 To NUMAS%
'       If CANENTAN(CIXI%) > 0.005 Then
'         TTXX$ = "ATENCION: Código '" + TRIM$(CODEXT$(CIXI%)) + " - " + DESCRIT0$(CIXI%) '"
'         TTXX$ = TTXX$ + "\  \Cantidad Ya Despachada Supera Pedido !!!"
'         Call COMUNI(TTXX$)
'       End If
'     Next CIXI%
'     '
'     ' FIN DE RECARGA CANTIDADES YA ENTREGADAS
'     '
'     RNC$ = RECFILT$("NOTACLI", 1, MKI$(NC%))
'     URN& = ULTREG&("NOTACLI")
'     MODPED07.Text6.TEXT = ""
'     '
'     For U& = 1 To Len(RNC$) Step 4
'       RENOTA& = CVS(Mid$(RNC$, U&, 4))
'       If RENOTA& > 0 Then
'         If RENOTA& <= URN& Then
'           TTYY$ = RTrim$(Mid$(REGLEIDO$("NOTACLI", RENOTA&), 3))
'           If MODPED07.Text6.TEXT <> "" Then
'             MODPED07.Text6.TEXT = MODPED07.Text6.TEXT + Chr$(13) + Chr$(10)
'           End If
'           MODPED07.Text6.TEXT = MODPED07.Text6.TEXT + RTrim$(TTYY$)
'         End If
'       End If
'     Next U&
     '
'8000 MODPED07.MARCONOTA.Visible = True
'     On Error Resume Next
'     MODPED07.Command6.SetFocus
'     On Error GoTo 0
'     Screen.MousePointer = 1
     '
     NOFRESHPRE% = 1
     'ESTO SE ACTIVA PARA QUE SI DESDE UNA LISTA DE PRECIOS SE CAMBIA UN CODIGO O SE ELIGE EL MISMO
     'SE TOME EL PRECIO ACTUALIZADO DE DICHA LISTA
     If InStr(1, EMPREAC$, "FERVI") > 0 Then NOFRESHPRE% = 0
     
     MODPED07.Show 1
     NOFRESHPRE% = 0
     
     '
     GoTo 2110
     '
9999 Call CIERRARCH("*.*")
     Screen.MousePointer = 1
     '
End Sub
'
Sub CARGA_PEDIDO(NUPEDIDO&)
    
     EPAC$ = TRIM$(UCase$(EMPREAC$))

261  Screen.MousePointer = 11

     '
     SQLX$ = "SELECT * FROM PEDENCA "
     SQLX$ = SQLX$ + " WHERE NroPed = " & NUPEDIDO&
     '
     Dim PEDENCATEMP As ADODB.Recordset
     Set PEDENCATEMP = New ADODB.Recordset
     PEDENCATEMP.CursorLocation = adUseClient
26   On Error Resume Next
     PEDENCATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
        GoTo 26
     End If
     On Error GoTo 0
     'jandy esto asi estaba'
     
     On Error Resume Next
     With PEDENCATEMP
       .MoveFirst
       If Err > 0 Then
         Call MENSERR(24, "Imposible Modificar Este Pedido")
         Exit Sub
       End If
       '
       '23/03/07 (OT-07-0121)'NO PERMITE MODIFICAR UN PEDIDO A UN CLIENTE CON LA
       NC = !NROCLIE        'CUENTA SUSPENDIDA
       LMCR# = LIMICRE(NC)
       If LMCR# < 0 Then
          Call COMUNI("Imposible Modificar Pedido.\ \La Cuenta esta Suspendida")
          Exit Sub
       End If
       '23/03/07 (FIN)
       '
         FEX = CVINT(!FECHEMIS)
       MODPED07.Text9.TEXT = FECHATEX$(FEX)
       MODPED07.Text11.TEXT = TRIM$(Str$(NUPEDIDO&))
         NCLIE = !NROCLIE
       MODPED07.LTEXT1 = TRIM$(Str$(NCLIE))
       MODPED07.Text1.TEXT = TRIM$(Str$(NCLIE))
       '
         NCLIEN = NCLIE
       MODPED07.Text2.TEXT = rasocli$(NCLIEN)
       MODPED07.Text3.TEXT = DOMICLI$(NCLIEN)
       MODPED07.Text4.TEXT = CPOSCLI$(NCLIEN)
       MODPED07.Text5.TEXT = LOCACLI$(NCLIEN)
       MODPED07.Text7.TEXT = CSTRING$(MKS$(LIMICRE(NCLIEN)), 3, 13, 2, 2)
       MODPED07.Text12.TEXT = TRIM$(POSIVAC$(NCLIEN)) + " - CUIT " + CuitCli$(NCLIEN)
         CREDISPO# = LIMICRE(NCLIEN) - SALDCLI#(NCLIEN) - CHEPEACRE#(NCLIEN)
       MODPED07.Text8.ForeColor = RGB(0, 0, 0)
         If CREDISPO# < 0 Then MODPED07.Text8.ForeColor = RGB(255, 0, 0)
       MODPED07.Text8.TEXT = CSTRING$(MKD$(Abs(CREDISPO#)), 3, 13, 2, 2)
       '
       MODPED07.Text10.TEXT = !NROOCOM
       MODPED07.Text21(0).TEXT = TRIM$(Str$(!ListPre))
       MODPED07.Text21(1).TEXT = TRIM$(Str$(!CondPag))
         If XVALO(MODPED07.Text21(1).TEXT) < 1 Then MODPED07.Text21(1).TEXT = Str$(CPAGCLI%(NCLIEN))
       MODPED07.Text21(2).TEXT = TRIM$(Str$(!VENDED))
         If XVALO(MODPED07.Text21(2).TEXT) < 1 Then MODPED07.Text21(2).TEXT = Str$(VENDCLI%(NCLIEN))
         If XVALO(MODPED07.Text21(2).TEXT) < 1 Then MODPED07.Eco21(2).TEXT = "GERENCIA"
       POCOMI = XVALO(!PorComVend)
         If Left$(UCase$(MODPED07.Label21(4)), 5) = "LETRA" Then
              MODPED07.Text44 = Chr$(64 + POCOMI)
            Else
              MODPED07.Text44 = TRIM(FORMATNUM$(POCOMI, "F4.1"))
         End If
       MODPED07.Text21(3).TEXT = TRIM$(Str$(!TIPOINGRE))
       MODPED07.TXTFLETE = SAFETEXT$(!FLETE_TEX)       '
       MODPED07.PORDESCU.TEXT = SAFETEXT$(!TDescApli)
       MODPED07.LENTREGA(0) = SAFETEXT$(!DOMIENT)
       MODPED07.LENTREGA(1) = SAFETEXT$(!LOCAENT)
       MODPED07.LENTREGA(2) = SAFETEXT$(!TranspEnt)
       If InStr(EPAC$, "CONTEPLAST") > 0 Then
         MODPED07.Text21(4) = XVALO(!AtributoComercial)
       End If
       '
       MODPED07.Text19.TEXT = SAFETEXT$(!NTranspo)
       MODPED07.Text20.TEXT = SAFETEXT$(!NSucent)
       '
       MODPED07.Text13.TEXT = SAFETEXT$(!Observa)
       MODPED07.Label11.ForeColor = RGB(0, 0, 255)
       MODPED07.Label11 = "Venta"
       '
       'MOTIVO DE PEDIDO
       MODPED07.Text16 = XVALO(!OV_Motivo)
       NUMOT% = XVALO(MODPED07.Text16)
       If EXISTE(LUDAT$ + "\TAMOTPRE.RFS") > 0 Then
          MODPED07.Label28 = ECOARCH$("TAMOTPRE", Chr$(NUMOT%))
       End If
       '
       If !VenRep = "R" Then
          MODPED07.Label11.ForeColor = RGB(255, 0, 0)
          MODPED07.Label11 = "Reparación"
          MODPED07.Label11.ToolTipText = "Doble-Click para 'Venta'"
       End If
       '
       
       If InStr(EPAC$, "GABAL") > 0 Then
          DESCURETIRA$ = TRIM$(SAFETEXT$(!Dto_Ret_Cte))
          If DESCURETIRA$ <> "" Then
             Call MODPED07.Image1_Click
             MODPED07.Check2.Value = 1
             MODPED07.Label31 = DESCURETIRA$
          End If
         DESCURETIRA$ = TRIM$(SAFETEXT$(!DtoxCant))
          If DESCURETIRA$ <> "" Then
             Call MODPED07.Image1_Click
             MODPED07.Text18 = DESCURETIRA$
             MODPED07.Labeldto.Visible = True
             MODPED07.Text17.Visible = True
             MODPED07.Text18.Visible = True
          End If
       End If
       TEXTOFLETE$ = SAFETEXT$(!FLETE_TEX)
       MODPED07.TXTFLETE = TEXTOFLETE$
       MODPED07.Label34 = 1 + XVALO(!Revision)
     End With
     PEDENCATEMP.Close
     Set PEDENCATEMP = Nothing
     
     '
     Call SETULTREG("!CARDETPE", 0)
     JJ& = 0
     REBLA$ = REGBLAN$("CARDETPE")
     '

     
     SQLX$ = "SELECT * FROM PEDDETA "
     SQLX$ = SQLX$ + "WHERE NroPed = " & NUPEDIDO& & " "
     '\\\OT-06-0520-RG-01/12/06///
     SQLX$ = SQLX$ + "AND Status < 8 "
     '\\\OT-06-0520-RG-FIN///
     SQLX$ = SQLX$ + "ORDER BY NuorItem ASC "
     'Set PEDDETTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     'jandy sql
     Dim PEDDETTEMP As ADODB.Recordset
     Set PEDDETTEMP = New ADODB.Recordset
     PEDDETTEMP.CursorLocation = adUseClient
27   On Error Resume Next
     PEDDETTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
        GoTo 27
     End If
     On Error GoTo 0
     'jandy esto asi estaba
     
     YAMONE% = 0
     On Error Resume Next
     PEDDETTEMP.MoveFirst
     If Err < 1 Then
       Do While Not PEDDETTEMP.EOF
         If YAMONE% = 0 Then
           YAMONE% = 1
           MONEMI% = PEDDETTEMP!MoneEmis
           MONEMA% = MONEMI%
           If MONEMI% < 1 Then MONEMI% = 1
           'If MONEMI% = 1 Then
               MODPED07.MONEMIS = TRIM$(DEMONECO$(MONEMI%))
               MODPED07.TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
           '  Else
           '    MONEMI% = 2
           '    MODPED07.MONEMIS = "Dólares"
           'End If
         End If
         X$ = REBLA$
           CIXI% = PEDDETTEMP!CODIINT
         Call REPLA(X$, MKI$(CIXI%), 1, 2)
         Call REPLA(X$, PEDDETTEMP!TCOLOR, 3, 16)
         
         Call REPLA(X$, PEDDETTEMP!TTALLE, 19, 8)
           TLAR = PEDDETTEMP!TLargo
         Call REPLA(X$, MKS$(TLAR), 27, 4)
           TCHO = PEDDETTEMP!TAncho
         Call REPLA(X$, MKS$(TCHO), 31, 4)
           TESP = PEDDETTEMP!TEspesor
         Call REPLA(X$, MKS$(TESP), 35, 4)
           'PREL = PEDDETTEMP!PreUnid * TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
           VENTANACARGA$ = TRIM$(Control$("", "VENCARPE"))
         If UCase$(VENTANACARGA$) <> "CARDETLO" Then
            If PEDDETTEMP!PreList > 0.00005 Then
                PREL = PEDDETTEMP!PreList * TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
            Else
                PREL = PEDDETTEMP!PreUnid * TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
            End If
         ElseIf UCase$(VENTANACARGA$) = "CARDETLO" Then
           'VENTANA CARDETLO ()VENTANA DE CARGA CON LOTES DE ARRASTRE DE PRESUPUESTO AHP
           PREL = PEDDETTEMP!PreUnid * TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
         End If
         
         Call REPLA(X$, MKS$(PREL), 39, 4)
           PDESCU = 0
           If Abs(XVALO(PEDDETTEMP!PreList)) >= 0.005 Then
              PDESCU = 100 * (XVALO(PEDDETTEMP!PreList) - XVALO(PEDDETTEMP!PreUnid)) / XVALO(PEDDETTEMP!PreList)
           End If
           If Abs(PDESCU) < 0.05 Then PDESCU = 0
         Call REPLA(X$, MKS$(PDESCU), 69, 4)
           '
           ' SI NO HAY DESCUENTOS PREDEFINIDOS, LO APLICA AL PRECIO
           If CantRegistros("DESCOMBI") < 1 Then
              PREL = PREL * (100 - PDESCU) / 100
              PDESCU = 0
              Call REPLA(X$, MKS$(PREL), 39, 4)
              Call REPLA(X$, MKS$(PDESCU), 69, 4)
           End If
           '
           CAN = PEDDETTEMP!CanPed
         Call REPLA(X$, MKS$(CAN), 47, 4)
           FSE% = CVINT(PEDDETTEMP!FeSolEnt)
         Call REPLA(X$, MKI$(FSE%), 63, 2) ' FECHA SOLICITADA DE ENTREGA
           'CAENTAN = PEDDETTEMP!CantEnt
           'FULENTAN% = CVINT(PEDDETTEMP!FeUltEnt)
           NLISSTA = XVALO(PEDDETTEMP!lista_de_item)
           Call REPLA(X$, MKS$(NLISSTA), 85, 4)
           
           IDIT$ = PEDDETTEMP!TPedItem
         Call REPLA(X$, IDIT$, 93, 16)
         Call REPLA(X$, MKS$(0), 117, 4)   ' ESTO SE ACTUALIZA MAS ABAJO
         Call REPLA(X$, MKI$(0), 121, 2)
         
         'esto agregado por ahp
         NPROV% = XVALO(PEDDETTEMP!Num_Prov)
         Call REPLA(X$, MKI$(NPROV%), 123, 2)
         'fin agregado ahp
         
           SUCOD$ = ""
           If IsNull(PEDDETTEMP!SubCod) = False Then SUCOD$ = PEDDETTEMP!SubCod
         Call REPLA(X$, SUCOD$, 109, 8)
         JJ& = JJ& + 1
         Call GRAREG("!CARDETPE", X$, JJ&)
       PEDDETTEMP.MoveNext
       Loop
     End If
     PEDDETTEMP.Close
     Set PEDDETTEMP = Nothing
     
     Call SETULTREG("!CARDETPE", JJ&)
     Call CIERRARCH("CARDETPE")
     On Error GoTo 0
     '
     ' REVISAR Y ACTUALIZAR CAENTAN
     Dim CANENTAN(32767), FULENT%(32767)
     CMO$ = TRIM$(Control$("", "CODREFAC"))
     If CMO$ = "" Then CMO$ = "RT"
     NMS% = NUMAS%
     '
     Erase CANENTAN()
     If InStr(EPAC$, "AHP") > 0 Then
       On Error Resume Next
       CONDI$ = "NROPED = " & NUPEDIDO&
       CONDI$ = CONDI$ + " AND NROCLIE = " & NC
       EntregaDomicilio% = XVALO(VALOBADA("DATOENTREGA", "EntEnDomi", CONDI$, "NOMESS"))
       MODPED07.CheckAhpRetiraDep = EntregaDomicilio%
       On Error GoTo 0
    End If
'     Call ABRESTOCKS
'     SQLX$ = "SELECT * FROM MOVISTO with(nolock)"
'     SQLX$ = SQLX$ + " WHERE CodiMovi = '" & CMO$ & "' "
'     SQLX$ = SQLX$ + " AND NuPedCli = " & NUPEDIDO& & " "
     'Set MOVITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     'jandy sql
'     Dim MOVITMP As ADODB.Recordset
'     Set PEDDETTEMP = New ADODB.Recordset
'     PEDDETTEMP.CursorLocation = adUseClient
'28   On Error Resume Next
'     Set PEDDETTEMP = READSET(SQLX$)
'     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'        On Error GoTo 0
'        Call CapturaErrorOpen
'        GoTo 28
'     End If
'     On Error GoTo 0
     
     'jandy esto asi estaba
'     On Error Resume Next
     Call ABRESTOCKS
     SQLX$ = "SELECT * FROM MOVISTO with(nolock)"
     SQLX$ = SQLX$ + " WHERE CodiMovi = '" & CMO$ & "' "
     SQLX$ = SQLX$ + " AND NuPedCli = " & NUPEDIDO& & " "

     Set MOVITMP = READSET(SQLX$)

     With MOVITMP
         Do While Not .EOF
           CIXXI% = XVALO(!cod_inte)
           If CIXXI% > 0 Then
             If CIXXI% <= NMS% Then
               CANENTAN(CIXXI%) = CANENTAN(CIXXI%) + !Cantsalid - !CANTINGRE
               If XVALO(!Cantsalid) >= 0.0005 Then
                 If CVINT(!FechMov) > FULENT%(CIXXI%) Then
                   FULENT%(CIXXI%) = CVINT(!FechMov)
                 End If
               End If
             End If
           End If
         .MoveNext
         Loop
     End With
     MOVITMP.Close
     Set MOVITMP = Nothing
     '
     For JJ& = 1 To ULTREG&("!CARDETPE")
       X$ = REGLEIDO$("!CARDETPE", JJ&)
       CIXXI% = CVI(Left$(X$, 2))
       If CIXXI% > 0 Then
         If CIXXI% <= NMS% Then
           CANPE = CVS(Mid$(X$, 47, 4))
           CAYAEN = CVS(Mid$(X$, 117, 4))
           CASAL = CANPE - CAYAEN
           FULTIEN% = CVI(Mid$(X$, 121, 2))
           If CASAL >= CANENTAN(CIXXI%) Then
               CAYAEN = CAYAEN + CANENTAN(CIXXI%)
               CANENTAN(CIXXI%) = 0
             Else
               CAYAEN = CANPE
               CANENTAN(CIXXI%) = CANENTAN(CIXXI%) - CASAL
           End If
           If CASAL >= 0.0005 Then
             If FULENT%(CIXXI%) > FULTIEN% Then
               FULTIEN% = FULENT%(CIXXI%)
             End If
           End If
           Call REPLA(X$, MKS$(CAYAEN), 117, 4)
           Call REPLA(X$, MKI$(FULTIEN%), 121, 2)
           Call GRAREG("!CARDETPE", X$, JJ&)
         End If
       End If
     Next JJ&
     '
     For CIXI% = 1 To NUMAS%
       If CANENTAN(CIXI%) > 0.005 Then
         TTXX$ = "ATENCION: Código '" + TRIM$(CODEXT$(CIXI%)) + " - " + DESCRIT0$(CIXI%) '"
         TTXX$ = TTXX$ + "\  \Cantidad Ya Despachada Supera Pedido !!!"
         Call COMUNI(TTXX$)
       End If
     Next CIXI%
     '
     ' FIN DE RECARGA CANTIDADES YA ENTREGADAS
     '
     RNC$ = RECFILT$("NOTACLI", 1, MKS$(NC))
     URN& = ULTREG&("NOTACLI")
     MODPED07.Text6.TEXT = ""
     '
     For U& = 1 To Len(RNC$) Step 4
       RENOTA& = CVS(Mid$(RNC$, U&, 4))
       If RENOTA& > 0 Then
         If RENOTA& <= URN& Then
           TTYY$ = RTrim$(Mid$(REGLEIDO$("NOTACLI", RENOTA&), 3))
           If MODPED07.Text6.TEXT <> "" Then
             MODPED07.Text6.TEXT = MODPED07.Text6.TEXT + Chr$(13) + Chr$(10)
           End If
           MODPED07.Text6.TEXT = MODPED07.Text6.TEXT + RTrim$(TTYY$)
         End If
       End If
     Next U&
       
    MODPED07.MARCONOTA.Visible = True
    On Error Resume Next
    MODPED07.Command6.SetFocus
    On Error GoTo 0
    Screen.MousePointer = 1

End Sub
Sub ANULAPED()
     '
2000 YAINDI% = 0
     '
2100 If YAINDI% = 0 Then
        'NUEVO
       Screen.MousePointer = 11
       Call ABREPEDCLI
       Encabezado$ = "Pedido/I10;Detalle/A48"
       Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Encabezado$, FRMZELECHO)
       Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Encabezado$, FRMZELECHO)
       
       SQLX$ = " SELECT DISTINCT PEDENCA.NroPed, PEDENCA.NroClie, PEDENCA.FechEmis, PEDENCA.NroOcom "
       SQLX$ = SQLX$ + " FROM PEDENCA "
       SQLX$ = SQLX$ + " INNER JOIN  PEDDETA "
       SQLX$ = SQLX$ + " ON (PEDENCA.NroPed = PEDDETA.NroPed) "
       SQLX$ = SQLX$ + " WHERE PEDDETA.Status < 8 "   ' para que no este anulada ni cumplida
       SQLX$ = SQLX$ + " AND (PEDDETA.CantEnt < (PEDDETA.CanPed - 0.005) OR PEDDETA.CanPed < 0.05) "   ' para incluir pedidos abiertos
       SQLX$ = SQLX$ + " AND PEDENCA.Status < 9 "
       SQLX$ = SQLX$ + " AND PEDENCA.CODIEMPR = " & CodiEmp% & " "
       SQLX$ = SQLX$ + " AND PEDDETA.CODIEMPR = " & CodiEmp% & " "
       SQLX$ = SQLX$ + " ORDER BY PEDENCA.NroPed ASC "
       
       Dim PediAnulTemp As ADODB.Recordset
       Set PediAnulTemp = New ADODB.Recordset
       PediAnulTemp.CursorLocation = adUseClient
25     On Error Resume Next
       PediAnulTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
          On Error GoTo 0
          Call CapturaErrorOpen
          GoTo 25
       End If
       On Error GoTo 0
       
       JJ& = 0
       With PediAnulTemp
         On Error Resume Next
         .MoveFirst
         If Err < 1 Then
           On Error GoTo 0
           Do While Not .EOF
               NRODES& = !NROPED
             TTXX$ = MKS$(NRODES&)
               NOCO$ = !NROOCOM
             If NOCO$ <> "" Then
               TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
             End If
               FEX = CVINT(!FECHEMIS)
             TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
               NCLIEN = !NROCLIE
             TTXX$ = TTXX$ + TRIM$(rasocli$(NCLIEN)) + Space$(52)
             JJ& = JJ& + 1
             'Call GRAREG("INDIPEP", Left$(TTXX$, 52) + MKS$(REGENCA&), jj&)
             
             With FRMZELECHO
                .MSF2.Rows = JJ& + 1
                .MSF2.TextMatrix(JJ&, 1) = CStr(NRODES&)
                .MSF2.TextMatrix(JJ&, 2) = Mid$(TTXX$, 5)
             End With
             
           .MoveNext
           Loop
         End If
         On Error GoTo 0
       End With
       'Call SETULTREG("INDIPEP", jj&)
       'Call CIERRARCH("INDIPEP")
       'Call FRESHECHO("INDIPEP")
       '
       Screen.MousePointer = 1
       YAINDI% = 1
     End If
     '
'     If ULTREG&("INDIPEP") < 1 Then
'        Call COMUNI("No hay Pedidos Abiertos\que Puedan Anularse")
'        Call CIERRARCH("*.*")
'        Exit Sub
'     End If
     '
2210 FRMZELECHO.Caption = "Pedidos pendientes"
     FRMZELECHO.Width = 8000
     FRMZELECHO.Show 1
    
     NUPEDIDO1& = XVALO(RESP_ZELE_VECT(1))
     If NUPEDIDO1& < 1 Then GoTo 9999

     'X$ = REGSEL$("INDIPEP/Pedidos Pendientes para Anulación")
     'If Len(X$) <= 4 Then GoTo 9999
     '
     'NUPEDIDO1& = CVS(Left$(X$, 4))
     
     If ESTAPREPARADO%(NUPEDIDO1&) > 0 Then
        Call MENSERR(24, "Pedido ya Preparado - Imposible Anular")
        GoTo 9999
     End If
     
     ' VERIFICA PARA TER QUE EL PEDIDO NO CONTENGA ITEMS EN ESTADO ARMADO
     If InStr(1, EMPREAC$, "TER") > 0 Then
        NumeroPedido& = NUPEDIDO1&
        If CantRegistros&("TANUMSE", "NUMPEDIDO=" & CStr(NumeroPedido&) & " AND NUMEROSERIE<>''") > 0 Then
            Call MENSERR(24, "Imposible Eliminar Pedido Cuando Posee Items Armados.\ \Puede Deshacer los Pasos Hasta Eliminar el Número de Serie\y Luego Volver a Intentarlo.")
            GoTo 9999
        End If
     End If
     
     'VERIFICA QUE EL PEDIDO NO TENGA ASIGNADO UN ARTICULO CON NUMERO DE SERIE
     PEDSER = XVALO(VALOBADA("PEDDETA", "count(NUMEROSERIE)", "(NUMEROSERIE <> '' or (NUMEROSERIE is not null and NUMEROSERIE <> '')) and NROPED = " & NUPEDIDO1&))
     If PEDSER > 0 Then
        Call MENSERR(24, "Imposible Anular Pedido con Nº de Serie Asignado.\Primero Debe Desasignar Nº de Serie.")
        GoTo 9999
     End If
     '
     Call ABREORFAB
     TEXNORFA$ = TEXT_ORD_FAB_ASOCIADAS(NUPEDIDO1&)
     OPXCONORFA% = 0
     If TEXNORFA$ <> "" Then
           If Control$("ADMIPED", "ANUPEDOF") <> "SI" Then
             Call COMUNI("Imposible Anular Este Pedido\Orden/es de Fabricación Asociada/s\ (" & TEXNORFA$ & ")\Debe Primero Anular la/s Orden/es de Fabricación\ Para Luego Anular el Pedido")
             GoTo 2210
           Else
             OPXCONORFA% = COMALTER%("Atención!!!!  Este Pedido Posee\Orden/es de Fabricación Asociada/s\ (" & TEXNORFA$ & ")\Sí Anula El Pedido Van a Quedar las O.Fabricación Activas\Debería Anular Primero las O.Fabricación Asociadas y Luego Anular el Pedido\\ Cancelar\Anular Pedido Dejando Activas las O.Fabricación")
             If OPXCONORFA% < 2 Then GoTo 2210
             If OPXCONORFA% = 2 Then GoTo 50
           End If
     End If
50   If COMALTER%("Confirme la Anulación del Pedido Número " + TRIM$(Str$(NUPEDIDO1&)) + "\Correspondiente a: " + TRIM$(Mid$(X$, 5, 48)) + "\\Cancelar\Confirmar Anulación") <> 2 Then
        GoTo 2210
     End If
     '
     strSQL = " Update PedEnca set Status = 9, "
     strSQL = strSQL & "Referencia = '(anul)' + left(Referencia,42) "
     strSQL = strSQL & ", FechAnul = '" & CVDAT(HOY(HO$)) & "'"
     strSQL = strSQL & " where Status < 9"
     strSQL = strSQL & " And NroPed =" & NUPEDIDO1&
     FLEXCONN.Execute (FILTSQL$(strSQL))
     '
     strSQL = strSQL & " Update PedDeta set Status = case "
     strSQL = strSQL & " when (CantEnt = 0) then 9 "
     strSQL = strSQL & " when ( CantEnt > 0) then 8 end "
     strSQL = strSQL & " where Status <> 3 and NroPed = " & NUPEDIDO1&
      FLEXCONN.Execute (FILTSQL$(strSQL))
     '
     If InStr(EPAC$, "ARCON") > 0 Then
       Call OCOMSUCU.GENERAR_ARCHIVO_GENERO_PEDIDO(NCLIE, NROPED&, NUOC&, "PEDIDOANUL") 'GENERA ARCHIVO INFORMATIVO HACIA SUCURSAL NEUQUEN
     End If
     
     HOII% = HOY(HOS$)
     TTXX$ = Space$(96)
     Call REPLA(TTXX$, HOS$ + "  " + "Anulación Interactiva", 1, 56)
     Call REPLA(TTXX$, USERTER$, 69, 24)
     Call AGREGA_ANOTAPED(NUPEDIDO1&, TTXX$)
     TTXX$ = Space$(96)
     Call REPLA(TTXX$, "Ordenes de Fabricación Activas", 11, 56)
     Call AGREGA_ANOTAPED(NUPEDIDO1&, TTXX$)
     
     If OPXCONORFA% = 2 Then ' SI ANULO CON ORDEN DE FABRICACION ACTIVAS
        Call FRATEXTO(TEXNORFA$, 50)
        ORDFABASOC$ = RETEX$(1)
        TTXX$ = Space$(96)
        Call REPLA(TTXX$, ORDFABASOC$, 11, 56)
        Call AGREGA_ANOTAPED(NUPEDIDO1&, TTXX$)
        For KKU1% = 2 To CARETEX%
            ORDFABASOC$ = RETEX$(KKU1%)
            TTXX$ = Space$(96)
            Call REPLA(TTXX$, ORDFABASOC$, 11, 56)
            Call AGREGA_ANOTAPED(NUPEDIDO1&, TTXX$)
        Next KKU1%
     End If
     
     Call AGREGA_ANOTAPED(NUPEDIDO1&, String$(88, "="))
     '
     'Call CARGALISEL("INDIPEDI", "PEDENCA", "NROPED/F6.0 ; REFERENCIA/A60", "NROPED > 0 ORDER BY NROPED DESC")
     EPAC$ = TRIM$(UCase$(EMPREAC$))
     If InStr(EPAC$, "TANIT") < 1 Then
      Call CARGALISEL("INDIPEDI", "PEDENCA", "NROPED/F6.0 ; REFERENCIA/A60", "NROPED > 0 ORDER BY NROPED DESC")
     Else
      Call CARGALISEL("INDIPEDI", "PEDENCA", "NROPED/F6.0 ; REFERENCIA/A60", "NROPED > 199000 ORDER BY NROPED DESC") 'modificado para tanit
     End If
     

'     Call BLOQARCH("INDIPEDI")
'     TTXX$ = FILTERGETRECORD$("INDIPEDI", 1, MKS$(NUPEDIDO1&))
'     REFE$ = TRIM$(Mid$(TTXX$, 5))
'     If UCase$(TRIM$(Left$(REFE$, 6))) <> "(ANUL)" Then
'        REFE$ = "(Anul) " + REFE$
'     End If
'     TTXX$ = MKS$(NUPEDIDO1&) + REFE$
'     Call FILTERPUTRECORD("INDIPEDI", 1, MKS$(NUPEDIDO1&), TTXX$)
'     Call FILESORT("INDIPEDI", "", 1, 1, "DESC")
'     Call CIERRARCH("INDIPEDI")
     '
     Call COMUNI("Transacción Completa")
     '
9999 Call CIERRARCH("*.*")
     Screen.MousePointer = 1
'     '
End Sub
'

Sub DESANUPED()
     '
     YAINDI% = 0
     Call COPYSTRU("INDIPEP", "INDIPEQ")
     '
     
10   Screen.MousePointer = 11
     Encabezado$ = "Pedido/F10;Cliente/A64"
     Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Encabezado$, FRMZELECHO)
     Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Encabezado$, FRMZELECHO)

     Call ABREPEDCLI
     SQLX$ = " SELECT DISTINCT PEDENCA.NroPed, PEDENCA.NroClie, PEDENCA.FechEmis, PEDENCA.NroOcom,PEDENCA.Referencia ,PedEnca.RASOCLIE"
     SQLX$ = SQLX$ + " FROM PEDENCA WITH(NOLOCK)"
     SQLX$ = SQLX$ + " INNER JOIN  PEDDETA "
     SQLX$ = SQLX$ + " ON (PEDENCA.NroPed = PEDDETA.NroPed) WHERE "
     CONDI$ = " PEDENCA.Status = 9  "   ' Selecciona las Anuladas
     CONDI$ = CONDI$ + " AND PEDDETA.CantEnt < (PEDDETA.CanPed - 0.005) "
     CONDI$ = CONDI$ + " AND PEDENCA.Status = 9 "      ' Selecciona los Anulados
     CONDI$ = CONDI$ + " AND PEDENCA.CODIEMPR = " & CodiEmp% & " "
     CONDI$ = CONDI$ + " AND PEDDETA.CODIEMPR = " & CodiEmp% & " "
     CONDI$ = CONDI$ + " ORDER BY PEDENCA.NroPed ASC "
     SQLX$ = SQLX$ + CONDI$
     '
     Set RS = READSET(SQLX$)
     With RS
      If .EOF Then
        Call COMUNI("No Existen Pedidos que Revertir.")
        RS.Close
        Screen.MousePointer = 1
        Exit Sub
      End If
      JJ& = 0
      Do While Not .EOF
        With FRMZELECHO
                  JJ& = JJ& + 1
                  .MSF2.Rows = JJ& + 1
                  .MSF2.TextMatrix(JJ&, 1) = CStr(XVALO(RS!NROPED))
                  .MSF2.TextMatrix(JJ&, 2) = SAFETEXT$(RS!RasoClie)
        End With
        .MoveNext
      Loop
      End With
      RS.Close
      Set RS = Nothing
      Screen.MousePointer = 1
      FRMZELECHO.Show 1
     'Call CARGALISEL("!INDIPEQ", "PEDENCA+PEDDETA/PEDENCA.NROPED = PEDDETA.NROPED", "PEDENCA.NROPED/F8.0;PEDENCA.REFERENCIA/A48", Condi$, "DISTINCT")
'     Call CARGALISEL("!INDIPEQ", "PEDENCA+PEDDETA/PEDENCA.NROPED = PEDDETA.NROPED", "PEDENCA.NROPED/F8.0;PEDENCA.RASOCLIE/A48", CONDI$, "DISTINCT")
'     Dim PeDesAnTemp As ADODB.Recordset
'     Set PeDesAnTemp = New ADODB.Recordset
'25   On Error Resume Next
'     PeDesAnTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
'     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'        On Error GoTo 0
'        Call CapturaErrorOpen
'        GoTo 25
'     End If
'     On Error GoTo 0
'     '
'     JJ& = 0
'     Y$ = REGBLAN("INDIPEQ")
'     With PeDesAnTemp
'        On Error Resume Next
'        .MoveFirst
'        If Err < 1 Then
'           On Error GoTo 0
'           Do While Not .EOF
'               Call REPLA(Y$, MKS$(!NROPED), 1, 4)
'               Call REPLA(Y$, !REFERENCIA, 5, 48)
'               JJ& = JJ& + 1
'               Call GRAREG("INDIPEQ", Y$, JJ&)
'           .MoveNext
'           Loop
'        End If
'        On Error GoTo 0
'     End With
'     PeDesAnTemp.Close
'     Set PeDesAnTemp = Nothing
'     Call SETULTREG("INDIPEQ", JJ&)
'     Call CIERRARCH("INDIPEQ")
'     Call FRESHECHO("INDIPEQ")
'    '
'23   Screen.MousePointer = 1
'     X$ = REGSEL$("INDIPEQ/Pedidos Anulados - Revertir Anulación")
'     If Len(X$) <= 4 Then
'        Call CIERRARCH("*.*")
'        Exit Sub
'     End If
     '
'23   Call SELECHO("!INDIPEQ")
     Screen.MousePointer = 1
23   NROPE& = XVALO(RESP_ZELE_VECT(1))
     If NROPE& < 1 Then Exit Sub
     '
     '
     'VALIDACION SI TIENE ORDEN DE COMPRA INGRESADA EN OTRO PEDIDO
     If Control$("ADMIPED", "NOREPOC") = "SI" Then
        NORCOM$ = TRIM$(VALOBADA("PEDENCA", "NROOCOM", "NROPED= " & NROPE&, "NOMESS"))
        If NORCOM$ <> "" Then
          NC% = XVALO(VALOBADA("PEDENCA", "NROCLIE", "NROPED= " & NROPE&, "NOMESS"))
          If OCOMEXISTE%(NC%, NORCOM$, NROPE&) > 0 Then
           Call MENSERR(24, "Imposible Revertir Anulación de Pedido Seleccionado\Orden de Compra Asignada (" & NORCOM$ & ") \Existente en Otro Pedido\Para el Cliente " & rasocli$(NC%) & "")
           GoTo 23
          End If
        End If
     End If
     '
     'ACA CAMBIO EL ENCABEZADO
     SQLX$ = "SELECT * FROM PEDENCA"
     SQLX$ = SQLX$ + " WHERE NroPed= " & NROPE& & ""
     SQLX$ = SQLX$ + " AND Status = 9"
     Dim PEDDDENCA As ADODB.Recordset
     Set PEDDDENCA = New ADODB.Recordset
26   On Error Resume Next
     PEDDDENCA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
        GoTo 26
     End If
     On Error GoTo 0
     With PEDDDENCA
        On Error Resume Next
        .MoveFirst
        If Err < 1 Then
           On Error GoTo 0
           FECHEMII$ = FECHATEX$(CVINT(!FECHEMIS))
           CLIRASO$ = !RasoClie
           NUMOCOMP$ = !NROOCOM
           If NUMOCOMP$ = "" Then     'MARCO LA REFERENCIA sin O/c
               !REFERENCIA = Left$(CLIRASO$ & "- " & FECHEMII$, 48)
             Else                       'MARCO LA REFERENCIA con O/c
               REFNUE$ = CLIRASO$ & "- " & NUMOCOMP$ & " - " & FECHEMII$
               !REFERENCIA = Left$(REFNUE$, 48)
           End If
           !Status = 1            'MARCO EL STAT
           .Update
        End If
     End With
     PEDDDENCA.Close
     Set PEDDDENCA = Nothing
     '
     'ACA CAMBIO LOS ITEMS
     SQLX$ = "UPDATE PEDDETA"
     SQLX$ = SQLX$ + " SET Status = 1"
     SQLX$ = SQLX$ + " WHERE NroPed= " & NROPE& & ""
     SQLX$ = SQLX$ + " AND Status >= 8"
     SQLX$ = SQLX$ + "AND CanPed > CantEnt"
      FLEXCONN.Execute (FILTSQL$(SQLX$))
     '
     HOII% = HOY(HOS$)
     TTXX$ = Space$(96)
     Call REPLA(TTXX$, HOS$ + "  " + "Revierte Anulación Interactiva", 1, 56)
     Call REPLA(TTXX$, USERTER$, 69, 24)
     Call AGREGA_ANOTAPED(NROPE&, TTXX$)
     Call AGREGA_ANOTAPED(NROPE&, String$(88, "="))
     '
     Screen.MousePointer = 1
     Call COMUNI("Anulación Revertida")
     GoTo 10
     '
End Sub
'

'
Sub ANULITPED()
   Dim Pedianuittemp As ADODB.Recordset
       '
     EPAC$ = TRIM$(UCase$(EMPREAC$))

2000 YAINDI% = 0
     '
2100 If YAINDI% = 0 Then
        'nuevo
         Screen.MousePointer = 11
       Call ABREPEDCLI
       SQLX$ = " SELECT DISTINCT PEDENCA.NroPed, PEDENCA.NroClie, PEDENCA.FechEmis, PEDENCA.NroOcom "
       SQLX$ = SQLX$ + " FROM PEDENCA "
       SQLX$ = SQLX$ + " INNER JOIN  PEDDETA "
       SQLX$ = SQLX$ + " ON (PEDENCA.NroPed = PEDDETA.NroPed) "
       SQLX$ = SQLX$ + " WHERE PEDDETA.Status < 8 "   ' para que no este anulada ni cumplida
       SQLX$ = SQLX$ + " AND PEDDETA.CantEnt < (PEDDETA.CanPed - 0.005) "
       SQLX$ = SQLX$ + " AND PEDENCA.Status < 9 "
       SQLX$ = SQLX$ + " AND PEDENCA.CODIEMPR = " & CodiEmp% & " "
       SQLX$ = SQLX$ + " AND PEDDETA.CODIEMPR = " & CodiEmp% & " "
       SQLX$ = SQLX$ + " ORDER BY PEDENCA.NroPed ASC "
        'jandy sql
       Set Pedianuittemp = New ADODB.Recordset
       Pedianuittemp.CursorLocation = adUseClient
25     On Error Resume Next
       Pedianuittemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
          On Error GoTo 0
         Call CapturaErrorOpen
       GoTo 25
       End If
       On Error GoTo 0
       'Set Pedianuittemp = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
       '
       JJ& = 0
       With Pedianuittemp
         On Error Resume Next
         .MoveFirst
         If Err < 1 Then
           Do While Not .EOF
               NRODES& = !NROPED
             TTXX$ = MKS$(NRODES&)
               NOCO$ = !NROOCOM
             If NOCO$ <> "" Then
               TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
             End If
               FEX = CVINT(!FECHEMIS)
             TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
               NCLIE = !NROCLIE
             TTXX$ = TTXX$ + TRIM$(rasocli$(NCLIE)) + Space$(52)
             JJ& = JJ& + 1
             Call GRAREG("INDIPEP", Left$(TTXX$, 52) + MKS$(REGENCA&), JJ&)
           .MoveNext
           Loop
         End If
       End With
       Pedianuittemp.Close
       Set Pedianuittemp = Nothing
       
       Call SETULTREG("INDIPEP", JJ&)
       Call CIERRARCH("INDIPEP")
       Call FRESHECHO("INDIPEP")
       Screen.MousePointer = 1
       YAINDI% = 1
     End If
     '
     If ULTREG&("INDIPEP") < 1 Then
        Call COMUNI("No hay Pedidos Abiertos\que Puedan Anularse")
        Call CIERRARCH("*.*")
        Exit Sub
     End If
     '
2210 X$ = REGSEL$("INDIPEP/Pedidos Pendientes para Anulación por Items")
     If Len(X$) <= 4 Then GoTo 9999
     '
     NUPEDIDO1& = CVS(Left$(X$, 4))
     REGENCA& = CVS(Mid$(X$, 53, 4))
     '
     DPEDI$ = TRIM$(Str$(NUPEDIDO1&)) + " - " + TRIM$(Mid$(X$, 5, 48))
     UREODES& = ULTREG&("PEDDETA")
     JJKK& = 0
     '   If TIDOC% <> 1 Then GoTo 1211
     '
     CAPESEL% = 1
     TTXX$ = DPEDI$
     NUODSEL& = XVALO(Mid$(TTXX$, 15, 10))
     REPRESU& = XVALO(Mid$(TTXX$, 121, 8))
     '
1205 TINI = Timer
     '
     SELITPED.List2.Clear
     '
     'nuevo
     Call ABREPEDCLI
     SQLX$ = " SELECT  * "
     SQLX$ = SQLX$ + " FROM PEDDETA "
     SQLX$ = SQLX$ + " INNER JOIN  PEDENCA "
     SQLX$ = SQLX$ + " ON (PEDDETA.NroPed = PEDENCA.NroPed) "
     SQLX$ = SQLX$ + " WHERE PEDDETA.Status < 8 "   ' para que no este anulada ni cumplida
     SQLX$ = SQLX$ + " AND PEDDETA.CantEnt < (PEDDETA.CanPed - 0.005) "
     SQLX$ = SQLX$ + " AND PEDENCA.Status < 9 "
     SQLX$ = SQLX$ + " AND PEDDETA.Status <> 3  "
     SQLX$ = SQLX$ + " AND PEDDETA.NroPed = " & NUPEDIDO1& & " "
     SQLX$ = SQLX$ + " AND PEDDETA.CODIEMPR = " & CodiEmp% & " "
     SQLX$ = SQLX$ + " AND PEDENCA.CODIEMPR = " & CodiEmp% & " "
     SQLX$ = SQLX$ + " ORDER BY PEDENCA.NroPed ASC "
     'jandy sql
     Set Pedianuittemp = New ADODB.Recordset
     Pedianuittemp.CursorLocation = adUseClient
26   On Error Resume Next
     Pedianuittemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adCmdText
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
        GoTo 26
     End If
     On Error GoTo 0
     
     'Set Pedianuittemp = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     With Pedianuittemp
       .MoveFirst
       Do While (.EOF = False)
         CAN = !CanPed - !cantent
         If CAN >= 0.05 Then
           CI% = !CODIINT
           TTYY$ = Space$(128)
           Call REPLA(TTYY$, CODEXT$(CI%), 1, 16)
           Call REPLA(TTYY$, DESCRIT0$(CI%), 19, 38)
           '\\\OT-05-0254-WAR-04/05/05///
           Call REPLA(TTYY$, FORMATNUM$(!CanPed, "F10.1"), 57, 10)
           Call REPLA(TTYY$, FORMATNUM$(!cantent, "F10.1"), 67, 10)
           Call REPLA(TTYY$, FORMATNUM$(CAN, "F10.1"), 77, 10)
           '\\\OT-05-0254-WAR-FIN///
           Call REPLA(TTYY$, FORMATNUM$(!NuOrItem, "F8.0"), 113, 8)
           SELITPED.List2.AddItem TTYY$
           .MoveNext
         End If
       Loop
     End With
     Pedianuittemp.Close
     Set Pedianuittemp = Nothing
       
     '
1209 SELITPED.Caption = "Anulacion por Items Pedido " + DPEDI$
     SELITPED.Show 1
     DoEvents
     '
     'VALIDACION DE CIERRE DE ORDENES DE FABRICACION GENERADAS A TRAVES DE PEDIDOS.
            
     If SELITPED.List2.ListCount > 0 Then
       Screen.MousePointer = 11
       '
       'VERIFICA QUE CADA ITEM A CANCELAR NO TENGAN ASIGNADO UN ARTICULO CON NUMERO DE SERIE
       For UY& = 1 To SELITPED.List2.ListCount
         TTXX$ = SELITPED.List2.List(UY& - 1)
         If SELITPED.List2.Selected(UY& - 1) = True Then
            CoIn% = CODINT%(Left$(TTXX$, 16))
            NUOR& = XVALO(Mid$(TTXX$, 113, 8))
            CONDI$ = "NroPed = " & NUPEDIDO1& & " and CodiInt = " & CoIn% & " and NuOrItem = " & NUOR&
            NUMSERI$ = VALOBADA("PEDDETA", "NUMEROSERIE", CONDI$, "NOMESS")
            If TRIM$(NUMSERI$) <> "" Then
               Call MENSERR(24, "Imposible Cancelar Codigo " & CODEXT(CoIn%) & ".\ \ Pedido con Nº de Serie Asignado.\Primero Debe Desasignar Nº de Serie.")
               Screen.MousePointer = 1
               GoTo 1209
            End If
            
            ' VERIFICA PARA TER QUE EL ITEM NO ESTE ARMADO Y PENDIENTE DE SER DESPACHADO
            NumeroPedido& = NUPEDIDO1&
            CodigoProductoInterno% = CoIn%
            NumeroOrdenPedido% = NUOR&
            CondicionSql$ = "NUMPEDIDO=" & CStr(NumeroPedido&) & " AND COD_INTE=" & CStr(CodigoProductoInterno%) & " AND NumOrdPedido=" & CStr(NumeroOrdenPedido%) & " AND NUMEROSERIE<>'' AND FECHENTREGA='" & FETEXCOR1$(FECHANUM("01/01/80")) & "'"
            If CantRegistros("TANUMSE", CondicionSql$) > 0 Then
                Call MENSERR(24, "Imposible Eliminar '" & CODEXT$(CodigoProductoInterno%) & "', Posee Armado Parcial o Total.\ \Para Quitar el Item del Pedido Debe Deshacer los Armados.")
                Screen.MousePointer = 1
                GoTo 1209
            End If
         End If
       Next UY&
       '
       'VERIFICA QUE CADA ITEM A CANCELAR NO TENGAN ASIGNADO UN ARTICULO CON NUMERO DE SERIE
       TEXNORFA$ = ""
       For UY& = 1 To SELITPED.List2.ListCount
         TTXX$ = SELITPED.List2.List(UY& - 1)
         If SELITPED.List2.Selected(UY& - 1) = True Then
            CoIn% = CODINT%(Left$(TTXX$, 16))
            NUOR& = XVALO(Mid$(TTXX$, 113, 8))
            CONDI$ = "NroPed = " & NUPEDIDO1& & " and CodiInt = " & CoIn% & " and NuOrItem = " & NUOR&
            NroPedido = XVALO(VALOBADA("PEDDETA", "NROPED", CONDI$, "NOMESS"))
            CONDI2$ = " NroPed = " & NroPedido & " AND STATUORF < 9 AND COD_INTE = " & CoIn%
            NORFA$ = TRIM$(VALOBADA("ORDEFABR", "IDSUBOR", CONDI2$, "NOMESS"))
            If NORFA$ <> "" Then
              If InStr(EPAC$, "J Y M") Then
30                  OPX% = COMALTER%("Se ha Detectado Para el Código: " & CODEXT(CoIn%) & ".\Que el Mismo Posee Orden de Fabricación Asociada(" & NORFA$ & " )\Desea Realizar el Cierre y Consumo de la Misma ?\\No, Continuar sin Cerrar\Si, Cerrar Orden de Fabricación")
                    If OPX% > 1 Then
                       CANPE = XVALO(VALOBADA("PEDDETA", "CanPed", CONDI$, "NOMESS"))
                       CANTE = XVALO(VALOBADA("PEDDETA", "CantEnt", CONDI$, "NOMESS"))
                       SALDO = CANPE - CANTE
                       ARGU$ = TRIM$(Str(SALDO))
10                     ARGU$ = UCase$(TRIM$(InputBox$("Cantidad", "Ingrese Cantidad a Cerrar Para la O.Fabricación", ARGU$)))
                       'If XVALO(ARGU$) >= SALDO Then
                       If XVALO(ARGU$) >= 0.05 Then
                          If SALDO >= 0 Then Call CIERRORFACTU(CoIn%, NORFA$, XVALO(ARGU$), HOY(HO$))
                          '
                          ' REGISTRA LA CANTIDAD FABRICADA COMO ENTREGADA EN EL PEDIDO
                          CantidadCierrePedido# = XVALO(ARGU$)
                          If CantidadCierrePedido# > 0.05 And CantidadCierrePedido# <= SALDO Then
                            CANTE = CANTE + CantidadCierrePedido#
                            Call ACTUREGISTRO("peddeta", "cantent", CONDI$, TRIM$(FORMATNUM$(CANTE, "F14.1")), REGAF&)
                            ' PONE MARCA DE CIERRE PARCIAL PARA QUE EL SISTEMA NO CAMBIE EL ESTADO Y LO MANTENGA PENDIENTE
                            ' SALVO QUE LA CANTIDAD ENTREGADA SEA IGUAL O MAYOR A LA PEDIDO
                            FLAG$ = "CIERRPAR"
                            If CANTE >= CANPE Then FLAG$ = "CIERRTOT"
                            Call REPLA(TTXX$, FLAG$, 200, 8)
                            SELITPED.List2.List(UY& - 1) = TTXX$
                          End If
                       Else
                          GoTo 30
                       '   OPX% = COMALTER%("La Cantidad a Cerrar Debe Ser Igual o Mayor que La Cantidad Pendiente del Pedido\\Modificar Cantidad\Cancelar Transacción")
                       '   If OPX% = 1 Then ARGU$ = TRIM$(Str(SALDO)): GoTo 10               'VUELVE AL INPUTBOX
                       '   If OPX% < 1 Or OPX% > 1 Then GoTo 9999 'CANCELA O CIERRA VENTANA
                       End If
                    Else
                       ' PONE MARCA DE CIERRE TOTAL DEL ITEM, PARA QUE MAS ADELANTE CAMBIE EL STATUS A 8 Y PONGA EN 0 LA CANTIDAD ENTREGADA
                       FLAG$ = "CIERRTOT"
                       Call REPLA(TTXX$, FLAG$, 200, 8)
                       SELITPED.List2.List(UY& - 1) = TTXX$
                    End If
              Else
                    'SI TIENE OPCION DE ANULAR ITEM CON OF
                    If NORFA$ <> "" Then
                          If Control$("ADMIPED", "ANUPEDOF") <> "SI" Then
                            Call COMUNI("Imposible Anular Este Pedido\Tiene Orden de Fabricación Asociada\ (" & NORFA$ & ")\Debe Primero Anular la Ordende Fabricación\ Para Luego Anular el Pedido")
                            GoTo 2210
                          Else
                            i& = i& + 1
                            If TEXNORFA$ <> "" Then
                              If i& > 5 And i& < 6 Then ' si supera las 5 le agrega puntos suspensivos
                                TEXNORFA$ = TEXNORFA$ + " ...."
                              End If
                              TEXNORFA$ = TEXNORFA$ & " ; " & NORFA$
                            Else
                              TEXNORFA$ = NORFA$
                            End If
                          End If
                    End If
              End If
            End If
         End If
       Next UY&
       '
       If TEXNORFA$ <> "" Then
             OPXCONORFA% = COMALTER%("Atención!!!!  Este Pedido Posee\Orden/es de Fabricación Asociada/s\ (" & TEXNORFA$ & ")\Sí Anula Los Items Selecciondos Van a Quedar las O.Fabricación Activas\Debería Anular Primero las O.Fabricación Asociadas y Luego Anular Dichos Items\\ Cancelar\Anular Pedido Dejando Activas las O.Fabricación")
             If OPXCONORFA% < 2 Then GoTo 1209
       End If
       
       FLAGANUL% = 1
       For U& = 1 To SELITPED.List2.ListCount
         TTXX$ = SELITPED.List2.List(U& - 1)
         '\\\OT-05-0254-WAR-05/05/05///
         FLAG$ = Mid$(SELITPED.List2.List(U& - 1), 200, 8)
         If SELITPED.List2.Selected(U& - 1) = True And FLAG$ <> "CIERRPAR" Then 'Mid$(TTXX$, 71, 5) = "anul." Then
         '\\\OT-05-0254-WAR-FIN///
           CoIn% = CODINT%(Left$(TTXX$, 16))
           NUOR& = XVALO(Mid$(TTXX$, 113, 8))
           'jandy sql actualizo upadte
           'Busco el registro
'           With PedDeta
             abuss$ = " CodiInt = " & CoIn% & " "
             abuss$ = abuss$ + " AND NroPed = " & NUPEDIDO1& & " "
             abuss$ = abuss$ + " AND NuOrItem = " & NUOR& & " "
             abuss$ = abuss$ + " AND Status < 8 "
             abuss$ = abuss$ + " AND Status <> 3 "
             strSQL = "Update PEDDETA set Status=8"
             If FLAG$ = "CIERRTOT" Then strSQL = "Update PEDDETA set Status=8, CantEnt=0"
             strSQL = strSQL & " where  "
             strsql1 = strSQL & abuss$
             FLEXCONN.Execute (FILTSQL$(strsql1))
             FLAGANUL% = 1
         End If
       Next U&
 
       Call BLOQARCH("PEDDETP")
       Call RECAPEDPE
       Call CIERRARCH("PEDDETP")
       Screen.MousePointer = 1
       '
       If FLAGANUL% = 1 Then
            HOII% = HOY(HOS$)
            TTXX1$ = Space$(96)
            Call REPLA(TTXX1$, HOS$ + "  " + "Anulación Parcial de Items Pedidos", 1, 56)
            Call REPLA(TTXX1$, USERTER$, 69, 24)
            Call AGREGA_ANOTAPED(NUPEDIDO1&, TTXX1$)
            For U& = 1 To SELITPED.List2.ListCount
              TTXX$ = SELITPED.List2.List(U& - 1)
              If SELITPED.List2.Selected(U& - 1) = True Then
                TTXX1$ = Space$(96)
                Codigo$ = Left$(TTXX$, 16)
                Call REPLA(TTXX1$, Space(10) + Codigo$, 1, 56)
                Call AGREGA_ANOTAPED(NUPEDIDO1&, TTXX1$)
              End If
            Next U&
            If OPXCONORFA% = 2 Then ' SI ANULO CON ORDEN DE FABRICACION ACTIVAS
                TTXX$ = Space$(96)
                Call REPLA(TTXX$, "Ordenes de Fabricación Activas", 11, 56)
                Call AGREGA_ANOTAPED(NUPEDIDO1&, TTXX$)
                Call FRATEXTO(TEXNORFA$, 50)
                ORDFABASOC$ = RETEX$(1)
                TTXX$ = Space$(96)
                Call REPLA(TTXX$, ORDFABASOC$, 11, 56)
                Call AGREGA_ANOTAPED(NUPEDIDO1&, TTXX$)
                For KKU1% = 2 To CARETEX%
                    ORDFABASOC$ = RETEX$(KKU1%)
                    TTXX$ = Space$(96)
                    Call REPLA(TTXX$, ORDFABASOC$, 11, 56)
                    Call AGREGA_ANOTAPED(NUPEDIDO1&, TTXX$)
                Next KKU1%
            End If
            Call AGREGA_ANOTAPED(NUPEDIDO1&, String$(88, "="))

     End If

       '
     End If
     '
     Call CIERRARCH("*.*")
     Call BAJALDISCO
     Screen.MousePointer = 1
     DoEvents
     '
     GoTo 2000
     '
9999 Call CIERRARCH("*.*")
     Screen.MousePointer = 1
     '
End Sub


Sub ANULITPED_MODFICADA_NO_FUNCIONA_CORRECTAMENTE()
   Dim Pedianuittemp As ADODB.Recordset
       '
2000 YAINDI% = 0
     '
2100 If YAINDI% = 0 Then
        'nuevo
         Screen.MousePointer = 11
       Call ABREPEDCLI
       SQLX$ = " SELECT DISTINCT PEDENCA.NroPed, PEDENCA.NroClie, PEDENCA.FechEmis, PEDENCA.NroOcom "
       SQLX$ = SQLX$ + " FROM PEDENCA "
       SQLX$ = SQLX$ + " INNER JOIN  PEDDETA "
       SQLX$ = SQLX$ + " ON (PEDENCA.NroPed = PEDDETA.NroPed) "
       SQLX$ = SQLX$ + " WHERE PEDDETA.Status < 8 "   ' para que no este anulada ni cumplida
       SQLX$ = SQLX$ + " AND PEDDETA.CantEnt < (PEDDETA.CanPed - 0.005) "
       SQLX$ = SQLX$ + " AND PEDENCA.Status < 9 "
       SQLX$ = SQLX$ + " AND PEDENCA.CODIEMPR = " & CodiEmp% & " "
       SQLX$ = SQLX$ + " AND PEDDETA.CODIEMPR = " & CodiEmp% & " "
       SQLX$ = SQLX$ + " ORDER BY PEDENCA.NroPed ASC "
        'jandy sql
       Set Pedianuittemp = New ADODB.Recordset
       Pedianuittemp.CursorLocation = adUseClient
25     On Error Resume Next
       Pedianuittemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
          On Error GoTo 0
         Call CapturaErrorOpen
       GoTo 25
       End If
       On Error GoTo 0
       'Set Pedianuittemp = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
       '
       JJ& = 0
       With Pedianuittemp
         On Error Resume Next
         .MoveFirst
         If Err < 1 Then
           Do While Not .EOF
               NRODES& = !NROPED
             TTXX$ = MKS$(NRODES&)
               NOCO$ = !NROOCOM
             If NOCO$ <> "" Then
               TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
             End If
               FEX = CVINT(!FECHEMIS)
             TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
               NCLIE = !NROCLIE
             TTXX$ = TTXX$ + TRIM$(rasocli$(NCLIE)) + Space$(52)
             JJ& = JJ& + 1
             Call GRAREG("INDIPEP", Left$(TTXX$, 52) + MKS$(REGENCA&), JJ&)
           .MoveNext
           Loop
         End If
       End With
       Pedianuittemp.Close
       Set Pedianuittemp = Nothing
       
       Call SETULTREG("INDIPEP", JJ&)
       Call CIERRARCH("INDIPEP")
       Call FRESHECHO("INDIPEP")
       Screen.MousePointer = 1
       YAINDI% = 1
     End If
     '
     If ULTREG&("INDIPEP") < 1 Then
        Call COMUNI("No hay Pedidos Abiertos\que Puedan Anularse")
        Call CIERRARCH("*.*")
        Exit Sub
     End If
     '
2210 X$ = REGSEL$("INDIPEP/Pedidos Pendientes para Anulación por Items")
     If Len(X$) <= 4 Then GoTo 9999
     '
     NUPEDIDO1& = CVS(Left$(X$, 4))
     REGENCA& = CVS(Mid$(X$, 53, 4))
     '
     DPEDI$ = TRIM$(Str$(NUPEDIDO&)) + " - " + TRIM$(Mid$(X$, 5, 48))
     UREODES& = ULTREG&("PEDDETA")
     JJKK& = 0
     '   If TIDOC% <> 1 Then GoTo 1211
     '
     CAPESEL% = 1
     TTXX$ = DPEDI$
     NUODSEL& = XVALO(Mid$(TTXX$, 15, 10))
     REPRESU& = XVALO(Mid$(TTXX$, 121, 8))
     '
1205 TINI = Timer
     '
     SELITPED.List2.Clear
     '
     'nuevo
     Call ABREPEDCLI
     SQLX$ = " SELECT  * "
     SQLX$ = SQLX$ + " FROM PEDDETA "
     SQLX$ = SQLX$ + " INNER JOIN  PEDENCA "
     SQLX$ = SQLX$ + " ON (PEDDETA.NroPed = PEDENCA.NroPed) "
     SQLX$ = SQLX$ + " WHERE PEDDETA.Status < 8 "   ' para que no este anulada ni cumplida
     SQLX$ = SQLX$ + " AND PEDDETA.CantEnt < (PEDDETA.CanPed - 0.005) "
     SQLX$ = SQLX$ + " AND PEDENCA.Status < 9 "
     SQLX$ = SQLX$ + " AND PEDDETA.Status <> 3  "
     SQLX$ = SQLX$ + " AND PEDDETA.NroPed = " & NUPEDIDO& & " "
     SQLX$ = SQLX$ + " AND PEDDETA.CODIEMPR = " & CodiEmp% & " "
     SQLX$ = SQLX$ + " AND PEDENCA.CODIEMPR = " & CodiEmp% & " "
     SQLX$ = SQLX$ + " ORDER BY PEDENCA.NroPed ASC "
     'jandy sql
     Set Pedianuittemp = New ADODB.Recordset
     Pedianuittemp.CursorLocation = adUseClient
26   On Error Resume Next
     Pedianuittemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adCmdText
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
        GoTo 26
     End If
     On Error GoTo 0
     
     'Set Pedianuittemp = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     With Pedianuittemp
'       .MoveFirst
       Do While (.EOF = False)
         CAN = !CanPed - !cantent
         If CAN >= 0.05 Then
           CI% = !CODIINT
           TTYY$ = Space$(128)
           Call REPLA(TTYY$, CODEXT$(CI%), 1, 16)
           Call REPLA(TTYY$, DESCRIT0$(CI%), 19, 38)
           '\\\OT-05-0254-WAR-04/05/05///
           Call REPLA(TTYY$, FORMATNUM$(!CanPed, "F10.1"), 57, 10)
           Call REPLA(TTYY$, FORMATNUM$(!cantent, "F10.1"), 67, 10)
           Call REPLA(TTYY$, FORMATNUM$(CAN, "F10.1"), 77, 10)
           '\\\OT-05-0254-WAR-FIN///
           Call REPLA(TTYY$, FORMATNUM$(!NuOrItem, "F8.0"), 113, 8)
           SELITPED.List2.AddItem TTYY$
           .MoveNext
         End If
       Loop
     End With
     Pedianuittemp.Close
     Set Pedianuittemp = Nothing
       
     '
1209 SELITPED.Caption = "Anulacion por Items Pedido " + DPEDI$
     SELITPED.Show 1
     DoEvents
     '
     'VALIDACION DE CIERRE DE ORDENES DE FABRICACION GENERADAS A TRAVES DE PEDIDOS.
            
     If SELITPED.List2.ListCount > 0 Then
       Screen.MousePointer = 11
       '
       'VERIFICA QUE CADA ITEM A CANCELAR NO TENGAN ASIGNADO UN ARTICULO CON NUMERO DE SERIE
       For UY& = 1 To SELITPED.List2.ListCount
         TTXX$ = SELITPED.List2.List(UY& - 1)
         If SELITPED.List2.Selected(UY& - 1) = True Then
            CoIn% = CODINT%(Left$(TTXX$, 16))
            NUOR& = XVALO(Mid$(TTXX$, 113, 8))
            CONDI$ = "NroPed = " & NUPEDIDO& & " and CodiInt = " & CoIn% & " and NuOrItem = " & NUOR&
            NUMSERI$ = VALOBADA("PEDDETA", "NUMEROSERIE", CONDI$, "NOMESS")
            If TRIM$(NUMSERI$) <> "" Then
               Call MENSERR(24, "Imposible Cancelar Codigo " & CODEXT(CoIn%) & ".\ \ Pedido con Nº de Serie Asignado.\Primero Debe Desasignar Nº de Serie.")
               Screen.MousePointer = 1
               GoTo 1209
            End If
         End If
       Next UY&
       '
       'VERIFICA QUE CADA ITEM A CANCELAR NO TENGAN ASIGNADO UN ARTICULO CON NUMERO DE SERIE
       For UY& = 1 To SELITPED.List2.ListCount
         TTXX$ = SELITPED.List2.List(UY& - 1)
         If SELITPED.List2.Selected(UY& - 1) = True Then
            CoIn% = CODINT%(Left$(TTXX$, 16))
            NUOR& = XVALO(Mid$(TTXX$, 113, 8))
            CONDI$ = "NroPed = " & NUPEDIDO& & " and CodiInt = " & CoIn% & " and NuOrItem = " & NUOR&
            NroPedido = XVALO(VALOBADA("PEDDETA", "NROPED", CONDI$, "NOMESS"))
            CONDI2$ = " NroPed = " & NroPedido & " AND COD_INTE = " & CoIn%
            NORFA$ = TRIM$(VALOBADA("ORDEFABR", "IDSUBOR", CONDI2$, "NOMESS"))
            If NORFA$ <> "" Then
               OPX% = COMALTER%("Se ha Detectado Para el Código: " & CODEXT(CoIn%) & ".\Que el Mismo Posee Orden de Fabricación Asociada(" & NORFA$ & " )\Desea Realizar el Cierre y Consumo de la Misma ?\\No, Continuar sin Cerrar, \Si, Cerrar Orden de Fabricación")
               If OPX% > 1 Then
                  CANPE = XVALO(VALOBADA("PEDDETA", "CanPed", CONDI$, "NOMESS"))
                  CANTE = XVALO(VALOBADA("PEDDETA", "CantEnt", CONDI$, "NOMESS"))
                  SALDO = CANPE - CANTE
                  ARGU$ = TRIM$(Str(SALDO))
10                ARGU$ = UCase$(TRIM$(InputBox$("Cantidad", "Ingrese Cantidad a Cerrar Para la O.Fabricación", ARGU$)))
                  If XVALO(ARGU$) >= SALDO Then
                     If SALDO >= 0 Then Call CIERRORFACTU(CoIn%, NORFA$, XVALO(ARGU$), HOY(HO$))
                  Else
                     OPX% = COMALTER%("La Cantidad a Cerrar Debe Ser Igual o Mayor que La Cantidad Pendiente del Pedido\\Modificar Cantidad\Cancelar Transacción")
                     If OPX% = 1 Then ARGU$ = TRIM$(Str(SALDO)): GoTo 10               'VUELVE AL INPUTBOX
                     If OPX% < 1 Or OPX% > 1 Then GoTo 9999 'CANCELA O CIERRA VENTANA
                  End If
               End If
            End If
         End If
       Next UY&
       '
       FLAGANUL% = 1
       For U& = 1 To SELITPED.List2.ListCount
         TTXX$ = SELITPED.List2.List(U& - 1)
         '\\\OT-05-0254-WAR-05/05/05///
         If SELITPED.List2.Selected(U& - 1) = True Then 'Mid$(TTXX$, 71, 5) = "anul." Then
         '\\\OT-05-0254-WAR-FIN///
           CoIn% = CODINT%(Left$(TTXX$, 16))
           NUOR& = XVALO(Mid$(TTXX$, 113, 8))
           'jandy sql actualizo upadte
           'Busco el registro
'           With PedDeta
             abuss$ = " CodiInt = " & CoIn% & " "
             abuss$ = abuss$ + " AND NroPed = " & NUPEDIDO& & " "
             abuss$ = abuss$ + " AND NuOrItem = " & NUOR& & " "
             abuss$ = abuss$ + " AND Status < 8 "
             abuss$ = abuss$ + " AND Status <> 3 "
             strSQL = "Update PEDDETA set Status=8"
             strSQL = strSQL & " where  "
             strsql1 = strSQL & abuss$
             FLEXCONN.Execute (FILTSQL$(strsql1))
             FLAGANUL% = 1
         End If
       Next U&
 
       Call BLOQARCH("PEDDETP")
       Call RECAPEDPE
       Call CIERRARCH("PEDDETP")
       Screen.MousePointer = 1
       '
       If FLAGANUL% = 1 Then
            HOII% = HOY(HOS$)
            TTXX1$ = Space$(96)
            Call REPLA(TTXX1$, HOS$ + "  " + "Anulación Parcial de Items Pedidos", 1, 56)
            Call REPLA(TTXX1$, USERTER$, 69, 24)
            Call AGREGA_ANOTAPED(NUPEDIDO&, TTXX1$)
            For U& = 1 To SELITPED.List2.ListCount
              TTXX$ = SELITPED.List2.List(U& - 1)
              If SELITPED.List2.Selected(U& - 1) = True Then
                TTXX1$ = Space$(96)
                Codigo$ = Left$(TTXX$, 16)
                Call REPLA(TTXX1$, Space(10) + Codigo$, 1, 56)
                Call AGREGA_ANOTAPED(NUPEDIDO&, TTXX1$)
              End If
            Next U&
            Call AGREGA_ANOTAPED(NUPEDIDO&, String$(88, "="))
       End If
       '
     End If
     '
     Call CIERRARCH("*.*")
     Call BAJALDISCO
     Screen.MousePointer = 1
     DoEvents
     '
     GoTo 2000
     '
9999 Call CIERRARCH("*.*")
     Screen.MousePointer = 1
     '
End Sub

Sub CIERRORFACTU(CI1%, NORFA$, CANTI, FFACT%)
    Call ENRAPRODUCTO_TERMINADO(CI1%, NORFA$, CANTI, FFACT%)
    '
    CAN# = CANTI
    Call CARCOMPO(CODEXT$(CI1%), NORFA$, CAN#, 0, 1)
    Call CONSUCOMPONETESXOF(CI1%, NORFA$, CANTI, FFACT%)
    '
    'ACTUALIZA LA RESERVA
    FIN& = ULTREG&("!MATEROF")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      ZZ$ = REGLEIDO$("!MATEROF", U&)
      CICOMPO% = CVI(Left$(ZZ$, 2))
      CANCONSUM = CVS(Mid$(ZZ$, 5, 4))
      If TRIM$(UCase$(Unimed$(CICOMPO%))) = "GK" Then CANCONSUM = CANCONSUM / 1000
      '
      FECHA% = FFACT%
      Call ACTURESERVA(CICOMPO%, NORFA$, CANCONSUM)
    Next U&
    '
    Call ANOTAENORFA_FACTURADO(CI1%, NORFA$, CANTI, FFACT%)
    Call ACTUALIZAORFA(CI1%, NORFA$, CANTI, FFACT%)
End Sub
Sub Command5_Click()
  Command5.Enabled = False
  APRONP07.Show 1
  Command5.Enabled = True
End Sub

Private Sub Command6_Click()
    '
    Command6.Enabled = False
    '
    Dim SCP_CLI%(4096)        ' VECTOR CLIENTES
    Dim SCP_PED&(4096)        ' VECTOR PEDIDOS
    Dim SCP_COD%(4096)        ' VECTOR CODIGOS
    Dim SCP_SAL(4096)         ' VECTOR CANTIDAD PENDIENTE
    Dim SCP_ENT(4096)
    Dim SCP_ACU(4096)
    Dim SCP_PEX$(4096)        ' VECTOR PEDIDOS
    '
35  Call DESHASFECHA(Des%, Has%, PERIO$)
    If TRIM$(PERIO$) = "" Then GoTo 99
    '
    FEDES = FETEXCOR1$(Des%)
    FEHAS = FETEXCOR1$(Has%)
    '
    Call ABREPEDCLI
    Call ABRESTOCKS
    '
40  SQLX$ = " SELECT * "
    SQLX$ = SQLX$ + " FROM PEDDETA "
    SQLX$ = SQLX$ + " WHERE FechEmis >= '" & FEDES & "'"
    SQLX$ = SQLX$ + " AND FechEmis <= '" & FEHAS & "'"
    SQLX$ = SQLX$ + " AND STATUS >= 0 "
    '
    'Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim PEDTEMP As ADODB.Recordset
    Set PEDTEMP = New ADODB.Recordset
25  On Error Resume Next
    PEDTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0

    With PEDTEMP
      On Error Resume Next
      .MoveLast
      If Err > 1 Then
        Call COMUNI("No Existen Pedidos en el Período Seleccionado")
        GoTo 99
      End If
      FIN& = .RecordCount
      On Error GoTo 0
      '
      .MoveFirst
      KU& = 0
      Do While Not .EOF
         KU& = KU& + 1
         Call TRACE(20, KU&, FIN&)
         ' esto se agrega para considerar pedidos anulados pero con movimiento
         ' SEGúN PEDIDO DE ASEMBLI (OT-6-21)
         If !Status >= 8 Then    ' CASO DE PEDIDOS ANULADOS
           NUCLI7% = !NROCLIE
           NUPED7& = !NROPED
           CODIN7% = !CODIINT
           SQLY$ = " SELECT * "
           SQLY$ = SQLY$ + " FROM MOVISTO "
           SQLY$ = SQLY$ + " WHERE Nume_Clie = " & NUCLI7% & ""
           SQLY$ = SQLY$ + " AND NuPedCli = " & NUPED7& & ""
           SQLY$ = SQLY$ + " AND Cod_Inte = " & CODIN7% & ""
           '
           'Set STOBUS = Stocks.OpenRecordset(SQLY$, dbOpenDynaset)
           Dim STOBUS As ADODB.Recordset
           Set STOBUS = FLEXCONN.Execute(FILTSQL$(SQLY$))

           With STOBUS
             On Error Resume Next
             .MoveFirst
             If Err > 1 Then
               On Error GoTo 0
               GoTo 49        ' NO TIENE REMITOS
             End If
             On Error GoTo 0
           End With
         End If
         ' ***************************************************************
         For JJ1& = 1 To JJ&
           If SCP_CLI(JJ1&) = !NROCLIE Then
             If SCP_PED(JJ1&) = !NROPED Then
               If SCP_COD(JJ1&) = !CODIINT Then
                 SCP_SAL(JJ1&) = SCP_SAL(JJ1&) + !CanPed
                 SCP_ENT(JJ1&) = SCP_ENT(JJ1&) + !cantent
                 SCP_PEX$(JJ&) = SCP_PEX$(JJ&) + Chr$(124) + !TPedItem
                 GoTo 49
               End If
             End If
           End If
         Next JJ1&
         '
         JJ& = JJ& + 1
         SCP_CLI(JJ&) = !NROCLIE
         SCP_PED(JJ&) = !NROPED
         SCP_COD(JJ&) = !CODIINT
         SCP_SAL(JJ&) = !CanPed
         SCP_ENT(JJ&) = !cantent
         SCP_PEX$(JJ&) = !TPedItem
         '
49    .MoveNext
      Loop
      '
    End With
    PEDTEMP.Close
    Set PEDTEMP = Nothing
    STOBUS.Close
    Set STOBUS = Nothing
    '
    SQLX$ = " SELECT * "
    SQLX$ = SQLX$ + " FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE Nume_Clie > 0 "
    SQLX$ = SQLX$ + " AND NuPedCli > 0 "
    SQLX$ = SQLX$ + " AND CodiMovi = 'RT' "
    '
    'Set STOCKTEMP = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
    Dim STOCKTEMP As ADODB.Recordset
    Set STOCKTEMP = New ADODB.Recordset
26  On Error Resume Next
    STOCKTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 26
    End If
    On Error GoTo 0

    With STOCKTEMP
      On Error Resume Next
      .MoveLast
      If Err > 1 Then GoTo 80
      FIN& = .RecordCount
      On Error GoTo 0
      '
      .MoveFirst
      KU& = 0
      Do While Not .EOF
        KU& = KU& + 1
        Call TRACE(20, KU&, FIN&)
        NCL% = !NUME_CLIE
        NPED& = !NUPEDCLI
        CIXI% = !cod_inte
        CASAL = !Cantsalid - !CANTINGRE
        '
        For U& = 1 To JJ&
          If SCP_CLI(U&) = NCL% Then
            If SCP_PED(U&) = NPED& Then
              If SCP_COD(U&) = CIXI% Then
                SCP_ACU(U&) = SCP_ACU(U&) + CASAL
                If SCP_SAL(U&) >= CASAL Then
                    SCP_SAL(U&) = SCP_SAL(U&) - CASAL
                    SCP_ENT(U&) = SCP_ENT(U&) - CASAL
                    CASAL = 0
                    GoTo 59
                  Else
                    CASAL = CASAL - SCP_SAL(U&)
                    SCP_ENT(U&) = SCP_ENT(U&) - SCP_SAL(U&)
                    SCP_SAL(U&) = 0
                End If
              End If
            End If
          End If
        Next U&
        '
        'TTXX$ = "Sobre-Embarque Remito " + !DOPRICOR + " - Pedido " + TRIM$(Str$(NPED&)) + " - Codigo: " + CODEXT$(CIXI%)
        'Call MENSERR(24, TTXX$)
        '
59    .MoveNext
      Loop
    End With
    STOCKTEMP.Close
    Set STOCKTEMP = Nothing
    '
80  FIN& = JJ&
    HAYRECU% = 0
    'With PedDeta
    SQLX$ = " SELECT * FROM PEDDETA"
    SQLX$ = SQLX$ + " WHERE NroClie = " & SCP_CLI(U&) & ""
    SQLX$ = SQLX$ + " AND NroPed = " & SCP_PED(U&) & ""
    SQLX$ = SQLX$ + " AND CodiInt = " & SCP_COD(U&) & ""
    Dim PEDDETTTA As ADODB.Recordset
    Set PEDDETTTA = New ADODB.Recordset
27  On Error Resume Next
    PEDDETTTA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 27
    End If
    On Error GoTo 0
    With PEDDETTTA
      For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
         If Abs(SCP_ENT(U&)) > 0.01 Then
           TTXX$ = "Error de Saldo Pendiente\En Pedido " & SCP_PED(U&) & "\Codigo " + CODEXT$(SCP_COD(U&))
           OPX% = COMALTER%(TTXX$ + "\\Continuar\Recuperar")
           If OPX% = 2 Then
             HAYRECU% = HAYRECU% + 1
             SALENT = SCP_ACU(U&)
'             SQLX$ = " NroClie = " & SCP_CLI(U&) & " "
'             SQLX$ = SQLX$ & "AND NroPed = " & SCP_PED(U&) & " "
'             SQLX$ = SQLX$ & "AND CodiInt = " & SCP_COD(U&) & " "
'             '
'             .FindFirst SQLX$
             'If .NoMatch = False Then
               '.Edit
             If Not (.EOF And .BOF) Then
                If SALENT <= !CanPed Then
                    !cantent = SALENT
                    !CanSaldo = !CanPed - !cantent
                    SALENT = 0
                  Else
                    SALENT = SALENT - !CanPed
                    !cantent = !CanPed
                    !CanSaldo = 0
                End If
                .Update
                '.FindNext SQLX$
                .MoveNext
             End If
           End If
         End If
      Next U&
    End With
    PEDDETTTA.Close
    Set PEDDETTTA = Nothing
    '
    TTXX$ = "Fin del Control"
    If HAYRECU% > 0 Then
        TTXX$ = TTXX$ + " con " + TRIM$(Str$(HAYRECU%)) + " Error(es) Recuperados"
      Else
        TTXX$ = TTXX$ + " sin Errores Recuperados"
    End If
    Call COMUNI(TTXX$)
    '
99  Screen.MousePointer = 1
    Command6.Enabled = True
    '
End Sub

 Sub Command61_Click()
    '\\\OT-05-0250-WAR-20/05/05///
    Command61.Enabled = False
'    If Option1 Then
'      Call COMUNI("Listado Solo Disponible por Impresión.")
'    End If
    Call LIENCPED
    Command61.Enabled = True
    '\\\OT-05-0250-WAR-FIN///
End Sub

 Sub Command62_Click()
    Dim PEDTEMP As ADODB.Recordset
    Command62.Enabled = False
    '
    MODOGRAMY% = 0
    If Control$("ADMIPED", "FORMALIS") = "MONEMORI" Then
      If EXISTE%(LUDAT$ + "LIPEDID.RFS") > 0 Then
        MODOGRAMY% = 1
      End If
    End If
    '
    '\\\OT-05-0249-WAR-04/05/05///
    If MODOGRAMY% < 1 Then Call MENSPED
    '\\\OT-05-0249-WAR-FIN///
    '
5   OPX% = ALTERNA%("Detalle Pedidos Pendientes\Seguimiento Comercial")
    '
    If OPX% = 1 Then
        If Option1.Value = True Then
           'Consultas por Pantalla
7          NC = ListaSeleccionClientes()
           Call PEDIPENGRILLA(NC, 0)
'           GoTo 7
           
'           Call RECAPEDPE
'           sali$ = "*COPEDCL"
'           If FORMAPED% = 2 Then sali$ = "*COQEDCL"
'           Call FINAL(sali$)
        Else
           'Listados
           Screen.MousePointer = 11
           JJ& = 0: LL& = 0
           '
           MODO% = COMALTER%("Clasificación por Fecha de Entrega Solicitada:\\Mes en Curso\Posteriores\Total")
           If MODO% < 1 Then
               GoTo 99
             ElseIf MODO% = 1 Then
               FLI1% = -1: FLI2% = FINMESACT%
             ElseIf MODO% = 2 Then
               FLI1% = FINMESACT% + 1: FLI2% = 32767
10             FNMA = FINMESACT%
               FLI1% = DIPOST(FNMA): FLI2% = 32767
               VDEF$ = MKI$(FLI1%) + MKI$(FLI2%)
               OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
               If OBX$ = "" Then GoTo 99
               '
               FLI1% = CVI(Left$(OBX$, 2))
               FLI2% = CVI(Mid$(OBX$, 3, 2))
               If FLI1% > FLI2% Then GoTo 10
               FEX = FLI1%: PERIO$ = FECHATEX$(FEX)
               FEX = FLI2%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
             ElseIf MODO% = 3 Or MODO% = 4 Then
               FLI1% = -1: FLI2% = 32767
           End If
           '
           OPX% = ALTERNA%("Solo Cantidades\Cantidades y Precios")
           If OPX% < 1 Or OPX% > 2 Then GoTo 99
           '
           Call SELRACLIENUE(NCL1, NCL2, VENDSEL%)
           If NCL1 < 1 Then GoTo 99
           '
           If ULTREG&("GRUCOCLI") > 0 Then
                Call SELECHO("GRUCOCLI")
                GrupoCliente% = XVALO(VALACT1$("GRUCOCLI"))
           End If
           '
           FORMASA$ = ""
           If EXISTE%(LUDAT$ + "DETPEPEN.RFS") Then
             FORMASA$ = "ASEMBLI"
           End If
           '
           VECPEDI$ = ""
           'nuevo
           ' ver ernesto no se puede usar la Sub RECAPEDPE()??

           Call ABREPEDCLI
'           SQLX$ = " SELECT * "
'           SQLX$ = SQLX$ + " FROM PEDDETA "
'           SQLX$ = SQLX$ + " WHERE NroClie >= " & NCL1 & " "
'           SQLX$ = SQLX$ + " AND  NroClie <= " & NCL2 & " "
'           SQLX$ = SQLX$ + " AND Status < 8 "
'           If MODOGRAMY% < 1 Then SQLX$ = SQLX$ + " AND STATUS >= 0 "
'           'sqlx$ = sqlx$ + " AND (((CanPed - CantEnt) >= 0.05) OR (TipoPed = 'ABIERTO' AND CDBL(FinVigen) >= " & Int(CDbl(Now)) & ")) "
'           SQLX$ = SQLX$ + " AND (((CanPed - CantEnt) >= 0.05) OR (TipoPed = 'ABIERTO' AND FinVigen >= getdate()))"
'           'ORDEN AGREGADO A LA CONSULTA, POR QUE FILESORT DABA ERROR
'           SQLX$ = SQLX$ + " ORDER BY NroClie,CODIINT,NROOCOM"
'           '
'           'Set PEDTEMP = Ventas.OpenRecordset(sqlx$, dbOpenSnapshot)
'           '
'           Set PEDTEMP = New ADODB.Recordset
'           PEDTEMP.CursorLocation = adUseClient
'25         On Error Resume Next
'           PEDTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
'           If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'              On Error GoTo 0
'              Call CapturaErrorOpen
'              GoTo 25
'           End If
'           On Error GoTo 0
'CREO UNA TABLA TEMPORAL (EL PROBLEMA ERA QUE EL ORDENAMIENTO HACIA QUE DEMORE DEMASIADO TIEMPO Y EL FILTRO OR
SQLX$ = "SELECT * INTO #TempPEDDETA "
SQLX$ = SQLX$ + " From PedDeta WITH(NOLOCK)"
SQLX$ = SQLX$ + " WHERE NroClie >= " & NCL1 & " "
SQLX$ = SQLX$ + " AND  NroClie <= " & NCL2 & " "
SQLX$ = SQLX$ + " AND Status < 8 "
SQLX$ = SQLX$ + " AND STATUS >= 0 "
SQLX$ = SQLX$ + "  AND (((CanPed - CantEnt) >= 0.05) OR (TipoPed = 'ABIERTO' AND FinVigen >= GETDATE()));"
FLEXCONN.Execute (SQLX$)

' Consultar y ordenar desde la tabla TEMPORAL *
SQLX$ = " SELECT * "
SQLX$ = SQLX$ + " FROM #TempPEDDETA WITH(NOLOCK)"
SQLX$ = SQLX$ + " ORDER BY NroClie, CODIINT, NROOCOM;"
Set PEDTEMP = READSET(SQLX$)


           With PEDTEMP
             On Error Resume Next
             .Move first
             '
             JJ& = 0
             Y$ = REGBLAN("LIPEDPE")
             '
             Do While (.EOF = False)
                CI% = XVALO(!CODIINT)
             If CI% > 0 And CI% <= NUMAS% Then
                CAPEN = !CanPed - !cantent
              '
              If CAPEN < 0 Then CAPEN = 0
                FENTRESOL% = CVINT(!FeSolEnt)
                '
                If FENTRESOL% >= FLI1% Then
                  If FENTRESOL% <= FLI2% Then
                    NCLIE = !NROCLIE
                    If (VENDSEL% < 1 Or VENDCLI%(NCLIE) = VENDSEL%) And (GrupoCliente% < 1 Or GRUCOCLI%(NCLIE) = GrupoCliente%) Then
                          Call REPLA(Y$, Mid$(X$, 7, 2), 1, 2)
                          Call REPLA(Y$, Mid$(X$, 7, 2), 45, 2)
                          Call REPLA(Y$, Mid$(X$, 33, 42), 3, 42)
                          CAPEN = !CanPed - !cantent
                        '
                        If CAPEN < 0 Then CAPEN = 0
                          CANOASI = CAPEN - !CanAsig
                        '
                        If CANOASI < 0 Then CANOASI = 0
                          MONEMISPED% = !MoneEmis
                        PREUNI = !PreUnid * (100 - PORDESCLI(NCLIE)) / 100 * TICAMONE(MONEMISPED%)
                        '
                        If OPX% <> 2 Then PREUNI = 0
                        ITOTA# = CDbl(CAPEN * PREUNI)
                        INOAS# = CDbl(CANOASI * PREUNI)
                        '
                        Call REPLA(Y$, MKS$(NCLIE), 1, 4)
                        Call REPLA(Y$, MKI$(!CODIINT), 5, 2)
                        Call REPLA(Y$, !TCOLOR, 7, 16)
                        Call REPLA(Y$, !TTALLE, 23, 8)
                        Call REPLA(Y$, MKS$(!TLargo), 31, 4)
                        Call REPLA(Y$, MKS$(!TAncho), 35, 4)
                        Call REPLA(Y$, MKS$(!TEspesor), 39, 4)
                        Call REPLA(Y$, MKS$(!NROPED), 51, 4)
                        Call REPLA(Y$, MKI$(FENTRESOL%), 73, 2)
                        Call REPLA(Y$, MKS$(!CanPed), 83, 4)
                        Call REPLA(Y$, MKS$(CAPEN), 103, 4)
                        Call REPLA(Y$, MKS$(PREUNI), 87, 4)
                        Call REPLA(Y$, MKD$(ITOTA#), 91, 8)
                        Call REPLA(Y$, MKS$(!cantent), 99, 4)
                        Call REPLA(Y$, MKS$(CANOASI), 113, 4)
                        Call REPLA(Y$, MKD$(INOAS#), 121, 8)
                        Call REPLA(Y$, !NROOCOM, 129, 16)
                        Call REPLA(Y$, Space$(18), 145, 18)
                        '
                        If FORMASA$ = "ASEMBLI" Then
                          STAX% = !Status: If STAX% < 0 Then STAX% = STAX% + 256
                          Call REPLA(Y$, Chr$(STAX%), 159, 1)
                          Call REPLA(Y$, Chr$(MONEMISPED%), 160, 1)
                        End If
                        '
                        JJ& = JJ& + 1
                        Call GRAREG("LIPEDPE", Y$, JJ&)
                        '
                    End If
                  End If
                End If
             End If
             .MoveNext
             Loop
             On Error GoTo 0
           End With
           PEDTEMP.Close
           Set PEDTEMP = Nothing
           'ELIMINO LA TABLA TEMPORAL
           SQLX$ = "DROP TABLE #TempPEDDETA"
           FLEXCONN.Execute (SQLX$)
           '
30         Call HEADERS("LIPEDPE", "")
           If MODO% = 2 Then
              Call HEADERS("LIPEDPE", "Entrega Solicitada: " + PERIO$)
           End If
           '
           Call SETULTREG("LIPEDPE", JJ&)
           Call CIERRARCH("LIPEDPE")
           '
           VARIANTE$ = "Totales por Entregar"
           Call HEADERS("LIPEDPE", "Corresponde a: " + VARIANTE$)
           If VENDSEL% > 0 Then
             Call HEADERS("LIPEDPE", "Vendedor: " + ECOARCH$("VENDEDOR", Chr$(VENDSEL%)))
           End If
           If GrupoCliente% > 0 Then
             Call HEADERS("LIPEDPE", "Grupo Cliente: " + ECOARCH$("GRUCOCLI", CStr(GrupoCliente%)))
           End If
           '
           If FORMASA$ = "ASEMBLI" Then
             Call FILESORT("LIPEDPE", "", 3, 3, "ASC") ' POR CODIGO
             Call FILESORT("LIPEDPE", "", 13, 13, "ASC") ' POR PEDIDO
             Call FILESORT("LIPEDPE", "", 1, 1, "ASC") ' POR CLIENTE
             '
             JJ& = 0
             For NCL% = NCL1 To NCL2
               RFF$ = RECFILT$("LIPEDPE", 1, MKS$(NCL%))
               If Len(RFF$) > 0 Then
                 REBLA$ = REGBLAN$("DETPEPEN")
                 Call REPLA(REBLA$, MKI$(NCL%), 1, 2)
                    RASO1$ = TRIM$(Str$(NCL%)) + " - " + rasocli$(NCL%)
                 Call REPLA(REBLA$, RASO1$, 3, 48)
                 Call REPLA(REBLA$, DOMICLI$(NCL%), 51, 48)
                 Call REPLA(REBLA$, LOCACLI$(NCL%), 99, 48)
                 '
                 For KL& = 1 To Len(RFF$) Step 4
                   RELI& = CVS(Mid$(RFF$, KL&, 4))
                   LIPE$ = REGLEIDO$("LIPEDPE", RELI&)
                   ZZ$ = REBLA$
                     '
                     NPED& = CVS(Mid$(LIPE$, 51, 4))
                   Call REPLA(ZZ$, MKS$(NPED&), 147, 4)
                     STA$ = "  "
                     If Asc(Mid$(LIPE$, 159, 1)) = 255 Then STA$ = "P."
                     If Asc(Mid$(LIPE$, 159, 1)) = 8 Then STA$ = "AN"
                     If Asc(Mid$(LIPE$, 159, 1)) = 9 Then STA$ = "AN"
                   Call REPLA(ZZ$, STA$, 151, 2)
                     NOCO$ = Mid$(LIPE$, 129, 16)
                   Call REPLA(ZZ$, NOCO$, 155, 16)
                     CIXI% = CVI(Mid$(LIPE$, 5, 2))
                   Call REPLA(ZZ$, MKI$(CIXI%), 171, 2)
                     FECHENT% = CVI(Mid$(LIPE$, 73, 2))
                   Call REPLA(ZZ$, MKI$(FECHENT%), 173, 2)
                     CANORI = CVS(Mid$(LIPE$, 83, 4))
                   Call REPLA(ZZ$, MKS$(CANORI), 175, 4)
                     CANENT = CVS(Mid$(LIPE$, 99, 4))
                   Call REPLA(ZZ$, MKS$(CANENT), 179, 4)
                     CANSAL = CVS(Mid$(LIPE$, 103, 4))
                   Call REPLA(ZZ$, MKS$(CANSAL), 183, 4)
                     MONEM% = Asc(Mid$(LIPE$, 160, 1))
                   Call REPLA(ZZ$, SIMBOLPES$(MONEM%), 187, 4)
                     PREUN# = CVS(Mid$(LIPE$, 87, 4)) / TICAMONE(MONEM%)
                     ' SI EL TIPO DE MONEDA ES 0 O 1 EL P.UNITARIO SE
                     ' GUARDA EN PESOS Y SE CALCULA EL P.UNITARIO EN DOLARES
                     If MONEM% < 2 Then
                       PREPES# = PREUN#
                       Call REPLA(ZZ$, MKD$(PREPES#), 191, 8)
                       On Error Resume Next
                       PREDOL# = PREUN# * TICAMONE(MONEM%) / TICAMONE(2)
                       On Error GoTo 0
                       Call REPLA(ZZ$, MKD$(PREDOL#), 199, 8)
                     End If
                     ' SI EL TIPO DE MONEDA ES 2 EL P.UNITARIO SE GUARDA EN
                     ' DOLARES Y SE CALCULA EL P. UNITARIO EN PESOS
                     If MONEM% = 2 Then
                       PREDOL# = PREUN#
                       Call REPLA(ZZ$, MKD$(PREDOL#), 199, 8)
                       On Error Resume Next
                       PREPES# = PREUN# * TICAMONE(2)
                       On Error GoTo 0
                       Call REPLA(ZZ$, MKD$(PREPES#), 191, 8)
                     End If
                   Call REPLA(ZZ$, MKD$(PREDOL# * CANSAL), 207, 8)
                   JJ& = JJ& + 1
                   Call GRAREG("DETPEPEN", ZZ$, JJ&)
                 Next KL&
                 '
                 JJ& = JJ& + 1
                 Call GRAREG("DETPEPEN", String$(256, "="), JJ&)
               End If
             Next NCL%
             Call SETULTREG("DETPEPEN", JJ&)
             '
             Call HEADERS("DETPEPEN", "")
             'Call HEADERS("DETPEPEN", "T.Cambio $ / U$S: " + TRIM$(FORMATNUM$(TICAMONE(2), "F12.4")))
             '
             Call FINAL("*DETPEPEN")
             GoTo 99
             '
           End If
           '
           Screen.MousePointer = 1
           '\\\OT-05-0248-WAR-04/05/05///
'           Call FILESORT("LIPEDPE", "", 16, 16, "ASC")
'           Call FILESORT("LIPEDPE", "", 2, 2, "ASC")
'           Call FILESORT("LIPEDPE", "", 1, 1, "ASC")
           '\\\OT-05-0248-WAR-FIN///
           'Call FILESORT("LIPEDPE", "", 1, 9, "ASC")
           sali$ = "*LIPENCL"
           Call FINAL(sali$)
        End If
    ElseIf OPX% = 2 Then
        '\\\OT-5-686-FG-20/10/05///
        Dim SCP_CLI(4096)         ' VECTOR CLIENTES
        Dim SCP_PED&(4096)        ' VECTOR PEDIDOS
        Dim SCP_COD%(4096)        ' VECTOR CODIGOS
        Dim SCP_REG&(4096)        ' VECTOR REGISTRO PEDIDO
        Dim SCP_STAT%(4096)       ' VECTOR STATUS PEDIDO
        Dim COEFPES#(4096)   ' VECTOR TOTAL PEDIDO
        Dim COEFDOL#(4096)   ' VECTOR ANTICIPO EN DOLARES
        '
35      Call DESHASFECHA(Des%, Has%, PERIO$)
        If TRIM$(PERIO$) = "" Then GoTo 5
        '
        FEDES = CDbl(CVDAT(Des%))
        FEHAS = CDbl(CVDAT(Has%))
        '
        Call ABREPEDCLI
        Call ABRESTOCKS
        '
        '\\\OT-06-0098-WAR-20/03/06///
37      R1% = COMALTER%("Seleccione la Opción de Filtro\\Todos\Por Cliente")
        If R1% = 0 Then GoTo 35
        If R1% = 2 Then
          Call SELECHO("DEUDOR1")
          NCLIE = XVALO(TRIM$(VALACT1$("DEUDOR1")))
          If NCLIE <= 0 Then GoTo 37
        End If
        ''\\\OT-06-0098-WAR-FIN///
        Desde = FETEXCOR1$(Des%)
        Hasta = FETEXCOR1$(Has%)

40      SQLX$ = " SELECT * "
        SQLX$ = SQLX$ + " FROM PEDDETA "
        'sqlx$ = sqlx$ + " WHERE INT(CDBL(FechEmis)) >= " & FEDES & ""
        'sqlx$ = sqlx$ + " AND INT(CDBL(FechEmis)) <= " & FEHAS & ""
        SQLX$ = SQLX$ + " WHERE FechEmis >= '" & Desde & "'"
        SQLX$ = SQLX$ + " AND FechEmis <= '" & Hasta & "'"
        'Jandy ESTO YA ESTABA ASI
        'SQLX$ = SQLX$ + " AND Status < 8 "
        If MODOGRAMY% < 1 Then SQLX$ = SQLX$ + " AND STATUS >= 0 "
        '\\\OT-06-0098-WAR-20/03/06///
        If R1% = 2 Then
          If NCLIE > 0 Then
            SQLX$ = SQLX$ + "AND NroClie = " & NCLIE & ""
          End If
        End If
        '\\\OT-06-0098-WAR-FIN///
        '
        'Set PEDTEMP = Ventas.OpenRecordset(sqlx$, dbOpenSnapshot)
        '
        
        Set PEDTEMP = New ADODB.Recordset
        PEDTEMP.CursorLocation = adUseClient
26      On Error Resume Next
        PEDTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 26
        End If
        On Error GoTo 0
        With PEDTEMP
          On Error Resume Next
          .MoveFirst
          If Err > 1 Then
            Call COMUNI("No Existen Pedidos en el Período Seleccionado")
            GoTo 99
          End If
          '
          Call BLANARCH("SECOMPED")
          JJ& = 0
          REBLA$ = REGBLAN("SECOMPED")
          '
          Do While Not .EOF
            ' esto se agrega para considerar pedidos anulados pero con movimiento
            ' SEGúN PEDIDO DE ASEMBLI (OT-6-21)
            If !Status >= 8 Then    ' CASO DE PEDIDOS ANULADOS
               NUCLI7 = !NROCLIE
               NUPED7& = !NROPED
               CODIN7% = !CODIINT
               SQLY$ = " SELECT * FROM MOVISTO "
               SQLY$ = SQLY$ + " WHERE Nume_Clie = " & NUCLI7
               'SQLY$ = SQLY$ + " AND CDBL(NuPedCli) = " & NUPED7&
               'jandy sql
               SQLY$ = SQLY$ + " AND NuPedCli = " & NUPED7&
               SQLY$ = SQLY$ + " AND Cod_Inte = " & CODIN7%
               '
               'Set STOBUS = Stocks.OpenRecordset(SQLY$, dbOpenDynaset)
               Dim STOBUS As ADODB.Recordset
               Set STOBUS = FLEXCONN.Execute(FILTSQL$(SQLY$))

  'ver ernesto
  
               With STOBUS
                 On Error Resume Next
                 .MoveFirst
                 If Err > 1 Then
                   On Error GoTo 0
                   GoTo 49        ' NO TIENE REMITOS
                 End If
                 On Error GoTo 0
               End With
            End If
            ' ***************************************************************
          
            Y$ = REBLA$
            '
            Call REPLA(Y$, MKS$(!NROCLIE), 1, 4)
            Call REPLA(Y$, MKS$(!NROPED), 149, 4)
            Call REPLA(Y$, !NROOCOM, 153, 16)
            Call REPLA(Y$, MKI$(!CODIINT), 169, 2)
            Call REPLA(Y$, Chr$(1), 171, 1) ' MARCA DE QUE ES UN PEDIDO
              FECHENT% = CVINT%(!FeSolEnt)
            Call REPLA(Y$, MKI$(FECHENT%), 172, 2)
              MONEORIG% = !MoneEmis
              TICAM# = !TIPOCAM: If TICAM# < 0.01 Then TICAM# = 1
              PRUNI# = !PreUnid
              If MONEORIG% <= 1 Then
                PRUNI# = PRUNI# / TICAM#
              End If
            Call REPLA(Y$, MKD$(PRUNI#), 174, 8)  'precio unitario en dolares
              CANTI# = (-1) * !CanPed
            Call REPLA(Y$, MKD$(CANTI#), 182, 8)
            Call REPLA(Y$, String(32, " "), 190, 32)
            Call REPLA(Y$, MKD$(TICAM#), 222, 8)
             ' SI MONEDA ES PESOS (0 - 1), GUARDA PESOS Y CALCULA DOLARES
              If MONEORIG% <= 1 Then
                  NETPES# = (-1) * !ImpoTot
                  IMPDOL# = NETPES# / TICAM#
                Else
                  IMPDOL# = (-1) * !ImpoTot
                  NETPES# = IMPDOL# * TICAM#
              End If
            Call REPLA(Y$, MKD$(IMPDOL#), 230, 8)
            Call REPLA(Y$, MKD$(NETPES#), 238, 8)
              GIVA% = GRUPIVA%(!CODIINT)
              TASA# = CVS(Mid$(FILTERGETRECORD("GRUPIVA", 1, MKI$(GIVA%)), 13, 4))
              IMPIVA# = (NETPES# * TASA#) / 100
            Call REPLA(Y$, MKD$(IMPIVA#), 246, 8)
              IMPTOT# = NETPES# + IMPIVA#
            Call REPLA(Y$, MKD$(IMPTOT#), 254, 8)
              IANTI# = 0
              On Error Resume Next
              IANTI# = !ImpoAnti
              On Error GoTo 0
              ' SI MONEDA ES PESOS (0 - 1), GUARDA PESOS Y CALCULA DOLARES
              If MONEORIG% <= 1 Then
                  ANTIPES# = IANTI#
                  ANTIDOL# = IANTI# / TICAM#
                Else
                  ANTIDOL# = IANTI#
                  ANTIPES# = IANTI# * TICAM#
              End If
            Call REPLA(Y$, MKD$(ANTIPES#), 270, 8)
            Call REPLA(Y$, MKD$(ANTIDOL#), 268, 8)
            '
            JJ& = JJ& + 1
            Call GRAREG("SECOMPED", Y$, JJ&)
            '
            SCP_CLI(JJ&) = !NROCLIE
            SCP_PED(JJ&) = !NROPED
            SCP_COD(JJ&) = !CODIINT
            SCP_REG&(JJ&) = JJ&           ' VECTOR REGISTRO PEDIDO
            SCP_STAT%(JJ&) = !Status      ' VECTOR STATUS PEDIDO
            If Abs(IMPDOL#) > 0.1 Then
              COEFDOL#(JJ&) = ANTIDOL# / IMPDOL#
            End If
            If Abs(NETPES#) > 0.1 Then
               COEFPES#(JJ&) = ANTIPES# / NETPES#
            End If
            '
49        .MoveNext
          Loop
          '
        End With
        PEDTEMP.Close
        Set PEDTEMP = Nothing
        '
        
        FIN& = JJ&
        For U& = 1 To FIN&
           '
           If SCP_COD(U&) < 1 Then GoTo 89
           '
           XREG$ = REGLEIDO$("SECOMPED", SCP_REG&(U&))
           CANACU# = CVD(Mid$(XREG$, 180, 8))
           '
           
           SQLX$ = " SELECT * FROM MOVISTO "
           SQLX$ = SQLX$ + " WHERE Nume_Clie = " & SCP_CLI(U&)
           'SQLX$ = SQLX$ + " AND CDBL(NuPedCli) = " & SCP_PED(U&)
           SQLX$ = SQLX$ + " AND NuPedCli = " & SCP_PED(U&)
           SQLX$ = SQLX$ + " AND Cod_Inte = " & SCP_COD(U&)
           '
           'Set STOCKTEMP = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
           
           Dim STOCKTEMP As ADODB.Recordset
           Set STOCKTEMP = New ADODB.Recordset
27         On Error Resume Next
           STOCKTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
           If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
              On Error GoTo 0
              Call CapturaErrorOpen
              GoTo 27
           End If
           On Error GoTo 0
           With STOCKTEMP
              On Error Resume Next
              .MoveFirst
              If Err > 1 Then GoTo 80
              On Error GoTo 0
              '
              Do While Not .EOF
                Y$ = REBLA$
                '
                Call REPLA(Y$, MKS$(SCP_CLI(U&)), 1, 4)
                Call REPLA(Y$, MKS$(SCP_PED(U&)), 149, 4)
                'Call REPLA(Y$, String(16, " "), 151, 16)
                NPEDX1& = SCP_PED(U&)
                Call REPLA(Y$, NOCOMCLI$(NPEDX1&), 153, 16)
 
                Call REPLA(Y$, MKI$(SCP_COD(U&)), 169, 2)
                Call REPLA(Y$, Chr$(2), 169, 1) ' MARCA DE QUE ES UN REMITO
                  FECHENT% = CVINT%(!FechMov)
                Call REPLA(Y$, MKI$(FECHENT%), 172, 2)
                  VUNI# = !ValoUnit
                  MONERE% = 1: TICAM# = 0
                  On Error Resume Next
                  MONERE% = !MoneVal
                  TICAM# = !TIPOCAM
                  On Error GoTo 0
                  '
                  NROX$ = !DOSECLAR
                  FEX = FECHENT%
                  NCLIEN% = SCP_CLI(U&)
                  MONERE1% = MONEMISOR%(NROX$, FEX, NCLIEN%)
                  If MONERE1% <> MONERE% Then
                    '.Edit
                    !MoneVal = MONERE1%
                    .Update
                    MONERE% = MONERE1%
                  End If
                  '
                  If TICAM# < 0.01 Then
                    If FECHENT% = FECHANUM("21/03/05") Then TICAM# = 2.92: GoTo 14
                    If FECHENT% = FECHANUM("11/04/05") Then TICAM# = 2.917: GoTo 14
                    If FECHENT% = FECHANUM("25/04/05") Then TICAM# = 2.861: GoTo 14
                    If FECHENT% = FECHANUM("26/05/05") Then TICAM# = 2.886: GoTo 14
                    If FECHENT% = FECHANUM("08/06/05") Then TICAM# = 2.854: GoTo 14
                    If FECHENT% = FECHANUM("09/06/05") Then TICAM# = 2.895: GoTo 14
                    If FECHENT% = FECHANUM("17/06/05") Then TICAM# = 2.874: GoTo 14
                    If FECHENT% = FECHANUM("28/06/05") Then TICAM# = 2.849: GoTo 14
                    If FECHENT% = FECHANUM("29/06/05") Then TICAM# = 2.889: GoTo 14
                    If FECHENT% = FECHANUM("30/06/05") Then TICAM# = 2.855: GoTo 14
                    If FECHENT% = FECHANUM("13/07/05") Then TICAM# = 2.886: GoTo 14
                    If FECHENT% = FECHANUM("14/07/05") Then TICAM# = 2.869: GoTo 14
                    If FECHENT% = FECHANUM("15/07/05") Then TICAM# = 2.864: GoTo 14
                    If FECHENT% = FECHANUM("18/07/05") Then TICAM# = 2.863: GoTo 14
                    If FECHENT% = FECHANUM("19/07/05") Then TICAM# = 2.864: GoTo 14
                    If FECHENT% = FECHANUM("21/07/05") Then TICAM# = 2.862: GoTo 14
                    If FECHENT% = FECHANUM("22/07/05") Then TICAM# = 2.863: GoTo 14
                    If FECHENT% = FECHANUM("25/07/05") Then TICAM# = 2.863: GoTo 14
                    If FECHENT% = FECHANUM("26/07/05") Then TICAM# = 2.862: GoTo 14
                    If FECHENT% = FECHANUM("28/07/05") Then TICAM# = 2.865: GoTo 14
                    If FECHENT% = FECHANUM("29/07/05") Then TICAM# = 2.863: GoTo 14
                    If FECHENT% = FECHANUM("02/08/05") Then TICAM# = 2.86: GoTo 14
                    If FECHENT% = FECHANUM("03/08/05") Then TICAM# = 2.859: GoTo 14
                    If FECHENT% = FECHANUM("05/08/05") Then TICAM# = 2.863: GoTo 14
                    If FECHENT% = FECHANUM("08/08/05") Then TICAM# = 2.865: GoTo 14
                    If FECHENT% = FECHANUM("10/08/05") Then TICAM# = 2.873: GoTo 14
                    If FECHENT% = FECHANUM("12/08/05") Then TICAM# = 2.877: GoTo 14
                    If FECHENT% = FECHANUM("18/08/05") Then TICAM# = 2.89: GoTo 14
                    If FECHENT% = FECHANUM("19/08/05") Then TICAM# = 2.907: GoTo 14
                    If FECHENT% = FECHANUM("22/08/05") Then TICAM# = 2.907: GoTo 14
                    If FECHENT% = FECHANUM("24/08/05") Then TICAM# = 2.908: GoTo 14
                    If FECHENT% = FECHANUM("25/08/05") Then TICAM# = 2.867: GoTo 14
                    If FECHENT% = FECHANUM("29/08/05") Then TICAM# = 2.907: GoTo 14
                    If FECHENT% = FECHANUM("31/08/05") Then TICAM# = 2.873: GoTo 14
                    If FECHENT% = FECHANUM("05/09/05") Then TICAM# = 2.907: GoTo 14
                    If FECHENT% = FECHANUM("06/09/05") Then TICAM# = 2.907: GoTo 14
                    If FECHENT% = FECHANUM("07/09/05") Then TICAM# = 2.907: GoTo 14
                    If FECHENT% = FECHANUM("09/09/05") Then TICAM# = 2.91: GoTo 14
                    If FECHENT% = FECHANUM("13/09/05") Then TICAM# = 2.91: GoTo 14
                    If FECHENT% = FECHANUM("15/09/05") Then TICAM# = 2.919: GoTo 14
                    If FECHENT% = FECHANUM("16/09/05") Then TICAM# = 2.919: GoTo 14
                    If FECHENT% = FECHANUM("19/09/05") Then TICAM# = 2.911: GoTo 14
                    If FECHENT% = FECHANUM("20/09/05") Then TICAM# = 2.911: GoTo 14
                    If FECHENT% = FECHANUM("21/09/05") Then TICAM# = 2.919: GoTo 14
                    If FECHENT% = FECHANUM("22/09/05") Then TICAM# = 2.914: GoTo 14
                    If FECHENT% = FECHANUM("23/09/05") Then TICAM# = 2.914: GoTo 14
                    If FECHENT% = FECHANUM("26/09/05") Then TICAM# = 2.91: GoTo 14
                    If FECHENT% = FECHANUM("27/09/05") Then TICAM# = 2.91: GoTo 14
                    If FECHENT% = FECHANUM("28/09/05") Then TICAM# = 2.915: GoTo 14
                    If FECHENT% = FECHANUM("29/09/05") Then TICAM# = 2.915: GoTo 14
                    If FECHENT% = FECHANUM("30/09/05") Then TICAM# = 2.91: GoTo 14
                    If FECHENT% = FECHANUM("04/10/05") Then TICAM# = 2.909: GoTo 14
                    If FECHENT% = FECHANUM("06/10/05") Then TICAM# = 2.909: GoTo 14
                    If FECHENT% = FECHANUM("07/10/05") Then TICAM# = 2.909: GoTo 14
                    If FECHENT% = FECHANUM("12/10/05") Then TICAM# = 2.909: GoTo 14
                    If FECHENT% = FECHANUM("19/10/05") Then TICAM# = 2.909: GoTo 14
                    '
                    TICAM# = XVALO(InputBox$("Actualizar Cotización Dolar", "Tipo de Cambio al " + FECHATEX$(FECHENT%)))
'14                 Movisto.FindFirst "INT(CDBL(FECHMOV)) = " & Int(CDbl(CDate(CVDAT(FECHENT%))))
'16                   If Movisto.NoMatch = False Then
'                      Movisto.Edit
14                  SQLX$ = " SELECT * FROM MOVISTO"
                    SQLX$ = " WHERE FECHMOV = '" & FETEXCOR1$(FECHENT%) & "'"
                    Dim MOVVVISTO As ADODB.Recordset
                    Set MOVVVISTO = New ADODB.Recordset
28                  On Error Resume Next
                    MOVVVISTO.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
                    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                       On Error GoTo 0
                       Call CapturaErrorOpen
                       GoTo 28
                    End If
                    On Error GoTo 0
                    With MOVVVISTO
                      On Error Resume Next
                      If Not (.EOF And .BOF) Then
                        On Error GoTo 0
                        Do While (Not .EOF)
                          Movisto!TIPOCAM = TICAM#
                          Movisto.Update
  '                       Movisto.FindNext "INT(CDBL(FECHMOV)) = " & Int(CDbl(CDate(CVDAT(FECHENT%))))
  '                       GoTo 16
                          .MoveNext
                        Loop
                      End If
                    End With
                  End If
                  If MONERE% <= 1 Then
                    VUNI# = VUNI# / TICAM#
                  End If
                Call REPLA(Y$, MKD$(VUNI#), 174, 8)
                Call REPLA(Y$, MKD$(!Cantsalid), 182, 8)
                  CANACU# = CANACU# + !Cantsalid
                Call REPLA(Y$, !DoPriCor, 190, 16)
                Call REPLA(Y$, !DoSecCor, 206, 16)
                Call REPLA(Y$, MKD$(TICAM#), 222, 8)
                ' SI MONEDA ES PESOS (0 - 1), GUARDA PESOS Y CALCULA DOLARES
                If MONERE% <= 1 Then
                    NETPES# = !Cantsalid * !ValoUnit
                    IMPDOL# = NETPES# / TICAM#
                  Else
                    IMPDOL# = !Cantsalid * !ValoUnit
                    NETPES# = IMPDOL# * TICAM#
                End If
                Call REPLA(Y$, MKD$(IMPDOL#), 230, 8)
                Call REPLA(Y$, MKD$(NETPES#), 238, 8)
                GIVA% = GRUPIVA%(SCP_COD(U&))
                  TASA# = CVS(Mid$(FILTERGETRECORD("GRUPIVA", 1, MKI$(GIVA%)), 13, 4))
                  IMPIVA# = (NETPES# * TASA#) / 100
                Call REPLA(Y$, MKD$(IMPIVA#), 246, 8)
                  IMPTOT# = NETPES# + IMPIVA#
                Call REPLA(Y$, MKD$(IMPTOT#), 254, 8)
' falta revisar de aqui en adelante
                  ' SI MONEDA ES PESOS (0 - 1), GUARDA PESOS
                  ' Y CALCULA DOLARES
                  ANTIPES# = NETPES# * COEFPES#(U&)
                  ANTIDOL# = IMPDOL# * COEFDOL#(U&)
                Call REPLA(Y$, MKD$(ANTIDOL#), 262, 8)
                Call REPLA(Y$, MKD$(ANTIPES#), 270, 8)
                '
                JJ& = JJ& + 1
                Call GRAREG("SECOMPED", Y$, JJ&)
                '
80            .MoveNext
              Loop
           End With
           '
' aqui poner algo para netear pedidos anulados
           If SCP_STAT%(U&) >= 8 Then
             If Abs(CANACU#) >= 0.05 Then
               XREG$ = REGLEIDO$("SECOMPED", SCP_REG&(U&))
               CANPEDOR# = CVD(Mid$(XREG$, 180, 8))
               Call REPLA(XREG$, MKD$((-1) * CANACU#), 180, 8)
               If Abs(CANPEDOR#) >= 0.05 Then
                   For CLX% = 228 To 252 Step 8
                     VALOPED# = (-1) * CANACU# * CVD(Mid$(XREG$, CLX%, 8)) / CANPEDOR#
                     Call REPLA(XREG$, MKD$(VALOPED#), CLX%, 8)
                   Next CLX%
                 Else
                   Call REPLA(XREG$, String$(32, 0), 230, 32)
               End If
               Call REPLA(XREG$, "Anul.", 190, 16)
               Call REPLA(XREG$, "Anul.", 206, 16)
               JJ& = JJ& + 1
               Call GRAREG("SECOMPED", XREG$, JJ&)
             End If
           End If
           '
           For U2& = U& + 1 To FIN&
             If SCP_CLI(U2&) = SCP_CLI(U&) Then
               If SCP_PED(U2&) = SCP_PED(U&) Then
                 If SCP_COD(U2&) = SCP_COD(U&) Then
                   SCP_CLI(U2&) = 0
                   SCP_PED(U2&) = 0
                   SCP_COD(U2&) = 0
                   SCP_REG&(U2&) = 0
                   SCP_STAT%(U2&) = 0
                 End If
               End If
             End If
           Next U2&
           '
89      Next U&
        '
        Call SETULTREG("SECOMPED", JJ&)
        Call FILESORT("SECOMPED", "", 10, 10, "ASC")
        Call FILESORT("SECOMPED", "", 9, 9, "ASC")
        Call FILESORT("SECOMPED", "", 7, 7, "ASC")
        Call FILESORT("SECOMPED", "", 5, 5, "ASC")
        Call FILESORT("SECOMPED", "", 1, 1, "ASC")
        '
        Call TRALOCAL("SECOMPED", "SECOMPED")
        Call BLANARCH("SECOMPED")
        '
        JJ& = 0
        NCLA% = 0
        For U& = 1 To ULTREG&("!SECOMPED")
          X$ = REGLEIDO$("!SECOMPED", U&)
          NCLI = CVS(Left$(X$, 4))
          If U& > 1 Then
            If NCLI <> NCLA Then
              JJ& = JJ& + 1
              Call GRAREG("SECOMPED", String$(512, "="), JJ&)
            End If
          End If
          NCLA = NCLI
          RASO$ = TRIM$(Str$(NCLI)) + " - " + rasocli$(NCLI)
          DOMI$ = DOMICLI$(NCLI)
          LOCA$ = LOCACLI$(NCLI)
          LIPE$ = X$
          Call REPLA(LIPE$, RASO$, 5, 48)
          Call REPLA(LIPE$, DOMI$, 53, 48)
          Call REPLA(LIPE$, LOCA$, 101, 48)
          JJ& = JJ& + 1
          Call GRAREG("SECOMPED", LIPE$, JJ&)
        Next U&
        '
        Call SETULTREG("SECOMPED", JJ&)
        '
        Call HEADERS("SECOMPED", "")
        '
        Call FINAL("*SECOMPED")
        '
    End If
    '\\\OT-5-686-FG-FIN///
99  Screen.MousePointer = 1
    Command62.Enabled = True
End Sub


 Sub Command63_Click()
    Command63.Enabled = False
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    MODOGRAMY% = 0
    If Control$("ADMIPED", "FORMALIS") = "MONEMORI" Then
      If EXISTE%(LUDAT$ + "LIPEDID.RFS") > 0 Then
        MODOGRAMY% = 1
      End If
    End If
    '
    '\\\OT-05-0249-WAR-04/05/05///
    If MODOGRAMY% < 1 Then Call MENSPED
    '\\\OT-05-0249-WAR-FIN///
    If Option1.Value = True Then
        'Por Pantalla
        Call RECAPEDPE
        sali$ = "*COPEDNO"
        If FORMAPED% = 2 Then sali$ = "*COPEDMA"
        Call FINAL(sali$)
      Else
        'Listados
        ESJYM% = 0
        If InStr(EPAC$, "J Y M") Then ESJYM% = 1
        Screen.MousePointer = 11
        JJ& = 0: LL& = 0
        MODO% = COMALTER%("Clasificación por Fecha de Entrega Solicitada:\\Mes en Curso\Posteriores\Total")
        If MODO% < 1 Then
          GoTo 99
         ElseIf MODO% = 1 Then
          FLI1% = -1: FLI2% = FINMESACT%
        '
         ElseIf MODO% = 2 Then
10              FNMA = FINMESACT%
             FLI1% = DIPOST(FNMA): FLI2% = 32767
             VDEF$ = MKI$(FLI1%) + MKI$(FLI2%)
             OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
             If OBX$ = "" Then GoTo 99
             '
             FLI1% = CVI(Left$(OBX$, 2))
             FLI2% = CVI(Mid$(OBX$, 3, 2))
             If FLI1% > FLI2% Then GoTo 10
             FEX = FLI1%: PERIO$ = FECHATEX$(FEX)
             FEX = FLI2%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
         ElseIf MODO% = 3 Or MODO% = 4 Then
             FLI1% = -1: FLI2% = 32767
        End If
        '
        OPX% = ALTERNA%("Solo Cantidades\Cantidades y Precios")
        If OPX% < 1 Or OPX% > 2 Then GoTo 99
        '
        Call SELRACLIENUE(NCL1, NCL2, VENDSEL%)
        If NCL1 < 1 Then GoTo 99
        '
        If ESJYM% = 1 And OPX% = 2 Then
           Call MOSTRARREPORTE_JYM(NCL1, NCL2, VENDSEL%, FLI1%, FLI2%)
           GoTo 99
        End If
        'nuevo
        Call ABREPEDCLI
        SQLX$ = " SELECT * "
        SQLX$ = SQLX$ + " FROM PEDDETA "
        SQLX$ = SQLX$ + " WHERE NroClie >= " & NCL1 & " "
        SQLX$ = SQLX$ + " AND  NroClie <= " & NCL2 & " "
        SQLX$ = SQLX$ + " AND Status < 8 "
        If MODOGRAMY% < 1 Then SQLX$ = SQLX$ + " AND STATUS >= 0 "
        'SQLX$ = SQLX$ + " AND (((CanPed - CantEnt) >= 0.05) OR (TipoPed = 'ABIERTO' AND CDBL(FinVigen) >= " & Int(CDbl(Now)) & ")) "
        SQLX$ = SQLX$ + " AND (((CanPed - CantEnt) >= 0.05) "
        SQLX$ = SQLX$ + " OR (TipoPed = 'ABIERTO' AND FinVigen >= GETDATE() )) "

        '
        'Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
        '
        'jandy sql
        Dim PEDTEMP As ADODB.Recordset
        Set PEDTEMP = New ADODB.Recordset
        PEDTEMP.CursorLocation = adUseClient
        PEDTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText

        With PEDTEMP
          On Error Resume Next
          .Move first
          '
          JJ& = 0
          '
          Do While (.EOF = False)
            Y$ = REGBLAN("LIPEDPE")
            CI% = !CODIINT
            If CI% > 0 And CI% <= NUMAS% Then
              CAPEN = !CanPed - !cantent
              If CAPEN < 0 Then CAPEN = 0
              FENTRESOL% = CVINT(!FeSolEnt)
              If FENTRESOL% >= FLI1% Then
                If FENTRESOL% <= FLI2% Then
                  NCLIE = XVALO(!NROCLIE)
                  If VENDSEL% < 1 Or VENDCLI%(NCLIE) = VENDSEL% Then
                    Y$ = String$(46, 0) + Left$(X$, 6) + Mid$(X$, 9, 24) + Mid$(X$, 75)
                    Call REPLA(Y$, Mid$(X$, 7, 2), 1, 2)
                    Call REPLA(Y$, Mid$(X$, 7, 2), 45, 2)
                    Call REPLA(Y$, Mid$(X$, 33, 42), 3, 42)
                    '
                    CANOASI = CAPEN - !CanAsig
                    If CANOASI < 0 Then CANOASI = 0
                    NCLIE = !NROCLIE
                      MONEMISPED% = !MoneEmis
                    PREUNI = !PreUnid * (100 - PORDESCLI(NCLIE)) / 100 * TICAMONE(MONEMISPED%)
                    If MODOGRAMY% = 1 Then
                       PREUNI = PREUNI / TICAMONE(2)
                    End If
                    '
                    If OPX% <> 2 Then PREUNI = 0
                    ITOTA# = CDbl(CAPEN * PREUNI)
                    INOAS# = CDbl(CANOASI * PREUNI)
                    '
                    Call REPLA(Y$, MKS$(NCLIE), 1, 4)
                    Call REPLA(Y$, MKI$(!CODIINT), 5, 2)
                    Call REPLA(Y$, MKS$(NCLIE), 47, 4)
                    Call REPLA(Y$, MKS$(!NROPED), 51, 4)
                    Call REPLA(Y$, MKI$(FENTRESOL%), 73, 2)
                    Call REPLA(Y$, MKS$(!CanPed), 83, 4)
                    Call REPLA(Y$, MKS$(CAPEN), 103, 4)
                    Call REPLA(Y$, MKS$(PREUNI), 87, 4)
                    Call REPLA(Y$, MKD$(ITOTA#), 91, 8)
                    Call REPLA(Y$, MKS$(!cantent), 99, 4)
                    Call REPLA(Y$, MKS$(CANOASI), 113, 4)
                    Call REPLA(Y$, MKD$(INOAS#), 121, 8)
                    Call REPLA(Y$, Space$(34), 129, 34)
                    JJ& = JJ& + 1
                    Call GRAREG("LIPEDPE", Y$, JJ&)
                    '
                    If MODOGRAMY% = 1 Then
                       Z$ = REBLZ$
                       Y0$ = Left$(Y$, 46) ' LOS PRIMEROS BYTES COINCIDEN
                       Call REPLA(Z$, Y0$, 1, 46)
                         NOCO$ = !NROOCOM
                       Call REPLA(Z$, NOCO$, 47, 16)
                       Call REPLA(Z$, MKS$(!NROPED), 63, 4)
                       Call REPLA(Z$, MKI$(FEMIPED%), 67, 2)
                       Call REPLA(Z$, MKI$(FENTRESOL%), 69, 2)
                       If !Status < 0 Then
                          Call REPLA(Z$, "P", 71, 1)
                       End If
                       Call REPLA(Z$, MKS$(!CanPed), 81, 4)
                       Call REPLA(Z$, MKS$(!cantent), 85, 4)
                       Call REPLA(Z$, MKS$(CAPEN), 89, 4)
                       Call REPLA(Z$, MKS$(PREUNI), 97, 4)
                       Call REPLA(Z$, MKD$(ITOTA#), 101, 8)
                       Call GRAREG("LIPEDID", Z$, JJ&)
                    End If
                  End If
                End If
              End If
            End If
          .MoveNext
          Loop
          On Error GoTo 0
        End With
        PEDTEMP.Close
        Set PEDTEMP = Nothing
        
        '
30      If MODOGRAMY% <> 1 Then
             Call HEADERS("LIPEDPE", "")
             If MODO% = 2 Then
                Call HEADERS("LIPEDPE", "Entrega Solicitada: " + PERIO$)
             End If
             '
             Call SETULTREG("LIPEDPE", JJ&)
             Call CIERRARCH("LIPEDPE")
             Screen.MousePointer = 1
             '
             Call FILESORT("LIPEDPE", "", 13, 13, "ASC")
             sali$ = "*LIPENNO"
             Call FINAL(sali$)
           Else
             Call HEADERS("LIPEDID", "")
             TTXX$ = ""
             If MODO% = 2 Then
                Call HEADERS("LIPEDID", "Entrega Solicitada: " + PERIO$)
             End If
             '
             Call SETULTREG("LIPEDID", JJ&)
             Call CIERRARCH("LIPEDID")
             Screen.MousePointer = 1
             '
             Call FILESORT("LIPEDID", "", 14, 14, "ASC") ' POR NUMERO DE PEDIDO
             sali$ = "*LIPGRNO"
             Call FINAL(sali$)
        End If
    End If
    '
99  Screen.MousePointer = 1
    Command63.Enabled = True
End Sub

Sub MOSTRARREPORTE_JYM(NCL1, NCL2, VENDSEL%, FLI1%, FLI2%)
Campos$ = "Cte/F7"
Campos$ = Campos$ + ";Cliente/A32"
Campos$ = Campos$ + ";CodigoA32"
Campos$ = Campos$ + ";Descripcion/A32"
Campos$ = Campos$ + ";NoPed/f8"
Campos$ = Campos$ + ";Entreg./D8"
Campos$ = Campos$ + ";Cant/F12.1"
Campos$ = Campos$ + ";Moneda/F8"
Campos$ = Campos$ + ";Pr.Unit. $/F14.4"
Campos$ = Campos$ + ";Pr.Unit. U$s/F14.4"
Campos$ = Campos$ + ";Importe $/F12.2"
Campos$ = Campos$ + ";Importe U$s/F12.2"
Campos$ = Campos$ + ";Entreg/F12.1"
Campos$ = Campos$ + ";Saldo/F12.1"
Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)

        Call ABREPEDCLI
        SQLX$ = " SELECT * "
        SQLX$ = SQLX$ + " FROM PEDDETA "
        SQLX$ = SQLX$ + " WHERE NroClie >= " & NCL1 & " "
        SQLX$ = SQLX$ + " AND  NroClie <= " & NCL2 & " "
        SQLX$ = SQLX$ + " AND Status < 8 "
        SQLX$ = SQLX$ + " AND (((CanPed - CantEnt) >= 0.05) "
        SQLX$ = SQLX$ + " OR (TipoPed = 'ABIERTO' AND FinVigen >= GETDATE() )) "
        SQLX$ = SQLX$ + " ORDER BY NROPED"
        Dim PEDTEMP As ADODB.Recordset
        Set PEDTEMP = New ADODB.Recordset
        PEDTEMP.CursorLocation = adUseClient
        PEDTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText

        With PEDTEMP
          On Error Resume Next
          .Move first
          '
          JJ& = 0
          '
          Do While (.EOF = False)
            Y$ = REGBLAN("LIPEDPE")
            CI% = !CODIINT
            If CI% > 0 And CI% <= NUMAS% Then
              CAPEN = XVALO(!CanPed) - XVALO(!cantent)
              If CAPEN < 0 Then CAPEN = 0
              FENTRESOL% = CVINT(!FeSolEnt)
              If FENTRESOL% >= FLI1% Then
                If FENTRESOL% <= FLI2% Then
                  NCLIE = XVALO(!NROCLIE)
                  If VENDSEL% < 1 Or VENDCLI%(NCLIE) = VENDSEL% Then
                    JJ& = JJ& + 1
                    GRILLAGRAL.GRID.Rows = JJ& + 1
                    CANOASI = CAPEN - XVALO(!CanAsig)
                    If CANOASI < 0 Then CANOASI = 0
                    NCLIE = XVALO(!NROCLIE)
                    MONEMISPED% = XVALO(!MoneEmis)
                    PREUNI = XVALO(!PreUnid) * (100 - PORDESCLI(NCLIE)) / 100

                    If NROPED_Ant > 0 And XVALO(!NROPED) <> NROPED_Ant Then
                       ACUMULATOTALDOLAR = 0: ACUMULATOTALPESOS = 0
                    End If
                    GRILLAGRAL.GRID.TextMatrix(JJ&, 1) = NCLIE
                    GRILLAGRAL.GRID.TextMatrix(JJ&, 2) = SAFETEXT$(!RasoClie)
                    GRILLAGRAL.GRID.TextMatrix(JJ&, 3) = SAFETEXT$(!CODIEXT)
                    GRILLAGRAL.GRID.TextMatrix(JJ&, 4) = DESCRIT0$(XVALO(!CODIINT))
                    GRILLAGRAL.GRID.TextMatrix(JJ&, 5) = XVALO(!NROPED)
                    GRILLAGRAL.GRID.TextMatrix(JJ&, 6) = FECHATEX$(FENTRESOL1%)
                    GRILLAGRAL.GRID.TextMatrix(JJ&, 7) = FORMATNUM$(XVALO(!CanPed), "F12.1")
                    GRILLAGRAL.GRID.TextMatrix(JJ&, 8) = MONEMISPED%
                    IMPORTE = XVALO(PREUNI) * XVALO(!CanPed)
                    If MONEMISPED% = 1 Then
                       GRILLAGRAL.GRID.TextMatrix(JJ&, 9) = FORMATNUM$(XVALO(PREUNI), "F14.4")
                       GRILLAGRAL.GRID.TextMatrix(JJ&, 11) = FORMATNUM$(IMPORTE, "F12.2")
                       ACUMULATOTALPESOS = ACUMULATOTALPESOS + IMPORTE
                    ElseIf MONEMISPED% = 2 Then
                       GRILLAGRAL.GRID.TextMatrix(JJ&, 10) = FORMATNUM$(XVALO(PREUNI), "F14.4")
                       GRILLAGRAL.GRID.TextMatrix(JJ&, 12) = FORMATNUM$(IMPORTE, "F12.2")
                       ACUMULATOTALDOLAR = ACUMULATOTALDOLAR + IMPORTE
                    End If
                    GRILLAGRAL.GRID.TextMatrix(JJ&, 13) = FORMATNUM$(XVALO(!cantent), "F12.2")
                    SALDO = XVALO(!CanPed) - XVALO(!cantent): If SALDO < 0 Then SALDO = 0
                    GRILLAGRAL.GRID.TextMatrix(JJ&, 14) = FORMATNUM$(SALDO, "F12.2")
                    GRILLAGRAL.GRID.TextMatrix(JJ&, 15) = FORMATNUM$(ACUMULATOTALPESOS, "F12.2")
                    GRILLAGRAL.GRID.TextMatrix(JJ&, 16) = FORMATNUM$(ACUMULATOTALDOLAR, "F12.2")
                    NROPED_Ant = XVALO(!NROPED)
                    '

                  End If
                End If
              End If
            End If
          .MoveNext
          Loop
          On Error GoTo 0
        End With
        PEDTEMP.Close
        Set PEDTEMP = Nothing

        TX$ = "Pedidos Pendientes Moneda Original"
        GRILLAGRAL.Caption = TX$
        '*******************************************
        GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
        GRILLAGRAL.mnuMenu1.Visible = True
        GRILLAGRAL.GRID.ZOrder 0
        GRILLAGRAL.TXTBUSCAR = ""
    
        GRILLAGRAL.Show 1
        Dim MSF As msFlexGrid
        If GRILLAGRAL.MENU1 > 0 Then
            Set MSF = GRILLAGRAL.GRID
            If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
            Call CARGA_TABLA_IMPRE(MSF, Campos$, "PEPEMOOR", TX$, "TABLA_IMPRE", RET_NAMECONS$, 1)
            Call FINAL("?" & RET_NAMECONS$)
            GRILLAGRAL.MENU1 = 0
            Screen.MousePointer = 1
        End If
        Set obj = Nothing
    
        '
99      Screen.MousePointer = 1
        Command63.Enabled = True
End Sub
    
 Sub Command64_Click()
    Command64.Enabled = False
    '
    MODOGRAMY% = 0
    If Control$("ADMIPED", "FORMALIS") = "MONEMORI" Then
      If EXISTE%(LUDAT$ + "LIPEDID.RFS") > 0 Then
        MODOGRAMY% = 1
      End If
    End If
    '
    '\\\OT-05-0249-WAR-04/05/05///
    If MODOGRAMY% < 1 Then Call MENSPED
    '\\\OT-05-0249-WAR-FIN///
    If Option1.Value = True Then
           'Por Pantalla
7          CodigoArticulo$ = ListaSeleccionArticulos
           CI1% = CODINT%(CodigoArticulo$)
           If CI1% > 0 Then
             Call PEDIPENGRILLA(0, CI1%)
             GoTo 7
           End If

'           Call RECAPEDPE
'           sali$ = "*COPEDMA"
'           If FORMAPED% = 2 Then sali$ = "*COQEDMA"
'           Call FINAL(sali$)
        Else
           'Listados
           Screen.MousePointer = 11
           JJ& = 0: LL& = 0
           FIN& = ULTREG&("PEDDETA")
           MODO% = COMALTER%("Clasificación por Fecha de Entrega Solicitada:\\Mes en Curso\Posteriores\Total")
           If MODO% < 1 Then
             GoTo 99
            ElseIf MODO% = 1 Then
             FLI1% = -1: FLI2% = FINMESACT%
           '
            ElseIf MODO% = 2 Then
10              FNMA = FINMESACT%
                FLI1% = DIPOST(FNMA): FLI2% = 32767
                VDEF$ = MKI$(FLI1%) + MKI$(FLI2%)
                OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
                If OBX$ = "" Then GoTo 99
                '
                FLI1% = CVI(Left$(OBX$, 2))
                FLI2% = CVI(Mid$(OBX$, 3, 2))
                If FLI1% > FLI2% Then GoTo 10
                FEX = FLI1%: PERIO$ = FECHATEX$(FEX)
                FEX = FLI2%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
            ElseIf MODO% = 3 Or MODO% = 4 Then
                FLI1% = -1: FLI2% = 32767
           End If
           '
           OPX% = ALTERNA%("Solo Cantidades\Cantidades y Precios")
           If OPX% < 1 Or OPX% > 2 Then GoTo 99
           '
           Call SELRACLIENUE(NCL1, NCL2, VENDSEL%)
           If NCL1 < 1 Then GoTo 99
           '
           'nuevo
           Call ABREPEDCLI
           'jandy
           SQLX$ = " SELECT * FROM PEDDETA "
           SQLX$ = SQLX$ + " WHERE NroClie >= " & NCL1 & " "
           SQLX$ = SQLX$ + " AND  NroClie <= " & NCL2 & " "
           SQLX$ = SQLX$ + " AND Status < 8 "
           If MODOGRAMY% < 1 Then SQLX$ = SQLX$ + " AND STATUS >= 0 "
           'sqlx$ = sqlx$ + " AND (((CanPed - CantEnt) >= 0.05) OR (TipoPed = 'ABIERTO' AND CDBL(FinVigen) >= " & Int(CDbl(Now)) & ")) "
           SQLX$ = SQLX$ + " AND (((CanPed - CantEnt) >= 0.05) OR (TipoPed = 'ABIERTO' AND FinVigen >= getdate()))"
           '
           'Set PEDTEMP = Ventas.OpenRecordset(sqlx$, dbOpenSnapshot)
           Dim PEDTEMP As ADODB.Recordset
           Set PEDTEMP = New ADODB.Recordset
           PEDTEMP.CursorLocation = adUseClient
25         On Error Resume Next
           PEDTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
           If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
              On Error GoTo 0
              Call CapturaErrorOpen
              GoTo 25
            End If
            On Error GoTo 0
           '
           With PEDTEMP
             On Error Resume Next
             .Move first
             '
             JJ& = 0
             Y$ = REGBLAN("LIPEDPE")
             If MODOGRAMY% = 1 Then
                REBLZ$ = REGBLAN$("LIPEDID")
             End If
             '
             
             '
             Do While (.EOF = False)
                CI% = !CODIINT
                '20/03/07 (OT-07-0118)
                NC% = !NROCLIE
                '20/03/07 (FIN)
             If CI% > 0 And CI% <= NUMAS% Then
                CAPEN = !CanPed - !cantent
               '
               If CAPEN < 0 Then CAPEN = 0
               FEMIPED% = CVINT(!FECHEMIS)
               FENTRESOL% = CVINT(!FeSolEnt)
               NCLIE = XVALO(!NROCLIE)
               '
               If FENTRESOL% >= FLI1% Then
                 If FENTRESOL% <= FLI2% Then
                   If VENDSEL% < 1 Or VENDCLI%(NCLIE) = VENDSEL% Then
                     Y$ = String$(46, 0) + Left$(X$, 6) + Mid$(X$, 9, 24) + Mid$(X$, 75)
                     Call REPLA(Y$, Mid$(X$, 7, 2), 1, 2)
                     Call REPLA(Y$, Mid$(X$, 7, 2), 45, 2)
                     Call REPLA(Y$, Mid$(X$, 33, 42), 3, 42)
                     CAPEN = !CanPed - !cantent
                     '
                     If CAPEN < 0 Then CAPEN = 0
                       CANOASI = CAPEN - !CanAsig
                     '
                     If CANOASI < 0 Then CANOASI = 0
                     'NC% = !NroClie
                       MONEMISPED% = !MoneEmis
                     PREUNI = !PreUnid * (100 - PORDESCLI(NCLIE)) / 100 * TICAMONE(MONEMISPED%)
                     If MODOGRAMY% = 1 Then
                        PREUNI = PREUNI / TICAMONE(2)
                     End If
                     '
                     If OPX% <> 2 Then PREUNI = 0
                     ITOTA# = CDbl(CAPEN * PREUNI)
                     INOAS# = CDbl(CANOASI * PREUNI)
                     '
                     Call REPLA(Y$, MKS$(NCLIE), 1, 4)
                     Call REPLA(Y$, MKI$(!CODIINT), 5, 2)
                     Call REPLA(Y$, MKS$(NCLIE), 47, 4)
                     Call REPLA(Y$, MKS$(!NROPED), 51, 4)
                     Call REPLA(Y$, MKI$(FENTRESOL%), 73, 2)
                     Call REPLA(Y$, MKS$(!CanPed), 83, 4)
                     Call REPLA(Y$, MKS$(CAPEN), 103, 4)
                     Call REPLA(Y$, MKS$(PREUNI), 87, 4)
                     Call REPLA(Y$, MKD$(ITOTA#), 91, 8)
                     Call REPLA(Y$, MKS$(!cantent), 99, 4)
                     Call REPLA(Y$, MKS$(CANOASI), 113, 4)
                     Call REPLA(Y$, MKD$(INOAS#), 121, 8)
                     Call REPLA(Y$, Space$(34), 129, 34)
                     JJ& = JJ& + 1
                     Call GRAREG("LIPEDPE", Y$, JJ&)
                     '
                     If MODOGRAMY% = 1 Then
                        Z$ = REBLZ$
                        Y0$ = Left$(Y$, 46) ' LOS PRIMEROS BYTES COINCIDEN
                        Call REPLA(Z$, Y0$, 1, 46)
                          NOCO$ = !NROOCOM
                        Call REPLA(Z$, MKS$(NCLIE), 47, 4)
                        Call REPLA(Z$, NOCO$, 51, 16)
                        Call REPLA(Z$, MKS$(!NROPED), 67, 4)
                        Call REPLA(Z$, MKI$(FEMIPED%), 71, 2)
                        Call REPLA(Z$, MKI$(FENTRESOL%), 73, 2)
                        If !Status < 0 Then
                           Call REPLA(Z$, "P", 75, 1)
                        End If
                        Call REPLA(Z$, MKS$(!CanPed), 83, 4)
                        Call REPLA(Z$, MKS$(!cantent), 87, 4)
                        Call REPLA(Z$, MKS$(CAPEN), 91, 4)
                        Call REPLA(Z$, MKS$(PREUNI), 99, 4)
                        Call REPLA(Z$, MKD$(ITOTA#), 103, 8)
                        Call GRAREG("LIPEDID", Z$, JJ&)
                     End If
                   End If
                 End If
               End If
             End If
             .MoveNext
             Loop
             On Error GoTo 0
           End With
           PEDTEMP.Close
           Set PEDTEMP = Nothing
           
30         If MODOGRAMY% <> 1 Then
               Call HEADERS("LIPEDPE", "")
               If MODO% = 2 Then
                  Call HEADERS("LIPEDPE", "Entrega Solicitada: " + PERIO$)
               End If
               '
               Call SETULTREG("LIPEDPE", JJ&)
               Call CIERRARCH("LIPEDPE")
               Screen.MousePointer = 1
               '
               '\\\OT-05-0734-WAR-11/11/05///
               Call FILESORT("LIPEDPE", "", 16, 16, "ASC") ' POR F.SOLIC.ENTREGA
               '\\\OT-05-0734-WAR-FIN///
               Call FILESORT("LIPEDPE", "", 3, 9, "ASC")   ' POR CODIGO DE ARTICULO
               sali$ = "*LIPENMA"
               Call FINAL(sali$)
            Else
               Call HEADERS("LIPEDID", "")
               TTXX$ = ""
               If MODO% = 2 Then
                  Call HEADERS("LIPEDID", "Entrega Solicitada: " + PERIO$)
               End If
               '
               Call SETULTREG("LIPEDID", JJ&)
               Call CIERRARCH("LIPEDID")
               Screen.MousePointer = 1
               '
               '\\\OT-05-0734-WAR-11/11/05///
               Call FILESORT("LIPEDID", "", 16, 16, "ASC") ' POR F.SOLIC.ENTREGA
               '\\\OT-05-0734-WAR-FIN///
               Call FILESORT("LIPEDID", "", 3, 9, "ASC")   ' POR CODIGO DE ARTICULO
               sali$ = "*LIPGRMA"
               Call FINAL(sali$)
           End If
    End If
99  Screen.MousePointer = 1
    Command64.Enabled = True
End Sub

Sub Command8_Click()
    Command8.Enabled = False
    '
10  TTXX$ = "Ranking de Pedidos\Rastreabilidad Facturas/Pedidos"
    OPMAX% = 2
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "DAKOT") > 0 Then
       TTXX$ = TTXX$ + "\Listados Clasificados"
       OPMAX% = 3
    End If
    OPX% = ALTERNA%(TTXX$)
    '
    If OPX% = 1 Then
        Call RANKPED
        GoTo 10
      ElseIf OPX% = 2 Then
        OPX2% = ALTERNA%("De Factura a Pedido \ De Pedido a Factura")
        Select Case OPX2%
           Case 1
            Call FILARTCLI
           Case 2
            Call PEDAFACT
           Case Else
            GoTo 10
        End Select
      ElseIf OPX% = OPMAX% Then
        Call LISCLASI
        GoTo 10
    End If
    Command8.Enabled = True
    '
End Sub

Sub LISCLASI()
    OPX% = ALTERNA%("Agrupados Número de Pedido\Agrupados por Color Base\Agrupados por Base\Agrupados por Base y Tela")
    If OPX% >= 1 And OPX% <= 4 Then
        Call MODELO(OPX% + 1)
    End If
End Sub
'
Sub RANKPED()
    Dim TOTVENTA#(32767)
    Dim TOTCANT#(32767)
    '
    SELVENDE.BackColor = &HC0FFC0
    SELVENDE.Frame6.BackColor = &HC0FFC0
10  SELVENDE.Show 1
    If SELVENDE.Label2 <> "OK" Then GoTo 99
    '
    opt% = COMALTER%("Opciones de Formato de Salida\\Ranking de Pedidos\Ranking de Margen Bruto")
    If opt% = 0 Then GoTo 99
    '
    Screen.MousePointer = 11
    CANSELE% = XVALO(SELVENDE.Label5)
    Des% = FECHANUM(SELVENDE.Label7)
    Has% = FECHANUM(SELVENDE.Label3)
    FEX = Des%: PERIO$ = FECHATEX$(FEX)
    FEX = Has%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    DESDAT = Int(CDbl(CVDAT(Des%)))
    HASDAT = Int(CDbl(CVDAT(Has%)))
    '
    If CANSELE% < 1 Then
       VENDSEL% = 0
       VENDDSC$ = "Todos Juntos"
       GoTo 15
    End If
    '
    For IVENDI% = 1 To SELVENDE.LICOPEN.ListCount
       If SELVENDE.LICOPEN.Selected(IVENDI% - 1) = True Then
          VENDSEL% = XVALO(Left$(SELVENDE.LICOPEN.List(IVENDI% - 1), 3))
          VENDDSC$ = TRIM$(Mid$(SELVENDE.LICOPEN.List(IVENDI% - 1), 11))
          '
15        J& = 0: TOTAGEN# = 0
          For U& = 0 To 32767
            TOTVENTA#(U&) = 0
            TOTCANT#(U&) = 0
          Next U&
          '
          Call ABREPEDCLI
          'jandy
          Desde = FETEXCOR1$(Des%)
          Hasta = FETEXCOR1$(Has%)
          SQLX$ = "SELECT * FROM PEDDETA "
          'sqlx$ = sqlx$ + "WHERE INT(CDBL(FechEmis)) >= " & DESDAT & " "
          'sqlx$ = sqlx$ + "AND INT(CDBL(FechEmis)) <= " & HASDAT & " "
          '
          SQLX$ = SQLX$ + "WHERE FechEmis >= '" & Desde & "' "
          SQLX$ = SQLX$ + "AND FechEmis <= '" & Hasta & "' "
          
          'On Error GoTo 0
          'Set PEDDETTEMP = Ventas.OpenRecordset(sqlx$, dbOpenSnapshot)
          '
          Dim PEDDETTEMP As ADODB.Recordset
          Set PEDDETTEMP = New ADODB.Recordset
          PEDDETTEMP.CursorLocation = adUseClient
25        On Error Resume Next
          PEDDETTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
          If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
             On Error GoTo 0
             Call CapturaErrorOpen
             GoTo 25
          End If
          On Error GoTo 0
          With PEDDETTEMP
             On Error Resume Next
             .MoveLast
             If Err > 0 Then
                On Error GoTo 0
                Call MENSERR(24, "Sin Movimientos en el Período Elegido")
                GoTo 99
             End If
             On Error GoTo 0
             FIN& = .RecordCount
             .MoveFirst
             Do While Not .EOF
                NCLIE = !NROCLIE
                'TOMO EL VENDEDOR INGRESADO EN EL PEDIDO
                NUPED& = XVALO(!NROPED)
                CONDIPE$ = "NROPED = " & NUPED& & ""
                VENPED% = XVALO(VALOBADA("PEDENCA", "VENDED", CONDIPE$, "NOMESS"))
                If VENPED% > 0 Then
                  If VENDSEL% = 0 Or VENDSEL% = VENPED% Then
                     GoTo 50
                  End If
                ElseIf VENDSEL% = 0 Or VENDSEL% = VENDCLI%(NCLIE) Then
50                NART% = !CODIINT
                  If NART% < 0 Then NART% = 0
                  If NART% > ARTMAX% Then ARTMAX% = NART%
                  TOTCANT#(NART%) = TOTCANT#(NART%) + !CanPed
                  NETO# = !CanPed * !PreUnid
                  If !MoneEmis > 1 Then
                     NETO# = NETO# * TICAMONE(!MoneEmis)
                  End If
                  NETO# = ROUND(NETO#, 2)
                  TOTVENTA#(NART%) = TOTVENTA#(NART%) + NETO#
                End If
             .MoveNext
             Loop
          End With
          PEDDETTEMP.Close
          Set PEDDETTEMP = Nothing
          '
          TOTAGEN# = 0
          For i% = 0 To ARTMAX%
            TOTAGEN# = TOTAGEN# + TOTVENTA#(i%)
          Next i%
          '
          List1.Clear
          For NCLIE = 0 To 32767
             If Abs(TOTVENTA#(NCLIE)) > 0.005 Then
                RELIB$ = "A" + CSTRING$(MKD$(TOTAGEN# - TOTVENTA#(NCLIE)), 4, 12, 2, 2)
                PORPAR = 100 * TOTVENTA#(NCLIE) / TOTAGEN#
                RELIB$ = RELIB$ + CSTRING$(MKS$(NCLIE), 3, 6, 0, 2)
                RELIB$ = RELIB$ + CSTRING$(MKS$(PORPAR), 4, 8, 2, 2)
                RELIB$ = RELIB$ + CSTRING$(MKD$(TOTVENTA#(NCLIE)), 4, 12, 2, 2)
                List1.AddItem RELIB$
             End If
          Next NCLIE
          '
          J& = 0: PORACU = 0
          FIN& = List1.ListCount
          '
          For U& = 1 To FIN&
              List1.ListIndex = U& - 1
              NCLIE = XVALO(Mid$(List1.TEXT, 14, 6))
              TONEVEN# = XVALO(Mid$(List1.TEXT, 28, 12))
              PORPAR = XVALO(Mid$(List1.TEXT, 20, 8))
              PORACU = PORACU + PORPAR
              '
              W$ = REGBLAN$("RANMABRU")
              '
              J& = J& + 1: JI% = J&
              Call REPLA(W$, MKI$(JI%), 1, 2)
              Call REPLA(W$, MKI$(NCLIE), 3, 2)
              Call REPLA(W$, MKD$(TOTCANT#(NCLIE)), 5, 8)
              Call REPLA(W$, MKD$(TONEVEN#), 13, 8)
              CU# = COSPESOS(NC%)
              Call REPLA(W$, MKD$(CU#), 21, 8)
              CT# = CU# * TOTCANT#(NC%)
              Call REPLA(W$, MKD$(CT#), 29, 8)
              UB# = TONEVEN# - CT#
              Call REPLA(W$, MKD$(UB#), 37, 8)
              '\\\OT-07-0105-RG-14/03/07///
              Call REPLA(W$, MKS$(PORPAR), 45, 4)
              Call REPLA(W$, MKS$(PORACU), 49, 4)
              IMPIVA# = (TONEVEN# * PORCEIVA(CInt(NCLIE))) / 100
              Call REPLA(W$, MKD$(IMPIVA#), 53, 8)
              ITOVE# = TONEVEN# + IMPIVA#
              Call REPLA(W$, MKD$(ITOVE#), 61, 8)
              '\\\OT-07-0105-RG-FIN///
              Call GRAREG("RANMABRU", W$, J&)
              '
          Next U&
          '
          Call SETULTREG("RANMABRU", J&)
          Call CIERRARCH("RANMABRU")
          If CANSELE% > 1 Then
             If J& < 1 Then
                GoTo 90
             End If
          End If
          '
          Call HEADERS("RANMABRU", "")
          Call HEADERS("RANMABRU", "Periodo: " + PERIO$)
          '\\\OT-07-0010-FG-12/01/07///
          TEXVEN$ = "0 - Todos los Vendedores"
          If VENDSEL% <> 0 Then TEXVEN$ = Str(VENDSEL%) & " - " & VENDDSC$
          Call HEADERS("RANMABRU", "Vendedor: " + TEXVEN$)
          '\\\OT-07-0010-FG-FIN///
          '
          FORMAEM$ = ""
          If CANSELE% > 1 Then FORMAEM$ = "/LPT1"
          '
          If opt% = 1 Then
               Call FINAL("*RANPEDI" + FORMAEM$)
            ElseIf opt% = 2 Then
                If DERACCE%("VERCOSTO", "Visualizar Costos y Precios") > 1 Then
                  Call FINAL("*RANMABRU" + FORMAEM$)
                End If
          End If
          '
90        List1.Clear
       End If
    Next IVENDI%
    '
    Screen.MousePointer = 1
    If CANSELE% > 0 Then GoTo 10
    '
99  Screen.MousePointer = 1
    Call CIERRARCH("*.*")
    Command8.Enabled = True
    '
End Sub


 Sub Command9_Click()
    Dim PEDTEMP As ADODB.Recordset
    Command9.Enabled = False
    '
    MODOGRAMY% = 0
    If Control$("ADMIPED", "FORMALIS") = "MONEMORI" Then
      If EXISTE%(LUDAT$ + "LIPEDID.RFS") > 0 Then
        MODOGRAMY% = 1
      End If
    End If
    '
    '\\\OT-05-0249-WAR-04/05/05///
    If MODOGRAMY% < 1 Then Call MENSPED
    '\\\OT-05-0249-WAR-FIN///
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    ESJYM% = 0
    If InStr(EPAC$, "J Y M") Then
      ESJYM% = 1
      INICIALIZO% = 0 'HAY QUE INICIALIZARLO FUERA DEL BUCLE DONDE SE LLAMA A CARGAGRILLAGRAL
    Else
      Call MENSERR(24, "No Disponible por el Momento.")
      GoTo 99
    End If
    
5   'OPX% = ALTERNA%("Detalle Pedidos Pendientes\Seguimiento Comercial")
    OPX% = 1  'FUERZO DERTALLE DE PEDIDOS PENDIENTES
    Option2.Value = True ' IMPRESION
    '
    If OPX% = 1 Then
        If Option1.Value = True Then
           'Consultas por Pantalla
7          NC = ListaSeleccionClientes()
           Call PEDIPENGRILLA(NC, 0)
'           GoTo 7
           
'           Call RECAPEDPE
'           sali$ = "*COPEDCL"
'           If FORMAPED% = 2 Then sali$ = "*COQEDCL"
'           Call FINAL(sali$)
        Else
           'Listados
           Screen.MousePointer = 11
           JJ& = 0: LL& = 0
           '
           MODO% = COMALTER%("Clasificación por Fecha de Entrega Solicitada:\\Mes en Curso\Posteriores\Total")
           If MODO% < 1 Then
               GoTo 99
             ElseIf MODO% = 1 Then
               FLI1% = -1: FLI2% = FINMESACT%
             ElseIf MODO% = 2 Then
               FLI1% = FINMESACT% + 1: FLI2% = 32767
10             FNMA = FINMESACT%
               FLI1% = DIPOST(FNMA): FLI2% = 32767
               VDEF$ = MKI$(FLI1%) + MKI$(FLI2%)
               OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
               If OBX$ = "" Then GoTo 99
               '
               FLI1% = CVI(Left$(OBX$, 2))
               FLI2% = CVI(Mid$(OBX$, 3, 2))
               If FLI1% > FLI2% Then GoTo 10
               FEX = FLI1%: PERIO$ = FECHATEX$(FEX)
               FEX = FLI2%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
             ElseIf MODO% = 3 Or MODO% = 4 Then
               FLI1% = -1: FLI2% = 32767
           End If
           '
           'OPX% = ALTERNA%("Solo Cantidades\Cantidades y Precios")
           OPX% = 2 'FUERZO A CANTIDADES Y PRECIOS
           If OPX% < 1 Or OPX% > 2 Then GoTo 99
           '
           Call SELRACLIENUE(NCL1, NCL2, VENDSEL%)
           If NCL1 < 1 Then GoTo 99
           '
           Screen.MousePointer = 1
           If ULTREG&("GRUCOCLI") > 0 Then
                Call SELECHO("GRUCOCLI")
                GrupoCliente% = XVALO(VALACT1$("GRUCOCLI"))
           End If
           Screen.MousePointer = 11
           '
           FORMASA$ = ""
           If EXISTE%(LUDAT$ + "DETPEPEN.RFS") Then
             FORMASA$ = "ASEMBLI"
           End If
           '
           VECPEDI$ = ""
           'nuevo
           ' ver ernesto no se puede usar la Sub RECAPEDPE()??
           If ESJYM% > 0 Then
            Campos$ = "Cta/F6; Cliente/A32; Codigo/A16; Descripcion/A24;NoPed/F6;F.entreg/D8; Cant./F7.1; Pr.Unit/f10.2"
            Campos$ = Campos$ + ";Importe/F12.2;M.O.R./F8.2;Tot.M.O.R./F10.2;M.O./F8.2;Tot.M.O./F10.2; Entreg/F7.1; Saldo/F7.1"
           End If
           Call ABREPEDCLI
           SQLX$ = " SELECT * "
           SQLX$ = SQLX$ + " FROM PEDDETA "
           SQLX$ = SQLX$ + " WHERE NroClie >= " & NCL1 & " "
           SQLX$ = SQLX$ + " AND  NroClie <= " & NCL2 & " "
           SQLX$ = SQLX$ + " AND Status < 8 "
           If MODOGRAMY% < 1 Then SQLX$ = SQLX$ + " AND STATUS >= 0 "
           'sqlx$ = sqlx$ + " AND (((CanPed - CantEnt) >= 0.05) OR (TipoPed = 'ABIERTO' AND CDBL(FinVigen) >= " & Int(CDbl(Now)) & ")) "
           SQLX$ = SQLX$ + " AND (((CanPed - CantEnt) >= 0.05) OR (TipoPed = 'ABIERTO' AND FinVigen >= getdate()))"
           '
           'Set PEDTEMP = Ventas.OpenRecordset(sqlx$, dbOpenSnapshot)
           '
           Set PEDTEMP = New ADODB.Recordset
           PEDTEMP.CursorLocation = adUseClient
25         On Error Resume Next
           PEDTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
           If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
              On Error GoTo 0
              Call CapturaErrorOpen
              GoTo 25
           End If
           On Error GoTo 0
           With PEDTEMP
             On Error Resume Next
             .Move first
             '
             JJ& = 0
             Y$ = REGBLAN("LIPEDPE")
             '
             ACUMMOR = 0
             ACUMMO = 0
             Do While (.EOF = False)
                CI% = !CODIINT
             If CI% > 0 And CI% <= NUMAS% Then
                CAPEN = !CanPed - !cantent
              '
              If CAPEN < 0 Then CAPEN = 0
                FENTRESOL% = CVINT(!FeSolEnt)
                '
                If FENTRESOL% >= FLI1% Then
                  If FENTRESOL% <= FLI2% Then
                    NCLIE = !NROCLIE
                    If (VENDSEL% < 1 Or VENDCLI%(NCLIE) = VENDSEL%) And (GrupoCliente% < 1 Or GRUCOCLI%(NCLIE) = GrupoCliente%) Then
                          Call REPLA(Y$, Mid$(X$, 7, 2), 1, 2)
                          Call REPLA(Y$, Mid$(X$, 7, 2), 45, 2)
                          Call REPLA(Y$, Mid$(X$, 33, 42), 3, 42)
                          CAPEN = !CanPed - !cantent
                        '
                        If CAPEN < 0 Then CAPEN = 0
                          CANOASI = CAPEN - !CanAsig
                        '
                        If CANOASI < 0 Then CANOASI = 0
                          MONEMISPED% = !MoneEmis
                        PREUNI = !PreUnid * (100 - PORDESCLI(NCLIE)) / 100 * TICAMONE(MONEMISPED%)
                        '
                        If OPX% <> 2 Then PREUNI = 0
                        ITOTA# = CDbl(CAPEN * PREUNI)
                        INOAS# = CDbl(CANOASI * PREUNI)
                        '
                        Call REPLA(Y$, MKS$(NCLIE), 1, 4)
                        Call REPLA(Y$, MKI$(!CODIINT), 5, 2)
                        Call REPLA(Y$, !TCOLOR, 7, 16)
                        Call REPLA(Y$, !TTALLE, 23, 8)
                        Call REPLA(Y$, MKS$(!TLargo), 31, 4)
                        Call REPLA(Y$, MKS$(!TAncho), 35, 4)
                        Call REPLA(Y$, MKS$(!TEspesor), 39, 4)
                        Call REPLA(Y$, MKS$(!NROPED), 51, 4)
                        Call REPLA(Y$, MKI$(FENTRESOL%), 73, 2)
                        Call REPLA(Y$, MKS$(!CanPed), 83, 4)
                        Call REPLA(Y$, MKS$(CAPEN), 103, 4)
                        Call REPLA(Y$, MKS$(PREUNI), 87, 4)
                        Call REPLA(Y$, MKD$(ITOTA#), 91, 8)
                        Call REPLA(Y$, MKS$(!cantent), 99, 4)
                        Call REPLA(Y$, MKS$(CANOASI), 113, 4)
                        Call REPLA(Y$, MKD$(INOAS#), 121, 8)
                        Call REPLA(Y$, !NROOCOM, 129, 16)
                        Call REPLA(Y$, Space$(18), 145, 18)
                        '
                        CI1% = XVALO(!CODIINT)
                        If FORMASA$ = "ASEMBLI" Then
                          STAX% = !Status: If STAX% < 0 Then STAX% = STAX% + 256
                          Call REPLA(Y$, Chr$(STAX%), 159, 1)
                          Call REPLA(Y$, Chr$(MONEMISPED%), 160, 1)
                        End If
                        '

                        JJ& = JJ& + 1
                        MOR = MANO_OBRA_REAL(CI1%)
                        If ESJYM% > 0 Then
                           CONDI$ = "CODICONJ =" & CI1% & " and NumeOper = 10"
                           MINU# = XVALO(VALOBADA("SECOPER", "MINUSTAN", CONDI$, "NOMESS"))
                           TEXTOGRILLA$ = NCLIE & "|" & rasocli$(NCLIE) & "|" & CODEXT$(CI1%)
                           TEXTOGRILLA$ = TEXTOGRILLA$ & "|" & DESCRIT0$(CI1%) & "|" & XVALO(!NROPED) & "|" & FECHATEX$(FENTRESOL%)
                           TEXTOGRILLA$ = TEXTOGRILLA$ & "|" & FORMATNUM$(XVALO(!CanPed), "F7.1") & "|" & FORMATNUM$(PREUNI, "F11.2") & "|" & FORMATNUM$(ITOTA#, "F11.2")
                           TEXTOGRILLA$ = TEXTOGRILLA$ & "|" _
                           & FORMATNUM$(MOR, "F10.2") & "|" _
                           & FORMATNUM$(MOR * XVALO(!CanPed), "F10.2") & "|" _
                           & FORMATNUM$(MINU#, "F10.2") & "|" _
                           & FORMATNUM$(MINU# * XVALO(!CanPed), "F10.2") & "|" _
                           & XVALO(!cantent) & "|" & XVALO(CANOASI)
    
                           Call CARGA_GRILLA_GENERICO(GRILLAGRAL, GRILLAGRAL.GRID, Campos$, TEXTOGRILLA$, INICIALIZO%)
                           
                        Else
                           Call GRAREG("LIPEDPE", Y$, JJ&)
                        End If
                        
                        '
                    End If
                  End If
                End If
             End If
             .MoveNext
             Loop
             On Error GoTo 0
           End With
           PEDTEMP.Close
           Set PEDTEMP = Nothing
           '.
            TX$ = "Pedidos Pendientes "
              If VENDSEL% > 0 Then
               TX$ = "Vendedor: " + ECOARCH$("VENDEDOR", Chr$(VENDSEL%))
              End If
              If GrupoCliente% > 0 Then
               TX$ = TX$ + " - Grupo Cliente: " + ECOARCH$("GRUCOCLI", CStr(GrupoCliente%))
           End If
           If ESJYM% > 0 Then
'                OPX% = ALTERNA%("Emitir Reporte Tradicional\Emitir Reporte Con M.Obra.")
'                If OPX% < 1 Or OPX% > 2 Then GoTo 99
'                If OPX% = 1 Then GoTo 30
                ACUMIMPORTE = SUMACOLUMNA#(GRILLAGRAL.GRID, 9, 1)
                ACUMMOR = SUMACOLUMNA#(GRILLAGRAL.GRID, 11, 1)
                ACUMMO = SUMACOLUMNA#(GRILLAGRAL.GRID, 13, 1)
                GRILLAGRAL.GRID.Rows = GRILLAGRAL.GRID.Rows + 2
                JJ& = GRILLAGRAL.GRID.Rows - 1
                GRILLAGRAL.GRID.TextMatrix(JJ&, 9) = FORMATNUM$(ACUMIMPORTE, "F12.2")
                GRILLAGRAL.GRID.TextMatrix(JJ&, 11) = FORMATNUM$(ACUMMOR, "F12.2")
                GRILLAGRAL.GRID.TextMatrix(JJ&, 13) = FORMATNUM$(ACUMMO, "F12.2")
                Call COLOREAR_GRILLA_SOLAFILA(GRILLAGRAL.GRID, &HE0E0E0, JJ&)
                Call numerargrilla(GRILLAGRAL.GRID, 0, 0, 0, 1)
                GRILLAGRAL.GRID.ColWidth(0) = 800
                GRILLAGRAL.Caption = TX$
                '*******************************************
                GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
                GRILLAGRAL.mnuMenu1.Visible = True
                GRILLAGRAL.GRID.ZOrder 0
                GRILLAGRAL.TXTBUSCAR = ""
              
                Screen.MousePointer = 1
                GRILLAGRAL.Show 1
                Dim MSF As msFlexGrid
                If GRILLAGRAL.MENU1 > 0 Then
                    Set MSF = GRILLAGRAL.GRID
                    X = MSF.Rows
                    If GRILLAGRAL.TXTBUSCAR <> "" Then Set MSF = GRILLAGRAL.MSF_2
                    Call CARGA_TABLA_IMPRE(MSF, Campos$, "PEDPEJYM", TX$, "TABLA_IMPRE", RET_NAMECONS$, 1)
                    Call FINAL("?" & RET_NAMECONS$)
                    GRILLAGRAL.MENU1 = 0
                    Screen.MousePointer = 1
                End If
                GoTo 99
           End If
           
30         Call HEADERS("LIPEDPE", "")
           If MODO% = 2 Then
              Call HEADERS("LIPEDPE", "Entrega Solicitada: " + PERIO$)
           End If
           '
           Call SETULTREG("LIPEDPE", JJ&)
           Call CIERRARCH("LIPEDPE")
           '
           VARIANTE$ = "Totales por Entregar"
           Call HEADERS("LIPEDPE", "Corresponde a: " + VARIANTE$)
           If VENDSEL% > 0 Then
             Call HEADERS("LIPEDPE", "Vendedor: " + ECOARCH$("VENDEDOR", Chr$(VENDSEL%)))
           End If
           If GrupoCliente% > 0 Then
             Call HEADERS("LIPEDPE", "Grupo Cliente: " + ECOARCH$("GRUCOCLI", CStr(GrupoCliente%)))
           End If
           '
           If FORMASA$ = "ASEMBLI" Then
             Call FILESORT("LIPEDPE", "", 3, 3, "ASC") ' POR CODIGO
             Call FILESORT("LIPEDPE", "", 13, 13, "ASC") ' POR PEDIDO
             Call FILESORT("LIPEDPE", "", 1, 1, "ASC") ' POR CLIENTE
             '
             JJ& = 0
             For NCL% = NCL1 To NCL2
               RFF$ = RECFILT$("LIPEDPE", 1, MKS$(NCL%))
               If Len(RFF$) > 0 Then
                 REBLA$ = REGBLAN$("DETPEPEN")
                 Call REPLA(REBLA$, MKI$(NCL%), 1, 2)
                    RASO1$ = TRIM$(Str$(NCL%)) + " - " + rasocli$(NCL%)
                 Call REPLA(REBLA$, RASO1$, 3, 48)
                 Call REPLA(REBLA$, DOMICLI$(NCL%), 51, 48)
                 Call REPLA(REBLA$, LOCACLI$(NCL%), 99, 48)
                 '
                 For KL& = 1 To Len(RFF$) Step 4
                   RELI& = CVS(Mid$(RFF$, KL&, 4))
                   LIPE$ = REGLEIDO$("LIPEDPE", RELI&)
                   ZZ$ = REBLA$
                     '
                     NPED& = CVS(Mid$(LIPE$, 51, 4))
                   Call REPLA(ZZ$, MKS$(NPED&), 147, 4)
                     STA$ = "  "
                     If Asc(Mid$(LIPE$, 159, 1)) = 255 Then STA$ = "P."
                     If Asc(Mid$(LIPE$, 159, 1)) = 8 Then STA$ = "AN"
                     If Asc(Mid$(LIPE$, 159, 1)) = 9 Then STA$ = "AN"
                   Call REPLA(ZZ$, STA$, 151, 2)
                     NOCO$ = Mid$(LIPE$, 129, 16)
                   Call REPLA(ZZ$, NOCO$, 155, 16)
                     CIXI% = CVI(Mid$(LIPE$, 5, 2))
                   Call REPLA(ZZ$, MKI$(CIXI%), 171, 2)
                     FECHENT% = CVI(Mid$(LIPE$, 73, 2))
                   Call REPLA(ZZ$, MKI$(FECHENT%), 173, 2)
                     CANORI = CVS(Mid$(LIPE$, 83, 4))
                   Call REPLA(ZZ$, MKS$(CANORI), 175, 4)
                     CANENT = CVS(Mid$(LIPE$, 99, 4))
                   Call REPLA(ZZ$, MKS$(CANENT), 179, 4)
                     CANSAL = CVS(Mid$(LIPE$, 103, 4))
                   Call REPLA(ZZ$, MKS$(CANSAL), 183, 4)
                     MONEM% = Asc(Mid$(LIPE$, 160, 1))
                   Call REPLA(ZZ$, SIMBOLPES$(MONEM%), 187, 4)
                     PREUN# = CVS(Mid$(LIPE$, 87, 4)) / TICAMONE(MONEM%)
                     ' SI EL TIPO DE MONEDA ES 0 O 1 EL P.UNITARIO SE
                     ' GUARDA EN PESOS Y SE CALCULA EL P.UNITARIO EN DOLARES
                     If MONEM% < 2 Then
                       PREPES# = PREUN#
                       Call REPLA(ZZ$, MKD$(PREPES#), 191, 8)
                       On Error Resume Next
                       PREDOL# = PREUN# * TICAMONE(MONEM%) / TICAMONE(2)
                       On Error GoTo 0
                       Call REPLA(ZZ$, MKD$(PREDOL#), 199, 8)
                     End If
                     ' SI EL TIPO DE MONEDA ES 2 EL P.UNITARIO SE GUARDA EN
                     ' DOLARES Y SE CALCULA EL P. UNITARIO EN PESOS
                     If MONEM% = 2 Then
                       PREDOL# = PREUN#
                       Call REPLA(ZZ$, MKD$(PREDOL#), 199, 8)
                       On Error Resume Next
                       PREPES# = PREUN# * TICAMONE(2)
                       On Error GoTo 0
                       Call REPLA(ZZ$, MKD$(PREPES#), 191, 8)
                     End If
                   Call REPLA(ZZ$, MKD$(PREDOL# * CANSAL), 207, 8)
                   JJ& = JJ& + 1
                   Call GRAREG("DETPEPEN", ZZ$, JJ&)
                 Next KL&
                 '
                 JJ& = JJ& + 1
                 Call GRAREG("DETPEPEN", String$(256, "="), JJ&)
               End If
             Next NCL%
             Call SETULTREG("DETPEPEN", JJ&)
             '
             Call HEADERS("DETPEPEN", "")
             'Call HEADERS("DETPEPEN", "T.Cambio $ / U$S: " + TRIM$(FORMATNUM$(TICAMONE(2), "F12.4")))
             '
             Call FINAL("*DETPEPEN")
             GoTo 99
             '
           End If
           '
           Screen.MousePointer = 1
           '\\\OT-05-0248-WAR-04/05/05///
           Call FILESORT("LIPEDPE", "", 16, 16, "ASC")
           Call FILESORT("LIPEDPE", "", 2, 2, "ASC")
           Call FILESORT("LIPEDPE", "", 1, 1, "ASC")
           '\\\OT-05-0248-WAR-FIN///
           'Call FILESORT("LIPEDPE", "", 1, 9, "ASC")
           sali$ = "*LIPENCL"
           Call FINAL(sali$)
        End If
    ElseIf OPX% = 2 Then
        '\\\OT-5-686-FG-20/10/05///
        Dim SCP_CLI(4096)         ' VECTOR CLIENTES
        Dim SCP_PED&(4096)        ' VECTOR PEDIDOS
        Dim SCP_COD%(4096)        ' VECTOR CODIGOS
        Dim SCP_REG&(4096)        ' VECTOR REGISTRO PEDIDO
        Dim SCP_STAT%(4096)       ' VECTOR STATUS PEDIDO
        Dim COEFPES#(4096)   ' VECTOR TOTAL PEDIDO
        Dim COEFDOL#(4096)   ' VECTOR ANTICIPO EN DOLARES
        '
35      Call DESHASFECHA(Des%, Has%, PERIO$)
        If TRIM$(PERIO$) = "" Then GoTo 5
        '
        FEDES = CDbl(CVDAT(Des%))
        FEHAS = CDbl(CVDAT(Has%))
        '
        Call ABREPEDCLI
        Call ABRESTOCKS
        '
        '\\\OT-06-0098-WAR-20/03/06///
37      R1% = COMALTER%("Seleccione la Opción de Filtro\\Todos\Por Cliente")
        If R1% = 0 Then GoTo 35
        If R1% = 2 Then
          Call SELECHO("DEUDOR1")
          NCLIE = XVALO(TRIM$(VALACT1$("DEUDOR1")))
          If NCLIE <= 0 Then GoTo 37
        End If
        ''\\\OT-06-0098-WAR-FIN///
        Desde = FETEXCOR1$(Des%)
        Hasta = FETEXCOR1$(Has%)

40      SQLX$ = " SELECT * "
        SQLX$ = SQLX$ + " FROM PEDDETA "
        'sqlx$ = sqlx$ + " WHERE INT(CDBL(FechEmis)) >= " & FEDES & ""
        'sqlx$ = sqlx$ + " AND INT(CDBL(FechEmis)) <= " & FEHAS & ""
        SQLX$ = SQLX$ + " WHERE FechEmis >= '" & Desde & "'"
        SQLX$ = SQLX$ + " AND FechEmis <= '" & Hasta & "'"
        'Jandy ESTO YA ESTABA ASI
        'SQLX$ = SQLX$ + " AND Status < 8 "
        If MODOGRAMY% < 1 Then SQLX$ = SQLX$ + " AND STATUS >= 0 "
        '\\\OT-06-0098-WAR-20/03/06///
        If R1% = 2 Then
          If NCLIE > 0 Then
            SQLX$ = SQLX$ + "AND NroClie = " & NCLIE & ""
          End If
        End If
        '\\\OT-06-0098-WAR-FIN///
        '
        'Set PEDTEMP = Ventas.OpenRecordset(sqlx$, dbOpenSnapshot)
        '
        
        Set PEDTEMP = New ADODB.Recordset
        PEDTEMP.CursorLocation = adUseClient
26      On Error Resume Next
        PEDTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 26
        End If
        On Error GoTo 0
        With PEDTEMP
          On Error Resume Next
          .MoveFirst
          If Err > 1 Then
            Call COMUNI("No Existen Pedidos en el Período Seleccionado")
            GoTo 99
          End If
          '
          Call BLANARCH("SECOMPED")
          JJ& = 0
          REBLA$ = REGBLAN("SECOMPED")
          '
          Do While Not .EOF
            ' esto se agrega para considerar pedidos anulados pero con movimiento
            ' SEGúN PEDIDO DE ASEMBLI (OT-6-21)
            If !Status >= 8 Then    ' CASO DE PEDIDOS ANULADOS
               NUCLI7 = !NROCLIE
               NUPED7& = !NROPED
               CODIN7% = !CODIINT
               SQLY$ = " SELECT * FROM MOVISTO "
               SQLY$ = SQLY$ + " WHERE Nume_Clie = " & NUCLI7
               'SQLY$ = SQLY$ + " AND CDBL(NuPedCli) = " & NUPED7&
               'jandy sql
               SQLY$ = SQLY$ + " AND NuPedCli = " & NUPED7&
               SQLY$ = SQLY$ + " AND Cod_Inte = " & CODIN7%
               '
               'Set STOBUS = Stocks.OpenRecordset(SQLY$, dbOpenDynaset)
               Dim STOBUS As ADODB.Recordset
               Set STOBUS = FLEXCONN.Execute(FILTSQL$(SQLY$))

  'ver ernesto
  
               With STOBUS
                 On Error Resume Next
                 .MoveFirst
                 If Err > 1 Then
                   On Error GoTo 0
                   GoTo 49        ' NO TIENE REMITOS
                 End If
                 On Error GoTo 0
               End With
            End If
            ' ***************************************************************
          
            Y$ = REBLA$
            '
            Call REPLA(Y$, MKS$(!NROCLIE), 1, 4)
            Call REPLA(Y$, MKS$(!NROPED), 149, 4)
            Call REPLA(Y$, !NROOCOM, 153, 16)
            Call REPLA(Y$, MKI$(!CODIINT), 169, 2)
            Call REPLA(Y$, Chr$(1), 171, 1) ' MARCA DE QUE ES UN PEDIDO
              FECHENT% = CVINT%(!FeSolEnt)
            Call REPLA(Y$, MKI$(FECHENT%), 172, 2)
              MONEORIG% = !MoneEmis
              TICAM# = !TIPOCAM: If TICAM# < 0.01 Then TICAM# = 1
              PRUNI# = !PreUnid
              If MONEORIG% <= 1 Then
                PRUNI# = PRUNI# / TICAM#
              End If
            Call REPLA(Y$, MKD$(PRUNI#), 174, 8)  'precio unitario en dolares
              CANTI# = (-1) * !CanPed
            Call REPLA(Y$, MKD$(CANTI#), 182, 8)
            Call REPLA(Y$, String(32, " "), 190, 32)
            Call REPLA(Y$, MKD$(TICAM#), 222, 8)
             ' SI MONEDA ES PESOS (0 - 1), GUARDA PESOS Y CALCULA DOLARES
              If MONEORIG% <= 1 Then
                  NETPES# = (-1) * !ImpoTot
                  IMPDOL# = NETPES# / TICAM#
                Else
                  IMPDOL# = (-1) * !ImpoTot
                  NETPES# = IMPDOL# * TICAM#
              End If
            Call REPLA(Y$, MKD$(IMPDOL#), 230, 8)
            Call REPLA(Y$, MKD$(NETPES#), 238, 8)
              GIVA% = GRUPIVA%(!CODIINT)
              TASA# = CVS(Mid$(FILTERGETRECORD("GRUPIVA", 1, MKI$(GIVA%)), 13, 4))
              IMPIVA# = (NETPES# * TASA#) / 100
            Call REPLA(Y$, MKD$(IMPIVA#), 246, 8)
              IMPTOT# = NETPES# + IMPIVA#
            Call REPLA(Y$, MKD$(IMPTOT#), 254, 8)
              IANTI# = 0
              On Error Resume Next
              IANTI# = !ImpoAnti
              On Error GoTo 0
              ' SI MONEDA ES PESOS (0 - 1), GUARDA PESOS Y CALCULA DOLARES
              If MONEORIG% <= 1 Then
                  ANTIPES# = IANTI#
                  ANTIDOL# = IANTI# / TICAM#
                Else
                  ANTIDOL# = IANTI#
                  ANTIPES# = IANTI# * TICAM#
              End If
            Call REPLA(Y$, MKD$(ANTIPES#), 270, 8)
            Call REPLA(Y$, MKD$(ANTIDOL#), 268, 8)
            '
            JJ& = JJ& + 1
            Call GRAREG("SECOMPED", Y$, JJ&)
            '
            SCP_CLI(JJ&) = !NROCLIE
            SCP_PED(JJ&) = !NROPED
            SCP_COD(JJ&) = !CODIINT
            SCP_REG&(JJ&) = JJ&           ' VECTOR REGISTRO PEDIDO
            SCP_STAT%(JJ&) = !Status      ' VECTOR STATUS PEDIDO
            If Abs(IMPDOL#) > 0.1 Then
              COEFDOL#(JJ&) = ANTIDOL# / IMPDOL#
            End If
            If Abs(NETPES#) > 0.1 Then
               COEFPES#(JJ&) = ANTIPES# / NETPES#
            End If
            '
49        .MoveNext
          Loop
          '
        End With
        PEDTEMP.Close
        Set PEDTEMP = Nothing
        '
        
        FIN& = JJ&
        For U& = 1 To FIN&
           '
           If SCP_COD(U&) < 1 Then GoTo 89
           '
           XREG$ = REGLEIDO$("SECOMPED", SCP_REG&(U&))
           CANACU# = CVD(Mid$(XREG$, 180, 8))
           '
           
           SQLX$ = " SELECT * FROM MOVISTO "
           SQLX$ = SQLX$ + " WHERE Nume_Clie = " & SCP_CLI(U&)
           'SQLX$ = SQLX$ + " AND CDBL(NuPedCli) = " & SCP_PED(U&)
           SQLX$ = SQLX$ + " AND NuPedCli = " & SCP_PED(U&)
           SQLX$ = SQLX$ + " AND Cod_Inte = " & SCP_COD(U&)
           '
           'Set STOCKTEMP = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
           
           Dim STOCKTEMP As ADODB.Recordset
           Set STOCKTEMP = New ADODB.Recordset
27         On Error Resume Next
           STOCKTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
           If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
              On Error GoTo 0
              Call CapturaErrorOpen
              GoTo 27
           End If
           On Error GoTo 0
           With STOCKTEMP
              On Error Resume Next
              .MoveFirst
              If Err > 1 Then GoTo 80
              On Error GoTo 0
              '
              Do While Not .EOF
                Y$ = REBLA$
                '
                Call REPLA(Y$, MKS$(SCP_CLI(U&)), 1, 4)
                Call REPLA(Y$, MKS$(SCP_PED(U&)), 149, 4)
                'Call REPLA(Y$, String(16, " "), 151, 16)
                NPEDX1& = SCP_PED(U&)
                Call REPLA(Y$, NOCOMCLI$(NPEDX1&), 153, 16)
 
                Call REPLA(Y$, MKI$(SCP_COD(U&)), 169, 2)
                Call REPLA(Y$, Chr$(2), 169, 1) ' MARCA DE QUE ES UN REMITO
                  FECHENT% = CVINT%(!FechMov)
                Call REPLA(Y$, MKI$(FECHENT%), 172, 2)
                  VUNI# = !ValoUnit
                  MONERE% = 1: TICAM# = 0
                  On Error Resume Next
                  MONERE% = !MoneVal
                  TICAM# = !TIPOCAM
                  On Error GoTo 0
                  '
                  NROX$ = !DOSECLAR
                  FEX = FECHENT%
                  NCLIEN% = SCP_CLI(U&)
                  MONERE1% = MONEMISOR%(NROX$, FEX, NCLIEN%)
                  If MONERE1% <> MONERE% Then
                    '.Edit
                    !MoneVal = MONERE1%
                    .Update
                    MONERE% = MONERE1%
                  End If
                  '
                  If TICAM# < 0.01 Then
                    If FECHENT% = FECHANUM("21/03/05") Then TICAM# = 2.92: GoTo 14
                    If FECHENT% = FECHANUM("11/04/05") Then TICAM# = 2.917: GoTo 14
                    If FECHENT% = FECHANUM("25/04/05") Then TICAM# = 2.861: GoTo 14
                    If FECHENT% = FECHANUM("26/05/05") Then TICAM# = 2.886: GoTo 14
                    If FECHENT% = FECHANUM("08/06/05") Then TICAM# = 2.854: GoTo 14
                    If FECHENT% = FECHANUM("09/06/05") Then TICAM# = 2.895: GoTo 14
                    If FECHENT% = FECHANUM("17/06/05") Then TICAM# = 2.874: GoTo 14
                    If FECHENT% = FECHANUM("28/06/05") Then TICAM# = 2.849: GoTo 14
                    If FECHENT% = FECHANUM("29/06/05") Then TICAM# = 2.889: GoTo 14
                    If FECHENT% = FECHANUM("30/06/05") Then TICAM# = 2.855: GoTo 14
                    If FECHENT% = FECHANUM("13/07/05") Then TICAM# = 2.886: GoTo 14
                    If FECHENT% = FECHANUM("14/07/05") Then TICAM# = 2.869: GoTo 14
                    If FECHENT% = FECHANUM("15/07/05") Then TICAM# = 2.864: GoTo 14
                    If FECHENT% = FECHANUM("18/07/05") Then TICAM# = 2.863: GoTo 14
                    If FECHENT% = FECHANUM("19/07/05") Then TICAM# = 2.864: GoTo 14
                    If FECHENT% = FECHANUM("21/07/05") Then TICAM# = 2.862: GoTo 14
                    If FECHENT% = FECHANUM("22/07/05") Then TICAM# = 2.863: GoTo 14
                    If FECHENT% = FECHANUM("25/07/05") Then TICAM# = 2.863: GoTo 14
                    If FECHENT% = FECHANUM("26/07/05") Then TICAM# = 2.862: GoTo 14
                    If FECHENT% = FECHANUM("28/07/05") Then TICAM# = 2.865: GoTo 14
                    If FECHENT% = FECHANUM("29/07/05") Then TICAM# = 2.863: GoTo 14
                    If FECHENT% = FECHANUM("02/08/05") Then TICAM# = 2.86: GoTo 14
                    If FECHENT% = FECHANUM("03/08/05") Then TICAM# = 2.859: GoTo 14
                    If FECHENT% = FECHANUM("05/08/05") Then TICAM# = 2.863: GoTo 14
                    If FECHENT% = FECHANUM("08/08/05") Then TICAM# = 2.865: GoTo 14
                    If FECHENT% = FECHANUM("10/08/05") Then TICAM# = 2.873: GoTo 14
                    If FECHENT% = FECHANUM("12/08/05") Then TICAM# = 2.877: GoTo 14
                    If FECHENT% = FECHANUM("18/08/05") Then TICAM# = 2.89: GoTo 14
                    If FECHENT% = FECHANUM("19/08/05") Then TICAM# = 2.907: GoTo 14
                    If FECHENT% = FECHANUM("22/08/05") Then TICAM# = 2.907: GoTo 14
                    If FECHENT% = FECHANUM("24/08/05") Then TICAM# = 2.908: GoTo 14
                    If FECHENT% = FECHANUM("25/08/05") Then TICAM# = 2.867: GoTo 14
                    If FECHENT% = FECHANUM("29/08/05") Then TICAM# = 2.907: GoTo 14
                    If FECHENT% = FECHANUM("31/08/05") Then TICAM# = 2.873: GoTo 14
                    If FECHENT% = FECHANUM("05/09/05") Then TICAM# = 2.907: GoTo 14
                    If FECHENT% = FECHANUM("06/09/05") Then TICAM# = 2.907: GoTo 14
                    If FECHENT% = FECHANUM("07/09/05") Then TICAM# = 2.907: GoTo 14
                    If FECHENT% = FECHANUM("09/09/05") Then TICAM# = 2.91: GoTo 14
                    If FECHENT% = FECHANUM("13/09/05") Then TICAM# = 2.91: GoTo 14
                    If FECHENT% = FECHANUM("15/09/05") Then TICAM# = 2.919: GoTo 14
                    If FECHENT% = FECHANUM("16/09/05") Then TICAM# = 2.919: GoTo 14
                    If FECHENT% = FECHANUM("19/09/05") Then TICAM# = 2.911: GoTo 14
                    If FECHENT% = FECHANUM("20/09/05") Then TICAM# = 2.911: GoTo 14
                    If FECHENT% = FECHANUM("21/09/05") Then TICAM# = 2.919: GoTo 14
                    If FECHENT% = FECHANUM("22/09/05") Then TICAM# = 2.914: GoTo 14
                    If FECHENT% = FECHANUM("23/09/05") Then TICAM# = 2.914: GoTo 14
                    If FECHENT% = FECHANUM("26/09/05") Then TICAM# = 2.91: GoTo 14
                    If FECHENT% = FECHANUM("27/09/05") Then TICAM# = 2.91: GoTo 14
                    If FECHENT% = FECHANUM("28/09/05") Then TICAM# = 2.915: GoTo 14
                    If FECHENT% = FECHANUM("29/09/05") Then TICAM# = 2.915: GoTo 14
                    If FECHENT% = FECHANUM("30/09/05") Then TICAM# = 2.91: GoTo 14
                    If FECHENT% = FECHANUM("04/10/05") Then TICAM# = 2.909: GoTo 14
                    If FECHENT% = FECHANUM("06/10/05") Then TICAM# = 2.909: GoTo 14
                    If FECHENT% = FECHANUM("07/10/05") Then TICAM# = 2.909: GoTo 14
                    If FECHENT% = FECHANUM("12/10/05") Then TICAM# = 2.909: GoTo 14
                    If FECHENT% = FECHANUM("19/10/05") Then TICAM# = 2.909: GoTo 14
                    '
                    TICAM# = XVALO(InputBox$("Actualizar Cotización Dolar", "Tipo de Cambio al " + FECHATEX$(FECHENT%)))
'14                 Movisto.FindFirst "INT(CDBL(FECHMOV)) = " & Int(CDbl(CDate(CVDAT(FECHENT%))))
'16                   If Movisto.NoMatch = False Then
'                      Movisto.Edit
14                  SQLX$ = " SELECT * FROM MOVISTO"
                    SQLX$ = " WHERE FECHMOV = '" & FETEXCOR1$(FECHENT%) & "'"
                    Dim MOVVVISTO As ADODB.Recordset
                    Set MOVVVISTO = New ADODB.Recordset
28                  On Error Resume Next
                    MOVVVISTO.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
                    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                       On Error GoTo 0
                       Call CapturaErrorOpen
                       GoTo 28
                    End If
                    On Error GoTo 0
                    With MOVVVISTO
                      On Error Resume Next
                      If Not (.EOF And .BOF) Then
                        On Error GoTo 0
                        Do While (Not .EOF)
                          Movisto!TIPOCAM = TICAM#
                          Movisto.Update
  '                       Movisto.FindNext "INT(CDBL(FECHMOV)) = " & Int(CDbl(CDate(CVDAT(FECHENT%))))
  '                       GoTo 16
                          .MoveNext
                        Loop
                      End If
                    End With
                  End If
                  If MONERE% <= 1 Then
                    VUNI# = VUNI# / TICAM#
                  End If
                Call REPLA(Y$, MKD$(VUNI#), 174, 8)
                Call REPLA(Y$, MKD$(!Cantsalid), 182, 8)
                  CANACU# = CANACU# + !Cantsalid
                Call REPLA(Y$, !DoPriCor, 190, 16)
                Call REPLA(Y$, !DoSecCor, 206, 16)
                Call REPLA(Y$, MKD$(TICAM#), 222, 8)
                ' SI MONEDA ES PESOS (0 - 1), GUARDA PESOS Y CALCULA DOLARES
                If MONERE% <= 1 Then
                    NETPES# = !Cantsalid * !ValoUnit
                    IMPDOL# = NETPES# / TICAM#
                  Else
                    IMPDOL# = !Cantsalid * !ValoUnit
                    NETPES# = IMPDOL# * TICAM#
                End If
                Call REPLA(Y$, MKD$(IMPDOL#), 230, 8)
                Call REPLA(Y$, MKD$(NETPES#), 238, 8)
                GIVA% = GRUPIVA%(SCP_COD(U&))
                  TASA# = CVS(Mid$(FILTERGETRECORD("GRUPIVA", 1, MKI$(GIVA%)), 13, 4))
                  IMPIVA# = (NETPES# * TASA#) / 100
                Call REPLA(Y$, MKD$(IMPIVA#), 246, 8)
                  IMPTOT# = NETPES# + IMPIVA#
                Call REPLA(Y$, MKD$(IMPTOT#), 254, 8)
' falta revisar de aqui en adelante
                  ' SI MONEDA ES PESOS (0 - 1), GUARDA PESOS
                  ' Y CALCULA DOLARES
                  ANTIPES# = NETPES# * COEFPES#(U&)
                  ANTIDOL# = IMPDOL# * COEFDOL#(U&)
                Call REPLA(Y$, MKD$(ANTIDOL#), 262, 8)
                Call REPLA(Y$, MKD$(ANTIPES#), 270, 8)
                '
                JJ& = JJ& + 1
                Call GRAREG("SECOMPED", Y$, JJ&)
                '
80            .MoveNext
              Loop
           End With
           '
' aqui poner algo para netear pedidos anulados
           If SCP_STAT%(U&) >= 8 Then
             If Abs(CANACU#) >= 0.05 Then
               XREG$ = REGLEIDO$("SECOMPED", SCP_REG&(U&))
               CANPEDOR# = CVD(Mid$(XREG$, 180, 8))
               Call REPLA(XREG$, MKD$((-1) * CANACU#), 180, 8)
               If Abs(CANPEDOR#) >= 0.05 Then
                   For CLX% = 228 To 252 Step 8
                     VALOPED# = (-1) * CANACU# * CVD(Mid$(XREG$, CLX%, 8)) / CANPEDOR#
                     Call REPLA(XREG$, MKD$(VALOPED#), CLX%, 8)
                   Next CLX%
                 Else
                   Call REPLA(XREG$, String$(32, 0), 230, 32)
               End If
               Call REPLA(XREG$, "Anul.", 190, 16)
               Call REPLA(XREG$, "Anul.", 206, 16)
               JJ& = JJ& + 1
               Call GRAREG("SECOMPED", XREG$, JJ&)
             End If
           End If
           '
           For U2& = U& + 1 To FIN&
             If SCP_CLI(U2&) = SCP_CLI(U&) Then
               If SCP_PED(U2&) = SCP_PED(U&) Then
                 If SCP_COD(U2&) = SCP_COD(U&) Then
                   SCP_CLI(U2&) = 0
                   SCP_PED(U2&) = 0
                   SCP_COD(U2&) = 0
                   SCP_REG&(U2&) = 0
                   SCP_STAT%(U2&) = 0
                 End If
               End If
             End If
           Next U2&
           '
89      Next U&
        '
        Call SETULTREG("SECOMPED", JJ&)
        Call FILESORT("SECOMPED", "", 10, 10, "ASC")
        Call FILESORT("SECOMPED", "", 9, 9, "ASC")
        Call FILESORT("SECOMPED", "", 7, 7, "ASC")
        Call FILESORT("SECOMPED", "", 5, 5, "ASC")
        Call FILESORT("SECOMPED", "", 1, 1, "ASC")
        '
        Call TRALOCAL("SECOMPED", "SECOMPED")
        Call BLANARCH("SECOMPED")
        '
        JJ& = 0
        NCLA% = 0
        For U& = 1 To ULTREG&("!SECOMPED")
          X$ = REGLEIDO$("!SECOMPED", U&)
          NCLI = CVS(Left$(X$, 4))
          If U& > 1 Then
            If NCLI <> NCLA Then
              JJ& = JJ& + 1
              Call GRAREG("SECOMPED", String$(512, "="), JJ&)
            End If
          End If
          NCLA = NCLI
          RASO$ = TRIM$(Str$(NCLI)) + " - " + rasocli$(NCLI)
          DOMI$ = DOMICLI$(NCLI)
          LOCA$ = LOCACLI$(NCLI)
          LIPE$ = X$
          Call REPLA(LIPE$, RASO$, 5, 48)
          Call REPLA(LIPE$, DOMI$, 53, 48)
          Call REPLA(LIPE$, LOCA$, 101, 48)
          JJ& = JJ& + 1
          Call GRAREG("SECOMPED", LIPE$, JJ&)
        Next U&
        '
        Call SETULTREG("SECOMPED", JJ&)
        '
        Call HEADERS("SECOMPED", "")
        '
        Call FINAL("*SECOMPED")
        '
    End If
    '\\\OT-5-686-FG-FIN///
99  Screen.MousePointer = 1
    Command9.Enabled = True
End Sub




Private Sub Form_Load()
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "DOSIVA") > 0 Then
     Command4.Enabled = False
     Label4.Enabled = False
   End If
   '
   If InStr(EPAC$, "J Y M") Then
     Command9.Visible = True
   End If
   
   If InStr(EPAC$, "SABO") > 0 Then
     If EXISTE%("ODESPA02.EXE") > 0 Then
       Command9.Enabled = True
       Command4.Enabled = False
       Label4.Enabled = False
     End If
   End If
   '
   If Control$("ADMIPED", "GENDESPA") = "REVISTOCK" Then
     Command4.Enabled = False
     Label4.Enabled = False
   End If
   '
   If InStr(EPAC$, "DOSIVA") > 0 Then
     OFVER = 1450
     Height = Height + OFVER
     MARCOBARRA.Top = MARCOBARRA.Top + OFVER
     Image3.Top = Image3.Top + OFVER
     'Line1.Y1 = Line1.Y1 + OFVER
     'Line1.Y2 = Line1.Y2 + OFVER
     DoEvents
   End If
   '
   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me)
    
End Sub

Sub RECAPEDPE()
   '
   Call ABREPEDCLI
   SQLX$ = " SELECT * "
   SQLX$ = SQLX$ + " FROM PEDDETA WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE Status < 8 "
   SQLX$ = SQLX$ + " AND STATUS <> -1 " ' EXCLUYE LAS NO APROBADAS
   'SQLX$ = SQLX$ + " AND (((CanPed - CantEnt) >= 0.05) OR (TipoPed = 'ABIERTO' AND CDBL(FinVigen) >= " & Int(CDbl(Now)) & ")) "
   SQLX$ = SQLX$ + " AND (((CanPed - CantEnt) >= 0.05) OR (TipoPed = 'ABIERTO' AND FinVigen >= getdate()))"
   
   '
   'Set PEDTEMP1 = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   Dim PEDTEMP1 As ADODB.Recordset
   Set PEDTEMP1 = New ADODB.Recordset
   PEDTEMP1.CursorLocation = adUseClient
25 On Error Resume Next
   PEDTEMP1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   
   With PEDTEMP1
     On Error Resume Next
     .MoveFirst
     On Error GoTo 0
     '
     V11& = 0
     Y$ = REGBLAN("PEDDETP")
     V11& = 0
     Do While (.EOF = False)
       '
       Call REPLA(Y$, MKS$(!NROPED), 1, 4)
         FECHAPED% = CVINT(!FECHEMIS)
       Call REPLA(Y$, MKI$(FECHAPED%), 5, 2)
       Call REPLA(Y$, MKS$(!NROCLIE), 7, 4)
       Call REPLA(Y$, !NROOCOM, 11, 16)
         FECHASOL% = CVINT(!FeSolEnt)
       Call REPLA(Y$, MKI$(FECHASOL%), 27, 2)
       Call REPLA(Y$, Chr$(!Status), 29, 1)
       Call REPLA(Y$, Chr$(!ListPre), 30, 1)
       Call REPLA(Y$, Chr$(!CondPag), 31, 1)
       Call REPLA(Y$, Chr$(!VENDED), 32, 1)
       Call REPLA(Y$, MKI$(!PorComVend), 33, 2)
       Call REPLA(Y$, MKI$(!CODIINT), 35, 2)
       Call REPLA(Y$, MKS$(!CanPed), 81, 4)
       Call REPLA(Y$, MKS$(!PreUnid), 85, 4)
       Call REPLA(Y$, MKS$(!cantent), 97, 4)
         CAPEN = !CanPed - !cantent
       Call REPLA(Y$, MKS$(CAPEN), 101, 4)
       V11& = V11& + 1
       Call GRAREG("PEDDETP", Y$, V11&)
       '
       .MoveNext
     Loop
   End With
   PEDTEMP1.Close
   Set PEDTEMP1 = Nothing
   '
   Call SETULTREG("PEDDETP", V11&)
   Call CIERRARCH("PEDDETP")
   Call CIERRARCH("PEDDETA")
   Call BAJALDISCO
   '
   Screen.MousePointer = SCMA
   '
End Sub
'\\\OT-05-0249-WAR-04/05/05///
Sub MENSPED()
    '
    Call ABREPEDCLI   ' SE ABRE CONEXION EN CASO QUE SE ACCEDA DESDE EL MENU
    SQLX$ = "SELECT * FROM PEDENCA"
    SQLX$ = SQLX$ + " WHERE STATUS = -1 "
    Dim PEPEAPTMP As ADODB.Recordset
    Set PEPEAPTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    'Set PEPEAPTMP = Ventas.OpenRecordset(sqlx$, 4)
    With PEPEAPTMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        Exit Sub
      End If
      On Error GoTo 0
      Call COMUNI("Hay Pedidos Pendientes de Aprobación\que no Serán Listados.")
    End With
    PEPEAPTMP.Close
    Set PEPEAPTMP = Nothing
    '
End Sub
Sub LIENCPED_NEW()
    '
    Has% = HOY(HOS$)
    FEX = Has%
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    Des% = FECHANUM(DESS$)
    VDEF$ = MKI$(Des%) + MKI$(Has%)
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
10  OKKX% = 0
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
       Exit Sub
    End If
    Screen.MousePointer = 11
    Campos$ = "NroPed/F8;Cta/f7;R.Social/a48;F.Emisión/d8;O.C./A16;Mon/I3;Imp.Orig./F12.2;I.Total($)/F12.2;Observaciones/A27"
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
    Des% = CVI(Left$(OBX$, 2)): DES1 = CDbl(CVDAT(Des%))
    Has% = CVI(Mid$(OBX$, 3, 2)): HAS1 = CDbl(CVDAT(Has%))
    If Has% < Des% Then GoTo 10
    FEX = Des%: PERIO$ = "Periodo: " + FECHATEX$(FEX)
    FEX = Has%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    Desde = FETEXCOR1$(Des%)
    Hasta = FETEXCOR1$(Has%)
    'CARGO EL RECORDSET
    ACUMUIMPED# = 0
    SQLX$ = "SELECT PN.STATUS,PN.Observa,PN.NROPED,PN.NROCLIE,PN.FECHEMIS,PN.NROOCOM,PD.MoneEmis,PD.ImpoTot,PN.RaSoClie FROM PEDENCA PN WITH(NOLOCK)"
    SQLX$ = SQLX$ + " INNER JOIN PEDDETA  PD  ON  PD.NROPED = PN.NROPED "
    SQLX$ = SQLX$ + " WHERE PN.FechEmis >= '" & Desde & "'"
    SQLX$ = SQLX$ + " AND PN.FechEmis <= '" & Hasta & "'"
    SQLX$ = SQLX$ + " ORDER BY PN.NROPED,PN.FechEmis ASC"
    Set PEDTMP = READSET(SQLX$)
    With PEDTMP
      If .EOF Then
        Call COMUNI("No Existen Pedidos que Listar.")
        PEDTMP.Close
        Screen.MousePointer = 1
        Exit Sub
      End If
      JJ& = 0
      YY$ = REGBLAN("ENCPED")
      Do While Not (.EOF)
'        CAMPOS$ = "NroPed/F8;Cta/f7;R.Social/a24;F.Emisión/d8;O.C./A16;I.Total/F12.2;St/A4"

        OBSE$ = SAFETEXT$(!Observa)
        RETEX$(1) = ""
        NP& = XVALO(!NROPED)
        NC& = XVALO(!NROCLIE)
        FE% = CVINT(!FECHEMIS)
        NCOMP$ = SAFETEXT$(!NROOCOM)
        MONEMISPED% = XVALO(!MoneEmis)
        
        If MONEMISPED% < 1 Then MONEMISPED% = 1
        IMPED# = IMPED# + XVALO(!ImpoTot) * TICAMONE(MONEMISPED%)
        IMPEDORIG# = IMPEDORIG# + XVALO(!ImpoTot)
'        ACUMUIMPED# = ACUMUIMPED# + IMPED#

        STX$ = "  "
        If XVALO(!Status) <= 0 Then STX$ = "P."
        If XVALO(!Status) >= 8 Then STX$ = "An"
        '
        If NP_ANT& <> NP& Then
           IMPED# = 0
           IMPED# = IMPED# + XVALO(!ImpoTot) * TICAMONE(MONEMISPED%)
           IMPEDORIG# = 0
           IMPEDORIG# = IMPEDORIG# + XVALO(!ImpoTot)
           J& = J& + 1
           GRILLAGRAL.GRID.Rows = J& + 1
        End If
           Call FRATEXTO(OBSE$, 27)
           GRILLAGRAL.GRID.TextMatrix(J&, 1) = NP&
           GRILLAGRAL.GRID.TextMatrix(J&, 2) = NC&
           GRILLAGRAL.GRID.TextMatrix(J&, 3) = SAFETEXT$(!RasoClie)
           GRILLAGRAL.GRID.TextMatrix(J&, 4) = FECHATEX$(FE%)
           GRILLAGRAL.GRID.TextMatrix(J&, 5) = NCOMP$
           GRILLAGRAL.GRID.TextMatrix(J&, 6) = CStr(MONEMISPED%)
           If MONEMISPED% = 1 Then IMPEDORIG# = 0
           If MONEMISPED% > 1 Then IMPED# = 0
           GRILLAGRAL.GRID.TextMatrix(J&, 7) = FORMATNUM$(IMPEDORIG#, "f12.2")
           GRILLAGRAL.GRID.TextMatrix(J&, 8) = FORMATNUM$(IMPED#, "f12.2")
           GRILLAGRAL.GRID.TextMatrix(J&, 9) = RETEX$(1)
         NP_ANT& = NP&
        .MoveNext
      Loop
    End With
    PEDTMP.Close
    Set PEDTMP = Nothing
    
    '*******************************************
    'solicitud de ambest que considera que si hay un pedido anulado en el rango seleccionado
    'debe figurar en negativo, suponiendo que el mismo se tuvo en cuenta en algun momento.
    'si en el rango esta el pedido y la anulacion va aparecer una vez en positivo y otra en negativo.
    If InStr(1, EPAC$, "AMBEST") > 0 Then
        SQLX$ = "SELECT PN.STATUS,PN.Observa,PN.NROPED,PN.NROCLIE,PN.FECHEMIS,PN.NROOCOM,PD.MoneEmis,PD.ImpoTot,PN.RaSoClie FROM PEDENCA PN WITH(NOLOCK)"
        SQLX$ = SQLX$ + " INNER JOIN PEDDETA  PD  ON  PD.NROPED = PN.NROPED "
        SQLX$ = SQLX$ + " WHERE PN.FechAnul >= '" & Desde & "'"
        SQLX$ = SQLX$ + " AND PN.FechAnul <= '" & Hasta & "'"
        SQLX$ = SQLX$ + " AND PN.Status = 9 "
        SQLX$ = SQLX$ + " ORDER BY PN.NROPED,PN.FechEmis ASC"
        Set PEDTMP = READSET(SQLX$)
        With PEDTMP
          YY$ = REGBLAN("ENCPED")
          Do While Not (.EOF)
            OBSE$ = "(Anul) " & SAFETEXT$(!Observa)
            RETEX$(1) = ""
            NP& = XVALO(!NROPED)
            NC& = XVALO(!NROCLIE)
            FE% = CVINT(!FECHEMIS)
            NCOMP$ = SAFETEXT$(!NROOCOM)

            MONEMISPED% = XVALO(!MoneEmis)
            If MONEMISPED% < 1 Then MONEMISPED% = 1
            IMPED# = IMPED# + XVALO(!ImpoTot) * TICAMONE(MONEMISPED%)
            IMPEDORIG# = IMPEDORIG# + XVALO(!ImpoTot)
            
            If NP_ANT& <> NP& Then
               IMPED# = 0
               IMPED# = IMPED# + XVALO(!ImpoTot) * TICAMONE(MONEMISPED%)
               IMPEDORIG# = 0
               IMPEDORIG# = IMPEDORIG# + XVALO(!ImpoTot)
               J& = J& + 1
               GRILLAGRAL.GRID.Rows = J& + 1
            End If
               Call FRATEXTO(OBSE$, 27)
               GRILLAGRAL.GRID.TextMatrix(J&, 1) = NP&
               GRILLAGRAL.GRID.TextMatrix(J&, 2) = NC&
               GRILLAGRAL.GRID.TextMatrix(J&, 3) = SAFETEXT$(!RasoClie)
               GRILLAGRAL.GRID.TextMatrix(J&, 4) = FECHATEX$(FE%)
               GRILLAGRAL.GRID.TextMatrix(J&, 5) = NCOMP$
               GRILLAGRAL.GRID.TextMatrix(J&, 6) = CStr(MONEMISPED%)
               If MONEMISPED% = 1 Then IMPEDORIG# = 0
               If MONEMISPED% > 1 Then IMPED# = 0
               GRILLAGRAL.GRID.TextMatrix(J&, 7) = FORMATNUM$(-1 * IMPEDORIG#, "f12.2")
               GRILLAGRAL.GRID.TextMatrix(J&, 8) = FORMATNUM$(-1 * IMPED#, "f12.2")
               GRILLAGRAL.GRID.TextMatrix(J&, 9) = RETEX$(1)
               'GRILLAGRAL.GRID.TextMatrix(J&, 6) = FORMATNUM$(-1 * IMPED#, "f12.2")
               'GRILLAGRAL.GRID.TextMatrix(J&, 7) = RETEX$(1)
               NP_ANT& = NP&
               
            .MoveNext
          Loop
        End With
        PEDTMP.Close
        Set PEDTMP = Nothing
    End If
    '********************************************
    Screen.MousePointer = 1
    If GRILLAGRAL.GRID.Rows < 2 Then
      Call COMUNI("No se Registraron Pedidos")
      Exit Sub
    End If
    'TOTALIZA
    GRILLAGRAL.GRID.Rows = J& + 2
    GRILLAGRAL.GRID.TextMatrix(J& + 1, 5) = "TOTAL"
    
    'TOTALIZO CON FUNCION PARA QUE LEA LOS DATOS FORMATEADOS Y NO HAYA DIFERENCIA POR REDONDEOS PARCIALES
    GRILLAGRAL.GRID.TextMatrix(J& + 1, 7) = FORMATNUM$(SUMACOLUMNA#(GRILLAGRAL.GRID, 7, 1), "f12.2")
    GRILLAGRAL.GRID.TextMatrix(J& + 1, 8) = FORMATNUM$(SUMACOLUMNA#(GRILLAGRAL.GRID, 8, 1), "f12.2")
    
    Call COLOREAR_GRILLA_SOLAFILA(GRILLAGRAL.GRID, &HE0E0E0, J& + 1)

    TX$ = "Encabezado de Pedidos " & PERIO$
    GRILLAGRAL.Caption = TX$
    '*******************************************
    GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.TXTBUSCAR = ""
    Call POSICIONAR_ULTIMAFILA(GRILLAGRAL.GRID)

    GRILLAGRAL.Show 1
    Dim MSF As msFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(MSF, Campos$, "ENCAPEDI", TX$, "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
    End If
    Set obj = Nothing

''        Call FRATEXTO(OBSE$, 27)
''        Call REPLA(YY$, MKS$(!NROPED), 1, 4)
''        Call REPLA(YY$, MKS$(!NROCLIE), 5, 4)
''        Call REPLA(YY$, MKI$(CVINT(!FECHEMIS)), 9, 2)
''        Call REPLA(YY$, !NROOCOM, 11, 16)
'        '
'        'jandy sql
'        strSQL = "Select * from PedDeta where NroPed=" & !NROPED
'        Dim PedDeta As ADODB.Recordset
'        Set PedDeta = New ADODB.Recordset
'        PedDeta.CursorLocation = adUseClient
'26      On Error Resume Next
'        PedDeta.Open FILTSQL$(strSQL), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
'        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'           On Error GoTo 0
'           Call CapturaErrorOpen
'           GoTo 26
'        End If
'        On Error GoTo 0
'        On Error Resume Next
'        IMPED# = 0
'        With PedDeta
'          Do While Not (.EOF)
'           MONEMISPED% = PedDeta!MoneEmis
'            IMPED# = IMPED# + PedDeta.Fields("IMPOTOT") * TICAMONE(MONEMISPED%)
'          .MoveNext
'          Loop
'        End With
'        PedDeta.Close
'        Set PedDeta = Nothing
'        'PedDeta.FindFirst (" NROPED = " & !NroPed & " ")
'
'25      'If PedDeta.NoMatch = False Then
'          'MONEMISPED% = PedDeta!MoneEmis
'          'IMPED# = IMPED# + PedDeta.Fields("IMPOTOT") * TICAMONE(MONEMISPED%)
'          'PedDeta.FindNext (" NROPED = " & !NroPed & " ")
'          'GoTo 25
'        'End If
'
'        If MODOGRAMY% = 1 Then
'          IMPED# = IMPED# / TICAMONE(2)  ' PARA LLEVARLO A DOLARES
'        End If
'
'        '
'        Call REPLA(YY$, MKD$(IMPED#), 27, 8)
'        Call REPLA(YY$, MKI$(!Status), 35, 2)
'        If MODOGRAMY% = 1 Then
'            STX$ = "  "
'            If !Status < 0 Then STX$ = "P."
'            If !Status >= 8 Then STX$ = "An"
'            Call REPLA(YY$, STX$, 35, 2)
'        End If
'        Call REPLA(YY$, RETEX$(1), 37, 27)
'        jj& = jj& + 1
'        Call GRAREG("ENCPED", YY$, jj&)
'        '
'        For KKU& = 2 To CARETEX%
'           Call REPLA(YY$, MKD$(0), 27, 8) ' AGREGADO PARA QUE CUANDO REPITE LA FILA NO REPITA IMPORTE
'           Call REPLA(YY$, RETEX$(KKU&), 37, 27)
'           jj& = jj& + 1
'           Call GRAREG("ENCPED", YY$, jj&)
'        Next KKU&
'      .MoveNext
'      Loop
'    End With
'    PEDTMP.Close
'    Set PEDTMP = Nothing
'    '
'    Call SETULTREG("ENCPED", jj&)
'    Call FINAL("*LIENPED")
'End Sub
''\\\OT-05-0250-WAR-FIN///
'
'
'
'\\\OT-05-0249-WAR-FIN///
'\\\OT-05-0250-WAR-19/05/05///
End Sub
Sub LIENCPED()
    '
    Call LIENCPED_NEW
    Exit Sub
    Has% = HOY(HOS$)
    FEX = Has%
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    Des% = FECHANUM(DESS$)
    VDEF$ = MKI$(Des%) + MKI$(Has%)
    '
10  OKKX% = 0
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
       Exit Sub
    End If
    OKKX% = 1
    Des% = CVI(Left$(OBX$, 2)): DES1 = CDbl(CVDAT(Des%))
    Has% = CVI(Mid$(OBX$, 3, 2)): HAS1 = CDbl(CVDAT(Has%))
    If Has% < Des% Then GoTo 10
    FEX = Des%: PERIO$ = "Periodo: " + FECHATEX$(FEX)
    FEX = Has%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    MODOGRAMY% = 0
    If Control$("ADMIPED", "FORMALIS") = "MONEMORI" Then
      If EXISTE%(LUDAT$ + "LIPEDID.RFS") > 0 Then
        MODOGRAMY% = 1
      End If
    End If
    '
    Call HEADERS("ENCPED", "")
    Call HEADERS("ENCPED", PERIO$)
    'jandy sql
    Desde = FETEXCOR1$(Des%)
    Hasta = FETEXCOR1$(Has%)
    Call ABREPEDCLI
    SQLX$ = "SELECT * FROM PEDENCA"
    'SQLX$ = SQLX$ + " WHERE INT(CDBL(FechEmis)) >= " & DES1 & ""
    'SQLX$ = SQLX$ + " AND INT(CDBL(FechEmis)) <= " & HAS1 & ""
    SQLX$ = SQLX$ + " WHERE FechEmis >= '" & Desde & "'"
    SQLX$ = SQLX$ + " AND FechEmis <= '" & Hasta & "'"
    SQLX$ = SQLX$ + " ORDER BY FechEmis ASC, NroPed ASC "
    'Set PEDTMP = Ventas.OpenRecordset(SQLX$, 4)
    Dim PEDTMP As ADODB.Recordset
    Set PEDTMP = New ADODB.Recordset
    PEDTMP.CursorLocation = adUseClient
27  On Error Resume Next
    PEDTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 27
    End If
    On Error GoTo 0
    With PEDTMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call COMUNI("No Existen Pedidos que Listar.")
        On Error GoTo 0
        Exit Sub
      End If
      On Error GoTo 0
      JJ& = 0
      YY$ = REGBLAN("ENCPED")
      Do While .EOF = False
      
        OBSE$ = !Observa
        RETEX$(1) = ""
        Call FRATEXTO(OBSE$, 27)
        Call REPLA(YY$, MKS$(!NROPED), 1, 4)
        Call REPLA(YY$, MKS$(!NROCLIE), 5, 4)
        Call REPLA(YY$, MKI$(CVINT(!FECHEMIS)), 9, 2)
        Call REPLA(YY$, !NROOCOM, 11, 16)
        '
        'jandy sql
        strSQL = "Select * from PedDeta where NroPed=" & !NROPED
        Dim PedDeta As ADODB.Recordset
        Set PedDeta = New ADODB.Recordset
        PedDeta.CursorLocation = adUseClient
26      On Error Resume Next
        PedDeta.Open FILTSQL$(strSQL), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 26
        End If
        On Error GoTo 0
        On Error Resume Next
        IMPED# = 0
        With PedDeta
          Do While Not (.EOF)
           MONEMISPED% = PedDeta!MoneEmis
            IMPED# = IMPED# + PedDeta.Fields("IMPOTOT") * TICAMONE(MONEMISPED%)
          .MoveNext
          Loop
        End With
        PedDeta.Close
        Set PedDeta = Nothing
        'PedDeta.FindFirst (" NROPED = " & !NroPed & " ")
        
25      'If PedDeta.NoMatch = False Then
          'MONEMISPED% = PedDeta!MoneEmis
          'IMPED# = IMPED# + PedDeta.Fields("IMPOTOT") * TICAMONE(MONEMISPED%)
          'PedDeta.FindNext (" NROPED = " & !NroPed & " ")
          'GoTo 25
        'End If
        
        If MODOGRAMY% = 1 Then
          IMPED# = IMPED# / TICAMONE(2)  ' PARA LLEVARLO A DOLARES
        End If
           
        '
        Call REPLA(YY$, MKD$(IMPED#), 27, 8)
        Call REPLA(YY$, MKI$(!Status), 35, 2)
        If MODOGRAMY% = 1 Then
            STX$ = "  "
            If !Status < 0 Then STX$ = "P."
            If !Status >= 8 Then STX$ = "An"
            Call REPLA(YY$, STX$, 35, 2)
        End If
        Call REPLA(YY$, RETEX$(1), 37, 27)
        JJ& = JJ& + 1
        Call GRAREG("ENCPED", YY$, JJ&)
        '
        For KKU& = 2 To CARETEX%
           Call REPLA(YY$, MKD$(0), 27, 8) ' AGREGADO PARA QUE CUANDO REPITE LA FILA NO REPITA IMPORTE
           Call REPLA(YY$, RETEX$(KKU&), 37, 27)
           JJ& = JJ& + 1
           Call GRAREG("ENCPED", YY$, JJ&)
        Next KKU&
      .MoveNext
      Loop
    End With
    PEDTMP.Close
    Set PEDTMP = Nothing
    '
    Call SETULTREG("ENCPED", JJ&)
    Call FINAL("*LIENPED")
End Sub
'\\\OT-05-0250-WAR-FIN///

Private Sub IMPUTAPEDI()
   '
   SQLX$ = "SELECT * FROM PEDENCA "
'   If PROLLAMA$ = "ABIE" Then
'        SQLX$ = SQLX$ + "WHERE TipoPed = 'ABIERTO' "
'      Else
'        SQLX$ = SQLX$ + "WHERE TipoPed <> 'ABIERTO' "
'        SQLX$ = SQLX$ + "OR IsNull(TipoPed) = True "
'   End If
   SQLX$ = SQLX$ + "ORDER BY NroPed ASC"
   'Set PEDENCTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim PEDENCTEMP As ADODB.Recordset
   Set PEDENCTEMP = New ADODB.Recordset
25 On Error Resume Next
   PEDENCTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0

   JJ& = 0
   With PEDENCTEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        Do While Not .EOF
             NRODES& = !NROPED
             TTXX$ = MKS$(NRODES&)
               NOCO$ = !NROOCOM
             If NOCO$ <> "" Then
               TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
             End If
               FEX = CVINT(!FECHEMIS)
             TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
               NCLIE% = !NROCLIE
               ANULX$ = "": If !Status = 9 Then ANULX$ = "(Anul) "
             TTXX$ = TTXX$ + TRIM$(rasocli$(NCLIE%)) + ANULX$ 'Space$(52)
             JJ& = JJ& + 1
             Call GRAREG("!INDIPEDI", TTXX$, JJ&)
        .MoveNext
        Loop
      End If
   End With
   PEDENCTEMP.Close
   Set PEDENCTEMP = Nothing
   '
   Call SETULTREG("!INDIPEDI", JJ&)
   Call CIERRARCH("!INDIPEDI")
   Call FRESHECHO("!INDIPEDI")
   Screen.MousePointer = 1
   '
   'Dim FORMORIG As Form
   'Set FORMORIG = FORPED04
   'If PROLLAMA = "ABIE" Then Set FORMORIG = PEDABIE04
   '
10 Call SELECHO("!INDIPEDI/INDICE GENERAL DE PEDIDOS")
   NPSEL& = XVALO(VALACT1$("!INDIPEDI"))
   TIPAN$ = REPLACAR$(TRIM$(VALACT2$("!INDIPEDI")), "/", "-")
   If NPSEL& > 0 Then
        JJ& = 0
        REBLA$ = REGBLAN$("CARDETPE")
        SQLX$ = "SELECT * FROM PEDDETA "
        SQLX$ = SQLX$ + "WHERE NroPed = " & NPSEL& & " "
        SQLX$ = SQLX$ + "ORDER BY NuorItem ASC "
        Set PEDDETTEMP = New ADODB.Recordset
26      On Error Resume Next
        PEDDETTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 26
        End If
        On Error GoTo 0

        On Error Resume Next
        PEDDETTEMP.MoveFirst
        If Err < 1 Then
          Do While Not PEDDETTEMP.EOF
           'For U& = 1 To ULTREG&("!CARDETPE")
            X$ = REBLA$
              CIXI% = PEDDETTEMP!CODIINT
            Call REPLA(X$, MKI$(CIXI%), 1, 2)                    ' CODIGO
              MONEMIORI% = PEDDETTEMP!MoneEmis
            Call REPLA(X$, MON$, 19, 3)                          ' MONEDA
            'Call REPLA(X$, PEDDETTEMP!TColor, 3, 16)
            'Call REPLA(X$, PEDDETTEMP!TTalle, 19, 8)
            '  TLAR = PEDDETTEMP!TLargo
            'Call REPLA(X$, MKS$(TLAR), 27, 4)
            '  TCHO = PEDDETTEMP!TAncho
            'Call REPLA(X$, MKS$(TCHO), 31, 4)
            '  TESP = PEDDETTEMP!TEspesor
            'Call REPLA(X$, MKS$(TESP), 35, 4)
              PREL = PEDDETTEMP!PreList
            Call REPLA(X$, MKS$(PREL), 51, 4)                    ' PRECIO
              CAN = PEDDETTEMP!CanPed
            Call REPLA(X$, MKS$(CAN), 47, 4)                     ' CANTIDAD
            'Call REPLA(X$, MKI$(HOII%), 63, 2) ' FECHA SOLICITADA DE ENTREGA
            '  SUCOD$ = ""
            '  If IsNull(PEDDETTEMP!SubCod) = False Then SUCOD$ = PEDDETTEMP!SubCod
            'Call REPLA(X$, SUCOD$, 109, 8)
              TOT# = PEDDETTEMP!ImpoTot
            Call REPLA(X$, MKD$(TOT#), 55, 8)             ' TOTAL
              ANTI# = 0
              On Error Resume Next
              ANTI# = PEDDETTEMP!ImpoAnti
              On Error GoTo 0
            Call REPLA(X$, MKD$(ANTI#), 77, 8)             ' ANTICIPO
            Call REPLA(X$, MKS$(PEDDETTEMP!NuOrItem), 125, 4)    ' NUMERO DE ORDEN
            JJ& = JJ& + 1
            Call GRAREG("!CARDETPE", X$, JJ&)
          PEDDETTEMP.MoveNext
          Loop
        End If
        Call SETULTREG("!CARDETPE", JJ&)
        Call CIERRARCH("CARDETPE")
        On Error GoTo 0
        '
        TIPAN$ = TIPAN$ + " (" + DEMONECO$(MONEMIORI%) + ")"
        XX% = VENTABU%("IMPUPED/TITUPAN=" + TIPAN$)
        '
        If XX% >= 0 Then
          For KU& = 1 To ULTREG&("!CARDETPE")
            X$ = REGLEIDO$("!CARDETPE", KU&)
            NORI& = CVS(Mid$(X$, 125, 4))
            ANTI# = CVD(Mid$(X$, 77, 8))
            '
            SQLX$ = "SELECT * FROM PEDDETA "
            SQLX$ = SQLX$ + "WHERE NroPed = " & NPSEL& & " "
            SQLX$ = SQLX$ + "and NuOrItem = " & NORI&
            Set PEDDETTEMP = New ADODB.Recordset
27          On Error Resume Next
            PEDDETTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
            If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
               On Error GoTo 0
               Call CapturaErrorOpen
               GoTo 27
            End If
            On Error GoTo 0
            '
            With PEDDETTEMP
               On Error Resume Next
               .MoveFirst
               If Err Then
                  On Error GoTo 0
                  GoTo 29
               End If
               On Error GoTo 0
               On Error Resume Next
               !ImpoAnti = ANTI#
               On Error GoTo 0
               .Update
            End With
29          PEDDETTEMP.Close
            Set PEDDETTEMP = Nothing
          Next KU&
        End If
        '
        GoTo 10
        '
    End If
   '
End Sub

Sub FILARTCLI()
   '
   '\\\OT-08-0200-OD-29/04/08///
15 Call SELECHO("DEUDOR12")
   NCLIE = XVALO(TRIM$(VALACT1$("DEUDOR12")))
   If NCLIE > 0 Then
20   Call DESHASFECHA(Des%, Has%, PERIO$)
     If TRIM$(PERIO$) = "" Then GoTo 15
     DESDAT = FETEXCOR1$(Des%)
     HASDAT = FETEXCOR1$(Has%)
     '
     JJ& = 0
     Call APERBASDAT("CABAN")
     SQLX$ = "SELECT CodiCom_Lar, FechConta, ValAbsComp FROM CAJABANC "
     SQLX$ = SQLX$ + " where FechConta >= '" & DESDAT & "' "
     SQLX$ = SQLX$ + " AND FechConta <= '" & HASDAT & "' "
     SQLX$ = SQLX$ + " AND NuClien = " & NCLIE & ""
     SQLX$ = SQLX$ + " AND TipoMovim = 'FVEN'"
     SQLX$ = SQLX$ + " AND OpciMovim = 'FC'"
     SQLX$ = SQLX$ + " AND ValAbsComp > 0"
     SQLX$ = SQLX$ + " ORDER BY FechConta ASC, CodiCom_Lar ASC"
     'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
     Dim CABATEMP As ADODB.Recordset
     Set CABATEMP = New ADODB.Recordset
25   On Error Resume Next
     CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
        GoTo 25
     End If
     On Error GoTo 0

     With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err > 0 Then
       Call MENSERR(24, "No se Encuentran Facturas en el Rango Elegido ")
       GoTo 20
      End If
       Y$ = REGBLAN("INDICOM")
       On Error GoTo 0
         Do While Not .EOF
          JJ& = JJ& + 1
          ORDEN% = JJ&
          'CARGAR EL ARCHIVO PARA SELECCION
          FECON% = CVINT(!fechconta)
          Call REPLA(Y$, MKI(ORDEN%), 1, 2)
          Call REPLA(Y$, !CODICOM_LAR + "  -  " & FECHATEX$(CVINT(!fechconta)) & "  -  $ " & FORMATNUM$(!ValAbsComp, "F12.2"), 3, 56)
          Call GRAREG("!INDICOM", Y$, JJ&)
         .MoveNext
         Loop
     End With
     CABATEMP.Close
     Set CABATEMP = Nothing
     '
     Call SETULTREG("!INDICOM", JJ&)
30   Call FRESHECHO("!INDICOM")
     Call SELECHO("!INDICOM")
     
     NUMFACT$ = (TRIM$(VALACT2$("!INDICOM")))
     If NUMFACT$ = "" Then
        GoTo 20
       Else:
        'PEDIDOS QUE COINCIDAN CON NUMERO DE FACTURA
        '
        NUMFACT1$ = TRIM$(Left$(NUMFACT$, 18))
        JJ& = 0
        Call ABRESTOCKS
        Call ABREPEDCLI
        SQLX$ = "SELECT DISTINCT NuPedCli FROM MOVISTO "
        SQLX$ = SQLX$ + " WHERE DoSecLar = '" & NUMFACT1$ & "'"
        SQLX$ = SQLX$ + " ORDER BY NuPedCli ASC"
        'Set MOVISTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
        Dim MOVISTMP As ADODB.Recordset
        Set MOVISTMP = New ADODB.Recordset
26      On Error Resume Next
        MOVISTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 26
        End If
        On Error GoTo 0

        With MOVISTMP
         On Error Resume Next
         .MoveFirst
         If Err > 0 Then
          Call MENSERR(24, "No se Encuentran Pedidos en el Rango Elegido ")
          GoTo 30
         End If
         Y$ = REGBLAN("INDIPEP")
         On Error GoTo 0
         Do While Not .EOF
            NPEDI& = !NUPEDCLI
            SQLY$ = "SELECT FechEmis, Referencia FROM PEDENCA WHERE nroPed = " & NPEDI&
            'Set ENCAPEDI = Ventas.OpenRecordset(SQLY$, dbOpenSnapshot)
            Dim ENCAPEDI As ADODB.Recordset
            Set ENCAPEDI = New ADODB.Recordset
27          On Error Resume Next
            ENCAPEDI.Open FILTSQL$(SQLY$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
            If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
               On Error GoTo 0
               Call CapturaErrorOpen
               GoTo 27
            End If
            On Error GoTo 0

            RAZSOC$ = ""
            On Error Resume Next
            ENCAPEDI.MoveFirst
            If Err < 1 Then
               On Error GoTo 0
               RAZSOC$ = FECHATEX$(CVINT(ENCAPEDI!FECHEMIS)) + " - " + ENCAPEDI!REFERENCIA
             Else
              Call MENSERR(24, "No se Encuentran Pedidos Relacionados con esa Factura ")
              GoTo 30
            End If
            On Error GoTo 0
            NUMPEDI$ = MKS$(NPEDI&)
            Call REPLA(Y$, NUMPEDI$, 1, 4)
            Call REPLA(Y$, RAZSOC$, 5, 48)
            JJ& = JJ& + 1
            Call GRAREG("!INDIPEP", Y$, JJ&)
            .MoveNext
         Loop
        End With
        MOVISTMP.Close
        Set MOVISTMP = Nothing
     End If
     Call SETULTREG("!INDIPEP", JJ&)
2110 Call FRESHECHO("!INDIPEP")
     Call SELECHO("!INDIPEP/Indice De Pedidos .")
     NUMEPEDI = XVALO(VALACT1$("!INDIPEP"))
     If XVALO(NUMEPEDI) <= 0 Then GoTo 30
     NUMEP% = NUMEPEDI
     '
     NPX& = NUMEP%
     If NPX& < 1 Then GoTo 9999
     '
     TIDOCUM$ = "Pedido de Cliente"
     TIDOCUX$ = "Pedido ABIERTO"
     TIDOCUY$ = "Pedido Interno"
     If DERACCE%("VERCOSTO/NOMESS", "Visualizar Costos y Precios") < 1 Then
       TIDOCUM$ = "Orden de Despacho"
       TIDOCUX$ = "Orden de Despacho"
       TIDOCUY$ = "Orden de Despacho"
     End If
    '
     NUDOCU$ = TRIM$(Str$(NPX&))
     While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
     NUDOCU$ = NUDOCU$ + " "
     Screen.MousePointer = 11
    '
4    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Or InStr(XX$, TIDOCUX$) > 4 Or InStr(XX$, TIDOCUY$) > 4 Then
        PPXX% = InStr(XX$, "Nro.")
        If PPXX% > 0 Then
          XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
        End If
        If InStr(XX$, NUDOCU$) > 4 Then
          DOCUNU& = CVS(Left$(XX$, 4))
          GoTo 5
        End If
      End If
     Next UI&
    '
     Screen.MousePointer = 1
     Call COMUNI("Documento no Encontrado")
     Exit Sub
    '
5    For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
     Next KKI%
    '
     REDOCU$ = REGACT$("PDOCLST")
     URE& = ULTREG&("PDOCSPL")
     LI& = 0: LS& = URE&
6    E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
     L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
     DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
     If DCC& < DOCUNU& Then LI& = L&: GoTo 6
     If DCC& > DOCUNU& Then LS& = L&: GoTo 6
     GoTo 8
    '
7    Screen.MousePointer = 1
     Call MENSERR(24, "Documento no Registrado")
     GoTo 9999
    '
8    TPSP$ = ""
     Screen.MousePointer = 11
     While L& > 1
      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
      If DCC& < DOCUNU& Then GoTo 9
      L& = L& - 1
     Wend
    '
9    CAPAGINAS% = 0
     For KKL& = L& To URE&
      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
      End If
     Next KKL&
    '
95   SPOOLSTRING$ = TPSP$
     For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).Top = 0
      EDITFORM.Picture1(KKI%).Refresh
      EDITFORM.Picture1(KKI%).Height = 7170
      EDITFORM.Picture1(KKI%).Width = 11700
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
     Next KKI%
     ALTUPAGINA = 7170
     TOPEPAGINA = 0
     PAGINACTIVA% = 0
     CAPAGINAS% = 0
    '
10   PPXX% = InStr(TPSP$, Chr$(255))
     If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call SHOWCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
     End If
    '
     Screen.MousePointer = 1
     PAGINACTIVA% = 0
     For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Visible = False
     Next KKI%
     EDITFORM.Picture1(PAGINACTIVA%).Visible = True
     EDITFORM.Picture1(PAGINACTIVA%).Refresh
     EDITFORM.VScroll1.Min = 0
     EDITFORM.VScroll1.Value = 0
     EDITFORM.VScroll1.Max = 1
     If ALTUPAGINA > EDITFORM.Frame1.Height Then
      EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
     End If
     EDITFORM.Command3.Enabled = False
     'EDITFORM.Command5.Enabled = False
     EDITFORM.Show 1
     Screen.MousePointer = 1
     '
     GoTo 2110
    '
    Screen.MousePointer = 1
   End If
   Call CIERRARCH("!INDICOM")
   Call FRESHECHO("!INDICOM")
9999 Screen.MousePointer = 1
   '
   '\\\OT-08-0200-OD-FIN///
End Sub

Sub MODELO(OPC%)

    If OPC% = 2 Then 'Por numero de pedido
        'Call PEDPENDAKOT(OPC%)
        Call LIENCPED
    ElseIf OPC% = 3 Then ' por color base
        Call PEDPENDAKOT(OPC%)
    ElseIf OPC% = 4 Then ' por color tela
        Call PEDPENDAKOT(OPC%)
    ElseIf OPC% = 5 Then ' por tela
        Call PEDPENDAKOT(OPC%)
    End If

End Sub
Sub PEDPENDAKOT(ORDEN%)
    
    Has% = HOY(HOS$)
    FEX = Has%
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    Des% = FECHANUM(DESS$)
    VDEF$ = MKI$(Des%) + MKI$(Has%)
    '
10  OKKX% = 0
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
       Exit Sub
    End If
    OKKX% = 1
    Des% = CVI(Left$(OBX$, 2)): DES1 = CDbl(CVDAT(Des%))
    Has% = CVI(Mid$(OBX$, 3, 2)): HAS1 = CDbl(CVDAT(Has%))
    If Has% < Des% Then GoTo 10
    FEX = Des%: PERIO$ = "Periodo: " + FECHATEX$(FEX)
    FEX = Has%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    Call ABREPEDCLI
    SQLX$ = "SELECT * FROM PEDDETA"
    SQLX$ = SQLX$ + " WHERE INT(CDBL(FeSolEnt)) >= " & DES1 & ""
    SQLX$ = SQLX$ + " AND INT(CDBL(FeSolEnt)) <= " & HAS1 & ""
    SQLX$ = SQLX$ + " AND INT(CDBL(STATUS)) < " & 8 & ""
    Set PEDTMP = Ventas.OpenRecordset(SQLX$, 4)
    With PEDTMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call COMUNI("No Existen Pedidos que Listar.")
        On Error GoTo 0
        Exit Sub
      End If
      On Error GoTo 0
      JJ& = 0
      Y$ = REGBLAN("PEPENDAK")
      Do While .EOF = False
        Call REPLA(Y$, MKS$(!NROPED), 1, 4)
        Call REPLA(Y$, MKI$(!NROCLIE), 5, 2)
        
        FECEMI% = CVINT(!FECHEMIS)
        Call REPLA(Y$, MKI$(FECEMI%), 7, 2)
        FECENT% = CVINT(!FeSolEnt)
        Call REPLA(Y$, MKI$(FECENT%), 9, 2)
        '
        COLBASE$ = ""
        If IsNull(!CBase) = False Then COLBASE$ = !CBase
        Call REPLA(Y$, COLBASE$, 11, 16)
        '
        Call REPLA(Y$, MKI$(!CODIINT), 27, 2)
        Call REPLA(Y$, COLBASE$, 29, 16)
        '
        TELA$ = ""
        If IsNull(!TTALLE) = False Then TELA$ = !TTALLE
        Call REPLA(Y$, TELA$, 45, 16)
        '
        COLTELA$ = ""
        If IsNull(!TCOLOR) = False Then COLTELA$ = !TCOLOR
        Call REPLA(Y$, COLTELA$, 61, 16)
        '
        Call REPLA(Y$, MKS$(!CanSaldo), 77, 4)
        JJ& = JJ& + 1
        Call GRAREG("PEPENDAK", Y$, JJ&)
        .MoveNext
      Loop
      On Error GoTo 0
    End With
    '
    Call SETULTREG("PEPENDAK", JJ&)
    Call CIERRARCH("PEPENDAK")
    Call HEADERS("PEPENDAK", "")
    Call HEADERS("PEPENDAK", PERIO$)
    '
'    If ORDEN% = 2 Then ' por numero de pedido
'        Call FILESORT("PEPENDAK", "", 1, 1, "ASC")
'        Call FINAL("*PEPENDAK")
    If ORDEN% = 3 Then ' por color base
        Call FILESORT("PEPENDAK", "", 7, 7, "ASC")
        Call FILESORT("PEPENDAK", "", 6, 6, "ASC")
        Call FINAL("*PEPENCB")
    ElseIf ORDEN% = 4 Then ' por color tela
        Call FILESORT("PEPENDAK", "", 6, 6, "ASC")
        Call FILESORT("PEPENDAK", "", 7, 7, "ASC")
        Call FINAL("*PEPENCOD")
    ElseIf ORDEN% = 5 Then ' por tela
        Call FILESORT("PEPENdAK", "", 10, 10, "ASC")
        Call FILESORT("PEPENDAK", "", 7, 7, "ASC")
        Call FINAL("*PEPENTEL")
    End If
    '
End Sub


Sub PEDAFACT()
'\\\OT-08-0200-OD-29/04/08///
15 Call SELECHO("DEUDOR12")
   If TRIM$(VALACT1$("DEUDOR12")) = "" Then Exit Sub
   NCLIE = XVALO(TRIM$(VALACT1$("DEUDOR12")))
   If NCLIE > 0 Then
20   Call DESHASFECHA(Des%, Has%, PERIO$)
     If TRIM$(PERIO$) = "" Then GoTo 15
     DESDAT = FETEXCOR1$(Des%)
     HASDAT = FETEXCOR1$(Has%)
     '
     SQLX$ = "SELECT  DISTINCT NroPed , NroClie, FechEmis FROM PEDENCA "
     SQLX$ = SQLX$ + " where FechEmis >= '" & DESDAT & "' "
     SQLX$ = SQLX$ + " AND FechEmis <= '" & HASDAT & "' "
     SQLX$ = SQLX$ + " AND NroClie = " & NCLIE & ""
     '
     JJ& = 0
     Dim PEDENTEMP As ADODB.Recordset
     Set PEDENTEMP = New ADODB.Recordset
25   On Error Resume Next
     PEDENTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
        GoTo 25
     End If
     On Error GoTo 0
     '
     With PEDENTEMP
       On Error Resume Next
         .MoveFirst
         If Err > 0 Then
          Call MENSERR(24, "No se Encuentran Pedidos Coincidentes con el Rango Elegido ")
          GoTo 20
         End If
         Y$ = REGBLAN("INDIPEP")
         Call BLANARCH("!INDIPEP")
         On Error GoTo 0
         Do While Not .EOF
            NPEDI& = !NROPED
            NUMPEDI$ = MKS$(NPEDI&)
            NUMCLIEN$ = !NROCLIE
            RAZSOC$ = rasocli$(NUMCLIEN$)
            'CARGAR EL ARCHIVO PARA SELECCION
            Call REPLA(Y$, NUMPEDI$, 1, 4)
            Call REPLA(Y$, RAZSOC$ & " - " & !FECHEMIS, 5, 48)
            JJ& = JJ& + 1
            Call GRAREG("!INDIPEP", Y$, JJ&)
            .MoveNext
         Loop
        End With
        PEDENTEMP.Close
        Set PEDENTEMP = Nothing
     Call SETULTREG("!INDIPEP", JJ&)
30   Call FRESHECHO("!INDIPEP")
     Call SELECHO("!INDIPEP/Indice De Pedidos .")
     NUMEPEDI = XVALO(VALACT1$("!INDIPEP"))
     If XVALO(NUMEPEDI) <= 0 Then GoTo 20
     NUMEP& = NUMEPEDI
     '
     JJ& = 0
     Call ABRESTOCKS
     SQLX$ = "SELECT  DISTINCT DOSECLAR ,FechMov  FROM MOVISTO "
     SQLX$ = SQLX$ + " WHERE NuPedCli = " & NUMEP& & ""
     SQLX$ = SQLX$ + " AND CODIEMPR = " & CodiEmp% & " "
     SQLX$ = SQLX$ + " ORDER BY DOSECLAR ASC "
     'Set MOVISTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     Dim MOVISTMP As ADODB.Recordset
     Set MOVISTMP = New ADODB.Recordset
26   On Error Resume Next
     MOVISTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
        GoTo 26
     End If
     On Error GoTo 0

     With MOVISTMP
      On Error Resume Next
      .MoveFirst
      If Err > 0 Then
       Call MENSERR(24, "No se Encuentran Despachos en el Rango Elegido ")
       GoTo 30
      End If
      Y$ = REGBLAN("INDICOM")
      Call BLANARCH("!INDICOM")
      On Error GoTo 0
      Do While Not .EOF
        '***
        'CARGAR EL ARCHIVO PARA SELECCION
        DSLAR$ = SAFETEXT$(!DOSECLAR)
        FEC% = CVINT%(!FechMov)
        FEMO = FECHATEX$(FEC%)
        If InStr(DSLAR$, "P/F") < 1 Then    ' DSLAR$ = " Pendientes de Facturación "
           JJ& = JJ& + 1
           ORDEN% = JJ&
           Call REPLA(Y$, MKI(ORDEN%), 1, 2)
           Call REPLA(Y$, DSLAR$ & " - " & FEMO, 3, 36)
           Call GRAREG("!INDICOM", Y$, JJ&)
        End If
        .MoveNext
      Loop
     End With
     MOVISTMP.Close
     Set MOVISTMP = Nothing
     '
     Call SETULTREG("!INDICOM", JJ&)
     Call FRESHECHO("!INDICOM")
40   Call SELECHO("!INDICOM")
     NFACT$ = TRIM(VALACT2$("!INDICOM"))
     NFACTU$ = Mid(NFACT$, 9, 8)
     If NFACTU$ = "" Then GoTo 30
     RAZSOCIAL$ = TRIM$(Left$(rasocli$(XVALO(NUMCLIEN$)), 8))
     
     Call VerFactElec(NFACTU$, RAZSOCIAL$) ' Muestra formulario electronico
     GoTo 40
   End If
   '\\\OT-08-0200-OD-FIN///
End Sub

Sub VerFactElec(NUDOCU$, RACLI$)
   Screen.MousePointer = 11
   For UI& = ULTREG&("PDOCLST") To 1 Step -1
     XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
     If InStr(XX$, "Factura") > 4 Then
       If InStr(XX$, NUDOCU$) > 4 Then
         If InStr(XX$, RACLI$) > 16 Then
           On Error Resume Next
           DOCUNU& = CVS(Left$(XX$, 4))
           If Err Then
              NUERRO = Err.Number
              On Error GoTo 0
              Call MENSERR(24, "Se ha Producido Error " & NUERRO & " (1).\Imposible Abrir Documento en este Momento.\Base de Datos de Documentos Digitalizados\Probablemente Dañada. Consulte.")
              Exit Sub
           End If
           On Error GoTo 0
           '
           GoTo 5
         End If
       End If
     End If
   Next UI&
   Screen.MousePointer = 1
   Call COMUNI("Imposible Abrir Documento Desde esta Aplicación")
   Exit Sub
   '
5  For KKI% = 0 To 5
     EDITFORM.Picture1(KKI%).Cls
     EDITFORM.Picture1(KKI%).DrawWidth = 3
     EDITFORM.Picture1(KKI%).DrawStyle = 0
     EDITFORM.Picture1(KKI%).FillStyle = 1
   Next KKI%
   '
   REDOCU$ = REGACT$("PDOCLST")
   URE& = ULTREG&("PDOCSPL")
   LI& = 0: LS& = URE&
6  E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
   L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
   On Error Resume Next
   DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
   If Err Then
      NUERRO = Err.Number
      On Error GoTo 0
      Call MENSERR(24, "Se ha Producido Error " & NUERRO & " (2)\al Intentar Leer el Registro " & L& & " de 'PDOCSPL'.\Imposible Abrir Documento en este Momento.\Base de Datos de Documentos Digitalizados\Probablemente Dañada. Consulte.")
      Exit Sub
   End If
   On Error GoTo 0
   '
   If DCC& < DOCUNU& Then LI& = L&: GoTo 6
   If DCC& > DOCUNU& Then LS& = L&: GoTo 6
   GoTo 8
   '
7  Screen.MousePointer = 1
   Call MENSERR(24, "Documento no Registrado")
   Exit Sub
   '
8  TPSP$ = ""
   Screen.MousePointer = 11
   While L& > 1
     On Error Resume Next
     DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
     If Err Then
        NUERRO = Err.Number
        On Error GoTo 0
        Call MENSERR(24, "Se ha Producido Error " & NUERRO & " (3)\al Intentar Leer el Registro " & DCC& & " de 'PDOCSPL'.\Imposible Abrir Documento en este Momento.\Base de Datos de Documentos Digitalizados\Probablemente Dañada. Consulte.")
        Exit Sub
     End If
     On Error GoTo 0
     '
     If DCC& < DOCUNU& Then GoTo 9
     L& = L& - 1
   Wend
   '
9  CAPAGINAS% = 0
   For KKL& = L& To URE&
     TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
     If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
     If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
       TPSP$ = TPSP$ + Mid$(TBUF$, 5)
     End If
   Next KKL&
   '
95 SPOOLSTRING$ = TPSP$
   For KKI% = 0 To 5
     EDITFORM.Picture1(KKI%).Cls
     EDITFORM.Picture1(KKI%).Top = 0
     EDITFORM.Picture1(KKI%).Refresh
     EDITFORM.Picture1(KKI%).Height = 7170
     EDITFORM.Picture1(KKI%).Width = 11700
     EDITFORM.Picture1(KKI%).DrawWidth = 3
     EDITFORM.Picture1(KKI%).DrawStyle = 0
     EDITFORM.Picture1(KKI%).FillStyle = 1
   Next KKI%
   ALTUPAGINA = 7170
   TOPEPAGINA = 0
   PAGINACTIVA% = 0
   CAPAGINAS% = 0
   '
10 PPXX% = InStr(TPSP$, Chr$(255))
   If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call SHOWCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
   End If
   '
   Screen.MousePointer = 1
   PAGINACTIVA% = 0
   For KKI% = 0 To 5
     EDITFORM.Picture1(KKI%).Visible = False
     If KKI% = PAGINACTIVA% Then EDITFORM.Picture1(KKI%).Visible = True
   Next KKI%
   EDITFORM.Picture1(PAGINACTIVA%).Refresh
   EDITFORM.VScroll1.Min = 0
   EDITFORM.VScroll1.Value = 0
   EDITFORM.VScroll1.Max = 1
   If ALTUPAGINA > EDITFORM.Frame1.Height Then
      EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
   End If
   EDITFORM.Command3.Enabled = False
   'EDITFORM.Command5.Enabled = False
   EDITFORM.Show 1
   '
End Sub

Private Sub Command26_Click_ORIGINAL()
   '
   Command26.Enabled = False
   '
10 Call DESHASFECHA(Des%, Has%, PERIO$)
   If PERIO$ = "" Then GoTo 99
   '
   Screen.MousePointer = 11
    ' INDICE DE CUMPLIMIENTO DE CLIENTES
   FIN& = ULTREG&("MOVISTO")
   JJ& = 0: PRIREMI& = 9999999: ULTREMI& = 0
   REBLA$ = REGBLAN$("COCUENT")
   '
   DESDEX = FETEXCOR1$(Des%)
   HASTAX = FETEXCOR1$(Has%)
   SQLX$ = " SELECT * FROM MOVISTO WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE FECHMOV >= '" & DESDEX & "'"
   SQLX$ = SQLX$ + " AND FECHMOV <= '" & HASTAX & "'"
   SQLX$ = SQLX$ + " AND DoPriCor LIKE 'RT%'"
   '
   Dim MOVVVISTO As ADODB.Recordset
   Set MOVVVISTO = New ADODB.Recordset
25 On Error Resume Next
   MOVVVISTO.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   With MOVVVISTO
     On Error Resume Next
     If Not (.EOF And .BOF) Then
       On Error GoTo 0
       Do While (Not .EOF)
         FF% = CVINT(!FechMov)
         DOCUPRI$ = UCase$(SAFETEXT$(!DoPriCor))
         CLI% = XVALO(!NUME_CLIE)
         If CLI% > 0 Then
            CAN = XVALO(!Cantsalid) - XVALO(!CANTINGRE)
            If Abs(CAN) > 0.05 Then
               CI% = XVALO(!cod_inte)
               Z0$ = REBLA$
               Call REPLA(Z0$, MKI(CLI%), 1, 2)
               Call REPLA(Z0$, MKS$(-1), 3, 4)
               Call REPLA(Z0$, MKI(CI%), 7, 2)
               Call REPLA(Z0$, MKS(CAN), 15, 4)
               Call REPLA(Z0$, MKI(FF%), 19, 2)
               Call REPLA(Z0$, DOCUPRI$, 21, 16)
               JJ& = JJ& + 1
               Call GRAREG("COCUENT", Z0$, JJ&)
               NREMITO& = XVALO(Mid$(Z0$, 24, 7))
               If NREMITO& > ULTREMI& Then ULTREMI& = NREMITO&
               If NREMITO& < PRIREMI& Then PRIREMI& = NREMITO&
            End If
         End If
       .MoveNext
       Loop
     End If
   End With
   '
   Call SETULTREG("COCUENT", JJ&)
   Call CIERRARCH("COCUENT")
   '
' ESTA PARTE ES CODIGO NUEVO PARA LO QUE SALE PRIMERO CON REMITO Y LUEGO CON FACTURA
   FIN& = JJ&
   For JJ& = 1 To FIN&
      Call TRACE(20, JJ&, FIN&)
      Z0$ = REGLEIDO$("COCUENT", JJ&)
      CAN = CVS(Mid$(Z0$, 15, 4))
      If CAN < (-0.05) Then
         CLI% = CVI(Mid$(Z0$, 1, 2))
         CI% = CVI(Mid$(Z0$, 7, 2))
         DOCUPRI$ = TRIM$(Mid$(Z0$, 21, 16))
         For JJ1& = JJ& + 1 To FIN&
            Z1$ = REGLEIDO$("COCUENT", JJ1&)
            If CVI(Mid$(Z1$, 1, 2)) = CLI% Then
               If CVI(Mid$(Z1$, 7, 2)) = CI% Then
                  If TRIM$(Mid$(Z1$, 21, 16)) = DOCUPRI$ Then
                     CANX = CVS(Mid$(Z1$, 15, 4)) + CAN
                     Call REPLA(Z1$, MKS$(CANX), 15, 4)
                     Call GRAREG("COCUENT", Z1$, JJ1&)
                     Call REPLA(Z0$, MKS$(0), 15, 4)
                     Call GRAREG("COCUENT", Z0$, JJ&)
                     GoTo 39
                  End If
               End If
            End If
         Next JJ1&
      End If
39 Next JJ&
' FIN DE LA PARTE QUE SALE PRIMERO CON REMITO Y DESPUES CON FACTURA

' ELIMINA MOVIMIENTOS NULOS
   JJ& = 0
   For U& = 1 To ULTREG&("COCUENT")
     Z0$ = REGLEIDO$("COCUENT", U&)
     If CVS(Mid$(Z0$, 15, 4)) > 0.05 Then
       JJ& = JJ& + 1
       Call GRAREG("COCUENT", Z0$, JJ&)
     End If
   Next U&
   Call SETULTREG("COCUENT", JJ&)
   Call CIERRARCH("COCUENT")
   '
   'CANTIDAD DE REGISTROS PARA BARRA TRAZADORA
   CONDI$ = " PEDDETA.NroPed > 0 AND  PEDENCA.Status < 9 AND PEDDETA.Status <> 9"
   CONDI$ = CONDI$ + " AND PEDENCA.CODIEMPR = " & CodiEmp%
   CONDI$ = CONDI$ + " AND  PEDDETA.CODIEMPR = " & CodiEmp% & ""
   FIN& = CantRegistros&("PEDDETA INNER JOIN PEDENCA ON PEDDETA.NroPed = PEDENCA.NroPed ", CONDI$)
   '
   SQLX$ = "SELECT *, PEDDETA.FESOLENT AS FESOL FROM PEDDETA WITH(NOLOCK)"
   SQLX$ = SQLX$ + " INNER JOIN PEDENCA"
   SQLX$ = SQLX$ + " ON PEDDETA.NroPed = PEDENCA.NroPed"
   SQLX$ = SQLX$ + " WHERE PEDDETA.NroPed > 0 "
   SQLX$ = SQLX$ + " AND PEDENCA.Status < 9 "                ' NO ANULADOS
   SQLX$ = SQLX$ + " AND PEDDETA.Status <> 9 "
   SQLX$ = SQLX$ + " AND PEDENCA.CODIEMPR = " & CodiEmp% & " "
   SQLX$ = SQLX$ + " AND PEDDETA.CODIEMPR = " & CodiEmp% & " "
   'SQLX$ = SQLX$ + " AND PEDDETA.NroRemi >= " & PRIREMI&
   'SQLX$ = SQLX$ + " AND PEDDETA.NroRemi <= " & ULTREMI&
   Dim IPEPE As ADODB.Recordset
   Set IPEPE = New ADODB.Recordset
   Set IPEPE = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   II& = 0
   With IPEPE
      Do While Not .EOF
         II& = II& + 1
         Call TRACE(20, II&, FIN&)
         NREMITO& = XVALO(Mid$(!NroRemi, 11))
         If NREMITO& >= PRIREMI& Then
           If NREMITO <= ULTREMI& Then
              CLI% = XVALO(!NROCLIE)
              CII% = XVALO(!CODIINT)
              CANOC& = XVALO(!CanPed)
              FESOLE% = CVINT(!FESOL)
              FEPEDI% = CVINT(!FECHEMIS)
              If FESOLE% < FEPEDI% Then FESOLE% = FEPEDI%
              NOC& = XVALO(!NROPED)
              '
              RFF$ = RECFILT$("COCUENT", 1, MKI$(CLI%))
              For KKL& = 1 To Len(RFF$) Step 4
                  U& = CVS(Mid$(RFF$, KKL&, 4))
                  X$ = REGLEIDO$("COCUENT", U&)
                  CANENT& = CVS(Mid$(X$, 15, 4))
                  CANMA& = CANOC&: If CANMA& > CANENT& Then CANMA& = CANENT&
                  If CVI(Mid$(X$, 7, 2)) = CII% Then
                     If CVS(Mid$(X$, 3, 4)) < 1 Then ' NO TIENE ASIGNADO PEDIDO
                        If Val(Mid$(X$, 24, 7)) = NREMITO& Then
                           FR% = CVI(Mid$(X$, 19, 2))
                           da% = DIENTRE(FESOLE%, FR%)
                           Call REPLA(X$, MKS$(NOC&), 3, 4)
                           Call REPLA(X$, MKS(CANOC), 9, 4)
                           Call REPLA(X$, MKI(FESOLE%), 13, 2)
                           If da% > 0 Then
                              Call REPLA(X$, MKI(da%), 37, 2)
                           End If
                           Call GRAREG("COCUENT", X$, U&)
                           GoTo 37
                        End If
                     End If
                  End If
              Next KKL&
            End If
          End If
37    .MoveNext
      Loop
   End With
   '
   Call FILESORT("COCUENT", "COCUENT", 1, 3, "ASC")
   Call HEADERS("COCUENT", "")
   Call HEADERS("COCUENT", PERIO$)
   Call FINAL("*LICUENT")
   Screen.MousePointer = 1
   '
99 Command26.Enabled = True
End Sub


Private Sub Command26_Click()
   '
   Command26.Enabled = False
   '
10 Call DESHASFECHA(Des%, Has%, PERIO$)
   If PERIO$ = "" Then GoTo 99
   '
   Screen.MousePointer = 11
    ' INDICE DE CUMPLIMIENTO DE CLIENTES
   FIN& = ULTREG&("MOVISTO")
   JJ& = 0: PRIREMI& = 9999999: ULTREMI& = 0
   REBLA$ = REGBLAN$("COCUENT")
   '
   DESDEX = FETEXCOR1$(Des%)
   HASTAX = FETEXCOR1$(Has%)
   SQLX$ = " SELECT SUM(CANTINGRE) AS CAIN, SUM(CANTSALID) AS CASA, COD_INTE, CODIMOVI, NUPEDCLI, NUME_CLIE FROM MOVISTO WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE FECHMOV >= '" & DESDEX & "'"
   SQLX$ = SQLX$ + " AND FECHMOV <= '" & HASTAX & "'"
   SQLX$ = SQLX$ + " AND CODIMOVI='PP' GROUP BY CODIMOVI, NUPEDCLI, COD_INTE, NUME_CLIE"
   '
   Dim MOVVVISTO As ADODB.Recordset
   Set MOVVVISTO = New ADODB.Recordset
25 On Error Resume Next
   MOVVVISTO.Open FILTSQL$(SQLX$), FLEXCONN, adOpenForwardOnly, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   With MOVVVISTO
     On Error Resume Next
     If Not (.EOF And .BOF) Then
       On Error GoTo 0
       Do While (Not .EOF)
         DOCUPRI$ = UCase$(SAFETEXT$(!NUPEDCLI))
         CLI% = XVALO(!NUME_CLIE)
         CONDI$ = "FECHMOV >= '" & DESDEX & "' AND FECHMOV <= '" & HASTAX & "'" & _
         " AND NUPEDCLI=" & DOCUPRI$ & " AND NUME_CLIE=" & CLI% & " ORDER BY FECHMOV DESC"
         FF% = CVINT(VALOBADA("MOVISTO", "FECHMOV", CONDI$))
         If CLI% > 0 Then
            CAN = XVALO(!CAIN) - XVALO(!CASA)
            If Abs(CAN) > 0.05 Then
               CI% = XVALO(!cod_inte)
               Z0$ = REBLA$
               Call REPLA(Z0$, MKI(CLI%), 1, 2)
               Call REPLA(Z0$, MKS$(-1), 3, 4)
               Call REPLA(Z0$, MKI(CI%), 7, 2)
               Call REPLA(Z0$, MKS(CAN), 15, 4)
               Call REPLA(Z0$, MKI(FF%), 19, 2)
               Call REPLA(Z0$, DOCUPRI$, 21, 16)
               JJ& = JJ& + 1
               Call GRAREG("COCUENT", Z0$, JJ&)
            End If
         End If
       .MoveNext
       Loop
     End If
   End With
   '
   Call SETULTREG("COCUENT", JJ&)
   Call CIERRARCH("COCUENT")
   '
' ELIMINA MOVIMIENTOS NULOS
   JJ& = 0
   For U& = 1 To ULTREG&("COCUENT")
     Z0$ = REGLEIDO$("COCUENT", U&)
     If CVS(Mid$(Z0$, 15, 4)) > 0.05 Then
       JJ& = JJ& + 1
       Call GRAREG("COCUENT", Z0$, JJ&)
     End If
   Next U&
   Call SETULTREG("COCUENT", JJ&)
   Call CIERRARCH("COCUENT")
   '
    FIN& = ULTREG&("COCUENT")
    For KKL& = 1 To ULTREG&("COCUENT")
        X$ = REGLEIDO$("COCUENT", KKL&)
        Call TRACE(24, KKL&, FIN&)
        CANENT& = CVS(Mid$(X$, 15, 4))
        NOC& = Val(Mid$(X$, 21, 16))
        CII% = CVI(Mid$(X$, 7, 2))
        CONDI$ = "NROPED=" & NOC& & " AND CODIINT=" & CII% & " AND Status < 8"
        If CantRegistros("PEDDETA", CONDI$) > 0 Then
            CANOC& = XVALO(VALOBADA("PEDDETA", "CANPED", CONDI$))
            FESOLE% = CVINT(VALOBADA("PEDDETA", "FESOLENT", CONDI$))
            CANMA& = CANOC&
            If CANMA& > CANENT& Then CANMA& = CANENT&
            FR% = CVI(Mid$(X$, 19, 2))
            da% = DIENTRE(FESOLE%, FR%)
            Call REPLA(X$, MKS$(NOC&), 3, 4)
            Call REPLA(X$, MKS(CANOC&), 9, 4)
            Call REPLA(X$, MKI(FESOLE%), 13, 2)
            If da% > 0 Then
               Call REPLA(X$, MKI(da%), 37, 2)
            End If
            Call GRAREG("COCUENT", X$, KKL&)
        End If
   Next KKL&
   '
40 Call FILESORT("COCUENT", "COCUENT", 1, 3, "ASC")
   Call HEADERS("COCUENT", "")
   Call HEADERS("COCUENT", "Periodo: " & PERIO$)
   Call FINAL("*LICUENT")
   Screen.MousePointer = 1
   '
99 Command26.Enabled = True
End Sub


Sub ANUPEDMAS()
   If DERACCE%("DEPUPED", "Depuracion de Pedidos") >= 2 Then
10  Call DESHASFECHA(Des%, Has%, PERIO$)
    Call ABREPEDCLI
    If Des% > 0 And Has% > 0 Then
      Desde$ = FETEXCOR1(Des%)
      Hasta$ = FETEXCOR1(Has%)
      CONDI$ = "STATUS < 3 "
      CONDI$ = CONDI$ + " AND (FECHEMIS >='" & Desde$ & "' "
      CONDI$ = CONDI$ + " AND FECHEMIS <= '" & Hasta$ & "') "
      CONDI$ = CONDI$ + " AND (CANTENT < CANPED)"
      '
      TXT$ = "Realmente Desea Anular Todos Los Pedidos \Dentro del Rango de Fecha Seleccionado"
      'AGREGADO PARA QUE FILT RE POR CLIENTE , SI NO ELIGE LO HACE PARA TODOS
      OBX$ = OBJPLA$("SELECLI", RECORD$)
      If OBX$ <> "" Then
         RECORD$ = OBX$
         NC% = CVI(Mid$(RECORD$, 1, 2))
         If NC% > 0 Then
            CONDI$ = CONDI$ + " AND NROCLIE = " & NC%
            TXT$ = TXT$ + "\Del Cliente: " & rasocli$(NC%)
         End If
      End If
      '
      If COMALTER(TXT$ & "\\Cancelar\Anular Pedidos") = 2 Then
        Call ACTUREGISTRO("PEDDETA", "STATUS", CONDI$, 8, REG&)
        Call COMUNI("Pedidos Anulados Correctamente")
      End If
      '
      GoTo 10 'vuelve a selfecha
    End If
   End If
End Sub

Sub CAMBIAFECHAEQUIPO()

    If DERACCE%("CAMFERE", "Cambio Fecha de Regreso de Equipo") >= 2 Then
        Call CARGALISEL("INDIPEDI", "PEDENCA", "NROPED/F6.0 ; REFERENCIA/A60", "NROPED > 0 AND (OV_MOTIVO > 0 AND OV_MOTIVO <=3) ORDER BY NROPED DESC")
5       Call SELECHO("INDIPEDI/Indice General de Pedidos con Motivo")
        NPED$ = TRIM$(VALACT1$("INDIPEDI"))
        
        ' VALIDA SI ELIGE PEDIDO
        If NPED$ <> "" Then
            CONDI$ = "NROPED=" & NPED$
            FECHAREGRESO = VALOBADA("PEDENCA", "OV_FEREGRESO", CONDI$, "NOMESS")
            FF% = CVINT%(FECHAREGRESO)
            ' ESTABLECE VALORES POR DEFAULT
            VDEF$ = MKI$(FF%) & MKI$(FECHANUM(HOY(HO$)))
            
            ' PRESENTA VENTANA PARA EL CAMBIO DE FECHA
10          OBX$ = OBJPLA$("CAMFEREGRE", VDEF$)
            
            ' VALIDA QUE NO HAYA CANCELADO
            If OBX$ <> "" Then
                FF1% = CVI(Mid$(OBX$, 3, 2))
                ' VERIFICA SI APRUEBA SIN INGRESAR UNA NUEVA FECHA DE REGRESO
                If FF1% = 0 Then GoTo 5
                ' VERIFICA SI LA FECHA ES MENOR A LA ANTERIOR. AVISA Y CONSULTA SI CONTINUA
                If FF1% < FF% Then
                    If COMALTER%("La Fecha Ingresada es Menor a la Anterior\¿Desea Continuar con la Grabacion?\\No, Abandonar\Si, Continuar") = 1 Then GoTo 5
                End If
                ' ACTUALIZA FECHA DE REGRESO DEL EQUIPO
                valor$ = FETEXCOR1$(FF1%)
                Call ACTUREGISTRO("PEDENCA", "OV_FEREGRESO", CONDI$, valor$, REG&)
            End If
            GoTo 5
        End If
    End If
End Sub

Private Sub mnuArchivo_Click()

End Sub
