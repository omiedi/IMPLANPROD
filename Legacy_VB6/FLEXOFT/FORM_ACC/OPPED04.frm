VERSION 5.00
Begin VB.Form OPPED04 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pedidos de Clientes - Modificaciones, Consultas y Listados"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5940
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
   ScaleHeight     =   6510
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Generación de Pedidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   210
      TabIndex        =   41
      Top             =   2940
      Width           =   4635
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Index           =   8
         Left            =   1050
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   42
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
            Height          =   600
            Left            =   0
            Picture         =   "OPPED04.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Copia de Pedido"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Generación por Copia de Pedido"
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
            TabIndex        =   44
            Top             =   80
            Width           =   1530
         End
      End
   End
   Begin VB.CommandButton Command9 
      Caption         =   "List"
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
      Height          =   700
      Left            =   4830
      Picture         =   "OPPED04.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Asignación de Materiales - Ordenes de Despacho "
      Top             =   7245
      Width           =   650
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000016&
      Height          =   645
      Index           =   3
      Left            =   105
      ScaleHeight     =   585
      ScaleWidth      =   2370
      TabIndex        =   37
      Top             =   7140
      Width           =   2430
      Begin VB.CommandButton Command43 
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
         Picture         =   "OPPED04.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Revertir Anulación de Pedidos"
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Revertir Anula- ción de Pedidos "
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
         Top             =   50
         Width           =   1425
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004040&
      Height          =   6735
      Left            =   5040
      ScaleHeight     =   6675
      ScaleWidth      =   1110
      TabIndex        =   28
      Top             =   0
      Width           =   1170
      Begin VB.CommandButton Command8 
         Caption         =   "More"
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
         Left            =   105
         Picture         =   "OPPED04.frx":059E
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Listados complementarios"
         Top             =   3360
         Width           =   650
      End
      Begin VB.CommandButton Command14 
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "OPPED04.frx":06E8
         TabIndex        =   50
         ToolTipText     =   "Recuperación de Precios de Venta"
         Top             =   2520
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Stk."
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
         Left            =   105
         Picture         =   "OPPED04.frx":09F2
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Listados Estadísticos de Ventas y Pedidos"
         Top             =   4740
         Width           =   650
      End
      Begin VB.CommandButton Command10 
         Caption         =   "List"
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
         Left            =   105
         Picture         =   "OPPED04.frx":0CFC
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Histórico de Precios por Pedido"
         Top             =   4050
         Width           =   650
      End
      Begin VB.ListBox List1 
         Height          =   300
         Left            =   105
         Sorted          =   -1  'True
         TabIndex        =   46
         Top             =   5145
         Visible         =   0   'False
         Width           =   645
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
         Left            =   525
         Picture         =   "OPPED04.frx":1006
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Control de Integridad de la B.Datos"
         Top             =   2520
         Visible         =   0   'False
         Width           =   650
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
         Height          =   700
         Left            =   105
         Picture         =   "OPPED04.frx":1310
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Salir"
         Top             =   210
         Width           =   650
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
         Height          =   700
         Left            =   105
         Picture         =   "OPPED04.frx":145A
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Ayuda"
         Top             =   900
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   30
         Top             =   5565
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   31
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Direct"
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
         Left            =   105
         Picture         =   "OPPED04.frx":1764
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Accesos Directos"
         Top             =   1600
         Width           =   650
      End
      Begin VB.Image Image3 
         Height          =   390
         Left            =   -315
         Picture         =   "OPPED04.frx":1A6E
         Top             =   5910
         Width           =   1515
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Consultas y Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   210
      TabIndex        =   8
      Top             =   4305
      Width           =   4635
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   3735
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
            Height          =   600
            Left            =   3150
            Picture         =   "OPPED04.frx":3990
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
            Height          =   600
            Left            =   2550
            Picture         =   "OPPED04.frx":3C9A
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
            Height          =   600
            Left            =   0
            Picture         =   "OPPED04.frx":40DC
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
            Height          =   600
            Left            =   1950
            Picture         =   "OPPED04.frx":43E6
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
            TabIndex        =   17
            Top             =   105
            Width           =   1485
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Dispositivo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
            BackColor       =   &H00C0FFFF&
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
            Left            =   525
            TabIndex        =   11
            Top             =   270
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0FFFF&
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
            Left            =   2205
            TabIndex        =   10
            Top             =   210
            Width           =   1170
         End
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Control"
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
      Left            =   3255
      Picture         =   "OPPED04.frx":46F0
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6930
      Width           =   855
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Base de Datos de Pedidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2730
      Left            =   210
      TabIndex        =   0
      Top             =   120
      Width           =   4635
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Index           =   4
         Left            =   1575
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   34
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
            Height          =   600
            Left            =   -20
            Picture         =   "OPPED04.frx":49FA
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Aprobar Pedidos de Clientes."
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Aprobar  Pedidos de Clientes."
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
            TabIndex        =   36
            Top             =   105
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Index           =   1
         Left            =   1260
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   25
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
            Height          =   600
            Left            =   0
            Picture         =   "OPPED04.frx":4D04
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Anulación de Pedidos"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Anulación por Items de Pedidos "
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
            Left            =   720
            TabIndex        =   27
            Top             =   50
            Width           =   1545
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Index           =   2
         Left            =   945
         ScaleHeight     =   585
         ScaleWidth      =   2370
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
            Picture         =   "OPPED04.frx":4E4E
            Style           =   1  'Graphical
            TabIndex        =   49
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
            Height          =   600
            Left            =   0
            Picture         =   "OPPED04.frx":4F98
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Anulación de Pedidos"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Anulación       de Pedidos "
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
            TabIndex        =   22
            Top             =   45
            Width           =   1425
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Index           =   0
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   18
         Top             =   400
         Width           =   2430
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
            Height          =   600
            Left            =   0
            Picture         =   "OPPED04.frx":52A2
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
            TabIndex        =   19
            Top             =   45
            Width           =   1530
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
   Begin VB.Image Image4 
      Height          =   435
      Left            =   2415
      Top             =   3045
      Width           =   1170
   End
   Begin VB.Image Image2 
      Height          =   435
      Left            =   2415
      Top             =   3045
      Width           =   1170
   End
   Begin VB.Image Image1 
      Height          =   435
      Left            =   2415
      Top             =   3045
      Width           =   1170
   End
   Begin VB.Line Line1 
      X1              =   5250
      X2              =   6120
      Y1              =   5280
      Y2              =   5280
   End
End
Attribute VB_Name = "OPPED04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




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
       Call DESHASFECHA(DES%, HAS%, PERIO$)
       If PERIO$ = "" Then GoTo 99
          DESDAT = CDbl(CVDAT(DES%))
          HASDAT = CDbl(CVDAT(HAS%))
       '
       Call ABREPEDCLI
       SQLX$ = "SELECT * FROM PEDDETA "
       SQLX$ = SQLX$ + "WHERE INT(CDBL(FechEmis)) >= " & DESDAT & " "
       SQLX$ = SQLX$ + "AND INT(CDBL(FechEmis)) <= " & HASDAT & " "
       '
       On Error GoTo 0
       Set PEDDETTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
       '
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
           NC% = !NroClie
           If NC% > 0 Then
              NART% = !CodiInt
              FEDOC% = CVINT(!FECHEMIS)
              NPED& = !NroPed
              ORCOM$ = !NroOcom
              CANTI# = !CANPED
              PUNIT# = !PREUNID
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
              Call REPLA(Z$, MKI$(NC%), 1, 2)
              Call REPLA(Z$, MKI$(NART%), 3, 2)
              Call REPLA(Z$, MKI$(FEDOC%), 5, 2)
              Call REPLA(Z$, MKS$(NPED&), 7, 4)
              Call REPLA(Z$, ORCOM$, 11, 16)
              PRECONV# = PUNIT#
              If !MoneEmis > 1 Then
                 '28/03/07 (OT-07-0092)
                 PRECONV# = PUNIT# * !TipoCam
                 '28/03/07 (FIN)
              End If
              Call REPLA(Z$, MKD$(PRECONV#), 27, 8)
              Call REPLA(Z$, MKS$(CANTI#), 35, 4)
              IMPO# = CANTI# * PRECONV#
              Call REPLA(Z$, MKD$(IMPO#), 39, 8)
              Call REPLA(Z$, PREORIG$, 47, 20)
              '
              X& = X& + 1
              Call GRAREG("HISPRECI", Z$, X&)
              '
            End If
            '
            .MoveNext
         Loop
        End With
   End If
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
   '
99 Screen.MousePointer = 1
   Command10.Enabled = True
   '\\\OT-07-0092-RG-FIN///
   '
End Sub

Sub Command11_Click()
   Command11.Enabled = False
   SQLX$ = "SELECT * FROM PEDENCA "
   If PROLLAMA$ = "ABIE" Then
        SQLX$ = SQLX$ + "WHERE TipoPed = 'ABIERTO' "
      Else
        SQLX$ = SQLX$ + "WHERE TipoPed <> 'ABIERTO' "
        SQLX$ = SQLX$ + "OR IsNull(TipoPed) = True "
   End If
   SQLX$ = SQLX$ + "ORDER BY NroPed ASC"
   Set PEDENCTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   JJ& = 0
   With PEDENCTEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        Do While Not .EOF
           'CODIGO ANTIGUO A OT-05-0682-WAR-18/10/05
'          NPX& = !NroPed
'          TTXX$ = Space$(64)
'             TPX$ = "      ": If !Status = 9 Then ANULX$ = "(Anul)"
'          Call REPLA(TTXX$, TPX$, 1, 6)
'             FEX = CVINT(!FECHEMIS)
'          Call REPLA(TTXX$, FECHATEX$(FEX), 9, 8)
'             NCLIE% = !NroClie
'          Call REPLA(TTXX$, RASOCLI$(NCLIE%), 19, 42)
'          YY$ = MKS$(NPX&) + TTXX
'          JJ& = JJ& + 1
'          Call GRAREG("!INDIPEDI", YY$, JJ&)
            '\\\OT-05-0682-WAR-18/10/05///
             NRODES& = !NroPed
             TTXX$ = MKS$(NRODES&)
             '\\\OT-05-0829-WAR-17/11/05///
               'NOCO$ = !NroOcom
               NOCO$ = TRIM$(!NroOcom)
             '\\\OT-05-0829-WAR-FIN///
             If NOCO$ <> "" Then
               TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
             End If
               FEX = CVINT(!FECHEMIS)
             TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
               NCLIE% = !NroClie
               ANULX$ = "": If !Status = 9 Then ANULX$ = "(Anul) "
             TTXX$ = TTXX$ + TRIM$(RASOCLI$(NCLIE%)) + ANULX$ 'Space$(52)
             JJ& = JJ& + 1
             Call GRAREG("!INDIPEDI", TTXX$, JJ&)
             '\\\OT-05-0682-WAR-FIN///
        .MoveNext
        Loop
      End If
   End With
   '
   Call SETULTREG("!INDIPEDI", JJ&)
   Call CIERRARCH("!INDIPEDI")
   Call FRESHECHO("!INDIPEDI")
   Screen.MousePointer = 1
   '
   Dim FORMORIG As Form
   Set FORMORIG = FORPED04
   If PROLLAMA = "ABIE" Then Set FORMORIG = PEDABIE04
   '
   Call SELECHO("!INDIPEDI/INDICE GENERAL DE PEDIDOS")
   NPSEL& = XVALO(VALACT1$("!INDIPEDI"))
   If NPSEL& > 0 Then
      SQLX$ = "SELECT * FROM PEDENCA "
      SQLX$ = SQLX$ + "WHERE NroPed = " & NPSEL& & " "
      Set PEDENCTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
      '
      With FORMORIG
        Call .BLANFORMU
        NCLIE% = PEDENCTEMP!NroClie
        .Text1 = NCLIE%
        Call .CARDACLI
        '
        .TEXT21(0) = TRIM$(Str$(PEDENCTEMP!ListPre))
        .TEXT21(1) = TRIM$(Str$(PEDENCTEMP!CondPag))
        .TEXT21(2) = TRIM$(Str$(PEDENCTEMP!VENDED))
        .LENTREGA(0) = TRIM$(Str$(PEDENCTEMP!DOMIENT))
        .LENTREGA(1) = TRIM$(Str$(PEDENCTEMP!LOCAENT))
        .LENTREGA(2) = TRIM$(Str$(PEDENCTEMP!TranspEnt))
        .Text10 = PEDENCTEMP!NroOcom
        .Text13 = PEDENCTEMP!Observa
        '\\\OT-05-0812-WAR-17/11/05///
        .PORDESCU = TRIM$(PEDENCTEMP!TDescApli)
        '\\\OT-05-0812-WAR-FIN///
        '
        JJ& = 0
        REBLA$ = REGBLAN$("CARDETPE")
        SQLX$ = "SELECT * FROM PEDDETA "
        SQLX$ = SQLX$ + "WHERE NroPed = " & NPSEL& & " "
        SQLX$ = SQLX$ + "ORDER BY NuorItem ASC "
        Set PEDDETTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
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
              CIXI% = PEDDETTEMP!CodiInt
            Call REPLA(X$, MKI$(CIXI%), 1, 2)
            Call REPLA(X$, PEDDETTEMP!TColor, 3, 16)
            Call REPLA(X$, PEDDETTEMP!TTalle, 19, 8)
              TLAR = PEDDETTEMP!TLargo
            Call REPLA(X$, MKS$(TLAR), 27, 4)
              TCHO = PEDDETTEMP!TAncho
            Call REPLA(X$, MKS$(TCHO), 31, 4)
              TESP = PEDDETTEMP!TEspesor
            Call REPLA(X$, MKS$(TESP), 35, 4)
              PREL = PEDDETTEMP!PreList
            Call REPLA(X$, MKS$(PREL), 39, 4)
              PREU = PEDDETTEMP!PREUNID
            Call REPLA(X$, MKS$(PREU), 51, 4)
              CAN = PEDDETTEMP!CANPED
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
End Sub

Sub Command12_Click()
    Command12.Enabled = False
    
    Dim TOTVENTA#(32767)
    Dim TOTIVA#(32767)
    '
    SELVENDE.BackColor = &HFFFFC0
    SELVENDE.Frame6.BackColor = &HFFFFC0
10  SELVENDE.Show 1
    If SELVENDE.Label2 <> "OK" Then GoTo 99
    '
    OPT% = COMALTER%("Opciones de Listado\\Estadística de Facturación\Estadística de Pedidos")
    If OPT% = 0 Then GoTo 99
    '
    Screen.MousePointer = 11
    CANSELE% = XVALO(SELVENDE.Label5)
    DES% = FECHANUM(SELVENDE.Label7)
    HAS% = FECHANUM(SELVENDE.Label3)
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    DESDAT = Int(CDbl(CVDAT(DES%)))
    HASDAT = Int(CDbl(CVDAT(HAS%)))
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
            TOTIVA#(U&) = 0
          Next U&
          '
          NCMAX% = 0
          If OPT% = 1 Then
               Call ABRECABAN
               SQLX$ = "SELECT * FROM CAJABANC "
               SQLX$ = SQLX$ + "where FechEmisi >= " & DESDAT & " "
               SQLX$ = SQLX$ + "AND FechEmisi <= " & HASDAT & " "
               Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
               '
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
                    NC% = !NuClien
                    If VENDSEL% = 0 Or VENDSEL% = VENDCLI%(NC%) Then
                       If NC% < 0 Then NC% = 0
                       If NC% > NCMAX% Then NCMAX% = NC%
                       ITOTA# = !IHABECOMP - !IDEBECOMP
                       If TIMO$ = "VENTA" Then
                            TOTVENTA#(NC%) = TOTVENTA#(NC%) + ITOTA#
                          ElseIf TIMO$ = "IVA" Then
                            TOTIVA#(NC%) = TOTIVA#(NC%) + ITOTA#
                       End If
                    End If
19                .MoveNext
                  Loop
               End With
               '
            ElseIf OPT% = 2 Then
               Call ABREPEDCLI
               SQLX$ = "SELECT * FROM PEDDETA "
               SQLX$ = SQLX$ + "WHERE INT(CDBL(FechEmis)) >= " & DESDAT & " "
               SQLX$ = SQLX$ + "AND INT(CDBL(FechEmis)) <= " & HASDAT & " "
               SQLX$ = SQLX$ + "AND Status<8"
               '
               On Error GoTo 0
               Set PEDDETTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
               '
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
                     NC% = !NroClie
                     If VENDSEL% = 0 Or VENDSEL% = VENDCLI%(NC%) Then
                       If NC% < 0 Then NC% = 0
                       If NC% > NCMAX% Then NCMAX% = NC%
                       NETO# = !CANPED * !PREUNID
                       If !MoneEmis > 1 Then
                          NETO# = NETO# * TICAMONE(!MoneEmis)
                       End If
                       NETO# = ROUND(NETO#, 2)
                       TOTVENTA#(NC%) = TOTVENTA#(NC%) + NETO#
                       ' AQUI VER COMO CALCULAR EL iva
                       NIVA# = (NETO# * PORCEIVA(NC%)) / 100
                       TOTIVA#(NC%) = TOTIVA#(NC%) + NIVA#
                     End If
                  .MoveNext
                  Loop
               End With
          End If
          '
          For U& = 1 To NCMAX%
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
             If OPT% = 1 Then
                  Call FINAL("*STATVENT" + FORMAEM$)
                ElseIf OPT% = 2 Then
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
   Call MENSERR(24, "No Compatible con la Presente Versión")
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   If DERACCE%("RECUPEDI", "Recuperacion de Precios de Pedidos") = 2 Then
      Call APERBASDAT("PEDCLI")
      SQLX$ = "SELECT * FROM PEDDETA "
      SQLX$ = SQLX$ + "WHERE NroPed >= 170 "
      SQLX$ = SQLX$ + "AND (PreList > 0 and PreUnid <0.005) "
      SQLX$ = SQLX$ + "ORDER BY NroPed ASC"
      '
      Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenDynaset)
      '
      With PEDTEMP
        On Error Resume Next
        .MoveLast
        If Err > 1 Then
          Call COMUNI("No Existen Pedidos en el Período Seleccionado")
          GoTo 40
        End If
        FIN& = .RecordCount
        On Error GoTo 0
        '
        .MoveFirst
        KU& = 0
        Do While Not .EOF
           KU& = KU& + 1
           Call TRACE(20, KU&, FIN&)
           .Edit
           PRELI = !PreList
           CANTU = !CANPED
           ITOTA = PRELI * CANTU
           !PREUNID = PRELI
           !ImpoTot = ITOTA
           .Update
        .MoveNext
        Loop
      End With
      PEDTEMP.Close
      '
40    FLIMI = CDbl(CVDAT(FECHANUM("30/08/07")))
      Call APERBASDAT("STOCKS")
      SQLX$ = "SELECT * FROM MOVISTO "
      SQLX$ = SQLX$ + "WHERE FECHMOV >= " & FLIMI & " "
      SQLX$ = SQLX$ + "and (CODIMOVI = 'RT' OR CODIMOVI = 'CI')"
      SQLX$ = SQLX$ + "AND NuPedCli > 0 "
      SQLX$ = SQLX$ + "AND ValoUnit < 0.005 "
      '
      Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
      '
      With MoviTemp
        On Error Resume Next
        .MoveLast
        If Err < 1 Then
           FIN& = .RecordCount
           On Error GoTo 0
           '
           .MoveFirst
           KU& = 0
           Do While Not .EOF
              KU& = KU& + 1
              Call TRACE(20, KU&, FIN&)
              NPED& = !NuPedCli
              CIXI% = !Cod_Inte
              ncli% = !Nume_Clie
              PREPE# = PRECIOPED#(NPED&, CIXI%, MONEPE%)
              .Edit
              !ValoUnit = PREPE#
              If MONEPE% > 0 Then
                 !MoneVal = MONEPE%
              End If
              .Update
           .MoveNext
           Loop
        End If
      End With
      MoviTemp.Close
      '
      Call COMUNI("Fin Recuperación")
   End If
99 Screen.MousePointer = 1
   Command14.Enabled = True
End Sub


Private Sub Command15_Click()

End Sub

Private Sub Command2_Click()
'Sub COPEFAC()
   '
   RECU% = 0: If HOY(HOS$) <= FECHANUM("24/10/01") Then RECU% = 1
   HOII% = HOY(HOS$)
   JJ& = 0
   FIN& = ULTREG("PEDDETA")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("PEDDETA", U&)
     NC% = CVI(Mid$(XX$, 7, 2))
     If NC% > 0 Then
       If Asc(Mid$(XX$, 27, 1)) <> 9 Then ' no esta anulado
         CI% = CVI(Mid$(XX$, 33, 2))
         If CI% > 0 And CI% <= NUMAS% Then
           FF% = CVI(Mid$(XX$, 5, 2))
           NF& = CVS(Left$(XX$, 4))
           DOCU$ = "PC-" + TRIM$(Str$(NF&))
           CAPED = CVS(Mid$(XX$, 79, 4))
           SALPE = CAPED - CVS(Mid$(XX$, 95, 4))
           '
           If RECU% = 1 Then
             If NF& = 77 Then
               If CI% = CODINT%("065008T-0000") Then
                 TOENTREGA = 1
                 SALPENDIEN = 0
                 REMI = 1681
                 Call REPLA(XX$, MKS$(TOENTREGA), 95, 4)
                 Call REPLA(XX$, MKS$(SALPENDIEN), 99, 4)
                 Call REPLA(XX$, MKI$(HOII%), 103, 2)
                 Call REPLA(XX$, MKS$(REMI), 105, 4)
                 Call GRAREG("PEDDETA", XX$, U&)
               End If
             End If
             '
             If NF& = 85 Then
               If CI% = CODINT%("042261T-PHH0") Then
                 TOENTREGA = 1
                 SALPENDIEN = 0
                 REMI = 1685
                 Call REPLA(XX$, MKS$(TOENTREGA), 95, 4)
                 Call REPLA(XX$, MKS$(SALPENDIEN), 99, 4)
                 Call REPLA(XX$, MKI$(HOII%), 103, 2)
                 Call REPLA(XX$, MKS$(REMI), 105, 4)
                 Call GRAREG("PEDDETA", XX$, U&)
               End If
             End If
             '
             If NF& = 21 Then
               If CI% = CODINT%("44132/2K") Then
                 TOENTREGA = 5
                 SALPENDIEN = 0
                 REMI = 1628
                 Call REPLA(XX$, MKS$(TOENTREGA), 95, 4)
                 Call REPLA(XX$, MKS$(SALPENDIEN), 99, 4)
                 Call REPLA(XX$, MKI$(HOII%), 103, 2)
                 Call REPLA(XX$, MKS$(REMI), 105, 4)
                 Call GRAREG("PEDDETA", XX$, U&)
               End If
             End If
             '
             If NF& = 81 Then
               If CI% = CODINT%("R410104-PVH0") Then
                 TOENTREGA = 1
                 SALPENDIEN = 0
                 REMI = 1652
                 Call REPLA(XX$, MKS$(TOENTREGA), 95, 4)
                 Call REPLA(XX$, MKS$(SALPENDIEN), 99, 4)
                 Call REPLA(XX$, MKI$(HOII%), 103, 2)
                 Call REPLA(XX$, MKS$(REMI), 105, 4)
                 Call GRAREG("PEDDETA", XX$, U&)
               End If
             End If
             '
           End If
           '
           YY$ = REGBLAN$("COPEFAC")
           Call REPLA(YY$, MKI$(NC%), 1, 2)
           Call REPLA(YY$, MKI$(CI%), 3, 2)
           Call REPLA(YY$, MKI$(FF%), 5, 2)
           Call REPLA(YY$, DOCU$, 7, 12)
           Call REPLA(YY$, MKS$(SALPE), 19, 4)
           Call REPLA(YY$, MKS$(CAPED), 27, 4)
           '
           JJ& = JJ& + 1
           Call GRAREG("COPEFAC", YY$, JJ&)
           '
         End If
       End If
     End If
   Next U&
   '
   FIN& = ULTREG("MOVISTO")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("MOVISTO", U&)
      DOPRI$ = Mid$(XX$, 23, 10): DOSEC$ = Mid$(XX$, 33, 10)
      CMX1$ = Left$(DOPRI$, 2): CMX2$ = Left$(DOSEC$, 2)
      If CMX1$ = "FC" Or CMX2$ = "FC" Or CMX1$ = "RT" Or CMX2$ = "RT" Then
         NC% = CVI(Mid$(XX$, 84, 2))
         If NC% > 0 Then
           CI% = CVI(Mid$(XX$, 3, 2))
           FF% = CVI(Mid$(XX$, 1, 2))
           DOCU$ = DOPRI$
           CAENT = CVD(Mid$(XX$, 96, 8))
           '
           YY$ = REGBLAN$("COPEFAC")
           Call REPLA(YY$, MKI$(NC%), 1, 2)
           Call REPLA(YY$, MKI$(CI%), 3, 2)
           Call REPLA(YY$, MKI$(FF%), 5, 2)
           Call REPLA(YY$, DOCU$, 7, 12)
           Call REPLA(YY$, MKS$(CAENT), 31, 4)
           '
           JJ& = JJ& + 1
           Call GRAREG("COPEFAC", YY$, JJ&)
           '
         End If
      End If
   Next U&
   '
   If EXISTE%(LUDAT$ + "MOVIREPA.RFS") < 1 Then
     Call COPYSTRU("MOVISTO", "MOVIREPA")
   End If
   '
   FIN& = ULTREG("MOVIREPA")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("MOVIREPA", U&)
      DOPRI$ = Mid$(XX$, 23, 10): DOSEC$ = Mid$(XX$, 33, 10)
      CMX1$ = Left$(DOPRI$, 2): CMX2$ = Left$(DOSEC$, 2)
      If CMX1$ = "FC" Or CMX2$ = "FC" Or CMX1$ = "RT" Or CMX2$ = "RT" Then
         NC% = CVI(Mid$(XX$, 84, 2))
         If NC% > 0 Then
           CI% = CVI(Mid$(XX$, 3, 2))
           FF% = CVI(Mid$(XX$, 1, 2))
           DOCU$ = DOPRI$
           CAENT = CVD(Mid$(XX$, 96, 8))
           '
           YY$ = REGBLAN$("COPEFAC")
           Call REPLA(YY$, MKI$(NC%), 1, 2)
           Call REPLA(YY$, MKI$(CI%), 3, 2)
           Call REPLA(YY$, MKI$(FF%), 5, 2)
           Call REPLA(YY$, DOCU$, 7, 12)
           Call REPLA(YY$, MKS$(CAENT), 31, 4)
           '
           JJ& = JJ& + 1
           Call GRAREG("COPEFAC", YY$, JJ&)
           '
         End If
      End If
   Next U&
   '
   Call SETULTREG("COPEFAC", JJ&)
   Call CIERRARCH("COPEFAC")
   Call CIERRARCH("*.*")
   Call FINAL("*SOPEFAC")
   '
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
10 OPX% = ALTERNA%("Imputacion de Anticipos\Etiqueta de Despacho\Otros Accesos Directos")
    If OPX% = 1 Then
       Call IMPUTAPEDI
       GoTo 10
     ElseIf OPX% = 2 Then
       ETIDESPA.Show 1
       GoTo 10
     ElseIf OPX% = 3 Then
       Call ACCDIR("PEDIDO02")
   End If
99 Command3.Enabled = True
End Sub

 Sub Command4_Click()
    Command4.Enabled = False
    If FORCAPED% < 1 Or FORCAPED% > 2 Then
       FORCAPED% = 1
       If CONTROL$("ADMIPED", "FORMACAR") = "ITEMIZADO" Then FORCAPED% = 2
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
       If CONTROL$("ADMIPED", "FORMACAR") = "ITEMIZADO" Then FORCAPED% = 2
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
       If CONTROL$("ADMIPED", "FORMACAR") = "ITEMIZADO" Then FORCAPED% = 2
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
       If CONTROL$("ADMIPED", "FORMACAR") = "ITEMIZADO" Then FORCAPED% = 2
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

Sub CAMBIAPED()
     '
2100 If YAINDI% = 0 Then
       '
       Screen.MousePointer = 11
       
       Call ABREPEDCLI
       SQLX$ = " SELECT DISTINCT PEDENCA.NroPed, PEDENCA.NroClie, PEDENCA.FechEmis, PEDENCA.NroOcom "
       SQLX$ = SQLX$ + " FROM PEDENCA "
       SQLX$ = SQLX$ + " INNER JOIN  PEDDETA "
       SQLX$ = SQLX$ + " ON (PEDENCA.NroPed = PEDDETA.NroPed) "
       SQLX$ = SQLX$ + " WHERE PEDDETA.Status < 8 "   ' para que no este anulada ni cumplida
       SQLX$ = SQLX$ + " AND PEDDETA.CantEnt < (PEDDETA.CanPed - 0.005) "
       SQLX$ = SQLX$ + " AND PEDENCA.Status < 9 "
       SQLX$ = SQLX$ + " ORDER BY PEDENCA.NroPed ASC "
       Set PediModifTemp = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
       '
       JJ& = 0
       With PediModifTemp
         On Error Resume Next
         .MoveFirst
         If Err < 1 Then
           Do While Not .EOF
               NRODES& = !NroPed
             TTXX$ = MKS$(NRODES&)
               NOCO$ = !NroOcom
             If NOCO$ <> "" Then
               TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
             End If
               FEX = CVINT(!FECHEMIS)
             TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
               NCLIE% = !NroClie
             TTXX$ = TTXX$ + TRIM$(RASOCLI$(NCLIE%)) + Space$(52)
             JJ& = JJ& + 1
             Call GRAREG("INDIPEP", Left$(TTXX$, 52) + MKS$(REGENCA&), JJ&)
           .MoveNext
           Loop
         End If
       End With
       '
       Call SETULTREG("INDIPEP", JJ&)
       Call CIERRARCH("INDIPEP")
       Call FRESHECHO("INDIPEP")
       Call CIERRARCH("*.*")
       Call BAJALDISCO
       '
       Screen.MousePointer = 1
       YAINDI% = 1
     End If
     '
2112 If ULTREG&("INDIPEP") < 1 Then
       Call COMUNI("No hay Pedidos Abiertos\que Puedan Modificarse")
       GoTo 9999
     End If
     '
2110 X$ = REGSEL$("INDIPEP/Pedidos Pendientes de Entrega")
     If Len(X$) <= 4 Then
        GoTo 9999
     End If
     '
     Screen.MousePointer = 11
     NUPEDIDO& = CVS(Left$(X$, 4))
     SQLX$ = "SELECT * FROM PEDENCA "
     SQLX$ = SQLX$ + " WHERE NroPed = " & NUPEDIDO&
     Set PEDENCATEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     On Error Resume Next
     With PEDENCATEMP
       .MoveFirst
       If Err > 0 Then
         Call MENSERR(24, "Imposible Modificar Este Pedido")
         GoTo 9999
       End If
       '
       '23/03/07 (OT-07-0121)'NO PERMITE MODIFICAR UN PEDIDO A UN CLIENTE CON LA
       NC% = !NroClie        'CUENTA SUSPENDIDA
       LMCR# = LIMICRE(NC%)
       If LMCR# < 0 Then
          Call COMUNI("Imposible Modificar Pedido.\ \La Cuenta esta Suspendida")
          GoTo 9999
       End If
       '23/03/07 (FIN)
       '
         FEX = CVINT(!FECHEMIS)
       MODPED04.Text9.TEXT = FECHATEX$(FEX)
       MODPED04.Text11.TEXT = TRIM$(Str$(NUPEDIDO&))
         NCLIE% = !NroClie
       MODPED04.LTEXT1 = TRIM$(Str$(NCLIE%))
       MODPED04.Text1.TEXT = TRIM$(Str$(NCLIE%))
       '
         NC% = NCLIE%
       MODPED04.Text2.TEXT = RASOCLI$(NC%)
       MODPED04.Text3.TEXT = DOMICLI$(NC%)
       MODPED04.Text4.TEXT = CPOSCLI$(NC%)
       MODPED04.Text5.TEXT = LOCACLI$(NC%)
       MODPED04.Text7.TEXT = CSTRING$(MKS$(LIMICRE(NC%)), 3, 13, 2, 2)
       MODPED04.Text12.TEXT = TRIM$(POSIVAC$(NC%)) + " - CUIT " + CUITCLI$(NC%)
         CREDISPO# = LIMICRE(NC%) - SALDCLI#(NC%) - CHEPEACRE#(NC%)
       MODPED04.Text8.ForeColor = RGB(0, 0, 0)
         If CREDISPO# < 0 Then MODPED04.Text8.ForeColor = RGB(255, 0, 0)
       MODPED04.Text8.TEXT = CSTRING$(MKD$(Abs(CREDISPO#)), 3, 13, 2, 2)
       '
       MODPED04.Text10.TEXT = !NroOcom
          LPREX% = !ListPre
       MODPED04.TEXT21(0).TEXT = TRIM$(Str$(LPREX%))
       MODPED04.TEXT21(1).TEXT = TRIM$(Str$(!CondPag))
         If XVALO(MODPED04.TEXT21(1).TEXT) < 1 Then MODPED04.TEXT21(1).TEXT = Str$(CPAGCLI%(NC%))
       MODPED04.TEXT21(2).TEXT = TRIM$(Str$(!VENDED))
         If XVALO(MODPED04.TEXT21(2).TEXT) < 1 Then MODPED04.TEXT21(2).TEXT = Str$(VENDCLI%(NC%))
         If XVALO(MODPED04.TEXT21(2).TEXT) < 1 Then MODPED04.Eco21(2).TEXT = "GERENCIA"
         '
       MODPED04.PORDESCU.TEXT = !TDescApli
       MODPED04.LENTREGA(0) = !DOMIENT
       MODPED04.LENTREGA(1) = !LOCAENT
       MODPED04.LENTREGA(2) = !TranspEnt
       '
       MODPED04.Text19.TEXT = !NTranspo
       MODPED04.Text20.TEXT = !NSucent
       '
       MODPED04.Text13.TEXT = !Observa
       MODPED04.Label11.ForeColor = RGB(0, 0, 255)
       MODPED04.Label11 = "Venta"
       '
     End With
     '
     Call SETULTREG("!CARDETPE", 0)
     '
     ' CONTROL DE REPETICIONES DE NUMERO DE ORDEN EN PEDIDO O NUMERO DE ORDEN EN 0
     '
     Call REVISANUMERO(NUPEDIDO&)
     '
     JJ& = 0
     REBLA$ = REGBLAN$("CARDETPE")
     SQLX$ = "SELECT * FROM PEDDETA "
     SQLX$ = SQLX$ + "WHERE NroPed = " & NUPEDIDO& & " "
     '\\\OT-06-0520-RG-01/12/06///
     SQLX$ = SQLX$ + "AND Status < 8 "
     '\\\OT-06-0520-RG-FIN///
     SQLX$ = SQLX$ + "ORDER BY NuorItem ASC "
     Set PEDDETTEMP = Ventas.OpenRecordset(SQLX$, dbOpenDynaset)
     '
     MODODAKOT% = 0
     If InStr(UCase$(EMPREAC$), "DAKOT") < 0 Then MODODAKOT% = 1
     '
     VALINORI$ = String$(1024, "0")
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
               MODPED04.MONEMIS = DEMONECO$(MONEMI%)
               MODPED04.TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
         End If
         '
         X$ = REBLA$
         DLISTA = 0: DGRUPO = 0: DTOTA = 0
         CIXI% = PEDDETTEMP!CodiInt
         Call REPLA(X$, MKI$(CIXI%), 1, 2)
         Call REPLA(X$, PEDDETTEMP!TColor, 3, 16)
         Call REPLA(X$, PEDDETTEMP!TTalle, 19, 8)
           TLAR = PEDDETTEMP!TLargo
         Call REPLA(X$, MKS$(TLAR), 27, 4)
           TCHO = PEDDETTEMP!TAncho
         Call REPLA(X$, MKS$(TCHO), 31, 4)
           TESP = PEDDETTEMP!TEspesor
         Call REPLA(X$, MKS$(TESP), 35, 4)
         If CONTROL$("PEDIDO", "VARCOLOR") = "SI" Then  ' si la opcion de pedidos por variante esta en true
             Call REPLA(X$, SAFETEXT$(PEDDETTEMP!CBase), 3, 12)
             Call REPLA(X$, SAFETEXT$(PEDDETTEMP!TTalle), 15, 12)
             Call REPLA(X$, SAFETEXT$(PEDDETTEMP!TColor), 27, 12)
         End If
           PREL = PEDDETTEMP!PreList * TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
         Call REPLA(X$, MKS$(PREL), 39, 4)
           PREU = PEDDETTEMP!PREUNID * TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
         Call REPLA(X$, MKS$(PREU), 51, 4)
         If MODODAKOT% < 1 Then
           If PREU > 0.005 Then
             Call REPLA(X$, MKS$(PREU), 39, 4)
           End If
         End If
           DLISTA = XVALO(PEDDETTEMP!DescLista)
           If DLISTA = 0 Then DLISTA = DESCULISPRE(LPREX%)
         Call REPLA(X$, MKS$(DLISTA), 109, 4)
           DTOTA = XVALO(PEDDETTEMP!DescTotal)
           If MODODAKOT% = 1 Then
             If DTOTA = 0 Then
               If PREL > 0.005 Then DTOTA = 100 * (1 - PREU / PREL)
             End If
           End If
         Call REPLA(X$, MKS$(DTOTA), 69, 4)
           DGRUPO = XVALO(PEDDETTEMP!DescGrupo)
           If DGRUPO = 0 Then
              COEFITO# = (100 - DTOTA) / 100
              COEFILI# = (100 - DLISTA) / 100
              COEFIGR# = 0: If COEFILI# > 0 Then COEFIGR# = COEFITO# / COEFILI#
              DGRUPO = 100 * (1 - COEFIGR#)
           End If
         Call REPLA(X$, MKS$(DGRUPO), 113, 4)
           CAN = PEDDETTEMP!CANPED
         Call REPLA(X$, MKS$(CAN), 47, 4)
           FSE% = CVINT(PEDDETTEMP!FeSolEnt)
         Call REPLA(X$, MKI$(FSE%), 63, 2) ' FECHA SOLICITADA DE ENTREGA
           NORIT% = XVALO(PEDDETTEMP!NuorItem)
           ' AQUÍ VALIDA Y ACTUALIZA EL NUMERO DE ORDEN DEL ITEM
           While Len(VALINORI$) < NORIT%: VALINORI$ = VALINORI$ + "0": Wend
           If Mid$(VALINORI$, NORIT%, 1) <> "0" Then
              NORIT% = InStr(VALINORI$, "0")
              PEDDETTEMP.Edit
              PEDDETTEMP!NuorItem = NORIT%
              PEDDETTEMP.Update
           End If
         Call REPLA(X$, MKI$(NORIT%), 123, 2)
           DCOMBI$ = TRIM$(PEDDETTEMP!DescCombi)
         Call REPLA(X$, DCOMBI$, 93, 16)
         Call REPLA(X$, MKS$(PORCEDESCU(DECOMBI$)), 113, 4)
         Call REPLA(X$, MKS$(0), 117, 4)   ' ESTO SE ACTUALIZA MAS ABAJO
         Call REPLA(X$, MKI$(0), 121, 2)
           SUCOD$ = ""
           If IsNull(PEDDETTEMP!SubCod) = False Then SUCOD$ = PEDDETTEMP!SubCod
         If SUCOD$ <> "" Then Call REPLA(X$, SUCOD$, 3, 8)
         JJ& = JJ& + 1
         Call GRAREG("!CARDETPE", X$, JJ&)
         '
       PEDDETTEMP.MoveNext
       Loop
     End If
     PEDDETTEMP.Close
     Call SETULTREG("!CARDETPE", JJ&)
     Call CIERRARCH("CARDETPE")
     Call SETULTREG("!CARDETP1", JJ&)
     Call CIERRARCH("CARDETP1")
     On Error GoTo 0
     '
     ' REVISAR Y ACTUALIZAR CAENTAN
     Dim CANENTAN(32767), FULENT%(32767)
     CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
     If CMO$ = "" Then CMO$ = "RT"
     NMS% = NUMAS%
     '
     Call ABRESTOCKS
     SQLX$ = "SELECT * FROM MOVISTO "
     SQLX$ = SQLX$ + " WHERE CodiMovi = '" & CMO$ & "' "
     SQLX$ = SQLX$ + " AND NuPedCli = " & NUPEDIDO& & " "
     SQLX$ = SQLX$ + " AND Nume_Clie = " & NCLIE% & " " ' SE POR SILCAR PARA ADEMAS COINCIDA EL CLIENTE
     Set MOVITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     On Error Resume Next
     With MOVITMP
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           CIXXI% = !Cod_Inte
           If CIXXI% > 0 Then
             If CIXXI% <= NMS% Then
               If DESCOMPO%(CIXXI%) > 0 Then ' PARA EL CASO DE LOS ITEMS CON DESCOMP.AUTOMATICA
                 CANENTAN(CIXXI%) = CANENTAN(CIXXI%) + !Cantsalid
               Else
                 CANENTAN(CIXXI%) = CANENTAN(CIXXI%) + !Cantsalid - !CantIngre
               End If
               If !Cantsalid >= 0.0005 Then
                 If CVINT(!fechMov) > FULENT%(CIXXI%) Then
                   FULENT%(CIXXI%) = CVINT(!fechMov)
                 End If
               End If
             End If
           End If
         .MoveNext
         Loop
       End If
     End With
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
     MODPED04.Text6.TEXT = ""
     '
     For U& = 1 To Len(RNC$) Step 4
       RENOTA& = CVS(Mid$(RNC$, U&, 4))
       If RENOTA& > 0 Then
         If RENOTA& <= URN& Then
           TTYY$ = RTrim$(Mid$(REGLEIDO$("NOTACLI", RENOTA&), 3))
           If MODPED04.Text6.TEXT <> "" Then
             MODPED04.Text6.TEXT = MODPED04.Text6.TEXT + Chr$(13) + Chr$(10)
           End If
           MODPED04.Text6.TEXT = MODPED04.Text6.TEXT + RTrim$(TTYY$)
         End If
       End If
     Next U&
     '
8000 MODPED04.MARCONOTA.Visible = True
     On Error Resume Next
     MODPED04.Command6.SetFocus
     On Error GoTo 0
     Screen.MousePointer = 1
     '
     NOFRESHPRE% = 1
     MODPED04.Show 1
     NOFRESHPRE% = 0
     '
     GoTo 2110
     '
9999 Call CIERRARCH("*.*")
     Screen.MousePointer = 1
     '
End Sub

'
Sub ANULAPED()
     '
2000 YAINDI% = 0
     '
2100 If YAINDI% = 0 Then
        'NUEVO
       Screen.MousePointer = 11
       Call ABREPEDCLI
       SQLX$ = " SELECT DISTINCT PEDENCA.NroPed, PEDENCA.NroClie, PEDENCA.FechEmis, PEDENCA.NroOcom "
       SQLX$ = SQLX$ + " FROM PEDENCA "
       SQLX$ = SQLX$ + " INNER JOIN  PEDDETA "
       SQLX$ = SQLX$ + " ON (PEDENCA.NroPed = PEDDETA.NroPed) "
       SQLX$ = SQLX$ + " WHERE (PEDDETA.Status < 8 "   ' para que no este anulada ni cumplida
       SQLX$ = SQLX$ + " AND (PEDDETA.CantEnt < (PEDDETA.CanPed - 0.005) OR PEDDETA.CanPed < 0.05) "   ' para incluir pedidos abiertos
       SQLX$ = SQLX$ + " AND PEDENCA.Status < 9) "
       SQLX$ = SQLX$ + " OR PEDENCA.STATUS < 0 "        ' PENDIENTE DE APROBACION
       SQLX$ = SQLX$ + " ORDER BY PEDENCA.NroPed ASC "
       Set PediAnulTemp = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
       '
       JJ& = 0
       With PediAnulTemp
         On Error Resume Next
         .MoveFirst
         If Err < 1 Then
           Do While Not .EOF
               NRODES& = !NroPed
             TTXX$ = MKS$(NRODES&)
               NOCO$ = !NroOcom
             If NOCO$ <> "" Then
               TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
             End If
               FEX = CVINT(!FECHEMIS)
             TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
               NCLIE% = !NroClie
               REGENCA& = NCLIE%
             TTXX$ = TTXX$ + TRIM$(RASOCLI$(NCLIE%)) + Space$(52)
             JJ& = JJ& + 1
             Call GRAREG("INDIPEP", Left$(TTXX$, 52) + MKS$(REGENCA&), JJ&)
           .MoveNext
           Loop
         End If
       End With
       Call SETULTREG("INDIPEP", JJ&)
       Call CIERRARCH("INDIPEP")
       Call FRESHECHO("INDIPEP")
       '
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
2210 X$ = REGSEL$("INDIPEP/Pedidos Pendientes para Anulación")
     If Len(X$) <= 4 Then GoTo 9999
     '
     NUPEDIDO& = CVS(Left$(X$, 4))
     NCLIEN& = CVS(Mid$(X$, 53, 4))
     '
     If COMALTER%("Confirme la Anulación del Pedido Número " + TRIM$(Str$(NUPEDIDO&)) + "\Correspondiente a: " + TRIM$(Mid$(X$, 5, 48)) + "\\Cancelar\Confirmar Anulación") <> 2 Then
        GoTo 2210
     End If
     
     'ACA ANULO el encabezado
     With PedEnca
       ABUS$ = "NroPed = " & NUPEDIDO& & "AND Status < 9 AND NroClie = " & NCLIEN&
       .FindFirst ABUS$
       If .NoMatch = False Then
         .Edit
         !Status = 9                                 'cambio status
         !Referencia = "(anul) " + !Referencia       'cambio referen
         .Update
       End If
     End With
     '
     ' ACA MARCO LOS ITEMS DEL PEDIDO
     ITEMREP% = 0
     With PedDeta
       ABUS$ = "NroPed= " & NUPEDIDO & " AND NroClie = " & NCLIEN&
20     If ITEMREP% = 0 Then
         .FindFirst (ABUS$)
       Else
         .FindNext (ABUS$)
       End If
       If .NoMatch = False Then
         .Edit
         If !Status <> 3 Then
           !Status = 9
           If CantEnt > 0 Then !Status = 8
         End If
         .Update
         ITEMREP% = 1
         GoTo 20
       End If
     End With
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
     'NUEVO
     Screen.MousePointer = 11
    Call ABREPEDCLI
       SQLX$ = " SELECT DISTINCT PEDENCA.NroPed, PEDENCA.NroClie, PEDENCA.FechEmis, PEDENCA.NroOcom,PEDENCA.Referencia "
       SQLX$ = SQLX$ + " FROM PEDENCA "
       SQLX$ = SQLX$ + " INNER JOIN  PEDDETA "
       SQLX$ = SQLX$ + " ON (PEDENCA.NroPed = PEDDETA.NroPed) "
       SQLX$ = SQLX$ + " WHERE PEDDETA.Status = 9  "   ' Selecciona las Anuladas
       SQLX$ = SQLX$ + " AND PEDDETA.CantEnt < (PEDDETA.CanPed - 0.005) "
       SQLX$ = SQLX$ + " AND PEDENCA.Status = 9 "      ' Selecciona los Anulados
       SQLX$ = SQLX$ + " ORDER BY PEDENCA.NroPed ASC "
       Set PeDesAnTemp = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
       '
       JJ& = 0
       Y$ = REGBLAN("INDIPEQ")
       With PeDesAnTemp
         On Error Resume Next
         .MoveFirst
         If Err < 1 Then
           Do While Not .EOF
               Call REPLA(Y$, MKS$(!NroPed), 1, 4)
               Call REPLA(Y$, !Referencia, 5, 48)
               JJ& = JJ& + 1
               Call GRAREG("INDIPEQ", Y$, JJ&)
           .MoveNext
           Loop
         End If
       End With
       Call SETULTREG("INDIPEQ", JJ&)
       Call CIERRARCH("INDIPEQ")
       Call FRESHECHO("INDIPEQ")
    '
2300 Screen.MousePointer = 11
'     JJ& = 0
'     Call FRESHECHO("INDIPEQ")
'     RFF$ = RECFILT$("PEDENCA", 14, Chr$(9))
'     For UU& = 1 To Len(RFF$) Step 4
'       I& = CVS(Mid$(RFF$, UU&, 4))
'       X$ = REGLEIDO$("PEDENCA", I&)
'       If Asc(Mid$(X$, 267, 1)) = 9 Then
'          Call VERIREFPED(X$, I&)
'          JJ& = JJ& + 1
'          Call GRAREG("INDIPEQ", Left$(X$, 52) + MKS$(I&), JJ&)
'       End If
'     'Next I&
'     Next UU&
'     Call SETULTREG("INDIPEQ", JJ&)
'     Call CIERRARCH("INDIPEQ")
'     Screen.MousePointer = 1
'     '
'     If ULTREG&("INDIPEQ") < 1 Then
'        Call COMUNI("No hay Pedidos Anulados\que Puedan Re-Abrirse")
'        Call CIERRARCH("*.*")
'        Exit Sub
'     End If
'     '
2310 X$ = REGSEL$("INDIPEQ/Pedidos Anulados - Revertir Anulación")
     If Len(X$) <= 4 Then
        Call CIERRARCH("*.*")
        Exit Sub
     End If
     '
     NROPE& = CVS(Left$(X$, 4))
     REGENCA& = CVS(Mid$(X$, 53, 4))
'     If REGENCA& < 1 Or REGENCA& > ULTREG&("PEDENCA") Then
'        Call CIERRARCH("*.*")
'        Exit Sub
'     End If
     '
     RGL% = Int(REGENCA&)
     Call BLOQARCH("PEDDETP")
     '
     'ACA CAMBIO EL ENCABEZADO
     With PedEnca
       ABUS$ = " NroPed= " & NROPE& & " AND Status = 9"
       .FindFirst (ABUS$)
       If .NoMatch = False Then
         .Edit
         !Status = 1            'MARCO EL STAT
         FECHEMII$ = Format(!FECHEMIS, "dd/mm/yy")
         CLIRASO$ = !RasoClie
         NUMOCOMP$ = !NroOcom
         If NUMOCOMP$ = "" Then     'MARCO LA REFERENCIA sin O/c
           !Referencia = CLIRASO$ & "- " & FECHEMII$
         Else                       'MARCO LA REFERENCIA con O/c
         !Referencia = CLIRASO$ & "- " & NUMOCOMP$ & "- " & FECHEMII$
         End If
         .Update
       End If
     End With
     '
     'ACA CAMBIO LOS ITEMS
     '
     ITEMREP% = 0
     With PedDeta
       ABUS$ = "NroPed= " & NROPE& & " AND Status = 9 "
       ABUS$ = ABUS$ + " AND CanPed > CantEnt "
20     If ITEMREP% = 0 Then
         .FindFirst (ABUS$)
       Else
         .FindNext (ABUS$)
       End If
       If .NoMatch = False Then
         .Edit
         !Status = 1
         .Update
         ITEMREP% = 1
         GoTo 20
       End If
     End With
    '
    Call FRESHECHO("INDIPEQ")
    '
    
'     Screen.MousePointer = 11
'     X$ = REGLEIDO$("PEDENCA", REGENCA&)
'     Call REPLA(X$, Chr$(1), 267, 1)
'     Call GRAREG("PEDENCA", X$, REGENCA&)
'     Call CIERRARCH("PEDENCA")
'     '
'     J& = 0
'     FIN& = ULTREG&("PEDDETA")
'     For I& = 1 To FIN&
'        Call TRACE(20, I&, FIN&)
'        X$ = REGLEIDO$("PEDDETA", I&)
'        PENDE = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
'        If PENDE < 0 Then PENDE = 0
'        Call REPLA(X$, MKS$(PENDE), 99, 4)
'        If CVS(Left$(X$, 4)) = NROPe& Then
'        Call REPLA(X$, Chr$(1), 27, 1)
'        Call GRAREG("PEDDETA", X$, I&)
'        '
'        PDC$ = X$
'        CANOR = CVS(Mid$(X$, 79, 4))
'        Call REPLA(PDC$, Chr$(1), 27, 1)
'        Call REPLA(PDC$, String$(38, 0), 71, 38)
'        Call REPLA(PDC$, MKS$(CANOR), 75, 4)
'        Call REPLA(PDC$, MKS$(CANOR), 99, 4)
'        Call REGAPP("DESPAPED", PDC$)
'        '
'        Call REPLA(X$, MKS$(I&), 109, 4)       ' puntero a PEDDETA
'        Call REGAPP("PEDDETP", X$)
'        End If
'     Next I&
'     Call CIERRARCH("PEDDETP")
'     Call CIERRARCH("PEDDETA")
'     '
'     HOII% = HOY(HOS$)
'     TTXX$ = Space$(96)
'     Call REPLA(TTXX$, HOS$ + "  Reversión de Anulación Interactiva", 1, 56)
'     Call REPLA(TTXX$, USERTER$, 69, 24)
'     Call REGAPP("ANOTAPED", Left$(MKS$(NROPe&) + TTXX$, 96))
'     Call REGAPP("ANOTAPED", MKS$(NROPe&) + String$(88, "="))
'     Call CIERRARCH("ANOTAPED")
'     '
      Call CIERRARCH("*.*")
      Call BAJALDISCO
     '
      Screen.MousePointer = 1
     '
     ' GoTo 2300
     '
End Sub
'

Sub ANULITPED()
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
       SQLX$ = SQLX$ + " WHERE (PEDDETA.Status < 8 "   ' para que no este anulada ni cumplida
       SQLX$ = SQLX$ + " AND PEDDETA.CantEnt < (PEDDETA.CanPed - 0.005) "
       SQLX$ = SQLX$ + " AND PEDENCA.Status < 9) "
       SQLX$ = SQLX$ + " OR PEDENCA.STATUS < 0 "        ' PENDIENTE DE APROBACION
       SQLX$ = SQLX$ + " ORDER BY PEDENCA.NroPed ASC "
       Set pedianuittemp = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
       '
       JJ& = 0
       With pedianuittemp
         On Error Resume Next
         .MoveFirst
         If Err < 1 Then
           Do While Not .EOF
               NRODES& = !NroPed
             TTXX$ = MKS$(NRODES&)
               NOCO$ = !NroOcom
             If NOCO$ <> "" Then
               TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
             End If
               FEX = CVINT(!FECHEMIS)
             TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
               NCLIE% = !NroClie
               REGENCA& = NCLIE%
             TTXX$ = TTXX$ + TRIM$(RASOCLI$(NCLIE%)) + Space$(52)
             JJ& = JJ& + 1
             Call GRAREG("INDIPEP", Left$(TTXX$, 52) + MKS$(REGENCA&), JJ&)
           .MoveNext
           Loop
         End If
       End With
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
     NUPEDIDO& = CVS(Left$(X$, 4))
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
     SQLX$ = SQLX$ + " ORDER BY PEDENCA.NroPed ASC "
     Set pedianuittemp = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     With pedianuittemp
       .MoveFirst
       Do While (.EOF = False)
         CAN = !CANPED - !CantEnt
         If CAN >= 0.05 Then
           CI% = !CodiInt
           TTYY$ = Space$(128)
           Call REPLA(TTYY$, CODEXT$(CI%), 1, 16)
           Call REPLA(TTYY$, DESCRIT0$(CI%), 19, 38)
           '\\\OT-05-0254-WAR-04/05/05///
           Call REPLA(TTYY$, FORMATNUM$(!CANPED, "F10.1"), 57, 10)
           Call REPLA(TTYY$, FORMATNUM$(!CantEnt, "F10.1"), 67, 10)
           Call REPLA(TTYY$, FORMATNUM$(CAN, "F10.1"), 77, 10)
           '\\\OT-05-0254-WAR-FIN///
           Call REPLA(TTYY$, FORMATNUM$(!NuorItem, "F8.0"), 113, 8)
           SELITPED.List2.AddItem TTYY$
           .MoveNext
         End If
       Loop
     End With
     '
1209 SELITPED.Caption = "Anulacion por Items Pedido " + DPEDI$
     SELITPED.Show 1
     DoEvents
     '
     If SELITPED.List2.ListCount > 0 Then
       Screen.MousePointer = 11
       '
       For U& = 1 To SELITPED.List2.ListCount
         TTXX$ = SELITPED.List2.List(U& - 1)
         '\\\OT-05-0254-WAR-05/05/05///
         If SELITPED.List2.Selected(U& - 1) = True Then 'Mid$(TTXX$, 71, 5) = "anul." Then
         '\\\OT-05-0254-WAR-FIN///
           CoIn% = CODINT%(Left$(TTXX$, 16))
           NUOR& = XVALO(Mid$(TTXX$, 113, 8))
           'Busco el registro
           With PedDeta
             ABUSS$ = " CodiInt = " & CoIn% & " "
             ABUSS$ = ABUSS$ + " AND NroPed = " & NUPEDIDO& & " "
             ABUSS$ = ABUSS$ + " AND NuOrItem = " & NUOR& & " "
             ABUSS$ = ABUSS$ + " AND Status < 8 "
             ABUSS$ = ABUSS$ + " AND Status <> 3 "
             '
             .FindFirst (ABUSS$)
             If .NoMatch = False Then
               .Edit
               !Status = 8
               .Update
             End If
           End With
         End If
       Next U&
 
       Call BLOQARCH("PEDDETP")
       Call RECAPEDPE
       Call CIERRARCH("PEDDETP")
       Screen.MousePointer = 1
       '
       HOII% = HOY(HOS$)
       TTXX$ = Space$(96)
       Call REPLA(TTXX$, HOS$ + "  " + "Anulación Parcial de Items Pedidos", 1, 56)
       Call REPLA(TTXX$, USERTER$, 69, 24)
       Call REGAPP("ANOTAPED", Left$(MKS$(NROPE&) + TTXX$, 96))
       Call REGAPP("ANOTAPED", MKS$(NROPE&) + String$(88, "="))
       Call CIERRARCH("ANOTAPED")
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

Sub Command5_Click()
  Command5.Enabled = False
  APRONP04.Show 1
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
35  Call DESHASFECHA(DES%, HAS%, PERIO$)
    If TRIM$(PERIO$) = "" Then GoTo 99
    '
    FEDES = CDbl(CVDAT(DES%))
    FEHAS = CDbl(CVDAT(HAS%))
    '
    Call ABREPEDCLI
    Call ABRESTOCKS
    '
40  SQLX$ = " SELECT * "
    SQLX$ = SQLX$ + " FROM PEDDETA "
    SQLX$ = SQLX$ + " WHERE INT(CDBL(FechEmis)) >= " & FEDES & ""
    SQLX$ = SQLX$ + " AND INT(CDBL(FechEmis)) <= " & FEHAS & ""
    SQLX$ = SQLX$ + " AND STATUS >= 0 "
    '
    Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
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
           NUCLI7% = !NroClie
           NUPED7& = !NroPed
           CODIN7% = !CodiInt
           SQLY$ = " SELECT * "
           SQLY$ = SQLY$ + " FROM MOVISTO "
           SQLY$ = SQLY$ + " WHERE Nume_Clie = " & NUCLI7%
           SQLY$ = SQLY$ + " AND CDBL(NuPedCli) = " & NUPED7&
           SQLY$ = SQLY$ + " AND Cod_Inte = " & CODIN7%
           '
           Set STOBUS = Stocks.OpenRecordset(SQLY$, dbOpenDynaset)
           '
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
           If SCP_CLI(JJ1&) = !NroClie Then
             If SCP_PED(JJ1&) = !NroPed Then
               If SCP_COD(JJ1&) = !CodiInt Then
                 SCP_SAL(JJ1&) = SCP_SAL(JJ1&) + !CANPED
                 SCP_ENT(JJ1&) = SCP_ENT(JJ1&) + !CantEnt
                 SCP_PEX$(JJ&) = SCP_PEX$(JJ&) + Chr$(124) + !TPedItem
                 GoTo 49
               End If
             End If
           End If
         Next JJ1&
         '
         JJ& = JJ& + 1
         SCP_CLI(JJ&) = !NroClie
         SCP_PED(JJ&) = !NroPed
         SCP_COD(JJ&) = !CodiInt
         SCP_SAL(JJ&) = !CANPED
         SCP_ENT(JJ&) = !CantEnt
         SCP_PEX$(JJ&) = !TPedItem
         '
49    .MoveNext
      Loop
      '
    End With
    '
    SQLX$ = " SELECT * "
    SQLX$ = SQLX$ + " FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE Nume_Clie > 0 "
    SQLX$ = SQLX$ + " AND CDBL(NuPedCli) > 0 "
    SQLX$ = SQLX$ + " AND CodiMovi = 'RT' "
    '
    Set STOCKTEMP = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
    '
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
        NCL% = !Nume_Clie
        NPED& = !NuPedCli
        CIXI% = !Cod_Inte
        CASAL = !Cantsalid - !CantIngre
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
    '
80  FIN& = JJ&
    HAYRECU% = 0
    With PedDeta
      For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
         If Abs(SCP_ENT(U&)) > 0.01 Then
           TTXX$ = "Error de Saldo Pendiente\En Pedido " & SCP_PED(U&) & "\Codigo " + CODEXT$(SCP_COD(U&))
           OPX% = COMALTER%(TTXX$ + "\\Continuar\Recuperar")
           If OPX% = 2 Then
             HAYRECU% = HAYRECU% + 1
             SALENT = SCP_ACU(U&)
             SQLX$ = " NroClie = " & SCP_CLI(U&) & " "
             SQLX$ = SQLX$ & "AND NroPed = " & SCP_PED(U&) & " "
             SQLX$ = SQLX$ & "AND CodiInt = " & SCP_COD(U&) & " "
             '
             .FindFirst SQLX$
             If .NoMatch = False Then
               .Edit
               If SALENT <= !CANPED Then
                   !CantEnt = SALENT
                   !CanSaldo = !CANPED - !CantEnt
                   SALENT = 0
                 Else
                   SALENT = SALENT - !CANPED
                   !CantEnt = !CANPED
                   !CanSaldo = 0
               End If
               .Update
               .FindNext SQLX$
             End If
           End If
         End If
      Next U&
    End With
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
    If Option1 Then
      Call COMUNI("Listado Solo Disponible por Impresión.")
    End If
    Call LIENCPED
    Command61.Enabled = True
    '\\\OT-05-0250-WAR-FIN///
End Sub

'
 Sub Command62_Click()
    Command62.Enabled = False
    '
    modogramy% = 0
    If CONTROL$("ADMIPED", "FORMALIS") = "MONEMORI" Then
      If EXISTE%(LUDAT$ + "LIPEDID.RFS") > 0 Then
        modogramy% = 1
      End If
    End If
    '
    '\\\OT-05-0249-WAR-04/05/05///
    If modogramy% < 1 Then Call MENSPED
    '\\\OT-05-0249-WAR-FIN///
    '
5   OPX% = ALTERNA%("Detalle Pedidos Pendientes\Seguimiento Comercial\Cuadro de Proyección")
    '
    If OPX% = 3 Then
       Call CUAPROPED
       GoTo 99
    End If
    '
    If OPX% = 1 Then
        If Option1.Value = True Then
           'Consultas por Pantalla
           Call RECAPEDPE
           sali$ = "*COPEDCL"
           If FORMAPED% = 2 Then sali$ = "*COQEDCL"
           Call FINAL(sali$)
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
           Call SELRACLIEN(NCL1%, NCL2%, VENDSEL%)
           If NCL1% < 1 Then GoTo 99
           '
           FORMASA$ = ""
           If EXISTE%(LUDAT$ + "DETPEPEN.RFS") Then
             FORMASA$ = "ASEMBLI"
           End If
           '
           VECPEDI$ = ""
           'nuevo
           Call ABREPEDCLI
           SQLX$ = " SELECT * "
           SQLX$ = SQLX$ + " FROM PEDDETA "
           SQLX$ = SQLX$ + " WHERE NroClie >= " & NCL1% & " "
           SQLX$ = SQLX$ + " AND  NroClie <= " & NCL2% & " "
           SQLX$ = SQLX$ + " AND Status < 8 "
           If modogramy% < 1 Then SQLX$ = SQLX$ + " AND STATUS >= 0 "
           SQLX$ = SQLX$ + " AND (((CanPed - CantEnt) >= 0.05) OR (TipoPed = 'ABIERTO' AND CDBL(FinVigen) >= " & Int(CDbl(Now)) & ")) "
           '
           Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
           '
           With PEDTEMP
             On Error Resume Next
             .Move first
             '
             JJ& = 0
             Y$ = REGBLAN("LIPEDPE")
             '
             Do While (.EOF = False)
                CI% = !CodiInt
             If CI% > 0 And CI% <= NUMAS% Then
                CAPEN = !CANPED - !CantEnt
              '
              If CAPEN < 0 Then CAPEN = 0
                FENTRESOL% = CVINT(!FeSolEnt)
                '
                If FENTRESOL% >= FLI1% Then
                  If FENTRESOL% <= FLI2% Then
                    NC% = !NroClie
                    If VENDSEL% < 1 Or VENDCLI%(NC%) = VENDSEL% Then
                          Call REPLA(Y$, Mid$(X$, 7, 2), 1, 2)
                          Call REPLA(Y$, Mid$(X$, 7, 2), 45, 2)
                          Call REPLA(Y$, Mid$(X$, 33, 42), 3, 42)
                          CAPEN = !CANPED - !CantEnt
                        '
                        If CAPEN < 0 Then CAPEN = 0
                          CANOASI = CAPEN - !CanAsig
                        '
                        If CANOASI < 0 Then CANOASI = 0
                          MONEMISPED% = !MoneEmis
                        PREUNI = !PREUNID * (100 - PORDESCLI(NC%)) / 100 * TICAMONE(MONEMISPED%)
                        '
                        If OPX% <> 2 Then PREUNI = 0
                        ITOTA# = CDbl(CAPEN * PREUNI)
                        INOAS# = CDbl(CANOASI * PREUNI)
                        '
                        Call REPLA(Y$, MKI$(NC%), 1, 2)
                        Call REPLA(Y$, MKI$(!CodiInt), 3, 2)
                        Call REPLA(Y$, !TColor, 5, 16)
                        Call REPLA(Y$, !TTalle, 21, 8)
                        Call REPLA(Y$, MKS$(!TLargo), 29, 4)
                        Call REPLA(Y$, MKS$(!TAncho), 33, 4)
                        Call REPLA(Y$, MKS$(!TEspesor), 37, 4)
                        Call REPLA(Y$, MKS$(!NroPed), 47, 4)
                        Call REPLA(Y$, MKI$(FENTRESOL%), 69, 2)
                        Call REPLA(Y$, MKS$(!CANPED), 81, 4)
                        Call REPLA(Y$, MKS$(CAPEN), 101, 4)
                        Call REPLA(Y$, MKS$(PREUNI), 85, 4)
                        Call REPLA(Y$, MKD$(ITOTA#), 89, 8)
                        Call REPLA(Y$, MKS$(!CantEnt), 97, 4)
                        Call REPLA(Y$, MKS$(CANOASI), 111, 4)
                        Call REPLA(Y$, MKD$(INOAS#), 119, 8)
                        Call REPLA(Y$, !NroOcom, 127, 16)
                        Call REPLA(Y$, Space$(18), 143, 18)
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
           '
           If FORMASA$ = "ASEMBLI" Then
             Call FILESORT("LIPEDPE", "", 3, 3, "ASC") ' POR CODIGO
             Call FILESORT("LIPEDPE", "", 13, 13, "ASC") ' POR PEDIDO
             Call FILESORT("LIPEDPE", "", 1, 1, "ASC") ' POR CLIENTE
             '
             JJ& = 0
             For NCL% = NCL1% To NCL2%
               RFF$ = RECFILT$("LIPEDPE", 1, MKI$(NCL%))
               If Len(RFF$) > 0 Then
                 REBLA$ = REGBLAN$("DETPEPEN")
                 Call REPLA(REBLA$, MKI$(NCL%), 1, 2)
                    RASO1$ = TRIM$(Str$(NCL%)) + " - " + RASOCLI$(NCL%)
                 Call REPLA(REBLA$, RASO1$, 3, 48)
                 Call REPLA(REBLA$, DOMICLI$(NCL%), 51, 48)
                 Call REPLA(REBLA$, LOCACLI$(NCL%), 99, 48)
                 '
                 For KL& = 1 To Len(RFF$) Step 4
                   RELI& = CVS(Mid$(RFF$, KL&, 4))
                   LIPE$ = REGLEIDO$("LIPEDPE", RELI&)
                   ZZ$ = REBLA$
                     '
                     NPED& = CVS(Mid$(LIPE$, 47, 4))
                   Call REPLA(ZZ$, MKS$(NPED&), 147, 4)
                     STA$ = "  "
                     If Asc(Mid$(LIPE$, 159, 1)) = 255 Then STA$ = "P."
                     If Asc(Mid$(LIPE$, 159, 1)) = 8 Then STA$ = "AN"
                     If Asc(Mid$(LIPE$, 159, 1)) = 9 Then STA$ = "AN"
                   Call REPLA(ZZ$, STA$, 151, 2)
                     NOCO$ = Mid$(LIPE$, 127, 16)
                   Call REPLA(ZZ$, NOCO$, 155, 16)
                     CIXI% = CVI(Mid$(LIPE$, 3, 2))
                   Call REPLA(ZZ$, MKI$(CIXI%), 171, 2)
                     FECHENT% = CVI(Mid$(LIPE$, 69, 2))
                   Call REPLA(ZZ$, MKI$(FECHENT%), 173, 2)
                     CANORI = CVS(Mid$(LIPE$, 81, 4))
                   Call REPLA(ZZ$, MKS$(CANORI), 175, 4)
                     CANENT = CVS(Mid$(LIPE$, 97, 4))
                   Call REPLA(ZZ$, MKS$(CANENT), 179, 4)
                     CANSAL = CVS(Mid$(LIPE$, 101, 4))
                   Call REPLA(ZZ$, MKS$(CANSAL), 183, 4)
                     MONEM% = Asc(Mid$(LIPE$, 160, 1))
                   Call REPLA(ZZ$, SIMBOLPES$(MONEM%), 187, 4)
                     PREUN# = CVS(Mid$(LIPE$, 85, 4)) / TICAMONE(MONEM%)
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
        Dim SCP_CLI%(4096)        ' VECTOR CLIENTES
        Dim SCP_PED&(4096)        ' VECTOR PEDIDOS
        Dim SCP_COD%(4096)        ' VECTOR CODIGOS
        Dim SCP_REG&(4096)        ' VECTOR REGISTRO PEDIDO
        Dim SCP_STAT%(4096)       ' VECTOR STATUS PEDIDO
        Dim COEFPES#(4096)   ' VECTOR TOTAL PEDIDO
        Dim COEFDOL#(4096)   ' VECTOR ANTICIPO EN DOLARES
        '
35      Call DESHASFECHA(DES%, HAS%, PERIO$)
        If TRIM$(PERIO$) = "" Then GoTo 5
        '
        FEDES = CDbl(CVDAT(DES%))
        FEHAS = CDbl(CVDAT(HAS%))
        '
        Call ABREPEDCLI
        Call ABRESTOCKS
        '
        '\\\OT-06-0098-WAR-20/03/06///
37      R1% = COMALTER%("Seleccione la Opción de Filtro\\Todos\Por Cliente")
        If R1% = 0 Then GoTo 35
        If R1% = 2 Then
          Call SELECHO("DEUDOR1")
          NC% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
          If NC% <= 0 Then GoTo 37
        End If
        ''\\\OT-06-0098-WAR-FIN///
40      SQLX$ = " SELECT * "
        SQLX$ = SQLX$ + " FROM PEDDETA "
        SQLX$ = SQLX$ + " WHERE INT(CDBL(FechEmis)) >= " & FEDES & ""
        SQLX$ = SQLX$ + " AND INT(CDBL(FechEmis)) <= " & FEHAS & ""
        'SQLX$ = SQLX$ + " AND Status < 8 "
        If modogramy% < 1 Then SQLX$ = SQLX$ + " AND STATUS >= 0 "
        '\\\OT-06-0098-WAR-20/03/06///
        If R1% = 2 Then
          If NC% > 0 Then
            SQLX$ = SQLX$ + "AND NroClie = " & NC% & ""
          End If
        End If
        '\\\OT-06-0098-WAR-FIN///
        '
        Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
        '
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
               NUCLI7% = !NroClie
               NUPED7& = !NroPed
               CODIN7% = !CodiInt
               SQLY$ = " SELECT * "
               SQLY$ = SQLY$ + " FROM MOVISTO "
               SQLY$ = SQLY$ + " WHERE Nume_Clie = " & NUCLI7%
               SQLY$ = SQLY$ + " AND CDBL(NuPedCli) = " & NUPED7&
               SQLY$ = SQLY$ + " AND Cod_Inte = " & CODIN7%
               '
               Set STOBUS = Stocks.OpenRecordset(SQLY$, dbOpenDynaset)
               '
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
            Call REPLA(Y$, MKI$(!NroClie), 1, 2)
            Call REPLA(Y$, MKS$(!NroPed), 147, 4)
            Call REPLA(Y$, !NroOcom, 151, 16)
            Call REPLA(Y$, MKI$(!CodiInt), 167, 2)
            Call REPLA(Y$, Chr$(1), 169, 1) ' MARCA DE QUE ES UN PEDIDO
              FECHENT% = CVINT%(!FeSolEnt)
            Call REPLA(Y$, MKI$(FECHENT%), 170, 2)
              MONEORIG% = !MoneEmis
              TICAM# = !TipoCam: If TICAM# < 0.01 Then TICAM# = 1
              PRUNI# = !PREUNID
              If MONEORIG% <= 1 Then
                PRUNI# = PRUNI# / TICAM#
              End If
            Call REPLA(Y$, MKD$(PRUNI#), 172, 8)  'precio unitario en dolares
              CANTI# = (-1) * !CANPED
            Call REPLA(Y$, MKD$(CANTI#), 180, 8)
            MOTIVO$ = String(32, " ")
            ' AGREGADO A PEDIDO DE ROSSANE DE GRAMY
                If InStr(1, EMPREAC$, "GRAMY") > 0 And (!NroPed = 1460 Or !NroPed = 1447) Then
                    Call REPLA(MOTIVO$, "Repetido", 1, 8)
                End If
            Call REPLA(Y$, MOTIVO$, 188, 32)
            Call REPLA(Y$, MKD$(TICAM#), 220, 8)
             ' SI MONEDA ES PESOS (0 - 1), GUARDA PESOS Y CALCULA DOLARES
              If MONEORIG% <= 1 Then
                  NETPES# = (-1) * !ImpoTot
                  IMPDOL# = NETPES# / TICAM#
                Else
                  IMPDOL# = (-1) * !ImpoTot
                  NETPES# = IMPDOL# * TICAM#
              End If
            Call REPLA(Y$, MKD$(IMPDOL#), 228, 8)
            Call REPLA(Y$, MKD$(NETPES#), 236, 8)
              GIVA% = GRUPIVA%(!CodiInt)
              TASA# = CVS(Mid$(FILTERGETRECORD("GRUPIVA", 1, MKI$(GIVA%)), 13, 4))
              IMPIVA# = (NETPES# * TASA#) / 100
            Call REPLA(Y$, MKD$(IMPIVA#), 244, 8)
              IMPTOT# = NETPES# + IMPIVA#
            Call REPLA(Y$, MKD$(IMPTOT#), 252, 8)
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
            Call REPLA(Y$, MKD$(ANTIPES#), 268, 8)
            Call REPLA(Y$, MKD$(ANTIDOL#), 260, 8)
            '
            JJ& = JJ& + 1
            Call GRAREG("SECOMPED", Y$, JJ&)
            '
            SCP_CLI(JJ&) = !NroClie
            SCP_PED(JJ&) = !NroPed
            SCP_COD(JJ&) = !CodiInt
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
        '
        FIN& = JJ&
        For U& = 1 To FIN&
           '
           If SCP_COD(U&) < 1 Then GoTo 89
           '
           XREG$ = REGLEIDO$("SECOMPED", SCP_REG&(U&))
           CANACU# = CVD(Mid$(XREG$, 180, 8))
           '
           SQLX$ = " SELECT * "
           SQLX$ = SQLX$ + " FROM MOVISTO "
           SQLX$ = SQLX$ + " WHERE Nume_Clie = " & SCP_CLI(U&)
           SQLX$ = SQLX$ + " AND CDBL(NuPedCli) = " & SCP_PED(U&)
           SQLX$ = SQLX$ + " AND Cod_Inte = " & SCP_COD(U&)
           '
           Set STOCKTEMP = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
           '
           With STOCKTEMP
              On Error Resume Next
              .MoveFirst
              If Err > 1 Then GoTo 80
              On Error GoTo 0
              '
              Do While Not .EOF
                Y$ = REBLA$
                '
                Call REPLA(Y$, MKI$(SCP_CLI(U&)), 1, 2)
                Call REPLA(Y$, MKS$(SCP_PED(U&)), 147, 4)
                'Call REPLA(Y$, String(16, " "), 151, 16)
                NPEDX1& = SCP_PED(U&)
                Call REPLA(Y$, NOCOMCLI$(NPEDX1&), 151, 16)
 
                Call REPLA(Y$, MKI$(SCP_COD(U&)), 167, 2)
                Call REPLA(Y$, Chr$(2), 169, 1) ' MARCA DE QUE ES UN REMITO
                  FECHENT% = CVINT%(!fechMov)
                Call REPLA(Y$, MKI$(FECHENT%), 170, 2)
                  VUNI# = !ValoUnit
                  MONERE% = 1: TICAM# = 0
                  On Error Resume Next
                  MONERE% = !MoneVal
                  TICAM# = !TipoCam
                  On Error GoTo 0
                  '
                  NROX$ = !DoSecLar
                  FEX = FECHENT%
                  NCLIEN% = SCP_CLI(U&)
                  If Left$(UCase$(NROX$), 2) = "FC" Then ' solo si está facturado
                     MONERE1% = MONEMISOR%(NROX$, FEX, NCLIEN%)
                     If MONERE1% <> MONERE% Then
                       .Edit
                       !MoneVal = MONERE1%
                       .Update
                       MONERE% = MONERE1%
                     End If
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
14                  Movisto.FindFirst "INT(CDBL(FECHMOV)) = " & Int(CDbl(CDate(CVDAT(FECHENT%))))
16                    If Movisto.NoMatch = False Then
                      Movisto.Edit
                      Movisto!TipoCam = TICAM#
                      Movisto.Update
                      Movisto.FindNext "INT(CDBL(FECHMOV)) = " & Int(CDbl(CDate(CVDAT(FECHENT%))))
                      GoTo 16
                    End If
                  End If
                  If MONERE% <= 1 Then
                    VUNI# = VUNI# / TICAM#
                  End If
                Call REPLA(Y$, MKD$(VUNI#), 172, 8)
                Call REPLA(Y$, MKD$(!Cantsalid), 180, 8)
                  CANACU# = CANACU# + !Cantsalid
                Call REPLA(Y$, !DoPriCor, 188, 16)
                Call REPLA(Y$, !DoSecCor, 204, 16)
                Call REPLA(Y$, MKD$(TICAM#), 220, 8)
                ' SI MONEDA ES PESOS (0 - 1), GUARDA PESOS Y CALCULA DOLARES
                If MONERE% <= 1 Then
                    NETPES# = !Cantsalid * !ValoUnit
                    IMPDOL# = NETPES# / TICAM#
                  Else
                    IMPDOL# = !Cantsalid * !ValoUnit
                    NETPES# = IMPDOL# * TICAM#
                End If
                Call REPLA(Y$, MKD$(IMPDOL#), 228, 8)
                Call REPLA(Y$, MKD$(NETPES#), 236, 8)
                GIVA% = GRUPIVA%(SCP_COD(U&))
                  TASA# = CVS(Mid$(FILTERGETRECORD("GRUPIVA", 1, MKI$(GIVA%)), 13, 4))
                  IMPIVA# = (NETPES# * TASA#) / 100
                Call REPLA(Y$, MKD$(IMPIVA#), 244, 8)
                  IMPTOT# = NETPES# + IMPIVA#
                Call REPLA(Y$, MKD$(IMPTOT#), 252, 8)
' falta revisar de aqui en adelante
                  ' SI MONEDA ES PESOS (0 - 1), GUARDA PESOS
                  ' Y CALCULA DOLARES
                  ANTIPES# = NETPES# * COEFPES#(U&)
                  ANTIDOL# = IMPDOL# * COEFDOL#(U&)
                Call REPLA(Y$, MKD$(ANTIDOL#), 260, 8)
                Call REPLA(Y$, MKD$(ANTIPES#), 268, 8)
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
                   Call REPLA(XREG$, String$(32, 0), 228, 32)
               End If
               Call REPLA(XREG$, "Anul.", 188, 16)
               Call REPLA(XREG$, "Anul.", 204, 16)
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
          NCL% = CVI(Left$(X$, 2))
          If U& > 1 Then
            If NCL% <> NCLA% Then
              JJ& = JJ& + 1
              Call GRAREG("SECOMPED", String$(512, "="), JJ&)
            End If
          End If
          NCLA% = NCL%
          RASO$ = TRIM$(Str$(NCL%)) + " - " + RASOCLI$(NCL%)
          DOMI$ = DOMICLI$(NCL%)
          LOCA$ = LOCACLI$(NCL%)
          LIPE$ = X$
          Call REPLA(LIPE$, RASO$, 3, 48)
          Call REPLA(LIPE$, DOMI$, 51, 48)
          Call REPLA(LIPE$, LOCA$, 99, 48)
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
    modogramy% = 0
    If CONTROL$("ADMIPED", "FORMALIS") = "MONEMORI" Then
      If EXISTE%(LUDAT$ + "LIPEDID.RFS") > 0 Then
        modogramy% = 1
      End If
    End If
    '
    '\\\OT-05-0249-WAR-04/05/05///
    If modogramy% < 1 Then Call MENSPED
    '\\\OT-05-0249-WAR-FIN///
    If Option1.Value = True Then
        'Por Pantalla
        Call RECAPEDPE
        sali$ = "*COPEDNO"
        If FORMAPED% = 2 Then sali$ = "*COQEDMA"
        Call FINAL(sali$)
      Else
        'Listados
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
        Call SELRACLIEN(NCL1%, NCL2%, VENDSEL%)
        If NCL1% < 1 Then GoTo 99
        '
        'nuevo
        Call ABREPEDCLI
        SQLX$ = " SELECT * "
        SQLX$ = SQLX$ + " FROM PEDDETA "
        SQLX$ = SQLX$ + " WHERE NroClie >= " & NCL1% & " "
        SQLX$ = SQLX$ + " AND  NroClie <= " & NCL2% & " "
        SQLX$ = SQLX$ + " AND Status < 8 "
        If modogramy% < 1 Then SQLX$ = SQLX$ + " AND STATUS >= 0 "
        SQLX$ = SQLX$ + " AND (((CanPed - CantEnt) >= 0.05) OR (TipoPed = 'ABIERTO' AND CDBL(FinVigen) >= " & Int(CDbl(Now)) & ")) "
        '
        Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
        '
        With PEDTEMP
          On Error Resume Next
          .Move first
          '
          JJ& = 0
          '
          Do While (.EOF = False)
            Y$ = REGBLAN("LIPEDPE")
            CI% = !CodiInt
            If CI% > 0 And CI% <= NUMAS% Then
              CAPEN = !CANPED - !CantEnt
              If CAPEN < 0 Then CAPEN = 0
              FENTRESOL% = CVINT(!FeSolEnt)
              If FENTRESOL% >= FLI1% Then
                If FENTRESOL% <= FLI2% Then
                  If VENDSEL% < 1 Or VENDCLI%(NC%) = VENDSEL% Then
                    Y$ = String$(46, 0) + Left$(X$, 6) + Mid$(X$, 9, 24) + Mid$(X$, 75)
                    Call REPLA(Y$, Mid$(X$, 7, 2), 1, 2)
                    Call REPLA(Y$, Mid$(X$, 7, 2), 45, 2)
                    Call REPLA(Y$, Mid$(X$, 33, 42), 3, 42)
                    '
                    CANOASI = CAPEN - !CanAsig
                    If CANOASI < 0 Then CANOASI = 0
                    NC% = !NroClie
                      MONEMISPED% = !MoneEmis
                    PREUNI = !PREUNID * (100 - PORDESCLI(NC%)) / 100 * TICAMONE(MONEMISPED%)
                    If modogramy% = 1 Then
                       PREUNI = PREUNI / TICAMONE(2)
                    End If
                    '
                    If OPX% <> 2 Then PREUNI = 0
                    ITOTA# = CDbl(CAPEN * PREUNI)
                    INOAS# = CDbl(CANOASI * PREUNI)
                    '
                    Call REPLA(Y$, MKI$(NC%), 1, 2)
                    Call REPLA(Y$, MKI$(!CodiInt), 3, 2)
                    Call REPLA(Y$, MKI$(NC%), 45, 2)
                    Call REPLA(Y$, MKS$(!NroPed), 47, 4)
                    Call REPLA(Y$, MKI$(FENTRESOL%), 69, 2)
                    Call REPLA(Y$, MKS$(!CANPED), 81, 4)
                    Call REPLA(Y$, MKS$(CAPEN), 101, 4)
                    Call REPLA(Y$, MKS$(PREUNI), 85, 4)
                    Call REPLA(Y$, MKD$(ITOTA#), 89, 8)
                    Call REPLA(Y$, MKS$(!CantEnt), 97, 4)
                    Call REPLA(Y$, MKS$(CANOASI), 111, 4)
                    Call REPLA(Y$, MKD$(INOAS#), 119, 8)
                    Call REPLA(Y$, Space$(34), 127, 34)
                    JJ& = JJ& + 1
                    Call GRAREG("LIPEDPE", Y$, JJ&)
                    '
                    If modogramy% = 1 Then
                       Z$ = REBLZ$
                       Y0$ = Left$(Y$, 46) ' LOS PRIMEROS BYTES COINCIDEN
                       Call REPLA(Z$, Y0$, 1, 46)
                         NOCO$ = !NroOcom
                       Call REPLA(Z$, NOCO$, 47, 16)
                       Call REPLA(Z$, MKS$(!NroPed), 63, 4)
                       Call REPLA(Z$, MKI$(FEMIPED%), 67, 2)
                       Call REPLA(Z$, MKI$(FENTRESOL%), 69, 2)
                       If !Status < 0 Then
                          Call REPLA(Z$, "P", 71, 1)
                       End If
                       Call REPLA(Z$, MKS$(!CANPED), 81, 4)
                       Call REPLA(Z$, MKS$(!CantEnt), 85, 4)
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
        '
30      If modogramy% <> 1 Then
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

Sub CUAPROPED()
    '
    Dim PEDIPRO#(32767, 12), TIENEPED%(32767)
'    Call MESERR(24, "No Disponible en la Presente Versión")
'    Exit Sub
    '
    modogramy% = 0
    If CONTROL$("ADMIPED", "FORMALIS") = "MONEMORI" Then
       modogramy% = 1
    End If
    '
    Call GENVECMES
    VECPEDI$ = ""
    Call ABREPEDCLI
    SQLX$ = " SELECT * FROM PEDDETA "
    SQLX$ = SQLX$ + " WHERE Status < 8 "
    If modogramy% < 1 Then SQLX$ = SQLX$ + " AND STATUS >= 0 "
    SQLX$ = SQLX$ + " AND (((CanPed - CantEnt) >= 0.05) OR (TipoPed = 'ABIERTO' AND CDBL(FinVigen) >= " & Int(CDbl(Now)) & ")) "
    Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With PEDTEMP
       On Error Resume Next
       .Move first
       If Err < 1 Then
             '
           On Error GoTo 0
           Do While (.EOF = False)
              CI% = !CodiInt
              If CI% > 0 And CI% <= NUMAS% Then
                 CAPEN = !CANPED - !CantEnt: If CAPEN < 0 Then CAPEN = 0
                 NC% = !NroClie
                 MONEMISPED% = !MoneEmis
                 PREUNI = !PREUNID * (100 - PORDESCLI(NC%)) / 100 * TICAMONE(MONEMISPED%)
                 ITOTA# = CDbl(CAPEN * PREUNI)
                 If modogramy% = 1 Then
                     ITOTA# = ITOTA# / TICAMONE(2) ' LO VUELVE A PASAR A DOLARES
                 End If
                 FENTRESOL% = CVINT(!FeSolEnt)
                 '
                 If NC% > 0 Then
                    If Abs(ITOTA#) > 0 Then
                      For KKI% = 0 To 13
                        If FENTRESOL% >= INIMES%(KKI%) Then
                          If FENTRESOL% < INIMES%(KKI% + 1) Then
                            PEDIPRO#(NC%, KKI%) = PEDIPRO#(NC%, KKI%) + ITOTA#
                            TIENEPED%(NC%) = 1
                            GoTo 19
                          End If
                        End If
                      Next KKI%
                    End If
                 End If
                 '
              End If
19         .MoveNext
           Loop
           '
       End If
       On Error GoTo 0
    End With
    '
    JJ& = 0
    For NC1% = 1 To 32766
      If TIENEPED%(NC1%) > 0 Then
        TTXX$ = AJUSTI$(FANTCLI$(NC1%), 16)
        SUMATO# = 0
        For KKI% = 0 To 12
           TTXX$ = TTXX$ + MKD$(PEDIPRO#(NC1%, KKI%))
           SUMATO# = SUMATO# + PEDIPRO#(NC1%, KKI%)
        Next KKI%
        TTXX$ = TTXX$ + MKD$(SUMATO#)
        JJ& = JJ& + 1
        Call GRAREG("CUAPROPE", TTXX$, JJ&)
      End If
    Next NC1%
    Call SETULTREG("CUAPROPE", JJ&)
    '
    ' ESTO ES PARA PONER EL ENCABEZADO DEL LISTADOR
    TTXX$ = AJUSTI$("Cliente", 16)
    For KKII% = 0 To 12
      TTXX$ = TTXX$ + AJUSTD$(DENOMES$(KKII%), 12)
    Next KKII%
    TTXX$ = TTXX$ + AJUSTD$("Total", 12)
    Call SAVEFILE(LUCOM$ + "CUAPROPE.HDR", TTXX$)
    '
    Call HEADERS("CUAPROPE", "")
    TTXX$ = "Importes: Netos sin IVA expresados en "
    If modogramy% < 1 Then
        TTXX$ = TTXX$ + DEMONECO$(1)
      Else
        TTXX$ = TTXX$ + "Dolares"
    End If
    Call HEADERS("CUAPROPE", TTXX$)
    '
30  Call FINAL("*CUAPROPE")
    '
End Sub

Sub Command64_Click()
    Command64.Enabled = False
    '
    modogramy% = 0
    If CONTROL$("ADMIPED", "FORMALIS") = "MONEMORI" Then
      If EXISTE%(LUDAT$ + "LIPEDID.RFS") > 0 Then
        modogramy% = 1
      End If
    End If
    '
    MODOSILCAR% = 0
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "SILCAR") > 0 Or InStr(EPAC$, "TELSO") > 0 Then
        NX% = FILEOPEN("FLEXMUL.DRV", 0, 79)
        Dim FMX As String * 79
        '
        For I% = 1 To LOF(NX%) / 79
          Get #NX%, I%, FMX$
          SUCAR% = 0: For J% = 1 To 77: SUCAR% = SUCAR% + Asc(Mid$(FMX$, J%, 1)): Next J%
          If CVI(Mid$(FMX$, 78, 2)) = SUCAR% Then
            RASO$ = DESENCRI$(Left$(FMX$, 30))
            PACCE$ = DESENCRI$(Mid$(FMX$, 32, 46))
            If I% = 1 Then LUDAPRI$ = TRIM$(PACCE$) + "\"
            If I% = 2 Then
              If TRIM$(PACCE$) + "\" = LUDAT$ Then
                MODOSILCAR% = 1
              End If
            End If
          End If
        Next I%
        Close #NX%
    End If
    '
    '\\\OT-05-0249-WAR-04/05/05///
    If modogramy% < 1 Then Call MENSPED
    '\\\OT-05-0249-WAR-FIN///
    If Option1.Value = True Then
           'Por Pantalla
           Call RECAPEDPE
           sali$ = "*COPEDMA"
           If FORMAPED% = 2 Then sali$ = "*COQEDMA"
           Call FINAL(sali$)
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
           Call SELRACLIEN(NCL1%, NCL2%, VENDSEL%)
           If NCL1% < 1 Then GoTo 99
           '
           'nuevo
           Call ABREPEDCLI
25         SQLX$ = " SELECT * "
           SQLX$ = SQLX$ + " FROM PEDDETA "
           SQLX$ = SQLX$ + " WHERE NroClie >= " & NCL1% & " "
           SQLX$ = SQLX$ + " AND  NroClie <= " & NCL2% & " "
           SQLX$ = SQLX$ + " AND Status < 8 "
           If modogramy% < 1 Then SQLX$ = SQLX$ + " AND STATUS >= 0 "
           SQLX$ = SQLX$ + " AND (((CanPed - CantEnt) >= 0.05) OR (TipoPed = 'ABIERTO' AND CDBL(FinVigen) >= " & Int(CDbl(Now)) & ")) "
           '
           Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
           '
           With PEDTEMP
             On Error Resume Next
             .Move first
             '
             If MODOSILCAR% < 2 Then JJ& = 0
             Y$ = REGBLAN("LIPEDPE")
             If modogramy% = 1 Then
                REBLZ$ = REGBLAN$("LIPEDID")
             End If
             '
             Do While (.EOF = False)
                CI% = !CodiInt
                '20/03/07 (OT-07-0118)
                NC% = !NroClie
                '20/03/07 (FIN)
             If CI% > 0 And CI% <= NUMAS% Then
                CAPEN = !CANPED - !CantEnt
               '
               If CAPEN < 0 Then CAPEN = 0
               FEMIPED% = CVINT(!FECHEMIS)
               FENTRESOL% = CVINT(!FeSolEnt)
               '
               If FENTRESOL% >= FLI1% Then
                 If FENTRESOL% <= FLI2% Then
                   If VENDSEL% < 1 Or VENDCLI%(NC%) = VENDSEL% Then
                     Y$ = String$(46, 0) + Left$(X$, 6) + Mid$(X$, 9, 24) + Mid$(X$, 75)
                     Call REPLA(Y$, Mid$(X$, 7, 2), 1, 2)
                     Call REPLA(Y$, Mid$(X$, 7, 2), 45, 2)
                     Call REPLA(Y$, Mid$(X$, 33, 42), 3, 42)
                     CAPEN = !CANPED - !CantEnt
                     '
                     If CAPEN < 0 Then CAPEN = 0
                       CANOASI = CAPEN - !CanAsig
                     '
                     If CANOASI < 0 Then CANOASI = 0
                     'NC% = !NroClie
                       MONEMISPED% = !MoneEmis
                     PREUNI = !PREUNID * (100 - PORDESCLI(NC%)) / 100 * TICAMONE(MONEMISPED%)
                     If modogramy% = 1 Then
                        PREUNI = PREUNI / TICAMONE(2)
                     End If
                     '
                     If OPX% <> 2 Then PREUNI = 0
                     ITOTA# = CDbl(CAPEN * PREUNI)
                     INOAS# = CDbl(CANOASI * PREUNI)
                     '
                     Call REPLA(Y$, MKI$(NC%), 1, 2)
                     Call REPLA(Y$, MKI$(!CodiInt), 3, 2)
                     Call REPLA(Y$, MKI$(NC%), 45, 2)
                     Call REPLA(Y$, MKS$(!NroPed), 47, 4)
                     Call REPLA(Y$, MKI$(FENTRESOL%), 69, 2)
                     Call REPLA(Y$, MKS$(!CANPED), 81, 4)
                     Call REPLA(Y$, MKS$(CAPEN), 101, 4)
                     Call REPLA(Y$, MKS$(PREUNI), 85, 4)
                     Call REPLA(Y$, MKD$(ITOTA#), 89, 8)
                     Call REPLA(Y$, MKS$(!CantEnt), 97, 4)
                     Call REPLA(Y$, MKS$(CANOASI), 111, 4)
                     Call REPLA(Y$, MKD$(INOAS#), 119, 8)
                     Call REPLA(Y$, Space$(34), 127, 34)
                     JJ& = JJ& + 1
                     Call GRAREG("LIPEDPE", Y$, JJ&)
                     '
                     If modogramy% = 1 Then
                        Z$ = REBLZ$
                        Y0$ = Left$(Y$, 46) ' LOS PRIMEROS BYTES COINCIDEN
                        Call REPLA(Z$, Y0$, 1, 46)
                          NOCO$ = !NroOcom
                        Call REPLA(Z$, NOCO$, 47, 16)
                        Call REPLA(Z$, MKS$(!NroPed), 63, 4)
                        Call REPLA(Z$, MKI$(FEMIPED%), 67, 2)
                        Call REPLA(Z$, MKI$(FENTRESOL%), 69, 2)
                        If !Status < 0 Then
                           Call REPLA(Z$, "P", 71, 1)
                        End If
                        Call REPLA(Z$, MKS$(!CANPED), 81, 4)
                        Call REPLA(Z$, MKS$(!CantEnt), 85, 4)
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
           '
           If MODOSILCAR% = 1 Then
              MODOSILCAR% = 2
              PedEnca.Close
              PedDeta.Close
              Ventas.Close
              On Error Resume Next
              Set Ventas = OpenDatabase(LUDAPRI$ + "\BDAccess\VENTAS.MDB")
              If Err Then
                 On Error GoTo 0
                 Call FINAL("")
              End If
              On Error GoTo 0
              GoTo 25
           End If
           '
30         If modogramy% <> 1 Then
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
           If MODOSILCAR% = 2 Then
              Call FINAL("")
           End If
    End If
99  Screen.MousePointer = 1
    Command64.Enabled = True
End Sub

Sub Command8_Click()
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
    '
End Sub

Sub LISCLASI()
10  OPX% = ALTERNA%("Agrupados Número de Pedido\Agrupados por Color Base\Agrupados por Base\Agrupados por Base y Tela")
    If OPX% >= 1 And OPX% <= 4 Then
    OPC% = OPX% + 1
    If OPC% = 2 Then 'Por numero de pedido
        'Call PEDPENDAKOT(OPC%)
        Call LIENCPED
    ElseIf OPC% = 3 Then ' por color base
        Call PEDPENDAKOT(OPC%)
        GoTo 10
    ElseIf OPC% = 4 Then ' por color tela
        Call PEDPENDAKOT(OPC%)
        GoTo 10
    ElseIf OPC% = 5 Then ' por tela
        Call PEDPENDAKOT(OPC%)
        GoTo 10
    End If

        
    End If

End Sub
'
Private Sub Command9_Click()
   Command9.Enabled = False
   Call CONECRUN("ODESPA02", "Ordenes de Despacho Pedidos Clientes")
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
   If InStr(EPAC$, "SABO") > 0 Then
     If EXISTE%("ODESPA02.EXE") > 0 Then
       Command9.Enabled = True
       Command4.Enabled = False
       Label4.Enabled = False
     End If
   End If
   '
   If CONTROL$("ADMIPED", "GENDESPA") = "REVISTOCK" Then
     Command4.Enabled = False
     Label4.Enabled = False
   End If
   '
   ' RECUPERACION DE PRECIOS DE VENTA DE PEDIDOS
   HOII% = HOY(HOS$)
   FLIMI = FECHANUM("12/09/07")
   If HOII% >= FLIMI And HOII% <= FLIMI + 1 Then
      Command14.Visible = True
   End If
End Sub

Sub RECAPEDPE()
   '
   Call ABREPEDCLI
   SQLX$ = " SELECT * "
   SQLX$ = SQLX$ + " FROM PEDDETA "
   SQLX$ = SQLX$ + " WHERE Status < 8 "
   SQLX$ = SQLX$ + " AND STATUS <> -1 " ' EXCLUYE LAS NO APROBADAS
   SQLX$ = SQLX$ + " AND (((CanPed - CantEnt) >= 0.05) OR (TipoPed = 'ABIERTO' AND CDBL(FinVigen) >= " & Int(CDbl(Now)) & ")) "
   '
   Set PEDTEMP1 = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   With PEDTEMP1
     On Error Resume Next
     .Move first
     '
     V1& = 0
     Y$ = REGBLAN("PEDDETP")
     V1& = 0
     Do While (.EOF = False)
       '
       Call REPLA(Y$, MKS$(!NroPed), 1, 4)
       Call REPLA(Y$, MKI$(!NroClie), 7, 2)
         FECHAPED% = CVINT(!FECHEMIS)
       Call REPLA(Y$, MKI$(FECHAPED%), 5, 2)
       Call REPLA(Y$, MKI$(!NroClie), 7, 2)
       Call REPLA(Y$, !NroOcom, 9, 16)
         FECHASOL% = CVINT(!FeSolEnt)
       Call REPLA(Y$, MKI$(FECHASOL%), 25, 2)
       Call REPLA(Y$, Chr$(!Status), 27, 1)
       Call REPLA(Y$, Chr$(!ListPre), 28, 1)
       Call REPLA(Y$, Chr$(!CondPag), 29, 1)
       Call REPLA(Y$, Chr$(!VENDED), 30, 1)
       Call REPLA(Y$, MKI$(!PorComVend), 31, 2)
       Call REPLA(Y$, MKI$(!CodiInt), 33, 2)
       Call REPLA(Y$, MKS$(!CANPED), 79, 4)
       Call REPLA(Y$, MKS$(!PREUNID), 83, 4)
       Call REPLA(Y$, MKS$(!CantEnt), 95, 4)
         CAPEN = !CANPED - !CantEnt
       Call REPLA(Y$, MKS$(CAPEN), 99, 4)
       V1& = V1& + 1
       Call GRAREG("PEDDETP", Y$, V1&)
       '
       .MoveNext
     Loop
   End With
   '
   Call SETULTREG("PEDDETP", V1&)
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
    SQLX$ = "SELECT * FROM PEDENCA"
    SQLX$ = SQLX$ + " WHERE STATUS = -1 "
    Set PEPEAPTMP = Ventas.OpenRecordset(SQLX$, 4)
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
    '
End Sub
'\\\OT-05-0249-WAR-FIN///
'\\\OT-05-0250-WAR-19/05/05///
Sub LIENCPED()
    '
    HAS% = HOY(HOS$)
    FEX = HAS%
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    VDEF$ = MKI$(DES%) + MKI$(HAS%)
    '
10  OKKX% = 0
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
       Exit Sub
    End If
    OKKX% = 1
    DES% = CVI(Left$(OBX$, 2)): DES1 = CDbl(CVDAT(DES%))
    HAS% = CVI(Mid$(OBX$, 3, 2)): HAS1 = CDbl(CVDAT(HAS%))
    If HAS% < DES% Then GoTo 10
    FEX = DES%: PERIO$ = "Periodo: " + FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    modogramy% = 0
    If CONTROL$("ADMIPED", "FORMALIS") = "MONEMORI" Then
      If EXISTE%(LUDAT$ + "LIPEDID.RFS") > 0 Then
        modogramy% = 1
      End If
    End If
    '
    Call HEADERS("ENCPED", "")
    Call HEADERS("ENCPED", PERIO$)
    Call ABREPEDCLI
    SQLX$ = "SELECT * FROM PEDENCA"
    SQLX$ = SQLX$ + " WHERE INT(CDBL(FechEmis)) >= " & DES1 & ""
    SQLX$ = SQLX$ + " AND INT(CDBL(FechEmis)) <= " & HAS1 & ""
    SQLX$ = SQLX$ + " ORDER BY FechEmis ASC, NroPed ASC "
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
      YY$ = REGBLAN("ENCPED")
      Do While .EOF = False
        OBSE$ = TRIM$(SAFETEXT(!Observa))
        RETEX$(1) = ""
        Call FRATEXTO(OBSE$, 30)
        Call REPLA(YY$, MKS$(!NroPed), 1, 4)
        Call REPLA(YY$, MKI$(!NroClie), 5, 2)
        Call REPLA(YY$, MKI$(CVINT(!FECHEMIS)), 7, 2)
        Call REPLA(YY$, !NroOcom, 9, 16)
        '
        TDESCU$ = !TDescApli
        PORDESCU = PORCEDESCU(TDESCU$)
        '
        IMPED# = 0
        PedDeta.FindFirst (" NROPED = " & !NroPed & " ")
25      If PedDeta.NoMatch = False Then
          MONEMISPED% = PedDeta!MoneEmis
          IMPED# = IMPED# + PedDeta.Fields("IMPOTOT") * TICAMONE(MONEMISPED%)
          PedDeta.FindNext (" NROPED = " & !NroPed & " ")
          GoTo 25
        End If
        If modogramy% = 1 Then
          IMPED# = IMPED# / TICAMONE(2)  ' PARA LLEVARLO A DOLARES
        End If
        '
        Call REPLA(YY$, MKD$(IMPED# * (1 - PORDESCU / 100)), 25, 8)
        Call REPLA(YY$, MKI$(!Status), 33, 2)
        If modogramy% = 1 Then
            STX$ = "  "
            If !Status < 0 Then STX$ = "P."
            If !Status >= 8 Then STX$ = "An"
            Call REPLA(YY$, STX$, 33, 2)
        End If
        Call REPLA(YY$, RETEX$(1), 35, 28)
        JJ& = JJ& + 1
        Call GRAREG("ENCPED", YY$, JJ&)
        '
        For KKU& = 2 To CARETEX%
           YY$ = REGBLAN$("ENCPED") ' SE ABRE NUEVO REGISTRO PARA QUE AL EXPORTARLO A EXCEL NO REPITA
           Call REPLA(YY$, RETEX$(KKU&), 35, 28)
           JJ& = JJ& + 1
           Call GRAREG("ENCPED", YY$, JJ&)
        Next KKU&
      .MoveNext
      Loop
    End With
    '
    Call SETULTREG("ENCPED", JJ&)
    Call FINAL("*LIENPED")
End Sub
'\\\OT-05-0250-WAR-FIN///

Private Sub IMPUTAPEDI()
   '
   If DERACCE%("IMPUANTI", "Imputación de Anticipos") >= 2 Then

     SQLX$ = "SELECT * FROM PEDENCA "
  '   If PROLLAMA$ = "ABIE" Then
  '        SQLX$ = SQLX$ + "WHERE TipoPed = 'ABIERTO' "
  '      Else
  '        SQLX$ = SQLX$ + "WHERE TipoPed <> 'ABIERTO' "
  '        SQLX$ = SQLX$ + "OR IsNull(TipoPed) = True "
  '   End If
     SQLX$ = SQLX$ + "ORDER BY NroPed ASC"
     Set PEDENCTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     JJ& = 0
     With PEDENCTEMP
        On Error Resume Next
        .MoveFirst
        If Err < 1 Then
          Do While Not .EOF
               NRODES& = !NroPed
               TTXX$ = MKS$(NRODES&)
                 NOCO$ = !NroOcom
               If NOCO$ <> "" Then
                 TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
               End If
                 FEX = CVINT(!FECHEMIS)
               TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
                 NCLIE% = !NroClie
                 ANULX$ = "": If !Status = 9 Then ANULX$ = "(Anul) "
               TTXX$ = TTXX$ + TRIM$(RASOCLI$(NCLIE%)) + ANULX$ 'Space$(52)
               JJ& = JJ& + 1
               Call GRAREG("!INDIPEDI", TTXX$, JJ&)
          .MoveNext
          Loop
        End If
     End With
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
10   Call SELECHO("!INDIPEDI/INDICE GENERAL DE PEDIDOS")
     NPSEL& = XVALO(VALACT1$("!INDIPEDI"))
     TIPAN$ = REPLACAR$(TRIM$(VALACT2$("!INDIPEDI")), "/", "-")
     If NPSEL& > 0 Then
          JJ& = 0
          REBLA$ = REGBLAN$("CARDETPE")
          SQLX$ = "SELECT * FROM PEDDETA "
          SQLX$ = SQLX$ + "WHERE NroPed = " & NPSEL& & " "
          SQLX$ = SQLX$ + "ORDER BY NuorItem ASC "
          Set PEDDETTEMP = Ventas.OpenRecordset(SQLX$, dbOpenDynaset)
          '
          On Error Resume Next
          PEDDETTEMP.MoveFirst
          If Err < 1 Then
            Do While Not PEDDETTEMP.EOF
             'For U& = 1 To ULTREG&("!CARDETPE")
              X$ = REBLA$
                CIXI% = PEDDETTEMP!CodiInt
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
                CAN = PEDDETTEMP!CANPED
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
              Call REPLA(X$, MKS$(PEDDETTEMP!NuorItem), 125, 4)    ' NUMERO DE ORDEN
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
              PEDDETTEMP.FindFirst "NuOrItem = " & NORI&
              If PEDDETTEMP.NoMatch = False Then
                PEDDETTEMP.Edit
                On Error Resume Next
                PEDDETTEMP!ImpoAnti = ANTI#
                On Error GoTo 0
                PEDDETTEMP.Update
              End If
            Next KU&
          End If
          '
          GoTo 10
          '
      End If
   End If
   '
End Sub

Sub PEDPENDAKOT(ORDEN%)
    '
    Call DESHASFECHA(DES%, HAS%, PERIO$)
    If TRIM$(PERIO$) = "" Then Exit Sub
    '
    DES1 = Int(CDbl(CVDAT(DES%)))
    HAS1 = Int(CDbl(CVDAT(HAS%)))
    PERIO$ = "Periodo: " + PERIO$
    If TRIM$(PERIO$) = "" Then Exit Sub
    '
    '\\\OT-08-0233-OD-12/05/08///
    'SELECCION DE ZONA DE REPARTO
    OPX% = ALTERNA%("Incluir Todos los Pendientes \ Solo Pendientes Aprobados")
    If OPX% < 1 Or OPX% > 2 Then Exit Sub
30  Call SELECHO("ZONAREP")
    ZR$ = TRIM$(VALACT1$("ZONAREP"))
    DESZR$ = TRIM$(VALACT2$("ZONAREP"))
    '
    'Esta consulta devuelve todos los pedidos  filtrados pendientes
    'o solo los pendientes aprobados , por Fecha y por zona en un list de Seleccion
    SQLX$ = " SELECT DISTINCT PEDENCA.NroPed, PEDENCA.NroClie, PEDENCA.FechEmis, PEDENCA.NroOcom "
    SQLX$ = SQLX$ + " FROM PEDENCA "
    SQLX$ = SQLX$ + " INNER JOIN  PEDDETA "
    SQLX$ = SQLX$ + " ON (PEDENCA.NroPed = PEDDETA.NroPed) "
    SQLX$ = SQLX$ + " WHERE PEDDETA.CantEnt < (PEDDETA.CanPed - 0.005) "
    If OPX% = 2 Then SQLX$ = SQLX$ + " AND PEDDETA.STATUS >= 0 "  ' APROBADOS
    SQLX$ = SQLX$ + " AND PEDDETA.Status < 8 "    ' para que no este anulada ni cumplida
    SQLX$ = SQLX$ + " AND PEDENCA.Status < 9 "    ' para que no esté anulado
    SQLX$ = SQLX$ + " AND INT(CDBL(PEDDETA.FeSolEnt)) >= " & DES1 & " "
    SQLX$ = SQLX$ + " AND INT(CDBL(PEDDETA.FeSolEnt)) <= " & HAS1 & " "
    SQLX$ = SQLX$ + " ORDER BY PEDENCA.NroPed ASC "
    Set PediModifTemp = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    'Carga el list de seleccion
    SELPEDIPEN.LICOPEN.Clear
    SELPEDIPEN.Label5 = "0"
    HAYPEDIPE% = 0
    With PediModifTemp
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
         Do While Not .EOF
          NCLIE% = !NroClie
          NP% = !NroPed
          NSUC% = NUMESUC$(NP%)
          If ZR$ = "" Or ZR$ = ZONASUCENTRE$(NCLIE%, NSUC%) Then
            TTXX$ = FORMATNUM$(NP%, "F6.0") + "  "
            NOCO$ = TRIM$(!NroOcom) 'si tiene orden de compra lo imprime en el list
            If NOCO$ <> "" Then
             TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
            End If
             FEX = CVINT(!FECHEMIS)
            TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
            TTXX$ = TTXX$ + TRIM$(RASOCLI$(NCLIE%)) + Space$(52)
            SELPEDIPEN.LICOPEN.AddItem TTXX$
            HAYPEDIPE% = 1
          End If
          .MoveNext
         Loop
      End If
    End With
    '
    Screen.MousePointer = 1
    If HAYPEDIPE% > 0 Then
        SELPEDIPEN.Show 1
      Else
        Call COMUNI("No Hay Pedidos APROBADOS Pendientes\de Entrega para el Cliente Elegido.")
        Exit Sub
    End If
    '
    CANSELE% = XVALO(SELPEDIPEN.Label5)
    If CANSELE% < 1 Then
      JJ& = 0
    End If
    '
    'Muestra el selecho de grupo contable de ventas (linea)
    Call SELECHO("GRUCOVEN")
    GC% = XVALO(VALACT1$("GRUCOVEN"))
    DESGC$ = TRIM$(VALACT2$("GRUCOVEN"))
    
    '\\\OT-08-0233-OD-FIN///
    RTA1% = COMALTER%("Seleccione Formato de Presentación : \\ Resumen\Detalle")
    If RTA1% < 1 Or RTA1% > 2 Then Exit Sub
    '
    VUELTA% = 0
    Call ABREPEDCLI
    SQLX$ = "SELECT * FROM PEDDETA"
    SQLX$ = SQLX$ + " WHERE INT(CDBL(FeSolEnt)) >= " & DES1 & ""
    SQLX$ = SQLX$ + " AND INT(CDBL(FeSolEnt)) <= " & HAS1 & ""
'    If OPX% = 1 Then 'todos los pedidos menos anulados
'      SQLX$ = SQLX$ + " AND INT(CDBL(STATUS)) < " & 8 & ""
'     Else 'solo los pedidos aprobados
'      SQLX$ = SQLX$ + " AND Status = 0"
'    End If
    SQLX$ = SQLX$ + " AND CANTENT < CANPED "
    
'   AGREGA A LA CONSULTA LA CONDICION PARA QUE TRAIGA SOLO LOS
'    PEDIDOS SELECCIONADOS EN EL LIST

    For KKU& = 1 To SELPEDIPEN.LICOPEN.ListCount
      If SELPEDIPEN.LICOPEN.Selected(KKU& - 1) = True Then
         NPESEL& = XVALO(Left$(SELPEDIPEN.LICOPEN.List(KKU& - 1), 6))
         If VUELTA% < 1 Then ' si esta seleccionado el pedido lo trae a la consulta
             SQLX$ = SQLX$ + "AND ( NroPed = " & NPESEL& & " "
             VUELTA% = 1
           Else
             ' si esta seleccionado el pedido lo trae a la consulta
             SQLX$ = SQLX$ + "OR NroPed = " & NPESEL & " "
         End If
      End If
    Next KKU&
    SQLX$ = SQLX$ + ") "

    SQLX$ = SQLX$ + " ORDER BY TColor,TTalle,CBase,CodiInt "
    '
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
        '\\\OT-08-0233-OD-12/05/08///
        NC% = !NroClie
        CI% = !CodiInt
          '
          If GC% < 1 Or GC% = GRUCOVEI%(CI%) Then 'SI NO SE SELECCIONO GRUPO CLIENTE  O COINCIDE
            Call REPLA(Y$, MKS$(!NroPed), 1, 4)
             Call REPLA(Y$, MKI$(!NroClie), 5, 2)
            
             FECEMI% = CVINT(!FECHEMIS)
             Call REPLA(Y$, MKI$(FECEMI%), 7, 2)
             FECENT% = CVINT(!FeSolEnt)
             Call REPLA(Y$, MKI$(FECENT%), 9, 2)
             '
             COLBASE$ = SAFETEXT$(!CBase)
             Call REPLA(Y$, COLBASE$, 11, 16)
             '
             Call REPLA(Y$, MKI$(!CodiInt), 27, 2)
             Call REPLA(Y$, COLBASE$, 29, 16)
             '
             TELA$ = SAFETEXT$(!TTalle)
             Call REPLA(Y$, TELA$, 45, 16)
             '
              COLTELA$ = SAFETEXT$(!TColor)
             Call REPLA(Y$, COLTELA$, 61, 16)
             '
             Call REPLA(Y$, MKI$(FECEMI%), 77, 2)
             Call REPLA(Y$, MKI$(FECENT%), 79, 2)
             '
             SALDDO = (!CANPED - !CantEnt)
             If SALDDO < 0 Then SALDDO = 0
             Call REPLA(Y$, MKS$(SALDDO), 81, 4)
             JJ& = JJ& + 1
             Call GRAREG("PEPENDAK", Y$, JJ&)
          End If
          .MoveNext
      Loop
      On Error GoTo 0
    End With
    '
    Call SETULTREG("PEPENDAK", JJ&)
    Call CIERRARCH("PEPENDAK")
    Call HEADERS("PEPENDAK", "")
    Call HEADERS("PEPENDAK", PERIO$)
    Call HEADERS("PEPENDA1", "")
    Call HEADERS("PEPENDA1", PERIO$)
    '
    If DESZR$ <> "" Then
     Call HEADERS("PEPENDA1", "ZONA: " + DESZR$)
     Call HEADERS("PEPENDAK", "ZONA: " + DESZR$)
    End If
    '
    If DESGC$ <> "" Then
     Call HEADERS("PEPENDA1", "Linea: " + DESGC$)
     Call HEADERS("PEPENDAK", "Linea: " + DESGC$)
    End If

'    If ORDEN% = 2 Then ' por numero de pedido
'        Call FILESORT("PEPENDAK", "", 1, 1, "ASC")
'        Call FINAL("*PEPENDAK")
    Call COPYSTRU("PEPENDAK", "PEPENDA1")
    Call SETULTREG("PEPENDA1", 0)
    UU& = 0
    If ORDEN% = 3 Then ' por color base
        Call FILESORT("PEPENDAK", "", 7, 7, "ASC")
        Call FILESORT("PEPENDAK", "", 6, 6, "ASC")
      If RTA1% = 1 Then
        Y$ = "": VANTE$ = "": U1& = 0
        For U& = 1 To ULTREG("PEPENDAK")
           TXXT$ = REGLEIDO$("PEPENDAK", U&)
           TX$ = Mid$(TXXT$, 11, 16)
           TX$ = TX$ + Mid$(TXXT$, 27, 2)
           If UCase$(TX$) = VANTE$ Then
               TYYT$ = REGLEIDO$("PEPENDA1", U1&)
               CANNT = CVS(Mid$(TXXT$, 81, 4)) + CVS(Mid$(TYYT$, 81, 4))
               Call REPLA(TYYT$, MKS$(CANNT), 81, 4)
               Call GRAREG("PEPENDA1", TYYT$, U1&)
             Else
               U1& = U1& + 1
               Call GRAREG("PEPENDA1", TXXT$, U1&)
               VANTE$ = UCase$(TX$)
           End If
        Next U&
        Call SETULTREG("PEPENDA1", U1&)
        Call FINAL("*PEPENCB1")
        Exit Sub
      End If
      '
        Call FINAL("*PEPENCB")
    ElseIf ORDEN% = 4 Then ' por color tela
        Call FILESORT("PEPENDAK", "", 6, 6, "ASC")
        Call FILESORT("PEPENDAK", "", 7, 7, "ASC")
      If RTA1% = 1 Then
        Y$ = "": VANTE$ = "": U1& = 0
        For U& = 1 To ULTREG("PEPENDAK")
           TXXT$ = REGLEIDO$("PEPENDAK", U&)
           If UCase$(Mid$(TXXT$, 27, 18)) = VANTE$ Then
               TYYT$ = REGLEIDO$("PEPENDA1", U1&)
               CANNT = CVS(Mid$(TXXT$, 81, 4)) + CVS(Mid$(TYYT$, 81, 4))
               Call REPLA(TYYT$, MKS$(CANNT), 81, 4)
               Call GRAREG("PEPENDA1", TYYT$, U1&)
             Else
               U1& = U1& + 1
               Call GRAREG("PEPENDA1", TXXT$, U1&)
               VANTE$ = UCase$(Mid$(TXXT$, 27, 18))
           End If
           '
        Next U&
        Call SETULTREG("PEPENDA1", U1&)
        Call FINAL("*PEPENCO1")
        Exit Sub
      End If
        '
        Call FINAL("*PEPENCOD")
        '
        
    ElseIf ORDEN% = 5 Then ' por tela
        Call FILESORT("PEPENdAK", "", 10, 10, "ASC")
        Call FILESORT("PEPENDAK", "", 7, 7, "ASC")
      If RTA1% = 1 Then
        Y$ = "": VANTE$ = "": U1& = 0
        For U& = 1 To ULTREG("PEPENDAK")
           TXXT$ = REGLEIDO$("PEPENDAK", U&)
           TX$ = Mid$(TXXT$, 27, 2)
           TX$ = TX$ + Mid$(TXXT$, 29, 48)
           If UCase$(TX$) = VANTE$ Then
               TYYT$ = REGLEIDO$("PEPENDA1", U1&)
               CANNT = CVS(Mid$(TXXT$, 81, 4)) + CVS(Mid$(TYYT$, 81, 4))
               Call REPLA(TYYT$, MKS$(CANNT), 81, 4)
               Call GRAREG("PEPENDA1", TYYT$, U1&)
             Else
               U1& = U1& + 1
               Call GRAREG("PEPENDA1", TXXT$, U1&)
               VANTE$ = UCase$(TX$)
           End If
        Next U&
        Call SETULTREG("PEPENDA1", U1&)
        Call FINAL("*PEPENTE1")
        Exit Sub
      End If
        Call FINAL("*PEPENTEL")
    End If
    '
End Sub

Sub RANKPED()
    
    '\\\OT-06-0501-RG-22/11/06///
    Command8.Enabled = False
    Dim TOTVENTA#(32767)
    Dim TOTCANT#(32767)
    Dim TOTACANT#(32767)
    '
    SELVENDE.BackColor = &HC0FFC0
    SELVENDE.Frame6.BackColor = &HC0FFC0
10  SELVENDE.Show 1
    If SELVENDE.Label2 <> "OK" Then GoTo 99
    '
    OPT% = COMALTER%("Opciones de Formato de Salida\\Ranking de Pedidos\Ranking de Margen Bruto")
    If OPT% = 0 Then GoTo 99
    '
    Screen.MousePointer = 11
    CANSELE% = XVALO(SELVENDE.Label5)
    DES% = FECHANUM(SELVENDE.Label7)
    HAS% = FECHANUM(SELVENDE.Label3)
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    DESDAT = Int(CDbl(CVDAT(DES%)))
    HASDAT = Int(CDbl(CVDAT(HAS%)))
    '
    OPXX% = COMALTER%("Opcion de Ordenamiento\\Por Importe\Por Unidades")
    If OPXX% < 1 Or OPXX% > 2 Then GoTo 99

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
15        J& = 0: TOTAGEN# = 0: TOTACAN# = 0
          For U& = 0 To 32767
            TOTVENTA#(U&) = 0
            TOTCANT#(U&) = 0
            TOTACANT#(U&) = 0
          Next U&
          '
          Call ABREPEDCLI
          SQLX$ = "SELECT * FROM PEDDETA "
          SQLX$ = SQLX$ + "WHERE INT(CDBL(FechEmis)) >= " & DESDAT & " "
          SQLX$ = SQLX$ + "AND INT(CDBL(FechEmis)) <= " & HASDAT & " "
          SQLX$ = SQLX$ + "AND Status<8"
          '
          On Error GoTo 0
          Set PEDDETTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
          '
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
                NC% = !NroClie
                '\\\OT-07-0010-FG-12/01/07///
                If VENDSEL% = 0 Or VENDSEL% = VENDCLI%(NC%) Then
                '\\\OT-07-0010-FG-FIN///
                  NART% = !CodiInt
                  If NART% < 0 Then NART% = 0
                  If NART% > ARTMAX% Then ARTMAX% = NART%
                  TOTCANT#(NART%) = TOTCANT#(NART%) + !CANPED
                  NETO# = !CANPED * !PREUNID
                  If !MoneEmis > 1 Then
                     NETO# = NETO# * TICAMONE(!MoneEmis)
                  End If
                  NETO# = ROUND(NETO#, 2)
                  TOTVENTA#(NART%) = TOTVENTA#(NART%) + NETO#
                End If
             .MoveNext
             Loop
          End With
          '
          TOTAGEN# = 0
          TOTACAN# = 0
          For I% = 0 To ARTMAX%
            TOTAGEN# = TOTAGEN# + TOTVENTA#(I%)
            TOTACAN# = TOTACAN# + TOTCANT#(I%)
          Next I%
          '
          List1.Clear
          For NCI% = 0 To 32766
             If Abs(TOTVENTA#(NCI%)) > 0.005 Then
                RELIB$ = "A" + CSTRING$(MKD$(TOTAGEN# - TOTVENTA#(NCI%)), 4, 12, 2, 2)
                If OPXX% = 2 Then RELIB$ = "A" + CSTRING$(MKD$(TOTACAN# - TOTCANT#(NCI%)), 4, 12, 2, 2)
                PORPAR = 100 * TOTVENTA#(NCI%) / TOTAGEN#
                RELIB$ = RELIB$ + CSTRING$(MKI$(NCI%), 1, 6, 0, 2)
                RELIB$ = RELIB$ + CSTRING$(MKS$(PORPAR), 4, 8, 2, 2)
                RELIB$ = RELIB$ + CSTRING$(MKD$(TOTVENTA#(NCI%)), 4, 12, 2, 2)
                List1.AddItem RELIB$
             End If
          Next NCI%
          '
          J& = 0: PORACU = 0
          FIN& = List1.ListCount
          '
          For U& = 1 To FIN&
              List1.ListIndex = U& - 1
              NC% = XVALO(Mid$(List1.TEXT, 14, 6))
              TONEVEN# = XVALO(Mid$(List1.TEXT, 28, 12))
              PORPAR = XVALO(Mid$(List1.TEXT, 20, 8))
              PORACU = PORACU + PORPAR
              '
              W$ = REGBLAN$("RANMABRU")
              '
              J& = J& + 1: JI% = J&
              Call REPLA(W$, MKI$(JI%), 1, 2)
              Call REPLA(W$, MKI$(NC%), 3, 2)
              Call REPLA(W$, MKD$(TOTCANT#(NC%)), 5, 8)
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
              IMPIVA# = (TONEVEN# * PORCEIVA(NC%)) / 100
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
          If OPT% = 1 Then
               Call FINAL("*RANPEDI" + FORMAEM$)
            ElseIf OPT% = 2 Then
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

Sub FILARTCLI()
   '
15 Call SELECHO("DEUDOR1")
   NC% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
   If NC% > 0 Then
20   Call DESHASFECHA(DES%, HAS%, PERIO$)
     If TRIM$(PERIO$) = "" Then GoTo 15
     DESDAT = Int(CDbl(CVDAT(DES%)))
     HASDAT = Int(CDbl(CVDAT(HAS%)))
     '
     JJ& = 0
     Call APERBASDAT("CABAN")
     SQLX$ = "SELECT CodiCom_Lar, FechConta, ValAbsComp FROM CAJABANC "
     SQLX$ = SQLX$ + " where Int(CDbl(FechConta)) >= " & DESDAT & " "
     SQLX$ = SQLX$ + " AND Int(CDbl(FechConta)) <= " & HASDAT & " "
     SQLX$ = SQLX$ + " AND NuClien = " & NC% & ""
     SQLX$ = SQLX$ + " AND TipoMovim = 'FVEN'"
     SQLX$ = SQLX$ + " AND OpciMovim = 'FC'"
     Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
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
          FECON% = CVINT(!FechConta)
          Call REPLA(Y$, MKI(ORDEN%), 1, 2)
          Call REPLA(Y$, !CODICOM_LAR + "  -  " & FECHATEX$(CVINT(!FechConta)) & "  -  $ " & FORMATNUM$(!ValAbsComp, "F12.2"), 3, 56)
          Call GRAREG("INDICOM", Y$, JJ&)
         .MoveNext
         Loop
     End With
     '
     Call SETULTREG("INDICOM", JJ&)
30   Call FRESHECHO("INDICOM")
     Call SELECHO("INDICOM")
     
     NUMFACT$ = (TRIM$(VALACT2$("INDICOM")))
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
        Set MOVISTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
        With MOVISTMP
         On Error Resume Next
         .MoveFirst
         If Err > 0 Then
              Call MENSERR(24, "No se Encuentran Pedidos Relacionados con esa Factura ")
          GoTo 30
         End If
         Y$ = REGBLAN("INDIPEP")
         On Error GoTo 0
         Do While Not .EOF
            NPEDI& = !NuPedCli
            SQLY$ = "SELECT FechEmis, Referencia FROM PEDENCA WHERE nroPed = " & NPEDI&
            Set ENCAPEDI = Ventas.OpenRecordset(SQLY$, dbOpenSnapshot)
            RAZSOC$ = ""
            On Error Resume Next
            ENCAPEDI.MoveFirst
            If Err < 1 Then
               On Error GoTo 0
               RAZSOC$ = FECHATEX$(CVINT(ENCAPEDI!FECHEMIS)) + " - " + ENCAPEDI!Referencia
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
   Call CIERRARCH("INDICOM")
   Call FRESHECHO("INDICOM")
9999 Screen.MousePointer = 1
   '
End Sub
Sub PEDAFACT()
15 Call SELECHO("DEUDOR1")
   NC% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
   If NC% > 0 Then
20   Call DESHASFECHA(DES%, HAS%, PERIO$)
     If TRIM$(PERIO$) = "" Then GoTo 15
     DESDAT = Int(CDbl(CVDAT(DES%)))
     HASDAT = Int(CDbl(CVDAT(HAS%)))
     '
     SQLX$ = "SELECT  DISTINCT NroPed , NroClie, FechEmis FROM PEDENCA "
     SQLX$ = SQLX$ + " where Int(CDbl(FechEmis)) >= " & DESDAT & " "
     SQLX$ = SQLX$ + " AND Int(CDbl(FechEmis)) <= " & HASDAT & " "
     SQLX$ = SQLX$ + " AND NroClie = " & NC% & ""
     '
     JJ& = 0
     Set PEDENTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
     With PEDENTEMP
       On Error Resume Next
         .MoveFirst
         If Err > 0 Then
          Call MENSERR(24, "No se Encuentran Pedidos Coincidentes con el Rango Elegido ")
          GoTo 20
         End If
         Y$ = REGBLAN("INDIPEP")
         Call BLANARCH("INDIPEP")
         On Error GoTo 0
         Do While Not .EOF
            NPEDI& = !NroPed
            NUMPEDI$ = MKS$(NPEDI&)
            NUMCLIEN$ = !NroClie
            RAZSOC$ = RASOCLI$(NUMCLIEN$)
            'CARGAR EL ARCHIVO PARA SELECCION
            Call REPLA(Y$, NUMPEDI$, 1, 4)
            Call REPLA(Y$, RAZSOC$ & " - " & !FECHEMIS, 5, 48)
            JJ& = JJ& + 1
            Call GRAREG("INDIPEP", Y$, JJ&)
            .MoveNext
         Loop
        End With
     Call SETULTREG("INDIPEP", JJ&)
30   Call FRESHECHO("INDIPEP")
     Call SELECHO("INDIPEP/Indice De Pedidos .")
     NUMEPEDI = XVALO(VALACT1$("INDIPEP"))
     If XVALO(NUMEPEDI) <= 0 Then GoTo 20
     NUMEP% = NUMEPEDI
     '
     JJ& = 0
     Call ABRESTOCKS
     SQLX$ = "SELECT  DISTINCT  NuPedCli ,DOSECLAR ,FechMov  FROM MOVISTO "
     SQLX$ = SQLX$ + " WHERE NuPedCli = " & NUMEP% & ""
     Set MOVISTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     With MOVISTMP
      On Error Resume Next
      .MoveFirst
      If Err > 0 Then
       Call MENSERR(24, "No se Encuentran Pedidos en el Rango Elegido ")
       GoTo 30
      End If
      Y$ = REGBLAN("INDICOM")
      Call BLANARCH("INDICOM")
      On Error GoTo 0
      Do While Not .EOF
        '***
        JJ& = JJ& + 1
        ORDEN% = JJ&
        'CARGAR EL ARCHIVO PARA SELECCION
        Call REPLA(Y$, MKI(ORDEN%), 1, 2)
        DSLAR$ = SAFETEXT$(!DoSecLar)
        FEC% = CVINT%(!fechMov)
        FEMO = FECHATEX$(FEC%)
        If InStr(DSLAR$, "P/F") = 1 Then DSLAR$ = " Pendientes de Facturación "
        Call REPLA(Y$, DSLAR$ & " - " & FEMO, 3, 36)
        Call GRAREG("INDICOM", Y$, JJ&)
        .MoveNext
      Loop
     End With
     '
     Call SETULTREG("INDICOM", JJ&)
     Call FRESHECHO("INDICOM")
40   Call SELECHO("INDICOM")
     NFACT$ = TRIM(VALACT2$("INDICOM"))
     NFACTU$ = Mid(NFACT$, 9, 8)
     If NFACTU$ = "" Then GoTo 30
     RAZSOCIAL$ = TRIM$(Left$(RASOCLI$(XVALO(NUMCLIEN$)), 8))
     
     Call VerFactElec(NFACTU$, RAZSOCIAL$) ' Muestra formulario electronico
     GoTo 40
   End If
End Sub

Private Sub IMPUTAPEDI_VIE()
   '
   SQLX$ = "SELECT * FROM PEDENCA "
'   If PROLLAMA$ = "ABIE" Then
'        SQLX$ = SQLX$ + "WHERE TipoPed = 'ABIERTO' "
'      Else
'        SQLX$ = SQLX$ + "WHERE TipoPed <> 'ABIERTO' "
'        SQLX$ = SQLX$ + "OR IsNull(TipoPed) = True "
'   End If
   SQLX$ = SQLX$ + "ORDER BY NroPed ASC"
   Set PEDENCTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   JJ& = 0
   With PEDENCTEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        Do While Not .EOF
             NRODES& = !NroPed
             TTXX$ = MKS$(NRODES&)
               NOCO$ = !NroOcom
             If NOCO$ <> "" Then
               TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
             End If
               FEX = CVINT(!FECHEMIS)
             TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
               NCLIE% = !NroClie
               ANULX$ = "": If !Status = 9 Then ANULX$ = "(Anul) "
             TTXX$ = TTXX$ + TRIM$(RASOCLI$(NCLIE%)) + ANULX$ 'Space$(52)
             JJ& = JJ& + 1
             Call GRAREG("!INDIPEDI", TTXX$, JJ&)
        .MoveNext
        Loop
      End If
   End With
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
        Set PEDDETTEMP = Ventas.OpenRecordset(SQLX$, dbOpenDynaset)
        '
        On Error Resume Next
        PEDDETTEMP.MoveFirst
        If Err < 1 Then
          Do While Not PEDDETTEMP.EOF
           'For U& = 1 To ULTREG&("!CARDETPE")
            X$ = REBLA$
              CIXI% = PEDDETTEMP!CodiInt
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
              CAN = PEDDETTEMP!CANPED
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
            Call REPLA(X$, MKS$(PEDDETTEMP!NuorItem), 125, 4)    ' NUMERO DE ORDEN
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
            PEDDETTEMP.FindFirst "NuOrItem = " & NORI&
            If PEDDETTEMP.NoMatch = False Then
              PEDDETTEMP.Edit
              On Error Resume Next
              PEDDETTEMP!ImpoAnti = ANTI#
              On Error GoTo 0
              PEDDETTEMP.Update
            End If
          Next KU&
        End If
        '
        GoTo 10
        '
    End If
   '
End Sub


Sub PEDPENDAKOT_VIE(ORDEN%)
    
    '
    Call DESHASFECHA(DES%, HAS%, PERIO$)
    If TRIM$(PERIO$) = "" Then Exit Sub
    '
    DES1 = Int(CDbl(CVDAT(DES%)))
    HAS1 = Int(CDbl(CVDAT(HAS%)))
    PERIO$ = "Periodo: " + PERIO$
    If TRIM$(PERIO$) = "" Then Exit Sub
    '
    
    '\\\OT-08-0233-OD-12/05/08///
    'SELECCION DE ZONA DE REPARTO
    OPX% = ALTERNA%("Todos los Pendientes \ Solo Aprobados")
    If OPX% < 1 Or OPX% > 2 Then Exit Sub
30  Call SELECHO("ZONAREP")
    ZR$ = TRIM$(VALACT1$("ZONAREP"))
    Call SELECHO("GRUCOVEN")
    GC$ = TRIM$(VALACT1$("GRUCOVEN"))
    DESGC$ = TRIM$(VALACT2$("GRUCOVEN"))
    
    '\\\OT-08-0233-OD-FIN///
    RTA1% = COMALTER%("Seleccione Formato de Presentación : \\ Resumen\Detalle")
    If RTA1% < 1 Or RTA1% > 2 Then Exit Sub
    '
    Call ABREPEDCLI
    SQLX$ = "SELECT * FROM PEDDETA"
    SQLX$ = SQLX$ + " WHERE INT(CDBL(FeSolEnt)) >= " & DES1 & ""
    SQLX$ = SQLX$ + " AND INT(CDBL(FeSolEnt)) <= " & HAS1 & ""
    SQLX$ = SQLX$ + " AND INT(CDBL(STATUS)) < " & 8 & ""
    SQLX$ = SQLX$ + " AND CANTENT < CANPED "
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
        '\\\OT-08-0233-OD-12/05/08///
        NC% = !NroClie 'SI SE SELECCIONO FILTRA
        CI% = !CodiInt

        If ZR$ = "" Then GoTo 20 'SI NO SE SELECCIONO ZONA REPARTO
          If ZR$ = ZONACLI$(NC%) Then
             '\\\OT-08-0233-OD-FIN///
20          If GC$ = "" Then GoTo 40 'SI NO SE SELECCIONO GRUPO CLIENTE
            If GC$ = GRUCOVEI%(CI%) Then
40           Call REPLA(Y$, MKS$(!NroPed), 1, 4)
             Call REPLA(Y$, MKI$(!NroClie), 5, 2)
            
             FECEMI% = CVINT(!FECHEMIS)
             Call REPLA(Y$, MKI$(FECEMI%), 7, 2)
             FECENT% = CVINT(!FeSolEnt)
             Call REPLA(Y$, MKI$(FECENT%), 9, 2)
             '
             COLBASE$ = SAFETEXT$(!CBase)
             Call REPLA(Y$, COLBASE$, 11, 16)
             '
             Call REPLA(Y$, MKI$(!CodiInt), 27, 2)
             Call REPLA(Y$, COLBASE$, 29, 16)
             '
             TELA$ = SAFETEXT$(!TTalle)
             Call REPLA(Y$, TELA$, 45, 16)
             '
              COLTELA$ = SAFETEXT$(!TColor)
             Call REPLA(Y$, COLTELA$, 61, 16)
             '
             Call REPLA(Y$, MKI$(FECEMI%), 77, 2)
             Call REPLA(Y$, MKI$(FECENT%), 79, 2)
             '
             SALDDO = (!CANPED - !CantEnt)
             If SALDDO < 0 Then SALDDO = 0
             Call REPLA(Y$, MKS$(SALDDO), 81, 4)
             JJ& = JJ& + 1
             Call GRAREG("PEPENDAK", Y$, JJ&)
            End If
          End If
          .MoveNext
      Loop
      On Error GoTo 0
    End With
    '
    Call SETULTREG("PEPENDAK", JJ&)
    Call CIERRARCH("PEPENDAK")
    Call HEADERS("PEPENDAK", "")
    Call HEADERS("PEPENDAK", PERIO$)
    Call HEADERS("PEPENDA1", "")
    Call HEADERS("PEPENDA1", PERIO$)
    Call HEADERS("PEPENDA1", "")
    Call HEADERS("PEPENDA1", "Linea : " + DESGC$)

    '
'    If ORDEN% = 2 Then ' por numero de pedido
'        Call FILESORT("PEPENDAK", "", 1, 1, "ASC")
'        Call FINAL("*PEPENDAK")
    Call COPYSTRU("PEPENDAK", "PEPENDA1")
    Call SETULTREG("PEPENDA1", 0)
    UU& = 0
    If ORDEN% = 3 Then ' por color base
        Call FILESORT("PEPENDAK", "", 7, 7, "ASC")
        Call FILESORT("PEPENDAK", "", 6, 6, "ASC")
      If RTA1% = 1 Then
        Y$ = "": VANTE$ = "": U1& = 0
        For U& = 1 To ULTREG("PEPENDAK")
           TXXT$ = REGLEIDO$("PEPENDAK", U&)
           TX$ = Mid$(TXXT$, 11, 16)
           TX$ = TX$ + Mid$(TXXT$, 27, 2)
           If UCase$(TX$) = VANTE$ Then
               TYYT$ = REGLEIDO$("PEPENDA1", U1&)
               CANNT = CVS(Mid$(TXXT$, 81, 4)) + CVS(Mid$(TYYT$, 81, 4))
               Call REPLA(TYYT$, MKS$(CANNT), 81, 4)
               Call GRAREG("PEPENDA1", TYYT$, U1&)
             Else
               U1& = U1& + 1
               Call GRAREG("PEPENDA1", TXXT$, U1&)
               VANTE$ = UCase$(TX$)
           End If
        Next U&
        Call SETULTREG("PEPENDA1", U1&)
        Call FINAL("*PEPENCB1")
        Exit Sub
      End If
      '
        Call FINAL("*PEPENCB")
    ElseIf ORDEN% = 4 Then ' por color tela
        Call FILESORT("PEPENDAK", "", 6, 6, "ASC")
        Call FILESORT("PEPENDAK", "", 7, 7, "ASC")
      If RTA1% = 1 Then
        Y$ = "": VANTE$ = "": U1& = 0
        For U& = 1 To ULTREG("PEPENDAK")
           TXXT$ = REGLEIDO$("PEPENDAK", U&)
           If UCase$(Mid$(TXXT$, 27, 18)) = VANTE$ Then
               TYYT$ = REGLEIDO$("PEPENDA1", U1&)
               CANNT = CVS(Mid$(TXXT$, 81, 4)) + CVS(Mid$(TYYT$, 81, 4))
               Call REPLA(TYYT$, MKS$(CANNT), 81, 4)
               Call GRAREG("PEPENDA1", TYYT$, U1&)
             Else
               U1& = U1& + 1
               Call GRAREG("PEPENDA1", TXXT$, U1&)
               VANTE$ = UCase$(Mid$(TXXT$, 27, 18))
           End If
           '
        Next U&
        Call SETULTREG("PEPENDA1", U1&)
        Call FINAL("*PEPENCO1")
        Exit Sub
      End If
        '
        Call FINAL("*PEPENCOD")
        '
        
    ElseIf ORDEN% = 5 Then ' por tela
        Call FILESORT("PEPENdAK", "", 10, 10, "ASC")
        Call FILESORT("PEPENDAK", "", 7, 7, "ASC")
      If RTA1% = 1 Then
        Y$ = "": VANTE$ = "": U1& = 0
        For U& = 1 To ULTREG("PEPENDAK")
           TXXT$ = REGLEIDO$("PEPENDAK", U&)
           TX$ = Mid$(TXXT$, 27, 2)
           TX$ = TX$ + Mid$(TXXT$, 29, 48)
           If UCase$(TX$) = VANTE$ Then
               TYYT$ = REGLEIDO$("PEPENDA1", U1&)
               CANNT = CVS(Mid$(TXXT$, 81, 4)) + CVS(Mid$(TYYT$, 81, 4))
               Call REPLA(TYYT$, MKS$(CANNT), 81, 4)
               Call GRAREG("PEPENDA1", TYYT$, U1&)
             Else
               U1& = U1& + 1
               Call GRAREG("PEPENDA1", TXXT$, U1&)
               VANTE$ = UCase$(TX$)
           End If
        Next U&
        Call SETULTREG("PEPENDA1", U1&)
        Call FINAL("*PEPENTE1")
        Exit Sub
      End If
        Call FINAL("*PEPENTEL")
    End If
    '
End Sub

Sub CAMBIAPED_VIE()
     '
2100 If YAINDI% = 0 Then
       '
       Screen.MousePointer = 11
       
       Call ABREPEDCLI
       SQLX$ = " SELECT DISTINCT PEDENCA.NroPed, PEDENCA.NroClie, PEDENCA.FechEmis, PEDENCA.NroOcom "
       SQLX$ = SQLX$ + " FROM PEDENCA "
       SQLX$ = SQLX$ + " INNER JOIN  PEDDETA "
       SQLX$ = SQLX$ + " ON (PEDENCA.NroPed = PEDDETA.NroPed) "
       SQLX$ = SQLX$ + " WHERE PEDDETA.Status < 8 "   ' para que no este anulada ni cumplida
       SQLX$ = SQLX$ + " AND PEDDETA.CantEnt < (PEDDETA.CanPed - 0.005) "
       SQLX$ = SQLX$ + " AND PEDENCA.Status < 9 "
       SQLX$ = SQLX$ + " ORDER BY PEDENCA.NroPed ASC "
       Set PediModifTemp = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
       '
       JJ& = 0
       With PediModifTemp
         On Error Resume Next
         .MoveFirst
         If Err < 1 Then
           Do While Not .EOF
               NRODES& = !NroPed
             TTXX$ = MKS$(NRODES&)
               NOCO$ = !NroOcom
             If NOCO$ <> "" Then
               TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
             End If
               FEX = CVINT(!FECHEMIS)
             TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
               NCLIE% = !NroClie
             TTXX$ = TTXX$ + TRIM$(RASOCLI$(NCLIE%)) + Space$(52)
             JJ& = JJ& + 1
             Call GRAREG("INDIPEP", Left$(TTXX$, 52) + MKS$(REGENCA&), JJ&)
           .MoveNext
           Loop
         End If
       End With
       '
       Call SETULTREG("INDIPEP", JJ&)
       Call CIERRARCH("INDIPEP")
       Call FRESHECHO("INDIPEP")
       Call CIERRARCH("*.*")
       Call BAJALDISCO
       '
       Screen.MousePointer = 1
       YAINDI% = 1
     End If
     '
2112 If ULTREG&("INDIPEP") < 1 Then
       Call COMUNI("No hay Pedidos Abiertos\que Puedan Modificarse")
       GoTo 9999
     End If
     '
2110 X$ = REGSEL$("INDIPEP/Pedidos Pendientes de Entrega")
     If Len(X$) <= 4 Then
        GoTo 9999
     End If
     '
     Screen.MousePointer = 11
     NUPEDIDO& = CVS(Left$(X$, 4))
     SQLX$ = "SELECT * FROM PEDENCA "
     SQLX$ = SQLX$ + " WHERE NroPed = " & NUPEDIDO&
     Set PEDENCATEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     On Error Resume Next
     With PEDENCATEMP
       .MoveFirst
       If Err > 0 Then
         Call MENSERR(24, "Imposible Modificar Este Pedido")
         GoTo 9999
       End If
       '
       '23/03/07 (OT-07-0121)'NO PERMITE MODIFICAR UN PEDIDO A UN CLIENTE CON LA
       NC% = !NroClie        'CUENTA SUSPENDIDA
       LMCR# = LIMICRE(NC%)
       If LMCR# < 0 Then
          Call COMUNI("Imposible Modificar Pedido.\ \La Cuenta esta Suspendida")
          GoTo 9999
       End If
       '23/03/07 (FIN)
       '
         FEX = CVINT(!FECHEMIS)
       MODPED04.Text9.TEXT = FECHATEX$(FEX)
       MODPED04.Text11.TEXT = TRIM$(Str$(NUPEDIDO&))
         NCLIE% = !NroClie
       MODPED04.LTEXT1 = TRIM$(Str$(NCLIE%))
       MODPED04.Text1.TEXT = TRIM$(Str$(NCLIE%))
       '
         NC% = NCLIE%
       MODPED04.Text2.TEXT = RASOCLI$(NC%)
       MODPED04.Text3.TEXT = DOMICLI$(NC%)
       MODPED04.Text4.TEXT = CPOSCLI$(NC%)
       MODPED04.Text5.TEXT = LOCACLI$(NC%)
       MODPED04.Text7.TEXT = CSTRING$(MKS$(LIMICRE(NC%)), 3, 13, 2, 2)
       MODPED04.Text12.TEXT = TRIM$(POSIVAC$(NC%)) + " - CUIT " + CUITCLI$(NC%)
         CREDISPO# = LIMICRE(NC%) - SALDCLI#(NC%) - CHEPEACRE#(NC%)
       MODPED04.Text8.ForeColor = RGB(0, 0, 0)
         If CREDISPO# < 0 Then MODPED04.Text8.ForeColor = RGB(255, 0, 0)
       MODPED04.Text8.TEXT = CSTRING$(MKD$(Abs(CREDISPO#)), 3, 13, 2, 2)
       '
       MODPED04.Text10.TEXT = !NroOcom
          LPREX% = !ListPre
       MODPED04.TEXT21(0).TEXT = TRIM$(Str$(LPREX%))
       MODPED04.TEXT21(1).TEXT = TRIM$(Str$(!CondPag))
         If XVALO(MODPED04.TEXT21(1).TEXT) < 1 Then MODPED04.TEXT21(1).TEXT = Str$(CPAGCLI%(NC%))
       MODPED04.TEXT21(2).TEXT = TRIM$(Str$(!VENDED))
         If XVALO(MODPED04.TEXT21(2).TEXT) < 1 Then MODPED04.TEXT21(2).TEXT = Str$(VENDCLI%(NC%))
         If XVALO(MODPED04.TEXT21(2).TEXT) < 1 Then MODPED04.Eco21(2).TEXT = "GERENCIA"
         '
       MODPED04.PORDESCU.TEXT = !TDescApli
       MODPED04.LENTREGA(0) = !DOMIENT
       MODPED04.LENTREGA(1) = !LOCAENT
       MODPED04.LENTREGA(2) = !TranspEnt
       '
       MODPED04.Text19.TEXT = !NTranspo
       MODPED04.Text20.TEXT = !NSucent
       '
       MODPED04.Text13.TEXT = !Observa
       MODPED04.Label11.ForeColor = RGB(0, 0, 255)
       MODPED04.Label11 = "Venta"
       '
     End With
     '
     Call SETULTREG("!CARDETPE", 0)
     JJ& = 0
     REBLA$ = REGBLAN$("CARDETPE")
     SQLX$ = "SELECT * FROM PEDDETA "
     SQLX$ = SQLX$ + "WHERE NroPed = " & NUPEDIDO& & " "
     '\\\OT-06-0520-RG-01/12/06///
     SQLX$ = SQLX$ + "AND Status < 8 "
     '\\\OT-06-0520-RG-FIN///
     SQLX$ = SQLX$ + "ORDER BY NuorItem ASC "
     Set PEDDETTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
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
               MODPED04.MONEMIS = DEMONECO$(MONEMI%)
               MODPED04.TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
         End If
         X$ = REBLA$
         DLISTA = 0: DGRUPO = 0: DTOTA = 0
         CIXI% = PEDDETTEMP!CodiInt
         Call REPLA(X$, MKI$(CIXI%), 1, 2)
         Call REPLA(X$, PEDDETTEMP!TColor, 3, 16)
         Call REPLA(X$, PEDDETTEMP!TTalle, 19, 8)
           TLAR = PEDDETTEMP!TLargo
         Call REPLA(X$, MKS$(TLAR), 27, 4)
           TCHO = PEDDETTEMP!TAncho
         Call REPLA(X$, MKS$(TCHO), 31, 4)
           TESP = PEDDETTEMP!TEspesor
         Call REPLA(X$, MKS$(TESP), 35, 4)
         If CONTROL$("PEDIDO", "VARCOLOR") = "SI" Then  ' si la opcion de pedidos por variante esta en true
             Call REPLA(X$, SAFETEXT$(PEDDETTEMP!CBase), 3, 12)
             Call REPLA(X$, SAFETEXT$(PEDDETTEMP!TTalle), 15, 12)
             Call REPLA(X$, SAFETEXT$(PEDDETTEMP!TColor), 27, 12)
         End If
           PREL = PEDDETTEMP!PreList * TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
         Call REPLA(X$, MKS$(PREL), 39, 4)
           PREU = PEDDETTEMP!PREUNID * TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
         Call REPLA(X$, MKS$(PREU), 51, 4)
           DLISTA = XVALO(PEDDETTEMP!DescLista)
           If DLISTA = 0 Then DLISTA = DESCUENLIS(LPREX%)
         Call REPLA(X$, MKS$(DLISTA), 109, 4)
           DTOTA = XVALO(PEDDETTEMP!DescTotal)
           If DTOTA = 0 Then
              If PREL > 0.005 Then DTOTA = 100 * (1 - PREU / PREL)
           End If
         Call REPLA(X$, MKS$(DTOTA), 69, 4)
           DGRUPO = XVALO(PEDDETTEMP!DescGrupo)
           If DGRUPO = 0 Then
              COEFITO# = (100 - DTOTA) / 100
              COEFILI# = (100 - DLISTA) / 100
              COEFIGR# = 0: If COEFILI# > 0 Then COEFIGR# = COEFITO# / COEFILI#
              DGRUPO = 100 * (1 - COEFIGR#)
           End If
         Call REPLA(X$, MKS$(DGRUPO), 113, 4)
           CAN = PEDDETTEMP!CANPED
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
         If SUCOD$ <> "" Then Call REPLA(X$, SUCOD$, 3, 8)
         JJ& = JJ& + 1
         Call GRAREG("!CARDETPE", X$, JJ&)
         '
       PEDDETTEMP.MoveNext
       Loop
     End If
     Call SETULTREG("!CARDETPE", JJ&)
     Call CIERRARCH("CARDETPE")
     Call SETULTREG("!CARDETP1", JJ&)
     Call CIERRARCH("CARDETP1")
     On Error GoTo 0
     '
     ' REVISAR Y ACTUALIZAR CAENTAN
     Dim CANENTAN(32767), FULENT%(32767)
     CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
     If CMO$ = "" Then CMO$ = "RT"
     NMS% = NUMAS%
     '
     Call ABRESTOCKS
     SQLX$ = "SELECT * FROM MOVISTO "
     SQLX$ = SQLX$ + " WHERE CodiMovi = '" & CMO$ & "' "
     SQLX$ = SQLX$ + " AND NuPedCli = " & NUPEDIDO& & " "
     Set MOVITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     On Error Resume Next
     With MOVITMP
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           CIXXI% = !Cod_Inte
           If CIXXI% > 0 Then
             If CIXXI% <= NMS% Then
               CANENTAN(CIXXI%) = CANENTAN(CIXXI%) + !Cantsalid - !CantIngre
               If !Cantsalid >= 0.0005 Then
                 If CVINT(!fechMov) > FULENT%(CIXXI%) Then
                   FULENT%(CIXXI%) = CVINT(!fechMov)
                 End If
               End If
             End If
           End If
         .MoveNext
         Loop
       End If
     End With
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
     MODPED04.Text6.TEXT = ""
     '
     For U& = 1 To Len(RNC$) Step 4
       RENOTA& = CVS(Mid$(RNC$, U&, 4))
       If RENOTA& > 0 Then
         If RENOTA& <= URN& Then
           TTYY$ = RTrim$(Mid$(REGLEIDO$("NOTACLI", RENOTA&), 3))
           If MODPED04.Text6.TEXT <> "" Then
             MODPED04.Text6.TEXT = MODPED04.Text6.TEXT + Chr$(13) + Chr$(10)
           End If
           MODPED04.Text6.TEXT = MODPED04.Text6.TEXT + RTrim$(TTYY$)
         End If
       End If
     Next U&
     '
8000 MODPED04.MARCONOTA.Visible = True
     On Error Resume Next
     MODPED04.Command6.SetFocus
     On Error GoTo 0
     Screen.MousePointer = 1
     '
     NOFRESHPRE% = 1
     MODPED04.Show 1
     NOFRESHPRE% = 0
     '
     GoTo 2110
     '
9999 Call CIERRARCH("*.*")
     Screen.MousePointer = 1
     '
End Sub
'

