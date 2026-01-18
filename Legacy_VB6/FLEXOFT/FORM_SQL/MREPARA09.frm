VERSION 5.00
Begin VB.Form MREPARA09 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reparaciones - Actualización y Consultas"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   7740
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MREPARA09.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   7740
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      Caption         =   "REPARACIONES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3345
      Left            =   3720
      TabIndex        =   0
      Top             =   1080
      Width           =   3945
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   1400
         ScaleHeight     =   615
         ScaleWidth      =   2295
         TabIndex        =   25
         Top             =   2450
         Width           =   2325
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
            Height          =   630
            Left            =   1740
            Picture         =   "MREPARA09.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Revertir Estado de Reparaciòn (de 5 a 4)"
            Top             =   0
            Visible         =   0   'False
            Width           =   560
         End
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
            Height          =   630
            Left            =   0
            Picture         =   "MREPARA09.frx":0454
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Entrega de Reparación"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Entrega de Reparación"
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
            Top             =   60
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   1070
         ScaleHeight     =   615
         ScaleWidth      =   2295
         TabIndex        =   1
         Top             =   1950
         Width           =   2325
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
            Left            =   0
            Picture         =   "MREPARA09.frx":075E
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Cierre de Orden de Reparación"
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
            TabIndex        =   3
            Top             =   60
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   800
         ScaleHeight     =   615
         ScaleWidth      =   2295
         TabIndex        =   32
         Top             =   1450
         Width           =   2325
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
            Height          =   630
            Left            =   0
            Picture         =   "MREPARA09.frx":0A68
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Describir Falla y Reparación"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Describir Falla y Reparación"
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
            TabIndex        =   33
            Top             =   60
            Width           =   1485
         End
      End
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   500
         ScaleHeight     =   615
         ScaleWidth      =   2295
         TabIndex        =   29
         Top             =   950
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
            Height          =   630
            Left            =   0
            Picture         =   "MREPARA09.frx":0BB2
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Transferencia Materiales de Depósito "
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Transferencia de Materiales"
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
            TabIndex        =   30
            Top             =   60
            Width           =   1485
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   240
         ScaleHeight     =   615
         ScaleWidth      =   2295
         TabIndex        =   4
         Top             =   450
         Width           =   2325
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
            Picture         =   "MREPARA09.frx":0CFC
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Generar una Orden de Reparación"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Generar Orden de Reparación"
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
            TabIndex        =   6
            Top             =   70
            Width           =   1485
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   3720
      ScaleHeight     =   1035
      ScaleWidth      =   4005
      TabIndex        =   21
      Top             =   0
      Width           =   4035
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   950
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   38
         Top             =   650
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   39
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command8 
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
         Picture         =   "MREPARA09.frx":1006
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Control Final de Reparaciones"
         Top             =   180
         Visible         =   0   'False
         Width           =   650
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
         Left            =   1800
         Picture         =   "MREPARA09.frx":1310
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   180
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command1 
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
         Height          =   650
         Left            =   3230
         Picture         =   "MREPARA09.frx":161A
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   180
         Width           =   650
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Setup"
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
         Left            =   2500
         Picture         =   "MREPARA09.frx":1764
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   180
         Visible         =   0   'False
         Width           =   650
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2D3C0&
      Caption         =   "CONSULTAS - REIMPRESION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Left            =   120
      TabIndex        =   15
      Top             =   2640
      Width           =   3480
      Begin VB.OptionButton Option6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Estado de Reparaciones"
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
         Height          =   380
         Left            =   525
         TabIndex        =   37
         Top             =   1050
         Width           =   2685
      End
      Begin VB.OptionButton Option5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Notas de Entrega (Remito)"
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
         Height          =   380
         Left            =   480
         TabIndex        =   28
         Top             =   1365
         Width           =   2685
      End
      Begin VB.OptionButton Option4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Orden de Reparación"
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
         Height          =   435
         Left            =   1950
         TabIndex        =   19
         Top             =   630
         Width           =   1380
      End
      Begin VB.OptionButton Option3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Presupuesto al Cliente"
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
         Height          =   435
         Left            =   1950
         TabIndex        =   18
         Top             =   210
         Width           =   1380
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Presupuesto Interno "
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
         Height          =   435
         Left            =   210
         TabIndex        =   17
         Top             =   630
         Width           =   1380
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Planilla de Reparación"
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
         Height          =   435
         Left            =   210
         TabIndex        =   16
         Top             =   210
         Value           =   -1  'True
         Width           =   1380
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E2D3C0&
      Caption         =   "PREPARACION"
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
      Left            =   105
      TabIndex        =   11
      Top             =   105
      Width           =   3480
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   240
         ScaleHeight     =   615
         ScaleWidth      =   2940
         TabIndex        =   12
         Top             =   360
         Width           =   2970
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
            Left            =   0
            Picture         =   "MREPARA09.frx":1BA6
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Alta de Reparación"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command20 
            BackColor       =   &H8000000A&
            Height          =   630
            Left            =   2380
            Picture         =   "MREPARA09.frx":1EB0
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Búsqueda Cliente por CUIT, Domicilio, Nombre de Fantasía, etc"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H8000000A&
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
            Left            =   1800
            Picture         =   "MREPARA09.frx":3216
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Planilla de Presupuestaciones"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Planilla de Presupuesto"
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
            Left            =   650
            TabIndex        =   14
            Top             =   120
            Width           =   1590
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      Caption         =   "PRESUPUESTOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   120
      TabIndex        =   7
      Top             =   1330
      Width           =   3480
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   525
         ScaleHeight     =   615
         ScaleWidth      =   2400
         TabIndex        =   8
         Top             =   360
         Width           =   2430
         Begin VB.CommandButton Command3 
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
            Picture         =   "MREPARA09.frx":3520
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Generar Presupuesto por una Reparación"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Generar y Emitir un Presupuesto"
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
            TabIndex        =   10
            Top             =   105
            Width           =   1380
         End
      End
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7680
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnuTransacciones 
      Caption         =   "Transacciones"
      Begin VB.Menu AltaRepa 
         Caption         =   "Alta de Reparación"
      End
      Begin VB.Menu Planipre 
         Caption         =   "Planilla de Presupuestación"
         Begin VB.Menu Pladir 
            Caption         =   "Directa"
         End
         Begin VB.Menu Plabuscli 
            Caption         =   "Búsqueda de Cliente"
         End
      End
      Begin VB.Menu plh3 
         Caption         =   "-"
      End
      Begin VB.Menu Genpresu 
         Caption         =   "Generar Presupuesto"
      End
      Begin VB.Menu Ordrep 
         Caption         =   "Orden de Reparación"
      End
      Begin VB.Menu Cierrorep 
         Caption         =   "Cierre de O/Reparación"
      End
      Begin VB.Menu mnuRevCierrreRepa 
         Caption         =   "Revertir Cierre de Reparación"
      End
      Begin VB.Menu mnuRecuperaFechaCierreReparacion 
         Caption         =   "Recuperacion Fechas de Cierre de Reparacion"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuControlFinal 
         Caption         =   "Control Final"
      End
      Begin VB.Menu plh2 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Terminar"
      End
   End
   Begin VB.Menu mnuConsultas 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu mnuConsRepa 
         Caption         =   "Consulta Reparaciones"
      End
      Begin VB.Menu Anota 
         Caption         =   "Anotador de Seguimiento"
      End
      Begin VB.Menu mnuReparTecExt 
         Caption         =   "Reparaciones de Técnicos Externos"
      End
      Begin VB.Menu linea2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAgrpDescTarea 
         Caption         =   "Agrupado por Descripción de Tareas"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuReportesVarios 
         Caption         =   "Reportes Varios"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Ver"
      Begin VB.Menu Plarepa 
         Caption         =   "Planilla de Reparación"
      End
      Begin VB.Menu Presuint 
         Caption         =   "Presupuesto Interno"
      End
      Begin VB.Menu Presucli 
         Caption         =   "Presupuesto al Cliente"
      End
      Begin VB.Menu Orderep 
         Caption         =   "Orden de Reparación"
      End
      Begin VB.Menu RtoRep 
         Caption         =   "Entregas de Reparación"
      End
      Begin VB.Menu mnuRemitosEntrega 
         Caption         =   "Remitos de Entrega de Repuestos"
      End
   End
   Begin VB.Menu mnuUtilidades 
      Caption         =   "Utilidades"
      Begin VB.Menu mnuAgRemi 
         Caption         =   "Agrupamiento de Remitos"
      End
      Begin VB.Menu mnuApli 
         Caption         =   "Aplicación de Reparaciones a Pagos"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuConfiguración 
      Caption         =   "Configuración"
      Begin VB.Menu mnuVincularReportes 
         Caption         =   "Vincular Reportes al Formulario"
      End
      Begin VB.Menu mnuTablasAuxiliares 
         Caption         =   "Tablas Auxiliares"
         Visible         =   0   'False
      End
      Begin VB.Menu linea3 
         Caption         =   "-"
      End
      Begin VB.Menu Setup 
         Caption         =   "Configuración General"
      End
   End
   Begin VB.Menu mnuAccesosDirectos 
      Caption         =   "Accesos Directos"
      Begin VB.Menu Master 
         Caption         =   "Maestro de Artículos"
      End
      Begin VB.Menu Clientes 
         Caption         =   "Padrón de Clientes"
      End
   End
End
Attribute VB_Name = "MREPARA09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ModoAutomaticoPorImporteMinimo%
Public ModoAutomaticoPorImporteMinimo_Reparacion&
Sub cambiarEstado5a4()
   'CAMBIA LOS QUE ESTAN EN ESTADO 5 A ESTADO 4 SEGUN EMANUEL NO DEBERIA ESTAR  ESE ESTADO
   'EL 5 SE ASIGNA CUANDO SE REMITE PERO ACA EN REPARACION Y HAY UNA RUTINA EN EL VENGES07 QUE SI ESTA EN GARANTIA LO CAMBIA  A 5
   'ENTONCES CON ESTA RUTINA VAN A QUEDAR COMO MAXIMO EN ESTADO 4
   CONDI$ = "STATUS = 5"
   Call ACTUREGISTRO("REPARA", "STATUS", CONDI$, 4, REGAF&, "NOMESS")
End Sub

Sub GENERAORDENREPARACION_EYM(NroRep&)
    'SOLO SE CAMBIA EL ESTADO
    'NO IMPRIMEN, TAMPOCO GENERA RESERVAS POR QUE TODAVIA NO ESTAN INFORMADOS LOS MATERIALES
    'SE INFORMAN EN EL CIERRE DI
    
    Call ACTUREGISTRO("REPARA", "STATUS", "NUMREPA = " & NroRep&, 3, REGAF&, "NOMESS")
    Call ACTUREGISTRO("REPARA", "FechOrdRep", "NUMREPA = " & NroRep&, FETEXCOR1$(HOY(HO$)), REGAF&, "NOMESS")
    '
    CONDI$ = "NUMREPA = " & NroRep&
    NUMESER$ = VALOBADA("REPARA", "NUMEROSERIE", CONDI$, "NOMESS")
    RACLI$ = VALOBADA("REPARA", "Raso_Cli", CONDI$, "NOMESS")
    CUERPO$ = "Generación de Orden de Reparación" & vbCrLf
    CUERPO$ = CUERPO$ + "Número de Reparación: " & NroRep& & vbCrLf
    CUERPO$ = CUERPO$ + "Cliente: " & RACLI$
    Call ANOTANUSE(NUMESER$, CUERPO$, , "NOMESS")
    '
    Call COMUNI("Orden de Reparación Generada")

End Sub

Function LISTAREPARACION&(CONDI$)

    Screen.MousePointer = 11
    ' CARGA LISTA DE TECNICOS A UN VECTOR
    Dim ListaDeTecnicos() As String
    Dim RsTecnicos As ADODB.Recordset
    ConsultaSql$ = "SELECT * FROM DATASQL WITH(NOLOCK) WHERE LISTA='TECNIEXT' AND CAMPO1 <> ''"
    Set RsTecnicos = READSET(ConsultaSql$)
    
    ReDim ListaDeTecnicos(1)
    With RsTecnicos
        Do While Not .EOF
            IdTecnico% = XVALO(!CAMPO1)
            NombreTecnico$ = SAFETEXT$(!CAMPO2)
             
            If IdTecnico% > UBound(ListaDeTecnicos) Then ReDim Preserve ListaDeTecnicos(IdTecnico%)
            ListaDeTecnicos(IdTecnico%) = NombreTecnico$
            
            .MoveNext
        Loop
    End With
    On Error Resume Next
    RsTecnicos.Close
    Set RsTecnicos = Nothing
    On Error GoTo 0
    ' FIN DE LA CARGA LISTA DE TECNICOS
    
    TituloLista$ = "LISTA DE REPARACIONES"
    FRMZELECHO.Caption = TituloLista$

    Campos$ = "Reparación/I12;N-Serie/A14;Referencia/A48;N.Tec/F6;Nombre Técnico/A32"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$
    
    SQLX$ = "SELECT R.NUMREPA,R.REFEREN,R.TEC_EXT, R.NUMEROSERIE FROM REPARA R  WITH(NOLOCK) WHERE " & CONDI$
    SQLX$ = SQLX$ + " ORDER BY NUMREPA"
    Set rs = READSET(SQLX$)
    With rs
        If Not (.EOF And .BOF) Then
            FIN& = rs.RecordCount
            FRMZELECHO.MSF2.Rows = FIN& + 1 ' ESTABLECE DE ANTEMANO LA CANTIDAD DE FILAS A UTILIZAR
        End If
        Do While Not .EOF
          NR = XVALO(!numrepa)
          RF = SAFETEXT$(!Referen)
          TEC& = XVALO(!TEC_EXT)
          NOM$ = ListaDeTecnicos(TEC&)
          NUMEROSERIE$ = SAFETEXT$(!NUMEROSERIE)
          
          J& = J& + 1
          Call TRACE(24, J&, FIN&)
          FRMZELECHO.MSF2.TextMatrix(J&, 1) = NR
          FRMZELECHO.MSF2.TextMatrix(J&, 2) = NUMEROSERIE$
          FRMZELECHO.MSF2.TextMatrix(J&, 3) = RF
          FRMZELECHO.MSF2.TextMatrix(J&, 4) = TEC&
          FRMZELECHO.MSF2.TextMatrix(J&, 5) = NOM$
          If TEC& > 0 Then Call COLOREAR_GRILLA_SOLO_UNA_AFILA(FRMZELECHO.MSF2, &HDAE4EB, J&)
          .MoveNext
        Loop
    End With
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    '
    ReDim ListaDeTecnicos(0)
    On Error GoTo 0
    
    Screen.MousePointer = 1
    
    FRMZELECHO.Width = 12000
50  FRMZELECHO.Show 1

    LISTAREPARACION& = XVALO(RESP_ZELE_VECT(1))

End Function

Function NUPROA1$()
   '
   NUPROA1$ = ""
   '
   FEASX = COMEJ%
   FECHAS$ = FECHATEX$(FEASX)
   FECHAU% = FEASX
   FEASYYYY = FETEXCOR1$(FECHAU%)
   NUIDAS$ = Mid$(FECHAS$, 7, 2) + Mid$(FECHAS$, 4, 2) + Mid$(FECHAS$, 1, 2)
   If Left$(NUIDAS$, 1) = "0" Then NUIDAS$ = Mid$(NUIDAS$, 2)
   ORDAS% = 0
   '
   ' SUCUI% = UNINEGO%
   Dim MIRS As ADODB.Recordset
   Set MIRS = New ADODB.Recordset
   strSQL = "SELECT TOP 1 OrFeAsien FROM EncaBEAsien WHERE FECHASIEN = '" & FEASYYYY & "'"
   strSQL = strSQL & " ORDER BY ORFEASIEN DESC"
   MIRS.Open FILTSQL$(strSQL), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   On Error Resume Next
   MIRS.MoveFirst
   If Err Then GoTo 20
   ORDAS% = MIRS!OrFeAsien
   MIRS.Close
   Set MIRS = Nothing
   '
20 On Error GoTo 0
      
   ORDAS% = ORDAS% + 1
   ORDA$ = TRIM$(Str$(ORDAS%))
   While Len(ORDA$) < 3: ORDA$ = "0" + ORDA$: Wend
   NUPROA1$ = NUIDAS$ + "-" + ORDA$
   '
End Function





Sub VALIDARLOTES_Y_STOCK(NUREP&, NUMESER$, OKVL%)
    OKVL% = -1
    NUMEREPA$ = "OR-" & TRIM$(Str$(NUREP&))
    Dim MALCONSU%(32767): Dim DESCRIMAT$(32767)
    DepDest% = XVALO(CONTROL$("", "DEPREDES"))

    SQLX$ = "SELECT * FROM RESERVA  WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE IDSUBOR = '" & NUMEREPA$ & "'"
    SQLX$ = SQLX$ + " ORDER BY NUORITEM "
    Dim REPURTMP As ADODB.Recordset
    Set REPURTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With REPURTMP
      Do While Not .EOF
       CI1% = XVALO(!COD_INTE)
       CANTI = XVALO(!CANTRES)
       LOTE$ = SAFETEXT$(!idenlote)
       CODI$ = SAFETEXT$(!COD_EXTE)
       DESC$ = SAFETEXT$(!Descrip)
       If CI1% > 0 Then
          ' CANTIDAD RESERVADA
          CONDI$ = "IDSUBOR= '" & NUMEREPA$ & "' AND COD_INTE = " & CI1% & ""
          CANTIRESER# = XVALO(VALOBADA("RESERVA", "SUM(CANTRES)", CONDI$, "NOMESS"))
          ' CANTIDAD INGRESADA EN MOVISTO
          CONDI$ = "Cod_Inte = " & CI1% & " AND DOSECCOR = '" & TRIM$(NUMEREPA$) & "' And DepoMovi = " & DepDest%
          CANTING = XVALO(VALOBADA("MOVISTO", "SUM(CantIngre)", CONDI$, "NOMESS"))
          ' CANTIDAD SALIDA MOVISTO
          CONDI$ = "CodiMovi = 'TR' AND Cod_Inte = " & CI1% & " AND DOPRICOR = '" & TRIM$(NUMEREPA$) & "' AND DepoMovi = " & DepDest%
          CANTISL = XVALO(VALOBADA("MOVISTO", "SUM(CantSalid)", CONDI$, "NOMESS"))
          CANTOTA# = XVALO(CANTING) - XVALO(CANTISL)
          '
          If CANTIRESER# - CANTOTA# > 0.05 Then
             Call COMUNI("Material Insuficiente para Código: " & CODI$)
             MALCONSU%(CI1%) = 1  'CARGO VECTOR PARA GUARDAR INFORMACION DE LOS ITEMS Q´NO FUERON TRANSFERIDOS
             DESCRIMAT$(CI1%) = TRIM$(Mid$(TTXX$, 17, 40))
             'command1.Enabled = True
'             Exit Sub SOLO AVISA Y PERMITE CONTINUAR
          End If
          '
          If ESTRAZABLE%(CI1%) = 1 Then
              If TRIM$(Mid$(TTXX$, 72, 12)) = "" Then
                 Call COMUNI("Falta Lote Para Código: " & CODI$)
                 If MALCONSU%(CI1%) < 1 Then MALCONSU%(CI1%) = 2: DESCRIMAT$(CI1%) = TRIM$(Mid$(TTXX$, 17, 40)) 'CARGO VECTOR PARA GUARDAR INFORMACION DE LOS ITEMS Q´NO FUERON TRANSFERIDOS
                  'command1.Enabled = True
                  'Exit Sub SOLO AVISA Y PERMITE CONTINUAR
              End If
          End If
        End If
       .MoveNext
      Loop
    End With
    REPURTMP.Close
    Set REPURTMP = Nothing
    OKVL% = 1 ' SE LLEGO HASTA AQUI DA POR BUENA LA VALIDACION DEL STOCK
    
   'SE CONSIDERA QUE SI ES CON TEC.EXT.
    TECEXT& = TECNICO_REPA&(NUREP&)

    PUSOTITULO% = 0: CUERPO3$ = ""
    For L% = 0 To UBound(MALCONSU%)
       '
       If MALCONSU%(L%) = 1 Then
          If PUSOTITULO% = 0 Then
             CUERPO3$ = CUERPO3$ + "Consumo Reparación: " & NUREP& & vbCrLf
             CUERPO3$ = CUERPO3$ + "Técnico Externo: " & TECEXT& & vbCrLf
             CUERPO3$ = CUERPO3$ + "Items No Consumidos de Manera Normalizada" + vbCrLf
             PUSOTITULO% = 1
          End If
          CUERPO3$ = CUERPO3$ + DESCRIMAT$(L%) + vbCrLf
          CUERPO3$ = CUERPO3$ + "Material Insuficiente " + vbCrLf & String$(56, "-") & vbCrLf
          
       ElseIf MALCONSU%(L%) = 2 Then
          If PUSOTITULO% = 0 Then
             CUERPO3$ = CUERPO3$ + "Consumo Reparación: " & NUREP& & vbCrLf
             CUERPO3$ = CUERPO3$ + "Items No Consumidos de Manera Normalizada" & vbCrLf
             CUERPO3$ = CUERPO3$ + "Técnico Externo: " & TECEXT& + vbCrLf
             PUSOTITULO% = 1
          End If
          CUERPO3$ = CUERPO3$ + DESCRIMAT$(L%) + vbCrLf
          CUERPO3$ = CUERPO3$ + "Falta Lote" + vbCrLf & String$(56, "-") & vbCrLf
       End If
       If L% >= 32767 Then Exit For
    Next L%
    
    Call ANOTANUSE(NUMESER$, CAMPO3$)
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

Private Sub AltaRepa_Click()
   Call Command9_Click
End Sub

Private Sub Anota_Click()
   
   Call ABREORFAB
   '
   Call COPYSTRU("INDIPRES", "INDRECFI")
   
10 Call CARGALISEL("INDRECFI", "REPARA", "NUMREPA/F6.0;REFEREN/A56", "NUMREPA > 0", "DISTINCT")
   Call SELECHO("INDRECFI/Indice General de Reparaciónes ")
   NroRep& = XVALO(TRIM$(VALACT1("INDRECFI")))
   '
   If NroRep& < 1 Then Exit Sub
   '
   NUSERIE$ = VALOBADA("REPARA", "NUMEROSERIE", "NUMREPA = '" & NroRep& & "'", "NOMESS")
   
   CONUSE07.Label7 = NUSERIE$
   CONUSE07.Show 1
   GoTo 10
   
End Sub

Private Sub Cierrorep_Click()
   Call Command2_Click
End Sub

Private Sub Clientes_Click()
   Call CIERRARCH("*.*")
   ACONEC$ = "ARCHIG07/AMACLI"
   Call CONECRUN(ACONEC$, "Padron Maestro de Clientes")

'   Call PLAPRERE07.Command11_Click
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Sub VALEMATPORREP()
   '
   Dim ORFASEL$(8192)
   Dim CANECE(32767)
   '
10 NroRep& = LISTAREPARACION&("NUMREPA > 0 AND STATUS = 3")

'   Call ABREORFAB
'   Call COPYSTRU("INDIPRES", "INDERTRA")
'10 Call CARGALISEL("INDERTRA", "REPARA", "NUMREPA/F6.0;REFEREN/A56", "NUMREPA > 0 AND STATUS = 3", "DISTINCT")
'   Call SELECHO("INDERTRA/Indice General de Reparaciónes Generadas")
'   NROREP& = XVALO(TRIM$(VALACT1("INDERTRA")))
   If NroRep& < 1 Then Exit Sub
    '
   NUMEREPA$ = "OR-" & TRIM$(Str$(NroRep&))
    '
   For U& = 0 To 32767
      CANECE(U&) = 0
   Next U&
   '
   If CONTROL$("", "DEPRECEN") <> "" Then
     DEPOCENTRAL% = CONTROL$("", "DEPRECEN") ' TRAE SEGUN CONFIGURACION
   End If
   If CONTROL$("", "DEPREDES") <> "" Then
     DEPODESTINO% = CONTROL$("", "DEPREDES")
   End If
   '
   'SE CONSIDERA QUE SI ES CON TEC.EXT. CAMBIA AL DEPOSITO CONFIGURADO EN EL SETUP DE VENTAS CON DEPOSITO PAÑOL DE FACTURACION.
   If TECNICO_REPA&(NroRep&) > 0 Then 'dentro del punto 2 se transfiere al 44 en vez del 40
      DEPODESTINO% = XVALO(CONTROL$("", "DPREPPAÑ"))
   End If
   '
   VDEF$ = String$(64, 0)
   Call REPLA(VDEF$, NUMEREPA$, 1, 12)
   Call REPLA(VDEF$, Chr$(DEPODESTINO%), 14, 1)
   Call REPLA(VDEF$, Chr$(DEPOCENTRAL%), 13, 1)
   '
12 OBX$ = OBJPLA$("SEREPDEP", VDEF$)
   If OBX$ = "" Then Exit Sub
   '
   DEPODEST% = Asc(Mid$(OBX$, 14, 1))
   DepOrig% = Asc(Mid$(OBX$, 13, 1))
   '
   'VALIDACION QUE TENGA DERECHOS PARA CAMBIAR DEPOSITOS
   If DERACCE%("CAMBIDEP/NOMESS", "CAMBIO DE DEPOSITOS") < 2 Then
     If DEPOCENTRAL% <> DepOrig% Or DEPODESTINO% <> DEPODEST% Then
        Call MENSERR(24, "Usuario no Habilitado para Modificar Depósitos\Módulo: CAMBIDEP")
        GoTo 12
     End If
   End If
   '
   'VALIDA DEPOSITOS
   If DEPODEST% < 1 Or ECOARCH$("TADEPOSI", Chr$(DEPODEST%)) = "" Then 'VALIDACION
     Call MENSERR(24, "Depósito Destino No Válido")
      GoTo 12
   End If
   
   If DepOrig% < 1 Or ECOARCH$("TADEPOSI", Chr$(DepOrig%)) = "" Then
      Call MENSERR(24, "Depósito Origen No Válido")
      GoTo 12
   End If
   
   Call ABRESTOCKS
   'RECORRE LA TABLA RESERVAS Y ARMA EL ARCHIVO RESERVAS
   SQLX$ = " SELECT * FROM RESERVA "
   SQLX$ = SQLX$ + " WHERE CantRes-CantCons > 0 "
   SQLX$ = SQLX$ + " AND IDSUBOR = '" & NUMEREPA$ & "'"
    '
   Dim RESERTMP As ADODB.Recordset
   Set RESERTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))

   With RESERTMP
      On Error Resume Next
      .MoveFirst
      If Err Then
       On Error GoTo 0
       Call MENSERR(24, " No Existen Reservas efectuadas. ")
       GoTo 12
      End If
      On Error GoTo 0
      JJ& = 0
      CONTA% = 0
      FIN& = .RecordCount
      Call SETULTREG("!RESERVA", 0)
      Do While (.EOF = False)
        Y$ = REGBLAN("!RESERVA")
        '
        CII% = XVALO(!COD_INTE)
        If ESTRAZABLE(CII%) = 2 Then
           CONTA% = CONTA% + 1 ' valida si hay codigos trazables
           GoTo 13
        End If
            Call REPLA(Y$, MKI$(!COD_INTE), 1, 2)
            Call REPLA(Y$, !IDSUBOR, 3, 12)
            Call REPLA(Y$, MKI$(!COD_INTE), 15, 2)
            Call REPLA(Y$, MKI$(CVINT(!FechRes)), 17, 2)
            Call REPLA(Y$, MKD$(!CANTRES), 19, 8)
            Call REPLA(Y$, MKD$(!CANTCONS), 27, 8)
            Call REPLA(Y$, MKD$(!Saldo_Entre), 35, 8)
            Call REPLA(Y$, !idenlote, 43, 16)
            Call REPLA(Y$, Chr$(!DEPOSITO), 60, 1)
            Call REPLA(Y$, MKS$(!USOUNIT), 75, 4)
            '
            JJ& = JJ& + 1
            Call GRAREG("!RESERVA", Y$, JJ&)
        '
13    .MoveNext
      Loop
   End With
   Set RESERTMP = Nothing
   Call SETULTREG("!RESERVA", JJ&)
   Call CIERRARCH("!RESERVA")
   If CONTA% > 0 Then
      Call MENSERR(24, "Se Han Encontrado Códigos con Trazabilidad = 2\Los Cuales Seran Excluidos de la Lista de Componentes")
   End If
   '
'   Call ASIGNALOTE_RESERVA(NUMEREPA$, OK%)
'MsgBox OK%
   FIN& = ULTREG&("!RESERVA")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("!RESERVA", U&)
      NUMEREPA$ = Mid$(XX$, 3, 12)
      Ci% = CVI(Left$(XX$, 2))
      CANTI = CVD(Mid$(XX$, 19, 8)) '- CVD(Mid$(XX$, 27, 8))
      '
      If CANTI < 0 Then CANTI = 0
      '
      If TRIM$(UCase$(Unimed$(Ci%))) = "GK" Then
         CANTI = CANTI / 1000
      End If
      '
15    If Ci% > 0 Then
         If Ci% <= NUMAS% Then
            CANECE(Ci%) = CANECE(Ci%) + CANTI
         End If
      End If
      '
19 Next U&
   '
   Call SETULTREG("!VALEMAT", 0)
   FIN& = NUMAS%
   Screen.MousePointer = 11
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      Ci% = U&
      If CANECE(Ci%) > 0 Then
          II1$ = CODEXT$(Ci%)
          CONDI$ = " DoSecLar = '" & TRIM$(NUMEREPA$) & "' AND Cod_Inte = " & Ci% & " AND DepoMovi = " & DEPODEST% & ""
          CANTXX = XVALO(VALOBADA("MOVISTO", "SUM(CantIngre)", CONDI$, "NOMESS"))
          CANTSAL = XVALO(VALOBADA("MOVISTO", "SUM(CantSalid)", CONDI$, "NOMESS"))
          STODEST = CANTXX - CANTSAL 'STODEPOS(CI%, DEPODEST%)
          STOCENT = STODEPOS(Ci%, DepOrig%)
          NECNETA = CANECE(Ci%) - STODEST
          If NECNETA < 0 Then NECNETA = 0
          
          VALEMATX = NECNETA ' EL VALE SE PRESENTA POR LA NECESIDAD NETA
          If NECNETA > STOCENT Then ' SI NO HAY STOCK CONSUME HASTA EL STOCK
             VALEMATX = STOCENT
          End If
          '
          If VALEMATX < 0.05 Then VALEMATX = 0
           '
          ZZ$ = REGBLAN$("VALEMAT")
          Call REPLA(ZZ$, MKI$(Ci%), 1, 2)
          Call REPLA(ZZ$, MKS$(CANECE(Ci%)), 3, 4) ' NECESIDAD BRUTA
          Call REPLA(ZZ$, MKS$(STODEST), 7, 4)     ' STOCK DEPOSITO DESTINO
          Call REPLA(ZZ$, MKS$(NECNETA), 11, 4)    ' NECESIDAD NETA
          Call REPLA(ZZ$, MKS$(STOCENT), 15, 4)    ' STOCK DEP.CENTRAL
          Call REPLA(ZZ$, MKS$(VALEMATX), 19, 4)   ' VALE A EMITIR
          Call REPLA(ZZ$, MKS$(SECCI), 23, 4)
          Call REPLA(ZZ$, "......", 27, 6)
          Call REPLA(ZZ$, UBISTO$(Ci%), 33, 16)
 '         Call REPLA(ZZ$, IDLO$, 49, 16)
          Call REGAPP("!VALEMAT", ZZ$)
          'End If
      End If
   Next U&
   Call CIERRARCH("!VALEMAT")
   '
   Screen.MousePointer = 1
   JJ& = ULTREG&("!VALEMAT")
25 TITUPA$ = "Vale de TRANSFERENCIA" + " - de " + TRIM$(ECOARCH("TADEPOSI", MKI$(DepOrig%))) + " a " + TRIM$(ECOARCH$("TADEPOSI", MKI$(DEPODEST%)))
   VTB% = VENTABU%("VALEMAT" + "/POINDEL=NO/POSEDIT(1)=2/TITUPAN=" + TITUPA$)
   If VTB% < 0 Then GoTo 10
   '
   ' BLOQUE AGREGADO PARA CONTROL DE NEGATIVOS
   CONTA% = 0
   '
   For JJ1& = 1 To ULTREG&("!VALEMAT") ' valida si hay codigos trazables
      XX$ = REGLEIDO$("!VALEMAT", JJ1&)
      Ci% = CVI(Mid$(YY$, 1, 2))
      If ESTRAZABLE(Ci%) = 2 Then
        CONTA% = CONTA% + 1
      End If
   Next JJ1&
   '
   If CONTA% > 0 Then
      Call MENSERR(24, "Se Han Encontrado Códigos con Trazabilidad = 2\Los Cuales Deben ser Excluidos de la Lista de Componentes")
      GoTo 25
   End If
   '
   If CONTROL("VANEDECE", "DEPOCEN") = "SI" Then ' SI CONFIGURACION DE CONTROL DE NEGATIVOS
    For YX& = 1 To ULTREG&("!VALEMAT")
       YY$ = REGLEIDO$("!VALEMAT", YX&)
       Ci% = CVI(Mid$(YY$, 1, 2))
       CAN# = CVS(Mid$(YY$, 19, 4))
           If CAN# > 0.005 Then
             STOCXCENT# = STODEPOS(Ci%, DepOrig%) 'CVS(Mid$(YY$, 15, 4))
             VALMAT# = CVS(Mid$(YY$, 19, 4))
             If STOCXCENT# < VALMAT# Or STOCXCENT# < 0 Then
               Call MENSERR(24, "Transferencia No Procesable\Código: " & CODEXT$(Ci%) & "\Generará Un Stock Negativo en Depósito de Origen")
               GoTo 25
             End If
           End If
    Next YX&
   End If
   If PUEDEMOVI%(DepOrig%) < 0 Then GoTo 25 ' CANCELA SALIDA
   ' FIN CONTROL DE NEGATIVOS
   '
   If COMALTER%("Confirme la Transferencia a Realizar\\Cancelar\Confirmar Transferencia") <> 2 Then GoTo 10
   '
28 ' RECORRER LA TABLA Y CARGAR EL LIST DE FRM TRAZCAMPO
   TRAZCOMPO.List2.Clear
   For YX& = 1 To ULTREG&("!VALEMAT")
       YY$ = REGLEIDO$("!VALEMAT", YX&)
       Ci% = CVI(Mid$(YY$, 1, 2))
       If ESTRAZABLE%(Ci%) = 1 Then
         CAN# = CVS(Mid$(YY$, 19, 4))
         If CAN# > 0.005 Then
            Call REPLA(TX$, CODEXT$(Ci%), 1, 16)
            Call REPLA(TX$, DESCRIT$(Ci%), 17, 32)
            Call REPLA(TX$, FORMATNUM(CAN#, "F10.1"), 50, 11)
            TRAZCOMPO.List2.AddItem TX$
         End If
       End If
   Next YX&
   
   'Presentar formulario para asignar lotes
   If TRAZCOMPO.List2.ListCount > 0 Then
     TRAZCOMPO.Label7 = NUMEREPA$
     TRAZCOMPO.Show 1
     Call FRESHALL
     If TRAZCOMPO.List2.ListCount < 1 Then
          TRAZCOMPO.LIMPIAR
          GoTo 25
     End If
   End If
   '
   For JJI% = 1 To TRAZCOMPO.List2.ListCount
      XXI$ = TRAZCOMPO.List2.List(JJI% - 1)
      CODE$ = TRIM$(Left$(XXI$, 16))
      For JJJ% = JJI% + 1 To TRAZCOMPO.List2.ListCount
         XXJ$ = TRAZCOMPO.List2.List(JJJ% - 1)
         If TRIM$(Left$(XXJ$, 16)) = CODE$ Then
            Call MENSERR(24, "En Reparaciones no Válido Combinar Lotes")
            GoTo 28
         End If
      Next JJJ%
   Next JJI%
   '
   For U& = 1 To ULTREG&("!VALEMAT")
      XX$ = REGLEIDO$("!VALEMAT", U&)
      CAN = CVS(Mid$(XX$, 19, 4))
      LOTEX$ = ""
      If CAN >= 0.05 Then
         CI1% = CVI(Mid$(XX$, 1, 2))
         If ESTRAZABLE%(CI1%) > 0 Then
            CODE$ = TRIM$(CODEXT$(CI1%))
            For JJI% = 1 To TRAZCOMPO.List2.ListCount
              XYX$ = TRAZCOMPO.List2.List(JJI% - 1)
              If TRIM$(Left$(XYX$, 16)) = CODE$ Then
                 LOTEX$ = TRIM$(Mid$(XYX$, 62, 16))
                 GoTo 17
              End If
            Next JJI%
         End If
      End If
17    Call REPLA(XX$, LOTEX$, 49, 16)
      Call GRAREG("!VALEMAT", XX$, U&)
   Next U&
   '
   TRAZCOMPO.List2.Clear
   ' GRABAR TRANSFERENCIA
   '
   CMOV$ = TRIM$(CONTROL$("", "CODVALMA"))
   If CMOV$ = "" Then CMOV$ = "TR"
   '
   NRO& = ProNroComprobante(CMOV$)
   '
   Call BLOQARCH("MOVISTO")
   '
   Screen.MousePointer = 11
   NC% = 0
   VUNI# = 0
   MONE$ = ""
   'LOTE$ = ""
   FECHA% = HOY(HOS$)
   DOPRI$ = "TR." + TRIM$(Str$(NRO&))
   'DOSEC$ = "TR." + Mid$(HO$, 8, 1) + Mid$(HO$, 4, 2) + Left$(HO$, 2)
   DOSEC$ = TRIM$(NUMEREPA$) ' NUMERO DE ORDEN DE REPARACION
   While Len(DOSEC$) < 10
     DOSEC$ = DOSEC$ + " "
   Wend
                              '
   NORI% = 0
   For U& = 1 To ULTREG&("!VALEMAT")
      XX$ = REGLEIDO$("!VALEMAT", U&)
      If CVS(Mid$(XX$, 19, 4)) >= 0.05 Then
         '
         CI1% = CVI(Mid$(XX$, 1, 2))
         CAN = CVS(Mid$(XX$, 19, 4))
         REFE$ = "TR.a " + TRIM$(ECOARCH$("TADEPOSI", MKI$(DepOrig%)))
         '
         CANTI = (-1) * CAN
         NORFI% = NORI% + 1
         '
         LOTE$ = TRIM$(Mid$(XX$, 49, 16)) ' INICIALIZAMOS POR SI NO ENTRA NO QUEDE ANTERIOR
         '
         Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NORI%, VUNI#, 0, NC%, DepOrig%, CANTI)
         '
         REFE$ = "TR.de " + TRIM$(ECOARCH$("TADEPOSI", MKI$(DEPODEST%)))
         '
         CANTI = CAN
         CODI = CODEXT$(CI1%)
         Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NORI%, VUNI#, 0, NC%, DEPODEST%, CANTI)
         '
         Call ACTUREGISTRO("RESERVA", "IDENLOTE", "COD_INTE = " & CI1% & " AND IDSUBOR = '" & NUMEREPA$ & "'", LOTE$, RAFE&, "NOMESS")
      End If
   Next U&
   '
   Call BAJALDISCO
   Call LIBARCH("MOVISTO")
   '
'==============================================================================

'==============================================================================
   FORIPRE$ = CONTROL$("", "FORVALMA")
   If FORIPRE$ <> "" Then
      '
      FECHDOC$ = HOS$
      NUMEDOC$ = TRIM$(Str$(NRO&))
      While Len(NUMEDOC$) < 6
         NUMEDOC$ = "0" + NUMEDOC$
      Wend
      TIPODOC$ = "Vale de Materiales"
      '
      CODPARAM$(1) = "DEPOSITO": DOCPARAM$(1) = TRIM$(Str$(DepOrig%)) + " - " + TRIM$(ECOARCH$("TADEPOSI", MKI$(DepOrig%)))
      CODPARAM$(2) = "DESTINO": DOCPARAM$(2) = TRIM$(Str$(DEPODEST%)) + " - " + TRIM$(ECOARCH$("TADEPOSI", MKI$(DEPODEST%)))
      CODPARAM$(3) = "DOCUORIG": DOCPARAM$(3) = TRIM$(NUMEREPA$)
      CAPARDOC% = 3
      '
      CODTABU1$(1) = "COD-MAT"
      CODTABU1$(2) = "DES-MAT"
      CODTABU1$(3) = "UNIMED"
      CODTABU1$(4) = "CANTIDAD"
      CODTABU1$(5) = "CAKILOS"
      CODTABU1$(6) = "REF-MAT1"
      CODTABU1$(7) = "NRO-LOTE"
      CACOLTAB1% = 7
      '
      DESTIDOCNUE = 0
      CAITAB1% = 0
      CAITCRU% = 0
      '
      For U& = 1 To ULTREG&("!VALEMAT")
         XX$ = REGLEIDO$("!VALEMAT", U&)
         If CVS(Mid$(XX$, 19, 4)) >= 0.05 Then
            '
            CICRU% = CVI(Mid$(XX$, 1, 2))
            CACOL1$ = Str$(CVS(Mid$(XX$, 19, 4)))
            KICOL1$ = Str$(CVS(Mid$(XX$, 19, 4)) * PESUNI(CICRU%))
            LOTE$ = TRIM$(Mid$(XX$, 49, 16))
            '
            CAITAB1% = CAITAB1% + 1
            TETABU1$(1, CAITAB1%) = CODEXT$(CICRU%)
            TETABU1$(2, CAITAB1%) = DESCRIT0$(CICRU%)
            TETABU1$(3, CAITAB1%) = Unimed$(CICRU%)
            TETABU1$(4, CAITAB1%) = CACOL1$
            TETABU1$(5, CAITAB1%) = KICOL1$
            TETABU1$(6, CAITAB1%) = UBISTO$(CICRU%)
            TETABU1$(7, CAITAB1%) = LOTE$
            '
         End If
      Next U&
      '
      If CAITAB1% < 1 Then
         Screen.MousePointer = 1
         GoTo 10
      End If
      '
      Call PRINTDOC("FORVALMA")   ' Vale de Materiales
      Call BAJALDISCO
      '
   End If
   '
   Screen.MousePointer = 1
   GoTo 10
   '
End Sub
Private Sub Command11_Click()
    Command11.Enabled = False
    OPX% = ALTERNA%("Vale de Transferencia\Devolución por O/R")
    If OPX% < 1 Or OPX% > 2 Then GoTo 99
    If OPX% = 1 Then
      If DERACCE%("VALREPA", "Vale de Transferencia de Repuestos") = 2 Then Call VALEMATPORREP: GoTo 99
    End If
    If OPX% = 2 Then
      If DERACCE%("DEVOREP", "Devolucion de Repuestos") = 2 Then Call DEVOPORREP
    End If
99  Command11.Enabled = True
End Sub


Private Sub Command2_Click()
    '

    Command2.Enabled = False
    EPAC$ = TRIM$(EMPREAC$)
   '
    If DERACCE%("CIEREPA", "Cierre de Orden de Reparación") < 2 Then
      GoTo 99
    End If
    '
10  CONDI$ = "NUMREPA > 0 AND STATUS = 3"
    NroRep& = LISTAREPARACION&(CONDI$)

'    Call ABREORFAB
'    Call COPYSTRU("INDIPRES", "INDEREPR")
'10  CONDI$ = "NUMREPA > 0 AND STATUS = 3"
'    Call CARGALISEL("INDEREPR", "REPARA", "NUMREPA/F6.0;REFEREN/A56", CONDI$, "DISTINCT")
'    Call SELECHO("INDEREPR/Indice General Planillas de Reparación")
'    NROREP& = XVALO(TRIM$(VALACT1("INDEREPR")))
    If NroRep& < 1 Then GoTo 99
    '
    'VALIDAR QUE HAYA PASADO POR EL CONTROL FINAL
    '********************************************
    'TRAIGO FECHA DE RECEPCION Y NUMERO DE SERIE DE REPARA
    CONDI$ = "NUMREPA = " & NroRep&
    FECHARECEP% = CVINT(VALOBADA("REPARA", "fechrecep", CONDI$))
    FECHARECEPCION = FETEXCOR1$(FECHARECEP%)
    NUSERIE$ = VALOBADA("REPARA", "NUMEROSERIE", CONDI$)
    '
    '
    If InStr(UCase$(EPAC$), "EYM") > 0 Then
        CONDI$ = "FECH_ENSA >= '" & FECHARECEPCION & "' AND NUMEROSERIE = '" & NUSERIE$ & "'"
        If CantRegistros&("CONTFINEYM", CONDI$, "NOMESS") > 0 Then
          CIERREP09.Label25 = NroRep&
          CIERREP09.Show 1
          GoTo 10
        Else
          OPX% = COMALTER%("Reparación Número : " & NroRep& & " Falta Realizar  Control Final.\\Si, Continuar Sin Realizar Control\No Cancelar")
          If OPX% = 1 Then 'se quita que tenga derechos para cerrar sin control final
            CIERREP09.Label25 = NroRep&
            CIERREP09.Show 1
            GoTo 10
          Else
            GoTo 10
          End If
        End If
    End If
    
    'CONTROLO QUE HAYA UN CONTROL POSTERIOR A LA RECEPCION DE ESTE NUMERO DE SERIE
    CONDI$ = "FECH_ENSA >= '" & FECHARECEPCION & "' AND NUMEROSERIE = '" & NUSERIE$ & "'"
    If CantRegistros&("CONTROLFIN", CONDI$, "NOMESS") < 1 Then
       OPX% = COMALTER%("Reparación Número : " & NroRep& & " Falta Realizar  Control Final.\\Si, Continuar Sin Realizar Control\No Cancelar")
'       Call COMUNI("Reparación Numero : " & NROREP& & " Falta Realizar Control Final")
       If OPX% = 1 Then 'se quita que tenga derechos para cerrar sin control final
         'If DERACCE%("CIERSINC", "CIERRE SIN CONTROL FINAL") > 1 Then
            CIERREP09.Label25 = NroRep&
            CIERREP09.Show 1
            GoTo 10
         'End If
       End If
       GoTo 10
    Else
       CIERREP09.Label25 = NroRep&
       CIERREP09.Show 1
       GoTo 10
    End If
    '
99  Command2.Enabled = True
    '
End Sub

Private Sub Command20_Click()
     
    Command20.Enabled = False
    Call MENSERR(24, "No Disponible en la Presente Versión")
    'Call PLAPREPB
    Command20.Enabled = True
End Sub

Private Sub Command23_Click()
    Command23.Enabled = False
    '

    If DERACCE%("DESPARE", "Despacho de Reparación") < 2 Then
      GoTo 99
    End If
    '
15  CONDI$ = "TEC_EXT < 1 AND STATUS = 4" 'SE AGREGA QUE EL EQUIPO SI TIENE TECNICO NO SE PUEDA REALIZAR EL REMITO DADO QUE EL EQUIPO NO ESTA EN  MEDITEA
    NroRep& = LISTAREPARACION&(CONDI$)

'    Call ABREORFAB
'    '
'    Call COPYSTRU("INDIPRES", "INDORREM")
'    TEXNICO% = TECNICO_REPA&(NROREP&)
'    'CONDICION AGREGADA DEL PUNTO 2
'15  Call CARGALISEL("INDORREM", "REPARA", "NUMREPA/F6.0;REFEREN/A56", CONDI$, "DISTINCT")
'    Call SELECHO("INDORREM")
'    NROREP& = XVALO(TRIM$(VALACT1("INDORREM")))
    If NroRep& < 1 Then GoTo 99
    '
    REMI$ = VALOBADA("REPARA", "NUMREMITO", "NUMREPA = " & NroRep&, "NOMESS")
    '
    CODXT$ = ""
    Ci% = XVALO(VALOBADA("REPARA", "COD_INTE", "NUMREPA = " & NroRep&, "NOMESS"))
    If Ci% > 0 Then
       CODXT$ = CODEXT$(Ci%)
    End If
    '

    If TRIM$(REMI$) = "" Then
       PRONREM& = 1 + XVALO(VALOBADA("REPARA", "max(cast(substring(NUMREMITO, 9, 8)as int))", "NUMREPA > 0", "NOMESS"))
       REP$ = TRIM$(FORMATNUM$(PRONREM&, "F8.0"))
       While Len(REP$) < 8: REP$ = "0" + REP$: Wend
       REMI$ = "RP-0001-" & REP$
    End If
    '
    NCLIE = XVALO(VALOBADA("REPARA", "NUME_CLI", "NUMREPA = " & NroRep&, "NOMESS"))
    'SI ESTA EN GARANTIA Y ES INTERNA (QUE NO TIENE TECNICO) 'ESTO ES LA ESPECIFICACION DEL PUNTO 1
    If VALOBADA("REPARA", "Status_Gar", "NUMREPA = " & NroRep& & "", "NOMESS") > 0 Then
         TEXNICO% = TECNICO_REPA&(NroRep&)
         If TEXNICO% < 1 Then
            Call CIERREP09.GENERA_MOV_CBLE(NroRep&)
         End If
    End If

    FREP = HOY(HO$)
    FILTX$ = TRIM$(Str$(NroRep&)) & ";" & TRIM$(REMI$) & ";" & TRIM$(Str$(FREP)) & ";" & TRIM$(CODXT$)
    Call STDFORM("REMX-MDT", FILTX$)
    If STDSCREEN.APRO <> "OK" Then
       GoTo 15
    End If
    '
    'SE MODIFICA AGREGANDO OPCION POR QUE EN MEDITEA ABREN EL FRM DEL REMITO Y GRABAN
    'EL IMPORTE PERO AUN NO REMITEN, EN ESE CASO NO DEBE CAMBIAR EL ESTADO
    OPPX% = COMALTER%("Confirmación de Entrega\\Aún Nó Se Entrega\Dar por Remitido")
    
20  If OPPX% = 2 Then
         'REPUESTOS PARA REPARACIONES EXTERNAS DE EQUIPOS QUE NO ESTAN GARANTIA PUNTO 3
         TEXNICO% = TECNICO_REPA&(NroRep&)
         If TEXNICO% > 0 Then  'ES REPARACION EXTERNA
            If VALOBADA("REPARA", "Status_Gar", "NUMREPA = " & NroRep& & "", "NOMESS") < 1 Then 'NO ESTA EN GARANTIA
               Call GENERAPEDIDO(NroRep&, REMI$, RET_NPED&, OKX%)
               Call CIERREP09.CARGAELPEDIDO(RET_NPED&)
               If OKX% > 0 Then
                  Call COMUNI("Atención !!!!\Reparación Externa Sin Garantía\Se Generó la Orden de Venta N°: " & RET_NPED& & "\Se Deberán Facturar Los Repuestos a Entregar.")
               End If
               GoTo 99
            End If
         End If
        '
        
        'grabo fecha de reparacion y numero de reparacion en tanumse
        NUSERI$ = VALOBADA("REPARA", "NUMEROSERIE", "NUMREPA = " & NroRep&, "NOMESS")
        'NCLIE = XVALO(VALOBADA("REPARA", "NUME_CLI", "NUMREPA = " & NROREP&, "NOMESS")) LO OBTENGO MAS ARRIBA
        FE% = FREP
        FE1% = DIASPOST(FE%, 180)
        FEVEGREP = FETEXCOR1$(FE1%)
'        Call ACTUREGISTRO("REPARA", "STATUS", "NUMREPA = " & NROREP&, 5, REGAF&, "NOMESS")
'        Call ACTUREGISTRO("REPARA", "FechEnt", "NUMREPA = " & NROREP&, FETEXCOR1$(FE%), REGAF&, "NOMESS")
        SQLX$ = "SELECT * FROM REPARA "
        SQLX$ = SQLX$ + " WHERE  NUMREPA = " & NroRep&
        Dim RS2 As ADODB.Recordset
        Set RS2 = New ADODB.Recordset
27      On Error Resume Next
        RS2.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
           '
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                 On Error GoTo 0
                 Call CapturaErrorOpen
                 GoTo 27
        End If
        On Error GoTo 0
        
        With RS2
          Do While Not .EOF
            !Status = 5
            !FECHENT = CVDAT(FE%)
            .Update
            .MoveNext
          Loop
        End With
        RS2.Close
        Set RS2 = Nothing
        '
'        Call ACTUREGISTRO("TANUMSE", "FevenGaRep", "NUMEROSERIE = '" & NUSERI$ & "'", FEVEGREP, REGAF&, "NOMESS")
'        Call ACTUREGISTRO("TANUMSE", "UnumRep", "NUMEROSERIE = '" & NUSERI$ & "'", NROREP&, REGAF&, "NOMESS")
        '
        SQLX$ = "SELECT * FROM TANUMSE "
        SQLX$ = SQLX$ + " WHERE  NUMEROSERIE = '" & NUSERI$ & "'"
        Dim RS3 As ADODB.Recordset
        Set RS3 = New ADODB.Recordset
28      On Error Resume Next
        RS3.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        '
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                 On Error GoTo 0
                 Call CapturaErrorOpen
                 GoTo 28
        End If
        On Error GoTo 0
        
        With RS3
          Do While Not .EOF
            !FevenGaRep = CVDAT(FE1%)
            !UnumRep = NroRep&
            .Update
            .MoveNext
          Loop
        End With
        RS3.Close
        Set RS3 = Nothing
        '
        CUERPO$ = "Equipo Remitido" & vbCrLf
        CUERPO$ = CUERPO$ + "Número de Reparación: " & NroRep& & vbCrLf '
        CUERPO$ = CUERPO$ + "Número de Remito: " & REMI$ & vbCrLf
        CUERPO$ = CUERPO$ + "Cliente: " & rasocli$(NCLIE)
        Call ANOTANUSE(NUSERI$, CUERPO$, , "NOMESS")
    
    End If
    'BORRO LOS REGISTROS DEL REPUREPA DE ESTA ORDEN (ANTES SE BORRABA EN EL CIERRE, SE TRASLADO AQUI POR QUE LUEGO DEL CIERRE SE PUEDE MODIFICAR EL PRESUPESTO
    CONDI$ = "NUMREPA = " & NroRep& & ""
    Call BORRAREGISTROS("REPUREPA", CONDI$, REGAF&, "NOMESS")
    '

99  Command23.Enabled = True

End Sub

Private Sub Command3_Click()

    Command3.Enabled = False
    '
    EPAC$ = TRIM$(EMPREAC$)

    If DERACCE%("PRESURE", "Presupuesto de Reparación") < 2 Then
      GoTo 99
    End If
    '
    
50   NroRep& = LISTAREPARACION&(" NUMREPA > 0 AND STATUS < 5")

     If NroRep& < 1 Then GoTo 99
     If InStr(UCase$(EPAC$), "EYM") > 0 Then
        COTREPEYM.Label25 = NroRep&
        COTREPEYM.Show 1
        GoTo 50
     End If

    PRESURE09.Label25 = NroRep&
    PRESURE09.Label2 = ""
    CONDI$ = "Status_Gar > 0 AND NUMREPA = " & NroRep&
    If CantRegistros&("REPARA", CONDI$) > 0 Then
       PRESURE09.Label2 = "EN GARANTÍA"
    End If
    '
    PRESURE09.Show 1
    GoTo 50
    '
99  Command3.Enabled = True
    '
End Sub

Sub Command4_Click()
    Command4.Enabled = False
    
    '
    If DERACCE%("GENREPA", "Generación de Orden de Reparación") < 2 Then
      GoTo 99
    End If
    '
    EPAC$ = TRIM$(EMPREAC$)
    
10  If ModoAutomaticoPorImporteMinimo% = 1 Then
        NroRep& = ModoAutomaticoPorImporteMinimo_Reparacion&
        Call ACTUREGISTRO("REPARA", "OR_COMPRA", "NUMREPA=" & CStr(NroRep&), "AUTO", REGAF&)
        Call ACTUREGISTRO("REPARA", "FOR_COMPRA", "NUMREPA=" & CStr(NroRep&), FETEXCOR1$(HOY(HO$)), REGAF&)
        NUMEREPA$ = "OR-" & TRIM$(Str$(NroRep&))
        GoTo 50
    End If
    If InStr(EPAC$, "EYM") > 0 Then
      NroRep& = LISTAREPARACION&("(Status_Gar > 0 AND STATUS = 1) OR (STATUS = 2 AND STAT_APROBACION > 0)")
    Else
      NroRep& = LISTAREPARACION&("(Status_Gar > 0 AND STATUS = 1) OR (STATUS = 2)")
    End If
    
'    Call ABREORFAB
'
'    Call COPYSTRU("INDIPRES", "INDORPAR")
'    'CARGA INDICE CON LOS QUE ESTAN EN STATUS 2 O 3 (PRESUPUESTADOS O CON ORDEN DE REPARACION (PARA SU EDICION)) O EQUIPOS EN GARANTIA
'15  Call CARGALISEL("INDORPAR", "REPARA", "NUMREPA/F6.0;REFEREN/A56", "(Status_Gar > 0 AND STATUS = 1) OR (STATUS = 2)", "DISTINCT")
'    If ULTREG&("INDORPAR") < 1 Thens
'        Call MENSERR(24, "No Hay Ordenes de Reparacion por Emitir")
'        GoTo 99
'    End If
'    '
'    Call SELECHO("INDORPAR")
'    NROREP& = XVALO(TRIM$(VALACT1("INDORPAR")))
    If NroRep& < 1 Then GoTo 99
    
    If InStr(EPAC$, "EYM") > 0 Then
      If EstadoAprobacionReparaEym%(NroRep&) > 0 Then
         Call GENERAORDENREPARACION_EYM(NroRep&)
      Else
         Call COMUNI("Atención !! Imposible Generar Orden de Reparación\Cotización No Aprobada")
      End If
      Command4.Enabled = True
      Exit Sub
    End If

    NUMEREPA$ = "OR-" & TRIM$(Str$(NroRep&))
    '
    OrdenCerradaConItemsAgregado& = CantRegistros("REPUREPA", "NUMREPA=" & CStr(NroRep&) & " AND MARCAITEMAGREGADO=1")
    If OrdenCerradaConItemsAgregado& > 0 Then GoTo 50
    '
    ' CONTROL QUE NO HAYAN MATERIALES TRANSFERIDOS
    DepDest% = XVALO(CONTROL$("", "DEPREDES"))
    Set RESERMAT = FLEXCONN.Execute(FILTSQL$("SELECT * FROM RESERVA WHERE IDSUBOR='" & NUMEREPA$ & "'"))
    
    With RESERMAT
        Do While Not .EOF
            Ci% = !COD_INTE
            CANTIRESER# = XVALO(!CANTRES)
            If Ci% > 0 Then
              ' CANTIDAD INGRESADA EN MOVISTO
              CONDI$ = "Cod_Inte = " & Ci% & " AND DOSECCOR = '" & TRIM$(NUMEREPA$) & "' And DepoMovi = " & DepDest%
              CANTING = XVALO(VALOBADA("MOVISTO", "SUM(CantIngre)", CONDI$, "NOMESS"))
              ' CANTIDAD SALIDA MOVISTO
              CONDI$ = "CodiMovi = 'TR' AND Cod_Inte = " & Ci% & " AND DOSECCOR = '" & TRIM$(NUMEREPA$) & "' AND DepoMovi = " & DepDest%
              CANTISL = XVALO(VALOBADA("MOVISTO", "SUM(CantSalid)", CONDI$, "NOMESS"))
              CANTOTA# = XVALO(CANTING) - XVALO(CANTISL)
              '
              If CANTOTA# > 0.05 Then
                 Call COMUNI("Imposible Re-Generar Reparaciones\que Poseen Materiales Transferidos.\Realice una Devolucion y Reintentelo Nuevamente")
                 Command4.Enabled = True
                 Exit Sub
              End If
            End If
            .MoveNext
        Loop
    End With

20  FILTX$ = TRIM$(Str$(NroRep&)) ' & ";" & TRIM$(Str$(FREP))
    Call STDFORM("OREP-MDT", FILTX$)
    If STDSCREEN.APRO <> "OK" Then
       GoTo 10
    End If
    '
    FEORCOM% = XVALO(VALOFORM("Text-1"))
    ORCOMPRA$ = VALOFORM("Text-3")
    If TRIM$(ORCOMPRA$) = "" Or FEORCOM% < 1 Then
        Call COMUNI("No se han Ingresado Fecha y Nº de Orden de Compra")
        GoTo 20
    End If
    '
    'TRASLADO DE REPUREPA A RESERVA
50  Call BORRAREGISTROS("RESERVA", "IDSUBOR = '" & NUMEREPA$ & "'", REGAF&, "NOMESS")

    SQLX$ = "SELECT * FROM REPUREPA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE NUMREPA = " & NroRep&
    SQLX$ = SQLX$ + " ORDER BY NUORD_ITEM"
    '
    Dim REPUTMP As ADODB.Recordset
    Set REPUTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With REPUTMP
    Do While Not .EOF
         
         If XVALO(!MarcaItemAgregado) = 0 And OrdenCerradaConItemsAgregado& > 0 Then GoTo 60
         
         SQLX1$ = "SELECT * FROM RESERVA"
         SQLX1$ = SQLX1$ + " WHERE COD_INTE < 1 "
         Dim RESETMP As ADODB.Recordset
         Set RESETMP = New ADODB.Recordset
         RESETMP.Open FILTSQL$(SQLX1$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
         With RESETMP
           .AddNew   ' YA CONVERTIDO
           On Error Resume Next
           !CodiEmpr = CodiEmp%
           On Error GoTo 0
           Ci% = XVALO(REPUTMP!COD_INTE)
           CANTI = XVALO(REPUTMP!Cant_Sol)
           !COD_INTE = Ci%
           !COD_EXTE = CODEXT$(Ci%)
           !Descrip = Left$(DESCRIT$(Ci%), 64) 'modificado por que devolvia un error de datos
           !IDSUBOR = NUMEREPA$
           !FechRes = Now
           !CANTRES = CANTI
           !CANTCONS = 0
           !Saldo_Entre = CANTI
           !idenlote = ""
           !DEPOSITO = 0
           !USOUNIT = CANTI
           !NuOrItem = XVALO(REPUTMP!NuOrd_Item)
           .Update
         End With
         RESETMP.Close
         Set RESETMP = Nothing

60  .MoveNext
    Loop
    End With
    
    REPUTMP.Close
    Set REPUTMP = Nothing
    '
    Call ACTUREGISTRO("REPARA", "STATUS", "NUMREPA = " & NroRep&, 3, REGAF&, "NOMESS")
    '
    NUMESER$ = VALOFORM("PARAM-10")
    RACLI$ = VALOFORM("PARAM-4")
    CUERPO$ = "Generación de Orden de Reparación" & vbCrLf
    CUERPO$ = CUERPO$ + "Número de Reparación: " & NroRep& & vbCrLf
    CUERPO$ = CUERPO$ + "Cliente: " & RACLI$
    Call ANOTANUSE(NUMESER$, CUERPO$, , "NOMESS")
    '
    Call COMUNI("Orden de Reparación Generada")
    '
99  Command4.Enabled = True
End Sub

Private Sub Command43_Click()
    Call ABREORFAB
    '
    Call COPYSTRU("INDIPRES", "INDIRREM")
15  Call CARGALISEL("INDIRREM", "REPARA", "NUMREPA/F6.0;REFEREN/A56", "STATUS = 5", "DISTINCT")
    Call SELECHO("INDIRREM")
    NroRep& = XVALO(TRIM$(VALACT1("INDIRREM")))
    If NroRep& < 1 Then Exit Sub
    '
    CONDI$ = "NUMREPA = " & NroRep&
    Call ACTUREGISTRO("REPARA", "STATUS", CONDI$, 4, REGAF&, "NOMESS")
    '
    If REGAF& > 0 Then
      Call COMUNI("Transacciòn Realizada Exitosamente")
    Else
      Call COMUNI("No Se Realizò la Transacciòn Consulte a su Soporte de Sistema")
    End If
    '
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    '
    If DERACCE%("PLAPRERE", "Planilla de Presupuestación de Reparaciones") < 2 Then
      GoTo 99
    End If
    '
50  NroRep& = LISTAREPARACION&(" STATUS < 5")
    If NroRep& > 0 Then
      PLAPRERE09.Label25 = NroRep&
      PLAPRERE09.Show 1
      GoTo 50
    End If
    '
99  Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   '
   If DERACCE%("STPREPA", "Configuracion de Funcionamiento de Reparación") < 2 Then
     GoTo 99
   End If
   '
   CONFREPA09.Show 1 ' SETUP
   '
99 Command6.Enabled = True
End Sub

Private Sub Command7_Click()
    '
    Command7.Enabled = False
    If DERACCE%("CIEREPA", "Cierre de Orden de Reparación") < 2 Then
      GoTo 99
    End If
    '
10  NroRep& = LISTAREPARACION&("NUMREPA > 0 AND STATUS = 3")

'    Call ABREORFAB
'    Call COPYSTRU("INDIPRES", "INDEREPR")
'10  Call CARGALISEL("INDEREPR", "REPARA", "NUMREPA/F6.0;REFEREN/A56", "NUMREPA > 0 AND STATUS = 3", "DISTINCT")
'    Call SELECHO("INDEREPR/Indice General Planillas de Reparación")
'    NROREP& = XVALO(TRIM$(VALACT1("INDEREPR")))
    If NroRep& < 1 Then GoTo 99
    '
    DESCRIREP.Label25 = NroRep&
    NUSERIE$ = VALOBADA("REPARA", "NUMEROSERIE", "NUMREPA = '" & NroRep& & "'", "NOMESS")
    DESCRIREP.Label51 = NUSERIE$
    DESCRIREP.Show 1
    GoTo 10

99  Command7.Enabled = True
    '
End Sub

Sub MATEROREP()
    '
    Screen.MousePointer = 11
    Call COPYSTRU("PRESREPA", "OREPARPE")
    JJ& = 0
    For U& = 1 To ULTREG&("PRESREPA")
      XX$ = REGLEIDO$("PRESREPA", U&)
      If CVS(Mid$(XX$, 125, 4)) >= 0.5 Then
        JJ& = JJ& + 1
        Call GRAREG("OREPARPE", XX$, JJ&)
      End If
    Next U&
    Call SETULTREG("OREPARPE", JJ&)
    Call CIERRARCH("OREPARPE")
    Screen.MousePointer = 1
    '
    Call FINAL("*COREPUES")

99 End Sub
'
Sub OPEROREP()
   '
   Call COPYSTRU("INDIOTRA", "INDIOREP") ' INDICE DE TODAS LAS o/t'S
   '
   Screen.MousePointer = 11
   JJ& = 0
   REBLA$ = REGBLAN$("INDIOREP")
   FIN& = ULTREG&("PRESREPA")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("PRESREPA", U&)
     If CVS(Mid$(XX$, 125, 4)) >= 0.5 Then
       NOREPARA& = CVS(Left$(XX$, 4))
       NORFA$ = "OR-" + TRIM$(Str$(NOREPARA&))
       DORFA$ = TRIM$(Mid$(XX$, 5, 60))
       FEMI% = CVI(Mid$(XX$, 121, 2))
       '
       ZZ$ = REBLA$
       Call REPLA(ZZ$, NORFA$, 1, 12)
       Call REPLA(ZZ$, DORFA$, 13, 48)
       Call REPLA(ZZ$, MKI$(FEMI%), 61, 2)
       JJ& = JJ& + 1
       Call GRAREG("INDIOREP", ZZ$, JJ&)
     End If
   Next U&
   Call SETULTREG("INDIOREP", JJ&)
   Call CIERRARCH("INDIOREP")
   Screen.MousePointer = 1
   '
   Call FINAL("*COMOBRAD")
   '
End Sub
'
Sub PRESEDOC()
    '
    Call ABREORFAB
3   If Option1.Value = True Then ' VER
          ASEL$ = "PRESREPA"
          TSEL$ = "Indice General Planillas de Reparación"
          TIDOCX$ = TRIM$("Planilla de Reparación")
        ElseIf Option2.Value = True Then ' MODIFICADA
          '
          If DERACCE%("VEPREPAR", "Visualizar Presupuestos de Reparación") < 2 Then
            Exit Sub
          End If
          '
          Screen.MousePointer = 11
          Call COPYSTRU("PRESREPA", "PREPAMI")
          Call CARGALISEL("PREPAMI", "REPARA", "NUMREPA/F6.0;REFEREN/A56", "NUMREPA > 0 AND STATUS > 1", "DISTINCT")
          Call CIERRARCH("PREPAMI")
          Screen.MousePointer = 1
          ASEL$ = "PREPAMI"
          TSEL$ = "Indice General Presupuestos de Reparación"
          TIDOCX$ = TRIM$("Presupuesto de Reparación")
        ElseIf Option3.Value = True Then ' MODIFICADA
          '
          If DERACCE%("VEPREPAR", "Visualizar Presupuestos de Reparación") < 2 Then
            Exit Sub
          End If
          '
          Screen.MousePointer = 11
          Call COPYSTRU("PRESREPA", "PREPAEMI")
          Call CARGALISEL("PREPAMI", "REPARA", "NUMREPA/F6.0;REFEREN/A56", "NUMREPA > 0 AND STATUS > 1", "DISTINCT")
          '
          Call CIERRARCH("PREPAMI")
          Screen.MousePointer = 1
          ASEL$ = "PREPAMI"
          TSEL$ = "Indice General Presupuestos de Reparación"
          'TIPODOC$ = "Presup.Reparación (Externo)" MODIFICADO POR QUE CON EL PARENTESIS NO ENVIABA EL MAIL
          TIDOCX$ = TRIM$("Presup.Reparación Externo")
        ElseIf Option4.Value = True Then
'          Screen.MousePointer = 11
'          Call COPYSTRU("PRESREPA", "OREPARPE")
'          JJ& = 0
'          For U& = 1 To ULTREG&("PRESREPA")
'            XX$ = REGLEIDO$("PRESREPA", U&)
'            If CVS(Mid$(XX$, 125, 4)) >= 0.5 Then
'              JJ& = JJ& + 1
'              Call GRAREG("OREPARPE", XX$, JJ&)
'            End If
'          Next U&
'          Call SETULTREG("OREPARPE", JJ&)
'          Call CIERRARCH("OREPARPE")
'          If ULTREG&("OREPARPE") < 1 Then
'             Call MENSERR(24, "No Disponible en la Presente Versión")
'             Screen.MousePointer = 1
'             GoTo 99
'          End If
'          Screen.MousePointer = 1
'          ASEL$ = "OREPARPE"
'          TSEL$ = "Indice General Ordenes de Reparación"
'          TIDOCX$ = TRIM$("Orden de Reparación")

            Call CARGALISEL("INDERTRA", "REPARA", "NUMREPA/F6.0;REFEREN/A56", "NUMREPA > 0 AND STATUS >= 3", "DISTINCT")
103         Call SELECHO("INDERTRA/Indice General de Reparaciónes Generadas")
            NroRep& = XVALO(TRIM$(VALACT1("INDERTRA")))
            If NroRep& < 1 Then GoTo 99
            
            FILTX$ = TRIM$(Str$(NroRep&))
            Call STDFORM("OREP-MDT", FILTX$, "READONLY")
            GoTo 103
    End If
    '
4   Call FRESHECHO(ASEL$)
    If ULTREG&(ASEL$) < 1 Then
       Call MENSERR(24, "No Disponible en la Presente Versión")
       GoTo 99
    End If
    '
    Call SELECHO(ASEL$ + "/" + TSEL$)
    NPX& = Val(VALACT1$(ASEL$))
    If NPX& < 1 Then GoTo 99
    '
    Screen.MousePointer = 11
    NUDOCU$ = TRIM$(Str$(NPX&))
    TIDOCU$ = TIDOCX$ + " " + NUDOCU$ + " "
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCU$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 5
      End If
    Next UI&
    '
    Screen.MousePointer = 1
    Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a esta Reparación.")
    On Error Resume Next
    GoTo 99
    '
5   Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
    GoTo 3
    '
99  End Sub

Sub PLAPREPA()
   '
10 Call SELECHO("DEUDOR1")
   NC% = Val(VALACT1$("DEUDOR1"))
   If NC% > 0 Then
     If rasocli$(NC%) <> "" Then
       PLAPRERE07.Label11 = TRIM$(Str$(NC%))
       PLAPRERE07.Label18 = TRIM$(Str$(NUPRESUP&))
       PLAPRERE07.Show 1
       GoTo 10
     End If
   End If
   '
End Sub

Sub PLAPREPB()
   Static ARGU$
   Command20.Enabled = False
10 ARGU$ = UCase$(TRIM$(InputBox$("Buscar", "Búsqueda Cliente por Texto", ARGU$)))
   If ARGU$ <> "" Then
     Screen.MousePointer = 11
     JJ& = 0
     Call COPYSTRU("DEUDOR1", "SELECLI")
     '
     Call APERBASDAT("CLIEN")
     SQLX$ = "SELECT * FROM DEUDOR12"
     Dim DEUDDDOR12 As ADODB.Recordset
     Set DEUDDDOR12 = New ADODB.Recordset
     DEUDDDOR12.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
     On Error Resume Next
     With DEUDDDOR12
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
           For KKI% = 1 To .Fields.Count
             On Error Resume Next
             VATEX$ = ""
             VATEX$ = .Fields(KKI% - 1).Value
             On Error GoTo 0
             If InStr(VATEX$, ARGU$) > 0 Then
               NCLIE% = !Nume_Cli
               RACLI$ = !Raso_Cli
               XX$ = MKI$(NCLIE%) + RACLI$
               JJ& = JJ& + 1
               Call GRAREG("!SELECLI", XX$, JJ&)
               Exit For
             End If
           Next KKI%
         .MoveNext
         Loop
       End If
     End With
     Set DEUDDDOR12 = Nothing
     Call SETULTREG("!SELECLI", JJ&)
     '
     Screen.MousePointer = 1
     If JJ& < 1 Then
         Call MENSERR(24, "No se Encontraron Coincidencias")
         GoTo 10
       ElseIf JJ& = 1 Then
         PLAPRERE07.Label11 = TRIM$(Str$(CVI(Left$(REGLEIDO$("!SELECLI", 1), 2))))
         PLAPRERE07.Label18 = TRIM$(Str$(NUPRESUP&))
         PLAPRERE07.Show 1
         GoTo 10
       Else
         Call FRESHECHO("!SELECLI")
         Call SELECHO("!SELECLI")
         NCLIX$ = VALACT1$("!SELECLI")
         If NCLIX$ <> "" Then
           PLAPRERE07.Label11 = TRIM$(NCLIX$)
           PLAPRERE07.Label18 = TRIM$(Str$(NUPRESUP&))
           PLAPRERE07.Show 1
           GoTo 10
         End If
     End If
   End If
   '
End Sub

Sub PRESUREPA()
   '
   
'   If DERACCE%("PRESREPA", "Generar Presupuestos Reparaciones") < 2 Then
'     Exit Sub
'   End If
   '
'3  Call FRESHECHO("PRESREPA")
'   PRESURE09.Label18 = ""
'   Call SELECHO("PRESREPA/Indice General Planillas de Reparación")
'   NPX& = Val(VALACT1$("PRESREPA"))
'   If NPX& < 1 Then Exit Sub
'   '
'   For UI& = ULTREG&("PRESREPA") To 1 Step -1
'     XX$ = REGLEIDO$("PRESREPA", UI&)
'     If CVS(Left$(XX$, 4)) = NPX& Then GoTo 10
'   Next UI&
'   Call COMUNI("Error de Indices - Consulte.")
'   Exit Sub
'   '
'10 REPRESUP = TRIM$(Str$(UI&))
'   XX$ = REGLEIDO$("PRESREPA", Val(REPRESUP))
'   '
'   If CVS(Mid$(XX$, 125, 4)) > 0.5 Then
'     Call COMUNI("Imposible Modificar este Presupuesto.\Corresponde O.Venta " + TRIM$(Str$(CVS(Mid$(XX$, 125, 4)))))
'     Exit Sub
'   End If
'   '
'   PRESURE07.Label18 = TRIM$(Str$(NPX&))
'   PRESURE07.Show 1
'   GoTo 3
   '
End Sub

Sub OREPARA()
   '
   If DERACCE%("OREPARA", "Generar / Cerrar Orden de Reparación") < 2 Then
     Exit Sub
   End If
   '
3  Screen.MousePointer = 11
   Call COPYSTRU("PRESREPA", "PREPAEMI")
   JJ& = 0
   For U& = 1 To ULTREG&("PRESREPA")
     XX$ = REGLEIDO$("PRESREPA", U&)
     If CVI(Mid$(XX$, 121, 2)) > 0 Then
       If CVS(Mid$(XX$, 125, 4)) < 0.5 Then
         JJ& = JJ& + 1
         Call GRAREG("PREPAEMI", XX$, JJ&)
       End If
     End If
   Next U&
   Call SETULTREG("PREPAEMI", JJ&)
   Call CIERRARCH("PREPAEMI")
   Screen.MousePointer = 1
   '
   If JJ& < 1 Then
     Call COMUNI("No Hay Presupuestos Pendientes\para Emitir Orden de Reparación.")
     Exit Sub
   End If
   '
   Call FRESHECHO("PREPAEMI")
   Call SELECHO("PREPAEMI/Indice Presupuestos de Reparación Pendientes")
   NPX& = Val(VALACT1$("PREPAEMI"))
   If NPX& < 1 Then Exit Sub
   '
   For UI& = ULTREG&("PRESREPA") To 1 Step -1
     XX$ = REGLEIDO$("PRESREPA", UI&)
     If CVS(Left$(XX$, 4)) = NPX& Then GoTo 10
   Next UI&
   Call COMUNI("Error de Indices - Consulte.")
   Exit Sub
   '
10 REPRESUP = TRIM$(Str$(UI&))
   XX$ = REGLEIDO$("PRESREPA", Val(REPRESUP))
   If CVI(Mid$(XX$, 121, 2)) < 1 Then
      Call COMUNI("Falta Generar y Emitir Presupuesto.")
      PRESURE07.Label18 = TRIM$(Str$(NPX&))
      PRESURE07.Show 1
      GoTo 3
   End If
   '
   If CVS(Mid$(XX$, 125, 4)) > 0.5 Then
      Call COMUNI("Orden de Reparación ya Emitida.")
      GoTo 3
   End If
   '
   ORDEREPA07.Label18 = ""
   ORDEREPA07.Label18 = TRIM$(Str$(NPX&))
   ORDEREPA07.Show 1
   GoTo 3
   '
End Sub

Sub CIERROREPA()
   '
   If DERACCE%("OREPARA", "Generar / Cerrar Orden de Reparación") < 2 Then
     Exit Sub
   End If
   '
3  Screen.MousePointer = 11
   Call COPYSTRU("PRESREPA", "OREPARPE")
   JJ& = 0
   For U& = 1 To ULTREG&("PRESREPA")
     XX$ = REGLEIDO$("PRESREPA", U&)
     If CVS(Mid$(XX$, 125, 4)) >= 0.5 Then
       If CVI(Mid$(XX$, 123, 2)) < 1 Then
         JJ& = JJ& + 1
         Call GRAREG("OREPARPE", XX$, JJ&)
       End If
     End If
   Next U&
   Call SETULTREG("OREPARPE", JJ&)
   Call CIERRARCH("OREPARPE")
   Screen.MousePointer = 1
   '
   If JJ& < 1 Then
     Call COMUNI("No hay Ordenes de Reparación Pendientes\que Puedan Cerrarse.")
     Exit Sub
   End If
   '
   Call SELECHO("OREPARPE/Indice de Ordenes de Reparación en Curso")
   NPX& = Val(VALACT1$("OREPARPE"))
   If NPX& < 1 Then Exit Sub
   '
   For UI& = ULTREG&("PRESREPA") To 1 Step -1
     XX$ = REGLEIDO$("PRESREPA", UI&)
     If CVS(Left$(XX$, 4)) = NPX& Then GoTo 10
   Next UI&
   Call COMUNI("Error de Indices - Consulte.")
   Exit Sub
   '
10 REPRE& = Str$(UI&)
   XX$ = REGLEIDO$("PRESREPA", REPRE&)
   If CVI(Mid$(XX$, 123, 2)) > 0 Then
      Call COMUNI("Orden de Reparación ya Cerrada")
      GoTo 3
   End If
   '
   If CVS(Mid$(XX$, 125, 4)) < 0.5 Then
      Call COMUNI("Falta Generar Orden de Reparación")
      ORDEREPA.Label18 = TRIM$(Str$(NPX&))
      ORDEREPA.Show 1
      GoTo 3
   End If
   '
   NSERIE$ = Mid$(XX$, 71, 15)
   NCLIE% = CVI(Mid$(XX$, 67, 2))
   TALOX$ = TRIM$(Mid$(XX$, 95, 6))
   NROPE& = CVS(Mid$(XX$, 125, 4))
   VDEF$ = XX$ + String$(128, 0)
   HOII% = HOY(HOS$)
   Call REPLA(VDEF$, MKI$(HOII%), 123, 2)
   DepOrig% = Val(CONTROL$("", "DEPOCEN"))
   Call REPLA(VDEF$, Chr$(DepOrig%), 129, 1)
20 OBX$ = OBJPLA$("CIERROREPA", VDEF$)
   If OBX$ = "" Then
     Exit Sub
   End If
   '
   DP2% = Asc(Mid$(OBX$, 129, 1))
   If DP2% < 1 Or TRIM$(ECOARCH$("ECODEP", Chr$(DP2%))) = "" Then
      Call MENSERR(24, "Falta Depósito de Consumo")
      GoTo 20
   End If
   FECIE% = CVI(Mid$(OBX$, 123, 2))
   '
   Screen.MousePointer = 11
   Call SETULTREG("!MATEROF", 0)
   NPI% = NPX& Mod 10000
   RFF$ = RECFILT$("REPUREPA", 2, MKI$(NPI%))
   '
   URE& = ULTREG&("REPUREPA")
30 For KK& = 1 To Len(RFF$) Step 4
     REG& = CVS(Mid$(RFF$, KK&, 4))
     If REG& > 0 Then
       If REG& <= URE& Then
         XX$ = REGLEIDO$("REPUREPA", REG&)
         If CVS(Left$(XX$, 4)) = NPX& Then
           CII% = CVI(Mid$(XX$, 7, 2))
           CANTI = CVS(Mid$(XX$, 9, 4))
           '
           ZZ$ = REGBLAN$("MATEROF")
           Call REPLA(ZZ$, MKI$(CII%), 1, 2)
           Call REPLA(ZZ$, Unimed$(CII%), 3, 2)
           Call REPLA(ZZ$, MKS$(CANTI), 5, 4)
           Call REPLA(ZZ$, MKD$(CDbl(CANTI)), 9, 8)
           Call REPLA(ZZ$, MKS$(REG&), 17, 4)
           STDIS = STOCKDIS(CII%)
           Call REPLA(ZZ$, MKS$(STDIS), 21, 4)
           FALTA = CANTI - STDIS: If FALTA < 0 Then FALTA = 0
           Call REPLA(ZZ$, MKS$(FALTA), 25, 4)
           Call REGAPP("!MATEROF", ZZ$)
           '
         End If
       End If
     End If
   Next KK&
   Call CIERRARCH("!MATEROF")
   Screen.MousePointer = 1
   '
   TITUPA$ = "Consumo de Repuestos O.Rep." + TRIM$(Str$(NPX&))
   CIXI% = CVI(Mid$(OBX$, 69, 2))
   CODPRO$ = CODEXT$(CIXI%)
   TITUPA$ = TITUPA$ + " - " + TRIM$(CODPRO$) + " - " + TRIM$(DESCRIT$(CIXI%)) + " - Cant.=" + Str$(Int(CACONJU# + 0.5))
   If NCLIE% > 0 Then TITUPA$ = TITUPA$ + " - " + rasocli$(NCLIE%)
   '
   VTB% = VENTABU%("MATEROF/TITUPAN=" + TITUPA$)
   If VTB% < 0 Then
      Exit Sub
   End If
   '
   Call CIERRARCH("MOVISTO")
   Call BLOQARCH("MOVISTO")
   Call BLOQARCH("REPUREPA")
   '
   XX$ = REGLEIDO$("PRESREPA", REPRE&)
   Call REPLA(XX$, MKI$(FECIE%), 123, 2)
   Call GRAREG("PRESREPA", XX$, REPRE&)
   Call CIERRARCH("PRESREPA")
   '
   DOPRI$ = "OR-" + TRIM$(Str$(NPX&))
   DOSEC$ = "TL-" + TRIM$(TALOX$)
   REFE$ = "Reparación " + CODPRO$
   '
   For U& = 1 To ULTREG&("!MATEROF")
     ZZ$ = REGLEIDO$("!MATEROF", U&)
     '
     CI2% = CVI(Mid$(ZZ$, 1, 2))
     CAN = (-1) * CVS(Mid$(ZZ$, 5, 4))
     '
     USOUNI# = CAN
     RERES& = CVS(Mid$(ZZ$, 17, 4))
     If CI2% > 0 Then
       If CI2% <= NUMAS% Then
         If Abs(CAN) >= 0.05 Then
           LOTE$ = ""
           CMOV$ = "RP"
           '
           If Abs(CAN) >= 0.05 Then
              DEPX% = DP2%: If DP2% = 99 Then DEPX% = DEPOPRE%(CI2%)
              'Call MOVISTOK(FECIE%, CI2%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPX%, CAN)
              Call MOVISTOK(FECIE%, CI2%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, 1, NC%, DEPX%, CAN)
           End If
           REGMOVI& = ULTREG&("MOVISTO")
           '
           If RERES& > 0 Then
             If RERES& <= ULTREG&("REPUREPA") Then
               YY$ = REGLEIDO$("REPUREPA", RERES&)
               CARESE = CVS(Mid$(YY$, 21, 4))
               Call REPLA(YY$, MKS$((-1) * CAN), 25, 4)
               Call GRAREG("REPUREPA", YY$, RERES&)
               GoTo 99
             End If
           End If
           '
           YY$ = REGBLAN$("REPUREPA")
           Call REPLA(YY$, MKS$(NPX&), 1, 4)
             NPI% = NPX& Mod 10000
           Call REPLA(YY$, MKI$(NPI%), 5, 2)
           Call REPLA(YY$, MKI$(CI2%), 7, 2)
           Call REPLA(YY$, MKS$((-1) * CAN), 25, 4)
           Call REPLA(YY$, "OR-" + TRIM$(Str$(NPX&)), 33, 12)
           Call REGAPP("REPUREPA", YY$)
         End If
       End If
     End If
99 Next U&
   Call CIERRARCH("MOVISTO")
   Call BLOQARCH("MOVISTO")
   Call CIERRARCH("REPUREPA")
   '
   ' CANCELAR TODAS LAS RESERVAS
   Call ABRESERVA
   SQLX$ = "SELECT * FROM RESERVA"
   SQLX$ = SQLX$ + " WHERE IdSubOr = '" & DOPRI$ & "' "
   'Set RESELE = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
   Dim RESELE As ADODB.Recordset
   Set RESELE = New ADODB.Recordset
   RESELE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText

   With RESELE
101   On Error Resume Next
      .MoveFirst
      If Err < 1 Then
         On Error Resume Next
         .Delete
         If Err Then GoTo 102
         GoTo 101
      End If
      On Error GoTo 0
   End With
102 RESELE.Close
   '
   Call LIBARCH("REPUREPA")
   '
   Call BLOQARCH("PRESREPA")
   '
   HOII% = HOY(HOS$)
   TTXX$ = Space$(96)
   Call REPLA(TTXX$, HOS$ + "  " + "Cierre Orden de Reparación " + TRIM$(Str$(NPX&)), 1, 56)
   Call REPLA(TTXX$, USERTER$, 69, 24)
   Call REGAPP("ANOTAREP", Left$(MKS$(NPX&) + TTXX$, 96))
   Call REGAPP("ANOTAREP", MKS$(NPX&) + String$(88, "="))
   Call CIERRARCH("ANOTAREP")
   '
   TTXX$ = Space$(96)
   Call REPLA(TTXX$, HOS$ + "  " + "Cierre O/Rep " + TRIM$(Str$(NPX&)) + " - Talón " + TRIM$(TALOX$), 1, 56)
   Call REPLA(TTXX$, USERTER$, 69, 24)

   Call AGREGA_ANOTAPED(NROPE&, TTXX$)
   Call AGREGA_ANOTAPED(NROPE&, String$(88, "="))
   '
   'GRABACION EN HISTORIAL DE NUMERO DE SERIE
   '
   If NSERIE$ <> "" Then
    CUERPO$ = "Cierre de Orden de Reparación"
    NSERIE$ = TRIM$(Text1.TEXT)
    Call ANOTANUSE(NSERIE$, CUERPO$, , "NOMESS")
   End If
   '
   SQLX$ = "UPDATE PEDDETA SET StaDespa = 1, AcuPrepa = 1 "
   SQLX$ = SQLX$ + " WHERE NroPed = " & NROPE& & " "   ' COINCIDE NÚMERO DE PEDIDO
   SQLX$ = SQLX$ + " AND CodiInt = " & CIXI% & " AND NroClie = " & NCLIE% & " "
   '
   On Error Resume Next
   FLEXCONN.Execute SQLX$
   If Err Then
      On Error GoTo 0
      Call MENSERR(24, "No se Encontró la Nota de Venta\Correspondiente a esta Reparación.\  \Consulte a su Soporte de Sistemas.")
   End If
   On Error GoTo 0
   '
90 Call LIBARCH("PRESREPA")
   Call CIERRARCH("DESPAPED")
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
   GoTo 3
   '
End Sub
'
Function PRONUOCO&()
   '
   PRONU& = 1
   REGXX& = ULTREG&("OCOMENCA")
   If REGXX& > 0 Then
     PRONU& = 1 + CVS(Left$(REGLEIDO$("OCOMENCA", REGXX&), 4))
   End If
   If PRONU& <= REGXX& Then
     Screen.MousePointer = 11
     For U& = 1 To REGXX&
       X0$ = REGLEIDO$("OCOMENCA", U&)
       If CVS(Left$(X0$, 4)) >= PRONU& Then
         PRONU& = 1 + CVS(Left$(X0$, 4))
       End If
     Next U&
     Screen.MousePointer = 1
   End If
   '
   PRONUOCO& = PRONU&
   '
End Function

Function PRONUOTER&()    ' numero orden de servicio terceros
   '
   PRONU& = 1
   REGXX& = ULTREG&("OPERTER")
   If REGXX& > 0 Then
     PRONU& = 1 + CVS(Mid$(REGLEIDO$("OPERTER", REGXX&), 75, 4))
   End If
   If PRONU& <= REGXX& Then
     Screen.MousePointer = 11
     For U& = 1 To REGXX&
       X0$ = REGLEIDO$("OPERTER", U&)
       If CVS(Mid$(X0$, 75, 4)) >= PRONU& Then
         PRONU& = 1 + CVS(Mid$(X0$, 75, 4))
       End If
     Next U&
     Screen.MousePointer = 1
   End If
   '
   PRONUOTER& = PRONU&
   '
End Function



Private Sub Command8_Click()
   Command8.Enabled = False
   Call ABREORFAB
   '
   Call COPYSTRU("INDIPRES", "INDRECFI")
   
10 Call CARGALISEL("INDRECFI", "REPARA", "NUMREPA/F6.0;REFEREN/A56", "NUMREPA > 0 AND STATUS >= 3", "DISTINCT")
   Call SELECHO("INDRECFI/Indice General de Reparaciónes ")
   NroRep& = XVALO(TRIM$(VALACT1("INDRECFI")))
   '
   If NroRep& < 1 Then GoTo 99
   '
   NUSERIE$ = VALOBADA("REPARA", "NUMEROSERIE", "NUMREPA = '" & NroRep& & "'", "NOMESS")
   
   RECONRE07.Label24 = NUSERIE$
   RECONRE07.Show 1
   GoTo 10
99 Command8.Enabled = True
End Sub


Private Sub Command9_Click_VIE()
    '
    If DERACCE%("NOTAREP", "Nota de Reparación") < 2 Then
      Exit Sub
    End If

    Call ABREORFAB
    Call CREATABLA_EQUIPOS

10  OPX% = ALTERNA%("Nueva Reparación\Corregir Datos")
    If OPX% < 1 Or OPX% > 2 Then Exit Sub
    '
    If OPX% = 1 Then
        'NUEVO
        Call SELECHO("DEUDOR1")
        NC% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
        If NC% < 1 Then GoTo 10 ' ESTO SE ESTA VIENDO EL FRM COMO MODIFICAR PARA
        '                       ' QUE PRESENTE LA INFORMACION DEL DEUDOR12 Y GRABE EN REPARA
        CONDI$ = "NUMREPA > 0"
        NroRep& = 1 + XVALO(VALOBADA("REPARA", "MAX(NUMREPA)", CONDI$, "NOMESS"))
        FREP = HOY(HO$)
      ElseIf OPX% = 2 Then
        'CORREGIR DATOS
        Call COPYSTRU("INDIPRES", "INDERPCI")
15      Call CARGALISEL("INDERPCI", "REPARA", "NUMREPA/F6.0;REFEREN/A56", "NUMREPA > 0  AND STATUS < 4", "DISTINCT")
        Call SELECHO("INDERPCI/Indice General de Reparación")
        NroRep& = XVALO(TRIM$(VALACT1("INDERPCI")))
        If NroRep& < 1 Then GoTo 10
        NC% = XVALO(VALOBADA("REPARA", "NUME_CLI", "NUMREPA = " & NroRep& & ""))
        FREP = CVINT(VALOBADA("REPARA", "FechRecep", "NUMREPA = " & NroRep& & ""))
        If FREP < 1 Then FREP = HOY(HO$)
    End If
    '
    FILTX$ = TRIM$(Str$(NroRep&)) & ";" & TRIM$(Str$(FREP)) & ";" & TRIM$(Str$(NC%))
    Call STDFORM("NREP-MDT", FILTX$)
    If STDSCREEN.APRO <> "OK" Then
       If OPX% = 2 Then GoTo 15
       Exit Sub
    End If
    '
'AQUI INTRODUCIR EL CODIGO INTERNO Y ARMAR EL TEXTO DE LA REFERENCIA
    DESKRIPCION$ = VALOFORM("Text-4")
    Ci% = 0
    NUMESER$ = VALOFORM("Text-7")
    '
    CONDI$ = "NUMEROSERIE = '" & NUMESER$ & "'"
    Ci% = XVALO(VALOBADA("TANUMSE", "COD_INTE", CONDI$, "NOMESS"))
    If Ci% > 0 Then
       DESKRIPCION$ = DESCRIT$(Ci%)
    End If
    '
    'VERIFICAR SI EL EQUIPO SE ENCUENTRA EN GARANTIA
    FE% = HOY(HO$)
    ESTREP% = 0
    '
    FECHCARANT% = FeVenGarant(NUMESER$)
    If FECHCARANT% > 0 Then
      If FE% < FECHCARANT% Then
         ESTREP% = 1
         Call COMUNI("Equipo en Garantía")
         GoTo 20
      End If
    End If
    '
    FECHAGARAREP% = FeVenGarantRep(NUMESER$) ' GARANTIA POR REPARACION
    If FECHAGARAREP% > 0 Then
       If FE% < FECHAGARAREP% Then
         ESTREP% = 2
         Call COMUNI("Reparación en Garantía")
       End If
    End If
    '
20  Call ACTUREGISTRO("REPARA", "COD_INTE", "NUMREPA = " & NroRep& & "", Ci%, REGAF&, "NOMESS")
    Call ACTUREGISTRO("REPARA", "DESCRIPCION", "NUMREPA = " & NroRep& & "", DESKRIPCION$, REGAF&, "NOMESS")
    Call ACTUREGISTRO("REPARA", "Status_Gar", "NUMREPA = '" & NroRep& & "'", ESTREP%, REGAF&, "NOMESS")
    '
    RACLI$ = VALOFORM("Param-5")
    REFER$ = TRIM$(Left$(DESKRIPCION$, 22)) & " - " & RACLI$
    '
    Call ACTUREGISTRO("REPARA", "REFEREN", "NUMREPA = " & NroRep& & "", REFER$, REGAF&, "NOMESS")
    Call ACTUREGISTRO("REPARA", "STATUS", "NUMREPA = " & NroRep& & "", "0", REGAF&, "NOMESS")
    '
    CUERPO$ = "Recepción del Equipo para Reparación" & vbCrLf
    If OPX% = 2 Then ' CORREGIR DATOS
       CUERPO$ = "Modificacion de Recepción de Equipo para Reparación" & vbCrLf
    End If
    CUERPO$ = CUERPO$ + "Número de Reparación: " & NroRep& & vbCrLf
    CUERPO$ = CUERPO$ + "Cliente: " & RACLI$
    Call ANOTANUSE(NUMESER$, CUERPO$, , "NOMESS")

    
    If OPX% = 2 Then GoTo 15
    '
    Call COMUNI("Equipo Ingresado Para Presupuestación")
    '
End Sub

Private Sub Command9_Click()
    '
    Command9.Enabled = False
    Call ABREORFAB
    Call CREATABLA_EQUIPOS
    '
    If DERACCE%("NOTAREP", "Nota de Reparación") > 1 Then
       Call RECEP_EQUIP
    End If
    '
    Command9.Enabled = True
    '
End Sub

Private Sub Exit_Click()
   Call Command1_Click
End Sub

Private Sub Form_Load()
    '
    Command43.Visible = True
    Show
    Refresh
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Close: End
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    EPAC$ = TRIM$(EMPREAC$)
    Call ABRECONEXION
    If InStr(EPAC$, "EYM") > 0 Then
        Planipre.Visible = False
        mnuRevCierrreRepa.Visible = False
        mnuRecuperaFechaCierreReparacion.Visible = False
        mnuReparTecExt.Visible = False
        mnuApli.Visible = False
        mnuAgrpDescTarea.Visible = True
        Call cambiarEstado5a4
        
    End If

    Call GRATITFOR(Caption)
    '
    Call InstalaFuenteAnchoFijo
    Call CREATABLA_OFABRIC
    Call GENERAR_ESTRUCTURAS_SQL
    '
    Call CREACAMPO("OFABRIC", "REPARA", "SecRepar", 10, 48)
    Call CREACAMPO("OFABRIC", "REPARA", "Usu_Recep", 10, 48)
    Call CREACAMPO("OFABRIC", "REPARA", "Ctac_Cli", 10, 48)
    Call CREACAMPO("OFABRIC", "REPARA", "Reacondi", 12, 0)
    Call CREACAMPO("OFABRIC", "REPARA", "Tec_Repara", 10, 64)
    Call CREACAMPO("OFABRIC", "REPARA", "DeTipFalla", 12, 0)
    Call CREACAMPO("OFABRIC", "REPARA", "DeCauFalla", 12, 0)
    '
    Call CREACAMPO("OFABRIC", "TIPOFALL", "NUMREPA", 4, 0, 1)
    Call CREACAMPO("OFABRIC", "TIPOFALL", "ORIGEN", 3, 0, 1)     ' 1 CLIENTE; 2 TECNICO
    Call CREACAMPO("OFABRIC", "TIPOFALL", "CODIFALL", 4, 0, 1)
    Call CREACAMPO("OFABRIC", "TIPOFALL", "DESCFALL", 10, 128, 0)
    '
    Call CREACAMPO("OFABRIC", "CAUSFALL", "NUMREPA", 4, 0, 1)
    Call CREACAMPO("OFABRIC", "CAUSFALL", "CODIFALL", 4, 0, 1)
    Call CREACAMPO("OFABRIC", "CAUSFALL", "DESCFALL", 10, 128, 0)
    '
    Call CREACAMPO("OFABRIC", "ACCEREP", "NUMREPA", 4, 0, 1)
    Call CREACAMPO("OFABRIC", "ACCEREP", "DESCRIACC", 10, 128, 0)
    Call CREACAMPO("OFABRIC", "ACCEREP", "CANTIDAD", 4, 0, 1)
    Call CREACAMPO("OFABRIC", "ACCEREP", "OBSERVA", 10, 128)
    '
    Call ABRECABAN
End Sub
Function NUPRESUP&()
   '
   NPP& = 1
   '
   FIN& = ULTREG&("PRESREPA")
   If FIN& > 0 Then
      For U& = FIN& To FIN& - 128 Step -1
         If U& < 1 Then GoTo 99
         XX$ = REGLEIDO$("PRESREPA", U&)
         NPPX& = CVS(Mid$(XX$, 1, 4))
         If NPPX& >= NPP& Then
            NPP& = 1 + NPPX&
         End If
      Next U&
   End If
   '
99 NUPRESUP& = NPP&
   '
End Function
'

Private Sub Genpresu_Click()
   Call Command3_Click
End Sub

Private Sub Lispre_Click()
   Call CONFREPA07.Command5_Click
End Sub



Private Sub Master_Click()
    Call CIERRARCH("*.*")
    ACONEC$ = "ARCHIG07/AMASTO"
    Call CONECRUN(ACONEC$, "Maestro de Artículos de Stock")

'   Call PLAPRERE07.Command8_Click
End Sub

Private Sub Mater_Orep_Click()
   Option5.Value = True
   Call Command7_Click
End Sub

Private Sub Oper_Orep_Click()
   Option6.Value = True
   Call Command7_Click
End Sub

Private Sub mnuAgRemi_Click()
   REMIDATO.Show 1
End Sub

Private Sub mnuAgrpDescTarea_Click()

10  OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
         Exit Sub
    End If
     '
    
    VDEF$ = OBX$
    Des% = CVI(Left$(OBX$, 2))
    Has% = CVI(Mid$(OBX$, 3, 2))
    '
    If Des% > Has% Then
         GoTo 10
    End If
    Call PoneEnCeroNull("COTIEYM", "DESTAREA", REGAF&, "NOMESS")
    Desde$ = FETEXCOR1$(Des%)
    Hasta$ = FETEXCOR1$(Has%)
    SQLX$ = "SELECT COUNT(DESTAREA) as tot, DESTAREA   "
    SQLX$ = SQLX$ + " FROM COTIEYM c WITH(NOLOCK) inner join repara r on r.numrepa = c.numrepa "
    SQLX$ = SQLX$ + " where (destarea <> '_' AND destarea <> '')"
    SQLX$ = SQLX$ + "  AND FECHRECEP >= '" & Desde$ & "'"
    SQLX$ = SQLX$ + "  AND FECHRECEP <= '" & Hasta$ & "'"
    
    SQLX$ = SQLX$ + " group by   c.destarea"
    SQLX$ = SQLX$ + " having count(r.numrepa) > 0 order by tot desc"
    
    Set RS1 = READSET(SQLX$)
    Campos$ = "Cantidad/F10;Descripción de Tarea/A84"
    If CANTTIPO% > 0 Then Campos$ = Campos$ + ";Tipo Reparación/A24"
    
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    
    Dim obj As New RECORXET
    With RS1
      Do While Not .EOF
         CANTI = XVALO(!TOT)
         REF$ = SAFETEXT$(!DESTAREA)
         '
         i& = i& + 1
         GRILLAGRAL.GRID.Rows = i& + 1
         GRILLAGRAL.GRID.TextMatrix(i&, 1) = CANTI
         GRILLAGRAL.GRID.TextMatrix(i&, 2) = REF$
         .MoveNext
      Loop
    End With
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    '
    GRILLAGRAL.Caption = " Acumulado de Tareas (Rango: " & FECHATEX$(Des%) & " al " & FECHATEX$(Has%) & ")"
    GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.mnuMenu2.Visible = False
    GRILLAGRAL.mnuMenu3.Visible = False
    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.TXTBUSCAR = ""
    Call numerargrilla(GRILLAGRAL.GRID, 0)
    GRILLAGRAL.Show 1
'    Call CARGA_TABLA_IMPRE(GRILLAGRAL.GRID, Campos$, "DESTARFE", " Acumulado de Tareas (Rango: " & FECHATEX$(Des%) & " al " & FECHATEX$(Has%) & ")", "TABLA_IMPRE", , 1)


   Dim MSF As msFlexGrid
   If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(GRILLAGRAL.GRID, Campos$, "DESTARFE", " Acumulado de Tareas (Rango: " & FECHATEX$(Des%) & " al " & FECHATEX$(Has%) & ")", "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
        Exit Sub
   End If
   
   GoTo 10

End Sub

Private Sub mnuApli_Click()
   APLIREPA.Show 1
End Sub



Private Sub mnuConsRepa_Click()
'SELECT NUMREPA,NUMEROSERIE,REFEREN, FECHRECEP,TELE_CLI,
'Case Status
'WHEN  1 THEN 'RECEPCIONADO'
'WHEN  2 THEN 'COTIZADO'
'End
'AS ESTADO
',NUMREMITO,FECHENT,
'Case STATUS_GAR
'WHEN  0 THEN 'SIN GARANTIA'
'WHEN  1 THEN 'EN GARANTIA DE FÁBRICA'
'WHEN  2 THEN 'EN GARANTÍA DE REPARACIÓN'
'End
'AS GARANTIA
', FACTURA
' From REPARA
'   'CONSIDERA EN EL CASO DE EYM INDICAN EL TIPO DE REPARACION DESDE UNA LISTA QUE CARGAN ELLOS
     'CONSIDERANDO SI ES RECLAMO , RECLAMO EN CATANTIA, NO CONFORMIDAD ETC.
    CANTTIPO% = CantRegistros&("DATASQL", "LISTA = 'TIPOREP'", "NOMESS")
    If CANTTIPO% > 0 Then
       Dim TIPOREPARACION$()
       ReDim Preserve TIPOREPARACION$(CANTTIPO%)
       SQLX$ = "SELECT CAMPO1,CAMPO2 FROM DATASQL WITH(NOLOCK) WHERE LISTA = 'TIPOREP'"
       Set rs = READSET(SQLX$)
       With rs
        Do While Not .EOF
          IND% = XVALO(!CAMPO1)
          If IND% > 0 Then
             TIPOREPARACION$(IND%) = SAFETEXT$(!CAMPO2)
          End If
          .MoveNext
        Loop
       End With
    End If
    'FIN CARGA VECTOR EYM
    
    SQLX$ = " SELECT distinct R.NUMREPA,R.NUMEROSERIE,R.REFEREN, R.FECHRECEP,R.FechOrdRep, R.FechaCierreRepa,R.TELE_CLI,"
    
    SQLX$ = SQLX$ + " Case R.Status "
    SQLX$ = SQLX$ + " WHEN  0 THEN 'A COTIZAR' "
    SQLX$ = SQLX$ + " WHEN  2 THEN 'COTIZADO' "
    SQLX$ = SQLX$ + " WHEN  3 THEN 'A REPARAR' "
    SQLX$ = SQLX$ + " WHEN  4 THEN 'CERRADO' " 'ANTES REPARADO
'    SQLX$ = SQLX$ + " WHEN  4 THEN 'A ENTREGAR' "
    SQLX$ = SQLX$ + " WHEN  5 THEN 'A ENTREGAR' " 'le deberian dar salida, pero hasta no realizar el remito real no se entrego
    SQLX$ = SQLX$ + " End"
    SQLX$ = SQLX$ + " AS ESTADO "
    
    SQLX$ = SQLX$ + " ,R.NUMREMITO,R.FECHENT, "
    
    SQLX$ = SQLX$ + " Case R.STATUS_GAR "
    SQLX$ = SQLX$ + " WHEN  0 THEN 'SIN GARANTIA' "
    SQLX$ = SQLX$ + " WHEN  1 THEN 'EN GARANTIA DE FÁBRICA' "
    SQLX$ = SQLX$ + " WHEN  2 THEN 'EN GARANTÍA DE REPARACIÓN' "
    SQLX$ = SQLX$ + " End "
    SQLX$ = SQLX$ + " AS GARANTIA "
    
    SQLX$ = SQLX$ + " , R.FACTURA "
    
    'SI ESTA CARGADA LA TABLA DATASQL CON LOS TIPOS DE REPARACION DE EYM
    If CANTTIPO% > 0 Then
        SQLX$ = SQLX$ + " ,Case C.TIPOREPA "
        For J% = 1 To CANTTIPO%
           SQLX$ = SQLX$ + " WHEN  " & J% & " THEN  '" & TIPOREPARACION$(J%) & "'"
        Next J%
        SQLX$ = SQLX$ + " End"
        SQLX$ = SQLX$ + " AS TIPO_REPARACION"
    End If
'   FIN DE WHEN  PARA EYM
    
    SQLX$ = SQLX$ + "  From REPARA R WITH(NOLOCK)"
    
    'CONTINUA CON LOS TIPO DE REPARACION DE EYM
    If CANTTIPO% > 0 Then
        SQLX$ = SQLX$ + " left JOIN COTIEYM C "
        SQLX$ = SQLX$ + " ON R.NUMREPA = C.NUMREPA "
    End If
    'FIN
    SQLX$ = SQLX$ + " order by NUMREPA DESC"
    Set RS1 = READSET(SQLX$)
    Campos$ = "N° Rep./F9;N° de Serie/A16;Referencia/A24;Recep./D8;Teléfono/A16;Estado/A12;Gen.Rep/d8;Remito/A18;Cierre/D8;Entrega/d8;Gar.en Rep./a12;Factura/A18;nada/A0"
    If CANTTIPO% > 0 Then Campos$ = Campos$ + ";Tipo Reparación/A24"
    
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    
    Dim obj As New RECORXET
    With RS1
      Do While Not .EOF
         NR = XVALO(!numrepa)
         NUSERX$ = SAFETEXT$(!NUMEROSERIE)
         REF$ = SAFETEXT$(!Referen)
         FERECEP% = CVINT(!FECHRECEP)
         TELEF$ = SAFETEXT$(!TELE_CLI)
         ESTA$ = SAFETEXT$(!ESTADO)
         REMI$ = SAFETEXT$(!NUMREMITO)
         FENTRE% = CVINT(!FECHENT)
         GARANTENREPARA$ = SAFETEXT$(!GARANTIA)
         FACTU$ = SAFETEXT$(!FACTURA)
         GENREP% = CVINT(!FechOrdRep)
         CIERRE% = CVINT(!FechaCierreRepa)
         If CANTTIPO% > 0 Then TIPOREPAX$ = SAFETEXT$(!TIPO_REPARACION)
         '
         i& = i& + 1
         GRILLAGRAL.GRID.Rows = i& + 1
         GRILLAGRAL.GRID.TextMatrix(i&, 1) = NR
         GRILLAGRAL.GRID.TextMatrix(i&, 2) = NUSERX$
         GRILLAGRAL.GRID.TextMatrix(i&, 3) = REF$
         GRILLAGRAL.GRID.TextMatrix(i&, 4) = FECHATEX$(FERECEP%)
         GRILLAGRAL.GRID.TextMatrix(i&, 5) = TELEF$
         GRILLAGRAL.GRID.TextMatrix(i&, 6) = ESTA$
         GRILLAGRAL.GRID.TextMatrix(i&, 7) = FECHATEX$(GENREP%)
         GRILLAGRAL.GRID.TextMatrix(i&, 8) = REMI$
         GRILLAGRAL.GRID.TextMatrix(i&, 9) = FECHATEX$(CIERRE%)
         
         GRILLAGRAL.GRID.TextMatrix(i&, 10) = FECHATEX$(FENTRE%)
         GRILLAGRAL.GRID.TextMatrix(i&, 11) = GARANTENREPARA$
         GRILLAGRAL.GRID.TextMatrix(i&, 12) = FACTU$
         GRILLAGRAL.GRID.TextMatrix(i&, 13) = "" 'NADA
         If CANTTIPO% > 0 Then GRILLAGRAL.GRID.TextMatrix(i&, 14) = TIPOREPAX$
         
         .MoveNext
      Loop
    End With
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    '
    GRILLAGRAL.Caption = " REPARACIONES(Rango: " & FECHATEX$(Des%) & " al " & FECHATEX$(Has%) & ")"
    GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.mnuMenu2.Visible = False
    GRILLAGRAL.mnuMenu3.Visible = False
    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.TXTBUSCAR = ""
    Call numerargrilla(GRILLAGRAL.GRID, 0)
    GRILLAGRAL.Show 1
'    Call CARGA_TABLA_IMPRE(GRILLAGRAL.GRID, CAMPOS$, "PROYOBRA", "PROYECCION DE MANO DE OBRA", "TABLA_IMPRE", , 1)
'    Call STDFORM("PMO-JYM", FILTX$)

End Sub

Private Sub mnuControlFinal_Click()
   Call ABREORFAB
   '
   EPAC$ = TRIM$(EMPREAC$)
   If InStr(EPAC$, "EYM") > 0 Then
      RegCFinEym.Command4.Enabled = False
      RegCFinEym.Show 1
      Exit Sub
   End If

   Call COPYSTRU("INDIPRES", "INDRECFI")
   
10 Call CARGALISEL("INDRECFI", "REPARA", "NUMREPA/F6.0;REFEREN/A56", "NUMREPA > 0 AND STATUS >= 3", "DISTINCT")
   Call SELECHO("INDRECFI/Indice General de Reparaciónes ")
   NroRep& = XVALO(TRIM$(VALACT1("INDRECFI")))
   '
   If NroRep& < 1 Then Exit Sub
   '
   NUSERIE$ = VALOBADA("REPARA", "NUMEROSERIE", "NUMREPA = '" & NroRep& & "'", "NOMESS")
   EPAC$ = TRIM$(EMPREAC$)

   RECONRE07.Label45 = CStr(NroRep&)
   RECONRE07.Label24 = NUSERIE$
   RECONRE07.Show 1
   GoTo 10
End Sub

Private Sub mnuRecuperaFechaCierreReparacion_Click()
    
    Dim Rst As ADODB.Recordset
    ConsultaSql$ = "SELECT NUMREPA, FECHACIERREREPA FROM REPARA WITH(NOLOCK) WHERE FECHRECEP >= '" & FETEXCOR1$(FECHANUM("01/05/18")) & "' AND FECHACIERREREPA IS NULL"
    Set Rst = READSET(ConsultaSql$)
    FIN& = Rst.RecordCount
    With Rst
        Do While Not .EOF
            Call TRACE(24, i&, FIN&)
            i& = i& + 1
            
            DocumentoPrimario$ = CStr(!numrepa)
            FECHACIERRE% = CVINT(VALOBADA("MOVISTO", "FECHMOV", "CODIMOVI='RP' AND REFERCOR LIKE '%" & DocumentoPrimario$ & "%'"))
            'Call COMUNI("Documento: " & DocumentoPrimario$ & " - Cierre: " & FECHATEX$(FechaCierre%))
            If FECHACIERRE% > 0 Then
                'Call COMUNI("Recupera " & DocumentoPrimario$)
                Call ACTUREGISTRO("REPARA", "FECHACIERREREPA", "NUMREPA=" & CStr(!numrepa), CVDAT(FECHACIERRE%), REAF&)
            End If
            
            .MoveNext
        Loop
    End With
    On Error Resume Next
    Rst.Close
    Set Rst = Nothing
    On Error GoTo 0
    
    Call COMUNI("Proceso de Recuperación Finalizado")
    
    
End Sub

Private Sub mnuRemitosEntrega_Click()
     Call CARSELMESA
     '
10   Screen.MousePointer = 1
     Call SELECHO("SELMESA")
     MESASEL$ = VALACT1$("SELMESA")
     If TRIM$(MESASEL$) = "" Then
        GoTo 99
     End If
     Call ABRESTOCKS
     '
     CORRESQ$ = MESANO$(MESASEL$, Des%, Has%)
     '
     Desde = FETEXCOR1$(Des%)
     Hasta = FETEXCOR1$(Has%)
     '
     CMO$ = "RP"
     
     CONDI$ = "FechMov >= '" & Desde & "'"
     CONDI$ = CONDI$ + "AND FechMov <= '" & Hasta & "'"
     CONDI$ = CONDI$ + " AND substring(REFERCOR,1,4) <> 'ANUL' " 'PARA EVITAR COMPROBANTES ARRUINADOS
     CONDI$ = CONDI$ + " AND (CodiMovi =  '" & CMO$ & "'"
     CONDI$ = CONDI$ + " OR CodiMovi =  '" & CMO1$ & "')"
     CONDI$ = CONDI$ + " AND substring(DOPRILAR,1,2) = 'RP'"
     '
     Screen.MousePointer = 11
     '
     Call COPYSTRU("SELDOCAP", "SELREMIT")
12   Call CARGALISEL("SELREMIT", "MOVISTO", "DOPRILAR/A18;FechMov/D8;REFERCOR/A24", CONDI$, "DISTINCT")
     '
     If ULTREG&("SELREMIT") < 1 Then
        Call MENSERR(24, "No se Encontraron Remitos en el Rango Seleccionado.")
        GoTo 10
     End If

15   Call SELECHO("SELREMIT")
     '
     Screen.MousePointer = 1
     '
     NUDOCU$ = TRIM$(VALACT1$("SELREMIT"))
'     If Left$(VALACT1$("SELREMIT"), 2) = "RP" Then
'       NUDOCU$ = TRIM$(Mid$(VALACT1$("SELREMIT"), 4))
'     End If
     FECHACBTE$ = Mid$(VALACT2$("SELREMIT"), 1, 8)
     If NUDOCU$ = "" Then GoTo 10
    
'==========================================================================
'     NUDOCU$ = "RP." & TRIM$(Mid$(VALACT1$("SELREMIT"), 11))
     TIDOCUM$ = "REM-REPA"
     nubus = XVALO(Mid$(NUDOCU$, 4))
     '
     Screen.MousePointer = 11
     '
     DOCUNU& = 0
     'LE PASO EN LA PARAMETRO DE CONTROL EL NOMESS PARA QUE SI NO ENCUENTRA EL DOCUMENTO
     'NO PRESENTE MENSAJE Y CONTINUE Y LO BUSQUE DE LA MANERA QUE ESTABA ANTES
     Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, Str(nubus), NUDOCU$, "NOMESS")
     Screen.MousePointer = 1
     GoTo 15

'========================================================================
99   Command5.Enabled = True

End Sub

Private Sub mnuReparTecExt_Click()

10  OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
         Exit Sub
    End If
     '
    
    VDEF$ = OBX$
    Des% = CVI(Left$(OBX$, 2))
    Has% = CVI(Mid$(OBX$, 3, 2))
    '
    If Des% > Has% Then
         GoTo 10
    End If
    Desde$ = FETEXCOR1$(Des%)
    Hasta$ = FETEXCOR1$(Has%)
      
    SQLX$ = "SELECT R.COD_INTE,R.NUMREPA,RA.NUME_CLI,RA.RASO_CLI,M.CODIGO ,R.CANT_SOL,RA.STATUS,RA.TEC_EXT,DD.CAMPO2 "
    SQLX$ = SQLX$ + " FROM REPUREPA R WITH(NOLOCK)"
    SQLX$ = SQLX$ + " INNER JOIN REPARA RA ON RA.NUMREPA=R.NUMREPA"
    SQLX$ = SQLX$ + " INNER JOIN  datasql dd on CONVERT(int, dd.campo1)  = ra.tec_ext "
    SQLX$ = SQLX$ + " INNER JOIN MASTER M ON M.CODINT = RA.COD_INTE "
    SQLX$ = SQLX$ + " WHERE RA.TEC_EXT > 0 AND DD.LISTA = 'TECNIEXT' and RA.status <= 4"
    SQLX$ = SQLX$ + " and Ra.Fechrecep >=  '" & Desde$ & "'"
    SQLX$ = SQLX$ + " and Ra.Fechrecep <=  '" & Hasta$ & "'"
    SQLX$ = SQLX$ + " ORDER BY R.NUMREPA"
    Set RS1 = READSET(SQLX$)
    
    Campos$ = "Ci/A0;N° Rep./F9;Stat./f5;Cta./F7;R.Social/A24;Código/A16;Cant.Sol/f8.1;NTec/A0;Técnico/A12"
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Dim obj As New RECORXET
    With RS1
      Do While Not .EOF
         CI1% = XVALO(!COD_INTE)
         NR = XVALO(!numrepa)
         NC = XVALO(!Nume_Cli)
         rs$ = SAFETEXT$(!Raso_Cli)
         CODI$ = SAFETEXT$(!Codigo)
         CANTI = XVALO(!Cant_Sol)
         ST% = XVALO(!Status)
         TEC% = XVALO(!TEC_EXT)
         TECNICO$ = SAFETEXT$(!CAMPO2)
         '
         i& = i& + 1
         GRILLAGRAL.GRID.Rows = i& + 1
         GRILLAGRAL.GRID.TextMatrix(i&, 1) = CI1%
         GRILLAGRAL.GRID.TextMatrix(i&, 2) = NR
         GRILLAGRAL.GRID.TextMatrix(i&, 3) = ST%
         GRILLAGRAL.GRID.TextMatrix(i&, 4) = NC
         GRILLAGRAL.GRID.TextMatrix(i&, 5) = rs$
         GRILLAGRAL.GRID.TextMatrix(i&, 6) = CODI$
         
         GRILLAGRAL.GRID.TextMatrix(i&, 7) = FORMATNUM$(CANTI, "F9.1")
         GRILLAGRAL.GRID.TextMatrix(i&, 8) = TEC%
         GRILLAGRAL.GRID.TextMatrix(i&, 9) = TECNICO$
         
         .MoveNext
      Loop
    End With
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    '
    GRILLAGRAL.Caption = "Reparaciones De Técnicos Externos Pendientes (Rango: " & FECHATEX$(Des%) & " al " & FECHATEX$(Has%) & ")"
    GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.mnuMenu2.Visible = False
    GRILLAGRAL.mnuMenu3.Visible = False
    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.TXTBUSCAR = ""
    Call numerargrilla(GRILLAGRAL.GRID, 0)
    GRILLAGRAL.Show 1
'    Call CARGA_TABLA_IMPRE(GRILLAGRAL.GRID, CAMPOS$, "PROYOBRA", "PROYECCION DE MANO DE OBRA", "TABLA_IMPRE", , 1)
'    Call STDFORM("PMO-JYM", FILTX$)




End Sub

Private Sub mnuReportesVarios_Click()
    Call AGenRep_Click
End Sub

Private Sub mnuRevCierrreRepa_Click()

'10  TituloLista$ = "Ordenes de Reparaciones Cerradas"
'    CAMPOS$ = "NRep./F8,Referencia/A24;F.Recep/D8;Cta./F7,R.Social/A24,N.Serie/A16;N°Remito/A16;N°Pedido/F8;IdAsiento/A0;N°Téc./F6"
'    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO)
'    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO)
'    FRMZELECHO.Caption = TituloLista$
'
'    Dim obj As New RECORXET
'    Set RS1 = obj.RS_LISREPACERRADAS()
'    Call Carga_MSF_Recordset(RS1, CAMPOS$, FRMZELECHO.msf2, 1)
'    On Error Resume Next
'    Set obj = Nothing
'    RS1.Close
'    Set RS1 = Nothing
'    On Error GoTo 0
'
'    FRMZELECHO.Width = 14000
'50  FRMZELECHO.Show 1
'
'    NUREP& = XVALO(RESP_ZELE_VECT(1))
'    If NUREP& > 0 Then
'      OPX% = COMALTER%("Realmente Desea Revertir el Cierre de La Reparación N°: " & NUREP& & "\\Cancelar\Revertir Cierre")
'      If OPX% < 2 Then
'         GoTo 10
'      End If
'    Else
'      Exit Sub
'    End If
'    Screen.MousePointer = 11
'    '================================================================================================================
'    'VALIDO SI TIENE PEDIDO ASOCIADO CON ENTREGAS
'    '================================================================================================================
'    NPED& = XVALO(RESP_ZELE_VECT(8))
'    If NPED& > 0 Then
'        CANTI_ENT = XVALO(VALOBADA("PEDDETA", "SUM(CANTENT)", "NROPED = " & NPED&, "NOMESS"))
'        If CANTI_ENT > 0 Then
'           Call COMUNI("Imposible Revertir Cierre\La Orden de Reparación N°: " & NUREP& & "\Generó el Pedido N°: " & NPED& & "\El Mismo Yá Tiene Entregas Realizadas")
'           GoTo 10
'        End If
'    End If
'    '================================================================================================================
'    'ANULACIÓN DEL PEDIDO Y ESCRIBE EL ANOTADOR
'    '================================================================================================================
'    strsql = " Update PedEnca set Status = 9, "
'    strsql = strsql & "Referencia = '(anul)' + left(Referencia,42) "
'    strsql = strsql & ", FechAnul = '" & CVDAT(HOY(HO$)) & "'"
'    strsql = strsql & " where Status < 9"
'    strsql = strsql & " And NroPed =" & NPED&
'    FLEXCONN.Execute (FILTSQL$(strsql))
'    '
'    strsql = strsql & " Update PedDeta set Status = case "
'    strsql = strsql & " when (CantEnt = 0) then 9 "
'    strsql = strsql & " when ( CantEnt > 0) then 8 end "
'    strsql = strsql & " where Status <> 3 and NroPed = " & NPED&
'    FLEXCONN.Execute (FILTSQL$(strsql))
'
'    HOII% = HOY(HOS$)
'    TTXX$ = Space$(96)
'    Call REPLA(TTXX$, HOS$ + "  " + "Anulación Automática Desde Reparación (N°" & NUREP& & ")", 1, 56)
'    Call REPLA(TTXX$, USERTER$, 69, 24)
'    Call AGREGA_ANOTAPED(NPED&, TTXX$)
'    Call AGREGA_ANOTAPED(NPED&, String$(88, "="))
'    '================================================================================================================
'    '================================================================================================================
'
'    '================================================================================================================
'    'RECUPERO EL ULTIMO NUMERO DE SERIE
'    '================================================================================================================
'    VUELTA% = 0
'    Nserie$ = TRIM(RESP_ZELE_VECT(6))
'    SQLX$ = "SELECT TOP 2 NUMREPA,FECHRECEP FROM REPARA WITH(NOLOCK) WHERE NUMEROSERIE = '" & Nserie$ & "' ORDER BY NUMREPA DESC"
'    Set RS = READSET(SQLX$)
'    With RS
'      Do While Not .EOF
'        VUELTA% = VUELTA% + 1
'        ULTIMAREPA& = XVALO(!NumRepa)
'        ULTFECHAREP% = CVINT(!FECHRECEP)
'        .MoveNext
'      Loop
'    End With
'    '
'    'SI NO DIO LAS 2 VUELTAS ERA LA PRIMER REPARACION ENTONCES LIMPIA LA INFO
'    If VUELTA% < 2 Then ULTIMAREPA& = 0: ULTFECHAREP% = 0
'    Call ACTUREGISTRO("TANUMSE", "UnumRep", "NUMEROSERIE = '" & Nserie$ & "'", ULTIMAREPA&, REGAF&, "NOMESS")
'    Call ACTUREGISTRO("TANUMSE", "FevenGaRep", "NUMEROSERIE = '" & Nserie$ & "'", FETEXCOR1$(ULTFECHAREP%), REGAF&, "NOMESS")
'    '================================================================================================================
'    '================================================================================================================
'
'    '================================================================================================================
'    'BORRO EL ASIENTO
'    '================================================================================================================
'    ASIENTO$ = TRIM(RESP_ZELE_VECT(9))
'    Call BORRAREGISTROS("DETAASIEN", "NUIDASIEN = '" & ASIENTO$ & "'", REGAF&, "NOMESS")
'    Call BORRAREGISTROS("ENCABEASIEN", "NUIDASIEN = '" & ASIENTO$ & "'", REGAF&, "NOMESS")
'    '================================================================================================================
'    '================================================================================================================
'
'    '================================================================================================================
'    'RECUPERACION DEL RESERVA EN BASE AL MOVISTO
'    '================================================================================================================
'    NUMEREPA$ = "OR-" & TRIM$(Str$(NUREP&))
'    'TRASLADO DE MOVISTO A RESERVA
'    Call BORRAREGISTROS("RESERVA", "IDSUBOR = '" & NUMEREPA$ & "'", REGAF&, "NOMESS")
'
'    SQLX$ = "SELECT * FROM MOVISTO WITH(NOLOCK)"
'    SQLX$ = SQLX$ + " WHERE IDCALIDAD = " & NUREP&
'    SQLX$ = SQLX$ + " ORDER BY NUORITEM"
'    '
'    Dim RSMOVI As ADODB.Recordset
'    Set RSMOVI = FLEXCONN.Execute(FILTSQL$(SQLX$))
'    With RSMOVI
'    Do While Not .EOF
'         SQLX1$ = "SELECT * FROM RESERVA"
'         SQLX1$ = SQLX1$ + " WHERE COD_INTE < 1 "
'         Dim RESETMP As ADODB.Recordset
'         Set RESETMP = New ADODB.Recordset
'         RESETMP.Open FILTSQL$(SQLX1$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
'         With RESETMP
'           .AddNew   ' YA CONVERTIDO
'           On Error Resume Next
'           !CodiEmpr = CodiEmp%
'           On Error GoTo 0
'           CI% = XVALO(RSMOVI!COD_INTE)
'           CANTI = XVALO(RSMOVI!CANTSALID)
'           If CI% > 0 And CANTI > 0 Then
'            !COD_INTE = CI%
'            !COD_EXTE = CODEXT$(CI%)
'            !Descrip = Left$(DESCRIT$(CI%), 64) 'modificado por que devolvia un error de datos
'            !idsubor = NUMEREPA$
'            !FechRes = RSMOVI!FechMov
'            !CANTRES = CANTI
'            !CANTCONS = 0
'            !Saldo_Entre = CANTI
'            !IDENLOTE = RSMOVI!IDENLOTE
'            !DEPOSITO = RSMOVI!DEPOMOVI
'            !USOUNIT = CANTI
'            !NuOrItem = XVALO(RSMOVI!NuOrItem)
'            .Update
'           End If
'         End With
'         RESETMP.Close
'         Set RESETMP = Nothing
'    .MoveNext
'    Loop
'    End With
'
'    RSMOVI.Close
'    Set RSMOVI = Nothing
'
'    '================================================================================================================
'    'ANULACION DE MOVIMIENTOS EN EL MOVISTO
'    '================================================================================================================
'    SQLX$ = "DELETE  FROM MOVISTO WHERE IDCALIDAD = " & NUREP&
'    FLEXCONN.Execute (SQLX$)
'    '================================================================================================================
'    '================================================================================================================
'    '================================================================================================================
'    'ACTUALIZO EL STATUS DEL REPARA A 3
'    '================================================================================================================
'    Call ACTUREGISTRO("REPARA", "STATUS", "NUMREPA = " & NUREP&, 3, REGAF&, "NOMESS")
'    '================================================================================================================
'    '================================================================================================================
'    '================================================================================================================
'    'ANOTACION EN EL ANOTADOR
'    '================================================================================================================
'    CUERPO$ = vbCrLf & "Se Revirtió Cierre de Reparación N°: " & NUREP& & vbCrLf
'    CUERPO$ = CUERPO$ + " En Forma Autómatica "
'    Call ANOTANUSE(NUMESER$, CUERPO$, , "NOMESS")
'    '================================================================================================================
'    '================================================================================================================
'    Screen.MousePointer = 1
'    Call COMUNI("Se Revirtió Cierre de Reparación N°: " & NUREP&)
End Sub


Private Sub mnuTabkaCostosDeReparaciones_Click()
  Call GENERAR_ESTRUCTURAS_SQL
  IDLIST& = ObtenerIdTablaAuxiliar("TABCOSREPA")
  
  Call ABM_TABLA(IDLIST&)
End Sub

Private Sub mnuTablasAuxiliares_Click()
    Call CONFREPA09.mnuTablasAuxiliares_Click
End Sub

Private Sub mnuVincularReportes_Click()
    Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
End Sub

Private Sub Option1_DblClick()
   Call PRESEDOC
End Sub

Private Sub Option2_DBLClick()
   EPAC$ = TRIM$(EMPREAC$)
   If InStr(EPAC$, "EYM") > 0 Then
       Call MENSERR(24, "No Disponible en la Presente Versión")
   Else
       Call PRESEDOC
   End If
End Sub

Private Sub Option3_DblClick()
    EPAC$ = TRIM$(EMPREAC$)
    If InStr(EPAC$, "EYM") > 0 Then
       COTREPEYM.Show 1
    Else
       Call PRESEDOC
    End If
End Sub

 
Private Sub Option4_DblClick()
   EPAC$ = TRIM$(EMPREAC$)
   If InStr(EPAC$, "EYM") > 0 Then
      Call MENSERR(24, "No Disponible en la Presente Versión")
   Else
       Call PRESEDOC
   End If
End Sub

Private Sub Option5_DblClick()
    Call ABREORFAB
    '
    Call COPYSTRU("INDIPRES", "INDIRREM")
15  Call CARGALISEL("INDIRREM", "REPARA", "NUMREPA/F6.0;REFEREN/A56", "STATUS = 5", "DISTINCT")
    Call SELECHO("INDIRREM")
    NroRep& = XVALO(TRIM$(VALACT1("INDIRREM")))
    If NroRep& < 1 Then Exit Sub
    '
    REMI$ = VALOBADA("REPARA", "NUMREMITO", "NUMREPA = " & NroRep&, "NOMESS")
    CODXT$ = ""
    Ci% = XVALO(VALOBADA("REPARA", "COD_INTE", "NUMREPA = " & NroRep&, "NOMESS"))
    If Ci% > 0 Then
       CODXT$ = CODEXT$(Ci%)
    End If
    '
    FREP = CVINT(VALOBADA("REPARA", "FECHENT", "NUMREPA = " & NroRep&, "NOMESS"))
    '
    FILTX$ = TRIM$(Str$(NroRep&)) & ";" & TRIM$(REMI$) & ";" & TRIM$(Str$(FREP)) & ";" & TRIM$(CODXT$)
    Call STDFORM("REMX-MDT", FILTX$, "PRINT")
    '
End Sub


Private Sub Option6_DblClick()
   'Presentar Cargalisel
'   Call CARGALISEL("ESTAREPA", "REPARA", "NUMREPA/F6.0;REFEREN/A32;NUMEROSERIE/A14", "NUMREPA > 0 ", "DISTINCT")
'10 Call SELECHO("ESTAREPA/Indice General de Reparaciones Generadas")
'   NROREP& = XVALO(TRIM$(VALACT1("ESTAREPA")))
10  NroRep& = LISTAREPARACION&("NUMREPA > 0")
   
   If NroRep& < 1 Then Exit Sub
   '
   'ESTADO
   '*********
   EST% = XVALO(VALOBADA("REPARA", "STATUS", "NUMREPA = " & NroRep& & "", "NOMESS"))
   Select Case EST%
      Case 0
        ESTADO$ = "Ingresado"
      Case 1
        ESTADO$ = "Planilla de Presupuestación"
      Case 2
        ESTADO$ = "Generación de Presupuesto"
      Case 3
        ESTADO$ = "Generación de O.Reparación"
      Case 4
        ESTADO$ = "Orden de Reparación Cerrada"
      Case 5
        ESTADO$ = "Reparación ya Entregada"
      Case 9
        ESTADO$ = "ANULADA"
   End Select
   '
   'PRESENTAR FRM DE ESTADO
   '**********************
   FILTX$ = TRIM$(Str$(NroRep&)) & ";" & TRIM$(ESTADO$)
   Call STDFORM("NSTA-MDT", FILTX$)
   
   '
   GoTo 10
   'presentar standscreen
   
End Sub

Private Sub Orderep_Click()
   Option4.Value = True
   Option4_DblClick
End Sub

Private Sub Ordrep_Click()
   Call Command4_Click
End Sub

Private Sub Plabuscli_Click()
   Call Command20_Click
End Sub

Private Sub Pladir_Click()
   Call Command5_Click
End Sub

Private Sub Plarepa_Click()
   Option1.Value = True
   Option1_DblClick
End Sub

Private Sub Presucli_Click()
   Option3.Value = True
   Option3_DblClick
End Sub

Private Sub Presuint_Click()
   Option2.Value = True
   Option2_DBLClick
End Sub

Private Sub RtoRep_Click()
   Call Option5_DblClick
End Sub

Private Sub Setup_Click()
   Call Command6_Click
End Sub

Private Sub ESTAREP()
  '
  Call DESHASFECHA(Des%, Has%, PERIO$)
  '
  If TRIM$(PERIO$) = "" Then Exit Sub 'SI CANCELO
  '
  Screen.MousePointer = 11
  FIN& = ULTREG&("PRESREPA")
  JJ& = 0
  For U& = 1 To FIN&
    Call TRACE(20, U&, FIN&)
    XX$ = REGLEIDO$("PRESREPA", U&)
    FEIN1% = CVI(Mid$(XX$, 65, 2))   ' FECHA DE RECEPCION
    '
    If FEIN1% >= Des% Then
      If FEIN1% <= Has% Then
        JJ& = JJ& + 1
        NO1% = JJ&
        NUREP1& = CVS(Mid$(XX$, 1, 4))
        CLIE1% = CVI(Mid$(XX$, 67, 2))
        CI1% = CVI(Mid$(XX$, 69, 2))
        TALON$ = TRIM$(Mid$(XX$, 95, 6))
        FEPRE1% = CVI(Mid$(XX$, 121, 2))   ' FECHA PRESUPUESTO
        FECIE1% = CVI(Mid$(XX$, 123, 2))   ' FECHA DE CIERRE
        NUOVEN1 = CVS(Mid$(XX$, 125, 4))
        '
        NUOVE& = XVALO(NUOVEN1)
        'OBTENGO LA FECHA DE LA O/VEN
'        XZ1$ = FILTERGETRECORD$("PEDENCA", 1, MKS$(NUOVEN1))
'        FEOREP1% = CVI(Mid$(XZ1$, 53, 2))  ' FECHA DE APROBACION
        '
        '\\\OT-08-1033-OD-11/11/08
        Call ABREPEDCLI
        FEOREP1% = CVINT(VALOBADA("PEDENCA", "FechEmis", "NroPed = " & NUOVE& & ""))
        '\\\OT-08-1033-OD-FIN
        '
        YY$ = REGBLAN$("ESTAREP")
        '
        Call REPLA(YY$, MKI$(NO1%), 1, 2)
        Call REPLA(YY$, MKS$(NUREP1&), 3, 4)
        Call REPLA(YY$, MKI$(CLIE1%), 7, 2)
        Call REPLA(YY$, MKI$(CI1%), 9, 2)
        Call REPLA(YY$, MKI$(FEIN1%), 11, 2)
        Call REPLA(YY$, MKI$(FEPRE1%), 13, 2)
        Call REPLA(YY$, MKI$(FEOREP1%), 15, 2)
        Call REPLA(YY$, MKS$(NUOVEN1), 17, 4)
        Call REPLA(YY$, MKI$(FECIE1%), 21, 2)
        Call REPLA(YY$, TALON$, 23, 6)
        '
        Call GRAREG("ESTAREP", YY$, JJ&)
      End If
    End If
  Next U&
  '
  Call SETULTREG("ESTAREP", JJ&)
  Call CIERRARCH("ESTAREP")
  Screen.MousePointer = 1
  '
  If JJ& > 0 Then
    Call HEADERS("ESTAREP", "")
    Call HEADERS("ESTAREP", "Período: " & PERIO$)
    Call FINAL("*ESTAREP")
  End If
  '
End Sub
Private Sub OREPABI()
   '
   FIN& = ULTREG&("PRESREPA")
   For U& = 1 To FIN&
     Call TRACE(20, UI&, FIN&)
     XX$ = REGLEIDO$("PRESREPA", U&)
     If CVS(Mid$(XX$, 125, 4)) >= 0.5 Then
        If CVI(Mid$(XX$, 123, 2)) < 1 Then
           JJ& = JJ& + 1
           Call GRAREG("OREPARPE", XX$, JJ&)
        End If
     End If
   Next U&
   Call SETULTREG("OREPARPE", JJ&)
   Screen.MousePointer = 1
   '
   Call FINAL("*LIREPAPE")
   '
End Sub

Sub DEVOPORREP()
   '
   Dim MATRANSFE#(32767)
   '
10 NroRep& = LISTAREPARACION&("NUMREPA > 0 AND STATUS = 3")

'   Call ABREORFAB
'   Call COPYSTRU("INDIPRES", "INDERTRA")
'10 Call CARGALISEL("INDERTRA", "REPARA", "NUMREPA/F6.0;REFEREN/A56", "NUMREPA > 0 AND STATUS = 3", "DISTINCT")
'   Call SELECHO("INDERTRA/Indice General de Reparaciónes Generadas")
'   NROREP& = XVALO(TRIM$(VALACT1("INDERTRA")))
   If NroRep& < 1 Then Exit Sub
   '
   NUMEREPA$ = "OR-" & TRIM$(Str$(NroRep&))
   '
   For t& = 1 To NUMAS%
      MATRANSFE#(t&) = 0
   Next t&
   '
   If CONTROL$("", "DEPREDES") <> "" Then
     DEPORIGINAL% = CONTROL$("", "DEPREDES")
   End If
   '
   Call REPLA(VDEF$, NUMEREPA$, 1, 12)
   Call REPLA(VDEF$, Chr$(DEPORIGINAL%), 13, 1) ' DEPOSITO TALLER
   '
12 OBX$ = OBJPLA$("SEREPDEP", VDEF$)
   If OBX$ = "" Then GoTo 10
   '
   DepOrig% = Asc(Mid$(OBX$, 13, 1))
   DEPODEST% = Asc(Mid$(OBX$, 14, 1))
   '
'   'VALIDACION QUE TENGA DERECHOS PARA CAMBIAR DEPOSITOS
   If DERACCE%("CAMBIDEP/NOMESS", "CAMBIO DE DEPOSITOS") < 2 Then
     If DEPORIGINAL% <> DEPODEST% Then
        Call MENSERR(24, "Usuario no Habilitado para Modificar Depósito\Módulo: CAMBIDEP")
        GoTo 12
     End If
   End If
   '
   'VALIDA DEPOSITOS
   If DepOrig% < 1 Or ECOARCH$("TADEPOSI", Chr$(DepOrig%)) = "" Then 'VALIDACION
     Call MENSERR(24, "Depósito Origen No Válido")
     VDEF$ = OBX$
     GoTo 12
   End If
   '
   If DEPODEST% < 1 Or ECOARCH$("TADEPOSI", Chr$(DEPODEST%)) = "" Then
      Call MENSERR(24, "Depósito Destino No Válido")
      VDEF$ = OBX$
      GoTo 12
   End If
   '
   Screen.MousePointer = 11
   '
   ' RECORRO MOVISTO Y CARGO EL VECTOR CON LA CANTIDAD INGRESADA POR ARTICULO SEGUN LA O.F.
   SQLX$ = "SELECT * FROM MOVISTO  WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE DOSECCOR = '" & NUMEREPA$ & "' "
   SQLX$ = SQLX$ + " AND CODIMOVI = 'TR' "
   SQLX$ = SQLX$ + " AND DEPOMOVI = " & DepOrig% & ""
   Set MoviTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With MoviTemp
     Do While Not .EOF
       Ci% = XVALO(!COD_INTE)
       CAN = XVALO(!CANTINGRE) - XVALO(!CANTSALID)
       '
       MATRANSFE#(Ci%) = MATRANSFE#(Ci%) + CAN
       .MoveNext
     Loop
   End With
   MoviTemp.Close
   Set MoviTemp = Nothing
   '
   JJ& = 0
   Call SETULTREG("!MATEROF", 0)
   For i% = 1 To NUMAS%
     Ci% = i%
     UM1$ = Unimed$(Ci%)
     CAN1 = MATRANSFE#(Ci%)
     If MATRANSFE#(Ci%) > 0.000005 Then
        JJ& = JJ& + 1
        Y$ = REGBLAN$("!MATEROF")
        Call REPLA(Y$, MKI$(Ci%), 1, 2)
        Call REPLA(Y$, UM1$, 3, 2)
        Call REPLA(Y$, MKS$(CAN1), 5, 4)
        Call GRAREG("!MATEROF", Y$, JJ&)
     End If
   Next i%
   If JJ& < 1 Then
      Call COMUNI("No se Encontró Material sin Consumir Para Esta Orden de Fabricación")
      Screen.MousePointer = 1
      GoTo 10
   End If
   
   Screen.MousePointer = 1
   VTB% = VENTABU%("MATEROF1")
   
   If VTB% < 0 Then GoTo 10
   
   OPPX% = COMALTER%("Confirmación de Devolución\\Realizar Devolución\Cancelar")
   If OPPX% < 1 Or OPPX% > 1 Then
       GoTo 10
   End If
   '
   'Devolucion de Materiales
   CMOV$ = TRIM$(CONTROL$("", "CODVALMA"))
   If CMOV$ = "" Then CMOV$ = "TR"
   HOII% = HOY(HOS$)
   FECHA% = HOII%
   LOTE$ = NUMEREPA$
'   CMOV$ = "TR"
   DOSEC$ = "TR-" + TRIM$(Mid$(NUMEREPA$, 4))
   NRO& = ProNroComprobante(CMOV$)
   DOPRI$ = "TR." + TRIM$(Str$(NRO&))
   VUNI# = 0
   MONE$ = ""
   DOSEC$ = TRIM$(NUMEREPA$) ' NUMERO DE ORDEN DE FABRICACION
   While Len(DOSEC$) < 10
      DOSEC$ = DOSEC$ + " "
   Wend

   'Recorro el Materof

   FIN& = ULTREG&("!MATEROF")
   For h& = 1 To FIN&
      YY$ = REGLEIDO$("!MATEROF", h&)
      Ci% = CVI(Mid$(YY$, 1, 2))
      CAN = CVS(Mid$(YY$, 5, 4))
      If CAN > 0.05 Then
         LOTE$ = ""
         If ESTRAZABLE(Ci%) = 1 Then
            CONDIX$ = "COD_INTE = " & Ci% & " AND DOSECCOR = '" & NUMEREPA$ & "' AND CODIMOVI = 'TR' AND DEPOMOVI = " & DepOrig% & ""
            LOTE$ = VALOBADA("MOVISTO", "IdenLote", CONDIX$, "NOMESS")
         End If
         '
         REFE$ = "TR.de " + TRIM$(ECOARCH$("TADEPOSI", MKI$(DepOrig%)))
          '
         Call MOVISTOK(FECHA%, Ci%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NORI%, VUNI#, 0, NC%, DEPODEST%, CAN)
          '
         CAN = (-1) * CAN
         REFE$ = "TR.a " + TRIM$(ECOARCH$("TADEPOSI", MKI$(DEPODEST%)))
         Call MOVISTOK(FECHA%, Ci%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NORI%, VUNI#, 0, NC%, DepOrig%, CAN)
         
      End If
   Next h&
   '
   Screen.MousePointer = 1
   Call COMUNI("Proceso Terminado")
   
End Sub

Sub RECEP_EQUIP()             ' RECEPCION EQUIPO PARA REPARACION
   '
5  RECEPEQ09.Text1 = "."
   RECEPEQ09.Text1 = ""
   RECEPEQ09.Text2 = ""
   RECACCE09.Text1 = ""
   RECACCE09.Text2 = ""
   RECEPEQ09.Text13 = ""
   RECEPEQ09.Text14 = ""
   RECEPEQ09.Text12 = ""
   EPAC$ = TRIM$(EMPREAC$)

   
   Call RECACCE09.Form_Load
   '
   OPX% = ALTERNA%("Nueva Reparación\Corregir Datos")
   If OPX% < 1 Or OPX% > 2 Then Exit Sub
   '
   If OPX% = 1 Then
        'NUEVO
        CONDI$ = "NUMREPA > 0 "
        NroRep& = 1 + XVALO(VALOBADA("REPARA", "MAX(NUMREPA)", CONDI$, "NOMESS"))
        
        'SE AGREGA ESTO POR QUE HAY UN VALOR CONFIGURADO DE NRO DE INICIO
        'SI ESTE ES MAYOR COMIENZA A PARTIR DE ESTE,
        NROINICIO& = XVALO(CONTROL("REPARA", "NUEVANUM"))
        If NroRep& < NROINICIO& Then NroRep& = NROINICIO&
        '
        RECEPEQ09.Label15 = NroRep&
        RECEPEQ09.Label16 = Format(Now, "dd/mm/yy - hh:mm")
        RECEPEQ09.Label17 = USERNAME$
      ElseIf OPX% = 2 Then
        'CORREGIR DATOS
        Call COPYSTRU("INDIPRES", "INDERPCI")
15      Call CARGALISEL("INDERPCI", "REPARA", "NUMREPA/F6.0;REFEREN/A56", "NUMREPA > 0  AND STATUS < 4", "DISTINCT")
        Call SELECHO("INDERPCI/Indice General de Reparación")
        NroRep& = XVALO(TRIM$(VALACT1("INDERPCI")))
        If NroRep& < 1 Then GoTo 5
        RECEPEQ09.Label15 = NroRep&
           FREP = VALOBADA("REPARA", "FechRecep", "NUMREPA = " & NroRep& & "", "NOMESS")
           If FREP = "" Then FREP = Now
        RECEPEQ09.Label16 = Format(FREP, "dd/mm/yy - hh:mm")
           USRECEP$ = VALOBADA("REPARA", "USU_RECEP", "NUMREPA = " & NroRep& & "", "NOMESS")
           If USRECEP$ = "" Then USRECEP$ = USERNAME$
        RECEPEQ09.Label17 = USRECEP$      ' TOMAR DE LA BASE DE DATOS
        RECEPEQ09.Text2 = TRIM$(VALOBADA("REPARA", "NUMEROSERIE", "NUMREPA = " & NroRep& & "", "NOMESS"))
        RECEPEQ09.Text1 = XVALO(VALOBADA("REPARA", "NUME_CLI", "NUMREPA = " & NroRep& & "", "NOMESS"))
        RECEPEQ09.Label14 = FECHATEX$(CVINT(VALOBADA("REPARA", "FechEnSol", "NUMREPA = " & NroRep& & "", "NOMESS")))
        RECEPEQ09.Text9 = VALOBADA("REPARA", "Ctac_Cli", "NUMREPA = " & NroRep& & "", "NOMESS")
        RECEPEQ09.Text11 = VALOBADA("REPARA", "SecRepar", "NUMREPA = " & NroRep& & "", "NOMESS")
        RECEPEQ09.LBLSTATGARAN_ORIGINAL = 0
        RECEPEQ09.Label7 = "NO"
        RECEPEQ09.Text13 = XVALO(VALOBADA("REPARA", "TEC_EXT", "NUMREPA = " & NroRep& & "", "NOMESS"))
        RECEPEQ09.Text14 = VALOBADA("REPARA", "TALON_TEC", "NUMREPA = " & NroRep& & "", "NOMESS")
        RECEPEQ09.Text12 = NOMBRE_TEC_EXT$(XVALO(RECEPEQ09.Text13))
           If VALOBADA("REPARA", "Status_Gar", "NUMREPA = " & NroRep& & "", "NOMESS") > 0 Then
              RECEPEQ09.Label7 = "SI"
              RECEPEQ09.LBLSTATGARAN_ORIGINAL = 1
           End If
   End If
   '
10 RECEPEQ09.APRO = ""
   On Error Resume Next
   RECEPEQ09.Text2.SetFocus
   On Error GoTo 0
   If InStr(UCase$(EPAC$), "EYM") > 0 Then
      If OPX% = 2 Then RECEPEQ09.Text2.Enabled = False
      If NroRep& = 1 Then
        RECEPEQ09.Text2.Enabled = True
      End If
   End If
12 RECEPEQ09.Show 1
   If RECEPEQ09.APRO <> "NEXT" Then Exit Sub
   '
   If OPX% = 1 Then
      ' validacion de fecha de entrega solicitada solo en primera entrada.
      FESOLENTRE = FECHANUM(RECEPEQ09.Label14.Caption)
      If FESOLENTRE < HOY(HOS$) Then
         Call MENSERR(24, "La Fecha Solicitada de Entrega\Debe ser POSTERIOR a la Fecha Actual.")
         GoTo 12
      End If
   End If
   '
   RECACCE09.Label14 = RECEPEQ09.Text2
   RECACCE09.Label15 = RECEPEQ09.Label15
   RECACCE09.Label9 = RECEPEQ09.Label9
   RECACCE09.Label7 = RECEPEQ09.Label7
   RECACCE09.Label10 = RECEPEQ09.Label10
   RECACCE09.APRO = ""
   '
   If OPX% = 2 Then   ' MODIFICACION DE DATOS
       '
       RECACCE09.Text1 = VALOBADA("REPARA", "OBSERVA", "NUMREPA = " & NroRep& & "", "NOMESS")
       '
       ' tipo de falla
       SQLX$ = "SELECT * FROM TIPOFALL WHERE NUMREPA = " & NroRep& & " AND ORIGEN = 1"
       Dim FASEL As ADODB.Recordset
       Set FASEL = New ADODB.Recordset
       Set FASEL = FLEXCONN.Execute(SQLX$)
       With FASEL
          Do While Not .EOF
             NUFAL& = XVALO(!CODIFALL)
             DEFAL$ = SAFETEXT$(!DESCFALL)
             For KKI% = 1 To RECACCE09.List1.ListCount
                If XVALO(Left$(RECACCE09.List1.List(KKI% - 1), 4)) = NUFAL Then
                   RECACCE09.List1.Selected(KKI% - 1) = True
                   GoTo 14
                End If
             Next KKI%
             '
             If RECACCE09.Text2 <> "" Then RECACCE09.Text2 = RECACCE09.Text2 + vbCrLf
             RECACCE09.Text2 = RECACCE09.Text2 + DEFAL$
14        .MoveNext
          Loop
       End With
       FASEL.Close
       Set FASEL = Nothing
       '
       ' accesorios
       SQLX$ = "SELECT * FROM ACCEREP WHERE NUMREPA = " & NroRep&
       Dim FACCE As ADODB.Recordset
       Set FACCE = New ADODB.Recordset
       Set FACCE = FLEXCONN.Execute(SQLX$)
       With FACCE
          Do While Not .EOF
             CANACC& = XVALO(!Cantidad)
             If CANACC& > 0 Then
                DEACC$ = TRIM$(SAFETEXT$(!DESCRIACC))
                OBSEX$ = TRIM$(SAFETEXT$(!Observa))
                For KKI% = 2 To RECACCE09.GRID.Rows
                   TGRID$ = TRIM$(RECACCE09.GRID.TextMatrix(KKI% - 1, 0))
                   If UCase$(TGRID$) = UCase$(DEACC$) Or TGRID$ = "" Then
                      RECACCE09.GRID.TextMatrix(KKI% - 1, 0) = DEACC$
                      RECACCE09.GRID.TextMatrix(KKI% - 1, 1) = CANACC&
                      RECACCE09.GRID.TextMatrix(KKI% - 1, 2) = OBSEX$
                      GoTo 24
                   End If
                Next KKI%
                '
                RECACCE09.GRID.Rows = RECACCE09.GRID.Rows + 1
                KKI% = RECACCE09.GRID.Rows
                RECACCE09.GRID.TextMatrix(KKI% - 1, 0) = DEACC$
                RECACCE09.GRID.TextMatrix(KKI% - 1, 1) = CANACC&
                RECACCE09.GRID.TextMatrix(KKI% - 1, 2) = OBSEX$
             End If
24        .MoveNext
          Loop
       End With
       FACCE.Close
       Set FACCE = Nothing
   End If
   '
   RECACCE09.Show 1
   If RECACCE09.APRO = "PREV" Then GoTo 10
   If RECACCE09.APRO <> "NEXT" Then Exit Sub
   '
   NroRep& = XVALO(RECEPEQ09.Label15)
' AQUI VERIFICAR QUE ES MODIFICACION O NUMERO NUEVO
   Ci% = CODINT%(RECEPEQ09.Label9)
   NUMESER$ = TRIM$(RECEPEQ09.Text2)
   NCXX& = XVALO(RECEPEQ09.Text1)
   TTXX$ = "Confirma Recepción de Equipo\'" + TRIM$(CODEXT$(Ci%)) + " - " + DESCRIT0$(Ci%) + "'\Serie Numero '" + NUMESER$ + "'\para Presupuesto de Reparación ?\\Si, Recibir\No, Cancelar"
   If COMALTER%(TTXX$) <> 1 Then GoTo 10
   '
   Screen.MousePointer = 11
   DESKRIPCION$ = RECEPEQ09.Label10
   Call CREATABLA_OFABRIC
   CONDI$ = "NUMREPA = " & NroRep&
   If CantRegistros("REPARA", CONDI$) < 1 Then
      Call CreaRegistro("REPARA", "NUMREPA", NroRep&, "NOMESS")
   End If
   '
20 RACLI$ = TRIM$(RECEPEQ09.Text3)
   REFER$ = TRIM$(Left$(DESKRIPCION$, 22)) & " - " & RACLI$
   '
   Screen.MousePointer = 11
   
''''   Call ACTUREGISTRO("REPARA", "COD_INTE", CONDI$, CI%, REGAF&, "NOMESS")
''''   Call ACTUREGISTRO("REPARA", "DESCRIPCION", CONDI$, DESKRIPCION$, REGAF&, "NOMESS")
''''   Call ACTUREGISTRO("REPARA", "NumeroSerie", CONDI$, NUMESER$, REGAF&, "NOMESS")
''''      ESTREP% = 0: If RECEPEQ09.Label7 = "SI" Then ESTREP% = 1
''''   Call ACTUREGISTRO("REPARA", "Status_Gar", CONDI$, ESTREP%, REGAF&, "NOMESS")
''''   '
''''   Call ACTUREGISTRO("REPARA", "REFEREN", CONDI$, REFER$, REGAF&, "NOMESS")
''''   Call ACTUREGISTRO("REPARA", "STATUS", CONDI$, "0", REGAF&, "NOMESS")
''''   Call ACTUREGISTRO("REPARA", "FechRecep", CONDI$, FETEXCOR1$(FECHANUM(Left$(RECEPEQ09.Label16, 8))), RAFE&, "NOMESS")
''''   Call ACTUREGISTRO("REPARA", "SecRepar", CONDI$, RECEPEQ09.Text11, RAFE&, "NOMESS")
''''   Call ACTUREGISTRO("REPARA", "Usu_Recep", CONDI$, RECEPEQ09.Label17, RAFE&, "NOMESS")
''''   Call ACTUREGISTRO("REPARA", "FechEnSol", CONDI$, FETEXCOR1$(FECHANUM(RECEPEQ09.Label14.Caption)), RAFE&, "NOMESS")
''''   Call ACTUREGISTRO("REPARA", "Nume_Cli", CONDI$, NCXX&, REGAF&, "NOMESS")
''''   Call ACTUREGISTRO("REPARA", "Raso_Cli", CONDI$, RECEPEQ09.Text3, REGAF&, "NOMESS")
''''   Call ACTUREGISTRO("REPARA", "Domi_Cli", CONDI$, RECEPEQ09.Text4, REGAF&, "NOMESS")
''''   Call ACTUREGISTRO("REPARA", "Loca_Cli", CONDI$, RECEPEQ09.Text6, REGAF&, "NOMESS")
'''''     Call ACTUREGISTRO("REPARA", "Pais_Cli", CONDI$, RECEPEQ09.Text3, REGAF&, "NOMESS")
''''   Call ACTUREGISTRO("REPARA", "Tele_Cli", CONDI$, RECEPEQ09.Text7, REGAF&, "NOMESS")
''''   Call ACTUREGISTRO("REPARA", "Mail_Cli", CONDI$, RECEPEQ09.Text10, REGAF&, "NOMESS")
''''   Call ACTUREGISTRO("REPARA", "Ctac_Cli", CONDI$, RECEPEQ09.Text9, REGAF&, "NOMESS")
''''
'''''     Call ACTUREGISTRO("REPARA", "Accesorios", 12, 0)
''''   Call ACTUREGISTRO("REPARA", "Observa", CONDI$, RECACCE09.Text1.TEXT, REGAF&, "NOMESS")
''''   Call ACTUREGISTRO("REPARA", "Cuit_Cli", CONDI$, CUITCLI$(NCXX&), REGAF&, "NOMESS")
''''   Call ACTUREGISTRO("REPARA", "Piva_Cli", CONDI$, PIVACLI%(NCXX&), REGAF&, "NOMESS")
'      Call ACTUREGISTRO("REPARA", "Status_Gar", 3, 0)
'      Call ACTUREGISTRO("REPARA", "TipoFalla",  3, 0)
'      Call ACTUREGISTRO("REPARA", "DeTipFalla", 10, 128)
'      Call ACTUREGISTRO("REPARA", "Referen2", 10, 64)
   '
   SQLX$ = "SELECT * FROM REPARA "
   SQLX$ = SQLX$ + " WHERE NUMREPA = " & NroRep&
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
25 On Error Resume Next
   rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 25
   End If
   On Error GoTo 0

   With rs
      Do While Not .EOF
       !COD_INTE = Ci%
       !Descripcion = Mid$(DESKRIPCION$, 1, 128)
       !NUMEROSERIE = Mid$(NUMESER$, 1, 32)
        ESTREP% = 0: If RECEPEQ09.Label7 = "SI" Then ESTREP% = 1
       !STATUS_GAR = ESTREP%
       !Referen = Mid$(REFER$, 1, 128)
       !Status = 0
       !FECHRECEP = CVDAT(FECHANUM(Left$(RECEPEQ09.Label16, 8)))
       !SecRepar = Mid$(RECEPEQ09.Text11, 1, 48)
       !Usu_Recep = Mid$(RECEPEQ09.Label17, 1, 48)
       !FechEnSol = CVDAT(FECHANUM(RECEPEQ09.Label14.Caption))
       !Nume_Cli = NCXX&
       !Raso_Cli = Mid$(RECEPEQ09.Text3, 1, 64)
       !Domi_Cli = Mid$(RECEPEQ09.Text4, 1, 64)
       !Loca_Cli = Mid$(RECEPEQ09.Text6, 1, 64)
       !TELE_CLI = Mid$(RECEPEQ09.Text7, 1, 24)
       !Mail_Cli = Mid$(RECEPEQ09.Text10, 1, 48)
       !Ctac_Cli = Mid$(RECEPEQ09.Text9, 1, 48)
       !Observa = RECACCE09.Text1.TEXT
       !Cuit_Cli = Mid$(CuitCli$(NCXX&), 1, 24)
       !Piva_Cli = PIVACLI%(NCXX&)
       !TEC_EXT = XVALO(RECEPEQ09.Text13.TEXT)
       !TALON_TEC = TRIM$(RECEPEQ09.Text14.TEXT)
       .Update
       .MoveNext
      Loop
   End With
   '
   On Error Resume Next
   rs.Close
   Set rs = Nothing
   On Error GoTo 0

   ' DESCRIPCION DE LOS FALLOS
   DELARFA$ = ""
   Call BORRAREGISTROS("TIPOFALL", CONDI$, RAFE&, "NOMESS")
'   For KKI% = 1 To RECACCE09.List1.ListCount
'      If RECACCE09.List1.Selected(KKI% - 1) = True Then
'         If DELARFA$ <> "" Then DELARFA$ = DELARFA$ + vbCrLf
'         DELARFA$ = DELARFA$ + TRIM$(RECACCE09.List1.List(KKI% - 1))
'
'28       RNDX& = XVALO(RANDSTRING$(-6))
'         CONDI1$ = "NUMREPA = " & RNDX&
'         If CantRegistros("TIPOFALL", CONDI1$, "NOMESS") > 0 Then GoTo 28
'         Call CreaRegistro("TIPOFALL", "NUMREPA", RNDX&, "NOMESS")
'         Call ACTUREGISTRO("TIPOFALL", "ORIGEN", CONDI1$, 1, RAFE&, "NOMESS")  ' 1 CLIENTE; 2 TECNICO
'         Call ACTUREGISTRO("TIPOFALL", "CODIFALL", CONDI1$, XVALO(Left$(RECACCE09.List1.List(KKI% - 1), 4)), RAFE&, "NOMESS")
'         Call ACTUREGISTRO("TIPOFALL", "DESCFALL", CONDI1$, TRIM$(Mid$(RECACCE09.List1.List(KKI% - 1), 6)), RAFE&, "NOMESS")
'         Call ACTUREGISTRO("TIPOFALL", "NUMREPA", CONDI1$, NROREP&, RAFE&, "NOMESS")
'      End If
'   Next KKI%
   '
   SQLX$ = "SELECT * FROM TIPOFALL "
   SQLX$ = SQLX$ + " WHERE  NUMREPA = " & NroRep&
   Dim RS1 As ADODB.Recordset
   Set RS1 = New ADODB.Recordset
26 On Error Resume Next
   RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 26
   End If
   On Error GoTo 0

   With RS1
    For KKI% = 1 To RECACCE09.List1.ListCount
       If RECACCE09.List1.Selected(KKI% - 1) = True Then
        If DELARFA$ <> "" Then DELARFA$ = DELARFA$ + vbCrLf
        DELARFA$ = DELARFA$ + TRIM$(RECACCE09.List1.List(KKI% - 1))
        
        .AddNew
        !numrepa = NroRep&
        !Origen = 1
        !CODIFALL = XVALO(Left$(RECACCE09.List1.List(KKI% - 1), 4))
        !DESCFALL = Mid$(TRIM$(Mid$(RECACCE09.List1.List(KKI% - 1), 6)), 1, 128)
        .Update
        
       End If
    Next KKI%
   End With
   On Error Resume Next
   RS1.Close
   Set RS1 = Nothing
   On Error GoTo 0
   Screen.MousePointer = 11
   '
'   If TRIM$(RECACCE09.Text2) <> "" Then
'       If DELARFA$ <> "" Then DELARFA$ = DELARFA$ + vbCrLf
'       DELARFA$ = DELARFA$ + TRIM$(RECACCE09.Text2)
'       '
'38     RNDX& = XVALO(RANDSTRING$(-6))
'       CONDI1$ = "NUMREPA = " & RNDX&
'       If CantRegistros("TIPOFALL", CONDI1$, "NOMESS") > 0 Then GoTo 38
'       Call CreaRegistro("TIPOFALL", "NUMREPA", RNDX&, "NOMESS")
'       Call ACTUREGISTRO("TIPOFALL", "ORIGEN", CONDI1$, 1, RAFE&, "NOMESS")  ' 1 CLIENTE; 2 TECNICO
'       Call ACTUREGISTRO("TIPOFALL", "CODIFALL", CONDI1$, 999, RAFE&, "NOMESS")
'       Call ACTUREGISTRO("TIPOFALL", "DESCFALL", CONDI1$, RECACCE09.Text2, RAFE&, "NOMESS")
'       Call ACTUREGISTRO("TIPOFALL", "NUMREPA", CONDI1$, NROREP&, RAFE&, "NOMESS")
'   End If
   SQLX$ = "SELECT * FROM TIPOFALL "
   SQLX$ = SQLX$ + " WHERE  NUMREPA = " & NroRep&
   Dim RS2 As ADODB.Recordset
   Set RS2 = New ADODB.Recordset
27 On Error Resume Next
   RS2.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 27
   End If
   On Error GoTo 0

   With RS2
       If TRIM$(RECACCE09.Text2) <> "" Then
        If DELARFA$ <> "" Then DELARFA$ = DELARFA$ + vbCrLf
        DELARFA$ = DELARFA$ + TRIM$(RECACCE09.Text2)
        .AddNew
        !numrepa = NroRep&
        !Origen = 1
        !CODIFALL = 999
        !DESCFALL = Mid$(RECACCE09.Text2, 1, 128)
        .Update
        
       End If
   End With
   On Error Resume Next
   RS2.Close
   Set RS2 = Nothing
   On Error GoTo 0

   Call ACTUREGISTRO("REPARA", "Defec_Infor", CONDI$, DELARFA$, RAFE&, "NOMESS")
   '
   Screen.MousePointer = 11
   '
   ' ACCESORIOS
   ACCESOR$ = ""
   Call BORRAREGISTROS("ACCEREP", CONDI$, RAFE&, "NOMESS")
'   For KKI% = 2 To RECACCE09.GRID.Rows
'      CANTU& = XVALO(RECACCE09.GRID.TextMatrix(KKI% - 1, 1))
'      If CANTU& > 0 Then
'         If ACCESOR$ <> "" Then ACCESOR$ = ACCESOR$ + vbCrLf
'         ACCESOR$ = ACCESOR$ + TRIM$(RECACCE09.GRID.TextMatrix(KKI% - 1, 0)) + " (" + TRIM$(Str$(CANTU&)) + ")"
'         If TRIM$(RECACCE09.GRID.TextMatrix(KKI% - 1, 2)) <> "" Then
'            ACCESOR$ = ACCESOR$ + " - " + TRIM$(RECACCE09.GRID.TextMatrix(KKI% - 1, 2))
'         End If
'         '
'48       RNDX& = XVALO(RANDSTRING$(-6))
'         CONDI1$ = "NUMREPA = " & RNDX&
'         If CantRegistros("ACCEREP", CONDI1$, "NOMESS") > 0 Then GoTo 48
'         Call CreaRegistro("ACCEREP", "NUMREPA", RNDX&, "NOMESS")
'            DESCRIACC$ = RECACCE09.GRID.TextMatrix(KKI% - 1, 0)
'         Call ACTUREGISTRO("ACCEREP", "DESCRIACC", CONDI1$, DESCRIACC$, RAFE&, "NOMESS")
'         Call ACTUREGISTRO("ACCEREP", "CANTIDAD", CONDI1$, CANTU&, RAFE&, "NOMESS")
'            OBSEX$ = RECACCE09.GRID.TextMatrix(KKI% - 1, 2)
'         Call ACTUREGISTRO("ACCEREP", "OBSERVA", CONDI1$, OBSEX$, RAFE&, "NOMESS")
'         Call ACTUREGISTRO("ACCEREP", "NUMREPA", CONDI1$, NROREP&, RAFE&, "NOMESS")
'      End If
'   Next KKI%
   SQLX$ = "SELECT * FROM ACCEREP "
   SQLX$ = SQLX$ + " WHERE  NUMREPA = " & NroRep&
   Dim RS3 As ADODB.Recordset
   Set RS3 = New ADODB.Recordset
30 On Error Resume Next
   RS3.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 30
   End If
   On Error GoTo 0

   With RS3
   For KKI% = 2 To RECACCE09.GRID.Rows
       CANTU& = XVALO(RECACCE09.GRID.TextMatrix(KKI% - 1, 1))
       If CANTU& > 0 Then
        If ACCESOR$ <> "" Then ACCESOR$ = ACCESOR$ + vbCrLf
        ACCESOR$ = ACCESOR$ + TRIM$(RECACCE09.GRID.TextMatrix(KKI% - 1, 0)) + " (" + TRIM$(Str$(CANTU&)) + ")"
        If TRIM$(RECACCE09.GRID.TextMatrix(KKI% - 1, 2)) <> "" Then
            ACCESOR$ = ACCESOR$ + " - " + TRIM$(RECACCE09.GRID.TextMatrix(KKI% - 1, 2))
         End If
        .AddNew
           DESCRIACC$ = RECACCE09.GRID.TextMatrix(KKI% - 1, 0)
        !numrepa = NroRep&
        !DESCRIACC = Mid$(DESCRIACC$, 1, 1024)
        !Cantidad = CANTU&
           OBSEX$ = RECACCE09.GRID.TextMatrix(KKI% - 1, 2)
        !Observa = Mid$(OBSEX$, 1, 512)
        .Update
        
       End If
    Next KKI%
   End With
   On Error Resume Next
   RS3.Close
   Set RS3 = Nothing
   On Error GoTo 0

   Call ACTUREGISTRO("REPARA", "Accesorios", CONDI$, ACCESOR$, RAFE&, "NOMESS")
   '
   '
   Screen.MousePointer = 11
   '
   CUERPO$ = "Recepción del Equipo para Reparación" & vbCrLf
   If OPX% = 2 Then ' CORREGIR DATOS
      CUERPO$ = "Modificacion de Recepción de Equipo para Reparación" & vbCrLf
   End If
   CUERPO$ = CUERPO$ + vbCrLf + "N-Rep: " & NroRep& & " - Cliente: " & RACLI$ + vbCrLf
   '
   If DELARFA$ <> "" Then
      CUERPO$ = CUERPO$ + vbCrLf
      CUERPO$ = CUERPO$ + "Fallas Informadas por Cliente:" + vbCrLf
      CUERPO$ = CUERPO$ + DELARFA$ + vbCrLf
   End If
   '
   If ACCESOR$ <> "" Then
      CUERPO$ = CUERPO$ + vbCrLf
      CUERPO$ = CUERPO$ + "Accesorios Recibidos:" + vbCrLf
      CUERPO$ = CUERPO$ + ACCESOR$ + vbCrLf
   End If
   '
   If TRIM$(RECACCE09.Text1) <> "" Then
      CUERPO$ = CUERPO$ + vbCrLf
      CUERPO$ = CUERPO$ + "Observaciones:" + vbCrLf
      CUERPO$ = CUERPO$ + RECACCE09.Text1 + vbCrLf
   End If
   '
   'SE ESCRIBE SI SE CAMBIO EL ESTADO DE LA GARANTIA
   If XVALO(RECEPEQ09.LBLSTATGARAN_ORIGINAL) <> ESTREP% Then
      GARAN_ORIGINAL$ = "No Estar en Garantía"
      If XVALO(RECEPEQ09.LBLSTATGARAN_ORIGINAL) > 0 Then GARAN_ORIGINAL$ = "Sí Estar En Garantía"
      GARAN_MODIFICADA$ = "Sí Estar En Garantía"
      If ESTREP% < 1 Then GARAN_MODIFICADA$ = "No Estar en Garantía"
      CUERPO$ = CUERPO$ + vbCrLf
      CUERPO$ = CUERPO$ + "Se Modificó el Estado de Garantía" + vbCrLf
      CUERPO$ = CUERPO$ + "De: " + GARAN_ORIGINAL$ + vbCrLf
      CUERPO$ = CUERPO$ + "Modificado a: " + GARAN_MODIFICADA$ + vbCrLf
   End If
   
   Call ANOTANUSE(NUMESER$, CUERPO$, , "NOMESS")
   '
   Screen.MousePointer = 1
   If OPX% = 2 Then GoTo 15
   '
   Call COMUNI("Equipo Ingresado Para Presupuestación")
   '
   FREP = HOY(HOS$)
   '
   ARCSQR$ = TRIM$(CONTROL$("FOREPAR", "DOCUREP" + "5"))
   If ARCSQR$ <> "" Then
      'PARA EYM
      FILTX$ = TRIM$(Str$(NroRep&)) & ";" & TRIM$(Str$(FREP)) & ";" & TRIM$(Str$(NC))
      Call STDFORM(ARCSQR$, FILTX$)
      
      'PARA EL CLIENTE
      ARCSQR$ = TRIM$(CONTROL$("FOREPAR", "DOCUREP" + "7"))
      FILTX$ = TRIM$(Str$(NroRep&)) & ";" & TRIM$(Str$(FREP)) & ";" & TRIM$(Str$(NC))
      Call STDFORM(ARCSQR$, FILTX$)
      Exit Sub
   End If
   '
   FILTX$ = TRIM$(Str$(NroRep&)) & ";" & TRIM$(Str$(FREP)) & ";" & TRIM$(Str$(NC))
   TEX& = XVALO(RECEPEQ09.Text13)
   If TEX& > 0 Then
            FILTX$ = FILTX$ & ";" & "-TECNICO EXTERNO-"
            Call STDFORM("NREP-MDT", FILTX$)
            Exit Sub
   Else
            FILTX$ = FILTX$ & ";" & "===================="
            Call STDFORM("NREP-MDT", FILTX$, "PRINT")
   End If

    '
End Sub

