VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form OPREMSTA07 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Remitos de Expedición - Opciones"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   4815
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   1470
      ScaleHeight     =   270
      ScaleWidth      =   270
      TabIndex        =   47
      Top             =   525
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Caption         =   "REIMPRESION DE COMPROBANTES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Index           =   1
      Left            =   120
      TabIndex        =   37
      Top             =   5930
      Width           =   3690
      Begin VB.PictureBox Picture12 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   720
         ScaleHeight     =   615
         ScaleWidth      =   2400
         TabIndex        =   39
         Top             =   315
         Width           =   2430
         Begin VB.CommandButton Command27 
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
            Left            =   1850
            Picture         =   "OPREMSTA07.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   50
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
            Height          =   630
            Left            =   0
            Picture         =   "OPREMSTA07.frx":066A
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Consulta y Re-Impresión de Remitos"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Consulta y Re-Impresión"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   600
            TabIndex        =   40
            Top             =   75
            Width           =   1245
         End
      End
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
         Left            =   760
         Picture         =   "OPREMSTA07.frx":0CD4
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Consulta y Re-Impresión de Remitos"
         Top             =   350
         Width           =   560
      End
   End
   Begin VB.CommandButton Command11 
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
      Height          =   630
      Left            =   210
      Picture         =   "OPREMSTA07.frx":133E
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Indice de Cumplimiento de Entregas"
      Top             =   6930
      Visible         =   0   'False
      Width           =   650
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   7635
      Left            =   3960
      ScaleHeight     =   7605
      ScaleWidth      =   4245
      TabIndex        =   21
      Top             =   -360
      Width           =   4275
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   100
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   51
         Top             =   7080
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   52
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Tranf."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   100
         TabIndex        =   46
         Top             =   5880
         Visible         =   0   'False
         Width           =   640
      End
      Begin VB.CommandButton Command22 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Code"
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
         Left            =   100
         Picture         =   "OPREMSTA07.frx":1C08
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Control de Despacho con Codigo de Barras"
         Top             =   5130
         Width           =   640
      End
      Begin VB.CommandButton Command20 
         Caption         =   "Cump"
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
         Left            =   100
         Picture         =   "OPREMSTA07.frx":1F12
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Indice de Cumplimiento de Entregas"
         Top             =   4440
         Width           =   650
      End
      Begin VB.CommandButton Command18 
         Caption         =   "P.List"
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
         Left            =   100
         Picture         =   "OPREMSTA07.frx":27DC
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Lista de Empaque"
         Top             =   3840
         Visible         =   0   'False
         Width           =   650
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
         Left            =   100
         Picture         =   "OPREMSTA07.frx":2926
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   3260
         Width           =   650
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Recup. F.15359"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10
         TabIndex        =   36
         ToolTipText     =   "Recuperacion de Factura 15342"
         Top             =   4920
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Check"
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
         Left            =   100
         Picture         =   "OPREMSTA07.frx":2C30
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Control de Remitos Anulados"
         Top             =   2580
         Width           =   650
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Import"
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
         Picture         =   "OPREMSTA07.frx":2F3A
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Importación de Remitos"
         Top             =   2580
         Width           =   650
      End
      Begin VB.CommandButton Command1 
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
         Left            =   100
         Picture         =   "OPREMSTA07.frx":3084
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Configuracion de Funcionamiento"
         Top             =   6300
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
         Height          =   690
         Left            =   100
         Picture         =   "OPREMSTA07.frx":34C6
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   1190
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
         Left            =   100
         Picture         =   "OPREMSTA07.frx":37D0
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   500
         Width           =   650
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Equiv"
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
         Left            =   100
         Picture         =   "OPREMSTA07.frx":391A
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Códigos Equivalentes de Clientes"
         Top             =   1890
         Width           =   650
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Caption         =   "RESUMENES DE ENTREGAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3320
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   3690
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   2880
         Top             =   480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.PictureBox Picture9 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   980
         ScaleHeight     =   615
         ScaleWidth      =   2400
         TabIndex        =   33
         Top             =   2520
         Width           =   2430
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
            Height          =   630
            Left            =   0
            Picture         =   "OPREMSTA07.frx":3A64
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Consulta Detalle por Numero de  Remito"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Consulta Detalle por Nro. Remito"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   735
            TabIndex        =   35
            Top             =   105
            Width           =   1605
         End
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   840
         ScaleHeight     =   615
         ScaleWidth      =   2400
         TabIndex        =   15
         Top             =   1995
         Width           =   2430
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
            Height          =   630
            Left            =   0
            Picture         =   "OPREMSTA07.frx":3BAE
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Consulta / Listado de Material en Consignación (Clientes)"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Entregas Pendientes de Facturar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   735
            TabIndex        =   17
            Top             =   55
            Width           =   1605
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   640
         ScaleHeight     =   615
         ScaleWidth      =   2400
         TabIndex        =   11
         Top             =   1470
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
            Picture         =   "OPREMSTA07.frx":3EB8
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Listado Secuencial por Número de Comprobante"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Secuencial por Nro. de Comprobante"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   630
            TabIndex        =   13
            Top             =   55
            Width           =   1750
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   470
         ScaleHeight     =   615
         ScaleWidth      =   2400
         TabIndex        =   5
         Top             =   945
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
            Picture         =   "OPREMSTA07.frx":41C2
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Listado de Entregas por Cliente "
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Resumen de Entregas por Cliente"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   630
            TabIndex        =   7
            Top             =   55
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   300
         ScaleHeight     =   615
         ScaleWidth      =   2400
         TabIndex        =   8
         Top             =   420
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
            Picture         =   "OPREMSTA07.frx":4604
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Listado de Entregas por Artículo "
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Resumen de Entregas por Producto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   630
            TabIndex        =   10
            Top             =   55
            Width           =   1725
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Caption         =   "OPERACIONES COMPLEMENTARIAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3690
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   735
         ScaleHeight     =   615
         ScaleWidth      =   2610
         TabIndex        =   27
         Top             =   1470
         Width           =   2640
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
            Left            =   2050
            Picture         =   "OPREMSTA07.frx":490E
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Re-Asignación de Remitos a Pedidos"
            Top             =   0
            Width           =   560
         End
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
            Picture         =   "OPREMSTA07.frx":4A58
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Corrección Cantidades Despachadas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Corrección de Remitos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   720
            TabIndex        =   29
            Top             =   105
            Width           =   1005
         End
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   525
         ScaleHeight     =   615
         ScaleWidth      =   2610
         TabIndex        =   18
         Top             =   945
         Width           =   2640
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
            Left            =   2050
            Picture         =   "OPREMSTA07.frx":4D62
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Archivo Transferencia Mercedes Benz"
            Top             =   0
            Visible         =   0   'False
            Width           =   560
         End
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
            Picture         =   "OPREMSTA07.frx":5234
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Devolución Material en Consignación"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Devolución de Despachos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   720
            TabIndex        =   19
            Top             =   45
            Width           =   1095
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   315
         ScaleHeight     =   615
         ScaleWidth      =   2625
         TabIndex        =   1
         Top             =   420
         Width           =   2655
         Begin VB.CommandButton Command6 
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
            Left            =   550
            Picture         =   "OPREMSTA07.frx":537E
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Eliminación Solo de Lista de Pendientes de Factura"
            Top             =   0
            Width           =   570
         End
         Begin VB.CommandButton Command25 
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
            Picture         =   "OPREMSTA07.frx":5688
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Anulación General de Remito / Orden de Despacho Emitidos"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Anulación de Comprobantes"
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
            Left            =   1260
            TabIndex        =   3
            ToolTipText     =   "Click Aqui Importa Archivo GETRAMO"
            Top             =   45
            Width           =   1740
         End
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "OPREMSTA07.frx":5992
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "OPREMSTA07.frx":5BC6
      TabIndex        =   48
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuRegRemParaenvap 
         Caption         =   "Re-Generar Archivo Remito Conteplast"
      End
      Begin VB.Menu mnuVinculaPediRemi 
         Caption         =   "Vincular Remito c/Pedido (Solo Remitos sin Pedido Aplicado)"
      End
      Begin VB.Menu mnuPasarRemitosDropbox 
         Caption         =   "Pasar Remitos A DropBox"
         Visible         =   0   'False
      End
      Begin VB.Menu mnusalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnulistados 
      Caption         =   "Listados"
      Begin VB.Menu mnuResumenEntrexCliente 
         Caption         =   "Resumen de Entregas x Cliente"
         Begin VB.Menu mnuDetxCliente 
            Caption         =   "Detalle por Cliente"
         End
         Begin VB.Menu mnuAcumxCliente 
            Caption         =   "Acumulado por Cliente"
         End
      End
   End
End
Attribute VB_Name = "OPREMSTA07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub ACTUCOMPRASUCU(CI1%, CARECIORIGINAL, Remito&, NROPEDAVELLANEDA&)
          '
          
          'SEGUN LA NOTA DE VENTA QUE LEO DE AVELLANEDA BUSCO EL NRO DE PEDIDO DE REPOSICION
          OCOMPRA& = 0
          SQLX$ = "Select * from DetaOVenta WITH(NOLOCK)"
          SQLX$ = SQLX$ + " WHERE nroped = " & NROPEDAVELLANEDA&
          SQLX$ = SQLX$ + "  AND  CODINT = " & CI1%
          Set RS1 = READSET(SQLX$)
          With RS1
            If Not .EOF Then OCOMPRA& = XVALO(!nro_ocom)
          End With
          RS1.Close: Set RS1 = Nothing
          '
          'ACA DESCUENTA DE OCOMDETA LAS CANTIDADES RECEPCIONADA
          SQLX$ = "SELECT * FROM OCOMDETA"
          SQLX$ = SQLX$ + " WHERE COD_INTE = " & CI1%
          SQLX$ = SQLX$ + " AND NUME_OCOM = " & OCOMPRA&
          SQLX$ = SQLX$ + " AND Cant_Rec < (Cant_Ocom - 0.005) "
          SQLX$ = SQLX$ + " ORDER BY Fentre_Sol,NuOrd_Item "
          Dim rs As ADODB.Recordset
          Set rs = New ADODB.Recordset
          rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
          With rs
            Do While Not .EOF
              CI1% = XVALO(!cod_inte)
              COEFUNI = XVALO(!Coef_Unids)
                If COEFUNI < 0.0001 Then COEFUNI = 1
              CANTI_RECI_BD = XVALO(!Cant_Rec)
              CARECI = XVALO(CARECIORIGINAL / COEFUNI)
              as1 = CODEXT$(CI1%)
              CANTI_COMPRA = XVALO(!cant_ocom)
              AB = !Fentre_Sol
              SOBRAN = 0
              CARECITOTAL = CARECI + CANTI_RECI_BD
              If CARECITOTAL > CANTI_COMPRA Then
                 SOBRAN = CARECITOTAL - CANTI_COMPRA
                 CARECITOTAL = CANTI_COMPRA
              End If
              'SI SE COMPLETO LO COMPRADO SE PONE EN STATUS = 3
              If CARECITOTAL >= CANTI_COMPRA Then
                 Stat_Ocom = 3
              End If
              !Cant_Rec = CARECITOTAL
              .Update
              If SOBRAN > 0 Then
                 CARECIORIGINAL = SOBRAN ' SI NO SOBRA POR QUE APLICO TODO
              Else
                 CARECIORIGINAL = 0
              End If
              Call ANOTARECEPCION_TRAZAREC(CI1%, XVALO(!nume_ocom), Remito&, (CARECI - SOBRAN), NROPEDAVELLANEDA&)
              'SI YA TERMINO DE ASIGNAR LA CANTIDAD REMITIDA EN LAS O.C. SALE.
              If CARECIORIGINAL = 0 Then
                Exit Do
              End If
              .MoveNext
            Loop
          End With
          rs.Close
          Set rs = Nothing
          'IF HAY SOBRANTE O BIEN NO ENTRO POR LA CONSULTA PARA ASIGNAR INFORMA EN LA TABLA TRAZAREC LO RECIBIDO _
          VA A SER CON PEDIDO DE REPOSICION = 0 O PUEDE SER QUE HAYA ENCONTRADO LA RELACDION P.REP = N.VENTA _
          PERO QUE NO ENTRO EN LA CONSULTA.
          If CARECIORIGINAL > 0 Then Call ANOTARECEPCION_TRAZAREC(CI1%, OCOMPRA&, Remito&, CARECIORIGINAL, NROPEDAVELLANEDA&)
          

End Sub

Sub ANULALEMPAQ()

10  Call CARGALISEL("LEMPEXIS", "ENCEMPAQ", "NUMLTAEMP/I5;REFERENCIA/A36", "NUMLTAEMP > 0", "DISTINCT")
    Call SELECHO("LEMPEXIS/LISTA DE EMPAQUES")
    NUMLEMP& = XVALO(VALACT1$("LEMPEXIS"))
    If NUMLEMP& < 1 Then Exit Sub
    '
    'VALIDO SI LA SELECCIONADA NO ESTA CERRADA
    If LEMPCER%(NUMLEMP&) > 0 Then
       Call COMUNI("La Lista de Empaque Seleccionada Se Encuentra Cerrada\Para Anularse Deberá Revertir el Cierre de la Misma")
       Exit Sub
    End If
    '
    OPX% = COMALTER%("Confirmación de Anulación de Lista de Empaque\\Cancelar\Anular")
    '
    If OPX% = 2 Then 'ANULAR
        CONDI$ = " NumLtaEmp = " & NUMLEMP&
        Call BORRAREGISTROS("LISEMPAQ", CONDI$, REGAF&, "NOMESS")
        Call BORRAREGISTROS("ENCEMPAQ", CONDI$, REGAF&, "NOMESS")
    End If
    GoTo 10
    '
End Sub


Sub ArmaArchivoDeDespachos(SQLX$, nroenvio&, NCLIE)

    'sqlx$ = "select * from movisto with(nolock) where codimovi='rt' and doprilar like 'rt%' and fechmov>='" & Desde & "' and fechmov<='" & HASTA & "' and (nroenvio=0 or nroenvio is null) and nume_clie=" & NCLIE & " AND " & CONDI$ & " order by doprilar asc"
    Dim vectorcodigos() As String * 128
    Dim resumenentrega(9) As Long
    Dim remitosdespachados() As String
    '
    Set Rst = READSET(SQLX$)
    remitoant$ = ""
    ReDim Preserve vectorcodigos(2)
    contadorremito& = 0
    '
    ' *** registro encabezado
    II& = 0
    vectorcodigos(II&) = Space(128)
    Call REPLA(vectorcodigos(II&), "&&##AOC115620146#ADC0MB565#BOWESA#Z", 1, 40) ' !!!!! configurar numero de cliente // SE MODIFICA MB155 x MB565
    '
    ' *** registro 711
    II& = 1
    vectorcodigos(II&) = Space(128)
    Call REPLA(vectorcodigos(II&), "71103", 1, 5)
        'nrocliemb$ = NPROCLI$(NCLIE)
        'While Len(nrocliemb$) < 9: nrocliemb$ = "0" + nrocliemb$: Wend
        'nrocliefv$ = "3"
        'While Len(nrocliefv$) < 9: nrocliefv$ = "0" + nrocliefv$: Wend
    Call REPLA(vectorcodigos(II&), "MB565    15620146 ", 6, 18)                 ' // SE MODIFICA MB155 x MB565
    'Call REPLA(vectorcodigos(II&), nrocliefv$, 15, 9)
    If nroenvio& = 0 Then
        ultnroenvio$ = SAFETEXT$(XVALO(VALOBADA("movisto", "nroenvio", "nroenvio>0 order by nroenvio desc")))
        While Len(ultnroenvio$) < 5: ultnroenvio$ = "0" + ultnroenvio$: Wend
        proxnroenvio$ = CStr(XVALO(ultnroenvio$) + 1)
        While Len(proxnroenvio$) < 5: proxnroenvio$ = "0" + proxnroenvio$: Wend
    Else
        ultnroenvio$ = CStr(nroenvio& - 1): If nroenvio& = 0 Then ultnroenvio$ = "0"
        While Len(ultnroenvio$) < 5: ultnroenvio$ = "0" + ultnroenvio$: Wend
        proxnroenvio$ = CStr(nroenvio&)
        While Len(proxnroenvio$) < 5: proxnroenvio$ = "0" + proxnroenvio$: Wend
    End If
    Call REPLA(vectorcodigos(II&), ultnroenvio$, 24, 5)
    Call REPLA(vectorcodigos(II&), proxnroenvio$, 29, 5)
        If nroenvio& = 0 Then
            FECHA = FECHATEX$(HOY(HO$))
        Else
            FF% = CVINT(VALOBADA("movisto", "FHor_Verif", "nroenvio=" & CStr(nroenvio&)))
            FECHA = FECHATEX$(FF%)
        End If
        fechaenvio$ = Mid$(FECHA, 7, 2) & Mid$(FECHA, 4, 2) & Mid$(FECHA, 1, 2)
    Call REPLA(vectorcodigos(II&), fechaenvio$, 34, 6)
    '
    ' *** registro 712
    II& = 2
    vectorcodigos(II&) = Space(128)
    Call REPLA(vectorcodigos(II&), "71203", 1, 5)
    Call REPLA(vectorcodigos(II&), "565", 14, 3)                ' // SE MODIFICA 155 x 565
    '
    With Rst
        Do While Not .EOF
            '
            Remito$ = SAFETEXT$(!DOPRILAR)
            If Remito$ <> remitoant$ Then
                contadorremito& = contadorremito& + 1
                ReDim Preserve remitosdespachados(contadorremito&)
                remitosdespachados(contadorremito&) = Remito$
                '
                ' *** registro 715
'                If remitoant$ <> "" Then
'                    findetalle& = II&
'                    For IX& = iniciodetalle& To findetalle&
'                        II& = II& + 1
'                        ReDim Preserve vectorcodigos(II&)
'                        vectorcodigos(II&) = Space(128)
'                        Call REPLA(vectorcodigos(II&), "71503", 1, 5)
'                        Call REPLA(vectorcodigos(II&), "2032", 6, 4)
'                            codigoart$ = TRIM$(Mid$(vectorcodigos(IX&), 28, 22))
'                            CIXI% = CODINT%(codigoart$)
'                                qantitxt$ = Mid$(vectorcodigos(IX&), 53, 10) & "." & Mid$(vectorcodigos(IX&), 63, 3)
'                            qanti# = XVALO(qantitxt$)
'                            cantiporbulto& = LINDIVEN(CIXI%)
'                            totalbultos& = Int(qanti# / cantiporbulto&)
'                            RESTO# = qanti# - (cantiporbulto& * totalbultos&)
'                            If RESTO# > 0 Then
'                               totalbultos& = totalbultos& + 1
'                            End If
'                            totalbultostxt$ = CStr(totalbultos&)
'                            While Len(totalbultostxt$) < 13: totalbultostxt$ = "0" + totalbultostxt$: Wend
'                        Call REPLA(vectorcodigos(II&), totalbultostxt$, 50, 13)
'                     Next IX&
'                End If
                '
                ' *** registro 713
                II& = II& + 1
                ReDim Preserve vectorcodigos(II&)
                vectorcodigos(II&) = Space(128)
                    remitoant$ = SAFETEXT$(!DOPRILAR)
                Call REPLA(vectorcodigos(II&), "71303", 1, 5)
                Call REPLA(vectorcodigos(II&), Mid$(Remito$, 11, 8), 6, 8)
                    FECHA = FECHATEX$(CVINT%(!FECHMOV))
                    fechaenvio$ = Mid$(FECHA, 7, 2) & Mid$(FECHA, 4, 2) & Mid$(FECHA, 1, 2)
                Call REPLA(vectorcodigos(II&), fechaenvio$, 14, 6)
                Call REPLA(vectorcodigos(II&), "266", 20, 3)
                Call REPLA(vectorcodigos(II&), "20", 25, 2)
                iniciodetalle& = II& + 1
            End If
            '
            CIXI% = XVALO(!cod_inte)
            Codigo$ = TRIM$(CODEXT$(CIXI%))
            If Codigo$ <> "29-0020" And Codigo$ <> "29-0030" Then ' obvia a las capachos, contemplados en estos dos codigos
                ' *** registro 714
                II& = II& + 1
                ReDim Preserve vectorcodigos(II&)
                vectorcodigos(II&) = Space(128)
                Call REPLA(vectorcodigos(II&), "71403", 1, 5)
    
                Call REPLA(vectorcodigos(II&), CODORIG$(CIXI%, CInt(NCLIE)), 6, 22) ' hay que reemplazar por el de mercedes
                Call REPLA(vectorcodigos(II&), Codigo$, 28, 22)
                Call REPLA(vectorcodigos(II&), "999", 50, 3)
                    qanti# = XVALO(!CANTSALID)
                    cantitxt$ = TRIM$(FORMATNUM$(qanti#, "F13.3"))
                    cantitxt$ = REPLACAR(cantitxt$, ".", "")
                    While Len(cantitxt$) < 13: cantitxt$ = "0" + cantitxt$: Wend
                Call REPLA(vectorcodigos(II&), cantitxt$, 53, 13)
                    nrorden$ = CStr(XVALO(!NuOrItem))
                    While Len(nrorden$) < 3: nrorden$ = "0" + nrorden$: Wend
                Call REPLA(vectorcodigos(II&), "ST", 81, 2)
                Call REPLA(vectorcodigos(II&), nrorden$, 87, 3)
                Call REPLA(vectorcodigos(II&), "S", 106, 1)
                Call REPLA(vectorcodigos(II&), "X1", 115, 2)
                '
                ' *** registro 715 *** codigo que reemplaza lo que esta comentado
                II& = II& + 1
                ReDim Preserve vectorcodigos(II&)
                vectorcodigos(II&) = Space(128)
                Call REPLA(vectorcodigos(II&), "71503", 1, 5)
                Call REPLA(vectorcodigos(II&), "2032", 6, 4)
                    cantiporbulto& = LINDIVEN(CIXI%)
                    totalbultos& = Int(qanti# / cantiporbulto&)
                    Resto# = qanti# - (cantiporbulto& * totalbultos&)
                    If Resto# > 0 Then
                       totalbultos& = totalbultos& + 1
                    End If
                    totalbultostxt$ = CStr(totalbultos&)
                    While Len(totalbultostxt$) < 13: totalbultostxt$ = "0" + totalbultostxt$: Wend
                Call REPLA(vectorcodigos(II&), totalbultostxt$, 50, 13)
                '
            End If
            .MoveNext
        Loop
    End With
    '
    On Error Resume Next
    Rst.Close
    Set Rst = Nothing
    On Error GoTo 0
    '
'    ' *** registro 715
'    findetalle& = II&
'    For IX& = iniciodetalle& To findetalle&
'        II& = II& + 1
'        ReDim Preserve vectorcodigos(II&)
'        vectorcodigos(II&) = Space(128)
'        Call REPLA(vectorcodigos(II&), "71503", 1, 5)
'        Call REPLA(vectorcodigos(II&), "2032", 6, 4)
'            codigoart$ = TRIM$(Mid$(vectorcodigos(IX&), 28, 22))
'            CIXI% = CODINT%(codigoart$)
'                qantitxt$ = Mid$(vectorcodigos(IX&), 53, 10) & "." & Mid$(vectorcodigos(IX&), 63, 3)
'            qanti# = XVALO(qantitxt$)
'            cantiporbulto& = LINDIVEN(CIXI%)
'            totalbultos& = Int(qanti# / cantiporbulto&)
'            RESTO# = qanti# - (cantiporbulto& * totalbultos&)
'            If RESTO# > 0 Then
'               totalbultos& = totalbultos& + 1
'            End If
'            totalbultostxt$ = CStr(totalbultos&)
'            While Len(totalbultostxt$) < 13: totalbultostxt$ = "0" + totalbultostxt$: Wend
'        Call REPLA(vectorcodigos(II&), totalbultostxt$, 50, 13)
'     Next IX&
     '
     ' *** registro 719
     For IX& = 1 To II&
        codigoreg% = XVALO(Mid$(vectorcodigos(IX&), 1, 3))
        Select Case codigoreg%
            Case 711
                    resumenentrega(1) = resumenentrega(1) + 1
            Case 712
                    resumenentrega(2) = resumenentrega(2) + 1
            Case 713
                    resumenentrega(3) = resumenentrega(3) + 1
            Case 714
                    resumenentrega(4) = resumenentrega(4) + 1
            Case 715
                    resumenentrega(5) = resumenentrega(5) + 1
            Case 716
                    resumenentrega(6) = resumenentrega(6) + 1
            Case 717
                    resumenentrega(7) = resumenentrega(7) + 1
            Case 718
                    resumenentrega(8) = resumenentrega(8) + 1
        End Select
     Next IX&
     '
     II& = II& + 1
     ReDim Preserve vectorcodigos(II&)
     vectorcodigos(II&) = Space(128)
     Call REPLA(vectorcodigos(II&), "71902", 1, 5)
     For IX& = 1 To 8
        totalcodigotxt$ = CStr(resumenentrega(IX&))
        While Len(totalcodigotxt$) < 7: totalcodigotxt$ = "0" + totalcodigotxt$: Wend
        Call REPLA(vectorcodigos(II&), totalcodigotxt$, (IX& - 1) * 7 + 5, 7)
     Next IX&
     Call REPLA(vectorcodigos(II&), "0000001", (IX& - 1) * 7 + 5, 7)
     '
     ' *** pie de archivo
     II& = II& + 1
     ReDim Preserve vectorcodigos(II&)
     vectorcodigos(II&) = Space(128)
     Call REPLA(vectorcodigos(II&), "&&##Z", 1, 5)
     '
     For IX& = 0 To II&
        Debug.Print vectorcodigos(IX&)
     Next IX&
     '
     ' *** guarda archivo y actualiza movisto con el numero de envio
     OPX% = COMALTER%("¿Desea Generar el Archivo de Despachos?\\Si, Generar\No, Cancelar")
     If OPX% = 1 Then
        If nroenvio& = 0 Then ' no graba si se trata de un regeneracion de un envio preexistente
            For IX& = 1 To contadorremito&
                Remito$ = remitosdespachados(IX&)
                CONDI$ = "doprilar='" & Remito$ & "'"
                Call ACTUREGISTRO("movisto", "nroenvio", CONDI$, XVALO(proxnroenvio$), REG&)
                Call ACTUREGISTRO("movisto", "FHor_Verif", CONDI$, FECHATEX$(HOY(HO$)), REG&)
            Next IX&
        End If
        '
        CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A EXPORTA
        CommonDialog1.DialogTitle = "Archivo Origen de Datos"
        CommonDialog1.Filter = "Archivo de Texto(*.ASI)|*.ASI|"
        CommonDialog1.FilterIndex = 1
        CommonDialog1.FileName = "mercedezbenz_envio_" & CStr(proxnroenvio$) & ".asi"
        On Error GoTo 99
        CommonDialog1.ShowOpen
        On Error GoTo 0
        DoEvents
        NOMBASALID$ = UCase$(CommonDialog1.FileName)
        ASALID1$ = NOMBASALID$
        Screen.MousePointer = 11
        NX1% = FILEOPEN%(ASALID1$, 2, 128)
        For IX& = 0 To II&
            If IX& = II& Or IX& = 0 Then
                TX$ = TRIM$(vectorcodigos(IX&))
            Else
                TX$ = vectorcodigos(IX&)
            End If
            Print #NX1%, TX$
        Next IX&
        Close #NX1%
        ASALID2$ = REPLACAR(ASALID1$, "\", "/")
        Screen.MousePointer = 1
        Call COMUNI("Archivos de Transferencia\ '" + ASALID2$ + "' \Generado con Correctamente")
     End If
     '
99
     ' *** fin de rutina

End Sub

Sub CARPETADEINTERCAMBIO()

    If CLACONTRA% < 1 Then
        Call MENSERR(24, "Santo Y Seña Inválido")
        Exit Sub
    End If
    TXT$ = TRIM$(Control("", "CARPEINT"))
    TXT$ = TRIM$(InputBox$("Ruta", "Carpeta de Intercambio", TXT$))
    Call GRACONTROL("", "CARPEINT", TXT$)

End Sub

Sub CERRAEMPAQ()
    'PRESENTO LISTA DE EMPAQUES SIN CERRAR
10  CONDI$ = "NUMLTAEMP > 0 AND (EmpaCerra is null or EmpaCerra < 1)"
    Call CARGALISEL("CERRAEMPA", "ENCEMPAQ", "NUMLTAEMP/I5;REFERENCIA/A36", CONDI$, "DISTINCT")
    Call SELECHO("CERRAEMPA/Lista de Empaques Abiertas")
    '
    'VALIDO
    NUMLEMPA& = XVALO(VALACT1$("CERRAEMPA"))
    If NUMLEMPA& < 1 Then Exit Sub
    '
    'VALIDO QUE TENGA TODOS LOS DATOS DEL ARCHIVO MAESTRO
    SQLX$ = "SELECT CODINT FROM LISEMPAQ WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE NumLtaEmp = " & NUMLEMPA&
    SQLX$ = SQLX$ + " GROUP BY CODINT "
    Dim LIMPTMP As ADODB.Recordset
    Set LIMPTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With LIMPTMP
      Do While Not .EOF
        CI% = XVALO(!CODINT)
        If FALTANDATOS%(CI%) > 0 Then
          Call MENSERR(24, "Imposible Cerrar Lista de Empaque Nro. " & NUMLEMPA& & "\Faltan Datos en Archivo Maestro para Código: " & CODEXT$(CI%))
          LIMPTMP.Close
          GoTo 10
        End If
      .MoveNext
      Loop
    End With
    LIMPTMP.Close
    
    'GRABO LA MARCA DE CERRADA
    CONDI$ = " NUMLTAEMP = " & NUMLEMPA&
    Call ACTUREGISTRO("ENCEMPAQ", "EmpaCerra", CONDI$, 1, REGAF&, "NOMESS")
    '
    'PRESENTA MENSAJE
    Call COMUNI("Lista de Empaque Nro. " & NUMLEMPA& & " Marcada Como Cerrada ")
    GoTo 10
End Sub

Sub CREATABLA_TRAZAREC()
   Call ABRECONEXION
   Call CREACAMPO("", "TRAZAREC", "ID", 4, 0, 1)
   Call CREACAMPO("", "TRAZAREC", "NROREMI", 4, 0, 1)
   Call CREACAMPO("", "TRAZAREC", "Fech_Recep", 8, 0)
   Call CREACAMPO("", "TRAZAREC", "Nume_Ocom", 4, 4, 1)
   Call CREACAMPO("", "TRAZAREC", "Cant_Remi", 6, 4)
   Call CREACAMPO("", "TRAZAREC", "Cod_Inte", 3, 0, 1)
   Call CREACAMPO("", "TRAZAREC", "NroPed", 4, 0, 1)
End Sub
Sub ANOTARECEPCION_TRAZAREC(CI1%, NUMOC&, Remito&, CARECIORIGINAL, NROPEDAVELLANEDA&)
    SQLX$ = "SELECT  TOP 1 * FROM TRAZAREC "
    SQLX$ = SQLX$ + " ORDER BY ID DESC"
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With rs
      IDX& = 1 + XVALO(!ID)
      .AddNew
      !ID = IDX&
      !NROREMI = Remito&
      !Fech_Recep = CVDAT(HOY(HO$))
      !nume_ocom = NUMOC&
      !Cant_Remi = CARECIORIGINAL
      !cod_inte = CI1%
      !NROPED = NROPEDAVELLANEDA&
      .Update
    End With
    rs.Close
    Set rs = Nothing
End Sub


Function FALTANDATOS%(CI%)
    '1 Faltan Datos
    '0 DATOS CORRECTOS
    INDAT% = 0
    '
    'TOMO LOS DATOS DEL A.MAESTRO
    CONDI$ = "CODINT = " & CI%
    M3CAX = XVALO(VALOBADA("MASTER", "M3CAJA", CONDI$, "NOMESS"))
    KILNETCAJ = XVALO(VALOBADA("MASTER", "KILONETOCAJA", CONDI$, "NOMESS"))
    KILBRCAJ = XVALO(VALOBADA("MASTER", "KILOBRUTOCAJA", CONDI$, "NOMESS"))
    CANTIDXCAJA = XVALO(VALOBADA("MASTER", "PIEZASXCAJA", CONDI$, "NOMESS"))
    '
    'VALIDO
    If M3CAX < 0.005 Then INDAT% = 1: GoTo 99
    If KILNETCAJ < 0.005 Then INDAT% = 1: GoTo 99
    If KILBRCAJ < 0.005 Then INDAT% = 1: GoTo 99
    If CANTIDXCAJA < 0.005 Then INDAT% = 1: GoTo 99
    '
99  FALTANDATOS% = INDAT%
End Function

Sub FRMLIEMPAQ()

    'CARGO EL LISTADO DE SELECCION
10  CONDI$ = "NUMLTAEMP > 0 AND EmpaCerra > 0 "
    Call CARGALISEL("FRMEMPAQ", "ENCEMPAQ", "NUMLTAEMP/I5;REFERENCIA/A36", CONDI$, "DISTINCT")
    Call SELECHO("FRMEMPAQ/Lista de Empaques Cerradas")
    '
    NUMLEMPA& = XVALO(VALACT1$("FRMEMPAQ"))
    If NUMLEMPA& < 1 Then Exit Sub
    '
    CONDI$ = " NUMLTAEMP = " & NUMLEMPA&
    FREP = CVINT(VALOBADA("ENCEMPAQ", "FECHEMIS", CONDI$, "NOMESS"))
    '
    'VALIDO
    If XVALO(FREP) < 1 Then FREP = HOY(HO$)
    '
    'PRESENTO EL FRM
    FILTX$ = TRIM$(Str$(NUMLEMPA&)) & ";" & TRIM$(Str$(FREP))
    Call STDFORM("LISTEMP", FILTX$)
    GoTo 10
    '
End Sub

Function LEMPCER%(NUMLEMP&)
    LECER% = 0
    CONDI$ = "NUMLTAEMP = " & NUMLEMP& & " AND EmpaCerra > 0 "
    LEMPCERRA& = CantRegistros&("ENCEMPAQ", CONDI$)
    If LEMPCERRA& > 0 Then
       LECER% = 1
    End If
    LEMPCER% = LECER%
    

End Function

Sub LUGARENTREGA()

   Call SELECHO("DEUDOR12")
   NCLI = XVALO(VALACT1$("DEUDOR12"))
   If NCLI = 0 Then Exit Sub
   '
   NCXI% = NCLI
   If ECOARCH("DEUDOR12", MKS$(NCXI%)) = "" Then
      Call MENSERR(24, "Ingrese Número de Cliente")
      Exit Sub
   End If
   '
   RFF$ = RECFILT$("ALMAENT", 3, MKI$(NCXI%))
   JJ& = 0
   For U& = 1 To Len(RFF$) Step 4
     REGE& = CVS(Mid$(RFF$, U&, 4))
     X$ = REGLEIDO$("ALMAENT", REGE&)
     JJ& = JJ& + 1
     Call GRAREG("!ALMAENT", X$, JJ&)
   Next U&
   Call SETULTREG("!ALMAENT", JJ&)
   '
   VTB% = VENTABU%("ALMAENT")
   If VTB% >= 0 Then
     URN& = ULTREG&("ALMAENT")
     RNC$ = RECFILT$("ALMAENT", 3, MKI$(NCXI%))
     kk& = (-3)
     '
     For U& = 1 To ULTREG&("!ALMAENT")
      TTYY$ = AJUSTI$(REGLEIDO$("!ALMAENT", U&), 30) + MKI$(NCXI%)
      kk& = kk& + 4
      If kk& <= Len(RNC$) - 3 Then
           RENOTA& = CVS(Mid$(RNC$, kk&, 4))
         Else
           URN& = URN& + 1
           RENOTA& = URN&
      End If
      Call GRAREG("ALMAENT", TTYY$, RENOTA&)
     Next U&
   End If
   '
20 While kk& < Len(RNC$) - 3
      kk& = kk& + 4
      RENOTA& = CVS(Mid$(RNC$, kk&, 4))
      Call GRAREG("ALMAENT", Space$(30) + MKI$(0), RENOTA&)
   Wend
   '
   Call CIERRARCH("ALMAENT")
   
End Sub

Sub ORDENALISTAMP(NUMLISTAEMP&)
        '
        'MARCO LOS CAMPOS EXISTENETES
        CONDI$ = " NumLtaEmp = " & NUMLISTAEMP&
        Call ACTUREGISTRO("LISEMPAQ", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")
        '
        'ABRO UN RECORDSET PARA ESCRITURA DE LOS DATOS ORDENADOS
        sqlx1$ = "SELECT * FROM LISEMPAQ WHERE NumLtaEmp >1"
        Dim LISTUPDTE As ADODB.Recordset
        Set LISTUPDTE = New ADODB.Recordset
26      On Error Resume Next
        LISTUPDTE.Open FILTSQL$(sqlx1$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 26
        End If
        On Error GoTo 0
        '
        'ABRO OTRO RECORDSET DE LECTURA POR QUE NO PUEDO ORDENAR EL RECORDSET DE ESCRITURA
        NUMEORD% = 0
        SQLX$ = "SELECT * FROM LISEMPAQ WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE "
        SQLX$ = SQLX$ + " NumLtaEmp = " & NUMLISTAEMP&
        SQLX$ = SQLX$ + " ORDER BY NUMPALLET , KILOBRUTOCAJA DESC "
        Dim LIMPTMP1 As ADODB.Recordset
        Set LIMPTMP1 = FLEXCONN.Execute(FILTSQL$(SQLX$))
        With LIMPTMP1
          Do While Not .EOF
            NUMEPALE% = XVALO(!Numpallet)
            If NUMEPALE% <> NUMEPALEANT% Then
               NUMEORD% = 0
            End If
            NUMEORD% = NUMEORD% + 1
            NUMEPALEANT% = NUMEPALE%
            '
            'ACTUALIZO EL RECORDSET CON PERMISO DE ESCRITURA
            LISTUPDTE.AddNew
            
            On Error Resume Next
            LISTUPDTE!CODIEMPR = !CODIEMPR
            On Error GoTo 0

            LISTUPDTE!NumLtaEmp = !NumLtaEmp
            LISTUPDTE!Numpallet = !Numpallet
            LISTUPDTE!CODINT = !CODINT
            '
            LISTUPDTE!DESDECAJA = !DESDECAJA
            LISTUPDTE!HASTACAJA = !HASTACAJA
            '
            LISTUPDTE!CANTICAJAS = !CANTICAJAS
            LISTUPDTE!PiezasXCaja = !PiezasXCaja
            LISTUPDTE!PIEZASTOTAL = !PIEZASTOTAL
            LISTUPDTE!M3TOTAL = !M3TOTAL
            LISTUPDTE!M3Caja = !M3Caja
            '
            LISTUPDTE!KILONETOCAJA = !KILONETOCAJA
            LISTUPDTE!KILONETOTOT = !KILONETOTOT
            LISTUPDTE!KILOBRUTOCAJA = !KILOBRUTOCAJA
            LISTUPDTE!KILOBRUTOTOT = !KILOBRUTOTOT
            LISTUPDTE!NROPED = !NROPED
            '
            LISTUPDTE!NUORDPALLET = NUMEORD%
            LISTUPDTE!MARBORRA = 0
            LISTUPDTE.Update
            .MoveNext
          Loop
        End With
        LISTUPDTE.Close
        Set LISTUPDTE = Nothing
        '
        LIMPTMP1.Close
        Set LIMPTMP1 = Nothing
        '
        'BORRO REGISTROS ANTERIORES
        CONDI$ = " NumLtaEmp = " & NUMLISTAEMP& & " AND MARBORRA = 1"
        Call BORRAREGISTROS("LISEMPAQ", CONDI$, REGAF&, "NOMESS")

        
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
    '
    Command1.Enabled = False
    If DERACCE%("STPREMIT", "Configuración - Setup de Remitos") = 2 Then
        OPREMIT07.Show 1
    End If
    Command1.Enabled = True
    '
End Sub

Sub Command10_Click()
   Command10.Enabled = False
   If DERACCE%("CORRECRE", "Correccion Cantidades por Remito") >= 2 Then
     '
     FactuParcial% = 0
     OPX% = COMALTER%("Opciones de Trabajo\\Cancelar\Corrección de  Cantidades de Remito\Modificar Cantidad Para Facturación Parcializada")
     If OPX% <= 1 Then GoTo 99
     If OPX% = 3 Then FactuParcial% = 1
     
     CMO$ = TRIM$(Control$("", "CODREFAC"))
     If Len(CMO$) <> 2 Then
        CMO$ = "RT"
     End If
     '
     Call CARSELMESA
     '
10   Screen.MousePointer = 1
     Call SELECHO("SELMESA")
     MESASEL$ = VALACT1$("SELMESA")
     If TRIM$(MESASEL$) = "" Then
        GoTo 99
     End If
     CORRESQ$ = MESANO$(MESASEL$, Des%, Has%)
     NUMES% = NUMMES%(CORRESQ$)
     CORREMI07.Label44.Caption = MESASEL$
     CORREMI07.Label45.Caption = NUMES%
     

     '
     Call FRESHECHO("SELREMI")
     Screen.MousePointer = 11
     '
     DES1 = FETEXCOR1$(Des%)
     HAS1 = FETEXCOR1$(Has%)
     '
     SQLX$ = " SELECT DISTINCT DoPriCor, Nume_Clie, FechMov FROM MOVISTO WITH(NOLOCK)"
     SQLX$ = SQLX$ + " WHERE FechMov >= '" & DES1 & "' "
     SQLX$ = SQLX$ + " AND FechMov <= '" & HAS1 & "' "
     'PARA EVITAR COMPROBANTES ARRUINADOS
     SQLX$ = SQLX$ + " AND SUBSTRING(REFERCOR,1,4) <> 'ANUL' "
     SQLX$ = SQLX$ + " AND CodiMovi =  '" & CMO$ & "'  "
     '
     Call APERBASDAT("STOCKS")
     Set MOVITMP = READSET(SQLX$)
'     Set MOVITMP = New ADODB.Recordset
'     MOVITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
     '
     With MOVITMP
       On Error Resume Next
       .MoveFirst
       If Err Then
         Call MENSERR(24, " No Existen Registros que Cumplan\con la Condición SQL.")
         On Error GoTo 0
         GoTo 10
       End If
       On Error GoTo 0
       '
       J& = 0
       Call FRESHECHO("!SELREMI")
       Do While (.EOF = False)
         CLISE = XVALO(!Nume_Clie)
         REMI$ = SAFETEXT$(!DoPriCor)
         FF% = CVINT(!FECHMOV)
         '\\\OT-06-0083-WAR-07/03/06///
         F1$ = FECHATEX$(FF%)
         REFE1$ = F1$ + " - " + rasocli$(CLISE)
         '\\\OT-06-0083-WAR-FIN///
         aa2 = !Nume_Clie
         Z$ = Space$(40) + String$(8, 0)
         Call REPLA(Z$, REMI$, 1, 10)
         Call REPLA(Z$, REFE1$, 11, 28) 'RASOCLI$(CLISE), 11, 30)
         Call REPLA(Z$, MKS$(U&), 39, 4)
         Call REPLA(Z$, MKI$(FF%), 43, 2)
         Call REPLA(Z$, MKS$(CLISE), 45, 4)
         J& = J& + 1
         Call GRAREG("!SELREMI", Z$, J&)
       .MoveNext
       Loop
     End With
     '
     Call SETULTREG("!SELREMI", J&)
     Call CIERRARCH("!SELREMI")
     Screen.MousePointer = 1
     MOVITMP.Close
     Set MOVITMP = Nothing
     '
     If J& < 1 Then
        Call MENSERR(24, "No se Encontraron Remitos\Posibles de Corregir Cantidades\en el Período Elegido.")
        GoTo 10
     End If
     '
     CORREMI07.Command2.Caption = "Registrar"
     CORREMI07.Label15.Caption = "Cantidad Corregida:"
     'SI ELIGIO MODIFICAR CANTIDADES PARA FACTURACION PARCIAL CAMBIA EL CAPTION COMO BANDERA PARA LA TRANSACCION A REALIZAR
     If FactuParcial% = 1 Then
         CORREMI07.Command2.Caption = "Guardar Cantidad Parcial"
         CORREMI07.Label15.Caption = "Generar Parcial Por: "
     End If
CORREMI07.Show 1
GoTo 99

15   RERE$ = REGSEL$("!SELREMI")
     If Len(RERE$) <= 4 Then GoTo 10
     '
     REMISEL$ = TRIM$(Left$(RERE$, 10))
     CLIE$ = Mid$(RERE$, 11, 30)
     FEXI% = CVI(Mid$(RERE$, 45, 2))
     FEX = FEXI%
     FEMOVD = FETEXCOR1$(FEXI%)
     CLISEL% = CVI(Mid$(RERE$, 47, 2))
     '
20   Call APERBASDAT("STOCKS")
     Screen.MousePointer = 11
     SQLX$ = "SELECT * FROM MOVISTO "
     SQLX$ = SQLX$ + "WHERE DoPriCor = '" & REMISEL$ & "' "
     SQLX$ = SQLX$ + "AND Nume_Clie = " & CLISEL% & " "
     SQLX$ = SQLX$ + " AND FechMov = '" & FEMOVD & "' "
     'SQLX$ = SQLX$ + "ORDER BY NuOrItem ASC"
     '
     Set MOVITMP = New ADODB.Recordset
     MOVITMP.CursorType = adOpenDynamic
     MOVITMP.LockType = adLockPessimistic
     MOVITMP.Open FILTSQL$(SQLX), FLEXCONN, , , adCmdText
     NORIMAX% = 0
     With MOVITMP
       On Error Resume Next
       .MoveLast
       If Err Then
         Call MENSERR(24, " No Existen Registros que Cumplan\con la Condición SQL.")
         On Error GoTo 0
         GoTo 10
       End If
       On Error GoTo 0
       '
       NORIMAX% = XVALO(!NuOrItem)
       DOPRI$ = REMISEL$
       DOPRILA$ = SAFETEXT$(!DOPRILAR)
       DOSEC$ = ""       ' QUEDA LIBRE PARA PONERLE EL NUMERO DE FACTURA
       DOSELA$ = ""      ' QUEDA LIBRE PARA PONERLE EL NUMERO DE FACTURA
       MONEVA% = XVALO(!MoneVal)
       DEPOMO% = XVALO(!depomovi)
       REFEX8$ = SAFETEXT$(!ReferCor)
       '
       HOII% = HOY(HOS$)
       '
       .MoveFirst
       J& = 0: NORI$ = ""
       Do While (.EOF = False)
         NORIT% = XVALO(!NuOrItem)
         If NORIT% < 1 Then
            NORIT% = NORIMAX% + 1
            !NuOrItem = NORIT%
            .Update
            NORIMAX% = NORIT%
            GoTo 29
         End If
         '
         For KKI% = 1 To Len(NORI$) Step 2
           If CVI(Mid$(NORI$, KKI%, 2)) = NORIT% Then
               NORIT% = NORIMAX% + 1
               !NuOrItem = NORIT%
               .Update
               NORIMAX% = NORIT%
               GoTo 29
           End If
         Next KKI%
         '
29       NORI$ = NORI$ + MKI$(NORIT%)
         '
         If Abs(XVALO(!CANTINGRE) - XVALO(!CANTSALID)) >= 0.005 Then
           XX$ = REGBLAN$("BOLREDET")
           Call REPLA(XX$, MKI$(XVALO(!NuOrItem)), 105, 2)
             If XVALO(!NuOrItem) > NORIMAX% Then NORIMAX% = XVALO(!NuOrItem)
           Call REPLA(XX$, MKI$(XVALO(!cod_inte)), 83, 2)
           Call REPLA(XX$, Unimed$(XVALO(!cod_inte)), 85, 4)
           Call REPLA(XX$, MKD$(XVALO(!CANTSALID) - XVALO(!CANTINGRE)), 89, 8)
           Call REPLA(XX$, SAFETEXT$(!idenlote), 119, 16)
             VUNIT = XVALO(!VALOUNIT)
             NUPEDI& = XVALO(!NUPEDCLI) 'NUMERO DE PEDIDO
           Call REPLA(XX$, MKS$(NUPEDI&), 151, 4) 'Nº PEDIDO.
           Call REPLA(XX$, MKS$(VUNIT), 159, 4)
           Call REPLA(XX$, MKD$(0), 171, 8)
           J& = J& + 1
           Call GRAREG("!BOLREDET", XX$, J&)
         End If
       .MoveNext
       Loop
     End With
     Call SETULTREG("!BOLREDET", J&)
     Screen.MousePointer = 1
     '
30   FETEXI$ = REPLACAR$(FECHATEX$(FEX), "/", "-")
     VTB% = VENTABU%("CORREMIT/TITUPAN=Corrección " + REMISEL$ + " - " + TRIM$(CLIE$) + " - " + FETEXI$)
     If VTB% < 0 Then GoTo 15
     '
     Screen.MousePointer = 11
     Call APERBASDAT("STOCKS")
     Screen.MousePointer = 11
     J& = ULTREG&("!BOLREDET")
     For KKIL& = 1 To J&
       XX$ = REGLEIDO$("!BOLREDET", KKIL&)
       CIXI% = CVI(Mid$(XX$, 83, 2))
       LOTE$ = Mid$(XX$, 119, 16)
       NUORIT% = CVI(Mid$(XX$, 105, 2))
       CACORREC = CVD(Mid$(XX$, 89, 8))
       NUPEDID1& = CVS(Mid$(XX$, 151, 4))
       VUNIT = CVS(Mid$(XX$, 159, 4))
       DOSELA$ = "P/F - PC." + TRIM$(Str$(NUPEDID1&))
       DOSEC$ = "PC." + TRIM$(Str$(NUPEDID1&))
       '
       '
       With MOVITMP ' Movisto
         If NUORIT% < 1 Then   ' para tener en cuenta items agregados
           If Abs(CACORREC) >= 0.05 Then
             .AddNew   ' YA CONVERTIDO
             On Error Resume Next
             !CODIEMPR = CodiEmp%
             On Error GoTo 0
             !FECHMOV = CVDAT(FEXI%)
             !cod_inte = CIXI% 'GUARDA EL CODIGO INTERNO EDITADO
             !Cod_Exte = CODEXT(CIXI%) 'GUARDA EL COD. EXT EDITADO
             !Descrip = Left$(DESCRIT(CIXI%), 64) 'GUARDA LA DESCRIPCION.
             !idenlote = Left$(LOTE$, 16) 'GUARDA EL LOTE EDITADO
             !CODIMOVI = CMO$
             !DoPriCor = Left$(DOPRI$, 12)
             !DOPRILAR = Left$(DOPRILA$, 24)
             !DoSecCor = Left$(DOSEC$, 12)
             !DOSECLAR = Left$(DOSELA$, 24)
             !NumRem_Cli = ""
               NORIMAX% = NORIMAX% + 1
             !NuOrItem = NORIMAX%
             !NUPEDCLI = NUPEDID1& 'GUARDA EL Nº PED. EDITADO
             !ReferCor = Left$("(Agr.Man) " + REFEX8$, 64)
             !VALOUNIT = VUNIT 'GUARDA EL VALOR UNITARIO EDITADO
             !MoneVal = MONEVA%
             !depomovi = DEPOMO%
             !Nume_Clie = CLISEL%
             !Nume_Prov = 0
             !CANTSALID = 0: !CANTINGRE = 0
             If CACORREC > 0 Then
                 !CANTSALID = CACORREC 'GUARDA LA CANTIDAD EDITADA
               ElseIf CACORREC < 0 Then
                 !CANTINGRE = Abs(CACORREC)
             End If
             !CantSaldo = 0
             !FeReMovi = Now
             !NumUsuar = USNBR%
             .Update
           End If
           GoTo 39
         End If
         '
         SQLX$ = "SELECT * FROM MOVISTO "
         SQLX$ = SQLX$ + "WHERE DoPriCor = '" & REMISEL$ & "' "
         SQLX$ = SQLX$ + "AND Nume_Clie = " & CLISEL% & " "
         SQLX$ = SQLX$ + " AND NuOrItem = " & NUORIT% & ""
         '
         Set MOVITMP1 = New ADODB.Recordset
         MOVITMP1.CursorType = adOpenDynamic
         MOVITMP1.LockType = adLockPessimistic
         MOVITMP1.Open FILTSQL$(SQLX), FLEXCONN, , , adCmdText
         '
         With MOVITMP1
            On Error Resume Next
            .MoveFirst
            If Err Then
               On Error GoTo 0
               GoTo 39
            End If
            On Error GoTo 0
            '
            If !CODIMOVI = CMO$ Then
               If CACORREC >= 0 Then
                   !CANTSALID = CACORREC 'GUARDA LA CANTIDAD EDITADA
                   !CANTINGRE = 0
                   GoTo 35
                 Else
                   Call MENSERR(24, "Imposible Corregir Item Número " + TRIM$(Str$(NUORIT%)))
                   GoTo 39
               End If
            End If
            '
            If !CODIMOVI = "DV" Then
               If CACORREC <= 0 Then
                   !CANTINGRE = Abs(CACORREC) 'GUARDA LA CANTIDAD EDITADA
                   !CANTSALID = 0
                   GoTo 35
                 Else
                   Call MENSERR(24, "Imposible Corregir Item Número " + TRIM$(Str$(NUORIT%)))
                   GoTo 39
               End If
            End If
            '
            If Abs(CACORREC) > 0.005 Then
               If Sgn(CACORREC) <> Sgn(!CANTSALID - !CANTINGRE) Then
                  Call MENSERR(24, "Imposible Corregir Item Número " + TRIM$(Str$(NUORIT%)))
                  GoTo 39
               End If
               If CACORREC > 0 Then
                   !CANTSALID = CACORREC 'GUARDA LA CANTIDAD EDITADA
                   !CANTINGRE = 0
                   GoTo 35
                 Else
                   !CANTINGRE = Abs(CACORREC)
                   !CANTSALID = 0
                   GoTo 35
               End If
            End If
            '
            If Abs(CACORREC) < 0.005 Then
               !CANTSALID = 0: !CANTINGRE = 0
            End If
            '
35          !VALOUNIT = VUNIT 'GUARDA EL VALOR UNITARIO EDITADO
            !idenlote = LOTE$ 'GUARDA EL LOTE EDITADO
            !cod_inte = CIXI% 'GUARDA EL CODIGO INTERNO EDITADO
            !Cod_Exte = CODEXT(CIXI%) 'GUARDA EL COD. EXT EDITADO
            !NUPEDCLI = NUPEDID1& 'GUARDA EL Nº PED. EDITADO
            !Descrip = DESCRIT(CIXI%) 'GUARDA LA DESCRIPCION.
            .Update
            MOVITMP1.Close
            Set MOVITMP1 = Nothing
         End With
       End With
       '
39   Next KKIL&
     Screen.MousePointer = 1
     '
     Call COMUNI("Corrección Procesada")
     GoTo 15
     '
   End If
   '
99 Command10.Enabled = True
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call CIERRARCH("*.*")
   Call CONECRUN("COCUMENT", "Control de Cumplimiento de Entregas")
   Command11.Enabled = True
End Sub

Sub Command12_Click()
'\\\OT-05-0437-WAR-FIN///
  If DERACCE%("CORREMIT", "Corrección de Remitos") < 2 Then
    Exit Sub
  End If
  '
  EPAC$ = TRIM$(UCase$(EMPREAC$))
  If InStr(EPAC$, "ENVAPLAST") > 0 Or InStr(EPAC$, "CAMBACERES") > 0 Then
    Call AsignarPedidoaRemito
  End If
  FEX = HOY(HO$)
  Has% = FEX
  HASS$ = FECHATEX$(FEX)
  DESS$ = "01" + Mid$(HASS$, 3)
  Des% = FECHANUM(DESS$)
  '
  NCLIE% = Val(Text1)
  RECORD$ = Space(15)
  Call REPLA(RECORD$, MKI$(NCLIE%), 1, 2)
  Call REPLA(RECORD$, MKI(Des%), 3, 2)
  Call REPLA(RECORD$, MKI(Has%), 5, 2)
  Call REPLA(RECORD$, MKS(0), 7, 4)
  Call REPLA(RECORD$, MKS(0), 11, 4)
  '
10 OBX$ = OBJPLA$("SELEREMPED", RECORD$)
  If OBX$ <> "" Then
    '
    RECORD$ = OBX$
    DES2% = CVI(Mid$(OBX$, 3, 2))
    HAS2% = CVI(Mid$(OBX$, 5, 2))
    DESD3$ = CVDAT(DES2%)
    HAST3$ = CVDAT(HAS2%)
    NCLIE% = CVI(Mid$(OBX$, 1, 2))
    NPED = CVS(Mid$(OBX$, 7, 4))
    NPEDRE = CVS(Mid$(OBX$, 11, 4))
    '
    If NCLIE% < 1 Or REGICLI&(NCLIE%) < 1 Then
      Call MENSERR(24, "Cliente no Válido")
      GoTo 10
    End If
    '
    If NPED < 1 Or NPEDRE < 1 Then
      Call MENSERR(24, "Números de Pedido no Válidos")
      GoTo 10
    End If
    '
    Call APERBASDAT("PEDCLI")
'    SQLX2$ = "NroPed = " & NPED & ""
'    PedEnca.FindFirst SQLX2$
    SQLX$ = "SELECT * FROM PEDENCA"
    SQLX$ = SQLX$ + " WHERE  NroPed = " & NPED & ""
    Dim PEDDDIDO As ADODB.Recordset
    Set PEDDDIDO = New ADODB.Recordset
    PEDDDIDO.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    On Error Resume Next
    With PEDDDIDO
      .MoveFirst
      'If PedEnca.NoMatch = True Then
      If Err > 0 Then
        Call MENSERR(24, "Pedido Origen no Encontrado")
        On Error GoTo 0
        GoTo 10
      End If
    End With
    Set PEDDDIDO = Nothing
    '
    'SQLX2$ = "NroPed = " & NPEDRE & ""
    'PedEnca.FindFirst SQLX2$
    
    SQLX$ = "SELECT * FROM PEDENCA"
    SQLX$ = SQLX$ + " WHERE  NroPed = " & NPEDRE & ""
    'Dim PEDDDIDO As ADODB.Recordset
    Set PEDDDIDO = New ADODB.Recordset
    PEDDDIDO.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    On Error Resume Next
    With PEDDDIDO
      .MoveFirst
      'If PedEnca.NoMatch = True Then
      If Err > 0 Then
        Call MENSERR(24, "Pedido Reemplazo no Encontrado")
        On Error GoTo 0
        GoTo 10
      End If
      On Error GoTo 0
    End With
    Set PEDDDIDO = Nothing
    '
    CMO$ = TRIM$(Control$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
      CMO$ = "RT"
    End If
    Screen.MousePointer = 11
    Call APERBASDAT("STOCKS")
    '
    SQLX$ = " SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + "WHERE Nume_Clie = " & NCLIE% & " "
    SQLX$ = SQLX$ + " AND (CodiMovi = '" & CMO$ & "' "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
    SQLX$ = SQLX$ + " AND DoPriLar <> '' "
    SQLX$ = SQLX$ + " AND DoSecLar IS NULL or (DoSecLar = '') OR (DoSecLar LIKE 'P/F%')) "
    SQLX$ = SQLX$ + " AND FechMov >= '" & DESD3$ & "' "
    SQLX$ = SQLX$ + " AND FechMov <= '" & HAST3$ & "' "
    SQLX$ = SQLX$ + " AND VAL(MID(DoSecCor,4))= " & NPED & " "
    '
    'Set PEDTMP = Stocks.OpenRecordset(SQLX$, 2)
    Dim PEDTMP As ADODB.Recordset
    Set PEDTMP = New ADODB.Recordset
    PEDTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText

    With PEDTMP
      '
      On Error Resume Next
        .MoveFirst
      If Err Then
        On Error GoTo 0
        Call COMUNI("Ningun Registro Coincide con Condición SQL.")
        Exit Sub
      End If
      On Error GoTo 0
      '
      Do While .EOF = False
        CIXI% = !cod_inte
        CANREM = !CANTSALID - !CANTINGRE
        '
        SQLX2$ = "SELECT * FROM PEDDETA"
        SQLX2$ = SQLX2$ + " WHERE CodiInt = " & CIXI% & " "
        SQLX2$ = SQLX2$ + " AND NroPed = " & NPED & ""
        Call APERBASDAT("PEDCLI")
        Dim PEDEEETA As ADODB.Recordset
        Set PEDEEETA = New ADODB.Recordset
        PEDEEETA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
        'PedDeta.FindFirst SQLX2$
        'If PedDeta.NoMatch = False Then
        On Error Resume Next
        With PEDEEETA
          If Err < 1 Then
             On Error GoTo 0
            TOREM = !CANTENT - CANREM
            If TOREM < 0 Then TOREM = 0
              'PedDeta.Edit
              !CANTENT = TOREM
              .Update
            End If
            On Error GoTo 0
          'End If
        End With
        Set PEDEEETA = Nothing
        '
        SQLX2$ = "SELECT * FROM PEDDETA"
        SQLX2$ = SQLX$ + " WHERE CodiInt = " & CIXI% & " "
        SQLX2$ = SQLX2$ + " AND NroPed = " & NPEDRE & ""
        Call APERBASDAT("PEDCLI")
        Dim PEDEEETA1 As ADODB.Recordset
        Set PEDEEETA1 = New ADODB.Recordset
        PEDEEETA1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
        On Error Resume Next
        With PEDEEETA1
            If Err < 1 Then
            'PedDeta.FindFirst SQLX2$
            'If PedDeta.NoMatch = False Then
            On Error GoTo 0
            PUNI1 = PedDeta!PreUnid
            TOREM = PedDeta!CANTENT + CANREM
            'PedDeta.Edit
            !CANTENT = TOREM
            .Update
          Else
            On Error GoTo 0
            PUNI1 = PEDTMP!VALOUNIT
          End If
        End With
        Set PEDEEETA1 = Nothing
        '
        '.Edit 'RUTINA DE EDICIÓN
          !DoSecCor = "PC." & NPEDRE
          !DOSECLAR = "P/F - PC." & NPEDRE
          !NUPEDCLI = NPEDRE
          !VALOUNIT = PUNI1
        .Update
        '
        .MoveNext
      Loop
    End With
    Set PEDTMP = Nothing
    Call COMUNI("Reemplazo Procesado")
  End If
  Screen.MousePointer = 1
  '
  '\\\OT-05-0556-WAR-FIN///

End Sub
Sub AsignarPedidoaRemito()
  FEX = HOY(HO$)
  Has% = FEX
  HASS$ = FECHATEX$(FEX)
  DESS$ = "01" + Mid$(HASS$, 3)
  Des% = FECHANUM(DESS$)
  '
  NCLIE% = Val(Text1)
  RECORD$ = Space(15)
  Call REPLA(RECORD$, MKI$(NCLIE%), 1, 2)
  Call REPLA(RECORD$, MKI(Des%), 3, 2)
  Call REPLA(RECORD$, MKI(Has%), 5, 2)
  Call REPLA(RECORD$, MKS(0), 7, 4)
  Call REPLA(RECORD$, MKS(0), 11, 4)
  '
10 OBX$ = OBJPLA$("SELEREMPED", RECORD$)
  If OBX$ <> "" Then
    '
    RECORD$ = OBX$
    DES2% = CVI(Mid$(OBX$, 3, 2))
    HAS2% = CVI(Mid$(OBX$, 5, 2))
    DESD3$ = CVDAT(DES2%)
    HAST3$ = CVDAT(HAS2%)
    NCLIE% = CVI(Mid$(OBX$, 1, 2))
    NPED = CVS(Mid$(OBX$, 7, 4))
    NPEDRE = CVS(Mid$(OBX$, 11, 4))
    '
    If NCLIE% < 1 Or REGICLI&(NCLIE%) < 1 Then
      Call MENSERR(24, "Cliente no Válido")
      GoTo 10
    End If
    '
    If NPED < 1 Or NPEDRE < 1 Then
      Call MENSERR(24, "Números de Pedido no Válidos")
      GoTo 10
    End If
    '
    Call APERBASDAT("PEDCLI")
'    SQLX2$ = "NroPed = " & NPED & ""
'    PedEnca.FindFirst SQLX2$

    SQLX$ = "SELECT * FROM PEDENCA"
    SQLX$ = SQLX$ + " WHERE  NroPed = " & NPED & ""
    Dim PEDDDIDO As ADODB.Recordset
    Set PEDDDIDO = New ADODB.Recordset
    PEDDDIDO.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    On Error Resume Next
    With PEDDDIDO
      .MoveFirst
      'If PedEnca.NoMatch = True Then
      If Err > 0 Then
        Call MENSERR(24, "Pedido Origen no Encontrado")
        On Error GoTo 0
        GoTo 10
      End If
    End With
    Set PEDDDIDO = Nothing
    '
    'SQLX2$ = "NroPed = " & NPEDRE & ""
    'PedEnca.FindFirst SQLX2$
    If NPED > 0 Then
        SQLX$ = "SELECT * FROM PEDENCA"
        SQLX$ = SQLX$ + " WHERE  NroPed = " & NPEDRE & ""
        'Dim PEDDDIDO As ADODB.Recordset
        Set PEDDDIDO = New ADODB.Recordset
        PEDDDIDO.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
        On Error Resume Next
        With PEDDDIDO
          .MoveFirst
          'If PedEnca.NoMatch = True Then
          If Err > 0 Then
            Call MENSERR(24, "Pedido Reemplazo no Encontrado")
            On Error GoTo 0
            GoTo 10
          End If
          On Error GoTo 0
        End With
        Set PEDDDIDO = Nothing
    End If
    '
    CMO$ = TRIM$(Control$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
      CMO$ = "RT"
    End If
    Screen.MousePointer = 11
    Call APERBASDAT("STOCKS")
    '
    SQLX$ = " SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + "WHERE Nume_Clie = " & NCLIE% & " "
    SQLX$ = SQLX$ + " AND (CodiMovi = '" & CMO$ & "' "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
    SQLX$ = SQLX$ + " AND DoPriLar <> '' "
    SQLX$ = SQLX$ + " AND DoSecLar IS NULL or (DoSecLar = '') OR (DoSecLar LIKE 'P/F%')) "
    SQLX$ = SQLX$ + " AND FechMov >= '" & DESD3$ & "' "
    SQLX$ = SQLX$ + " AND FechMov <= '" & HAST3$ & "' "
    SQLX$ = SQLX$ + " AND VAL(MID(DoSecCor,4))= " & NPED & " "
    '
    'Set PEDTMP = Stocks.OpenRecordset(SQLX$, 2)
    Dim PEDTMP As ADODB.Recordset
    Set PEDTMP = New ADODB.Recordset
    PEDTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText

    With PEDTMP
      '
      On Error Resume Next
        .MoveFirst
      If Err Then
        On Error GoTo 0
        Call COMUNI("Ningun Registro Coincide con Condición SQL.")
        Exit Sub
      End If
      On Error GoTo 0
      '
      Do While .EOF = False
        CIXI% = !cod_inte
        CANREM = !CANTSALID - !CANTINGRE
        '
        SQLX2$ = "SELECT * FROM PEDDETA"
        SQLX2$ = SQLX2$ + " WHERE CodiInt = " & CIXI% & " "
        SQLX2$ = SQLX2$ + " AND NroPed = " & NPED & ""
        Call APERBASDAT("PEDCLI")
        Dim PEDEEETA As ADODB.Recordset
        Set PEDEEETA = New ADODB.Recordset
        PEDEEETA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
        'PedDeta.FindFirst SQLX2$
        'If PedDeta.NoMatch = False Then
        On Error Resume Next
        With PEDEEETA
          If Err < 1 Then
             On Error GoTo 0
            TOREM = !CANTENT - CANREM
            If TOREM < 0 Then TOREM = 0
              'PedDeta.Edit
              !CANTENT = TOREM
              .Update
            End If
            On Error GoTo 0
          'End If
        End With
        Set PEDEEETA = Nothing
        '
        SQLX2$ = "SELECT * FROM PEDDETA"
        SQLX2$ = SQLX$ + " WHERE CodiInt = " & CIXI% & " "
        SQLX2$ = SQLX2$ + " AND NroPed = " & NPEDRE & ""
        Call APERBASDAT("PEDCLI")
        Dim PEDEEETA1 As ADODB.Recordset
        Set PEDEEETA1 = New ADODB.Recordset
        PEDEEETA1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
        On Error Resume Next
        With PEDEEETA1
            If Err < 1 Then
            'PedDeta.FindFirst SQLX2$
            'If PedDeta.NoMatch = False Then
            On Error GoTo 0
            PUNI1 = PedDeta!PreUnid
            TOREM = PedDeta!CANTENT + CANREM
            'PedDeta.Edit
            !CANTENT = TOREM
            .Update
          Else
            On Error GoTo 0
            PUNI1 = PEDTMP!VALOUNIT
          End If
        End With
        Set PEDEEETA1 = Nothing
        '
        '.Edit 'RUTINA DE EDICIÓN
          !DoSecCor = "PC." & NPEDRE
          !DOSECLAR = "P/F - PC." & NPEDRE
          !NUPEDCLI = NPEDRE
          !VALOUNIT = PUNI1
        .Update
        '
        .MoveNext
      Loop
    End With
    Set PEDTMP = Nothing
    Call COMUNI("Reemplazo Procesado")
  End If
  Screen.MousePointer = 1
  '
  '\\\OT-05-0556-WAR-FIN///

End Sub


Private Sub Command13_Click()
  '
  If DERACCE("VEREMANU", "Control de Remitos Anulados") < 2 Then
     Exit Sub
  End If
  '
  Command13.Enabled = False
  PREFIREM$ = TRIM$(Control$("", "PREFIREM"))
  If PREFIREM$ = "" Then PREFIREM$ = "0001"
  '
5 Call DESHASFECHA(Des%, Has%, PERIO$)
  If TRIM$(PERIO$) = "" Then GoTo 99
  Call APERBASDAT("STOCKS")
  '
  Screen.MousePointer = 11
  CMO1$ = TRIM$(Control$("", "CODREFAC"))
  If Len(CMO$) <> 2 Then
     CMO1$ = "RT"
  End If
  Call SETULTREG("!SELREMI", 0)
  Call FRESHECHO("!SELREMI")
  Screen.MousePointer = 11
  '
  DESD1$ = FETEXCOR1$(Des%)
  HAST1$ = FETEXCOR1$(Has%)
  '
  ABUX$ = "RT-X-" + PREFIREM$
  'SQLX$ = " SELECT MAX(DOPRILAR) FROM MOVISTO"
  CONDI$ = "FechMov >= '" & DESD1$ & "' "
  CONDI$ = CONDI$ + " AND FechMov <= '" & HAST1$ & "' "
  CONDI$ = CONDI$ + " AND CodiMovi =  '" & CMO1$ & "' "
  CONDI$ = CONDI$ + " AND DOPRILAR LIKE '" & ABUX$ & "%' "
  PRIREMI$ = VALOBADA("MOVISTO", "MIN(DOPRILAR)", CONDI$)
  '
  CONDI$ = "FechMov < '" & DESD1$ & "' "
  CONDI$ = CONDI$ + " AND CodiMovi =  '" & CMO1$ & "' "
  CONDI$ = CONDI$ + " AND DOPRILAR LIKE '" & ABUX$ & "%' "
  CONDI$ = CONDI$ + " AND DOPRILAR < '" & PRIREMI$ & "' "
  UREMIAN$ = VALOBADA("MOVISTO", "MAX(DOPRILAR)", CONDI$)
  '
  CONDI$ = "FechMov <= '" & HAST1$ & "' "
  CONDI$ = CONDI$ + " AND FechMov >= '" & DESD1$ & "' "
  CONDI$ = CONDI$ + " AND CodiMovi =  '" & CMO1$ & "' "
  CONDI$ = CONDI$ + " AND DOPRILAR LIKE '" & ABUX$ & "%' "
  ULTREMI$ = VALOBADA("MOVISTO", "MAX(DOPRILAR)", CONDI$)
  '
  NREMI& = XVALO(Mid$(PRIREMI$, 11, 8))
  UREMI& = XVALO(Mid$(ULTREMI$, 11, 8))
  CANREMI& = UREMI& - NREMI&: If CANREMI& < 1 Then CANREMI& = 1
  If CANREMI& > 42048 Then
     Call MENSERR(24, "Demasiados Remitos - Elija menor Rango de Fechas.")
     GoTo 5
  End If
  '
  JJ& = 0: J& = 0
  '
10  NREMI& = NREMI& + 1
    JJ& = JJ& + 1
    Call TRACE(20, JJ&, CANREMI&)
    NRX$ = TRIM$(Str$(NREMI&))
    While Len(NRX$) < 8: NRX$ = "0" + NRX$: Wend
    NRX$ = ABUX$ + "-" + NRX$
    If NRX$ > ULTREMI$ Then GoTo 9
    '
    CONDI$ = "FechMov >= '" & DESD1$ & "' "
    CONDI$ = CONDI$ + " AND FechMov <= '" & HAST1$ & "' "
    CONDI$ = CONDI$ + " AND CodiMovi =  '" & CMO1$ & "' "
    CONDI$ = CONDI$ + " AND DOPRILAR = '" & NRX$ & "' "
    CRG& = CantRegistros("MOVISTO", CONDI$)
    If CRG& < 1 Then
       Call COMUNI("Remito '" + NRX$ + " Anulado.")
       Z$ = Space$(40) + String$(8, 0)
       Call REPLA(Z$, "." + FORMATNUM$(JJ&, "F8.0"), 1, 10)
       Call REPLA(Z$, NRX$, 11, 30)
       J& = J& + 1
       Call GRAREG("!SELREMI", Z$, J&)
    End If
  GoTo 10
  '
9 Call SETULTREG("!SELREMI", J&)
  Screen.MousePointer = 1
  Call MENSERR(24, "Fin del Control")
  '
  If J& > 0 Then
15   Call SELECHO("!SELREMI")
     resel$ = TRIM$(VALACT2$("!SELREMI"))
     If resel$ <> "" Then
        TIDOCUM$ = "Remito"
        NUDOCU$ = resel$
        Screen.MousePointer = 11
        '
        For UI& = ULTREG&("PDOCLST") To 1 Step -1
           XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
           If InStr(XX$, TIDOCUM$) > 4 Then
             If InStr(XX$, NUDOCU$) > 4 Then
               DOCUNU& = CVS(Left$(XX$, 4))
               Call MUESTRADOC(DOCUNU&)
               GoTo 15
             End If
           End If
        Next UI&
        Screen.MousePointer = 1
        Call COMUNI("Documento no Encontrado")
        GoTo 15
     End If
  End If
   '
99 Screen.MousePointer = 1
  Command13.Enabled = True
  '
End Sub

Sub Command14_Click()
   Command14.Enabled = False
   NREMI$ = InputBox$("Numero de Remito:", "Seleccion del Comprobante")
   If NREMI$ <> "" Then
      If XVALO(NREMI$) > 0 Then
         ABUS$ = "RT." + TRIM$(Str$(XVALO(FOREMIT07.Label10))) + "-" + TRIM$(NREMI$)
         Call CONECRUN("GENREP07/CONREMIT(" + ABUS$ + ")", "Consulta por Número de Remito")
      End If
   End If
   Command14.Enabled = True
End Sub

Private Sub Command15_Click()
    '
    Call ABRECONEXION
    '
    Des% = FECHANUM("19/03/09")
    DES1$ = FETEXCOR1$(Des%)
    Has% = FECHANUM("21/03/09")
    HAS1$ = FETEXCOR1$(Has%)
    '
    SQLX$ = "SELECT * FROM MOVISTO"
    SQLX$ = SQLX$ + " WHERE FECHMOV >= '" & DES1$ & "' AND FECHMOV<= '" & HAS1$ & "'"
    SQLX$ = SQLX$ + " AND CODIMOVI = 'RT'"
    Dim MOVITMP As ADODB.Recordset
    Set MOVITMP = New ADODB.Recordset
    MOVITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With MOVITMP
      On Error Resume Next
      .MoveFirst
      If Not (.EOF And .BOF) Then
        J% = 0
        Do While Not .EOF
          CI% = XVALO(!cod_inte)
          If STATITEM%(CI%) < 0 Then
             Codigo$ = TRIM$(Mid$(CODEXT$(CI%), 1, 11))
             CIXI% = XVALO(VALOBADA("MASTER", "CODINT", "CODIGO LIKE '" & Codigo$ & "%' AND STATUS > 0"))
             If CIXI% > 0 Then
               !cod_inte = CIXI%
               !Cod_Exte = CODEXT$(CIXI%)
               .Update
               J% = J% + 1
             End If
          End If
          .MoveNext
        Loop
      End If
      On Error GoTo 0
    End With
    On Error Resume Next
    MOVITMP.Close
    Set MOVITMP = Nothing
    On Error GoTo 0
    '
    Call COMUNI("Proceso de Recuperación Finalizado con " & J% & " Correcciones")
    Exit Sub
    '
    Call ABRECONEXION
    '
    Des% = FECHANUM("23/02/09")
    DES1$ = FETEXCOR1$(Des%)
    Has% = FECHANUM("25/02/09")
    HAS1$ = FETEXCOR1$(Has%)
    CONDI$ = "NUME_CLIE = 228 AND CODIMOVI = 'RT' AND (DOSECLAR like 'P/F%' OR DOSECLAR = '' OR DOSECLAR IS NULL) "
    CONDI$ = CONDI$ + "AND FECHMOV >= '" & DES1$ & "' AND FECHMOV <= '" & HAS1$ & "'"
     FLEXCONN.Execute (FILTSQL$("UPDATE MOVISTO SET DOSECCOR = 'FC-015359-A' WHERE " & CONDI$))
     FLEXCONN.Execute (FILTSQL$("UPDATE MOVISTO SET DOSECLAR = 'FC-0001-00015359-A' WHERE " & CONDI$))
    '
    Des% = FECHANUM("26/02/09")
    DES1$ = FETEXCOR1$(Des%)
    Has% = FECHANUM("28/02/09")
    HAS1$ = FETEXCOR1$(Has%)
    CONDI$ = "NUME_CLIE = 228 AND CODIMOVI = 'RT' AND (DOSECLAR like 'P/F%' OR DOSECLAR = '' OR DOSECLAR IS NULL) "
    CONDI$ = CONDI$ + "AND FECHMOV >= '" & DES1$ & "' AND FECHMOV <= '" & HAS1$ & "'"
     FLEXCONN.Execute (FILTSQL$("UPDATE MOVISTO SET DOSECCOR = 'FC-015393-A' WHERE " & CONDI$))
     FLEXCONN.Execute (FILTSQL$("UPDATE MOVISTO SET DOSECLAR = 'FC-0001-00015393-A' WHERE " & CONDI$))
    '
    MsgBox "FIN"
Exit Sub
    Command15.Enabled = False
    Screen.MousePointer = 11
    '
    Call ABREPEDCLI ' RECUPERACION DE DATOS PARA GUIDI DE PEDIDOS QUE NO GUARDARON LA CANTIDAD ENTREGADA
    Call ABRESTOCKS
    '
    SQLX$ = "SELECT * FROM PEDDETA"
    SQLX$ = SQLX$ + " WHERE CANPED > 0 "
    'SQLX$ = SQLX$ + " AND NROPED > 244"
    SQLX$ = SQLX$ + " order by NROPED"
    Dim PEDTMP As ADODB.Recordset
    Set PEDTMP = New ADODB.Recordset
    PEDTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With PEDTMP
      Do While Not .EOF
       NPED& = XVALO(!NROPED)
       NUOR% = XVALO(!NuOrItem)
       CI% = XVALO(!CODIINT)
       CONDI$ = "NuPedCli = " & NPED& & " AND Cod_Inte = " & CI% & ""
       Entrada# = XVALO(VALOBADA("MOVISTO", "SUM(CantIngre) ", CONDI$))
       SALIDA# = XVALO(VALOBADA("MOVISTO", "SUM(CantSalid) ", CONDI$))
       TOTALENTREG# = SALIDA# - Entrada#
       !CANTENT = TOTALENTREG#
       CANNNSALDO# = XVALO(!CANPED) - TOTALENTREG#
       If CANNNSALDO# < 0 Then CANNNSALDO# = 0
       !CanSaldo = CANNNSALDO#
       .Update
       .MoveNext
      Loop
    End With
    '
    PEDTMP.Close
    Set PEDTMP = Nothing
    Call MENSERR(24, "Proceso Terminado")
    Command15.Enabled = True
    Screen.MousePointer = 1
    '
End Sub

Private Sub Command17_Click()
     '
     Command17.Enabled = False
     Screen.MousePointer = 11
     CMO$ = TRIM$(Control$("", "CODREFAC"))
     If Len(CMO$) <> 2 Then
        CMO$ = "RT"
     End If
     '
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
     CONDI$ = "FechMov >= '" & Desde & "'"
     CONDI$ = CONDI$ + "AND FechMov <= '" & Hasta & "'"
     CONDI$ = CONDI$ + " AND substring(REFERCOR,1,4) <> 'ANUL' " 'PARA EVITAR COMPROBANTES ARRUINADOS
     CONDI$ = CONDI$ + " AND CodiMovi =  '" & CMO$ & "'"
     CONDI$ = CONDI$ + " AND substring(DOPRILAR,1,2) = 'RT' "
     '
     Screen.MousePointer = 11
     '
     Call COPYSTRU("SELDOCAP", "SELREMIT")
     Call CARGALISEL("SELREMIT", "MOVISTO", "DOPRILAR/A18;FechMov/D8;REFERCOR/A24", CONDI$, "DISTINCT")
     '
     If ULTREG&("SELREMIT") < 1 Then
        Call MENSERR(24, "No se Encontraron Remitos en el Rango Seleccionado.")
        GoTo 10
     End If

15   Call SELECHO("SELREMIT")
     '
     Screen.MousePointer = 1
     '
     NUDOCU$ = TRIM$(Mid$(VALACT1$("SELREMIT"), 6))
     If NUDOCU$ = "" Then GoTo 10
     TIDOCUM$ = "Remito"
     nubus = XVALO(Mid$(NUDOCU$, 6))
     '
     Screen.MousePointer = 11
     '
     DOCUNU& = 0
     'LE PASO EN LA PARAMETRO DE CONTROL EL NOMESS PARA QUE SI NO ENCUENTRA EL DOCUMENTO
     'NO PRESENTE MENSAJE Y CONTINUE Y LO BUSQUE DE LA MANERA QUE ESTABA ANTES
     Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, Str(nubus), NUDOCU$, "NOMESS")
     
     If DOCUNU& > 0 Then GoTo 15
     
     '
     For UI& = ULTREG&("PDOCLST") To 1 Step -1
        XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
        '\\\OT-08-0646-OD-20/08/08///
         If InStr(XX$, TIDOCUM$) > 4 Then
            If InStr(XX$, NUDOCU$) > 4 Then
                PPXX% = InStr(XX$, "Nro.")
                If PPXX% > 0 Then
                 XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
                End If
                If InStr(XX$, NUDOCU$) > 4 Then
                 DOCUNU& = CVS(Left$(XX$, 4))
                 GoTo 50
                End If
            End If
         End If
     Next UI&
     Screen.MousePointer = 1
     Call ReimpresionDeRemito(NUDOCU$, OK%)
     If OK% = 1 Then GoTo 15
     
40   Call COMUNI("No se Encontró el Documento Impreso")
     GoTo 15
     '
50   Call MUESTRADOC(DOCUNU&)
     GoTo 15
     '
99   Command17.Enabled = True
     '
End Sub
Sub ReimpresionDeRemito(rto$, OK%)
    
    PRINTERPORT$ = TRIM$(Control$(IDENTER$, "PORTREM"))
    FORIPRE$ = TRIM$(Control$("", "FOREMIFA"))
    
    REPLAORI% = 0
    If Control$("REMSTAN", "REPLAORI") = "SI" Then
       REPLAORI% = 1
    End If

    OK% = 0
    If TRIM$(FORIPRE$) <> "" Then
      '
      Call LIMPIATETAB
      '
      CONDI$ = "Doprilar LIKE '%" & rto$ & "%'"
      FF% = CVINT(VALOBADA("MOVISTO", "fechmov", CONDI$))
      NCLI = XVALO(VALOBADA("MOVISTO", "nume_clie", CONDI$))
      NROREMIT$ = rto$
      
      '
      FEX = FF%
      FECHDOC$ = FECHATEX$(FEX)
      NUMEDOC$ = NROREMIT$
      TIPODOC$ = "Remito"
      '
      CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = TRIM$(FOREMIT07.Text8.TEXT)
      CODPARAM$(2) = "NUBULTOS": DOCPARAM$(2) = TRIM$(FOREMIT07.Text15.TEXT)
      CODPARAM$(3) = "CUIT-TRA": DOCPARAM$(3) = TRIM$(FOREMIT07.Text12.TEXT)
      CODPARAM$(4) = "VALO-DEC": DOCPARAM$(4) = TRIM$(FOREMIT07.Text17.TEXT)
      CODPARAM$(5) = "DOMI-TRA": DOCPARAM$(5) = TRIM$(FOREMIT07.Text14.TEXT)
      CODPARAM$(6) = "DOCUCOND": DOCPARAM$(6) = TRIM$(FOREMIT07.Text13.TEXT)
      CODPARAM$(7) = "DEPOSAL ": DOCPARAM$(7) = TRIM$(DEPOSAL$)
      'CODPARAM$(8) = "NRO-REMI": DOCPARAM$(8) = TRIM$(NROREMIT$)
      CODPARAM$(8) = "NRO-REMI": DOCPARAM$(8) = Mid$(NROREMIT$, 6) ' SE IMPRIME FORMATO 0000-00000000
      CODPARAM$(9) = "NRO-OCOM": DOCPARAM$(9) = TRIM$(FOREMIT07.Text20.TEXT)
      CODPARAM$(10) = "PATE-TRA": DOCPARAM$(10) = TRIM$(FOREMIT07.Text18.TEXT)
      CODPARAM$(11) = "N-AUTENT": DOCPARAM$(11) = TRIM$(NORVENTA)
      CODPARAM$(12) = "COND-TRA": DOCPARAM$(12) = TRIM$(FOREMIT07.Text19)
      CODPARAM$(13) = "LOCA-TRA": DOCPARAM$(13) = TRIM$(FOREMIT07.Text21)
      CODPARAM$(14) = "NRO-AENT": DOCPARAM$(14) = TRIM$(FOREMIT07.Text22)

      CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = ""
      CODPARAM$(2) = "NUBULTOS": DOCPARAM$(2) = ""
      CODPARAM$(3) = "CUIT-TRA": DOCPARAM$(3) = ""
      CODPARAM$(4) = "VALO-DEC": DOCPARAM$(4) = ""
      CODPARAM$(5) = "DOMI-TRA": DOCPARAM$(5) = ""
      CODPARAM$(6) = "DOCUCOND": DOCPARAM$(6) = ""
      CODPARAM$(7) = "DEPOSAL ": DOCPARAM$(7) = ""
      CODPARAM$(8) = "NRO-REMI": DOCPARAM$(8) = NROREMIT$
      CODPARAM$(9) = "NRO-OCOM": DOCPARAM$(9) = ""
      CODPARAM$(10) = "PATE-TRA": DOCPARAM$(10) = ""
      CODPARAM$(11) = "N-AUTENT": DOCPARAM$(11) = ""
      CODPARAM$(12) = "COND-TRA": DOCPARAM$(12) = ""
      CODPARAM$(13) = "LOCA-TRA": DOCPARAM$(13) = ""
      CODPARAM$(14) = "NRO-AENT": DOCPARAM$(14) = ""

      '\\\OT-07-0007-RG-08/01/07///
      HORA$ = HORAENT$(NCLI)
      CODPARAM$(15) = "Z-PLANO": DOCPARAM$(15) = ""
      CODPARAM$(16) = "IMP-TOTA": DOCPARAM$(16) = ""
      CODPARAM$(17) = "DOCU-APL": DOCPARAM$(17) = ""
      CODPARAM$(18) = "DESTINO": DOCPARAM$(18) = ""
      
      '\\\OT-07-0007-RG-FIN///
      CAPARDOC% = 18
      '
      CODTABU1$(1) = "COD-MAT"
      CODTABU1$(2) = "DES-MAT"
      CODTABU1$(3) = "CANTIDAD"
      CODTABU1$(4) = "NRO-LOTE"
      CODTABU1$(5) = "UNIMED"
      CODTABU1$(6) = "N-OCOMPR"
      CODTABU1$(7) = "ORD-ITEM"
      CODTABU1$(8) = "REF-MAT1"
      CODTABU1$(9) = "REF-MAT2"
      CODTABU1$(10) = "PRE-NETO"
      CODTABU1$(11) = "NET-ITEM"
      CODTABU1$(12) = "REF-MAT6"
      CODTABU1$(13) = "REF-MAT5"
      CACOLTAB1% = 13
      '
      DESTIDOCNUE = NCLI
      CAITAB1% = 0
      '
      SQLX$ = "SELECT * FROM MOVISTO WITH(NOLOCK) WHERE " & CONDI$
      Dim RstMovisto As ADODB.Recordset
      Set RstMovisto = READSET(SQLX$)
      
      REBLAN$ = REGBLAN$("BOLREDET")
      JJ& = 0
      With RstMovisto
        Do While Not .EOF
            YY$ = REBLAN$
            Call REPLA(YY$, MKI$(XVALO(!cod_inte)), 83, 2)
            Call REPLA(YY$, MKD$(XVALO(!CANTSALID) - XVALO(!CANTINGRE)), 89, 8)
            CI1% = XVALO(!cod_inte)
            
'            JJ& = JJ& + 1
            CAITAB1% = CAITAB1% + 1
            CODDX$ = CODEXT$(CI1%)
            If REPLAORI% = 1 Then
                CODDX$ = CODORIG$(CI%, NCLI)
            End If
            '
            NPDX& = XVALO(!NUPEDCLI)
            CANS$ = FORMATNUM$(XVALO(!CANTSALID) - XVALO(!CANTINGRE), "F10." & CNTDC$ & "")
            TETABU1$(1, CAITAB1%) = CODDX$
            TETABU1$(2, CAITAB1%) = SAFETEXT$(!Descrip)
            TETABU1$(3, CAITAB1%) = CANS$
            TETABU1$(4, CAITAB1%) = SAFETEXT$(!idenlote)
            TETABU1$(5, CAITAB1%) = Unimed$(CI1%)
            TETABU1$(6, CAITAB1%) = NUOCO$
            TETABU1$(7, CAITAB1%) = Str$(CAITAB1%)
            TETABU1$(8, CAITAB1%) = NUOCOCLI$(NPDX&, CI1%)
            TETABU1$(9, CAITAB1%) = NUPLANO$(CI%) '
            
            'VALORIZACION ALEJPA SEGUN LISTA VIGENTE
            IMPOUNI# = 0: IMPORTEX# = 0
            If InStr(UCase$(EMPREAC$), "ALEJPA") > 0 Then
                IMPOUNI# = PRELISTACT#(CI1%, TIMONE%)
                IMPOUNI# = IMPOUNI# * TICAMONE(TIMONE%) 'VALOR UNITARIO EN PESOS
                IMPORTEX# = CAN# * IMPOUNI#             'IMPORTE EN PESOS
                ACUMLAIMPOX# = ACUMLAIMPOX# + IMPORTEX# 'IMPORTE ACUMULADO EN PESOS
            End If
            TETABU1$(10, CAITAB1%) = FORMATNUM$(IMPOUNI#, "F12.2")
            TETABU1$(11, CAITAB1%) = FORMATNUM$(IMPORTEX#, "F12.2")
            CORIG$ = TRIM$(CODORIG$(CI%, NCLI))
             'If CORIG$ = TRIM$(CODDX$) Then CORIG$ = ""
              If CORIG$ = CODEXT$(CI1%) Then CORIG$ = ""
              'If CORIG$ <> CODEXT$(CI%) Then CORIG$ = CODEXT$(CI%) 'ESTA LINEA HABRIA QUE QUITAR
              TETABU1$(12, CAITAB1%) = CORIG$
              'FIN ALEJPA

              TETABU1$(13, CAITAB1%) = SAFETEXT$(ORD&) ' NUMERO DE ORDEN INGRESADO EN VENTANA DE ARCON NEUQUEN.
              'IMPRESION DE ESPECIFICACION
              If Control("FACTUREM", "INCLUSPE") = "SI" Then ' IMPRIME ESPECIFICACIONES DEL ARTICULO
                XX$ = ESPECIFIT$(CI1%)
                Call PRINTEXT(XX$, 2, "DES-MAT", FORIPRE$)
              End If
              Call CARLINSEP         'Carga linea de separacion entre codigo y codigo
              '
              If SERINU$ <> "" Then
                CAITAB1% = CAITAB1% + 1
                For KI1% = 1 To 20
                  TETABU1$(KI1%, CAITAB1%) = ""
                Next KI1%
                TETABU1$(2, CAITAB1%) = "S.No." + SERINU$
              End If
            .MoveNext
        Loop
      End With
      Call SETULTREG("!BOLREDET", JJ&)
      Call CIERRARCH("!BOLREDET")
      '
      Call PRINTDOC("FOREMIFA")   ' Remito de Facturacion
      
      OK% = 1
      '
    End If


End Sub

Sub Command18_Click()
      Command18.Enabled = False
      'CREO LA TABLA DE ENCABEZADO
      Call ABRESTOCKS
      Call CREACAMPO("STOCKS", "ENCEMPAQ", "NumLtaEmp", 3, 0, 1)
      Call CREACAMPO("STOCKS", "ENCEMPAQ", "CODI_CLI", 3, 0, 1)
      Call CREACAMPO("STOCKS", "ENCEMPAQ", "REFERENCIA", 10, 64)
      Call CREACAMPO("STOCKS", "ENCEMPAQ", "NumFactura", 10, 18)
      Call CREACAMPO("STOCKS", "ENCEMPAQ", "EmpaCerra ", 3, 0)
      Call CREACAMPO("STOCKS", "ENCEMPAQ", "FechEmis ", 8, 0)

      '
10    OPX% = ALTERNA%("Abrir Nuevo\Abrir Existente\Cerrar Lista de Empaque\Revertir Cierre de Lista de Empaque\Impresión de Lista de Empaque\Anular Lista de Empaque")
      If OPX% < 1 Or OPX% > 6 Then GoTo 99
      '
      If OPX% = 1 Then Call NUEVALEMPAQ 'NUEVA LISTA DE EMPAQUE
      If OPX% = 2 Then Call EMPAQEXISTE 'ABRIR EXISTENTE
      If OPX% = 3 Then Call CERRAEMPAQ  'CERRAR LISTA DE EMPAQUE
      If OPX% = 4 Then Call ABREEMPAQ   'REABRIR LISTA DE EMPAQUE CERRADA
      If OPX% = 5 Then Call FRMLIEMPAQ  'FORMULARIO DE LISTA DE EMPAQUE
      If OPX% = 6 Then Call ANULALEMPAQ 'ANULAR LISTA DE EMPAQUE
      GoTo 10
      '
99    Command18.Enabled = True

End Sub
Sub ABREEMPAQ()
    'PRESENTO VENTANA DE LISTA DE EMPAQUES SIN CERRAR
10  CONDI$ = "NUMLTAEMP > 0 AND EmpaCerra > 0 AND (NUMFACTURA IS NULL OR NUMFACTURA = '')"
    Call CARGALISEL("REABREMP", "ENCEMPAQ", "NUMLTAEMP/I5;REFERENCIA/A36", CONDI$, "DISTINCT")
    Call SELECHO("REABREMP/Lista de Empaques Cerradas")
    '
    'VALIDO
    NUMLEMPA& = XVALO(VALACT1$("REABREMP"))
    If NUMLEMPA& < 1 Then Exit Sub
    '
    'GRABO
    CONDI$ = " NUMLTAEMP = " & NUMLEMPA&
    Call ACTUREGISTRO("ENCEMPAQ", "EmpaCerra", CONDI$, 0, REGAF&, "NOMESS")
    '
    'MENSAJE
    Call COMUNI("Lista de Empaque Nro. " & NUMLEMPA& & " Re-Abierta ")
    GoTo 10
End Sub
Sub EMPAQEXISTE()


10  Call CARGALISEL("LEMPEXIS", "ENCEMPAQ", "NUMLTAEMP/I5;REFERENCIA/A36", "NUMLTAEMP > 0", "DISTINCT")
    Call SELECHO("LEMPEXIS/LISTA DE EMPAQUES")
    NUMLEMP& = XVALO(VALACT1$("LEMPEXIS"))
    If NUMLEMP& < 1 Then Exit Sub
    '
    'VALIDO SI LA SELECCIONADA NO ESTA CERRADA
    If LEMPCER%(NUMLEMP&) > 0 Then
       Call COMUNI("La Lista de Empaque Seleccionada Se Encuentra Cerrada\La Misma Podrá ser Visualizada\Pero No Podrá ser Editada")
    End If
    '
    'PRESENTA VENTANA DE CARGA Y O MODIFICACION
    '
    'RECORRO LA TABLA PARA CARGAR LA VENTANA
    '
    i& = 0
    Call SETULTREG("!LISTAEMP", 0)
    SQLX$ = "SELECT * FROM LISEMPAQ WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE "
    SQLX$ = SQLX$ + " NumLtaEmp = " & NUMLEMP&
    SQLX$ = SQLX$ + " ORDER BY NUMPALLET , NUORDPALLET"
    Dim LIMPTMP As ADODB.Recordset
    Set LIMPTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With LIMPTMP
      Do While Not .EOF
        i& = i& + 1
        NPALL% = XVALO(!Numpallet)
        CI% = XVALO(!CODINT)
        CANTIXCAJA% = XVALO(!PiezasXCaja)
        CANTICAJAS% = XVALO(!CANTICAJAS)
        NUORPL% = XVALO(!NUORDPALLET)
        NUPED& = XVALO(!NROPED)
        Call REPLA(XX$, MKI$(NPALL%), 1, 2)
        Call REPLA(XX$, MKI$(CI%), 3, 2)
        Call REPLA(XX$, MKS$(CANTIXCAJA%), 5, 4)
        Call REPLA(XX$, MKI$(CANTICAJAS%), 9, 2)
        Call REPLA(XX$, MKS$(NUPED&), 11, 4)
        Call REPLA(XX$, MKI$(NUORPL%), 15, 2)
        Call GRAREG("!LISTAEMP", XX$, i&)
        .MoveNext
      Loop
    End With
    LIMPTMP.Close
    Set LIMPTMP = Nothing
    Call SETULTREG("!LISTAEMP", i&)
    '
    ' PRESENTO LA VENTANA
    '
    ATRI$ = "/TITUPAN=LISTAEMP/TITUPAN=Lista de Empaque Nro.: " & NUMLEMP&

    Call VENTALIEMPA(NUMLEMP&, ATRI$)
    GoTo 10
    
End Sub


Sub NUEVALEMPAQ()
    '
    NULPQUE& = 1 + XVALO(VALOBADA("ENCEMPAQ", "MAX(NumLtaEmp)", "NumLtaEmp > 0"))
    FREP = HOY(HO$)
     
10  FILTX$ = TRIM$(Str$(NULPQUE&)) & ";" & TRIM$(Str$(FREP))
    Call STDFORM("NUELPQ", FILTX$)
    If STDSCREEN.APRO <> "OK" Then Exit Sub
    '
    'Validaciones
    NC% = XVALO(VALOFORM("Text-2"))
    
    
    If NC% < 1 Then
      Call COMUNI("Falta Ingresar Cta. Cliente")
      GoTo 10
    End If
     
    REFE$ = TRIM$(VALOFORM("Text-1"))
    If REFE$ = "" Then
      Call COMUNI("Falta Ingresar Referencia")
      GoTo 10
    End If
    '
20  OPX% = ALTERNA%(" Por Pedido \Edición Directa")
    If OPX% < 1 Or OPX% > 2 Then GoTo 10
    '
    ' OPCION DE CARGA DE LISTA POR PEDIDO
    If OPX% = 1 Then
         Call LISTEMPAPEDIDO(NC%, NULPQUE&)
      ElseIf OPX% = 2 Then
          Call SETULTREG("!LISTAEMP", 0)
          ATRI$ = "/TITUPAN=LISTAEMP/TITUPAN=Lista de Empaque Nro.: " & NULPQUE&
          Call VENTALIEMPA(NULPQUE&, ATRI$)
    End If
    GoTo 20
    '
'    Call ORDENALISTAMP(NULPQUE&)
End Sub
Sub LISTEMPAPEDIDO(NC%, NULPQUE&)
     
        CONDI$ = "PEDENCA.Status < 9 and PEDENCA.NROCLIE = " & NC%
        CONDI$ = CONDI$ + " AND PEDDETA.CantEnt < (PEDDETA.CanPed - 0.005) "
        If CODCARPE$ <> "" Then
          CONDI$ = CONDI$ + " AND PEDENCA.CODCARPIMP = '" & CODCARPE$ & "' " ' Carpeta de Exportacion.
        End If
        CONDI$ = CONDI$ + " AND PEDDETA.Status < 8 ORDER BY PEDENCA.NROPED"
        Call CARGALISEL("!INDIPEDI", "PEDENCA+PEDDETA/PEDENCA.NroPed = PEDDETA.NroPed", "PEDENCA.NroPed/F6;PEDENCA.REFERENCIA/A32", CONDI$, "NOMESS/DISTINCT")
        '
        'VALIDO SI EL CLIENTE SELECCIONADO TIENE PEDIDOS PENDIENTES
        If ULTREG("!INDIPEDI") < 1 Then
           Call COMUNI("Cliente sin Pedidos Pendientes")
           Exit Sub
        End If
        
15      Call SELECHO("!INDIPEDI/Pedidos Pendientes")
        NPED& = XVALO(VALACT1$("!INDIPEDI"))
        If NPED& < 1 Then Exit Sub
        '
        'VALIDO QUE EL PEDIDO NO TENGA LISTA DE EMPAQUE
        CONDI$ = "NROPED = " & NPED&
        If CantRegistros("LISEMPAQ", CONDI$) > 0 Then
          Call COMUNI("Pedido Seleccionado Ya Tiene Lista de Empaque")
          GoTo 15
        End If
        '
        Call SETULTREG("!LISTAEMP", 0)
        i& = 1: IANTE& = 0
        XX$ = REGBLAN$("LISTAEMP")
        Call GRAREG("!LISTAEMP", XX$, i&)
        NUMERODEPALET% = 1
        MECUBPORPALET = 0
        MAXMECUBPORPALET = XVALO(Control$("REMSTAN", "LIMITEM3"))
        If MAXMECUBPORPALET < 1 Then
           Call MENSERR(24, "Imposible Continuar !!!\Falta Ingresar Límite de M3 Por Pallet")
        End If
        '
        'RECORRO EL DETALLE DEL PEDIDO SELECCIONADO
        SQLX$ = "SELECT * FROM PEDDETA WITH(NOLOCK) "
        SQLX$ = SQLX$ + " WHERE NroPed = " & NPED& & " "
        SQLX$ = SQLX$ + " AND Status < 8 "
        SQLX$ = SQLX$ + " AND CantEnt < (CanPed - 0.005)"
        SQLX$ = SQLX$ + " ORDER BY NUORITEM "
        Dim PEDDETTEMP As ADODB.Recordset
        Set PEDDETTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
        With PEDDETTEMP
            Do While Not .EOF
              'CODIGO Y CANTIDAD DE RECORDSET
              CI% = XVALO(!CODIINT)
              CANTI = XVALO(!CANPED) - XVALO(!CANTENT)
              If CANTI < 0.05 Then
                 Call COMUNI("Artículo : " & CODEXT$(CI%) & " Falta Cantidad ")
                 Exit Sub
              End If
              '
              'DATOS DE ARCHIVO MAESTROS SEGUN ARTICULO
              CONDI$ = "CODINT = " & CI%
              CANTIDPORCAJA = XVALO(VALOBADA("MASTER", "PIEZASXCAJA", CONDI$, "NOMESS"))
              M3CAX = XVALO(VALOBADA("MASTER", "M3CAJA", CONDI$, "NOMESS"))
              KILNETCAJ = XVALO(VALOBADA("MASTER", "KILONETOCAJA", CONDI$, "NOMESS"))
              KILBRCAJ = XVALO(VALOBADA("MASTER", "KILOBRUTOCAJA", CONDI$, "NOMESS"))
              '
              If CANTIDPORCAJA < 1 Then
                 Call MENSERR(24, "Artículo : " & CODEXT$(CI%) & " Falta Cantidad por Caja")
                 Exit Sub
              End If
              CANTICAJAS% = 0
              Resto = CANTI
              Do While Resto > 0
                If MECUBPORPALET + M3CAX > MAXMECUBPORPALET Then
                   ' ABRO UN NUEVO PALLET PORQUE LA CAJA ADICIONAL NO CABE
                   i& = i& + 1: If i& > ULTREG&("!LISTAEMP") Then i& = 1 + ULTREG&("!LISTAEMP")
                   XX$ = REGBLAN$("LISTAEMP")
                   Call GRAREG("!LISTAEMP", XX$, i&)
                   NUMERODEPALET% = NUMERODEPALET% + 1
                   MECUBPORPALET = 0
                   CANTICAJAS% = 0
                End If
                '
                CANDELACAJA = CANTIDPORCAJA
                If CANDELACAJA > Resto Then 'MODIFICADO POR QUE EL REMANENTE
                   CANDELACAJA = Resto       ' DEBE CREAR UN NUEVO REGISTRO
                   If MECUBPORPALET > 0 Then
                     CANTICAJAS% = 0
                     i& = i& + 1: If i& > ULTREG&("!LISTAEMP") Then i& = 1 + ULTREG&("!LISTAEMP")
                     XX$ = REGBLAN$("LISTAEMP")
                     Call GRAREG("!LISTAEMP", XX$, i&)
                   End If
                End If
                If i& > IANTE& + 1 Then i& = IANTE& + 1
                CANTICAJAS% = CANTICAJAS% + 1
                XX$ = REGLEIDO$("!LISTAEMP", i&)
                Call REPLA(XX$, MKI$(NUMERODEPALET%), 1, 2)
                Call REPLA(XX$, MKI$(CI%), 3, 2)
                Call REPLA(XX$, MKS$(CANDELACAJA), 5, 4)
                Call REPLA(XX$, MKI$(CANTICAJAS%), 9, 2)
                Call REPLA(XX$, MKS$(NPED&), 11, 4)
                Call REPLA(XX$, MKI$(0), 15, 2)
 
                Call GRAREG("!LISTAEMP", XX$, i&)
                IANTE& = i&
                MECUBPORPALET = MECUBPORPALET + M3CAX 'AGREGADO FALTABA ACUMULAR
                Resto = Resto - CANDELACAJA ' AGREGADO POR QUE NO DESCONTABA Y ENTRABA EN UN BUCLE
              Loop
              i& = i& + 1: If i& > ULTREG&("!LISTAEMP") Then i& = 1 + ULTREG&("!LISTAEMP")
              '
            .MoveNext
            Loop
        End With
              
'''''              MECUBPORPALET = MECUBPORPALET + M3TOT ' ACUMULO M3
'''''
'''''              'CALCULO CANTIDAD DE CAJAS SEGUN CANTIDAD DE UNIDADES
'''''              CAJAEXTRA% = 0
'''''              CANTICAJAS% = Int(CANTI / CANTIDXCAJA&)
'''''              RESTO& = CANTI - (CANTICAJAS% * CANTIDXCAJA&)
'''''              If RESTO& > 0 Then
'''''                CAJAEXTRA% = 1
'''''              End If
'''''              '
'''''              'CALCULO DE M3 Y KILOS
'''''              M3TOT = CANTICAJAS% * M3CAX
'''''              PESNETOT = CANTICAJAS% * KILNETCAJ
'''''              PESBRUTOT = CANTICAJAS% * KILBRCAJ
'''''              TOTPIE = CANTIDXCAJA& * CANTICAJAS%
'''''              '
'''''              'VALIDO SI NO EXEDE EL LIMITE DE VOLUMEN POR PALLET
'''''              If MECUBPORPALET + M3TOT > MAXMECUBPORPALET Then
'''''                NUMERODEPALET = NUMERODEPALET + 1
'''''                MECUBPORPALET = 0
'''''              End If
'''''              MECUBPORPALET = MECUBPORPALET + M3TOT ' ACUMULO M3
'''''              '
'''''              'CALCULO DE DESDE Y HASTA CAJAS
'''''              DESDEX% = HASTAX% + 1
'''''              If HASTAX% < 1 Then HASTAX% = 1
'''''              HASTAX% = DESDEX% + CANTICAJAS% - 1
'''''              '
'''''              I& = I& + 1
'''''              XX$ = REGBLAN$("!LISTAEMP")
'''''              Call REPLA(XX$, MKI$(NUMERODEPALET%), 1, 2)
'''''              Call REPLA(XX$, MKI$(CI%), 3, 2)
'''''              Call REPLA(XX$, MKS$(CANTIDXCAJA&), 5, 4)
'''''              Call REPLA(XX$, MKI$(CANTICAJAS%), 9, 2)
'''''              Call REPLA(XX$, MKS$(NPED&), 11, 4)
'''''              'Call REPLA(XX$, MKI$(NUORPL%), 15, 2)
'''''              Call GRAREG("!LISTAEMP", XX$, I&)
'''''              '
'''''              ' SI HAY EXCEDENTE VA EN CAJA EXTRA Y GENERO NUEVO REGISTRO PARA ESTAS
'''''              If CAJAEXTRA% > 0 Then
'''''
''''''                'CALCULO CANTIDAD DE CAJAS SEGUN CANTIDAD DE UNIDADES
''''''                M3TOT = CAJAEXTRA% * M3CAX
''''''                PESNETOT = ROUND#((KILNETCAJ / CANTIDXCAJA&) * RESTO&, 2)
''''''                PESBRUTOT = ROUND#(PESNETOT + (KILBRCAJ - KILNETCAJ), 2)
''''''                TOTPIE = RESTO&
'''''                '
'''''                'VALIDO SI NO EXEDE EL LIMITE DE VOLUMEN POR PALLET
'''''                If MECUBPORPALET + M3TOT > MAXMECUBPORPALET Then
'''''                  NUMERODEPALET = NUMERODEPALET + 1
'''''                  MECUBPORPALET = 0
'''''                End If
'''''                MECUBPORPALET = MECUBPORPALET + M3TOT ' ACUMULO M3
'''''
'''''                'CALCULO DE DESDE Y HASTA CAJAS
'''''                DESDEX% = HASTAX% + 1
'''''                If HASTAX% < 1 Then HASTAX% = 1
'''''                HASTAX% = DESDEX% + CAJAEXTRA% - 1
'''''                '
'''''                'GRABO EN EL ARCHIVO
'''''                I& = I& + 1
'''''                XX$ = REGBLAN$("!LISTAEMP")
'''''                Call REPLA(XX$, MKI$(NUMERODEPALET%), 1, 2)
'''''                Call REPLA(XX$, MKI$(CI%), 3, 2)
'''''                Call REPLA(XX$, MKS$(RESTO&), 5, 4)
'''''                Call REPLA(XX$, MKI$(CAJAEXTRA%), 9, 2)
'''''                Call REPLA(XX$, MKS$(NPED&), 11, 4)
'''''                'Call REPLA(XX$, MKI$(NUORPL%), 15, 2)
'''''                Call GRAREG("!LISTAEMP", XX$, I&)
'''''
'''''              End If
        '
        'ASIGNO LOS CAMPOS NO EDITABLES
        ATRI$ = "/TITUPAN=LISTAEMP/TITUPAN=Lista de Empaque Nro.: " & NUMLEMP&
        ATRI$ = ATRI$ + "/POSEDIT(1)=2"
        ATRI$ = ATRI$ + "/POSEDIT(2)=2"
        ATRI$ = ATRI$ + "/POSEDIT(3)=2"
        ATRI$ = ATRI$ + "/POSEDIT(6)=2"
        '
        Call VENTALIEMPA(NULPQUE&, ATRI$)
        GoTo 15
        '
   End Sub




     '
     
 Sub VENTALIEMPA(NUMLEMP&, Optional ATRI$)
    '
    'VALIDO SI NO ESTA CERRADA
    If LEMPCER%(NUMLEMP&) > 0 Then
       Call COMUNI("Lista de Empaque Cerrada Imposible Editar")
       Exit Sub
    End If
    '
    'PRESENTO VENTANA
10  VTB% = VENTABU%("LISTAEMP" + ATRI$)
    If VTB% < 0 Then Exit Sub
    '
    'PERMITE EDITAR PACKING LIST
    If COMALTER%("¿Desea Modificar el Packing List Activo?\(Deberá Tener el Privilegio Correspondiente)\\No, Continuar\Sí, Modificar") = 2 Then
        If DERACCE%("EDIPALIS", "Edición de Packing por Pedido") >= 2 Then
            Call MODIFICAPACKLIST(OK%, NUMLEMP&)
            If OK% < 1 Then Exit Sub
            GoTo 25
        Else
            GoTo 10
        End If
    End If
   
    'VALIDACION DE DATOS
    '
    NPALLANT% = 0
    For U& = 1 To ULTREG&("!LISTAEMP")
       XY$ = REGLEIDO$("!LISTAEMP", U&)
       NPALL% = CVI(Mid$(XY$, 1, 2))
       CI% = CVI(Mid$(XY$, 3, 2))
       CANTXCAJA = CVS(Mid$(XY$, 5, 4))
       CANTICAJAS% = CVI(Mid$(XY$, 9, 2))
       NUPED& = CVS(Mid$(XY$, 11, 4))
       CONDI$ = "CODINT = " & CI%
       CANPXCJA& = XVALO(VALOBADA("MASTER", "PIEZASXCAJA", CONDI$, "NOMESS"))
       '
       'VALIDO QUE TENGA TODOS LOS DATOS DEL ARCHIVO MAESTRO
       If FALTANDATOS%(CI%) > 0 Then
          Call MENSERR(24, "Para el Código: " & CODEXT$(CI%) & " Falta Datos Del Archivo Maestro\Deberá Ingresarlos Antes del Cierre")
       End If
       '
       If CANTXCAJA > CANPXCJA& Then Call COMUNI("Piezas x Caja  No Válido para Registro " & U&): Exit Sub
       If NPALL% < 1 Then Call MENSERR(24, "Número de Pall No Válido para Registro " & U&): Exit Sub
       '
       If NPALL% < NPALLANT% Or NPALL% > NPALLANT% + 1 Then
          Call MENSERR(24, "Los Pallets Deben ir en Número de\Orden Creciente y Correlativo !!!")
          Exit Sub
       End If
       '
       If CI% < 1 Then Call MENSERR(24, "Número de Código No Válido para Registro " & U&): Exit Sub
       If CANTICAJAS% < 1 Then Call MENSERR(24, "Número de Cajas No Válido para Registro " & U&): Exit Sub
       '
       'VALIDO EL NUMERO DE PEDIDO SI SE INGRESO
       If NUPED& > 0 Then
          NC% = XVALO(VALOBADA("ENCEMPAQ", "CODI_CLI", " NumLtaEmp = " & NUMLEMP&, " NOMESS"))
          CONDI$ = "NROPED = " & NUPED& & "  AND CODIINT = " & CI% & " And Status < 8 And NROCLIE = " & NC%
          CANTRE& = CantRegistros("PEDDETA", CONDI$)
          If CANTRE& < 1 Then
             Call MENSERR(24, "El Artículo " & CODEXT$(CI%) & " No Corresponde al Pedido: " & NUPED&)
             Exit Sub
          End If
       End If
       NPALLANT% = NPALL%
       '
       'CALCULO Y GUARDO EL PESO BRUTO
       CONDI$ = "CODINT = " & CI%
       KILBRCAJ = XVALO(VALOBADA("MASTER", "KILOBRUTOCAJA", CONDI$, "NOMESS"))
       XY$ = REGLEIDO$("!LISTAEMP", U&)
       Call REPLA(XY$, MKS(KILBRCAJ), 17, 4)
       Call GRAREG("LISTAEMP", XY$, U&)
       '
    Next U&
    '
    Call SETULTREG("LISTAEMP", U& - 1)
    'ORDENO EL ARCHIVO
    Call FILESORT("LISTAEMP", "", 8, 8, "DESC")
    Call FILESORT("LISTAEMP", "", 1, 1, "ASC")
    '
    Call TRALOCAL("LISTAEMP", "")
    '
25  CONDI$ = " NumLtaEmp = " & NUMLEMP&
    '
    Call ACTUREGISTRO("LISEMPAQ", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")
    '
    NUORPL% = 0
    DESDEX% = 0
    HASTAX% = 0
    '
    SQLX$ = "SELECT * FROM LISEMPAQ "
    SQLX$ = SQLX$ + " WHERE NumLtaEmp < 1 "
    Dim LIMPTMP1 As ADODB.Recordset
    Set LIMPTMP1 = New ADODB.Recordset
26  On Error Resume Next
    LIMPTMP1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
     On Error GoTo 0
     Call CapturaErrorOpen
     GoTo 26
    End If
    On Error GoTo 0
    '
    For U& = 1 To ULTREG&("!LISTAEMP")
       '
       ' LEEMOS DATOS DE LA GRILLA
       XY$ = REGLEIDO$("!LISTAEMP", U&)
       NPALL% = CVI(Mid$(XY$, 1, 2))
       'If NPALLANT% <> NPALL% Then NUORPL% = 0: DESDEX% = 0: HASTAX% = 0
       NUORPL% = NUORPL% + 1
       NPALL% = CVI(Mid$(XY$, 1, 2))
       CI% = CVI(Mid$(XY$, 3, 2))
       CANTXCAJA = CVS(Mid$(XY$, 5, 4))
       CANTICAJAS% = CVI(Mid$(XY$, 9, 2))
       NUPED& = CVS(Mid$(XY$, 11, 4))
       '
       ' LEEMOS DATOS DEL MAESTRO
       CONDI$ = "CODINT = " & CI%
       'CANTIDXCAJA& = XVALO(VALOBADA("MASTER", "PIEZASXCAJA", CONDI$, "NOMESS"))
       M3CAX = XVALO(VALOBADA("MASTER", "M3CAJA", CONDI$, "NOMESS"))
       KILNETCAJ = XVALO(VALOBADA("MASTER", "KILONETOCAJA", CONDI$, "NOMESS"))
       KILBRCAJ = XVALO(VALOBADA("MASTER", "KILOBRUTOCAJA", CONDI$, "NOMESS"))
       '
       'CALCULO
       M3TOT = CANTICAJAS% * M3CAX
       PESNETOT = CANTICAJAS% * KILNETCAJ
       PESBRUTOT = CANTICAJAS% * KILBRCAJ
       TOTPIE = CANTXCAJA * CANTICAJAS%
'       VOLUTOT = M3TOT
       '
       DESDEX% = HASTAX% + 1
       If HASTAX% < 1 Then HASTAX% = 1
       HASTAX% = DESDEX% + CANTICAJAS% - 1
       '
       'GRABO
       With LIMPTMP1
          .AddNew
          On Error Resume Next
          !CODIEMPR = CodiEmp%
          On Error GoTo 0
          !NumLtaEmp = NUMLEMP&
          !Numpallet = NPALL%
          !NUORDPALLET = NUORPL%
          !DESDECAJA = DESDEX%
          !HASTACAJA = HASTAX%
          '
          !CODINT = CI%
          !CANTICAJAS = CANTICAJAS%
          !PiezasXCaja = CANTXCAJA
          !PIEZASTOTAL = TOTPIE
          '
          !M3TOTAL = M3TOT
          !M3Caja = M3CAX
          '
          !KILONETOCAJA = KILNETCAJ
          !KILONETOTOT = PESNETOT
          !KILOBRUTOCAJA = KILBRCAJ
          !KILOBRUTOTOT = PESBRUTOT
          '
          !MARBORRA = 0
          !NROPED = NUPED&
          '
          .Update
       End With
       NPALLANT% = NPALL%
       
    Next U&
    '
    LIMPTMP1.Close
    Set LIMPTMP1 = Nothing
    '
    'BORRO LOS DATOS DE LA TABAL MARCADOS ANTERIORMENTE
    CONDI$ = " NumLtaEmp = " & NUMLEMP& & "  AND  MARBORRA = 1"
    Call BORRAREGISTROS("LISEMPAQ", CONDI$, REGAF&, "NOMESS")
    '
    Call COMUNI("Lista de Empaque Guardada")
End Sub


Private Sub Command19_Click()
   Command19.Enabled = False
   Call FINAL("*CODORCL")
   Command19.Enabled = True
End Sub

Private Sub Command2_Click()
    '
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command20_Click()

   Command20.Enabled = False
   Call CONECRUN("COCUME09", "Control de Cumplimiento de Entregas")
   Command20.Enabled = True

End Sub

Private Sub Command21_Click()
   '\\\OT-05-0831-WAR-17/11/05///
   Call MENSERR(24, "nNo Disponible en la Presente Versión")
   Exit Sub
   '
   '
   Command21.Enabled = False
   IMPOREM07.Show 1
99 Command21.Enabled = True
   '
End Sub

Private Sub Command22_Click()
   Command22.Enabled = False
   
10 OPX% = COMALTER%("Opciones de Picking:\\Picking de Remitos de Venta\Picking de Remitos de Consiganción")
   If OPX% < 1 Or OPX% > 2 Then GoTo 99
   
   If OPX% = 1 Then
        PICKDES09.Show 1
        GoTo 10
     ElseIf OPX% = 2 Then
        PICKCON09.Show 1
        GoTo 10
   End If

99 Command22.Enabled = True
End Sub

Private Sub Command23_Click()
    '
    Command23.Enabled = False
10  OPX% = ALTERNA%("Recibir desde Central Avellaneda\Enviar Informacion a Central Avellaneda\Recepcion de Materiales\Lugares de Entrega\Transfiere Datos a SQL\Carpeta de Sincronización")
    If OPX% = 0 Then GoTo 99
    Select Case OPX%
        Case 1: Call RECIBECEN: GoTo 10 'lee el atrasuc
        Case 2: Call GENARTRA: GoTo 10
        Case 3: Call RECECOMO: GoTo 10 ' aca ingresa solo los que no se han recibido
        Case 4: Call LUGARENTREGA: GoTo 10
        Case 5: Call TRANSFESQL: GoTo 10
        Case 6: Call CARPETADEINTERCAMBIO: GoTo 10
    End Select
    '
99  Command23.Enabled = True
    '
End Sub

Private Sub Command24_Click()
    '
    Dim vectorcodigos() As String * 128
    Dim resumenentrega(9) As Long
    Dim remitosdespachados() As String
    '
    NCLIE = 3
5   Call DESHASFECHA(Des%, Has%, PERIO$)
    If TRIM$(PERIO$) = "" Then Exit Sub
    Desde = FETEXCOR1$(Des%)
    Hasta = FETEXCOR1$(Has%)
    CANTREG& = CantRegistros("movisto", "codimovi='rt' and doprilar like 'rt%' and fechmov>='" & Desde & "' and fechmov<='" & Hasta & "' and (nroenvio=0 or nroenvio is null) and nume_clie=" & NCLIE)
    If CANTREG& = 0 Then
        Call COMUNI("No se Han Encontrado Remitos Pendientes de Despachar")
        GoTo 5
    End If
    '
    DESREMB07.List1.Clear
    SQLX$ = "select distinct doprilar, fechmov from movisto with(nolock) where codimovi='rt' and doprilar like 'rt%' and fechmov>='" & Desde & "' and fechmov<='" & Hasta & "' and (nroenvio=0 or nroenvio is null) and nume_clie=" & NCLIE & " order by doprilar asc"
    Set rstseleccion = READSET(SQLX$)
    Do While Not rstseleccion.EOF
        Remito$ = SAFETEXT$(rstseleccion!DOPRILAR)
            FF% = CVINT(rstseleccion!FECHMOV)
        FECHA$ = FECHATEX$(FF%)
        '
        TXT$ = FECHA$ & " - " & Remito$
        DESREMB07.List1.AddItem TXT$
        '
        rstseleccion.MoveNext
    Loop
    rstseleccion.Close
    Set rstseleccion = Nothing
    DESREMB07.Show 1
    '
    CONDI$ = "("
    For IX& = 0 To DESREMB07.List1.ListCount - 1
        If DESREMB07.List1.Selected(IX&) Then
            Remito$ = Mid$(DESREMB07.List1.List(IX&), 12)
            CONDI$ = CONDI$ & "doprilar='" & Remito$ & "' or "
        End If
    Next IX&
    CONDI$ = Mid$(CONDI$, 1, Len(CONDI$) - 3) & ")"
    '
    ' *** validaciones
    SQLX$ = "select * from movisto with(nolock) where codimovi='rt' and doprilar like 'rt%' and fechmov>='" & Desde & "' and fechmov<='" & Hasta & "' and (nroenvio=0 or nroenvio is null) and nume_clie=" & NCLIE & " AND " & CONDI$ & " order by doprilar asc"
    Set rstvalida = READSET(SQLX$)
    Do While Not rstvalida.EOF
        CIXI% = XVALO(rstvalida!cod_inte)
        qbultos& = LINDIVEN(CIXI%)
        If qbultos& = 0 Then
            Call COMUNI("Imposible Continuar. Falta Cantidad de Bultos en Código '" & CODEXT$(CIXI%) & "'")
            GoTo 99
        End If
        CODORI$ = TRIM$(CODORIG$(CIXI%, CInt(NCLIE)))
        If CODORI$ = CODEXT$(CIXI%) Then
            Call COMUNI("Imposible Continuar. Falta Código Mercedes Para el Código '" & CODEXT$(CIXI%) & "'")
            GoTo 99
        End If
        rstvalida.MoveNext
    Loop
    '
    rstvalida.Close
    Set rstvalida = Nothing
    '
    SQLX$ = "select * from movisto with(nolock) where codimovi='rt' and doprilar like 'rt%' and fechmov>='" & Desde & "' and fechmov<='" & Hasta & "' and (nroenvio=0 or nroenvio is null) and nume_clie=" & NCLIE & " AND " & CONDI$ & " order by doprilar asc"
    Set Rst = READSET(SQLX$)
    remitoant$ = ""
    ReDim Preserve vectorcodigos(2)
    contadorremito& = 0
    '
    ' *** registro encabezado
    II& = 0
    vectorcodigos(II&) = Space(128)
    Call REPLA(vectorcodigos(II&), "&&##C115620146#ADC0MB155#BOWESA#Z", 1, 40)
    '
    ' *** registro 711
    II& = 1
    vectorcodigos(II&) = Space(128)
    Call REPLA(vectorcodigos(II&), "71103", 1, 5)
        nrocliemb$ = NPROCLI$(3)
        While Len(nrocliemb$) < 9: nrocliemb$ = "0" + nrocliemb$: Wend
        nrocliefv$ = "3"
        While Len(nrocliefv$) < 9: nrocliefv$ = "0" + nrocliefv$: Wend
    Call REPLA(vectorcodigos(II&), nrocliemb$, 6, 9)
    Call REPLA(vectorcodigos(II&), nrocliefv$, 15, 9)
        ultnroenvio$ = SAFETEXT$(XVALO(VALOBADA("movisto", "nroenvio", "nroenvio>0")))
        While Len(ultnroenvio$) < 5: ultnroenvio$ = "0" + ultnroenvio$: Wend
        proxnroenvio$ = CStr(XVALO(ultnroenvio$) + 1)
        While Len(proxnroenvio$) < 5: proxnroenvio$ = "0" + proxnroenvio$: Wend
    Call REPLA(vectorcodigos(II&), ultnroenvio$, 24, 5)
    Call REPLA(vectorcodigos(II&), proxnroenvio$, 29, 5)
        FECHA = FECHATEX$(HOY(HO$))
        fechaenvio$ = Mid$(FECHA, 7, 2) & Mid$(FECHA, 4, 2) & Mid$(FECHA, 1, 2)
    Call REPLA(vectorcodigos(II&), fechaenvio$, 34, 6)
    '
    ' *** registro 712
    II& = 2
    vectorcodigos(II&) = Space(128)
    Call REPLA(vectorcodigos(II&), "71203", 1, 5)
    Call REPLA(vectorcodigos(II&), "155", 14, 3)
    '
    With Rst
        Do While Not .EOF
            '
            Remito$ = SAFETEXT$(!DOPRILAR)
            If Remito$ <> remitoant$ Then
                contadorremito& = contadorremito& + 1
                ReDim Preserve remitosdespachados(contadorremito&)
                remitosdespachados(contadorremito&) = Remito$
                '
                ' *** registro 715
                If remitoant$ <> "" Then
                    findetalle& = II&
                    For IX& = iniciodetalle& To findetalle&
                        II& = II& + 1
                        ReDim Preserve vectorcodigos(II&)
                        vectorcodigos(II&) = Space(128)
                        Call REPLA(vectorcodigos(II&), "71503", 1, 5)
                        Call REPLA(vectorcodigos(II&), "2032", 6, 4)
                            codigoart$ = TRIM$(Mid$(vectorcodigos(IX&), 28, 22))
                            CIXI% = CODINT%(codigoart$)
                                qantitxt$ = Mid$(vectorcodigos(IX&), 53, 10) & "." & Mid$(vectorcodigos(IX&), 63, 3)
                            qanti# = XVALO(qantitxt$)
                            cantiporbulto& = LINDIVEN(CIXI%)
                            totalbultos& = Int(qanti# / cantiporbulto&)
                            Resto# = qanti# - (cantiporbulto& * totalbultos&)
                            If Resto# > 0 Then
                               totalbultos& = totalbultos& + 1
                            End If
                            totalbultostxt$ = CStr(totalbultos&)
                            While Len(totalbultostxt$) < 13: totalbultostxt$ = "0" + totalbultostxt$: Wend
                        Call REPLA(vectorcodigos(II&), totalbultostxt$, 50, 13)
                     Next IX&
                End If
                '
                ' *** registro 713
                II& = II& + 1
                ReDim Preserve vectorcodigos(II&)
                vectorcodigos(II&) = Space(128)
                    remitoant$ = SAFETEXT$(!DOPRILAR)
                Call REPLA(vectorcodigos(II&), "71303", 1, 5)
                Call REPLA(vectorcodigos(II&), Mid$(Remito$, 11, 8), 6, 8)
                    FECHA = FECHATEX$(CVINT%(!FECHMOV))
                    fechaenvio$ = Mid$(FECHA, 7, 2) & Mid$(FECHA, 4, 2) & Mid$(FECHA, 1, 2)
                Call REPLA(vectorcodigos(II&), fechaenvio$, 14, 6)
                Call REPLA(vectorcodigos(II&), "266", 20, 3)
                Call REPLA(vectorcodigos(II&), "20", 25, 2)
                iniciodetalle& = II& + 1
            End If
            '
            II& = II& + 1
            ReDim Preserve vectorcodigos(II&)
            vectorcodigos(II&) = Space(128)
            Call REPLA(vectorcodigos(II&), "71403", 1, 5)
                CIXI% = XVALO(!cod_inte)
            Call REPLA(vectorcodigos(II&), CODORIG$(CIXI%, CInt(NCLIE)), 6, 22) ' hay que reemplazar por el de mercedes
            Call REPLA(vectorcodigos(II&), CODEXT$(CIXI%), 28, 22)
            Call REPLA(vectorcodigos(II&), "999", 50, 3)
                qanti# = XVALO(!CANTSALID)
                cantitxt$ = TRIM$(FORMATNUM$(qanti#, "F13.3"))
                cantitxt$ = REPLACAR(cantitxt$, ".", "")
                While Len(cantitxt$) < 13: cantitxt$ = "0" + cantitxt$: Wend
            Call REPLA(vectorcodigos(II&), cantitxt$, 53, 13)
                nrorden$ = CStr(XVALO(!NuOrItem))
                While Len(nrorden$) < 3: nrorden$ = "0" + nrorden$: Wend
            Call REPLA(vectorcodigos(II&), "ST", 81, 2)
            Call REPLA(vectorcodigos(II&), nrorden$, 87, 3)
            Call REPLA(vectorcodigos(II&), "S", 106, 1)
            Call REPLA(vectorcodigos(II&), "X1", 115, 2)
            '
            .MoveNext
        Loop
    End With
    '
    On Error Resume Next
    Rst.Close
    Set Rst = Nothing
    On Error GoTo 0
    '
    ' *** registro 715
    findetalle& = II&
    For IX& = iniciodetalle& To findetalle&
        II& = II& + 1
        ReDim Preserve vectorcodigos(II&)
        vectorcodigos(II&) = Space(128)
        Call REPLA(vectorcodigos(II&), "71503", 1, 5)
        Call REPLA(vectorcodigos(II&), "2032", 6, 4)
            codigoart$ = TRIM$(Mid$(vectorcodigos(IX&), 28, 22))
            CIXI% = CODINT%(codigoart$)
                qantitxt$ = Mid$(vectorcodigos(IX&), 53, 10) & "." & Mid$(vectorcodigos(IX&), 63, 3)
            qanti# = XVALO(qantitxt$)
            cantiporbulto& = LINDIVEN(CIXI%)
            totalbultos& = Int(qanti# / cantiporbulto&)
            Resto# = qanti# - (cantiporbulto& * totalbultos&)
            If Resto# > 0 Then
               totalbultos& = totalbultos& + 1
            End If
            totalbultostxt$ = CStr(totalbultos&)
            While Len(totalbultostxt$) < 13: totalbultostxt$ = "0" + totalbultostxt$: Wend
        Call REPLA(vectorcodigos(II&), totalbultostxt$, 50, 13)
     Next IX&
     '
     ' *** registro 719
     For IX& = 1 To II&
        codigoreg% = XVALO(Mid$(vectorcodigos(IX&), 1, 3))
        Select Case codigoreg%
            Case 711
                    resumenentrega(1) = resumenentrega(1) + 1
            Case 712
                    resumenentrega(2) = resumenentrega(2) + 1
            Case 713
                    resumenentrega(3) = resumenentrega(3) + 1
            Case 714
                    resumenentrega(4) = resumenentrega(4) + 1
            Case 715
                    resumenentrega(5) = resumenentrega(5) + 1
            Case 716
                    resumenentrega(6) = resumenentrega(6) + 1
            Case 717
                    resumenentrega(7) = resumenentrega(7) + 1
            Case 718
                    resumenentrega(8) = resumenentrega(8) + 1
        End Select
     Next IX&
     '
     II& = II& + 1
     ReDim Preserve vectorcodigos(II&)
     vectorcodigos(II&) = Space(128)
     Call REPLA(vectorcodigos(II&), "71902", 1, 5)
     For IX& = 1 To 8
        totalcodigotxt$ = CStr(resumenentrega(IX&))
        While Len(totalcodigotxt$) < 7: totalcodigotxt$ = "0" + totalcodigotxt$: Wend
        Call REPLA(vectorcodigos(II&), totalcodigotxt$, (IX& - 1) * 7 + 5, 7)
     Next IX&
     Call REPLA(vectorcodigos(II&), "0000001", (IX& - 1) * 7 + 5, 7)
     '
     ' *** pie de archivo
     II& = II& + 1
     ReDim Preserve vectorcodigos(II&)
     vectorcodigos(II&) = Space(128)
     Call REPLA(vectorcodigos(II&), "&&##", 1, 4)
     '
     For IX& = 0 To II&
        Debug.Print vectorcodigos(IX&)
     Next IX&
     '
     ' *** guarda archivo y actualiza movisto con el numero de envio
     OPX% = COMALTER%("¿Desea Generar el Archivo de Despachos?\\Si, Generar\No, Cancelar")
     If OPX% = 1 Then
        For IX& = 1 To contadorremito&
            Remito$ = remitosdespachados(IX&)
            CONDI$ = "doprilar='" & Remito$ & "'"
            Call ACTUREGISTRO("movisto", "nroenvio", CONDI$, XVALO(proxnroenvio$), REG&)
        Next IX&
        '
        CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
        CommonDialog1.DialogTitle = "Archivo Origen de Datos"
        CommonDialog1.Filter = "Archivo de Texto(*.TXT)|*.TXT|"
        CommonDialog1.FilterIndex = 1
        On Error GoTo 99
        CommonDialog1.ShowOpen
        On Error GoTo 0
        DoEvents
        NOMBASALID$ = UCase$(CommonDialog1.FileName)
        ASALID1$ = NOMBASALID$
        Screen.MousePointer = 11
        NX1% = FILEOPEN%(ASALID1$, 2, 128)
        For IX& = 0 To II&
            TX$ = vectorcodigos(IX&)
            Print #NX1%, TX$
        Next IX&
        Close #NX1%
        ASALID2$ = REPLACAR(ASALID1$, "\", "/")
        Screen.MousePointer = 1
        Call COMUNI("Archivos de Transferencia\ '" + ASALID2$ + "' \Generado con Correctamente")
     End If
     '
99   ' *** fin de rutina
     On Error Resume Next
     rstvalida.Close
     Set rstvalida = Nothing
     On Error GoTo 0

End Sub

Sub Command25_Click()
    Command25.Enabled = False
    '
    StringCodySerie$ = "" 'VARIABLE PARA GENERAR UN STRING CON CODIGOS Y LOTES PARA LUEGO REALIZAR UN ACSALOTE
    If DERACCE("ANUREMI", "Marcar Sin Cargo Remitos") < 2 Then GoTo 99

    EsReteplast% = 0
    If InStr(1, EMPREAC$, "RETEPLAST") > 0 Then EsReteplast% = 1
    
    CMO$ = TRIM$(Control$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
       CMO$ = "RT"
    End If
    '
    Call CARSELMESA
    '
10  Screen.MousePointer = 1
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) = "" Then
       GoTo 99
    End If
    CORRESQ$ = MESANO$(MESASEL$, Des%, Has%)
    '
    Call FRESHECHO("!SELREMI")
    Screen.MousePointer = 11
    '
    'DES1$ = CVDAT(DES%)
    'HAS1$ = CVDAT(HAS%)
    DES1$ = FETEXCOR1$(Des%)
    HAS1$ = FETEXCOR1$(Has%)
    TABLAMOVI$ = "MOVISTO"
    ULTIVALO& = 0
    '
12  SQLX$ = " SELECT DISTINCT DoPriCor, Nume_Clie, FechMov, ReferCor FROM " & TABLAMOVI$ ' MOVISTO "
    SQLX$ = SQLX$ + " WHERE FechMov >= '" & DES1$ & "' "
    SQLX$ = SQLX$ + " AND FechMov <= '" & HAS1$ & "' "
    If TRIM$(TABLAMOVI$) = "MOVISTO" Then
        SQLX$ = SQLX$ + " AND (CodiMovi =  '" & CMO$ & "' "
        SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
        SQLX$ = SQLX$ + " AND CantSalid > 0 "
    Else
        SQLX$ = SQLX$ + " AND (CodiMovi = '*R') "
    End If
    'SQLX$ = SQLX$ + "ORDER BY val(MID$(DoPriCor,4)) asc "
    Call APERBASDAT("STOCKS")
    'Set MOVITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim MOVITMP As ADODB.Recordset
    Set MOVITMP = New ADODB.Recordset
    MOVITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

    With MOVITMP
      On Error Resume Next
      .MoveFirst
      If Err And ULTIVALO& < 1 Then
        Call MENSERR(24, " No Existen Registros que Cumplan\con la Condición SQL.")
        On Error GoTo 0
        GoTo 10
      End If
      On Error GoTo 0
      '
      J& = ULTIVALO&: TXCONTROL$ = ""
      Do While (.EOF = False)
        If Left$(UCase$(!ReferCor), 4) <> "ANUL" Then   ' para evitar comprobantes arruinados
          NCLI = XVALO(!Nume_Clie)
          REMI$ = SAFETEXT$(!DoPriCor)
          If TRIM$(REMI$) <> "" Then
            If InStr(TXCONTROL$, REMI$) < 1 Then
              FF% = CVINT(!FECHMOV)
              '
              Z$ = Space$(40) + String$(8, 0)
'              Call REPLA(Z$, REMI$, 1, 12)
'                RASOX$ = TRIM$(RASOCLI$(NCLI))
'                If RASOX$ = "" Then RASOX$ = "CLIENTE Nro." + TRIM$(Str$(NCLI))
'              Call REPLA(Z$, RASOX$, 13, 26)
'              Call REPLA(Z$, MKS$(U&), 39, 4)
'              Call REPLA(Z$, MKI$(FF%), 43, 2)
'              Call REPLA(Z$, MKS$(NCLI), 45, 4)
              
              Call REPLA(Z$, REMI$, 1, 10)
                RASOX$ = TRIM$(rasocli$(NCLI))
                If RASOX$ = "" Then RASOX$ = "CLIENTE Nro." + TRIM$(Str$(NCLI))
              Call REPLA(Z$, RASOX$, 11, 28)
              Call REPLA(Z$, MKS$(U&), 39, 4)
              Call REPLA(Z$, MKI$(FF%), 43, 2)
              Call REPLA(Z$, MKS$(NCLI), 45, 4)
              J& = J& + 1
              Call GRAREG("!SELREMI", Z$, J&)
              TXCONTROL$ = TXCONTROL$ + AJUSTI$(REMI$, 24)
            End If
          End If
        End If
      .MoveNext
      Loop
    End With
    Set MOVITMP = Nothing
    '
    Call SETULTREG("!SELREMI", J&)
    Call CIERRARCH("!SELREMI")
    Call FRESHECHO("!SELREMI")
    Screen.MousePointer = 1
    '
    If TABLAMOVI$ = "MOVISTO" Then
        TABLAMOVI$ = "MOVIREPA"
        ULTIVALO& = J&
        GoTo 12
    End If
    '
    If J& < 1 Then
       Call MENSERR(24, "No se Encontraron Remitos Posibles de Anular\en el Período Elegido.")
       GoTo 10
    End If
    '
    RERE$ = REGSEL$("!SELREMI")
    If Len(RERE$) <= 4 Then GoTo 10
    '
    REMISEL$ = TRIM$(Left$(RERE$, 10))
    CLIE$ = Mid$(RERE$, 11, 28)
    FEXI% = CVI(Mid$(RERE$, 43, 2))
    FEMOV = FETEXCOR1$(FEXI%)
    CLISE = CVS(Mid$(RERE$, 45, 4))
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "ENVAPLAST") > 0 Then
        REMILARX$ = TRIM$(VALOBADA("MOVISTO", "DOPRILAR", "DoPriCor = '" & REMISEL$ & "'", "NOMESS"))
        NXL& = CLISE
        NCLX$ = FORMATOCON0$(NXL&, 7)
        Dim ENVA As New Envap
        YAFACTUWALD% = ENVA.REMITO_YA_FACTU_EN_WALDBOTT%(REMILARX$)
        Set ENVA = Nothing
        If YAFACTUWALD% > 0 Then
           Call COMUNI("Atención !!! Imposible Anular Remito\El Mismo Está Facturado en Sistema Waldbott")
           GoTo 10
        End If
    End If
    
'    REMISEL$ = TRIM$(Left$(RERE$, 12))
'    CLIE$ = Mid$(RERE$, 11, 26)
'    FEXI% = CVI(Mid$(RERE$, 43, 2))
'    FEMOV = FETEXCOR1$(FEXI%)
'    CLISE = CVS(Mid$(RERE$, 45, 4))
    
    COMOV$ = TRIM$(VALOBADA("MOVISTO", "CODIMOVI", "DoPriCor = '" & REMISEL$ & "'", "NOMESS"))
    If COMOV$ = "" Then COMOV$ = TRIM$(VALOBADA("MOVIREPA", "CODIMOVI", "DoPriCor = '" & REMISEL$ & "'", "NOMESS"))
    
    '
    If InStr(1, EMPREAC$, "NEUQUEN") > 0 Or InStr(1, EMPREAC$, "METALURGICASIAM-ARCON") < 1 And (InStr(1, EMPREAC$, "ENVAPLAST") < 1 And InStr(1, EMPREAC$, "CONTEPLAST") < 1 And InStr(1, EMPREAC$, "CAMBACERES") < 1) Then 'SI NO ES ARCON COMODORO O NEUQUEN
      TXT$ = "Remito Número " + TRIM$(REMISEL$) + "\Cliente: " + TRIM$(CLIE$) + "\Fecha: " + FECHATEX$(FEXI) + "\   \\Abandonar\Anulación Reutilizable\Anular con Reserva"
      ALTE% = COMALTER%(TXT$)
      If (ALTE% < 2 Or ALTE% > 3) Then GoTo 10
    Else
      'EN ESTE CASO ANULA SOLO CON RESERVA
      TXT$ = "Confirmar Anulación del Remito " + TRIM$(REMISEL$) + "\Cliente: " + TRIM$(CLIE$) + "\Fecha: " + FECHATEX$(FEXI) + "\   \\Abandonar\Anular (Con Reserva de Número)"
      ALTE% = COMALTER%(TXT$)
      If ALTE% < 2 Then
        GoTo 10
      Else
        ALTE% = 3 ' LE ASIGNA 3 POR QUE ES EN REALIDAD LA OPCION RESERVA DE NUMERO
      End If
    End If
    '

    '\\\OT-5-0400-FG-14/06/05///
    Call APERBASDAT("STOCKS")
    TABLAMOVI$ = "MOVISTO"
    If COMOV$ = "*R" Then TABLAMOVI$ = "MOVIREPA"
30  SQLX$ = " SELECT * FROM " & TABLAMOVI$ 'MOVISTO "
    SQLX$ = SQLX$ + " WHERE DoPriCor = '" & REMISEL$ & "' "
    SQLX$ = SQLX$ + " AND Nume_Clie = " & CLISE & " "
    SQLX$ = SQLX$ + " AND FechMov = '" & FEMOV & "'"
    'Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
    Dim MoviTemp As ADODB.Recordset
    Set MoviTemp = New ADODB.Recordset
    MoviTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    ENTROEN_RS_CAJABANC% = 0
    With MoviTemp
      On Error Resume Next
      .MoveFirst
      If Err > 1 Then
        On Error GoTo 0
        Call MENSERR(24, "No se Encontró Detalle del Remito")
        GoTo 10
      Else
        'RECORRE EL DETALLE Y GENERA UN STRING DE LOS TRAZABLES PARA REALIZAR UN ACSALOTE AL FINAL SI SE GRABA BIEN LA ANULACION
        VUELTA% = 0
        Do While Not .EOF
           CI1X% = XVALO(!cod_inte)
           If ESTRAZABLE%(CI1X%) > 0 Then
             LOTE$ = TRIM$(SAFETEXT$(!idenlote))
             If LOTE$ <> "" Then
               If VUELTA% = 0 Then
                StringCodySerie$ = CI1X% & "|" & LOTE$
                VUELTA% = 1
               Else
                StringCodySerie$ = StringCodySerie$ & ";" & CI1X% & "|" & LOTE$
                ''EJ.:StringCodySerie$ = "325|OF-000012-1;325|OF-000012-2;455|OF-000045-1"
               End If
            End If
          End If
          .MoveNext
        Loop
      End If
      '
      On Error GoTo 0
      ' VERIFICA SI EL REMITO ESTA FACTURADO
      Do While Not .EOF
        TIDOCSEC$ = ""
        On Error Resume Next
        TIDOCSEC$ = UCase$(Mid$(TRIM$(SAFETEXT$(!DoSecCor)), 1, 2))
        On Error GoTo 0
        If TIDOCSEC$ = "FC" Then
          NUMERFAC$ = TRIM$(SAFETEXT$(!DoSecCor))
          Call APERBASDAT("CABAN")
          Screen.MousePointer = 11
          'CajaBanc.FindFirst "TipoMovim = 'FVEN' AND CodiCom_Cor = '" & NUMERFAC$ & "'"
          SQLX$ = " SELECT * FROM CajaBanc "
          SQLX$ = SQLX$ + "WHERE TipoMovim = 'FVEN'"
          SQLX$ = SQLX$ + " AND CodiCom_Cor = '" & NUMERFAC$ & "'"
          Dim CAJJJABANC As ADODB.Recordset
          Set CAJJJABANC = New ADODB.Recordset
          CAJJJABANC.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
          With CAJJJABANC
            ENTROEN_RS_CAJABANC% = 1
            On Error Resume Next
            .MoveFirst
            'If CajaBanc.NoMatch = False Then
            If Err < 1 Then
              On Error GoTo 0
              If (XVALO(!Idebecomp) + XVALO(!IHABECOMP)) < 0.005 Then GoTo 16
            End If
            On Error GoTo 0
            Call COMUNI("Imposible Anular. Remito Ya Facturado\Mediante Comprobante " & NUMERFAC$ & "." _
                      & "\   \Debe Anular la Factura y Re-Intentar.")
            
          End With
          GoTo 10
        End If
      .MoveNext
      Loop
    End With
    If ENTROEN_RS_CAJABANC% = 1 Then
        CAJJJABANC.Close
        Set CAJJJABANC = Nothing
        ENTROEN_RS_CAJABANC% = 0
    End If
    '
16  Call APERBASDAT("STOCKS")
    Screen.MousePointer = 11
    SQLX$ = "SELECT * FROM " & TABLAMOVI$ ' MOVISTO "
    SQLX$ = SQLX$ + " WHERE DoPriCor = '" & REMISEL$ & "' "
    SQLX$ = SQLX$ + " AND Nume_Clie = " & CLISE & " "
    SQLX$ = SQLX$ + " AND FechMov = '" & FEMOV & "'"
    Dim MOVVVISTO As ADODB.Recordset
    Set MOVVVISTO = New ADODB.Recordset
    MOVVVISTO.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

'20  MoviTemp.FindFirst SQLX$
    
'21  If MoviTemp.NoMatch = False Then
      ' VERIFICA SI EL REMITO TIENE ASOCIADO UN PEDIDO
    With MOVVVISTO
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While (Not .EOF)
          NROPEDID& = XVALO(!NUPEDCLI)
          RemitoLargo$ = SAFETEXT$(!DOPRILAR)
          If NROPEDID& > 0 Then
          ' ACTUALIZA LOS PEDIDOS QUE FUERON ENTREGADOS CON EL REMITO, EL PROBLEMA ES QUE CUANDO SE REMITE MAS DE LO DEL PEDIDO
          ' NO ACTUALIZA EN EL PEDIDO Y SI SE MODIFICA SI SE REALIZAN ENTREGAS PARCIALES Y ANULACIONES
          ' LA CANTIDAD ENTREGADA REAL NO VA A COINCIDIR CON EL PENDIENTE.
            CODINTE% = XVALO(!cod_inte)
            CANTSALD# = XVALO(!CANTSALID)
            Call APERBASDAT("PEDCLI")
            SQLY$ = "SELECT * FROM PEDDETA "
            SQLY$ = SQLY$ + " where NroPed = " & NROPEDID& & " "
            SQLY$ = SQLY$ + "AND CodiInt = " & CODINTE% & " "
            SQLY$ = SQLY$ + "AND NroClie = " & CLISE & " "
            SQLY$ = SQLY$ + "AND CantEnt >= " & FORMATNUM$((CANTSALD# - 0.05), "F15.2") & " "
            SQLY$ = SQLY$ & " AND Status < 9 "
            Dim PEDETTTA As ADODB.Recordset
            Set PEDETTTA = New ADODB.Recordset
            'MsgBox sqly$
            PEDETTTA.Open FILTSQL$(SQLY$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
  
            'PedDeta.FindFirst SQLY$
            With PEDETTTA
              On Error Resume Next
              .MoveFirst
              'If PedDeta.NoMatch = False Then
              If Err < 1 Then
                On Error GoTo 0
                'PedDeta.Edit
                ENTREGA# = XVALO(!CANTENT) - CANTSALD#
                SALDO# = XVALO(!CanSaldo) + CANTSALD#
                !CANTENT = ENTREGA#
                !CanSaldo = SALDO#
                !Status = 0 'SI SE ESTA INGRESANDO POR ANULACION DE REMITO CAMBIA EL STATUS A 0 PARA QUITAR MARCA DE CUMPLIDO
                !NROREMI = ""   'ESTO SE MODIFICO POR QUE SI LUEGO DE ANULAR EL REMITO SE ANULA EL PEDIDO, SI EL STATUS ESTA EN 3 NO LO CONSIDERA
                .Update
              End If
              On Error GoTo 0
            End With
            Set PEDETTTA = Nothing
          End If
          .MoveNext
        Loop
      End If
      On Error GoTo 0
      '
    End With
    Set MOVVVISTO = Nothing
    '
    'BLANQUEA EL NUMERO DE REMITO,FECHA, CLIENTE EN TANUMSE DE LOS EQUIPOS QUE SALIERON CON EL REMITO
    'BORRA EN EL PEDDETA EL NRO DE SERIE Y REGISTRA EN EL ANOTADOR
    '******QUEDA LIBERADO EL NUMERO DE SERIE******
    Call ANULARSALIDANUMEROSERIESEGUNREMITO(RemitoLargo$, CLISE, FEMOV)

    'METALES ARCON - GENERA ARCHIVO CON LA INFORMACION QUE LUEGO SE TRANSFIERE A AVELLANEDA
    If InStr(1, EMPREAC$, "NEUQUEN") > 0 Or InStr(1, EMPREAC$, "METALURGICASIAM-ARCON") > 0 Then
       Call ACTU_ENTRECOM(REMISEL$, CLISE, FEMOV) ' GENERA UN MOVIMIENTO INVERSO EN LA TABLA ENRTRECOM
       ALTE% = 3 'SOLO ANULACION RE-UTILIZABLE
    End If

    If ALTE% = 2 Then
'        SQLX$ = "DELETE FROM " & TABLAMOVI$ ' MOVISTO "
'        SQLX$ = SQLX$ + " WHERE DoPriCor = '" & REMISEL$ & "' "
'        SQLX$ = SQLX$ + " AND Nume_Clie = " & CLISE & " "
'        SQLX$ = SQLX$ + " AND FechMov = '" & FEMOV & "'"
        On Error Resume Next 'SE CAMBIO EL DELETE COMENTADO ARRIBA P/Q  AL REALIZAR UN DELETE QUE AFECTA
                             'MAS DE UN REGISTRO DA ERROR POR EL TRIGGER DE ESTE MODO VA ELIMINANDO DE UN REGISTRO
        SQLX$ = "select DoPriCor,Nume_Clie,FechMov FROM " & TABLAMOVI$ ' MOVISTO "
        SQLX$ = SQLX$ + " WHERE DoPriCor = '" & REMISEL$ & "' "
        SQLX$ = SQLX$ + " AND Nume_Clie = " & CLISE & " "
        SQLX$ = SQLX$ + " AND FechMov = '" & FEMOV & "'"
        Dim RST1 As ADODB.Recordset
        Set RST1 = New ADODB.Recordset
        RST1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        With RST1
          Do While Not .EOF
            .Delete
            .MoveNext
          Loop
        End With
        RST1.Close
        Set RST1 = Nothing
        On Error GoTo 0
    Else
        SQLX$ = "UPDATE " & TABLAMOVI$ & " SET CANTSALID=0, REFERCOR='(Anul) ' + REFERCOR"
        If InStr(EPAC$, "CONTEPLAST") > 0 Or InStr(EPAC$, "CAMBACERES") > 0 Then
          SQLX$ = SQLX$ + " , PESOKG = 0 "
          SQLX$ = SQLX$ + " , METROS = 0 "
        End If

        SQLX$ = SQLX$ + " WHERE DoPriCor = '" & REMISEL$ & "' "
        SQLX$ = SQLX$ + " AND Nume_Clie = " & CLISE & " "
        SQLX$ = SQLX$ + " AND FechMov = '" & FEMOV & "'"
    End If
    '
    If InStr(EPAC$, "CONTEPLAST") > 0 Then
      'EN DOCUORIG ESTA GUARDADO EL NRO DE REMITO (ES EL REGISTRO QUE SE GENERA CUANDO SE CIERRA LA OF, SE GUARDA CUANDO SE REMITE
      RemitoLargo$ = TRIM$(VALOBADA("MOVISTO", "DOPRILAR", "DoPriCor = '" & REMISEL$ & "'", "NOMESS"))
      CONDI$ = "DocuOrig = '" & RemitoLargo$ & "'"
      Call ACTUREGISTRO("MOVISTO", "DocuOrig", CONDI$, "", REGAF&, "NOMESS")
    End If
    
    ' *** LIBERA PEDIDO / LISTA DE DESPACHO ASOCIADA AL REMITO
    If EsReteplast% = 1 Then
        RemitoLargo$ = TRIM$(VALOBADA("MOVISTO", "DOPRILAR", "DoPriCor = '" & REMISEL$ & "'", "NOMESS"))
        Call ACTUREGISTRO("bultosfabric", "remito", "remito='" & RemitoLargo$ & "'", "", REGAF&)
    End If
    '
    If REMILARX$ <> "" Then
        NXL& = CLISE
        NCLX$ = FORMATOCON0$(NXL&, 7)
        REMILARX$ = FORMATO_RTO_WALDB(REMILARX$)
        sqlx1$ = "DELETE FROM PDCABEZA"
        sqlx1$ = sqlx1$ + " WHERE NROPED = '" & REMILARX$ & "' AND TIPO = 'R' AND CLIENTE = '" & NCLX$ & "'"
        
        SQLX2$ = "DELETE FROM PDITEMS"
        SQLX2$ = SQLX2$ + " WHERE NROPED = '" & REMILARX$ & "' AND TIPO = 'R' AND CLIENTE = '" & NCLX$ & "'"
        Call abreBaDaWald
        On Error Resume Next
        BASEWALD.BeginTrans
        BASEWALD.Execute sqlx1$
        BASEWALD.Execute SQLX2$
        If Err.Number <> 0 Then
          BASEWALD.RollbackTrans
          Call COMUNI("Atención No se Pudo Eliminar Correctamente\El Remito " & REMISEL$ & " Del Cliente " & NCLX$ & "\De La Base de Datos de Waldbott\Informe esta Situación a su Soporte de Sistema")
        Else
          BASEWALD.CommitTrans
        End If
    End If
    '
    
    If ALTE% <> 2 Then  ' SI ALTE% ES IGUAL A 2 LO EJECUTA EN LA CONSULTA DIRECTAMENTE
      On Error Resume Next
      FLEXCONN.Execute SQLX$
    End If
    
    If Err Then
         Call MENSERR(24, "No fue Posible Completar la Anulación\del Remito '" + REMISEL$ + "'")
       Else
         CONDI$ = "Dopricor = '" & REMISEL$ & "'"
         Call ACTUREGISTRO("DETACARTEM", "Cant_Entre", CONDI$, 0, REGAF&)
         Call ACTUREGISTRO("DETACARTEM", "DopriLar", CONDI$, "", REGAF&)
         Call ACTUREGISTRO("DETACARTEM", "DopriCor", CONDI$, "", REGAF&)
         Call ACTUREGISTRO("DETACARTEM", "FechEmis", CONDI$, CVDAT(0), REGAF&)
         Call ACTUREGISTRO("DETACARTEM", "NuorItem", CONDI$, 0, REGAF&)
         Call ACSALOTE_STRING_CODIGO_Y_LOTE(StringCodySerie$) ' RUTINA ACTUALIZA EL LOTINGRE
         Call COMUNI("Remito " + REMISEL$ + " Anulado.")

    End If
    GoTo 10
    '
    Call ACTUALIZAR_TABLA_MAXREMITO_CUANDO_CAMBIA(0) ' SOLO REFRESCA POR QUE LE PASO FECHA 0
    '


    
99  Call CIERRARCH("*.*")
    Call BAJALDISCO
    '
    FOREMIT07.Label10 = ""
    PREFIREM$ = TRIM$(Control$("", "PREFIREM"))
    If PREFIREM$ <> "*" Then
      FOREMIT07.Label10 = PREFIREM$
    End If
    NROPRO& = 0
    Call NUREMIFA(NROPRO&, PREFIREM$)
    FOREMIT07.Text11.TEXT = CMO$ + "." + TRIM$(Str$(NROPRO&))
    '
    Screen.MousePointer = 1
    Command25.Enabled = True
    '
End Sub

Sub ACTU_ENTRECOM(REMISEL$, CLISE, FEMOV)
        SQLX$ = "SELECT * FROM MOVISTO WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE DoPriCor = '" & REMISEL$ & "' "
        SQLX$ = SQLX$ + " AND Nume_Clie = " & CLISE & " "
        SQLX$ = SQLX$ + " AND FechMov = '" & FEMOV & "'"
        Set RS2 = READSET(SQLX$)
          Do While Not RS2.EOF
            SQLX$ = "SELECT * FROM ENTRECOM WHERE IDOCUM < 1"
            Dim RST1 As ADODB.Recordset
            Set RST1 = New ADODB.Recordset
26:                 RST1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
            If Err Then
             On Error GoTo 0
             Call CapturaErrorOpen
             GoTo 26
            End If
            On Error GoTo 0
            '
            With RST1
                .AddNew
                !FECHA = CVDAT(HOY(HO$))
                !DOCUMENTO = REMISEL$
                !TIPODOC = "DEVOLUCION"
                !nume_cli = CLISE
                NROO& = XVALO(Mid$(REMISEL$, 4))
                IDOCUM& = 10000& * (FF% Mod 1000) + 100& * (NROO& Mod 100)
                !IDOCUM = IDOCUM&
                !cod_inte = XVALO(RS2!cod_inte)
                !Cod_Exte = SAFETEXT$(RS2!Cod_Exte)
                !Cantidad = XVALO(RS2!CANTSALID) * (-1)
                !Status = 1
                !VECESENVIADO = 0
                .Update
            End With
            '
            On Error Resume Next
            RST1.Close
            Set RST1 = Nothing
            On Error GoTo 0
            RS2.MoveNext
          Loop
          RS2.Close
          Set RS2 = Nothing
End Sub
Private Sub Command26_Click()
    '
    Command26.Enabled = False
    '
2   OPX% = COMALTER%("Opciones varias:\\Generar Archivo de Despachos\Re-generar Archivo de Despachos Existente\Eliminar Envío\Configuración de Cliente")
    '
    ' *** lee numero de cliente
    NCLIE = XVALO(Control("REMITO", "DESPAMB"))
    If NCLIE = 0 Then
        Call COMUNI("Imposible Continuar. Debe Primero Configurar la Cuenta de Mercedes Benz")
        GoTo 2
    End If
    '
    If OPX% = 0 Then GoTo 99
    If OPX% = 2 Then GoTo 50
    If OPX% = 3 Then GoTo 60
    If OPX% = 4 Then GoTo 70
    '
    ' *** genera archivo de despacho
5   Call DESHASFECHA(Des%, Has%, PERIO$)
    If TRIM$(PERIO$) = "" Then GoTo 2
    Desde = FETEXCOR1$(Des%)
    Hasta = FETEXCOR1$(Has%)
    CANTREG& = CantRegistros("movisto", "codimovi='rt' and doprilar like 'rt%' and fechmov>='" & Desde & "' and fechmov<='" & Hasta & "' and (nroenvio=0 or nroenvio is null) and nume_clie=" & NCLIE)
    If CANTREG& = 0 Then
        Call COMUNI("No se Han Encontrado Remitos Pendientes de Despachar")
        GoTo 5
    End If
    '
    DESREMB07.List1.Clear
    SQLX$ = "select distinct doprilar, fechmov from movisto with(nolock) where codimovi='rt' and doprilar like 'rt%' and fechmov>='" & Desde & "' and fechmov<='" & Hasta & "' and (nroenvio=0 or nroenvio is null) and nume_clie=" & NCLIE & " order by doprilar asc"
    Set rstseleccion = READSET(SQLX$)
    Do While Not rstseleccion.EOF
        Remito$ = SAFETEXT$(rstseleccion!DOPRILAR)
            FF% = CVINT(rstseleccion!FECHMOV)
        FECHA$ = FECHATEX$(FF%)
        '
        TXT$ = FECHA$ & " - " & Remito$
        DESREMB07.List1.AddItem TXT$
        '
        rstseleccion.MoveNext
    Loop
    rstseleccion.Close
    Set rstseleccion = Nothing
    DESREMB07.Show 1
    '
    CONDI$ = "(": ELIGIOALGUNO% = 0
    For IX& = 0 To DESREMB07.List1.ListCount - 1
        If DESREMB07.List1.Selected(IX&) Then
            Remito$ = Mid$(DESREMB07.List1.List(IX&), 12)
            CONDI$ = CONDI$ & "doprilar='" & Remito$ & "' or "
            ELIGIOALGUNO% = 1
        End If
    Next IX&
    If ELIGIOALGUNO% = 0 Then
        Call COMUNI("Imposible Continuar. No Se Han Elegido Remitos")
        GoTo 5
    End If
    CONDI$ = Mid$(CONDI$, 1, Len(CONDI$) - 3) & ")"
    '
    ' *** validaciones
    SQLX$ = "select * from movisto with(nolock) where codimovi='rt' and doprilar like 'rt%' and fechmov>='" & Desde & "' and fechmov<='" & Hasta & "' and (nroenvio=0 or nroenvio is null) and nume_clie=" & NCLIE & " AND " & CONDI$ & " order by doprilar asc"
    Set rstvalida = READSET(SQLX$)
    Do While Not rstvalida.EOF
        CIXI% = XVALO(rstvalida!cod_inte)
        qbultos& = LINDIVEN(CIXI%)
        If qbultos& = 0 Then
            Call COMUNI("Imposible Continuar. Falta Cantidad de Bultos en Código '" & CODEXT$(CIXI%) & "'")
            GoTo 2
        End If
        CODORI$ = TRIM$(CODORIG$(CIXI%, CInt(NCLIE)))
        If CODORI$ = CODEXT$(CIXI%) Then
            Call COMUNI("Imposible Continuar. Falta Código Mercedes Para el Código '" & CODEXT$(CIXI%) & "'")
            GoTo 2
        End If
        rstvalida.MoveNext
    Loop
    '
    rstvalida.Close
    Set rstvalida = Nothing
    '
    SQLX$ = "select * from movisto with(nolock) where codimovi='rt' and doprilar like 'rt%' and fechmov>='" & Desde & "' and fechmov<='" & Hasta & "' and (nroenvio=0 or nroenvio is null) and nume_clie=" & NCLIE & " AND " & CONDI$ & " order by doprilar, NuorItem asc"
    Call ArmaArchivoDeDespachos(SQLX$, 0, NCLIE)
    GoTo 2
    ' *** fin de rutina
    '
    ' *** regenera archivo de despacho de un envio
50  Call CARGALISEL("LIENVIO", "MOVISTO", "FHor_Verif/D8;NROENVIO/I16", "NROENVIO>0", "DISTINCT")
55  Call SELECHO("LIENVIO/Lista de Envios")
    nroenvio& = XVALO(VALACT2$("LIENVIO"))
    If nroenvio& = 0 Then GoTo 2
    SQLX$ = "select * from movisto with(nolock) where nroenvio=" & nroenvio& & " order by doprilar, nuoritem asc"
    Call ArmaArchivoDeDespachos(SQLX$, nroenvio&, NCLIE)
    GoTo 55
    '
    ' *** Elimina un envio
60  Call CARGALISEL("LIENVIO", "MOVISTO", "FHor_Verif/D8;NROENVIO/I16", "NROENVIO>0", "DISTINCT")
65  Call SELECHO("LIENVIO/Lista de Envios")
    nroenvio& = XVALO(VALACT2$("LIENVIO"))
    nroenviomax& = XVALO(VALOBADA("movisto", "max(nroenvio)", "codimovi='rt'"))
    If nroenvio& = 0 Then
        GoTo 2
    Else
        If nroenvio& < nroenviomax& Then
            Call COMUNI("Imposible Anular un Envio Existiendo Otros Posteriores.\Debe Primero Eliminar los más Recientes")
            GoTo 65
        End If
        OPX% = COMALTER%("¿Esta Seguro Que Desea Eliminar el Envio Nro. " & CStr(nroenvio&) & "?\\No, Cancelar\Si, Continuar")
        If OPX% = 1 Then GoTo 2
    End If
    CONDI$ = "nroenvio=" & nroenvio&
    Call ACTUREGISTRO("movisto", "nroenvio", CONDI$, "0", REG&)
    GoTo 60
    '
    ' ***
70  nclieregistrado = XVALO(Control("REMITO", "DESPAMB"))
    If nclieregistrado = 0 Then
        Call SELECHO("deudor1")
        codcuentatxt$ = VALACT1$("deudor1")
        If codcuentatxt$ = "" Then GoTo 2
        Call GRACONTROL("REMITO", "DESPAMB", codcuentatxt$)
        Call COMUNI("Configuración Registrada")
    Else
        RASOC$ = rasocli$(nclieregistrado)
        OPX% = COMALTER%("¿Desea Cambiar La Cuenta Actualmente Configurada\" & CStr(nclieregistrado) & " - " & Mid$(RASOC$, 1, 32) & "?\\No, Mantener\Si, Cambiar")
        If OPX% = 2 Then
            Call SELECHO("deudor1")
            codcuentatxt$ = VALACT1$("deudor1")
            If codcuentatxt$ = "" Then GoTo 2
            Call GRACONTROL("REMITO", "DESPAMB", codcuentatxt$)
            Call COMUNI("Configuración Registrada")
        End If
    End If
    GoTo 2
    '
99  Command26.Enabled = True
    '
End Sub

Private Sub Command27_Click()
   Command27.Enabled = False
     
     REMIDATO.Show 1
     GoTo 99
     Screen.MousePointer = 11
     CMO$ = TRIM$(Control$("", "CODREFAC"))
     If Len(CMO$) <> 2 Then
        CMO$ = "RT"
     End If
     '
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
     CONDI$ = "FechMov >= '" & Desde & "'"
     CONDI$ = CONDI$ + "AND FechMov <= '" & Hasta & "'"
     CONDI$ = CONDI$ + " AND substring(REFERCOR,1,4) <> 'ANUL' " 'PARA EVITAR COMPROBANTES ARRUINADOS
     CONDI$ = CONDI$ + " AND CodiMovi =  '" & CMO$ & "'"
     CONDI$ = CONDI$ + " AND substring(DOPRILAR,1,2) = 'RT' "
     '
     Screen.MousePointer = 11
     '
     Call COPYSTRU("SELDOCAP", "SELREMIT")
     Call CARGALISEL("SELREMIT", "MOVISTO", "DOPRILAR/A18;FechMov/D8;REFERCOR/A24", CONDI$, "DISTINCT")
     '
     If ULTREG&("SELREMIT") < 1 Then
        Call MENSERR(24, "No se Encontraron Remitos en el Rango Seleccionado.")
        GoTo 10
     End If

15   Call SELECHO("SELREMIT")
     '
     Screen.MousePointer = 1
     '
'     NUDOCU$ = TRIM$(Mid$(VALACT1$("SELREMIT"), 6))
'     If NUDOCU$ = "" Then GoTo 10
'     TIDOCUM$ = "Remito"
'     nubus = XVALO(Mid$(NUDOCU$, 6))
'     '
'     Screen.MousePointer = 11
'     '
'     DOCUNU& = 0
'     'LE PASO EN LA PARAMETRO DE CONTROL EL NOMESS PARA QUE SI NO ENCUENTRA EL DOCUMENTO
'     'NO PRESENTE MENSAJE Y CONTINUE Y LO BUSQUE DE LA MANERA QUE ESTABA ANTES
'     Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, Str(nubus), NUDOCU$, "NOMESS")
'
'     If DOCUNU& > 0 Then GoTo 15
'
'     '
'     For UI& = ULTREG&("PDOCLST") To 1 Step -1
'        XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
'        '\\\OT-08-0646-OD-20/08/08///
'         If InStr(XX$, TIDOCUM$) > 4 Then
'            If InStr(XX$, NUDOCU$) > 4 Then
'                PPXX% = InStr(XX$, "Nro.")
'                If PPXX% > 0 Then
'                 XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
'                End If
'                If InStr(XX$, NUDOCU$) > 4 Then
'                 DOCUNU& = CVS(Left$(XX$, 4))
'                 GoTo 50
'                End If
'            End If
'         End If
'     Next UI&
'     Screen.MousePointer = 1

     Call ReimpresionDeRemito(NUDOCU$, OK%)
     If OK% = 1 Then GoTo 15
     
40   Call COMUNI("No se Encontró el Documento Impreso")
     GoTo 15
     '
50   Call MUESTRADOC(DOCUNU&)
     GoTo 15
     '
     '
   
   
99 Command27.Enabled = True
End Sub

'
Sub Command3_Click()
    '
    Command3.Enabled = False
    '
    
    OPX% = COMALTER%("Formato de Salida:\\Detalle por Cliente\Acumulado por Cliente")
    If OPX% = 1 Then
       Call mnuDetxCliente_Click
    ElseIf OPX% = 2 Then
       mnuAcumxCliente_Click
    End If
    Command3.Enabled = True
    Exit Sub
    
    
    'codigo anterior
    
    ' detalle de entregas por cliente
    '
    CMO$ = TRIM$(Control$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
       CMO$ = "RT"
    End If
    '
    Dim PUNMO$(32767)
    '
10  FEX = HOY(HO$)
    Has% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    Des% = FECHANUM(DESS$)
    VDEF$ = MKI$(Des%) + MKI$(Has%)
    '
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    Call BLOQARCH("DTCLIEN")
    '
    Des% = CVI(Left$(OBX$, 2))
    Has% = CVI(Mid$(OBX$, 3, 2))
    If Des% > Has% Then
        GoTo 10
    End If
    FEX = Des%: PERIO$ = FECHATEX$(FEX)
    FEX = Has%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    Screen.MousePointer = 11
    Call HEADERS("DTCLIEN", "")
    Call HEADERS("DTCLIEN", "Periodo: " + PERIO$)
    '
    'nuevo
'    DES1# = CDbl(CVDAT(DES%))
'    HAS1# = CDbl(CVDAT(HAS%))
    DESD1 = FETEXCOR1$(Des%)
    HASD1 = FETEXCOR1$(Has%)
    '
    SQLX$ = " SELECT * FROM MOVISTO WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE FechMov >= '" & DESD1 & "' "
    SQLX$ = SQLX$ + " AND FechMov <= '" & HASD1 & "' "
     'PARA EVITAR COMPROBANTES ARRUINADOS
    SQLX$ = SQLX$ + " AND LEFT(REFERCOR,4) <> 'ANUL' "
    SQLX$ = SQLX$ + " AND (CodiMovi =  '" & CMO$ & "'  "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV')  "
    Call APERBASDAT("STOCKS")
    'Set MOVITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    Dim MOVITMP As ADODB.Recordset
    Set MOVITMP = New ADODB.Recordset
    MOVITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

    With MOVITMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, " No Existen Registros que Cumplan\con la Condición SQL.")
        On Error GoTo 0
        GoTo 10
      End If
      On Error GoTo 0
      '
      J& = 0
      REBLB$ = REGBLAN$("CONESREM")
      Do While (.EOF = False)
        NCLI = !Nume_Clie
        REMIFAC$ = !DoPriCor
        OBSERX$ = ""
        On Error Resume Next
        NUMEFAC$ = TRIM$(!DoSecCor)
          If Left$(NUMEFAC$, 2) <> "FC" Then
            NUMEFAC$ = ""
            If Left$(UCase$(!DOSECLAR), 9) = "SIN CARGO" Then
              OBSERX$ = !DOSECLAR
            End If
          End If
        On Error GoTo 0
        NPEDX& = !NUPEDCLI
        CANTU# = !CANTSALID - !CANTINGRE
        CII% = !cod_inte
        NOCOMX$ = NUOCOCLI$(NPEDX&, CII%)
        FF% = CVINT(!FECHMOV)
        AAA = FECHATEX$(FF%)
        Debug.Print AAA
        If Left$(REMIFAC$, 4) = "DV.0" Then
            REMIFAC$ = !DoSecCor
        End If
        '
        Z$ = REGBLAN$("DTCLIEN")
        Call REPLA(Z$, MKS$(NCLI), 1, 4)
        Call REPLA(Z$, MKI$(FF%), 5, 2)
        Call REPLA(Z$, REMIFAC$, 7, 10)
        Call REPLA(Z$, MKI$(CII%), 17, 2)
        Call REPLA(Z$, MKD$(CANTU#), 19, 8)
        Call REPLA(Z$, MKS$(!VALOUNIT), 27, 4)
          TOTI# = !VALOUNIT * (!CANTSALID - !CANTINGRE)
        Call REPLA(Z$, MKD$(TOTI#), 35, 8)
        J& = J& + 1
        Call GRAREG("DTCLIEN", Z$, J&)
        '
        Z$ = REBLB$
        Call REPLA(Z$, MKS$(NCLI), 1, 4)
        Call REPLA(Z$, MKI$(FF%), 5, 2)
        Call REPLA(Z$, REMIFAC$, 7, 12)
        Call REPLA(Z$, NUMEFAC$, 19, 12)
        Call REPLA(Z$, MKS$(NPEDX&), 31, 4)
        Call REPLA(Z$, NOCOMX$, 37, 16)
        Call REPLA(Z$, MKI$(CII%), 53, 2)
        Call REPLA(Z$, MKD$(CANTU#), 55, 8)
        Call REPLA(Z$, MKS$(!VALOUNIT), 63, 4)
        Call REPLA(Z$, MKD$(TOTI#), 71, 8)
        Call REPLA(Z$, OBSERX$, 87, 23)
        Call GRAREG("CNESREM", Z$, J&)
        '
        .MoveNext
      Loop
    End With
    Set MOVITMP = Nothing
    '
    Call SETULTREG("DTCLIEN", J&)
    Call FILESORT("DTCLIEN", "", 1, 5, "ASC")
    Call CIERRARCH("DTCLIEN")
    Call SETULTREG("CNESREM", J&)
    Call FILESORT("CNESREM", "", 1, 5, "ASC")
    Call CIERRARCH("CNESREM")
    '
    '
    OPX% = COMALTER%("Formato de Salida:\\Detalle por Cliente\Acumulado por Cliente\Control de Estado")
    If OPX% = 1 Then
        'Call FINAL("*DTACLI")
        Call FINAL("*DETACLI")
      ElseIf OPX% = 2 Then
        Call COPYSTRU("DTCLIEN", "ACUCLI")
        Call SETULTREG("ACUCLI", 0)
        NCLIANT = 0: CIXIANT% = 0: IMPO# = 0: TOTCANT# = 0: JJ& = 0
        Call FILESORT("DTCLIEN", "", 5, 5, "ASC")
        Call FILESORT("DTCLIEN", "", 1, 1, "ASC")
        '
        FIN& = ULTREG&("DTCLIEN")
        For II& = 1 To FIN&
            XX$ = REGLEIDO$("DTCLIEN", II&)
            NCLI = CVS(Mid$(XX$, 1, 4))
            CIXI% = CVI(Mid$(XX$, 17, 2))
            If CIXIANT = 0 And NCLIANT = 0 Then CIXIANT% = CIXI%
            If NCLIANT = 0 Then NCLIANT = NCLI
            '
            If NCLIANT = NCLI And CIXIANT% = CIXI% And II& < FIN& Then
                CANTI# = CVD(Mid$(XX$, 19, 8))
                VALO# = CVS(Mid$(XX$, 27, 4))
                TOTCANT# = TOTCANT# + CANTI#
                IMPO# = IMPO# + (CANTI# * VALO#)
            Else
                If II& = FIN& Then
                    CANTI# = CVD(Mid$(XX$, 19, 8))
                    VALO# = CVS(Mid$(XX$, 27, 4))
                    TOTCANT# = TOTCANT# + CANTI#
                    IMPO# = IMPO# + (CANTI# * VALO#)
                End If
                Z$ = REGBLAN$("DTCLIEN")
                Call REPLA(Z$, MKS$(NCLIANT), 1, 4)
                Call REPLA(Z$, MKI$(CIXIANT%), 17, 2)
                Call REPLA(Z$, MKD$(TOTCANT#), 19, 8)
                Call REPLA(Z$, MKD$(IMPO#), 35, 8)
                JJ& = JJ& + 1
                Call GRAREG("ACUCLI", Z$, JJ&)
                '
                CIXIANT% = CIXI%: NCLIANT = NCLI
                CANTI# = CVD(Mid$(XX$, 19, 8))
                VALO# = CVS(Mid$(XX$, 27, 4))
                TOTCANT# = CANTI#
                IMPO# = CANTI# * VALO#
            End If
        Next II&
        '
    Call SETULTREG("ACUCLI", JJ&)
    Call CIERRARCH("ACUCLI")
      
        Call FINAL("*ACUCLI")
      ElseIf OPX% = 3 Then
        Call FINAL("*CNESREM")
    End If
    '
99  Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    Command3.Enabled = True
    '
End Sub
Sub Command4_Click()
    '
    Command4.Enabled = False
    '
    ' detalle de entregas por artículo
    '
    CMO$ = TRIM$(Control$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
       CMO$ = "RT"
    End If
    '
10  FEX = HOY(HO$)
    Has% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    Des% = FECHANUM(DESS$)
    VDEF$ = MKI$(Des%) + MKI$(Has%)
    '
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    Des% = CVI(Left$(OBX$, 2))
    Has% = CVI(Mid$(OBX$, 3, 2))
    If Des% > Has% Then
        GoTo 10
    End If
    FEX = Des%: PERIO$ = FECHATEX$(FEX)
    FEX = Has%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    Screen.MousePointer = 11
    Call HEADERS("DETARTIN", "")
    Call HEADERS("DETARTIN", "Periodo: " + PERIO$)
    '
    'nuevo
'    DES1# = CDbl(CVDAT(DES%))
'    HAS1# = CDbl(CVDAT(HAS%))
    DESD1$ = CVDAT(Des%)
    HAST1$ = CVDAT(Has%)

    '
    SQLX$ = " SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE FechMov >= '" & DESD1$ & "' "
    SQLX$ = SQLX$ + " AND  FechMov <= '" & HAST1$ & "' "
     'PARA EVITAR COMPROBANTES ARRUINADOS
    SQLX$ = SQLX$ + " AND LEFT(REFERCOR,4) <> 'ANUL' "
    SQLX$ = SQLX$ + " AND (CodiMovi =  '" & CMO$ & "'  "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV')  "
    SQLX$ = SQLX$ + " ORDER BY COD_INTE"
    Call APERBASDAT("STOCKS")
    'Set MOVITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim MOVITMP As ADODB.Recordset
    Set MOVITMP = New ADODB.Recordset
    MOVITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

    With MOVITMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, " No Existen Registros que Cumplan\con la Condición SQL.")
        On Error GoTo 0
        GoTo 10
      End If
      On Error GoTo 0
      '
      J& = 0
      Do While (.EOF = False)
        NCLI = !Nume_Clie
        REMIFAC$ = !DoPriCor
        CANTU# = !CANTSALID - !CANTINGRE
        CII% = !cod_inte
        FF% = CVINT(!FECHMOV)
        If Left$(REMIFAC$, 4) = "DV.0" Then
            REMIFAC$ = !DoSecCor
        End If
        '
        Z$ = REGBLAN$("DETARTIN")
        Call REPLA(Z$, MKI$(CII%), 1, 2)
        Call REPLA(Z$, MKI$(FF%), 3, 2)
        Call REPLA(Z$, REMIFAC$, 5, 10)
        Call REPLA(Z$, MKS$(NCLI), 15, 4)
        Call REPLA(Z$, MKD$(CANTU#), 19, 8)
        J& = J& + 1
        Call GRAREG("DETARTIN", Z$, J&)
        .MoveNext
      Loop
    End With
    '


    Call SETULTREG("DETARTIN", J&)
    'Call FILESORT("DETARTIN", "DETARTIN", 1, 5, "ASC")
    Call CIERRARCH("DETARTIN")
    Call FINAL("*DETARTIN")
    '
99  Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    Command4.Enabled = True
    '
End Sub

Sub Command5_Click()
    '
    Command5.Enabled = False
    '
    CMO$ = TRIM$(Control$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
       CMO$ = "RT"
    End If
    '
10  FEX = HOY(HO$)
    Has% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    Des% = FECHANUM(DESS$)
    VDEF$ = MKI$(Des%) + MKI$(Has%)
    '
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    Des% = CVI(Left$(OBX$, 2))
    Has% = CVI(Mid$(OBX$, 3, 2))
    If Des% > Has% Then
        GoTo 10
    End If
    FEX = Des%: PERIO$ = FECHATEX$(FEX)
    FEX = Has%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    'nuevo
'    DES1# = CDbl(CVDAT(DES%))
'    HAS1# = CDbl(CVDAT(HAS%))
    DESD1$ = CVDAT(Des%)
    HAST1$ = CVDAT(Has%)

    '
    SQLX$ = " SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE FechMov >= '" & DESD1$ & "' "
    SQLX$ = SQLX$ + " AND FechMov <= '" & HAST1$ & "' "
    'PARA EVITAR COMPROBANTES ARRUINADOS
    SQLX$ = SQLX$ + " AND LEFT(REFERCOR,4) <> 'ANUL' "
    SQLX$ = SQLX$ + " AND (CodiMovi =  '" & CMO$ & "' "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV')  "
    SQLX$ = SQLX$ + "ORDER BY FechMov, DoPriCor ASC"
    Call APERBASDAT("STOCKS")
    'Set MOVITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    Dim MOVITMP As ADODB.Recordset
    Set MOVITMP = New ADODB.Recordset
    MOVITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

    With MOVITMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, " No Existen Registros que Cumplan\con la Condición SQL.")
        On Error GoTo 0
        GoTo 10
      End If
      On Error GoTo 0
      '
      Screen.MousePointer = 11
      J& = 0
      Do While (.EOF = False)
        J& = J& + 1
        '
        DOPRI$ = !DoPriCor
        FRI% = CVINT(!FECHMOV)
        NCLI = !Nume_Clie
        CIXI% = !cod_inte
        CAN# = !CANTSALID - !CANTINGRE
        '
        Z$ = REGBLAN$("LIREMIT")
        Call REPLA(Z$, DOPRI$, 1, 10)
        Call REPLA(Z$, MKI$(FRI%), 11, 2)
         RCCX$ = TRIM$(TEL2CLI$(NCLI))
          If RCCX$ = "" Then RCCX$ = rasocli$(NCLI)
        Call REPLA(Z$, RCCX$, 13, 10)
        Call REPLA(Z$, MKI$(CIXI%), 23, 2)
        Call REPLA(Z$, MKD$(CAN#), 25, 8)
        Call GRAREG("LIREMIT", Z$, J&)
       .MoveNext
      Loop
    End With
    '
    Call SETULTREG("LIREMIT", J&)
    Call CIERRARCH("LIREMIT")
    '
    Call HEADERS("LIREMIT", "")
    Call HEADERS("LIREMIT", "Periodo: " + PERIO$)
    '
    Screen.MousePointer = 1
    Call FILESORT("LIREMIT", "LIREMIT", 1, 4, "ASC")
    Call FINAL("*LIREMIT")
    '
99  Command5.Enabled = True
    '

'    Screen.MousePointer = 11
'    J& = 0
'    FIN& = ULTREG&("MATVEN")
'    '
'    For U& = 1 To FIN&
'        Call TRACE(20, U&, FIN&)
'        X$ = REGLEIDO$("MATVEN", U&)
'        FRI% = CVI(Mid$(X$, 7, 2))
'        If FRI% >= DES% Then
'            If FRI% <= HAS% Then
'                NC% = CVI(Left$(X$, 2))
'                DOPRI$ = "RT." + TRIM$(CVS(Mid$(X$, 23, 4)))
'                CIXI% = CVI(Mid$(X$, 9, 2))
'                CAN# = CVS(Mid$(X$, 19, 4))
'                '
'                J& = J& + 1
'                Z$ = REGBLAN$("LIREMIT")
'                Call REPLA(Z$, DOPRI$, 1, 10)
'                Call REPLA(Z$, MKI$(FRI%), 11, 2)
'                  RCCX$ = TRIM$(TEL2CLI$(NC%)) ' fantcli$(nc%)
'                  If RCCX$ = "" Then RCCX$ = RASOCLI$(NC%)
'                Call REPLA(Z$, RCCX$, 13, 10)
'                Call REPLA(Z$, MKI$(CIXI%), 23, 2)
'                Call REPLA(Z$, MKD$(CAN#), 25, 8)
'                Call GRAREG("LIREMIT", Z$, J&)
'            End If
'        End If
'    Next U&
'    '
'    FIN& = ULTREG&("MOVISTO")
'    For U& = 1 To FIN&
'        Call TRACE(20, U&, FIN&)
'        X$ = REGLEIDO$("MOVISTO", U&)
'        FRI% = CVI(Left$(X$, 2))
'        If FRI% >= DES% Then
'          If FRI% <= HAS% Then
'             If UCase$(Mid$(X$, 21, 2)) = CMO$ Then
'                NC% = CVI(Mid$(X$, 84, 2))
'                DOPRI$ = Mid$(X$, 23, 10)
'                CIXI% = CVI(Mid$(X$, 3, 2))
'                CAN# = CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))
'                '
'                J& = J& + 1
'                Z$ = REGBLAN$("LIREMIT")
'                Call REPLA(Z$, DOPRI$, 1, 10)
'                Call REPLA(Z$, MKI$(FRI%), 11, 2)
'                  RCCX$ = TRIM$(TEL2CLI$(NC%))
'                  If RCCX$ = "" Then RCCX$ = RASOCLI$(NC%)
'                Call REPLA(Z$, RCCX$, 13, 10)
'                Call REPLA(Z$, MKI$(CIXI%), 23, 2)
'                Call REPLA(Z$, MKD$(CAN#), 25, 8)
'                Call GRAREG("LIREMIT", Z$, J&)
'             End If
'          End If
'        End If
'    Next U&
'    '
'    Call COPYSTRU("MOVISTO", "MOVIREPA")
'    FIN& = ULTREG&("MOVIREPA")
'    For U& = 1 To FIN&
'        Call TRACE(20, U&, FIN&)
'        X$ = REGLEIDO$("MOVIREPA", U&)
'        FRI% = CVI(Left$(X$, 2))
'        If FRI% >= DES% Then
'          If FRI% <= HAS% Then
'            NC% = CVI(Mid$(X$, 84, 2))
'            DOPRI$ = Mid$(X$, 23, 10)
'            CIXI% = CVI(Mid$(X$, 3, 2))
'            CAN# = CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))
'            '
'            J& = J& + 1
'            Z$ = REGBLAN$("LIREMIT")
'            Call REPLA(Z$, DOPRI$, 1, 10)
'            Call REPLA(Z$, MKI$(FRI%), 11, 2)
'              RCCX$ = TRIM$(TEL2CLI$(NC%))
'              If RCCX$ = "" Then RCCX$ = RASOCLI$(NC%)
'            Call REPLA(Z$, RCCX$, 13, 10)
'            Call REPLA(Z$, MKI$(CIXI%), 23, 2)
'            Call REPLA(Z$, MKD$(CAN#), 25, 8)
'            Call GRAREG("LIREMIT", Z$, J&)
'          End If
'        End If
'    Next U&
'    '
'    Call SETULTREG("LIREMIT", J&)
'    Call CIERRARCH("LIREMIT")
'    '
'    Call HEADERS("LIREMIT", "")
'    Call HEADERS("LIREMIT", "Periodo: " + PERIO$)
'    '
'    Screen.MousePointer = 1
'    Call FINAL("*SOREMIT")
'    '
'99  Command5.Enabled = True
'    '
'End Sub
End Sub

Sub Command6_Click()
    '
    
    Command6.Enabled = False
    '
    Call NUEVO_COMAND6_CLICK
    GoTo 99
    '
    CMO$ = TRIM$(Control$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
       CMO$ = "RT"
    End If
    '
    Call CARSELMESA
    '
10  Screen.MousePointer = 1
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) = "" Then
       GoTo 99
    End If
    CORRESQ$ = MESANO$(MESASEL$, Des%, Has%)
    '
    Call FRESHECHO("!SELREMI")
    Screen.MousePointer = 11
    '
    'DES1$ = FECHATEX$(DES%)
    'HAS1$ = FECHATEX$(HAS%)
    DES1$ = FETEXCOR1$(Des%)
    HAS1$ = FETEXCOR1$(Has%)
    '
    SQLX$ = " SELECT DISTINCT DoPriCor, Nume_Clie, FechMov FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE FechMov >= '" & DES1$ & "' "
    SQLX$ = SQLX$ + " AND FechMov <= '" & HAS1$ & "' "
    SQLX$ = SQLX$ + " AND (CodiMovi =  '" & CMO$ & "'  "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV')  "
    'PARA EVITAR COMPROBANTES ARRUINADOS
    SQLX$ = SQLX$ + " AND LEFT(REFERCOR,4) <> 'ANUL' "
    SQLX$ = SQLX$ + " AND (DoSecLar IS NULL or (DoSecLar = '') OR (DoSecLar LIKE 'P/F%')) "
    '
    Call APERBASDAT("STOCKS")
    'Set MOVITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim MOVITMP As ADODB.Recordset
    Set MOVITMP = New ADODB.Recordset
    MOVITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

    With MOVITMP
      '
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, " No Existen Registros que Cumplan\con la Condición SQL.")
        On Error GoTo 0
        GoTo 10
      End If
      On Error GoTo 0
      '
      J& = 0
      Do While (.EOF = False)
        CLI% = !Nume_Clie
        REMI$ = !DoPriCor
        FF% = CVINT(!FECHMOV)
        '
        Z$ = Space$(40) + String$(8, 0)
        Call REPLA(Z$, REMI$, 1, 10)
        Call REPLA(Z$, rasocli$(CLI%), 11, 30)
        Call REPLA(Z$, MKS$(U&), 41, 4)
        Call REPLA(Z$, MKI$(FF%), 45, 2)
        Call REPLA(Z$, MKI$(CLI%), 47, 2)
        J& = J& + 1
        Call GRAREG("!SELREMI", Z$, J&)
      .MoveNext
      Loop
    End With
    Set MOVITMP = Nothing
    '
    Call SETULTREG("!SELREMI", J&)
    Call CIERRARCH("!SELREMI")
    Screen.MousePointer = 1
    '
    If J& < 1 Then
       Call MENSERR(24, "No se Encontraron Remitos Posibles de Anular\en el Período Elegido.")
       GoTo 10
    End If
    '
    RERE$ = REGSEL$("!SELREMI")
    If Len(RERE$) <= 4 Then GoTo 10
    '
    REMISEL$ = TRIM$(Left$(RERE$, 10))
    CLIE$ = Mid$(RERE$, 11, 30)
    FEXI% = CVI(Mid$(RERE$, 45, 2))
    FEX = FEXI%
    'FEMOVD = Int(CDbl(CVDAT(FEXI%)))
    FEMOV$ = CVDAT(FEXI%)
    CLISEL% = CVI(Mid$(RERE$, 47, 2))
    '
    TXT$ = "Remito Número " + TRIM$(REMISEL$) + "\Cliente: " + TRIM$(CLIE$) + "\Fecha: " + FECHATEX$(FEX) + "\   \\Abandonar\Marcar 'Sin Cargo'"
    ALTE% = COMALTER%(TXT$)
    If ALTE% <> 2 Then GoTo 10
    '
    Call APERBASDAT("STOCKS")
    
    SQLX$ = "UPDATE MOVISTO SET DOSECCOR = 'S/Cargo'"
    SQLX$ = SQLX$ + ", DoSecLar ='Sin Cargo' "
    SQLX$ = SQLX$ + " WHERE DoPriCor = '" & REMISEL$ & "'"
    SQLX$ = SQLX$ + "AND Nume_Clie = " & CLISEL% & " "
    SQLX$ = SQLX$ + "AND FechMov = '" & FEMOV$ & "'"
     FLEXCONN.Execute (FILTSQL$(SQLX$))
    '
    Call COMUNI("Movimiento " + REMISEL$ + " Marcado 'SinCargo'.")
    GoTo 10
    '
99  Command6.Enabled = True
    '
End Sub
'

Private Sub Command7_Click()
    Call HELPONLINE("REMITO")
End Sub

Sub Command8_Click()
    Dim RESP%
    Dim NUMEOCXX$
    '
    Command8.Enabled = False
    MODO% = COMALTER%("Opciones de Salida\\Pantalla\Impresora\Cancelar")
    If MODO% < 1 Or MODO% > 2 Then GoTo 99
    '
    CMO$ = TRIM$(Control$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
      CMO$ = "RT"
    End If
10  FEX = HOY(HO$)
    Has% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    Des% = FECHANUM(DESS$)
    VDEF$ = MKI$(Des%) + MKI$(Has%)
    '
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    Des% = CVI(Left$(OBX$, 2))
    Has% = CVI(Mid$(OBX$, 3, 2))
    If Des% > Has% Then
        GoTo 10
    End If
    FEX = Des%: PERIO$ = FECHATEX$(FEX)
    FEX = Has%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
'    DES1# = CDbl(CVDAT(DES%))
'    HAS1# = CDbl(CVDAT(HAS%))
'    DESD1$ = CVDAT(DES%)
'    HAST1$ = CVDAT(HAS%)
    DESD1$ = FETEXCOR1$(Des%) ' SE CAMBIA EL FORMATO DE FECHA PARA SQL
    HAST1$ = FETEXCOR1$(Has%)
    
    If MODO% = 1 Then GoTo 56
    '////////******NUEVO FILTRO CLIENTE*******///////
    NCMIN = 999999: NCMAX = 0
    FIN& = ULTREG&("DEUDOR12")
    For U& = 1 To FIN&
        NCX = CVI(Left$(REGLEIDO$("DEUDOR12", U&), 2))
        If NCX > 0 Then
            If NCX > NCMAX Then NCMAX = NCX
            If NCX < NCMIN Then NCMIN = NCX
        End If
    Next U&
    '
    With RACLIRPF
      .Text10 = TRIM$(Str$(NCMIN))
      .Text9 = rasocli$(NCMIN)
      .Text8 = TRIM$(Str$(NCMAX))
      .Text7 = rasocli$(NCMAX)
      .Show 1
    End With
    '
    'SI CANCELA EN EL FORMULARIO DE RANGO DE CTAS. DE CLIENTES
    If TRIM$(RACLIRPF.Text10.TEXT) = "" And TRIM$(RACLIRPF.Text8.TEXT) = "" Then
        Command8.Enabled = True
        Exit Sub
    End If
    '
    CLINFE = Val(RACLIRPF.Text10)
    CLISUP = Val(RACLIRPF.Text8)
    '//////********HASTA ACA NUEVO FILTRO CLIENTE****///////
56  Screen.MousePointer = 11
    Call APERBASDAT("STOCKS")
    '
    SQLX$ = " SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE (CodiMovi = '" & CMO$ & "' "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
    SQLX$ = SQLX$ + " AND DoPriLar <> '' "
    SQLX$ = SQLX$ + " AND ((DoSecLar IS NULL) or (DoSecLar = '') OR (DoSecLar LIKE 'P/F%')) " ' SE CAMBIA EN LIKE EL FORMATO SIMBOLO * POR % PARA SQL
    SQLX$ = SQLX$ + " AND FechMov >= '" & DESD1$ & "' "
    SQLX$ = SQLX$ + " AND FechMov <= '" & HAST1$ & "' "
    'PARA EVITAR COMPROBANTES ARRUINADOS
    'PARA EVITAR COMPROBANTES ARRUINADOS
    SQLX$ = SQLX$ + " AND LEFT(REFERCOR,4) <> 'ANUL' "
    'FILTRO CLIENTE
    If MODO% <> 1 Then
      SQLX$ = SQLX$ + " AND Nume_Clie >= " & CLINFE & ""
      SQLX$ = SQLX$ + " AND Nume_Clie<= " & CLISUP & ""
    End If
    '
    'Set CONSITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim CONSITMP As ADODB.Recordset
    Set CONSITMP = New ADODB.Recordset
    CONSITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

    With CONSITMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, "No Existen Remitos Pendientes de Facturación.")  'Generados para este Cliente")
        On Error GoTo 0
        Screen.MousePointer = 1
        Command8.Enabled = True
        Exit Sub
      End If
      On Error GoTo 0
      '
      RECONSI$ = REGBLAN("CNSICLI")
      RECONSJ$ = REGBLAN("RPENFAC")
      J& = 0
      Do While .EOF = False
         NCLI = !Nume_Clie
         Call REPLA(RECONSI$, MKS$(NCLI), 1, 4)
         Call REPLA(RECONSI$, MKI$(CVINT(!FECHMOV)), 5, 2)
         Call REPLA(RECONSI$, "", 7, 18)
         Call REPLA(RECONSI$, !DoPriCor, 25, 12)
         Call REPLA(RECONSI$, MKI$(!cod_inte), 37, 2)
         Call REPLA(RECONSI$, MKD$(!CANTSALID), 39, 8)
         Call REPLA(RECONSI$, MKD$(!CANTINGRE), 47, 8)
         Call REPLA(RECONSI$, MKD$(!CantSaldo), 55, 8)
         Call REPLA(RECONSI$, !ReferCor, 63, 46)
         Call REPLA(RECONSI$, String$(12, 0), 109, 12)
         Call REPLA(RECONSI$, MKS$(Val(Mid(!DoPriCor, 5))), 109, 4)
         Call REPLA(RECONSI$, MKI$(CVINT(!FeReMovi)), 122, 2)
          HOR = Format(!FeReMovi, "HH:MM:SS")
         Call REPLA(RECONSI$, MKI$(CVINT(HOR)), 124, 2)
         Call REPLA(RECONSI$, MKI$(XVALO(!NumUsuar)), 126, 2)
         '
         Call REPLA(RECONSJ$, MKS$(NCLI), 1, 4)
         Call REPLA(RECONSJ$, MKI$(!cod_inte), 5, 2)
         Call REPLA(RECONSJ$, !DoPriCor, 7, 12)
         Call REPLA(RECONSJ$, !DOPRILAR, 19, 18)
         Call REPLA(RECONSJ$, MKI$(CVINT(!FECHMOV)), 37, 2)
         If IsNull(!NUPEDCLI) Then
            Call REPLA(RECONSJ$, MKS$(0), 39, 4)
         Else
            Call REPLA(RECONSJ$, MKS$(!NUPEDCLI), 39, 4)
         End If
         If IsNull(!NUPEDCLI) Then
             NUMEOCXX$ = ""
           Else
             NUMEOCXX$ = NUOCOCLI$(!NUPEDCLI, !cod_inte)
         End If
         Call REPLA(RECONSJ$, NUMEOCXX$, 43, 24)
         Call REPLA(RECONSJ$, MKD$(!CANTSALID), 67, 8)
         Call REPLA(RECONSJ$, MKD$(!CANTINGRE), 75, 8)
         Call REPLA(RECONSJ$, MKD$(!CantSaldo), 83, 8)
         Call REPLA(RECONSJ$, MKD$(!VALOUNIT), 91, 8)
           TOTI# = !VALOUNIT * (!CANTSALID - !CANTINGRE)
         Call REPLA(RECONSJ$, MKD$(TOTI#), 99, 8)
         '
         J& = J& + 1
         Call GRAREG("CNSICLI", RECONSI$, J&)
         Call GRAREG("RPENFAC", RECONSJ$, J&)
      .MoveNext
      Loop
    End With
    '
    Call SETULTREG("CNSICLI", J&)
    Call FILESORT("CNSICLI", "", 1, 6, "ASC")
    Call CIERRARCH("CNSICLI")
    '
    Call SETULTREG("RPENFAC", J&)
    Call FILESORT("RPENFAC", "", 1, 6, "ASC")
    Call CIERRARCH("RPENFAC")
    '
    If MODO% = 1 Then
        Call CONECRUN("*CRPENFA", "Remitos Sin Facturar (Clientes)")
      Else
        RESP% = COMALTER%("Opciones de Agrupamiento\\Comprobante\Articulo\Cancelar")
        If RESP% = 1 Then
            Call CONECRUN("*LREPEFA", "Remitos Sin Facturar (Clientes)")
          ElseIf RESP% = 2 Then
            Call HEADERS("RPENFAC", "")
            If CLINFE = CLISUP Then
                NCXX = CLINFE
                Call HEADERS("RPENFAC", "Cliente: " + rasocli$(NCXX))
                Call CONECRUN("*LRPENFB", "Remitos Sin Facturar (Clientes)")
              Else
                Call CONECRUN("*LRPENFA", "Remitos Sin Facturar (Clientes)")
            End If
        End If
    End If
    '
99  Command8.Enabled = True
End Sub

 Sub Command9_Click()
    Command9.Enabled = False
    '
    '\\\OT-05-0437-WAR-05/07/05///
        MODO% = COMALTER%("Modalidad de Devolución\\Por Remito Propio\Por Remito de Cliente\Cancelar")
        If MODO% < 1 Or MODO% > 2 Then GoTo 99
        If MODO% = 2 Then 'Devolución de Remito de Cliente.
          Call DEVRECLI
          GoTo 99
        End If
        '\\\OT-05-0437-WAR-FIN///
    CMO$ = TRIM$(Control$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
       CMO$ = "RT"
    End If
    '
    Call CARSELMESA
    '
10  Screen.MousePointer = 1
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) = "" Then
       GoTo 99
    End If
    CORRESQ$ = MESANO$(MESASEL$, Des%, Has%)
    '
    Call FRESHECHO("SELREMI")
    Screen.MousePointer = 11
    '
    DES1$ = CVDAT(Des%)
    HAS1$ = CVDAT(Has%)
    '
    SQLX$ = " SELECT DISTINCT DoPriCor, Nume_Clie, FechMov FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE FechMov >= '" & DES1$ & "' "
    SQLX$ = SQLX$ + " AND FechMov <= '" & HAS1$ & "' "
    'PARA EVITAR COMPROBANTES ARRUINADOS
    SQLX$ = SQLX$ + " AND LEFT(REFERCOR,4) <> 'ANUL' "
    'SQLX$ = SQLX$ + " AND Len(DoSecCor) < 1 "  ' PARA QUE PUEDA DEVOLVERSE AUN FACTURADO
    SQLX$ = SQLX$ + " AND CodiMovi =  '" & CMO$ & "'  "
    '
    Call APERBASDAT("STOCKS")
    'Set MOVITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim MOVITMP As ADODB.Recordset
    Set MOVITMP = New ADODB.Recordset
    MOVITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    
    '
    With MOVITMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, " No Existen Registros que Cumplan\con la Condición SQL.")
        On Error GoTo 0
        GoTo 10
      End If
      On Error GoTo 0
      '
      J& = 0
      Do While (.EOF = False)
        NCLI = !Nume_Clie
        REMI$ = !DoPriCor
        FF% = CVINT(!FECHMOV)
        '\\\OT-06-0083-WAR-07/03/06///
        F1$ = FECHATEX$(FF%)
        REFE1$ = F1$ + " - " + rasocli$(NCLI)
        '\\\OT-06-0083-WAR-FIN///
        '
        Z$ = Space$(40) + String$(8, 0)
        Call REPLA(Z$, REMI$, 1, 10)
        Call REPLA(Z$, REFE1$, 11, 28) ' RASOCLI$(NCLI), 11, 30)
        Call REPLA(Z$, MKS$(U&), 39, 4)
        Call REPLA(Z$, MKI$(FF%), 43, 2)
        Call REPLA(Z$, MKS$(NCLI), 45, 4)
        J& = J& + 1
        Call GRAREG("!SELREMI", Z$, J&)
      .MoveNext
      Loop
    End With
    '
    Call SETULTREG("!SELREMI", J&)
    Call CIERRARCH("!SELREMI")
    Call FRESHECHO("!SELREMI")
    Screen.MousePointer = 1
    '
    If J& < 1 Then
       Call MENSERR(24, "No se Encontraron Remitos Posibles de Devolución\en el Período Elegido.")
       GoTo 10
    End If
    '
28  RERE$ = REGSEL$("!SELREMI")
    If Len(RERE$) <= 4 Then GoTo 10
    '
    REMISEL$ = TRIM$(Left$(RERE$, 10))
    CLIE$ = Mid$(RERE$, 11, 28)
    FEXI% = CVI(Mid$(RERE$, 43, 2))
    FEX = FEXI%
    FEMOVD = Int(CDbl(CVDAT(FEXI%)))
    CLISE = CVS(Mid$(RERE$, 45, 4))
    '
    'TXT$ = "Remito Número " + TRIM$(REMISEL$) + "\Cliente: " + TRIM$(CLIE$) + "\Fecha: " + FECHATEX$(FEX) + "\   \\Abandonar\Marcar 'Sin Cargo'"
    '
    Call APERBASDAT("STOCKS")
    SQLX$ = "SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + "WHERE DoPriCor = '" & REMISEL$ & "' "
    SQLX$ = SQLX$ + " AND CodiMovi =  '" & CMO$ & "'  "
    'SQLX$ = SQLX$ + "AND Len(DoSecCor) < 1 "   ' PARA QUE PUEDAN DEVOLVERSE AUN FACTURADOS
    SQLX$ = SQLX$ + "AND Nume_Clie = " & CLISE & " "
    'SQLX$ = SQLX$ + "AND INT(CDBL(FechMov)) = " & FEMOVD
    'Set MOVITMP1 = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    Dim MOVITMP1 As ADODB.Recordset
    Set MOVITMP1 = New ADODB.Recordset
    MOVITMP1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

    With MOVITMP1
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, " No Existen Registros que Cumplan\con la Condición SQL.")
        On Error GoTo 0
        GoTo 10
      End If
      On Error GoTo 0
      '
      DOPRI$ = REMISEL$
      DOPRILA$ = !DOPRILAR
      DOSEC$ = ""       ' QUEDA LIBRE PARA PONERLE EL NUMERO DE FACTURA
      DOSELA$ = ""      ' QUEDA LIBRE PARA PONERLE EL NUMERO DE FACTURA
      MONEVA% = !MoneVal
      DEPOMO% = !depomovi
      '
      HOII% = HOY(HOS$)
      '
      J& = 0
      Do While (.EOF = False)
        XX$ = REGBLAN$("BOLREDET")
        Call REPLA(XX$, MKS$(!NuOrItem), 1, 4)
        Call REPLA(XX$, MKI$(!cod_inte), 83, 2)
        Call REPLA(XX$, Unimed$(!cod_inte), 85, 4)
        Call REPLA(XX$, MKD$(!CANTSALID), 89, 8)
        Call REPLA(XX$, !idenlote, 119, 16)
          VUNIT = !VALOUNIT
        Call REPLA(XX$, MKS$(VUNIT), 159, 4)
        Call REPLA(XX$, MKD$(0), 171, 8)
        J& = J& + 1
        Call GRAREG("!BOLREDET", XX$, J&)
      .MoveNext
      Loop
    End With
    Set MOVITMP1 = Nothing
    Call SETULTREG("!BOLREDET", J&)
    '
    '\\\OT-06-0084-WAR-07/03/06///
29  VDEF$ = Space(4)
    Call REPLA(VDEF$, MKI$(FEXI%), 1, 2)
    Call REPLA(VDEF$, MKI$(DEPOMO%), 3, 1)
    OBJ1$ = OBJPLA$("ENTRAFECHADEVU", VDEF$)
    If TRIM$(OBJ1$) = "" Then GoTo 28
    FEDEVU% = CVI(Mid$(OBJ1$, 1, 2))
    DEPOMO1% = Asc(Mid$(OBJ1$, 3, 1))
    If ECOARCH$("TADEPOSI", MKI$(DEPOMO1%)) = "" Then
      Call MENSERR(24, "Depósito no Válido.")
      GoTo 29
    End If
    '\\\OT-06-0084-WAR-FIN///
    'FEDEVU% = CVI(OBJPLA$("ENTRAFECHADEVU", MKI$(FEXI%)))
    If FEDEVU% < 1 Or FEDEVU% > HOY(HOS$) Then GoTo 10
    FEX = FEDEVU%
    '
30  VTB% = VENTABU%("DEVOLCLI/TITUPAN=Devolucion " + REMISEL$ + " - " + TRIM$(CLIE$) + " - " + REPLACAR$(FECHATEX$(FEX), "/", "-"))
    If VTB% < 0 Then GoTo 10
    '
    For KKIL& = 1 To J&
      XX$ = REGLEIDO$("!BOLREDET", KKIL&)
      If CVD(Mid$(XX$, 171, 8)) > CVD(Mid$(XX$, 89, 8)) + 0.05 Then
        Call MENSERR(24, "La Cantidad Devuelta no Puede Superar\la Despachada Originalmente.")
        GoTo 30
      End If
    Next KKIL&
    '
    For KKIL& = 1 To J&
      XX$ = REGLEIDO$("!BOLREDET", KKIL&)
      If CVD(Mid$(XX$, 171, 8)) > 0 Then
        CIXI% = CVI(Mid$(XX$, 83, 2))
        LOTE$ = Mid$(XX$, 119, 16)
        NUORIT% = CVS(Left$(XX$, 4)) - 1
        CADEVU = CVD(Mid$(XX$, 171, 8))
        VUNI# = CVS(Mid$(XX$, 159, 4))
        '
        Call MOVISTOK(FEDEVU%, CIXI%, LOTE$, "DV", DOPRILA$, DOPRI$, DOSELA$, DOSEC$, "Devol." + rasocli$(CLISE), NUORIT%, VUNI#, MONEVA%, CLISE, DEPOMO1%, CADEVU)
        '\\\OT-06-0105-WAR-27/03/06///
        If DESCOMPO(CIXI%) > 0 Then
          'SI HAY DEVULUCION DE KIT
          'COMPENSO CONJUNTO
          Call MOVISTOK(FEDEVU%, CIXI%, LOTE$, "CI", DOPRILA$, DOPRI$, DOSELA$, DOSEC$, "Devol." + rasocli$(CLISE), NUORIT%, VUNI#, MONEVA%, CLISE, DEPOMO1%, CADEVU * (-1))
          'COMPENSO COMPONENTES
          Call AREXPLO(CIXI%, 1, 1)
          For KKJ% = 1 To CAIT%
            CICO% = CIJ%(KKJ%)
            CACICO = USOI(KKJ%) * CADEVU
            Call MOVISTOK(FEDEVU%, CICO%, LOTE$, "CI", DOPRILA$, DOPRI$, DOSELA$, DOSEC$, "Devol." + rasocli$(CLISE), NUORIT%, VUNI#, MONEVA%, CLISE, DEPOMO1%, CACICO)
          Next KKJ%
        End If
        '\\\OT-06-0105-WAR-FIN///
      End If
    Next KKIL&
    'Movisto.Close
    '
    Call COMUNI("Devolución Procesada")
    GoTo 10
    '
99  Command9.Enabled = True
    '
End Sub

Sub RECONCLI(NCX7%)
   '
   Dim SALCONSI(32767)
   Call SETULTREG("RECONCLI", 0)
   Call SETULTREG("RECONMAT", 0)
   '
   If NCX7% > 0 Then
     If ECOARCH$("DEUDOR1", MKI$(NCX7%)) <> "" Then
       Screen.MousePointer = 11
       FIN& = ULTREG&("CONSICLI")
       For KLK& = 1 To 32767: SALCONSI(KLK&) = 0: Next KLK&
       For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
         XX$ = REGLEIDO$("CONSICLI", U&)
         NCX8% = CVI(Left$(XX$, 2))
         If NCX8% = NCX7% Then
           CIXI% = CVI(Mid$(XX$, 35, 2))
           If CIXI% > 0 Then
             If CIXI% <= NUMAS% Then
               CACON = CVD(Mid$(XX$, 37, 8)) - CVD(Mid$(XX$, 45, 8))
               SALCONSI(CIXI%) = SALCONSI(CIXI%) + CACON
             End If
           End If
         End If
       Next U&
       '
       JJ& = 0
       FIN& = NUINV%
       For U& = 1 To NUINV%
         Call TRACE(20, U&, FIN&)
         II$ = REGLEIDO$("INVERT", U& + 1)
         CIXI% = CVI(Mid$(II$, 17, 2))
         If CIXI% > 0 Then
           If CIXI% <= NUMAS% Then
             If SALCONSI(CIXI%) > 0 Then
               XX$ = REGBLAN$("RECONCLI")
               YY$ = REGBLAN$("RECONMAT")
               Call REPLA(XX$, MKI$(NCX7%), 1, 2)
               Call REPLA(XX$, MKI$(CIXI%), 3, 2)
               Call REPLA(XX$, MKS$(SALCONSI(CIXI%)), 5, 4)
               Call REPLA(XX$, String$(16, "."), 9, 16)
               Call REPLA(YY$, MKI$(CIXI%), 1, 2)
               Call REPLA(YY$, MKS$(SALCONSI(CIXI%)), 3, 4)
               JJ& = JJ& + 1
               Call GRAREG("RECONCLI", XX$, JJ&)
               Call GRAREG("RECONMAT", YY$, JJ&)
             End If
           End If
         End If
       Next U&
       Call SETULTREG("RECONCLI", JJ&)
       Call CIERRARCH("RECONCLI")
       Call SETULTREG("RECONMAT", JJ&)
       Call CIERRARCH("RECONMAT")
       Screen.MousePointer = 1
       '
     End If
   End If

End Sub

Private Sub Form_Activate()
    '
    'Command19.Enabled = False
    'If Control$("REMSTAN", "REPLAORI") = "SI" Then
    '  Command19.Enabled = True
    'End If
    '
    If EXISTE%("COCUMENT.EXE") > 0 Then
      Command11.Enabled = True
    End If
    '
    If USNBR% Mod 100 <> 1 Then
       Command15.Visible = False
    End If
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "GABAL") Or InStr(EPAC$, "MARBE") > 0 Then
      Command18.Visible = True
    End If
    '
    If InStr(EPAC$, "ARCON") > 0 Then
      Command23.Visible = True
    End If
    '
    If InStr(EPAC$, "FERVI") > 0 Then
      Command26.Visible = True
    End If
    '
    If InStr(1, EMPREAC$, "NEUQUEN") > 0 Then
      mnuPasarRemitosDropbox.Visible = True
    End If
    '
    End Sub

Sub DEVRECLI()
    '\\\OT-06-0084-WAR-08/03/06///
10  DP1% = Val(Control$("", "DEPOEXPE"))
    Y$ = Space(64)
    Call REPLA(Y$, MKI$(0), 1, 2)
    Call REPLA(Y$, MKI$(0), 3, 2)
    Call REPLA(Y$, MKI$(0), 5, 2)
    Call REPLA(Y$, MKI$(DP1%), 47, 1)
    '\\\OT-06-0084-WAR-FIN///
    '
    OBJE$ = "DEVOCLI"
    PEDY$ = OBJPLA$(OBJE$, Y$)
    If PEDY$ = "" Then Exit Sub
    NC% = CVI(Mid$(PEDY$, 1, 2))
    CI% = CVI(Mid$(PEDY$, 5, 2))
    FEDEVU% = CVI(TRIM(Mid$(PEDY$, 3, 2)))
    'FE2# = CDbl(CVDAT(FEDEVU%))
    FECH2$ = FETEXCOR1$(FEDEVU%)
    NURE$ = TRIM(Mid$(PEDY$, 7, 16))
    '\\\OT-06-0084-WAR-08/03/06///
    DEPOMO% = Asc(Mid$(PEDY$, 47, 1))
    If ECOARCH$("TADEPOSI", MKI$(DEPOMO%)) = "" Then
      Call MENSERR(24, "Depósito no Válido.")
      GoTo 10
    End If
    '\\\OT-06-0084-WAR-FIN///
    'FALTA VALIDAR QUE NO EXISTA EL REMITO DEL CLIENTE EN UNA DEVOLUCION
    ' en !NumRem_cli CON NURE$
    '
    CMO$ = TRIM$(Control$("", "CODREFAC"))
    If CMO$ = "" Then CMO$ = "RT"
    '
    Call APERBASDAT("STOCKS")
    SQLX$ = "SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE CodiMovi =  '" & CMO$ & "'  "
    SQLX$ = SQLX$ + " AND LEFT(DOSECCOR,2) <> 'FC' " 'PARA QUE NO TRAIGA LOS FACTURADOS
    SQLX$ = SQLX$ + "AND Nume_Clie = " & NC% & " "
    SQLX$ = SQLX$ + " AND COD_INTE = " & CI% & ""
    SQLX$ = SQLX$ + " AND LEFT(REFERCOR,4) <> 'ANUL' "
    SQLX$ = SQLX$ + "AND FechMov <= '" & FECH2$ & "' "
    SQLX$ = SQLX$ + " ORDER BY DOPRILAR "
    '
    'Set MOVITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim MOVITMP As ADODB.Recordset
    Set MOVITMP = New ADODB.Recordset
    MOVITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

    With MOVITMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, " No Existen Registros que Cumplan\con la Condición SQL.")
        On Error GoTo 0
        GoTo 10
      End If
      On Error GoTo 0
      '
      HOII% = HOY(HOS$)
      '
      J& = 0
      Do While (.EOF = False)
         XX$ = REGBLAN$("BOLREDET")
         '
         DOPRILA$ = !DOPRILAR
         DOSEC$ = ""       ' QUEDA LIBRE PARA PONERLE EL NUMERO DE FACTURA
         DOSELA$ = ""      ' QUEDA LIBRE PARA PONERLE EL NUMERO DE FACTURA
         MONEVA% = !MoneVal
         'DEPOMO% = !DepoMovi  'EL DEPOSITO LO ASGINA EL USURARIO'OT-06-0084-WAR-08/03/06///
         DOPRICO1$ = !DoPriCor 'HAY QUE GRABARLO EN EL ARCHI.
         '
        Call REPLA(XX$, MKS$(!NuOrItem), 1, 4)
          REMISEL$ = !DOPRILAR
          DOPRI$ = REMISEL$
        Call REPLA(XX$, REMISEL$, 5, 45) 'ACA PONGO EL REMITO LARGO
        Call REPLA(XX$, DOPRICO1$, 50, 14) 'ACA GRABO EL REMITO CORTO
          FECHMO = !FECHMOV
        Call REPLA(XX$, MKI(CVINT(FECHMO)), 81, 2)
        Call REPLA(XX$, MKI$(!cod_inte), 83, 2)
        Call REPLA(XX$, Unimed$(!cod_inte), 85, 4)
        Call REPLA(XX$, MKD$(!CANTSALID), 89, 8)
        Call REPLA(XX$, Chr(DEPOMO%), 97, 1)
        Call REPLA(XX$, !idenlote, 119, 16)
          VUNIT = !VALOUNIT
        Call REPLA(XX$, MKS$(VUNIT), 159, 4)
        Call REPLA(XX$, MKD$(0), 171, 8)
        J& = J& + 1
        Call GRAREG("!BOLREDET", XX$, J&)
      .MoveNext
      Loop
    End With
    Call SETULTREG("!BOLREDET", J&)
    '
40  VTB% = VENTABU%("DEVCLIRT/TITUPAN=Devolucion " + DESCRIT0(CI%) + " - " + TRIM$(rasocli(NC%)) + " - " + REPLACAR$(FECHATEX$(FEX), "/", "-"))
    If VTB% < 0 Then GoTo 10
    '
    For KKIL& = 1 To J&
      XX$ = REGLEIDO$("!BOLREDET", KKIL&)
      If CVD(Mid$(XX$, 163, 8)) > CVD(Mid$(XX$, 89, 8)) + 0.05 Then
        Call MENSERR(24, "La Cantidad Devuelta no Puede Superar\la Despachada Originalmente.")
        GoTo 40
      End If
    Next KKIL&
    '
    For KKIL& = 1 To J&
      XX$ = REGLEIDO$("!BOLREDET", KKIL&)
      If CVD(Mid$(XX$, 163, 8)) > 0 Then
        DOPRI$ = TRIM(Mid$(XX$, 50, 14))
        CIXI% = CVI(Mid$(XX$, 83, 2))
        LOTE$ = Mid$(XX$, 119, 16)
        NUORIT% = CVS(Left$(XX$, 4)) - 1
        CADEVU = CVD(Mid$(XX$, 163, 8))
        VUNI# = CVS(Mid$(XX$, 159, 4))
        DEPOMO% = Asc(Mid$(XX$, 97, 1))
        Call MOVISTOK(FEDEVU%, CIXI%, LOTE$, "DV", DOPRILA$, DOPRI$, DOSELA$, DOSEC$, "Devol." + rasocli$(NC%), NUORIT%, VUNI#, MONEVA%, NC%, DEPOMO%, CADEVU)
        If DESCOMPO(CIXI%) > 0 Then
          'SI HAY DEVULUCION DE KIT
          'COMPENSO EL CONJUNTO
            Call MOVISTOK(FEDEVU%, CIXI%, LOTE$, "CI", DOPRILA$, DOPRI$, DOSELA$, DOSEC$, "Devol." + rasocli$(NC%), NUORIT%, VUNI#, MONEVA%, NC%, DEPOMO%, CADEVU * (-1))
          'COMPENSO LOS COMPONENTES
          Call AREXPLO(CIXI%, 1, 1)
          For KKJ% = 1 To CAIT%
            CICO% = CIJ%(KKJ%)
            CACICO = USOI(KKJ%) * CADEVU
            Call MOVISTOK(FEDEVU%, CICO%, LOTE$, "CI", DOPRILA$, DOPRI$, DOSELA$, DOSEC$, "Devol." + rasocli$(NC%), NUORIT%, VUNI#, MONEVA%, NC%, DEPOMO%, CACICO)
          Next KKJ%
        End If
      End If
    Next KKIL&
    '
    Screen.MousePointer = 1
    Call COMUNI("Devolución Procesada")
    '
999 End Sub
Sub NUEVO_COMAND6_CLICK()
    
    If DERACCE("SINCARRE", "Marcar Sin Cargo Remitos") < 2 Then Exit Sub

    CMO$ = TRIM$(Control$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
       CMO$ = "RT"
    End If
    '
    'RANGO DE FECHAS DE REMITO
    Call CARSELMESA
10  Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) = "" Then
       Exit Sub
    End If
    CORRESQ$ = MESANO$(MESASEL$, Des%, Has%)

    '
    DES1 = FETEXCOR1$(Des%)
    HAS1 = FETEXCOR1$(Has%)

    '
    'AGREAGADO PARA QUITAR DE LA TABLA SQL_LISELEC EL REGISTRO DONDE SE CREO !SELREMI
    'POR QUE TIENE EL MISMO NOMBRE QUE UN ARCHIVO DAT PERO CON DIFERENTE ESTRUCTURA
    'Y SI NO SE BORRABA NO SE GENERABA EL DAT SIEMPRE MOSTRABA LO MISMO
    Call BORRAREGISTROS("SQL_LISELEC", "NOMLISELE = '!SELREMI'", REGAF&, "NOMESS")

    'SELECCION REMITOS DE MOVISTO Y MOVIREPA
20  CONDI$ = "FechMov >= '" & DES1 & "' "
    CONDI$ = CONDI$ + " AND FechMov <= '" & HAS1 & "' "
    CONDI$ = CONDI$ + " AND (CodiMovi =  '" & CMO$ & "'  "
    CONDI$ = CONDI$ + " OR CodiMovi =  'DV' OR CodiMovi =  '*R')"
    'PARA EVITAR COMPROBANTES ARRUINADOS
    CONDI$ = CONDI$ + " AND LEFT(REFERCOR,4) <> 'ANUL' "
    CONDI$ = CONDI$ + " AND (DoSecLar IS NULL or (DoSecLar = '') OR (DoSecLar LIKE 'P/F%')) "

    Call CARGALISEL("!SELREMI1", "MOVISTO", "DOPRICOR/A12;NUME_CLIE/I6+DEUDOR12/A24;FECHMOV/D8", CONDI$, "DISTINCT")

    Call CARGALISEL("!SELREMI1", "MOVIREPA", "DOPRICOR/A12;NUME_CLIE/I6+DEUDOR12/A24;FECHMOV/D8", CONDI$, "DISTINCT/NOCLEAR")

30  Call SELECHO("!SELREMI1")
    '

    'REMITO SELECCIONADO
    REMISEL$ = TRIM$(VALACT1$("!SELREMI1"))

    If REMISEL$ = "" Then GoTo 10
    NCLI = XVALO(VALOBADA("MOVISTO", "NUME_CLIE", CONDI$ + " AND DOPRICOR = '" & REMISEL$ & "'", "NOMESS"))
    If NCLI > 0 Then
       FE% = CVINT(VALOBADA("MOVISTO", "FECHMOV", CONDI$ + " AND DOPRICOR = '" & REMISEL$ & "'", "NOMESS"))
    Else
       NCLI = XVALO(VALOBADA("MOVIREPA", "NUME_CLIE", CONDI$ + " AND DOPRICOR = '" & REMISEL$ & "'", "NOMESS"))
       FE% = CVINT(VALOBADA("MOVIREPA", "FECHMOV", CONDI$ + " AND DOPRICOR = '" & REMISEL$ & "'", "NOMESS"))
    End If
    '
    CLIE$ = rasocli$(NCLI)
    '
    'MENSAJE DE CONFIRMACION DE MARCA
    TXT$ = "Remito Número " + TRIM$(REMISEL$) + "\Cliente: " + TRIM$(CLIE$) + "\Fecha: " + FECHATEX$(FE%) + "\   \\Abandonar\Marcar 'Sin Cargo'"
    ALTE% = COMALTER%(TXT$)
    If ALTE% <> 2 Then GoTo 30
    '
    'GRABAR MARCA DE SIN CARGO
    FEMOV = FETEXCOR1$(FE%)
    
    CONDI$ = "DoPriCor = '" & REMISEL$ & "'"
    CONDI$ = CONDI$ + " AND Nume_Clie = " & NCLI & " "
    CONDI$ = CONDI$ + " AND FechMov = '" & FEMOV & "'"
    'CONDICION AGREGADA PARA QUE SOLO MARQUE LOS QUE NO ESTAN FACTURADOS - CASOS DE REMITOS FACTURADOS PARCIALMENTE QUEDABAN DESVINCULADOS DE LA FACTURA
    CONDI$ = CONDI$ + " AND ((DOSECCOR = '' OR DOSECCOR IS NULL) OR (DOSECLAR LIKE 'P/F%'))"
    
    Call ACTUREGISTRO("MOVISTO", "DOSECLAR", CONDI$, "SIN CARGO", REGAF&, "NOMESS")
    Call ACTUREGISTRO("MOVISTO", "DOSECCOR", CONDI$, "S/Cargo", REGAF&, "NOMESS")
    Call ACTUREGISTRO("MOVIREPA", "DOSECLAR", CONDI$, "SIN CARGO", REGAF&, "NOMESS")
    Call ACTUREGISTRO("MOVIREPA", "DOSECCOR", CONDI$, "S/Cargo", REGAF&, "NOMESS")
    '
    'PRESENTA MENSAJE INDICANDO FIN DE LA  TRANSACCION
    Call COMUNI("Movimiento " + REMISEL$ + " Marcado 'Sin Cargo'.")
    GoTo 20

End Sub

Sub MODIFICAPACKLIST(OK%, NUMLEMP&)
    
 
10  VTB% = VENTABU%("LISTAEMP")
    If VTB% < 0 Then OK% = 0: Exit Sub
    OK% = 1
    
    NPALLANT% = 0
    MAXMECUBPORPALET = XVALO(Control$("REMSTAN", "LIMITEM3"))
    M3ACUMULADOS = 0
    For U& = 1 To ULTREG&("!LISTAEMP")
       XY$ = REGLEIDO$("!LISTAEMP", U&)
       NPALL% = CVI(Mid$(XY$, 1, 2))
       CI% = CVI(Mid$(XY$, 3, 2))
       CANTXCAJA = CVS(Mid$(XY$, 5, 4))
       CANTICAJAS% = CVI(Mid$(XY$, 9, 2))
       NUPED& = CVS(Mid$(XY$, 11, 4))
       CONDI$ = "CODINT = " & CI%
       CANPXCJA& = XVALO(VALOBADA("MASTER", "PIEZASXCAJA", CONDI$, "NOMESS"))
       '
       'VALIDO QUE TENGA TODOS LOS DATOS DEL ARCHIVO MAESTRO
       If FALTANDATOS%(CI%) > 0 Then
          Call MENSERR(24, "Para el Código: " & CODEXT$(CI%) & " Falta Datos Del Archivo Maestro\Deberá Ingresarlos Antes del Cierre")
       End If
       '
       If CANTXCAJA > CANPXCJA& Then Call COMUNI("Piezas x Caja  No Válido para Registro " & U&): GoTo 10
       If NPALL% < 1 Then Call MENSERR(24, "Número de Pall No Válido para Registro " & U&): GoTo 10
       '
       If NPALL% < NPALLANT% Or NPALL% > NPALLANT% + 1 Then
          Call MENSERR(24, "Los Pallets Deben ir en Número de\Orden Creciente y Correlativo !!!")
          GoTo 10
       End If
       '
       If CI% < 1 Then Call MENSERR(24, "Número de Código No Válido para Registro " & U&): GoTo 10
       If CANTICAJAS% < 1 Then Call MENSERR(24, "Número de Cajas No Válido para Registro " & U&): GoTo 10
       '
       'VALIDACION QUE EL MISMO PALLET NO SUPERE LA CANTIDAD DE METROS CUBICOS CONFIGURADOS
       If NPALL% <> NPALLANT% Then M3ACUMULADOS = 0
       CONDI$ = "CODINT = " & CI%
       M3CAX = XVALO(VALOBADA("MASTER", "M3CAJA", CONDI$, "NOMESS"))
       M3TOT = CANTICAJAS% * M3CAX
       If (M3ACUMULADOS + M3TOT) > MAXMECUBPORPALET Then
          Call MENSERR(24, "Cantidad de M3 Para el Pallet N: " & NPALL% & " \Supera el Límite Máximo de M3")
          GoTo 10
       End If
       M3ACUMULADOS = M3ACUMULADOS + M3TOT
       '
       'VALIDO EL NUMERO DE PEDIDO SI SE INGRESO
       If NUPED& > 0 Then
          NC% = XVALO(VALOBADA("ENCEMPAQ", "CODI_CLI", " NumLtaEmp = " & NUMLEMP&, " NOMESS"))
          CONDI$ = "NROPED = " & NUPED& & "  AND CODIINT = " & CI% & " And Status < 8 And NROCLIE = " & NC%
          CANTRE& = CantRegistros("PEDDETA", CONDI$)
          If CANTRE& < 1 Then
             Call MENSERR(24, "El Artículo " & CODEXT$(CI%) & " No Corresponde al Pedido: " & NUPED&)
             GoTo 10
          End If
       End If
       NPALLANT% = NPALL%
       '
       'CALCULO Y GUARDO EL PESO BRUTO
       CONDI$ = "CODINT = " & CI%
       KILBRCAJ = XVALO(VALOBADA("MASTER", "KILOBRUTOCAJA", CONDI$, "NOMESS"))
       XY$ = REGLEIDO$("!LISTAEMP", U&)
       Call REPLA(XY$, MKS(KILBRCAJ), 17, 4)
       Call GRAREG("LISTAEMP", XY$, U&)
       '
    Next U&
    '
    Call SETULTREG("LISTAEMP", U& - 1)
    'ORDENO EL ARCHIVO
    Call FILESORT("LISTAEMP", "", 8, 8, "DESC")
    Call FILESORT("LISTAEMP", "", 1, 1, "ASC")
    '
    Call TRALOCAL("LISTAEMP", "")
    '
End Sub

Private Sub RECIBECEN()
    'ESTO FUNCIONA SI EN LA PC DE LOS USUARIOS QUE LO EJECUTEN TIENEN LA UNIDAD DE RED I
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "NEUQUEN") < 1 Then
       Call COMUNI("Esta Opción es Para Que se Reciba de Avellaneda a Neuquen")
       Exit Sub
    End If
       
    Screen.MousePointer = 11
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Screen.MousePointer = 1
    '
5   OPX% = COMALTER%("Archivo de Transferencia\\Recibir Vía Web\Desde Unidad de Disco")
    If OPX% < 1 Or OPX% > 2 Then Exit Sub
    If OPX% = 1 Then
        CARPETASINCRO$ = TRIM$(Control("", "CARPEINT"))
        If CARPETASINCRO$ <> "" Then
            ORIGARCH$ = CARPETASINCRO$
            GoTo 18
        Else
            Call COMUNI("Imposible Continuar.\Falta Definir Carpeta de Intercambio")
            GoTo 5
        End If
    End If
    '
    FILEORIG.Show 1
    If FILEORIG.Label1.Caption = "" Then Exit Sub
    '
10  LU$ = LUDAT$
    ORIGARCH$ = Left$(FILEORIG.Label1.Caption, 2)
    If Mid$(ORIGARCH$, 2, 1) <> ":" Then
       Call MENSERR(24, "Unidad '" + ORIGARCH$ + "' No Disponible.")
       Exit Sub
    End If
18  On Error GoTo 20
    FileCopy ORIGARCH$ + "\ATRASUC.DAT", LU$ + "ATRASUC.DAT"
    On Error GoTo 0
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Call COMUNI("Unidad de Transferencia '" + REPLACAR$(ORIGARCH$, "\", "/") + "'\Leída Exitosamente.\  \Pulse 'Aceptar Mensaje' para\Iniciar Importación de Datos.")
    GoTo 30
    '
20  Resume 21
21  On Error GoTo 0
    Call MENSERR(24, "Imposible Leer Unidad de Transferencia.")
    Exit Sub
    '
30  Screen.MousePointer = 11
    Dim FINUAR%(128), ARTRAS$(128), UREAR&(128)
    Dim BUFGRA As String * 128
    CANAR% = 0
    LU$ = LUDAT$
    FIN& = ULTREG&("ATRASUC")
    For U& = 1 To FIN&
      X$ = REGLEIDO$("ATRASUC", U&)
      ANOM$ = Left$(X$, 8)
      TTXX$ = Mid$(X$, 9, 128)
      '
      For KKI% = 1 To CANAR%
        If ARTRAS$(KKI%) = ANOM$ Then GoTo 35
      Next KKI%
      CANAR% = CANAR% + 1
      KKI% = CANAR%
      FINUAR%(KKI%) = FILEOPEN%(LU$ + TRIM$(ANOM$) + ".DAT", 0, 128)
      ARTRAS$(KKI%) = ANOM$
      UREAR&(KKI%) = 0
      '
35    UREAR&(KKI%) = UREAR&(KKI%) + 1
      BUFGRA$ = TTXX$
      Put #FINUAR%(KKI%), UREAR&(KKI%), BUFGRA$
    Next U&
    '
    For KKI% = 1 To CANAR%
      FIN& = LOF(FINUAR%(KKI%)) / 128
      If LOF(FINUAR%(KKI%)) > 128 * FIN& Then FIN& = FIN& + 1
      For U& = UREAR&(KKI%) + 1 To FIN&
        BUFGRA$ = String$(128, 0)
        Put #FINUAR%(KKI%), U&, BUFGRA$
      Next U&
      Close #FINUAR%(KKI%)
      FINUAR%(KKI%) = 0
    Next KKI%
    '
    For KKI% = 1 To CANAR%
      X1$ = REGLEIDO$(ARTRAS$(KKI%), 1)
      Call GRAREG(ARTRAS$(KKI%), X1$, 1)
      Call CIERRARCH(ARTRAS$(KKI%))
    Next KKI%
    '
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    '
    URET& = ULTREG&("RTRASLA")
    For U& = 1 To ULTREG("RTRASLA")
      XX$ = REGLEIDO$("RTRASLA", U&)
      IDENDO$ = Mid$(XX$, 19, 4)
      NROPED& = CVS(Mid$(XX$, 117, 4))
      RFF$ = RECFILT$("ATRASLA", 4, IDENDO$)
      For UJ& = 1 To Len(RFF$) Step 4
        RELE& = CVS(Mid$(RFF$, UJ&, 4))
        If RELE& > 0 Then
          If RELE& <= ULTREG&("ATRASLA") Then
            YY$ = REGLEIDO$("ATRASLA", RELE&)
            If Left$(YY$, 112) = Left$(XX$, 112) Then GoTo 45 ' YA RECIBIDO
          End If
        End If
      Next UJ&
      '
      YY$ = Left$(XX$, 112) + String$(256, 0)
      Call REPLA(YY$, MKS$(NROPED&), 129, 4)
      Call REGAPP("ATRASLA", YY$)
45  Next U&
    '
    Call ACTUALIZAMASTER
    '
    Call CIERRARCH("ATRASLA")
    Call CIERRARCH("RTRASLA")
99  Call CIERRARCH("*.*")
    Call BAJALDISCO
    '
    Screen.MousePointer = 1
    Call COMUNI("Importación Completa.")
    '
End Sub

Sub GENARTRA()
    '
    Screen.MousePointer = 11
    Call COPYSTRU("ENTRECOM", "TRAENTRE")
    Call BLOQARCH("ATRACEN")
    Call BLANARCH("ATRACEN")
    '
'    JJ& = 0
'    For U& = 1 To ULTREG("ENTRECOM")
'      XX$ = REGLEIDO$("ENTRECOM", U&)
'      STATU% = CVI(Mid$(x$, 191, 2))
'      If STATU% <> 9 Then        ' NO ESTA ANULADO
'        VENVI% = CVI(Mid$(XX$, 193, 2))     ' veces enviado
'        If VENVI% < 5 Then
'          Call REPLA(XX$, MKI$(VENVI% + 1), 193, 2)
'          Call GRAREG("ENTRECOM", XX$, U&)
'          JJ& = JJ& + 1
'          Call GRAREG("TRAENTRE", XX$, JJ&)
'        End If
'      End If
'    Next U&
    Call ABRECONEXION
    '
    SQLX$ = "SELECT * FROM ENTRECOM WHERE IDOCUM > 0"
    Dim Rst As ADODB.Recordset
    Set Rst = New ADODB.Recordset
26: Rst.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 26
    End If
    On Error GoTo 0
    '
    With Rst
       Do While Not .EOF
            STATU% = XVALO(!Status)
            If STATU% <> 9 Then        ' NO ESTA ANULADO
              VENVI% = XVALO(!VECESENVIADO)
              If VENVI% < 5 Then
                !VECESENVIADO = VENVI% + 1
                .Update
                Call REPLA(XX$, MKI$(STATU%), 191, 2)
                Call REPLA(XX$, MKI$(VENVI%), 193, 2)    ' veces enviado
                    FF% = CVINT%(!FECHA)
                Call REPLA(XX$, MKI$(FF%), 1, 2)
                Call REPLA(XX$, SAFETEXT$(!DOCUMENTO), 3, 14)
                Call REPLA(XX$, SAFETEXT$(!TIPODOC), 17, 16)
                    NC% = XVALO(!nume_cli)
                Call REPLA(XX$, MKI$(NC%), 33, 2)
                    RACLI$ = rasocli$(NC%)
                Call REPLA(XX$, RACLI$, 35, 30)
                Call REPLA(XX$, MKS$(!IDOCUM), 97, 4)
                Call REPLA(XX$, Space(16), 101, 16)
                    CI% = XVALO(!cod_inte)
                Call REPLA(XX$, MKI$(CI%), 117, 2)
                Call REPLA(XX$, SAFETEXT$(!Cod_Exte), 119, 16)
                Call REPLA(XX$, DESCRIT0$(CI%), 135, 48)
                Call REPLA(XX$, MKD$(!Cantidad), 183, 8)
                
                JJ& = JJ& + 1
                Call GRAREG("TRAENTRE", XX$, JJ&)
              End If
            End If
            .MoveNext
       Loop
    End With

    On Error Resume Next
    Rst.Close
    Set Rst = Nothing
    On Error GoTo 0
    
    Call SETULTREG("TRAENTRE", JJ&)
    Call CIERRARCH("ENTRECOM")
    Call CIERRARCH("TRAENTRE")
    '
    Dim BUFLE As String * 128
    LU$ = LUDAT$
    n1% = FILEOPEN(LU$ + "TRAENTRE.DAT", 0, 128)
    For KU& = 1 To 1 + LOF(n1%) / 128
      Get #n1%, KU&, BUFLE$
      TTXX$ = Space$(8) + String$(128, 0)
      Call REPLA(TTXX$, "TRAENTRE", 1, 8)
      Call REPLA(TTXX$, BUFLE$, 9, 128)
      Call REGAPP("ATRACEN", TTXX$)
    Next KU&
    Close #n1%
    '
    Call CIERRARCH("ATRACEN")
    Call CIERRARCH("*.*")
    '
    Screen.MousePointer = 1
    '
35  OPX% = COMALTER%("Archivo de Transferencia\\Enviar Vía Web\Descargar en Disco")
    If OPX% < 1 Or OPX% > 2 Then GoTo 99
    If OPX% = 1 Then
        CARPETASINCRO$ = TRIM$(Control("", "CARPEINT"))
        If CARPETASINCRO$ <> "" Then
            ORIGARCH$ = CARPETASINCRO$
            GoTo 45
        Else
            Call COMUNI("Imposible Continuar.\Falta Definir Carpeta de Intercambio")
            GoTo 35
        End If
    End If
    '
    FILEORIG.Caption = "Destino del Archivo"
40  FILEORIG.Show 1
    If FILEORIG.Label1.Caption = "" Then GoTo 99
    '
    LU$ = LUDAT$
    ORIGARCH$ = Left$(FILEORIG.Label1.Caption, 2)
    If Mid$(ORIGARCH$, 2, 1) <> ":" Then
       Call MENSERR(24, "Unidad '" + ORIGARCH$ + "' No Disponible.")
       Exit Sub
    End If
45  On Error GoTo 50
    FileCopy LU$ + "ATRACEN.DAT", ORIGARCH$ + "\ATRACEN.DAT"
    On Error GoTo 0
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Call COMUNI("Archivo de Transferencia Electrónica\para Central Avellaneda\" & REPLACAR("Generado Exitosamente en '" + ORIGARCH$ + "'.", "\", "/"))
    GoTo 99
    '
50  Resume 51
51  On Error GoTo 0
    Call MENSERR(24, "Error de Copiado a Unidad.\   \ Intente de Nuevo.")
    GoTo 99
    '
99  Screen.MousePointer = 1
End Sub

Sub ACTUALIZAMASTER()

    FIN& = ULTREG&("MASTER")
    For REG& = 1 To FIN&
        Call TRACE(24, REG&, FIN&)
        XX$ = REGLEIDO$("MASTER", REG&)
        CI% = CVI(Mid$(XX$, 1, 2))
        ' SI NO EXISTE EL CODIGO LO CREA
        CONDI$ = "CODINT=" & CI%
        If CantRegistros("MASTER", CONDI$, "NOMESS") = 0 Then
            ABUS$ = CStr(CI%)
            Call CreaRegistro("MASTER", "CODINT", ABUS$, "NOMESS")
        End If
        '
        'ACTUALIZA MASTER CON LO QUE VIENE DE AVELLANEDA
        SQLX$ = "SELECT * FROM MASTER WHERE " & CONDI$
        Dim Rst As ADODB.Recordset
        Set Rst = New ADODB.Recordset
        Rst.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        'On Error Resume Next
        With Rst
                CIEXT$ = Mid$(XX$, 77, 16)
            !Codigo = CIEXT$
                DENOMINACION$ = Mid(XX$, 5, 64)
            !Descripcion = DENOMINACION$
                PESU = CVS(Mid$(XX$, 99, 4))
            !PESUNI = PESU
                TIPOMATE% = Asc(Mid$(XX$, 93, 1))
            !TIMATE = TIPOMATE%
                GRUIVA% = Asc(Mid$(XX$, 3, 1))
            !GRUPIVA = GRUIVA%
                CRITREPOS% = Asc(Mid$(XX$, 94, 1))
            !CRITREP = CRITREPOS%
                Monecos% = Asc(Mid$(XX$, 95, 1))
            !Monecos = Monecos%
                TCOLOR$ = Asc(Mid$(XX$, 96, 1))
            !COLOR = TCOLOR$
                TTALLE$ = Asc(Mid$(XX$, 97, 1))
            !TALLE = TTALLE$
                FECHREV% = CVI(Mid$(XX$, 121, 2))
            !Ferevis = CVDAT(FECHREV%)
                STOCKMIN# = CVI(Mid$(XX$, 105, 2))
            '!STOMIN = STOCKMIN# comentado para que no pise los stock minimos generados en la sucursal de comodoro o neuqen
                LOTEREP% = CVI(Mid$(XX$, 107, 2))
            !LOREPOS = LOTEREP%
                LOTEMAI% = CVI(Mid$(XX$, 103, 2))
            '!STOMAX = LOTEMAI%   comentado para que no pise los stock minimos generados en la sucursal de comodoro o neuqen
                Unidad$ = Mid$(XX$, 109, 2)
            !UMEDIDA = Unidad$
                COSTO# = CVS(Mid$(XX$, 111, 4))
            !COSUNI = COSTO#
                FECHACOS% = CVI(Mid$(XX$, 115, 2))
            !FECOSTO = CVDAT(FECHACOS%)
                NACIONA% = CVI(Mid$(XX$, 127, 2))
            !PORNACIO = NACIONA%
                PROBHAB% = CVI(Mid$(XX$, 123, 2))
            !PROVHABI = PROBHAB%
                GRUPOCON% = CVI(Mid$(XX$, 125, 2))
            !Grucoven = GRUPOCON%
                STAT% = 1: If Mid$(XX$, 92, 1) = Chr$(96) Then STAT% = (-1)
            !Status = STAT%
            .Update
        End With
        Rst.Close
        Set Rst = Nothing

10  Next REG&
    
    UPDATMASTER% = 1
    Call GENECOMAS
    UPDATMASTER% = 0
    '
End Sub

Sub RECECOMO() ' recepcion de materiales Comodoro
        '
   'Call BLIMESS("Seleccionando ...")
   Screen.MousePointer = 11
   FIN& = ULTREG&("ATRASLA")
   Call BLANARCH("!ATRASLA")
   CABOPE% = 0
   FF% = HOY(FECHAREC$)
   '
   For i& = 1 To FIN&
     X$ = REGLEIDO$("ATRASLA", i&)
     REMI& = XVALO(Mid$(X$, 6, 11))
     If REMI& = 3247 Then
      A = 1
     End If
     NROPED& = CVS(Mid$(X$, 129, 4))
     If CVD(Mid$(X$, 105, 8)) >= 0.05 Then
       If CVD(Mid$(X$, 117, 8)) < 0.05 Then
         CANOM# = CVD(Mid$(X$, 105, 8))
         Call REPLA(X$, MKD$(CANOM#), 117, 8)
         Call REPLA(X$, MKS$(i&), 125, 4) ' registro
         Call REPLA(X$, MKS$(NROPED&), 129, 4)
         Call REGAPP("!ATRASLA", X$)
       End If
     End If
   Next i&
   '
   Screen.MousePointer = 1
   If ULTREG&("!ATRASLA") < 1 Then
     Call COMUNI("No hay Recepciones\Pendientes de Ingreso.")
     Exit Sub
   End If
   '
   FERECEP$ = REPLACAR$(FECHAREC$, "/", "-")
   VTB% = VENTABU%("RECECOM/NC/TITUPAN=Boleta de Recepción - " + FERECEP$)
   '
   Screen.MousePointer = 11
   '
   JJ& = 0
   If VTB% >= 0 Then
     If COMALTER%("Confirme Ingreso de Materiales Recibidos\\Cancelar\Confirmar Ingreso") = 2 Then
       AINDI% = 0
       For U& = 1 To ULTREG&("!ATRASLA")
         X$ = REGLEIDO$("!ATRASLA", U&)
         APRO# = CVD(Mid$(X$, 117, 8))
         If APRO# > 0.05 Then
           Call REPLA(X$, MKI$(FF%), 115, 2)
           REGORI& = CVS(Mid$(X$, 125, 4))
           If REGORI& > 0 Then
             Y$ = REGLEIDO$("ATRASLA", REGORI&)
           End If
           If REGORI& < 1 Or Left$(Y$, 112) <> Left$(X$, 112) Then
               ' Call REGAPP("ATRASLA", X$)
               REMIT$ = TRIM$(Mid$(Y$, 6, 14))
               ARTICU$ = TRIM$(Mid$(Y$, 41, 16))
               Call MENSERR(24, "Problema al Recepcionar Articulo '" & ARTICU$ & "' del Remito '" & REMIT$ & "'.\Vuelva a Intentarlo.")
               GoTo 30
             Else
               Call GRAREG("ATRASLA", X$, REGORI&)
           End If
           '
           CI1% = CVI(Mid$(X$, 39, 2))
           REFE$ = "Recepción " + FECHAREC$
           NROPED& = CVS(Mid$(X$, 129, 4))
           AINDI% = AINDI% + 1
           LOTE$ = TRIM$(Mid$(X$, 3, 14)) + "-" + TRIM$(Str$(AINDI%))
           DOPRI$ = Mid$(X$, 3, 14)
           DOSEC$ = "BR." + FECHAREC$
           DEPX% = 1
           CANTI = APRO#
           NROORD% = NROORD% + 1
           REMIXTO& = XVALO(Mid$(DOPRI$, 4))
           Call MOVISTOK(FF%, CI1%, LOTE$, "BR", DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NROORD%, VUNI#, 1, NC1%, DEPX%, CANTI, NROPED&)
           Call ACTUCOMPRASUCU(CI1%, CANTI, REMIXTO&, NROPED&)
           'Call Movisto(FF%, CI1%, LOTE$, "BR", DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC1%, DEPX%, CANTI)
         End If
30     Next U&
     End If
   End If
   '
   Call CIERRARCH("MOVISTO")
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
   Screen.MousePointer = 1
   '
End Sub
Sub TRANSFESQL()
    '
    If CLACONTRA% < 1 Then                      ' iNGRESO DE CLAVE Y CONTRASEÑA
       Call MENSERR(24, "Santo Y Seña Inválido")
       Exit Sub
    End If
    '
    Call ABRECONEXION
    '
    Call CREACAMPO("", "ENTRECOM", "Fecha", 8, 0)
    Call CREACAMPO("", "ENTRECOM", "Documento", 10, 14)
    Call CREACAMPO("", "ENTRECOM", "TipoDoc", 10, 16)
    Call CREACAMPO("", "ENTRECOM", "Nume_Cli", 3, 0)
    Call CREACAMPO("", "ENTRECOM", "IDocum", 4, 0)
    Call CREACAMPO("", "ENTRECOM", "Cod_Inte", 3, 0)
    Call CREACAMPO("", "ENTRECOM", "Cod_Exte", 10, 16)
    Call CREACAMPO("", "ENTRECOM", "Cantidad", 7, 0)
    Call CREACAMPO("", "ENTRECOM", "Status", 3, 0)
    Call CREACAMPO("", "ENTRECOM", "VecesEnviado", 3, 0)
    
    For U& = 1 To ULTREG&("ENTRECOM")
        XX$ = REGLEIDO$("ENTRECOM", U&)
        '
        STATU% = CVI(Mid$(XX$, 191, 2))
        If STATU% <> 9 Then        ' NO ESTA ANULADO
            VENVI% = CVI(Mid$(XX$, 193, 2))     ' veces enviado
            If VENVI% < 5 Then
                FF% = CVI(Mid$(XX$, 1, 2))
                REMI$ = Mid$(XX$, 3, 14)
                TIPODO$ = Mid$(XX$, 17, 16)
                NC% = CVI(Mid$(XX$, 33, 2))
                IDOCUM& = CVS(Mid$(XX$, 97, 4))
                OCOMPRA$ = Mid$(XX$, 101, 16)
                CI% = CVI(Mid$(XX$, 117, 2))
                Codigo$ = Mid$(XX$, 119, 16)
                CAN# = CVD(Mid$(XX$, 183, 8))
                '
                CONDI$ = "IDOCUM=" & IDOCUM&
                If CantRegistros("ENTRECOM", CONDI$, "NOMESS") < 1 Then
                    SQLX$ = "SELECT * FROM ENTRECOM WHERE IDOCUM < 1"
                    Dim RST1 As ADODB.Recordset
                    Set RST1 = New ADODB.Recordset
26:                 RST1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
                    If Err Then
                     On Error GoTo 0
                     Call CapturaErrorOpen
                     GoTo 26
                    End If
                    On Error GoTo 0
                    '
                    With RST1
                        .AddNew
                        !FECHA = CVDAT(FF%)
                        !DOCUMENTO = Left(REMI$, 14)
                        !TIPODOC = Left(TIPODOC$, 16)
                        !nume_cli = NC%
                        !IDOCUM = IDOCUM&
                        !cod_inte = CI%
                        !Cod_Exte = Left(Codigo$, 16)
                        !Cantidad = CAN#
                        !Status = STATU%
                        !VECESENVIADO = VENVI%
                        .Update
                    End With
                    '
                    On Error Resume Next
                    RST1.Close
                    Set RST1 = Nothing
                    On Error GoTo 0
                End If
            End If
        End If
    Next U&
    '
    Exit Sub
    '
    '
    OPX% = COMALTER%("Opciones de Transferencia\\Consumo Reparacion\Ajustes de Stock")
    If OPX% < 1 Or OPX% > 2 Then Exit Sub
    '
    If OPX% = 2 Then GoTo 50
    '
    SQLX$ = "SELECT * FROM MOVISTO WITH(NOLOCK) WHERE CODIMOVI='RP'"
    Dim Rst As ADODB.Recordset
    Set Rst = READSET(SQLX$)
    '
    With Rst
        Do While Not .EOF
            CONDI$ = "DOCUMENTO='" & SAFETEXT$(!DoPriCor) & "' AND COD_INTE=" & XVALO(!cod_inte)
            If CantRegistros("ENTRECOM", CONDI$, "NOMESS") < 1 Then
                SQLX$ = "SELECT * FROM ENTRECOM WHERE IDOCUM < 1"
                Dim RST2 As ADODB.Recordset
                Set RST2 = New ADODB.Recordset
27:                 RST2.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
                If Err Then
                 On Error GoTo 0
                 Call CapturaErrorOpen
                 GoTo 27
                End If
                On Error GoTo 0
                '
                FF% = CVINT(Rst!FECHMOV): NROO& = XVALO(Mid$(Rst!DoPriCor, 4))
                IDOCUM& = 10000& * (FF% Mod 1000) + 100& * (NROO& Mod 100)
30:             If CantRegistros("ENTRECOM", "IDOCUM=" & IDOCUM&) > 0 Then
                    IDOCUM& = IDOCUM& + 1
                    GoTo 30
                End If
                '
                With RST2
                    .AddNew
                    !FECHA = Rst!FECHMOV
                    !DOCUMENTO = Left(SAFETEXT$(Rst!DoPriCor), 14)
                    !TIPODOC = "Cons.Reparacion"
                    !nume_cli = XVALO(Rst!Nume_Clie)
                    !IDOCUM = IDOCUM&
                        CI% = XVALO(Rst!cod_inte)
                    !cod_inte = CI%
                    !Cod_Exte = Left(CODEXT$(CI%), 16)
                    !Cantidad = XVALO(Rst!CANTSALID)
                    !Status = 1
                    !VECESENVIADO = 0
                    .Update
                End With
                '
                On Error Resume Next
                RST2.Close
                Set RST2 = Nothing
                On Error GoTo 0
            End If
            .MoveNext
        Loop
    End With
    '
    On Error Resume Next
    Rst.Close
    Set Rst = Nothing
    On Error GoTo 0
    Exit Sub
    '
    ' TRANSFERENCIA DE AJUSTES DE STOCK
50: SQLX$ = "SELECT * FROM MOVISTO WITH(NOLOCK) WHERE CODIMOVI='AJ'"
    Dim RST0 As ADODB.Recordset
    Set RST0 = READSET(SQLX$)
    '
    With RST0
        Do While Not .EOF
            CONDI$ = "DOCUMENTO='" & SAFETEXT$(!DoPriCor) & "' AND COD_INTE=" & XVALO(!cod_inte)
            If CantRegistros("ENTRECOM", CONDI$, "NOMESS") < 1 Then
                SQLX$ = "SELECT * FROM ENTRECOM WHERE IDOCUM < 1"
                Dim RST20 As ADODB.Recordset
                Set RST20 = New ADODB.Recordset
37:                 RST20.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
                If Err Then
                 On Error GoTo 0
                 Call CapturaErrorOpen
                 GoTo 37
                End If
                On Error GoTo 0
                '
                FF% = CVINT(RST0!FECHMOV): NROO& = XVALO(Mid$(RST0!DoPriCor, 4))
                IDOCUM& = 10000& * (FF% Mod 1000) + 100& * (NROO& Mod 100)
40:             If CantRegistros("ENTRECOM", "IDOCUM=" & IDOCUM&) > 0 Then
                    IDOCUM& = IDOCUM& + 1
                    GoTo 40
                End If
                '
                With RST20
                    .AddNew
                    !FECHA = RST0!FECHMOV
                    !DOCUMENTO = Left(SAFETEXT$(RST0!DoPriCor), 14)
                    !TIPODOC = "Ajuste de Stock"
                    !nume_cli = XVALO(RST0!Nume_Clie)
                    !IDOCUM = IDOCUM&
                        CI% = XVALO(RST0!cod_inte)
                    !cod_inte = CI%
                    !Cod_Exte = Left(CODEXT$(CI%), 16)
                    !Cantidad = XVALO(RST0!CANTSALID)
                    !Status = 1
                    !VECESENVIADO = 0
                    .Update
                End With
                '
                On Error Resume Next
                RST20.Close
                Set RST20 = Nothing
                On Error GoTo 0
            End If
            .MoveNext
        Loop
    End With
    '
    On Error Resume Next
    RST0.Close
    Set RST0 = Nothing
    On Error GoTo 0
    
End Sub
Private Sub Form_Load()
   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me)

   Call CREATABLA_TRAZAREC
End Sub
Sub mnuPasarRemitosDropbox_Click()
'1.-borrar los movimientos de la tabla movctp
'2.-pasar primero los del mes de mayo del 2024 y luego ir hacia atras.
'3.-Anotar que meses se registraron
'4.-si se pierde el control consultar en Arcon Avellaneda los meses en la tabla movisto_ctp
   If CLACONTRA% < 1 Then
       Exit Sub
   End If
     CARPETASINCRO$ = RutaCarpetaRemitosNeuquen$
     Ruta$ = CARPETASINCRO$
     If Ruta$ = "" Then
           Call COMUNI("Imposible Continuar.\Falta Definir Ruta de DropBox\En Tablas Auxiliares configurar en la anteultima columna como 'REMITOSNEUQUEN'\ y su ruta en la ultima columna ")
           Exit Sub
     End If

     
     Call CARSELMESA

     Screen.MousePointer = 1
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

     CONTA = 0
     CONDI$ = "FechMov >= '" & Desde & "'"
     CONDI$ = CONDI$ + "AND FechMov <= '" & Hasta & "'"
     CONDI$ = CONDI$ + " AND substring(REFERCOR,1,4) <> 'ANUL' " 'PARA EVITAR COMPROBANTES ARRUINADOS
     CONDI$ = CONDI$ + " AND CodiMovi =  'RT'"
     CONDI$ = CONDI$ + " AND substring(DOPRILAR,1,2) = 'RT' "
     CTRG& = CantRegistros("MOVISTO", CONDI$, "NOMESS")
     
     SQLX$ = "SELECT DOPRILAR,NUPEDCLI FROM MOVISTO WITH(NOLOCK) WHERE " & CONDI$
     Set rs = READSET(SQLX$)
     With rs
       Do While Not .EOF
         CONTA = CONTA + 1
         NROREMITO$ = SAFETEXT$(!DOPRILAR)
         NROPED& = XVALO(!NUPEDCLI)
         Call FOREMIT07.GeneraArchivoRemitoNeuquen(NROREMITO$, NROPED&)
         Frame2(2).Caption = CTRG& - CONTA
         Call T_Espera(1)
        .MoveNext
      Loop
     End With
     rs.Close
     Set rs = Nothing
     
     Frame2(2).Caption = "RESUMENES DE ENTREGAS"


99
End Sub
Private Sub mnuAcumxCliente_Click()
    CMO$ = TRIM$(Control$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
       CMO$ = "RT"
    End If
    '
    Dim PUNMO$(32767)
    '
10  FEX = HOY(HO$)
    Has% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    Des% = FECHANUM(DESS$)
    VDEF$ = MKI$(Des%) + MKI$(Has%)
    '
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    Call BLOQARCH("DTCLIEN")
    '
    Des% = CVI(Left$(OBX$, 2))
    Has% = CVI(Mid$(OBX$, 3, 2))
    If Des% > Has% Then
        GoTo 10
    End If
    FEX = Des%: PERIO$ = FECHATEX$(FEX)
    FEX = Has%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    Screen.MousePointer = 11
    '
    DESD1 = FETEXCOR1$(Des%)
    HASD1 = FETEXCOR1$(Has%)
    '
    SQLX$ = " SELECT * FROM MOVISTO WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE FechMov >= '" & DESD1 & "' "
    SQLX$ = SQLX$ + " AND FechMov <= '" & HASD1 & "' "
     'PARA EVITAR COMPROBANTES ARRUINADOS
    SQLX$ = SQLX$ + " AND LEFT(REFERCOR,4) <> 'ANUL' "
    SQLX$ = SQLX$ + " AND (CodiMovi =  '" & CMO$ & "'  "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV')  AND NUME_CLIE > 0 AND COD_INTE > 0 "
'    sqlx$ = sqlx$ + " GROUP BY Nume_Clie, COD_INTE  "
    SQLX$ = SQLX$ + " ORDER BY Nume_Clie, COD_INTE "
    Call APERBASDAT("STOCKS")
    'Set MOVITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    Dim MOVITMP As ADODB.Recordset
    Set MOVITMP = New ADODB.Recordset
    MOVITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

    With MOVITMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, " No Existen Registros que Cumplan\con la Condición SQL.")
        On Error GoTo 0
        GoTo 10
      End If
      On Error GoTo 0
      '
      Screen.MousePointer = 11
      Call VectCliente.cargarvector
      Call VectArticulos.cargarvector
      Campos$ = "Cta/F7;Razon Social/A32;Código/A16;Descripción/A24;Cant/F9.1;Importe/F12.2;S.T.Cliente/F14.2"

      J& = 0
      REBLB$ = REGBLAN$("CONESREM")
      Do While (.EOF = False)
        NCLI = XVALO(!Nume_Clie)
        If NCLI = 49 Then
         A = 1
        End If
        CII% = XVALO(!cod_inte)
        RASO$ = VectCliente.nombre(NCLI)
        
        If NCLI_ANT > 0 And NCLI <> NCLI_ANT Then
           'SUBTOTAL DEL CLIENTE
           TEXTOGRILLA$ = NCLI_ANT & "|" & RASO_ANT$ & "|" & "--------------------" & "|" & "--------------------" & "|" & "--------------------" & "|" & "--------------------" & "|" & FORMATNUM$(TOT_ACU_CTE#, "F12.2")
           Call CARGA_GRILLA_SHOWREP2(SHOWREP07, SHOWREP07.GRID, Campos$, TEXTOGRILLA$, INICIALIZO%, 1)
           TOT_ACU_CTE# = 0
        End If
        
        If NCLI <> NCLI_ANT Or CII_ANT% <> CII% Then
           TOTI# = 0: CANTU# = 0
        End If
        
        On Error GoTo 0
        CANTU# = CANTU# + XVALO(!CANTSALID) - XVALO(!CANTINGRE)
        TOTI# = TOTI# + (XVALO(!VALOUNIT) * (XVALO(!CANTSALID) - XVALO(!CANTINGRE)))
        TOT_ACU_CTE# = TOT_ACU_CTE# + (XVALO(!VALOUNIT) * (XVALO(!CANTSALID) - XVALO(!CANTINGRE)))
        KODIX$ = VectArticulos.Codigo(CII%)
        DESCRIPCIONX$ = VectArticulos.Descripcion(CII%)
        '
        NCLI_ANT = NCLI
        CII_ANT% = CII%
        RASO_ANT$ = RASO$

        .MoveNext
        
NCLI1 = XVALO(!Nume_Clie)
CI1% = XVALO(!cod_inte)
If NCLI1 <> NCLI_ANT Or CII_ANT% <> CI1% Then
'IMPRIME EL ACUMULADO POR ITEM
TEXTOGRILLA$ = NCLI & "|" & RASO$ & "|" & KODIX$ & "|" & DESCRIPCIONX$ & "|" & FORMATNUM$(CANTU#, "F9.1") & "|" & FORMATNUM$(TOTI#, "F12.2")
Call CARGA_GRILLA_SHOWREP2(SHOWREP07, SHOWREP07.GRID, Campos$, TEXTOGRILLA$, INICIALIZO%)
End If
      Loop
    End With
    Set MOVITMP = Nothing
    '
'MUESTRA FRM DE GENREP07
SHOWREP07.Label7 = SHOWREP07.GRID.Rows - 1
SHOWREP07.Frame1.Visible = False
SHOWREP07.Label3 = "Resumen de Entrega x Cliente (Acumulado)  Periodo: " + PERIO$
Screen.MousePointer = 1
SHOWREP07.NO_ORDENAR = 1
Call COLOREAR_GRILLASQL2(SHOWREP07.GRID, &HC0C0C0, 1)

SHOWREP07.Show 1
99
Screen.MousePointer = 1

End Sub

Private Sub mnuDetxCliente_Click()
    '
    CMO$ = TRIM$(Control$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
       CMO$ = "RT"
    End If
    '
    Dim PUNMO$(32767)
    '
10  FEX = HOY(HO$)
    Has% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    Des% = FECHANUM(DESS$)
    VDEF$ = MKI$(Des%) + MKI$(Has%)
    '
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    Call BLOQARCH("DTCLIEN")
    '
    Des% = CVI(Left$(OBX$, 2))
    Has% = CVI(Mid$(OBX$, 3, 2))
    If Des% > Has% Then
        GoTo 10
    End If
    FEX = Des%: PERIO$ = FECHATEX$(FEX)
    FEX = Has%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    Screen.MousePointer = 11
    Call HEADERS("DTCLIEN", "")
    Call HEADERS("DTCLIEN", "Periodo: " + PERIO$)
    '
    DESD1 = FETEXCOR1$(Des%)
    HASD1 = FETEXCOR1$(Has%)
    '
    SQLX$ = " SELECT * FROM MOVISTO WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE FechMov >= '" & DESD1 & "' "
    SQLX$ = SQLX$ + " AND FechMov <= '" & HASD1 & "' "
     'PARA EVITAR COMPROBANTES ARRUINADOS
    SQLX$ = SQLX$ + " AND LEFT(REFERCOR,4) <> 'ANUL' "
    SQLX$ = SQLX$ + " AND (CodiMovi =  '" & CMO$ & "'  "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV')  AND NUME_CLIE > 0 AND COD_INTE > 0 "
    SQLX$ = SQLX$ + " ORDER BY Nume_Clie "
    Call APERBASDAT("STOCKS")
    'Set MOVITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    Dim MOVITMP As ADODB.Recordset
    Set MOVITMP = New ADODB.Recordset
    MOVITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

    With MOVITMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, " No Existen Registros que Cumplan\con la Condición SQL.")
        On Error GoTo 0
        GoTo 10
      End If
      On Error GoTo 0
      '
      Screen.MousePointer = 11
      Call VectCliente.cargarvector
      Call VectArticulos.cargarvector

      J& = 0
      REBLB$ = REGBLAN$("CONESREM")
      Do While (.EOF = False)
        NCLI = !Nume_Clie
        REMIFAC$ = !DoPriCor
        OBSERX$ = ""
        On Error Resume Next
        NUMEFAC$ = TRIM$(!DoSecCor)
          If Left$(NUMEFAC$, 2) <> "FC" Then
            NUMEFAC$ = ""
            If Left$(UCase$(!DOSECLAR), 9) = "SIN CARGO" Then
              OBSERX$ = !DOSECLAR
            End If
          End If
        On Error GoTo 0
        NPEDX& = XVALO(!NUPEDCLI)
        CANTU# = XVALO(!CANTSALID) - XVALO(!CANTINGRE)
        CII% = !cod_inte
        NOCOMX$ = NUOCOCLI$(NPEDX&, CII%)
        FF% = CVINT(!FECHMOV)
        If Left$(REMIFAC$, 4) = "DV.0" Then
            REMIFAC$ = !DoSecCor
        End If
        '
        TOTI# = !VALOUNIT * (XVALO(!CANTSALID) - XVALO(!CANTINGRE))
        RASO$ = VectCliente.nombre(NCLI)
        KODIX$ = VectArticulos.Codigo(CII%)
        DESCRIPCIONX$ = VectArticulos.Descripcion(CII%)
        '
TEXTOGRILLA$ = NCLI & "|" & RASO$ & "|" & FECHATEX(FF%) & "|" & REMIFAC$ & "|"
TEXTOGRILLA$ = TEXTOGRILLA$ & NUMEFAC$ & "|" & NPEDX& & "|" & NOCOMX$ & "|"
TEXTOGRILLA$ = TEXTOGRILLA$ & KODIX$ & "|" & DESCRIPCIONX$ & "|" & FORMATNUM$(CANTU#, "F9.1") & "|" & FORMATNUM$(XVALO(!VALOUNIT), "F12.2") & "|"
TEXTOGRILLA$ = TEXTOGRILLA$ & FORMATNUM$(TOTI#, "F12.2") & "|" & OBSERX$
Campos$ = "Cta/F7;Razon Social/A32;F.Mov./D8;Remito/A14;Factura/A18;Pedido/F8;O.Comp./A12;Código/A16;Descripción/A24;Cant/F9.1;P.Unit./F12.2;Importe/F12.2;Observaciones/A24"
Call CARGA_GRILLA_SHOWREP2(SHOWREP07, SHOWREP07.GRID, Campos$, TEXTOGRILLA$, INICIALIZO%)
        
        .MoveNext
      Loop
    End With
    Set MOVITMP = Nothing
    '

SHOWREP07.Label7 = SHOWREP07.GRID.Rows - 1
Screen.MousePointer = 1
SHOWREP07.Frame1.Visible = False
SHOWREP07.Label3 = "Resumen de Entrega x Cliente (Detalle)  Periodo: " + PERIO$
Call COLOREAR_GRILLASQL2(SHOWREP07.GRID, &HC0C0C0, 1)
Screen.MousePointer = 1
SHOWREP07.Show 1
99
Screen.MousePointer = 1

End Sub

Private Sub mnuRegRemParaenvap_Click()
   If CLACONTRA% < 1 Then
       Exit Sub
   End If
   
10  TituloLista$ = "Lista de Remitos "
    Campos$ = "Fecha/D8;Nro.Remito/A20;Cta./f8;Razón Social/A40"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$

    Dim RS1 As New ADODB.Recordset
    
12  SQLX$ = " SELECT FechMov,DOPRILAR,Nume_Clie"
    SQLX$ = SQLX$ + " FROM MOVISTO  WITH(NOLOCK)"
    SQLX$ = SQLX$ + " GROUP BY FechMov,DOPRILAR,Nume_Clie"
    SQLX$ = SQLX$ + " ORDER BY FechMov DESC,DOPRILAR DESC"
    
    
    Set RS1 = READSET(SQLX$)
              
    
    Call Carga_MSF_Recordset(RS1, Campos$, FRMZELECHO.msf2, 1)
    
    On Error Resume Next
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    On Error GoTo 0
    Screen.MousePointer = 1
    FRMZELECHO.Width = 15000
    FRMZELECHO.Show 1
     '
     Screen.MousePointer = 1
     '
     NROREMIT$ = TRIM$(RESP_ZELE_VECT(2))
     If NROREMIT$ = "" Then Exit Sub

     If NROREMIT$ <> "" Then
       Call FOREMIT07.GeneraArchivoRemitoConteplast(NROREMIT$, 0)
     End If

End Sub

Private Sub mnuSalir_Click()
   Call Command2_Click
End Sub


 Function RemtiosSinPedidoVinculado$()
    Erase RESP_ZELE_VECT
    RemtiosSinPedidoVinculado$ = ""
    Campos$ = "Remito/A18;Fecha/D8;N°Cta./F9;Razòn Social/A40"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = ""
    Dim obj As New RECORXET
    CAMPOSTABLA$ = "DoPriLar,fechmov,nume_clie,Raso_Cli"
    TABLA$ = " movisto inner join Deudor12 with(nolock) on Nume_Cli = Nume_Clie "
    CONDI$ = " nupedcli < 1 and codimovi = 'rt' group by FechMov,DoPriLar,nume_clie,Raso_Cli order by doprilar"
    
    Set RS1 = obj.RS_GENERAL(TABLA$, CAMPOSTABLA$, CONDI$)
    Call Carga_MSF_Recordset(RS1, Campos$, FRMZELECHO.msf2, 1)
    On Error Resume Next
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 10000
50  FRMZELECHO.Show 1
    If TRIM$(RESP_ZELE_VECT(1)) <> "" Then '
       RemtiosSinPedidoVinculado$ = RESP_ZELE_VECT(1)
    End If
End Function
Sub CodigodeRemitoSelec(NROREMITO$)
    Erase RESP_ZELE_VECT

    Campos$ = "Remito/A18;Fecha/D8;N°Cta./F9;Razòn Social/A40,CI/A0;Còdigo/A16;Descrip./A32;Idenlote/A16;Salida/f9.1;Ingreso/F9.1;Pedido/F8,ID/A0"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = ""
    Dim obj As New RECORXET
    CAMPOSTABLA$ = " DoPriLar,fechmov,nume_clie,Raso_Cli,cod_Inte,Cod_Exte,Descrip,idenlote,sum(CantSalid),sum(Cantingre),NuPedCli,MOVISTO_ID"
    TABLA$ = " movisto with(nolock)  inner join Deudor12 on Nume_Cli = Nume_Clie "
    CONDI$ = CONDI$ + " codimovi = 'rt' and doprilar = '" & NROREMITO$ & "' group by FechMov,DoPriLar ,nume_clie,Raso_Cli,Cod_Inte,Cod_Exte,Descrip,Idenlote,CantSalid,CantSalid,NuPedCli,MOVISTO_ID order by doprilar"
    
    Set RS1 = obj.RS_GENERAL(TABLA$, CAMPOSTABLA$, CONDI$)
    'Call Carga_MSF_Recordset(RS1, Campos$, FRMZELECHO.MSF2, 1)
    Call Carga_MSF_Recordset(RS1, Campos$, VENTABNEW.MSF)
    On Error Resume Next
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 18000
50  FRMZELECHO.Show 1
    If XVALO(RESP_ZELE_VECT(5)) > 0 Then '
       'CodigodeRemitoSelec& = XVALO(RESP_ZELE_VECT(5))
    End If
    
    
        Unload VENTABNEW
        VENTANA.Inicializar = 0
        VENTANA.CampEditables = "11"
        VENTANA.AgregaRegistro = 0
        VENTANA.UtilizaFormula = 1
        VENTANA.HabilitaBorrar = 1
        VENTANA.HabilitaInsertar = 1
        VENTANA.NoMostrarBotonTilde = 0
        VENTANA.TITULO = "Vincular Pedido a Entregas del Remito: " & NROREMITO$
        

        VENTANA.Campox = Campos$
        VENTANA.Inicializar = 1
'        Call numerargrilla(VENTABNEW.MSF, 0, 0, 0, 0)
        Dim MSF1 As msFlexGrid
        Set MSF1 = VENTABNEW.MSF

30      VENTABNEW.Show 1
        If VENTABNEW.APROBADO% < 1 Then Exit Sub
        '
        'PRIMERO VALIDO
        For REG& = 1 To MSF1.Rows - 1
                  CI1% = XVALO(MSF1.TextMatrix(REG&, 5))
                  LOTE$ = MSF1.TextMatrix(REG&, 8)
                  NROPED& = XVALO(MSF1.TextMatrix(REG&, 11))
                  
                  If CI1% > 0 And NROPED& > 0 Then
                     If CantidadPedidaItemEnPedido&(NROPED&, CI1%) < 1 Then
                        Call COMUNI("El Item: " & TRIM$(MSF1.TextMatrix(REG&, 6)) & " En Fila: " & REG& & " No Corresponde al Pedido: " & NROPED&)
                        GoTo 30
                     End If
                  End If
        Next REG&
        'GRABO EN EL MOVISTO EL NRO DE PEDIDO A CADA ITEM DEL REMITO QUE POSEA NRO DE PEDIDO
        For REG& = 1 To MSF1.Rows - 1
                  CI1% = XVALO(MSF1.TextMatrix(REG&, 5))
                  NROPED& = XVALO(MSF1.TextMatrix(REG&, 11))
                  idMov& = XVALO(MSF1.TextMatrix(REG&, 12))
                  
                  If CI1% > 0 And NROPED& > 0 Then
                     CONDI$ = " MOVISTO_ID = " & idMov&
                     SQLX$ = "UPDATE MOVISTO SET NuPedCli = 1 NROPED&"
        FLEXCONN.Execute (SQLX$)

                  End If
        Next REG&
        

        'SI PASO VALIDACION GRABA
        'TRANSACCIÓN GRABACION
        On Error GoTo ERROR_GUARDAR
        FLEXCONN.BeginTrans

        CONDI$ = " CONFSCRAP_ID > 0"
        SQLX$ = "UPDATE CONFSCRAP SET MARBORRA = 1 WHERE " & CONDI$
        FLEXCONN.Execute (SQLX$)
        
        SQLX$ = "SELECT * FROM CONFSCRAP"
        SQLX$ = SQLX$ + " WHERE CONFSCRAP_ID < 1"
        FE% = HOY(HO$)
        FEC$ = FETEXCOR1$(FE%)
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

        With rs
              For REG& = 1 To MSF1.Rows - 1
                  .AddNew
                  !CodSector = XVALO(MSF1.TextMatrix(REG&, 1))
                  !DescSector = Left$(MSF1.TextMatrix(REG&, 2), 32)
                  !CodCateg = XVALO(MSF1.TextMatrix(REG&, 3))
                  !DescCateg = Left$(MSF1.TextMatrix(REG&, 4), 64)
                  !CodArt = Left$(MSF1.TextMatrix(REG&, 5), 24)
                  !DescArt = Left$(MSF1.TextMatrix(REG&, 6), 64)
                  !DepIngre = XVALO(MSF1.TextMatrix(REG&, 7))
                  !CiArt = CODINT%(Left$(MSF1.TextMatrix(REG&, 5), 24))
                  !FECHA = FEC$
                  !CODIEMPR = CodiEmp%
                  !MARBORRA = 0
                  .Update
              Next REG&
        End With
        
        rs.Close
        Set rs = Nothing
        
        CONDI$ = CONDI$ + " AND  MARBORRA = 1  "
        SQLX$ = "DELETE FROM CONFSCRAP WHERE " & CONDI$
        FLEXCONN.Execute (SQLX$)
        
        Set VENTANA = Nothing
        
ERROR_GUARDAR:
    If Err <> 0 Then
       FLEXCONN.RollbackTrans
       On Error GoTo 0
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION")
       Exit Sub
    Else
      FLEXCONN.CommitTrans
      Call COMUNI("Grabación Realizado Exitosamente")
    End If
    Unload VENTABNEW



End Sub



Private Sub mnuVinculaPediRemi_Click()
10   Remito$ = TRIM$(RemtiosSinPedidoVinculado$)
  If Remito$ = "" Then Exit Sub
  Call CodigodeRemitoSelec&(Remito$)
  If CI1% < 1 Then GoTo 10
  
      'Campos$ = "Remito/A18;Fecha/D8;N°Cta./F9;Razòn Social/A40,CI/A0;Còdigo/A16;Descrip./A32;Idenlote/A16;Salida/f9.1;Ingreso/F9.1;Pedido/F8"

End Sub
