VERSION 5.00
Begin VB.Form MENDEPPT 
   BackColor       =   &H00F7E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Demostración FLEXOFT - Distribución de Secuencias"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
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
      Left            =   6435
      Picture         =   "MENDEPPT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   84
      ToolTipText     =   "PhamaFlex"
      Top             =   4305
      Width           =   560
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   6420
      ScaleHeight     =   585
      ScaleWidth      =   2160
      TabIndex        =   81
      Top             =   4970
      Width           =   2220
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
         Picture         =   "MENDEPPT.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   83
         ToolTipText     =   "Batch Record Electronico"
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Batch Record Electronico"
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
         Left            =   690
         TabIndex        =   82
         Top             =   90
         Width           =   1485
      End
   End
   Begin VB.PictureBox Picture25 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   9555
      ScaleHeight     =   585
      ScaleWidth      =   2055
      TabIndex        =   77
      Top             =   2680
      Width           =   2115
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
         Height          =   600
         Left            =   0
         Picture         =   "MENDEPPT.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Maestro de Clientes"
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Maestro de Clientes"
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
         TabIndex        =   78
         Top             =   105
         Width           =   1395
      End
   End
   Begin VB.PictureBox Picture24 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   9555
      ScaleHeight     =   585
      ScaleWidth      =   2055
      TabIndex        =   75
      Top             =   3360
      Width           =   2115
      Begin VB.CommandButton Command24 
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
         Picture         =   "MENDEPPT.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "Maestro de Proveedores"
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Maestro de Proveedores"
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
         TabIndex        =   76
         Top             =   105
         Width           =   1395
      End
   End
   Begin VB.PictureBox Picture23 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   3360
      ScaleHeight     =   585
      ScaleWidth      =   2055
      TabIndex        =   72
      Top             =   4270
      Width           =   2115
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
         Height          =   600
         Left            =   0
         Picture         =   "MENDEPPT.frx":0D60
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Gestión de Stock e Inventarios"
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Gestión Stock e Inventarios"
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
         TabIndex        =   74
         Top             =   105
         Width           =   1275
      End
   End
   Begin VB.PictureBox Picture22 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   3360
      ScaleHeight     =   585
      ScaleWidth      =   2055
      TabIndex        =   69
      Top             =   7680
      Width           =   2115
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
         Height          =   600
         Left            =   0
         Picture         =   "MENDEPPT.frx":1B8E
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Despacho de Materiales"
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Despacho de Materiales"
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
         TabIndex        =   70
         Top             =   105
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture21 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   6405
      ScaleHeight     =   585
      ScaleWidth      =   2160
      TabIndex        =   66
      Top             =   7680
      Width           =   2220
      Begin VB.CommandButton Command21 
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
         Picture         =   "MENDEPPT.frx":1E98
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "Sistema Gestión de Mantenimiento Industrial"
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Mantenimiento Industrial"
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
         TabIndex        =   68
         Top             =   105
         Width           =   1485
      End
   End
   Begin VB.PictureBox Picture20 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   9555
      ScaleHeight     =   585
      ScaleWidth      =   2055
      TabIndex        =   63
      Top             =   6780
      Width           =   2115
      Begin VB.CommandButton Command20 
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
         Picture         =   "MENDEPPT.frx":21A2
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Sistema de Planeamiento Financiero"
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Presupuesto Financiero "
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
         Left            =   615
         TabIndex        =   65
         Top             =   105
         Width           =   1485
      End
   End
   Begin VB.PictureBox Picture19 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   6405
      ScaleHeight     =   585
      ScaleWidth      =   2160
      TabIndex        =   60
      Top             =   6330
      Width           =   2220
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
         Height          =   600
         Left            =   0
         Picture         =   "MENDEPPT.frx":24AC
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Sistema de Administración de Documentos según ISO 9000"
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Documentación ISO 9000"
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
         TabIndex        =   62
         Top             =   105
         Width           =   1485
      End
   End
   Begin VB.PictureBox Picture18 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   6405
      ScaleHeight     =   585
      ScaleWidth      =   2160
      TabIndex        =   57
      Top             =   5650
      Width           =   2220
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
         Left            =   0
         Picture         =   "MENDEPPT.frx":308E
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   " Administración de Documentos Gráficos"
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Planera Electrónica"
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
         TabIndex        =   59
         Top             =   105
         Width           =   1485
      End
   End
   Begin VB.PictureBox Picture17 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   6405
      ScaleHeight     =   585
      ScaleWidth      =   2160
      TabIndex        =   55
      Top             =   4270
      Width           =   2220
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "PharmaFlex"
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
         Left            =   690
         TabIndex        =   56
         Top             =   90
         Width           =   1485
      End
   End
   Begin VB.PictureBox Picture16 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   6405
      ScaleHeight     =   585
      ScaleWidth      =   2160
      TabIndex        =   52
      Top             =   3600
      Width           =   2220
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
         Left            =   0
         Picture         =   "MENDEPPT.frx":3398
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Maestro de Stock"
         Top             =   15
         Width           =   560
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Maestro de Stock"
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
         TabIndex        =   54
         Top             =   90
         Width           =   1485
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C000&
      Height          =   580
      Index           =   4
      Left            =   2835
      ScaleHeight     =   525
      ScaleWidth      =   2055
      TabIndex        =   48
      Top             =   2835
      Width           =   2115
      Begin VB.Image Image5 
         Height          =   375
         Index           =   1
         Left            =   70
         Picture         =   "MENDEPPT.frx":4142
         Stretch         =   -1  'True
         ToolTipText     =   "Sistemas de Control de Fabricación"
         Top             =   70
         Width           =   420
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Control de Manufactura"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   3
         Left            =   525
         TabIndex        =   49
         Top             =   70
         Width           =   1485
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H008080FF&
      Height          =   580
      Index           =   3
      Left            =   210
      ScaleHeight     =   525
      ScaleWidth      =   2055
      TabIndex        =   46
      Top             =   1550
      Width           =   2115
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Características Generales"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   5
         Left            =   525
         TabIndex        =   47
         Top             =   70
         Width           =   1485
      End
      Begin VB.Image Image5 
         Height          =   375
         Index           =   0
         Left            =   70
         Picture         =   "MENDEPPT.frx":444C
         Stretch         =   -1  'True
         Top             =   70
         Width           =   420
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   580
      Index           =   2
      Left            =   8310
      ScaleHeight     =   525
      ScaleWidth      =   2055
      TabIndex        =   44
      Top             =   1550
      Width           =   2115
      Begin VB.Image Image4 
         Height          =   270
         Left            =   1680
         Picture         =   "MENDEPPT.frx":4756
         Stretch         =   -1  'True
         Top             =   40
         Width           =   300
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Gestion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   1200
         TabIndex        =   45
         Top             =   315
         Width           =   1275
      End
      Begin VB.Image Image2 
         Height          =   525
         Index           =   2
         Left            =   -50
         Picture         =   "MENDEPPT.frx":48A0
         ToolTipText     =   "Sub Sistema de Gestión Administrativo - Financiera"
         Top             =   0
         Width           =   5850
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   580
      Index           =   1
      Left            =   4310
      ScaleHeight     =   525
      ScaleWidth      =   2055
      TabIndex        =   42
      Top             =   1550
      Width           =   2115
      Begin VB.Image Image3 
         Height          =   270
         Left            =   1680
         Picture         =   "MENDEPPT.frx":E91E
         Stretch         =   -1  'True
         Top             =   40
         Width           =   300
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Industrial"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   1100
         TabIndex        =   43
         Top             =   315
         Width           =   1275
      End
      Begin VB.Image Image2 
         Height          =   525
         Index           =   1
         Left            =   -50
         Picture         =   "MENDEPPT.frx":EC28
         ToolTipText     =   "Sub-Sistema de Gestión Industrial"
         Top             =   0
         Width           =   5850
      End
   End
   Begin VB.PictureBox Picture12 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   6405
      ScaleHeight     =   585
      ScaleWidth      =   2160
      TabIndex        =   38
      Top             =   7000
      Width           =   2220
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
         Picture         =   "MENDEPPT.frx":18CA6
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Sistema de Análisis de Costos y Presupuestación de Productos"
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Presupuesto    de Producto"
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
         TabIndex        =   40
         Top             =   105
         Width           =   1485
      End
   End
   Begin VB.PictureBox Picture10 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   3360
      ScaleHeight     =   585
      ScaleWidth      =   2055
      TabIndex        =   35
      Top             =   7000
      Width           =   2115
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
         Picture         =   "MENDEPPT.frx":18FB0
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Gestion de Compras de Materiales y Servicios Productivos"
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Gestion de Compras"
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
         TabIndex        =   37
         Top             =   105
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command9 
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
      Height          =   750
      Left            =   10920
      Picture         =   "MENDEPPT.frx":19B1A
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   210
      Width           =   750
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   630
      ScaleHeight     =   585
      ScaleWidth      =   2055
      TabIndex        =   31
      Top             =   4305
      Width           =   2115
      Begin VB.CommandButton Command1 
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
         Picture         =   "MENDEPPT.frx":19C64
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Instrucciones para Ejecutar este Programa de Demostración"
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " Cómo usar esta Demo"
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
         TabIndex        =   33
         Top             =   105
         Width           =   1275
      End
   End
   Begin VB.PictureBox Picture15 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   3360
      ScaleHeight     =   585
      ScaleWidth      =   2055
      TabIndex        =   28
      Top             =   3600
      Width           =   2115
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
         Height          =   600
         Left            =   0
         Picture         =   "MENDEPPT.frx":1B95E
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Deposito Caotico"
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Depósito Caótico"
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
         TabIndex        =   30
         Top             =   105
         Width           =   1275
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   3360
      ScaleHeight     =   585
      ScaleWidth      =   2055
      TabIndex        =   25
      Top             =   4970
      Width           =   2115
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
         Picture         =   "MENDEPPT.frx":1C8D0
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Ordenes de Fabricacion"
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Ordenes de Fabricación"
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
         Top             =   105
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture13 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   3360
      ScaleHeight     =   585
      ScaleWidth      =   2055
      TabIndex        =   22
      Top             =   5650
      Width           =   2115
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
         Picture         =   "MENDEPPT.frx":1D6E2
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Programacion de Materiales"
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Programacion      de Materiales"
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
         TabIndex        =   24
         Top             =   105
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture11 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   3360
      ScaleHeight     =   585
      ScaleWidth      =   2055
      TabIndex        =   19
      Top             =   6330
      Width           =   2115
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
         Picture         =   "MENDEPPT.frx":1E48C
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Secuenciación de Carga de Máquinas"
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Secuencia de      Carga Maquina"
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
         Left            =   705
         TabIndex        =   21
         Top             =   105
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture9 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   9555
      ScaleHeight     =   585
      ScaleWidth      =   2055
      TabIndex        =   16
      Top             =   4035
      Width           =   2115
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
         Height          =   600
         Left            =   0
         Picture         =   "MENDEPPT.frx":1F0CE
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   " Gestión de Ventas y Cuentas Corrientes"
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Gestion Ventas y Deudores"
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
         TabIndex        =   18
         Top             =   105
         Width           =   1395
      End
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   9555
      ScaleHeight     =   585
      ScaleWidth      =   2055
      TabIndex        =   13
      Top             =   4740
      Width           =   2115
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
         Height          =   600
         Left            =   0
         Picture         =   "MENDEPPT.frx":1FCC8
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Gestión de Acreedores Comerciales"
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Acreedores Comerciales"
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
         Left            =   525
         TabIndex        =   15
         Top             =   105
         Width           =   1500
      End
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   9555
      ScaleHeight     =   585
      ScaleWidth      =   2055
      TabIndex        =   10
      Top             =   5430
      Width           =   2115
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
         Left            =   0
         Picture         =   "MENDEPPT.frx":20A52
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Contabiliodad General"
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Contabilidad General"
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
         Left            =   525
         TabIndex        =   12
         Top             =   105
         Width           =   1500
      End
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   9555
      ScaleHeight     =   585
      ScaleWidth      =   2055
      TabIndex        =   7
      Top             =   6105
      Width           =   2115
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
         Left            =   0
         Picture         =   "MENDEPPT.frx":216C4
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Sistema de Tesorería"
         Top             =   -15
         Width           =   560
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tesoreria y Gestion Valores"
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
         TabIndex        =   9
         Top             =   105
         Width           =   1485
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   630
      ScaleHeight     =   585
      ScaleWidth      =   2055
      TabIndex        =   4
      Top             =   5985
      Width           =   2115
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
         Height          =   600
         Left            =   0
         Picture         =   "MENDEPPT.frx":22386
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Genere su Fax Pidiendo el Envío de más Información"
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Pedir más Información "
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
         TabIndex        =   6
         Top             =   105
         Width           =   1170
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   630
      ScaleHeight     =   585
      ScaleWidth      =   2055
      TabIndex        =   1
      Top             =   5145
      Width           =   2115
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
         Left            =   -15
         Picture         =   "MENDEPPT.frx":22690
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Características Generales de los Sistemas FLEXOFT"
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Características Generales"
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
         Left            =   525
         TabIndex        =   3
         Top             =   105
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   550
      Index           =   0
      Left            =   4310
      ScaleHeight     =   495
      ScaleWidth      =   1980
      TabIndex        =   0
      Top             =   210
      Width           =   2040
      Begin VB.Image Image1 
         Height          =   270
         Index           =   0
         Left            =   1650
         Picture         =   "MENDEPPT.frx":227DA
         Stretch         =   -1  'True
         Top             =   45
         Width           =   315
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Soluciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   1050
         TabIndex        =   41
         Top             =   315
         Width           =   1275
      End
      Begin VB.Image Image2 
         Height          =   525
         Index           =   0
         Left            =   -50
         Picture         =   "MENDEPPT.frx":22AE4
         Top             =   -10
         Width           =   5850
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000C0C0&
      Height          =   580
      Index           =   5
      Left            =   5775
      ScaleHeight     =   525
      ScaleWidth      =   2055
      TabIndex        =   50
      Top             =   2835
      Width           =   2115
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ingenieria de Producto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   4
         Left            =   525
         TabIndex        =   51
         Top             =   70
         Width           =   1485
      End
      Begin VB.Image Image5 
         Height          =   375
         Index           =   2
         Left            =   70
         Picture         =   "MENDEPPT.frx":2CB62
         Stretch         =   -1  'True
         ToolTipText     =   "Sistemas de Ingeniería de Producto"
         Top             =   70
         Width           =   420
      End
   End
   Begin VB.Line Line11 
      X1              =   6090
      X2              =   6450
      Y1              =   5310
      Y2              =   5310
   End
   Begin VB.Line Line35 
      X1              =   9360
      X2              =   9570
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line34 
      X1              =   9360
      X2              =   9570
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line33 
      X1              =   3150
      X2              =   3360
      Y1              =   7875
      Y2              =   7875
   End
   Begin VB.Line Line32 
      X1              =   6090
      X2              =   6405
      Y1              =   7875
      Y2              =   7875
   End
   Begin VB.Line Line31 
      X1              =   9345
      X2              =   9660
      Y1              =   6975
      Y2              =   6975
   End
   Begin VB.Line Line30 
      X1              =   6090
      X2              =   6510
      Y1              =   3945
      Y2              =   3945
   End
   Begin VB.Line Line29 
      X1              =   6090
      X2              =   6615
      Y1              =   4605
      Y2              =   4605
   End
   Begin VB.Line Line27 
      X1              =   6090
      X2              =   6615
      Y1              =   6405
      Y2              =   6405
   End
   Begin VB.Line Line26 
      X1              =   6090
      X2              =   6405
      Y1              =   7140
      Y2              =   7140
   End
   Begin VB.Line Line25 
      X1              =   6090
      X2              =   6090
      Y1              =   3360
      Y2              =   7875
   End
   Begin VB.Line Line24 
      X1              =   6825
      X2              =   6825
      Y1              =   2415
      Y2              =   2940
   End
   Begin VB.Line Line23 
      X1              =   3885
      X2              =   3885
      Y1              =   2415
      Y2              =   2835
   End
   Begin VB.Line Line5 
      X1              =   3885
      X2              =   6825
      Y1              =   2415
      Y2              =   2415
   End
   Begin VB.Line Line2 
      X1              =   4200
      X2              =   4410
      Y1              =   6615
      Y2              =   6615
   End
   Begin VB.Line Line1 
      X1              =   3150
      X2              =   3360
      Y1              =   7140
      Y2              =   7140
   End
   Begin VB.Line Line22 
      X1              =   9345
      X2              =   9555
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line21 
      X1              =   9345
      X2              =   9555
      Y1              =   5610
      Y2              =   5610
   End
   Begin VB.Line Line20 
      X1              =   9345
      X2              =   9555
      Y1              =   4905
      Y2              =   4905
   End
   Begin VB.Line Line19 
      X1              =   9345
      X2              =   9555
      Y1              =   4260
      Y2              =   4260
   End
   Begin VB.Line Line6 
      X1              =   9345
      X2              =   9345
      Y1              =   2100
      Y2              =   7000
   End
   Begin VB.Line Line18 
      X1              =   3150
      X2              =   3360
      Y1              =   6405
      Y2              =   6405
   End
   Begin VB.Line Line17 
      X1              =   3150
      X2              =   3360
      Y1              =   5300
      Y2              =   5300
   End
   Begin VB.Line Line16 
      X1              =   3150
      X2              =   3465
      Y1              =   4600
      Y2              =   4600
   End
   Begin VB.Line Line15 
      X1              =   3150
      X2              =   3360
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line14 
      X1              =   3150
      X2              =   3150
      Y1              =   3150
      Y2              =   7875
   End
   Begin VB.Line Line13 
      X1              =   1260
      X2              =   1260
      Y1              =   1680
      Y2              =   1155
   End
   Begin VB.Line Line12 
      X1              =   1260
      X2              =   9345
      Y1              =   1155
      Y2              =   1155
   End
   Begin VB.Line Line10 
      X1              =   420
      X2              =   630
      Y1              =   6300
      Y2              =   6300
   End
   Begin VB.Line Line9 
      X1              =   420
      X2              =   630
      Y1              =   5460
      Y2              =   5460
   End
   Begin VB.Line Line8 
      X1              =   420
      X2              =   630
      Y1              =   4620
      Y2              =   4620
   End
   Begin VB.Line Line7 
      X1              =   405
      X2              =   405
      Y1              =   1965
      Y2              =   6315
   End
   Begin VB.Line Line4 
      X1              =   9345
      X2              =   9345
      Y1              =   1155
      Y2              =   1680
   End
   Begin VB.Line Line3 
      X1              =   5355
      X2              =   5355
      Y1              =   525
      Y2              =   2415
   End
   Begin VB.Image Image1 
      Height          =   8700
      Index           =   1
      Left            =   -15
      Picture         =   "MENDEPPT.frx":2CE6C
      Stretch         =   -1  'True
      Top             =   -180
      Width           =   12000
   End
End
Attribute VB_Name = "MENDEPPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command1_Click()
'   Call SECUDEMO("INSTRU")
    Call ShellExecute(Me.hwnd, "Open", "Instrucciones demo.pps", "", "", 1)
End Sub

Private Sub Command10_Click()
   'Call SECUDEMO("COMPRA")
   Call ShellExecute(Me.hwnd, "Open", "Gestion de Compras.pps", "", "", 1)
End Sub

Private Sub Command11_Click()
   'Call SECUDEMO("CARMAQ")
   Call ShellExecute(Me.hwnd, "Open", "Secuencia de Carga de Maquinas.pps", "", "", 1)
End Sub

Private Sub Command12_Click()
   'Call SECUDEMO("PROMAT")
   Call ShellExecute(Me.hwnd, "Open", "Programacion de Materiales.pps", "", "", 1)
End Sub

Private Sub Command13_Click()
   'Call SECUDEMO("ORDEFA")
   Call ShellExecute(Me.hwnd, "Open", "Ordenes de Produccion.pps", "", "", 1)
End Sub

Private Sub Command14_Click()
   'Call SECUDEMO("CSTOCK")
   Call ShellExecute(Me.hwnd, "Open", "Deposito Caotico.pps", "", "", 1)
End Sub

Private Sub Command15_Click()
'   Call SECUDEMO("ANACOS")
'   MsgBox "Opcion no Disponible"
    Call ShellExecute(Me.hwnd, "Open", "Presupuestos a Clientes.pps", "", "", 1)
End Sub

Private Sub Command16_Click()
   'Call SECUDEMO("FORPRO")
   Call ShellExecute(Me.hwnd, "Open", "Maestro de Items de Stock.pps", "", "", 1)
End Sub

Private Sub Command17_Click()
   'Call SECUDEMO("HORUTA")
   Call ShellExecute(Me.hwnd, "Open", "PharmaFlex.pps", "", "", 1)
End Sub

Private Sub Command18_Click()
'   Call SECUDEMO("AGRAFI")
'    MsgBox "Opcion no Disponible"
Call ShellExecute(Me.hwnd, "Open", "Planera Electronica.pps", "", "", 1)
End Sub

Private Sub Command19_Click()
'   Call SECUDEMO("ISO-9K")
    MsgBox "Opcion no Disponible"
End Sub

Private Sub Command2_Click()
    MsgBox "Opcion no Disponible"
'   MsgBox "Documento no Encontrado"
   'Call SECUDEMO("CARGEN")
End Sub

Private Sub Command20_Click()
'   Call SECUDEMO("PREFIN")
    MsgBox "Opcion no Disponible"
End Sub

Private Sub Command21_Click()
'   Call SECUDEMO("GEMANTIN")
    MsgBox "Opcion no Disponible"
End Sub

Private Sub Command22_Click()
  'Call SECUDEMO("REPARA")
   Call ShellExecute(Me.hwnd, "Open", "Despacho de Materiales.pps", "", "", 1)
End Sub

Private Sub Command23_Click()
   Call ShellExecute(Me.hwnd, "Open", "Gestion de Stock.pps", "", "", 1)
End Sub

Private Sub Command24_Click()
   Call ShellExecute(Me.hwnd, "Open", "Maestro de Proveedores.pps", "", "", 1)
End Sub

Private Sub Command25_Click()
   Call ShellExecute(Me.hwnd, "Open", "Maestro de clientes.pps", "", "", 1)
End Sub

Private Sub Command26_Click()
'Call SECUDEMO("HORUTA")
 Call ShellExecute(Me.hwnd, "Open", "PharmaFlex.pps", "", "", 1)
End Sub

Private Sub Command3_Click()
   MASINPPT.Show 1
End Sub

Private Sub Command4_Click()
   '
'   Command4.Enabled = False
'   LISECUEN.Show 1
'   Command4.Enabled = True
'   TTXX$ = Trim$(LISECUEN.List1.Text)
'   If TTXX$ <> "" Then
'     SECUDE$ = Trim$(Left$(TTXX$, 10))
'     Call SECUDEMO(SECUDE$)
'   End If
   '
   Call ShellExecute(Me.hwnd, "Open", "Batch Record Electronico.pps", "", "", 1)
   '
End Sub

Private Sub Command5_Click()
'   Call SECUDEMO("GESVAL")
'    MsgBox "Opcion no Disponible"
Call ShellExecute(Me.hwnd, "Open", "Tesoreria y Gestion de Valores.pps", "", "", 1)
End Sub

Private Sub Command6_Click()
'   Call SECUDEMO("CONTAB")
'    MsgBox "Opcion no Disponible"
    Call ShellExecute(Me.hwnd, "Open", "Contabilidad General.pps", "", "", 1)
End Sub

Private Sub Command7_Click()
'   Call SECUDEMO("GESACR")
'    MsgBox "Opcion no Disponible"
    Call ShellExecute(Me.hwnd, "Open", "Acreedores Comerciales.pps", "", "", 1)
End Sub

Private Sub Command8_Click()
'   Call SECUDEMO("GESVEN")
'    MsgBox "Opcion no Disponible"
    Call ShellExecute(Me.hwnd, "Open", "Gestion de Ventas.pps", "", "", 1)
End Sub

Private Sub Command9_Click()
   OPCIDEPPT.Show
   Unload Me
End Sub

Private Sub Form_Load()
    Call CENTRAFORM(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   '
   For K% = 0 To 6
      DEMOSPPT.MODULO(K%).Visible = True
   Next K%
   '
   DEMOSPPT.Command1.Visible = True
   '
End Sub

