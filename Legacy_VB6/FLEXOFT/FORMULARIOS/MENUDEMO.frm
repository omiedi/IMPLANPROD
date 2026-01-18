VERSION 5.00
Begin VB.Form MENUDEMO 
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
   Begin VB.PictureBox Picture22 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   3360
      ScaleHeight     =   585
      ScaleWidth      =   2055
      TabIndex        =   73
      Top             =   7560
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
         Picture         =   "MENUDEMO.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Sistema de Gestión de Reparaciones"
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Gestion de Reparaciones"
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
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture21 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   6405
      ScaleHeight     =   585
      ScaleWidth      =   2160
      TabIndex        =   70
      Top             =   7560
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
         Picture         =   "MENUDEMO.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   71
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
         TabIndex        =   72
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
      TabIndex        =   67
      Top             =   6300
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
         Picture         =   "MENUDEMO.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   68
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
         Left            =   630
         TabIndex        =   69
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
      TabIndex        =   63
      Top             =   6090
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
         Picture         =   "MENUDEMO.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   64
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
         TabIndex        =   65
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
      TabIndex        =   60
      Top             =   5355
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
         Picture         =   "MENUDEMO.frx":1500
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Sistema de Administración de Documentos Gráficos"
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
         TabIndex        =   62
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
      TabIndex        =   58
      Top             =   4620
      Width           =   2220
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
         Left            =   0
         Picture         =   "MENUDEMO.frx":180A
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   "Hojas de Ruta - Secuencia de Fabricación"
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Secuencia de Fabricación"
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
   Begin VB.PictureBox Picture16 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   6405
      ScaleHeight     =   585
      ScaleWidth      =   2160
      TabIndex        =   55
      Top             =   3885
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
         Picture         =   "MENUDEMO.frx":1B14
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Estructuras y Fórmulas de Producto"
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Formulas y Estructuras"
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
         TabIndex        =   57
         Top             =   105
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
      TabIndex        =   51
      Top             =   2835
      Width           =   2115
      Begin VB.Image Image5 
         Height          =   375
         Index           =   1
         Left            =   70
         Picture         =   "MENUDEMO.frx":28BE
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
         TabIndex        =   52
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
      TabIndex        =   49
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
         TabIndex        =   50
         Top             =   70
         Width           =   1485
      End
      Begin VB.Image Image5 
         Height          =   375
         Index           =   0
         Left            =   70
         Picture         =   "MENUDEMO.frx":2BC8
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
      TabIndex        =   47
      Top             =   1550
      Width           =   2115
      Begin VB.Image Image4 
         Height          =   270
         Left            =   1680
         Picture         =   "MENUDEMO.frx":2ED2
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
         TabIndex        =   48
         Top             =   315
         Width           =   1275
      End
      Begin VB.Image Image2 
         Height          =   525
         Index           =   2
         Left            =   -50
         Picture         =   "MENUDEMO.frx":301C
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
      TabIndex        =   45
      Top             =   1550
      Width           =   2115
      Begin VB.Image Image3 
         Height          =   270
         Left            =   1680
         Picture         =   "MENUDEMO.frx":D09A
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
         TabIndex        =   46
         Top             =   315
         Width           =   1275
      End
      Begin VB.Image Image2 
         Height          =   525
         Index           =   1
         Left            =   -50
         Picture         =   "MENUDEMO.frx":D3A4
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
      TabIndex        =   41
      Top             =   6825
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
         Picture         =   "MENUDEMO.frx":17422
         Style           =   1  'Graphical
         TabIndex        =   42
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
         TabIndex        =   43
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
      TabIndex        =   38
      Top             =   6825
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
         Picture         =   "MENUDEMO.frx":1772C
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Sistemas de Compras de Materiales y Servicios Productivos"
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
         Left            =   735
         TabIndex        =   40
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
      Picture         =   "MENUDEMO.frx":18296
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   210
      Width           =   750
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   630
      ScaleHeight     =   585
      ScaleWidth      =   2055
      TabIndex        =   34
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
         Picture         =   "MENUDEMO.frx":183E0
         Style           =   1  'Graphical
         TabIndex        =   35
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
         TabIndex        =   36
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
      TabIndex        =   31
      Top             =   3885
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
         Picture         =   "MENUDEMO.frx":1A0DA
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Sistemas de Control de Stocks y Gestión de Inventarios"
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label15 
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
         TabIndex        =   33
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
      TabIndex        =   28
      Top             =   4620
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
         Picture         =   "MENUDEMO.frx":1AF08
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Sistemas de Gestión de Fabricación"
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
         Left            =   735
         TabIndex        =   30
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
      TabIndex        =   25
      Top             =   5355
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
         Picture         =   "MENUDEMO.frx":1BD1A
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Sistemas de Programación de Materiales por Técnica MRP"
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
         TabIndex        =   27
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
      TabIndex        =   22
      Top             =   6090
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
         Picture         =   "MENUDEMO.frx":1CAC4
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Sistemas de Secuenciación de Carga de Máquinas por Técnicas CRP"
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
         Left            =   735
         TabIndex        =   24
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
      TabIndex        =   19
      Top             =   2940
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
         Picture         =   "MENUDEMO.frx":1D706
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Sistema de Gestión de Ventas y Cuentas Corrientes"
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
         TabIndex        =   21
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
      TabIndex        =   16
      Top             =   3780
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
         Picture         =   "MENUDEMO.frx":1E300
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Sistema de Gestión de Acreedores Comerciales"
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
         TabIndex        =   18
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
      TabIndex        =   13
      Top             =   4620
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
         Picture         =   "MENUDEMO.frx":1F08A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Sistema de Contabiliodad General"
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
         TabIndex        =   15
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
      TabIndex        =   10
      Top             =   5460
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
         Picture         =   "MENUDEMO.frx":1FCFC
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Sistema de Tesorería"
         Top             =   0
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
         TabIndex        =   12
         Top             =   105
         Width           =   1485
      End
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   630
      ScaleHeight     =   585
      ScaleWidth      =   2055
      TabIndex        =   7
      Top             =   6825
      Width           =   2115
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
         Picture         =   "MENUDEMO.frx":209BE
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Acceso a Todas las Secuencias de Demostración Disponibles"
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Indice General de Demos "
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
         TabIndex        =   9
         Top             =   105
         Width           =   1275
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
         Picture         =   "MENUDEMO.frx":20CC8
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
         Left            =   0
         Picture         =   "MENUDEMO.frx":20FD2
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
         Picture         =   "MENUDEMO.frx":2111C
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
         TabIndex        =   44
         Top             =   315
         Width           =   1275
      End
      Begin VB.Image Image2 
         Height          =   525
         Index           =   0
         Left            =   -50
         Picture         =   "MENUDEMO.frx":21426
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
      TabIndex        =   53
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
         TabIndex        =   54
         Top             =   70
         Width           =   1485
      End
      Begin VB.Image Image5 
         Height          =   375
         Index           =   2
         Left            =   70
         Picture         =   "MENUDEMO.frx":2B4A4
         Stretch         =   -1  'True
         ToolTipText     =   "Sistemas de Ingeniería de Producto"
         Top             =   70
         Width           =   420
      End
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
      Y1              =   6615
      Y2              =   6615
   End
   Begin VB.Line Line30 
      X1              =   6090
      X2              =   6510
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line29 
      X1              =   6090
      X2              =   6615
      Y1              =   4935
      Y2              =   4935
   End
   Begin VB.Line Line28 
      X1              =   6090
      X2              =   6405
      Y1              =   5670
      Y2              =   5670
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
      Y1              =   5775
      Y2              =   5775
   End
   Begin VB.Line Line21 
      X1              =   9345
      X2              =   9555
      Y1              =   4935
      Y2              =   4935
   End
   Begin VB.Line Line20 
      X1              =   9345
      X2              =   9555
      Y1              =   4095
      Y2              =   4095
   End
   Begin VB.Line Line19 
      X1              =   9345
      X2              =   9555
      Y1              =   3255
      Y2              =   3255
   End
   Begin VB.Line Line6 
      X1              =   9345
      X2              =   9345
      Y1              =   2100
      Y2              =   6615
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
      Y1              =   5670
      Y2              =   5670
   End
   Begin VB.Line Line16 
      X1              =   3150
      X2              =   3465
      Y1              =   4935
      Y2              =   4935
   End
   Begin VB.Line Line15 
      X1              =   3150
      X2              =   3360
      Y1              =   4200
      Y2              =   4200
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
   Begin VB.Line Line11 
      X1              =   420
      X2              =   630
      Y1              =   7140
      Y2              =   7140
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
      X1              =   420
      X2              =   420
      Y1              =   1995
      Y2              =   7140
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
      Left            =   0
      Picture         =   "MENUDEMO.frx":2B7AE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "MENUDEMO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Call SECUDEMO("INSTRU")
End Sub

Private Sub Command10_Click()
   Call SECUDEMO("COMPRA")
End Sub

Private Sub Command11_Click()
   Call SECUDEMO("CARMAQ")
End Sub

Private Sub Command12_Click()
   Call SECUDEMO("PROMAT")
End Sub

Private Sub Command13_Click()
   Call SECUDEMO("ORDEFA")
End Sub

Private Sub Command14_Click()
   Call SECUDEMO("CSTOCK")
End Sub

Private Sub Command15_Click()
   Call SECUDEMO("ANACOS")
End Sub

Private Sub Command16_Click()
   Call SECUDEMO("FORPRO")
End Sub

Private Sub Command17_Click()
   Call SECUDEMO("HORUTA")
End Sub

Private Sub Command18_Click()
   Call SECUDEMO("AGRAFI")
End Sub

Private Sub Command19_Click()
   Call SECUDEMO("ISO-9K")
End Sub

Private Sub Command2_Click()
   MsgBox "Documento no Encontrado"
   'Call SECUDEMO("CARGEN")
End Sub

Private Sub Command20_Click()
   Call SECUDEMO("PREFIN")
End Sub

Private Sub Command21_Click()
   Call SECUDEMO("GEMANTIN")
End Sub

Private Sub Command22_Click()
  Call SECUDEMO("REPARA")
End Sub

Private Sub Command3_Click()
   MASINFO.Show 1
End Sub

Private Sub Command4_Click()
   '
   Command4.Enabled = False
   LISECUEN.Show 1
   Command4.Enabled = True
   TTXX$ = Trim$(LISECUEN.List1.Text)
   If TTXX$ <> "" Then
     SECUDE$ = Trim$(Left$(TTXX$, 10))
     Call SECUDEMO(SECUDE$)
   End If
   '
End Sub

Private Sub Command5_Click()
   Call SECUDEMO("GESVAL")
End Sub

Private Sub Command6_Click()
   Call SECUDEMO("CONTAB")
End Sub

Private Sub Command7_Click()
   Call SECUDEMO("GESACR")
End Sub

Private Sub Command8_Click()
   Call SECUDEMO("GESVEN")
End Sub

Private Sub Command9_Click()
   OPCIDEMO.Show
   Unload Me
End Sub

Private Sub Form_Load()
    Call CENTRAFORM(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   '
   For K% = 0 To 6
      DEMOFLE1.MODULO(K%).Visible = True
   Next K%
   '
   DEMOFLE1.Command1.Visible = True
   '
End Sub

