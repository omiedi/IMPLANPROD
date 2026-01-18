VERSION 5.00
Begin VB.Form FLESCRIT 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema Preconfigurado FLEXOFT (r) - V.9.03"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10875
   ClipControls    =   0   'False
   FillColor       =   &H00FFFF80&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   10875
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Interval        =   60000
      Left            =   840
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Interval        =   15000
      Left            =   420
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   2430
      Left            =   4200
      ScaleHeight     =   2370
      ScaleWidth      =   2775
      TabIndex        =   38
      Top             =   2500
      Visible         =   0   'False
      Width           =   2840
      Begin VB.ListBox List1 
         Height          =   2010
         Left            =   0
         TabIndex        =   41
         Top             =   315
         Width           =   2745
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00800000&
         Height          =   330
         Left            =   0
         ScaleHeight     =   270
         ScaleWidth      =   3000
         TabIndex        =   39
         Top             =   0
         Width           =   3060
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Actualizaciones FLEXOFT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   225
            Left            =   105
            TabIndex        =   40
            Top             =   25
            Width           =   2640
         End
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFEB&
      Caption         =   "SIG - Sistema de Gestión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   5985
      TabIndex        =   32
      Top             =   3770
      Visible         =   0   'False
      Width           =   4635
      Begin VB.CommandButton Command5 
         Caption         =   "Contable"
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
         Left            =   315
         Picture         =   "FLESCRIT.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Contabilidad General - Caja y Bancos"
         Top             =   420
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Acreedor"
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
         Left            =   2415
         Picture         =   "FLESCRIT.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Cuentas Corrientes Acreedoras"
         Top             =   420
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Ventas"
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
         Left            =   1365
         Picture         =   "FLESCRIT.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Gestión de Ventas y Deudores"
         Top             =   420
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Personal"
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
         Left            =   3465
         Picture         =   "FLESCRIT.frx":0B8E
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Sueldos - Jornales - Administración de Personal"
         Top             =   420
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "SIM - Sistema de Manufactura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   5985
      TabIndex        =   22
      Top             =   1155
      Visible         =   0   'False
      Width           =   4635
      Begin VB.CommandButton Command6 
         Caption         =   "I-Temp"
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
         Left            =   3990
         Picture         =   "FLESCRIT.frx":0E98
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Sistema de Importacion Temporaria"
         Top             =   1260
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Compras"
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
         Left            =   2415
         Picture         =   "FLESCRIT.frx":0FE2
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Gestión de Pedidos y Ordenes de Compra a Terceros."
         Top             =   1470
         Width           =   855
      End
      Begin VB.CommandButton Command16 
         Caption         =   "News"
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
         Left            =   1365
         Picture         =   "FLESCRIT.frx":12EC
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Informes de Producción"
         Top             =   1470
         Width           =   855
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Costos"
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
         Left            =   3465
         Picture         =   "FLESCRIT.frx":172E
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Costos Industriales - Presupuestación de Productos"
         Top             =   1470
         Width           =   855
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Factory"
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
         Left            =   315
         Picture         =   "FLESCRIT.frx":1A38
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Ordenes de Trabajo y Fabricación"
         Top             =   1470
         Width           =   855
      End
      Begin VB.CommandButton Command14 
         Caption         =   "C.R.P"
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
         Left            =   3465
         Picture         =   "FLESCRIT.frx":1D42
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Programación de Carga de Máquinas y Equipos - Técnica CRP"
         Top             =   420
         Width           =   855
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Stock"
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
         Left            =   1365
         Picture         =   "FLESCRIT.frx":204C
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Control y Movimientos de Stocks"
         Top             =   420
         Width           =   855
      End
      Begin VB.CommandButton Command13 
         Caption         =   "MRP-I"
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
         Left            =   2415
         Picture         =   "FLESCRIT.frx":2356
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Planificación de Necesidades de Materiales - Técnica MRP-II"
         Top             =   420
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Files"
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
         Left            =   315
         Picture         =   "FLESCRIT.frx":2660
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Actualización y Consulta de Archivos Maestros"
         Top             =   420
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Utilities"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   5985
      TabIndex        =   18
      Top             =   5330
      Visible         =   0   'False
      Width           =   4635
      Begin VB.CommandButton Command24 
         Caption         =   "Mail"
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
         Left            =   3465
         Picture         =   "FLESCRIT.frx":296A
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Correo Interno FLEXOFT"
         Top             =   420
         Width           =   855
      End
      Begin VB.CommandButton Command22 
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
         Height          =   855
         Left            =   1365
         Picture         =   "FLESCRIT.frx":2C74
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Configuración de Funcionamiento"
         Top             =   420
         Width           =   855
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Qwery"
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
         Left            =   2415
         Picture         =   "FLESCRIT.frx":30B6
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Consultas y Listados"
         Top             =   420
         Width           =   855
      End
      Begin VB.CommandButton Command21 
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
         Height          =   855
         Left            =   315
         Picture         =   "FLESCRIT.frx":3200
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Ayuda en Línea"
         Top             =   420
         Width           =   855
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "End"
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
      Left            =   9765
      Picture         =   "FLESCRIT.frx":350A
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Terminar Sesión de Trabajo - Cerrar Sistema"
      Top             =   210
      Width           =   855
   End
   Begin VB.FileListBox File1 
      Height          =   1650
      Hidden          =   -1  'True
      Left            =   210
      TabIndex        =   15
      Top             =   5145
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Acceso al Sistema FLEXOFT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2640
      Left            =   6615
      TabIndex        =   0
      Top             =   4095
      Width           =   3585
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2205
         TabIndex        =   3
         Top             =   630
         Width           =   1170
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   210
         TabIndex        =   2
         Top             =   1365
         Width           =   3165
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2205
         TabIndex        =   1
         Top             =   1995
         Width           =   1170
      End
      Begin VB.Image Image5 
         Height          =   540
         Left            =   210
         Picture         =   "FLESCRIT.frx":3654
         Stretch         =   -1  'True
         Top             =   1890
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Image Image4 
         Height          =   525
         Left            =   210
         Picture         =   "FLESCRIT.frx":395E
         Stretch         =   -1  'True
         Top             =   1890
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Image Image2 
         Height          =   525
         Left            =   210
         Picture         =   "FLESCRIT.frx":3C68
         Stretch         =   -1  'True
         Top             =   1890
         Width           =   585
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Acceso"
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
         Left            =   1485
         TabIndex        =   16
         Top             =   2205
         Width           =   645
      End
      Begin VB.Image Image1 
         Height          =   585
         Left            =   315
         Picture         =   "FLESCRIT.frx":3F72
         Stretch         =   -1  'True
         Top             =   420
         Width           =   690
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario"
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
         Left            =   1455
         TabIndex        =   5
         Top             =   735
         Width           =   660
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clave de"
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
         Left            =   1350
         TabIndex        =   4
         Top             =   1995
         Width           =   765
      End
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   10920
      Y1              =   6870
      Y2              =   6870
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "         Soluciones Informaticas Integrales "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   7140
      TabIndex        =   14
      Top             =   7035
      Width           =   3690
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2835
      TabIndex        =   13
      Top             =   4515
      Width           =   135
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2835
      TabIndex        =   12
      Top             =   3465
      Width           =   135
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2835
      TabIndex        =   11
      Top             =   2415
      Width           =   135
   End
   Begin VB.Image Image3 
      Height          =   885
      Left            =   0
      Picture         =   "FLESCRIT.frx":427C
      Stretch         =   -1  'True
      Top             =   6890
      Width           =   10920
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ernesto P. Bauer "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   210
      TabIndex        =   10
      Top             =   6090
      Width           =   5475
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Buenos Aires, marzo 4 de 1999.-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   210
      TabIndex        =   9
      Top             =   5145
      Width           =   5580
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "por el sistema denominado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   210
      TabIndex        =   8
      Top             =   2940
      Width           =   5580
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "bajo la supervision y responsabilidad de"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   210
      TabIndex        =   7
      Top             =   3990
      Width           =   5580
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "En mi caracter de autor del  Sistema  Informatico Preconfigurado Flexoft (r)  otorgo licencia de uso a favor de"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   210
      TabIndex        =   6
      Top             =   1155
      Width           =   5580
   End
End
Attribute VB_Name = "FLESCRIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Image6_Click()
    Call FUERASYS
End Sub

Private Sub Image3_DblClick()
   Call UPDATES
End Sub

Private Sub Label11_Click()
    Call FUERASYS
End Sub

Private Sub command1_Click()
    TERMINAR.Show 1
End Sub

Private Sub List1_DblClick()
    '
    MENUPRIN$ = Mid$(List1.TEXT, 57, 8)
    Call GRARAM(13, AJUSTI$(MENUPRIN$, 8))
    FLEMENUS.Show
    '
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
    '
    A% = KeyAscii
    If A% = 27 Then
          Call FUERASYS
       ElseIf A% = 13 Then
          MENUPRIN$ = Mid$(List1.TEXT, 57, 8)
          If TRIM$(MENUPRIN$) <> "" Then
              Call GRARAM(13, AJUSTI$(MENUPRIN$, 8))
              FLEMENUS.Show
          End If
    End If
    '
End Sub


Private Sub Command11_Click()
    Command11.Enabled = False
    ACONEC$ = "ARCHIG04"
    If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "ARCHIMA"
    If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "ARCH_GES"
    Call CONECRUN(ACONEC$, "")
    Command11.Enabled = True
End Sub

Private Sub Command12_Click()
    Command12.Enabled = False
    If EXISTE%("CONSTO04.EXE") > 0 Then
          Call CONECRUN("CONSTO04", "Control y Movimientos de Stock")
        ElseIf EXISTE%("CONSTO02.EXE") > 0 Then
          Call CONECRUN("CONSTO02", "Control y Movimientos de Stock")
        ElseIf EXISTE%("CONSTO01.EXE") > 0 Then
          Call CONECRUN("CONSTO01", "Control y Movimientos de Stock")
        ElseIf EXISTE%("CONSTOCK.EXE") > 0 Then
          Call CONECRUN("CONSTOCK", "Control y Movimientos de Stock")
        ElseIf EXISTE%("STOCK.MNU") > 0 Then
          MENUPRIN$ = "STOCK"
          Call GRARAM(13, MENUPRIN$ + "   ")
          FLEMENUS.Show
        ElseIf EXISTE%("EXPEDI.MNU") > 0 Then
          MENUPRIN$ = "EXPEDI"
          Call GRARAM(13, MENUPRIN$ + "   ")
          FLEMENUS.Show
    End If
    Command12.Enabled = True
End Sub

Private Sub Command13_Click()
    Command13.Enabled = False
    If EXISTE%("PROMRP04.EXE") > 0 Then
         Call CONECRUN("PROMRP04", "Programacion de Materiales")
       ElseIf EXISTE%("PROMRP02.EXE") > 0 Then
         Call CONECRUN("PROMRP02", "Programacion de Materiales")
       ElseIf EXISTE%("PROMRP01.EXE") > 0 Then
         Call CONECRUN("PROMRP01", "Programacion de Materiales")
       Else
         Call CONECRUN("PROGMAT", "Programacion de Materiales")
    End If
    Command13.Enabled = True
End Sub

Private Sub Command14_Click()
    Command14.Enabled = False
    If EXISTE%("CARMAQ04.EXE") > 0 Then
        Call CONECRUN("CARMAQ04", "Cálculo de Carga de Máquinas")
      ElseIf EXISTE%("CARMAQ03.EXE") > 0 Then
        Call CONECRUN("CARMAQ03", "Cálculo de Carga de Máquinas")
      ElseIf EXISTE%("CARMAQ02.EXE") > 0 Then
        Call CONECRUN("CARMAQ02", "Cálculo de Carga de Máquinas")
      Else
        Call CONECRUN("CARGMAQ", "Cálculo de Carga de Máquinas")
    End If
    Command14.Enabled = True
End Sub

Private Sub Command15_Click()
    Command15.Enabled = False
    If EXISTE%("ORFABR04.EXE") > 0 Then
        Call CONECRUN("ORFABR04", "Ordenes de Fabricación")
      ElseIf EXISTE%("ORFABR02.EXE") > 0 Then
        Call CONECRUN("ORFABR02", "Ordenes de Fabricación")
      Else
        Call CONECRUN("ORDEFABR", "Ordenes de Fabricación")
    End If
    Command15.Enabled = True
End Sub

Private Sub Command16_Click()
    Command16.Enabled = False
    If EXISTE%("IFABRI04.EXE") > 0 Then
         Call CONECRUN("IFABRI04", "Informes de Producción")
       ElseIf EXISTE%("IFABRI02.EXE") > 0 Then
         Call CONECRUN("IFABRI02", "Informes de Producción")
       ElseIf EXISTE%("INFOFAB.EXE") > 0 Then
         Call CONECRUN("INFOFAB", "Informes de Producción")
    End If
    Command16.Enabled = True
End Sub

Private Sub Command17_Click()
    If EXISTE%("ANACOS04.EXE") > 0 Then
          Call CONECRUN("ANACOS04", "Analisis y Presupuestacion de Producto")
        ElseIf EXISTE%("ANACOS02.EXE") > 0 Then
          Call CONECRUN("ANACOS02", "Analisis y Presupuestacion de Producto")
        ElseIf EXISTE%("ANACOS01.EXE") > 0 Then
          Call CONECRUN("ANACOS01", "Analisis y Presupuestacion de Producto")
        ElseIf EXISTE%("COSTOS.EXE") > 0 Then
          Call CONECRUN("COSTOS", "Analisis y Presupuestacion de Producto")
        ElseIf EXISTE%("COSTOS.MNU") > 0 Then
          MENUPRIN$ = "COSTOS"
          Call GRARAM(13, MENUPRIN$ + "   ")
          FLEMENUS.Show
    End If
End Sub

Private Sub Command18_Click()
    Command18.Enabled = False
    If InStr(EMPREAC$, "SABO") > 0 Then
        ACONEC$ = "AUTENT03"
      Else
        ACONEC$ = "ORCOMP04"
    End If
    If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "ORCOMP03"
    If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "ORCOMP02"
    If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "OCOMPRA"
    Call CONECRUN(ACONEC$, "Ordenes de Compra Terceros")
    Command18.Enabled = True
End Sub

Private Sub Command2_Click()
    If EXISTE%("ADMIFLEX.EXE") > 0 Then
      Command2.Enabled = False
      Call CONECRUN("ADMIFLEX", "")
      Command2.Enabled = True
      Exit Sub
    End If
    MENUPRIN$ = "SUELDOS"
    Call GRARAM(13, MENUPRIN$ + " ")
    FLEMENUS.Show
End Sub

Private Sub Command22_Click()
    Form4.Show
End Sub

Private Sub Command23_Click()
    Call GRAJOBID
    Call BAJRAM
    Shell LUBAS$ + "WINTABU " + "#" + JOBID$ + "#", 1
End Sub

Private Sub Command24_Click()
    Command24.Enabled = False
    Call CONECRUN("FLEXMAIL", "")
    Command24.Enabled = True
End Sub

Private Sub Command3_Click()
    '
    Command3.Enabled = False
    If EXISTE%("VENGES04.EXE") > 0 Then
        Call CONECRUN("VENGES04", "Gestión de Ventas y Deudores")
        GoTo 99
      ElseIf EXISTE%("VENTAS.MNU") > 0 Then
       MENUPRIN$ = "VENTAS"
       Call GRARAM(13, MENUPRIN$ + "  ")
       FLEMENUS.Show
       GoTo 99
    End If
    '
    If EXISTE%("PEDIDO02.EXE") > 0 Then
        Call CONECRUN("PEDIDO02", "Sub-Sistema de Pedidos de Clientes")
      ElseIf EXISTE%("PEDIDO01.EXE") > 0 Then
        Call CONECRUN("PEDIDO01", "Sub-Sistema de Pedidos de Clientes")
    End If
    '
99  Command3.Enabled = True
    '
End Sub

Private Sub Command4_Click()
    If EXISTE%("ACRGES04.EXE") > 0 Then
        Call CONECRUN("ACRGES04", "Gestión de Compras y Acreedores")
      Else
        MENUPRIN$ = "COMPRAS"
        Call GRARAM(13, MENUPRIN$ + " ")
        FLEMENUS.Show
    End If
End Sub

Private Sub Command5_Click()
    Call SELEJER(EJUI%)
    If Abs(EJUI%) <> 1 Then
       Exit Sub
    End If
    '
    If EXISTE%("CONTAB04.EXE") > 0 Then
          Call CONECRUN("CONTAB04", "Sistema de Contabilidad General")
       ElseIf EXISTE%("CONTAGEN.EXE") > 0 Then
          Call CONECRUN("CONTAGEN", "Sistema de Contabilidad General")
       Else
          MENUPRIN$ = "CONTAGEN"
          Call GRARAM(13, AJUSTI$(MENUPRIN$, 8))
          FLEMENUS.Show
    End If
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   If EXISTE%("SIMPTE03.EXE") > 0 Then
     Call CONECRUN("SIMPTE03", "Sistema de Importacion Temporaria")
   End If
   Command6.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 1
   TERMINAR.Show 1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    '
          Static CLA$, CLB$
          '
          If Text1.TEXT = "" Then
              CLA$ = "": CLB$ = ""
          End If
          '
          A% = KeyAscii: KeyAscii = 0
          Text1.TEXT = CLB$
          If A% = 13 Then
              Call BUSER
              If UCase$(Left$(CONTROL$("SYSTEM", "STATUS"), 8)) = "MANTENIM" Then
                MsgBox "En Mantenimiento - Intente más Tarde"
                If NUSER% > 1 Then
                  Call FUERASYS
                  Exit Sub
                End If
              End If
AVC% = 0
If CONTROL$("FLEXOFT", "LMTBNMNT") = "SI" Then
   AVC% = 1
End If
Call GRARAM(307, TRIM$(Str$(AVC%)))
          '
             Exit Sub
          End If
          If A% = 27 Then Call FUERASYS
          If A% < 48 Then Exit Sub
         
          If A% > 96 Then A% = A% - 32
          If Len(Text1.TEXT) < 12 Then
             CLB$ = CLB$ + "*"
             Text1.TEXT = CLB$
             CLA$ = CLA$ + Chr$(A%)
             CLAVE$ = AJUSTI$(CLA$, 12)
          End If
          '
End Sub

Private Sub Form_Load()
    '
'   If HOY(HO$) > 8350 Then
'       Call FUERASYS
'   End If
    '
    Call TRANSLABELS(Name)
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    Label6.Caption = "Buenos Aires, " + TRIM$(LCase$(FETEXTO$(FELICEN%))) + ".-"
    Label8.Caption = EMPRELI$
    Label9.Caption = DENOSYS$
    Label10.Caption = RESPOSI$
    '
    Text1.TEXT = ""
    Text1.Enabled = True
    '
    If EXISTE%("ARCHIG04.EXE") < 1 Then
      If EXISTE%("ARCHIMA.EXE") < 1 Then
        If EXISTE%("ARCH_GES.EXE") < 1 Then
          Command11.Enabled = False
        End If
      End If
    End If
    '
    If EXISTE%("CONSTO04.EXE") < 1 Then
      If EXISTE%("CONSTO02.EXE") < 1 Then
        If EXISTE%("CONSTO01.EXE") < 1 Then
          If EXISTE%("CONSTOCK.EXE") < 1 Then
            If EXISTE%("STOCK.MNU") < 1 Then
              Command12.Enabled = False
            End If
          End If
        End If
      End If
    End If
    '
    If Command12.Enabled = False Then
       If EXISTE%("EXPEDI.MNU") > 0 Then
          Command12.Caption = "Remitos"
          Command12.ToolTipText = "Sistema de Expedicion - Remitos de Facturacion."
          Command12.Enabled = True
       End If
    End If
    '
    If EXISTE%("PROMRP04.EXE") < 1 Then
      If EXISTE%("PROMRP02.EXE") < 1 Then
        If EXISTE%("PROMRP01.EXE") < 1 Then
          If EXISTE%("PROGMAT.EXE") < 1 Then
            Command13.Enabled = False
          End If
        End If
      End If
    End If
    '
    If EXISTE%("CARMAQ04.EXE") < 1 Then
      If EXISTE%("CARMAQ03.EXE") < 1 Then
        If EXISTE%("CARMAQ02.EXE") < 1 Then
          If EXISTE%("CARGMAQ.EXE") < 1 Then
            Command14.Enabled = False
          End If
        End If
      End If
    End If
    '
    If EXISTE%("ORDEFABR.EXE") < 1 Then
      If EXISTE%("ORFABR02.EXE") < 1 Then
        If EXISTE%("ORFABR04.EXE") < 1 Then
          Command15.Enabled = False
        End If
      End If
    End If
    '
    If EXISTE%("IFABRI04.EXE") < 1 Then
      If EXISTE%("IFABRI02.EXE") < 1 Then
        If EXISTE%("INFOFAB.EXE") < 1 Then
          Command16.Enabled = False
        End If
      End If
    End If
    '
    If EXISTE%("OCOMPRA.EXE") < 1 Then
      If EXISTE%("ORCOMP02.EXE") < 1 Then
        If EXISTE%("ORCOMP03.EXE") < 1 Then
          If EXISTE%("ORCOMP04.EXE") < 1 Then
            If EXISTE%("AUTENT03.EXE") < 1 Then
              Command18.Enabled = False
            End If
          End If
        End If
      End If
    End If
    '
    If EXISTE%("ANACOS04.EXE") < 1 Then
      If EXISTE%("ANACOS02.EXE") < 1 Then
        If EXISTE%("ANACOS01.EXE") < 1 Then
          If EXISTE%("COSTOS.EXE") < 1 Then
            If EXISTE%("COSTOS.MNU") < 1 Then
              Command17.Enabled = False
            End If
          End If
        End If
      End If
    End If
    '
    If InStr(EMPREAC$, "SABO") > 0 Then
      If EXISTE%("SIMPTE03.EXE") > 0 Then
        Command6.Top = Command17.Top
        Command6.Left = Command17.Left
        Command6.Visible = True
        Command6.Enabled = True
      End If
    End If
    '
    If EXISTE%("CONTAB04.EXE") < 1 Then
      If EXISTE%("CONTAGEN.EXE") < 1 Then
        If EXISTE%("CONTAGEN.MNU") < 1 Then
          Command5.Enabled = False
        End If
      End If
    End If
    '
    If EXISTE%("VENTAS.MNU") < 1 Then
        Command3.Enabled = False
    End If
    '
    If Command3.Enabled = False Then
       If EXISTE%("PEDIDO01.EXE") Or EXISTE%("PEDIDO02.EXE") > 0 Then
         Command3.Enabled = True
         Command3.Caption = "Pedidos"
         Command3.ToolTipText = "Sub-Sistema de Pedidos de Clientes"
       End If
    End If
    '
    If EXISTE%("COMPRAS.MNU") < 1 Then
      If EXISTE%("ACRGES04.EXE") < 1 Then
        Command4.Enabled = False
      End If
    End If
    '
    If EXISTE%("SUELDOS.MNU") < 1 Then
        Command2.Enabled = False
    End If
    '
    If EXISTE%("ADMIFLEX.EXE") > 0 Then
        Command2.Caption = "Admin."
        Command2.Enabled = True
    End If
End Sub


Private Sub Text3_KeyPress(KeyAscii As Integer)

          Static CLA$, CLB$
          '
          If Text3.TEXT = "" Then CLA$ = "": CLB$ = ""
          A% = KeyAscii
          KeyAscii = 0
          Text3.TEXT = CLB$
          If A% = 13 Then
'              If UCase$(Left$(CONTROL$("SYSTEM", "STATUS"), 8)) = "MANTENIM" Then
'                MsgBox "En Mantenimiento - Intente más Tarde"
'                Call FUERASYS
'                Exit Sub
'              End If
              '
              Call BUCLA
'              If USNBR% > 0 Then
'                 FLESCRIT.Image5.Visible = False
'                 FLESCRIT.Image2.Visible = False
'                 FLESCRIT.Image4.Visible = True
'                 tini = Timer: While Timer < tini + 1: Wend
'                 '
'                 If TRIM$(MENUPRIN$) <> "" Then
'                     FLEMENUS.Show
'                   Else
'                     Call MENUSEL
'                     FLESCRIT.Image4.Visible = False
'                     FLESCRIT.Image6.Visible = True
'                     FLESCRIT.Label11.Visible = True
'                     FLESCRIT.Frame1.Visible = False
'                     If FLESCRIT.List1.ListCount > 1 Then
'                            FLESCRIT.List1.Visible = True
'                            FLESCRIT.List1.SetFocus
'                        ElseIf FLESCRIT.List1.ListCount = 1 Then
'                           FLESCRIT.List1.ListIndex = 0
'                           Menu$ = TRIM$(Mid$(FLESCRIT.List1.Text, 57, 8))
'                           Call GRARAM(13, Menu$)
'                           FLEMENUS.Show
'                        Else
'                           Call MENSERR(24, "No se Encontraron\Menúes Activos.")
'                           Call FUERASYS
'                    End If
'                 End If
'              End If
          End If
          '
          If A% = 27 Then Call FUERASYS
          If A% < 48 Then Exit Sub
          If A% > 96 Then A% = A% - 32
          If Len(Text3.TEXT) < 12 Then
             CLB$ = CLB$ + "*"
             Text3.TEXT = CLB$
             CLA$ = CLA$ + Chr$(A%)
             CONTRA$ = AJUSTI$(CLA$, 12)
          End If
          '
End Sub

Private Sub Timer1_Timer()
    '
    Static CTX%
    '
    If CTX% = 0 Then
        '
        If USNBR% > 0 Then
            '
            CTX% = 1
            FLESCRIT.Image5.Visible = False
            FLESCRIT.Image2.Visible = False
            FLESCRIT.Image4.Visible = True
            '
        End If
        '
      ElseIf CTX% = 1 Then
        '
        CTX% = 2
        FLESCRIT.Image4.Visible = False
        FLESCRIT.Frame1.Visible = False
        FLESCRIT.Frame2.Visible = True
        FLESCRIT.Frame3.Visible = True
        FLESCRIT.Frame4.Visible = True
        '
VENCIABO% = 32767
'EPLI$ = TRIM$(LCase$(EMPRELI$))
'EPAC$ = TRIM$(LCase$(EMPREAC$))
'If InStr(EPLI$, "dosiva") > 0 Or InStr(EPAC$, "dosiva") > 0 Then
'   XYZT$ = REGLEIDO$("MASTER", 1)
'   VENCIABO% = CVI(Mid$(XYZT$, 122, 2))
'   If VENCIABO% > 0 Then
'     If VENCIABO% < HOY(HOS$) Then
'       Call GRACONTROL("FLEXOFT", "LMTBNMNT", "SI")
'     End If
'   End If
'   Call CIERRARCH("MASTER")
'End If
'
AVC% = 0
'If CONTROL$("FLEXOFT", "LMTBNMNT") = "SI" Then
'  AVC% = 1
'End If
Call GRARAM(307, TRIM$(Str$(AVC%)))
'
KKII% = 0
SEVOL$ = SERIVOL$(Left$(CurDir$, 2))
While Len(SEVOL$) < 13
  KKII% = KKII% + 1
  SEVOL$ = SEVOL$ + Chr$(18 + Asc(Mid$(SEVOL$, KKII%, 1)))
Wend
TENCRI$ = ENCRIPTA$(AJUSTI$(SEVOL$, 13))
Call GRARAM(308, TENCRI$)
'
        If VERDEMO% < 1 Then
          Call UPDATES
        End If
        '
        If CONTROL$(IDENTER$, "CONEGATI") = "SI" Then
            Call NEGASTO2
            If ULTREG&("NEGASTO") > 0 Then
                Call CONECRUN("*NEGASTS", "")
            End If
            Call CIERRARCH("*.*")
        End If
        '
        If CONTROL$(IDENTER$, "REVIPEPE") = "SI" Then
            Call CONECRUN("ADMIPED/REVIPEPE", "Revision Pedidos a Despachar")
        End If
        '
        If TRIM$(MENUPRIN$) <> "" Then
             If UCase$(MENUPRIN$) = "CONTAGEN" Then
                 Call SELEJER(EJUI%)
                 If Abs(EJUI%) <> 1 Then
                    Exit Sub
                 End If
             End If
             '
             If TRIM$(MENUPRIN$) = "CONTAGEN" Then
                If EXISTE%("CONTAGEN.EXE") > 0 Then
                   MENUPRIN$ = ""
                   Call GRARAM(13, AJUSTI$(MENUPRIN$, 8))
                   Call CONECRUN("CONTAGEN", "Sistema de Contabilidad General")
                   Exit Sub
                End If
             End If
             FLEMENUS.Show
             '
        End If
    End If
    '
    EPAC$ = EMPREAC$
    If InStr(EPAC$, "NEUMANN") > 0 Then
'       If Timer1.Interval < 3000 Then
'          Timer1.Interval = 3000
'       End If
'       '
'       HOII% = HOY(HOS$)
'       FECHORGRA$ = CONTROL$("", "FECHORA")
'       FEGRA% = FECHANUM(Left$(FECHORGRA$, 8))
'       TIGRA$ = Mid$(FECHORGRA$, 9)
'       FEX = HOII%
'       If FEGRA% > HOII% Then
'          FEX = FEGRA%
'          HOS$ = FECHATEX$(FEX)
'          Date$ = Mid$(HOS$, 4, 2) + "-" + Left$(HOS$, 2) + "-20" + Mid$(HOS$, 7, 2)
'          Time = TIGRA$
'       End If
'       Call GRACONTROL("", "FECHORA", HOS$ + Time$)
    End If
    '
End Sub

Sub UPDATES()
   '
   Screen.MousePointer = 11
   DIREPRO$ = App.Path
   DIREVIE$ = DIREPRO$ + "\VIEJOS\"
   If EXISTE%("F:\FLEXOFT\NOVEDAD\.") > 0 Then
       DIRENEW$ = "F:\FLEXOFT\NOVEDAD"
     Else
       DIRENEW$ = Left$(DIREPRO$, 2) + "\FLEXOFT\NOVEDAD"
   End If
   '
   YAMENSA% = 0
   INSTALAC% = 0
   List1.Clear
   '
   On Error Resume Next
   MkDir DIREVIE$
   On Error GoTo 0
   '
   vuelta% = 0
10 On Error Resume Next
   File1.Path = DIRENEW$
   If Err = 76 Then
     If vuelta% = 0 Then
       On Error GoTo 0
       vuelta% = 1
       DIRENEW$ = "F:" + Mid$(DIRENEW$, 3)
       GoTo 10
     End If
   End If
   '
   If Err < 1 Then
     On Error GoTo 0
     File1.Pattern = "*.*"
     '
     For KU& = 1 To File1.ListCount
       ANAME$ = TRIM$(UCase$(File1.List(KU& - 1)))
       PPXX% = InStr(ANAME$, ".")
       If PPXX% < 1 Then PPXX% = 1 + Len(ANAME$)
       EXTEN$ = TRIM$(UCase$(Mid$(ANAME$, PPXX% + 1)))
       '
       If Len(EXTEN$) = 3 Then
         '
         If EXTEN$ = "RFS" Or EXTEN$ = "DAT" Then
           GoTo 29    ' NO INSTALA DATOS
         End If
         '
         RUTACOM$ = DIREPRO$ + "\"
         '
         ' lo siguiente para forzar instalacion del FLX_MSGR
         If ANAME$ = "FLX_MSGR.EXE" And EXISTE%(RUTACOM$ + ANAME$) < 1 Then GoTo 12
         '
         If EXISTE%(RUTACOM$ + ANAME$) > 0 Then
           FECHANUE# = FileDateTime(DIRENEW$ + "\" + ANAME$)
           FECHAVIE# = FileDateTime(RUTACOM$ + ANAME$)
           If FECHANUE# > FECHAVIE# + 0.0001 Then  ' ESO SON APROX 6 SEGUNDOS
12           If YAMENSA% = 0 Then
               MESSA$ = "Hay Actualizaciones Listas para Instalarse." + Chr$(10) + Chr$(13) + Chr$(10) + Chr$(13)
               MESSA$ = MESSA$ + "Para Instalar las Actualizaciones Disponibles" + Chr$(10) + Chr$(13)
               MESSA$ = MESSA$ + "Cierre las Aplicaciones FLEXOFT en TODAS las" + Chr$(10) + Chr$(13)
               MESSA$ = MESSA$ + "Estaciones de Trabajo." + Chr$(10) + Chr$(13) + Chr$(10) + Chr$(13)
               MESSA$ = MESSA$ + "¿ Desea Ahora Actualizar su Sistema FLEXOFT ?"
               RESPU% = MsgBox(MESSA$, vbYesNoCancel)
               If RESPU% <> 6 Then
                 Screen.MousePointer = 1
                 Exit Sub     ' 6 es = SI
               End If
               YAMENSA% = 1
               Picture1.Visible = True
             End If
             '
             List1.AddItem ANAME$
             List1.Refresh
             EXTECO$ = Left$(EXTE$, 2)
             ANAMCO$ = Left$(ANAME$, Len(ANAME$) - 1)
             On Error Resume Next
             SetAttr RUTACOM$ + ANAMCO$ + "0", vbNormal
             Kill RUTACOM$ + ANAMCO$ + "0"
             SetAttr DIREVIE$ + ANAMCO$ + "9", vbNormal
             Kill DIREVIE$ + ANAMCO$ + "9"
             On Error GoTo 0
             If EXISTE%(DIREVIE$ + ANAMCO$ + "9") > 0 Then
               MsgBox "Imposible Instalar " + ANAME$ + " (1)"
               On Error GoTo 0
               GoTo 29
             End If
             '
             On Error Resume Next
             SetAttr DIREVIE$ + ANAME$, vbNormal
             On Error GoTo 0
             On Error Resume Next
             FileCopy DIRENEW$ + "\" + ANAME$, DIREVIE$ + ANAME$
             If Err Then
               MsgBox "Imposible Instalar " + ANAME$ + ". Aparentemente en Uso."
               On Error GoTo 0
               GoTo 29
             End If
             Kill DIREVIE$ + ANAME$
             On Error GoTo 0
             '
             On Error Resume Next
             For KI% = 8 To 0 Step -1
               SetAttr DIREVIE$ + ANAMCO$ + TRIM$(Str$(KI% + 1)), vbNormal
               Name DIREVIE$ + ANAMCO$ + TRIM$(Str$(KI%)) As DIREVIE$ + ANAMCO$ + TRIM$(Str$(KI% + 1))
             Next KI%
             On Error GoTo 0
             If EXISTE%(DIREVIE$ + ANAMCO$ + "0") > 0 Then
               MsgBox "Imposible Instalar " + ANAME$ + " (2)"
               On Error GoTo 0
               GoTo 29
             End If
             '
             If EXISTE%(RUTACOM$ + ANAME$) < 1 Then GoTo 18
             '
             On Error Resume Next
             Name RUTACOM$ + ANAME$ As RUTACOM$ + ANAMCO$ + "0"
             If Err Then
               MsgBox "Imposible Instalar " + ANAME$ + " (3)"
               On Error GoTo 0
               GoTo 29
             End If
             On Error Resume Next
             Name RUTACOM$ + ANAMCO$ + "0" As DIREVIE$ + ANAMCO$ + "0"
             If Err Then
               MsgBox "Imposible Instalar " + ANAME$ + " (4)"
               On Error GoTo 0
               GoTo 29
             End If
             '
18           On Error Resume Next
             FileCopy DIRENEW$ + "\" + ANAME$, RUTACOM$ + ANAME$
             If Err Then
               MsgBox "Imposible Instalar " + ANAME$ + " (5)"
               On Error GoTo 0
               GoTo 29
             End If
             INSTALAC% = 1
             '
           End If
         End If
       End If
29   Next KU&
     '
     If INSTALAC% > 0 Then MsgBox "Actualizaciones Instaladas"
     Picture1.Visible = False
     '
   End If
   Screen.MousePointer = 1
   '
End Sub

Private Sub Timer2_Timer()
   If UCase$(Left$(CONTROL$("SYSTEM", "STATUS"), 8)) = "MANTENIM" Then
     If AQUIMANTEN% <> 1 Then  ' para que no lo presente donde lo lancé
       If EXISTE%("ENMANTEN.EXE") > 0 Then
         ABCD = Shell("ENMANTEN.EXE", 4)
       End If
     End If
   End If
End Sub

Private Sub Timer3_Timer()
   'If EXISTE%("FLX_MSGR.EXE") > 0 Then
   '   AABC = Shell("FLX_MSGR")
   'End If
End Sub
