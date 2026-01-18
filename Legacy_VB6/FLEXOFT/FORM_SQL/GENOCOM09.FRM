VERSION 5.00
Begin VB.Form GENOCOM09 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0F8&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Escritorio de Compras"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4890
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "GENOCOM09.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   4890
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0F8&
      Caption         =   "Programacion de Materiales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   3690
      Begin VB.PictureBox Picture9 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   600
         ScaleHeight     =   585
         ScaleWidth      =   2385
         TabIndex        =   34
         Top             =   360
         Width           =   2445
         Begin VB.CommandButton Command34 
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
            Picture         =   "GENOCOM09.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Asigna y Reserva Materiales en Stock y Genera Necesidades de Compras"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Re-Calculo General de Necesidades"
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
            Left            =   630
            TabIndex        =   36
            Top             =   105
            Width           =   1860
         End
      End
      Begin VB.Frame Frame5 
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
         TabIndex        =   27
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command16 
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
            TabIndex        =   33
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command7 
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
            TabIndex        =   32
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command5 
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
            TabIndex        =   31
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command4 
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
            TabIndex        =   30
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command3 
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
            TabIndex        =   29
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command2 
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
            TabIndex        =   28
            Top             =   1890
            Width           =   1590
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0F8&
      Caption         =   "Programacion de Materiales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   3690
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   2220
         TabIndex        =   17
         Top             =   1000
         Width           =   2280
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
            Left            =   -15
            Picture         =   "GENOCOM09.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Genera y Emite Nota de Pedido de Materiales al Laboratorio"
            Top             =   0
            Width           =   530
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Nota de Pedido al Laboratorio"
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
            Left            =   720
            TabIndex        =   19
            Top             =   105
            Width           =   1365
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   600
         ScaleHeight     =   585
         ScaleWidth      =   2220
         TabIndex        =   20
         Top             =   480
         Width           =   2280
         Begin VB.CommandButton Command8 
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
            Left            =   0
            Picture         =   "GENOCOM09.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Registracion de Ordenes de Compras"
            Top             =   0
            Width           =   530
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Orden de Compras (Proveedores)"
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
            Left            =   615
            TabIndex        =   21
            Top             =   70
            Width           =   1605
         End
      End
      Begin VB.Frame Frame2 
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
         TabIndex        =   7
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command15 
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
            TabIndex        =   13
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command13 
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
            TabIndex        =   12
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command12 
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
            TabIndex        =   11
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command11 
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
            TabIndex        =   10
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command10 
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
            TabIndex        =   9
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command9 
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
            TabIndex        =   8
            Top             =   420
            Width           =   1590
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0F8&
      Caption         =   "Consultas "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   120
      TabIndex        =   14
      Top             =   3360
      Width           =   3690
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   600
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   15
         Top             =   360
         Width           =   2325
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
            Left            =   0
            Picture         =   "GENOCOM09.frx":076A
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
            Top             =   0
            Width           =   650
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Consultas y Listados"
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
            Left            =   800
            TabIndex        =   16
            Top             =   90
            Width           =   1560
         End
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00000080&
      Height          =   6840
      Left            =   3960
      ScaleHeight     =   6780
      ScaleWidth      =   6630
      TabIndex        =   0
      Top             =   -165
      Width           =   6690
      Begin VB.CommandButton Command18 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   120
         Picture         =   "GENOCOM09.frx":0A74
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Pendientes de Recepcion / Listados de Requerimientos"
         Top             =   1845
         Width           =   650
      End
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
         Height          =   640
         Left            =   120
         Picture         =   "GENOCOM09.frx":0D7E
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Ver por Pantalla y Re-Imprimir O/Compra y Notas de Pedido"
         Top             =   2445
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   4
         Top             =   4005
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   12000
         End
      End
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
         Height          =   640
         Left            =   105
         Picture         =   "GENOCOM09.frx":1088
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   650
      End
      Begin VB.CommandButton Command29 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "GENOCOM09.frx":11D2
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1020
         Width           =   650
      End
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
         Height          =   640
         Left            =   120
         Picture         =   "GENOCOM09.frx":14DC
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Configuración de Procesos y Formularios"
         Top             =   3165
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -270
         Picture         =   "GENOCOM09.frx":191E
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   4275
         Width           =   1515
      End
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   10395
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu Orfab 
      Caption         =   "Ordenes de Fabricacion"
      Visible         =   0   'False
      Begin VB.Menu NvaOrd 
         Caption         =   "Nueva Orden"
         Visible         =   0   'False
      End
      Begin VB.Menu php0 
         Caption         =   "-"
      End
      Begin VB.Menu LanOrd 
         Caption         =   "Lanzar Orden"
      End
      Begin VB.Menu Cerrar 
         Caption         =   "Cerrar Orden"
      End
      Begin VB.Menu php1 
         Caption         =   "-"
      End
      Begin VB.Menu SusOrd 
         Caption         =   "Suspensión de Orden"
      End
      Begin VB.Menu AnuOrd 
         Caption         =   "Anulación de Orden"
      End
      Begin VB.Menu ModOrd 
         Caption         =   "Modificación de Orden"
      End
      Begin VB.Menu php2 
         Caption         =   "-"
      End
      Begin VB.Menu VerImp 
         Caption         =   "Ver y Re-Imprimir"
      End
   End
   Begin VB.Menu Edicion 
      Caption         =   "Edición"
      Visible         =   0   'False
      Begin VB.Menu SuOrd 
         Caption         =   "Suspender Orden"
      End
      Begin VB.Menu AnulOrd 
         Caption         =   "Anular Orden"
      End
      Begin VB.Menu ModiOrd 
         Caption         =   "Modificar Orden"
      End
      Begin VB.Menu php3 
         Caption         =   "-"
      End
      Begin VB.Menu VeryReimpi 
         Caption         =   "Ver y Re-Imprimir"
      End
   End
   Begin VB.Menu AsiMaPri 
      Caption         =   "Materiales"
      Visible         =   0   'False
      Begin VB.Menu MatReser 
         Caption         =   "Reserva de Materiales"
         Begin VB.Menu GenReservOC 
            Caption         =   "Generar Reservas por Orden de Fabricación"
         End
         Begin VB.Menu CanceReser 
            Caption         =   "Cancelación de Reservas"
         End
         Begin VB.Menu ConOrFaCu 
            Caption         =   "Control Segun Ordenes de Fabricación en Curso"
         End
         Begin VB.Menu DepuBdD 
            Caption         =   "Depuración de Base de Datos"
         End
         Begin VB.Menu php3a 
            Caption         =   "-"
         End
         Begin VB.Menu LiToMat 
            Caption         =   "Listado de Totales por Material"
            Begin VB.Menu LiToPan 
               Caption         =   "por Pantalla"
            End
            Begin VB.Menu LiToImpre 
               Caption         =   "por Impresión"
            End
         End
         Begin VB.Menu CoGrlRe 
            Caption         =   "Consulta General de Reservas por Orden de Fabricación"
         End
         Begin VB.Menu ReXNumOF 
            Caption         =   "Reservas por Numero de Orden de Fabricación"
         End
         Begin VB.Menu ReXMat 
            Caption         =   "Reservas por Material"
         End
      End
      Begin VB.Menu ValMate 
         Caption         =   "Vale de Materiales"
         Begin VB.Menu ValPrep 
            Caption         =   "Vale de Preparacion"
         End
         Begin VB.Menu ValTran 
            Caption         =   "Vale de Transferencia"
         End
      End
   End
   Begin VB.Menu AcDirec 
      Caption         =   "Accesos Directos"
      Visible         =   0   'False
      Begin VB.Menu InfoPro 
         Caption         =   "Informes de Producción"
      End
      Begin VB.Menu CarMaq 
         Caption         =   "Carga de Maquinas"
      End
      Begin VB.Menu php15 
         Caption         =   "-"
      End
      Begin VB.Menu MaestArt 
         Caption         =   "Maestro de Articulos"
      End
      Begin VB.Menu php16 
         Caption         =   "-"
      End
      Begin VB.Menu otraccdir 
         Caption         =   "Otros Accesos Directos"
      End
   End
   Begin VB.Menu ConsuLis 
      Caption         =   "Consultas y Listados"
      Visible         =   0   'False
      Begin VB.Menu ConGrBdD 
         Caption         =   "Consulta General de la Base de Datos"
      End
      Begin VB.Menu php4 
         Caption         =   "-"
      End
      Begin VB.Menu PenSecPro 
         Caption         =   "Pendientes por Sector Productivo"
         Begin VB.Menu Xpanta 
            Caption         =   "por Pantalla"
         End
         Begin VB.Menu Ximpres 
            Caption         =   "por Impresión"
         End
      End
      Begin VB.Menu PenNumOF 
         Caption         =   "Pendientes por Numero de Orden de Fabricación"
      End
      Begin VB.Menu PenCodPro 
         Caption         =   "Pendientes por Código de Producto"
         Begin VB.Menu porPantall 
            Caption         =   "por Pantalla"
         End
         Begin VB.Menu Ximpreso 
            Caption         =   "por Impresión"
         End
      End
      Begin VB.Menu php5 
         Caption         =   "-"
      End
      Begin VB.Menu ConSelEst 
         Caption         =   "Consulta Selectiva por Estado"
         Begin VB.Menu Xpant 
            Caption         =   "por Pantalla"
         End
         Begin VB.Menu Ximpre 
            Caption         =   "por Impresión"
         End
      End
      Begin VB.Menu php6 
         Caption         =   "-"
      End
      Begin VB.Menu VeryReimp 
         Caption         =   "Ver y Re-Imprimir"
      End
      Begin VB.Menu php24 
         Caption         =   "-"
      End
      Begin VB.Menu ProCar 
         Caption         =   "Programación de Carga"
         Begin VB.Menu OperPro 
            Caption         =   "Operaciones Programadas"
         End
         Begin VB.Menu CaToEqui 
            Caption         =   "Carga Total por Equipo"
         End
      End
   End
   Begin VB.Menu ConfGral 
      Caption         =   "Configuración"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "GENOCOM09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
    Call FINAL("")
End Sub

Private Sub Command14_Click()
   '
   Command14.Enabled = False
   If DERACCE%("STPOFAB", "Configuracion de Funcionamiento de O/Compra") = 2 Then
      CFGNOPE09.Show 1
   End If
   Command14.Enabled = True
   '
End Sub

Private Sub Command18_Click()
   '
   Command18.Enabled = False
   OPX% = ALTERNA%("Pendientes de Recepcion\Balance de Requerimientos")
   If OPX% = 1 Then
       Call LIPENREC
       GoTo 99
     ElseIf OPX% = 2 Then
       GoTo 5
     Else
       GoTo 99
   End If
   '
   ' LISTADOS DE REQUERIMIENTOS
5  Call CARGALISEL("!ECORDEF", "ORDEFABR", "IDSUBOR/A12;REFEREN/A64", "IDSUBOR <> '' AND STATUORF < 3 AND CanApro < CanProy ORDER BY IDSUBOR ASC")
   Call FRESHECHO("!ECORDEF")
   '
   Call CIERRARCH("*.*")
10 Call SELECHO("!ECORDEF")
   NLIPIX$ = VALACT1$("!ECORDEF")
   '
   ' OJO!!! SI SE MODIFICA ESTA RUTINA(RECUOCPEN), TAMBIEN SE DEBE ACTUALIZAR EN
   ' EN EL PROYECTO DECAOT09, FORMULARIO MENCAOT09.
   Call RECUOCPEN(NLIPIX$)
   '
   If TRIM$(NLIPIX$) <> "" Then
25      OPX% = COMALTER%("Opciones de Listados de Requerimientos:\\Listado Plano\Balance de Detalle\Balance Acumulado")
        Select Case OPX%
          Case 1
            Call FINAL("?LISREQUERI(" + TRIM$(NLIPIX$) + "%)")
            GoTo 25
          Case 2
            Call FINAL("?BALREQUERI(" + TRIM$(NLIPIX$) + "%)")
            GoTo 25
          Case 3
            Call FINAL("?ACUREQUERI(" + TRIM$(NLIPIX$) + "%)")
            GoTo 25
        End Select
        GoTo 10
   End If
   '
99 Command18.Enabled = True
   '
End Sub

Sub LIPENREC()
   '
10 OPX% = ALTERNA%("Pendientes Por Fecha Ent. Solicitada\Pendientes Por Articulo\Pendientes Por Proveedor\Destino por Nota de Suministro")
   Select Case OPX%
     Case 1
      Call FINAL("?OCOMFES")
      GoTo 10
     Case 2
      Call FINAL("?OCOMART")
      GoTo 10
     Case 3
      Call FINAL("?OCOMPRO")
      GoTo 10
     Case 4
       Call CARGALISEL("PEDSUMI", "OCOMENCA", "ID_COMPROB/A16;Refe_OCom/A48", "ID_COMPROB <> '' ORDER BY FEMI_OCOM")
       Call SELECHO("PEDSUMI")
       IDSUM$ = TRIM(VALACT1("PEDSUMI"))
       If IDSUM$ <> "" Then
          Call FINAL("?DESNOTSUM(" + TRIM$(IDSUM$) + ")")
       End If
       GoTo 10
     Case Else
       Exit Sub
   End Select
   '
End Sub

Private Sub Command22_Click()
    '
    Command22.Enabled = False
10  OPX% = COMALTER%("Ver y Re-Imprimir\\Ordenes de Compras\Notas de Pedidos")
    If OPX% = 1 Then Call VEROCOMPRA: GoTo 10
    If OPX% = 2 Then Call VERNOTAPEDI: GoTo 10
    Command22.Enabled = True
    '
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("ORDEFABR")
End Sub

Private Sub Command34_Click()
   Command34.Enabled = False
   RECALMAT09.Show 1
   Command34.Enabled = True
End Sub

Private Sub Command6_Click()
   '
   Command6.Enabled = False
   Call ABRECLIEN
   '
10 OPX% = COMALTER%("Opciones de Trabajo con Notas de Pedido:\\Generar Nota de Pedido\Re-Imprimir N.P.\Anular N.P")
   Select Case OPX%
      Case 1
        GoTo 20
      Case 2
        GoTo 30
      Case 3
        GoTo 40
   End Select
   GoTo 99
   
20 Call CARGALISEL("!ECORDEF", "ORDEFABR", "IDSUBOR/A12;REFEREN/A64", "IDSUBOR <> '' AND STATUORF < 3 AND CanApro < CanProy ORDER BY IDSUBOR ASC")
   Call FRESHECHO("!ECORDEF")
   Call CIERRARCH("*.*")
25 Call SELECHO("!ECORDEF")
   '
   NORFA$ = VALACT1$("!ECORDEF")
   If NORFA$ <> "" Then
      CONDI$ = "IDSUBOR = '" & NORFA$ & "'"
      NCLIE% = XVALO(VALOBADA("ORDEFABR", "Nume_Cli", CONDI$))
      If NCLIE% < 1 Then
         Call MENSERR(24, "Falta Cliente para Solicitud de Fabricacion '" + NORFA$ + "'")
         GoTo 99
      End If
      '
      Call GENERANOTAPE(NORFA$, SAFETEXT$(NCLIE%))
   End If
   GoTo 10
   '
30 ' ver y re-imprimir nota de pedido
   CONDI$ = "Nume_Pedi > 0"
   Call CARGALISEL("PEDSUMI", "OCOMENCA", "ID_COMPROB/A16;Refe_OCom/A48", CONDI$ + " ORDER BY NUME_PEDI, FEMI_OCOM")
   Call FRESHECHO("PEDSUMI")
   Call SELECHO("PEDSUMI")
   NVALE$ = TRIM$(VALACT1$("PEDSUMI"))
   If NVALE$ = "" Then GoTo 10
   NVALE$ = TRIM$(Mid$(NVALE$, 4))
   While Len(NVALE$) < 6: NVALE$ = "0" + NVALE$: Wend
   '
   TIDOCUM$ = "Nota de Pedido " + NVALE$
   Screen.MousePointer = 11
   '
   For UI& = ULTREG&("PDOCLST") To 1 Step -1
     XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 39
      End If
   Next UI&
   Screen.MousePointer = 1
   Call COMUNI("Documento no Encontrado")
   GoTo 10

39 Call MUESTRADOC(DOCUNU&)
   GoTo 10
   '
   '
40 ' anulacion de Nota de Pedido
   ' PRIMERO RECUPERA REGISTROS DAÑADOS
   CONDI$ = "Nume_Pedi > 0 AND ID_COMPROB IS NULL"
   Dim ARREGLA As ADODB.Recordset
   Set ARREGLA = New ADODB.Recordset
   Set ARREGLA = FLEXCONN.Execute(FILTSQL$("SELECT * FROM OCOMENCA WHERE " & CONDI$))
   With ARREGLA
      Do While Not .EOF
         NPX& = XVALO(!Nume_Pedi)
         IDCOM$ = "NP." & NPX&
         Call ACTUREGISTRO("OCOMENCA", "ID_COMPROB", "Nume_Pedi = " & NPX&, IDCOM$, RAFE&)
      .MoveNext
      Loop
   End With
   ARREGLA.Close
   Set ARREGLA = Nothing
   '
   'ANULACION DE NOTA DE PEDIDO PROPIAMENTE DICHA
   CONDI$ = "Nume_Pedi > 0 AND STAT_OCOM < 3"   ' PARA EVITAR ANULADAS Y CUMPLIDAS
   Call CARGALISEL("PEDSUMI", "OCOMENCA", "ID_COMPROB/A16;Refe_OCom/A48", CONDI$ + " ORDER BY NUME_PEDI, FEMI_OCOM")
   Call FRESHECHO("PEDSUMI")
   Call SELECHO("PEDSUMI")
   IDSUM$ = TRIM(VALACT1("PEDSUMI"))
   '
   If IDSUM$ <> "" Then
      CONDI$ = "IDENLOTE = '" & IDSUM$ & "'"
      SAFE$ = SAFETEXT$(VALOBADA("BALAREQMO", "DESTINO", CONDI$))
      PREGU$ = "¿ Confirma la Anulación de la '" + IDSUM$ + "' ?"
      If SAFE$ <> "" Then PREGU$ = PREGU$ + "\  \(La Anulación Afectará la '" + SAFE$ + "')"
      PREGU$ = PREGU$ + "\  \\No, Mantenerla\Si, Anularla"
      If COMALTER%(PREGU$) = 2 Then
         NPX& = XVALO(VALOBADA("OCOMENCA", "Nume_Pedi", "ID_COMPROB = '" & IDSUM$ & "'"))
         ' anula registros en OCOMENCA
         Call ACTUREGISTRO("OCOMENCA", "Stat_Ocom", "ID_COMPROB = '" & IDSUM$ & "'", 9, RAFE&)
         ' anula registros en OCOMDETA
         Call ACTUREGISTRO("OCOMDETA", "Stat_Ocom", "Nume_Pedi = " & NPX&, 9, RAFE&)
         ' elimina registros en BALAREQMO
         Call BORRAREGISTROS("BALAREQMO", "IdenLote = '" & IDSUM$ & "'", RAFE&)
         Call COMUNI("Operacion Completada.")
         
      End If
   End If
   GoTo 10
   '
99 Command6.Enabled = True
   '
End Sub

Private Sub Command8_Click()
   '
   Command8.Enabled = False
   If EXISTE%("ORCOMP07.EXE") > 0 Then
     Call CONECRUN("ORCOMP07.EXE", "Gestion de Compras Industriales")
     GoTo 99
   End If
   '
   Call CARGALISEL("!ECOMAST", "OCOEMI", "COD_INTE/C16+MASTER/A40;", "COD_INTE <> 0")
   Call SELECHO("!ECOMAST")
10 Codigo$ = VALACT1("!ECOMAST")
   '
   NUMOCO% = 0: FECHAA% = 0: CANTX# = 0: EDITENCA = False: EDITDETA = False ' Variables del Procedimiento
   FechAlta% = 0
   CI% = CODINT%(Codigo$)
   '
   If CI% > 0 Then
      FILTRO$ = CI%
      Call STDFORM("REGOCOM", FILTRO$)
      If STDSCREEN.APRO = "OK" Then
       'Validaciones
       NUMOCO% = XVALO(VALOFORM("Text-1"))
       If NUMOCO% <= 0 Then
          COMUNI ("Falta Ingresar Nro de O/Compra")
          GoTo 10
       End If
       '
       FECHAA% = XVALO((VALOFORM("PARAM-9")))
       If FECHAA% <= 0 Then
          COMUNI ("Falta Ingresar Fecha")
          GoTo 10
       End If
       '
       CANTX# = XVALO(VALOFORM("TEXT-2"))
       If CANTX# <= 0 Then
          COMUNI ("Falta Ingresar Cantidad")
          GoTo 10
       End If
       '
       FechAlta% = XVALO(VALOFORM("TEXT-3"))
       If FechAlta% <= 0 Then
          COMUNI ("Falta Ingresar Fecha De Entrega Solicitada")
          GoTo 10
       End If
       '
       GoTo 20 'Dar de Alta
       '
      End If
      '
   End If
   GoTo 99

20
   If CantRegistros("OCOMENCA", "Nume_Ocom =" & NUMOCO%) > 0 Then
      EDITENCA = True
   End If
   '
   CONDI$ = "Nume_Ocom =" & NUMOCO% & " AND COD_INTE = " & CI%
   If CantRegistros("OCOMDETA", CONDI$) > 0 Then
      OPT1% = COMALTER%("Se Encuentra Registrado el Articulo " & Codigo$ & "\para la Orden de Compra Nro " & NUMOCO% & "\Desea Modificarla?\\No, Cancelar\Si, Confirmar")
      If OPT1% = 0 Then GoTo 99
      If OPT1% = 1 Then GoTo 10
      If OPT1% = 2 Then EDITDETA = True
      
   End If
  
   ' AQUI AGREGA UN UNICO REGISTRO DE ENCABEZADO EN LA TABLA OCOMENCA
   PROV$ = SAFETEXT(VALOFORM("PARAM-3"))
   FEMI% = FECHAA%
   Call APERBASDAT("COMPRAS")
   Dim OCOMENCA As ADODB.Recordset
   Set OCOMENCA = New ADODB.Recordset
   SQL$ = "Select * from OCOMENCA WHERE NUME_OCOM = " & NUMOCO%
   OCOMENCA.Open FILTSQL$(SQL$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   With OCOMENCA
         If EDITENCA = True Then
            .MoveFirst
            Else
            .AddNew   ' YA CONVERTIDO
            On Error Resume Next
            !CodiEmpr = CodiEmp%
            On Error GoTo 0
         End If
         '
         !Nume_Pedi = 0
         !Nume_Ocom = NUMOCO%
          REFE$ = FECHATEX$(HOY(HO$)) + " - " + RASOCLI$(XVALO(PROV$))
         !Refe_OcoM = Left$(REFE$, 64)
         !Femi_Ocom = CVDAT(FEMI%)
         !Nume_Cli = 0 'XVALO(PROV$) 'NPRO%
         !RASO_PROV = Left$(RASOCLI$(XVALO(PROV$) * (-1)), 48) 'Left$(RASOCLI$((-1) * NPRO%), 48) 'OT-06-0227-WAR-05/06/06///
         !TSum_Ocom = 0               ' materiales
'            STATUOC% = (-1)           ' SIGNIFICA PENDIENTE DE APROBACIÓN
'            If Left$(CONTROL$("OCOMPRA", "MODOAPRO"), 6) = "DIRECT" Then STATUOC% = 0
         !Stat_Ocom = 0 'STATUOC%        ' PASA A 0 (CERO) SI APROBACION DIRECTA
         !Fgen_Ocom = Now
         !Usua_Gen = USNBR%
             MONECOI% = 1 ' XVALO(Text15)
         !Mone_Cos = MONECOI%        ' MONEDA PARA COSTEO
         !Mone_Emis = MONECOI%        ' MONEDA DE EMISIÓN
         !Tcam_Ocom = 0               ' TIPO DE CAMBIO / COTIZACION
         !Cpag_Ocom = 0
         !Embala_Ocom = 0
         !AtSr_Ocom = ""
         '
            PORDESCUENTO$ = 0
            DESC$ = 0
            TOTANE# = 0
            IDESCUEN# = 0
            TOTANEDES# = 0
            TOTIVA# = 0
            TOTPER# = 0
            TOTATO# = 0
            '
         !ToBru_Ocom = TOTANE#
         !Pdes_Ocom = PORDESCU
         !IDes_Ocom = IDESCUEN#
         !Iiva_Ocom = TOTIVA#
         !Iotr1_Ocom = TOTPER#
         !Total_Ocom = TOTATO#
         !Lentrega = ""
         !Obse_Ocom = ""
         !Anex_Ocom = ""
         !PorcIva = ""
         !Anex_Ocom = "" 'SAFETEXT(NORFA$)
         !Id_Comprob = "OC." + SAFETEXT(NUMOCO%)
         .Update
   End With  'fin de OcomEnca
   OCOMENCA.Close
   Set OCOMENCA = Nothing
     

      'Dar de alta en OCOMDETA
      Dim OCOMDETA As ADODB.Recordset
      Set OCOMDETA = New ADODB.Recordset
      CONDI$ = "Nume_Ocom =" & NUMOCO% & " AND COD_INTE = " & CI%
      OCOMDETA.Open FILTSQL$("select * from OCOMDETA where " & CONDI$), FLEXCONN, adOpenKeyset, adLockPessimistic
      With OCOMDETA
         '
         If EDITDETA = True Then
            .MoveFirst
            Else
            .AddNew   ' YA CONVERTIDO
            On Error Resume Next
            !CodiEmpr = CodiEmp%
            On Error GoTo 0
         End If
         '
         CI0% = CI%
         FEMI% = FECHAA%
         
         
         !Nume_Ocom = NUMOCO%
         !Nume_Pedi = 0
         !NUORD_ITEM = ORDIT%
         !Femi_Ocom = CVDAT(FEMI%)
         !Nume_Cli = XVALO(PROV$)
         !RASO_PROV = Left$(RASOCLI$((-1) * XVALO(PROV$)), 64) 'OT-06-0227-WAR-05/06/06///
         !Medi_Ocom = PedInt$
         !COD_INTE = CI0%
         !COD_EXTE = CODEXT$(CI0%)
         !DESCRITEM = DESCRIT0$(CI0%)
         !Delibre = DESCRIT0$(CI0%)
         !Valo_Unid = 0
         !UNI_MAS = UNIMED$(CI0%) 'Mid$(ZZ$, 11, 4)
           CAUNIS = CANTX# 'XVALO(!KILO_SALDO)  ' CANT. EN UNIDADES DE STOCK
         !Unids_Ocom = CAUNIS
           CANTU = CANTX# 'XVALO(!KILO_SALDO)  ' CANT EN UNIDADES DE COMPRA
           'CANTU = REDON_INDIVIS(CI0%, CANTX#) 'XVALO(!KILO_SALDO))
         !CANT_OCOM = CANTU
         !Uni_Com = UNIMED$(CI0%) ' unidad de compra
         !Coef_Unids = 1
         !Pre_Unit = 0
         !Desc_Item = 0
         !PrUn_Neto = 0
         !INet_Item = 0
         !Fentre_Sol = FECHATEX(HOY(HO$))
         !CONSIMAT = ""
         !Cant_Rec = 0
         !Cod_Oper = 0
         !Doc_Orig = ""
         !N_OrServ = 0
         !Mone_Cos = 1        ' antes pos 83,1
         !Mone_Emis = 1        ' antes pos 84,1
         !TSum_Ocom = 0              ' materiales
         '!Stat_Ocom = APRODIR% - 1  ' status Pendiente de aprobacion
         '!IdenReque = IDREQ$ 'SAFETEXT(!IdenReque)
         !IContab = 0 ' Imputacion contable
         '!IContab = 0
         !Sect_Ocom = ""
         .Update
      End With
      OCOMDETA.Close
      Set OCOMDETA = Nothing
   
   '
99 Command8.Enabled = True
   '
End Sub

Private Sub ConfGral_Click()
Call Command14_Click
End Sub

Private Sub ConOrFaCu_Click()
Call RESERV07.Command7_Click
End Sub

Private Sub DepuBdD_Click()
Call RESERV07.Command5_Click
End Sub

Private Sub Form_Activate()
   WindowState = 0
   Openforms = DoEvents
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     Openforms = DoEvents
   Wend
End Sub

Private Sub Form_Load()
    '
    VERANTER$ = "ORDEFABR"
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    'Esta actualizacion evita el error de Valor Nulo cuando se lee el campo Priorid
    Call ABRECONEXION
    Call ACTUREGISTRO("ORDEFABR", "Priorid", "Priorid IS NULL", "0", REG&)
    '
    If APLINVO$ = "OPRESERV" Then
      RESERV07.Show 1
      Call FINAL("")
    End If
    '
    If UCase$(Left$(APLINVO$, 6)) = "ORFNUE" Then
       APLIXYZ$ = APLINVO$
    End If
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Call GRATITFOR(Caption)
    '
    If InStr(EPAC$, "BOREN") > 0 Then
       Command12.Visible = True
    End If
    '
    Call VERIDEST
    If InStr(UCase$(EPAC$), "BORENS") > 0 Then
       Picture1.Width = Picture1.Width + Command4.Width
       Picture1.Left = Picture1.Left - Command4.Width / 2
    End If
    '
    Openforms = DoEvents
    Show
    '
End Sub
'
Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call FINAL("")
End Sub
'
Sub GENERANOTAPE(NORFA$, PROV$)   'ATENCION, ESTO SE TRAJO DEL PROYECTO ORPROD09
   '
   FEMI% = HOY(HOS)
   '
   CONDI$ = "IDSUBOR ='" & NORFA$ & "'"
   NOCOM$ = VALOBADA("ordefabr", "NUME_OCOM", CONDI$)
   If TRIM$(NOCOM$) = "" Then
      COMUNI ("Imposible Continuar\Falta O/Compra Para " & NORFA$)
      GoTo 99
   End If
   Dim BALAREQ As ADODB.Recordset
   Set BALAREQ = New ADODB.Recordset
   '
   ' VERIFICA SI YA SE REALIZO UNA NOTA DE PEDIDO PARA ASI NO DUPLICAR LO PEDIDO
   CONDI$ = "Nume_Pedi > 0 AND Anex_Ocom='" & NORFA$ + "' "
   CONDI$ = CONDI$ + " AND Nume_Cli=" & PROV$
   CONDI$ = CONDI$ + " AND STAT_OCOM < 9 "        ' PARA DESCARTAR LAS ANULADAS"
   If CantRegistros("OCOMENCA", CONDI$) > 0 Then
      NUNOPE& = XVALO(VALOBADA("OCOMENCA", "Nume_Pedi", CONDI$))
      CONDI2$ = "Nume_Pedi = " & NUNOPE&
      CONDI2$ = CONDI2$ + " AND Nume_Cli=" & PROV$
      CONDI2$ = CONDI2$ + " AND CANT_REC <= CANT_OCOM"
      CONDI2$ = CONDI2$ + " AND STAT_OCOM < 3"
      If CantRegistros("OCOMDETA", CONDI2$) > 0 Then
        Call COMUNI("Ya se Ha Generado Nota de Pedido Nro " & NUNOPE& & "\para la '" & NORFA$ & "' Solicitada.\   \Para Refrescar, Anule y Re-Emita.")
        Exit Sub
      End If
   End If
   '
   Dim OCOMDETA As ADODB.Recordset
   Set OCOMDETA = New ADODB.Recordset
   OCOMDETA.Open FILTSQL$("select * from OCOMDETA where raso_prov ='x'"), FLEXCONN, adOpenKeyset, adLockPessimistic
   '
   Dim BALTEMP As ADODB.Recordset
   Set BALTEMP = New ADODB.Recordset
   BALTEMP.Open FILTSQL$("SELECT * FROM BALAREQMO WHERE IDENREQUE='X'"), FLEXCONN, adOpenKeyset, adLockOptimistic
   '
   SQL$ = "SELECT IDENREQUE, COD_INTE, SUM(KILO_REQUE) AS SUREQUE, SUM(KILO_ASIG) AS SUASIG FROM BALAREQMO WITH(NOLOCK) "
   SQL$ = SQL$ + " WHERE DESTINO = '" & NORFA$ & "' "
   SQL$ = SQL$ + " GROUP BY IDENREQUE, COD_INTE  "
   Set BALAREQ = FLEXCONN.Execute(FILTSQL$(SQL$))
   If BALAREQ.BOF = True And BALAREQ.EOF = True Then GoTo 99
   '
   ' AQUI AGREGA UN UNICO REGISTRO DE ENCABEZADO EN LA TABLA OCOMENCA
   Call APERBASDAT("COMPRAS")
   Dim OCOMENCA As ADODB.Recordset
   Set OCOMENCA = New ADODB.Recordset
   OCOMENCA.Open FILTSQL$("Select * from OCOMENCA WHERE RASO_PROV ='X'"), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   NUMOCOMPRA& = XVALO(VALOBADA("OCOMENCA", "Max(Nume_Pedi)", "Nume_pedi > 0", "NOMESS")) + 1
   With OCOMENCA
         .AddNew   ' YA CONVERTIDO
         On Error Resume Next
         !CodiEmpr = CodiEmp%
         On Error GoTo 0
          !Nume_Pedi = NUMOCOMPRA&
             'FEX = FEMI%
             REFE$ = FECHATEX$(HOY(HO$)) + " - " + RASOCLI$(XVALO(PROV$))
          !Refe_OcoM = Left$(REFE$, 64)
          !Femi_Ocom = CVDAT(FEMI%)
          !Nume_Cli = XVALO(PROV$) 'NPRO%
          '!RASO_PROV = Left$(RASOCLI$(XVALO(PROV$) * (-1)), 48) 'Left$(RASOCLI$((-1) * NPRO%), 48) 'OT-06-0227-WAR-05/06/06///
          !RASO_PROV = Left$(RASOCLI$(XVALO(PROV$)), 48)    ' SE QUITA EL "*(-1)" PORQUE LAS N/P SON EN BASE A CLIENTES
          !TSum_Ocom = 0               ' materiales
             STATUOC% = (-1)           ' SIGNIFICA PENDIENTE DE APROBACIÓN
             If Left$(CONTROL$("OCOMPRA", "MODOAPRO"), 6) = "DIRECT" Then STATUOC% = 0
          !Stat_Ocom = STATUOC%        ' PASA A 0 (CERO) SI APROBACION DIRECTA
          '
          !Fgen_Ocom = Now
          !Usua_Gen = USNBR%
             MONECOI% = 1 ' XVALO(Text15)
          !Mone_Cos = MONECOI%        ' MONEDA PARA COSTEO
          !Mone_Emis = MONECOI%        ' MONEDA DE EMISIÓN
          !Tcam_Ocom = 0               ' TIPO DE CAMBIO / COTIZACION
          !Cpag_Ocom = 0
          !Embala_Ocom = 0
          !AtSr_Ocom = ""
            '
            PORDESCUENTO$ = 0
            DESC$ = 0
            TOTANE# = 0
            IDESCUEN# = 0
            TOTANEDES# = 0
            TOTIVA# = 0
            TOTPER# = 0
            TOTATO# = 0
            '
         !ToBru_Ocom = TOTANE#
         !Pdes_Ocom = PORDESCU
         !IDes_Ocom = IDESCUEN#
         !Iiva_Ocom = TOTIVA#
         !Iotr1_Ocom = TOTPER#
         !Total_Ocom = TOTATO#
         !Lentrega = ""
         !Obse_Ocom = ""
         !Anex_Ocom = ""
         !PorcIva = ""
         !Anex_Ocom = SAFETEXT(NORFA$)
         !Id_Comprob = "NP." + SAFETEXT(NUMOCOMPRA&)
         .Update
   End With  'fin de OcomEnca
   OCOMENCA.Close
   Set OCOMENCA = Nothing
     
   'Variables utilizadas en Alta Detalle de Pedido
   APRODIR% = 0
   If Left$(CONTROL$("OCOMPRA", "MODOAPRO"), 6) = "DIRECT" Then
     APRODIR% = 1
   End If
   ORDIT% = 0
   '
   With BALAREQ
     .MoveFirst
     Do While Not (.EOF)
       CIXI% = XVALO(!COD_INTE)
       CODE$ = CODEXT$(CIXI%)
       IDREQ$ = SAFETEXT$(!IdenReque)
       If CRITREP%(CIXI%) = 1 Then
         If XVALO(!SUREQUE) > XVALO(!SUASIG) Then
            '
            If PROVHABI%(CIXI%) < 1 Then
              QTR& = CantRegistros("FORMUPAT", "CODICONJ = " & CIXI%)
              If QTR& > 0 Then                ' COMPONENTE CON FORMULA, PROGRAMA POR o/prod
                 TTXX$ = "Codigo '" + TRIM$(CODE$) + " - " + DESCRIT0$(CIXI%)
                 TTXX$ = TTXX$ + "\Requerido para esta S/Fabricación.\   \Debe Programarse por Orden de Produccion."
                 Call COMUNI(TTXX$)
                 GoTo 39
              End If
            End If
            'Registro de Detalle Pedido
            ORDIT% = ORDIT% + 1
            CI0% = XVALO(!COD_INTE)
            
            OCOMDETA.AddNew   ' YA CONVERTIDO
                    On Error Resume Next
            OCOMDETA!CodiEmpr = CodiEmp%
                    On Error GoTo 0
            OCOMDETA!Nume_Pedi = NUMOCOMPRA&
            OCOMDETA!NUORD_ITEM = ORDIT%
            OCOMDETA!Femi_Ocom = CVDAT(FEMI%)
            OCOMDETA!Nume_Cli = XVALO(PROV$)
            OCOMDETA!RASO_PROV = Left$(RASOCLI$((-1) * XVALO(PROV$)), 64) 'OT-06-0227-WAR-05/06/06///
            OCOMDETA!Medi_Ocom = PedInt$
            OCOMDETA!COD_INTE = CI0%
            OCOMDETA!COD_EXTE = CODEXT$(CI0%)
            OCOMDETA!DESCRITEM = DESCRIT0$(CI0%)
            OCOMDETA!Delibre = DESCRIT0$(CI0%)
            OCOMDETA!Valo_Unid = 0
            OCOMDETA!UNI_MAS = UNIMED$(CI0%) 'Mid$(ZZ$, 11, 4)
              CAUNIS = XVALO(!SUREQUE) - XVALO(!SUASIG) ' CANT. EN UNIDADES DE STOCK
            OCOMDETA!Unids_Ocom = CAUNIS
              'CANTU = XVALO(!KILO_SALDO)  ' CANT EN UNIDADES DE COMPRA
              CANTU = REDON_INDIVIS(CI0%, XVALO(!SUREQUE) - XVALO(!SUASIG))
            OCOMDETA!CANT_OCOM = CANTU
            OCOMDETA!Uni_Com = UNIMED$(CI0%) ' unidad de compra
            OCOMDETA!Coef_Unids = 1
            OCOMDETA!Pre_Unit = 0
            OCOMDETA!Desc_Item = 0
            OCOMDETA!PrUn_Neto = 0
            OCOMDETA!INet_Item = 0
            OCOMDETA!Fentre_Sol = FECHATEX(HOY(HO$))
            OCOMDETA!CONSIMAT = ""
            OCOMDETA!Cant_Rec = 0
            OCOMDETA!Cod_Oper = 0
            OCOMDETA!Doc_Orig = ""
            OCOMDETA!N_OrServ = 0
            OCOMDETA!Mone_Cos = 1        ' antes pos 83,1
            OCOMDETA!Mone_Emis = 1        ' antes pos 84,1
            OCOMDETA!TSum_Ocom = 0              ' materiales
            OCOMDETA!Stat_Ocom = APRODIR% - 1  ' status Pendiente de aprobacion
            OCOMDETA!IdenReque = SAFETEXT(!IdenReque)
            OCOMDETA!IContab = 0 ' Imputacion contable
            '!IContab = 0
            OCOMDETA!Sect_Ocom = ""
            OCOMDETA.Update
            'Fin Registro de Pedido
            '
            '
            'Se agrega Registro en Balareqmo Informando el Pedido Generado
            CONDI$ = "IDENREQUE = '" + IDREQ$ + "' AND KILO_REQUE > 0"
            NROPED$ = "NP." & SAFETEXT(NUMOCOMPRA&)
            BALTEMP.AddNew   ' YA CONVERTIDO
                   On Error Resume Next
            BALTEMP!CodiEmpr = CodiEmp%
                   On Error GoTo 0
            BALTEMP!CODICOMPRO = VALOBADA("BALAREQMO", "CODICOMPRO", CONDI$, "NOMESS")        ' SAFETEXT(!CODICOMPRO)
            BALTEMP!NUMECOMPRO = XVALO(VALOBADA("BALAREQMO", "NUMECOMPRO", CONDI$, "NOMESS")) ' XVALO(!NUMECOMPRO)
            BALTEMP!CODSUCU = XVALO(VALOBADA("BALAREQMO", "CODSUCU", CONDI$, "NOMESS"))       ' XVALO(!CODSUCU)
            BALTEMP!NUORITEM = XVALO(VALOBADA("BALAREQMO", "NUORITEM", CONDI$, "NOMESS"))     'XVALO(!NUORITEM)
            BALTEMP!IdenReque = IDREQ$                                                        ' SAFETEXT(!IDENREQUE)
            BALTEMP!COD_INTE = CIXI%                                                          ' XVALO(!Cod_Inte)
            BALTEMP!idenlote = NROPED$
            BALTEMP!Kilo_Asig = CANTU                                                         ' XVALO(!KILO_SALDO)
            BALTEMP!FeReMovi = FECHATEX$(HOY(HO))
            BALTEMP!fechMov = FECHATEX$(HOY(HO))
            BALTEMP!CREADOPOR = "FLEX"
            BALTEMP!Destino = VALOBADA("BALAREQMO", "DESTINO", CONDI$, "NOMESS")              'SAFETEXT(!DESTINO)
            BALTEMP.Update
            '
         End If
      End If
39    .MoveNext
      Loop
    End With
    OCOMDETA.Close
    Set OCOMDETA = Nothing
    '
    'Ventana de Opcion Impresion o Email
    PORPRIN% = 1: PORMAIL% = 0
    NCIX% = XVALO(PROV$) * (-1)
    EMAI$ = TRIM$(EMAILCLI$(NCIX%))
    If EMAI$ = "" Or InStr(EMAI$, "@") < 1 Then
       GoTo 50 'NO TIENE DIRECCION O DIRECCION NO VALIDA
    End If
    '
12  SALIPED07.Show 1
    Call FRESHALL
    If SALIPED07.Check1.Value < 1 Then
       If SALIPED07.Check2.Value < 1 Then
          Call MENSERR(24, "Debe Marcar al Menos\una Dirección de Salida.")
          GoTo 12
       End If
    End If
    PORPRIN% = 0: If SALIPED07.Check1.Value = 1 Then PORPRIN% = 1
    PORMAIL% = 0: If SALIPED07.Check2.Value = 1 Then PORMAIL% = 1
    '
50  If PORPRIN% > 0 Then      ' EMITE POR IMPRESORA
       Call PRICOMPRA(NUMOCOMPRA&)
    End If
    '
    If PORMAIL% > 0 Then      ' ENVIA POR MAIL
        SENDBYMAIL$ = EMAI$
        For Each PRX1 In Printers
           If InStr(UCase$(PRX1.DeviceName), "PDFWRITER") > 0 Then GoTo 25
        Next
        Call MENSERR(24, "Imposible Enviar por Mail.\   \Falta Impresora Virtual PDFWRITER.")
        SENDBYMAIL = ""
        '
25      Call PRICOMPRA(NUMOCOMPRA&)
        SENDBYMAIL$ = ""
    End If
    '
    BALAREQ.Close
    Set BALAREQ = Nothing
    '
99  End Sub

Sub PRICOMPRA(OCOM&) 'ATENCION, ESTO SE TRAJO DEL PROYECTO ORPROD09
    '
    PRIPLANOI% = 0
    PRIESPEC% = 0
    PRIHRINS% = 0
    '
    ' IMPRESION DE LA O/COMPRA
    '
    FORIPRE$ = CONTROL$("", "FORNOPE")
    '
    If TRIM$(FORIPRE$) <> "" Then
      '
      Screen.MousePointer = 11
      FECHEMII% = HOY(HO$) 'FECHANUM(Text9)
      FECHDOC$ = FETEXTO$(FECHEMII%)
      NUMEDOC$ = OCOM&
      While Len(NUMEDOC$) < 6
        NUMEDOC$ = "0" + NUMEDOC$
      Wend
      TIPODOC$ = "Nota de Pedido"
      '
      CODPARAM$(1) = "NUME-COM"
      CODPARAM$(2) = "FECH-EMI"
      CODPARAM$(3) = "DOCUORIG"
      CODPARAM$(4) = "DESTINO"
      CODPARAM$(5) = "NRO-LOTE"
      CODPARAM$(6) = "N-REMITO"
      CODPARAM$(7) = "N-OCOMPR"
      CODPARAM$(8) = "RASO-CLI"
      CODPARAM$(9) = "DOMI-CLI"
      CODPARAM$(10) = "CPOS-CLI"
      CODPARAM$(11) = "TRANSPOR"
      CODPARAM$(12) = "LOCA-CLI"
      
      CAPARDOC% = 12
      '
      'Encabezado
      '
      CONDIOCOM$ = "NUME_PEDI = " & OCOM&
      '
      NORFA$ = VALOBADA("OCOMENCA", "ANEX_OCOM", CONDIOCOM$)
      '
      CONDIFAB$ = "IDSUBOR ='" & NORFA$ & "'"
      '
      FECH$ = FECHATEX$(CVINT%(VALOBADA("OCOMENCA", "FGEN_OCOM", CONDIOCOM$)))
      '
      DESTI$ = VALOBADA("ORDEFABR", "DESTINO", CONDIFAB$)
      LOTE$ = VALOBADA("ORDEFABR", "IDENLOTE", CONDIFAB$)
      REMI$ = VALOBADA("ORDEFABR", "NumRem_Prov", CONDIFAB$)
      OCOMCLI$ = VALOBADA("ORDEFABR", "NUME_OCOM", CONDIFAB$)
      NPROV% = XVALO(VALOBADA("OCOMENCA", "NUME_CLI", CONDIOCOM$))
      
      NUTRA% = NUTRANCLI(NPROV%)
      DENOTRA$ = DENOTRANS$(NUTRA%)
     
     
      DOCPARAM$(1) = NUMEDOC$
      'DOCPARAM$(2) = FECH$
      DOCPARAM$(3) = NORFA$
      DOCPARAM$(4) = DESTI$                                    '  Text12
      DOCPARAM$(5) = LOTE$
      DOCPARAM$(6) = REMI$
      DOCPARAM$(7) = OCOMCLI$                                  ' Mid$(X1$, 147, 20)
      DOCPARAM$(8) = RASOCLI$(NPROV%)                          ' Mid$(X1$, 67, 40)
      DOCPARAM$(9) = DOMICLI$(NPROV%)                          ' Mid$(X1$, 167, 40)
      DOCPARAM$(10) = CPOSCLI$(NPROV%)                         ' Mid$(X1$, 107, 40)
      DOCPARAM$(11) = DENOTRA$
      DOCPARAM$(12) = LOCACLI$(NPROV%)
      
      
      HOII% = HOY(HOS$)
      FECHENT% = HOII%
'      DOCPARAM$(15) = Text14
'      If FECHENT% = HOII% Then
'          DOCPARAM$(15) = DOCPARAM$(15) + " (inmediato)"
'        ElseIf FECHENT% < HOII% + 4 Then
'          DOCPARAM$(15) = DOCPARAM$(15) + " (urgente!!)"
'      End If
'      DOCPARAM$(16) = TRIM$(Label11)
'      DOCPARAM$(17) = ""
'      If MONEMI% > 1 Then DOCPARAM$(17) = Str$(TICAMONE(MONEMI%))
'      DOCPARAM$(18) = Eco21(1)
'      DOCPARAM$(19) = Text10
'      DOCPARAM$(20) = Total(4)
'      '
'      ' esto es para poner el importe en letras
'      TTTT# = XVALO(Total(3))
'      TOTFAC$ = NUMTEX$(TTTT#)
'      If Label11 <> "" And Label11 <> "Pesos" Then
'         PPXX% = InStr(TOTFAC$, "PESOS")
'         If PPXX% > 0 Then
'           TOTFAC$ = Left$(TOTFAC$, PPXX% - 1) + UCase$(Label11) + Mid$(TOTFAC$, PPXX% + 5)
'         End If
'      End If
'      Call FRATEXTO(TOTFAC$, 40)
'      CODPARAM$(22) = "IMP-LET1": DOCPARAM$(22) = RETEX$(1)
'      CODPARAM$(23) = "IMP-LET2": DOCPARAM$(23) = RETEX$(2)
'      CODPARAM$(24) = "IMP-LET3": DOCPARAM$(24) = RETEX$(3)
'      '
'      CODTABU2$(1) = "REF-MAT3"
'      CACOLTAB2% = 1
      '
      CODTABU1$(1) = "COD-MAT"
      CODTABU1$(2) = "DES-MAT"
      CODTABU1$(3) = "UNIMED"
      CODTABU1$(4) = "CANTIDAD"
      CODTABU1$(5) = "ORD-ITEM"
      CODTABU1$(6) = "COD-CONJ"
      CACOLTAB1% = 6
      '
      DESTIDOC% = (-1) * XVALO(Text1)
      CAITAB1% = 0
      CAITCRU% = 0
      INCLUSPE% = 1
      REPLAORI% = 0
      RECOPROP% = 0
      REVICODX% = 0
      '
      
      '
      Dim OCOMDETA As ADODB.Recordset
      Set OCOMDETA = New ADODB.Recordset
      SQLX$ = "SELECT * FROM OCOMDETA WHERE NUME_PEDI = " & OCOM&
      Set OCOMDETA = FLEXCONN.Execute(FILTSQL$(SQLX$))
      If OCOMDETA.EOF And OCOMDETA.BOF Then GoTo 999
      '
      With OCOMDETA
       .MoveFirst
       CAITAB1% = 0
       Do While Not (.EOF)
       CAITAB1% = CAITAB1% + 1
       CODDX$ = SAFETEXT(!COD_EXTE)    ' TRIM$(Text7)
       TETABU1$(1, CAITAB1%) = CODDX$
       DECRIT$ = SAFETEXT(!DESCRITEM)   'Text18           ' TRIM$(DESCRIT0$(CI0%))
       TETABU1$(2, CAITAB1%) = DECRIT$
       TETABU1$(3, CAITAB1%) = SAFETEXT(!UNI_MAS)          ' Mid$(X$, 59, 4)
       TETABU1$(4, CAITAB1%) = XVALO(!CANT_OCOM)
       TETABU1$(5, CAITAB1%) = XVALO(!NUORD_ITEM)
       TETABU1$(6, CAITAB1%) = NUPLANO$(XVALO(!COD_INTE))
       '
       .MoveNext
       Loop
      End With
      OCOMDETA.Close
      Set OCOMDETA = Nothing
      '
      Call PRINTDOC("FORNOPE")                  'Pedidos  ' Orden de Compra
      '
      End If
999 Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
End Sub

Sub VEROCOMPRA()
    'Cargo la lista con las O/C
    Call APERBASDAT("COMPRAS")
    SQLX$ = " SELECT Nume_Ocom, Refe_Ocom FROM OCOMENCA"
    SQLX$ = SQLX$ + " ORDER BY Nume_Ocom "
    'Set VRIMPTEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim VRIMPTEMP As ADODB.Recordset
    Set VRIMPTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With VRIMPTEMP
        On Error Resume Next ' si hay algun error msgbox err
        .MoveFirst
        If Err Then
          On Error GoTo 0
          Call MENSERR(24, "No Aparecen Ordenes de Compra Emitidas")
          Exit Sub
        End If
        On Error GoTo 0
        JJ& = 0
        Do While (.EOF = False)
            XXY$ = REGBLAN$("INDIOCOM")
            If XVALO(!Nume_Ocom) > 0 Then
                Call REPLA(XXY$, MKS(!Nume_Ocom), 1, 4)
                Call REPLA(XXY$, !Refe_OcoM, 5, 44)
                JJ& = JJ& + 1
                Call GRAREG("!INDIOCOM", XXY$, JJ&)
            End If
            .MoveNext
        Loop
    End With
    VRIMPTEMP.Close
    Set VRIMPTEMP = Nothing
    Call SETULTREG("!INDIOCOM", JJ&)
    Call CIERRARCH("!INDIOCOM")
    Call FRESHECHO("!INDIOCOM")
    '
3   Call SELECHO("!INDIOCOM/Indice General de Ordenes de Compra (Pendientes-Cumplidas-Anuladas)")
    NPX& = XVALO(VALACT1$("!INDIOCOM")) ' ACA LEO DEL INDIOCOM
    If NPX& < 1 Then Exit Sub
    '
    TIDOCUM$ = "Nota de Pedido"
    NUDOCU$ = TRIM$(Str$(NPX&))
    While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
    NUDOCU1$ = " " + NUDOCU$ + " "
    NUDOCU2$ = " " + NUDOCU$ + "-"
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        PPXX% = InStr(XX$, "Nro.")
        If PPXX% > 0 Then
          XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
        End If
        If InStr(XX$, NUDOCU1$) > 4 Or InStr(XX$, NUDOCU2$) > 4 Then
          DOCUNU& = CVS(Left$(XX$, 4))
          GoTo 5
        End If
      End If
    Next UI&
    Screen.MousePointer = 1
    Call COMUNI("Documento no Encontrado")
    GoTo 3
    '
5   Call MUESTRADOC(DOCUNU&)
    GoTo 3
    '

End Sub

Sub VERNOTAPEDI()

    'Cargo la lista con las N/P
    Call APERBASDAT("COMPRAS")
    SQLX$ = " SELECT Nume_Pedi, Refe_Ocom FROM OCOMENCA"
    SQLX$ = SQLX$ + " ORDER BY Nume_Ocom "
    'Set VRIMPTEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim VRIMPTEMP As ADODB.Recordset
    Set VRIMPTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With VRIMPTEMP
        On Error Resume Next ' si hay algun error msgbox err
        .MoveFirst
        If Err Then
          On Error GoTo 0
          Call MENSERR(24, "No Aparecen Ordenes de Compra Emitidas")
          Exit Sub
        End If
        On Error GoTo 0
        JJ& = 0
        Do While (.EOF = False)
            XXY$ = REGBLAN$("INDIOCOM")
            If XVALO(!Nume_Pedi) > 0 Then
                Call REPLA(XXY$, MKS(!Nume_Pedi), 1, 4)
                Call REPLA(XXY$, !Refe_OcoM, 5, 44)
                JJ& = JJ& + 1
                Call GRAREG("!INDIOCOM", XXY$, JJ&)
            End If
            .MoveNext
        Loop
    End With
    VRIMPTEMP.Close
    Set VRIMPTEMP = Nothing
    Call SETULTREG("!INDIOCOM", JJ&)
    Call CIERRARCH("!INDIOCOM")
    Call FRESHECHO("!INDIOCOM")
    '
3   Call SELECHO("!INDIOCOM/Indice General de Notas de Pedidos")
    NPX& = XVALO(VALACT1$("!INDIOCOM")) ' ACA LEO DEL INDIOCOM
    If NPX& < 1 Then Exit Sub
    '
    TIDOCUM$ = "Nota de Pedido"
    NUDOCU$ = TRIM$(Str$(NPX&))
    While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
    NUDOCU1$ = " " + NUDOCU$ + " "
    NUDOCU2$ = " " + NUDOCU$ + "-"
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        PPXX% = InStr(XX$, "Nro.")
        If PPXX% > 0 Then
          XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
        End If
        If InStr(XX$, NUDOCU1$) > 4 Or InStr(XX$, NUDOCU2$) > 4 Then
          DOCUNU& = CVS(Left$(XX$, 4))
          GoTo 5
        End If
      End If
    Next UI&
    Screen.MousePointer = 1
    Call COMUNI("Documento no Encontrado")
    GoTo 3
    '
5   Call MUESTRADOC(DOCUNU&)
    GoTo 3
    '
End Sub

Sub RECUOCPEN(NLIPIX$)
   '
   SQLX$ = "SELECT * FROM BALAREQMO "
   SQLX$ = SQLX$ + " WHERE Kilo_Asig > 0 "
   SQLX$ = SQLX$ + " AND IDENREQUE LIKE '" & NLIPIX$ & "%' "
   SQLX$ = SQLX$ + " AND (IDENLOTE LIKE 'OC%' OR IDENLOTE LIKE 'NP%')"
   Dim BALRE As ADODB.Recordset
   Set BALRE = New ADODB.Recordset
   BALRE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   With BALRE
      Do While Not .EOF
        IDOC$ = SAFETEXT$(!idenlote)
        CIXI% = XVALO(!COD_INTE)
        SALCOMPRA = SALDOCOM(IDOC$, CIXI%)
        If SALCOMPRA < XVALO(!Kilo_Asig) Then
           !Kilo_Asig = SALCOMPRA
           .Update
        End If
      .MoveNext
      Loop
   End With
   '
   BALRE.Close
   Set BALRE = Nothing
   '
End Sub
