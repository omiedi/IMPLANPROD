VERSION 5.00
Begin VB.Form GENDESMA09 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0F8&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Escritorio de Despacho"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4890
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "GENDESMA09.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   4890
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0E0F8&
      Caption         =   "Remitos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   120
      TabIndex        =   18
      Top             =   1380
      Width           =   3690
      Begin VB.Frame Frame8 
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
         TabIndex        =   22
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command36 
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
         Begin VB.CommandButton Command37 
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
            TabIndex        =   27
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command38 
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
            TabIndex        =   26
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command39 
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
            TabIndex        =   25
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command40 
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
            TabIndex        =   24
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command42 
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
            TabIndex        =   23
            Top             =   420
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   780
         ScaleHeight     =   585
         ScaleWidth      =   2340
         TabIndex        =   19
         Top             =   840
         Width           =   2400
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
            Picture         =   "GENDESMA09.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Ver y Re-Imprimir Remito"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Re-Imprimir Remito"
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
            TabIndex        =   21
            Top             =   90
            Width           =   1440
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   600
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   29
         Top             =   360
         Width           =   2430
         Begin VB.CommandButton Command71 
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
            Picture         =   "GENDESMA09.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Generar Remito"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Imprimir Remito segun O/Despacho"
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
            Left            =   680
            TabIndex        =   31
            Top             =   30
            Width           =   1785
         End
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0F8&
      Caption         =   "Despacho de Materiales"
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
      TabIndex        =   10
      Top             =   120
      Width           =   3690
      Begin VB.PictureBox Picture8 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   600
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   32
         Top             =   360
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
            Height          =   600
            Left            =   0
            Picture         =   "GENDESMA09.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Orden de Despacho"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Orden de Despacho"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   855
            TabIndex        =   34
            Top             =   120
            Width           =   1665
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
         TabIndex        =   11
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
            TabIndex        =   17
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
            TabIndex        =   16
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
            TabIndex        =   15
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
            TabIndex        =   14
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
            TabIndex        =   13
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
            TabIndex        =   12
            Top             =   1890
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
      TabIndex        =   6
      Top             =   3165
      Width           =   3690
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   615
         ScaleHeight     =   585
         ScaleWidth      =   2340
         TabIndex        =   7
         Top             =   360
         Width           =   2400
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
            Picture         =   "GENDESMA09.frx":076A
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
            Top             =   0
            Width           =   615
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
            Left            =   915
            TabIndex        =   8
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
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   4
         Top             =   3870
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
         Picture         =   "GENDESMA09.frx":0A74
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
         Picture         =   "GENDESMA09.frx":0BBE
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
         Left            =   105
         Picture         =   "GENDESMA09.frx":0EC8
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Configuración de Procesos y Formularios"
         Top             =   2970
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -270
         Picture         =   "GENDESMA09.frx":130A
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   4080
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
Attribute VB_Name = "GENDESMA09"
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

'
Private Sub Command14_Click()
   '
   Command14.Enabled = False
   If DERACCE%("STPOFAB", "Configuracion de Funcionamiento de O/Compra") = 2 Then
      CFGRECEP.Show 1
   End If
   Command14.Enabled = True
   '
End Sub

Private Sub Command6_Click()
   '
   Screen.MousePointer = 11
   CMO$ = "RT" 'TRIM$(CONTROL$("", "CODREFAC"))
   If Len(CMO$) <> 2 Then
        CMO$ = "RT"
   End If
   '
   Screen.MousePointer = 1
   Call ABRESTOCKS
   '
   CONDI$ = "DOPRILAR <>''"
   Call CARGALISEL("!DESPMAT", "ENCAREMI", "DOPRILAR/A24;DESCRIPCION/A64", CONDI$)
   
   
   '
   If CantRegistros("ENCAREMI", CONDI$) < 1 Then
      Call MENSERR(24, "No se Encontraron Remitos en el Rango Seleccionado.")
      GoTo 99
   End If
15 Call SELECHO("!DESPMAT")
  
   '
   Screen.MousePointer = 1
   '
   NUDOCU$ = TRIM$(VALACT1$("!DESPMAT"))
   If NUDOCU$ = "" Then GoTo 99
   TIDOCUM$ = "Despacho de Material"
   Screen.MousePointer = 11
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
40   Call COMUNI("No se Encontró el Documento Impreso")
     GoTo 15
     '
50   Call MUESTRADOC(DOCUNU&)
     GoTo 15
     '
99   Command6.Enabled = True

End Sub



Private Sub Command8_Click()
    DESPMAT09.Show 1
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
    If Left$(APLINVO$, 7) = "REMITO" Then
       Call Command71_Click
       Call FINAL("")
    End If
    '
    If Left$(APLINVO$, 7) = "REIMPRE" Then
       Call Command6_Click
       Call FINAL("")
    End If
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

Private Sub Command71_Click()
   Command71.Enabled = False
   '
8
   CONDI$ = "DOPRILAR =''"
   Call CARGALISEL("!DESPMAT", "ENCAREMI", "DOSECLAR/A24;DESCRIPCION/A64", CONDI$)
   Call SELECHO("!DESPMAT")
   If VALACT1("!DESPMAT") <> "" Then
      OD$ = VALACT1$("!DESPMAT")
      '
      Call PreVisualizar(OD$)
      If DESMAVISPRE09.Ok = False Then GoTo 8
      '
10    NUM$ = InputBox("Ingresar Nro de Remito:", "Remito")
      If TRIM$(NUM$) = "" Then GoTo 99
      DOPRILAR$ = NroRemi(NUM$)
      If DOPRILAR$ = "" Then GoTo 10
      '
      'ACTUALIZAR CON NUMERO DE REMITO EN ENCAREMI
      CODREM$ = "RT"
      AA$ = XVALO(Mid(DOPRILAR$, 6, 4))
      BB$ = XVALO(Mid(DOPRILAR$, 12, 8))
      DoPriCor$ = "RT." & AA$ & "-" + BB$
      '
      'VALIDAR DUPLICADO
      If CantRegistros("ENCAREMI", "DOPRILAR ='" & DOPRILAR$ & "'") > 0 Then
         COMUNI ("Nro de Remito '" & DOPRILAR$ & "' ya fue Utilizado")
         GoTo 10
      End If
      '
      CONDI$ = "DOSECLAR = '" & OD$ & "'"
      Call ACTUREGISTRO("MOVISTO", "DOPRICOR", CONDI$, DoPriCor$, REG&)
      Call ACTUREGISTRO("MOVISTO", "DOPRILAR", CONDI$, DOPRILAR$, REG&)
      Call ACTUREGISTRO("ENCAREMI", "DOPRILAR", CONDI$, DOPRILAR$, REG&)
      '
      SQLX$ = "SELECT * FROM MOVISTO WHERE DOSECLAR = '" & OD$ & "'"
      '
      'SE OBTIENE EL DETALLE
      Dim DETALLE As ADODB.Recordset
      Set DETALLE = New ADODB.Recordset
      Set DETALLE = FLEXCONN.Execute(FILTSQL$(SQLX$))
      '
      'CARGAR ENCABELZADO
      TOTBUL% = XVALO(VALOBADA("ENCAREMI", "BULTOS", CONDI$))
      PROP$ = SAFETEXT(VALOBADA("ENCAREMI", "PROPIEDAD", CONDI$))
      FECHA$ = FECHATEX(CVINT(VALOBADA("ENCAREMI", "FECHMOV", CONDI$)))
      Producto$ = SAFETEXT(VALOBADA("ENCAREMI", "PRODUCTO", CONDI$))
      REMITO$ = SAFETEXT(VALOBADA("ENCAREMI", "DOPRILAR", CONDI$))
      CONCEP$ = SAFETEXT(VALOBADA("ENCAREMI", "CONCEPTO", CONDI$))
      PALET$ = SAFETEXT(VALOBADA("ENCAREMI", "CANTPALET", CONDI$))
      '
      'Se completan datos segun Destino (Laboratorio, Proveedor, Deposito)
      PROVE% = XVALO(VALOBADA("ENCAREMI", "NUME_PROV", CONDI$))
      DEPODEST% = XVALO(VALOBADA("ENCAREMI", "DEPOSITO", CONDI$))
      CLIE% = XVALO(VALOBADA("ENCAREMI", "NUME_CLIE", CONDI$))
      
      '
      If DEPODEST% > 0 Then
         RAZON$ = ECOARCH("TADEPOSI", Chr(SAFETEXT(DEPODEST%)))
      ElseIf PROVE% > 0 Then
         'Si DESTINO ES PROVEEDOR SE OBTIENEN DATOS del Proveedor
         RAZON$ = RASOCLI$(PROVE% * (-1))
         DOMI$ = DOMICLI$(PROVE% * (-1))
         CPOSTAL$ = CPOSCLI$(PROVE% * (-1))
         LOCLI$ = LOCACLI$(PROVE% * (-1))
         CUIT$ = CUITCLI$(PROVE% * (-1))
         CPAG$ = CPAGCLI%(PROVE% * (-1))
         CONDPAG$ = ECOARCH("CONPAG", Chr(TRIM(CPAG$)))
      ElseIf CLIE% > 0 Then
         RAZON$ = RASOCLI$(CLIE%)
         DOMI$ = DOMICLI$(CLIE%)
         CPOSTAL$ = CPOSCLI$(CLIE%)
         LOCLI$ = LOCACLI$(CLIE%)
         CUIT$ = CUITCLI$(CLIE%)
         CPAG$ = CPAGCLI%(CLIE%)
         CONDPAG$ = ECOARCH("CONPAG", Chr(TRIM(CPAG$)))
      End If
      '
      CODPARAM$(1) = "RASO-CLI": DOCPARAM$(1) = RAZON$
      CODPARAM$(2) = "DOMI-CLI": DOCPARAM$(2) = DOMI$
      CODPARAM$(3) = "CPOS-CLI": DOCPARAM$(3) = CPOSTAL$
      CODPARAM$(4) = "LOCA-CLI": DOCPARAM$(4) = LOCLI$    'localidad
      CODPARAM$(5) = "NUME-COM": DOCPARAM$(5) = REMITO$
      CODPARAM$(6) = "REF-MAT1": DOCPARAM$(6) = Producto$     'PRODUCTO
      CODPARAM$(7) = "CUIT-CLI": DOCPARAM$(7) = CUIT$
      CODPARAM$(8) = "REF-MAT2": DOCPARAM$(8) = TOTBUL%     'Deposito destino
      CODPARAM$(9) = "FECH-EMI": DOCPARAM$(9) = FECHA$
      CODPARAM$(10) = "COND-PAG": DOCPARAM$(10) = CONDPAG$
      CODPARAM$(11) = "REF-MAT4": DOCPARAM$(11) = CONCEP$
      CODPARAM$(12) = "REF-MAT3": DOCPARAM$(12) = PROP$
      CODPARAM$(13) = "REF-MAT5": DOCPARAM$(13) = PALET$
      '
      CAPARDOC% = 13
      '
      CODTABU1$(1) = "COD-MAT"
      CODTABU1$(2) = "DES-MAT"
      CODTABU1$(3) = "NRO-LOTE"
      CODTABU1$(4) = "UNIMED"
      CODTABU1$(5) = "CANTIDAD"
      CACOLTAB1% = 5
      '
      'CARGAR DETALLE
      CAITAB1% = 0
      With DETALLE
       Do While Not (.EOF)
       '
       CIXI% = XVALO(!Cod_Inte)
       Codigo$ = CODEXT(CIXI%)
       LOTE$ = SAFETEXT(!IdenLote)
       UNIMEDIDA$ = UNIMED$(CIXI%)
       Cantidad# = XVALO(!Cantsalid)
       '
       CAITAB1% = CAITAB1% + 1
       TETABU1$(1, CAITAB1%) = Codigo$
       TETABU1$(2, CAITAB1%) = DESCRIT$(CIXI%)
       TETABU1$(3, CAITAB1%) = LOTE$
       TETABU1$(4, CAITAB1%) = UNIMEDIDA$
       TETABU1$(5, CAITAB1%) = Cantidad#
       .MoveNext
      Loop
      End With
      'IMPRIMIR FORMULARIO
      TIPODOC$ = "Despacho de Material"
      NUMEDOC$ = DOPRILAR$
      Call PRINTDOC("@RTT1-BLP")
   End If
   '
   'ASIGNO MAIL DE DESPACHO Y LO VALIDO
   ' ENVIO DE MAIL SEGUN SI EL DEPOSITO DESTINO TIENE ASIGNADO UN MAIL
   
   NUMEDEP$ = TRIM$(Str$(DEPODEST%))
   EMAI$ = TRIM$(CONTROL("DIREMAIL", "DEPO-" & NUMEDEP$))
   If EMAI$ <> "" And InStr(EMAI$, "@") > 1 Then
       SENDBYMAIL$ = EMAI$
       '
       Call PRINTDOC("@RTT1-BLP")
       '
       SENDBYMAIL$ = ""
       '
   End If
   '
   
99 Screen.MousePointer = 1
   Command71.Enabled = True
End Sub
Function NroRemi(NROX$)
      Dim NRO$(10)
      NROY$ = NROX$
      LETRAFA$ = "X"
      TIDOCOM$ = "RT"
      While Len(NROY$) > 0
         If Asc(Left$(NROY$, 1)) < 48 Or Asc(Left$(NROY$, 1)) > 57 Then
              NROY$ = Mid$(NROY$, 2)
            Else
              GoTo 101
         End If
      Wend
      '
101   If NROY$ = "" Then NROY$ = "1-1"
      SUC% = 0
      C1% = 1: For C% = 1 To 10:  NRO$(C%) = "": Next C%
      '
      For C% = 1 To Len(NROY$)
         C2% = Asc(Mid$(NROY$, C%, 1))
         If C2% = 32 Then GoTo 102
         If (C2% < 48 Or C2% > 57) Then C1% = C1% + 1: GoTo 102
         NRO$(C1%) = NRO$(C1%) + Chr$(C2%)
102   Next C%
      '
      If C1% = 1 Then
         SUC% = 0
         VD# = XVALO(NRO$(1))
         NRO$(2) = NRO$(1)
         NRO$(1) = "0"
      End If
      '
      If XVALO(NRO$(2)) < 1 Or Len(NRO$(1)) > 4 Or Len(NRO$(2)) > 8 Then
          Call MENSERR(24, "Numero de Documento no Permitido")
          Exit Function
      End If
      '
      While Len(NRO$(1)) < 4: NRO$(1) = "0" + NRO$(1): Wend
      While Len(NRO$(2)) < 8: NRO$(2) = "0" + NRO$(2): Wend
      '
      
      PVCL% = PIVACLI%(-1 * NC%)
      NROY$ = TIDOCOM$ + "-" + LETRAFA$ + "-" + NRO$(1) + "-" + NRO$(2)
      NroRemi = NROY$
End Function

Sub PreVisualizar(OD$)
    'Visualizar la Orden de Despacho Previamente en pantalla
      CONDI$ = "DOSECLAR = '" & OD$ & "'"
      'Encabezados
      DESMAVISPRE09.LIMPIAR
      '!DoSecLar = NumODLar$
     ' !DOPRILAR = ""
      DESMAVISPRE09.Label13 = VALOBADA("ENCAREMI", "FechMov", CONDI$)
      DESMAVISPRE09.Label11 = VALOBADA("ENCAREMI", "Propiedad", CONDI$)
      DESMAVISPRE09.Label39 = VALOBADA("ENCAREMI", "Concepto", CONDI$)
      DESMAVISPRE09.Label40 = VALOBADA("ENCAREMI", "Producto", CONDI$)
      DESMAVISPRE09.Label12 = VALOBADA("ENCAREMI", "Bultos", CONDI$)
      DESMAVISPRE09.Label37 = VALOBADA("ENCAREMI", "Origen", CONDI$)
      DESMAVISPRE09.Label32 = VALOBADA("ENCAREMI", "Deposito", CONDI$)
      DESMAVISPRE09.Label33 = VALOBADA("ENCAREMI", "Nume_Clie", CONDI$)
      DESMAVISPRE09.Label34 = VALOBADA("ENCAREMI", "Nume_Prov", CONDI$)
      DESMAVISPRE09.Label19 = VALOBADA("ENCAREMI", "CantPalet", CONDI$)
      If XVALO(Label32) = 0 Then Label32 = ""
      If XVALO(Label33) = 0 Then Label33 = ""
      If XVALO(Label34) = 0 Then Label34 = ""
        
      '
      'Movimientos
      SQLX$ = "SELECT * FROM MOVISTO WHERE DOSECLAR ='" & OD$ & "'"
      SQLX$ = SQLX$ & " AND Cantsalid > 0 "
      Dim Movi As ADODB.Recordset
      Set Movi = New ADODB.Recordset
      Set Movi = FLEXCONN.Execute(FILTSQL(SQLX$))
      With Movi
        .MoveFirst
       
        Do While Not (.EOF)
          XX$ = ""
          CI% = !Cod_Inte
          LOTE$ = !IdenLote
          CAN# = XVALO(!Cantsalid)
          '
          Call REPLA(XX$, CODEXT(CI%), 1, 16)
          Call REPLA(XX$, DESCRIT0(CI%), 18, 32)
          Call REPLA(XX$, LOTE$, 51, 12)
          Call REPLA(XX$, FORMATNUM(CAN#, "F12." & CNTDC), 66, 14) 'CANTIDAD
          DESMAVISPRE09.List1.AddItem XX$
          .MoveNext
        Loop
        .Close
      End With
      DESMAVISPRE09.Show 1

End Sub
