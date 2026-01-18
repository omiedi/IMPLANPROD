VERSION 5.00
Begin VB.Form INIFLEX04 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FLEXOFT Industrial"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11880
   Icon            =   "INIFLEX04.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   11880
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
      Left            =   3675
      Picture         =   "INIFLEX04.frx":1D2A
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Expedición - Remito de Ventas - Clientes "
      Top             =   5145
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.CommandButton Command28 
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
      Left            =   3255
      Picture         =   "INIFLEX04.frx":2034
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Recepción de Proveedores Externos"
      Top             =   5985
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   2625
      Pattern         =   "*.EXK"
      TabIndex        =   30
      Top             =   210
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Timer Timer2 
      Interval        =   10000
      Left            =   420
      Top             =   0
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
      Height          =   600
      Left            =   8190
      Picture         =   "INIFLEX04.frx":233E
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Solicitud Interna de Materiales"
      Top             =   6050
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.PictureBox Picture9 
      Height          =   2010
      Left            =   420
      ScaleHeight     =   1950
      ScaleWidth      =   2055
      TabIndex        =   27
      Top             =   210
      Visible         =   0   'False
      Width           =   2115
      Begin VB.Image Image1 
         Height          =   1590
         Left            =   0
         Top             =   0
         Width           =   2010
      End
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
      Height          =   600
      Left            =   5880
      Picture         =   "INIFLEX04.frx":3038
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Informes de Producción"
      Top             =   7200
      Visible         =   0   'False
      Width           =   600
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
      Height          =   600
      Left            =   9450
      Picture         =   "INIFLEX04.frx":347A
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Programación por Pedidos de Clentes"
      Top             =   4410
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
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
      Left            =   6615
      Picture         =   "INIFLEX04.frx":3784
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Gestión de Pedidos y Ordenes de Compra a Terceros."
      Top             =   5040
      Width           =   600
   End
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
      Left            =   4305
      Picture         =   "INIFLEX04.frx":3A8E
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Ordenes de Trabajo y Fabricación"
      Top             =   6195
      Width           =   600
   End
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
      Left            =   1785
      Picture         =   "INIFLEX04.frx":3D98
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Control y Movimientos de Stocks"
      Top             =   5040
      Width           =   600
   End
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
      Left            =   7980
      Picture         =   "INIFLEX04.frx":40A2
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Planificación de Necesidades de Materiales - Técnica MRP-II"
      Top             =   3360
      Width           =   600
   End
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
      Left            =   630
      Picture         =   "INIFLEX04.frx":43AC
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Actualización y Consulta de Archivos Maestros"
      Top             =   3360
      Width           =   600
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1380
      Left            =   8295
      Picture         =   "INIFLEX04.frx":46B6
      ScaleHeight     =   1380
      ScaleWidth      =   1905
      TabIndex        =   7
      Top             =   210
      Width           =   1905
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   160
         TabIndex        =   16
         Top             =   450
         Width           =   225
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   780
         TabIndex        =   15
         Top             =   900
         Width           =   225
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1460
         TabIndex        =   14
         Top             =   470
         Width           =   225
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   800
         TabIndex        =   13
         Top             =   60
         Width           =   225
      End
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00800000&
      Height          =   8600
      Left            =   10900
      ScaleHeight     =   8535
      ScaleWidth      =   1215
      TabIndex        =   6
      Top             =   0
      Width           =   1275
      Begin VB.CommandButton Command16 
         Caption         =   "Reset"
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
         Left            =   140
         Picture         =   "INIFLEX04.frx":BE48
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Cerrar Sesión - Cambiar Usuario / Entorno"
         Top             =   1620
         Width           =   700
      End
      Begin VB.CommandButton Command7 
         Caption         =   "C.R.P."
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
         Height          =   710
         Left            =   140
         Picture         =   "INIFLEX04.frx":BF92
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Programación de Carga de Máquinas"
         Top             =   3510
         Width           =   700
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Util."
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
         Left            =   140
         Picture         =   "INIFLEX04.frx":C29C
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   7350
         Width           =   700
      End
      Begin VB.CommandButton Command4 
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
         Height          =   700
         Left            =   140
         Picture         =   "INIFLEX04.frx":C5A6
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   6640
         Width           =   700
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
         Height          =   700
         Left            =   140
         Picture         =   "INIFLEX04.frx":C9E8
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   5940
         Width           =   700
      End
      Begin VB.CommandButton Command9 
         Caption         =   "P-Cap"
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
         Height          =   710
         Left            =   140
         Picture         =   "INIFLEX04.frx":CCF2
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Planeamiento de Capacidad de Fábrica"
         Top             =   4215
         Width           =   700
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Costos"
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
         Left            =   140
         Picture         =   "INIFLEX04.frx":D934
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   2805
         Width           =   700
      End
      Begin VB.CommandButton Command2 
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
         Left            =   140
         Picture         =   "INIFLEX04.frx":DC3E
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   920
         Width           =   700
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
         Height          =   700
         Left            =   140
         Picture         =   "INIFLEX04.frx":DF48
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   210
         Width           =   700
      End
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   6720
      Picture         =   "INIFLEX04.frx":E092
      ScaleHeight     =   1170
      ScaleWidth      =   2115
      TabIndex        =   4
      Top             =   5250
      Width           =   2115
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   8085
      Picture         =   "INIFLEX04.frx":15B8C
      ScaleHeight     =   1170
      ScaleWidth      =   2115
      TabIndex        =   3
      Top             =   3570
      Width           =   2115
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   4350
      Picture         =   "INIFLEX04.frx":1D81E
      ScaleHeight     =   1170
      ScaleWidth      =   2010
      TabIndex        =   2
      Top             =   6405
      Width           =   2010
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1245
      Left            =   1890
      Picture         =   "INIFLEX04.frx":255E8
      ScaleHeight     =   1245
      ScaleWidth      =   2115
      TabIndex        =   1
      Top             =   5250
      Width           =   2115
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   735
      Picture         =   "INIFLEX04.frx":2D216
      ScaleHeight     =   1170
      ScaleWidth      =   2010
      TabIndex        =   0
      Top             =   3570
      Width           =   2010
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1485
      Left            =   2940
      Picture         =   "INIFLEX04.frx":34E44
      ScaleHeight     =   1485
      ScaleWidth      =   5055
      TabIndex        =   5
      Top             =   1575
      Width           =   5055
   End
   Begin VB.Line Line7 
      X1              =   5250
      X2              =   9135
      Y1              =   2310
      Y2              =   4200
   End
   Begin VB.Line Line6 
      X1              =   5355
      X2              =   7770
      Y1              =   2310
      Y2              =   5565
   End
   Begin VB.Line Line5 
      X1              =   5355
      X2              =   5355
      Y1              =   2310
      Y2              =   6300
   End
   Begin VB.Line Line4 
      X1              =   3045
      X2              =   5355
      Y1              =   5565
      Y2              =   2310
   End
   Begin VB.Line Line1 
      X1              =   1680
      X2              =   5355
      Y1              =   4095
      Y2              =   2310
   End
End
Attribute VB_Name = "INIFLEX04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Call SAVEFILE(LUCOM$ + "RCONTROL.DRV", "CLOSETAB")
  Call CIERRARCH("*.*")
  Call BAJALDISCO
  If PROCACTI% < 1 Then
    Call GENBACKUP
    Call CANCELUCOM
    Call FINAL("")
  End If
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   If EXISTE%("IFABRI04.EXE") > 0 Then
     Call CONECRUN("IFABRI04", "Informes de Fabricación")
   Else
     Call COMUNI("Falta módulo IFABRI04")
   End If
   Command10.Enabled = True
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   If EXISTE%("ARCHIG04.EXE") > 0 Then
     Call CONECRUN("ARCHIG04", "Archivos Maestros")
   Else
     Call COMUNI("Falta módulo ARCHIG04")
   End If
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   If EXISTE%("CONSTO04.EXE") > 0 Then
     Call CONECRUN("CONSTO04", "Control de Stocks")
   Else
     Call COMUNI("Falta módulo CONSTO04")
   End If
   Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   If EXISTE%("PROMRP04.EXE") > 0 Then
     Call CONECRUN("PROMRP04", "Programación de Necesidades MRP")
   Else
     Call COMUNI("Falta módulo PROMRP04")
   End If
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   Call CONECRUN("SOLIMA02", "Solicitud Interna de Materiales")
   Command14.Enabled = True
End Sub

Private Sub COMMAND15_CLICK()
   Command15.Enabled = False
'   If EXISTE%("ORFABR04.EXE") > 0 Then
'     Call CONECRUN("ORFABR04", "Sistema de Ordenes de Fabricación")
'   Else
'     Call COMUNI("Falta módulo ORFABR04")
'   End If
  '\\\OT-05-0415-WAR-22/06/05///
   If EXISTE%("ORCONF04.EXE") > 0 Then
     Call CONECRUN("ORCONF04", "Sistema de Ordenes de Fabricación y Envios de Prendas")
    ElseIf EXISTE%("ORFABR04.EXE") > 0 Then
       Call CONECRUN("ORFABR04", "Sistema de Ordenes de Fabricación")
      Else
       Call COMUNI("Falta módulo ORFABR04")
   End If
   '\\\OT-05-0415-WAR-FIN///
   Command15.Enabled = True
End Sub

Private Sub Command16_Click()
   Command16.Enabled = False
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
   If PROCACTI% < 1 Then
     Call CANCELUCOM
99   ' conecta con FLOGON.EXE
     On Error Resume Next
     RETOR$ = TRIM$(UCase$(App.EXEName))
     Shell "FLOGON.EXE " + RETOR$, vbNormalFocus
     If Err Then
        Call MERRFINAL("Falta Utilitario 'FLOGON.EXE'.")
     End If
     Call FINAL("")
   End If
   '
   Command16.Enabled = True
End Sub


Private Sub Command18_Click()
   Command18.Enabled = False
   If EXISTE%("ORCOMP04.EXE") > 0 Then
     Call CONECRUN("ORCOMP04", "Sistema de Ordenes de Compra")
   Else
     Call COMUNI("Falta módulo ORCOMP04")
   End If
   Command18.Enabled = True
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Call MENSERR(24, "Ayuda en Linea no Disponible\en Esta Localización.")
   Command2.Enabled = True
End Sub

Private Sub Command23_Click()
    Command23.Enabled = False
    MODU$ = "REMITO04"
    Call CONECRUN(MODU$, "Emision de Remito de Facturacion")
    Command23.Enabled = True
End Sub

Private Sub Command28_Click()
   Command28.Visible = False
   ACONEC$ = "BRECEP04"
   Call CONECRUN(ACONEC$, "Recepcion de Materiales")
   Command28.Visible = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   If EXISTE%("FLGEST04.EXE") > 0 Then
     OPX% = COMALTER%("Opciones de Conexión a Accesos Directos:\\FLEXOFT Gestión Administrativa\Otros Accesos Directos")
     If OPX% = 1 Then
         Call CONECRUN("FLGEST04", "Sistema de Gestión Administrativa")
       ElseIf OPX% = 2 Then
         GoTo 90
     End If
     GoTo 99
   End If
   '
90 Call ACCDIR("FLEXO00")
   '
99 Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call CONECRUN("FSETUP04", "Configuración de Funcionamiento")
   Call FINAL("")
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
If EXISTE%("PREACH04.EXE") > 0 Then
     Call CONECRUN("PREACH04", "Presupuesto ACHERNAR")
     GoTo 99
End If
   If EXISTE%("ANACOS04.EXE") > 0 Then
     Call CONECRUN("ANACOS04", "Análisis de Costos y Presupuestación de Producto")
   Else
     Call COMUNI("Falta módulo ANACOS04")
   End If
99 Command5.Enabled = True
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   If EXISTE%("CARMAQ04.EXE") > 0 Then
     Call CONECRUN("CARMAQ04", "Cálculo de Capacidad y Secuencia de Máquina")
   Else
     Call COMUNI("Falta módulo CARMAQ04")
   End If
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   OPX% = 2
   If EXISTE%("PRESUP04.EXE") > 0 Then
     OPX% = ALTERNA%("Generación de Presupuestos\Ordenes de Venta")
   End If
   If OPX% = 1 Then
        Call CONECRUN("PRESUP04", "Presupuestos a Clientes")
      '16/03/07(OT-01-0114)
      'PREGUNTO SI LA OPCION ELEGIDA POR EL USUARIO ES 'PEDIDOS DE CLIENTES'
      'SI ELIGE CANCELAR O HACE CLIC EN LA X, ME MUESTRA NUEVAMENTE EL MENU
      ElseIf OPX% = 2 Then
      '16/03/07(FIN)
        Call CONECRUN("PEDIDO04", "Pedidos de Clientes")
   End If
   Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   If EXISTE%("CALCAP04.EXE") > 0 Then
     Call CONECRUN("CALCAP04", "Cálculo de Capacidad")
   Else
     Call COMUNI("Falta módulo CALCAP04")
   End If
   Command9.Enabled = True
End Sub

Private Sub Form_Load()
   '
   CCOLO% = 192
   BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   Picture1.BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   Picture2.BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   Picture3.BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   Picture4.BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   Picture5.BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   Picture6.BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   Picture8.BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   '
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
10   Caption = Caption + "  -  " + EPAC$
     '  APPX$ = App.EXEName
     '  If EXISTE%(APPX$ + ".KKK") > 0 Then GoTo 10
     '  FMDX = FECHMODU%
     '  BB1 = HOY(HOS$)
     '  If DIENTRE%(FMDX, HOY(HOS$)) > 100 Then
     '    Call SAVEFILE(APPX$ + ".KKK", String$(128, 0))
     '  End If
     '
   End If
   Call GRATITFOR(Caption)
   '
   Call PONELOGO
   '
   If EXISTE%("IFABRI04.EXE") > 0 Then
      Command10.Visible = True
   End If
   '
   If EXISTE%("SOLIMA02.EXE") > 0 Then
     Command14.Visible = True
   End If
   '
   If EXISTE%("ANACOS04.EXE") > 0 Then
     Command5.Enabled = True
   End If
   '
   If EXISTE%("CARMAQ04.EXE") > 0 Then
     Command7.Enabled = True
   End If
   '
   If EXISTE%("CALCAP04.EXE") > 0 Then
     Command9.Enabled = True
   End If
   '
   If EXISTE%("BRECEP04.EXE") > 0 Then
     Command28.Visible = True
   End If
   '
   If EXISTE%("REMITO04.EXE") > 0 Then
     Command23.Visible = True
   End If
   '
   If EXISTE%("PEDIDO04.EXE") > 0 Then
      Command8.Visible = True
      If EXISTE%("PRESUP04.EXE") > 0 Then
        Command8.ToolTipText = "Acceso a Presupuestos y Ordenes de Venta"
      End If
   End If
   '
End Sub

Sub PONELOGO()
   '
   RUTARGRA$ = TRIM$(CONTROL$("", "RUTARGRA"))
   If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
   If Right$(RUTARGRA$, 1) <> "\" Then
      RUTARGRA$ = RUTARGRA$ + "\"
   End If
   '
   ISOLOGO$ = TRIM$(CONTROL$("", "ISOLOGO"))
   If ISOLOGO$ <> "" Then
      ARCHIMAG$ = RUTARGRA$ + ISOLOGO$
      If EXISTE%(ARCHIMAG$) = 1 Then
        Image1.Picture = LoadPicture("")
        On Error Resume Next
        Image1.Picture = LoadPicture(ARCHIMAG$)
        If Err Then
          ARXX$ = REPLACAR$(ARCHIMAG$, "\", "/")
          Call MENSERR(24, "Archivo Imagen Isologo\" + ARXX$ + "\No Válido o Inexistente.")
        End If
        On Error GoTo 0
        ' Call PRINTIMAGE("ISOLOGO", ARCHIMAG$)
        If EXISTE%(LUDAT$ + "LOGOEMPR.BMP") < 1 Then ' PARA CASO GENTILI
           Picture9.Visible = True
           IMAGEFILE$ = ARCHIMAG$     ' LUDAT$ + "LOGOEMPR.BMP"
           '
           FORMALTER.Image3.Picture = LoadPicture("")
           On Error Resume Next
           FORMALTER.Image3.Picture = LoadPicture(IMAGEFILE$)
           On Error GoTo 0
           GoTo 12
        End If
        '
      End If
   End If
   '
   Picture9.Visible = False
   If EXISTE%(LUDAT$ + "LOGOEMPR.BMP") > 0 Then
     Picture9.Visible = True
     IMAGEFILE$ = LUDAT$ + "LOGOEMPR.BMP"
     '
     FORMALTER.Image3.Picture = LoadPicture("")
     FORMALTER.Image3.Picture = LoadPicture(IMAGEFILE$)
     '
12   HIMAGEN = FORMALTER.Image3.Picture.Width
     VIMAGEN = FORMALTER.Image3.Picture.Height
     '
     HCUADRO = Picture9.Width
     VCUADRO = Picture9.Height
     '
     REDUH = HIMAGEN / HCUADRO
     REDUV = VIMAGEN / VCUADRO
     REDUC = REDUH: If REDUV > REDUH Then REDUC = REDUV
     If REDUC < 1 Then REDUC = 1
     '
     With Picture9
       .Height = VCUADRO * REDUV / REDUC
       .Width = HCUADRO * REDUH / REDUC
       .Left = .Left + (HCUADRO - Picture9.Width) / 2
         If .Left < 300 Then Left = 300
       .Top = .Top - (VCUADRO - Picture9.Height) / 2
         If .Top < 400 Then .Top = 400
     End With
     '
     Image1.Width = Picture9.Width
     Image1.Height = Picture9.Height
     Image1.Stretch = True
     On Error Resume Next
     Image1.Picture = LoadPicture(IMAGEFILE) ' FORMALTER.Image1.Picture ' , PH1 + BORIZQ, PV1 + BORSUP, ANCHCUA, ALTUCUA
     On Error GoTo 0
     '
     FORMALTER.Image3.Picture = LoadPicture("")
   End If
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
   Call SAVEFILE(LUCOM$ + "RCONTROL.DRV", "CLOSETAB")
   Call CIERRARCH("*.*")
   Cancel = 0
   Call FINAL("")
End Sub

Private Sub Timer1_Timer()
   Static LACTIVA%
   For KKI% = 0 To 3
     Label1(KKI%).FontSize = 12
     Label1(KKI%).FontBold = False
     Label1(KKI%).ForeColor = RGB(0, 0, 0)
     If KKI% = LACTIVA% Then
       Label1(KKI%).FontSize = 12
       Label1(KKI%).FontBold = True
       Label1(KKI%).ForeColor = RGB(255, 0, 0)
     End If
   Next KKI%
   DoEvents
   LACTIVA% = (LACTIVA% + 1) Mod 4
End Sub

Private Sub Timer2_Timer()
Exit Sub
   If EXISTE%("C:\FLEXOFT\PRUEBA.PDF") > 0 Then
     ACROL$ = FILEFIND$("C:", "ACRORD32.EXE", 0)
     If ACROL$ <> "" Then
       RUTOR$ = ACROL$ + " "
       TEXACT$ = "" '  "Adobe"
       RUTAR$ = "C:\FLEXOFT\"
       ANAME$ = RUTOR$ + RUTAR$ + "PRUEBA.PDF"
       RETURNVALUE = Shell(ANAME$, 1)
18     On Error Resume Next
       AppActivate TEXACT$ + " Reader - [PRUEBA.pdf]" ' Adobe Reader"
       'If Err Then
         'Resume 20
20       'On Error GoTo 0
         'TINI = Timer
         While Timer < TINI + 1: DoEvents: Wend
         'GoTo 18
       'End If
       On Error GoTo 0
       '
     End If
   End If
End Sub

Function PROCACTI%()
   '
   File1.Path = LUCOM$
   File1.Pattern = "*.EXK"
   File1.Refresh
   '
   For U& = 1 To File1.ListCount
      X$ = TRIM$(UCase$(File1.List(U& - 1)))
      If X$ <> "" Then
        If X$ <> App.EXEName + ".EXK" Then
          If FileLen(LUCOM$ + X$) > 0 Then
            On Error GoTo 10
            Kill LUCOM$ + X$
            On Error GoTo 0
          End If
        End If
      End If
   Next U&
   PROCACTI% = 0
   Exit Function
   '
10 Resume 11
11 On Error GoTo 0
   PROCACTI% = 1
   X1$ = Left$(X$, Len(X$) - 1) + "E"
   TTXX$ = "Imposible Cerrar Normalmente la Sesión de Trabajo.\   \"
   TTXX$ = TTXX$ + "Se Encuentra Abierto y Ejecutando el Proceso\"
   TTXX$ = TTXX$ + "'" + X1$ + "'\  \"
   TTXX$ = TTXX$ + "Cierre este Proceso Antes de Cerrar Sesión."
   Call COMUNI(TTXX$)
   '
End Function
