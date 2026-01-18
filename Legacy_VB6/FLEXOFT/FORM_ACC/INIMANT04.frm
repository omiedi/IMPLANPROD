VERSION 5.00
Begin VB.Form INIMANT04 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FLEXOFT Mantenimiento Industrial"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11880
   Icon            =   "INIMANT04.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   11880
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
      Picture         =   "INIMANT04.frx":1D2A
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Ordenes de Trabajo / Mantenimiento"
      Top             =   6195
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
      Height          =   600
      Left            =   3675
      Picture         =   "INIMANT04.frx":2034
      Style           =   1  'Graphical
      TabIndex        =   29
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
      Picture         =   "INIMANT04.frx":233E
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Recepción de Proveedores Externos"
      Top             =   5985
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   2625
      Pattern         =   "*.EXK"
      TabIndex        =   27
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
      Picture         =   "INIMANT04.frx":2648
      Style           =   1  'Graphical
      TabIndex        =   25
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
      TabIndex        =   24
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
      Picture         =   "INIMANT04.frx":3342
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Ordenes de Recorrida"
      Top             =   7200
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
      Picture         =   "INIMANT04.frx":3784
      Style           =   1  'Graphical
      TabIndex        =   21
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
      Picture         =   "INIMANT04.frx":3A8E
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Gestión de Pedidos y Ordenes de Compra a Terceros."
      Top             =   5040
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
      Picture         =   "INIMANT04.frx":3D98
      Style           =   1  'Graphical
      TabIndex        =   8
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
      Picture         =   "INIMANT04.frx":40A2
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Consultas"
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
      Picture         =   "INIMANT04.frx":43AC
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Actualización y Consulta de Archivos Maestros"
      Top             =   3360
      Width           =   600
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1380
      Left            =   8295
      Picture         =   "INIMANT04.frx":46B6
      ScaleHeight     =   1380
      ScaleWidth      =   1905
      TabIndex        =   5
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
      TabIndex        =   4
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
         Picture         =   "INIMANT04.frx":BE48
         Style           =   1  'Graphical
         TabIndex        =   26
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
         Picture         =   "INIMANT04.frx":BF92
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Programación de Carga de Máquinas"
         Top             =   9030
         Visible         =   0   'False
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
         Picture         =   "INIMANT04.frx":C29C
         Style           =   1  'Graphical
         TabIndex        =   18
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
         Picture         =   "INIMANT04.frx":C5A6
         Style           =   1  'Graphical
         TabIndex        =   16
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
         Picture         =   "INIMANT04.frx":C9E8
         Style           =   1  'Graphical
         TabIndex        =   17
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
         Picture         =   "INIMANT04.frx":CCF2
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Planeamiento de Capacidad de Fábrica"
         Top             =   9735
         Visible         =   0   'False
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
         Picture         =   "INIMANT04.frx":D934
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   8325
         Visible         =   0   'False
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
         Picture         =   "INIMANT04.frx":DC3E
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Picture         =   "INIMANT04.frx":DF48
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   210
         Width           =   700
      End
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   6720
      Picture         =   "INIMANT04.frx":E092
      ScaleHeight     =   1170
      ScaleWidth      =   2115
      TabIndex        =   2
      Top             =   5250
      Width           =   2115
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1245
      Left            =   1890
      Picture         =   "INIMANT04.frx":15B8C
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
      Picture         =   "INIMANT04.frx":1D7BA
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
      Picture         =   "INIMANT04.frx":253E8
      ScaleHeight     =   1485
      ScaleWidth      =   5055
      TabIndex        =   3
      Top             =   1575
      Width           =   5055
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   4350
      Picture         =   "INIMANT04.frx":3D66A
      ScaleHeight     =   1170
      ScaleWidth      =   2010
      TabIndex        =   30
      Top             =   6405
      Width           =   2010
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Ordenes de Trabajo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   360
         TabIndex        =   31
         Top             =   350
         Width           =   1290
      End
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   8085
      Picture         =   "INIMANT04.frx":45030
      ScaleHeight     =   1170
      ScaleWidth      =   2115
      TabIndex        =   32
      Top             =   3570
      Width           =   2115
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "    Consultas   "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   195
         TabIndex        =   33
         Top             =   420
         Width           =   1590
      End
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   10920
      Y1              =   5
      Y2              =   5
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
   Begin VB.Menu MnuArch 
      Caption         =   "Archivos"
      Begin VB.Menu MnuOrt 
         Caption         =   "Ordenes de Trabajo"
      End
      Begin VB.Menu MnuOco 
         Caption         =   "Ordenes de Compra"
      End
      Begin VB.Menu php0 
         Caption         =   "-"
      End
      Begin VB.Menu MnuIfu 
         Caption         =   "Informes de Uso"
      End
      Begin VB.Menu MnuRem 
         Caption         =   "Registros de Mantenimiento"
      End
      Begin VB.Menu php1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuTer 
         Caption         =   "Terminar"
      End
   End
   Begin VB.Menu MnuGes 
      Caption         =   "Gestión"
      Begin VB.Menu MnuStk 
         Caption         =   "Stocks"
      End
      Begin VB.Menu Mnucom 
         Caption         =   "Compras"
      End
      Begin VB.Menu MnuOtr 
         Caption         =   "Ordenes de Trabajo"
      End
   End
   Begin VB.Menu MnuRep 
      Caption         =   "Reportes"
      Begin VB.Menu MnuHis 
         Caption         =   "Historial"
      End
   End
   Begin VB.Menu MnuCon 
      Caption         =   "Configuración"
      Begin VB.Menu MnuPla 
         Caption         =   "Plantas"
      End
      Begin VB.Menu MnuSPro 
         Caption         =   "Secciones"
      End
      Begin VB.Menu MnuEqu 
         Caption         =   "Equipos"
      End
      Begin VB.Menu MnuTab 
         Caption         =   "Tablas"
         Begin VB.Menu MnuDep 
            Caption         =   "Depósitos"
         End
         Begin VB.Menu MnuSec 
            Caption         =   "Sectores Productivos"
         End
      End
      Begin VB.Menu php2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCpt 
         Caption         =   "Componentes"
      End
      Begin VB.Menu MnuPro 
         Caption         =   "Proveedores"
      End
      Begin VB.Menu php4 
         Caption         =   "-"
      End
      Begin VB.Menu MnuEsp 
         Caption         =   "Especialidades"
      End
      Begin VB.Menu MnuPer 
         Caption         =   "Personal"
      End
      Begin VB.Menu php5 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPdt 
         Caption         =   "Procedimientos"
      End
   End
   Begin VB.Menu MnuAyu 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "INIMANT04"
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
   If DERACCE%("ORDEREC", "Ordenes de Recorrida") > 1 Then
      ORECORR04.Show 1
   End If
   Command10.Enabled = True
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   MARCH_MAQ04.Show 1
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   If EXISTE%("CONSTO04.EXE") > 0 Then
       Call CONECRUN("CONSTO04", "Control de Stocks")
     Else
       CONSTMAQ04.Show 1
   End If
   Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   Call COMUNI("Consultas. No Disponible por el Momento.")
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   Call CONECRUN("SOLIMA02", "Solicitud Interna de Materiales")
   Command14.Enabled = True
End Sub

Private Sub COMMAND15_CLICK()
   Command15.Enabled = False
   If DERACCE%("ORDEMAN", "Ordenes de Trabajo de Mantenimiento") > 1 Then
      ORMANTEN04.Show 1
   End If
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
       FOCOMAQ04.Show 1
   End If
   Command18.Enabled = True
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Call MENSERR(24, "Ayuda en Línea no Disponible\en Esta Localización.")
   Command2.Enabled = True
End Sub

Private Sub Command23_Click()
    Command23.Enabled = False
    MODU$ = "REMITO04"
    Call CONECRUN(MODU$, "Emision de Remito de Facturacion")
    Command23.Enabled = True
End Sub

Private Sub Command28_Click()
   Command28.Enabled = False
   If EXISTE%("BRECEP04.EXE") > 0 Then
       Call CONECRUN("BRECEP04", "Recepcion de Materiales")
     Else
       FOREMAQ04.Show 1
   End If
   Command28.Enabled = True
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
   If EXISTE%("ANACOS04.EXE") > 0 Then
     Call CONECRUN("ANACOS04", "Análisis de Costos y Presupuestación de Producto")
   Else
     Call COMUNI("Falta módulo ANACOS04")
   End If
   Command5.Enabled = True
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
   If EXISTE%("PEDIDO04.EXE") > 0 Then
     Call CONECRUN("PEDIDO04", "Pedidos de Clientes")
   Else
     Call COMUNI("Falta módulo PEDIDO04")
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
   End If
   Call GRATITFOR(Caption)
   '
   Call PONELOGO
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

Private Sub Mnucom_Click()
  '
  Call Command18_Click
  '
End Sub

Private Sub MnuCpt_Click()
  '
  AMAS_MAQ04.Show 1
  '
End Sub

Private Sub MnuDep_Click()
   '
    Call BLANARCH("!TADEPOSI")
    J& = 0
    For U& = 1 To ULTREG&("SUCURSAL")
        X$ = REGLEIDO$("SUCURSAL", U&)
        CODEPO% = Asc(Mid$(X$, 64, 1))
        If CODEPO% > 0 Then
           If CODEPO% <> 99 Then
              If J& < 62 Then
                 J& = J& + 1
                 Y$ = REGBLAN$("TADEPOSI")
                 Call REPLA(Y$, Chr$(CODEPO%), 1, 1)
                 Call REPLA(Y$, Left$(X$, 63), 2, 63)
                 Call REGAPP("!TADEPOSI", Y$)
              End If
           End If
        End If
    Next U&
    Call CIERRARCH("SUCURSAL")
    Call CIERRARCH("!TADEPOSI")
    '
    VTB% = VENTABU%("TADEPOSI")
    If VTB% < 0 Then Exit Sub
    '
    J& = 0
    For U& = 1 To ULTREG&("!TADEPOSI")
        X$ = REGLEIDO$("!TADEPOSI", U&)
        CODEPO% = Asc(Left$(X$, 1))
        If CODEPO% > 0 Then
           If CODEPO% < 99 Then
             If J& < 62 Then
               J& = J& + 1
               Y$ = REGBLAN$("SUCURSAL")
               Call REPLA(Y$, Mid$(X$, 2, 63), 1, 63)
               Call REPLA(Y$, Chr$(CODEPO%), 64, 1)
               Call GRAREG("SUCURSAL", Y$, J&)
               Y$ = REGBLAN$("ECODEP")
               Call REPLA(Y$, Left$(X$, 21), 1, 21)
               Call GRAREG("ECODEP", Y$, J&)
               Call GRAREG("TADEPOSI", X$, J&)
             End If
           End If
        End If
    Next U&
    Call SETULTREG("SUCURSAL", J&)
    Call CIERRARCH("SUCURSAL")
    Call SETULTREG("TADEPOSI", J&)
    Call CIERRARCH("TADEPOSI")
    '
    Y$ = REGBLAN$("ECODEP")
    Call REPLA(Y$, Chr$(99) + "Predeterminado      ", 1, 21)
    J& = J& + 1
    Call GRAREG("ECODEP", Y$, J&)
    Call SETULTREG("ECODEP", J&)
    Call CIERRARCH("ECODEP")
    Call CIERRARCH("!TADEPOSI")
    '
End Sub

Private Sub MnuEqu_Click()
  '
  AMAEQUI04.Show 1
  '
End Sub

Private Sub MnuEsp_Click()
  '
  Call SETGEMA04.Option1_DblClick(4)
  '
End Sub

Private Sub MnuOco_Click()
  '
  Call Command18_Click
  '
End Sub

Private Sub MnuOrt_Click()
  '
  Call COMMAND15_CLICK
  '
End Sub

Private Sub MnuOtr_Click()
  '
  Call COMMAND15_CLICK
  '
End Sub

Private Sub MnuPer_Click()
  '
  PAD_MAQ04.Show 1
  '
End Sub

Private Sub MnuPla_Click()
  '
  SETGEMA04.Option1(1).Value = True
  Call SETGEMA04.Command5_Click
  '
End Sub

Private Sub MnuPro_Click()
  '
  Call MARCH_MAQ04.Command19_Click
  '
End Sub

Private Sub MnuSec_Click()
    '
    Call BLANARCH("!SECPROD")
    J& = 0
    For U& = 1 To ULTREG&("SECPROD")
        X$ = REGLEIDO$("SECPROD", U&)
        COSECC = CVS(Left$(X$, 4))
        If COSECC >= 1 Then
            Call REGAPP("!SECPROD", X$)
        End If
    Next U&
    Call CIERRARCH("SECPROD")
    Call CIERRARCH("!SECPROD")
    '
    VTB% = VENTABU%("SECPROD")
    If VTB% < 0 Then Exit Sub
    '
    J& = 0
    For U& = 1 To ULTREG&("!SECPROD")
        X$ = REGLEIDO$("!SECPROD", U&)
        COSECC = CVS(Left$(X$, 4))
        If COSECC >= 1 Then
            J& = J& + 1
            Call GRAREG("SECPROD", X$, J&)
        End If
    Next U&
    '
    Call SETULTREG("SECPROD", J&)
    Call CIERRARCH("SECPROD")
    Call CIERRARCH("!SECPROD")
    '
End Sub

Private Sub MnuStk_Click()
  '
  Call Command12_Click
  '
End Sub

Private Sub MnuTer_Click()
  Call Command1_Click
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
