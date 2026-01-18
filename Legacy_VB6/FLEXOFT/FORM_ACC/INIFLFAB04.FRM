VERSION 5.00
Begin VB.Form INIFLFAB04 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FLEXOFT Industrial"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11880
   Icon            =   "INIFLFAB04.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   11880
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
      Left            =   4200
      Picture         =   "INIFLFAB04.frx":1D2A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Gestión de Pedidos y Ordenes de Compra a Terceros."
      Top             =   5985
      Width           =   600
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   4410
      Picture         =   "INIFLFAB04.frx":2034
      ScaleHeight     =   1170
      ScaleWidth      =   2010
      TabIndex        =   30
      Top             =   6195
      Width           =   2010
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "No Conforme"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   420
         TabIndex        =   31
         Top             =   315
         Width           =   1290
      End
   End
   Begin VB.CommandButton Command9 
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
      Height          =   600
      Left            =   9660
      Picture         =   "INIFLFAB04.frx":9CC6
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Planeamiento de Capacidad de Fábrica"
      Top             =   4305
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
      Picture         =   "INIFLFAB04.frx":A908
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Programación de Carga de Máquinas"
      Top             =   3360
      Width           =   600
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   8085
      Picture         =   "INIFLFAB04.frx":AC12
      ScaleHeight     =   1170
      ScaleWidth      =   2115
      TabIndex        =   27
      Top             =   3570
      Width           =   2115
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "C.R.P. "
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
         Left            =   600
         TabIndex        =   28
         Top             =   420
         Width           =   780
      End
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
      Picture         =   "INIFLFAB04.frx":125D8
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Sub-Sistema ISO 9000"
      Top             =   5040
      Width           =   600
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1245
      Left            =   1890
      Picture         =   "INIFLFAB04.frx":128E2
      ScaleHeight     =   1245
      ScaleWidth      =   2115
      TabIndex        =   25
      Top             =   5250
      Width           =   2115
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ISO 9000"
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
         Left            =   525
         TabIndex        =   26
         Top             =   465
         Width           =   1020
      End
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   2625
      Pattern         =   "*.EXK"
      TabIndex        =   24
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
      Left            =   9555
      Picture         =   "INIFLFAB04.frx":1A3DC
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Solicitud Interna de Materiales"
      Top             =   7140
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.PictureBox Picture9 
      Height          =   2010
      Left            =   420
      ScaleHeight     =   1950
      ScaleWidth      =   2055
      TabIndex        =   21
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
      Left            =   8925
      Picture         =   "INIFLFAB04.frx":1B0D6
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Informes de Producción"
      Top             =   6150
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   0
      Top             =   0
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
      Left            =   7350
      Picture         =   "INIFLFAB04.frx":1B518
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Ordenes de Trabajo y Fabricación"
      Top             =   5145
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
      Picture         =   "INIFLFAB04.frx":1B822
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Actualización y Consulta de Archivos Maestros"
      Top             =   3360
      Width           =   600
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1380
      Left            =   8295
      Picture         =   "INIFLFAB04.frx":1BB2C
      ScaleHeight     =   1380
      ScaleWidth      =   1905
      TabIndex        =   4
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
      TabIndex        =   3
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
         Picture         =   "INIFLFAB04.frx":232BE
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Cerrar Sesión - Cambiar Usuario / Entorno"
         Top             =   1620
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
         Picture         =   "INIFLFAB04.frx":23408
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
         Picture         =   "INIFLFAB04.frx":23712
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
         Picture         =   "INIFLFAB04.frx":23B54
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   5940
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
         Picture         =   "INIFLFAB04.frx":23E5E
         Style           =   1  'Graphical
         TabIndex        =   19
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
         Picture         =   "INIFLFAB04.frx":24168
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
         Picture         =   "INIFLFAB04.frx":24472
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   210
         Width           =   700
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   7395
      Picture         =   "INIFLFAB04.frx":245BC
      ScaleHeight     =   1170
      ScaleWidth      =   2010
      TabIndex        =   1
      Top             =   5355
      Width           =   2010
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   735
      Picture         =   "INIFLFAB04.frx":2C386
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
      Picture         =   "INIFLFAB04.frx":33FB4
      ScaleHeight     =   1485
      ScaleWidth      =   5055
      TabIndex        =   2
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
Attribute VB_Name = "INIFLFAB04"
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
   If EXISTE%("ISO9K04.EXE") > 0 Then
     Call CONECRUN("ISO9K04", "Sub-Sistema ISO 9000")
   Else
     Call COMUNI("Falta Módulo ISO9K04")
   End If
   Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   If EXISTE%("CARMAQ04.EXE") > 0 Then
     Call CONECRUN("CARMAQ04", "Cálculo de Capacidad y Secuencia de Máquina")
   Else
     Call COMUNI("Falta módulo CARMAQ04")
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
   If EXISTE%("ANOCON04.EXE") > 0 Then
     Call CONECRUN("ANOCON04", "Analisis de No Conformidades")
   Else
     Call COMUNI("Falta módulo ANOCON04")
   End If
   Command18.Enabled = True
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Call MENSERR(24, "Ayuda en Linea no Disponible\en Esta Localización.")
   Command2.Enabled = True
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
   If EXISTE%("CALCAP04.EXE") > 0 Then
     Command9.Enabled = True
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
