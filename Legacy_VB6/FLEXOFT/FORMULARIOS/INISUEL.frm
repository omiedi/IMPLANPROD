VERSION 5.00
Begin VB.Form INISUEL 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema de Sueldos y Jornales"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11880
   Icon            =   "INISUEL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11880
   Begin VB.CommandButton Command10 
      Caption         =   "Command10"
      Height          =   540
      Left            =   105
      TabIndex        =   38
      Top             =   3045
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.CommandButton Command9 
      Caption         =   "CLIENTES"
      Height          =   375
      Left            =   0
      TabIndex        =   37
      Top             =   2625
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.CommandButton Command8 
      Caption         =   "BLOQUEAR"
      Height          =   375
      Left            =   0
      TabIndex        =   36
      Top             =   2205
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.ListBox List2 
      Height          =   645
      Left            =   5760
      TabIndex        =   35
      Top             =   480
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00800000&
      Height          =   8600
      Left            =   11010
      ScaleHeight     =   8535
      ScaleWidth      =   855
      TabIndex        =   5
      Top             =   0
      Width           =   915
      Begin VB.CommandButton Command7 
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   120
         TabIndex        =   27
         Top             =   3520
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   555
         TabIndex        =   25
         Top             =   5400
         Visible         =   0   'False
         Width           =   615
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   26
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Plan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   120
         Picture         =   "INISUEL.frx":1D2A
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Plan General de Cuentas"
         Top             =   5835
         Visible         =   0   'False
         Width           =   615
      End
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
         Height          =   675
         Left            =   120
         Picture         =   "INISUEL.frx":1E74
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2850
         Visible         =   0   'False
         Width           =   615
      End
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
         Height          =   675
         Left            =   120
         Picture         =   "INISUEL.frx":217E
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Consultas y Listados Varios"
         Top             =   7200
         Width           =   615
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
         Height          =   675
         Left            =   120
         Picture         =   "INISUEL.frx":2488
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Setup - Configuración de Funcionamiento"
         Top             =   6525
         Width           =   615
      End
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
         Height          =   675
         Left            =   120
         Picture         =   "INISUEL.frx":28CA
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1130
         Width           =   615
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
         Height          =   675
         Left            =   120
         Picture         =   "INISUEL.frx":2BD4
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Cierra y Abandona la Aplicacion"
         Top             =   450
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H00800000&
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   11835
      TabIndex        =   28
      Top             =   0
      Width           =   11895
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7440
         TabIndex        =   30
         Top             =   40
         Width           =   3015
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         Sorted          =   -1  'True
         TabIndex        =   29
         Top             =   45
         Width           =   3735
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Liquidación:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5760
         TabIndex        =   32
         Top             =   90
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   -360
         TabIndex        =   31
         Top             =   90
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture6 
      BorderStyle     =   0  'None
      Height          =   1485
      Left            =   1560
      Picture         =   "INISUEL.frx":2D1E
      ScaleHeight     =   1485
      ScaleWidth      =   4800
      TabIndex        =   23
      Top             =   3600
      Width           =   4800
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   0
      Top             =   480
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
      Left            =   5760
      Picture         =   "INISUEL.frx":1AFA0
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Impresión de Libros y  Listados"
      Top             =   6060
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
      Left            =   8160
      Picture         =   "INISUEL.frx":1B2AA
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Transferencia a Sistema  DGI-SIJP"
      Top             =   3600
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
      Left            =   5535
      Picture         =   "INISUEL.frx":1C8DC
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Configuración"
      Top             =   1230
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
      Left            =   1680
      Picture         =   "INISUEL.frx":1CA26
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Liquidaciones"
      Top             =   6300
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
      Left            =   1725
      Picture         =   "INISUEL.frx":1CD30
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Empleados"
      Top             =   900
      Width           =   600
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   1800
      Picture         =   "INISUEL.frx":1D03A
      ScaleHeight     =   1170
      ScaleWidth      =   2235
      TabIndex        =   4
      Top             =   6540
      Width           =   2235
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Liquidaciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   180
         TabIndex        =   19
         Top             =   450
         Width           =   1710
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   5820
      Picture         =   "INISUEL.frx":24A00
      ScaleHeight     =   1170
      ScaleWidth      =   2235
      TabIndex        =   3
      Top             =   6330
      Width           =   2235
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Impresión  "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   21
         Top             =   365
         Width           =   1590
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "de Libros "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   20
         Top             =   590
         Width           =   1590
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   8280
      Picture         =   "INISUEL.frx":2C3C6
      ScaleHeight     =   1170
      ScaleWidth      =   2220
      TabIndex        =   2
      Top             =   3885
      Width           =   2220
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "a SIST. DGI-SIJP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   18
         Top             =   555
         Width           =   1800
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Transferencia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   120
         TabIndex        =   17
         Top             =   315
         Width           =   1800
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1155
      Left            =   5640
      Picture         =   "INISUEL.frx":33D8C
      ScaleHeight     =   1155
      ScaleWidth      =   2115
      TabIndex        =   1
      Top             =   1500
      Width           =   2115
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Configuración"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   120
         TabIndex        =   16
         Top             =   450
         Width           =   1710
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   1830
      Picture         =   "INISUEL.frx":3B886
      ScaleHeight     =   1170
      ScaleWidth      =   2010
      TabIndex        =   0
      Top             =   1170
      Width           =   2010
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Empleados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   210
         TabIndex        =   22
         Top             =   450
         Width           =   1590
      End
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   1320
      Sorted          =   -1  'True
      TabIndex        =   33
      Top             =   480
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label Label13 
      Caption         =   "Label13"
      Height          =   255
      Left            =   5400
      TabIndex        =   34
      Top             =   600
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Line Line7 
      X1              =   3960
      X2              =   6720
      Y1              =   4440
      Y2              =   2040
   End
   Begin VB.Line Line6 
      X1              =   2760
      X2              =   3960
      Y1              =   1800
      Y2              =   4320
   End
   Begin VB.Line Line5 
      X1              =   3960
      X2              =   2760
      Y1              =   4320
      Y2              =   7200
   End
   Begin VB.Line Line4 
      X1              =   3840
      X2              =   9360
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line1 
      X1              =   3960
      X2              =   6720
      Y1              =   4320
      Y2              =   6960
   End
   Begin VB.Menu MnuArch 
      Caption         =   "Archivos"
      Begin VB.Menu MnuEmple 
         Caption         =   "Empleados"
         Begin VB.Menu MnuModExi 
            Caption         =   "Modificar Existente"
         End
         Begin VB.Menu MnuNueEmp 
            Caption         =   "Nuevo Empleado"
         End
         Begin VB.Menu MnuLisEmp 
            Caption         =   "Listados"
         End
         Begin VB.Menu MnuStpEmp 
            Caption         =   "Configuración"
         End
      End
      Begin VB.Menu php0 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFin 
         Caption         =   "Terminar"
      End
   End
   Begin VB.Menu MnuLiq 
      Caption         =   "Liquidaciones"
      Begin VB.Menu MnuIniLiq 
         Caption         =   "Inicialización de Liquidación"
         Begin VB.Menu MnuNueLiq 
            Caption         =   "Nueva Liquidación"
         End
         Begin VB.Menu MnuDatLiqAct 
            Caption         =   "Datos de Liquidación Activa"
         End
      End
      Begin VB.Menu MnuArcRec 
         Caption         =   "Archivado y Recuperación"
         Begin VB.Menu MnuArchiva 
            Caption         =   "Archivado"
         End
         Begin VB.Menu MnuRecupera 
            Caption         =   "Recuperación"
         End
      End
      Begin VB.Menu MnuRecal 
         Caption         =   "Recálculo de Antiguedad"
      End
      Begin VB.Menu MnuLiqOrd 
         Caption         =   "Liquidación Ordinaria"
      End
      Begin VB.Menu MnuLisInfHor 
         Caption         =   "Listado de Información de Horas"
      End
   End
   Begin VB.Menu Mnu_Qwery 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu MnuPlanHab 
         Caption         =   "Planilla de Haberes"
      End
      Begin VB.Menu MnuPlanDet 
         Caption         =   "Planillas de Detalle"
         Begin VB.Menu DetHoLiq 
            Caption         =   "Detalle de Horas Liquidadas"
         End
         Begin VB.Menu DetRemun 
            Caption         =   "Detalle de Remuneraciones"
         End
         Begin VB.Menu DetReten 
            Caption         =   "Detalle de Retenciones"
         End
         Begin VB.Menu DetAsiFam 
            Caption         =   "Detalle de Asignaciones Familiares"
         End
      End
      Begin VB.Menu ReLiCon 
         Caption         =   "Resumen de Liquidación por Concepto"
      End
      Begin VB.Menu MnuPlanCam 
         Caption         =   "Planilla de Cambio para el Banco"
      End
      Begin VB.Menu MnuRecHab 
         Caption         =   "Recibo de Haberes"
      End
      Begin VB.Menu MnuSueJor 
         Caption         =   "Libro de Sueldos y Jornales"
      End
      Begin VB.Menu PHP100 
         Caption         =   "-"
      End
      Begin VB.Menu MnuObSoc 
         Caption         =   "Listado de Obras Sociales"
         Enabled         =   0   'False
      End
      Begin VB.Menu MnuLisIn 
         Caption         =   "Listado de Sindicato"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu Config 
      Caption         =   "Configuración"
      Begin VB.Menu MnuStp 
         Caption         =   "Datos de la Empresa"
      End
      Begin VB.Menu MnuVCP 
         Caption         =   "Valores de Cambio para Pago"
      End
      Begin VB.Menu MnuCalFer 
         Caption         =   "Calendario de Feriados"
      End
      Begin VB.Menu MnuParInd 
         Caption         =   "Parámetros Individuales"
      End
      Begin VB.Menu MnuSalNom 
         Caption         =   "Salarios Nominales"
      End
   End
   Begin VB.Menu Mnu_Setup 
      Caption         =   "Herramientas"
      Begin VB.Menu Mnu_Import 
         Caption         =   "Transferencia a Sistema DGI-SIJP"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "INISUEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Contablori As Database
Dim Plancuenori As Recordset
Dim NUREPLANORI As Integer

Private Sub Asi_Caj_Click()
   If COMEJ% < 1 Then
     Call MENSERR(24, "Falta Abrir Ejercicio Contable")
     Exit Sub
   End If
   '
   Call ASIDIAR.Asi_Caj_Click
   ASIDIAR.Show 1
End Sub

Private Sub Asi_Com_Click()
   If COMEJ% < 1 Then
     Call MENSERR(24, "Falta Abrir Ejercicio Contable")
     Exit Sub
   End If
   '
   Call ASIDIAR.Asi_Com_Click
   ASIDIAR.Show 1
End Sub

Private Sub Asi_Man_Click()
   If COMEJ% < 1 Then
     Call MENSERR(24, "Falta Abrir Ejercicio Contable")
     Exit Sub
   End If
   '
   Call ASIDIAR.Asi_Man_Click
   ASIDIAR.Show 1
End Sub

Private Sub Asi_Pag_Click()
   If COMEJ% < 1 Then
     Call MENSERR(24, "Falta Abrir Ejercicio Contable")
     Exit Sub
   End If
   '
   Call ASIDIAR.Asi_Pag_Click
   ASIDIAR.Show 1
End Sub

Private Sub Asi_Ven_Click()
   If COMEJ% < 1 Then
     Call MENSERR(24, "Falta Abrir Ejercicio Contable")
     Exit Sub
   End If
   '
   Call ASIDIAR.Asi_Ven_Click
   ASIDIAR.Show 1
End Sub

Private Sub Borrempre_Click()
   Call CARMULTIEM
End Sub

Private Sub Combo1_Click()
  '
  Call CIERRARCH("*.*")
  Call BAJALDISCO
  '
  ICOMBO& = Combo1.ListIndex
  LUDA1$ = ULODATOS + TRIM$(Mid$(List1.List(ICOMBO&), 31)) + "."
  LUDA2$ = ULODATOS + TRIM$(Mid$(List1.List(ICOMBO&), 31)) + "\"
  '
  EXISZ$ = ""
  On Error Resume Next
  EXISZ$ = TRIM$(Dir(LUDA1$, 16))
  On Error GoTo 0
  '
  If EXISZ$ = "" Then
     MsgBox "No Existe Carpeta " + LUDA2$
     Call FUERASYS
  End If
  '
  EMPREAC$ = Left$(List1.List(ICOMBO&), 30)
  LUSELEC$ = TRIM$(Mid$(List1.List(ICOMBO&), 31))
  LUDAT$ = LUDA2$
  '
  EMPREAC$ = TRIM$(EMPREAC$) + " - " + VERNUM$ + " - " + USERNAME$
  EPAC$ = TRIM$(EMPREAC$)
  If EPAC$ <> "" Then
    Caption = "Sistema de Sueldos y Jornales  -  " + EPAC$
  End If
  '
  FF00$ = LOADFILE$("FLEXCFG.DAT")
  Call REPLA(FF00$, LUSELEC$, 51, 45)
  Call SAVEFILE("FLEXCFG.DAT", FF00$)
  '
  RAMSTRI$ = Left$(LOADFILE("C:\FLEXOFT\FLEXRAM.DRV") + String$(512, 0), 512)
  Call REPLA(RAMSTRI$, LUDA2$, 99, 46)
  Call REPLA(RAMSTRI$, EMPREAC$, 193, 30)
  Call SAVEFILE("C:\FLEXOFT\FLEXRAM.DRV", RAMSTRI$)
  '
  VUELTA% = 0
5 EXISZ$ = ""
  On Error Resume Next
  EXISZ$ = TRIM$(Dir(LUDAT$ + "BdAccess.", 16))
  On Error GoTo 0
  If EXISZ$ = "" Then
    If VUELTA% = 0 Then
      VUELTA% = VUELTA% + 1
      MkDir LUDAT$ + "BdAccess\"
      GoTo 5
    End If
    MsgBox "No Existe Carpeta " + LUDAT$ + "BdAccess\"
    Call FUERASYS
  End If
  '
  LUACC$ = LUDAT$ + "BdAccess\"
  '
  Call CARLISLIQ
  XX$ = REGLEIDO$("ULTILIQ", 1)
  NULI1% = CVI(Mid$(XX$, 104, 2))
  Call CARGLIQ(NULI1%) 'CARGA LAS VARIABLES PUBLICAS
  DENOLI1$ = TRIM$(Mid$(XX$, 1, 30))
  If DENOLI1$ <> "" Then
    Combo2.TEXT = AJUSTI$(DENOLI1$, 40) + FORMATNUM$(NULI1%, "I5")
  End If
  '
End Sub

Private Sub Combo2_Click()
  '
  Call Picture7_Click
  ICOMBO& = Combo2.ListIndex
  TTXX$ = Combo2.TEXT
  DENOLI1$ = TRIM$(Left$(TTXX$, 40))
  NULI1% = Val(Mid$(TTXX$, 41, 8))
  '
  Call RECULIQ(NULI1%)
  '
  Command13.SetFocus
  '
End Sub

Private Sub Combo2_LostFocus()
  '
  Combo2.SelStart = 0
  Combo2.SelLength = 0
  '
End Sub

Private Sub Command1_Click()
  Call GENBACKUP
  Call SAVEFILE(LUCOM$ + "RCONTROL.DRV", "CLOSETAB")
  Call CIERRARCH("*.*")
  Call FINAL("")
End Sub

Private Sub Command10_Click()
   'Call ABRECABAN
   With CajaBanc
     .MoveFirst
     Do While Not .EOF
       On Error Resume Next
       AA1$ = !NuAsiDiar
       If Err Then
         .Edit
         !NuAsiDiar = ""
         .Update
       End If
       On Error GoTo 0
     .MoveNext
     Loop
   End With
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   EMP_GES05.Show 1
99 Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   '
   COSUJO05.Show 1
   '
99 Command12.Enabled = True
End Sub
Private Sub Command13_Click()
  Command13.Enabled = False
  PopupMenu MnuLiq
  Command13.Enabled = True
End Sub
'Private Sub Command13_Click()
'  Command13.Enabled = False
'  '
'10 RTA% = ALTERNA%("Inicialización de Liquidación\Archivado y Recuperación\Recalculo de Antiguedad\Liquidación Ordinaria\Listado Información de Horas")
'   Select Case RTA%
'    Case 1
'15    RTA1% = ALTERNA%("Nueva Liquidación\Datos de Liquidación Activa")
'      Select Case RTA1%
'        Case 1 'NUEVA
'             If COMALTER%("¿Borra Liquidación Anterior?\\Borrar\Cancelar") = 1 Then
'               'borra ultliq.
'               Call BLANARCH("ULTILIQ")
'               OBJ$ = OBJPLA$("LIQUIDACI", VDEF$)
'               VALO$ = OBJ$
'               'SI ACEPTA GRABA CAMBIOS
'               If TRIM(VALO$) <> "" Then
'                 DESCRILIQ$ = Mid$(VALO$, 1, 30)
'                 PERIOLIQ$ = Mid$(VALO$, 31, 12)
'                 GRUP% = Asc(Mid$(VALO$, 43, 1))
'                 MESACU% = Asc(Mid$(VALO$, 44, 1))
'                 ANOACUM$ = Mid$(VALO$, 45, 2)
'                  ANOACU% = Val(ANOACUM$)
'                 MESIMPU% = Asc(Mid$(VALO$, 47, 1))
'                  ANOIMP$ = Mid$(VALO$, 48, 2)
'                 ANOIMPU% = Val(ANOIMP$)
'                 TIPOLIQ% = Asc(Mid$(VALO$, 50, 1))
'                 ASIFLI% = Asc(Mid$(VALO$, 51, 1))
'                 FEPA% = CVI(Mid$(VALO$, 52, 2))
'                 LUPA$ = Mid$(VALO$, 54, 20)
'                 PEULDE$ = Mid$(VALO$, 74, 16)
'                 FEDE% = CVI(Mid$(VALO$, 90, 2))
'                 BANCOLIQ$ = Mid$(VALO$, 92, 16)
'                 '
'                 XX$ = REGLEIDO$("ULTILIQ", 1)
'                 Call REPLA(XX$, DESCRILIQ$, 1, 30)
'                 Call REPLA(XX$, Chr$(GRUP%), 31, 1)
'                 Call REPLA(XX$, Chr$(MESACU%), 32, 1)
'                 Call REPLA(XX$, Chr$(ANOACU%), 33, 1)
'                 Call REPLA(XX$, Chr$(TIPOLIQ%), 34, 1)
'                 Call REPLA(XX$, Chr$(ASIFLI%), 35, 1)
'                 Call REPLA(XX$, MKI$(FEPA%), 36, 2)
'                 Call REPLA(XX$, LUPA$, 38, 20)
'                 Call REPLA(XX$, PEULDE$, 58, 16)
'                 Call REPLA(XX$, MKI$(FEDE%), 74, 2)
'                 Call REPLA(XX$, BANCOLIQ$, 76, 16)
'                 Call REPLA(XX$, PERIOLIQ$, 92, 12)
'                 Call REPLA(XX$, Chr$(MESIMPU%), 106, 1)
'                 Call REPLA(XX$, Chr$(ANOIMPU%), 107, 1)
'                 Call GRAREG("ULTILIQ", XX$, 1)
'                 '
'                 Command13.Enabled = True
'                 Exit Sub
'               End If
'             End If
'            GoTo 15
'        Case 2 'EDITA LIQUIDACION ACTIVA
'             XX$ = REGLEIDO$("ULTILIQ", 1)
'             VDEF$ = Space(128)
'             DESCRILIQ$ = Mid$(XX$, 1, 30)
'             GRUP% = Asc(Mid$(XX$, 31, 1))
'             MESACU% = Asc(Mid$(XX$, 32, 1))
'             ANOACU% = Asc(Mid$(XX$, 33, 1))
'             TIPOLIQ% = Asc(Mid$(XX$, 34, 1))
'             ASIFLI% = Asc(Mid$(XX$, 35, 1))
'             FEPA% = CVI(Mid$(XX$, 36, 2))
'             LUPA$ = Mid$(XX$, 38, 20)
'             PEULDE$ = Mid$(XX$, 58, 16)
'             FEDE% = CVI(Mid$(XX$, 74, 2))
'             BANCOLIQ$ = Mid$(XX$, 76, 16)
'             PERIOLIQ$ = Mid$(XX$, 92, 12)
'             MESIMPU% = Asc(Mid$(XX$, 106, 1))
'             ANOIMPU% = Asc(Mid$(XX$, 107, 1))
'             '
'             Call REPLA(VDEF$, DESCRILIQ$, 1, 30)
'             Call REPLA(VDEF$, PERIOLIQ$, 31, 12)
'             Call REPLA(VDEF$, Chr$(GRUP%), 43, 1)
'             Call REPLA(VDEF$, Chr$(MESACU%), 44, 1)
'              ANOACUM$ = CStr(ANOACU%)
'              Do While Len(ANOACUM$) < 2
'               ANOACUM$ = "0" + ANOACUM$
'              Loop
'             Call REPLA(VDEF$, ANOACUM$, 45, 2)
'             Call REPLA(VDEF$, Chr$(MESIMPU%), 47, 1)
'              ANOIMP$ = TRIM$(FORMATNUM$(ANOIMPU%, "F2.0"))
'              Do While Len(ANOIMP$) < 2
'               ANOIMP$ = "0" + ANOIMP$
'              Loop
'             Call REPLA(VDEF$, ANOIMP$, 48, 2)
'             Call REPLA(VDEF$, Chr$(TIPOLIQ%), 50, 1)
'             Call REPLA(VDEF$, Chr(ASIFLI%), 51, 1)
'             Call REPLA(VDEF$, MKI$(FEPA%), 52, 2)
'             Call REPLA(VDEF$, LUPA$, 54, 20)
'             Call REPLA(VDEF$, PEULDE$, 74, 16)
'             Call REPLA(VDEF$, MKI$(FEDE%), 90, 2)
'             Call REPLA(VDEF$, BANCOLIQ$, 92, 16)
'30           OBJ$ = OBJPLA$("LIQUIDACI", VDEF$)
'             VALO$ = OBJ$
'             'SI ACEPTA GRABA CAMBIOS
'             If TRIM(VALO$) <> "" Then
'               DESCRILIQ$ = Mid$(VALO$, 1, 30)
'               If TRIM(DESCRILIQ$) = "" Then
'                Call MENSERR(24, "Denominación de Liquidación no Válida.")
'                GoTo 30
'               End If
'               PERIOLIQ$ = Mid$(VALO$, 31, 12)
'               GRUP% = Asc(Mid$(VALO$, 43, 1))
'               MESACU% = Asc(Mid$(VALO$, 44, 1))
'                ANOACUM$ = Mid$(VALO$, 45, 2)
'               ANOACU% = Val(ANOACUM$)
'               MESIMPU% = Asc(Mid$(VALO$, 47, 1))
'                ANOIMP$ = Mid$(VALO$, 48, 2)
'               ANOIMPU% = Val(ANOIMP$)
'               TIPOLIQ% = Asc(Mid$(VALO$, 50, 1))
'               ASIFLI% = Asc(Mid$(VALO$, 51, 1))
'               FEPA% = CVI(Mid$(VALO$, 52, 2))
'               LUPA$ = Mid$(VALO$, 54, 20)
'               PEULDE$ = Mid$(VALO$, 74, 16)
'               FEDE% = CVI(Mid$(VALO$, 90, 2))
'               BANCOLIQ$ = Mid$(VALO$, 92, 16)
'               '
'               XX$ = REGLEIDO$("ULTILIQ", 1)
'               Call REPLA(XX$, DESCRILIQ$, 1, 30)
'               Call REPLA(XX$, Chr$(GRUP%), 31, 1)
'               Call REPLA(XX$, Chr$(MESACU%), 32, 1)
'               Call REPLA(XX$, Chr$(ANOACU%), 33, 1)
'               Call REPLA(XX$, Chr$(TIPOLIQ%), 34, 1)
'               Call REPLA(XX$, Chr$(ASIFLI%), 35, 1)
'               Call REPLA(XX$, MKI$(FEPA%), 36, 2)
'               Call REPLA(XX$, LUPA$, 38, 20)
'               Call REPLA(XX$, PEULDE$, 58, 16)
'               Call REPLA(XX$, MKI$(FEDE%), 74, 2)
'               Call REPLA(XX$, BANCOLIQ$, 76, 16)
'               Call REPLA(XX$, PERIOLIQ$, 92, 12)
'               Call REPLA(XX$, Chr$(MESIMPU%), 106, 1)
'               Call REPLA(XX$, Chr$(ANOIMPU%), 107, 1)
'               Call GRAREG("ULTILIQ", XX$, 1)
'               '
'             End If
'        Case 3
'        Case 4
'             REMUORD05.Show 1
'        Case Else
'        GoTo 10
'      End Select
'    Case 2
'35      RTA2% = ALTERNA%("Archivado\Recuperación")
'        Select Case RTA2%
'          Case 1
'            Call ARCHILIQ
'            GoTo 35
'          Case 2
'            Call RECUPELIQ
'            GoTo 35
'          Case Else
'            GoTo 10
'        End Select
'    Case 3
'        Call CALCUANTIGU
'    Case 4
'        'Call LIQUIORDI                    ' LIQUIDACION ORDINARIA
'        REMUORD05.Show 1
'    Case 5
'        A = 5
'   End Select
'  '
'  Command13.Enabled = True
'End Sub

Sub ARCHILIQ()
  '
  X$ = REGLEIDO$("ULTILIQ", 1)
  DENOLI$ = Mid$(X$, 1, 30)
  GRULI% = Asc(Mid$(X$, 31, 1))
    MEIMLI% = Asc(Mid$(X$, 32, 1))
    ANIMLI% = Asc(Mid$(X$, 33, 1))
  MESLI$ = TRIM$(ECOARCH("TAMESE", MKI$(MEIMLI%)))
  ANOIMP$ = TRIM$(FORMATNUM$(ANIMLI%, "F2.0"))
    Do While Len(ANOIMP$) < 2
      ANOIMP$ = "0" + ANOIMP$
    Loop
  PEACULI$ = MESLI$ + " " + ANOIMP$
  '
  VDEF$ = ""
  Call REPLA(VDEF$, DENOLI$, 1, 30)
  Call REPLA(VDEF$, Chr(GRULI%), 31, 1)
  Call REPLA(VDEF$, PEACULI$, 32, 30)
  Call OBJPLA$("DATOSLIQ", VDEF$)
  '
  OBJ1$ = VDEF$
  'GRABO LIQUIDACION
  If TRIM$(OBJ1$) <> "" Then
    DENLI$ = Mid$(OBJ1$, 1, 30)
    GRULI% = Asc(Mid$(OBJ1$, 31, 1))
    PEACULI$ = Mid$(OBJ1$, 32, 30)
    'OBTENGO NUMERO DE LIQUIDACION SIGUIENTE
    NULIQ% = PRONULIQ%
    X$ = REGLEIDO$("ULTILIQ", 1)
    MAR1% = CVI(Mid$(X$, 104, 2))
    If MAR1% > 0 Then
      NULIQ% = MAR1%
    End If
    DATLIQ$ = X$
    'GRABO EN EL ENCLIQUI
    Y$ = REGBLAN$("ENCLIQUI")
    Call REPLA(Y$, MKI$(NULIQ%), 1, 2)
    Call REPLA(Y$, DATLIQ$, 3, 126)
    NULIQUI& = Val(NULIQ%)
    Call GRAREG("ENCLIQUI", Y$, NULIQUI&)
    Call CIERRARCH("ENCLIQUI")
    'GRABA LA MARCA DE LIQUIDACION ARCHIVADA EN ULTLIQ
    Call REPLA(X$, MKI$(NULIQ%), 104, 2)
    Call GRAREG("ULTILIQ", X$, 1)
    '
    MESACU% = Asc(Mid$(Y$, 34, 1))
    ANOACU% = Asc(Mid$(Y$, 35, 1))
    MESI$ = Mid$(Str$(MESACU%), 2): While Len(MESI$) < 2: MESI$ = "0" + MESI$: Wend
    ANI$ = Mid$(Str$(ANOACU%), 2): While Len(ANI$) < 2: ANI$ = "0" + ANI$: Wend
    ARCHAC$ = "DTL" + MESI$ + ANI$ + ".DAT"
    TTXX$ = LOADFILE$(LUDAT$ + ARCHAC$)
    TTZZ$ = ""
    For KU& = 1 To Len(TTXX$) Step 24
      TTYY$ = Mid$(TTXX$, KU&, 24)
      If CVI(Left$(TTYY$, 2)) <> NULIQ% Then
        TTZZ$ = TTZZ$ + TTYY$
      End If
    Next KU&
    For KU& = 1 To ULTREG&("DETLIQUI")
      XX$ = REGLEIDO$("DETLIQUI", KU&)
      If CVI(Left$(XX$, 2)) > 0 Then
        TTYY$ = MKI$(NULIQ%) + XX$ + String$(6, 0)
        TTZZ$ = TTZZ$ + TTYY$
      End If
    Next KU&
    Call SAVEFILE(LUDAT$ + ARCHAC$, TTZZ$)
    '
  End If
  '
End Sub
Function PRONULIQ%()
  'DEVULEVEL EL PROXIMO NUMERO DE LIQUIDACION
  '
  For J& = 1 To ULTREG("ENCLIQUI")
    X$ = REGLEIDO$("ENCLIQUI", J&)
    NULI% = CVI(Mid(X$, 1, 2))
    If NUMASLI% < NULI% Then
      NUMASLI% = NULI%
    End If
  Next J&
  '
  PRONULIQ% = NUMASLI% + 1
  '
End Function
'
Sub RECUPELIQ()
  '
  Call BLANARCH("INLIQARC")
  For J& = 1 To ULTREG("ENCLIQUI")
    X$ = REGLEIDO$("ENCLIQUI", J&)
    NULI% = CVI(Mid(X$, 1, 2))
    DENLOL$ = Mid$(X$, 3, 30)
    DENOLI$ = Mid$(DENLOL$, 1, 20)
    GRULI% = Asc(Mid$(X$, 33, 1))
    GRULIQ$ = Mid$(ECOARCH$("DATGREM1", MKI$(GRULI%)), 1, 16)
        MEIMLI% = Asc(Mid$(X$, 34, 1))
    ANIMLI% = Asc(Mid$(X$, 35, 1))
    MESLI$ = TRIM$(ECOARCH("TAMESE", MKI$(MEIMLI%)))
    ANOIMP$ = TRIM$(FORMATNUM$(ANIMLI%, "F2.0"))
    Do While Len(ANOIMP$) < 2
      ANOIMP$ = "0" + ANOIMP$
    Loop
    '
    LA$ = Space$(50)
    Call REPLA(LA$, DENOLI$, 1, 20)
    Call REPLA(LA$, GRULIQ$, 21, 16)
    Call REPLA(LA$, MESLI$, 37, 10)
    Call REPLA(LA$, ANOIMP$, 47, 2)
    If MEIMLI% <= 0 Or MEIMLI% >= 13 Then
      LA$ = Space$(50)
      Call REPLA(LA$, "ANULADA", 37, 10)
    End If
    '
    LIQARC$ = Space(64)
    Call REPLA(LIQARC$, MKI$(NULI%), 1, 2)
    Call REPLA(LIQARC$, LA$, 3, 50)
    Call REGAPP("INLIQARC", LIQARC$)
    '
  Next J&
  '
20 Call SELECHO("INLIQARC")
  VDEVU$ = TRIM$(VALACT1$("INLIQARC"))
  If VDEVU$ <> "" Then
     LIQUI% = Val(VALACT1$("INLIQARC"))
     Call RECULIQ(LIQUI%)
  End If
  '
End Sub

Private Sub Command15_Click()
   Command15.Enabled = False
   Call TRANSFDGI
   Command15.Enabled = True
End Sub

Private Sub Command18_Click()
   Command18.Enabled = False
   PopupMenu Mnu_Qwery
   Command18.Enabled = True
End Sub

Private Sub Command2_Click()
  '
  Call HELPONLINE("INISUEL")
  '
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call CONECRUN("FLSETUP", "Configuración de Funcionamiento")
   Call FINAL("")
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   'AVAMRPSB.Show 1
   Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   Call MENSERR(24, "No Disponible en la Presente Versión")
   'Call GRAJOBID
   'Call BAJRAM
   'Shell LUBAS$ + "WINTABU " + "#" + JOBID$ + "#", 1
   Command6.Enabled = True
End Sub

Private Sub Command7_Click()
   '
   Command7.Enabled = False
   If COMALTER%("Está Seguro de Importar Asientos ?\\No, Cancelar\Si, Importar") <> 2 Then GoTo 99
   '
   CAREPLA& = 0
   On Error Resume Next
   Plancuen.MoveLast
   CAREPLA& = Plancuen.RecordCount
   On Error GoTo 0
   If CAREPLA& < 1 Then
     Call MENSERR(24, "Imposible Importar.\Abra y Guarde Plan General de Cuentas.\  \Luego Re-Intente Importación")
     Exit Sub
   End If
   '
   Screen.MousePointer = 11
   EJER% = (-1)
   EXTE$ = "OLD"
   '
   'Call ABREPLANCUEN
   Screen.MousePointer = 11
   Set Encasien = Contable.OpenRecordset("ENCABEASIEN", dbOpenTable)
   Set Detasien = Contable.OpenRecordset("DETAASIEN", dbOpenTable)
   '
   On Error Resume Next
   Encasien.MoveLast
   CARE% = 0
   CARE% = Encasien.RecordCount
   If CARE% > 0 Then
      While Encasien.RecordCount > 0
        Encasien.MoveLast
        Encasien.Delete
      Wend
   End If
   '
5  Detasien.MoveLast
   CARE% = 0
   CARE% = Detasien.RecordCount
   If CARE% > 0 Then
      While Detasien.RecordCount > 0
        Detasien.MoveLast
        Detasien.Delete
      Wend
   End If
   On Error GoTo 0
   '
6  Encasien.Close
   Detasien.Close
   Set Encasien = Contable.OpenRecordset("ENCABEASIEN", dbOpenTable)
   Set Detasien = Contable.OpenRecordset("DETAASIEN", dbOpenTable)
   '
10 FIN& = MAXENCDIARIO%
   ORDAS% = 0: NUIDANT$ = ""
   UDIA% = 0
   For KI% = 1 To FIN&
     U& = KI%: Call TRACE(20, U&, FIN&)
     'Call LEENCDIARIO(KI%)
     NUASI = CVS(E1$)
     NUASDEF& = CVI(E2$)
     FEASX = CVI(E3$)
       FECHAS$ = FECHATEX$(FEASX)
       NUIDAS$ = Mid$(FECHAS$, 7, 2) + Mid$(FECHAS$, 4, 2) + Mid$(FECHAS$, 1, 2)
       If NUIDAS$ <> NUIDANT$ Then ORDAS% = 0
       NUIDANT$ = NUIDAS$
       If Left$(NUIDAS$, 1) = "0" Then NUIDAS$ = Mid$(NUIDAS$, 2)
       ORDAS% = ORDAS% + 1
       ORDA$ = TRIM$(Str$(ORDAS%))
       While Len(ORDA$) < 3: ORDA$ = "0" + ORDA$: Wend
       NUIDAS$ = NUIDAS$ + "-" + ORDA$
     SUDEBE# = CVD(E4$)
     CONTA% = CVI(E6$)
     REFEAS$ = E7$
     TIASI% = Asc(Left$(E8$, 1))
     POEDIT% = 1: If (TIASI% > 0 And TIASI% < 4) Then POEDIT% = (-1)
     '
     
     With Encasien
        .AddNew
        !NUIDASIEN = NUIDAS$
        !Refeasien = REFEAS$
        !FECHASIEN = CVDAT(FECHANUM(FECHAS$))
        !OrFeAsien = ORDAS%
        !SucuAsien = 0
        !NuDfAsien = NUASDEF&
        !TIPOASIEN = TIASI%
        !StatAsien = 1  ' activo
        !EditAsien = POEDIT%
        !QPasAsien = CONTA%
        !DebeAsien = SUDEBE#
        !Observa = " "
        .Update
        .Bookmark = .LastModified
     End With
     '
     ORPASE% = 0: SUDEBE# = 0
     For K2% = UDIA% + 1 To MAXDIARIO%
       'Call LEEDIARIO(K2%)
       If CVS(E12$) > NUASI + 0.01 Then
          UDIA% = K2% - 1
          GoTo 35
       End If
       If Abs(CVS(E12$) - NUASI) < 0.005 Then
          ORPASE% = ORPASE% + 1
          CUENTAINT% = CVI(E11$)
            'COCUE$ = COCUENTA$(CUENTAINT%)
            'CUEIII% = CUEINT%(COCUE$)
            If CUEIII% < 1 Then
              VEDECUE$ = LOADFILE$(LUDAT$ + "DENOCUE." + EXTE$)
              IDECU& = 1 + 128& * (CUENTAINT% - 1)
              DECUE$ = TRIM$(Mid$(VEDECUE$, IDECU&, 100))
              Call AGRECUENTA(COCUE$, DECUE$, CUEIII%)
            End If
          IMPOPAS# = CVD(E16$)
          '
          DETAPAS$ = ""
          If K2% <= LOF(NX12%) / 32 Then
            Get #NX12%, K2%, E20$
            DETAPAS$ = TRIM$(E20$)
          End If
          'If DETAPAS$ = "" Then DETAPAS$ = DENOCUE$(CVI(E11$))
          If DETAPAS$ = "" Then DETAPAS$ = " "
          '
          IDEBE# = IMPOPAS#: IHABE# = 0
          If IMPOPAS# < 0 Then
            IDEBE# = 0: IHABE# = Abs(IMPOPAS#)
            SUDEBE# = SUDEBE# + IHABE#
          End If
          '
          AGREPLAN% = 1
          With Detasien
             .AddNew
             !NUIDASIEN = NUIDAS$
             !OrNuPase = ORPASE%
             !REFEPASE = DETAPAS$
             !FECHASIEN = CVDAT(FECHANUM(FECHAS$))
             !OrFeAsien = ORDAS%
             !SucuAsien = 0
             !NuDfAsien = NUASDEF&
             !StatPase = 1
             !CUECOINT = CUEIII%
             !CUECOEXT = COCUE$
                If !CUECOEXT = "" Then !CUECOEXT = " "
             !IDEBEPASE = IDEBE#
             !IHABEPASE = IHABE#
             .Update
             .Bookmark = .LastModified
          End With
          AGREPLAN% = 0
       End If
     Next K2%
     '
35   With Encasien
       .Edit
       !DebeAsien = SUDEBE#
       .Update
       .Bookmark = .LastModified
     End With
     '
   Next KI%
   '
   If EJER% < 0 Then
     'Call CIERRACONTA
     EJER% = 1
     EXTE$ = "DAT"
     GoTo 10
   End If
   '
   Encasien.Close
   Detasien.Close
   Screen.MousePointer = 1
   '
   Call COMUNI("Importacion Completa. Debe Ahora Abrir Manualmente\los Ejercicios Contables Importados.\  \Re-Inicie Ahora la Aplicación.")
   Call CIERRARCH("*.*")
   Call FINAL("")
   '
99 Command7.Enabled = True
   '
End Sub

Function COCUENTA$(CUEII%)
   '
   Static VECUENTA$, EXTEA$, LOREL&
   '
   If EXTE$ <> EXTEA$ Then
     LOREL& = RECLEN("CUENTAS")
     EXTEA$ = EXTE$
     VECUENTA$ = LOADFILE$(LUDAT$ + "CUENTAS." + EXTE$)
   End If
   '
   COCUENTA$ = TRIM$(Mid$(VECUENTA$, 1 + LOREL& * (CUEII% - 1), 12))
   '
End Function
'
Sub AGRECUENTA(COCUE$, DECUE$, CUEIII%)
   '
      '
20 On Error Resume Next
   Set Plancuen = Contable.OpenRecordset("PLANCUEN", dbOpenDynaset)
   If Err Then
      If COMALTER%("No se Puede Abrir 'PLANCUEN'\en Modo Exclusivo.\\Reintentar\Cancelar") <> 2 Then GoTo 20
      Close: End
   End If
   '
   CAREPLA& = 0
   On Error Resume Next
   Plancuen.MoveLast
   CAREPLA& = Plancuen.RecordCount
   On Error GoTo 0
   '
   UCUE% = 0
   With Plancuen
     For U& = 1 To CAREPLA&
       If U& = 1 Then
           .MoveFirst
         Else
           .MoveNext
       End If
       If !CODINCUE > UCUE% Then
         UCUE% = !CODINCUE
       End If
     Next U&
     CUEIII% = 1 + UCUE%
     '
     .AddNew
     !CODEXCUE = COCUE$
     !DENOCUE = DECUE$
     !CODINCUE = CUEIII%
     !STATCUE = 9   ' ANULADA
     !NIVAGRU = 4   ' ULTIMO NIVEL
     !TipoCue = "Debe"
     !POSIMPU = 0   ' NO IMPUTABLE
     .Update
   End With
   Plancuen.Close
   '
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   'Call BLOQTABLA("CONTAGEN/PLANCUEN")
   MsgBox "BLOQUEADA"
   'Call LIBERTABLA("CONTAGEN/PLANCUEN")
   Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   'Call ABRECLIEN
End Sub

Private Sub Con_sql_Click()
   Call Command6_Click
End Sub

Private Sub Copyplan_Click()
   ORICOPLAN.Show 1
   '
   LUACCORI$ = TRIM$(ORICOPLAN.LUACCORI)
   If LUACCORI$ <> "" Then
      On Error Resume Next
      Set Contablori = OpenDatabase(LUACCORI$ + "CONTAGEN.MDB")
      If Err Then
         MsgBox "No Existe " + LUACCORI$ + "CONTAGEN.MDB"
         Exit Sub
      End If
      '
      Load FPLANCUEN
      If FPLANCUEN.List1.ListCount > 0 Then
        If COMALTER%("Desea Sobre-Escribir\el Plan de Cuentas Existente ?\\No, Cancelar\Si, Sobre-Escribir") <> 2 Then
          Exit Sub
        End If
      End If
      Set Plancuenori = Contablori.OpenRecordset("PLANCUEN", dbOpenDynaset, dbReadOnly)
      NUREPLANORI% = 0
      On Error Resume Next
      Plancuenori.MoveLast
      NUREPLANORI% = Plancuenori.RecordCount
      On Error GoTo 0
      If NUREPLANORI% < 1 Then
        Call MENSERR(24, "No Hay Plan de Cuentas\en la Empresa Origen Elegida")
        Exit Sub
      End If
      '
      Call CARPLANORI
      FPLANCUEN.Show 1
      '
   End If
End Sub

Sub CARPLANORI()
   '
   Screen.MousePointer = 11
   FPLANCUEN.List1.Clear
   '
   'Call ABREPLANCUEN
   '
   FIN& = NUREPLANORI%
   With Plancuenori
     .MoveFirst
     U& = 1
     ZZ$ = Space$(256)
     '
10   COCUE$ = !CODEXCUE
     DECUE$ = !DENOCUE
     NIVEI% = !NIVAGRU
       NVX$ = Left$(String$(NIVEI%, ".") + TRIM$(Str$(NIVEI%)) + ".....", 5)
       '
     Call REPLA(ZZ$, NVX$, 1, 5)
     Call REPLA(ZZ$, COCUE$, 9, 12)
     Call REPLA(ZZ$, DECUE$, 23, 44)
     Call REPLA(ZZ$, DECUE$, 129, 128)
     '
     TICUEX$ = TRIM$(UCase$(Left$(!TipoCue, 1)))
     If TICUEX$ = "H" Then
         TICUEX$ = "Haber"
       ElseIf TICUEX$ = "I" Then
         TICUEX$ = "Ingresos"
       ElseIf TICUEX$ = "G" Then
         TICUEX$ = "Gastos"
       Else
         TICUEX$ = "Debe"
     End If
     Call REPLA(ZZ$, TICUEX$, 69, 5)
     '
     IMPUBLE$ = ""
     If !POSIMPU < 1 Then IMPUBLE$ = "No"
     Call REPLA(ZZ$, IMPUBLE$, 76, 2)
     '
     STATUX$ = ""
     If !STATCUE = 9 Then STATUX$ = "Baja"
     Call REPLA(ZZ$, STATUX$, 80, 4)
     '
     NUINT$ = AJUSTD$(Str$(!CODINCUE), 8)
     Call REPLA(ZZ$, NUINT$, 121, 8)
     '
     FPLANCUEN.List1.AddItem ZZ$
     '
     If U& < FIN& Then
       .MoveNext
       U& = U& + 1
       GoTo 10
     End If
     '
   End With
   '
90 FPLANCUEN.List1.Refresh
   Screen.MousePointer = 1
   '
End Sub
'

Private Sub Editplan_Click()
   Call Command3_Click
End Sub

Private Sub Estcuen_Click()
   OPBALAN04.Option1(2).Value = 1
   On Error Resume Next
   OPBALAN04.Command1.SetFocus
   On Error GoTo 0
   OPBALAN04.Show 1
End Sub

Private Sub Exit_Click()
   Call Command1_Click
End Sub

Private Sub Exit1_Click()
   Call Command1_Click
End Sub

Private Sub DetAsiFam_Click()
   ' DETALLE DE ASIGNACIONES FAMILIARES LIQUIDADAS
   Call CALDEASI
   '
   Call FINAL("*DETASLI")
   '
End Sub

Private Sub DetHoLiq_Click()
   ' DETALLE DE HORAS LIQUIDADAS
   Dim DECO$(100), HORA%(100)
   JJ& = 0: CACON% = 0
   AA11$ = LOADFILE(LUDAT$ + "CONCEP.DAT")
   For KKU& = 1 To Len(AA11$) Step 128
     aa2$ = Mid$(AA11$, KKU&, 128)
     ICOI% = ICOI% + 1
     If ICOI% <= 100 Then
       DECO$(ICOI%) = TRIM$(Mid$(aa2$, 1, 30))
       HORA%(ICOI%) = Asc(Mid$(aa2$, 33, 1))
       CACON% = ICOI%
     End If
   Next KKU&
   '
   Call FILESORT("DETLIQUI", "DETLIQJ", 1, 2, "ASC")
   '
   AA11$ = LOADFILE(LUDAT$ + "DETLIQJ.DAT")
   For KKU& = 1 To Len(AA11$) Step 16
     aa2$ = Mid$(AA11$, KKU&, 16)
     LEGAX% = CVI(Left$(aa2$, 2))
     If LEGAX% > 0 Then
       ICOI% = CVI(Mid$(aa2$, 3, 2))
       If ICOI% < 200 Then
         QTID = CVS(Mid$(aa2$, 5, 4))
         QTHOR = 0
         If Abs(QTID) >= 0.005 Then
           If HORA%(ICOI% Mod 100) = 2 Then
             QTHOR = QTID
             QTID = 0
           End If
           YY$ = REBLA
           Call REPLA(YY$, MKI$(LEGAX%), 1, 2)
           Call REPLA(YY$, MKI$(ICOI%), 3, 2)
           Call REPLA(YY$, MKS$(QTID), 5, 4)
           Call REPLA(YY$, MKS$(QTHOR), 9, 4)
           JJ& = JJ& + 1
           Call GRAREG("DETHOLIQ", YY$, JJ&)
         End If
       End If
     End If
   Next KKU&
   Call SETULTREG("DETHOLIQ", JJ&)
   Call HEADERS("DETHOLIQ", "")
   Call HEADERS("DETHOLIQ", "Liquidación: " + DESCRI_LIQ$)
   '
   Call FINAL("*DETHOLI")
   '
End Sub

Private Sub DetRemun_Click()
   ' DETALLE DE REMUNERACIONES LIQUIDADAS
   Call CALDEREM
   '
   Call FINAL("*DETRELI")
   '
End Sub

Private Sub DetReten_Click()
   ' DETALLE DE RETENCIONES LIQUIDADAS
   Call CALDERET
   '
   Call FINAL("*DETRTLI")
   '
End Sub

Private Sub Form_Load()
   '
   CCOLO% = 193
   BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   Picture1.BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   Picture2.BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   Picture3.BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   Picture4.BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   Picture5.BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   Picture6.BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   '
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   Call CENTRAFORM(Me)
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
     Caption = Caption + "  -  " + EPAC$
   End If
   Call GRATITFOR(Caption)
   '
   Call CARLISEMP
   RAMSTRI$ = Left$(LOADFILE("C:\FLEXOFT\FLEXRAM.DRV") + String$(512, 0), 512)
   LUDAT$ = TRIM$(Mid$(RAMSTRI$, 99, 46))
   Label13 = LUDAT$
   Combo1.TEXT = TRIM$(Mid$(RAMSTRI$, 193, 30))
   '
   VUELTA% = 0
5  EXISZ$ = ""
   On Error Resume Next
   EXISZ$ = TRIM$(Dir(LUDAT$ + "BdAccess.", 16))
   On Error GoTo 0
   If EXISZ$ = "" Then
     If VUELTA% = 0 Then
       VUELTA% = VUELTA% + 1
       MkDir LUDAT$ + "BdAccess\"
       GoTo 5
     End If
     MsgBox "No Existe Carpeta " + LUDAT$ + "BdAccess\"
     Call FUERASYS
   End If
   '
   LUACC$ = LUDAT$ + "BdAccess\"
   '
   Call CARLISLIQ
   XX$ = REGLEIDO$("ULTILIQ", 1)
   NULI1% = CVI(Mid$(XX$, 104, 2))
   Call CARGLIQ(NULI1%) 'CARGA LAS VARIABLES PUBLICAS
   DENOLI1$ = TRIM$(Mid$(XX$, 1, 30))
   If DENOLI1$ <> "" Then
     Combo2.TEXT = AJUSTI$(DENOLI1$, 40) + FORMATNUM$(NULI1%, "I5")
   End If
   '
End Sub

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

Sub CARLISEMP()
   '
   Combo1.Clear: List1.Clear
   If EXISTE%(CurDir + "\FLEXMUL.DRV") = 1 Then
     LIEMP$ = LOADFILE$(CurDir + "\FLEXMUL.DRV")
     Dim FMX As String * 79
     '
     For I% = 1 To Len(LIEMP$) Step 79
       FMX$ = Mid$(LIEMP$, I%, 79)
       SUCAR% = 0
       For J% = 1 To 77
         SUCAR% = SUCAR% + Asc(Mid$(FMX$, J%, 1))
       Next J%
       If CVI(Mid$(FMX$, 78, 2)) = SUCAR% Then
         RASO$ = DESENCRI$(Left$(FMX$, 30))
         PACCE$ = DESENCRI$(Mid$(FMX$, 32, 46))
         Z$ = Space$(96)
         Call REPLA(Z$, RASO$, 1, 30)
         Call REPLA(Z$, PACCE$, 51, 46)
         Combo1.AddItem TRIM$(Left$(Z$, 30))
         List1.AddItem Z$
       End If
     Next I%
     '
   End If
   '
End Sub

Sub CARLISLIQ()
   '
   Combo2.Clear
   For U& = 1 To ULTREG&("ENCLIQUI")
     XX$ = REGLEIDO$("ENCLIQUI", U&)
     NULI1% = CVI(Mid$(XX$, 1, 2))
     DENOLI1$ = TRIM$(Mid$(XX$, 3, 30))
     If NULI1% > 0 Then
       If DENOLI1$ <> "" Then
         Combo2.AddItem AJUSTI$(DENOLI1$, 40) + FORMATNUM$(NULI1%, "I5")
       End If
     End If
   Next U&
   '
End Sub


Private Sub Abrejer_Click()
   '
   If DERACCE%("ABREJER", "Apertura de Ejercicio Contable") = 2 Then
     '
     VDEF$ = String$(80, 0)
     Call REPLA(VDEF$, MKI$(COMEJ%), 1, 2)
     Call REPLA(VDEF$, DENOEJ$, 3, 32)
     FEX = FINEJ%
     CMJ1% = DIPOST(FEX)
     Call REPLA(VDEF$, MKI$(CMJ1%), 41, 2)
     Call REPLA(VDEF$, "", 43, 32)
     '
10   OBX$ = OBJPLA$("ABREJER", VDEF$)
     If OBX$ = "" Then Exit Sub
     '
     VDEF$ = OBX$
     DENUE$ = TRIM$(Mid$(OBX$, 43, 32))
     If DENUE$ = "" Then
       Call MENSERR(24, "Falta Denominación del Nuevo Ejercicio")
       GoTo 10
     End If
     '
     CMJACT% = FINEJ%
     APNUE% = CVI(Mid$(OBX$, 41, 2))
     'If APNUE% < CMJACT% Then
     '  Call MENSERR(24, "Nueva Apertura es Anterior\al Cierre del Ejercicio Actual")
     '  GoTo 10
     'End If
     '
     FEX = APNUE%
     APNU$ = FECHATEX$(FEX)
     FEAPNU$ = Left$(APNU$, 6) + "20" + Mid$(APNU$, 7)
     If Val(Mid$(APNU$, 7)) > 80 Then FEAPNU$ = Left$(APNU$, 6) + "19" + Mid$(APNU$, 7)
     ANX% = Val(Mid$(APNU$, 7, 2)) + 1: If ANX% > 99 Then ANX% = ANX% - 100
     AN$ = TRIM$(Str$(ANX%)): While Len(AN$) < 2: AN$ = "0" + AN$: Wend
     FE$ = Left$(APNU$, 6) + AN$
     FEX = FECHANUM(FE$)
     CIENUE% = DIANTE(FEX)
     FEXX = CIENUE%
     FE$ = FECHATEX$(FEXX)
     FECINU$ = Left$(FE$, 6) + "20" + Mid$(FE$, 7)
     If Val(Mid$(FE$, 7)) > 80 Then FECINU$ = Left$(FE$, 6) + "19" + Mid$(FE$, 7)
     '
     On Error Resume Next
     Tabejcon.Close
     On Error GoTo 0
     Set Tabejcon = Contable.OpenRecordset("TABEJCON", dbOpenTable)
     
     With Tabejcon
       UREI% = 0
       On Error Resume Next
       .MoveLast
       UREI% = .RecordCount
       On Error GoTo 0
       .AddNew
       !CODIEJER = 1 + UREI%
       !Denoejcon = DENUE$
       !Fechaper = CDate(FEAPNU$)
       !Fechcierr = CDate(FECINU$)
       !Status = 1
       !Diverif = 0
       !Marcacti = 1
       .Update
       .Bookmark = .LastModified
     End With
     Tabejcon.Close
     '
     DENOEJ$ = TRIM$(DENUE$)
     COMEJ% = APNUE%
     FINEJ% = CIENUE%
     NEJER$ = TRIM$(Str$(1 + UREI%)): While Len(NEJER$) < 2: NEJER$ = "0" + NEJER$: Wend
     '
     Call GRARAM(87, MKI$(COMEJ%))
     Call GRARAM(89, MKI$(FINEJ%))
     Call GRARAM(321, AJUSTI$(DENOEJ$, 48))
     Call GRARAM(369, NEJER$)
     Combo2.TEXT = DENOEJ$
     '
     Call COMUNI("Nuevo Ejercicio O.K.")
     '
   End If
   '
End Sub

Private Sub Help_Click()
  '
  Call Command2_Click
  '
End Sub

Sub Mayor_Click()
   If COMEJ% < 1 Then
     Call MENSERR(24, "Falta Abrir Ejercicio Contable")
     Exit Sub
   End If
   '
10 RG1$ = REGSEL$("ECUECOT")
   If Len(RG1$) > 4 Then
     COCUE$ = Left$(RG1$, 12)
     'CUEII% = CUEINT%(COCUE$)
     If CUEII% > 0 Then
       Call LMAYORPAN(CUEII%)
       GoTo 10
     End If
   End If
End Sub

Private Sub Mnu_Import_Click()
   Call Command15_Click
End Sub

Private Sub Modempre_Click()
   Call CARMULTIEM
End Sub

Private Sub Nuempre_Click()
   Call CARMULTIEM
End Sub

Private Sub Open_Asi_Click()
   If COMEJ% < 1 Then
     Call MENSERR(24, "Falta Abrir Ejercicio Contable")
     Exit Sub
   End If
   '
   Call ASIDIAR.Open_Asi_Click
   ASIDIAR.Show 1
End Sub

Private Sub Resumen_Click()
   OPBALAN04.Option1(1).Value = 1
   On Error Resume Next
   OPBALAN04.Command1.SetFocus
   On Error GoTo 0
   OPBALAN04.Show 1
End Sub

Private Sub Setup_Click()
   Call Command4_Click
End Sub

Private Sub Susaldos_Click()
   OPBALAN04.Option1(4).Value = 1
   On Error Resume Next
   OPBALAN04.Command1.SetFocus
   On Error GoTo 0
   OPBALAN04.Show 1
End Sub

Sub LMAYORPAN(CUEII%)
   '
   FEINI = Int(CDbl(CVDAT(COMEJ%)))
   FEFIN = Int(CDbl(CVDAT(FINEJ%)))
   '
   Call BLOQARCH("LIMAYOAC")
   Screen.MousePointer = 11
   '
   SQLX$ = "SELECT * FROM DetaAsien "
   SQLX$ = SQLX$ + "WHERE (CueCoInt=" & CUEII%
   SQLX$ = SQLX$ + " AND SucuAsien = " & SUCUI%
   'SQLX$ = SQLX$ + " AND CDBL(Fechasien) >= " & FEINI
   SQLX$ = SQLX$ + " AND CDBL(Fechasien) <= " & FEFIN
   SQLX$ = SQLX$ + " AND statpase >0 ) "
   SQLX$ = SQLX$ + "ORDER BY FECHASIEN, ORFEASIEN ASC;"
   Set PASEASIEN = Contable.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
   '
   JJ& = 0
   TRANSPOR# = 0
10 With PASEASIEN
     On Error Resume Next
     .MoveFirst
     If Err Then GoTo 30
     On Error GoTo 0
     Do While Not .EOF
       IDEBE# = !IDEBEPASE
       IHABE# = !IHABEPASE
       If Int(CDbl(!FECHASIEN)) < FEINI Then
           TRANSPOR# = TRANSPOR# + IDEBE# - IHABE#
         ElseIf Int(CDbl(!FECHASIEN)) <= FEFIN Then
           Z$ = REGBLAN$("LIMAYOAC")
           Call REPLA(Z$, MKI$(CVINT(!FECHASIEN)), 1, 2)
           Call REPLA(Z$, !NUIDASIEN, 3, 12)
             MACON$ = ""
             On Error Resume Next
             MACON$ = TRIM$(!MARCONSI)
             On Error GoTo 0
           Call REPLA(Z$, MACON$, 15, 1)
             DEPASE$ = !REFEPASE
           Call REPLA(Z$, DEPASE$, 17, 32)
           Call REPLA(Z$, MKD$(IDEBE#), 49, 8)
           Call REPLA(Z$, MKD$(IHABE#), 57, 8)
           JJ& = JJ& + 1
           Call GRAREG("LIMAYOAC", Z$, JJ&)
       End If
     .MoveNext
     Loop
   End With
   '
30 On Error GoTo 0
   'Call ABRECABAN
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "where (CueContaI = " & CUEII% & " "
   SQLX$ = SQLX$ + " AND (NuAsidiar = '' or isnull(NuAsiDiar) = true OR IsEmpty(NuAsiDiar) = true)) "
   'SQLX$ = SQLX$ + " AND CDBL(FechConta) >= " & FEINI
   SQLX$ = SQLX$ + " AND CDBL(FechConta) <= " & FEFIN & " "
   SQLX$ = SQLX$ + " ORDER BY FECHCONTA ASC, FEHOREAL ASC;"
   '
   Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   With CABATEMP
     On Error Resume Next
     .MoveFirst
     If Err Then
        GoTo 40
     End If
     On Error GoTo 0
     Do While Not .EOF
         IDEBE# = !IDEBECOMP
         IHABE# = !IHABECOMP
       If Int(CDbl(!FechConta)) < FEINI Then
           TRANSPOR# = TRANSPOR# + IDEBE# - IHABE#
         ElseIf Int(CDbl(!FechConta)) <= FEFIN Then
           Z$ = REGBLAN$("LIMAYOAC")
           Call REPLA(Z$, MKI$(CVINT(!FechConta)), 1, 2)
           Call REPLA(Z$, !CodiCom_Cor, 3, 12)
             MACON$ = ""
             On Error Resume Next
             MACON$ = TRIM$(!MARCONSI)
             On Error GoTo 0
           Call REPLA(Z$, MACON$, 15, 1)
             DEPASE$ = !RefeCom
           Call REPLA(Z$, DEPASE$, 17, 32)
           Call REPLA(Z$, MKD$(IDEBE#), 49, 8)
           Call REPLA(Z$, MKD$(IHABE#), 57, 8)
           JJ& = JJ& + 1
           Call GRAREG("LIMAYOAC", Z$, JJ&)
       End If
     .MoveNext
     Loop
   End With
   '
40 Call SETULTREG("LIMAYOAC", JJ&)
   Call FILESORT("LIMAYOAC", "", 1, 1, "ASC")
   '
   MAYORPAN.List1.Clear
   Z$ = Space$(128)
     FEX = DIANTE(COMEJ%)
     FEXA$ = FECHATEX$(FEX)
   Call REPLA(Z$, FEXA$, 1, 8)
     DEPASE$ = "TRANSPORTE .............................."
   Call REPLA(Z$, DEPASE$, 35, 30)
   Call REPLA(Z$, CSTRING$(MKD$(TRANSPOR#), 3, 14, 2, 2), 94, 14)
   MAYORPAN.List1.AddItem Z$
   MAYORPAN.List1.AddItem Space$(95) + " =========== "
   '
   SUDEBE1# = 0: SUHABE1# = 0
   For U& = 1 To ULTREG("LIMAYOAC")
     X$ = REGLEIDO$("LIMAYOAC", U&)
     Z$ = Space$(128)
       FEX = CVI(Left$(X$, 2))
       FEXA$ = FECHATEX$(FEX)
     Call REPLA(Z$, FEXA$, 1, 8)
       COCOM$ = TRIM$(Mid$(X$, 3, 12))
     Call REPLA(Z$, COCOM$, 10, 14)
       MACON$ = TRIM$(Mid$(X$, 15, 2))
     Call REPLA(Z$, MACON$, 24, 2)
       DEPASE$ = TRIM$(Mid$(X$, 17, 32))
     Call REPLA(Z$, DEPASE$, 35, 30)
       IDEBE# = CVD(Mid$(X$, 49, 8))
       IHABE# = CVD(Mid$(X$, 57, 8))
     If IDEBE# > 0 Then
         Call REPLA(Z$, CSTRING$(MKD$(IDEBE#), 3, 14, 2, 2), 66, 14)
         SUDEBE1# = SUDEBE1# + IDEBE#
       Else
         Call REPLA(Z$, CSTRING$(MKD$(IHABE#), 3, 14, 2, 2), 80, 14)
         SUHABE1# = SUHABE1# + IHABE#
     End If
       TRANSPOR# = TRANSPOR# + IDEBE# - IHABE#
     Call REPLA(Z$, CSTRING$(MKD$(TRANSPOR#), 3, 14, 2, 2), 94, 14)
     MAYORPAN.List1.AddItem Z$
   Next U&
   Call CIERRARCH("LIMAYOAC")
   '
   On Error GoTo 0
   If MAYORPAN.List1.ListCount > 0 Then
     MAYORPAN.List1.ListIndex = MAYORPAN.List1.ListCount - 1
   End If
   '
   'MAYORPAN.Ltext2 = CODCUE$(CUEII%)
   'MAYORPAN.Label4 = DENOCUE$(CUEII%)
   MAYORPAN.Label21 = TRIM$(FORMATNUM$(SUDEBE1#, "F11.2"))
   MAYORPAN.Label22 = TRIM$(FORMATNUM$(SUHABE1#, "F11.2"))
   Screen.MousePointer = 1
   MAYORPAN.Show 1
   '
End Sub

Sub CARMULTIEM()
   '
   Call BLANARCH("!CARMULTI")
   NX% = FILEOPEN("FLEXMUL.DRV", 0, 79)
   Dim FMX As String * 79
   '
   For I% = 1 To LOF(NX%) / 79
     Get #NX%, I%, FMX$
     If Left$(FMX$, 77) <> String$(77, 0) Then
       SUCAR% = 0
       For J% = 1 To 77
         SUCAR% = SUCAR% + Asc(Mid$(FMX$, J%, 1))
       Next J%
       If CVI(Mid$(FMX$, 78, 2)) = SUCAR% Then
         RASO$ = DESENCRI$(Left$(FMX$, 30))
         PACCE$ = DESENCRI$(Mid$(FMX$, 32, 46))
         Z$ = Space$(96)
         Call REPLA(Z$, RASO$, 1, 30)
         Call REPLA(Z$, PACCE$, 31, 46)
         Call REGAPP("!CARMULTI", Z$)
       End If
     End If
   Next I%
   '
   Close #NX%
   Call CIERRARCH("!CARMULTI")
   '
   VTB% = VENTABU%("CARMULTI")
   If VTB% >= 0 Then
     '
     NX% = FILEOPEN("FLEXMUL.DRV", 0, 79)
     Dim FMY As String * 79
     '
     For U& = 1 To ULTREG&("!CARMULTI")
       Z$ = REGLEIDO$("!CARMULTI", U&)
       RASO$ = AJUSTI$(Left$(Z$, 30), 30)
       PACCE$ = AJUSTI$(Mid$(Z$, 31, 46), 46)
       If TRIM$(RASO$) <> "" Then
         If TRIM$(PACCE$) <> "" Then
           Z1$ = String$(79, 0)
           Call REPLA(Z1$, ENCRIPTA$(RASO$), 1, 30)
           Call REPLA(Z1$, ENCRIPTA$(PACCE$), 32, 46)
           SUCAR% = 0
           For J% = 1 To 77
             SUCAR% = SUCAR% + Asc(Mid$(Z1$, J%, 1))
           Next J%
           Call REPLA(Z1$, MKI$(SUCAR%), 78, 2)
           V& = V& + 1
           LSet FMY$ = Z1$
           Put #NX%, V&, FMY$
         End If
       End If
     Next U&
     '
     For U& = V& + 1 To LOF(NX%) / 79
       LSet FMY$ = String$(79, 0)
       Put #NX%, U&, FMY$
     Next U&
     '
     Close #NX%
     Call CIERRARCH("!CARMULTI")
     '
   End If
   '
End Sub

Private Sub MnuArchiva_Click()
  '
  Call ARCHILIQ
  '
End Sub

Private Sub MnuCalFer_Click()
  '
  Call COSUJO05.Option1_DBLClick(7)
  '
End Sub

Private Sub MnuDatLiqAct_Click()
  'Edita Liquidacion Activa
  XX$ = REGLEIDO$("ULTILIQ", 1)
  VDEF$ = Space(128)
  DESCRILIQ$ = Mid$(XX$, 1, 30)
  GRUP% = Asc(Mid$(XX$, 31, 1))
  MESACU% = Asc(Mid$(XX$, 32, 1))
  ANOACU% = Asc(Mid$(XX$, 33, 1))
  TIPOLIQ% = Asc(Mid$(XX$, 34, 1))
  ASIFLI% = Asc(Mid$(XX$, 35, 1))
  FEPA% = CVI(Mid$(XX$, 36, 2))
  LUPA$ = Mid$(XX$, 38, 20)
  PEULDE$ = Mid$(XX$, 58, 16)
  FEDE% = CVI(Mid$(XX$, 74, 2))
  BANCOLIQ$ = Mid$(XX$, 76, 16)
  PERIOLIQ$ = Mid$(XX$, 92, 12)
  MESIMPU% = Asc(Mid$(XX$, 106, 1))
  ANOIMPU% = Asc(Mid$(XX$, 107, 1))
  '
  Call REPLA(VDEF$, DESCRILIQ$, 1, 30)
  Call REPLA(VDEF$, PERIOLIQ$, 31, 12)
  Call REPLA(VDEF$, Chr$(GRUP%), 43, 1)
  Call REPLA(VDEF$, Chr$(MESACU%), 44, 1)
   ANOACUM$ = CStr(ANOACU%)
   Do While Len(ANOACUM$) < 2
    ANOACUM$ = "0" + ANOACUM$
   Loop
  Call REPLA(VDEF$, ANOACUM$, 45, 2)
  Call REPLA(VDEF$, Chr$(MESIMPU%), 47, 1)
   ANOIMP$ = TRIM$(FORMATNUM$(ANOIMPU%, "F2.0"))
   Do While Len(ANOIMP$) < 2
    ANOIMP$ = "0" + ANOIMP$
   Loop
  Call REPLA(VDEF$, ANOIMP$, 48, 2)
  Call REPLA(VDEF$, Chr$(TIPOLIQ%), 50, 1)
  Call REPLA(VDEF$, Chr(ASIFLI%), 51, 1)
  Call REPLA(VDEF$, MKI$(FEPA%), 52, 2)
  Call REPLA(VDEF$, LUPA$, 54, 20)
  Call REPLA(VDEF$, PEULDE$, 74, 16)
  Call REPLA(VDEF$, MKI$(FEDE%), 90, 2)
  Call REPLA(VDEF$, BANCOLIQ$, 92, 16)
30   OBJ$ = OBJPLA$("LIQUIDACI", VDEF$)
  VALO$ = OBJ$
  'SI ACEPTA GRABA CAMBIOS
  If TRIM(VALO$) <> "" Then
    DESCRILIQ$ = Mid$(VALO$, 1, 30)
    If TRIM(DESCRILIQ$) = "" Then
     Call MENSERR(24, "Denominación de Liquidación no Válida.")
     GoTo 30
    End If
    PERIOLIQ$ = Mid$(VALO$, 31, 12)
    GRUP% = Asc(Mid$(VALO$, 43, 1))
    MESACU% = Asc(Mid$(VALO$, 44, 1))
     ANOACUM$ = Mid$(VALO$, 45, 2)
    ANOACU% = Val(ANOACUM$)
    MESIMPU% = Asc(Mid$(VALO$, 47, 1))
     ANOIMP$ = Mid$(VALO$, 48, 2)
    ANOIMPU% = Val(ANOIMP$)
    TIPOLIQ% = Asc(Mid$(VALO$, 50, 1))
    ASIFLI% = Asc(Mid$(VALO$, 51, 1))
    FEPA% = CVI(Mid$(VALO$, 52, 2))
    LUPA$ = Mid$(VALO$, 54, 20)
    PEULDE$ = Mid$(VALO$, 74, 16)
    FEDE% = CVI(Mid$(VALO$, 90, 2))
    BANCOLIQ$ = Mid$(VALO$, 92, 16)
    '
    XX$ = REGLEIDO$("ULTILIQ", 1)
    Call REPLA(XX$, DESCRILIQ$, 1, 30)
    Call REPLA(XX$, Chr$(GRUP%), 31, 1)
    Call REPLA(XX$, Chr$(MESACU%), 32, 1)
    Call REPLA(XX$, Chr$(ANOACU%), 33, 1)
    Call REPLA(XX$, Chr$(TIPOLIQ%), 34, 1)
    Call REPLA(XX$, Chr$(ASIFLI%), 35, 1)
    Call REPLA(XX$, MKI$(FEPA%), 36, 2)
    Call REPLA(XX$, LUPA$, 38, 20)
    Call REPLA(XX$, PEULDE$, 58, 16)
    Call REPLA(XX$, MKI$(FEDE%), 74, 2)
    Call REPLA(XX$, BANCOLIQ$, 76, 16)
    Call REPLA(XX$, PERIOLIQ$, 92, 12)
    Call REPLA(XX$, Chr$(MESIMPU%), 106, 1)
    Call REPLA(XX$, Chr$(ANOIMPU%), 107, 1)
    Call GRAREG("ULTILIQ", XX$, 1)
    '
  End If
  '
End Sub
Private Sub MnuFin_Click()
  '
  Call Command1_Click
  '
End Sub

Private Sub MnuLiqOrd_Click()
   FECALANT$ = TRIM$(CONTROL("SUELJOR", "FANTIGU"))
   If FECALANT$ = "" Then
      Call COMUNI("Debe Ahora Recalcular\la Antigüedad en la Empresa")
      GoTo 20
   End If
   TTXX$ = "La Antigüedad de Cálculo Registrada\Corresponde al '" + FECALANT$ + "'."
   TTXX$ = TTXX$ + "\   \¿ Desea Ahora Realizar el Recálculo a Fecha ?"
   If COMALTER%(TTXX$ + "\\Si, Recalcular\No, Continuar") = 2 Then GoTo 30
20 Call CALCUANTIGU(1)
   '
30 On Error Resume Next
   REMUORD05.Show 1
   On Error GoTo 0
End Sub

Private Sub MnuLisEmp_Click()
  '
  Call EMP_GES05.Command13_Click
  '
End Sub

Private Sub MnuLisInfHor_Click()
  '
  Call LISINFHOR
  '
End Sub

Private Sub MnuModExi_Click()
  '
  Call EMP_GES05.Command9_Click
  EMP_GES05.Show 1
  '
End Sub

Private Sub MnuNueEmp_Click()
  '
  Call EMP_GES05.Command21_Click
  '
End Sub

Private Sub MnuNueLiq_Click()
  '
  If COMALTER%("¿Borra Liquidación Anterior?\\Si, Borrar\No, Conservar") = 1 Then
    'borra ultliq.
    Call BLANARCH("ULTILIQ")
    OBJ$ = OBJPLA$("LIQUIDACI", VDEF$)
    VALO$ = OBJ$
    'SI ACEPTA GRABA CAMBIOS
    If TRIM(VALO$) <> "" Then
      DESCRILIQ$ = Mid$(VALO$, 1, 30)
      PERIOLIQ$ = Mid$(VALO$, 31, 12)
      GRUP% = Asc(Mid$(VALO$, 43, 1))
      MESACU% = Asc(Mid$(VALO$, 44, 1))
      ANOACUM$ = Mid$(VALO$, 45, 2)
       ANOACU% = Val(ANOACUM$)
      MESIMPU% = Asc(Mid$(VALO$, 47, 1))
       ANOIMP$ = Mid$(VALO$, 48, 2)
      ANOIMPU% = Val(ANOIMP$)
      TIPOLIQ% = Asc(Mid$(VALO$, 50, 1))
      ASIFLI% = Asc(Mid$(VALO$, 51, 1))
      FEPA% = CVI(Mid$(VALO$, 52, 2))
      LUPA$ = Mid$(VALO$, 54, 20)
      PEULDE$ = Mid$(VALO$, 74, 16)
      FEDE% = CVI(Mid$(VALO$, 90, 2))
      BANCOLIQ$ = Mid$(VALO$, 92, 16)
      '
      XX$ = REGLEIDO$("ULTILIQ", 1)
      Call REPLA(XX$, DESCRILIQ$, 1, 30)
      Call REPLA(XX$, Chr$(GRUP%), 31, 1)
      Call REPLA(XX$, Chr$(MESACU%), 32, 1)
      Call REPLA(XX$, Chr$(ANOACU%), 33, 1)
      Call REPLA(XX$, Chr$(TIPOLIQ%), 34, 1)
      Call REPLA(XX$, Chr$(ASIFLI%), 35, 1)
      Call REPLA(XX$, MKI$(FEPA%), 36, 2)
      Call REPLA(XX$, LUPA$, 38, 20)
      Call REPLA(XX$, PEULDE$, 58, 16)
      Call REPLA(XX$, MKI$(FEDE%), 74, 2)
      Call REPLA(XX$, BANCOLIQ$, 76, 16)
      Call REPLA(XX$, PERIOLIQ$, 92, 12)
      Call REPLA(XX$, Chr$(MESIMPU%), 106, 1)
      Call REPLA(XX$, Chr$(ANOIMPU%), 107, 1)
      Call GRAREG("ULTILIQ", XX$, 1)
      '
      Call SETULTREG("DETLIQUI", 0)
    End If
  End If
  '
End Sub

Private Sub MnuParInd_Click()
  '
  Call COSUJO05.Option1_DBLClick(9)
  '
End Sub

Private Sub MnuPlanCam_Click()
  'Listado de Planilla de Calculo
Dim NETOLEG1#(32767)
Dim VALBI(10)
Dim CABI%(32767)
Dim NUBI%(32767)
'1.-Neto para cada Empleado con Liquidación
  FIN& = ULTREG&("DETLIQUI")
  For J1& = 1 To FIN&
    X$ = REGLEIDO$("DETLIQUI", J1&)
    NLEG1% = CVI(Mid$(X$, 1, 2))
    CONCEP1% = CVI(Mid$(X$, 3, 2))
    IMPORT1# = CVD(Mid$(X$, 9, 8))
    '
    If CONCEP1% <= 199 Or CONCEP1% >= 401 Then
      NETOLEG1#(NLEG1%) = NETOLEG1#(NLEG1%) + (Int(100 * IMPORT1# + 0.5)) / 100
      GoTo 5
    End If
    NETOLEG1#(NLEG1%) = NETOLEG1#(NLEG1%) - (Int(100 * Abs(IMPORT1#) + 0.5)) / 100
    '
5 Next J1&
'2.-OBTENGO LOS VALORES DE LOS BILLETES
C1$ = REGLEIDO$("BILLETES", 1)
5020    For I% = 1 To 10
5022    VALBI(I%) = XVALO(CVS(Mid$(C1$, 4 * I% - 3, 4)))
5025    If VALBI(I%) = 0 Then GoTo 5045
5030    Next I%
5045    CAVAL = I% - 1
5077    VALMIN# = CDbl(VALBI(1))
5078    For I% = 1 To 10
5080    If VALBI(I%) <= 0 Then GoTo 5085
5082    If VALBI(I%) < VALMIN# Then VALMIN# = CDbl(VALBI(I%))
5085    Next I%
5086    If VALMIN# < 0.01 Then VALMIN# = 1
'2.1 VALIDACION DE QUE ESTE CARGADO ALGÚN BILLETE
If CAVAL = 0 Then
   Call MENSERR(24, "Tabla Valores de Cambio para Pago\Incompleta o No Definida.")
   Exit Sub
End If
'3.-CALCULO DE DISTRIBUCION
REBLA$ = REGBLAN$("LIBILLET")
JJJ& = 0
For XB% = 1 To 32767
If XB% = 32767 Then Exit For
  If NETOLEG1#(XB%) > 0.005 Then
    NETO# = NETOLEG1#(XB%)
    NETO1# = 100 * NETO#: MENVAL = 100 * VALBI(CAVAL): DIFE = NETO1# - MENVAL * Int(NETO1# / MENVAL)
19001     If DIFE >= MENVAL / 2 Then DIFE = DIFE - MENVAL
19002     NETO1# = NETO1# - DIFE: NETO# = (Int(NETO1# + 0.5)) / 100
19003     RESTO# = (Int(100 * NETO# + 0.5)) / 100
19004     '
19005     For J% = 1 To CAVAL
19010     CABI%(J%) = Int(RESTO# / VALBI(J%))
19015     '
19020     RESTO# = (Int(100 * RESTO# - 100 * CDbl(CABI%(J%)) * VALBI(J%) + 0.5)) / 100
19025     Next J%
19027     If RESTO# >= (VALBI(CAVAL)) / 2 Then CABI%(CAVAL) = CABI%(CAVAL) + 1
19030     If CABI%(CAVAL) < VALBI(CAVAL - 1) / VALBI(CAVAL) Then GoTo 19054
19035     CABI%(CAVAL - 1) = CABI%(CAVAL - 1) + 1
19040     DESCO% = VALBI(CAVAL - 1) / VALBI(CAVAL)
19045     CABI%(CAVAL) = CABI%(CAVAL) - DESCO%
19052     '
19054     For J% = CAVAL To 2 Step -1
19056     If CABI%(J%) * VALBI(J%) >= VALBI(J% - 1) Then CABI%(J%) = CABI%(J%) - Int(VALBI(J% - 1) / VALBI(J%)): CABI%(J% - 1) = CABI%(J% - 1) + 1: GoTo 19056
19058     Next J%
19060     '
19062     For J% = 1 To CAVAL
           On Error Resume Next
           TOTBIL = NUBI%(J%) + CABI%(J%)
           If Err.Number = 6 Then
              Call MENSERR(24, "Cantidad de Billetes no Permitida.\Aumente el Valor de Billetes para Cambio.")
              On Error GoTo 0
              Exit Sub
           End If
           On Error GoTo 0
           NUBI%(J%) = NUBI%(J%) + CABI%(J%)
          Next J%
    '4.-AGREGAR A ARCHIVO LISTADOR
    YY$ = REBLA$
    POSINI% = 2
    Call REPLA(YY$, MKI$(XB%), 2, 2)
    For J% = 1 To CAVAL
      NUBI%(J%) = NUBI%(J%) + CABI%(J%)
      POSINI% = POSINI% + 2
      Call REPLA(YY$, MKI$(CABI%(J%)), POSINI%, 2)
    Next J%
    Call REPLA(YY$, MKD$(NETOLEG1#(XB%)), 24, 8)
    Call REPLA(YY$, MKD$(VALES1#), 32, 8) 'VALES????
    IBRUTO1# = NETOLEG1#(XB%) - VALES1#
    Call REPLA(YY$, MKD$(IBRUTO1#), 40, 8)
    JJJ& = JJJ& + 1
    Call GRAREG("LIBILLET", YY$, JJJ&)
  End If
Next XB%
'5.-LISTAR
Call SETULTREG("LIBILLET", JJJ&)
Call HEADERS("LIBILLET", "")
Call HEADERS("LIBILLET", "Liquidación: " + DESCRI_LIQ$)
'6.- Headers con Valores
XY$ = ""
XJ% = 0
A1$ = ""
A2$ = ""
C1$ = REGLEIDO$("BILLETES", 1)
For I% = 1 To 10
  VALCA1 = XVALO(CVS(Mid$(C1$, 4 * I% - 3, 4)))
  XJ% = XJ% + 1
  Select Case XJ%
    Case 1
      A1$ = "A= " + TRIM$(FORMATNUM$(VALCA1, "F8.2"))
    Case 2
      A1$ = A1$ + " B = " + TRIM$(FORMATNUM$(VALCA1, "F8.2"))
    Case 3
      A1$ = A1$ + " C = " + TRIM$(FORMATNUM$(VALCA1, "F8.2"))
    Case 4
      A1$ = A1$ + " D = " + TRIM$(FORMATNUM$(VALCA1, "F8.2"))
    Case 5
      A1$ = A1$ + " E = " + TRIM$(FORMATNUM$(VALCA1, "F8.2"))
    Case 6
      A2$ = A2$ + " F = " + TRIM$(FORMATNUM$(VALCA1, "F8.2"))
    Case 7
      A2$ = A2$ + " G = " + TRIM$(FORMATNUM$(VALCA1, "F8.2"))
    Case 8
      A2$ = A2$ + " H = " + TRIM$(FORMATNUM$(VALCA1, "F8.2"))
    Case 9
      A2$ = A2$ + " I = " + TRIM$(FORMATNUM$(VALCA1, "F8.2"))
    Case 10
      A2$ = A2$ + " J = " + TRIM$(FORMATNUM$(VALCA1, "F8.2"))
  End Select
Next I%
Call HEADERS("LIBILLET", "Billetes: " + A1$ + A2$)
'
Call FINAL("*LIBILLET")
End Sub

Private Sub MnuPlanHab_Click()
   ' planilla de haberes
   Call PLANIHAB
   '
   Call FINAL("*PLANIHA")
   '
End Sub


Private Sub MnuRecal_Click()
  '
  Call CALCUANTIGU
  '
End Sub

Private Sub MnuRecHab_Click()
  '
  '1.-Indice con Legajos con Liquidación Activa
  Call INDILEGLIQ
  '2.- Seleccion de Modo de Impresión
  PRINUME& = 1 + XVALO(CONTROL$("NUMERECI", "UNURECI"))
2 NUMERECI& = XVALO(InputBox$("Número Primer Recibo", "Numeración de Recibos de Haberes", PRINUME&))
  If NUMERECI& < 1 Then Exit Sub
  If NUMERECI& > 900000 Then GoTo 2
  '
5 RTA% = ALTERNA%("Por Legajo\Por Liquidación")
  Select Case RTA%
    Case 1
10    Call FRESHECHO("!INDPADRO")
      Call SELECHO("!INDPADRO")
      NLEG% = Val(TRIM$(VALACT1$("!INDPADRO")))
      If NLEG% < 1 Then Exit Sub
      NUMEDOC$ = TRIM$(FORMATNUM$(NUMERECI&, "F8.0"))
      Call GRACONTROL("NUMERECI", "UNURECI", NUMEDOC$)
      While Len(NUMEDOC$) < 6: NUMEDOC$ = "0" + NUMEDOC$: Wend
      Call PRIRECISUEL(NLEG%)
      NUMERECI& = NUMERECI& + 1
      GoTo 10
    Case 2
      FIN& = ULTREG&("!INDPADRO")
      For J& = 1 To FIN&
        X$ = REGLEIDO$("!INDPADRO", J&)
        NLEG% = CVI(Mid$(X$, 1, 2))
        NUMEDOC$ = TRIM$(FORMATNUM$(NUMERECI&, "F8.0"))
        Call GRACONTROL("NUMERECI", "UNURECI", NUMEDOC$)
        While Len(NUMEDOC$) < 6: NUMEDOC$ = "0" + NUMEDOC$: Wend
        Call PRIRECISUEL(NLEG%)
        NUMERECI& = NUMERECI& + 1
      Next J&
      GoTo 5
  End Select
  '
End Sub
Sub INDILEGLIQ()
  '
  JJ& = 0
  ZBLA$ = REGBLAN$("INDPADRO")
  FIN& = ULTREG&("DETLIQUI")
  Call FILESORT("DETLIQUI", "DETLIQUI", 1, 1, "ASC")
  For J& = 1 To FIN&
    Z$ = ZBLA$
    XX$ = REGLEIDO$("DETLIQUI", J&)
    NLEG1% = CVI(Mid$(XX$, 1, 2))
    If NLEG1% <> NLEG0% Then
      Call REPLA(Z$, MKI$(NLEG1%), 1, 2)
      Call REPLA(Z$, ECOARCH$("PADREMP", MKI$(NLEG1%)), 3, 30)
      JJ& = JJ& + 1
      Call GRAREG("!INDPADRO", Z$, JJ&)
      NLEG0% = NLEG1%
    End If
  Next J&
  Call SETULTREG("!INDPADRO", JJ&)
  '
End Sub

Sub PRIRECISUEL(NLEGA%)
    '
    Dim SB(128)
    '
    FORIPRE$ = TRIM$(CONTROL$("", "RECIHABE"))  'REICBO HABERES
    FPXX$ = "RECIHABE"
    If FORIPRE$ = "" Then
       Call MENSERR(24, "Imposible Imprimir Recibo de Haberes.\Falta Definir Formulario Electrónico.")
       Exit Sub
    End If
    If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
       Call MENSERR(24, "Imposible Imprimir Recibo de Haberes.\No existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
       Exit Sub
    End If
    '
10  TIPODOC$ = "Recibo de Haberes"
    ' NUMEDOC$ = "000001" ' VIENE DE LA RUTINA CONTROLADORA
    HOII% = HOY(HO$)
    FECHDOC$ = TRIM$(FECHATEX$(HOII%))
    'ENCABEZADO
    Call CARGEMP(NLEGA%)
    LEGACUIL$ = CUIL_LEG$
    NOAPLEG$ = APNO_LEG$
    DOCULEG$ = DOCU_LEG$
    SECTLEG$ = FORMATNUM$(SECC_LEG, "F7.3")
    CATELEG$ = ECOARCH$("CATEGO", CATE_LEG$)
    TARELEG$ = TARE_LEG$
    FEINLEG$ = FECHATEX$(FEIN_LEG%)
    '
    Call CARGLIQACT
    BANCLEG$ = BANCO_LIQ$
    PERIBAN$ = PEULDE_LIQ$
    FECDEPO$ = FECHATEX$(FEDE_LIQ%)
    REMBASI$ = "" '???????
    PERIABO$ = PERIO_LIQ$
    LUGARBO$ = LUPA_LIQ$
    '
'    X1$ = REGLEIDO$("ULTILIQ", 1)
'    BANCLEG$ = TRIM$(Mid$(X1$, 76, 16))
'    PERIBAN$ = TRIM$(Mid$(X1$, 58, 16)) 'PERIODO LIQ. ANT.
'    FECDEPO$ = FECHATEX$(CVI(Mid$(X1$, 74, 2)))
'    TARELEG$ = TRIM$(Mid$(XX$, 161, 20))
'    FEINLEG$ = FECHATEX$(CVI(Mid$(XX$, 145, 2)))
'    REMBASI$ = "" '???????
'    PERIABO$ = TRIM$(Mid$(X1$, 91, 12)) 'PERIODO LIQUIDADO LIQ.
'    LUGARBO$ = TRIM$(Mid$(X1$, 38, 20))
    '
    XX1$ = RECFILT$("PADREMP", 1, MKI$(NLEGA%))
    J1& = CVS(XX1$)
    '
    For VX% = 1 To NUREM%: SB(VX%) = 0: Next VX%
    RBPI$ = REGLEIDO$("REMBAS", J1&)
    RB1$ = Left$(RBPI$, 48)
    For VX% = 1 To NUREM%
      SB(VX%) = CVS(Mid$(RB1$, 4 * VX% - 3))
    Next VX%
    '
    CODPARAM$(1) = "CUIT-CLI": DOCPARAM$(1) = LEGACUIL$
    CODPARAM$(2) = "COT-PLCO": DOCPARAM$(2) = NOAPLEG$
    CODPARAM$(3) = "REF-PLCO": DOCPARAM$(3) = DOCULEG$
    CODPARAM$(4) = "MED-PLCO": DOCPARAM$(4) = SECTLEG$
    CODPARAM$(5) = "CAN-PLCO": DOCPARAM$(5) = TRIM$(Str$(NLEGA%))
    CODPARAM$(6) = "FRE-PLCO": DOCPARAM$(6) = CATELEG$
    CODPARAM$(7) = "REA-PLCO": DOCPARAM$(7) = BANCLEG$
    CODPARAM$(8) = "LAR-CUTE": DOCPARAM$(8) = PERIBAN$
    CODPARAM$(9) = "ACH-CUTE": DOCPARAM$(9) = FECDEPO$
    CODPARAM$(10) = "LAR-MATE": DOCPARAM$(10) = TARELEG$
    CODPARAM$(11) = "ACH-MATE": DOCPARAM$(11) = FEINLEG$
    CODPARAM$(12) = "LAR-CUCO": DOCPARAM$(12) = REMBASI$
    CODPARAM$(13) = "ACH-CUCO": DOCPARAM$(13) = PERIABO$
    CODPARAM$(14) = "LAR-MACO": DOCPARAM$(14) = LUGARBO$
    CODPARAM$(22) = "R-PLANO": DOCPARAM$(22) = AFJP_LEG$
    CODPARAM$(23) = "N-PLANO": DOCPARAM$(23) = TCON_LEG$
    CODPARAM$(24) = "ACH-MACO": DOCPARAM$(24) = FECHATEX$(FEPA_LIQ%)
    CODPARAM$(25) = "LAR-CUCO": DOCPARAM$(25) = FORMATNUM$(SB(3) + SB(6), "F10.2")
    CAPARDOC% = 25
    'DETALLE
    '
    CODTABU1$(1) = "REF-MAT1"
    CODTABU1$(2) = "REF-MAT2"
    CODTABU1$(3) = "REF-MAT3"
    CODTABU1$(4) = "MEDIDA-1"
    CODTABU1$(5) = "MEDIDA-2"
    CODTABU1$(6) = "MEDIDA-3"
    CODTABU1$(7) = "IMP-NETO" 'REMUNERACIONES NO GRAVADAS
    CACOLTAB1% = 7
    '
    CAITAB1% = 0
    SUBDES = 0
    SUBREM = 0 'TOTAL REMUENERACIONES GRAVADAS
    SUBREMNO = 0 'TOTAL REMUENERACIONES NO GRAVADAS
    Call CARGAFORMULAS
    Call FILTERFILE("DETLIQUI", "!DETLIQUI", 1, MKI$(NLEGA%))
    FIN& = ULTREG&("!DETLIQUI")
    For U& = 1 To ULTREG&("!DETLIQUI")
      X$ = REGLEIDO$("!DETLIQUI", U&)
      CONC% = CVI(Mid$(X$, 3, 2))
      CAN = CVS(Mid$(X$, 5, 4))
      INETO = CVD(Mid$(X$, 9, 8))
      '
      CAITAB1% = CAITAB1% + 1
      TETABU1$(1, CAITAB1%) = ECOARCH$("ECONCEP", MKI$(CONC%))
      TETABU1$(2, CAITAB1%) = FORMATNUM$(CONC%, "F6.0")
      TETABU1$(3, CAITAB1%) = FORMATNUM$(CAN, "F6.0")
      '
      VAUNI1 = 0
      If CAN > 0 Then 'Para evitar Division por cero
        VAUNI1 = INETO / CAN
      End If
      TETABU1$(4, CAITAB1%) = FORMATNUM$(VAUNI1, "F12.4")
      If CONC% <= 200 Or CONC% >= 400 Then 'INETO > 0.005 Then
        If CONC% < 100 Then 'SI ES REMUNERATIVO
          If SURETLE%(CONC%) > 0 Then 'REM. GRAVADAS
            TETABU1$(5, CAITAB1%) = FORMATNUM$(INETO, "F12.2")
            SUBREM = SUBREM + (INETO)
          Else 'REM. NO GRAVADAS (POR MARCA)
            TETABU1$(7, CAITAB1%) = FORMATNUM$(INETO, "F12.2")
            SUBREMNO = SUBREMNO + (INETO)
          End If
        Else 'REM. NO GRAVADAS(POR ASIG. FLIARES,ETC)
          TETABU1$(7, CAITAB1%) = FORMATNUM$(INETO, "F12.2")
          SUBREMNO = SUBREMNO + (INETO)
        End If
      Else 'DESCUENTOS
        TETABU1$(6, CAITAB1%) = FORMATNUM$(INETO, "F12.2")
        SUBDES = SUBDES + (Abs(INETO))
      End If
      '
    Next U&
    '
    TOTREMU = SUBREM + SUBREMNO
    IMPOLET1$ = NUMTEX$(TOTREMU - SUBDES)
    IMPOLET2$ = ""
    CODPARAM$(15) = "NET-ITEM": DOCPARAM$(15) = FORMATNUM$(SUBREM, "F12.2") 'TOTAL REM. GRAV.
    CODPARAM$(16) = "BRU-ITEM": DOCPARAM$(16) = FORMATNUM$(SUBDES, "F12.2")
    CODPARAM$(17) = "IMP-TOTA": DOCPARAM$(17) = FORMATNUM$(TOTREMU - SUBDES, "F12.2")
    CODPARAM$(18) = "IMP-LET1": DOCPARAM$(18) = IMPOLET1$
    CODPARAM$(19) = "IMP-LET2": DOCPARAM$(19) = IMPOLET2$ '??????????'
    CODPARAM$(20) = "PRE-NETO": DOCPARAM$(20) = FORMATNUM$(SUBREMNO, "F12.2") 'TOT. REM. NO GRAV.
    CODPARAM$(21) = "IMPNEDES": DOCPARAM$(21) = FORMATNUM$(TOTREMU, "F12.2") 'TOTAL REMUNERACIONES
    '
    Call PRINTDOC("RECIHABE")
    '
    For KK1% = 1 To CACOLTAB1%
      For KK2% = 1 To CAITAB1%
        TETABU1$(KK1%, KK2%) = ""
      Next KK2%
    Next KK1%
    '
    For KK1% = 1 To CACOLTAB2%
      For KK2% = 1 To CAITAB2%
        TETABU2$(KK1%, KK2%) = ""
      Next KK2%
    Next KK1%
    '
    Call CIERRARCH("*.*")
    '
End Sub

Private Sub MnuRecupera_Click()
  '
  Call RECUPELIQ
  '
End Sub

Private Sub MnuSalNom_Click()
  '
  Call COSUJO05.Option1_DBLClick(10)
  '
End Sub

'
Private Sub MnuStp_Click()
  '
  Call COSUJO05.Option1_DBLClick(5)
  '
End Sub

Private Sub MnuStpEmp_Click()
  '
  Call EMP_GES05.Command14_Click
  '
End Sub

Private Sub MnuSueJor_Click()
  '
  Call LIBROSUEL
  Call MENSERR(24, "No Disponible en la Presente Versión")
  '
End Sub

Private Sub MnuVCP_Click()
  '
  Call COSUJO05.Option1_DBLClick(6)
  '
End Sub

Private Sub Picture7_Click()
   AAA = BBB
End Sub

Private Sub plh7_Click()
  '
  Call Command13_Click
  '
End Sub

Sub CALCUANTIGU(Optional MODO%)
    '
    If MODO% = 1 Then
       OPX% = 1
       GoTo 20
    End If
    '
    OPX% = COMALTER%("Opciones de Trabajo:\\Recálculo a Fecha\Listado Impreso")
    If OPX% < 1 Then Exit Sub
    If OPX% = 2 Then GoTo 100
    '
20  Call SELECHO("FECHA")
    If VALACT1$("FECHA") = "" Then Exit Sub
    FECALCU = FECHANUM(VALACT1$("FECHA"))
    '
    Screen.MousePointer = 11
    REBLA$ = REGBLAN$("ANTICAL")
    JJ& = 0
    For U& = 1 To ULTREG&("PADREMP")
      XX$ = REGLEIDO$("PADREMP", U&)
      LEGAX% = CVI(Left$(XX$, 2))
      GRU% = Asc(Mid$(XX$, 49, 1))
      If Asc(Mid$(XX$, 48, 1)) = 1 Then ' STATUS ACTIVO
        FEIN1 = CVI(Mid$(XX$, 46, 2))
        FANTIG1 = CVI(Mid$(XX$, 145, 2))
          If FANTIG1 = 0 Then FANTIG1 = FEIN1
          If FANTIG1 > FEIN1 Then FANTIG% = FEIN1
        FECA$ = FECHATEX$(FECALCU)
        FATG$ = FECHATEX$(FANTIG1)
        ANIOS% = Val(Mid$(FECA$, 7, 2)) - Val(Mid$(FATG$, 7, 2))
          If ANIOS% < 0 Then ANIOS% = ANIOS% + 100
        MESES% = Val(Mid$(FECA$, 4, 2)) - Val(Mid$(FATG$, 4, 2))
        DIAS% = Val(Mid$(FECA$, 1, 2)) - Val(Mid$(FATG$, 1, 2))
        If DIAS% < 0 Then
           DIAS% = DIAS% + 30
           MESES% = MESES% - 1
        End If
        If MESES% < 0 Then
           MESES% = MESES% + 12
           ANIOS% = ANIOS% - 1
        End If
        If ANIOS% < 1 Then ANIOS% = 0
        '
        YY$ = REBLA$
        Call REPLA(YY$, MKI$(LEGAX%), 1, 2)
        Call REPLA(YY$, MKI$(GRU%), 3, 2)
        Call REPLA(YY$, MKI$(Int(FEIN1)), 5, 2)
        Call REPLA(YY$, MKI$(Int(FANTIG1)), 7, 2)
        Call REPLA(YY$, MKI$(ANIOS%), 9, 2)
        Call REPLA(YY$, Chr$(MESES%), 11, 1)
        Call REPLA(YY$, Chr$(DIAS%), 12, 1)
        JJ& = JJ& + 1
        Call GRAREG("ANTICAL", YY$, JJ&)
      End If
    Next U&
    Call SETULTREG("ANTICAL", JJ&)
    Call FILESORT("ANTICAL", "", 1, 1, "ASC")
    Call COPYSTRU("ANTICAL", "LIANTIG")
    Call HEADERS("ANTICAL", "")
    Call HEADERS("ANTICAL", "Recálculo al: " + FECHATEX$(FECALCU))
    '
    Call GRACONTROL("SUELJOR", "FANTIGU", FECHATEX$(FECALCU))
    '
    Screen.MousePointer = 1
    '
    OPX% = COMALTER%("Fin del Proceso de Cálculo.\¿ Desea Emitir un Listado Impreso ?\\Si, Listar\No, Terminar")
    If OPX% <> 1 Then Exit Sub
    '
100 Call SELECHO("DATGREM1")
    VDEVU$ = TRIM$(VALACT1$("DATGREM1"))
    If VDEVU$ <> "" Then
      GRU% = Val(VDEVU$)
      Call FILTERFILE("ANTICAL", "LIANTIG", 3, MKI$(GRU%))
      Call CONECRUN("*LIANTIG", "Listado de Antiguedad")
      GoTo 100
    End If
    '
End Sub

Sub PLANIHAB()
   '
   Call CALDEREM
   Call CALDERET
   Call CALDEASI
   '
   JJ& = 0
   REBLA$ = REGBLAN$("PLANIHAB")
   '
   ' DETALLE DE REMUNERACIONES
   For KKU& = 1 To ULTREG&("DETRELIQ")
     aa2$ = REGLEIDO$("DETRELIQ", KKU&)
     LEGAX% = CVI(Left$(aa2$, 2))
     If LEGAX% > 0 Then
       ICOI% = CVI(Mid$(aa2$, 3, 2))
       CANTI = CVS(Mid$(aa2$, 5, 4))
       IEXEN# = CVD(Mid$(aa2$, 9, 8))
       IGRAV# = CVD(Mid$(aa2$, 25, 8))
       '
       YY$ = REBLA$
       Call REPLA(YY$, MKI$(LEGAX%), 1, 2)
       Call REPLA(YY$, MKI$(ICOI%), 3, 2)
       Call REPLA(YY$, MKS$(CANTI), 5, 4)
       Call REPLA(YY$, MKD$(IEXEN#), 65, 8)
       Call REPLA(YY$, MKD$(IGRAV#), 73, 8)
       '
       JJ& = JJ& + 1
       Call GRAREG("PLANIHAB", YY$, JJ&)
       '
     End If
   Next KKU&
   '
   ' DETALLE DE RETENCIONES
   For KKU& = 1 To ULTREG&("DETRTLIQ")
     aa2$ = REGLEIDO$("DETRTLIQ", KKU&)
     LEGAX% = CVI(Left$(aa2$, 2))
     If LEGAX% > 0 Then
       ICOI% = CVI(Mid$(aa2$, 3, 2))
       IRETE# = CVD(Mid$(aa2$, 5, 8))
       IAPOR# = CVD(Mid$(aa2$, 21, 8))
       '
       YY$ = REBLA$
       Call REPLA(YY$, MKI$(LEGAX%), 1, 2)
       Call REPLA(YY$, MKI$(ICOI%), 3, 2)
       Call REPLA(YY$, MKD$(IRETE#), 81, 8)
       Call REPLA(YY$, MKD$(IAPOR#), 89, 8)
       '
       JJ& = JJ& + 1
       Call GRAREG("PLANIHAB", YY$, JJ&)
       '
     End If
   Next KKU&
   '
   ' DETALLE DE ASIGNACIONES FAMILIARES
   For KKU& = 1 To ULTREG&("DETASLIQ")
     aa2$ = REGLEIDO$("DETASLIQ", KKU&)
     LEGAX% = CVI(Left$(aa2$, 2))
     If LEGAX% > 0 Then
       ICOI% = CVI(Mid$(aa2$, 3, 2))
       CANTI = CVS(Mid$(aa2$, 5, 4))
       IEXEN# = CVD(Mid$(aa2$, 9, 8))
       '
       YY$ = REBLA$
       Call REPLA(YY$, MKI$(LEGAX%), 1, 2)
       Call REPLA(YY$, MKI$(ICOI%), 3, 2)
       Call REPLA(YY$, MKS$(CANTI), 5, 4)
       Call REPLA(YY$, MKD$(IEXEN#), 65, 8)
       '
       JJ& = JJ& + 1
       Call GRAREG("PLANIHAB", YY$, JJ&)
       '
     End If
   Next KKU&
   '
   Call SETULTREG("PLANIHAB", JJ&)
   Call FILESORT("PLANIHAB", "", 1, 3, "ASC")
   Call HEADERS("PLANIHAB", "")
   Call HEADERS("PLANIHAB", "Liquidación: " + DESCRI_LIQ$)
   '
End Sub

Function TRXPRIN$(AAAAA$)

End Function

Private Sub ReLiCon_Click()
   ' resumen de liquidacion por concepto
   Call CALDEREM
   Call CALDERET
   Call CALDEASI
   '
   Dim COCON%(1024), DECON$(1024), CACON(1024), IMPO1#(1024), IMPO2#(1024), IMPO3#(1024), IMPO4#(1024)
   Dim DERE$(100), MOCAL%(100), PORET(100), PORAP(100), IMRET(100), IMAPO(100)
   CACONI% = 0: CARET% = 0
   '
   'CARGA VECTOR DE CONCEPTOS REMUNERATIVOS
   AA11$ = LOADFILE(LUDAT$ + "CONCEP.DAT")
   For KKU& = 1 To Len(AA11$) Step 128
     aa2$ = Mid$(AA11$, KKU&, 128)
     ICOI% = ICOI% + 1
     If ICOI% <= 100 Then
       CACONI% = CACONI% + 1
       COCON%(CACONI%) = ICOI%
       DECON$(CACONI%) = TRIM$(Mid$(aa2$, 1, 30))
       CACONI% = CACONI% + 1
       COCON%(CACONI%) = 100 + ICOI%
       DECON$(CACONI%) = "AJUSTE " + TRIM$(Mid$(aa2$, 1, 30))
     End If
   Next KKU&
   '
   AA11$ = LOADFILE(LUDAT$ + "RETLEG.DAT")
   ICOI% = 0
   For KKU& = 1 To Len(AA11$) Step 128
     aa2$ = Mid$(AA11$, KKU&, 128)
     ICOI% = ICOI% + 1
     If ICOI% <= 100 Then
       DERE$(ICOI%) = TRIM$(Mid$(aa2$, 1, 30))
       CACONI% = CACONI% + 1
       COCON%(CACONI%) = 200 + ICOI%
       DECON$(CACONI%) = TRIM$(Mid$(aa2$, 1, 30))
       CACONI% = CACONI% + 1
       COCON%(CACONI%) = 300 + ICOI%
       DECON$(CACONI%) = "AJUSTE " + TRIM$(Mid$(aa2$, 1, 30))
     End If
   Next KKU&
   '
   AA11$ = LOADFILE(LUDAT$ + "ASIGNA.DAT")
   ICOI% = 0
   For KKU& = 1 To Len(AA11$) Step 48
     aa2$ = Mid$(AA11$, KKU&, 48)
     ICOI% = ICOI% + 1
     If ICOI% <= 100 Then
       DERE$(ICOI%) = TRIM$(Mid$(aa2$, 1, 30))
       CACONI% = CACONI% + 1
       COCON%(CACONI%) = 400 + ICOI%
       DECON$(CACONI%) = TRIM$(Mid$(aa2$, 1, 30))
       CACONI% = CACONI% + 1
       COCON%(CACONI%) = 500 + ICOI%
       DECON$(CACONI%) = "AJUSTE " + TRIM$(Mid$(aa2$, 1, 30))
     End If
   Next KKU&
   '
   ' DETALLE DE REMUNERACIONES
   For KKU& = 1 To ULTREG&("DETRELIQ")
     aa2$ = REGLEIDO$("DETRELIQ", KKU&)
     LEGAX% = CVI(Left$(aa2$, 2))
     If LEGAX% > 0 Then
       ICOI% = CVI(Mid$(aa2$, 3, 2))
       CANTI = CVS(Mid$(aa2$, 5, 4))
       IEXEN# = CVD(Mid$(aa2$, 9, 8))
       IGRAV# = CVD(Mid$(aa2$, 25, 8))
       '
       For JI% = 1 To CACONI%
         If COCON%(JI%) = ICOI% Then
           CACON(JI%) = CACON(JI%) + CANTI
           IMPO1#(JI%) = IMPO1#(JI%) + IEXEN#
           IMPO2#(JI%) = IMPO2#(JI%) + IGRAV#
         End If
       Next JI%
       '
     End If
   Next KKU&
   '
   ' DETALLE DE RETENCIONES
   For KKU& = 1 To ULTREG&("DETRTLIQ")
     aa2$ = REGLEIDO$("DETRTLIQ", KKU&)
     LEGAX% = CVI(Left$(aa2$, 2))
     If LEGAX% > 0 Then
       ICOI% = CVI(Mid$(aa2$, 3, 2))
       IRETE# = CVD(Mid$(aa2$, 5, 8))
       IAPOR# = CVD(Mid$(aa2$, 21, 8))
       '
       For JI% = 1 To CACONI%
         If COCON%(JI%) = ICOI% Then
           IMPO3#(JI%) = IMPO3#(JI%) + IRETE#
           IMPO4#(JI%) = IMPO4#(JI%) + IAPOR#
         End If
       Next JI%
       '
     End If
   Next KKU&
   '
   ' DETALLE DE ASIGNACIONES FAMILIARES
   For KKU& = 1 To ULTREG&("DETASLIQ")
     aa2$ = REGLEIDO$("DETASLIQ", KKU&)
     LEGAX% = CVI(Left$(aa2$, 2))
     If LEGAX% > 0 Then
       ICOI% = CVI(Mid$(aa2$, 3, 2))
       CANTI = CVS(Mid$(aa2$, 5, 4))
       IEXEN# = CVD(Mid$(aa2$, 9, 8))
       '
       For JI% = 1 To CACONI%
         If COCON%(JI%) = ICOI% Then
           CACON(JI%) = CACON(JI%) + CANTI
           IMPO1#(JI%) = IMPO1#(JI%) + IEXEN#
         End If
       Next JI%
       '
     End If
   Next KKU&
   '
   JJ& = 0
   REBLA$ = REGBLAN$("RELICON")
   For JU& = 1 To CACONI%
     If Abs(CACON(JU&)) >= 0.005 Or Abs(IMPO1#(JU&)) > 0.005 Or Abs(IMPO2#(JU&)) > 0.005 Or Abs(IMPO3#(JU&)) > 0.005 Or Abs(IMPO4#(JU&)) > 0.005 Then
       YY$ = REBLA$
       IRUB% = 1 + Int(COCON%(JU&) / 200)
       Call REPLA(YY$, MKI$(IRUB%), 1, 2)
       Call REPLA(YY$, MKI$(COCON%(JU&)), 3, 2)
       Call REPLA(YY$, DECON$(JU&), 5, 30)
       Call REPLA(YY$, MKS$(CACON(JU&)), 35, 4)
       Call REPLA(YY$, MKD$(IMPO1#(JU&)), 65, 8)
       Call REPLA(YY$, MKD$(IMPO2#(JU&)), 73, 8)
       Call REPLA(YY$, MKD$(IMPO3#(JU&)), 81, 8)
       Call REPLA(YY$, MKD$(IMPO4#(JU&)), 89, 8)
       JJ& = JJ& + 1
       Call GRAREG("RELICON", YY$, JJ&)
     End If
   Next JU&
   Call SETULTREG("RELICON", JJ&)
   Call HEADERS("RELICON", "")
   Call HEADERS("RELICON", "Liquidación: " + DESCRI_LIQ$)
   '
   Call FINAL("*RELICON")
   '
End Sub

Sub CALDEREM()
   '
   Dim DECO$(100), HORA%(100), SURE%(100)
   JJ& = 0: CACON% = 0
   AA11$ = LOADFILE(LUDAT$ + "CONCEP.DAT")
   For KKU& = 1 To Len(AA11$) Step 128
     aa2$ = Mid$(AA11$, KKU&, 128)
     ICOI% = ICOI% + 1
     If ICOI% <= 100 Then
       DECO$(ICOI%) = TRIM$(Mid$(aa2$, 1, 30))
       HORA%(ICOI%) = Asc(Mid$(aa2$, 33, 1))
       SURE%(ICOI%) = Asc(Mid$(aa2$, 35, 1))
       CACON% = ICOI%
     End If
   Next KKU&
   '
   Call FILESORT("DETLIQUI", "DETLIQJ", 1, 2, "ASC")
   '
   AA11$ = LOADFILE(LUDAT$ + "DETLIQJ.DAT")
   For KKU& = 1 To Len(AA11$) Step 16
     aa2$ = Mid$(AA11$, KKU&, 16)
     LEGAX% = CVI(Left$(aa2$, 2))
     If LEGAX% > 0 Then
       ICOI% = CVI(Mid$(aa2$, 3, 2))
       If ICOI% > 0 Then
         If ICOI% < 200 Then
           QTID = CVS(Mid$(aa2$, 5, 4))
           QTHOR = 0
           If Abs(QTID) >= 0.005 Then
             If HORA%(ICOI% Mod 100) = 2 Then
               QTHOR = QTID
             End If
           End If
           '
           REMUD# = CVD(Mid$(aa2$, 9, 8))
           If Abs(REMUD#) >= 0.005 Then
             REGRA# = REMUD#: REXEN# = 0
             If SURE%(ICOI% Mod 100) = 0 Then
               REXEN# = REGRA#: REGRA# = 0
             End If
             YY$ = REBLA
             Call REPLA(YY$, MKI$(LEGAX%), 1, 2)
             Call REPLA(YY$, MKI$(ICOI%), 3, 2)
             Call REPLA(YY$, MKS$(QTHOR), 5, 4)
             Call REPLA(YY$, MKD$(REXEN#), 9, 8)
             Call REPLA(YY$, MKD$(REGRA#), 25, 8)
             JJ& = JJ& + 1
             Call GRAREG("DETRELIQ", YY$, JJ&)
           End If
         End If
       End If
     End If
   Next KKU&
   Call SETULTREG("DETRELIQ", JJ&)
   Call HEADERS("DETRELIQ", "")
   Call HEADERS("DETRELIQ", "Liquidación: " + DESCRI_LIQ$)
   '
End Sub

Sub CALDERET()
   '
   Dim DECO$(100), MOCAL%(100), PORET(100), PORAP(100), IMRET(100), IMAPO(100)
   JJ& = 0: CACON% = 0
   AA11$ = LOADFILE(LUDAT$ + "RETLEG.DAT")
   For KKU& = 1 To Len(AA11$) Step 128
     aa2$ = Mid$(AA11$, KKU&, 128)
     ICOI% = ICOI% + 1
     If ICOI% <= 100 Then
       DECO$(ICOI%) = TRIM$(Mid$(aa2$, 1, 30))
       MOCAL%(ICOI%) = Asc(Mid$(aa2$, 34, 1))
       PORET(ICOI%) = 0
       PORAP(ICOI%) = 0
       IMRET(ICOI%) = 0
       IMAPO(ICOI%) = 0
       If MOCAL%(ICOI%) = 2 Then
           PORET(ICOI%) = CVS(Mid$(aa2$, 39, 4))
           PORAP(ICOI%) = CVS(Mid$(aa2$, 43, 4))
         ElseIf MOCAL%(ICOI%) = 3 Then
           IMRET(ICOI%) = CVS(Mid$(aa2$, 39, 4))
           IMAPO(ICOI%) = CVS(Mid$(aa2$, 43, 4))
       End If
       CACON% = ICOI%
     End If
   Next KKU&
   '
   Call FILESORT("DETLIQUI", "DETLIQJ", 1, 2, "ASC")
   '
   AA11$ = LOADFILE(LUDAT$ + "DETLIQJ.DAT")
   For KKU& = 1 To Len(AA11$) Step 16
     aa2$ = Mid$(AA11$, KKU&, 16)
     LEGAX% = CVI(Left$(aa2$, 2))
     If LEGAX% > 0 Then
       ICOI% = CVI(Mid$(aa2$, 3, 2))
       If ICOI% >= 200 Then
         If ICOI% < 400 Then
           REMUD# = CVD(Mid$(aa2$, 9, 8))
           If Abs(REMUD#) >= 0.005 Then
             YY$ = REBLA
             Call REPLA(YY$, MKI$(LEGAX%), 1, 2)
             Call REPLA(YY$, MKI$(ICOI%), 3, 2)
             Call REPLA(YY$, MKD$(REMUD#), 5, 8)
             IMAPOD# = 0
             ICOJ% = ICOI% Mod 100
             If MOCAL%(ICOJ%) = 2 Or MOCAL(ICOJ%) = 3 Then ' PORCENTAJES
                 If PORET(ICOJ%) > 0 Then
                   IMAPOD# = REMUD# * PORAP(ICOJ%) / PORET(ICOJ%)
                 End If
             End If
             Call REPLA(YY$, MKD$(IMAPOD#), 21, 8)
             JJ& = JJ& + 1
             Call GRAREG("DETRTLIQ", YY$, JJ&)
           End If
         End If
       End If
     End If
   Next KKU&
   Call SETULTREG("DETRTLIQ", JJ&)
   Call HEADERS("DETRTLIQ", "")
   Call HEADERS("DETRTLIQ", "Liquidación: " + DESCRI_LIQ$)
   '
End Sub

Sub CALDEASI()
   '
   Dim DECO$(100), HORA%(100), SURE%(100)
   JJ& = 0: CACON% = 0
   AA11$ = LOADFILE(LUDAT$ + "ASIGNA.DAT")
   For KKU& = 1 To Len(AA11$) Step 48
     aa2$ = Mid$(AA11$, KKU&, 128)
     ICOI% = ICOI% + 1
     If ICOI% <= 100 Then
       DECO$(ICOI%) = TRIM$(Mid$(aa2$, 1, 30))
       CACON% = ICOI%
     End If
   Next KKU&
   '
   Call FILESORT("DETLIQUI", "DETLIQJ", 1, 2, "ASC")
   '
   AA11$ = LOADFILE(LUDAT$ + "DETLIQJ.DAT")
   For KKU& = 1 To Len(AA11$) Step 16
     aa2$ = Mid$(AA11$, KKU&, 16)
     LEGAX% = CVI(Left$(aa2$, 2))
     If LEGAX% > 0 Then
       ICOI% = CVI(Mid$(aa2$, 3, 2))
       If ICOI% > 400 Then
         If ICOI% < 600 Then
           QTID = CVS(Mid$(aa2$, 5, 4))
           '
           REMUD# = CVD(Mid$(aa2$, 9, 8))
           If Abs(REMUD#) >= 0.005 Then
             YY$ = REBLA
             Call REPLA(YY$, MKI$(LEGAX%), 1, 2)
             Call REPLA(YY$, MKI$(ICOI%), 3, 2)
             Call REPLA(YY$, MKS$(QTID), 5, 4)
             Call REPLA(YY$, MKD$(REMUD#), 9, 8)
             JJ& = JJ& + 1
             Call GRAREG("DETASLIQ", YY$, JJ&)
           End If
         End If
       End If
     End If
   Next KKU&
   Call SETULTREG("DETASLIQ", JJ&)
   Call HEADERS("DETASILIQ", "")
   Call HEADERS("DETASILIQ", "Liquidación: " + DESCRI_LIQ$)
   '
End Sub

Sub RECULIQ(NULI1%)
    '
    SCMP = Screen.MousePointer
    Screen.MousePointer = 11
    '
    LIQUI1& = Val(NULI1%)
    If LIQUI1& > 0 Then
      Call BLANARCH("ULTILIQ")
      Call CARGLIQ(NULI1%) 'CARGA LAS VARIABLES PUBLICAS
      XX$ = REGLEIDO$("ENCLIQUI", LIQUI1&)
      DENOLI$ = Mid$(XX$, 3, 30)
      GRUP% = Asc(Mid$(XX$, 33, 1))
      MESACU% = Asc(Mid$(XX$, 34, 1))
      ANOACU% = Asc(Mid$(XX$, 35, 1))
      TIPLIQ% = Asc(Mid$(XX$, 36, 1))
      ASIGFLI% = Asc(Mid$(XX$, 37, 1))
      FEPA% = CVI(Mid$(XX$, 38, 2))
      LUPA$ = TRIM$(Mid$(XX$, 40, 20))
      PEULDE$ = TRIM$(Mid$(XX$, 60, 16))
      FEDEPE% = CVI(Mid$(XX$, 76, 2))
      BCO1$ = TRIM$(Mid$(XX$, 78, 16))
      PERLIQ$ = TRIM$(Mid$(XX$, 94, 12))
      '
      Y$ = Mid$(XX$, 3) + String$(2, 0)
      Call REPLA(Y$, MKI$(NULI1%), 104, 2) 'LA MARCO COMO YA ARCHIVADA
      Call GRAREG("ULTILIQ", Y$, 1)
      '
      Combo2.TEXT = AJUSTI$(DENOLI$, 40) + FORMATNUM$(NULI1%, "I5")
      Combo2.Refresh
      '
      MESI$ = Mid$(Str$(MESACU%), 2): While Len(MESI$) < 2: MESI$ = "0" + MESI$: Wend
      ANI$ = Mid$(Str$(ANOACU%), 2): While Len(ANI$) < 2: ANI$ = "0" + ANI$: Wend
      ARCHAC$ = "DTL" + MESI$ + ANI$ + ".DAT"
      TTXX$ = LOADFILE$(LUDAT$ + ARCHAC$)
      JJ& = 0
      For KU& = 1 To Len(TTXX$) Step 24
        TTYY$ = Mid$(TTXX$, KU&, 24)
        If CVI(Left$(TTYY$, 2)) = NULI1% Then
          JJ& = JJ& + 1
          TTZZ$ = Mid$(TTYY$, 3, 16)
          Call GRAREG("DETLIQUI", TTZZ$, JJ&)
        End If
      Next KU&
      Call SETULTREG("DETLIQUI", JJ&)
      '
    End If
    Screen.MousePointer = SCMP
    '
End Sub

Sub LISINFHOR()
     '
     Call CARGLIQACT          ' datos de liquidacion activa
     Call CARGAFORMULAS       ' CARGAR CONCEPTOS y RETENCIONES
     JJ& = 0
     REBLA$ = REGBLAN$("LINFHOR")
     '
1000 For KU& = 1 To ULTREG&("PADREMP")
       PP0$ = REGLEIDO$("PADREMP", KU&)
       LEGAI% = CVI(Left$(PP0$, 2))
       Call CARGEMP(LEGAI%)
       If STAT_LEG% > 0 Then  ' ESTADO ACTIVO
         If GRUP_LEG% = GRUP_LIQ% Then
           '
           For CONI% = 1 To NUCON%
             If ESTCON%(CONI%) < 1 Then GoTo 8230 ' DESCARTO SUSPENDIDOS
             If FORCON%(CONI%) < 1 Then GoTo 8230 ' DESCARTO CONCEPTOS DE CALCULO
             If QLCON%(CONI%) <> 0 Then If QLCON%(CONI%) <> TIPO_LIQ% Then GoTo 8230
             If CARCON%(CONI%) = 3 Or CARCON%(CONI%) = GRUP_LIQ% Then GoTo 8220 ' MANDATORIO
             For X% = 1 To 96 Step 2
               X1% = CVI(Mid$(COPC_LEG$, X%))
               If X1% = CONI% Then GoTo 8220           ' PREASIGNADO
             Next X%
             GoTo 8230
             '
8220         YY$ = REBLA$
             Call REPLA(YY$, MKI$(LEGAI%), 1, 2)
             Call REPLA(YY$, MKI$(CONI%), 3, 2)
             Call REPLA(YY$, ECOARCH$("ECONCEP", MKI$(CONI%)), 5, 30)
             If FORCON%(CONI%) = 2 Then
                 Call REPLA(YY$, String$(10, "."), 37, 10)
               ElseIf FORCON%(CONI%) = 3 Then
                 Call REPLA(YY$, String$(10, "."), 49, 10)
               Else
                 Call REPLA(YY$, String$(10, "."), 61, 12)
             End If
             Call REPLA(YY$, String$(64, "."), 75, 64)
             JJ& = JJ& + 1
             Call GRAREG("LINFHOR", YY$, JJ&)
8230       Next CONI%
           '
           For CONI% = 1 To NURET%
             If ESTRET%(CONI%) < 1 Then GoTo 8270 ' DESCARTO SUSPENDIDOS
             If FORRET%(CONI%) <> 1 Then GoTo 8270 ' DESCARTO CONCEPTOS DE CALCULO
             If QLRET%(CONI%) <> 0 Then If QLRET%(CONI%) <> TIPO_LIQ% Then GoTo 8270
             If CARRET%(CONI%) = 3 Or CARRET%(CONI%) = GRUP_LIQ% Then GoTo 8260 ' MANDATORIO
             For X% = 1 To 96 Step 2
               X1% = CVI(Mid$(COPC_LEG$, X%))
               If X1% = 200 + CONI% Then GoTo 8260         ' PREASIGNADO
             Next X%
             GoTo 8270
             '
8260         YY$ = REBLA$
             Call REPLA(YY$, MKI$(LEGAI%), 1, 2)
             Call REPLA(YY$, MKI$(200 + CONI%), 3, 2)
             Call REPLA(YY$, ECOARCH$("ECONCEP", MKI$(200 + CONI%)), 5, 30)
             If FORRET%(CONI%) = 1 Then
               Call REPLA(YY$, String$(10, "."), 61, 12)
             End If
             Call REPLA(YY$, String$(64, "."), 75, 64)
             JJ& = JJ& + 1
             Call GRAREG("LINFHOR", YY$, JJ&)
8270       Next CONI%
           '
           For KKII% = 1 To 4
             YY$ = REBLA$
             Call REPLA(YY$, MKI$(LEGAI%), 1, 2)
             Call REPLA(YY$, MKI$(0), 3, 2)
             Call REPLA(YY$, String$(30, "."), 5, 30)
             Call REPLA(YY$, String$(10, "."), 37, 10)
             Call REPLA(YY$, String$(10, "."), 49, 10)
             Call REPLA(YY$, String$(10, "."), 61, 12)
             Call REPLA(YY$, String$(64, "."), 75, 64)
             JJ& = JJ& + 1
             Call GRAREG("LINFHOR", YY$, JJ&)
           Next KKII%
           '
         End If
       End If
     Next KU&
     '
     Call SETULTREG("LINFHOR", JJ&)
     Call HEADERS("LINFHOR", "")
     Call HEADERS("LINFHOR", "Corresponde a: " + DESCRI_LIQ$)
     Call FINAL("*LINFHOR")
     '
End Sub

Sub TRANSFDGI()
'
    LU$ = LUDAT$
    Dim LEGA%(512), REMGRA#(512), RENOGR#(512), ASFAM#(512)
    Dim SUAD#(512), RSAC#(512), HOEX#(512), RVAC#(512)

60 'determinacion de fecha actual
62  HO1% = HOY(HO$)
    '
105 Call SELECHO("SELMESA")
    CORRESP$ = VALACT1$("SELMESA")
    If CORRESP$ = "" Then
        Exit Sub
    End If
    '
    MES$ = Left$(UCase$(CORRESP$), 3)
    If MES$ = "ENE" Then MES$ = "01": GoTo 106
    If MES$ = "FEB" Then MES$ = "02": GoTo 106
    If MES$ = "MAR" Then MES$ = "03": GoTo 106
    If MES$ = "ABR" Then MES$ = "04": GoTo 106
    If MES$ = "MAY" Then MES$ = "05": GoTo 106
    If MES$ = "JUN" Then MES$ = "06": GoTo 106
    If MES$ = "JUL" Then MES$ = "07": GoTo 106
    If MES$ = "AGO" Then MES$ = "08": GoTo 106
    If MES$ = "SET" Then MES$ = "09": GoTo 106
    If MES$ = "OCT" Then MES$ = "10": GoTo 106
    If MES$ = "NOV" Then MES$ = "11": GoTo 106
    If MES$ = "DIC" Then MES$ = "12": GoTo 106
    If MES$ = "SAC" Then MES$ = "13": GoTo 106
    If MES$ = "AGU" Then MES$ = "13": GoTo 106
    If MES$ = "VAC" Then MES$ = "14": GoTo 106
    Call MENSERR(24, "Periodo Mensual no Permitido")
    GoTo 105

106 CORR1$ = MES$ + Right$(CORRESP$, 2)

110 VALO# = TOPE#
    TOPE1$ = NUMSTRI$(VALO#, 12, 2, 0)
    TOPE2$ = LTrim$(RTrim$(TOPE1$))

111 N0% = FreeFile
    ASALID$ = LUCOM$ + "TRASIJP.TXT"
    N0% = FILEOPEN%(ASALID$, 2, 192)
    If LOF(N0%) > 0 Then
        Call MENSERR(24, "Archivo de Transferencia Existente")
        A% = ALTERNA%("Abandonar Operacion\Sobre-escribir Archivo\Sobre-escribir con Respaldo Archivo Actual")
        On A% GoTo 112, 113, 113
        '
112     GoTo 3900
113     Close #N0%
        If A% = 3 Then
           FileCopy LUCOM$ + "TRASIJP.TXT", LUCOM$ + ARESP$
           Call MENSERR(24, "Archivo de Transferencia fue Respaldado como " + ARESP$)
        End If
        Kill LUCOM$ + "TRASIJP.TXT"
        GoTo 111
    End If
    Close #N0%
    '
    N0% = FILEOPEN%(ASALID$, 2, 192)

115 If ULTREG&("RESIJP") > 0 Then
       Call CIERRARCH("RESIJP")
       Kill LU$ + "RESIJP.DAT"
       GoTo 115
    End If
    '
    Call HEADERS("RESIJP", "")
    Call HEADERS("RESIJP", "Corresponde a: " + CORRESP$)

120 Call MENSERR(24, "DETERMINANDO CONCEPTOS REMUNERATIVOS GRAVADOS")
    Dim SURET%(100)
    For I% = 1 To 100
       II& = I%
       X$ = REGLEIDO$("CONCEP", II&): F1$ = Mid$(X$, 1, 30): F2$ = Mid$(X$, 31, 8)
       If Asc(F1$) < 1 Then GoTo 140
       SURET%(I%) = Asc(Mid$(F2$, 5, 1))
    Next I%

140 Call MENSERR(24, "RECORRIENDO ARCHIVO DE LIQUIDACION DEL MES DE " + LTrim$(RTrim$(CORRESP$)))
    Archi$ = "DTL" + CORR1$ ' + ".DAT"
    ARESP$ = "SJP" + CORR1$ ' + ".DAT"
    FIN& = ULTREG&(Archi$)
    
    For III& = 1 To FIN&
       X$ = REGLEIDO$(Archi$, III&)
       J0$ = Mid$(X$, 1, 2): J1$ = Mid$(X$, 3, 2): J2$ = Mid$(X$, 5, 2)
       J3$ = Mid$(X$, 7, 4): J4$ = Mid$(X$, 11, 8): J5$ = Mid$(X$, 19, 6)
       Call TRACE(20, III&, FIN&)
       If CVI(J0$) = 0 Then Call TRACE(20, FIN&, FIN&): GoTo 150
   
       LEGA1% = CVI(J1$)
       CODI% = CVI(J2$)
       IMPORT# = (Int(100 * CVD(J4$) + 0.5)) / 100
   
       For J% = CONTA% To 1 Step -1
          If LEGA%(J%) = LEGA1% Then GoTo 142
       Next J%
   
       CONTA% = CONTA% + 1
       J% = CONTA%
       LEGA%(J%) = LEGA1%
   
142    If CODI% >= 400 Then
         If CODI% <= 599 Then
           ASFAM#(J%) = ASFAM#(J%) + IMPORT#
           GoTo 149
         End If
       End If
       
       If CODI% > 0 Then
         If CODI% <= 199 Then
           CODI% = CODI% Mod 100
           If CODI% <> 72 Then
             If SURET%(CODI%) > 0 Then
              REMGRA#(J%) = REMGRA#(J%) + IMPORT#
            Else
              RENOGR#(J%) = RENOGR#(J%) + IMPORT#
             End If
             If CODI% = 73 Or CODI% = 91 Or CODI% = 93 Or CODI% = 95 Or CODI% = 97 Or CODI% = 98 Then
             RSAC#(J%) = RSAC#(J%) + IMPORT#
           ElseIf CODI% = 2 Or CODI% = 3 Or CODI% = 5 Or CODI% = 6 Then
             HOEX#(J%) = HOEX#(J%) + IMPORT#
           ElseIf CODI% = 26 Or CODI% = 27 Or CODI% = 28 Or CODI% = 65 Or CODI% = 66 Or CODI% = 67 Then
             RVAC#(J%) = RVAC#(J%) + IMPORT#
           ElseIf CODI% <> 82 Then
             SUAD#(J%) = SUAD#(J%) + IMPORT#
             End If
           End If
         End If
       End If
149 Next III&

150 Call MENSERR(24, "GRABANDO ARCHIVO DE TRANSFERENCIA")

    K& = 0
    For I% = 1 To CONTA%
        II& = I%: CONTA1& = CONTA%
        Call TRACE(20, II&, CONTA1&)
        LEGA1% = LEGA%(I%)
        APYNOM$ = ""
        FIN& = ULTREG&("PADRON")
        For JJ& = 1 To FIN&
          X$ = REGLEIDO$("PADRON", JJ&)
          P1$ = Mid$(X$, 1, 2): P2$ = Mid$(X$, 3, 30)
          AA = CVI(P1$)
          If CVI(P1$) = LEGA1% Then
            APYNOM$ = P2$
            GoTo 155
          End If
        Next JJ&

155     FIN& = ULTREG&("DATEMDGI")
        For JJ& = 1 To FIN&
            '
            X$ = REGLEIDO$("DATEMDGI", JJ&)
            X0$ = X$
            X1$ = Mid$(X$, 1, 2): X3$ = Mid$(X$, 3, 13): X4$ = Mid$(X$, 16, 2)
            X5$ = Mid$(X$, 18, 1): X6$ = Mid$(X$, 19, 1): X7$ = Mid$(X$, 20, 1)
            X8$ = Mid$(X$, 22, 4): X9$ = Mid$(X$, 26, 4): X10$ = Mid$(X$, 30, 2)
            X11$ = Mid$(X$, 32, 2)
            '
            If CVI(X1$) = LEGA1% Then
               '
               XXX3$ = X3$
               CUIL$ = Left$(REPLACAR$(XXX3$, "-", ""), 11)
               ZOGE$ = "01" ' X4$
               Rem   COAC$=X5$:IF VAL(COAC$)<1 THEN
               COAC$ = "49"
               GFAM$ = Mid$(Str$(Asc(X6$)), 2): While Len(GFAM$) < 2: GFAM$ = "0" + GFAM$: Wend
               ADGF$ = Mid$(Str$(Asc(X7$)), 2): While Len(ADGF$) < 2: ADGF$ = "0" + ADGF$: Wend
                   VALO# = CVS(X8$) 'VALO# = CVSMBF(X8$)
               OSOC$ = NUMSTRI0$(VALO#, 6, 0, 0)
                   VALO# = CVS(X9$) 'VALO# = CVSMBF(X9$)
               REDU$ = NUMSTRI0$(VALO#, 4, 2, 0): REDU$ = Left$(REDU$, 2)
                   VALO# = CVS(X10$) 'VALO# = CVSMBF(X10$)
               APAD$ = NUMSTRI0$(VALO#, 4, 2, 0): APAD$ = Left$(APAD$, 2)
                   VALO# = CVI(X11$)
               REBA$ = NUMSTRI0$(VALO#, 3, 0, 0)
               '
               CONYU$ = TRIM$(Str$(Asc(Mid$(X0$, 36, 1))))
               While Len(CONYU$) < 1: CONYU$ = "0" + CONYU$: Wend
               HIJOS$ = TRIM$(Str$(Asc(Mid$(X0$, 37, 1))))
               While Len(HIJOS$) < 2: HIJOS$ = "0" + HIJOS$: Wend
               CODSI$ = TRIM$(Mid$(X0$, 38, 1))
               If Val(CODSI$) <> 1 Then CODSI$ = "01"
               While Len(CODSI$) < 2: CODSI$ = "0" + CODSI$: Wend
               CODCO$ = TRIM$(Mid$(X0$, 39, 1))
               If Val(CODCO$) <> 1 Then CODCO$ = "01"
               While Len(CODCO$) < 2: CODCO$ = "0" + CODCO$: Wend
               MODCO$ = TRIM$(Mid$(X0$, 40, 3))
               If Val(MODCO$) < 1 Then MODCO$ = "008"
               While Len(MODCO$) < 3: MODCO$ = "0" + MODCO$: Wend
               COCOL$ = TRIM$(Mid$(X0$, 43, 1))
               While Len(COCOL$) < 1: COCOL$ = "0" + COCOL$: Wend
               '
               FIL1$ = String$(24, "0")
               '
               VALO# = REMGRA#(I%) + RENOGR#(I%)
               REMU$ = NUMSTRI0$(VALO#, 8, 2, 0)
                 REMU$ = Left$(REMU$, 6) + "." + Mid$(REMU$, 7)
               VALO# = REMGRA#(I%) ': If VALO# > TOPE# Then VALO# = TOPE#
               REIM$ = NUMSTRI0$(VALO#, 8, 2, 0)
                 REIM$ = Left$(REIM$, 6) + "." + Mid$(REIM$, 7)
               VALO# = ASFAM#(I%)
               AFAM$ = NUMSTRI0$(VALO#, 8, 2, 0)
                 AFAM$ = Left$(AFAM$, 6) + "." + Mid$(AFAM$, 7)
               VALO# = RENOGR#(I%)
               REMU5$ = NUMSTRI0$(VALO#, 8, 2, 0)
                 REMU5$ = Left$(REMU5$, 6) + "." + Mid$(REMU5$, 7)
                   '
               APVOL$ = "000000.00"
               ADIOS$ = "000000.00"
               EXESS$ = "000000.00"
               EXEOS$ = "000000.00"
               '
               RECORD$ = String$(354, "0")
               Call REPLA(RECORD$, CUIL$, 1, 11)
               Call REPLA(RECORD$, APYNOM$, 12, 30)
               Call REPLA(RECORD$, CONYU$, 42, 1)
               Call REPLA(RECORD$, HIJOS$, 43, 2)
               Call REPLA(RECORD$, CODSI$, 45, 2)
               Call REPLA(RECORD$, CODCO$, 47, 2)
               Call REPLA(RECORD$, COAC$, 49, 2)
               Call REPLA(RECORD$, ZOGE$, 51, 2)
               Call REPLA(RECORD$, APAD$, 56, 2)
               Call REPLA(RECORD$, MODCO$, 58, 3)
               Call REPLA(RECORD$, OSOC$, 61, 6)
               Call REPLA(RECORD$, ADGF$, 67, 2)
               Call REPLA(RECORD$, REMU$, 69, 9)
               Call REPLA(RECORD$, REIM$, 78, 9)
               Call REPLA(RECORD$, AFAM$, 87, 9)
               Call REPLA(RECORD$, APVOL$, 96, 9)
               Call REPLA(RECORD$, ADIOS$, 105, 9)
               Call REPLA(RECORD$, EXESS$, 114, 9)
               Call REPLA(RECORD$, EXEOS$, 123, 9)
               Call REPLA(RECORD$, "Buenos Aires - Lomas del Mirador                 ", 132, 50)
       
               Call REPLA(RECORD$, REIM$, 182, 9)
               Call REPLA(RECORD$, REIM$, 191, 9)
               Call REPLA(RECORD$, REIM$, 200, 9)
               '
               Call REPLA(RECORD$, "00", 209, 2)
               Call REPLA(RECORD$, "0", 211, 1)
               Call REPLA(RECORD$, "000000000", 212, 9)
               Call REPLA(RECORD$, "1", 221, 1)
               Call REPLA(RECORD$, "000000000", 222, 9)
               Call REPLA(RECORD$, "1", 231, 1)
               Call REPLA(RECORD$, "01", 232, 2) ' SITUACION DE REVISTA
               Call REPLA(RECORD$, "01", 234, 2) ' DIA INICIO SITUACION REVISTA
               Call REPLA(RECORD$, String$(62, "0"), 236, 62)
               '
               VALO# = SUAD#(I%)
               If VALO# < 0.005 Then
                 GoTo 159
               End If
               VALEX$ = NUMSTRI0$(VALO#, 8, 2, 0)
               VALEX$ = Left$(VALEX$, 6) + "." + Mid$(VALEX$, 7)
               Call REPLA(RECORD$, VALEX$, 244, 9)
               '
               REMUTOT$ = Mid$(REIM$, 1, 6) + Mid$(REIM$, 8) ' MODIFICADO
               'SUAD1# = Val(Mid$(RECORD$, 200, 9))
               SUAD1# = Val(REMUTOT$)
               VALO# = RSAC#(I%)
               VALEX$ = NUMSTRI0$(VALO#, 8, 2, 0)
               VALEX$ = Left$(VALEX$, 6) + "." + Mid$(VALEX$, 7)
               Call REPLA(RECORD$, VALEX$, 253, 9)
               SUAD1# = SUAD1# - Val(VALEX$)
               '
               VALO# = HOEX#(I%)
               VALEX$ = NUMSTRI0$(VALO#, 8, 2, 0)
               VALEX$ = Left$(VALEX$, 6) + "." + Mid$(VALEX$, 7)
               Call REPLA(RECORD$, VALEX$, 262, 9)
               SUAD1# = SUAD1# - Val(VALEX$)
               '
               VALO# = RVAC#(I%)
               VALEX$ = NUMSTRI0$(VALO#, 8, 2, 0)
               VALEX$ = Left$(VALEX$, 6) + "." + Mid$(VALEX$, 7)
               Call REPLA(RECORD$, VALEX$, 280, 9)
               SUAD1# = SUAD1# - Val(VALEX$)
               '
               VALO# = SUAD1#
               VALEX$ = NUMSTRI0$(VALO#, 8, 2, 0)
               VALEX$ = Left$(VALEX$, 6) + "." + Mid$(VALEX$, 7)
               Call REPLA(RECORD$, VALEX$, 244, 9)
               '
               Call REPLA(RECORD$, "000030.00", 289, 9) ' DIAS TRABAJADOS
               '
               Call REPLA(RECORD$, REMU5$, 298, 9)
               Call REPLA(RECORD$, COCOL$, 307, 1)
               TIPOP$ = Mid$(X0$, 50, 1)
               Call REPLA(RECORD$, TIPOP$, 317, 1)
               '
               Print #N0%, RECORD$
               '
               K& = K& + 1
               R1$ = CUIL$ ' LSet R1$ = CUIL$
               R2$ = MKI$(LEGA%(I%)) ' LSet R2$ = MKI$(LEGA%(I%))
               R4$ = MKD$(REMGRA#(I%)) ' LSet R4$ = MKDMBF$(REMGRA#(I%))
                   If REMGRA#(I%) > TOPE# Then R4$ = MKD$(TOPE#) 'LSet R4$ = MKDMBF$(TOPE#)
               R5$ = MKD$(ASFAM#(I%)) 'LSet R5$ = MKDMBF$(ASFAM#(I%))
               TTXX$ = R1$ + R2$ + R4$ + R5$
               Call GRAREG("RESIJP", TTXX$, K&)
       
               GoTo 159
            End If
        Next JJ&

    Call MENSERR(24, "No Figuran Datos de Legajo" + Str$(LEGA1%))
159 Next I%

    If N0% > 0 Then
       Close #N0%
       N0% = 0
    End If
    A% = ALTERNA%("Copia Archivo Transferencia a C:DGISIJP")
    If A% = 1 Then
       On Error Resume Next
       FileCopy LUCOM$ + "TRASIJP.TXT", "C:\DGISIJP\" + "TRASIJP.TXT"
       On Error GoTo 0
    End If

3900 On Error Resume Next
     Close N0%
     On Error GoTo 0
     Call CIERRARCH("*.*")
     '
End Sub

