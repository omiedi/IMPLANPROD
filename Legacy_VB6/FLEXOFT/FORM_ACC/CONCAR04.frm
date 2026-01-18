VERSION 5.00
Begin VB.Form CONCAR04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "C.R.P. - Programación de Carga de Máquinas"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5325
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "B.D. de Operaciones Pendientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   210
      TabIndex        =   12
      Top             =   105
      Width           =   4005
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2685
         TabIndex        =   13
         Top             =   525
         Width           =   2745
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
            Picture         =   "CONCAR04.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Operaciones Proyectadas por O/F"
            Top             =   0
            Width           =   590
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Operaciones Proyec- tadas por O/Fabric."
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
            TabIndex        =   14
            Top             =   90
            Width           =   1800
         End
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004040&
      Height          =   3900
      Left            =   4410
      ScaleHeight     =   3840
      ScaleWidth      =   1215
      TabIndex        =   8
      Top             =   0
      Width           =   1275
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
         Height          =   635
         Left            =   105
         Picture         =   "CONCAR04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   180
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   9
         Top             =   2730
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "CONCAR04.frx":0454
         Top             =   3045
         Width           =   1515
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Consultas y Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   210
      TabIndex        =   1
      Top             =   1575
      Width           =   4005
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   2730
         TabIndex        =   5
         Top             =   1050
         Width           =   2790
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
            Picture         =   "CONCAR04.frx":2376
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Operaciones Pendientes por Máquina o Equipo"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Por Máquina o Equi- po de Producción"
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
            Width           =   1800
         End
      End
      Begin VB.PictureBox Picture8 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   2685
         TabIndex        =   2
         Top             =   525
         Width           =   2745
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
            Left            =   0
            Picture         =   "CONCAR04.frx":2680
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Operaciones Pendientes por Orden de Fabricación"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Pendientes por Or- den de Fabricación"
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
            TabIndex        =   4
            Top             =   60
            Width           =   1800
         End
      End
   End
   Begin VB.CommandButton Command34 
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
      Height          =   645
      Left            =   1995
      Picture         =   "CONCAR04.frx":298A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
      Top             =   2835
      Width           =   750
   End
End
Attribute VB_Name = "CONCAR04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command15_Click()
   Command15.Enabled = False
   Call FINAL("*SECAREQ")
   Command15.Enabled = True
End Sub

Private Sub Command26_Click()
   Command26.Enabled = False
   Call GENEOPERFAP 'AGREGUE PARA GENERAR OPERFAH Y OPERFAP
   Call FINAL("*SECAROF")
   Command26.Enabled = True
End Sub

Private Sub Command34_Click()
    Call ACCDIR("CARGMAQ")
End Sub

Private Sub Command7_Click()
     Command7.Enabled = False
     '
     HO% = HOY(HOS$)
     LU$ = LUDAT$
     '
     TITU$ = "OPERACIONES PROYECTADAS POR ORDEN DE FABRICACION"

100  NROOF$ = ""

     Screen.MousePointer = 11
     Call COPYSTRU("ECORDEF", "ECORDEP")
     '
     Call ABREORFAB
     SQLX$ = "SELECT NumeOrFa, Referen, IdSubOr, CanProy, Cod_Inte FROM ORDEFABR "
     SQLX$ = SQLX$ + "WHERE StatuOrf < 3 "
     SQLX$ = SQLX$ + "AND CanApro < CanProy "
     SQLX$ = SQLX$ + "order by IdSubOr ASC"
     Set ORFANUTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     On Error Resume Next
     With ORFANUTMP
       .MoveFirst
       If Err Then
         Call MENSERR(24, "No Hay O/F's que Puedan Consultarse.")
         Screen.MousePointer = 1
         On Error GoTo 0
         GoTo 99
       End If
       On Error GoTo 0
       '
       JJ& = 0
       Do While Not .EOF
           ZZ$ = REGBLAN$("!ECORDEP")
           Call REPLA(YY$, !IdSubOr, 1, 12)
           Call REPLA(YY$, !Referen, 13, 52)
           JJ& = JJ& + 1
           Call GRAREG("!ECORDEP", YY$, JJ&)
       .MoveNext
       Loop
     End With
     Call SETULTREG("!ECORDEP", JJ&)
     Call CIERRARCH("!ECORDEP")
     Screen.MousePointer = 1
     '
1405 Call FRESHECHO("ECORDEP")
     If ULTREG&("ECORDEP") < 1 Then
         Call MENSERR(24, "No hay Ordenes de Fabricación Pendientes.")
         GoTo 99
     End If
     '
     Call SELECHO("ECORDEP/Ordenes de Trabajo en Curso")
1410 NORFA$ = AJUSTI$(VALACT1$("ECORDEP"), 12)
     If TRIM$(NORFA$) = "" Then
       GoTo 99
     End If
     '
     Screen.MousePointer = 11
     '
     SQLX$ = " SELECT * FROM ORDEFABR "
     SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "' "
     '
     Set CARGAOBJTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     With CARGAOBJTMP
       On Error Resume Next
      .MoveFirst
       On Error GoTo 0
     NROOFL = !NUMEoRFA
     Call REPLA(VDEF$, MKS$(NROOFL), 1, 4)
     Call REPLA(VDEF$, !Referen, 5, 64)
     Call REPLA(VDEF$, MKI$(CVINT(!FechGen)), 69, 2)
     Call REPLA(VDEF$, !Unidad, 73, 4)
     Call REPLA(VDEF$, MKD$(!CanProy), 89, 8)
     Call REPLA(VDEF$, !Destino, 101, 8)
     Call REPLA(VDEF$, MKI$(CVINT(!FechEnSol)), 109, 2)
     Call REPLA(VDEF$, Chr$(!Priorid), 111, 1)
     Call REPLA(VDEF$, Chr$(!StatuOrf), 112, 1)
     Call REPLA(VDEF$, MKI$(CVINT(!FechLan)), 113, 2)
     Call REPLA(VDEF$, MKI$(CVINT(!FechCierr)), 119, 2)
     End With
     '
1412 YY$ = OBJPLA$("ORFALAN/Datos de la Orden de Fabricación", VDEF$)
     If YY$ = "" Then GoTo 1405
     '
     Screen.MousePointer = 1
     '
     Screen.MousePointer = 11
     '
     Call FRESHECHO("!OPERINF") 'nuevo
     'ACA OBTENGO LOS REGISTROS DE OPERFAB, CUYO IdSubOr = NROOF$
     'Luego cambio las variable por los campos correspondientes
     '
     Call ABREORFAB
     SQLX$ = " SELECT * FROM OPERFAB "
     SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "' "
     SQLX$ = SQLX$ + "AND StatOper < 8 "                      ' no marca cumplidos"
     SQLX$ = SQLX$ + " ORDER BY NumeOper "
     '
     Set FABRTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     With FABRTMP
       On Error Resume Next
       .MoveFirst
       If Err Then
         Call MENSERR(24, "No hay Operaciones Programadas\para esta Orden de Fabricación.")
         On Error GoTo 0
         'GoTo 1405
       End If
       On Error GoTo 0
       JJ& = 0
       YY$ = REGBLAN("OPERFAB")
       Do While (.EOF = False)
          CAPROD# = !CantProg      'Cantidad Programada
          CAREA# = !CantReal       'Cantidad realizada
          SALDO# = CAPROD# - CAREA#: If SALDO# < 0 Then SALDO# = 0
          '
          If SALDO# > 0.05 Then
            NOPE% = !NumeOper      'numero de operación
            DOPE$ = !DescOper      'Descripcion de la operacion
            MAQUI$ = !Maquina      'Máquina
            CIXII% = !COD_INTE
            '
            Call REPLA(YY$, MKI$(NOPE%), 15, 2)
            Call REPLA(YY$, DOPE$, 17, 48)
            Call REPLA(YY$, MAQUI$, 73, 6)
            Call REPLA(YY$, MKD$(CAPROD#), 79, 8)
            Call REPLA(YY$, MKD$(CAREA#), 87, 8)
            '
            TPREP = !TiemPrep
            TFIJO = !TiemFijo
            TVARI = !TiemVari    'TPROD - TFIJO: If TVARI < 0 Then TVARI = 0
            TDEMO = !TiemDemo
            TMANT = !TiemMHer
            TLIMP = !TiemLimp
            TOTRO = !TiemOtro
            TPROTO = TDEMO + TMANT + TLIMP + TOTRO
            '
            Call REPLA(YY$, MKS$(TPREP), 103, 4)
            Call REPLA(YY$, MKS$(TFIJO), 107, 4)
            Call REPLA(YY$, MKS$(TVARI), 111, 4)
            Call REPLA(YY$, MKS$(TPROTO), 127, 4)
            Call REPLA(YY$, MKS$(!CantiOps), 147, 4)
            JJ& = JJ& + 1
            Call GRAREG("!OPERFAB", YY$, JJ&)
            '
          End If
       .MoveNext
       Loop
     End With
        
     Call SETULTREG("!OPERFAB", JJ&)
     Call CIERRARCH("!OPERFAB")
     '
     Screen.MousePointer = 1
     '
     VTB% = VENTABU%("ACOPEND/TITUPAN=Tiempor Básicos de Operación - " + NORFA$)
     If VTB% >= 0 Then
       Call ABREORFAB
       ABUS$ = "IdSubOr = '" & NORFA$ & "' AND StatOper < 8 "
       With OperFab
         On Error Resume Next
         .FindFirst ABUS$
35       If .NoMatch = False Then
           .Delete
           .FindNext ABUS$
           GoTo 35
         End If
         '
         For U& = 1 To ULTREG&("!OPERFAB")
           YY$ = REGLEIDO$("!OPERFAB", U&)
           NOPE% = CVI(Mid$(YY$, 15, 2))
           DOPE$ = Mid$(YY$, 17, 48)
           MAQUI$ = Mid$(YY$, 73, 6)
           CAPROD# = CVD(Mid$(YY$, 79, 8))
           CAREA# = CVD(Mid$(YY$, 87, 8))
           TPREP = CVS(Mid$(YY$, 103, 4))
           TFIJO = CVS(Mid$(YY$, 107, 4))
           TVARI = CVS(Mid$(YY$, 111, 4))
           TOTRO = CVS(Mid$(YY$, 127, 4))
           CANOP = CVS(Mid$(YY$, 147, 4))
           '
           TDEMO = 0
           TMANT = 0
           TLIMP = 0
           '
           .AddNew
           !IdSubOr = NORFA$
           !COD_INTE = CIXII%
           !Cod_Exte = CODEXT$(CIXII%)
           !Despcrip = DESCRIT0$(CIXII%)
           !NumeOper = NOPE%
           !DescOper = DOPE$
           !Maquina = MAQUI$
           !DescMaq = ECOARCH$("PADMAQ", MAQUI$)
           !EquiAux = ""
           !CantiOps = CANOP
           !CantProg = CAPROD#
           !CantReal = CAREA#
           !CantPend = CAPROD# - CAREA#
           !TiemPrep = TPREP
           !TiemFijo = TFIJO
           !TiemVari = TVARI
           !TiemDemo = TDEMO
           !TiemMHer = TMANT
           !TiemLimp = TLIMP
           !TiemOtro = TOTRO
           !TiemTota = TPREP + TFIJO + TVARI + TOTRO
           !TiemEjec = 0
           !TiemPend = 0
           !StatOper = 1
           !ComiePro = 0
           !TermiPro = 0
           !TermiTemp = 0
           !Observa = ""
           .Update
         Next U&
       End With
       '
     End If
     GoTo 1405
     '
99   Command7.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call BAJALDISCO
End Sub

