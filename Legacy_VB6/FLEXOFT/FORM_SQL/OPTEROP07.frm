VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form OPTEROP07 
   BackColor       =   &H00E0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tercerización de Operaciones"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5430
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   19
      Top             =   720
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   6840
      Left            =   4515
      ScaleHeight     =   6780
      ScaleWidth      =   2160
      TabIndex        =   12
      Top             =   0
      Width           =   2220
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   540
         TabIndex        =   16
         Top             =   2730
         Width           =   600
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   17
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
         Left            =   120
         Picture         =   "OPTEROP07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   315
         Width           =   650
      End
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
         Height          =   640
         Left            =   120
         Picture         =   "OPTEROP07.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   950
         Width           =   650
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
         Height          =   640
         Left            =   120
         Picture         =   "OPTEROP07.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Configuración de Procesos y Formularios"
         Top             =   1785
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "OPTEROP07.frx":0896
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   3060
         Width           =   1515
      End
   End
   Begin VB.ListBox SORTLIST 
      Height          =   300
      Left            =   1890
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0FFC0&
      Caption         =   "Actualización"
      Height          =   1905
      Left            =   210
      TabIndex        =   1
      Top             =   210
      Width           =   4110
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1050
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   4
         Top             =   1050
         Width           =   2325
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
            Left            =   1750
            Picture         =   "OPTEROP07.frx":27B8
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Anulación de Orden de Servicio"
            Top             =   0
            Width           =   530
         End
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
            Picture         =   "OPTEROP07.frx":2AC2
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Consultas / Anulacion"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   600
            TabIndex        =   6
            Top             =   100
            Width           =   1095
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   2
         Top             =   525
         Width           =   2325
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
            Picture         =   "OPTEROP07.frx":2DCC
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Generar una Nueva Orden de Compra"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Generar Orden de Servicios"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   840
            TabIndex        =   3
            Top             =   50
            Width           =   1380
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0FFC0&
      Caption         =   "Consultas - Listados"
      Height          =   1245
      Left            =   210
      TabIndex        =   0
      Top             =   2205
      Width           =   4110
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   9
         Top             =   420
         Width           =   2640
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
            Picture         =   "OPTEROP07.frx":30D6
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Pendientes por Número de O/F"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   945
            TabIndex        =   11
            Top             =   70
            Width           =   1590
         End
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "OPTEROP07.frx":33E0
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "OPTEROP07.frx":3614
      TabIndex        =   20
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
End
Attribute VB_Name = "OPTEROP07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    OPTEROP07.Hide
End Sub

Private Sub Command2_Click()
    '
    Command2.Enabled = False
    Screen.MousePointer = 11
    Call COPYSTRU("OPERTER", "LIOPTER")
    JJ& = 0
    '
    Call APERBASDAT("ORFAB")
    FIN& = ULTREG&("OPERTER")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("OPERTER", U&)
      NOFA$ = TRIM$(Left$(X$, 12))
      FEBAJA% = CVI(Mid$(X$, 157, 2))
      If FEBAJA% < 1 Or FEBAJA% = 8224 Then
        STATUOF% = XVALO(VALOBADA("ORDEFABR", "STATUORF", "IDSUBOR = '" & NOFA$ & "'"))
        If STATUOF% >= 3 Then
          FECIE% = CVINT(VALOBADA("ORDEFABR", "FechCierr", "IDSUBOR = '" & NOFA$ & "'"))
          Call REPLA(X$, MKI$(FECIE%), 157, 2)
          Call GRAREG("OPERTER", X$, U&)
          GoTo 99
        End If
        JJ& = JJ& + 1
        Call GRAREG("LIOPTER", X$, JJ&)
      End If
99  Next U&
    Call SETULTREG("LIOPTER", JJ&)
    Call CIERRARCH("LIOPTER")
    '
    Call CONECRUN("*LIOPTER", "Ordenes de Servicio Pendientes")
    '
    Screen.MousePointer = 1
    Command2.Enabled = True
    '
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call TERCOPER
    Command3.Enabled = True
End Sub

Private Sub Command5_Click()
    '
    Command5.Enabled = False
    Call APERBASDAT("ORFAB")
    '
3   Call SELECHO("INDIOTER/Indice General de Ordenes de Servicio (Pendientes-Cumplidas-Anuladas)")
    NPX& = Val(VALACT1$("INDIOTER"))
    NORFA$ = TRIM$(Mid$(VALACT2$("INDIOTER"), 1, 12))
    If NPX& < 1 Then GoTo 99
    '
    OPX% = COMALTER%("¿Esta Seguro Que Desea Anular La Orden '" & NPX& & "'?\\No, Cancelar\Si, Continuar")
    If OPX% < 2 Then GoTo 99
    '
    CONDI$ = "N_ORSERV=" & NPX& & " AND DOC_ORIG='" & NORFA$ & "'"
    CANTRECI = XVALO(VALOBADA("OCOMDETA", "Cant_Rec", CONDI$))
    NUMEOCOM = XVALO(VALOBADA("OCOMDETA", "Nume_Ocom", CONDI$))
    If CANTRECI > 0.05 Then
        Call COMUNI("Imposible Continuar. O-Compra '" & TRIM$(str$(NUMEOCOM)) & "' Recibida Parcialmente")
        GoTo 99
    End If
    '
    RFF$ = RECFILT$("OPERTER", 9, MKS$(NPX&))
    If RFF$ <> "" Then
        Reg& = CVS(Mid$(RFF$, 1, 4))
        XX$ = REGLEIDO$("OPERTER", Reg&)
        '
        NUMEOPER% = CVI(Mid$(XX$, 15, 2))
        Call REPLA(XX$, MKI$(HOY(HO$)), 157, 2)
        Call GRAREG("OPERTER", XX$, Reg&)
        CONDI$ = "NUME_OCOM=" & NUMEOCOM
        Call ACTUREGISTRO("OCOMDETA", "Stat_Ocom", CONDI$, 9, REGAF&)
        Call ACTUREGISTRO("OCOMENCA", "Stat_Ocom", CONDI$, 9, REGAF&)
        Call ACTUREGISTRO("OPERFAB", "CANTREAL", "IDSUBOR='" & NORFA$ & "'", 0, REGAF&)
        '
        Call COMUNI("Orden de Servicio '" & TRIM$(str$(NPX&)) & "' Anulada")
    End If
    '
99  Command5.Enabled = True
    '
End Sub





Private Sub Command6_Click()
   Command6.Enabled = False
   FOROPTER.Show 1
   Command6.Enabled = True
End Sub

Private Sub Command7_Click()
    '
    Command7.Enabled = False
    'Call MENSERR(24, "No Disponible Por Ahora")
    'GoTo 99
    '
3   Call SELECHO("INDIOTER/Indice General de Ordenes de Servicio (Pendientes-Cumplidas-Anuladas)")
    NPX& = Val(VALACT1$("INDIOTER"))
    If NPX& < 1 Then GoTo 99
    '
    TIDOCUM$ = "Orden de Servicio"
    NUDOCU$ = " " + TRIM$(str$(NPX&)) + " - "
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        PPXX% = InStr(XX$, "Nro.")
        If PPXX% > 0 Then
          XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
        End If
        If InStr(XX$, NUDOCU$) > 4 Then
          DOCUNU& = CVS(Left$(XX$, 4))
          GoTo 5
        End If
      End If
    Next UI&
    Screen.MousePointer = 1
    Call COMUNI("Documento no Encontrado")
    GoTo 99
    '
    '19/03/07 (OT-07-0113)
5   Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
    '19/03/07 (FIN)
'5   For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Cls
'      EDITFORM.Picture1(KKI%).DrawWidth = 3
'      EDITFORM.Picture1(KKI%).DrawStyle = 0
'      EDITFORM.Picture1(KKI%).FillStyle = 1
'    Next KKI%
'    '
'    REDOCU$ = REGACT$("PDOCLST")
'    URE& = ULTREG&("PDOCSPL")
'    LI& = 0: LS& = URE&
'6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
'    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
'    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
'    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
'    GoTo 8
'    '
'7   Screen.MousePointer = 1
'    Call MENSERR(24, "Documento no Registrado")
'    GoTo 99
'    '
'8   TPSP$ = ""
'    Screen.MousePointer = 11
'    While L& > 1
'      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'      If DCC& < DOCUNU& Then GoTo 9
'      L& = L& - 1
'    Wend
'    '
'9   CAPAGINAS% = 0
'    For KKL& = L& To URE&
'      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
'      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
'      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
'        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
'      End If
'    Next KKL&
'    '
'95  SPOOLSTRING$ = TPSP$
'    FIN& = Len(TPSP$)
'    For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Cls
'      EDITFORM.Picture1(KKI%).Top = 0
'      EDITFORM.Picture1(KKI%).Refresh
'      EDITFORM.Picture1(KKI%).Height = 7170
'      EDITFORM.Picture1(KKI%).Width = 11700
'      EDITFORM.Picture1(KKI%).DrawWidth = 3
'      EDITFORM.Picture1(KKI%).DrawStyle = 0
'      EDITFORM.Picture1(KKI%).FillStyle = 1
'    Next KKI%
'    ALTUPAGINA = 7170
'    TOPEPAGINA = 0
'    PAGINACTIVA% = 0
'    CAPAGINAS% = 0
'    '
'10  PPXX% = InStr(TPSP$, Chr$(255))
'    If PPXX% > 0 Then
'      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
'      TPSP$ = Mid$(TPSP$, PPXX% + 1)
'      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
'      '
'      If Len(COPRIAT$) > 0 Then
'        COPRI% = Asc(Left$(COPRIAT$, 1))
'        ATRIP$ = Mid$(COPRIAT$, 2)
'        Call SHOWCOMMAND(COPRI%, ATRIP$)
'      End If
'      GoTo 10
'      '
'    End If
'    '
'    UI& = UI& + 1
'    XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
'    If InStr(XX$, TIDOCUM$) > 4 Then
'      If InStr(XX$, NUDOCU$) > 4 Then
'        DOCUNU& = CVS(Left$(XX$, 4))
'        GoTo 7
'      End If
'    End If
'    '
'    Screen.MousePointer = 1
'    PAGINACTIVA% = 0
'    For KKI% = 0 To 5
'       EDITFORM.Picture1(KKI%).Visible = False
'    Next KKI%
'    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
'    EDITFORM.Picture1(PAGINACTIVA%).Refresh
'    EDITFORM.VScroll1.Min = 0
'    EDITFORM.VScroll1.Value = 0
'    EDITFORM.VScroll1.Max = 1
'    If ALTUPAGINA > EDITFORM.Frame1.Height Then
'       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
'    End If
'    EDITFORM.Command3.Enabled = False
'    EDITFORM.Command5.Enabled = False
'    EDITFORM.Show 1
    '
    GoTo 3
    '
99  Command7.Enabled = True
    '
End Sub

Sub TERCOPER()
   '
   HO% = HOY(HOS$)
   LU$ = LUDAT$
   NROOF$ = ""

    '
    HO% = HOY(HOS$)
    LU$ = LUDAT$
    '
    TITU$ = "INFORMES DE PRODUCCION - GENERALES"

100  NROOF$ = ""
     '
     Call GENRESDAT(1)
     'Call GENORDEFDAT(1)
     '
1405 Call FRESHECHO("ECORDEP")
     If ULTREG&("ECORDEP") < 1 Then
         Call MENSERR(24, "No hay Ordenes de Fabricación Pendientes.")
         Exit Sub
     End If
     '
     Call SELECHO("ECORDEP/Ordenes de Trabajo en Curso")
1410 NROOF$ = AJUSTI$(VALACT1$("ECORDEP"), 12)
     If TRIM$(NROOF$) = "" Then
       Exit Sub
     End If
     '
     Screen.MousePointer = 11
     '
     Call FRESHECHO("!OPERINF") 'nuevo
     'ACA OBTENGO LOS REGISTROS DE OPERFAB, CUYO IdSubOr = NROOF$
     'Luego cambio las variable por los campos correspondientes
     '
     Call ABREORFAB
     SQLX$ = " SELECT CANTPROG, CANTREAL, NUMEOPER, DESCOPER, MAQUINA FROM OPERFAB WITH(NOLOCK) "
     SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "' "
     SQLX$ = SQLX$ + "AND StatOper < 8 "                      ' no marca cumplidos"
     SQLX$ = SQLX$ + " ORDER BY NumeOper "
     '
     'Set FABRTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
'     Dim FABRTMP As ADODB.Recordset
'     Set FABRTMP = New ADODB.Recordset
'     FABRTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
     Set FABRTMP = READSET(SQLX$)
     '
     With FABRTMP
       On Error Resume Next
       .MoveFirst
       If Err Then
         Call MENSERR(24, "No hay Operaciones Programadas\para esta Orden de Fabricación.")
         On Error GoTo 0
         GoTo 1405
       End If
       On Error GoTo 0
       JJ& = 0
       YY$ = REGBLAN("OPERINF")
       Do While (.EOF = False)
          CAPROD# = !CANTPROG      'Cantidad Programada
          CAREA# = !CANTREAL       'Cantidad realizada
          SALDO# = CAPROD# - CAREA#: If SALDO# < 0 Then SALDO# = 0
          '
          If SALDO# > 0.05 Then
            NOPE% = !NUMEOPER      'numero de operación
            DOPE$ = !descoper      'Descripcion de la operacion
            MAQUI$ = !MAQUINA      'Máquina
            '
            JJ& = JJ& + 1: JJI% = JJ&
            YY$ = MKI$(JJI%) + Space$(58) + String(4, 0)
            NOPX$ = AJUSTI$(TRIM$(str$(NOPE%)) + " - " + TRIM$(Left$(DOPE$, 31)) + " - " + MAQUI$, 46)
            NOPX$ = NOPX$ + NUMSTRI$(SALDO#, 12, 1, 0)
            Call REPLA(YY$, NOPX$, 3, 58)
            Call GRAREG("!OPERINF", YY$, JJ&)
            '
          End If
       .MoveNext
       Loop
     End With
     Set FABRTMP = Nothing
        
     Call SETULTREG("!OPERINF", JJ&)
     Call CIERRARCH("!OPERINF")
     '
     Screen.MousePointer = 1
     '
     If JJ& < 1 Then
        Call MENSERR(24, "No hay Operaciones Programadas\para esta Orden de Fabricación.")
        GoTo 1405
     End If
     '
1415 Call SELECHO("!OPERINF/Operaciones Programadas para " + NROOF$)
     NORDE& = Val(VALACT1$("!OPERINF"))
     If NORDE& < 1 Then
       GoTo 1405
     End If
     '
     XX$ = REGLEIDO$("!OPERINF", NORDE&)
     NOPE% = Val(Mid$(XX$, 3, 5))
     SALDO# = Val(Mid$(XX$, 49, 12))
     Screen.MousePointer = 1
     '
     SQLX$ = "SELECT COD_INTE, NUMEOPER, DESCOPER, CANTPROG, CANTREAL FROM OPERFAB WITH(NOLOCK) "
     SQLX$ = SQLX$ + " WHERE IDSUBOR ='" + TRIM$(NROOF$) + "' AND NumeOper = " & NOPE%
'     Set OPERFABTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
     Set OPERFABTMP = READSET(SQLX$)
     With OPERFABTMP
       .MoveFirst
       CI% = !cod_inte
       NOPE% = !NUMEOPER
       DOPE$ = SAFETEXT(!descoper)
       CAPROD# = XVALO(!CANTPROG)
       CAREA# = XVALO(!CANTREAL)
       SALDO# = CAPROD# - CAREA#: If SALDO# < 0 Then SALDO# = 0
       CAINF# = SALDO#
       '
       ZZ$ = REGBLAN("OPERFAB")
       Call REPLA(ZZ$, NROOF$, 1, 12)
       Call REPLA(ZZ$, MKI$(CI%), 13, 2)
       Call REPLA(ZZ$, MKI$(NOPE%), 15, 2)
       Call REPLA(ZZ$, DOPE$, 17, 48)
' DESDE AQUI REVISAR
         NPRO% = PROVHABI%(CI%)
       Call REPLA(ZZ$, MKI$(NPRO%), 73, 2)
       Call REPLA(ZZ$, MKS$(0), 75, 4)
       Call REPLA(ZZ$, String$(178, 0), 79, 178)
       Call REPLA(ZZ$, MKD$(CAINF#), 79, 8)
       Call REPLA(ZZ$, MKI$(HO%), 135, 2)
       Call REPLA(ZZ$, MKI$(HO%), 139, 2)
       Call REPLA(ZZ$, Space$(30), 151, 30)
       Call REPLA(ZZ$, String$(76, 0), 181, 76)
     End With

1420  XX$ = OBJPLA$("TERCEROPER", ZZ$)
     If XX$ = "" Then GoTo 1415
     '
     CAINF# = CVD(Mid$(XX$, 79, 8))
     If CAINF# < 0.999 Then
       Call MENSERR(24, "Falta Cantidad")
       YY$ = XX$
       GoTo 1420
     End If
     Call REPLA(XX$, MKD$(CAINF#), 95, 8)
     '
     NPRO% = CVI(Mid$(XX$, 73, 2))
     If NPRO% < 1 Then
        Call MENSERR(24, "Falta Proveedor")
        YY$ = XX$
        GoTo 1420
      ElseIf ECOARCH$("ACREDOR", MKI$(NPRO%)) = "" Then
        Call MENSERR(24, "Número de Proveedor no Válido")
        YY$ = XX$
        GoTo 1420
     End If
     FENTRESOL% = CVI(Mid$(XX$, 139, 2))
     ATEN$ = Mid$(XX$, 221, 30)
     Observa$ = Mid$(XX$, 257, 320)
     '
     Screen.MousePointer = 11
     Call BLOQARCH("OPERFAB")
     Call BLOQARCH("OPERTER")
     Call BLOQARCH("OCOMENCA")
     Call BLOQARCH("OCOMDETA")
     '
    ' determinar numero de orden de servicio
     NORSE& = PRONUOTER&
     Call REPLA(XX$, MKS$(NORSE&), 75, 4)
     NUOCONUE& = PRONUOCO&
     Call REPLA(XX$, MKS$(NUOCONUE&), 251, 4)
     Call REPLA(XX$, String$(18, 0), 203, 18)
     '
     IXT$ = MKS$(NORSE&) + TRIM$(NROOF$) + " - " + TRIM$(DOPE$) + " " + TRIM$(CODEXT$(CI%)) + " - " + rasocli$((-1) * NPRO%)
     '
     FEBA% = HO%
     HOBA% = Int(HORANUM(Time$))
     USUA% = USNBR% Mod 100
     DBAJA$ = MKI$(FEBA%) + MKI$(HOBA%) + Chr$(USUA%) + Chr$(1) + "Tercerizacion"
     '
     NORFB$ = Left$(XX$, 12)
     CIIK% = CVI(Mid$(XX$, 13, 2))
     NOPE% = CVI(Mid$(XX$, 15, 2))
     CAREAL# = CAINF#
     '
'     ZZ$ = REGLEIDO$("OPERFAB", REGI&)
'     NORFA$ = Left$(ZZ$, 12)
'     If ZZ$ <> RANTEINF$ Or NORFA$ <> NORFB$ Or CVI(Mid$(ZZ$, 15, 2)) <> NOPE% Or Mid$(ZZ$, 177, 6) <> String$(6, 0) Then
'        Call CIERRARCH("OPERFAB")
'        Call MENSERR(24, "Imposible Generar esta Orden de Servicio.\  \Causa Probable: Se Ha Registrado\un Parte para esta Operación\desde otra Terminal de la Red.\  \Vuelva a Generar la Orden de Servicio.")
'        GoTo 1999
'     End If
'     '
     NORFA$ = Left$(ZZ$, 12)

     OBSERX$ = "Orden de Servicio " + TRIM$(str$(NORSE&)) + " - Orden de Fabricación " + NORFA$ + "  "
     Call REPLA(Observa$, OBSERX$, 1, 64)
     '
' Marca como "Cumplida" Operación en OPERFAB.DAT
     SALRE# = CVD(Mid$(ZZ$, 79, 8)) - CAREAL#
     If SALRE# < 0 Then SALRE# = 0
     '
 ' EPB AQUI REVISAR SALRE Y CAREAL Y HACER UN UPDATE EN OPERFAB WHERE COINCIDE NUMERO DE OF Y NOPE%
     '\\\OT-08-0667-OD-04/09/08///
     SQLX$ = "UPDATE OPERFAB SET CantReal = " & FORMATNUM$(CAREAL#, "F13.2") & " "
 'SE CONDICIONO LA CONSULTA POR QUE SI SALRE# ES = 0 FORMATNUM DEVUELVE UN STRING _
 VACIO Y AL CONCATERNARLO EN LA CONSULTA GENERA UN ERROR ()
     If SALRE# > 0.005 Then
       SQLX$ = SQLX$ + " ,CantPend = " & FORMATNUM$(SALRE#, "F13.2") & ""
     Else
       SQLX$ = SQLX$ + " ,CantPend = 0"
     End If
     SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFB$ & "'"
     SQLX$ = SQLX$ + " AND NumeOper = " & NOPE% & " "
     'MsgBox SQLX$
      FLEXCONN.Execute (FILTSQL$(SQLX$))
     '\\\OT-08-0667-OD-FIN///
''     NUOROP% = Asc(Mid$(ZZ$, 67, 1))
''     Call REPLA(ZZ$, MKD$(CAREAL#), 87, 8)
''     Call REPLA(ZZ$, MKD$(SALRE#), 95, 8)
''     Call REPLA(ZZ$, "3ROS  ", 151, 6)
''     Call REPLA(ZZ$, MKI$(HO%), 157, 2)
''     Call REPLA(ZZ$, MKI$(HORANUM%(Time$)), 159, 2)
''     Call REPLA(ZZ$, MKS$(0), 161, 4)
''     Call REPLA(ZZ$, String$(8, 0), 165, 8)
''     Call REPLA(ZZ$, MKS$(0), 173, 4)
''     Call REPLA(ZZ$, DBAJA$, 177, 22)
''     Call GRAREG("OPERFAB", ZZ$, REGI&)
     '
' Si se tercerizó un parcial, genera registro en OPERFAB.DAT
' por el saldo pendiente
     If SALRE# > 0.01 Then
       ' GENERA UN NUEVO MOVIMIENTO
''        NUOROP% = 1 + Asc(Mid$(ZZ$, 67, 1))
''        CIIK% = CVI(Mid$(ZZ$, 13, 2))
''        NOPE% = CVI(Mid$(ZZ$, 15, 2))
''        DALTA$ = Chr$(NUOROP%) + Mid$(DBAJA$, 2, 5)
''        INFOBLA$ = Space$(6) + String$(26, 0) + Space$(16)
''        PROGBLA$ = Space$(6) + String$(26, 0)
''        MINVARI = TIESTAN(CIIK%, NOPE%, SALRE#, NUOROP%)
''        '
''        Call REPLA(ZZ$, DALTA$, 67, 6)
''        Call REPLA(ZZ$, MKD$(SALRE#), 79, 8)
''        Call REPLA(ZZ$, MKD$(0), 87, 8)
''        Call REPLA(ZZ$, MKD$(SALRE#), 95, 8)
''        Call REPLA(ZZ$, String$(32, 0), 103, 32)
''        Call REPLA(ZZ$, MKS$(MINVARI), 111, 4)
''        Call REPLA(ZZ$, MKS$(MINVARI), 131, 4)
''        Call REPLA(ZZ$, String$(16, 0), 135, 16)
''        Call REPLA(ZZ$, INFOBLA$, 151, 48)
''        Call REPLA(ZZ$, PROGBLA$, 225, 32)
''        Call REGAPP("OPERFAB", ZZ$)
''        AGREGA% = 1
     End If
     '
' Carga nuevo registro en Archivo OPERTER.DAT
     Call REGAPP("OPERTER", XX$)
     Call REGAPP("INDIOTER", IXT$)
     Call CIERRARCH("OPERFAB")
     Call CIERRARCH("OPERTER")
     Call CIERRARCH("INDIOTER")
     '
' Genera la correspondiente Orden de Compra
     'ADDNEW EN OCOMDETA SQL
     '\\\OT-08-0667-OD-04/09/08///
     'AGREGADO PARA OBTENER EL PRECIO DE LA OPERACION A REALIZAR (DOSIVAC)
     CONDI$ = "STATOPER > 0 AND CODICONJ = " & CI% & " AND NUMEOPER = " & NOPE%
     PRECIO_UNIT# = XVALO(VALOBADA("SECOPER", "PreUnid", CONDI$, "NOMESS"))
     MONEX% = XVALO(VALOBADA("SECOPER", "MonePrec", CONDI$, "NOMESS"))
     '
     If MONEX% <= 1 Then MONE$ = "$ "
     If MONEX% = 2 Then MONE$ = "U$ "
     If MONEX% > 2 Then MONE$ = Left$(ECOARCH$("TAMONED", Chr$(MONEX%)), 3)
     TICAMBIO = 1
     FORMATOX$ = "F10.4" ' para aprovechar al maximo el espacio para $
     If MONEX% >= 2 Then TICAMBIO = TICAMONE(MONEX%): FORMATOX$ = "F9.4"
     VAUNI$ = MONE$ + FORMATNUM$(PRECIO_UNIT#, FORMATOX$)
            '
     SQLX$ = "SELECT * FROM OCOMDETA WHERE Nume_Ocom < 1 "
     Dim OCOMTMP As ADODB.Recordset
     Set OCOMTMP = New ADODB.Recordset
     OCOMTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     With OCOMTMP
        .AddNew   ' YA CONVERTIDO
        On Error Resume Next
        !CODIEMPR = CodiEmp%
        On Error GoTo 0
        !nume_ocom = NUOCONUE&
        !Femi_Ocom = CVDAT(HO%)
        !Npro_Ocom = NPRO%
        !cod_inte = CI%
        !Uni_Mas = Unimed$(CI%)
        !cant_ocom = CAREAL#
        !Uni_Com = Unimed$(CI%)
        !Coef_Unids = 1
        !Fentre_Sol = CVDAT(FENTRESOL%)
        !CONSIMAT = 1
        !Cant_Rec = 0
        !Cod_Oper = NOPE%
        !Doc_Orig = NROOF$
        !N_ORSERV = NORSE&
        '
        'ESTOS CAMPOS SE DEBERIAN AGREGAR PARA QUE NO QUEDEN NULOS
        !NuOrd_Item = 1
        !Raso_Prov = Left$(rasocli$((-1) * NPRO%), 48)
        !Cod_Exte = CODEXT$(CI%)
        !Descritem = DESCRIT0$(CI%)
        !Delibre = ""
        !Sect_Ocom = ""
        !Medi_Ocom = ""
        !Valo_Unid = VAUNI$
        !Unids_Ocom = CAREAL#
        !Desc_Item = 0
        !Tsum_Ocom = 0
        !Stat_Ocom = 0
        !Mone_Emis = MONEX%
        !mone_cos = MONEX%
        !IContab = 0
        !Pre_Unit = PRECIO_UNIT#
        !prun_neto = PRECIO_UNIT#
        !inet_item = CAREAL# * PRECIO_UNIT#
        .Update
     End With
     OCOMTMP.Close
     Set OCOMTMP = Nothing
     '\\\OT-08-0667-OD-FIN///

     '
     FECHEMI% = HO%
     FEX = FECHEMI%
     REFE$ = FECHATEX$(FEX) + " - " + TRIM$(NORFA$) + " - " + rasocli$((-1) * NPRO%)
     
    'ADDNEW EN OCOMENCA SQL
    '\\\OT-08-0667-OD-04/09/08///

     CPAG% = CPAGCLI%((-1) * NPRO%)

     SQLX$ = "SELECT * FROM OCOMENCA WHERE Nume_Ocom < 1 "
     Dim OCOMENCTMP As ADODB.Recordset
     Set OCOMENCTMP = New ADODB.Recordset
     OCOMENCTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     With OCOMENCTMP
        .AddNew   ' YA CONVERTIDO
        On Error Resume Next
        !CODIEMPR = CodiEmp%
        On Error GoTo 0
        !nume_ocom = NUOCONUE&
        !Refe_Ocom = Left$(REFE$, 64)
        !Femi_Ocom = CVDAT(FECHEMI%)
        !Npro_Ocom = NPRO%
        !Tsum_Ocom = 2 'FASSON
        !Stat_Ocom = 0 'ABIERTA
        !Cpag_Ocom = CPAG%
        !AtSr_Ocom = Left$(ATEN$, 64)
        !Obse_Ocom = Observa$
        '
        'ESTOS CAMPOS SE DEBERIAN AGREGAR PARA QUE NO QUEDEN NULOS
        !Raso_Prov = Left$(rasocli$((-1) * NPRO%), 48)
        !Fgen_Ocom = Now
        !Usua_Gen = USNBR%
        !tcam_ocom = TICAMBIO ' **
        !IDes_Ocom = 0
        !Embala_Ocom = 0
        !Obse_Ocom = ""
        !Pdes_Ocom = ""
        !Lentrega = ""
        !mone_cos = MONEX% ' **
        !Mone_Emis = MONEX% ' **
        !Anex_Ocom = ""
           TOTANE# = XVALO(CAREAL# * PRECIO_UNIT#)
           TOTIVA# = (Int(ALIVA * TOTANE# + 0.5)) / 100
           TOTATO# = TOTANE# + TOTIVA#
        !TOBRU_OCOM = TOTANE#
        !Iiva_Ocom = TOTIVA#
        !Iotr1_Ocom = 0
        !Total_Ocom = TOTATO#
        !PORCIVA = FORMATNUM$(ALIVA, "F5.2")
        .Update
     End With
     OCOMENCTMP.Close
     Set OCOMENCTMP = Nothing
     '\\\OT-08-0667-OD-04/09/08///
     '
     '
     JJ& = 0
     For U& = 1 To 319 Step 64
       TTXX$ = Mid$(Observa$, U&, 64)
       JJ& = JJ& + 1
       Call GRAREG("!OBSERVOF", TTXX$, JJ&)
     Next U&
     Call SETULTREG("!OBSERVOF", JJ&)
     Call CIERRARCH("!OBSERVOF")
     '
     'SI TIENE CONFIGURADO FRM , IMPRIME REMITO EN VEZ DE ORDEN DE SERVICIO
     If TRIM$(Control("", "REMISERV")) <> "" Then '
        GoTo 1600
     End If
     '
     ' imprimir Orden de Servicio
     FORIPRE$ = TRIM$(Control$("", "FORSERVI"))
     If FORIPRE$ <> "" Then
        LU$ = LUDAT$
        If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
          RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
          Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Orden de Servicio:\  \'" + RMCC$ + "'.")
          GoTo 1600
        End If
        CANCELPRINT% = 0
        If Control$("ORSERVI", "SUPRINT") = "SI" Then
          CANCELPRINT% = 1
        End If
        '
        FEX = FECHEMI%
        FECHDOC$ = FECHATEX$(FEX)
        NUMEDOC$ = TRIM$(str$(NORSE&))
        TIPODOC$ = "Orden de Servicio"
        '
        DTP$ = OBJPLA$("TRANSPOR", DATRANSPO$)
        DATRANSPO$ = DTP$
        If DATRANSPO$ = "" Then
          CANCELPRINT% = 1
          Call COMUNI("Imposible Cancelar.\Solo se Suprime la Impresion")
        End If
        '
        CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = TRIM$(Mid$(DTP$, 1, 32))
        CODPARAM$(2) = "DOMI-TRA": DOCPARAM$(2) = TRIM$(Mid$(DTP$, 33, 32))
        CODPARAM$(3) = "CUIT-TRA": DOCPARAM$(3) = TRIM$(Mid$(DTP$, 65, 16))
        CODPARAM$(4) = "PATE-TRA": DOCPARAM$(4) = TRIM$(Mid$(DTP$, 81, 16))
        CODPARAM$(5) = "COND-TRA": DOCPARAM$(5) = TRIM$(Mid$(DTP$, 97, 32))
        CODPARAM$(6) = "DOCUCOND": DOCPARAM$(6) = TRIM$(Mid$(DTP$, 129, 24))
        CODPARAM$(13) = "LOCA-TRA": DOCPARAM$(13) = TRIM$(Mid$(DTP$, 159, 32))
        CODPARAM$(14) = "NUBULTOS": DOCPARAM$(14) = TRIM$(Mid$(DTP$, 153, 6))
        FEX = FENTRESOL%
        CODPARAM$(7) = "FECH-VEN": DOCPARAM$(7) = FECHATEX$(FEX)
        CODPARAM$(8) = "DOCUORIG": DOCPARAM$(8) = NORFA$
        CODPARAM$(9) = "NOP-OPER": DOCPARAM$(9) = str$(NOPE%)
        CODPARAM$(10) = "DES-OPER": DOCPARAM$(10) = DOPE$
        CODPARAM$(11) = "CAN-COMP": DOCPARAM$(11) = str$(CAINF#)
        CODPARAM$(12) = "VENDEDOR": DOCPARAM$(12) = TRIM$(Mid$(ENCAOCOM$, 163, 30))          ' ATENCION SR.
        CAPARDOC% = 14
        '
        CODTABU1$(1) = "COD-MAT"
        CODTABU1$(2) = "DES-MAT"
        CODTABU1$(3) = "CANTIDAD"
        CODTABU1$(4) = "CAKILOS"
        CODTABU1$(5) = "UNIMED"
        CACOLTAB1% = 5
        '
        DESTIDOC% = (-1) * NPRO%
        CAITAB1% = 1
        '
        UNID$ = Unimed$(CI%)
        CANS$ = CSTRING$(MKD$(CAINF#), 3, 9, 1, 2)
        KILO$ = CSTRING$(MKD$(CAINF# * PESUNI(CI%)), 3, 12, 3, 2)
        NULOTE$ = TRIM$(Mid$(YY$, 69, 12))
        NUOCO$ = CSTRING$(Mid$(YY$, 151, 4), 3, 7, 0, 2)
        '
        CAITAB1% = 1
        TETABU1$(1, CAITAB1%) = CODEXT$(CI%)
        TETABU1$(2, CAITAB1%) = DESCRIT0$(CI%)
        TETABU1$(3, CAITAB1%) = CANS$
        TETABU1$(4, CAITAB1%) = KILO$
        TETABU1$(5, CAITAB1%) = UNID$
        '
        CAITAB1% = 2
        TETABU1$(1, CAITAB1%) = ""
        TETABU1$(2, CAITAB1%) = "ANTES DE " + TRIM$(DOPE$)
        TETABU1$(3, CAITAB1%) = ""
        TETABU1$(4, CAITAB1%) = ""
        TETABU1$(5, CAITAB1%) = ""
        '
        Call PRINTDOC("FORSERVI")   ' Orden de Servicio
        '
     End If
     '
1600 ' imprimir Orden de Compra
     '
     FORIPRE$ = TRIM$(Control$("", "FORMOCOM"))
     If FORIPRE$ <> "" Then
         LU$ = LUDAT$
         If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
           RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
           Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Orden de Compra:\  \'" + RMCC$ + "'.")
           GoTo 1999
         End If
         CANCELPRINT% = 0
         If Control$("OCOMPRA", "SUPRINT") = "SI" Then
           CANCELPRINT% = 1
         End If
         '
         Screen.MousePointer = 11
         FECHDOC$ = FETEXTO$(FECHEMI%)
         NUMEDOC$ = TRIM$(str$(NUOCONUE&))
         While Len(NUMEDOC$) < 6
           NUMEDOC$ = "0" + NUMEDOC$
         Wend
         TIPODOC$ = "Nota de Pedido"
         CANTU = CAREAL#
         PreUnid# = PRECIO_UNIT#
         IMPOXTE# = XVALO(CANTU * PreUnid#)
         TOTANEDES# = IMPOXTE#
         TOTANE# = IMPOXTE#
         TOTIVA# = (Int(ALIVA * TOTANE# + 0.5)) / 100
         TOTATO# = TOTANEDES# + TOTIVA#
         '
         CODPARAM$(1) = "COND-PAG"
         CODPARAM$(2) = "OBSERVA"
         CODPARAM$(3) = "IMP-NETO"
         CODPARAM$(4) = "POR-IVA"
         CODPARAM$(5) = "IMP-IVA"
         CODPARAM$(6) = "IMP-TOTA"
         CODPARAM$(7) = "VENDEDOR"
         CODPARAM$(8) = "LIST-PRE"
         CODPARAM$(9) = "ZONA-VEN"
         CODPARAM$(10) = "FECH-LAN"
         CODPARAM$(11) = "IMPNEDES"
         CODPARAM$(12) = "TRANSPOR"
         CODPARAM$(13) = "POR-DESC"
         CODPARAM$(14) = "IMP-DESC"
         CODPARAM$(15) = "TIPO-VAL"
         CAPARDOC% = 15
         '
         DOCPARAM$(1) = FORPAG$
         DOCPARAM$(2) = OBSE$
         DOCPARAM$(3) = str$(TOTANE#)
         DOCPARAM$(4) = str$(ALIVA)
         DOCPARAM$(5) = str$(TOTIVA#)
         DOCPARAM$(6) = str$(TOTATO#)
         DOCPARAM$(7) = Mid$(X1$, 147, 20)
         DOCPARAM$(8) = Mid$(X1$, 67, 40)
         DOCPARAM$(9) = Mid$(X1$, 167, 40)
         DOCPARAM$(10) = Mid$(X1$, 107, 40)
         DOCPARAM$(11) = str$(TOTANEDES#)
         DOCPARAM$(12) = TRIM$(Mid$(ENCAOCOM$, 163, 30)) ' ATENCION SR.
         DOCPARAM$(13) = TRIM$(PORDESCUENTO$)
         DOCPARAM$(14) = str$(IDESCUEN#)
         DOCPARAM$(15) = DEMONECO$(MONEX%)
         '
         CODTABU1$(1) = "COD-MAT"
         CODTABU1$(2) = "DES-MAT"
         CODTABU1$(3) = "UNIMED"
         CODTABU1$(4) = "CANTIDAD"
         CODTABU1$(5) = "FECH-VEN"
         CODTABU1$(6) = "REF-MAT1"
         CODTABU1$(7) = "PRE-LIST"
         CODTABU1$(8) = "PORDESCU"
         CODTABU1$(9) = "PRE-NETO"
         CODTABU1$(10) = "NET-ITEM"
         CODTABU1$(11) = "CAKILOS"
         CODTABU1$(12) = "F-PLANO"
         CODTABU1$(13) = "ORD-ITEM"
         CACOLTAB1% = 13
         '
         DESTIDOC% = (-1) * NPRO%
         CAITAB1% = 0
         CAITCRU% = 0
         '
         JJ& = 0
         For U& = 1 To 319 Step 64
           TTXX$ = Mid$(Observa$, U&, 64)
           JJ& = JJ& + 1
           Call GRAREG("!OBSERVOF", TTXX$, JJ&)
         Next U&
         Call SETULTREG("!OBSERVOF", JJ&)
         Call CIERRARCH("!OBSERVOF")
         '
         X$ = DETOCOM$
         CI0% = CI%
         PORDE = 0
         FEX = FENTRESOL%
         CONSIMAT% = 0: If TRIM$(Mid$(X$, 97, 2)) <> "" Then CONSIMAT% = 1
         '
         CAITAB1% = CAITAB1% + 1
         TETABU1$(1, CAITAB1%) = CODEXT$(CI0%)
         TETABU1$(2, CAITAB1%) = DESCRIT0$(CI0%)
         TETABU1$(3, CAITAB1%) = Mid$(X$, 59, 4)
         TETABU1$(4, CAITAB1%) = TRIM$(str$(CANTU))
         TETABU1$(5, CAITAB1%) = FECHATEX$(FEX)
         TETABU1$(6, CAITAB1%) = str$(CONSIMAT%)  'consigna m.prima
         TETABU1$(7, CAITAB1%) = str$(PreUnid#)
         TETABU1$(8, CAITAB1%) = str$(PORDE)
         TETABU1$(9, CAITAB1%) = str$(PreUnid# * (1 - PORDE / 100))
         TETABU1$(10, CAITAB1%) = str$(IMPOXTE#)
         TETABU1$(11, CAITAB1%) = str$(CANTU * PreUnid#)
         TETABU1$(12, CAITAB1%) = FEPLANO$(CI0%)
         TETABU1$(13, CAITAB1%) = TRIM$(str$(CAITAB%))
         '
         CAITAB1% = CAITAB1% + 1
         For KKII% = 1 To 13: TETABU1$(KKII%, CAITAB1%) = "": Next KKII%
         TETABU1$(2, CAITAB1%) = "Operacion: " + TRIM$(DOPE$)
         '
         Call PRINTDOC("FORMOCOM")   ' Orden de Compra
         '
     End If
     '
     FEX = FECHEMI%
     FECHDOC$ = FECHATEX$(FEX)
     'OBTENER EL PROXIMO NUMERO DE REMITO
     CMO$ = TRIM$(Control$("", "CODREFAC"))
     If CMO$ = "" Then CMO$ = "RT"
     PVENTA& = XVALO(Control$("", "PVTARTSE"))
     If PVENTA& < 1 Then
        PVENTA& = 1
     End If
     '
     For U& = 1 To 319 Step 64
       TTXX$ = Mid$(Observa$, U&, 64)
       XX1& = XX1& + 1
       Call GRAREG("!OBSERVOF", TTXX$, XX1&)
     Next U&
     Call SETULTREG("!OBSERVOF", XX1&)
     Call CIERRARCH("!OBSERVOF")

     NROPRO& = 0
     Call NUREMIFA(NROPRO&, SAFETEXT$(PVENTA&))
     If NROPRO& < 2 Then
5        NROPRO& = XVALO(InputBox$("Inicialización N° Remito P.Vta. " & PVENTA&, "Ingrese Número de Próximo Remito", NROPRO&))
         If NROPRO& < 1 Or NROPRO& > 9999 Then GoTo 5
     End If
     NROREMIT$ = SAFETEXT$(NROPRO&)
     PREFIREM$ = Right$("0000" & SAFETEXT$(PVENTA&), 4)
     While Len(NROREMIT$) < 8: NROREMIT$ = "0" + NROREMIT$: Wend
     NROREMIT$ = CMO$ + "-X-" + PREFIREM$ + "-" + NROREMIT$
     NURE2& = XVALO(Mid$(NROREMIT$, 11))
     SUC1% = XVALO(PREFIREM$)
     SUC2$ = TRIM$(str$(SUC1%))
     REMICOR1$ = CMO$ & "." + str$(NURE2&)
     DOSEC$ = NORFA$
  '
     NUORIT% = 0
     VUNI# = 0
     MONEII% = 0
     NC% = NPRO%
     REFE$ = "Remito de Servicio - " & NORFA$ & " - Op.: " & TRIM$(DOPE$)
     DELAR$ = TRIM$(str$(CANTU)) & " - " & CODEXT$(CI%) & " - " & DESCRIT0$(CI%)
     '
     'MOVIMIENTOS DE SALIDA
     Call MOVISTOK(FECHEMI%, 0, "", CMO$, NROREMIT$, REMICOR1$, DOSEC$, DOSEC$, REFE$, 1, VUNI#, MONEII%, NC%, 1, CANTU, 0, 0, DELAR$, "", "", 0, 0)
     '
     Call LIMPIATETAB
     FORIPRE$ = TRIM$(Control$("", "REMISERV"))
     If FORIPRE$ <> "" Then
       LU$ = LUDAT$
       If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
         RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
         Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Orden de Remito de Servicio:\  \'" + RMCC$ + "'.")
         GoTo 1999
       End If
     End If
     '
     CANCELPRINT% = 0
     If Control$("REMITOSE", "SUPRINT") = "SI" Then
        CANCELPRINT% = 1
     End If
     '
     NUMEDOC$ = Mid$(NROREMIT$, 6)
     TIPODOC$ = "Remito"
     DESTIDOC% = (-1) * NPRO%
     CAITAB1% = 0
      '
     DTP$ = OBJPLA$("TRANSPOR", DATRANSPO$)
     DATRANSPO$ = DTP$
     If DATRANSPO$ = "" Then
        CANCELPRINT% = 1
        Call COMUNI("Imposible Cancelar.\Solo se Suprime la Impresion")
     End If
      
     Call SETULTREG("!OBSERVOF", 0)
     CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = TRIM$(Mid$(DTP$, 1, 32))
     CODPARAM$(2) = "NUBULTOS": DOCPARAM$(2) = TRIM$(Mid$(DTP$, 153, 6))
     CODPARAM$(3) = "CUIT-TRA": DOCPARAM$(3) = TRIM$(Mid$(DTP$, 65, 16))
     CODPARAM$(4) = "REF-ASIE": DOCPARAM$(4) = TRIM$(Mid$(DTP$, 33, 32))
     CODPARAM$(5) = "DOCUCOND": DOCPARAM$(5) = TRIM$(Mid$(DTP$, 129, 24))
     CODPARAM$(6) = "DEN-CUEN": DOCPARAM$(6) = TRIM$(Mid$(DTP$, 159, 32))
'     CODPARAM$(7) = "OBSERVA": DOCPARAM$(7) = Observa$
     CODPARAM$(7) = "REF-MAT6": DOCPARAM$(7) = OBSERX$
     CAPARDOC% = 7
     '
     CAITAB1% = CAITAB1% + 1
     CODTABU1$(1) = "COD-MAT": TETABU1$(1, CAITAB1%) = CODEXT$(CI%)
     CODTABU1$(2) = "DES-MAT": TETABU1$(2, CAITAB1%) = DESCRIT0$(CI%)
     CODTABU1$(3) = "CANTIDAD": TETABU1$(3, CAITAB1%) = TRIM$(str$(CANTU))
     CACOLTAB1% = 3
     '
     CAITAB1% = CAITAB1% + 1
     TETABU1$(2, CAITAB1%) = "Operacion: " + TRIM$(DOPE$)
     '
     Call PRINTDOC("@" + FORIPRE$)  ' Remito de Facturacion
     '
1999 Call CIERRARCH("*.*")
     Call BAJALDISCO
     
End Sub


Function PRONUOTER&()    ' numero orden de servicio terceros
   '
   PRONU& = 1
   REGXX& = ULTREG&("OPERTER")
   If REGXX& > 0 Then
     PRONU& = 1 + CVS(Mid$(REGLEIDO$("OPERTER", REGXX&), 75, 4))
   End If
   If PRONU& <= REGXX& Then
     Screen.MousePointer = 11
     For U& = 1 To REGXX&
       X0$ = REGLEIDO$("OPERTER", U&)
       If CVS(Mid$(X0$, 75, 4)) >= PRONU& Then
         PRONU& = 1 + CVS(Mid$(X0$, 75, 4))
       End If
     Next U&
     Screen.MousePointer = 1
   End If
   '
   PRONUOTER& = PRONU&
   '
End Function

Sub GENRESDAT(Optional MODO%)
    '
    Screen.MousePointer = 11
    ' GENERA SOLO ECORDEP
    If MODO% = 1 Then GoTo 40
    '
    Call ABRESTOCKS
    '
    SQLX$ = " SELECT COD_INTE, IDSUBOR, COD_INTE, FechRes, CantRes, CANTCONS, Saldo_Entre, IdenLote,Deposito, UsoUnit "
    SQLX$ = SQLX$ + " FROM RESERVA WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE CantRes-CantCons > 0 "
    SQLX$ = SQLX$ + " ORDER BY IdSubOr "
    '
    'Set RESERTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
'    Dim RESERTMP As ADODB.Recordset
'    Set RESERTMP = New ADODB.Recordset
'    RESERTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

    Set RESERTMP = READSET(SQLX$)
    With RESERTMP
      On Error Resume Next
      .MoveFirst
      'If Err Then Call MENSERR(24, " No Existen Reservas efectuadas. ")
      On Error GoTo 0
      JJ& = 0
      Do While (.EOF = False)
        Y$ = REGBLAN("RESERVA")
        '
        Call REPLA(Y$, MKI$(!cod_inte), 1, 2)
        Call REPLA(Y$, !idsubor, 3, 12)
        Call REPLA(Y$, MKI$(!cod_inte), 15, 2)
        Call REPLA(Y$, MKI$(CVINT(!FechRes)), 17, 2)
        Call REPLA(Y$, MKD$(!CANTRES), 19, 8)
        Call REPLA(Y$, MKD$(!CANTCONS), 27, 8)
        Call REPLA(Y$, MKD$(!Saldo_Entre), 35, 8)
        Call REPLA(Y$, !idenlote, 43, 16)
        Call REPLA(Y$, Chr$(!DEPOSITO), 60, 1)
        Call REPLA(Y$, MKS$(!UsoUnit), 75, 4)
        '
        JJ& = JJ& + 1
        Call GRAREG("RESERVA", Y$, JJ&)
        '
      .MoveNext
      Loop
    End With
    RESERTMP.Close
    Set RESERTMP = Nothing
    Call SETULTREG("RESERVA", JJ&)
    Call CIERRARCH("RESERVA")
    '
40  Call ABREORFAB
    '
    SQLX$ = " SELECT DISTINCT IdSubOr FROM ORDEFABR WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE StatuOrf < 3 "
    SQLX$ = SQLX$ + " ORDER BY IdSubOr "
    '
'    Dim ORDEFABTMP As ADODB.Recordset
'    Set ORDEFABTMP = New ADODB.Recordset
'    ORDEFABTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    Set ORDEFABTMP = READSET(SQLX$)
    With ORDEFABTMP
      On Error Resume Next
      .MoveFirst
      If Err Then Call MENSERR(24, "No Existen Ordenes de Fabricación en Curso.")   '  Imposible Continuar la Carga.")
      On Error GoTo 0
      JJ& = 0
      Do While (.EOF = False)
        Y$ = REGBLAN("ECORDEP")
        '
        Call REPLA(Y$, !idsubor, 1, 12)
        Call REPLA(Y$, ECOARCH$("ECORDEF", !idsubor), 13, 52)
        '
        JJ& = JJ& + 1
        Call GRAREG("ECORDEP", Y$, JJ&)
        '
      .MoveNext
      Loop
    End With
    ORDEFABTMP.Close
    Set ORDEFABTMP = Nothing
    Call SETULTREG("ECORDEP", JJ&)
    Call CIERRARCH("ECORDEP")
    '
    Screen.MousePointer = 1
    '
End Sub


Private Sub Form_Load()
    Call APLICACIONSKINS(Me, Picture14)
End Sub
