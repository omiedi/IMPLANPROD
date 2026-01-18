VERSION 5.00
Begin VB.Form OPTEROPER 
   BackColor       =   &H00E0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tercerización de Operaciones"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5610
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
   ScaleHeight     =   3585
   ScaleWidth      =   5610
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4540
      ScaleHeight     =   75
      ScaleWidth      =   810
      TabIndex        =   11
      Top             =   2310
      Width           =   870
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.ListBox SORTLIST 
      Height          =   300
      Left            =   1890
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
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
      Left            =   4540
      Picture         =   "OPTEROPER.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1260
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0FFC0&
      Caption         =   "Actualización"
      Height          =   1905
      Left            =   210
      TabIndex        =   2
      Top             =   210
      Width           =   4110
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1050
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   6
         Top             =   1050
         Width           =   2325
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
            Picture         =   "OPTEROPER.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Consultas y Re-Impresión"
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
            Left            =   840
            TabIndex        =   8
            Top             =   105
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   3
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
            Picture         =   "OPTEROPER.frx":074C
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Generar una Nueva Orden de Compra"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Generar Orden de Servicios"
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
            Left            =   840
            TabIndex        =   4
            Top             =   0
            Width           =   1380
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0FFC0&
      Caption         =   "Consultas - Listados"
      Height          =   1245
      Left            =   210
      TabIndex        =   1
      Top             =   2205
      Width           =   4110
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   13
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
            Picture         =   "OPTEROPER.frx":0A56
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Pendientes por Número de O/F"
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
            Left            =   945
            TabIndex        =   15
            Top             =   105
            Width           =   1590
         End
      End
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
      Height          =   855
      Left            =   4540
      Picture         =   "OPTEROPER.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   315
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   3990
      Picture         =   "OPTEROPER.frx":0EAA
      Top             =   2940
      Width           =   2010
   End
End
Attribute VB_Name = "OPTEROPER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    OPTEROPER.Hide
End Sub

Private Sub Command2_Click()
    '
    Command2.Enabled = False
    Screen.MousePointer = 11
    Call COPYSTRU("OPERTER", "LIOPTER")
    JJ& = 0
    '
    FIN& = ULTREG&("OPERTER")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("OPERTER", U&)
      FEBAJA% = CVI(Mid$(X$, 157, 2))
      If FEBAJA% < 1 Or FEBAJA% = 8224 Then
        JJ& = JJ& + 1
        Call GRAREG("LIOPTER", X$, JJ&)
      End If
    Next U&
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
    NUDOCU$ = " " + TRIM$(Str$(NPX&)) + " - "
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
     BVOF% = BINVAL%(AJUSTI$(NROOF$, 12))
     RGFSX$ = RECFILT$("OPERFAB", 6, MKI$(BVOF%))
     RGFSY$ = ""
     For KJ& = 1 To Len(RGFSX$) Step 4
       REOPE& = CVS(Mid$(RGFSX$, KJ&, 4))
       If REOPE& > 0 Then
         If REOPE& <= ULTREG&("OPERFAB") Then
           XX$ = REGLEIDO$("OPERFAB", REOPE&)
           If Left$(XX$, 12) = NROOF$ Then
             If Mid$(XX$, 177, 6) = String$(6, 0) Then ' aun pendiente
               RGFSY$ = RGFSY$ + MKS$(REOPE&)
             End If
           End If
         End If
       End If
     Next KJ&
     
     If Len(RGFSY$) < 4 Then
         Call MENSERR(24, "No hay Operaciones Programadas\para esa Orden de Fabricación.")
         Screen.MousePointer = 1
         GoTo 1405
     End If
     '
     J% = 0
     Call FRESHECHO("!OPERINF")
     SORTLIST.Clear
     For K% = 1 To Len(RGFSY$) Step 4
        REGI& = CVS(Mid$(RGFSY$, K%, 4))
        XX$ = REGLEIDO$("OPERFAB", REGI&)
        NOPE% = CVI(Mid$(XX$, 15, 2))
        TSOR$ = CSTRING$(Mid$(XX$, 15, 2), 1, 6, 8, 2) + Str$(REGI&)
        SORTLIST.AddItem TSOR$
     Next K%
     '
     For K% = 1 To SORTLIST.ListCount
        REGI& = Val(Mid$(SORTLIST.List(K% - 1), 7))
        XX$ = REGLEIDO$("OPERFAB", REGI&)
        CAPROD# = CVD(Mid$(XX$, 79, 8))
        CAREA# = CVD(Mid$(XX$, 87, 8))
        SALDO# = CAPROD# - CAREA#: If SALDO# < 0 Then SALDO# = 0
        '
        If SALDO# > 0.05 Then
            '
            CI% = CVI(Mid$(XX$, 13, 2))
            NOPE% = CVI(Mid$(XX$, 15, 2))
            DOPE$ = Mid$(XX$, 17, 48)
            CAINF# = SALDO#
            MAQUI$ = Mid$(XX$, 73, 6)
            FEOP% = CVI(Mid$(XX$, 139, 2))
            HOOP% = CVI(Mid$(XX$, 141, 2))
            MALIB$ = "S"
            TITOT = CVS(Mid$(XX$, 131, 4))
            '
            J% = J% + 1
            YY$ = MKI$(J%) + Space$(58) + MKS$(REGI&)
            NOPX$ = AJUSTD$(Str$(NOPE%), 5) + " - " + Left$(DOPE$, 38)
            NOPX$ = NOPX$ + NUMSTRI$(SALDO#, 12, 1, 0)
            Call REPLA(YY$, NOPX$, 3, 58)
            JJ& = J%
            Call GRAREG("!OPERINF", YY$, JJ&)
            '
        End If
     Next K%
     '
     If J% < 1 Then
         Call MENSERR(24, "No hay Operaciones Programadas\para esa Orden de Fabricación.")
         Screen.MousePointer = 1
         GoTo 1405
     End If
     '
     Call SETULTREG("!OPERINF", JJ&)
     Call CIERRARCH("!OPERINF")
     '
     Screen.MousePointer = 1
     '
1415 Call SELECHO("!OPERINF/Operaciones Programadas para O/F " + NROOF$)
     NORDE& = Val(VALACT1$("!OPERINF"))
     If NORDE& < 1 Then
         GoTo 1405
     End If
     '
     XX$ = REGLEIDO$("!OPERINF", NORDE&)
     REGI& = CVS(Mid$(XX$, 61, 4))
     '
     ZZ$ = REGLEIDO$("OPERFAB", REGI&)
     RANTEINF$ = ZZ$
     '
     NROOF$ = Left$(ZZ$, 12)
     CI% = CVI(Mid$(ZZ$, 13, 2))
     NOPE% = CVI(Mid$(ZZ$, 15, 2))
     DOPE$ = Mid$(ZZ$, 17, 48)
        CAPROD# = CVD(Mid$(ZZ$, 79, 8))
        CAREA# = CVD(Mid$(ZZ$, 87, 8))
        SALDO# = CAPROD# - CAREA#: If SALDO# < 0 Then SALDO# = 0
     CAINF# = SALDO#
     '
     YY$ = ZZ$ + Space$(320)   ' espacios para observaciones
     Call REPLA(YY$, MKI$(0), 73, 2) ' NUMERO DE PROVEEDOR
     Call REPLA(YY$, MKS$(0), 75, 4) ' NUMERO DE REMITO
     Call REPLA(YY$, String$(178, 0), 79, 178)
     Call REPLA(YY$, MKD$(CAINF#), 79, 8) ' CANTIDAD PROYECTADA
     Call REPLA(YY$, MKI$(HO%), 135, 2)
     Call REPLA(YY$, MKI$(HO%), 139, 2)
     Call REPLA(YY$, Space$(30), 151, 30)
     Call REPLA(YY$, String$(76, 0), 181, 76)
     '
1420 XX$ = OBJPLA$("TERCEROPER", YY$)
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
     OBSERVA$ = Mid$(XX$, 257, 320)
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
     IXT$ = MKS$(NORSE&) + TRIM$(NROOF$) + " - " + TRIM$(DOPE$) + " " + TRIM$(CODEXT$(CI%)) + " - " + RASOCLI$((-1) * NPRO%)
     '
     FEBA% = HO%
     HOBA% = HORANUM%(Time$)
     USUA% = USNBR% Mod 100
     DBAJA$ = MKI$(FEBA%) + MKI$(HOBA%) + Chr$(USUA%) + Chr$(1) + "Tercerizacion"
     '
     NORFB$ = Left$(XX$, 12)
     CIIK% = CVI(Mid$(XX$, 13, 2))
     NOPE% = CVI(Mid$(XX$, 15, 2))
     CAREAL# = CAINF#
     '
     ZZ$ = REGLEIDO$("OPERFAB", REGI&)
     NORFA$ = Left$(ZZ$, 12)
     If ZZ$ <> RANTEINF$ Or NORFA$ <> NORFB$ Or CVI(Mid$(ZZ$, 15, 2)) <> NOPE% Or Mid$(ZZ$, 177, 6) <> String$(6, 0) Then
        Call CIERRARCH("OPERFAB")
        Call MENSERR(24, "Imposible Generar esta Orden de Servicio.\  \Causa Probable: Se Ha Registrado\un Parte para esta Operación\desde otra Terminal de la Red.\  \Vuelva a Generar la Orden de Servicio.")
        GoTo 1999
     End If
     '
     OBSERX$ = "Orden de Servicio " + TRIM$(Str$(NORSE&)) + " - Orden de Fabricación " + NORFA$
     Call REPLA(OBSERVA$, OBSERX$, 1, 64)
     '
' Marca como "Cumplida" Operación en OPERFAB.DAT
     SALRE# = CVD(Mid$(ZZ$, 79, 8)) - CAREAL#
     If SALRE# < 0 Then SALRE# = 0
     '
     NUOROP% = Asc(Mid$(ZZ$, 67, 1))
     Call REPLA(ZZ$, MKD$(CAREAL#), 87, 8)
     Call REPLA(ZZ$, MKD$(SALRE#), 95, 8)
     Call REPLA(ZZ$, "3ROS  ", 151, 6)
     Call REPLA(ZZ$, MKI$(HO%), 157, 2)
     Call REPLA(ZZ$, MKI$(HORANUM%(Time$)), 159, 2)
     Call REPLA(ZZ$, MKS$(0), 161, 4)
     Call REPLA(ZZ$, String$(8, 0), 165, 8)
     Call REPLA(ZZ$, MKS$(0), 173, 4)
     Call REPLA(ZZ$, DBAJA$, 177, 22)
     Call GRAREG("OPERFAB", ZZ$, REGI&)
     '
' Si se tercerizó un parcial, genera registro en OPERFAB.DAT
' por el saldo pendiente
     If SALRE# > 0.01 Then
        NUOROP% = 1 + Asc(Mid$(ZZ$, 67, 1))
        CIIK% = CVI(Mid$(ZZ$, 13, 2))
        NOPE% = CVI(Mid$(ZZ$, 15, 2))
        DALTA$ = Chr$(NUOROP%) + Mid$(DBAJA$, 2, 5)
        INFOBLA$ = Space$(6) + String$(26, 0) + Space$(16)
        PROGBLA$ = Space$(6) + String$(26, 0)
        MINVARI = TIESTAN(CIIK%, NOPE%, SALRE#, NUOROP%)
        '
        Call REPLA(ZZ$, DALTA$, 67, 6)
        Call REPLA(ZZ$, MKD$(SALRE#), 79, 8)
        Call REPLA(ZZ$, MKD$(0), 87, 8)
        Call REPLA(ZZ$, MKD$(SALRE#), 95, 8)
        Call REPLA(ZZ$, String$(32, 0), 103, 32)
        Call REPLA(ZZ$, MKS$(MINVARI), 111, 4)
        Call REPLA(ZZ$, MKS$(MINVARI), 131, 4)
        Call REPLA(ZZ$, String$(16, 0), 135, 16)
        Call REPLA(ZZ$, INFOBLA$, 151, 48)
        Call REPLA(ZZ$, PROGBLA$, 225, 32)
        Call REGAPP("OPERFAB", ZZ$)
        AGREGA% = 1
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
     FECHEMI% = HO%
     DETOCOM$ = REGBLAN$("OCOMDETA")
     Call REPLA(DETOCOM$, MKS$(NUOCONUE&), 1, 4)
     Call REPLA(DETOCOM$, MKI$(HO%), 5, 2)
     Call REPLA(DETOCOM$, MKI$(NPRO%), 7, 2)
     Call REPLA(DETOCOM$, MKI$(CI%), 9, 2)
     Call REPLA(DETOCOM$, UNIMED$(CI%), 11, 4)
     CACOMPRA = CAREAL#
     Call REPLA(DETOCOM$, MKS$(CACOMPRA), 55, 4)
     Call REPLA(DETOCOM$, UNIMED$(CI%), 59, 4)
     Call REPLA(DETOCOM$, MKS$(1), 63, 4)
     Call REPLA(DETOCOM$, String$(28, 0), 67, 28) ' precios en cero
     Call REPLA(DETOCOM$, MKI$(FENTRESOL%), 95, 2)
     Call REPLA(DETOCOM$, MKI$(1), 97, 2)         ' consigna MP
     Call REPLA(DETOCOM$, MKS$(0), 99, 4)         ' cant recibida
     Call REPLA(DETOCOM$, MKI$(NOPE%), 103, 2)
     Call REPLA(DETOCOM$, NORFA$, 105, 12)
     Call REPLA(DETOCOM$, MKS$(NORSE&), 117, 4)
     Call REPLA(DETOCOM$, String$(2, 0), 121, 2)
     Call REPLA(DETOCOM$, Chr$(2), 123, 1)        ' tipo fasson
     Call REPLA(DETOCOM$, Chr$(0), 124, 1)        ' activa
     '
123  ENCAOCOM$ = REGBLAN$("OCOMENCA")
     Call REPLA(ENCAOCOM$, MKS$(NUOCONUE&), 1, 4)
     FEX = FECHEMI%
     REFE$ = FECHATEX$(FEX) + " - " + TRIM$(NORFA$) + " - " + RASOCLI$((-1) * NPRO%)
     Call REPLA(ENCAOCOM$, REFE$, 5, 44)
     Call REPLA(ENCAOCOM$, MKI$(FECHEMI%), 49, 2)
     Call REPLA(ENCAOCOM$, MKI$(NPRO%), 51, 2)
     Call REPLA(ENCAOCOM$, Chr$(2), 53, 1)
     Call REPLA(ENCAOCOM$, Chr$(0), 54, 1)  ' abierta
     '
     HO% = HOY(HOS$)
     HP = 100 * Val(Left$(Time$, 2)) + Val(Mid$(Time$, 4, 2))
     Call REPLA(ENCAOCOM$, MKI$(HO%), 55, 2)
     Call REPLA(ENCAOCOM$, MKI$(Int(HP)), 57, 2)
     Call REPLA(ENCAOCOM$, MKI$(USNBR%), 59, 2)
     Call REPLA(ENCAOCOM$, String$(92, 0), 61, 92)
     CPAG% = CPAGCLI%((-1) * NPRO%)
     Call REPLA(ENCAOCOM$, MKI$(CPAG%), 153, 2) ' importes
     '
     Call REPLA(ENCAOCOM$, ATEN$, 163, 30)
     Call REPLA(ENCAOCOM$, OBSERVA$, 193, 320)
     '
     RENCA& = 1 + ULTREG&("OCOMENCA")
     PRIDETA& = 1 + ULTREG&("OCOMDETA")
     ULTDETA& = PRIDETA&
     '
     ' grabar encabezado
     Call REPLA(ENCAOCOM$, MKS$(PRIDETA&), 155, 4)
     Call REPLA(ENCAOCOM$, MKS$(ULTDETA&), 159, 4)
     Call GRAREG("OCOMENCA", ENCAOCOM$, RENCA&)
     Call CIERRARCH("OCOMENCA")
     '
     ' grabar detalle
     Call REPLA(DETOCOM$, MKS$(RENCA&), 125, 4)
     Call REGAPP("OCOMDETA", DETOCOM$)
     '
1498 Call CIERRARCH("OCOMDETA")
     Call CIERRARCH("OCOMENCA")
     Call CIERRARCH("OPERFAB")
     Call CIERRARCH("OPERTER")
     '
     Call CIERRARCH("*.*")
     Call BAJALDISCO
     '
     JJ& = 0
     For U& = 193 To 512 Step 64
       TTXX$ = Mid$(ENCAOCOM$, U&, 64)
       JJ& = JJ& + 1
       Call GRAREG("!OBSERVOF", TTXX$, JJ&)
     Next U&
     Call SETULTREG("!OBSERVOF", JJ&)
     Call CIERRARCH("!OBSERVOF")
     '
     ' imprimir Orden de Servicio
     '
     FORIPRE$ = TRIM$(CONTROL$("", "FORSERVI"))
     If FORIPRE$ <> "" Then
        LU$ = LUDAT$
        If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
          RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
          Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Orden de Servicio:\  \'" + RMCC$ + "'.")
          GoTo 1600
        End If
        CANCELPRINT% = 0
        If CONTROL$("ORSERVI", "SUPRINT") = "SI" Then
          CANCELPRINT% = 1
        End If
        '
        FEX = FECHEMI%
        FECHDOC$ = FECHATEX$(FEX)
        NUMEDOC$ = TRIM$(Str$(NORSE&))
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
        CODPARAM$(9) = "NOP-OPER": DOCPARAM$(9) = Str$(NOPE%)
        CODPARAM$(10) = "DES-OPER": DOCPARAM$(10) = DOPE$
        CODPARAM$(11) = "CAN-COMP": DOCPARAM$(11) = Str$(CAINF#)
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
        UNID$ = UNIMED$(CI%)
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
     FORIPRE$ = TRIM$(CONTROL$("", "FORMOCOM"))
     If FORIPRE$ <> "" Then
         LU$ = LUDAT$
         If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
           RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
           Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Orden de Compra:\  \'" + RMCC$ + "'.")
           GoTo 1999
         End If
         CANCELPRINT% = 0
         If CONTROL$("OCOMPRA", "SUPRINT") = "SI" Then
           CANCELPRINT% = 1
         End If
         '
         Screen.MousePointer = 11
         FECHDOC$ = FETEXTO$(FECHEMI%)
         NUMEDOC$ = TRIM$(Str$(NUOCONUE&))
         While Len(NUMEDOC$) < 6
           NUMEDOC$ = "0" + NUMEDOC$
         Wend
         TIPODOC$ = "Nota de Pedido"
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
         CAPARDOC% = 14
         '
         DOCPARAM$(1) = FORPAG$
         DOCPARAM$(2) = OBSE$
         DOCPARAM$(3) = Str$(TOTANE#)
         DOCPARAM$(4) = Str$(ALIVA)
         DOCPARAM$(5) = Str$(TOTIVA#)
         DOCPARAM$(6) = Str$(TOTATO#)
         DOCPARAM$(7) = Mid$(X1$, 147, 20)
         DOCPARAM$(8) = Mid$(X1$, 67, 40)
         DOCPARAM$(9) = Mid$(X1$, 167, 40)
         DOCPARAM$(10) = Mid$(X1$, 107, 40)
         DOCPARAM$(11) = Str$(TOTANEDES#)
         DOCPARAM$(12) = TRIM$(Mid$(ENCAOCOM$, 163, 30)) ' ATENCION SR.
         DOCPARAM$(13) = TRIM$(PORDESCUENTO$)
         DOCPARAM$(14) = Str$(IDESCUEN#)
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
         For U& = 193 To 512 Step 64
           TTXX$ = Mid$(ENCAOCOM$, U&, 64)
           JJ& = JJ& + 1
           Call GRAREG("!OBSERVOF", TTXX$, JJ&)
         Next U&
         Call SETULTREG("!OBSERVOF", JJ&)
         Call CIERRARCH("!OBSERVOF")
         '
         X$ = DETOCOM$
         CI0% = CVI(Mid$(X$, 9, 2))
         CANTU = CVS(Mid$(X$, 55, 4))
         PREUNID# = CVD(Mid$(X$, 67, 8))
         PORDE = CVS(Mid$(X$, 75, 4))
         FEX = CVI(Mid$(X$, 95, 2))
         CONSIMAT% = 0: If TRIM$(Mid$(X$, 97, 2)) <> "" Then CONSIMAT% = 1
         '
         CAITAB1% = CAITAB1% + 1
         TETABU1$(1, CAITAB1%) = CODEXT$(CI0%)
         TETABU1$(2, CAITAB1%) = DESCRIT0$(CI0%)
         TETABU1$(3, CAITAB1%) = Mid$(X$, 59, 4)
         TETABU1$(4, CAITAB1%) = TRIM$(Str$(CANTU))
         TETABU1$(5, CAITAB1%) = FECHATEX$(FEX)
         TETABU1$(6, CAITAB1%) = Mid$(X$, 97, 2)  'consigna m.prima
         TETABU1$(7, CAITAB1%) = Str$(PREUNID#)
         TETABU1$(8, CAITAB1%) = Str$(PORDE)
         TETABU1$(9, CAITAB1%) = Str$(PREUNID# * (1 - PORDE / 100))
         TETABU1$(10, CAITAB1%) = Str$(CANTU * PREUNID# * (1 - PORDE / 100))
         TETABU1$(11, CAITAB1%) = Str$(CANTU * PESUNI(CI0%))
         TETABU1$(12, CAITAB1%) = FEPLANO$(CI0%)
         TETABU1$(13, CAITAB1%) = TRIM$(Str$(CAITAB%))
         '
         CAITAB1% = CAITAB1% + 1
         For KKII% = 1 To 13: TETABU1$(KKII%, CAITAB1%) = "": Next KKII%
         TETABU1$(2, CAITAB1%) = "Operacion: " + TRIM$(DOPE$)
         '
         Call PRINTDOC("FORMOCOM")   ' Orden de Compra
         '
     End If
     '
1999 Call CIERRARCH("*.*")
     Call BAJALDISCO
     GoTo 1410
     '
End Sub

Function PRONUOCO&()
   '
   PRONU& = 1
   REGXX& = ULTREG&("OCOMENCA")
   If REGXX& > 0 Then
     PRONU& = 1 + CVS(Left$(REGLEIDO$("OCOMENCA", REGXX&), 4))
   End If
   If PRONU& <= REGXX& Then
     Screen.MousePointer = 11
     For U& = 1 To REGXX&
       X0$ = REGLEIDO$("OCOMENCA", U&)
       If CVS(Left$(X0$, 4)) >= PRONU& Then
         PRONU& = 1 + CVS(Left$(X0$, 4))
       End If
     Next U&
     Screen.MousePointer = 1
   End If
   '
   PRONUOCO& = PRONU&
   '
End Function

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
