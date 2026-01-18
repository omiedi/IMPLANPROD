VERSION 5.00
Begin VB.Form OPLIPRESU 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Presupuestos y Cotizaciones - Opciones"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5160
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4095
      ScaleHeight     =   75
      ScaleWidth      =   810
      TabIndex        =   8
      Top             =   1240
      Width           =   875
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command2 
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
      Height          =   750
      Left            =   4095
      Picture         =   "OPLIPRESU.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   420
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Consulas y Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Index           =   1
      Left            =   210
      TabIndex        =   0
      Top             =   315
      Width           =   3690
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Index           =   1
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   5
         Top             =   840
         Width           =   2430
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
            Picture         =   "OPLIPRESU.frx":014A
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Listado Presupuestos Emitidos por Rango de Fechas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Listados por Rango de Fechas"
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
            TabIndex        =   7
            Top             =   105
            Width           =   1635
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Index           =   0
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   1
         Top             =   420
         Width           =   2430
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
            Height          =   600
            Left            =   0
            Picture         =   "OPLIPRESU.frx":0454
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Visualizar y Re-Imprimir Presupuestos"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Consultas y Re-Impresion"
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
            TabIndex        =   3
            Top             =   0
            Width           =   1425
         End
      End
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   3540
      Picture         =   "OPLIPRESU.frx":0ABE
      Top             =   1480
      Width           =   2010
   End
End
Attribute VB_Name = "OPLIPRESU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Command1.Enabled = False
    Screen.MousePointer = 11
    JJ& = 0: NPX& = 0
    FIN& = ULTREG&("PRESUPU")
    If FIN& < 1 Then
      Call MENSERR(24, "Archivo Vacío")
      GoTo 99
    End If
    '
    FINI1& = ULTREG&("PRESUPU")
    XX$ = REGLEIDO$("PRESUPU", FINI1&)
    NPX1& = CVS(Left$(XX$, 4))
    FINI2& = ULTREG&("INDIPRES")
    YY$ = REGLEIDO$("INDIPRES", FINI2&)
    NPX2& = CVS(Left$(YY$, 4))
    If NPX1& = NPX2& Then GoTo 3
    '
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
2     XX$ = REGLEIDO$("PRESUPU", U&)
      NPY& = CVS(Left$(XX$, 4))
      If NPY& > 0 Then
        If NPY& <> NPX& Then
          NPX& = NPY&
          NCLIE% = CVI(Mid$(XX$, 7, 2))
          CIXI% = CVI(Mid$(XX$, 69, 2))
            '
            'Call RECODOSI(CIXI%, CIXJ%, NPX&)
            'If CIXJ% <> CIXI% Then
            '  Call REPLA(XX$, MKI$(CIXJ%), 69, 2)
            '  Call GRAREG("PRESUPU", XX$, U&)
            '  GoTo 2
            'End If
            '
          FEX = CVI(Mid$(XX$, 5, 2))
          YY$ = MKS$(NPX&) + TRIM$(RASOCLI$(NCLIE%)) + " - "
          If CVS(Left$(REGLEIDO$("PRESUPU", U& + 1), 4)) <> NPX& Then
              YY$ = YY$ + TRIM$(CODEXT$(CIXI%))
            Else
              YY$ = YY$ + "VARIOS"
          End If
          YY$ = YY$ + " - " + FECHATEX$(FEX)
          JJ& = JJ& + 1
          Call GRAREG("INDIPRES", YY$, JJ&)
        End If
      End If
    Next U&
    Call SETULTREG("INDIPRES", JJ&)
    '
3   Screen.MousePointer = 1
    Call SELECHO("INDIPRES")
    NPX& = Val(VALACT1$("INDIPRES"))
    If NPX& < 1 Then GoTo 99
    '
    TIDOCUM$ = "Presup"
    TIDOCUM2$ = "F. Proforma"
    NUDOCU$ = " " + TRIM$(Str$(NPX&)) + " -"
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Or InStr(XX$, TIDOCUM2$) > 4 Then
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
99  Command1.Enabled = True
    '
End Sub

Private Sub Command2_Click()
    '
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command6_Click()
    '
    Command6.Enabled = False
    '
    Dim NPRES&(32767)
    Dim FPRES%(32767)
    Dim NCLIEN%(32767)
    Dim ITOT#(32767)
    Dim FVEN%(32767)
    Dim NOVEN&(32767)
    '
    '\\\OT-5-550-FG-29/07/05///
    OP% = COMALTER%("Elija una opción de consulta de Presupuestos:\\Por Número\Por Cliente\Por Producto\Por Zona\Control")
    '\\\OT-5-550-FG-FIN///
    '
10  FEX = HOY(HO$)
    HAS% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    VDEF$ = MKI$(DES%) + MKI$(HAS%)
    '
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    DES% = CVI(Left$(OBX$, 2))
    FEX = DES%: DESX$ = FECHATEX$(FEX)
    HAS% = CVI(Mid$(OBX$, 3, 2))
    FEX = HAS%: HASS$ = FECHATEX$(FEX)
    If HAS% < DES% Then
      Call MENSERR(24, "Rango de Fechas no Válido")
      GoTo 10
    End If
    '
    Screen.MousePointer = 11
    Call COPYSTRU("PRESUPU", "LIPRESU")
    JJ& = 0
    JJ1% = 0: ULTPRE& = 0
    FIN& = ULTREG&("PRESUPU")
    '
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("PRESUPU", U&)
      FEMI% = CVI(Mid$(X$, 5, 2))
      If FEMI% >= DES% Then
        If FEMI% <= HAS% Then
          MONEMISI% = Asc(Mid$(X$, 68, 1))
          If MONEMISI% < 1 Then
            MONEMISI% = 1
            Call REPLA(X$, Chr$(MONEMISI%), 68, 1)
          End If
          ITOPES# = CVS(Mid$(X$, 97, 4)) * CVS(Mid$(X$, 101, 4)) * TICAMONE(MONEMISI%)
          Call REPLA(X$, MKD$(ITOPES#), 105, 8)
          JJ& = JJ& + 1
          Call GRAREG("LIPRESU", X$, JJ&)
          '
          NPRES1& = CVS(Mid$(X$, 1, 4))
          If NPRES1& <> ULTPRE& Then
            JJ1% = JJ1% + 1
            NPRES&(JJ1%) = NPRES1&
            ULTPRE& = NPRES1&
            FPRES%(JJ1%) = CVI(Mid$(X$, 5, 2))
            NCLIEN%(JJ1%) = CVI(Mid$(X$, 7, 2))
            ITOT#(JJ1%) = ITOPES#
            FVEN%(JJ1%) = CVI(Mid$(X$, 113, 2))
            NOVEN&(JJ1%) = CVS(Mid$(X$, 115, 4))
          Else
            ITOT#(JJ1%) = ITOT#(JJ1%) + ITOPES#
          End If
          '
        End If
      End If
    Next U&
    Call SETULTREG("LIPRESU", JJ&)
    Call CIERRARCH("LIPRESU")
    '
    REBLA$ = REGBLAN$("LICONPRE")
    For U& = 1 To JJ1%
      X$ = REBLA$
      Call REPLA(X$, MKS$(NPRES&(U&)), 1, 4)
      Call REPLA(X$, MKI$(FPRES%(U&)), 5, 2)
      Call REPLA(X$, MKI$(NCLIEN%(U&)), 7, 2)
      Call REPLA(X$, MKD$(ITOT#(U&)), 9, 8)
      Call REPLA(X$, MKI$(FVEN%(U&)), 17, 2)
      Call REPLA(X$, MKS$(NOVEN&(U&)), 19, 4)
      '
      '\\\OT-05-0756-WAR-25/11/05///
        Call REPLA(X$, MKI$(1), 39, 2)
      '\\\OT-05-0756-WAR-FIN///
      Call REPLA(X$, String$(16, 0), 23, 8)
      If NOVEN&(U&) > 0 Then
          Call REPLA(X$, MKD$(ITOT#(U&)), 23, 8)
          Call REPLA(X$, MKI$(1), 41, 2) '\\\OT-05-0756-WAR-25/11/05///
        Else
          Call REPLA(X$, MKD$(ITOT#(U&)), 31, 8)
          Call REPLA(X$, MKI$(1), 43, 2) '\\\OT-05-0756-WAR-25/11/05///
      End If
      Call GRAREG("LICONPRE", X$, U&)
    Next U&
    UL& = JJ1%
    Call SETULTREG("LICONPRE", UL&)
    '
    Call HEADERS("LIPRESU", "")
    '\\\OT-5-550-FG-29/07/05///
    DEPER$ = "Corresponde a: " + DESX$ + " - " + HASS$ + " / "
    '
    Select Case OP%
      Case 1
        Call FILESORT("LIPRESU", "", 1, 1, "ASC")
        DEPER$ = DEPER$ + "por Numero"
      Case 2
        Call FILESORT("LIPRESU", "", 3, 3, "ASC")
        DEPER$ = DEPER$ + "por Cliente"
      Case 3
        Call FILESORT("LIPRESU", "", 11, 11, "ASC")
        DEPER$ = DEPER$ + "por Producto"
      Case 4
        SELE$ = REGSEL$("PROVINB")
        If Len(SELE$) <= 4 Then GoTo 99
        CPROVI$ = Left$(SELE$, 1): DPROVI$ = Mid$(SELE$, 2, 25)
        FIN& = ULTREG&("LIPRESU")
        JJ& = 0
        For U& = 1 To FIN&
          Call TRACE(20, 1, FIN&)
          X$ = REGLEIDO$("LIPRESU", U&)
          NCLIE% = CVI(Mid$(X$, 7, 2))
          If PROVCLI$(NCLIE%) = CPROVI$ Then
            JJ& = JJ& + 1
            Call GRAREG("LIPRESU", X$, JJ&)
          End If
        Next U&
        Call SETULTREG("LIPRESU", JJ&)
        DEPER$ = DEPER$ + "Zona = " + DPROVI$
      Case 5
        Call HEADERS("LICONPRE", "")
        Call HEADERS("LICONPRE", "Corresponde a: " + DESX$ + " - " + HASS$)
        Call CONECRUN("*LICONPRE", "Control de Presupuestos y Ventas")
        GoTo 99
    End Select
    Call HEADERS("LIPRESU", DEPER$)
    '\\\OT-5-550-FG-FIN///
    '
    Call CONECRUN("*LIPRESU", "Listado de Presupuestos")
    '
99  Screen.MousePointer = 1
    Command6.Enabled = True
    '
End Sub

Sub RECODOSI(CIXI%, CIXJ%, NPX&)
   '
   CIBIN$ = MKI$(CIXI%)
   CIXJ% = CIXI%
   If Asc(Mid$(CIBIN$, 2, 1)) = 0 Then
     If NPX& > 1600 Then
       If NPX& <= 1614 Then
         If NPX& = 1601 Then CIXJ% = CODINT%("041130M-PVT0")
         If NPX& = 1602 Then CIXJ% = CODINT%("042121T-PHH0")
         If NPX& = 1603 Then CIXJ% = CODINT%("041015M-PVT0")
         If NPX& = 1604 Then CIXJ% = CODINT%("041015M-PVT0")
         If NPX& = 1605 Then CIXJ% = CODINT%("041130M-PVT0")
         If NPX& = 1606 Then CIXJ% = CODINT%("041130M-PVT0")
         If NPX& = 1607 Then CIXJ% = CODINT%("042501T-PHH0")
         If NPX& = 1608 Then CIXJ% = CODINT%("44234")
         If NPX& = 1609 Then CIXJ% = CODINT%("44234")
         If NPX& = 1610 Then CIXJ% = CODINT%("LD54S2-PHCA-D13")
         If NPX& = 1611 Then CIXJ% = CODINT%("044501T-PTT0")
         If NPX& = 1612 Then CIXJ% = CODINT%("044251T-PTT0")
         If NPX& = 1613 Then CIXJ% = CODINT%("041130M-PVT0")
         If NPX& = 1614 Then CIXJ% = CODINT%("A4190")
       End If
     End If
   End If
   If CIXJ% < 1 Then CIXJ% = CIXI%
   '
End Sub

