VERSION 5.00
Begin VB.Form OPLIPRE04 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Presupuestos y Cotizaciones - Opciones"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5100
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   210
      TabIndex        =   11
      Top             =   105
      Width           =   3690
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Index           =   2
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2055
         TabIndex        =   12
         Top             =   420
         Width           =   2115
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
            Picture         =   "OPLIPRE04.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Visualizar y Re-Imprimir Presupuestos"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Marcar Como ""No Vigente"""
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
            Top             =   105
            Width           =   1425
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Consultas y Listados"
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
      Top             =   1575
      Width           =   3690
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Index           =   1
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   4
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
            Picture         =   "OPLIPRE04.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   5
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
            TabIndex        =   6
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
            Picture         =   "OPLIPRE04.frx":0614
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
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      Height          =   3270
      Left            =   4170
      ScaleHeight     =   3210
      ScaleWidth      =   1530
      TabIndex        =   7
      Top             =   0
      Width           =   1590
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   9
         Top             =   2310
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -315
            TabIndex        =   10
            Top             =   105
            Width           =   12000
         End
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
         Height          =   640
         Left            =   105
         Picture         =   "OPLIPRE04.frx":0C7E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   315
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "OPLIPRE04.frx":0DC8
         Top             =   2625
         Width           =   1515
      End
   End
End
Attribute VB_Name = "OPLIPRE04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim BRTEMP As Recordset

Sub Command1_Click()
    Command1.Enabled = False
    Screen.MousePointer = 11
    'NUEVO
    Call ABREPRESVEN
    '
    SQLX$ = " SELECT * "
    SQLX$ = SQLX$ + " FROM PRESUVEN "
    SQLX$ = SQLX$ + " ORDER BY NroPres DESC "
    '
    Set LISPRETEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
    With LISPRETEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        XXX$ = REGBLAN("INDIPRES")
        JJ& = 0: NPXA& = 0
        Do While (.EOF = False)
          NPX& = !NroPres
          NCLIE% = !NroClie
          CIXI% = !CodiInt
          FEX = CVINT(!FechEmis)
          '
          '\\\OT-06-0168-WAR-09/05/06///
          CN1$ = "": MARNOVI$ = ""
          On Error Resume Next
          CR1$ = !Cod_Rubro
          NCR1 = !NumeCRubro
          CA1$ = !Cod_Anio
          '\\\OT-06-0497-RG-21/11/06///
          MARNOVI$ = !MARNOVIG
          '\\\OT-06-0497-RG-FIN///
          On Error GoTo 0
          '
          If TRIM$(CR1$) <> "" And XVALO(NCR1) > 0 Then
            NCR2$ = TRIM$(Str$(NCR1))
            While Len(NCR2$) < 4: NCR2$ = "0" + NCR2$: Wend
            CN1$ = CR1$ + "-" + NCR2$ + "-" + CA1$ + "   "
          End If
          '\\\OT-06-0168-WAR-FIN///
          YY2$ = TRIM$(AJUSTI$(CN1$ + TRIM$(RASOCLI$(NCLIE%)), 30)) + " - "
          If NPX& <> NPXA& Then
               YY2$ = YY2$ + TRIM$(CODEXT$(CIXI%))
             Else
               YY2$ = YY2$ + "VARIOS"
          End If
          YY2$ = YY2$ + " - " + FECHATEX$(FEX)
          '\\\OT-06-0497-RG-21/11/06///
          If MARNOVI$ <> "" Then
             YY2$ = "(N/Vig) " + YY2$
          End If
          '\\\OT-06-0497-RG-FIN///
          '
          Call REPLA(XXX$, MKS$(NPX&), 1, 4)
          Call REPLA(XXX$, YY2$, 5, 60)
          If NPX& <> NPXA& Then
            JJ& = JJ& + 1
          End If
          Call GRAREG("INDIPRES", XXX$, JJ&)
          '
          NPXA& = NPX&
        .MoveNext
        Loop
      End If
    End With
    '
    Call SETULTREG("INDIPRES", JJ&)
    Call CIERRARCH("INDIPRES")
    Call FRESHECHO("INDIPRES")
    '
3   Screen.MousePointer = 1
    Call SELECHO("INDIPRES")
    NPX& = XVALO(VALACT1$("INDIPRES"))
    If NPX& < 1 Then GoTo 99
    '
    TIDOCUM$ = "Presup"
    NUDOCU$ = " " + TRIM$(Str$(NPX&)) + " -"
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
5   For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    '
    REDOCU$ = REGACT$("PDOCLST")
    URE& = ULTREG&("PDOCSPL")
    LI& = 0: LS& = URE&
6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
    GoTo 8
    '
7   Screen.MousePointer = 1
    Call MENSERR(24, "Documento no Registrado")
    GoTo 99
    '
8   TPSP$ = ""
    Screen.MousePointer = 11
    While L& > 1
      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
      If DCC& < DOCUNU& Then GoTo 9
      L& = L& - 1
    Wend
    '
9   CAPAGINAS% = 0
    For KKL& = L& To URE&
      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
      End If
    Next KKL&
    '
95  SPOOLSTRING$ = TPSP$
    For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).Top = 0
      EDITFORM.Picture1(KKI%).Refresh
      EDITFORM.Picture1(KKI%).Height = 7170
      EDITFORM.Picture1(KKI%).Width = 11700
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    ALTUPAGINA = 7170
    TOPEPAGINA = 0
    PAGINACTIVA% = 0
    CAPAGINAS% = 0
    '
10  PPXX% = InStr(TPSP$, Chr$(255))
    If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call SHOWCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
    End If
    '
    UI& = UI& + 1
    XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
    If InStr(XX$, TIDOCUM$) > 4 Then
      If InStr(XX$, NUDOCU$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 7
      End If
    End If
    '
    Screen.MousePointer = 1
    PAGINACTIVA% = 0
    For KKI% = 0 To 5
       EDITFORM.Picture1(KKI%).Visible = False
    Next KKI%
    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
    EDITFORM.Picture1(PAGINACTIVA%).Refresh
    EDITFORM.VScroll1.Min = 0
    EDITFORM.VScroll1.Value = 0
    EDITFORM.VScroll1.Max = 1
    If ALTUPAGINA > EDITFORM.Frame1.Height Then
       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
    End If
    EDITFORM.Command3.Enabled = False
    EDITFORM.Show 1
    '
    GoTo 3
    '
99  Command1.Enabled = True
    Screen.MousePointer = 1
    '
End Sub

Private Sub Command2_Click()
    '
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Sub Command3_Click()
    Command3.Enabled = False
    Screen.MousePointer = 11
    'NUEVO
    Call ABREPRESVEN
    '
    SQLX$ = " SELECT * "
    SQLX$ = SQLX$ + " FROM PRESUVEN "
    SQLX$ = SQLX$ + " ORDER BY NroPres DESC "
    '
    Set LISPRETEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
    With LISPRETEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        XXX$ = REGBLAN("INDIPRES")
        JJ& = 0: NPXA& = 0
        Do While (.EOF = False)
          NPX& = !NroPres
          NCLIE% = !NroClie
          CIXI% = !CodiInt
          FEX = CVINT(!FechEmis)
          '
          '\\\OT-06-0168-WAR-09/05/06///
          CN1$ = "": MARNOVI$ = ""
          On Error Resume Next
          CR1$ = !Cod_Rubro
          NCR1 = !NumeCRubro
          CA1$ = !Cod_Anio
          '\\\OT-06-0497-RG-21/11/06///
          MARNOVI$ = !MARNOVIG
          '\\\OT-06-0497-RG-FIN///
          On Error GoTo 0
          '
          If TRIM$(CR1$) <> "" And XVALO(NCR1) > 0 Then
            NCR2$ = TRIM$(Str$(NCR1))
            While Len(NCR2$) < 4: NCR2$ = "0" + NCR2$: Wend
            CN1$ = CR1$ + "-" + NCR2$ + "-" + CA1$ + "   "
          End If
          '\\\OT-06-0168-WAR-FIN///
          YY2$ = TRIM$(AJUSTI$(CN1$ + TRIM$(RASOCLI$(NCLIE%)), 30)) + " - "
          If NPX& <> NPXA& Then
               YY2$ = YY2$ + TRIM$(CODEXT$(CIXI%))
             Else
               YY2$ = YY2$ + "VARIOS"
          End If
          YY2$ = YY2$ + " - " + FECHATEX$(FEX)
          '\\\OT-06-0497-RG-21/11/06///
          If MARNOVI$ <> "" Then
             YY2$ = "(N/Vig) " + YY2$
          End If
          '\\\OT-06-0497-RG-FIN///
          '
          Call REPLA(XXX$, MKS$(NPX&), 1, 4)
          Call REPLA(XXX$, YY2$, 5, 60)
          If NPX& <> NPXA& Then
            JJ& = JJ& + 1
          End If
          Call GRAREG("INDIPRES", XXX$, JJ&)
          '
          NPXA& = NPX&
        .MoveNext
        Loop
       
      End If
    End With
    '
    Call SETULTREG("INDIPRES", JJ&)
    Call CIERRARCH("INDIPRES")
    Call FRESHECHO("INDIPRES")
    '
3   Screen.MousePointer = 1
    Call SELECHO("INDIPRES")
    NPX& = XVALO(VALACT1$("INDIPRES"))
    If NPX& < 1 Then GoTo 99
    '\\\OT-06-0497-RG-21/11/06///
    RESPU% = COMALTER%("Confirma Marcar Presupuesto Nro." + Str$(NPX) + " como 'No Vigente' \\ No, Consevar \ Si, Marcar")
    If RESPU% <> 2 Then GoTo 99
    '
    Call ABREPRESVEN
    SQLZ$ = "SELECT * FROM PRESUVEN"
    SQLZ$ = SQLZ$ + " WHERE NROPRES = " & NPX&
    Set BRTEMP = Ventas.OpenRecordset(SQLZ, dbOpenDynaset)
    '
    With BRTEMP
      On Error Resume Next
      .MoveFirst
      If Err > 0 Then
         On Error GoTo 0
         GoTo 3
      End If
      On Error GoTo 0
      '
      Do While Not .EOF
         .Edit
         !MARNOVIG = "*"
         .Update
      .MoveNext
      Loop
    End With
    Call COMUNI("Transacción Completada")
    '\\\OT-06-0497-RG-FIN///
    '
    GoTo 3
    '
99  Command3.Enabled = True
    Screen.MousePointer = 1
    '
End Sub

 Sub Command6_Click()
    Command6.Enabled = False
5   OPT% = COMALTER%("Opciones de Consulta de Presupuestos:\\Por Numero\Por Cliente\Por Producto\Por Zona\Listado de Control")
    If OPT% < 1 Or OPT% > 5 Then
        GoTo 99
      ElseIf OPT% = 5 Then
        Call LICONPRE
    End If
    '
    SOLOVIGEN% = ALTERNA%("Todos \ Solo Vigentes")
    If SOLOVIGEN% < 1 Or SOLOVIGEN% > 2 Then GoTo 99
    
    'NCLIE1% = 0
    '
'         Call DESHASFECHA(DES%, HAS%, PERIO$)
'         If PERIO$ = "" Then GoTo 5
'
'      Else
'10      DES% = COMESACT%: HAS% = HOY(HOS$)
'        VDEF$ = MKI$(DES%) + MKI$(HAS%)
'        OBX$ = OBJPLA$("DESHASFECHCLI", VDEF$)
'        If OBX$ = "" Then GoTo 5
'        '
'        DES% = CVI(Left$(OBX$, 2))
'        HAS% = CVI(Mid$(OBX$, 3, 2))
'        If HAS% < DES% Then
'          Call MENSERR(24, "Rango de Fechas no Válido")
'          GoTo 10
'        End If
'        NCLIE1% = CVI(Mid$(OBX$, 5, 2))
'        '
'
'    End If
'    Call SELECHO("VENDEDOR")
'    NVEND% = XVALO(VALACT1$("VENDEDOR"))
    If OPT% <> 2 Then

    Call DESHASFECHA(DES%, HAS%, PERIO$)
    If PERIO$ = "" Then Exit Sub
    
10  VDEF$ = MKI$(NCMIN%) + MKI$(NCMAX%) + Chr$(0)
    OBX$ = OBJPLA$("DESHASCLIEVEN", VDEF$)
    If OBX$ = "" Then
     NCL1% = 0: NCL2% = 0
     Exit Sub
    End If
   '
    NCL1% = CVI(Left$(OBX$, 2))
    NCL2% = CVI(Mid$(OBX$, 3, 2))
    If NCL1% < 1 Or NCL2% < 1 Or NCL2% < NCL1% Then
      Call MENSERR(24, "Rango no Válido")
      GoTo 10
    End If
    NVEND% = Asc(Mid$(OBX$, 5, 1))
20  Screen.MousePointer = 11
    Call COPYSTRU("PRESUPU", "LIPRESU")
    '
    Call ABREPRESVEN
    SQLX$ = " SELECT * "
    SQLX$ = SQLX$ + " FROM PRESUVEN "
    SQLX$ = SQLX$ + " WHERE cdbl(FechEmis) >=   " & CDbl(CVDAT(DES%)) & "  "
    SQLX$ = SQLX$ + " AND cdbl(FechEmis) <=  " & CDbl(CVDAT(HAS%)) & "   "
    If NCL1% > 0 Then
      SQLX$ = SQLX$ + " AND NroClie >= " & NCL1% & " "
      SQLX$ = SQLX$ + " AND NroClie <= " & NCL2% & " "
    End If
    If NVEND% > 0 Then
        SQLX$ = SQLX$ + " AND VENDED = " & NVEND% & " "
    End If
    SQLX$ = SQLX$ + " ORDER BY NroPres ASC "

    '
    Set PRESUTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With PRESUTEMP
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        On Error GoTo 0
        FIN& = .RecordCount
        .MoveFirst
        JJ& = 0
        x$ = REGBLAN("LIPRESU")
        Do While (.EOF = False)
          Call TRACE(20, JJ&, FIN&)
          If SOLOVIGEN = 2 Then
            If (!MARNOVIG = "*") Then GoTo 50
          End If
          On Error Resume Next
          Call REPLA(x$, MKS$(!NroPres), 1, 4)
          Call REPLA(x$, MKI$(CVINT(!FechEmis)), 5, 2)
          Call REPLA(x$, MKI$(!NroClie), 7, 2)
            Monemisi% = !MonedEmi
            If Monemisi% < 1 Then Monemisi% = 1
          Call REPLA(x$, Chr$(!ConPag), 66, 1)
          Call REPLA(x$, Chr$(!VENDED), 67, 1)
          Call REPLA(x$, Chr$(Monemisi%), 68, 1)
          Call REPLA(x$, MKI$(!CodiInt), 69, 2)
          Call REPLA(x$, MKS$(!PreList), 71, 4)
          Call REPLA(x$, !TDESCUEN, 75, 14)
          Call REPLA(x$, MKS$(!PreUnit), 97, 4)
          Call REPLA(x$, MKS$(!Cantit), 101, 4)
            ITOPES# = !PreUnit * !Cantit * TICAMONE(Monemisi%)
          Call REPLA(x$, MKD$(ITOPES#), 105, 8)
          Call REPLA(x$, MKI$(CVINT(!FechVent)), 113, 2)
          Call REPLA(x$, MKS$(!NroPedi), 115, 4)
          On Error GoTo 0
          '
          JJ& = JJ& + 1
          Call GRAREG("LIPRESU", x$, JJ&)
50      .MoveNext
        Loop
      End If
      On Error GoTo 0
      Call SETULTREG("LIPRESU", JJ&)
      Call CIERRARCH("LIPRESU")
    End With
    '
    DESX$ = FECHATEX$(DES%)
    HASS$ = FECHATEX$(HAS%)
    DEPER$ = "Corresponde a: " + DESX$ + " - " + HASS$ + " / "
    '
    Screen.MousePointer = 1
    Select Case OPT%
        Case 1
            Call FILESORT("LIPRESU", "", 1, 1, "Asc")
            DEPER$ = DEPER$ + "por Numero"
        Case 2
            Call FILESORT("LIPRESU", "", 3, 3, "Asc")
            DEPER$ = DEPER$ + "por Cliente"
        Case 3
            Call FILESORT("LIPRESU", "", 11, 11, "Asc")
            DEPER$ = DEPER$ + "por Producto"
        Case 4
            SELE$ = REGSEL$("PROVINB")
            If Len(SELE$) <= 4 Then GoTo 99
            CPROVI$ = Left$(SELE$, 1): DPROVI$ = Mid$(SELE$, 2, 25)
            FIN& = ULTREG&("LIPRESU")
            JJ& = 0
            For U& = 1 To FIN&
              Call TRACE(20, 1, FIN&)
              x$ = REGLEIDO$("LIPRESU", U&)
              NCLIE% = CVI(Mid$(x$, 7, 2))
              If PROVCLI$(NCLIE%) = CPROVI$ Then
                JJ& = JJ& + 1
                Call GRAREG("LIPRESU", x$, JJ&)
              End If
            Next U&
            Call SETULTREG("LIPRESU", JJ&)
            DEPER$ = DEPER$ + "Zona = " + DPROVI$
    End Select
    '
    Call HEADERS("LIPRESU", "")
    Call HEADERS("LIPRESU", DEPER$)
    '
    Call CONECRUN("*LIPRESU", "Listado de Presupuestos")
    '
99  Screen.MousePointer = 1
    Command6.Enabled = True
  End If
End Sub

Sub LICONPRE()
    '
    Dim NPRES&(32767)
    Dim FPRES%(32767)
    Dim NCLIEN%(32767)
    Dim ITOT#(32767)
    Dim FVEN%(32767)
    Dim NOVEN&(32767)
    Dim IMPOVE#(32767)
    '
    SOLOVIGEN% = ALTERNA%("Todos \ Solo Vigentes")
    If SOLOVIGEN% < 1 Or SOLOVIGEN% > 2 Then GoTo 99

    OPT% = 5
    Call DESHASFECHA(DES%, HAS%, PERIO$)
    If PERIO$ = "" Then Exit Sub
    '
20  Screen.MousePointer = 11
    '
    Call ABREPRESVEN
    SQLX$ = " SELECT * "
    SQLX$ = SQLX$ + " FROM PRESUVEN "
    SQLX$ = SQLX$ + " WHERE cdbl(FechEmis) >=   " & CDbl(CVDAT(DES%)) & "  "
    SQLX$ = SQLX$ + " AND cdbl(FechEmis) <=  " & CDbl(CVDAT(HAS%)) & "   "
    If NCLIE1% > 0 Then
      SQLX$ = SQLX$ + " AND NroClie = " & NCLIE1% & " "
    End If
    SQLX$ = SQLX$ + "ORDER BY NroPres ASC "
    '
    Set PRESUTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With PRESUTEMP
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        On Error GoTo 0
        FIN& = .RecordCount
        .MoveFirst
        JJ& = 0: JJ1% = 0
        Do While (.EOF = False)
          Call TRACE(20, JJ&, FIN&)
          If SOLOVIGEN = 2 Then
            If !MARNOVIG = "*" Then GoTo 50 ' NO FUNCIONABA EN LA SELECT OD-09-05-08
          End If
          On Error Resume Next
          Monemisi% = !MonedEmi
          ITOPES# = !PreUnit * !Cantit * TICAMONE(Monemisi%)
          On Error GoTo 0
          '
          NPRES1& = !NroPres
          If NPRES1& <> ULTPRE& Then
              JJ1% = JJ1% + 1
              NPRES&(JJ1%) = NPRES1&
              ULTPRE& = NPRES1&
              FPRES%(JJ1%) = CVINT(!FechEmis)
              NCLIEN%(JJ1%) = !NroClie
              ITOT#(JJ1%) = ITOPES#
              FVEN%(JJ1%) = CVINT(!FechVent)
              NOVEN&(JJ1%) = !NroPedi
              IMPOVE#(JJ1%) = 0
          
              If !NroPedi > 0 Then
                 ' OBTIENE CUANTO DEL PRESUPUESTO SE TRANSFORMO EN PEDIDO
                 SQLX$ = " SELECT * FROM PEDDETA "
                 SQLX$ = SQLX$ + " WHERE NroPed =   " & !NroPedi & "  "
                 Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
                 With PEDTEMP
                    On Error Resume Next
                    .MoveFirst
                    If Err < 1 Then
                       On Error GoTo 0
                       Do While Not .EOF
                          MONE% = !MoneEmis
                          SUBTOTALPED# = !PreUnid * !CantEnt * TICAMONE(MONE%)
                          IMPOVE#(JJ1%) = IMPOVE#(JJ1%) + SUBTOTALPED#
                       .MoveNext
                       Loop
                    End If
                    On Error GoTo 0
                 End With
              End If
            Else
              ITOT#(JJ1%) = ITOT#(JJ1%) + ITOPES#
          End If
          '
          JJ& = JJ& + 1
50      .MoveNext
        Loop
      End If
      On Error GoTo 0
    End With
    '
    REBLA$ = REGBLAN$("LICONPRE")
    For U& = 1 To JJ1%
      x$ = REBLA$
      Call REPLA(x$, MKS$(NPRES&(U&)), 1, 4)
      Call REPLA(x$, MKI$(FPRES%(U&)), 5, 2)
      Call REPLA(x$, MKI$(NCLIEN%(U&)), 7, 2)
      Call REPLA(x$, MKD$(ITOT#(U&)), 9, 8)
      Call REPLA(x$, MKI$(FVEN%(U&)), 17, 2)
      Call REPLA(x$, MKS$(NOVEN&(U&)), 19, 4)
      Call REPLA(x$, MKD$(IMPOVE#(U&)), 23, 8)
      '
      NOAPRO# = ITOT#(U&) - IMPOVE#(U&)
      If NOAPRO# < 0 Then NOAPRO# = 0
      '
      Call REPLA(x$, MKD$(NOAPRO#), 31, 8)
      Call GRAREG("LICONPRE", x$, U&)
    Next U&
    UL& = JJ1%
    Call SETULTREG("LICONPRE", UL&)
    '
    Call HEADERS("LICONPRE", "")
    Call HEADERS("LICONPRE", "Corresponde a: " + DESX$ + " - " + HASS$)
    Call CONECRUN("*LICONPRE", "Control de Presupuestos y Ventas")
    '
99  Screen.MousePointer = 1
    '
End Sub
'
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


