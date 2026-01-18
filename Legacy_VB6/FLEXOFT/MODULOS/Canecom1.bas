Attribute VB_Name = "CANECOM1"
Dim LOST$(10)
Dim LOSTN(10)
Dim CII1%(1024), CANT1(1024), PAD1%(1024)
Dim INIPER%(52), CAPER(52), DEPER$(52)
'
Dim LOREC As Integer
Dim CACOL As Integer
Dim COSEMACT As Integer
Dim HOO As Integer
Dim HO As String
Dim DMX%(12)
'
Dim DISPACT(16384)
Dim LOREPOS(16384)
Dim DPI%(16384)
Dim PROVARTI%(16384)
'
Dim SST As String * 40
Dim M0 As String * 128
Dim RA As String * 128
Dim RB As String * 256
Dim NESTRU As Long
Dim PUNTCI As String
Dim PUNTRG As String
'
Dim SEMANAL As Integer
Dim PERIO As String
Dim PLADISTA As Integer
Dim REPOMIN As Single
'
Dim CTX As Integer
Dim CACO As Integer
Dim PUNTE$(16384)
'
Dim PROVINT%(32767)
Public OPCIPLAN%         ' 0=solo entregas / 1=entregas o fabricacion
Public MODOPLAN%         ' 0=programa entregas / 1=plan maestro de fabricacion

Sub PARACAL()

'   ************* determinacion parametros de calculo ******************

    CT$ = UCase$(CONTROL$("CANECOM", "PERIODO"))
    SEMANAL% = 0: FORCANEC.Option5.Value = True
    If InStr(CT$, "SEMAN") > 0 Then
        SEMANAL% = 1
        FORCANEC.Option4.Value = True
    End If
    '
    CT$ = UCase$(CONTROL$("CANECOM", "PLADISTA"))
    PLADISTA% = Val(CT$)
    If PLADISTA% < 0 Then PLADISTA% = 0
    If PLADISTA% > 6 Then PLADISTA% = 6
    FORCANEC.Text1.TEXT = TRIM$(Str$(PLADISTA%))
    '
    CT$ = UCase$(CONTROL$("CANECOM", "LOST-TER"))
    LOSTN(3) = Val(CT$): If LOSTN(3) < 0 Then LOSTN(3) = 0
    LOST$(3) = Mid$(Str$(LOSTN(3)), 2)
    FORCANEC.Text2.TEXT = LOST$(3)
    '
    CT$ = UCase$(CONTROL$("CANECOM", "LOST-MON"))
    LOSTN(2) = Val(CT$): If LOSTN(2) < 0 Then LOSTN(2) = 0
    LOST$(2) = Mid$(Str$(LOSTN(2)), 2)
    FORCANEC.Text3.TEXT = LOST$(2)
    '
    CT$ = UCase$(CONTROL$("CANECOM", "LOST-MEC"))
    LOSTN(1) = Val(CT$): If LOSTN(1) < 0 Then LOSTN(1) = 0
    LOST$(1) = Mid$(Str$(LOSTN(1)), 2)
    FORCANEC.Text4.TEXT = LOST$(1)
    '
    CT$ = UCase$(CONTROL$("CANECOM", "LOST-COM"))
    LOSTN(0) = Val(CT$): If LOSTN(0) < 0 Then LOSTN(0) = 0
    LOST$(0) = Mid$(Str$(LOSTN(0)), 2)
    FORCANEC.Text5.TEXT = LOST$(0)
    '
    CT$ = UCase$(CONTROL$("CANECOM", "REPOMIN"))
    REPOMIN = Val(CT$): If REPOMIN < 1 Then REPOMIN = 1
    REPOMINS$ = Mid$(Str$(REPOMIN), 2)
    '
    FORCANEC.Check1.Value = 0
    FORCANEC.Check2.Value = 0
    FORCANEC.Check3.Value = 0
    FORCANEC.Check4.Value = 0
    FORCANEC.Check5.Value = 0
    FORCANEC.Check6.Value = 0
    FORCANEC.Check7.Value = 0
    FORCANEC.Label1 = ""
    '
    CTX% = 0
    '
End Sub
'

Sub GRAPARCAL()
    '
    SEMANAL% = 0: PERIO$ = "MENSUAL"
    If FORCANEC.Option4.Value = True Then
        SEMANAL% = 1: PERIO$ = "SEMANAL"
    End If
    Call GRACONTROL("CANECOM", "PERIODO", PERIO$)
    '
    PLADISTA% = Val(FORCANEC.Text1.TEXT)
    Call GRACONTROL("CANECOM", "PLADISTA", TRIM$(Str$(PLADISTA%)))
    '
    LOST$(3) = TRIM$(FORCANEC.Text2.TEXT): LOSTN(3) = Val(LOST$(3))
    LOST$(2) = TRIM$(FORCANEC.Text3.TEXT): LOSTN(2) = Val(LOST$(2))
    LOST$(1) = TRIM$(FORCANEC.Text4.TEXT): LOSTN(1) = Val(LOST$(1))
    LOST$(0) = TRIM$(FORCANEC.Text5.TEXT): LOSTN(0) = Val(LOST$(0))
    REPOMINS$ = "1": REPOMIN = 1
    '
    Call GRACONTROL("CANECOM", "LOST-TER", LOST$(3))
    Call GRACONTROL("CANECOM", "LOST-MON", LOST$(2))
    Call GRACONTROL("CANECOM", "LOST-MEC", LOST$(1))
    Call GRACONTROL("CANECOM", "LOST-COM", LOST$(0))
    Call GRACONTROL("CANECOM", "REPOMIN", REPOMINS$)
    '
End Sub
'
Sub SETCALCU()
    '
    Screen.MousePointer = 11
    '
    HOO% = HOY(HO$)
    '
    LOREC% = 72  ' *************  caso programacion mensual
    CACOL% = 18  '                cantidad de periodos
    COSEMACT% = HOO%            ' comienzo semana actual
    '
    J% = 1
    IPERIO$ = "01" + Mid$(HO$, 3, 6)
    INIPER%(1) = FECHANUM(IPERIO$)
    DIBAS = INIPER%(1)
        '
        While J% < CACOL%
95         DIBAS = DIBAS + 1
           FF = DIBAS: DIA = Val(FECHATEX$(FF))
                       If DIA <> 1 Then GoTo 95
           FF = DIBAS: MES = Val(Mid$(FECHATEX$(FF), 4, 2))
                       If MES < 1 Then GoTo 95
                       If MES > 12 Then GoTo 95
           J% = J% + 1
           INIPER%(J%) = DIBAS
        Wend
        INIPER%(1) = HOO%
        '
    If SEMANAL% = 1 Then  ' ***** caso programacion semanal
        '
        LOREC% = 208
        CACOL% = 52
        '
        DMX%(1) = 31
        DMX%(2) = 28
        DMX%(3) = 31
        DMX%(4) = 30
        DMX%(5) = 31
        DMX%(6) = 30
        DMX%(7) = 31
        DMX%(8) = 31
        DMX%(9) = 30
        DMX%(10) = 31
        DMX%(11) = 30
        DMX%(12) = 31
        '
        FFX = HOO%
        SEMANACT% = SEMANU%(FFX, A, B)
        COSEMACT% = A
        '
        FFX = HOO%
        SEMAHOO% = SEMANU%(FFX, A, B)
        DIBAS = A
        SMNA% = SEMAHO%
        J% = 1
        INIPER%(1) = DIBAS
'FEX = INIPER%(1): FFSS$ = FECHATEX$(FEX)
        While J% < CACOL%
98         DIBAS = DIBAS + 1
           FF = DIBAS: DIAI% = Val(Left$(FECHATEX$(FF), 2))
                       If DIAI% < 1 Then GoTo 98
           FF = DIBAS: MESI% = Val(Mid$(FECHATEX$(FF), 4, 2))
                       If MESI% < 1 Then GoTo 98
                       If MESI% > 12 Then GoTo 98
99         MESX% = MESI%
           DIAMES% = DMX%(MESX%)
           If DIAMES% = 28 Then
               If (Val(Mid$(FECHATEX$(FF), 7, 2)) Mod 4) = 0 Then
                   DIAMES% = 29
               End If
           End If
           '
           If DIAI% > DIAMES% Then GoTo 98

           FF = DIBAS
           SMN% = SEMANU%(FF, A, B)
           If SMN% = SMNA% Then GoTo 98
           '
           FF = DIBAS
           FEX = FF
           If DIASEM%(FF) <> 1 Then GoTo 98

           SMNA% = SMN%
           J% = J% + 1: INIPER%(J%) = DIBAS
'FEX = INIPER%(J%): FFSS$ = FECHATEX$(FEX)
           DIBAS = DIBAS + 6
        Wend
    End If
    '
    Screen.MousePointer = 1
    '
End Sub
'

Sub CANECOM()
     '
     Dim BAGRE() As String
     Dim HAYPLAFAB%(32767)
     '
     UDIM& = 2048
     ReDim Preserve BAGRE$(UDIM&)
'CICONT% = CODINT%("19019")
     LU$ = LUDAT$
     HOO% = HOY(HO$)
     CACO% = 0
     Screen.MousePointer = 11
     '
100  ' ************************* ABRIENDO ARCHIVOS MAESTROS
     '
     FORCANEC.Check1.Value = 2
     Screen.MousePointer = 11
     NMS& = NUMAS%
     NVI% = NUINV%
     'N1% = FILENBR%("MASTER")
     N1% = FILEOPEN%(LUDAT$ + "MASTER.DAT", 0, 128)
     'N9% = FILENBR%("SALSTOCK")
     N9% = FILEOPEN%(LUDAT$ + "SALSTOCK.DAT", 0, 40)
     '
     GESUBOR% = 0
     If CONTROL$("", "GESUBOR") = "SI" Then
         GESUBOR% = 1
     End If
     '
     GoSub 5300                           ' estrunue - estrudos
     FORCANEC.Check1.Value = 1
     '
     If NESTRU& < 1 Then
         Call MENSERR(24, "Proceso Imposible.\ \Faltan Definir Fórmulas y\Estructuras de Producto.")
         Call FINAL("")
     End If
     '
     PLADISTA% = Val(FORCANEC.Text1.TEXT)
     If PLADISTA% > 6 Then PLADISTA% = 6
     '
'    **** DEPURA PROGRAMA DE ENTREGAS POR ENTREGAS REALIZADAS
     '
     APROGRA$ = "PROGENT"
     If MODOPLAN% = 1 Then APROGRA$ = "PROGFAB"
     '
     UREPRO& = ULTREG&(APROGRA$)
     JJ& = 0
     For I& = 1 To UREPRO&
       ZZ$ = REGLEIDO$(APROGRA$, I&)
       If CVS(Mid$(ZZ$, 7, 4)) > 0.05 Then
         Call REPLA(ZZ$, MKS$(0), 11, 4)
         Call REPLA(ZZ$, MKS$(0), 19, 4)
         JJ& = JJ& + 1
         Call GRAREG(APROGRA$, ZZ$, JJ&)
       End If
       If MODOPLAN% = 1 Then
         ' marca items programados para no considerar stock como disponibilidad
         CIXI% = CVI(Mid$(ZZ$, 1, 2))
         If CIXI% > 0 Then
           If CIXI% <= NUMAS% Then
             CAPROBRU = CVS(Mid$(ZZ$, 7, 4))
             If CAPROBRU > 0.05 Then
               HAYPLAFAB%(CIXI%) = 1
             End If
           End If
         End If
         '
       End If
     Next I&
     Call SETULTREG(APROGRA$, JJ&)
     Call CIERRARCH(APROGRA$)
     If JJ& < 1 Then
        On Error Resume Next
        Kill LUDAT$ + APROGRA$ + ".DAT"
        On Error GoTo 0
     End If
     '
     CMA% = COMESACT%
     If MODOPLAN% <> 1 Then
       FIN& = ULTREG&("MOVISTO")
       'FINUMO% = FILENBR%("MOVISTO")
       FINUMO% = FILEOPEN%(LUDAT$ + "MOVISTO.DAT", 0, 128)
       Dim RMVS As String * 128
       For I& = 1 To FIN&
         Call TRACE(20, I&, FIN&)
         Get #FINUMO%, I&, RMVS$
         XS$ = RMVS$
         FF% = CVI(Left$(XS$, 2))
         If FF% >= CMA% Then
           COMO$ = UCase$(Mid$(XS$, 33, 2))
           If COMO$ = "FC" Or COMO$ = "NC" Then
             CIXI% = CVI(Mid$(XS$, 3, 2))
             If CIXI% > 0 Then
               If CIXI% <= NUMAS% Then
                 CAENTREG = CVD(Mid$(XS$, 96, 8)) - CVD(Mid$(XS$, 88, 8))
                 '
                 RFF$ = RECFILT$(APROGRA$, 1, MKI$(CIXI%))
                 For KK1& = 1 To Len(RFF$) Step 4
                   REPRO& = CVS(Mid$(RFF$, KK1&, 4))
                   ZZ$ = REGLEIDO$(APROGRA$, REPRO&)
                   ACUENTRE = CVS(Mid$(ZZ$, 11, 4))
                   ENTREMA = CVS(Mid$(ZZ$, 7, 4)) - ACUENTRE
                   If CAENTREG <= ENTREMA Then
                      Call REPLA(ZZ$, MKS$(CAENTREG + ACUENTRE), 11, 4)
                      CAENTREG = 0
                    Else
                      Call REPLA(ZZ$, MKS$(ENTREMA + ACUENTRE), 11, 4)
                      CAENTREG = CAENTREG - ENTREMA
                   End If
                   Call GRAREG(APROGRA$, ZZ$, REPRO&)
                   If CAENTREG < 0.05 Then GoTo 102
                 Next KK1&
               End If
             End If
           End If
         End If
102    Next I&
       Call CIERRARCH(APROGRA$)
     End If
     '
'    **** AGREGA PEDIDOS PENDIENTES DE ENTREGA
'    **** DESCARTANDO L0S CORRESPONDIENTES A REPARACIONES
     '
    If MODOPLAN% <> 1 Then
      APEDPE$ = "PEDDETA"
      FIN& = ULTREG&(APEDPE$)
      '
      For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$(APEDPE$, U&)
        If CVI(Mid$(X$, 7, 2)) > 0 Then
          If UCase$(Mid$(X$, 51, 1)) <> "R" Then ' no es reparacion
            CIXI% = CVI(Mid$(X$, 33, 2))
            CAPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
            FENTRESOL% = CVI(Mid$(X$, 25, 2))
            If Asc(Mid$(X$, 27, 1)) > 1 Then CAPEN = 0 ' Cumplido o Anulado
            '
            If CAPEN > 0 Then
              RFF$ = RECFILT$(APROGRA$, 1, MKI$(CIXI%))
              '
              If Len(RFF$) < 4 Then     ' NO HAY PROGRAMA
                FENTRESO% = FENTRESOL%
                If FENTRESO% < HOO% Then FENTRESO% = HOO%
                ZZ$ = REGBLAN$(APROGRA$)
                Call REPLA(ZZ$, MKI$(CIXI%), 1, 2)
                Call REPLA(ZZ$, MKI$(FENTRESO%), 3, 2)
                Call REPLA(ZZ$, MKS$(CAPEN), 19, 4)
                Call REGAPP(APROGRA$, ZZ$)
                Call CIERRARCH(APROGRA$)
                GoTo 103
              End If
              '
              For KK1& = Len(RFF$) - 3 To 1 Step -4
                REPRO& = CVS(Mid$(RFF$, KK1&, 4))
                ZZ$ = REGLEIDO$(APROGRA$, REPRO&)
                FEPROG% = CVI(Mid$(ZZ$, 3, 2))
                If FEPROG% <= FENTRESOL% Or KK1& = 1 Then
                  CAPEPROMAX = CVS(Mid$(ZZ$, 7, 4)) - CVS(Mid$(ZZ$, 11, 4)) - CVS(Mid$(ZZ$, 19, 4))
                  '         PROGRAMA BRUTO - ENTREGAS - PEDIDOS PROGRAMADOS
                  If CAPEPROMAX >= CAPEN Then
                      CAPEDID = CVS(Mid$(ZZ$, 19, 4)) + CAPEN
                      Call REPLA(ZZ$, MKS$(CAPEDID), 19, 4)
                      Call GRAREG(APROGRA$, ZZ$, REPRO&)
                      CAPEN = 0
                      GoTo 103
                    Else
                      CAPEDID = CVS(Mid$(ZZ$, 19, 4)) + CAPEPROMAX
                      Call REPLA(ZZ$, MKS$(CAPEDID), 19, 4)
                      Call GRAREG(APROGRA$, ZZ$, REPRO&)
                      CAPEN = CAPEN - CAPEPROMAX
                  End If
                End If
              Next KK1&
              '
              If CAPEN > 0.05 Then
                FENTRESO% = FENTRESOL%
                If FENTRESO% < HOO% Then FENTRESO% = HOO%
                ZZ$ = REGBLAN$(APROGRA$)
                Call REPLA(ZZ$, MKI$(CIXI%), 1, 2)
                Call REPLA(ZZ$, MKI$(FENTRESO%), 3, 2)
                Call REPLA(ZZ$, MKS$(CAPEN), 19, 4)
                Call REGAPP(APROGRA$, ZZ$)
                Call CIERRARCH(APROGRA$)
              End If
              '
            End If
          End If
        End If
        '
103   Next U&
    End If
    '
    Call CIERRARCH(APROGRA$)
' hasta aqui esta corregido para entregas / fabricacion
' Y PARA DESCARTAR PEDIDOS POR REPARACIONES
'
' ************** lee existencias, stock minimo y lote reposicion *******************

     FORCANEC.Check2.Value = 2
     Screen.MousePointer = 11
     '
     FIN& = NESTRU&
     For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        RA$ = REGLEIDO$("ESTRUNUE", I&)
        CI% = CVI(Left$(RA$, 2))
        If CI% > 0 Then
            If CI% <= NMS& Then
                R2$ = Mid$(RA$, 3, 120)
                For K% = 1 To 120 Step 8
                    DPSX$ = Mid$(R2$, K%, 8)
                    If CVI(DPSX$) < 1 Or CVI(DPSX$) > NMS& Then GoTo 105
                    If CVS(Mid$(DPSX$, 3)) >= 0.00005 Then DPI%(CI%) = DPI%(CI%) + 1
                    If DPI%(CI%) > 1 Then GoTo 105
                Next K%
            End If
        End If
105  Next I&
     '
     HAYNE% = 0
     MODO% = 0
     If FORCANEC.Option1.Value = True Then MODO% = 1
     If FORCANEC.Option2.Value = True Then MODO% = 2
     '
     FFXY = HOY(HO$)
     SMXX% = SEMANU%(FFXY, AA, BB)
     COSEMACT% = AA
     '
     QLN& = 0
     Call CIERRARCH("COBERFA")
     'Call BLANARCH("COBERFA")
     Call CIERRARCH("ORFAEMI")
     'Call BLANARCH("ORFAEMI")
     Call CIERRARCH("OCOMEMI")
     'Call BLANARCH("OCOMEMI")
     '
     For CI% = 1 To NMS&
       CIIX& = CI%
       Call TRACE(20, CIIX&, NMS&)
       Get #N9%, CI% + 1, SST$
       SALSTO = 0
       If STATUIT%(CI%) >= 0 Then
         'For K% = 1 To 40 Step 4
         SALSTO = STOCKACT(CI%)  'SALSTO + CVS(Mid$(SST$, K%, 4))
         'Next K%
         '
'CIJI% = CI%
'Call VESARRAS(CIJI%)
'If STOCKACT(CIJI%) <> SALSTO Then
'ABC$ = CODEXT$(CIJI%)
'If Left$(ABC$, 2) = "42" Then
'AAA = BBB
'End If
'End If

         If MODO% = 1 Then SALSTO = 0
         If MODO% = 2 Then If SALSTO < 0 Then SALSTO = 0

         If SALSTO < 0 Then
            HAYNE% = HAYNE% + 1
            CI1% = CI%
            TTXX$ = REGBLAN$("LINEGAT")
            Call REPLA(TTXX$, MKI$(CI%), 1, 2)
            Call REPLA(TTXX$, DESCRIT$(CI%), 3, 30)
            Call REPLA(TTXX$, CSTRING$(MKS$(SALSTO), 4, 10, 1, 2), 33, 10)
            Call REPLA(TTXX$, UNIMED$(CI%), 45, 4)
            '
            QLN& = QLN& + 1
            Call GRAREG("LINEGAT", TTXX$, QLN&)
            '
         End If
         '
         STMIN = STOMIN(CI%): If MODO% = 1 Then STMIN = 0
         REPOS = LOTESTAN(CI%)
         DP% = DPI%(CI%): If DP% < 0 Or DP% > 3 Then DP% = 0

         If REPOS < 0.5 Then REPOS = LOSTN(DP%)
         If REPOS < 0.5 Then REPOS = REPOMIN

         DISPACT(CI%) = SALSTO - STMIN
         If HAYPLAFAB%(CI%) > 0 Then DISPACT(CI%) = 0
         LOREPOS(CI%) = REPOS
         If HAYPLAFAB%(CI%) < 1 Then
           If Abs(SALSTO) > 0.05 Then
'If CI% = CICONT% Then
'FEX = COSEMACT%: ABC$ = FECHATEX$(FEX)
'AAA = BBB
'End If
             Call COBERFA(COSEMACT%, CI%, "Stock al " + HO$, 10, SALSTO)
           End If
           '
           If Abs(STMIN) > 0.05 Then
'If CI% = CICONT% Then
'FEX = COSEMACT%: ABC$ = FECHATEX$(FEX)
'AAA = BBB
'End If
             Call COBERFA(COSEMACT%, CI%, "Stock Minimo del Item", 20, (-1) * STMIN)
           End If
         End If
       End If
     Next CI%
     '
     Call SETULTREG("LINEGAT", QLN&)
     Call CIERRARCH("LINEGAT")
     '
     If HAYNE% > 0 Then
        Call FILESORT("LINEGAT", "", 1, 1, "ASC")
        A% = COMALTER%("ATENCION: Se Detectaron Existencias Negativas\          en" + Str$(HAYNE%) + " Items de Stock.\\Revisar\Cancelar\Continuar")
        On A% GoTo 107, 106, 108
106     Call CIERRARCH("*.*")
        Screen.MousePointer = 1
        Exit Sub
107     Call FRESHECHO("LINEGAT")
        Call SELECHO("LINEGAT")
        VDEVU$ = VALACT1$("LINEGAT")
        If UCase$(Left$(App.EXEName, 5)) = "FLEXO" Then
          If VDEVU$ <> "" Then
            CONSTO00.Text1 = VDEVU$
            CONSTO00.Show 1
            GoTo 107
          End If
        End If
        '
        If ALTERNA%("Continuar Proceso") <> 1 Then
            Call CIERRARCH("*.*")
            Screen.MousePointer = 1
            Exit Sub
        End If
     End If
     '
108  FORCANEC.Check2.Value = 1
     FORCANEC.Refresh
     FORCANEC.Refresh
     '
     Screen.MousePointer = 11
     FINL& = ULTREG&("OCOMPRA")
     For IL& = 1 To FINL&
        '
        Call TRACE(20, IL&, FINL&)
        F$ = REGLEIDO$("OCOMPRA", IL&)
        '
        STATU% = Asc(Mid$(F$, 62, 1))
        If STATU% <= 0 Then
          CATOT# = CVS(Mid$(F$, 13, 4))
          CAPRA# = CVS(Mid$(F$, 58, 4))
          CAPEN = CATOT# - CAPRA#
          If CAPEN > 0 Then
            CIBX% = CVI(Mid$(F$, 11, 2))
            If STATUIT%(CIBX%) >= 0 Then
              NC% = CVI(Mid$(F$, 7, 2))
              COMPROB$ = "OC." + TRIM$(Str$(CVS(Left$(F$, 4))))
              FENTRE% = CVI(Mid$(F$, 9, 2))
              CICX% = CIBX%
'If CICX% = CICONT% Then
'AAA = BBB
'End If
              Call COBERFA(FENTRE%, CIBX%, COMPROB$ + " " + RASOCLI$((-1) * NC%), 13, CAPEN)
              If HAYPLAFAB%(CICX%) < 1 Then
                DISPACT(CICX%) = DISPACT(CICX%) + CAPEN
              End If
            End If
          End If
        End If
     Next IL&
     '
     FINL& = ULTREG&("OCOMDETA")
     For IL& = 1 To FINL&
        '
        Call TRACE(20, IL&, FINL&)
        F$ = REGLEIDO$("OCOMDETA", IL&)
        CIBX% = CVI(Mid$(F$, 9, 2))
        STATU% = Asc(Mid$(F$, 124, 1))
        If STATU% < 9 Then      ' no esta anulada
          NUORSE& = CVS(Mid$(F$, 117, 4))
          If NUORSE& < 1 Then
            COEFCAN = CVS(Mid$(F$, 63, 4))
            If COEFCAN < 0.005 Then COEFCAN = 1
            CATOT# = CVS(Mid$(F$, 55, 4)) * COEFCAN
            CAPRA# = CVS(Mid$(F$, 99, 4)) * COEFCAN
            CAPEN = CATOT# - CAPRA#
            If CAPEN > 0 Then
              If STATUIT%(CIBX%) >= 0 Then
                NC% = CVI(Mid$(F$, 7, 2))
                NROPED& = CVS(Left$(F$, 4))
                FENTRE% = CVI(Mid$(F$, 95, 2))
                COMPROB$ = "OC." + TRIM$(Str$(NROPED&))
                CICX% = CIBX%
'If CICX% = CICONT% Then
'AAA = BBB
'End If
                Call COBERFA(FENTRE%, CIBX%, COMPROB$ + " " + RASOCLI$((-1) * NC%), 13, CAPEN)
                If HAYPLAFAB%(CICX%) < 1 Then
                  DISPACT(CICX%) = DISPACT(CICX%) + CAPEN
                End If
              End If
            End If
          End If
        End If
     Next IL&
     '
' ******** agrega ordenes de fabricacion lanzadas y en curso ******
     '
     JJJ& = 0
     MODOCA% = 1
     If Left$(UCase$(EMPREAC$), 4) = "NOVO" Then
        MODOCA% = 2
     End If
     '
     FINL& = ULTREG&("ORDEFABR")
     For IL& = 1 To FINL&
       '
       Call TRACE(20, IL&, FINL&)
       F$ = REGLEIDO$("ORDEFABR", IL&)
       '
       STATU% = Asc(Mid$(F$, 108, 1))
       If STATU% <= 2 Then ' esta proyectada o lanzada
         '
         CI% = CVI(Mid$(F$, 83, 2))
         If STATUIT%(CI%) >= 0 Then
           If MODOPLAN% <> 1 Or HAYPLAFAB%(CI%) < 1 Then 'descarta ordenes de fabricación productos terminados
             CAPEN = CVD(Mid$(F$, 89, 8)) - CVD(Mid$(F$, 119, 8))
             NORFA$ = TRIM$(Mid$(F$, 69, 12))
             REFOR$ = Mid$(F$, 5, 64)
             NUORFA& = Val(Mid$(NORFA$, 4, 4))
             SUBOR% = Val(Mid$(NORFA$, 9, 2))
             FENTRE% = CVI(Mid$(F$, 147, 2))
             If FENTRE% < HOO% Then FENTRE% = CVI(Mid$(F$, 105, 2))
             If FENTRE% < HOO% Then FENTRE% = HOO%
             FELANZ% = CVI(Mid$(F$, 109, 2))
             If FELANZ% < HOO% Then FELANZ% = HOO%
             STAT$ = " (proyectada)"
             If STATU% = 1 Then STAT$ = " (lanzada)"
             If STATU% = 2 Then STAT$ = " (suspendida)"
             '
             If CAPEN > 0 Then
               If CI% > 0 Then
                 If CI% <= NUMAS% Then
                   DISPACT(CI%) = DISPACT(CI%) + CAPEN
                   '
                   CIBX% = CI%
                   COD$ = CODEXT$(CI%)
                   'Call LEEEXPLOMUL(COD$, CAPEN, 1)
                   ' SE REEMPLAZA POR BARRIDO ARCHIVO DE RESERVAS
                   COEF = 1
                   '
                   If MODOCA% = 2 Then ' CASO NOVOESPUMA
                     QQ% = 0
                     For III& = 1 To ULTREG&("PEDDETP")
                       Y$ = REGLEIDO$("PEDDETP", III&)
                       If CVS(Left$(Y$, 4)) = NUORFA& Then
                         QQ% = QQ% + 1
                         If QQ% = SUBOR% Then
                           If CVI(Mid$(Y$, 33, 2)) = CIBX% Then
                             If CVS(Mid$(Y$, 59, 4)) > 1 Then
                               If CVS(Mid$(Y$, 63, 4)) > 1 Then
                                 If CVS(Mid$(Y$, 67, 4)) > 1 Then
                                   COEF = (CVS(Mid$(Y$, 59, 4))) / 1000
                                   COEF = COEF * (CVS(Mid$(Y$, 63, 4))) / 1000
                                   COEF = COEF * (CVS(Mid$(Y$, 67, 4))) / 1000
                                 End If
                               End If
                             End If
                             GoTo 109
                           End If
                         End If
                       End If
                     Next III&
                     '
                     ' segunda vuelta valida solo codigo
                     For III& = 1 To ULTREG&("PEDDETP")
                       Y$ = REGLEIDO$("PEDDETP", III&)
                       If CVS(Left$(Y$, 4)) = NUORFA& Then
                         If CVI(Mid$(Y$, 33, 2)) = CIBX% Then
                           If CVS(Mid$(Y$, 59, 4)) > 1 Then
                             If CVS(Mid$(Y$, 63, 4)) > 1 Then
                               If CVS(Mid$(Y$, 67, 4)) > 1 Then
                                 COEF = (CVS(Mid$(Y$, 59, 4))) / 1000
                                 COEF = COEF * (CVS(Mid$(Y$, 63, 4))) / 1000
                                 COEF = COEF * (CVS(Mid$(Y$, 67, 4))) / 1000
                               End If
                             End If
                           End If
                           GoTo 109
                         End If
                       End If
                     Next III&
                     '
                     JJJ& = JJJ& + 1
                     ZJJ$ = REGBLAN$("OFNOPED")
                     Call REPLA(ZJJ$, NORFA$, 1, 12)
                     Call REPLA(ZJJ$, REFOR$, 13, 64)
                     Call GRAREG("OFNOPED", ZJJ$, JJJ&)
                     GoTo 1091
                     '
                   End If
                   '
109 '               '
'If CI% = CICONT% Then
'AAA = BBB
'End If
                   Call COBERFA(FENTRE%, CIBX%, NORFA$ + STAT$, 10, CAPEN)
                   '
                   ' SE REEMPLAZA POR BARRIDO ARCHIVO DE RESERVAS
                   'For K% = 1 To CAIT%
                   '    CI1% = CIJ%(K%)
                   '    If CI1% > 0 Then
                   '       If CI1% <= NUMAS% Then
                   '          CAN1 = CAPEN * USOI(K%) * COEF
                   '          DISPACT(CI1%) = DISPACT(CI1%) - CAN1
                   '          Call COBERFA(FELANZ%, CI1%, "Reserva para " + NORFA$, 20, (-1) * CAN1)
                   '        End If
                   '    End If
                   'Next K%
                 End If
               End If
             End If
           End If
         End If
       End If
1091 Next IL&
     '
     If JJJ& > 0 Then
        Call SETULTREG("OFNOPED", JJJ&)
        Call CIERRARCH("OFNOPED")
        TTXX$ = "Algunas Ordenes de Fabricación en Curso"
        TTXX$ = TTXX$ + "\No Pudieron Procesarse por no Estar Activo"
        TTXX$ = TTXX$ + "\el Pedido Correspondiente."
        TTXX$ = TTXX$ + "\\Ver Detalle\Listado\Continuar"
        On COMALTER%(TTXX$) GoTo 1092, 1093
        GoTo 110
        '
1092    Call FINAL("*COFNOPE")
        If COMALTER%("Puede Ahora Continuar o Cancelar Proceso\\Continuar\Cancelar Proceso") <> 1 Then
            Call FINAL("")
        End If
        GoTo 110
        '
1093    Call FINAL("*LOFNOPE")
        If COMALTER%("Puede Ahora Continuar o Cancelar Proceso\\Continuar\Cancelar Proceso") <> 1 Then
            Call FINAL("")
        End If
        GoTo 110
        '
     End If
     '
     Screen.MousePointer = 11
     FINL& = ULTREG&("RESERVA")
     For IL& = 1 To FINL&
       Call TRACE(20, IL&, FINL&)
       F$ = REGLEIDO$("RESERVA", IL&)
       CI1% = CVI(Left$(F$, 2))
       CAN1 = CVD(Mid$(F$, 19, 8)) - CVD(Mid$(F$, 27, 8))
       If CAN1 >= 0.05 Then
'If CI1% = CICONT% Then
'AAA = BBB
'End If
          DISPACT(CI1%) = DISPACT(CI1%) - CAN1
          FELANZ% = CVI(Mid$(F$, 17, 2))
          NORFA$ = TRIM$(Mid$(F$, 3, 12))
'If CI1% = CICONT% Then
'AAA = BBB
'End If
          Call COBERFA(FELANZ%, CI1%, "Reserva para " + NORFA$, 20, (-1) * CAN1)
       End If
     Next IL&
     Call CIERRARCH("RESERVA")
     '
     HOII% = HOY(HOS$)
     FINL& = ULTREG&("REPUREPA")
     For IL& = 1 To FINL&
       Call TRACE(20, IL&, FINL&)
       F$ = REGLEIDO$("REPUREPA", IL&)
       CI1% = CVI(Mid$(F$, 7, 2))
       CAN1 = CVS(Mid$(F$, 21, 4)) - CVS(Mid$(F$, 25, 4))
       If CAN1 >= 0.05 Then
'If CI1% = CICONT% Then
'AAA = BBB
'End If
          DISPACT(CI1%) = DISPACT(CI1%) - CAN1
          FELANZ% = HOII%
          NORFA$ = TRIM$(Mid$(F$, 33, 12))
          NOREPARA& = CVS(Left$(F$, 4))
          REREP& = REGBUS&("PRESREPA", 1, MKS$(NOREPARA&))
          If REREP& > 0 Then
            YY$ = REGLEIDO$("PRESREPA", REREP&)
            If CVI(Mid$(YY$, 123, 2)) > 0 Then GoTo 1095 ' O/Rep Cerrada
            If CVS(Mid$(YY$, 125, 4)) < 1 Then GoTo 1095 ' falta O/Venta
            If TRIM$(NORFA$) = "" Then
              NORFA$ = "OR-" + TRIM$(Str$(NOREPARA&))
            End If
'If CI1% = CICONT% Then
'AAA = BBB
'End If
            Call COBERFA(FELANZ%, CI1%, "Reserva p/Rep." + NORFA$, 20, (-1) * CAN1)
          End If
       End If
1095 Next IL&
     Call CIERRARCH("REPUREPA")
     '
' ******** inicializa necesidades secundarias de compra y fabricacion ********

110  FORCANEC.Check3.Value = 2
     Screen.MousePointer = 11
     '
111  N13% = FreeFile
     Open LU$ + "NEFACOM.DAT" For Random Access Read Write Shared As #N13% Len = LOREC%
     Dim PDS As String * 208
     Dim PDM As String * 72
     '
     If LOF(N13%) > 0 Then
         Close #N13%
         Kill LU$ + "NEFACOM.DAT"
         GoTo 111
     End If
     '
     For I& = 1 To NMS&
         CIIX& = I&
         Call TRACE(20, CIIX&, NMS&)
         If LOREC% = 208 Then
              Get #N13%, I&, PDS$
            Else
              Get #N13%, I&, PDM$
         End If
         PDS$ = String$(208, 0)
         PDM$ = String$(72, 0)
         If LOREC% = 208 Then
              Put #N13%, I&, PDS$
            Else
              Put #N13%, I&, PDM$
         End If
     Next I&

     FORCANEC.Check3.Value = 1


' **************** necesidades brutas segun entregas programadas **************

120  N9% = FreeFile
     Open LU$ + "NEBRUNE.DAT" For Random Access Read Write Shared As #N9% Len = 48
     Dim B00 As String * 48
     Dim B01 As String * 48
     X1& = 1: X2& = 0

     If LOF(N9%) > 0 Then
        Close #N9%
        Kill LU$ + "NEBRUNE.DAT"
        GoTo 120
     End If

     FORCANEC.Check4.Value = 2
     Screen.MousePointer = 11
     '
121  FIN& = NUMAS%
     For I& = 1 To FIN&
       CIIX& = I&
       Call TRACE(20, CIIX&, FIN&)
       CI% = I&
       If STATUIT%(CI%) >= 0 Then
         '
         If DISPACT(CI%) < 0 Then
'If CI% = CICONT% Then
'AAA = BBB
'End If
            CANTU = Abs(DISPACT(CI%))
            FF% = HOO%
            '
            X2& = X2& + 1
            B00$ = String$(48, 0)
            Call REPLA(B00$, MKI$(CI%), 1, 2)
            Call REPLA(B00$, MKS$(CANTU), 3, 4)
            Call REPLA(B00$, MKI$(FF%), 7, 2)
            Call REPLA(B00$, "PROGRAMA POR MINIMOS", 9, 29)
            Call REPLA(B00$, MKS$(0), 38, 4)
            Call REPLA(B00$, MKS$(0), 42, 4)
            Call REPLA(B00$, MKS$(0), 46, 2)
            Call REPLA(B00$, " ", 48, 1)
            '
            Put #N9%, X2&, B00$
            DISPACT(CI%) = 0
            '
         End If
       End If
122  Next I&
     '
     N4% = FreeFile
     Open LU$ + APROGRA$ + ".DAT" For Random Access Read Write Shared As #N4% Len = 24
     Dim P00 As String * 24
     FIN& = LOF(N4%) / 24
     For I& = 1 To FIN&
       CIIX& = I&
       Call TRACE(20, CIIX&, FIN&)
       Get #N4%, I&, P00$
       P1$ = Left$(P00$, 2)
       CI% = CVI(P1$)
       '
       If CI% > 0 Then
         If CI% <= NUMAS% Then
           If STATUIT%(CI%) >= 0 Then
             '
             P3$ = Mid$(P00$, 3, 2)
             P6$ = Mid$(P00$, 7, 4)
             P7$ = Mid$(P00$, 11, 4)
             '
             If CVI(P3$) < COMESACT% Then
                cix% = CVI(P1$)
                Call MENSERR(24, "Error: Codigo " + TRIM$(CODEXT$(cix%)) + " Mal Fecha Entrega Programada")
                GoTo 129
             End If
             '
             PLAVEN = CVS(P6$)
             ENTREG = CVS(P7$)
             PEDIPE = CVS(Mid$(P00$, 19, 4))
             CANTU = PLAVEN - ENTREG
             '
             PEDINOPRO = PEDIPE + ENTREG - PLAVEN
             If PEDINOPRO < 0 Then PEDINOPRO = 0
             '
             CI% = CVI(P1$)
             FF% = CVI(P3$)
             '
             If CANTU > 0 Then
               If CI% > 0 Then
                  If CI% <= NMS& Then
                     DPI%(CI%) = 3
                  End If
               End If
             End If
             '
             If CANTU > 0 Then
               '
               X2& = X2& + 1
               B00$ = String$(48, 0)
               Call REPLA(B00$, MKI$(CI%), 1, 2)
               Call REPLA(B00$, MKS$(CANTU), 3, 4)
               Call REPLA(B00$, MKI$(FF%), 7, 2)
               Call REPLA(B00$, "PROGRAMA ENTREGAS AL " + HO$, 9, 29)
               If MODOPLAN% = 1 Then
                 Call REPLA(B00$, "PLAN DE FABRICACION AL " + HO$, 9, 29)
               End If
               If FF% < INIPER%(1) Then
                   Call REPLA(B00$, MKI$(INIPER%(1)), 7, 2)
                   Call REPLA(B00$, "ATRASO DE ENTREGA AL " + HO$, 9, 29)
               End If
               Call REPLA(B00$, MKS$(0), 38, 4)
               Call REPLA(B00$, MKS$(0), 42, 4)
               Call REPLA(B00$, MKS$(0), 46, 2)
               Call REPLA(B00$, " ", 48, 1)
               '
               Put #N9%, X2&, B00$
               '
             End If
             '
             If MODOPLAN% = 0 Then
               If PEDINOPRO > 0 Then
                 '
                 X2& = X2& + 1
                 B00$ = String$(48, 0)
                 Call REPLA(B00$, MKI$(CI%), 1, 2)
                 Call REPLA(B00$, MKS$(PEDINOPRO), 3, 4)
                 Call REPLA(B00$, MKI$(FF%), 7, 2)
                 Call REPLA(B00$, "PEDIDOS NO PROGRAMADOS AL " + HO$, 9, 29)
                 Call REPLA(B00$, MKS$(0), 38, 4)
                 Call REPLA(B00$, MKS$(0), 42, 4)
                 Call REPLA(B00$, MKS$(0), 46, 2)
                 Call REPLA(B00$, " ", 48, 1)
                 '
                 Put #N9%, X2&, B00$
                 '
               End If
             End If
             '
           End If
         End If
       End If
129  Next I&
     '
'For U11& = 1 To X2&
'  Call TRACE(20, U11&, X2&)
'1295  Get #N9%, U11&, B00$
'  B001$ = B00$
'  FF11% = CVI(Mid$(B00$, 7, 2))
'  For U22& = U11& + 1 To X2&
'    Get #N9%, U22&, B00$
'    FF22% = CVI(Mid$(B00$, 7, 2))
'    If FF22% < FF11% Then
'      Put #N9%, U11&, B00$
'      B00$ = B001$
'      Put #N9%, U22&, B00$
'      GoTo 1295
'    End If
'  Next U22&
'Next U11&
     FORCANEC.Check4.Value = 1
     '
130  CICLO% = CICLO% + 1
     CICLOMA% = 9
     If InStr(UCase$(EMPREAC$), "YALTRE") > 0 Then CICLOMA% = 15
     '
     If CICLO% > CICLOMA% Then GoTo 150
     FORCANEC.Label1 = TRIM$(Str$(CICLO%))
     FORCANEC.Label1.Refresh
     '
' ******** CALCULANDO NECESIDADES NETAS SEGUN DISPONIBILIDAD
     '
     FORCANEC.Check5.Value = 2
     Screen.MousePointer = 11
     For I& = X1& To X2&
        '
        CIIX& = I&
        Call TRACE(20, CIIX&, X2&)
        '
        Get #N9%, I&, B00$
        B1$ = Left$(B00$, 2)
        B3$ = Mid$(B00$, 3, 4)
        B4$ = Mid$(B00$, 7, 2)
        '
        CI% = CVI(B1$)
        If CI% > 0 Then
          If CI% <= NMS& Then
            If STATUIT%(CI%) >= 0 Then
      'If CI% = 484 Then
      'ABC$ = CODEXT$(CI%)
      'CCC = CVS(B3$)
      'AAA = BBB
      'End If
'If CI% = CICONT% Then
'AAA = CVS(B3$)
'FEX = CVI(B4$): ABC$ = FECHATEX$(FEX)
'End If
                If DISPACT(CI%) >= CVS(B3$) Then
                    DISPACT(CI%) = DISPACT(CI%) - CVS(B3$)
                    Call REPLA(B00$, B3$, 38, 4)
                    Call REPLA(B00$, MKS$(0), 42, 4)
                    Put #N9%, I&, B00$
                    GoTo 139
                End If
                '
                CANEC = CVS(B3$) - DISPACT(CI%)
                FENEBRU% = CVI(B4$): CUANDO% = 1
                '
'FEX = FENEBRU: ABC$ = FECHATEX$(FEX)
                For UX% = 1 To CACOL%
                    If INIPER%(UX%) <= FENEBRU% Then CUANDO% = UX%
                Next UX%
                '
                If LOREC% = 208 Then
                      Get #N13%, CI%, PDS$
                      PDX$ = PDS$
                    Else
                      Get #N13%, CI%, PDM$
                      PDX$ = PDM$
                End If
                '
                For UX% = CUANDO% To 1 Step -1
                    PROANT = CVS(Mid$(PDX$, 4 * UX% - 3, 4))
                    If PROANT > 0 Then GoTo 135
                Next UX%
                UX% = CUANDO%
135             If PROANT >= LOREPOS(CI%) Then UX% = CUANDO%
                '
                PROANT = CVS(Mid$(PDX$, 4 * UX% - 3, 4))
                Mid$(PDX$, 4 * UX% - 3, 4) = MKS$(PROANT + CANEC)
                '
                If LOREC% = 208 Then
                      PDS$ = PDX$
                      Put #N13%, CI%, PDS$
                    Else
                      PDM$ = PDX$
                      Put #N13%, CI%, PDM$
                End If
                '
                FENENE% = INIPER%(UX%)
                '
                Call REPLA(B00$, MKS$(DISPACT(CI%)), 38, 4)
                DISPACT(CI%) = 0
                Call REPLA(B00$, MKS$(CANEC), 42, 4)
                Call REPLA(B00$, MKI$(FENENE%), 46, 2)
                Put #N9%, I&, B00$
                '
            End If
          End If
        End If
139 Next I&



140 ' Call MENSAJE(14, "GENERANDO NECESIDADES SECUNDARIAS SEGUN EXPLOSION")

    X3& = X2&
    IAGRE& = 0
    For I& = X1& To X2&
        CIIX& = I&
        UKI& = I& - X1& + 1
        Call TRACE(20, UKI&, X2& - X1& + 1)
        Get #N9%, I&, B00$
        Call REPLA(B00$, "C", 48, 1)
        Put #N9%, I&, B00$
        '
        B1$ = Left$(B00$, 2)
        B3$ = Mid$(B00$, 3, 4)
        B4$ = Mid$(B00$, 7, 2)
        B7$ = Mid$(B00$, 42, 4)
        B8$ = Mid$(B00$, 46, 2)
        '
        CI% = CVI(B1$)
        If STATUIT%(CI%) >= 0 Then
          CABAS = CVS(B7$)
          FENEC% = CVI(B8$): CUANEC% = 1
                          For UX% = 1 To CACOL%
                             If INIPER%(UX%) <= FENEC% Then CUANEC% = UX%
                          Next UX%
          If CABAS > 0 Then
1234        CIBAS% = CI%
            CII% = CIBAS%
            CODBAS$ = TRIM$(CODEXT$(CII%))
            Call LEEEXPLOMUL(CODBAS$, CABAS, 1)
            If CAIT% > 0 Then
                Call REPLA(B00$, "F", 48, 1)
                Put #N9%, I&, B00$
                For KJ% = 1 To CAIT%    ' J%
                  CI% = CIJ%(KJ%)     '  CII1%(KJ%)
                  If STATUIT%(CI%) >= 0 Then
                    CANTU = CABAS * USOI(KJ%) 'CANT1(KJ%)
                    PLADISPO% = PDI%(KJ%)
                    If PLADISPO% < 1 Then PLADISPO% = PLADISTA%
                    CLX% = CUANEC% - PLADISPO%  'PAD1%(KJ%)
                    If CLX% < 1 Then CLX% = 1
                    If CANTU > 0 Then
                        'X3& = X3& + 1
                        B00$ = String$(48, 0)
                        Call REPLA(B00$, MKI$(CIJ%(KJ%)), 1, 2)
                        Call REPLA(B00$, MKS$(CANTU), 3, 4)
                        Call REPLA(B00$, MKI$(INIPER%(CLX%)), 7, 2)
                        Call REPLA(B00$, "Reserva para " + CODBAS$, 9, 28)
                        Call REPLA(B00$, " ", 48, 1)
                        IAGRE& = IAGRE& + 1
                        If IAGRE& >= UDIM& Then
                          UDIM& = UDIM& + 1024
                          ReDim Preserve BAGRE$(UDIM&)
                        End If
                        For JAGRE& = IAGRE& - 1 To 1 Step -1
                          If CVI(Mid$(BAGRE$(JAGRE&), 7, 2)) <= INIPER%(CLX%) Then
                            BAGRE$(JAGRE& + 1) = B00$
                            GoTo 1235
                          End If
                          BAGRE$(JAGRE& + 1) = BAGRE$(JAGRE&)
                        Next JAGRE&
                        BAGRE$(1) = B00$
                    End If
                  End If
1235            Next KJ%
            End If
          End If
        End If
    Next I&
    '
    For JAGRE& = 1 To IAGRE&
      X3& = X3& + 1
      B00$ = BAGRE$(JAGRE&)
      Put #N9%, X3&, B00$
    Next JAGRE&
    '
    X1& = X2& + 1
    X2& = X3&: If X2& < X1& Then GoTo 150 ' terminar
    GoTo 130

150 FORCANEC.Check5.Value = 1
    FORCANEC.Check6.Value = 2
    Screen.MousePointer = 11
    '
    For I& = 1 To NMS&
        '
        CIIX& = I&
        Call TRACE(20, CIIX&, NMS&)
        If LOREC% = 208 Then
              Get #N13%, I&, PDS$
              PDX$ = PDS$
            Else
              Get #N13%, I&, PDM$
              PDX$ = PDM$
        End If
        '
        For J% = 1 To CACOL%
            CAFAB = CVS(Mid$(PDX$, 4 * J% - 3, 4))
            If CAFAB > 0 Then
                REFE$ = "O/F a emitir": TIPO% = 12
                If DPI%(I&) = 0 Then REFE$ = "Tentativo de Compras": TIPO% = 11
                UII% = I&
'If UII% = CICONT% Then
'FEX = INIPER%(J%): ABC$ = FECHATEX$(FEX)
'AAA = BBB
'End If
                Call COBERFA(INIPER%(J%), UII%, REFE$, TIPO%, CAFAB)
            End If
        Next J%
    Next I&
    '
    For I& = 1 To X3&
If I& Mod 1000 = 0 Then
AAA = BBB
End If
        CIIX& = I&
        Call TRACE(20, CIIX&, X3&)
        Get #N9%, I&, B00$
        B1$ = Left$(B00$, 2)
        B3$ = Mid$(B00$, 3, 4)
        B4$ = Mid$(B00$, 7, 2)
        B5$ = Mid$(B00$, 9, 29)
        B7$ = Mid$(B00$, 42, 4)
        B8$ = Mid$(B00$, 46, 2)
        '
        CI% = CVI(B1$)
        If STATUIT%(CI%) >= 0 Then
          CANEC = (-1) * CVS(B3$)
          FFX% = CVI(B4$)
          REFE$ = B5$
          TIPO% = 22
          If Left$(REFE$, 20) <> "PROGRAMA POR MINIMOS" Then
'If CI% = CICONT% Then
'FEX = FFX%: ABC$ = FECHATEX$(FEX)
'AAA = BBB
'End If
            Call COBERFA(FFX%, CI%, REFE$, TIPO%, CANEC)
          End If
        End If
    Next I&
    '
    FORCANEC.Check6.Value = 1
    '
3900 Call CIERRARCH("*.*")
     If N4% > 0 Then Close #N4%: N4% = 0
     If N5% > 0 Then Close #N5%: N5% = 0
     If N7% > 0 Then Close #N7%: N7% = 0
     If N9% > 0 Then Close #N9%: N9% = 0
     If N13% > 0 Then Close #N13%: N13% = 0
     '
     'Call BLANARCH("PERPROMA")
     KKLL& = 0
     For KKK% = 1 To 52
       If TRIM$(DEPER$(KKK%)) <> "" Then
         KKLL& = KKLL& + 1
         Call GRAREG("PERPROMA", DEPER$(KKK%) + Space$(16), KKLL&)
       End If
     Next KKK%
     Call SETULTREG("PERPROMA", KKLL&)
     Call CIERRARCH("PERPROMA")
     '
     For UI& = 0 To 32767
       PROVINT%(UI&) = 0
     Next UI&
     '
     For U& = 1 To ULTREG&("TAPROINT")
       X$ = REGLEIDO$("TAPROINT", U&)
       NP% = CVI(Left$(X$, 2))
       If NP% > 0 Then
         PROVINT%(NP%) = 1
       End If
     Next U&
     Call CIERRARCH("TAPROINT")
     '
     JC& = 0    'Call BLANARCH("PLACOMP")
     JF& = 0    'Call BLANARCH("PLAFABR")
     'Call BLANARCH("ORFAEMI")
     'Call BLANARCH("OCOMEMI")
     '
     FORCANEC.Check7.Value = 2
     Screen.MousePointer = 11
     '
     FIN& = ULTREG&("ARTIPRO")
     For III& = 1 To FIN&
        Call TRACE(20, III&, FIN&)
        X$ = REGLEIDO$("ARTIPRO", III&)
        CI% = CVI(Left$(X$, 2))
        If CI% > 0 Then
            If CI% <= NUMAS% Then
                PROVEX$ = Mid$(X$, 3, 2)
                If ECOARCH$("PROVED1", PROVEX$) <> "" Then
                    PROVARTI%(CI%) = CVI(PROVEX$)
                End If
            End If
        End If
     Next III&
     '
     URECO& = CACO%
     FIN& = URECO&
     For III& = 1 To FIN&
        '
        Call TRACE(20, III&, FIN&)
        X$ = REGLEIDO$("COBERFA", III&)
        TIMO% = CVI(Mid$(X$, 51, 2))
        CI% = CVI(Mid$(X$, 3, 2))
        FF% = CVI(Mid$(X$, 1, 2))
        PERIO$ = Mid$(X$, 5, 16)
        CANTX = CVS(Mid$(X$, 53, 4))
        ' FFI% = CVI(Mid$(X$, 61, 2))
        '
        If TIMO% = 12 Then
          PHB% = PROVHABI%(CI%) ' PROVEEDOR HABITUAL
          If PHB% > 0 Then
            If PROVINT%(PHB%) = 0 Then ' NO ES PROVEEDOR INTERNO
              TIMO% = 11
            End If
          End If
        End If
        '
        Call REPLA(X$, MKI$(TIMO%), 51, 2)
        Call REPLA(X$, MKS$(0), 61, 4)
        Call GRAREG("COBERFA", X$, III&)
        '
        If TIMO% = 11 Then
             '
             MONEI% = MONECOSTO%(CI%)
             COSUDOL = COSUNI(CI%)
             COSUN = COSUDOL * TICAMONE(MONEI%)
             PHB% = PROVHABI%(CI%) ' PROVEEDOR HABITUAL
             If PHB% > 0 Then
               If PROVINT%(PHB%) <> 0 Then ' ES PROVEEDOR INTERNO
                 COSUN = 0
               End If
             End If
             '
             Z$ = REGBLAN$("PLACOMP")
             Call REPLA(Z$, MKI$(CI%), 1, 2)
             Call REPLA(Z$, MKI$(FF%), 3, 2)
             Call REPLA(Z$, PERIO$, 5, 16)
             Call REPLA(Z$, MKS$(CANTX), 21, 4)
             Call REPLA(Z$, MKS$(0), 25, 4) ' tot.progra.por item
             Call REPLA(Z$, MKS$(0), 29, 4) ' filler
             Call REPLA(Z$, MKI$(PHB%), 33, 2)
             Call REPLA(Z$, MKS$(COSUN), 35, 4)
             Call REPLA(Z$, MKD$(COSUN * CANTX), 39, 8)
             '
             JC& = JC& + 1
             Call GRAREG("PLACOMP", Z$, JC&)
             '
             FF1$ = REGBLAN$("OCOMEMI")
             Call REPLA(FF1$, MKI$(CI%), 1, 2)
             Call REPLA(FF1$, DESCRIT$(CI%), 3, 26)
             Call REPLA(FF1$, CSTRING$(MKS$(CANTX), 4, 10, 1, 2), 29, 10)
             FEX = FF%
             Call REPLA(FF1$, FECHATEX$(FEX), 41, 8)
             Call REPLA(FF1$, MKD$(CDbl(CANTX)), 49, 8)
             Call REPLA(FF1$, MKI$(FF%), 57, 2)
             Call REPLA(FF1$, MKI$(CI%), 59, 2)
             Call GRAREG("OCOMEMI", FF1$, JC&)
             '
           ElseIf TIMO% = 12 Then
             '
             Z$ = REGBLAN$("PLAFABR")
             Call REPLA(Z$, MKI$(CI%), 1, 2)
             Call REPLA(Z$, MKI$(FF%), 3, 2)
             Call REPLA(Z$, PERIO$, 5, 16)
             Call REPLA(Z$, MKS$(CANTX), 21, 4)
             Call REPLA(Z$, MKS$(0), 29, 4)  ' FILLER
             JF& = JF& + 1
             Call GRAREG("PLAFABR", Z$, JF&)
             '
             FF1$ = REGBLAN$("ORFAEMI")
             Call REPLA(FF1$, MKI$(CI%), 1, 2)
             Call REPLA(FF1$, DESCRIT$(CI%), 3, 26)
             Call REPLA(FF1$, CSTRING$(MKS$(CANTX), 4, 10, 1, 2), 29, 10)
             FEX = FF%
             Call REPLA(FF1$, FECHATEX$(FEX), 41, 8)
             Call REPLA(FF1$, MKD$(CDbl(CANTX)), 49, 8)
             Call REPLA(FF1$, MKI$(FF%), 57, 2)
             Call REPLA(FF1$, MKI$(CI%), 59, 2)
             Call GRAREG("ORFAEMI", FF1$, JF&)
             '
        End If
        '
     Next III&
     FORCANEC.Check7.Value = 1
     '
     Screen.MousePointer = 1
     URECO& = CACO%
     Call SETULTREG("COBERFA", URECO&)
     Call CIERRARCH("COBERFA")
     Call SETULTREG("ORFAEMI", JF&)
     Call CIERRARCH("ORFAEMI")
     Call SETULTREG("OCOMEMI", JC&)
     Call CIERRARCH("OCOMEMI")
     Call SETULTREG("PLACOMP", JC&)
     Call CIERRARCH("PLACOMP")
     Call SETULTREG("PLAFABR", JF&)
     Call CIERRARCH("PLAFABR")
     Call CIERRARCH("PERPROMA")
     '
     FECHOR$ = HOS$ + "  " + Left$(Time$, 5)
     Call GRACONTROL("CANECOM", "FECHOCAL", FECHOR$)
     If UCase$(Left$(App.EXEName, 5)) = "FLEXO" Then
         PROMAT00.Label7.Caption = Left$(FECHOR$, 8): PROMAT00.Label7.Refresh
         PROMAT00.Label95.Caption = Mid$(FECHOR$, 11, 5): PROMAT00.Label95.Refresh
       Else
         PROMATER.Label7.Caption = Left$(FECHOR$, 8): PROMATER.Label7.Refresh
         PROMATER.Label95.Caption = Mid$(FECHOR$, 11, 5): PROMATER.Label95.Refresh
     End If
     Call CIERRARCH("*.*")
     Call BAJALDISCO
     '
     'Call FINAL(SALI$)
     Call FILESORT("COBERFA", "", 1, 2, "ASC")
     Call FILESORT("PLACOMP", "", 1, 3, "ASC")
     Call FILESORT("PLAFABR", "", 1, 3, "ASC")
     Call FILESORT("ORFAEMI", "", 4, 5, "ASC")
     Call FILESORT("OCOMEMI", "", 4, 5, "ASC")
     '
     CIAX% = 0
     DEPOSTAKK% = Val(TRIM$(CONTROL$("PROMATER", "DEPOSTOK")))
     '
     For U& = 1 To ULTREG&("PLACOMP")
       XX$ = REGLEIDO$("PLACOMP", U&)
       CII% = CVI(Left$(XX$, 2))
       If CII% <> CIAX% Then
         CIAX% = CII%
         STAKK = STODEPOS(CII%, DEPOSTAKK%): If STAKK < 0 Then STAKK = 0
         Call REPLA(XX$, MKS$(STAKK), 29, 4)
         Call REPLA(XX$, MKS$(STOMIN(CII%)), 55, 4)
         Call REPLA(XX$, MKS$(LOTESTAN(CII%)), 59, 4)
         Call GRAREG("PLACOMP", XX$, U&)
       End If
     Next U&
     Call CIERRARCH("PLACOMP")
     Call CARPROSEL
     '
     CIAX% = 0
     For U& = 1 To ULTREG&("PLAFABR")
       XX$ = REGLEIDO$("PLAFABR", U&)
       CII% = CVI(Left$(XX$, 2))
       If CII% <> CIAX% Then
         CIAX% = CII%
         STAKK = STODEPOS(CII%, DEPOSTAKK%): If STAKK < 0 Then STAKK = 0
         Call REPLA(XX$, MKS$(STAKK), 29, 4)
         Call REPLA(XX$, MKS$(STOMIN(CII%)), 55, 4)
         Call REPLA(XX$, MKS$(LOTESTAN(CII%)), 59, 4)
         Call GRAREG("PLAFABR", XX$, U&)
       End If
     Next U&
     Call CIERRARCH("PLAFABR")
     '
999  Call CIERRARCH("*.*")
     Call BAJALDISCO
     Call COMUNI("Re-Cálculo Completo.")
     '
     FORCANEC.Hide
     Exit Sub
     '
' *********************************************************
     '
5300 If NESTRU& > 0 Then GoTo 5380
     '
     PUNTCI$ = "": PUNTRG$ = ""
     NESTRU& = 0
     FIN& = ULTREG&("ESTRUNUE")
     For I& = 1 To FIN&
        RA$ = REGLEIDO$("ESTRUNUE", I&)
        RB$ = REGLEIDO$("ESTRUDOS", I&)
        CI% = CVI(Left$(RA$, 2))
        If CI% > 0 Then
           If CI% <= NUMAS% Then
              If DPI%(CI%) = 0 Then
                 DPI%(CI%) = 1
                 NESTRU& = NESTRU& + 1
                 Call GRAREG("ESTRUNUE", RA$, NESTRU&)
                 Call GRAREG("ESTRUDOS", RB$, NESTRU&)
                 PUNTCI$ = PUNTCI$ + MKI$(CI%)
                 NESTRI% = NESTRU&
                 PUNTRG$ = PUNTRG$ + MKI$(NESTRI%)
              End If
           End If
        End If
     Next I&
     '
     For I& = 1 + ULTREG&("ESTRUDOS") To NESTRU&
        RB$ = String$(256, 0)
        Call REPLA(RB$, MKS$(0.0001), 251, 4)
        Call GRAREG("ESTRUDOS", RB$, I&)
     Next I&
     '
     Call SETULTREG("ESTRUNUE", NESTRU&)
     Call SETULTREG("ESTRUDOS", NESTRU&)
     Call CIERRARCH("ESTRUNUE")
     Call CIERRARCH("ESTRUDOS")
     '
     For K% = 1 To NUMAS%: DPI%(K%) = 0: Next K%
     '
5380 Return
     '
17000 ENCU% = 0: PX1% = 1

17005 PX% = InStr(PX1%, PUNTCI$, MKI$(CIBAS%))
      If PX% < PX1% Then GoTo 17035
      If (PX% Mod 2) <> 1 Then PX1% = PX% + 1: GoTo 17005
      L% = CVI(Mid$(PUNTRG$, PX%, 2))
      ENCU% = L%
      '
      Get #N5%, L%, RA$
      R1$ = Left$(RA$, 2)
      R2$ = Mid$(RA$, 3, 120)
      R3$ = Mid$(RA$, 123, 6)
      '
17035 Return
17040 '

17200 Get #N5%, ENCU%, RA$
      R1$ = Left$(RA$, 2)
      R2$ = Mid$(RA$, 3, 120)
      R3$ = Mid$(RA$, 123, 6)
      DPS$ = R2$: J% = 0
      '
17205 For K% = 1 To 120 Step 8
17210    DPSX$ = Mid$(DPS$, K%, 8)
17215    If CVI(DPSX$) < 1 Or CVI(DPSX$) > NMS& Then GoTo 17280
         If CVS(Mid$(DPSX$, 3)) < 0.0005 Then GoTo 17225
17220    J% = J% + 1
         CII1%(J%) = CVI(DPSX$)
         CANT1(J%) = CABAS * CVS(Mid$(DPSX$, 3))
         PAD1%(J%) = Asc(Mid$(DPSX$, 7))
         If PAD1%(J%) < 1 Then
             PAD1%(J%) = PLADISTA%
         End If
17225 Next K%
      '
17230 Get #N7%, ENCU%, RB$
      R4$ = RB$
      DPS$ = R4$
17235 For K% = 1 To 256 Step 8
17240    DPSX$ = Mid$(DPS$, K%, 8)
17245    If CVI(DPSX$) < 1 Or CVI(DPSX$) > NMS& Then GoTo 17280
         If CVS(Mid$(DPSX$, 3)) < 0.0005 Then GoTo 17255
17250    J% = J% + 1
         CII1%(J%) = CVI(DPSX$)
         CANT1(J%) = CABAS * CVS(Mid$(DPSX$, 3))
         PAD1%(J%) = Asc(Mid$(DPSX$, 7))
         If PAD1%(J%) < 1 Then
             PAD1%(J%) = PLADISTA%
         End If
17255 Next K%
17260 '
17280 Return
17285 '

interrum:
      'RESP% = COMALTER%("Proceso Interrumpido a su Comando\\Reanudar\Abortar")
      'If RESP% <> 1 Then Call FINAL("")
      Return

End Sub
'

Sub COBERFA(FFI%, CI%, REFE$, TIPO%, CAN)
      '
      Static MESU$(12)
      Static CTX%
      '
      If CTX% = 0 Then
         '
         CACO% = 0
         '
         MESU$(1) = "ENERO"
         MESU$(2) = "FEBRERO"
         MESU$(3) = "MARZO"
         MESU$(4) = "ABRIL"
         MESU$(5) = "MAYO"
         MESU$(6) = "JUNIO"
         MESU$(7) = "JULIO"
         MESU$(8) = "AGOSTO"
         MESU$(9) = "SETIEMBRE"
         MESU$(10) = "OCTUBRE"
         MESU$(11) = "NOVIEMBRE"
         MESU$(12) = "DICIEMBRE"
         '
         For NMM% = 1 To NUMAS%
            PUNTE$(NMM%) = ""
         Next NMM%
         '
         CTX% = 1
      End If

      If SEMANAL% = 1 Then
         FFX = FFI%
         SMN% = SEMANU%(FFX, A, B)
         MESANIO$ = "S." + Mid$(Str$(SMN%), 2) + " - " + FECHATEX$(A)
         MESANIOT% = FFI%
         '
         FF$ = FECHATEX$(FFX)
         ANI% = Val(Mid$(FF$, 7, 2))
         MES% = Val(Mid$(FF$, 4, 2))
         ANIOX% = ANI%: If ANIOX% > 15 Then ANIOX% = 0
         MESANXX& = 30 * (54 * ANIOX% + SMN%) + TIPO%
         If MESANXX& > 32767 Then MESANXX& = 32767
         If MESANXX& < (-32767) Then MESANXX& = (-32767)
         MESANIOI% = MESANXX&
         '
         GoTo 25902
      End If
      '
      FFX = FFI%
      FF$ = FECHATEX$(FFX)
      ANI% = Val(Mid$(FF$, 7, 2))
      ANIO% = Val(Mid$(FF$, 7, 2)) - 90
      If ANIO% < 0 Then ANIO% = ANIO% + 100
      MES% = Val(Mid$(FF$, 4, 2))
      MESANIO$ = MESU$(MES%) + Str$(ANI%)
      MESANIOT% = 1000 * ANIO% + 10 * MES%
      '
      ANIOX% = ANI%: If ANIOX% > 20 Then ANIOX% = 0
      MESANIOI% = 100 * (13 * ANIOX% + MES%) + TIPO%
      '
25902 For UX% = 1 To Len(PUNTE$(CI%)) Step 2
         REG& = CVI(Mid$(PUNTE$(CI%), UX%, 2))
         FF0$ = REGLEIDO$("COBERFA", REG&)
            F1$ = Left$(FF0$, 2)
            F2$ = Mid$(FF0$, 3, 2)
            F4$ = Mid$(FF0$, 5, 16)
            F5$ = Mid$(FF0$, 21, 30)
            F0$ = Mid$(FF0$, 51, 2)
            F6$ = Mid$(FF0$, 53, 4)
            F7$ = Mid$(FF0$, 57, 4)
            '
            'If CVI(F1$) = MESANIOI% Then
            If CVI(F1$) = FFI% Then
                If CVI(F2$) = CI% Then
                    If TRIM$(F4$) = TRIM$(MESANIO$) Then
                        If TRIM$(F5$) = TRIM$(REFE$) Then
                            If CVI(F0$) = TIPO% Then
                                If Sgn(CAN) = Sgn(CVS(F6$) - CVS(F7$)) Then
                                    If CAN >= 0 Then
                                          Call REPLA(FF0$, MKS$(CVS(F6$) + CAN), 53, 4)
                                          F6$ = Mid$(FF0$, 53, 4)
                                        ElseIf CAN < 0 Then
                                          Call REPLA(FF0$, MKS$(CVS(F7$) - CAN), 57, 4)
                                          F7$ = Mid$(FF0$, 57, 4)
                                    End If
                                    Call GRAREG("COBERFA", FF0$, REG&)
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End If
            End If
      Next UX%
      '
      FF0$ = REGBLAN$("COBERFA")
      'Call REPLA(FF0$, MKI$(MESANIOI%), 1, 2)
      Call REPLA(FF0$, MKI$(FFI%), 1, 2)
      Call REPLA(FF0$, MKI$(CI%), 3, 2)
      Call REPLA(FF0$, MESANIO$, 5, 16)
      Call REPLA(FF0$, REFE$, 21, 30)
      Call REPLA(FF0$, MKI$(TIPO%), 51, 2)
      If CAN >= 0 Then
            Call REPLA(FF0$, MKS$(CAN) + MKS$(0) + MKS$(0), 53, 12)
          Else
            Call REPLA(FF0$, MKS$(0) + MKS$(Abs(CAN)) + MKS$(0), 53, 12)
      End If
      Call REPLA(FF0$, MKI$(FFI%), 61, 2)
      CACO% = CACO% + 1
      RECOBE& = CACO%
      Call GRAREG("COBERFA", FF0$, RECOBE&)
      '
      PUNTE$(CI%) = PUNTE$(CI%) + MKI$(CACO%)
      '
      If TRIM$(MESANIO$) <> "" Then
        For KKK% = 52 To 1 Step -1
          If MESANIOT% >= INIPER%(KKK%) Then
             If DEPER$(KKK%) = "" Then
                DEPER$(KKK%) = MESANIO$
             End If
             Exit Sub
          End If
        Next KKK%
      End If
      '
End Sub



Sub PRODISP(FF%, CI%, TIPO%, CAN)

Static SEMANAL%, CTX%, CAPER%, N9%, N98%, F1$, F2$, F4$, F5$, F7$, F8$, F9$, F10$, G1$, G2$, LU$
Static MES$()

If CTX% = 0 Then
      LU$ = LUDAT$
'      CT$ = UCase$(CTRLFAC$("PERIODO"))
      SEMANAL% = 0: If InStr(CT$, "SEMAN") > 0 Then SEMANAL% = 1

35901 N9% = FreeFile
'      OPEN "R",#N9%,LU$+"PERIPRO.DAT",32
'      FIELD #N9%,6 AS G1$,26 AS G2$
      If LOF(N9%) > 0 Then
         Close #N9%
         Kill LU$ + "PERIPRO.DAT"
         GoTo 35901
      End If
      CAPER% = 0

 '     Dim MES$(12)
'35902 Data ENERO, FEBRERO, MARZO, ABRIL, MAYO, JUNIO, JULIO, AGOSTO, SETIEMBRE, OCTUBRE, NOVIEMBRE, DICIEMBRE
'      RESTORE 35902
'      For M% = 1 To 12
'        READ MES$(M%)
'      Next M%

      CTX% = 1
End If

If SEMANAL% = 1 Then
'    FF = FF%: FF$ = FECHATEX$(FF)
'    FF = FF%: SMN% = SEMANU%(FF, A, B)
'    AN$ = Mid$(FF$, 7, 2): If Left$(AN$, 1) = "9" Then AN$ = " " + AN$ Else AN$ = "1" + AN$
'    SMN$ = Mid$(Str$(SMN%), 2): While Len(SMN$) < 2: SMN$ = "0" + SMN$: Wend
'    CODPER$ = AN$ + "-" + SMN$
    PERIO$ = "S." + Mid$(Str$(SMN%), 2) + " - "
'    A$ = FECHATEX$(A): B$ = FECHATEX$(B)
    If Mid$(A$, 3) = Mid$(B$, 3) Then A$ = Left$(A$, 2)
    If Mid$(A$, 6) = Mid$(B$, 6) Then A$ = Left$(A$, 5)
    If Left$(A$, 1) = "0" Then A$ = Mid$(A$, 2)
    PERIO$ = PERIO$ + A$ + " al " + B$
    GoTo 35905
End If

'    FF = FF%: FF$ = FECHATEX$(FF)
'    AN$ = Mid$(FF$, 7, 2): If Left$(AN$, 1) = "9" Then AN$ = " " + AN$ Else AN$ = "1" + AN$
'    Me = Mid$(FF$, 4, 2): MES% = Val(Me)
    CODPER$ = AN$ + "-" + Mexx
'    PERIO$ = MES$(MES%) + " " + AN$

35905 '

For G% = CAPER% To 1 Step -1
'    GET #N9%,G%
    If G1$ = CODPER$ Then GoTo 35990
Next G%

For G% = CAPER% To 1 Step -1
'    GET #N9%,G%
    If G1$ <= CODPER$ Then GoTo 35910
'    PUT #N9%,G%+1
Next G%
G% = 0

35910 '
G% = G% + 1
LSet G1$ = CODPER$
LSet G2$ = PERIO$
'PUT #N9%,G%
CAPER% = CAPER% + 1

35990 '
End Sub

Sub CANECOFER()
     '
     Dim BAGRE$(16384)
     Dim HAYPLAFAB%(32767)
     Dim CACOPEN(32767)
     '
     LU$ = LUDAT$
     HOO% = HOY(HO$)
     CACO% = 0
     Screen.MousePointer = 11
     '
100  ' ************************* ABRIENDO ARCHIVOS MAESTROS
     '
     FORCANEC.Check1.Value = 2
     Screen.MousePointer = 11
     NMS& = NUMAS%
     NVI% = NUINV%
     'N9% = FILENBR%("SALSTOCK")
     N9% = FILEOPEN%(LUDAT$ + "SALSTOCK.DAT", 0, 40)
     '
     GESUBOR% = 0
     GoSub 5300                           ' estrunue - estrudos
     FORCANEC.Check1.Value = 1
     '
     If NESTRU& < 1 Then
         Call MENSERR(24, "Proceso Imposible.\ \Faltan Definir Fórmulas y\Estructuras de Producto.")
         Call FINAL("")
     End If
     '
     PLADISTA% = Val(FORCANEC.Text1.TEXT)
     If PLADISTA% > 6 Then PLADISTA% = 6
     '
'    **** DEPURA PROGRAMA DE ENTREGAS POR ENTREGAS REALIZADAS
     '
     APROGRA$ = "PROGENT"
     If MODOPLAN% = 1 Then APROGRA$ = "PROGFAB"
     '
     UREPRO& = ULTREG&(APROGRA$)
     JJ& = 0
     For I& = 1 To UREPRO&
       ZZ$ = REGLEIDO$(APROGRA$, I&)
       If CVS(Mid$(ZZ$, 7, 4)) > 0.05 Then
         Call REPLA(ZZ$, MKS$(0), 11, 4)
         Call REPLA(ZZ$, MKS$(0), 19, 4)
         JJ& = JJ& + 1
         Call GRAREG(APROGRA$, ZZ$, JJ&)
       End If
       If MODOPLAN% = 1 Then
         ' marca items programados para no considerar stock como disponibilidad
         CIXI% = CVI(Mid$(ZZ$, 1, 2))
         If CIXI% > 0 Then
           If CIXI% <= NUMAS% Then
             CAPROBRU = CVS(Mid$(ZZ$, 7, 4))
             If CAPROBRU > 0.05 Then
               HAYPLAFAB%(CIXI%) = 1
             End If
           End If
         End If
         '
       End If
     Next I&
     Call SETULTREG(APROGRA$, JJ&)
     Call CIERRARCH(APROGRA$)
     '
CMXXX% = COMALTER%("Puede Descontar del Programa\las Entregas Realizadas en el Mes en Curso\\No Descontarlas\Si, Netear Programa")
If CMXXX% <> 2 Then GoTo 103
     '
     CMA% = COMESACT%
     If MODOPLAN% <> 1 Then
       FIN& = ULTREG&("MOVISTO")
       'FINUMO% = FILENBR%("MOVISTO")
       FINUMO% = FILEOPEN%(LUDAT$ + "MOVISTO.DAT", 0, 128)
       Dim RMVS As String * 128
       For I& = 1 To FIN&
         Call TRACE(20, I&, FIN&)
         Get #FINUMO%, I&, RMVS$
         XS$ = RMVS$
         FF% = CVI(Left$(XS$, 2))
         If FF% >= CMA% Then
           COMO$ = UCase$(Mid$(XS$, 33, 2))
           If COMO$ = "FC" Or COMO$ = "NC" Then
             CIXI% = CVI(Mid$(XS$, 3, 2))
             If CIXI% > 0 Then
               If CIXI% <= NUMAS% Then
                 CAENTREG = CVD(Mid$(XS$, 96, 8)) - CVD(Mid$(XS$, 88, 8))
                 '
                 RFF$ = RECFILT$(APROGRA$, 1, MKI$(CIXI%))
                 For KK1& = 1 To Len(RFF$) Step 4
                   REPRO& = CVS(Mid$(RFF$, KK1&, 4))
                   ZZ$ = REGLEIDO$(APROGRA$, REPRO&)
                   ACUENTRE = CVS(Mid$(ZZ$, 11, 4))
                   ENTREMA = CVS(Mid$(ZZ$, 7, 4)) - ACUENTRE
                   If CAENTREG <= ENTREMA Then
                      Call REPLA(ZZ$, MKS$(CAENTREG + ACUENTRE), 11, 4)
                      CAENTREG = 0
                    Else
                      Call REPLA(ZZ$, MKS$(ENTREMA + ACUENTRE), 11, 4)
                      CAENTREG = CAENTREG - ENTREMA
                   End If
                   Call GRAREG(APROGRA$, ZZ$, REPRO&)
                   If CAENTREG < 0.05 Then GoTo 102
                 Next KK1&
               End If
             End If
           End If
         End If
102    Next I&
       Call CIERRARCH(APROGRA$)
     End If
     '
'
' ************** lee existencias, stock minimo y lote reposicion *******************

103  FORCANEC.Check2.Value = 2
     Screen.MousePointer = 11
     '
     FIN& = NESTRU&
     For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        RA$ = REGLEIDO$("ESTRUNUE", I&)
        CI% = CVI(Left$(RA$, 2))
        If CI% > 0 Then
            If CI% <= NMS& Then
                R2$ = Mid$(RA$, 3, 120)
                For K% = 1 To 120 Step 8
                    DPSX$ = Mid$(R2$, K%, 8)
                    If CVI(DPSX$) < 1 Or CVI(DPSX$) > NMS& Then GoTo 105
                    If CVS(Mid$(DPSX$, 3)) >= 0.00005 Then DPI%(CI%) = DPI%(CI%) + 1
                    If DPI%(CI%) > 1 Then GoTo 105
                Next K%
            End If
        End If
105  Next I&
     '
     HAYNE% = 0
     MODO% = 0
     If FORCANEC.Option1.Value = True Then MODO% = 1
     If FORCANEC.Option2.Value = True Then MODO% = 2
     '
     FFXY = HOY(HO$)
     SMXX% = SEMANU%(FFXY, AA, BB)
     COSEMACT% = AA
     '
     QLN& = 0
     Call CIERRARCH("COBERFA")
     'Call BLANARCH("COBERFA")
     Call CIERRARCH("ORFAEMI")
     'Call BLANARCH("ORFAEMI")
     Call CIERRARCH("OCOMEMI")
     'Call BLANARCH("OCOMEMI")
     '
     For CI% = 1 To NMS&
       CIIX& = CI%
       Call TRACE(20, CIIX&, NMS&)
       Get #N9%, CI% + 1, SST$
       SALSTO = 0
       If STATUIT%(CI%) >= 0 Then
         For K% = 1 To 40 Step 4
            SALSTO = SALSTO + CVS(Mid$(SST$, K%, 4))
         Next K%
         '
         If MODO% = 1 Then SALSTO = 0
         If MODO% = 2 Then If SALSTO < 0 Then SALSTO = 0

         If SALSTO < 0 Then
            HAYNE% = HAYNE% + 1
            CI1% = CI%
            TTXX$ = REGBLAN$("LINEGAT")
            Call REPLA(TTXX$, MKI$(CI%), 1, 2)
            Call REPLA(TTXX$, DESCRIT$(CI%), 3, 30)
            Call REPLA(TTXX$, CSTRING$(MKS$(SALSTO), 4, 10, 1, 2), 33, 10)
            Call REPLA(TTXX$, UNIMED$(CI%), 45, 4)
            '
            QLN& = QLN& + 1
            Call GRAREG("LINEGAT", TTXX$, QLN&)
            '
         End If
         '
         STMIN = STOMIN(CI%): If MODO% = 1 Then STMIN = 0
         REPOS = LOTESTAN(CI%)
         DP% = DPI%(CI%): If DP% < 0 Or DP% > 3 Then DP% = 0

         If REPOS < 0.5 Then REPOS = LOSTN(DP%)
         If REPOS < 0.5 Then REPOS = REPOMIN

         DISPACT(CI%) = SALSTO - STMIN
         If HAYPLAFAB%(CI%) > 0 Then DISPACT(CI%) = 0
         LOREPOS(CI%) = REPOS
         If HAYPLAFAB%(CI%) < 1 Then
           If Abs(SALSTO) > 0.05 Then
             Call COBERFA(COSEMACT%, CI%, "Stock al " + HO$, 10, SALSTO)
           End If
           '
           If Abs(STMIN) > 0.05 Then
             Call COBERFA(COSEMACT%, CI%, "Stock Minimo del Item", 20, (-1) * STMIN)
           End If
         End If
       End If
     Next CI%
     '
     Call SETULTREG("LINEGAT", QNL&)
     Call CIERRARCH("LINEGAT")
     '
     If HAYNE% > 0 Then
        Call FILESORT("LINEGAT", "", 1, 1, "ASC")
        A% = COMALTER%("ATENCION: Se Detectaron Existencias Negativas\          en" + Str$(HAYNE%) + " Items de Stock.\\Revisar\Cancelar\Continuar")
        On A% GoTo 107, 106, 108
106     Call CIERRARCH("*.*")
        Screen.MousePointer = 1
        Exit Sub
107     Call SELECHO("LINEGAT")
        VDEVU$ = VALACT1$("LINEGAT")
        If UCase$(Left$(App.EXEName, 5)) = "FLEXO" Then
          If VDEVU$ <> "" Then
            CONSTO00.Text1 = VDEVU$
            CONSTO00.Show 1
            GoTo 107
          End If
        End If
        '
        If ALTERNA%("Continuar Proceso") <> 1 Then
            Call CIERRARCH("*.*")
            Screen.MousePointer = 1
            Exit Sub
        End If
     End If
     '
108  FORCANEC.Check2.Value = 1
     FORCANEC.Refresh
     FORCANEC.Refresh

' ******** inicializa necesidades secundarias de compra y fabricacion ********

110  FORCANEC.Check3.Value = 2
     Screen.MousePointer = 11
     '
111  N13% = FreeFile
     Open LU$ + "NEFACOM.DAT" For Random Access Read Write Shared As #N13% Len = LOREC%
     Dim PDS As String * 208
     Dim PDM As String * 72
     '
     If LOF(N13%) > 0 Then
         Close #N13%
         Kill LU$ + "NEFACOM.DAT"
         GoTo 111
     End If
     '
     For I& = 1 To NMS&
         CIIX& = I&
         Call TRACE(20, CIIX&, NMS&)
         If LOREC% = 208 Then
              Get #N13%, I&, PDS$
            Else
              Get #N13%, I&, PDM$
         End If
         PDS$ = String$(208, 0)
         PDM$ = String$(72, 0)
         If LOREC% = 208 Then
              Put #N13%, I&, PDS$
            Else
              Put #N13%, I&, PDM$
         End If
     Next I&

     FORCANEC.Check3.Value = 1


' **************** necesidades brutas segun entregas programadas **************

120  N9% = FreeFile
     Open LU$ + "NEBRUNE.DAT" For Random Access Read Write Shared As #N9% Len = 48
     Dim B00 As String * 48
     Dim B01 As String * 48
     X1& = 1: X2& = 0

     If LOF(N9%) > 0 Then
        Close #N9%
        Kill LU$ + "NEBRUNE.DAT"
        GoTo 120
     End If

     FORCANEC.Check4.Value = 2
     Screen.MousePointer = 11
     '
121  FIN& = NUMAS%
     For I& = 1 To FIN&
       CIIX& = I&
       Call TRACE(20, CIIX&, FIN&)
       CI% = I&
       If STATUIT%(CI%) >= 0 Then
         '
         If DISPACT(CI%) < 0 Then
            CANTU = Abs(DISPACT(CI%))
            FF% = HOO%
            '
            X2& = X2& + 1
            B00$ = String$(48, 0)
            Call REPLA(B00$, MKI$(CI%), 1, 2)
            Call REPLA(B00$, MKS$(CANTU), 3, 4)
            Call REPLA(B00$, MKI$(FF%), 7, 2)
            Call REPLA(B00$, "PROGRAMA POR MINIMOS", 9, 29)
            Call REPLA(B00$, MKS$(0), 38, 4)
            Call REPLA(B00$, MKS$(0), 42, 4)
            Call REPLA(B00$, MKS$(0), 46, 2)
            Call REPLA(B00$, " ", 48, 1)
            '
            Put #N9%, X2&, B00$
            DISPACT(CI%) = 0
            '
         End If
       End If
122  Next I&
     '
     N4% = FreeFile
     Open LU$ + APROGRA$ + ".DAT" For Random Access Read Write Shared As #N4% Len = 24
     Dim P00 As String * 24
     '
125  FIN& = LOF(N4%) / 24
     For I& = 1 To FIN&
       CIIX& = I&
       Call TRACE(20, CIIX&, FIN&)
       Get #N4%, I&, P00$
       P1$ = Left$(P00$, 2)
       CI% = CVI(P1$)
       '
       If CI% > 0 Then
         If CI% <= NUMAS% Then
           If STATUIT%(CI%) >= 0 Then
             '
             P3$ = Mid$(P00$, 3, 2)
             P6$ = Mid$(P00$, 7, 4)
             P7$ = Mid$(P00$, 11, 4)
             '
             If CVI(P3$) < COMESACT% Then
                cix% = CVI(P1$)
                Call MENSERR(24, "Error: Codigo " + TRIM$(CODEXT$(cix%)) + " Mal Fecha Entrega Programada")
                GoTo 129
             End If
             '
             PLAVEN = CVS(P6$)
             ENTREG = CVS(P7$)
             PEDIPE = CVS(Mid$(P00$, 19, 4))
             CANTU = PLAVEN - ENTREG
             '
             PEDINOPRO = PEDIPE + ENTREG - PLAVEN
             If PEDINOPRO < 0 Then PEDINOPRO = 0
             '
             CI% = CVI(P1$)
             FF% = CVI(P3$)
             '
             If CANTU > 0 Then
               If CI% > 0 Then
                  If CI% <= NMS& Then
                     DPI%(CI%) = 3
                  End If
               End If
             End If
             '
             If CANTU > 0 Then
               '
               X2& = X2& + 1
               B00$ = String$(48, 0)
               Call REPLA(B00$, MKI$(CI%), 1, 2)
               Call REPLA(B00$, MKS$(CANTU), 3, 4)
               Call REPLA(B00$, MKI$(FF%), 7, 2)
               Call REPLA(B00$, "PROGRAMA ENTREGAS AL " + HO$, 9, 29)
               If MODOPLAN% = 1 Then
                 Call REPLA(B00$, "PLAN DE FABRICACION AL " + HO$, 9, 29)
               End If
               If FF% < INIPER%(1) Then
                   Call REPLA(B00$, MKI$(INIPER%(1)), 7, 2)
                   Call REPLA(B00$, "ATRASO DE ENTREGA AL " + HO$, 9, 29)
               End If
               Call REPLA(B00$, MKS$(0), 38, 4)
               Call REPLA(B00$, MKS$(0), 42, 4)
               Call REPLA(B00$, MKS$(0), 46, 2)
               Call REPLA(B00$, " ", 48, 1)
               '
               Put #N9%, X2&, B00$
               '
             End If
             '
             If MODOPLAN% = 0 Then
               If PEDINOPRO > 0 Then
                 '
                 X2& = X2& + 1
                 B00$ = String$(48, 0)
                 Call REPLA(B00$, MKI$(CI%), 1, 2)
                 Call REPLA(B00$, MKS$(PEDINOPRO), 3, 4)
                 Call REPLA(B00$, MKI$(FF%), 7, 2)
                 Call REPLA(B00$, "PEDIDOS NO PROGRAMADOS AL " + HO$, 9, 29)
                 Call REPLA(B00$, MKS$(0), 38, 4)
                 Call REPLA(B00$, MKS$(0), 42, 4)
                 Call REPLA(B00$, MKS$(0), 46, 2)
                 Call REPLA(B00$, " ", 48, 1)
                 '
                 Put #N9%, X2&, B00$
                 '
               End If
             End If
             '
           End If
         End If
       End If
129  Next I&
     '
     FORCANEC.Check4.Value = 1
     '
130  CICLO% = CICLO% + 1
     CICLOMA% = 9
     If InStr(UCase$(EMPREAC$), "YALTRE") > 0 Then CICLOMA% = 15
     '
     If CICLO% > CICLOMA% Then GoTo 150
     FORCANEC.Label1 = TRIM$(Str$(CICLO%))
     FORCANEC.Label1.Refresh
     '
' ******** CALCULANDO NECESIDADES NETAS SEGUN DISPONIBILIDAD
     '
     FORCANEC.Check5.Value = 2
     Screen.MousePointer = 11
     For I& = X1& To X2&
        '
        CIIX& = I&
        Call TRACE(20, CIIX&, X2&)
        '
        Get #N9%, I&, B00$
        B1$ = Left$(B00$, 2)
        B3$ = Mid$(B00$, 3, 4)
        B4$ = Mid$(B00$, 7, 2)
        '
        CI% = CVI(B1$)
        If CI% > 0 Then
          If CI% <= NMS& Then
            If STATUIT%(CI%) >= 0 Then
      'If CI% = 484 Then
      'ABC$ = CODEXT$(CI%)
      'CCC = CVS(B3$)
      'AAA = BBB
      'End If
                If DISPACT(CI%) >= CVS(B3$) Then
                    DISPACT(CI%) = DISPACT(CI%) - CVS(B3$)
                    Call REPLA(B00$, B3$, 38, 4)
                    Call REPLA(B00$, MKS$(0), 42, 4)
                    Put #N9%, I&, B00$
                    GoTo 139
                End If
                '
                CANEC = CVS(B3$) - DISPACT(CI%)
                FENEBRU% = CVI(B4$): CUANDO% = 1
                '
'FEX = FENEBRU: ABC$ = FECHATEX$(FEX)
                For UX% = 1 To CACOL%
                    If INIPER%(UX%) <= FENEBRU% Then CUANDO% = UX%
                Next UX%
                '
                If LOREC% = 208 Then
                      Get #N13%, CI%, PDS$
                      PDX$ = PDS$
                    Else
                      Get #N13%, CI%, PDM$
                      PDX$ = PDM$
                End If
                '
                For UX% = CUANDO% To 1 Step -1
                    PROANT = CVS(Mid$(PDX$, 4 * UX% - 3, 4))
                    If PROANT > 0 Then GoTo 135
                Next UX%
                UX% = CUANDO%
135             If PROANT >= LOREPOS(CI%) Then UX% = CUANDO%
                '
                PROANT = CVS(Mid$(PDX$, 4 * UX% - 3, 4))
                Mid$(PDX$, 4 * UX% - 3, 4) = MKS$(PROANT + CANEC)
                '
                If LOREC% = 208 Then
                      PDS$ = PDX$
                      Put #N13%, CI%, PDS$
                    Else
                      PDM$ = PDX$
                      Put #N13%, CI%, PDM$
                End If
                '
                FENENE% = INIPER%(UX%)
                '
                Call REPLA(B00$, MKS$(DISPACT(CI%)), 38, 4)
                DISPACT(CI%) = 0
                Call REPLA(B00$, MKS$(CANEC), 42, 4)
                Call REPLA(B00$, MKI$(FENENE%), 46, 2)
                Put #N9%, I&, B00$
                '
            End If
          End If
        End If
139 Next I&



140 ' Call MENSAJE(14, "GENERANDO NECESIDADES SECUNDARIAS SEGUN EXPLOSION")

    X3& = X2&
    IAGRE% = 0
    For I& = X1& To X2&
        CIIX& = I&
        Call TRACE(20, CIIX&, X2&)
        Get #N9%, I&, B00$
        Call REPLA(B00$, "C", 48, 1)
        Put #N9%, I&, B00$
        '
        B1$ = Left$(B00$, 2)
        B3$ = Mid$(B00$, 3, 4)
        B4$ = Mid$(B00$, 7, 2)
        B7$ = Mid$(B00$, 42, 4)
        B8$ = Mid$(B00$, 46, 2)
        '
        CI% = CVI(B1$)
        If STATUIT%(CI%) >= 0 Then
          CABAS = CVS(B7$)
          FENEC% = CVI(B8$): CUANEC% = 1
                          For UX% = 1 To CACOL%
                             If INIPER%(UX%) <= FENEC% Then CUANEC% = UX%
                          Next UX%
          If CABAS > 0 Then
1234        CIBAS% = CI%
            CII% = CIBAS%
            CODBAS$ = TRIM$(CODEXT$(CII%))
            Call LEEEXPLOMUL(CODBAS$, CABAS, 1)
            If CAIT% > 0 Then
                Call REPLA(B00$, "F", 48, 1)
                Put #N9%, I&, B00$
                For KJ% = 1 To CAIT%    ' J%
                  CI% = CIJ%(KJ%)     '  CII1%(KJ%)
                  If STATUIT%(CI%) >= 0 Then
                    CANTU = CABAS * USOI(KJ%) 'CANT1(KJ%)
                    PLADISPO% = PDI%(KJ%)
                    If PLADISPO% < 1 Then PLADISPO% = PLADISTA%
                    CLX% = CUANEC% - PLADISPO%  'PAD1%(KJ%)
                    If CLX% < 1 Then CLX% = 1
                    If CANTU > 0 Then
                        'X3& = X3& + 1
                        B00$ = String$(48, 0)
                        Call REPLA(B00$, MKI$(CIJ%(KJ%)), 1, 2)
                        Call REPLA(B00$, MKS$(CANTU), 3, 4)
                        Call REPLA(B00$, MKI$(INIPER%(CLX%)), 7, 2)
                        Call REPLA(B00$, "Reserva para " + CODBAS$, 9, 28)
                        Call REPLA(B00$, " ", 48, 1)
                        IAGRE% = IAGRE% + 1
                        For JAGRE% = IAGRE% - 1 To 1 Step -1
                          If CVI(Mid$(BAGRE$(JAGRE%), 7, 2)) <= INIPER%(CLX%) Then
                            BAGRE$(JAGRE% + 1) = B00$
                            GoTo 1235
                          End If
                          BAGRE$(JAGRE% + 1) = BAGRE$(JAGRE%)
                        Next JAGRE%
                        BAGRE$(1) = B00$
                    End If
                  End If
1235            Next KJ%
            End If
          End If
        End If
    Next I&
    '
    For JAGRE% = 1 To IAGRE%
      X3& = X3& + 1
      B00$ = BAGRE$(JAGRE%)
      Put #N9%, X3&, B00$
    Next JAGRE%
    '
    X1& = X2& + 1
    X2& = X3&: If X2& < X1& Then GoTo 150 ' terminar
    GoTo 130

150 FORCANEC.Check5.Value = 1
    FORCANEC.Check6.Value = 2
    Screen.MousePointer = 11
    '
    For I& = 1 To NMS&
        '
        CIIX& = I&
        Call TRACE(20, CIIX&, NMS&)
        If LOREC% = 208 Then
              Get #N13%, I&, PDS$
              PDX$ = PDS$
            Else
              Get #N13%, I&, PDM$
              PDX$ = PDM$
        End If
        '
        For J% = 1 To CACOL%
            CAFAB = CVS(Mid$(PDX$, 4 * J% - 3, 4))
            If CAFAB > 0 Then
                REFE$ = "O/F a emitir": TIPO% = 12
                If DPI%(I&) = 0 Then REFE$ = "Tentativo de Compras": TIPO% = 11
                UII% = I&
                Call COBERFA(INIPER%(J%), UII%, REFE$, TIPO%, CAFAB)
            End If
        Next J%
    Next I&
    '
    For I& = 1 To X3&
        CIIX& = I&
        Call TRACE(20, CIIX&, X3&)
        Get #N9%, I&, B00$
        B1$ = Left$(B00$, 2)
        B3$ = Mid$(B00$, 3, 4)
        B4$ = Mid$(B00$, 7, 2)
        B5$ = Mid$(B00$, 9, 29)
        B7$ = Mid$(B00$, 42, 4)
        B8$ = Mid$(B00$, 46, 2)
        '
        CI% = CVI(B1$)
        If STATUIT%(CI%) >= 0 Then
          CANEC = (-1) * CVS(B3$)
          FFX% = CVI(B4$)
          REFE$ = B5$
          TIPO% = 22
          If Left$(REFE$, 20) <> "PROGRAMA POR MINIMOS" Then
            Call COBERFA(FFX%, CI%, REFE$, TIPO%, CANEC)
          End If
        End If
    Next I&
    '
    FORCANEC.Check6.Value = 1
    '
3900 Call CIERRARCH("*.*")
     If N4% > 0 Then Close #N4%: N4% = 0
     If N5% > 0 Then Close #N5%: N5% = 0
     If N7% > 0 Then Close #N7%: N7% = 0
     If N9% > 0 Then Close #N9%: N9% = 0
     If N13% > 0 Then Close #N13%: N13% = 0
     '
     'Call BLANARCH("PERPROMA")
     KKLL& = 0
     For KKK% = 1 To 52
       If TRIM$(DEPER$(KKK%)) <> "" Then
         KKLL& = KKLL& + 1
         Call GRAREG("PERPROMA", DEPER$(KKK%) + Space$(16), KKLL&)
       End If
     Next KKK%
     Call SETULTREG("PERPROMA", KKLL&)
     Call CIERRARCH("PERPROMA")
     '
     For UI& = 0 To 32767
       PROVINT%(UI&) = 0
     Next UI&
     '
     For U& = 1 To ULTREG&("TAPROINT")
       X$ = REGLEIDO$("TAPROINT", U&)
       NP% = CVI(Left$(X$, 2))
       If NP% > 0 Then
         PROVINT%(NP%) = 1
       End If
     Next U&
     Call CIERRARCH("TAPROINT")
     '
     JC& = 0    'Call BLANARCH("PLACOMP")
     JF& = 0    'Call BLANARCH("PLAFABR")
     'Call BLANARCH("ORFAEMI")
     'Call BLANARCH("OCOMEMI")
     '
     FORCANEC.Check7.Value = 2
     Screen.MousePointer = 11
     '
     FIN& = ULTREG&("ARTIPRO")
     For III& = 1 To FIN&
        Call TRACE(20, III&, FIN&)
        X$ = REGLEIDO$("ARTIPRO", III&)
        CI% = CVI(Left$(X$, 2))
        If CI% > 0 Then
            If CI% <= NUMAS% Then
                PROVEX$ = Mid$(X$, 3, 2)
                If ECOARCH$("PROVED1", PROVEX$) <> "" Then
                    PROVARTI%(CI%) = CVI(PROVEX$)
                End If
            End If
        End If
     Next III&
     '
     URECO& = CACO%
     FIN& = URECO&
     For III& = 1 To FIN&
        '
        Call TRACE(20, III&, FIN&)
        X$ = REGLEIDO$("COBERFA", III&)
        TIMO% = CVI(Mid$(X$, 51, 2))
        CI% = CVI(Mid$(X$, 3, 2))
        FF% = CVI(Mid$(X$, 1, 2))
        PERIO$ = Mid$(X$, 5, 16)
        CANTX = CVS(Mid$(X$, 53, 4))
        ' FFI% = CVI(Mid$(X$, 61, 2))
        '
        If TIMO% = 12 Then
          PHB% = PROVHABI%(CI%) ' PROVEEDOR HABITUAL
          If PHB% > 0 Then
            If PROVINT%(PHB%) = 0 Then ' NO ES PROVEEDOR INTERNO
              TIMO% = 11
            End If
          End If
        End If
        '
        Call REPLA(X$, MKI$(TIMO%), 51, 2)
        Call REPLA(X$, MKS$(0), 61, 4)
        Call GRAREG("COBERFA", X$, III&)
        '
        If TIMO% = 11 Then
             '
             MONEI% = MONECOSTO%(CI%)
             COSUDOL = COSUNI(CI%)
             COSUN = COSUDOL * TICAMONE(MONEI%)
             PHB% = PROVHABI%(CI%) ' PROVEEDOR HABITUAL
             If PHB% > 0 Then
               If PROVINT%(PHB%) <> 0 Then ' ES PROVEEDOR INTERNO
                 COSUN = 0
               End If
             End If
             '
             Z$ = REGBLAN$("PLACOMP")
             Call REPLA(Z$, MKI$(CI%), 1, 2)
             Call REPLA(Z$, MKI$(FF%), 3, 2)
             Call REPLA(Z$, PERIO$, 5, 16)
             Call REPLA(Z$, MKS$(CANTX), 21, 4)
             Call REPLA(Z$, MKS$(0), 25, 4) ' tot.progra.por item
             Call REPLA(Z$, MKS$(0), 29, 4) ' filler
             Call REPLA(Z$, MKI$(PHB%), 33, 2)
             Call REPLA(Z$, MKS$(COSUN), 35, 4)
             Call REPLA(Z$, MKD$(COSUN * CANTX), 39, 8)
             '
             JC& = JC& + 1
             Call GRAREG("PLACOMP", Z$, JC&)
             '
             FF1$ = REGBLAN$("OCOMEMI")
             Call REPLA(FF1$, MKI$(CI%), 1, 2)
             Call REPLA(FF1$, DESCRIT$(CI%), 3, 26)
             Call REPLA(FF1$, CSTRING$(MKS$(CANTX), 4, 10, 1, 2), 29, 10)
             FEX = FF%
             Call REPLA(FF1$, FECHATEX$(FEX), 41, 8)
             Call REPLA(FF1$, MKD$(CDbl(CANTX)), 49, 8)
             Call REPLA(FF1$, MKI$(FF%), 57, 2)
             Call REPLA(FF1$, MKI$(CI%), 59, 2)
             Call GRAREG("OCOMEMI", FF1$, JC&)
             '
           ElseIf TIMO% = 12 Then
             '
             Z$ = REGBLAN$("PLAFABR")
             Call REPLA(Z$, MKI$(CI%), 1, 2)
             Call REPLA(Z$, MKI$(FF%), 3, 2)
             Call REPLA(Z$, PERIO$, 5, 16)
             Call REPLA(Z$, MKS$(CANTX), 21, 4)
             Call REPLA(Z$, MKS$(0), 29, 4)  ' FILLER
             JF& = JF& + 1
             Call GRAREG("PLAFABR", Z$, JF&)
             '
             FF1$ = REGBLAN$("ORFAEMI")
             Call REPLA(FF1$, MKI$(CI%), 1, 2)
             Call REPLA(FF1$, DESCRIT$(CI%), 3, 26)
             Call REPLA(FF1$, CSTRING$(MKS$(CANTX), 4, 10, 1, 2), 29, 10)
             FEX = FF%
             Call REPLA(FF1$, FECHATEX$(FEX), 41, 8)
             Call REPLA(FF1$, MKD$(CDbl(CANTX)), 49, 8)
             Call REPLA(FF1$, MKI$(FF%), 57, 2)
             Call REPLA(FF1$, MKI$(CI%), 59, 2)
             Call GRAREG("ORFAEMI", FF1$, JF&)
             '
        End If
        '
     Next III&
     FORCANEC.Check7.Value = 1
     '
     Screen.MousePointer = 1
     URECO& = CACO%
     Call SETULTREG("COBERFA", URECO&)
     Call CIERRARCH("COBERFA")
     Call SETULTREG("ORFAEMI", JF&)
     Call CIERRARCH("ORFAEMI")
     Call SETULTREG("OCOMEMI", JC&)
     Call CIERRARCH("OCOMEMI")
     Call SETULTREG("PLACOMP", JC&)
     Call CIERRARCH("PLACOMP")
     Call SETULTREG("PLAFABR", JF&)
     Call CIERRARCH("PLAFABR")
     Call CIERRARCH("PERPROMA")
     '
     FECHOR$ = HOS$ + "  " + Left$(Time$, 5)
     Call GRACONTROL("CANECOM", "FECHOCAL", FECHOR$)
     If UCase$(Left$(App.EXEName, 5)) = "FLEXO" Then
         PROMAT00.Label7.Caption = Left$(FECHOR$, 8): PROMAT00.Label7.Refresh
         PROMAT00.Label95.Caption = Mid$(FECHOR$, 11, 5): PROMAT00.Label95.Refresh
       Else
         PROMATER.Label7.Caption = Left$(FECHOR$, 8): PROMATER.Label7.Refresh
         PROMATER.Label95.Caption = Mid$(FECHOR$, 11, 5): PROMATER.Label95.Refresh
     End If
     Call CIERRARCH("*.*")
     Call BAJALDISCO
     '
     'Call FINAL(SALI$)
     Call FILESORT("COBERFA", "", 1, 2, "ASC")
     Call FILESORT("PLACOMP", "", 1, 3, "ASC")
     Call FILESORT("PLAFABR", "", 1, 3, "ASC")
     Call FILESORT("ORFAEMI", "", 4, 5, "ASC")
     Call FILESORT("OCOMEMI", "", 4, 5, "ASC")
     '
     AA1% = COPYFILE%(LUDAT$ + "PLACOMP.DAT", LUDAT$ + "PLACOMQ.DAT")
     Call FILESORT("PLACOMQ", "", 3, 3, "ASC")
     '
     FINL& = ULTREG&("OCOMDETA")
     For IL& = 1 To FINL&
        '
        Call TRACE(20, IL&, FINL&)
        F$ = REGLEIDO$("OCOMDETA", IL&)
        CIBX% = CVI(Mid$(F$, 9, 2))
        STATU% = Asc(Mid$(F$, 124, 1))
        If STATU% < 9 Then      ' no esta anulada
          NUORSE& = CVS(Mid$(F$, 117, 4))
          If NUORSE& < 1 Then
            COEFCAN = CVS(Mid$(F$, 63, 4))
            If COEFCAN < 0.005 Then COEFCAN = 1
            CATOT# = CVS(Mid$(F$, 55, 4)) * COEFCAN
            CAPRA# = CVS(Mid$(F$, 99, 4)) * COEFCAN
            CAPEN = CATOT# - CAPRA#
            If CAPEN > 0 Then
              If STATUIT%(CIBX%) >= 0 Then
                CACOPEN(CIBX%) = CACOPEN(CIBX%) + CAPEN
              End If
            End If
          End If
        End If
     Next IL&
     '
     CIAX% = 0: JJ& = 0
     For U& = 1 To ULTREG&("PLACOMQ")
       XX$ = REGLEIDO$("PLACOMQ", U&)
       CII% = CVI(Left$(XX$, 2))
       If CII% <> CIAX% Then
         CIAX% = CII%
         OCEMI = CVS(Mid$(XX$, 21, 4))
         STAKK = CACOPEN(CII%)
         If STAKK > OCEMI Then STAKK = OCEMI
         Call REPLA(XX$, MKS$(STAKK), 25, 4)
         Call REPLA(XX$, MKS$(OCEMI - STAKK), 29, 4)
         Call REPLA(XX$, MKD$(CDbl(CVS(Mid$(XX$, 35, 4)) * (OCEMI - STAKK))), 39, 8)
         Call REPLA(XX$, MKS$(STOMIN(CII%)), 55, 4)
         Call REPLA(XX$, MKS$(LOTESTAN(CII%)), 59, 4)
         Call GRAREG("PLACOMQ", XX$, U&)
         CACOPEN(CII%) = CACOPEN(CII%) - STAKK
       End If
     Next U&
     Call CIERRARCH("PLACOMQ")
     Call FILESORT("PLACOMQ", "", 1, 3, "ASC")
     Call CARPROSEL
     '
     CIAX% = 0
     For U& = 1 To ULTREG&("PLAFABR")
       XX$ = REGLEIDO$("PLAFABR", U&)
       CII% = CVI(Left$(XX$, 2))
       If CII% <> CIAX% Then
         CIAX% = CII%
         STAKK = STOCKACT(CII%): If STAKK < 0 Then STAKK = 0
         Call REPLA(XX$, MKS$(STAKK), 29, 4)
         Call REPLA(XX$, MKS$(STOMIN(CII%)), 55, 4)
         Call REPLA(XX$, MKS$(LOTESTAN(CII%)), 59, 4)
         Call GRAREG("PLAFABR", XX$, U&)
       End If
     Next U&
     Call CIERRARCH("PLAFABR")
     '
999  Call CIERRARCH("*.*")
     Call BAJALDISCO
     Call COMUNI("Re-Cálculo Completo.")
     '
     FORCANEC.Hide
     Exit Sub
     '
' *********************************************************
     '
5300 If NESTRU& > 0 Then GoTo 5380
     '
     PUNTCI$ = "": PUNTRG$ = ""
     NESTRU& = 0
     FIN& = ULTREG&("ESTRUNUE")
     For I& = 1 To FIN&
        RA$ = REGLEIDO$("ESTRUNUE", I&)
        RB$ = REGLEIDO$("ESTRUDOS", I&)
        CI% = CVI(Left$(RA$, 2))
        If CI% > 0 Then
           If CI% <= NUMAS% Then
              If DPI%(CI%) = 0 Then
                 DPI%(CI%) = 1
                 NESTRU& = NESTRU& + 1
                 Call GRAREG("ESTRUNUE", RA$, NESTRU&)
                 Call GRAREG("ESTRUDOS", RB$, NESTRU&)
                 PUNTCI$ = PUNTCI$ + MKI$(CI%)
                 NESTRI% = NESTRU&
                 PUNTRG$ = PUNTRG$ + MKI$(NESTRI%)
              End If
           End If
        End If
     Next I&
     '
     For I& = 1 + ULTREG&("ESTRUDOS") To NESTRU&
        RB$ = String$(256, 0)
        Call REPLA(RB$, MKS$(0.0001), 251, 4)
        Call GRAREG("ESTRUDOS", RB$, I&)
     Next I&
     '
     Call SETULTREG("ESTRUNUE", NESTRU&)
     Call SETULTREG("ESTRUDOS", NESTRU&)
     Call CIERRARCH("ESTRUNUE")
     Call CIERRARCH("ESTRUDOS")
     '
     For K% = 1 To NUMAS%: DPI%(K%) = 0: Next K%
     '
5380 Return
     '
17000 ENCU% = 0: PX1% = 1

17005 PX% = InStr(PX1%, PUNTCI$, MKI$(CIBAS%))
      If PX% < PX1% Then GoTo 17035
      If (PX% Mod 2) <> 1 Then PX1% = PX% + 1: GoTo 17005
      L% = CVI(Mid$(PUNTRG$, PX%, 2))
      ENCU% = L%
      '
      Get #N5%, L%, RA$
      R1$ = Left$(RA$, 2)
      R2$ = Mid$(RA$, 3, 120)
      R3$ = Mid$(RA$, 123, 6)
      '
17035 Return
17040 '

17200 Get #N5%, ENCU%, RA$
      R1$ = Left$(RA$, 2)
      R2$ = Mid$(RA$, 3, 120)
      R3$ = Mid$(RA$, 123, 6)
      DPS$ = R2$: J% = 0
      '
17205 For K% = 1 To 120 Step 8
17210    DPSX$ = Mid$(DPS$, K%, 8)
17215    If CVI(DPSX$) < 1 Or CVI(DPSX$) > NMS& Then GoTo 17280
         If CVS(Mid$(DPSX$, 3)) < 0.0005 Then GoTo 17225
17220    J% = J% + 1
         CII1%(J%) = CVI(DPSX$)
         CANT1(J%) = CABAS * CVS(Mid$(DPSX$, 3))
         PAD1%(J%) = Asc(Mid$(DPSX$, 7))
         If PAD1%(J%) < 1 Then
             PAD1%(J%) = PLADISTA%
         End If
17225 Next K%
      '
17230 Get #N7%, ENCU%, RB$
      R4$ = RB$
      DPS$ = R4$
17235 For K% = 1 To 256 Step 8
17240    DPSX$ = Mid$(DPS$, K%, 8)
17245    If CVI(DPSX$) < 1 Or CVI(DPSX$) > NMS& Then GoTo 17280
         If CVS(Mid$(DPSX$, 3)) < 0.0005 Then GoTo 17255
17250    J% = J% + 1
         CII1%(J%) = CVI(DPSX$)
         CANT1(J%) = CABAS * CVS(Mid$(DPSX$, 3))
         PAD1%(J%) = Asc(Mid$(DPSX$, 7))
         If PAD1%(J%) < 1 Then
             PAD1%(J%) = PLADISTA%
         End If
17255 Next K%
17260 '
17280 Return
17285 '

interrum:
      'RESP% = COMALTER%("Proceso Interrumpido a su Comando\\Reanudar\Abortar")
      'If RESP% <> 1 Then Call FINAL("")
      Return

End Sub
'
Sub CARPROSEL()
   '
   Call COPYSTRU("PROVED1", "PROVEPLA")
   ASELPRO$ = "PROVEPLA"
   Screen.MousePointer = 11
   PPXXZ$ = String$(32767, 0)
   For U& = 1 To ULTREG("PLACOMP")
     XX$ = REGLEIDO$("PLACOMP", U&)
     NPRX% = CVI(Mid$(XX$, 33, 2))
     If NPRX% > 0 Then
       Mid$(PPXXZ$, NPRX%, 1) = Chr$(1)
     End If
   Next U&
   '
   JJ& = 0
   For U& = 1 To ULTREG&("ACREDOR")
     XX$ = REGLEIDO$("ACREDOR", U&)
     NPRX% = CVI(Left$(XX$, 2))
     If NPRX% > 0 Then
       If Asc(Mid$(PPXXZ$, NPRX%, 1)) > 0 Then
         JJ& = JJ& + 1
         Call GRAREG("PROVEPLA", XX$, JJ&)
       End If
     End If
   Next U&
   '
   Call SETULTREG("PROVEPLA", JJ&)
   Call CIERRARCH("PROVEPLA")
   '
   Screen.MousePointer = 1
   '
End Sub


