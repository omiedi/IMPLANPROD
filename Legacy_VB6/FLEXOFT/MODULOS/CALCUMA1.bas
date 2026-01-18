Attribute VB_Name = "CALCUMA"
Dim HAYFAB%(32767), NIVEFAB%(32767)
Dim FEFINANT%(10), HOFINANT%(10)
'
Dim ORFA$(2048)
Dim ORFB$(2048)
Dim OPRIN$(2048)
Dim ORFASEQ$(4096), STATU%(4096), PRIVI%(4096)
Dim REGO&(2048)
Dim REGOF&(1024)
Dim PRIO%(1024)
Dim FENEC%(1024)
Dim DNEC%(1024)
Dim HNEC%(1024)
Dim MQNEC$(1024)
Dim MQBAS$(1024)
Dim RGOPX&(1024)
Dim REGOPE&(256)
'
Dim HODE%(64), HOHA%(64), OPEDIS(64)
Dim EQ$(10)
'
Dim MAQUI$(512), MAEQUI$(512), COEFTIE(512), TIEPRO(512), VECHORES$(512)
Dim VECTOMAQ As String
Dim NUMAQ As Integer
Dim MQQ%(8192, 3)
Dim MAQUISE As String
'
Dim EQSEL$(512), PUNOP$(512)
'
Dim CAOR As Integer
Dim CADIA As Integer
Dim TXT As String * 255

Public COMPLECALCU%
Dim IFORCA%
'
Public XMUOFA, YMUOFA
'
Sub CALCUMAQ()   ' recalculo de carga de maquinas segun
                 ' ordenes de fabricacion en curso
    For UFORMU% = 1 To Forms.Count
      FORNA$ = TRIM$(UCase$(Forms(UFORMU% - 1).Name))
      If FORNA$ = "FORCAMAQ" Or FORNA$ = "CALCUCAR" Then
        On Error Resume Next
        FECOMPRO% = FECHANUM(Forms(UFORMU% - 1).Text1(9).TEXT)
        If Err < 1 Then GoTo 5
      End If
    Next UFORMU%
    Call MENSERR(24, "Proceso Imposible")
    Exit Sub
        
5   IFORCA% = UFORMU% - 1
    On Error GoTo 0
    EPAC$ = EMPREAC$
    FCAL% = HOY(HO$)
    FECOMPRO% = FECHANUM(Forms(IFORCA%).Text1(9).TEXT)
    If FECOMPRO% < FCAL% Then
        Call MENSERR(24, "Fecha no válida")
        Forms(IFORCA%).Command2.Enabled = True
        Forms(IFORCA%).Text1(9).TEXT = HO$
        Forms(IFORCA%).Text1(9).SetFocus
        Exit Sub
    End If
    '
    If Forms(IFORCA%).Option1.Value = True Then
        MODOCA% = 1
      ElseIf Forms(IFORCA%).Option2.Value = True Then
        MODOCA% = 2
      Else
        Call MENSERR(24, "Falta Seleccionar Modo de Ordenamiento")
        Forms(IFORCA%).Command2.Enabled = True
        Forms(IFORCA%).Option2.SetFocus
        Exit Sub
    End If
    '
    INCLUSUS% = 0
    If Forms(IFORCA%).Check7.Value = 1 Then
        INCLUSUS% = 1
    End If
    '
    HOCOMPRO% = HORANUM(Forms(IFORCA%).Text1(10).TEXT)
    '
    CAORFA% = 2048
    CAPRO% = 128
    CAOPSELMAX% = 256
    '
10  NUMAQ% = ULTREG&("PADMAQ")
    If NUMAQ% > 512 Then
        Call COMUNI("Demasiadas Máquinas Definidas.\Consulte al Soporte de Sistemas.")
        Forms(IFORCA%).Hide
        Exit Sub
    End If
    '
    VECTOMAQ$ = ""
    For U& = 1 To NUMAQ%
       REGMA$ = REGLEIDO$("PADMAQ", U&)
       MAQUI$(U&) = Left$(REGMA$, 6)
       VECTOMAQ$ = VECTOMAQ$ + "@@" + Left$(REGMA$, 6)
       MAEQUI$(U&) = Mid$(REGMA$, 43, 6)
       COEFI = CVI(Mid$(REGMA$, 49, 2))
       If COEFI < 1 Then COEFI = 100
       COEFTIE(U&) = COEFI
       TIEPRO(U&) = 0
    Next U&
    '

    Call BLANARCH("OPERFAP")
    '
    'CARGA VECHORES$ CON HORARIOS RESERVADOS POR PARADAS PROGRAMADAS
    For U& = 1 To ULTREG("PARAPROG")
      XXX$ = REGLEIDO$("PARAPROG", U&)
      MQX% = MAQNU%(Left$(XXX$, 6))
      FECOM% = CVI(Mid$(XXX$, 7, 2))
      HOCOM% = CVI(Mid$(XXX$, 9, 2))
      FEFIN% = CVI(Mid$(XXX$, 11, 2))
      HOFIN% = CVI(Mid$(XXX$, 13, 2))
      STH$ = MKI$(FECOM%) + MKI$(HOCOM%) + MKI$(FEFIN%) + MKI$(HOFIN%) + MKI$(FECOM%) + MKI$(HOCOM%)
      VECHORES$(MQX%) = VECHORES$(MQX%) + STH$
    Next U&
    '
    '
    ' ******************************* Seleccionando Ordenes Lanzadas
    '
    Forms(IFORCA%).Check1.Value = 2
    '
    CAOR% = 0
    FIN& = ULTREG&("ORDEFABR")
    If FIN& < 1 Then GoTo 99    ' no hay ordenes de fabricacion
    '
    'FORMS(IFORCA%).ProgressBar1.Min = 0
    'FORMS(IFORCA%).ProgressBar1.Max = FIN&
    '
    If MODOCA% = 2 Then
        FIN& = ULTREG&("SECUORF")
        For I& = 1 To FIN&
            X$ = REGLEIDO$("SECUORF", I&)
            ORFASEQ$(I&) = Left$(X$, 12)
        Next I&
        CAORFSEQ% = FIN&
    End If
    '
    Call COPYSTRU("ECORDEF", "ECORDEP")
    J& = 0
    STATUMI% = 1: STATUMA% = 1
    If INCLUSUS% = 1 Then STATUMA% = 2
    '
    FIN& = ULTREG&("ORDEFABR")
    For I& = 1 To FIN&
        '
        'FORMS(IFORCA%).ProgressBar1.Value = I&
        Call TRACE(20, I&, FIN&)
        F$ = REGLEIDO$("ORDEFABR", I&)
        '
        STATUOF% = Asc(Mid$(F$, 108, 1))
        If STATUOF% >= STATUMI% And STATUOF% <= STATUMA% Then ' esta lanzada
            PRIORI% = Asc(Mid$(F$, 107, 1))
            If PRIORI% < 1 Or PRIORI% > 9 Then PRIORI% = 10
            PRIORI% = 10 + PRIORI%
            '
            If MODOCA% = 2 Then   ' prioridad por secuencia
                ORFASE$ = Mid$(F$, 69, 12)
                For JI% = 1 To CAORFSEQ%
                   If ORFASE$ = ORFASEQ$(JI%) Then
                      PRIORI% = JI%
                      GoTo 49
                   End If
                Next JI%
                PRIORI% = PRIORI% + CAORFSEQ% + 100 * STATUOF%
            End If
            '
49          TXTX$ = Mid$(F$, 69, 12) + AJUSTI$(Str$(PRIORI%), 6) + Str$(I&)
            Forms(IFORCA%).LIORFALA.AddItem TXTX$
            CIOFXX% = CVI(Mid$(F$, 83, 2))
            If CIOFXX% > 0 Then
               If CIOFXX% <= NUMAS% Then
                  HAYFAB%(CIOFXX%) = 1
               End If
            End If
            '
            X$ = REGLEIDO$("ECORDEF", I&)
            J& = J& + 1
            Call GRAREG("ECORDEP", X$, J&)
            '
        End If
    Next I&
    '
    For I& = J& + 1 To ULTREG("ECORDEP")
       Call GRAREG("ECORDEP", String$(64, 0), I&)
    Next I&
    Call CIERRARCH("ECORDEP")
    '
    CAOR% = Forms(IFORCA%).LIORFALA.ListCount
    If CAOR% > CAORFA% Then
       Call COMUNI("Proceso Imposible.\Demasiadas Ordenes de Fabricación Lanzadas.")
       Forms(IFORCA%).Hide
       Exit Sub
    End If
    '
99  If CAOR% < 1 Then
       Call COMUNI("Proceso Imposible.\No hay Ordenes de Fabricacion\Lanzadas y Pendientes.")
       Forms(IFORCA%).Hide
       Exit Sub
    End If
    '
    For I& = 1 To CAOR%
       TXTX$ = Forms(IFORCA%).LIORFALA.List(I& - 1)
       ORFA$(I&) = Left$(TXTX$, 12)
       PRIO%(I&) = Val(Mid$(TXTX$, 13, 6))
       REGO&(I&) = Val(Mid$(TXTX$, 19))
    Next I&
    
    ' ************************ hasta aquí se cargó un vector con las
    '                          ordenes de fabricacion lanzadas
    '
    'DETERMINANDO NIVEL DE CADA ARTICULO
    vuelta% = 1
    '
20  RECA% = 0
    FIN& = ULTREG&("ESTRUNUE")
    Screen.MousePointer = 11
    For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       R0$ = Left$(REGLEIDO$("ESTRUNUE", I&), 122)
       If I& <= ULTREG&("ESTRUDOS") Then
            R0$ = R0$ + REGLEIDO$("ESTRUDOS", I&)
          Else
            R0$ = R0$ + String$(256, 0)
       End If
       CICON% = CVI(Left$(R0$, 2))
       If CICON% > 0 Then
         If CICON% <= NUMAS% Then
           If HAYFAB%(CICON%) > 0 Then
             NIVMAX% = 0
             DPS$ = Mid$(R0$, 3)
             For KQQ% = 1 To Len(DPS$) Step 8
               DPSX$ = Mid$(DPS$, KQQ%, 8)
               CIPAR% = CVI(Left$(DPSX$, 2))
               If CIPAR% > 0 Then
                 If CIPAR% <= NUMAS% Then
                   If CIPAR% <> CICON% Then
                     CAPAR = CVS(Mid$(DPSX$, 3, 4))
                     If CAPAR >= 0.00005 Then
                        If NIVMAX% < NIVEFAB%(CIPAR%) Then
                           NIVMAX% = NIVEFAB%(CIPAR%)
                        End If
                     End If
                   End If
                 End If
               End If
             Next KQQ%
             '
             If NIVEFAB%(CICON%) <= NIVMAX% Then
               NIVEFAB%(CICON%) = NIVMAX% + 1
               RECA% = 1
             End If
             '
             For KQQ% = 1 To Len(DPS$) Step 8
               DPSX$ = Mid$(DPS$, KQQ%, 8)
               CIPAR% = CVI(Left$(DPSX$, 2))
               If CIPAR% > 0 Then
                 If CIPAR% <= NUMAS% Then
                   If CIPAR% <> CICON% Then
                     CAPAR = CVS(Mid$(DPSX$, 3, 4))
                     If CAPAR <= (-0.00005) Then
                        If NIVEFAB%(CIPAR%) <= NIVMAX% Then
                           NIVEFAB%(CIPAR%) = NIVMAX% + 1
                           RECA% = 1
                        End If
                     End If
                   End If
                 End If
               End If
             Next KQQ%
             '
           End If
         End If
       End If
    Next I&
    '
    If RECA% = 1 Then
       If vuelta% < 9 Then
          vuelta% = vuelta% + 1
          GoTo 20
       End If
    End If
    '
    UOFAX& = ULTREG&("ORDEFABR")
    For I& = 1 To CAOR%
       ORFA$(I&) = AJUSTI$(ORFA$(I&), 12) + MKI$(0)
       REOFAX& = REGO&(I&)
       If REOFAX& > 0 Then
          If REOFAX& <= UOFAX& Then
             F$ = REGLEIDO$("ORDEFABR", REOFAX&)
             CIOFXX% = CVI(Mid$(F$, 83, 2))
             ORFA$(I&) = Left$(ORFA$(I&), 12) + MKI$(NIVEFAB%(CIOFXX%))
          End If
       End If
    Next I&
    Forms(IFORCA%).Check1.Value = 1
    '
    ' ************************ lectura de operaciones pendientes
    '
    Forms(IFORCA%).Check2.Value = 2
    FIN& = ULTREG&("OPERFAB")
    If FIN& < 1 Then
       Call COMUNI("Proceso Imposible.\No hay Operaciones de\Fabricación Programadas.")
       Forms(IFORCA%).Hide
       Exit Sub
    End If
    '
    For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       '
       O$ = REGLEIDO$("OPERFAB", I&)
       If Mid$(O$, 177, 6) <> String$(6, 0) Then GoTo 109 ' cumplida
       NOFA$ = Left$(O$, 12)
       '
       LI% = 0: LS% = CAOR% + 1
101    E% = Int((LS% - LI%) / 2): If E% < 1 Then GoTo 109
       L% = LI% + E%: If L% < 1 Or L% > CAOR% Then GoTo 109
       If NOFA$ < Left$(ORFA$(L%), 12) Then
           LS% = L%: GoTo 101
         ElseIf NOFA$ > Left$(ORFA$(L%), 12) Then
           LI% = L%: GoTo 101
       End If
       '
       CAPROG# = CVD(Mid$(O$, 79, 8))
       CAREA# = CVD(Mid$(O$, 87, 8))
       CAPEN# = CAPROG# - CAREA#
       '
       If CAPEN# >= 0.05 Then
         If CAPROG# >= 0.05 Then
           COEF = CAPEN# / CAPROG#
           If COEF > 1 Then COEF = 1
           TPREP = CVS(Mid$(O$, 103, 4)): If CAREA# >= 0.05 Then TPREP = 0
           TFIJO = CVS(Mid$(O$, 107, 4))
           TVARI = COEF * CVS(Mid$(O$, 111, 4))
           TDEMO = COEF * CVS(Mid$(O$, 115, 4))
           TMANT = COEF * CVS(Mid$(O$, 119, 4))
           TLIMP = CVS(Mid$(O$, 123, 4))
           TOTRO = COEF * CVS(Mid$(O$, 127, 4))
           TTOTA = TPREP + TFIJO + TVARI + TDEMO + TMANT + TLIMP + TOTRO
           '
           If TTOTA > 3000 Then
              If InStr(EPAC$, "NOVOES") > 0 Then
                 OPXX% = COMALTER%("Tiempos Aparentemente Incorrectos\en O.F. Número " + TRIM$(NOFA$) + "\\Refrescar\Continuar")
                 If OPXX% = 1 Then
                    CIIX% = CVI(Mid$(O$, 13, 2))
                    Call FRESHOPER(CIIX%)
                    GoTo 10
                 End If
              End If
           End If
           '
           DOPEXI$ = TRIM$(Str$(CVI(Mid$(O$, 15, 2)))) + "." + Mid$(O$, 17, 48)
           Call REPLA(O$, DOPEXI$, 17, 48)
           Call REPLA(O$, MKS$(TPREP), 103, 4)
           Call REPLA(O$, MKS$(TFIJO), 107, 4)
           Call REPLA(O$, MKS$(TVARI), 111, 4)
           Call REPLA(O$, MKS$(TDEMO), 115, 4)
           Call REPLA(O$, MKS$(TMANT), 119, 4)
           Call REPLA(O$, MKS$(TLIMP), 123, 4)
           Call REPLA(O$, MKS$(TOTRO), 127, 4)
           Call REPLA(O$, MKS$(TTOTA), 131, 4)
           Call REPLA(O$, String$(8, 0), 139, 8)
           Call REPLA(O$, MKS$(I&), 221, 4)
           '
           Call REGAPP("OPERFAP", O$)
           J& = ULTREG&("OPERFAP")
NUPENU% = CVI(Mid$(O$, 15, 2))
ORFANE$ = Left$(ORFA$(L%), 14)
PUNTEX$ = Mid$(ORFA$(L%), 15)
For KIKI% = 1 To Len(PUNTEX$) Step 4
  REJO& = CVS(Mid$(PUNTEX$, KIKI%, 4))
  OKX$ = REGLEIDO$("OPERFAP", REJO&)
  NUPEVI% = CVI(Mid$(OKX$, 15, 2))
  If NUPEVI% >= NUPENU% Then
    PUNTEX$ = Left$(PUNTEX$, KIKI% - 1) + MKS$(J&) + Mid$(PUNTEX$, KIKI%)
    GoTo 102
  End If
Next KIKI%
PUNTEX$ = PUNTEX$ + MKS$(J&)
102 ORFA$(L%) = ORFANE$ + PUNTEX$
O$ = REGLEIDO$("OPERFAP", J&)
           'ORFA$(L%) = ORFA$(L%) + MKS$(J&)
           If CAREA# > 0 Then
              If MODOCA% = 1 Then
                 If PRIO%(L%) > 10 Then
                     PRIO%(L%) = PRIO%(L%) - 10
                 End If
               ElseIf MODOCA% = 2 Then
                 If PRIO%(L%) > 0 Then
                     PRIO%(L%) = PRIO%(L%) - 32767
                 End If
               End If
           End If
           '
         End If
       End If
       '
109 Next I&
    '
    Forms(IFORCA%).Check2.Value = 1
    '
    ' *************************** Ordenando O/F's por Prioridad
    '
    Forms(IFORCA%).Check3.Value = 2
    '
    FIN& = CAOR%
    'FORMS(IFORCA%).ProgressBar1.Min = 0
    'FORMS(IFORCA%).ProgressBar1.Max = FIN&
    '
    For II% = 1 To CAOR%
        'FORMS(IFORCA%).ProgressBar1.Value = II%
        I& = II%
        Call TRACE(20, I&, FIN&)
        For JJ% = II% + 1 To CAOR%
            If PRIO%(II%) > PRIO%(JJ%) Then
                ORFAX$ = ORFA$(II%)
                REGOX& = REGO&(II%)
                PRIOX% = PRIO%(II%)
                ORFA$(II%) = ORFA$(JJ%)
                REGO&(II%) = REGO&(JJ%)
                PRIO%(II%) = PRIO%(JJ%)
                ORFA$(JJ%) = ORFAX$
                REGO&(JJ%) = REGOX&
                PRIO%(JJ%) = PRIOX%
            End If
        Next JJ%
    Next II%
    '
    Forms(IFORCA%).Check3.Value = 1
    '
    ' ************************** hasta aquí se ordenó el vector
    '
    '
    ' ********************************* Calculando Tiempos Standard
    '
400 Forms(IFORCA%).Check4.Value = 2
    '
    FIN& = ULTREG&("OPERFAP")
    FIN1& = FIN&: If FIN1& < 1 Then FIN1& = 1
    'FORMS(IFORCA%).ProgressBar1.Min = 0
    'FORMS(IFORCA%).ProgressBar1.Max = FIN1&
    '
    For I& = 1 To FIN&
        '
        'FORMS(IFORCA%).ProgressBar1.Value = I&
        Call TRACE(20, I&, FIN&)
        O$ = REGLEIDO$("OPERFAP", I&)
        '
        If CVI(Mid$(O$, 135, 2)) < 1 Then
            O$ = REPLACE$(O$, MKI$(FECOMPRO%), 135, 2)
            O$ = REPLACE$(O$, MKI$(HOCOMPRO%), 137, 2)
        End If
        '
        Call GRAREG("OPERFAP", O$, I&)
    Next I&
    Call CIERRARCH("OPERFAP")
    '
    Forms(IFORCA%).Check4.Value = 1
    '
500 Call BLANARCH("CRONOPE")
    Call BLANARCH("GANTMAQ")
    Call BLANARCH("PERIPRO")
    '
    ' ***************************** Secuenciando Carga de Maquinas
    '
    Forms(IFORCA%).Check5.Value = 2
    FIN& = ULTREG&("OPERFAP")
    FIN1& = FIN&: If FIN1& < 1 Then FIN1& = 1
    '
    NUOTA = 0
    CAOPERF& = ULTREG&("OPERFAP")
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    For I& = 1 To FIN&
       '
       Call TRACE(20, I&, FIN&)
       O$ = REGLEIDO$("OPERFAP", I&)
       NUOTI = Val(Mid$(O$, 4, 9))
       If NUOTI <> NUOTA Then MAQUISE$ = ""
       CAREA# = CVD(Mid$(O$, 87, 8))
       '
       Call GRAREG("OPERFAP", O$, I&)
       TTOT = CVS(Mid$(O$, 131, 4))
       MAQUIS$ = Mid$(O$, 73, 6) + Mid$(O$, 203, 18)
       If InStr(EPAC$, "DOSIVA") > 0 Then
         MAQUIS$ = Left$(MAQUIS$, 6)
       End If
       '
       For IMQ% = 0 To 3
          MQQ%(I&, IMQ%) = MAQNU%(Mid$(MAQUIS$, 6 * IMQ% + 1, 6))
       Next IMQ%
       '
       If CAREA# < 0.1 Then
          '
          Call REPLAMAQ(MAQUIS$, TTOT, MAQUIAL$, TALT)
          '
          COEFTIPRO = 1: If TTOT > 0 Then COEFTIPRO = TALT / TTOT
          STTIE$ = ""
          For IMQ% = 103 To 134 Step 4
             STTIE$ = STTIE$ + MKS$(COEFTIPRO * CVS(Mid$(O$, IMQ%, 4)))
          Next IMQ%
          Call REPLA(O$, Left$(MAQUIAL$, 6), 73, 6)
          Call REPLA(O$, Mid$(MAQUIAL$, 7, 18), 203, 18)
          Call REPLA(O$, STTIE$, 103, 32)
          Call GRAREG("OPERFAP", O$, I&)
          '
       End If
       '
       For IMQ% = 0 To 3
          maquiba$ = Mid$(MAQUIAL$, 6 * IMQ% + 1, 6)
          If TRIM$(maquiba$) <> "" Then
             For JJ% = 1 To NUMAQ%
                If MAQUI$(JJ%) = maquiba$ Then
                    TIEPRO(JJ%) = TIEPRO(JJ%) + TALT
                End If
             Next JJ%
          End If
       Next IMQ%
       '
       NUOTA = NUOTI
       '
503 Next I&
    '
    Call CIERRARCH("OPERFAP")
    Forms(IFORCA%).Check5.Value = 1
    '
505 CAOPRIN% = 0
    For KK% = 1 To CAOR%
       NUOPRIN$ = Mid$(ORFA$(KK%), 4, 6)
       For KK3% = 1 To CAOPRIN%
          If Left$(OPRIN$(KK3%), 6) = NUOPRIN$ Then GoTo 506
       Next KK3%
       CAOPRIN% = CAOPRIN% + 1
       OPRIN$(CAOPRIN%) = NUOPRIN$
       KK3% = CAOPRIN%
506    OPRIN$(KK3%) = OPRIN$(KK3%) + MKI$(KK%)
    Next KK%
    '
    FIN1& = CAOR%
FIN7& = CAOPRIN%
JVP& = 0
'tini1 = Timer
'sumatie = 0
For kk7% = 1 To CAOPRIN%
KK77& = kk7%
Call TRACE(20, KK77&, FIN7&)
'If kk7% = 6 Then
'   MsgBox Timer - tini1 & "   " & sumatie
'   Stop
'End If
For NIVEFA% = 0 To 9
FEFINANT%(NIVEFA%) = 0: HOFINANT%(NIVEFA%) = 0
Next NIVEFA%
For NIVEFA% = 0 To 9
If NIVEFA% > 0 Then
   AAA% = FEFINANT%(NIVEFA% - 1)
   BBB% = HOFINANT%(NIVEFA% - 1)
   If FEFINANT%(NIVEFA%) < AAA% Or (FEFINANT%(NIVEFA%) = AAA% And HOFINANT%(NIVEFA%) < BBB%) Then
      FEFINANT%(NIVEFA%) = AAA%
      HOFINANT%(NIVEFA%) = BBB%
   End If
End If
For kk8% = 7 To Len(OPRIN$(kk7%)) Step 2
KK% = CVI(Mid$(OPRIN$(kk7%), kk8%, 2))
'nvf1% = CVI(Mid$(ORFA$(KK%), 13, 2))
If CVI(Mid$(ORFA$(KK%), 13, 2)) = NIVEFA% Then
    'For KK% = 1 To CAOR%
        '
        'FORMS(IFORCA%).ProgressBar1.Value = KK%
        I& = KK%
        'Call TRACE(20, I&, FIN1&)
        '
        CAOPSEL% = ((Len(ORFA$(KK%))) - 14) / 4
        If CAOPSEL% < 1 Then
            GoTo 590                            ' NEXT KK%
        End If
        '
        For KL% = 1 To CAOPSEL%
            '
            REGOPE&(KL%) = CVS(Mid$(ORFA$(KK%), 4 * KL% + 11, 4))
            If KL% = 1 Then
                O$ = REGLEIDO$("OPERFAP", REGOPE&(KL%))
                Forms(IFORCA%).Text1(0).TEXT = Left$(O$, 12)
                CIII% = CVI(Mid$(O$, 13, 2))
                CAPROGG = CVD(Mid$(O$, 79, 8))
                Forms(IFORCA%).Text1(1).TEXT = TRIM$(CODEXT$(CIII%)) + " - " + TRIM$(DESCRIT$(CIII%)) + " (" + TRIM$(Str$(CAPROGG)) + ")"
                '
                FEFINA% = CVI(Mid$(O$, 135, 2))
                HOFINA% = CVI(Mid$(O$, 137, 2))
                If FEFINA% < FECOMPRO% - 100 Then
                    FEFINA% = FECOMPRO%
                    HOFINA% = HOCOMPRO%
                End If
                FEX = FEFINA%
                Forms(IFORCA%).Text1(2).TEXT = FECHATEX$(FEX)
             End If
            '
        Next KL%
        '
'MsgBox Timer - TINXX
        FECOM% = FEFINA%: HOCOM% = HOFINA%
        If NIVEFA% > 0 Then
           If FECOM% < FEFINANT%(NIVEFA% - 1) Or (FECOM% = FEFINANT%(NIVEFA% - 1) And HOCOM% < HOFINANT%(NIVEFA% - 1)) Then
              FECOM% = FEFINANT%(NIVEFA% - 1): HOCOM% = HOFINANT%(NIVEFA% - 1)
           End If
        End If
        FEFINABAS% = FEFINA%: HOFINABAS% = HOFINA%
        KX% = 0
        REGOPE0& = 0
        '
510     For KY% = 1 To CAOPSEL%
            '
'MsgBox "1   " & (Timer - TINXX): TINXX = Timer
            TURNO% = 1
            '
511         REGOPEL& = REGOPE&(KY%)
            VECTIPAR$ = ""
            O$ = REGLEIDO$("OPERFAP", REGOPEL&)
            '
            Forms(IFORCA%).Text1(3).TEXT = TRIM$(CVI(Mid$(O$, 15, 2)))
            Forms(IFORCA%).Text1(4).TEXT = TRIM$(Mid$(O$, 17, 48))
            '
            TNEC = CVS(Mid$(O$, 131, 4)) / 60
            TNEC0 = TNEC
            MAQUIS$ = Mid$(O$, 73, 6) + Mid$(O$, 203, 18)
            '
'MsgBox "2   " & (Timer - TINXX): TINXX = Timer
            If FECOM% <= FEFINA% Then
                FECOM% = FEFINA%
                HOCOM% = HOFINA%
            End If
            '
            If 10000 * CDbl(FECOM%) + CDbl(HOCOM%) < 10000 * CDbl(CVI(Mid$(O$, 135, 2))) + CDbl(CVI(Mid$(O$, 137, 2))) Then
                FECOM% = CVI(Mid$(O$, 135, 2))
                HOCOM% = CVI(Mid$(O$, 137, 2))
            End If
            '
'MsgBox "3   " & (Timer - TINXX): TINXX = Timer
            If NIVEFA% > 0 Then
               If FECOM% < FEFINANT%(NIVEFA% - 1) Or (FECOM% = FEFINANT%(NIVEFA% - 1) And HOCOM% < HOFINANT%(NIVEFA% - 1)) Then
                  FECOM% = FEFINANT%(NIVEFA% - 1): HOCOM% = HOFINANT%(NIVEFA% - 1)
               End If
            End If
            '
'MsgBox "4   " & (Timer - TINXX): TINXX = Timer
512         HAYEQ% = 0
            For UI% = 0 To 3
                EQ$(UI%) = Mid$(MAQUIS$, 6 * UI% + 1, 6)
                If TRIM$(EQ$(UI%)) <> "" Then
                    HAYEQ% = 1
'tini2 = Timer
                    Call HODISTUR(EQ$(UI%), FECOM%, TURNO%, DEHOR%, HAHOR%)
'sumatie = sumatie + Timer - tini2
                    If DEHOR% > DHR0% Then DHR0% = DEHOR%
                    If HAHOR% < HHR0% Then HHR0% = HAHOR%
                End If
            Next UI%
           If HAYEQ% = 0 Then
                Call MENSERR(24, "Falta Definir Equipo para " + TRIM$(Left$(O$, 12)))
                'Call CIERRARCH("*.*")
                'GoTo 3800
                GoTo 524
            End If
            '
'MsgBox "5   " & (Timer - TINXX): TINXX = Timer
            If DEHOR% >= HAHOR% Then
                TURNO% = TURNO% + 1
                If TURNO% > 6 Then
                   TURNO% = 1
                   FEX = FECOM%
                   FECOM% = DIPOST(FEX)
                End If
                HOCOM% = 0
                GoTo 512
            End If
            '
            If HOCOM% < DEHOR% Then
                HOCOM% = DEHOR%
            End If
            '
            FEFIN% = FECOM%
            HOFIN% = HOCOM%
'MsgBox "D   " & TNEC0 & "   " & HOCOM% & "   " & HOFIN%
            FEX1 = FECOM%
            '
'MsgBox "6   " & (Timer - TINXX): TINXX = Timer
515         DHR0% = 0: HHR0% = 2400
            '
            For UI% = 0 To 3
                EQ$(UI%) = Mid$(MAQUIS$, 6 * UI% + 1, 6)
                If TRIM$(EQ$(UI%)) <> "" Then
tini2 = Timer
                Call HODISTUR(EQ$(UI%), FEFIN%, TURNO%, DEHOR%, HAHOR%)
sumatie = sumatie + Timer - tini2
'MsgBox Timer - TINXX
                    If FEFIN% = FECOM% Then
                        If DEHOR% < HOCOM% Then
                            DEHOR% = HOCOM%
                        End If
                    End If
                End If
            Next UI%
            '
'MsgBox "7   " & (Timer - TINXX): TINXX = Timer
            If DEHOR% <= HAHOR% Then
                TDIS = DIFHORA(DEHOR%, HAHOR%)
              Else
                TDIS = 0
            End If
            '
            If TNEC <= TDIS Then
333               MINUDI% = 60 * Int(TNEC) + 60 * (TNEC - Int(TNEC))
                  MINUDI0% = MINUDI%
                  HOFIN% = DEHOR%
                  While MINUDI% > 0
                     HOFIN% = HOFIN% + 1
                     If (HOFIN% Mod 100) > 59 Then
                        HOFIN% = HOFIN% + 40
                     End If
                     MINUDI% = MINUDI% - 1
                  Wend
                  If HOFIN% > DEHOR% Then
                     VECTIPAR$ = VECTIPAR$ + MKI$(FEFIN%) + MKI$(DEHOR%) + MKI$(HOFIN%)
                  End If
'MsgBox FEFIN% & "   " & TURNO% & "   " & DEHOR% & "   " & HOFIN%
'If HOFIN% > 2400 Then
'AAA = BBB
'End If
'                  HOFIN% = DEHOR% + MINUDI%
'MsgBox "A   " & TNEC0 & "   " & HOCOM% & "   " & HOFIN%
                   'If (HOFIN% Mod 100) > 59 Then HOFIN% = HOFIN% + 40
'MsgBox "B   " & TNEC0 & "   " & HOCOM% & "   " & HOFIN%
'GoTo 333
                Else
                  HOFIN% = HAHOR%
                  If HOFIN% > DEHOR% Then
                     VECTIPAR$ = VECTIPAR$ + MKI$(FEFIN%) + MKI$(DEHOR%) + MKI$(HOFIN%)
                  End If
'MsgBox FEFIN% & "   " & TURNO% & "   " & DEHOR% & "   " & HOFIN%
'MsgBox "C   " & TNEC0 & "   " & HOCOM% & "   " & HOFIN%
                  TNEC = TNEC - TDIS
                  TURNO% = TURNO% + 1
                  If TURNO% > 6 Then
                     TURNO% = 1
                     FEX = FEFIN%
                     FEFIN% = DIPOST(FEX)
                  End If
                  GoTo 515
            End If
'MsgBox "8   " & (Timer - TINXX): TINXX = Timer
            '
            FEFINA% = FEFIN%
            HOFINA% = HOFIN%
            '
' verificar si hay interferencias - AQUI ES DONDE TENGO QUE ACTUAR 31/05/2004
            '
            DESPRO# = 10000# * CDbl(FECOM%) + CDbl(HOCOM%)
            HASPRO# = 10000# * CDbl(FEFIN%) + CDbl(HOFIN%)
            '
            REGOPEL& = REGOPE&(KY%)
            O$ = REGLEIDO$("OPERFAP", REGOPEL&)
            MAQUIS$ = Mid$(O$, 73, 6) + Mid$(O$, 203, 18)

            For IMQ% = 0 To 3
               MQX% = MAQNU%(Mid$(MAQUIS$, 6 * IMQ% + 1, 6))
               If MQX% > 0 Then
                  For IBLO% = 1 To Len(VECHORES$(MQX%)) Step 12
                     STH$ = Mid$(VECHORES$(MQX%), IBLO%, 8)
                     Desde# = 10000# * CDbl(CVI(Mid$(STH$, 1, 2))) + CDbl(CVI(Mid$(STH$, 3, 2)))
                     HASTA# = 10000# * CDbl(CVI(Mid$(STH$, 5, 2))) + CDbl(CVI(Mid$(STH$, 7, 2)))
                     If DESPRO# < HASTA# And HASPRO# > Desde# Then
                        FEFINA% = CVI(Mid$(STH$, 5, 2))
                        FECOM% = FEFINA%
                        HOFINA% = CVI(Mid$(STH$, 7, 2))
                        HOCOM% = HOFINA%
                        TURNO% = 1
                        GoTo 511
                     End If
                  Next IBLO%
                  'For I& = 1 To CAOPERF&
                  '   For TTI% = 0 To 3
                  '      If MQQ%(I&, TTI%) = MQX% Then
                  '         O$ = REGLEIDO$("OPERFAP", I&)
                  '         Desde# = 10000# * CDbl(CVI(Mid$(O$, 135, 2))) + CDbl(CVI(Mid$(O$, 137, 2)))
                  '         HASTA# = 10000# * CDbl(CVI(Mid$(O$, 139, 2))) + CDbl(CVI(Mid$(O$, 141, 2)))
                  '         '
'                           If DESPRO# >= DESDE# And DESPRO# < HASTA# Then
'                               FEFINA% = CVI(Mid$(O$, 139, 2))
'                               HOFINA% = CVI(Mid$(O$, 141, 2))
' MsgBox Str$(DESPRO#) + Str$(DESDE#) + Str$(HASTA#) + "VOY A 511"
'                               GoTo 511
'                           End If
                           '
'                            If HASPRO# >= DESDE# And HASPRO# < HASTA# Then
'                               FEFINA% = CVI(Mid$(O$, 139, 2))
'                               FECOM% = FEFINA%
'                               HOFINA% = CVI(Mid$(O$, 141, 2))
'                               HOCOM% = HOFINA%
'                               GoTo 511
'                           End If
                  '         If DESPRO# < HASTA# And HASPRO# > Desde# Then
                  '             FEFINA% = CVI(Mid$(O$, 139, 2))
                  '             FECOM% = FEFINA%
                  '             HOFINA% = CVI(Mid$(O$, 141, 2))
                  '             HOCOM% = HOFINA%
                  '             TURNO% = 1
                  '             GoTo 511
                  '         End If
                  '         '
                  '      End If
                  '   Next TTI%
                  'Next I&
               End If
            Next IMQ%
'MsgBox "9   " & (Timer - TINXX): TINXX = Timer
            '
            JVP& = JVP& + 1
            VP1$ = String$(256, 0)
            Call REPLA(VP1$, MKI$(Len(VECTIPAR$)), 1, 2)
            Call REPLA(VP1$, VECTIPAR$, 3, Len(VECTIPAR$))
            Call GRAREG("VECTIPAR", VP1$, JVP&)
            '
' AQUI VAN LOS TIEMPOS DE PREPARACION
O$ = REGLEIDO$("OPERFAP", REGOPEL&)
TNEPREP = CVS(Mid$(O$, 103, 4)) / 60 ' EN HORAS
FEFINPREP% = FECOM%
TURPREP% = 1
332 '
Call HODISTUR(EQ$(UI%), FEFINPREP%, TURPREP%, DEHORPREP%, HAHORPREP%)
If FEFINPREP% <= FECOM% Then
  If DEHORPREP% < HOCOM% Then
    DEHORPREP% = HOCOM%
  End If
End If
'
If DEHORPREP% <= HAHORPREP% Then
    TDISPREP = DIFHORA(DEHORPREP%, HAHORPREP%)
   Else
    TDISPREP = 0
End If
If TNEPREP <= TDISPREP Then          '
    MINUDIP% = 60 * Int(TNEPREP) + 60 * (TNEPREP - Int(TNEPREP))
      MINUDIQ% = MINUDIP%
      HOFINPREP% = DEHORPREP%
      While MINUDIP% > 0
         HOFINPREP% = HOFINPREP% + 1
         If (HOFINPREP% Mod 100) > 59 Then
            HOFINPREP% = HOFINPREP% + 40
         End If
         MINUDIP% = MINUDIP% - 1
      Wend
      '            If HOFIN% > DEHOR% Then
      '               VECTIPAR$ = VECTIPAR$ + MKI$(FEFIN%) + MKI$(DEHOR%) + MKI$(HOFIN%)
      '            End If
    Else
      HOFINPREP% = HAHORPREP%
      '            If HOFIN% > DEHOR% Then
      '               VECTIPAR$ = VECTIPAR$ + MKI$(FEFIN%) + MKI$(DEHOR%) + MKI$(HOFIN%)
      '            End If
      TNEPREP = TNEPREP - TDISPREP
      TURPREP% = TURPREP% + 1
      If TURPREP% > 6 Then
        TURPREP% = 1
        FEX = FEFINPREP%
        FEFINPREP% = DIPOST(FEX)
      End If
      GoTo 332
End If
' ***********************************
            STH$ = MKI$(FECOM%) + MKI$(HOCOM%) + MKI$(FEFIN%) + MKI$(HOFIN%) + MKI$(FEFINPREP%) + MKI$(HOFINPREP%)
            O$ = REGLEIDO$("OPERFAP", REGOPEL&)
            Call REPLA(O$, STH$, 135, 12)
            Call REPLA(O$, MKS$(JVP&), 253, 4)
            NOFAXI$ = Left$(O$, 12)
            OBSEX$ = "      ": If NOFAXI$ <> NOFAXJ$ Then OBSEX$ = "Op.Lib"
            Call REPLA(O$, OBSEX$, 225, 6)
            Call GRAREG("OPERFAP", O$, REGOPEL&)
            NOFAXJ$ = NOFAXI$
            '
            For IMQ% = 0 To 3
               MQX% = MAQNU%(Mid$(MAQUIS$, 6 * IMQ% + 1, 6))
               If MQX% > 0 Then
                  LVX% = Len(VECHORES$(MQX%))
                  If LVX% >= 8 Then
                     If Mid$(VECHORES$(MQX%), LVX% - 3, 4) = Left$(STH$, 4) Then
                        VECHORES$(MQX%) = Left$(VECHORES$(MQX%), LVX% - 4) + Mid$(STH$, 5, 8)
                        GoTo 519
                     End If
                  End If
                  VECHORES$(MQX%) = VECHORES$(MQX%) + STH$
               End If
519         Next IMQ%
'MsgBox Len(VECTIPAR$)
'If HOCOM% > 2400 Or HOFIN% > 2400 Then
' MsgBox TNEC0 & "  " & FECOM% & "  " & HOCOM% & "  " & FEFIN% & "  " & HOFIN%
'End If
            '
            If FEFINANT%(NIVEFA%) < FEFIN% Or (FEFINANT%(NIVEFA%) = FEFIN% And HOFINANT%(NIVEFA%) < HOFIN%) Then
               FEFINANT%(NIVEFA%) = FEFIN%
               HOFINANT%(NIVEFA%) = HOFIN%
            End If
            '
            MAQUIS$ = Mid$(O$, 73, 6) + Mid$(O$, 203, 18)
            For IMQ% = 0 To 3
               MQQ%(REGOPEL&, IMQ%) = MAQNU%(Mid$(MAQUIS$, 6 * IMQ% + 1, 6))
            Next IMQ%
            '
'MsgBox "A   " & (Timer - TINXX): TINXX = Timer
            FEX2 = FEFIN%
            Forms(IFORCA%).Text1(5).TEXT = FECHATEX$(FEX1)
            Forms(IFORCA%).Text1(6).TEXT = HORATEX$(HOCOM%)
            Forms(IFORCA%).Text1(7).TEXT = FECHATEX$(FEX2)
            Forms(IFORCA%).Text1(8).TEXT = HORATEX$(HOFIN%)
            '
   'MsgBox HOFIN%
            For III% = 0 To 8: Forms(IFORCA%).Text1(III%).Refresh: Next III%
            '
524     Next KY%
        '
        REGORFA& = REGO&(KK%)
        Z$ = REGLEIDO$("ORDEFABR", REGORFA&)
        Call REPLA(Z$, MKI$(FEFIN%), 147, 2)
        Call GRAREG("ORDEFABR", Z$, REGORFA&)
        '
590 'Next KK%
End If
'MsgBox Timer - TINXX
Next kk8%
Next NIVEFA%
Next kk7%
    '
5555 '
Call SETULTREG("VECTIPAR", JVP&)
Call CIERRARCH("VECTIPAR")
    '
    Forms(IFORCA%).Check6.Value = 2
    FIN& = ULTREG&("ORDEFABR")
    FIN1& = FIN&: If FIN1& < 1 Then FIN1& = 1
    '
    'FORMS(IFORCA%).ProgressBar1.Min = 0
    'FORMS(IFORCA%).ProgressBar1.Max = FIN1&
    '
    Call COPYSTRU("ORDEFABR", "ORDEFABP")
    J& = 0
    For I& = 1 To FIN&
        '
        'FORMS(IFORCA%).ProgressBar1.Value = I&
        Call TRACE(20, I&, FIN&)
        F$ = REGLEIDO$("ORDEFABR", I&)
        '
        If Asc(Mid$(F$, 108, 1)) = 1 Then ' esta lanzada
            J& = J& + 1
            Call GRAREG("ORDEFABP", F$, J&)
        End If
        '
    Next I&
    '
    Call SETULTREG("ORDEFABP", J&)
    Call CIERRARCH("*.*")
    Forms(IFORCA%).Check6.Value = 1
    '
GoTo 3800
    '
    '
600 'FORMS(IFORCA%).Check7.Value = 1
    'FORMS(IFORCA%).Check8.Value = 2
    '
    'Call BLANARCH("CRONOPF")
    '
    E0$ = ""
'    FIN& = ULTREG&("CRONOPE")
'    FORMS(IFORCA%).ProgressBar1.Min = 0
'    FORMS(IFORCA%).ProgressBar1.Max = FIN&
    '
    For I& = 1 To FIN&
        '
        'FORMS(IFORCA%).ProgressBar1.Value = I&
        Call TRACE(20, I&, FIN&)
        E1$ = Left$(REGLEIDO$("CRONOPE", I&), 72)
        MQ$ = Mid$(REGLEIDO$("CRONOPE", I&), 73, 6)
        E2$ = Mid$(REGLEIDO$("CRONOPE", I&), 79)
        '
        If E1$ + Mid$(E2$, 57, 2) = E0$ Then
             F$ = REGLEIDO$("CRONOPF", JL&)
             MQ$ = TRIM$(Mid$(F$, 73, 28)) + "-" + MQ$
             F$ = REPLACE$(F$, MQ$, 73, 28)
             Call GRAREG("CRONOPF", F$, JL&)
          Else
             F$ = E1$ + Space$(28) + E2$
             F$ = REPLACE$(F$, MQ$, 73, 28)
             Call REGAPP("CRONOPF", F$)
             JL& = ULTREG&("CRONOPF")
             E0$ = E1$ + Mid$(E2$, 57, 2)
       End If
    Next I&
    '
    'FORMS(IFORCA%).Check8.Value = 1
    '
700 'FORMS(IFORCA%).Check9.Value = 2
    '
    Call BLANARCH("GANTORF")
    '
    E0$ = ""
    FIN& = ULTREG&("GANTMAQ")
    'FORMS(IFORCA%).ProgressBar1.Min = 0
    'FORMS(IFORCA%).ProgressBar1.Max = FIN&
    '
    For I& = 1 To FIN&
        '
        'FORMS(IFORCA%).ProgressBar1.Value = I&
        Call TRACE(20, I&, FIN&)
        E1$ = Left$(REGLEIDO$("GANTMAQ", I&), 6)
        MQ$ = Mid$(REGLEIDO$("GANTMAQ", I&), 7, 6)
        E2$ = Mid$(REGLEIDO$("GANTMAQ", I&), 13, 72)
        E3$ = Mid$(REGLEIDO$("GANTMAQ", I&), 91)
        '
        If E1$ + E2$ + Mid$(E3$, 39, 60) = E0$ Then
             F$ = REGLEIDO$("GANTORF", JL&)
             MQ$ = TRIM$(Mid$(F$, 7, 28)) + "-" + MQ$
             F$ = REPLACE$(F$, MQ$, 7, 28)
             F$ = REPLACE$(F$, MQ$, 107, 28)
             Call GRAREG("GANTORF", F$, JL&)
          Else
             F$ = E1$ + Space$(28) + E2$ + Space$(28) + E3$
             F$ = REPLACE$(F$, MQ$, 7, 28)
             F$ = REPLACE$(F$, MQ$, 107, 28)
             Call REGAPP("GANTORF", F$)
             JL& = ULTREG&("GANTORF")
             E0$ = E1$ + E2$ + Mid$(E3$, 39, 60)
       End If
    Next I&
    '
'    XXX$ = CTRLFAC$("CALCUMA->" + HO$ + " a las " + Left$(Time$, 5) + " horas.")
    GoTo 3800

610 N88% = FreeFile
'    OPEN "R",#N88%,LU$+"PROCARG.DAT",128
'   FIELD #N88%,80 AS DX0$,2 AS DX18$,2 AS DX19$,2 AS DX20$,2 AS DX21$,40 AS DX22$
    If LOF(N88%) > 0 Then
        Close #N88%
        Kill LU$ + "PROCARG.DAT"
        GoTo 610
    End If

611 N22% = FreeFile
'    OPEN "R",N22%,LU$+"HORSECC.DAT",32
'    FIELD #N22%,2 AS X1$, 6 AS X2$, 4 AS X3$, 4 AS X4$,4 AS X5$,4 AS X6$,8 AS X7$
    REG& = 0
    If LOF(N22%) > 0 Then
        Close #N22%
        Kill LU$ + "HORSECC.DAT"
        GoTo 611
    End If


    DXL& = 0: DXX$ = ""
    For I& = 1 To CX&
 '       GET #N8%,I&
        GoSub 6000
        If CX0$ <> DXX$ Then
            If DXX$ <> "" Then
                DXL& = DXL& + 1
                LSet DX0$ = DXX$
                LSet DX18$ = MKI$(FECOM%)
                LSet DX19$ = MKI$(HOCOM%)
                LSet DX20$ = MKI$(FEFIN%)
                LSet DX21$ = MKI$(HOFIN%)
                LSet DX22$ = String$(22, 0)
  '              PUT #N88%,DXL&
            End If
            DXX$ = CX0$
651         FECOM% = CVI(CX18$): FEFIN% = CVI(CX18$)
652         HOCOM% = CVI(CX19$): HOFIN% = CVI(CX20$)
            NOPER = CVS(CX21$)
'            OPEDIS = CVS(CX22$)
653         GoTo 699
        End If
654     If CVI(CX18$) < FECOM% Then
6541        FECOM% = CVI(CX18$)
6542        HOCOM% = CVI(CX19$)
6543    End If
655     If CVI(CX18$) > FEFIN% Then FEFIN% = CVI(CX18$): HOFIN% = CVI(CX20$)
656     If CVI(CX18$) = FECOM% Then If CVI(CX19$) < HOCOM% Then HOCOM% = CVI(CX19$)
657     If CVI(CX18$) = FEFIN% Then If CVI(CX20$) > HOFIN% Then HOFIN% = CVI(CX20$)
699 Next I&
    DXL& = DXL& + 1
    LSet DX0$ = DXX$
    LSet DX18$ = MKI$(FECOM%)
    LSet DX19$ = MKI$(HOCOM%)
    LSet DX20$ = MKI$(FEFIN%)
    LSet DX21$ = MKI$(HOFIN%)
    LSet DX22$ = String$(22, 0)
'    PUT #N88%,DXL&

800 GoSub 7000

3800 Call CIERRARCH("*.*")
     Screen.MousePointer = 1
     COMPLECALCU% = 1
     Forms(IFORCA%).Hide
     Call CIERRARCH("*.*")
     Call BAJALDISCO
     Exit Sub

4000 ' grafica de GANTT
     '
     If TUNOR = 0 Then GoSub 4100
     '
     FF% = CVI(Mid$(X1$, 135, 2))
 '    FF = FF%: FF$ = FECHATEX$(FF)
     FF = FF%: SMN% = SEMANU%(FF, A, B)
 '    AN$ = Mid$(FF$, 7, 2): If Left$(AN$, 1) = "9" Then AN$ = " " + AN$ Else AN$ = "1" + AN$
'     SMN$ = Mid$(Str$(SMN%), 2): While Len(SMN$) < 2: SMN$ = "0" + SMN$: Wend
'     CODPER$ = Left$(AN$ + "-" + SMN$ + Space$(6), 6)
     PERIO$ = "S." + Mid$(Str$(SMN%), 2) + " - "
 '    A$ = FECHATEX$(A): B$ = FECHATEX$(B)
'     If Mid$(A$, 3) = Mid$(B$, 3) Then A$ = Left$(A$, 2)
'     If Mid$(A$, 6) = Mid$(B$, 6) Then A$ = Left$(A$, 5)
 '    If Left$(A$, 1) = "0" Then A$ = Mid$(A$, 2)
 '    PERIO$ = PERIO$ + A$ + " al " + B$
     '
     For G& = 1 To ULTREG&("PERIPRO")
         G1$ = REGLEIDO$("PERIPRO", G&)
         If Left$(G1$, 6) = CODPER$ Then GoTo 4010
     Next G&
     '
     Call REGAPP("PERIPRO", CODPER$ + PERIO$)
     '
4005 For G& = 1 To ULTREG&("PERIPRO") - 1
         G1$ = REGLEIDO$("PERIPRO", G&)
         G2$ = REGLEIDO$("PERIPRO", G& + 1)
         If G1$ > G2$ Then
             Call GRAREG("PERIPRO", G2$, G&)
             Call GRAREG("PERIPRO", G1$, G& + 1)
             GoTo 4005
         End If
     Next G&
     '
4010 YAESTA& = 0
     For XX& = ULTREG&("GANTMAQ") To 1 Step -1
         X2$ = REGLEIDO$("GANTMAQ", XX&)
         If Mid$(X2$, 13, 16) <> Left$(X1$, 16) Then GoTo 4020 ' distinta orden/operacion
         If Left$(X2$, 6) = CODPER$ Then                ' mismo periodo
             If Mid$(X2$, 85, 6) = Mid$(X1$, 73, 6) Then ' mismo equipo
                 YAESTA& = XX&
                 GoTo 4030
             End If
         End If
     Next XX&
     '
4020 X2$ = REGBLAN$("GANTMAQ")
     X2$ = REPLACE$(X2$, CODPER$, 1, 6)
     X2$ = REPLACE$(X2$, Mid$(X1$, 73, 6), 7, 6)
     X2$ = REPLACE$(X2$, Left$(X1$, 78), 13, 78)
     X2$ = REPLACE$(X2$, Mid$(X1$, 95, 8), 91, 8)
     '
4030 COLUI% = 128 + 8 * DIASEM%(FF) - 7
     DH% = Int((CVI(Mid$(X1$, 137, 2)) - INOR%) / 100 + 0.5)
     HH% = Int((CVI(Mid$(X1$, 139, 2)) - INOR%) / 100 + 0.5)
'     TIE = DIFHORA(CVI(Mid$(X1$, 137, 2)), CVI(Mid$(X1$, 139, 2))) + CVS(Mid$(X2$, 185, 4))
     DHX% = 16 * DH% / TUNOR
     TTXI% = 16 * (HH% - DH%) / TUNOR: If TTXI% < 1 Then If HH% > DD% Then TTXI% = 1
     BLK$ = Left$(Space$(DHX%) + String$(TTXI%, "*") + Space$(16), 16)
     '
     BLQ$ = "        "
     For TTI% = 1 To 16 Step 2
        If Mid$(BLK$, TTI%, 2) = "**" Then
            BB$ = "Û"
          ElseIf Mid$(BLK$, TTI%, 2) = "* " Then
            BB$ = "Ý"
          ElseIf Mid$(BLK$, TTI%, 2) = " *" Then
            BB$ = "Þ"
          Else
            BB$ = " "
        End If
        Mid$(BLQ$, (TTI% + 1) / 2, 1) = BB$
     Next TTI%
     '
     For TTI% = 1 To Len(BLQ$)
        If Mid$(BLQ$, TTI%, 1) <> " " Then
            X2$ = REPLACE$(X2$, Mid$(BLQ$, TTI%, 1), COLUI% + TTI% - 1, 1)
        End If
     Next TTI%
     X2$ = REPLACE$(X2$, MKS$(TIE), 185, 4)

     '
4040 If YAESTA& > 0 Then
          Call GRAREG("GANTMAQ", X2$, YAESTA&)
        Else
          Call REGAPP("GANTMAQ", X2$)
     End If
     '
4090 Return



4100 For U& = 1 To ULTREG&("TURGEN")
          X$ = REGLEIDO$("TURGEN", U&)
          DSX% = Asc(Left$(X$, 1))
          If DSX% = 1 Then
               DHR% = CVI(Mid$(X$, 2, 2))
               HHR% = CVI(Mid$(X$, 4, 2))
               INOR% = DHR%
               TUNOR = (HHR% - DHR%) / 100
          End If
     Next U&
     '
     If TUNOR < 8 Then
         Call MENSERR(24, "Error de Configuracion de Turnos Generales")
         Call FINAL("")
     End If
     '
     Return
     '

6000 FECHA% = CVI(CX18$)
     DESHOR% = CVI(CX19$)
     HASHOR% = CVI(CX20$)
     NOPER = CVS(CX21$)
'     OPEDIS = CVS(CX22$)
'     MAQUIS$= C9$

'     HORAS = DIFHORA(DESHOR%, HASHOR%)

     For JK% = 1 To LOF(N7%) / 128
'            GET #N7%,JK%
'            If Q1$ = MAQUIS$Then SECCI$ = Q3$: GoTo 6010
     Next JK%
'     Call MENSERR(24, "Seccion no Definida para Maquina '" + TRIM$(MAQUI$) + "'")
     GoTo 6099                                                     ' next i&


6010 For JKL& = 1 To REG&
'         GET #N22%,JKL&
         If CVI(X1$) = FECHA% Then
                If X2$ = SECCI$ Then
                     LSet X3$ = MKS$(CVS(X3$) + HORAS)
                     LSet X4$ = MKS$(CVS(X4$) + HORAS * NOPER)
                     LSet X5$ = MKS$(0)
                     LSet X6$ = MKS$(0)
                     LSet X7$ = String$(8, 0)
'                     PUT #N22%,JKL&
                     GoTo 6099
                End If
         End If
     Next JKL&

     REG& = REG& + 1
     LSet X1$ = MKI$(FECHA%)
     LSet X2$ = SECCI$
     LSet X3$ = MKS$(HORAS)
     LSet X4$ = MKS$(HORAS * NOPER)
     LSet X5$ = MKS$(0)
     LSet X6$ = MKS$(0)
     LSet X7$ = String$(8, 0)
 '    PUT #N22%,REG&

6099 Return

7000 For JKL& = 1 To REG&
'        GET #N22%,JKL&
        FECHA% = CVI(X1$)
        SECCI$ = X2$
7015    FEX = FECHA%
        DS% = DIASEM%(FEX)

        CAPAMOB = 0
        For JJ% = 1 To LOF(N11%) / 64
 '           GET #N11%,JJ%
            If CP1$ = SECCI$ Then
                If Asc(CP3$) = DS% Then
 '                   CAPAMOB = CAPAMOB + CVI(CP7$) * DIFHORA(CVI(CP5$), CVI(CP6$))
                End If
            End If
        Next JJ%

        CAPAMOX = 0
        For JJ% = 1 To LOF(N12%) / 64
'            GET #N12%,JJ%
            If CF1$ = SECCI$ Then
                If CVI(CF3$) = FECHA% Then
'                    CAPAMOX = CAPAMOX + CVI(CF6$) * DIFHORA(CVI(CF4$), CVI(CF5$))
                End If
            End If
        Next JJ%

        If CAPAMOX > 0 Then CAPAMOB = CAPAMOX

        LSet X5$ = MKS$(CAPAMOB)
        RELA = 0: If CVS(X5$) > 0 Then RELA = 100 * CVS(X4$) / CVS(X5$)
        LSet X6$ = MKS$(RELA)
'        PUT #N22%,JKL&
     Next JKL&
     Return

End Sub
' ***************************************************************************

Sub HODISPO(EQ$, DIA%, DEHOR%, HAHOR%)
     '
     ' devuelve sobre DEHOR% y HAHOR% los horarios de disponibilidad del
     ' equipo EQ$ en la fecha DIA%, en base a los archivos de
     ' disponibilidad horaria
     '
     Static DHOR%(7), HHOR%(7)
     '
     If CADIA% = 0 Then
         CADIA% = 7
         For U& = 1 To ULTREG&("TURGEN")
             X$ = REGLEIDO$("TURGEN", U&)
             DSX% = Asc(Left$(X$, 1))
             If DSX% > 0 Then
                 If DSX% <= 7 Then
                     DHR% = CVI(Mid$(X$, 2, 2))
                     HHR% = CVI(Mid$(X$, 4, 2))
                     If DHR% >= 0 Then
                         If DHR% <= 2400 Then
                             DHOR%(DSX%) = DHR%
                         End If
                     End If
                     '
                     HHOR%(DSX%) = DHOR%(DSX%)
                     If HHR% > DHOR%(DSX%) Then
                         If HHR% <= 2400 Then
                             HHOR%(DSX%) = HHR%
                         End If
                     End If
                 End If
             End If
         Next U&
         DHOR%(0) = DHOR%(7)
         HHOR%(0) = HHOR%(7)
         '
         TSEM = 0
         For U& = 1 To 7
            TSEM = TSEM + HHOR%(U&) - DHOR%(U&)
         Next U&
         '
         If TSEM < 1 Then
            OPP% = COMALTER%("Proceso Imposible.\Falta Definir Horarios\y Turnos Diarios de Trabajo.\\Definir Horarios\Terminar.")
            If OPP% = 1 Then
                Call FINAL("*TURGEN")
            End If
            Call CIERRARCH("*.*")
            Call FINAL("")
         End If
         '
     End If
     '
     FECHA = DIA%
     DS% = DIASEM%(FECHA)
     DEHOR% = DHOR%(DS%): HAHOR% = HHOR%(DS%)
     '
     For U& = 1 To ULTREG&("CALFERI")
         If CVI(Left$(REGLEIDO$("CALFERI", U&), 2)) = DIA% Then
             DEHOR% = 0
             HAHOR% = 0
             Exit Sub
         End If
     Next U&
     '
End Sub

Sub HODISTUR(EQ$, DIA%, TURNO%, DEHOR%, HAHOR%)
     '
     ' devuelve sobre DEHOR% y HAHOR% los horarios de disponibilidad del
     ' equipo EQ$ en la fecha DIA%, en base a los archivos de
     ' disponibilidad horaria
     '
     Static CADIA%, DHORT%(7, 6), HHORT%(7, 6)
     Static DHORM%(7, 6, 512), HHORM%(7, 6, 512)
     '
     If CADIA% = 0 Then
         '
         CADIA% = 7
         For U& = 1 To ULTREG&("TURGEN")
             X$ = REGLEIDO$("TURGEN", U&)
             DSX% = Asc(Left$(X$, 1))
             If DSX% > 0 Then
                 If DSX% <= 7 Then
                    For KKI% = 1 To 6
                       DHR% = CVI(Mid$(X$, 4 * KKI% - 2, 2))
                       HHR% = CVI(Mid$(X$, 4 * KKI%, 2))
                       If DHR% >= 0 Then
                         If DHR% <= 2400 Then
                             DHORT%(DSX%, KKI%) = DHR%
                         End If
                       End If
                       If DHORT%(DSX%, KKI%) < HHORT%(DSX%, KKI% - 1) Then
                          DHORT%(DSX%, KKI%) = HHORT%(DSX%, KKI% - 1)
                       End If
                       '
                       HHORT%(DSX%, KKI%) = DHORT%(DSX%, KKI%)
                       If HHR% > DHORT%(DSX%, KKI%) Then
                         If HHR% <= 2400 Then
                             HHORT%(DSX%, KKI%) = HHR%
                         End If
                       End If
                    Next KKI%
                 End If
             End If
         Next U&
         For KKI% = 1 To 6
            DHORT%(0, KKI%) = DHORT%(7, KKI%)
            HHORT%(0, KKI%) = HHORT%(7, KKI%)
         Next KKI%
         '
         TSEM = 0
         For U& = 1 To 7
            For KKI% = 1 To 6
               TSEM = TSEM + HHORT%(U&, KKI%) - DHORT%(U&, KKI%)
            Next KKI%
         Next U&
         '
         If TSEM < 1 Then
            OPP% = COMALTER%("Proceso Imposible.\Falta Definir Horarios\y Turnos Diarios de Trabajo.\\Definir Horarios\Terminar.")
            If OPP% = 1 Then
                Call FINAL("*TURGEN")
            End If
            Call CIERRARCH("*.*")
            Call FINAL("")
         End If
         '
         For MQPX% = 0 To NUMAQ%
            For DSX% = 0 To 7
               For KKI% = 1 To 6
                  DHORM%(DSX%, KKI%, MQPX%) = DHORT%(DSX%, KKI%)
                  HHORM%(DSX%, KKI%, MQPX%) = HHORT%(DSX%, KKI%)
               Next KKI%
            Next DSX%
         Next MQPX%
         '
         For U& = 1 To ULTREG&("TURMAQ")
            X$ = REGLEIDO$("TURMAQ", U&)
            DSX% = Asc(Left$(X$, 1))
            MQPX% = MAQNU%(Mid$(X$, 33, 6))
            For KKI% = 1 To 6
               DHORM%(DSX%, KKI%, MQPX%) = 0
               DHR% = CVI(Mid$(X$, 4 * KKI% - 2, 2))
               HHR% = CVI(Mid$(X$, 4 * KKI%, 2))
               If DHR% >= 0 Then
                  If DHR% <= 2400 Then
                     DHORM%(DSX%, KKI%, MQPX%) = DHR%
                  End If
               End If
               HHORM%(DSX%, KKI%, MQPX%) = DHORM%(DSX%, KKI%, MQPX%)
               If HHR% > DHORM%(DSX%, KKI%, MQPX%) Then
                  If HHR% <= 2400 Then
                     HHORM%(DSX%, KKI%, MQPX%) = HHR%
                  End If
               End If
               If DSX% = 7 Then
                  DHORM%(0, KKI%, MQPX%) = DHORM%(7, KKI%, MQPX%)
                  HHORM%(0, KKI%, MQPX%) = HHORM%(7, KKI%, MQPX%)
               End If
            Next KKI%
         Next U&
         '
     End If
     '
     DEHOR% = 0: HAHOR% = 0
     FECHA = DIA%
     DS% = DIASEM%(FECHA)
     MQP% = MAQNU%(EQ$)
     '
     ' VERIFICAR TURNOS ESPECIALES POR FECHA
     
     'If TURNO% > 0 Then
     '   If TURNO% <= 6 Then
     '      DEHOR% = DHORT%(DS%, TURNO%)
     '      HAHOR% = HHORT%(DS%, TURNO%)
     '   End If
     'End If
     '
     For U& = 1 To ULTREG&("CALFERI")
         If CVI(Left$(REGLEIDO$("CALFERI", U&), 2)) = DIA% Then
             DEHOR% = 0
             HAHOR% = 0
             GoTo 99
         End If
     Next U&
     '
     If DS% >= 0 Then
       If DS% <= 7 Then
         If TURNO% > 0 Then
           If TURNO% <= 6 Then
             If MQP% >= 0 Then
               If MQP% <= 512 Then
                 DEHOR% = DHORM%(DS%, TURNO%, MQP%)
                 HAHOR% = HHORM%(DS%, TURNO%, MQP%)
               End If
             End If
           End If
         End If
       End If
     End If
     'EQP$ = TRIM$(EQ$)
     'If TURNO% > 0 Then
     '   If TURNO% <= 6 Then
     '     For U& = 1 To ULTREG&("TURMAQ")
     '        X$ = REGLEIDO$("TURMAQ", U&)
     '        If Asc(Left$(X$, 1)) = DS% Then
     '           If TRIM$(Mid$(X$, 33, 6)) = EQP$ Then
     '              KKI% = TURNO%
     '              DEHOR% = 0
     '              DHR% = CVI(Mid$(X$, 4 * KKI% - 2, 2))
     '              HHR% = CVI(Mid$(X$, 4 * KKI%, 2))
     '              If DHR% >= 0 Then
     '                 If DHR% <= 2400 Then
     '                    DEHOR% = DHR%
     '                 End If
     '              End If
     '              HAHOR% = DEHOR%
     '              If HHR% > DEHOR% Then
     '                 If HHR% <= 2400 Then
     '                    HAHOR% = HHR%
     '                 End If
     '              End If
     '           End If
     '        End If
     '     Next U&
     '     '
          For U& = 1 To ULTREG&("TURESP")
             X$ = REGLEIDO$("TURESP", U&)
             If CVI(Mid$(X$, 31, 2)) = DIA% Then
                If TRIM$(Mid$(X$, 33, 6)) = EQP$ Then
                   KKI% = TURNO%
                   DHR% = CVI(Mid$(X$, 4 * KKI% - 2, 2))
                   HHR% = CVI(Mid$(X$, 4 * KKI%, 2))
                   DEHOR% = 0
                   If DHR% >= 0 Then
                      If DHR% <= 2400 Then
                         DEHOR% = DHR%
                      End If
                   End If
                   HAHOR% = DEHOR%
                   If HHR% > DEHOR% Then
                      If HHR% <= 2400 Then
                         HAHOR% = HHR%
                      End If
                   End If
                End If
             End If
          Next U&
          '
     '   End If
     'End If
     '
99 End Sub

Function MAQNU%(MQS$)
    '
    MAQNU% = 0
    PPXY% = InStr(VECTOMAQ$, MQS$)
    If PPXY% Mod 8 = 3 Then
       MAQNU% = (PPXY% + 5) / 8
    End If
    '
    'If TRIM$(MQS$) <> "" Then
    '    For UI% = 1 To NUMAQ%
    '        If MAQUI$(UI%) = MQS$ Then
    '            MAQNU% = UI%
    '            Exit Function
    '        End If
    '    Next UI%
    'End If
    '
End Function
'
'For I& = 1 To CAOR%
'   TXTX$ = Left$(ORFA$(I&), 12)
'   For KK% = 15 To Len(ORFA$(I&)) Step 4
'      TXTX$ = TXTX$ + Str$(CVS(Mid$(ORFA$(I&), KK%, 4)))
'   Next KK%
'   MsgBox TXTX$
'Next I&
'Stop

Sub PLACARMA()        ' GENERA PLANILLA DE CARGA DE MAQUINA
     '
     Dim MAQUISEL$(512)
     CAMAQUI% = 0
     HO% = HOY(HOS$)
     LU$ = LUDAT$
     '
     TITU$ = "INFORMES DE PRODUCCION - GENERALES"
     '
     FEX = HO%
95   DIADOS% = DIPOST(FEX)
     FEX = DIADOS%
     If DIASEM%(FEX) > 5 Or DIASEM%(FEX) < 1 Then GoTo 95
     '     Call BLANARCH("OPERFAQ")
100  VDEF$ = MKI$(HO%) + "Todos " + MKI$(DIADOS%)
101  DIAEQUI$ = OBJPLA$("DIAEQUI", VDEF$)
     If DIAEQUI$ = "" Then
         REGQ& = 0
         'GoTo 190
         Exit Sub
     End If
     '
     Screen.MousePointer = 11
     DIA% = CVI(Left$(DIAEQUI$, 2))
     FEX = DIA%
110  DIADOS% = CVI(Mid$(DIAEQUI$, 9, 2))
     If DIADOS% < DIA% Then Exit Sub
     MAQUIX$ = Mid$(DIAEQUI$, 3, 6)
     MODO% = 1: If MAQUIX$ = "Todos " Then MODO% = 2
     '
     Screen.MousePointer = 11
     FEX = DIA%
     CORRESP$ = "Corresponde a: " + FECHATEX$(FEX)
     FEX = DIADOS%
     CORRESP$ = CORRESP$ + " - " + FECHATEX$(FEX)
     '
     CAEQSEL% = 0
     For REGI& = 1 To ULTREG&("OPERFAP")
         XX$ = REGLEIDO$("OPERFAP", REGI&)
         DES% = CVI(Mid$(XX$, 135, 2))
         HAS% = CVI(Mid$(XX$, 139, 2))
         If DES% >= DIA% Then
            If HAS% <= DIADOS% Then
               CAPROD# = CVD(Mid$(XX$, 79, 8))
               CAREA# = CVD(Mid$(XX$, 87, 8))
               SALDO# = CAPROD# - CAREA#: If SALDO# < 0 Then SALDO# = 0
               Call REPLA(XX$, MKD$(SALDO#), 95, 8)
               Call GRAREG("OPERFAP", XX$, REGI&)
               '
               MAQUIY$ = Mid$(XX$, 73, 6)
               For J% = 1 To CAEQSEL%
                  If EQSEL$(J%) = MAQUIY$ Then
                     PUNOP$(J%) = PUNOP$(J%) + MKS$(REGI&)
                     GoTo 120
                  End If
               Next J%
               '
               If MODO% = 2 Or MAQUIY$ = MAQUIX$ Then
                  CAEQSEL% = CAEQSEL% + 1
                  EQSEL$(CAEQSEL%) = MAQUIY$
                  PUNOP$(CAEQSEL%) = MKS$(REGI&)
               End If
               '
            End If
         End If
120   Next REGI&
      '
      For I% = 1 To CAEQSEL%
         For J% = I% + 1 To CAEQSEL%
            If EQSEL$(I%) > EQSEL$(J%) Then
               EQRES$ = EQSEL$(J%)
               PURES$ = PUNOP$(J%)
               EQSEL$(J%) = EQSEL$(I%)
               PUNOP$(J%) = PUNOP$(I%)
               EQSEL$(I%) = EQRES$
               PUNOP$(I%) = PURES$
            End If
         Next J%
      Next I%
      '
130   FORIPRE$ = TRIM$(CONTROL$("", "PLACARMA"))
      If FORIPRE$ <> "" Then
         GoSub 500
         GoTo 199
      End If
      '
      REGQ& = 0
      Call HEADERS("OPERFAQ", "")
      Screen.MousePointer = 11
      For I% = 1 To CAEQSEL%
        If Len(PUNOP$(I%)) >= 4 Then
          MAQUIX$ = EQSEL$(I%)
          DEQUIX$ = ECOARCH$("PADMAQ", MAQUIX$)
          If REGQ& > 0 Then
            REGQ& = REGQ& + 1
            XX$ = REGBLAN$("OPERFAQ")
            Call GRAREG("OPERFAQ", XX$, REGQ&)
            REGQ& = REGQ& + 1
            Call GRAREG("OPERFAQ", String$(256, "="), REGQ&)
          End If
          REGQ& = REGQ& + 1
          XX$ = REGBLAN$("OPERFAQ")
          Call GRAREG("OPERFAQ", XX$, REGQ&)
          REGQ& = REGQ& + 1
          XX$ = REGBLAN$("OPERFAQ")
          Call REPLA(XX$, "@Equipo: " + TRIM$(MAQUIX$) + " - " + DEQUIX$, 17, 48)
          Call GRAREG("OPERFAQ", XX$, REGQ&)
          REGQ& = REGQ& + 1
          XX$ = REGBLAN$("OPERFAQ")
          Call REPLA(XX$, "@******* ", 17, 48)
          Call GRAREG("OPERFAQ", XX$, REGQ&)
          REGQ& = REGQ& + 1
          XX$ = REGBLAN$("OPERFAQ")
          Call GRAREG("OPERFAQ", XX$, REGQ&)
          '
          For JJ& = 1 To Len(PUNOP$(I%)) Step 4
            REGI& = CVS(Mid$(PUNOP$(I%), JJ&, 4))
            If REGI& > 0 Then
              XX$ = REGLEIDO$("OPERFAP", REGI&)
              DES% = CVI(Mid$(XX$, 135, 2))
              HAS% = CVI(Mid$(XX$, 139, 2))
              CAPROD# = CVD(Mid$(XX$, 79, 8))
              CAREA# = CVD(Mid$(XX$, 87, 8))
              SALDO# = CAPROD# - CAREA#: If SALDO# < 0 Then SALDO# = 0
              Call REPLA(XX$, MKD$(SALDO#), 95, 8)
              Call GRAREG("OPERFAP", XX$, REGI&)
              '
              'pasar tiempo de minutos a segundos
              For KK1% = 103 To 131 Step 4
                 TIESEGUN = 60 * CVS(Mid$(XX$, KK1%, 4))
                 Call REPLA(XX$, MKS$(TIESEGUN), KK1%, 4)
              Next KK1%
              REGQ& = REGQ& + 1
              Call GRAREG("OPERFAQ", XX$, REGQ&)
            End If
          Next JJ&
        End If
     Next I%
     '
     Screen.MousePointer = 1
     If REGQ& < 1 Then
        Call MENSERR(24, "No hay Operaciones Programadas\para esa Fecha y ese Equipo.")
        GoTo 100
     End If
     '
190  Call SETULTREG("OPERFAQ", REGQ&)
     '
     Call CIERRARCH("OPERFAQ")
     If ULTREG&("OPERFAQ") > 0 Then
        Call HEADERS("OPERFAQ", "")
        Call HEADERS("OPERFAQ", CORRESP$)
        Call CIERRARCH("*.*")
        Call CONECRUN("*OPERFAQ", "Planilla de Carga de Maquina")
     End If
     '
199  If MODO% = 1 Then GoTo 100
     Exit Sub
     '
500  Screen.MousePointer = 11
     HOII% = HOY(HOS$)
     FECHDOC$ = HOS$
     NUMEDOC$ = ""
     TIPODOC$ = "Planilla de Carga de Máquina"
     '
     CODPARAM$(1) = "EQUIPO1"
     CAPARDOC% = 1
     '
     CODTABU1$(1) = "NOP-OPER"
     CODTABU1$(2) = "DES-OPER"
     CODTABU1$(3) = "CANTIDAD"
     CODTABU1$(4) = "TIE-TOTA"
     CODTABU1$(5) = "FECH-VEN"
     CODTABU1$(6) = "TIE-PREP"
     CODTABU1$(9) = "FECH-REC"
     CODTABU1$(7) = "TIE-PROD"
     CODTABU1$(8) = "OBS-OPER"
     CACOLTAB1% = 9
     '
     For I% = 1 To CAEQSEL%
        If Len(PUNOP$(I%)) >= 4 Then
           MAQUIX$ = EQSEL$(I%)
           DEQUIX$ = ECOARCH$("PADMAQ", MAQUIX$)
           DOCPARAM$(1) = DEQUIX$
           '
           CAITAB1% = 0
           For JJ& = 1 To Len(PUNOP$(I%)) Step 4
             REGI& = CVS(Mid$(PUNOP$(I%), JJ&, 4))
             If REGI& > 0 Then
               XX$ = REGLEIDO$("OPERFAP", REGI&)
               NOFAX$ = Left$(XX$, 12)
               CIXI% = CVI(Mid$(XX$, 13, 2))
               CAPROD# = CVD(Mid$(XX$, 79, 8))
               CAREA# = CVD(Mid$(XX$, 87, 8))
               SALDOI = CAPROD# - CAREA#: If SALDOI < 0 Then SALDOI = 0
               TTOT = CVS(Mid$(XX$, 131, 4))
               FEXDES = CVI(Mid$(XX$, 135, 2))
               HDES% = CVI(Mid$(XX$, 137, 2))
               FEXHAS = CVI(Mid$(XX$, 139, 2))
               HHAS% = CVI(Mid$(XX$, 141, 2))
               '
               CAITAB1% = CAITAB1% + 1
               TETABU1$(1, CAITAB1%) = NOFAX$
               TETABU1$(2, CAITAB1%) = TRIM$(CODEXT$(CIXI%)) + " - " + DESCRIT0$(CIXI%)
               TETABU1$(3, CAITAB1%) = Str$(SALDOI)
               HORASIG$ = TRIM$(HORMINSE$(TTOT))
               If Left$(HORASIG$, 2) = "0:" Then HORASIG$ = "  " + Mid$(HORASIG$, 3)
               TETABU1$(4, CAITAB1%) = HORASIG$
               TETABU1$(5, CAITAB1%) = FECHATEX$(FEXDES)
               TETABU1$(6, CAITAB1%) = HORATEX$(HDES%)
               TETABU1$(9, CAITAB1%) = FECHATEX$(FEXHAS)
                 If TETABU1$(9, CAITAB1%) = TETABU1$(5, CAITAB1%) Then TETABU1$(9, CAITAB1%) = "---id---"
               TETABU1$(7, CAITAB1%) = HORATEX$(HHAS%)
               TETABU1$(8, CAITAB1%) = Mid$(XX$, 225, 6)
               '
             End If
           Next JJ&
           '
           Call PRINTDOC("PLACARMA")   ' planilla de carga
           '
        End If
     Next I%
     '
     Screen.MousePointer = 1
     Return
     '
End Sub
'

Sub SECUORF()  ' secuencia de ordenes de fabricacion
    '
    For UFORMU% = 1 To Forms.Count
      FORNA$ = TRIM$(UCase$(Forms(UFORMU% - 1).Name))
      If FORNA$ = "FORCAMAQ" Or FORNA$ = "CALCUCAR" Then
        On Error Resume Next
        FECOMPRO% = FECHANUM(Forms(UFORMU% - 1).Text1(9).TEXT)
        If Err < 1 Then GoTo 5
      End If
    Next UFORMU%
    Call MENSERR(24, "Proceso Imposible")
    Exit Sub
        
5   IFORCA% = UFORMU% - 1
    On Error GoTo 0
    '
    FIN& = ULTREG&("SECUORF")
    If FIN& > 4096 Then
        Call MENSERR(24, "Proceso Imposible - Vector fuera de Rango - Consulte")
        Exit Sub
    End If
    '
    SECCI = (-1): SEPROD$ = ""
    URESEC& = ULTREG&("SECPROD")
    If URESEC& > 0 Then
      If URESEC& = 1 Then
          XX$ = REGLEIDO$("SECPROD", 1)
        Else
          XX$ = REGSEL$("SECPROD")
      End If
      If Len(XX$) >= 4 Then
        SECCI = CVS(Left$(XX$, 4))
        SEPROD$ = Mid$(XX$, 5, 28)
      End If
    End If
    '
    Screen.MousePointer = 11
    For I& = 1 To FIN&
        X$ = REGLEIDO$("SECUORF", I&)
        ORFASEQ$(I&) = Left$(X$, 12)
        PRIVI%(I&) = 0
        STATU%(I&) = 0
    Next I&
    CAORFA% = FIN&
    CAORFB& = CAORFA%
    '
    FIN& = ULTREG&("ORDEFABR")
    For I& = 1 To FIN&
        F$ = REGLEIDO$("ORDEFABR", I&)
        ST% = Asc(Mid$(F$, 108, 1))
        If ST% = 1 Or ST% = 2 Then ' esta lanzada
            ORFASE$ = Mid$(F$, 69, 12)
            For J& = 1 To CAORFA%
                If ORFASEQ$(J&) = ORFASE$ Then
                    STATU%(J&) = ST%
                    If CVS(Mid$(F$, 113, 4)) = SECCI Then PRIVI%(J&) = 1
                    GoTo 20
                End If
            Next J&
            If ST% = 1 Or ST% = 2 Then
                DORFA$ = Mid$(F$, 5, 64)
                CAORFB& = CAORFB& + 1
                ORFASEQ$(CAORFB&) = ORFASE$
                STATU%(CAORFB&) = ST%
                PRIVI%(CAORFB&) = 0
                If CVS(Mid$(F$, 113, 4)) = SECCI Then PRIVI%(CAORFB&) = 1
                Call GRAREG("SECUORF", ORFASE$ + DORFA$, CAORFB&)
            End If
        End If
20  Next I&
    '
    J& = 0
    LISECUORF.LISECU.Clear
    For I& = 1 To CAORFB&
        If STATU%(I&) = 1 Then
            F$ = REGLEIDO$("SECUORF", I&)
            F$ = Left$(F$, 12) + Space$(6) + Mid$(F$, 13)
            If PRIVI%(I&) > 0 Then
              Call REPLA(F$, "<=>", 14, 4)
            End If
            LISECUORF.LISECU.AddItem F$
        End If
    Next I&
    '
    If Forms(IFORCA%).Check7.Value = 1 Then
       For I& = 1 To CAORFB&   ' incluye suspendidas
          If STATU%(I&) = 2 Then
             F$ = REGLEIDO$("SECUORF", I&)
             F$ = Left$(F$, 12) + Space$(6) + Mid$(F$, 13)
             If PRIVI%(I&) > 0 Then
               Call REPLA(F$, "<=>", 14, 4)
             End If
             LISECUORF.LISECU.AddItem F$
          End If
       Next I&
    End If
    '
    Screen.MousePointer = 1
    LISECUORF.Caption = "Secuenciación de Ordenes de Fabricación"
    If SEPROD$ <> "" Then
      LISECUORF.Caption = LISECUORF.Caption + " - Resalta: " + SEPROD$
    End If
    LISECUORF.Show 1
    '
End Sub
'

Sub CARHORAEQ()
    '
10  OBX$ = OBJPLA$("ENTRAEQUIPO", "")
    If OBX$ = "" Then
        Exit Sub
    End If
    '
    EQUI$ = AJUSTI$(OBX$, 6)
    YAESTA% = 0
    Call BLANARCH("!HORAEQUI")
    For U& = 1 To ULTREG&("HORAEQUI")
        X$ = REGLEIDO$("HORAEQUI", U&)
        If Left$(X$, 6) = EQUI$ Then
            YAESTA% = 1
            Call REGAPP("!HORAEQUI", X$)
        End If
    Next U&
    '
    If YAESTA% < 1 Then
        For UI% = 1 To 7
            X$ = REGBLAN$("HORAEQUI")
            Call REPLA(X$, Chr$(UI%), 7, 1)
            Call REPLA(X$, ECOARCH$("DIASEM", Chr$(UI%)), 8, 15)
            Call REGAPP("!HORAEQUI", X$)
        Next UI%
    End If
    Call CIERRARCH("!HORAEQUI")
    '
    VTB% = VENTABU%("HORAESP/TITUPAN=Horarios Especiales: " + TRIM$(EQUI$) + " (" + TRIM$(ECOARCH$("PADMAQ", EQUI$)) + ")")
    If VTB% < 0 Then GoTo 10
    '
    V& = 0
    For U& = 1 To ULTREG&("HORAEQUI")
        X$ = REGLEIDO$("HORAEQUI", U&)
        If Left$(X$, 6) <> EQUI$ Then
            V& = V& + 1
            Call GRAREG("HORAEQUI", X$, V&)
        End If
    Next U&
    Call SETULTREG("HORAEQUI", V&)
    Call CIERRARCH("HORAEQUI")
    '
    For U& = 1 To 7
        X$ = REGLEIDO$("!HORAEQUI", U&)
        TDES% = CVI(Mid$(X$, 23, 2))
        THAS% = CVI(Mid$(X$, 25, 2))
        If THAS% > TDES% Then GoTo 100
    Next U&
    GoTo 10
    '
100 For U& = 1 To 7
        X$ = REGLEIDO$("!HORAEQUI", U&)
        Call REPLA(X$, EQUI$, 1, 6)
        Call REGAPP("HORAEQUI", X$)
    Next U&
    Call CIERRARCH("HORAEQUI")
    GoTo 10
    '
End Sub
'
Sub REPLAMAQ(MAQUIBAS$, TTOT, MAQUIALT$, TALT)
    '
    MAQUIALT$ = ""
    TALT = TTOT
    COEFMAX = 0
    For IMQ% = 0 To 3
       '
       maquiba$ = Mid$(MAQUIBAS$, 6 * IMQ% + 1, 6)
       maquiasi$ = maquiba$
       COEFITIE = 0
       '
5      If TRIM$(maquiba$) <> "" Then
          '
          'For JJ% = 1 To Len(MAQUISE$) Step 6
          '   If Mid$(MAQUISE$, JJ%, 6) = MAQUIBA$ Then
          '      COEFITIE = 100
          '      GoTo 90
          '   End If
          'Next JJ%
          '
          ALTERMAQ$ = ""
          For JJ% = 1 To NUMAQ%
             If MAQUI$(JJ%) = maquiba$ Then
                TIMAQBA = TIEPRO(JJ%)
                ALTERMAQ$ = MAEQUI$(JJ%)
                COEFALT = COEFTIE(JJ%)
                GoTo 10
             End If
          Next JJ%
          '
10        COEFITIE = 100
          If TRIM$(ALTERMAQ$) <> "" Then
             For JJ% = 1 To NUMAQ%
                If MAQUI$(JJ%) = ALTERMAQ$ Then
                   TIMAQAL = TIEPRO(JJ%)
                End If
             Next JJ%
             '
             If TIMAQAL + TTOT * COEFALT / 100 < TIMAQBA + TTOT Then
                maquiasi$ = ALTERMAQ$
                COEFITIE = COEFALT
                maquiba$ = maquiasi$
                GoTo 5
             End If
             '
          End If
          '
       End If
       '
90     MAQUIALT$ = MAQUIALT$ + AJUSTI$(maquiasi$, 6)
       If TRIM$(maquiasi$) <> "" Then
          For JJ% = 1 To Len(MAQUISE$) Step 6
             If Mid$(MAQUISE$, JJ%, 6) = maquiasi$ Then GoTo 95
          Next JJ%
          MAQUISE$ = MAQUISE$ + maquiasi$
       End If
       '
95     If COEFITIE > COEFMAX Then
          COEFMAX = COEFITIE
       End If
       '
    Next IMQ%
    '
    TALT = TTOT
    If COEFMAX > 0 Then
       TALT = TTOT * COEFMAX / 100
    End If
    '
'If MAQUIALT$ <> MAQUIBAS$ Then
'MsgBox MAQUIBAS$ & "   " & MAQUIALT$
'End If
'
End Sub
