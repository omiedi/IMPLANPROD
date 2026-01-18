Attribute VB_Name = "COBERST1"
Public MODOCOBER%

Sub COBERSTO()
     'OK MOVISVIE
     ' generador analisis de cobertura de stocks
     '
     ' version modificada el 12/01/99 para controlar las diferencias entre
     ' subdiario de ventas y resumen estadistico
     '
     Dim NC1%(16384), TI1%(16384), NRO&(16384), FF1%(16384) ' dimensionamiento
     Dim IFAX#(16384), IDEX#(16384)
     Dim FEPRO%(12), CAPRO(12)
     '
200  OPX$ = OBJPLA$("DESHASFECHA", VDEF$)
     If OPX$ = "" Then Exit Sub
     DES% = CVI(Left$(OPX$, 2))
     HAS% = CVI(Mid$(OPX$, 3, 2))
     If DES% > HAS% Then Exit Sub
     DES1 = DES%: HAS1 = HAS%
     CORRESP$ = FECHATEX$(DES1) + " - " + FECHATEX$(HAS1)
     If CORRESP$ = "" Then Exit Sub
     '
     FEX = DES%: DESX$ = FECHATEX$(FEX)
     FEX = HAS%: HASX$ = FECHATEX$(FEX)
     CORR1$ = TRIM$(CORRESP$)
     HOII% = HOY(HO$)
     '
     ME1 = Val(Mid$(DESX$, 4, 2)): AN1 = Val(Mid$(DESX$, 7, 2)): If AN1 < 80 Then AN1 = AN1 + 100
     ME2 = Val(Mid$(HASX$, 4, 2)): AN2 = Val(Mid$(HASX$, 7, 2)): If AN2 < 80 Then AN2 = AN2 + 100
     CAMESES% = 12 * (AN2 - AN1) + (ME2 - ME1) + 1
     If CAMESES% < 1 Then
        Call CIERRARCH("*.*")
        Exit Sub   'GoTo 3900
     End If
     '
     'M0% = Val(Mid$(HO$, 4, 2)): A0% = Val(Mid$(HO$, 7, 2))
     M0% = Val(Mid$(HASX$, 4, 2)): A0% = Val(Mid$(HASX$, 7, 2))
     A1% = A0%: M1% = M0% - 1: If M1% < 1 Then M1% = M1% + 12: A1% = A1% - 1
     A2% = A1%: M2% = M1% - 1: If M2% < 1 Then M2% = M2% + 12: A2% = A2% - 1
     '
     M0S$ = TRIM$(Str$(M0%)): While Len(M0S$) < 2: M0S$ = "0" + M0S$: Wend
     A0S$ = TRIM$(Str$(A0%)): While Len(A0S$) < 2: A0S$ = "0" + A0S$: Wend
     DX0$ = "01/" + M0S$ + "/" + A0S$: DX0I% = FECHANUM(DX0$)
     HX0$ = "31/" + M0S$ + "/" + A0S$: HX0I% = FECHANUM(HX0$)
     '
     M1S$ = TRIM$(Str$(M1%)): While Len(M1S$) < 2: M1S$ = "0" + M1S$: Wend
     A1S$ = TRIM$(Str$(A1%)): While Len(A1S$) < 2: A1S$ = "0" + A1S$: Wend
     DX1$ = "01/" + M1S$ + "/" + A1S$: DX1I% = FECHANUM(DX1$)
     HX1$ = "31/" + M1S$ + "/" + A1S$: HX1I% = FECHANUM(HX1$)
     '
     M2S$ = TRIM$(Str$(M2%)): While Len(M2S$) < 2: M2S$ = "0" + M2S$: Wend
     A2S$ = TRIM$(Str$(A2%)): While Len(A2S$) < 2: A2S$ = "0" + A2S$: Wend
     DX2$ = "01/" + M2S$ + "/" + A2S$: DX2I% = FECHANUM(DX2$)
     HX2$ = "31/" + M2S$ + "/" + A2S$: HX2I% = FECHANUM(HX2$)
     '
     Screen.MousePointer = 11
     Call BLOQARCH("COBERST")
     Call BLOQARCH("PROREPO")
     '
     FIN& = NUMAS%
     FORCOBERSTO.Label6.Caption = "Inicializando Archivo de Cobertura"
     FORCOBERSTO.Label6.Refresh
     For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        MX0$ = REGLEIDO$("MASTER", I&)
        COU = CVS(Mid$(MX0$, 111, 4))
        CISX% = I&
        XXX$ = MKI$(CISX%) + String$(62, 0)
        Call REPLA(XXX$, MKS$(COU), 3, 4)
        Call GRAREG("!COBERST", XXX$, I&)
        Call GRAREG("!PROREPO", String$(30, 0) + MKI$(CISX%), I&)
     Next I&
     '
     Call CIERRARCH("MASTER")
     Call SETULTREG("!COBERST", FIN&)
     Call SETULTREG("!PROREPO", FIN&)
     '
     FORCOBERSTO.Label6.Caption = "Revisando Pedidos Pendientes"
     FORCOBERSTO.Label6.Refresh
     FIN& = ULTREG&("PEDDETP")
     For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        XXX$ = REGLEIDO$("PEDDETP", I&)
        NC% = CVI(Mid$(XXX$, 7, 2))
        CI% = CVI(Mid$(XXX$, 33, 2))
        If CI% > 0 Then
           If CI% <= NUMAS% Then
              CAPEN = CVS(Mid$(XXX$, 79, 4)) - CVS(Mid$(XXX$, 95, 4))
              If CAPEN > 0.05 Then
                 CIL& = CI%
                 CBX$ = REGLEIDO$("!COBERST", CIL&)
                 Y$ = CBX$
                 CAPENDE = CVS(Mid$(Y$, 43, 4)) + CAPEN
                 Call REPLA(Y$, MKS$(CAPENDE), 43, 4)
                 Call GRAREG("!COBERST", Y$, CIL&)
                 If NC% < 3000 Or (NC% > 3999 And NC% < 5000) Then
                    CIL& = CI%
                    CBY$ = REGLEIDO$("!PROREPO", CIL&)
                    Y$ = CBY$
                    CAPENDE = CVS(Mid$(Y$, 1, 4)) + CAPEN
                    Call REPLA(Y$, MKS$(CAPENDE), 1, 4)
                    Call GRAREG("!PROREPO", Y$, CIL&)
                 End If
              End If
           End If
        End If
     Next I&
     '
     Call CIERRARCH("PEDDETP")
     '
     GoSub 11000                               ' seleccionar facturas
     If CAFAC% < 1 Then                           '       del periodo elegido
         Call COMUNI("No hay Facturación\en el Período Elegido")
         Call CIERRARCH("*.*")
         Screen.MousePointer = 1
         Exit Sub  ' GoTo 3900
     End If
     '
     Dim OTMO#(5)
     GoSub 11100                               ' cargar facturas a vector
     GoSub 11200                               ' coeficiente de precios
     '
     FORCOBERSTO.Label6.Caption = "Revisando Movimientos de Stock"
     FORCOBERSTO.Label6.Refresh
     '
     ARMOVI$ = "MOVISTO"
     '
300  FIN& = ULTREG&(ARMOVI$)
     For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       MVX$ = REGLEIDO$(ARMOVI$, I&)
       XXX$ = MVX$
       '
       CI% = CVI(Mid$(XXX$, 3, 2))
       If CI% > 0 Then
         If CI% <= NUMAS% Then
           NC% = CVI(Mid$(XXX$, 84, 2))
           FF% = CVI(Left$(XXX$, 2))
           If FF% <= HAS% Then
               CAN = CVD(Mid$(XXX$, 96, 8)) - CVD(Mid$(XXX$, 88, 8))
               PRE# = CVD(Mid$(XXX$, 73, 8))
               TOT# = PRE# * CAN
               '
               CIL& = CI%
               CBX$ = REGLEIDO$("!COBERST", CIL&)
               Y$ = CBX$
               CATO = CVS(Mid$(Y$, 7, 4)) - CAN
               VATO# = CDbl(CATO * CVS(Mid$(Y$, 3, 4)))
               '
               Call REPLA(Y$, MKS$(CATO), 7, 4)
               Call REPLA(Y$, MKD$(VATO#), 11, 8)
               '
               COMO$ = UCase$(Mid$(XXX$, 33, 2))
               If COMO$ = "FC" Or COMO$ = "NC" Or COMO$ = "ND" Then
                 If FF% >= DX0I% And FF% <= HX0I% Then
                   CANX = CVS(Mid$(Y$, 59, 4)) + CAN
                   Call REPLA(Y$, MKS$(CANX), 59, 4)
                   If NC% > 0 Then
                      If NC% < 3000 Or (NC% > 3999 And NC% < 5000) Then
                         CIL& = CI%
                         CBY$ = REGLEIDO$("!PROREPO", CIL&)
                         YY$ = CBY$
                         CANY = CVS(Mid$(YY$, 5, 4)) + CAN
                         Call REPLA(YY$, MKS$(CANY), 5, 4)
                         Call GRAREG("!PROREPO", YY$, CIL&)
                      End If
                   End If
                 ElseIf FF% >= DX1I% And FF% <= HX1I% Then
                   CANX = CVS(Mid$(Y$, 55, 4)) + CAN
                   Call REPLA(Y$, MKS$(CANX), 55, 4)
                   If NC% > 0 Then
                      If NC% < 3000 Or (NC% > 3999 And NC% < 5000) Then
                         CIL& = CI%
                         CBY$ = REGLEIDO$("!PROREPO", CIL&)
                         YY$ = CBY$
                         CANY = CVS(Mid$(YY$, 9, 4)) + CAN
                         Call REPLA(YY$, MKS$(CANY), 9, 4)
                         Call GRAREG("!PROREPO", YY$, CIL&)
                      End If
                   End If
                 ElseIf FF% >= DX2I% And FF% <= HX2I% Then
                   CANX = CVS(Mid$(Y$, 51, 4)) + CAN
                   Call REPLA(Y$, MKS$(CANX), 51, 4)
                   If NC% > 0 Then
                      If NC% < 3000 Or (NC% > 3999 And NC% < 5000) Then
                         CIL& = CI%
                         CBY$ = REGLEIDO$("!PROREPO", CIL&)
                         YY$ = CBY$
                         CANY = CVS(Mid$(YY$, 13, 4)) + CAN
                         Call REPLA(YY$, MKS$(CANY), 13, 4)
                         Call GRAREG("!PROREPO", YY$, CIL&)
                      End If
                   End If
                 End If
               End If
               '
               If FF% >= DES% Then
                 If FF% <= HAS% Then
                   COMO$ = UCase$(Mid$(XXX$, 33, 2))
                   If COMO$ = "FC" Or COMO$ = "NC" Or COMO$ = "ND" Then
                     NC% = CVI(Mid$(XXX$, 84, 2))
                     NRO1& = Val(Mid$(XXX$, 36, 7))
                     GoSub 12100
                     If COEF > 1.005 Or ENCU% = 0 Then COEF = 1
                     CAVE = CVS(Mid$(Y$, 19, 4)) + CAN
                     VAVE# = CVD(Mid$(Y$, 23, 8)) + TOT# * COEF
                     '
                     Call REPLA(Y$, MKS$(CAVE), 19, 4)
                     Call REPLA(Y$, MKD$(VAVE#), 23, 8)
                     '
                     If NC% > 0 Then
                        If NC% < 3000 Or (NC% > 3999 And NC% < 5000) Then
                         CIL& = CI%
                         CBY$ = REGLEIDO$("!PROREPO", CIL&)
                           YY$ = CBY$
                           CANY = CVS(Mid$(YY$, 17, 4)) + CAN
                           Call REPLA(YY$, MKS$(CANY), 17, 4)
                           Call GRAREG("!PROREPO", YY$, CIL&)
                        End If
                     End If
                   End If
                 End If
               End If
               '
               CIL& = CI%
               Call GRAREG("!COBERST", Y$, CIL&)
               '
            End If
         End If
       End If
     Next I&
     If ARMOVI$ = "MOVISTO" Then
       If EXISTE%(LUDAT$ + "MOVISVIE.DAT") > 0 Then
         ARMOVI$ = "MOVISVIE"
         GoTo 300
       End If
     End If
     '
     FORCOBERSTO.Label6.Caption = "Calculando Promedios y Cobertura"
     FORCOBERSTO.Label6.Refresh
     FIN& = ULTREG&("!COBERST")
     For I& = 1 To FIN&
           Call TRACE(20, I&, FIN&)
           CBX$ = REGLEIDO$("!COBERST", I&)
           Y$ = CBX$
           CATO = CVS(Mid$(Y$, 7, 4))
           VATO# = CDbl(CATO * CVS(Mid$(Y$, 3, 4)))
           CAVE = CVS(Mid$(Y$, 19, 4))
           VAVE# = CVD(Mid$(Y$, 23, 8))
           CAPE = CVS(Mid$(Y$, 43, 4))
           PROMECA = CAVE / CAMESES%
           PROMEVA# = VAVE# / CAMESES%
           '
           COBER = 999
           If PROMECA >= 1 Then
               COBER = (CATO - CAPE) / PROMECA
           End If
           If COBER < 0 Then COBER = 0
           If COBER > 999 Then COBER = 999
           '
           Call REPLA(Y$, MKS$(PROMECA), 31, 4)
           Call REPLA(Y$, MKD$(PROMEVA#), 35, 8)
           Call REPLA(Y$, MKS$(COBER), 47, 4)
           '
           Call GRAREG("!COBERST", Y$, I&)
           '
           CBY$ = REGLEIDO$("!PROREPO", I&)
           YY$ = CBY$
           PROMECA = (CVS(Mid$(CBY$, 17, 4))) / CAMESES%
           MESANT1 = CVS(Mid$(CBY$, 9, 4))
           MESANT2 = CVS(Mid$(CBY$, 13, 4))
           CAMAX = PROMECA
           If MESANT1 > CAMAX Then CAMAX = MESANT1
           If MESANT2 > CAMAX Then CAMAX = MESANT2
           Call REPLA(YY$, MKS$(CAMAX), 21, 4)
           Call GRAREG("!PROREPO", YY$, I&)
           '
     Next I&
     '
     Call GRARGEN("!COBERST", TRIM$(CORRESP$))
     FORCOBERSTO.Label1.Caption = TRIM$(CORRAR$("!COBERST"))
     FEX = FEGEAR%("!COBERST")
     FORCOBERSTO.Label21.Caption = FECHATEX$(FEX)
     FORCOBERSTO.Label22.Caption = HORATEX$(HOGEAR%("!COBERST"))
     FORCOBERSTO.Label1.Refresh
     FORCOBERSTO.Label21.Refresh
     FORCOBERSTO.Label22.Refresh
     HOII% = HOY(HO$)
     '
     Call HEADERS("COBERFI", "")
     Call HEADERS("COBERFI", "Corresponde a: " + CORRESP$)
     Call HEADERS("COBERFI", "Estado de Stock al: " + HASX$ + " / Pedidos al " + HO$ + " - " + Time$ + " hs")
     '
     Call CIERRARCH("*.*")
     '
     JK& = 0
     FIN& = NUINV%
     For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        II$ = REGLEIDO$("INVERT", U& + 1)
        RECO& = CVI(Mid$(II$, 17, 2))
        If RECO& > 0 Then
           If RECO& <= NUMAS% Then
              JK& = JK& + 1
              XX$ = REGLEIDO$("!COBERST", RECO&)
              Call GRAREG("COBERST", XX$, JK&)
              XX$ = REGLEIDO$("!PROREPO", RECO&)
              Call GRAREG("PROREPO", XX$, JK&)
           End If
        End If
     Next U&
     Call SETULTREG("COBERST", JK&)
     Call CIERRARCH("COBERST")
     Call SETULTREG("PROREPO", JK&)
     Call CIERRARCH("PROREPO")
     '
     Screen.MousePointer = 1
     Exit Sub
     '

11000   Screen.MousePointer = 11
        '
        FORCOBERSTO.Label6.Caption = "Revisando Archivo de Facturación"
        FORCOBERSTO.Label6.Refresh
        FIN& = ULTREG&("FACTUR")
        CAFAC% = 0
        PRIRE& = 999999: ULTRE& = 0
        '
        For I& = 1 To FIN&
           Call TRACE(20, I&, FIN&)
           XXX$ = REGLEIDO$("FACTUR", I&)
           FF% = CVI(Mid$(XXX$, 7, 2))
           If FF% >= DES% Then
             If FF% <= HAS% Then
               CAFAC% = CAFAC% + 1
               If I& < PRIRE& Then PRIRE& = I&
               If I& > ULTRE& Then ULTRE& = I&
             End If
           End If
        Next I&
        '
        If CAFAC% < 1 Then
           Call COMUNI("Proceso no Realizable !!!\No hay Facturas Registradas\en el Periodo Elegido.")
        End If
        '
        Return
        '
        '
11100   CAFAC% = 0
        SUFAC# = 0
        For I& = PRIRE& To ULTRE&
           Call TRACE(20, I&, ULTRE&)
           XXX$ = REGLEIDO$("FACTUR", I&)
           FF% = CVI(Mid$(XXX$, 7, 2))
           If FF% >= DES% Then
               If FF% <= HAS% Then
                   CAFAC% = CAFAC% + 1
                   NC1%(CAFAC%) = CVI(Left$(XXX$, 2))
                   FF1%(CAFAC%) = FF%
                   TI1%(CAFAC%) = Asc(Mid$(XXX$, 13, 1))
                   If TI1%(CAFAC%) < 0 Or TI1%(CAFAC%) > 6 Then TI1%(CAFAC%) = 2 ' por default factura
                   NRO&(CAFAC%) = CVS(Mid$(XXX$, 3, 4))
                   '
                   If TI1%(CAFAC%) = 0 Then                     ' contempla ajustes de saldos
                      If CVD(Mid$(XXX$, 15, 8)) < 0 Then
                         TI1%(CAFAC%) = 7
                      End If
                   End If
                   '
                   IFAC# = Abs(CVD(Mid$(XXX$, 15, 8)))
                   IFAX#(CAFAC%) = IFAC#
                   '
                   SIG% = 1: If TI1%(CAFAC%) > 3 Then SIG% = (-1)
                   SUFAC# = SUFAC# + IFAC# * SIG%
               End If
           End If
        Next I&
        '
        Return
        '
11200   FORCOBERSTO.Label6.Caption = "Revisando Movimientos de Stock"
        FORCOBERSTO.Label6.Refresh
        'PRIRE& = 999999: ULTRE& = 0
        '
        ARMOVI$ = "MOVISTO"
        '
11300   FIN& = ULTREG&(ARMOVI$)
        For I& = 1 To FIN&
           Call TRACE(20, I&, FIN&)
           XXX$ = REGLEIDO$(ARMOVI$, I&)
           FF% = CVI(Left$(XXX$, 2))
           If FF% >= DES% Then
               If FF% <= HAS% Then
                   COMO$ = UCase$(Mid$(XXX$, 33, 2))
                   If COMO$ = "FC" Or COMO$ = "NC" Or COMO$ = "ND" Then
                       TP1% = 2: If COMO$ = "NC" Then TP1% = 6
                              If COMO$ = "ND" Then TP1% = 3
                       NRO1& = Val(Mid$(XXX$, 36, 7))
                       CAND# = CVD(Mid$(XXX$, 96, 8)) - CVD(Mid$(XXX$, 88, 8))
                       PRE# = CVD(Mid$(XXX$, 73, 8))
                       TOT# = Abs(PRE# * CAND#)
                       'If I& < PRIRE& Then PRIRE& = I&
                       'If I& > ULTRE& Then ULTRE& = I&
                       GoSub 12000
                   End If
               End If
           End If
        Next I&
        '
        If ARMOVI$ = "MOVISTO" Then
          If EXISTE%(LUDAT$ + "MOVISVIE.DAT") > 0 Then
            XXX$ = REGLEIDO$("MOVISTO", 1)
            FF% = CVI(Left$(XXX$, 2))
            If FF% >= DES% Then
              ARMOVI$ = "MOVISVIE"
              GoTo 11300
            End If
          End If
        End If
        '
        Return
        '
        '
12000   For JJ% = 1 To CAFAC%
          If FF1%(JJ%) = FF% Then
             If TI1%(JJ%) = TP1% Then
                 If NRO&(JJ%) = NRO1& Then
                    IDET# = IDEX#(JJ%)
                    IDET# = IDET# + TOT#
                    IDEX#(JJ%) = IDET#
                    GoTo 12090
                 End If
             End If
          End If
        Next JJ%
        '
12090   Return

12100   COEF = 1: ENCU% = 0
        For JJ% = 1 To CAFAC%
          If FF1%(JJ%) = FF% Then
             If NRO&(JJ%) = NRO1& Then
                 ENCU% = 1
                 IFAC# = IFAX#(JJ%)
                 IDET# = IDEX#(JJ%)
                 If IDET# > 0 Then
                     COEF = IFAC# / IDET#
                 End If
                 GoTo 12190
             End If
          End If
        Next JJ%
        '
12190   Return
'
End Sub
'

Sub GEPROGEN()
     '
     Dim FEPRO%(12)
     '
     Screen.MousePointer = 11
     FORCOBERSTO.Label6.Caption = "Generando Programa de Entregas"
     FORCOBERSTO.Label6.Refresh
     '
500  M0% = Val(Mid$(HO$, 4, 2)): A0% = Val(Mid$(HO$, 7, 2))
     For K% = 1 To 12
         M0S$ = TRIM$(Str$(M0%)): While Len(M0S$) < 2: M0S$ = "0" + M0S$: Wend
         A0S$ = TRIM$(Str$(A0%)): While Len(A0S$) < 2: A0S$ = "0" + A0S$: Wend
         DX0S$ = "01/" + M0S$ + "/" + A0S$: FEPRO%(K%) = FECHANUM(DX0S$)
         If FEPRO%(K%) < HOII% Then FEPRO%(K%) = HOII%
         '
         M0% = M0% + 1
         If M0% > 12 Then
             M0% = 1
             A0% = A0% + 1
             If A0% >= 100 Then A0% = A0% - 100
         End If
         '
     Next K%
     '
     J& = 0
     For I& = 1 To ULTREG&("PROGENT")
        XXX$ = REGLEIDO$("PROGENT", I&)
        NC% = CVI(Mid$(XXX$, 5, 2))
        If NC% >= 5000 Or (NC% >= 3000 And NC% <= 3999) Then
           J& = J& + 1
           Call GRAREG("PROGENT", XXX$, J&)
        End If
     Next I&
     '
     FIN& = ULTREG&("PROREPO")
     For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        Y$ = REGLEIDO$("PROREPO", I&)
        PROMECA = CVS(Mid$(Y$, 21, 4))
        PROMESA = PROMECA - CVS(Mid$(Y$, 5, 4))
        If CVS(Left$(Y$, 4)) > PROMESA Then PROMESA = CVS(Left$(Y$, 4))
        If PROMECA >= 50 Or PROMESA >= 50 Then
            CI% = I&
            For K% = 1 To 4
                CAPROG = PROMECA
                If K% = 1 Then CAPROG = PROMESA
                XXX$ = REGBLAN$("PROGENT")
                Call REPLA(XXX$, MKI$(CI%), 1, 2)
                Call REPLA(XXX$, MKI$(FEPRO%(K%)), 3, 2)
                Call REPLA(XXX$, MKS$(CAPROG), 7, 4)
                Call REPLA(XXX$, MKS$(CAPROG), 15, 4)
                J& = J& + 1
                Call GRAREG("PROGENT", XXX$, J&)
            Next K%
        End If
     Next I&
     '
     Call SETULTREG("PROGENT", J&)
     Call CIERRARCH("PROGENT")
     Call CIERRARCH("*.*")
     Screen.MousePointer = 1

End Sub
'

Sub ANACOBER()
   '
   If MODOCOBER% = 2 Then
      MECODE = Val(FORCOBERSTO.Text5.TEXT)
      If MECODE < 1 Or MECODE > 99 Then
         Call COMUNI("Falta Indicar Meses de Cobertura Requeridos")
         Call CIERRARCH("*.*")
         Exit Sub
       End If
   End If
   '
   Screen.MousePointer = 11
   Dim ACTI%(16384)
   '
   TIPOMA% = Val(FORCOBERSTO.Text1.TEXT)
   ORIGMA% = Val(FORCOBERSTO.Text3.TEXT)
   LIPREVA% = Val(FORCOBERSTO.Text6.TEXT)
   '
   FORCOBERSTO.Label6.Caption = "Seleccionando por Tipo y Origen"
   FORCOBERSTO.Label6.Refresh
   FIN& = NUINV%
   '
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      Y$ = REGLEIDO$("INVERT", U& + 1)
      If Mid$(Y$, 16, 1) <> Chr$(96) Then
         W& = CVI(Mid$(Y$, 17, 2))
         If W& > 0 Then
            If W <= NUMAS% Then
               X$ = REGLEIDO$("MASTER", W& + 1)
               CLASI1% = Asc(Mid$(X$, 69, 1))
               CLASI4% = Asc(Mid$(X$, 72, 1))
               '
               If CLASI1% = TIPOMA% Or TIPOMA% = 0 Then
                  If CLASI4% = ORIGMA% Or ORIGMA% = 0 Then
                     CIIX% = CVI(Left$(X$, 2))
                     ACTI%(CIIX%) = 1
                     If LIPREVA% = 0 Then
                        ACTI%(CIIX%) = 2
                     End If
                  End If
               End If
            End If
         End If
      End If
   Next U&
   '
   If LIPREVA% > 0 Then
      FORCOBERSTO.Label6.Caption = "Seleccionando por Lista de Precios"
      FORCOBERSTO.Label6.Refresh
      ALISTA$ = TRIM$(Str$(LIPREVA%))
      While Len(ALISTA$) < 3
         ALISTA$ = "0" + ALISTA$
      Wend
      ALISTA$ = "LIPRE" + ALISTA$
      '
      FIN& = ULTREG&(ALISTA$)
      '
      For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
         X$ = REGLEIDO$(ALISTA$, U&)
         CIIX% = CVI(Left$(X$, 2))
         If ACTI%(CIIX%) > 0 Then
            ACTI%(CIIX%) = 2
         End If
      Next U&
   End If
   '
   FORCOBERSTO.Label6.Caption = "Generando Archivo de Salida"
   FORCOBERSTO.Label6.Refresh
   Call BLANARCH("COBERFI")
   FIN& = ULTREG&("COBERST")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("COBERST", U&)
      CIIX% = CVI(Left$(X$, 2))
      If ACTI%(CIIX%) = 2 Then
         '
         If MODOCOBER% = 2 Then
            PROMECA = CVS(Mid$(X$, 31, 4))
            PEDPEN = CVS(Mid$(X$, 43, 4))
            If PROMECA < 0.05 And PEDPEN < 0.05 Then
               GoTo 99
            End If
            If CVS(Mid$(X$, 47, 4)) > MECODE Then
               GoTo 99
            End If
         End If
         '
         Call REGAPP("COBERFI", X$)
      End If
99 Next U&
   '
   Call CIERRARCH("*.*")
   FORCOBERSTO.Label6.Caption = "Selección Completada"
   FORCOBERSTO.Label6.Refresh
   Screen.MousePointer = 1
   '
   If FORCOBERSTO.Option5.Value = True Then
        Call FINAL("*COCOBER")
      Else
        Call FINAL("*LICOBER")
   End If
   '
End Sub
